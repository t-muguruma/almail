import os
import email
from email import policy
import sqlite3
import imaplib
import poplib
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from email.header import decode_header, Header
import re
import traceback
from pathlib import Path

def setup_database(db_path="mailbox.db"):
    """メールデータを保存するローカルDBの作成"""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            message_id TEXT,
            subject TEXT,
            sender TEXT,
            date TEXT,
            body TEXT,
            body_html TEXT,
            file_path TEXT,
            folder TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT,
            protocol TEXT DEFAULT 'IMAP',
            smtp_server TEXT,
            smtp_port INTEGER,
            imap_server TEXT,
            imap_port INTEGER,
            username TEXT,
            password TEXT,
            display_html_as_text INTEGER DEFAULT 0,
            minimize_to_tray INTEGER DEFAULT 0,
            notify_new_mail INTEGER DEFAULT 0,
            auto_receive_enabled INTEGER DEFAULT 0,
            auto_receive_interval INTEGER DEFAULT 10,
            signature TEXT,
            search_almail_at_startup INTEGER DEFAULT 0
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS address_book (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            email TEXT,
            nickname TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            message_id INTEGER,
            filename TEXT,
            data BLOB
        )
    ''')
    
    # 既存のDBに対するマイグレーション（カラム追加チェック）
    def add_column_if_not_exists(table, column, definition):
        cursor.execute(f"PRAGMA table_info({table})")
        if column not in [row[1] for row in cursor.fetchall()]:
            cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")

    add_column_if_not_exists("messages", "body_html", "TEXT")
    add_column_if_not_exists("accounts", "display_html_as_text", "INTEGER DEFAULT 0")
    add_column_if_not_exists("accounts", "minimize_to_tray", "INTEGER DEFAULT 0")
    add_column_if_not_exists("accounts", "notify_new_mail", "INTEGER DEFAULT 0")
    add_column_if_not_exists("accounts", "auto_receive_enabled", "INTEGER DEFAULT 0")
    add_column_if_not_exists("accounts", "auto_receive_interval", "INTEGER DEFAULT 10")
    add_column_if_not_exists("accounts", "signature", "TEXT")
    add_column_if_not_exists("accounts", "search_almail_at_startup", "INTEGER DEFAULT 0")
    add_column_if_not_exists("messages", "auth_results", "TEXT")
    add_column_if_not_exists("address_book", "nickname", "TEXT")
    add_column_if_not_exists("messages", "recipient", "TEXT")
    add_column_if_not_exists("messages", "headers", "TEXT")
    add_column_if_not_exists("messages", "account_id", "INTEGER")
    add_column_if_not_exists("accounts", "protocol", "TEXT DEFAULT 'IMAP'")

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS window_settings (
            name TEXT PRIMARY KEY,
            geometry TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS app_config (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    ''')

    # 救済措置: account_id が NULL の古いメッセージを、最初のアカウント(ID=1)に紐付ける
    cursor.execute("UPDATE messages SET account_id = (SELECT id FROM accounts LIMIT 1) WHERE account_id IS NULL")

    conn.commit()
    return conn

def import_address_book(adr_path, conn):
    """AL-Mailのアドレス帳(.adr)をインポート"""
    if not os.path.exists(adr_path): return
    cursor = conn.cursor()
    try:
        with open(adr_path, 'r', encoding='shift-jis', errors='replace') as f:
            for line in f:
                # AL-Mailの形式: ニックネーム=名前 <メールアドレス> または 名前 <メールアドレス>
                line = line.strip()
                if not line or line.startswith(';'): continue
                
                parts = line.split('=', 1)
                if len(parts) == 2:
                    nickname = parts[0].strip()
                    rest = parts[1].strip()
                else:
                    nickname = ""
                    rest = line

                match = re.search(r'([^<]+)?(?:<([^>]+)>|([^=\s]+))', rest)
                if match:
                    name = (match.group(1) or nickname or "Unknown").strip()
                    email_addr = (match.group(2) or match.group(3) or "").strip()
                    if email_addr:
                        # 重複チェック
                        cursor.execute("SELECT id FROM address_book WHERE email = ?", (email_addr,))
                        if not cursor.fetchone():
                            cursor.execute("INSERT INTO address_book (name, email, nickname) VALUES (?, ?, ?)", (name, email_addr, nickname))
        conn.commit()
    except Exception as e:
        print(f"Error importing address book: {e}")

def import_almail_settings(ini_path, conn):
    """AL-Mailの設定ファイル(.ini)からアカウント情報を抽出"""
    if not os.path.exists(ini_path): return None
    
    config = {}
    try:
        with open(ini_path, 'r', encoding='shift-jis', errors='replace') as f:
            content = f.read()
            config['email'] = re.search(r'MailAddress=([^\s\n]+)', content)
            config['smtp'] = re.search(r'SmtpServer=([^\s\n]+)', content)
            config['pop'] = re.search(r'PopServer=([^\s\n]+)', content)
            config['user'] = re.search(r'PopUserName=([^\s\n]+)', content)
            
        if config['email'] and config['pop']:
            cursor = conn.cursor()
            email_val = config['email'].group(1)
            # 既に同じメールアドレスのアカウントがないか確認
            cursor.execute("SELECT id FROM accounts WHERE email = ?", (email_val,))
            if not cursor.fetchone():
                cursor.execute("""
                    INSERT INTO accounts (email, protocol, smtp_server, smtp_port, imap_server, imap_port, username, password)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (email_val, 'POP3',
                      config['smtp'].group(1) if config['smtp'] else "", 465, 
                      config['pop'].group(1) if config['pop'] else "", 995, 
                      config['user'].group(1) if config['user'] else "", ""))
                conn.commit()
                return cursor.lastrowid
    except Exception as e:
        print(f"Error importing settings: {e}")
    return None

def process_and_save_message(msg_bytes, conn, folder_name, account_id, file_path=None):
    """メールのバイナリデータを解析してDBに保存する（重複チェック付き）"""
    cursor = conn.cursor()
    try:
        msg = email.message_from_bytes(msg_bytes, policy=policy.default)
        message_id = msg.get('Message-ID', '')
        
        # 重複チェック (Message-IDがある場合)
        if message_id:
            cursor.execute("SELECT id FROM messages WHERE message_id = ?", (message_id,))
            if cursor.fetchone():
                return # 登録済み

        subject = msg.get('subject', '(No Subject)')
        sender = msg.get('from', '(Unknown)')
        recipient = msg.get('to', '(Unknown)')
        date = msg.get('date', '')

        # 本文（プレーンテキストとHTML）の抽出
        body_plain = ""
        body_html = ""

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    body_plain += decode_payload(part)
                elif content_type == "text/html":
                    body_html += decode_payload(part)
        else:
            if msg.get_content_type() == "text/html":
                body_html = decode_payload(msg)
            else:
                body_plain = decode_payload(msg)

        # 認証ヘッダーの取得（標準および拡張ヘッダーの両方をチェック）
        auth_headers = msg.get_all('Authentication-Results') or []
        auth_headers += msg.get_all('X-Authentication-Results') or []
        auth_results = " ".join(auth_headers)

        # ヘッダー全文の構築
        headers_str = "\n".join([f"{k}: {v}" for k, v in msg.items()])

        cursor.execute(
            "INSERT INTO messages (message_id, subject, sender, recipient, date, body, body_html, file_path, folder, auth_results, headers, account_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (message_id, subject, sender, recipient, date, body_plain, body_html, file_path, folder_name, auth_results, headers_str, account_id)
        )
        msg_db_id = cursor.lastrowid

        # 添付ファイルの抽出
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            
            filename = part.get_filename()
            # 本文以外のパート、またはファイル名を持つパートを添付ファイルとして処理
            is_body = part.get_content_type() in ["text/plain", "text/html"] and part.get('Content-Disposition') is None
            
            if not is_body or filename:
                if filename:
                    decoded = decode_header(filename)
                    filename = "".join([s.decode(c or 'utf-8') if isinstance(s, bytes) else s for s, c in decoded])
                else:
                    filename = f"attachment_{msg_db_id}"

                file_data = part.get_payload(decode=True)
                if file_data:
                    cursor.execute(
                        "INSERT INTO attachments (message_id, filename, data) VALUES (?, ?, ?)",
                        (msg_db_id, filename, sqlite3.Binary(file_data))
                    )

    except Exception as e:
        print(f"Error processing message: {e}")

def import_from_almail(al_folder_path, conn, account_id):
    """AL-Mailのディレクトリから.alファイルを読み取ってDBへ保存"""
    path_obj = Path(al_folder_path)
    if not path_obj.is_dir():
        print(f"Warning: {al_folder_path} is not a directory.")
        return

    folder_name = path_obj.name
    try:
        # ディレクトリ内のファイルを走査
        for file_item in path_obj.iterdir():
            if file_item.is_file() and file_item.suffix.lower() == ".al":
                with file_item.open('rb') as f:
                    process_and_save_message(f.read(), conn, folder_name, account_id, file_path=str(file_item))
    except Exception as e:
        print(f"Failed to expand folder {al_folder_path}. Error: {e}")
        traceback.print_exc()

    conn.commit()
    print(f"Import from {folder_name} completed.")

def fetch_emails(conn, account_id=None):
    """IMAPサーバーから新着メールを取得する。account_idが指定されない場合は全アカウントを対象とする。"""
    cursor = conn.cursor()
    if account_id:
        cursor.execute("SELECT id, imap_server, imap_port, username, password, protocol FROM accounts WHERE id = ?", (account_id,))
    else:
        cursor.execute("SELECT id, imap_server, imap_port, username, password, protocol FROM accounts")
    
    accounts = cursor.fetchall()
    if not accounts:
        return "settings_missing"

    results = []
    for acc_id, server, port, user, pwd, protocol in accounts:
        if not all([server, port, user, pwd]):
            continue

        try:
            if protocol == 'POP3':
                # POP3 受信ロジック
                pop_conn = poplib.POP3_SSL(server, int(port))
                pop_conn.user(user)
                pop_conn.pass_(pwd)
                
                count, _ = pop_conn.stat()
                # 最新の20件を取得
                start = max(1, count - 19)
                for i in range(start, count + 1):
                    _, lines, _ = pop_conn.retr(i)
                    msg_content = b'\n'.join(lines)
                    process_and_save_message(msg_content, conn, "Inbox", acc_id)
                
                pop_conn.quit()
            else:
                # IMAP 受信ロジック
                mail = imaplib.IMAP4_SSL(server, int(port))
                mail.login(user, pwd)
                mail.select("inbox")

                status, data = mail.search(None, 'ALL')
                mail_ids = data[0].split()

                # 各アカウント最新20件を取得
                for m_id in mail_ids[-20:]:
                    status, data = mail.fetch(m_id, '(RFC822)')
                    process_and_save_message(data[0][1], conn, "Inbox", acc_id)

                mail.logout()
            
            results.append("success")
        except Exception as e:
            results.append(f"{user}: {str(e)}")

    conn.commit()
    return "success" if "success" in results else (results[0] if results else "no_accounts")

def send_email(conn, to_addr, subject, body, account_id, attachment_paths=None):
    """指定されたアカウントを使用してメールを送信する"""
    cursor = conn.cursor()
    cursor.execute("SELECT email, smtp_server, smtp_port, username, password FROM accounts WHERE id = ?", (account_id,))
    account = cursor.fetchone()
    
    if not account:
        return "settings_missing"

    from_addr, server, port, user, pwd = account
    if not all([from_addr, server, port, user, pwd]):
        return "settings_missing"

    try:
        # メッセージの作成
        if attachment_paths:
            msg = MIMEMultipart()
            msg.attach(MIMEText(body, 'plain'))
            for path in attachment_paths:
                if not os.path.exists(path): continue
                part = MIMEBase('application', "octet-stream")
                with open(path, 'rb') as f:
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                fname = os.path.basename(path)
                part.add_header('Content-Disposition', 'attachment', filename=Header(fname, 'utf-8').encode())
                msg.attach(part)
        else:
            msg = MIMEText(body)

        msg['Subject'] = subject
        msg['From'] = from_addr
        msg['To'] = to_addr
        msg['Date'] = formatdate(localtime=True)

        # SMTPサーバーに接続して送信 (SSLを想定)
        with smtplib.SMTP_SSL(server, int(port)) as smtp:
            smtp.login(user, pwd)
            smtp.send_message(msg)
        return "success"
    except Exception as e:
        return str(e)

def test_connection(imap_server, imap_port, smtp_server, smtp_port, username, password):
    """IMAPとSMTPの両方の接続テストを行う"""
    results = []
    
    # IMAP Test
    try:
        mail = imaplib.IMAP4_SSL(imap_server, int(imap_port))
        mail.login(username, password)
        mail.logout()
        results.append("✅ IMAP接続成功")
    except Exception as e:
        results.append(f"❌ IMAP接続失敗: {str(e)}")
        
    # SMTP Test
    try:
        with smtplib.SMTP_SSL(smtp_server, int(smtp_port)) as smtp:
            smtp.login(username, password)
        results.append("✅ SMTP接続成功")
    except Exception as e:
        results.append(f"❌ SMTP接続失敗: {str(e)}")
        
    return "\n".join(results)

def decode_payload(part):
    """ペイロードを適切なエンコーディングでデコードする"""
    payload = part.get_payload(decode=True)
    if not payload: return ""
    
    # 試行するエンコーディングのリスト
    charsets = [part.get_content_charset(), 'iso-2022-jp', 'cp932', 'utf-8']
    for charset in charsets:
        if not charset: continue
        try:
            return payload.decode(charset)
        except (UnicodeDecodeError, LookupError):
            continue
    return payload.decode('utf-8', errors='replace')

if __name__ == "__main__":
    # 使用例: conn = setup_database(); import_from_almail("C:/ALMail/Mail/Inbox", conn)
    pass