import sys
import io
import os
import urllib.request
import subprocess

# --noconsole (GUI) モード実行時に sys.stdout が None になり、
# tkinterweb 等がエラー出力を試みてクラッシュするのを防ぐ
if getattr(sys, 'frozen', False) and sys.stdout is None:
    # 実行ファイルと同じディレクトリにログを出力する
    exe_dir = os.path.dirname(sys.executable)
    log_path = os.path.join(exe_dir, "error_log.txt")
    try:
        sys.stdout = open(log_path, "w", encoding="utf-8")
        sys.stderr = sys.stdout
    except:
        pass

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sqlite3
import almail_importer as importer
import json
import re
import html
import urllib.parse
import webbrowser
import tempfile
import threading
import email.utils
import datetime
from tkinterweb import HtmlFrame
import tkinterweb_tkhtml
try:
    from PIL import Image, ImageDraw, ImageTk
except ImportError:
    Image = None
    ImageDraw = None
    ImageTk = None

try:
    import pystray
except ImportError:
    pystray = None

try:
    from win10toast import ToastNotifier
    toaster = ToastNotifier()
except ImportError:
    toaster = None

class ModernALMail:
    # 正規表現をクラス定数として事前コンパイル（軽量化）
    RE_HTML_TAG = re.compile(r'<[^>]*?>', re.DOTALL)
    RE_HTML_META = re.compile(r'<(script|style)[^>]*?>.*?</\1>', re.IGNORECASE | re.DOTALL)
    RE_HTML_NL = re.compile(r'<(br|p|div|li|tr|h[1-6])[^>]*>', re.IGNORECASE)
    RE_DATE_TIME = re.compile(r'(\d{1,2})月\s?(\d{1,2})日.*?\s?(\d{1,2}):(\d{2})')
    RE_MEET_URL = re.compile(r'https://meet\.google\.com/[\w-]+')
    RE_DOMAIN = re.compile(r'@([\w.-]+)')
    RE_AUTH = {"spf": re.compile(r'spf=([a-z]+)', re.I), "dkim": re.compile(r'dkim=([a-z]+)', re.I), "dmarc": re.compile(r'dmarc=([a-z]+)', re.I), "header_from": re.compile(r'header\.from=([\w.-]+)', re.I)}

    APP_VERSION = "1.1.0" # 現在のプログラムのバージョン
    UPDATE_URL = "https://raw.githubusercontent.com/t-muguruma/almail/main/version.json"
    DOWNLOAD_BASE_URL = "https://github.com/t-muguruma/almail/raw/main/"

    def __init__(self, root):
        self.root = root

        self.compose_win = None  # 作成画面の管理用
        self.auto_receive_job = None # 自動受信タイマー用
        self.preview_images = [] # 画像プレビューの参照保持用
        self.tray_icon = None # システムトレイアイコン用
        self.current_message_account_id = None # 現在選択中のメールが属するアカウントID
        self.tk_icon = None # ウィンドウアイコン用
        self.pil_icon = None # トレイアイコン用
        self._skip_auto_select_first = False # 受信時の自動フォーカス制御用フラグ

        # 1. 起動画面（スプラッシュスクリーン）を表示
        self.root.withdraw()  # メインウィンドウを一旦隠す
        splash = tk.Toplevel(self.root)
        splash.overrideredirect(True)  # 枠やタイトルバーを消す

        # ロゴ画面のデザイン（白背景にブルーのアクセント）
        w, h = 450, 250
        sw, sh = splash.winfo_screenwidth(), splash.winfo_screenheight()
        splash.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        splash.configure(bg='white', highlightbackground='#0052a3', highlightthickness=2)

        # タイトルロゴ風の配置
        logo_frame = tk.Frame(splash, bg='white')
        logo_frame.pack(expand=True)
        
        # ロゴ画面のデザイン（グラデーション背景とピンクのボーダー）
        splash.configure(highlightbackground='#ff80ab', highlightthickness=2)

        self.splash_canvas = tk.Canvas(splash, width=w, height=h, highlightthickness=0)
        self.splash_canvas.pack(fill="both", expand=True)

        # PILを使用して「Petal」らしいグラデーション画像を生成 (薄いピンク #fff0f5 から 白 #ffffff へ)
        if Image:
            grad = Image.new('RGB', (w, h), "#ffffff")
            draw = ImageDraw.Draw(grad)
            for i in range(h):
                # Y座標に応じて色を補間
                g = int(240 + (15 * i / h))
                b = int(245 + (10 * i / h))
                draw.line([(0, i), (w, i)], fill=(255, g, b))
            self.splash_bg_photo = ImageTk.PhotoImage(grad)
            self.splash_canvas.create_image(0, 0, anchor="nw", image=self.splash_bg_photo)

        # テキストの描画（Canvas上に描画することで背景を透過させる）
        self.splash_canvas.create_text(w//2, h//2 - 35, text="Simple & Elegant", 
                                       font=("Arial", 20, "italic"), fill='#666666')
        self.splash_canvas.create_text(w//2, h//2 + 15, text="Petal", 
                                       font=("Arial", 42, "bold"), fill='#d81b60')

        self.splash_status = tk.Label(splash, text="Initializing database...", font=("MS UI Gothic", 9), 
                                     bg='white', fg='#999999')
        # ステータスラベルをCanvasの下部に窓として配置
        self.splash_canvas.create_window(w//2, h - 30, window=self.splash_status)
        splash.update()  # 即座に描画を反映

        # データベースのセットアップ
        self.splash_status.config(text="Loading database...")
        splash.update()
        if getattr(sys, 'frozen', False):
            # インストール環境（Program Files）は書き込み禁止のため、AppData を使用する
            appdata_dir = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), "Petal")
            if not os.path.exists(appdata_dir):
                os.makedirs(appdata_dir)
            self.db_path = os.path.join(appdata_dir, "mailbox.db")

            # 移行措置：もし実行ファイルと同じ場所に既存のDBがあれば移動を試みる
            old_db = os.path.join(os.path.dirname(sys.executable), "mailbox.db")
            if os.path.exists(old_db) and not os.path.exists(self.db_path):
                import shutil
                try: shutil.move(old_db, self.db_path)
                except: pass
        else:
            # 通常のスクリプトとして実行されている場合
            base_dir = os.path.dirname(os.path.abspath(__file__))
            self.db_path = os.path.join(base_dir, "mailbox.db")

        print(f"DEBUG: Resolved DB Path: {self.db_path}")
        print(f"DEBUG: DB file exists at path: {os.path.exists(self.db_path)}")

        self.conn = importer.setup_database(self.db_path)

        # DBに現在のバージョンを記録
        self.conn.execute("INSERT OR REPLACE INTO app_config (key, value) VALUES ('version', ?)", (self.APP_VERSION,))
        self.conn.commit()

        # メインウィンドウの初期設定
        self.root.title("Petal")
        
        # アプリケーションアイコンの設定
        self._setup_app_icon()

        # ウィンドウ位置の記憶と復元 (DB接続が完了してから実行)
        saved_geometry = self._load_window_geometry("MainWindow")
        if saved_geometry:
            self.root.geometry(saved_geometry)
        else:
            self.root.geometry("1100x750") # デフォルトサイズ

        # メインウィンドウが閉じられたときの処理
        self.root.protocol("WM_DELETE_WINDOW", self.on_main_window_close)
        # 最小化イベントのバインド
        self.root.bind("<Unmap>", self.on_window_minimize)

        # 共通の右クリックメニュー作成
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="切り取り", command=lambda: self.root.focus_get().event_generate("<<Cut>>"))
        self.context_menu.add_command(label="コピー", command=lambda: self.root.focus_get().event_generate("<<Copy>>"))
        self.context_menu.add_command(label="貼り付け", command=lambda: self.root.focus_get().event_generate("<<Paste>>"))
        self.context_menu.add_separator()
        self.context_menu.add_command(label="すべて選択", command=lambda: self.root.focus_get().event_generate("<<SelectAll>>"))

        self.splash_status.config(text="Setting up interface...")
        splash.update()
        self.setup_ui()
        self.refresh_folders()
        self.select_home()
        # 起動時の自動探索（初回またはチェックが入っている時のみ）
        self.check_almail_on_startup()
        self.init_auto_receive()

        # 2. 準備ができたらスプラッシュを消してメインを表示
        splash.destroy()
        self.root.deiconify()

    # --- ウィンドウ管理共通メソッド ---
    def _load_window_geometry(self, name):
        res = self.conn.execute("SELECT geometry FROM window_settings WHERE name = ?", (name,)).fetchone()
        return res[0] if res else None

    def _save_window_geometry(self, name, geom):
        self.conn.execute("INSERT OR REPLACE INTO window_settings (name, geometry) VALUES (?, ?)", (name, geom))
        self.conn.commit()

    def _center_window_on_parent(self, child, parent, w, h):
        parent.update_idletasks()
        px, py = parent.winfo_x(), parent.winfo_y()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        return f"{w}x{h}+{px + (pw//2)-(w//2)}+{py + (ph//2)-(h//2)}"

    def _setup_sub_window(self, name, title, w, h, modal=False):
        win = tk.Toplevel(self.root)
        win.title(title)
        geom = self._load_window_geometry(name) or self._center_window_on_parent(win, self.root, w, h)
        win.geometry(geom)
        if modal: win.grab_set()
        win.protocol("WM_DELETE_WINDOW", lambda: self._on_sub_window_close(win, name))
        return win

    def _on_sub_window_close(self, win, name):
        self._save_window_geometry(name, win.geometry())
        win.destroy()

    def on_main_window_close(self):
        try:
            self._save_window_geometry("MainWindow", self.root.geometry())
        finally:
            self.root.quit()

    def _setup_app_icon(self):
        """プログラム内で生成したアイコンをウィンドウとタスクバーに適用"""
        # pystrayがなくてもPILがあればアイコン設定は可能
        if not ImageTk: return
        
        def apply_rounded_corners(img, rad):
            """画像を角丸に加工するヘルパー"""
            # 1. ロゴと同じサイズの白い角丸背景を作成
            bg = Image.new('RGBA', img.size, (255, 255, 255, 0))
            mask = Image.new('L', img.size, 0)
            mask_draw = ImageDraw.Draw(mask)
            mask_draw.rounded_rectangle((0, 0, img.size[0]-1, img.size[1]-1), radius=rad, fill=255)
            
            # 背景を白（またはPetalの薄いピンク）で塗りつぶす
            panel = Image.new('RGBA', img.size, (255, 255, 255, 255))
            bg.paste(panel, (0, 0), mask=mask)
            
            # 2. 元のロゴを少しだけ小さくリサイズして中央に配置（余白を作る）
            padding = int(img.size[0] * 0.1)
            inner_size = img.size[0] - (padding * 2)
            img_small = img.resize((inner_size, inner_size), Image.Resampling.LANCZOS)
            
            bg.paste(img_small, (padding, padding), mask=img_small if img_small.mode == 'RGBA' else None)
            return bg

        def get_resource_path(filename):
            if getattr(sys, 'frozen', False):
                # 1. 内部リソース 2. 実行ファイルと同階層
                paths = [os.path.join(sys._MEIPASS, filename), os.path.join(os.path.dirname(sys.executable), filename)]
            else:
                paths = [os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)]
            for p in paths:
                if os.path.exists(p): return p
            return None

        # 1. メインウィンドウ用アイコン (Petal_icon.png)
        main_icon_path = get_resource_path("Petal_icon.png")
        if main_icon_path:
            try:
                main_pil = Image.open(main_icon_path).convert("RGBA").resize((64, 64), Image.Resampling.LANCZOS)
                # メインアイコンに角丸を適用（半径14px）
                main_pil = apply_rounded_corners(main_pil, 14)
                self.tk_icon = ImageTk.PhotoImage(main_pil)
                self.root.iconphoto(True, self.tk_icon)
            except Exception as e:
                print(f"Failed to load main icon: {e}")

        # 2. タスクアイコン用 (Petal_small.png)
        small_icon_path = get_resource_path("Petal_small.png")
        if small_icon_path:
            try:
                self.pil_icon = Image.open(small_icon_path).convert("RGBA").resize((32, 32), Image.Resampling.LANCZOS)
                # 小さいアイコンにも角丸を適用（半径7px）
                self.pil_icon = apply_rounded_corners(self.pil_icon, 7)
            except Exception as e:
                print(f"Failed to load small icon: {e}")

        # フォールバック（画像がない場合）
        if not hasattr(self, 'tk_icon') or not self.tk_icon:
            dummy = Image.new('RGBA', (64, 64), color=(0, 82, 163, 255))
            draw = ImageDraw.Draw(dummy)
            draw.rectangle([10, 20, 54, 44], outline="white", width=4)
            draw.line([10, 20, 32, 32, 54, 20], fill="white", width=3)
            dummy = apply_rounded_corners(dummy, 14)
            self.tk_icon = ImageTk.PhotoImage(dummy)
            self.root.iconphoto(True, self.tk_icon)
        
        if not self.pil_icon:
            # トレイアイコンが読み込めなかった場合はメイン用をリサイズして代用
            if hasattr(self, 'tk_icon'):
                self.pil_icon = Image.new('RGBA', (32, 32), color=(0, 82, 163, 255))

    def on_window_minimize(self, event):
        """最小化されたときに設定に応じてトレイへ格納"""
        if event.widget == self.root and self.root.state() == 'iconic':
            cursor = self.conn.cursor()
            cursor.execute("SELECT minimize_to_tray FROM accounts LIMIT 1")
            row = cursor.fetchone()
            if row and row[0] and pystray:
                self.hide_to_tray()

    def hide_to_tray(self):
        """ウィンドウを隠してトレイアイコンを表示"""
        self.root.withdraw()
        
        if not self.tray_icon:
            menu = pystray.Menu(
                pystray.MenuItem("開く", self.show_from_tray, default=True),
                pystray.MenuItem("終了", self.on_main_window_close)
            )
            self.tray_icon = pystray.Icon("Petal", self.pil_icon, "Petal", menu)
        
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def show_from_tray(self, icon=None):
        """トレイから復帰"""
        if self.tray_icon:
            self.tray_icon.stop()
            self.tray_icon = None
        self.root.after(0, self.root.deiconify)
        self.root.after(0, lambda: self.root.state('normal'))

    def notify_new_mail(self, subject, sender):
        """Windows通知を表示"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT notify_new_mail FROM accounts LIMIT 1")
        row = cursor.fetchone()
        if row and row[0] and toaster:
            toaster.show_toast("新着メール - Petal", 
                               f"件名: {subject}\n差出人: {sender}",
                               duration=5, threaded=True)

    # --- UI構築 ---
    def setup_ui(self):
        # メニューバー
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="新規作成", command=self.open_compose_window)
        file_menu.add_command(label="メール受信", command=self.receive_mail)
        file_menu.add_command(label="アドレス帳...", command=self.open_address_book)
        file_menu.add_command(label="AL-Mailフォルダをインポート...", command=self.import_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self.root.quit)
        menubar.add_cascade(label="ファイル", menu=file_menu)

        # ツールメニューの追加
        tools_menu = tk.Menu(menubar, tearoff=0)
        export_menu = tk.Menu(tools_menu, tearoff=0)
        export_menu.add_command(label="アカウント情報...", command=lambda: self.run_export(['accounts']))
        export_menu.add_command(label="アドレス帳...", command=lambda: self.run_export(['address_book']))
        export_menu.add_command(label="メールデータ...", command=lambda: self.run_export(['messages']))
        export_menu.add_command(label="アカウント ＋ アドレス帳...", command=lambda: self.run_export(['accounts', 'address_book']))
        export_menu.add_command(label="全データ (アカウント ＋ アドレス帳 ＋ メール)...", command=lambda: self.run_export(['accounts', 'address_book', 'messages']))
        tools_menu.add_cascade(label="エクスポート", menu=export_menu)

        import_menu = tk.Menu(tools_menu, tearoff=0)
        import_menu.add_command(label="アカウント情報...", command=lambda: self.run_import(['accounts']))
        import_menu.add_command(label="アドレス帳...", command=lambda: self.run_import(['address_book']))
        import_menu.add_command(label="メールデータ...", command=lambda: self.run_import(['messages']))
        import_menu.add_command(label="アカウント ＋ アドレス帳...", command=lambda: self.run_import(['accounts', 'address_book']))
        import_menu.add_command(label="全データ (アカウント ＋ アドレス帳 ＋ メール)...", command=lambda: self.run_import(['accounts', 'address_book', 'messages']))
        tools_menu.add_cascade(label="インポート", menu=import_menu)
        tools_menu.add_separator()
        tools_menu.add_command(label="カレンダーに登録 (選択中メール)", command=self.schedule_to_google_calendar)
        menubar.add_cascade(label="ツール", menu=tools_menu)
        
        # 添付ファイルメニュー
        files_menu = tk.Menu(menubar, tearoff=0)
        files_menu.add_command(label="添付ファイル一覧...", command=self.open_attachments_manager)
        menubar.add_cascade(label="添付ファイル", menu=files_menu)

        menubar.add_command(label="設定", command=self.open_settings)

        # ヘルプメニューの追加
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="マニュアル", command=self.open_manual)
        help_menu.add_command(label="バージョンアップ確認", command=self.check_for_updates)
        help_menu.add_command(label="About", command=self.open_about)
        menubar.add_cascade(label="ヘルプ", menu=help_menu)

        self.root.config(menu=menubar)

        # ツールバー（アイコンバー）
        self.toolbar = tk.Frame(self.root, relief=tk.RAISED, borderwidth=1, bg="#f8f9fa")
        self.toolbar.pack(side=tk.TOP, fill=tk.X)

        btn_receive = tk.Button(self.toolbar, text="📥 受信", bg="#e3f2fd", command=self.receive_mail, relief=tk.FLAT, padx=10)
        btn_receive.pack(side=tk.LEFT, padx=2, pady=2)
        btn_compose = tk.Button(self.toolbar, text="📝 新規", bg="#e8f5e9", command=self.open_compose_window, relief=tk.FLAT, padx=10)
        btn_compose.pack(side=tk.LEFT, padx=2, pady=2)
        btn_addr = tk.Button(self.toolbar, text="📖 アドレス帳", bg="#fff3e0", command=self.open_address_book, relief=tk.FLAT, padx=10)
        btn_addr.pack(side=tk.LEFT, padx=2, pady=2)
        btn_delete = tk.Button(self.toolbar, text="🗑️ 削除", bg="#ffebee", command=self.delete_selected_mail, relief=tk.FLAT, padx=10)
        btn_delete.pack(side=tk.LEFT, padx=2, pady=2)
        btn_settings = tk.Button(self.toolbar, text="⚙ 設定", bg="#eeeeee", command=self.open_settings, relief=tk.FLAT, padx=10)
        btn_settings.pack(side=tk.LEFT, padx=2, pady=2)

        # メインパン（左右に分割）
        self.h_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.h_pane.pack(fill=tk.BOTH, expand=True)

        # 左：フォルダツリー (先に行う)
        self.folder_tree = ttk.Treeview(self.h_pane, show="tree", selectmode="browse")
        self.folder_tree.bind("<<TreeviewSelect>>", self.on_folder_select)

        # 右側のメインコンテナ（ホームとメール一覧を切り替えるため）
        self.right_container = ttk.Frame(self.h_pane)

        # --- ホームビュー ---
        self.home_frame = ttk.Frame(self.right_container)
        self.setup_home_view()
        self.home_frame.pack(fill=tk.BOTH, expand=True) # 初期状態として配置

        # --- メールビュー（上下分割） ---
        self.v_pane = ttk.PanedWindow(self.right_container, orient=tk.VERTICAL)

        # 右上：メール一覧
        columns = ("subject", "sender", "date", "account")
        self.msg_list = ttk.Treeview(self.v_pane, columns=columns, show="headings", selectmode="browse")
        self.msg_list.heading("subject", text="件名")
        self.msg_list.heading("sender", text="差出人")
        self.msg_list.heading("date", text="送信日時")
        self.msg_list.heading("account", text="アカウント")
        self.msg_list.column("subject", width=400)
        self.msg_list.column("sender", width=200)
        self.msg_list.column("date", width=150)
        self.msg_list.column("account", width=120)
        self.v_pane.add(self.msg_list, weight=2)
        self.msg_list.bind("<<TreeviewSelect>>", self.on_message_select)
        self.msg_list.bind("<Button-3>", self.show_msg_list_context_menu)

        # パンに追加する順番を確定
        self.h_pane.add(self.folder_tree, weight=1)
        self.h_pane.add(self.right_container, weight=4)

        self.content_container = ttk.Frame(self.v_pane)
        self.v_pane.add(self.content_container, weight=3)

        # メッセージ操作ツールバー（右下ペイン最上段）
        self.action_bar = tk.Frame(self.content_container, bg="#ffffff", borderwidth=1, relief=tk.FLAT)
        self.action_bar.pack(side=tk.TOP, fill=tk.X)
        self.btn_reply = tk.Button(self.action_bar, text="↩️ 返信", bg="#ffffff", command=self.reply_mail, relief=tk.FLAT, padx=10, font=("MS UI Gothic", 9))
        self.btn_reply.pack(side=tk.LEFT, padx=2, pady=2)
        self.btn_forward = tk.Button(self.action_bar, text="↪️ 転送", bg="#ffffff", command=self.forward_mail, relief=tk.FLAT, padx=10, font=("MS UI Gothic", 9))
        self.btn_forward.pack(side=tk.LEFT, padx=2, pady=2)
        self.btn_pane_delete = tk.Button(self.action_bar, text="🗑️ 削除", bg="#ffffff", fg="#c62828", command=self.delete_selected_mail, relief=tk.FLAT, padx=10, font=("MS UI Gothic", 9))
        self.btn_pane_delete.pack(side=tk.LEFT, padx=2, pady=2)
        self.btn_headers = tk.Button(self.action_bar, text="📄 ヘッダー表示", bg="#ffffff", command=self.view_headers, relief=tk.FLAT, padx=10, font=("MS UI Gothic", 9))
        self.btn_headers.pack(side=tk.LEFT, padx=2, pady=2)
        self.btn_browser = tk.Button(self.action_bar, text="🌐 ブラウザで表示", bg="#ffffff", command=self.view_in_browser, relief=tk.FLAT, padx=10, font=("MS UI Gothic", 9))
        self.btn_browser.pack(side=tk.LEFT, padx=2, pady=2)

        # 宛先・差出人表示バー（新設）
        self.info_bar = tk.Frame(self.content_container, bg="#fcfcfc", borderwidth=1, relief=tk.FLAT)
        self.info_bar.pack(side=tk.TOP, fill=tk.X)
        self.lbl_from_to = tk.Label(self.info_bar, text="", anchor="w", font=("MS UI Gothic", 9), bg="#fcfcfc", fg="#555555", padx=5, pady=2)
        self.lbl_from_to.pack(side=tk.LEFT, fill=tk.X)

        # 認証・ドメイン照合バナー（右下ペインの最上段）
        self.auth_banner = tk.Label(self.content_container, text="", font=("MS UI Gothic", 9, "bold"), pady=2)
        self.auth_banner.pack(side=tk.TOP, fill=tk.X)
        
        # 添付ファイル表示エリア（スクロールバー付きText）
        self.att_frame = tk.Frame(self.content_container)
        self.attachment_bar = tk.Text(self.att_frame, bg="#f0f0f0", height=1, padx=5, pady=2, 
                                     relief=tk.FLAT, font=("MS UI Gothic", 9), state=tk.DISABLED, cursor="arrow", wrap=tk.CHAR)
        self.att_scroll = ttk.Scrollbar(self.att_frame, orient=tk.VERTICAL, command=self.attachment_bar.yview)
        self.attachment_bar.config(yscrollcommand=self.att_scroll.set)
        self.attachment_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.att_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.att_frame.pack(side=tk.TOP, fill=tk.X)
        self.attachment_bar.pack_forget() # 初期は非表示
        self.att_frame.pack_forget()

        # メール本文と画像プレビューを分割するペイン
        self.body_preview_pane = ttk.PanedWindow(self.content_container, orient=tk.VERTICAL)
        self.body_preview_pane.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.body_container = ttk.Frame(self.body_preview_pane)
        self.preview_container = ttk.Frame(self.body_preview_pane)
        self.body_preview_pane.add(self.body_container, weight=4)
        # プレビューコンテナは画像がある時だけ add するため初期は追加しない

        # 画像プレビューエリア（HTML表示時用：スクロール可能に）
        self.preview_canvas = tk.Canvas(self.preview_container, bg="white", highlightthickness=0)
        self.preview_frame = tk.Frame(self.preview_canvas, bg="white")
        self.preview_scroll = ttk.Scrollbar(self.preview_container, orient=tk.VERTICAL, command=self.preview_canvas.yview)
        self.preview_canvas.configure(yscrollcommand=self.preview_scroll.set)
        self.preview_canvas.create_window((0, 0), window=self.preview_frame, anchor="nw")
        self.preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # テキスト表示用コンテナ（スクロールバー付き）
        self.body_frame = tk.Frame(self.body_container)
        self.body_text = tk.Text(self.body_frame, font=("MS Gothic", 11), padx=10, pady=10, state=tk.DISABLED)
        self.body_scroll = ttk.Scrollbar(self.body_frame, orient=tk.VERTICAL, command=self.body_text.yview)
        self.body_text.configure(yscrollcommand=self.body_scroll.set)
        
        self.body_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.body_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # HTML表示用 (埋め込み)
        self.html_view = HtmlFrame(self.body_container)

        # 初期状態はテキストを表示
        self.body_frame.pack(fill=tk.BOTH, expand=True)
        self.bind_context_menu(self.body_text)
        
        # HtmlFrameの内部ウィジェットにも右クリックをバインド
        # tkinterwebのHtmlFrameは .html (または内部の実体) にバインドが必要
        try:
            self.bind_context_menu(self.html_view)
        except:
            pass

        # キャンバスのスクロール範囲更新用
        def _on_preview_resize(e):
            self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
        self.preview_frame.bind("<Configure>", _on_preview_resize)

    def check_message_security(self, msg_id):
        """SPF/DKIM/DMARCの判定とバナー更新"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT sender, auth_results FROM messages WHERE id = ?", (msg_id,))
        res = cursor.fetchone()
        if not res: return
        sender, auth_res = res

        if not auth_res or not auth_res.strip():
            self.auth_banner.config(text="🛡️ 認証情報なし (古いメール、またはサーバー未対応)", bg="#9e9e9e", fg="white")
            return

        # ステータス解析
        s = self.RE_AUTH["spf"].search(auth_res)
        dk = self.RE_AUTH["dkim"].search(auth_res)
        dm = self.RE_AUTH["dmarc"].search(auth_res)
        h_from = self.RE_AUTH["header_from"].search(auth_res)

        s_val, dk_val, dm_val = [x.group(1).upper() if x else "UNKNOWN" for x in (s, dk, dm)]
        from_domain = self.RE_DOMAIN.search(sender).group(1).lower() if self.RE_DOMAIN.search(sender) else ""
        auth_domain = h_from.group(1).lower() if h_from else ""

        is_suspicious = (dm_val == "FAIL" or (auth_domain and from_domain and from_domain != auth_domain))
        status_text = f"SPF:{s_val} | DKIM:{dk_val} | DMARC:{dm_val}"

        if is_suspicious:
            self.auth_banner.config(text=f"❌ 警告: なりすましの疑い | {status_text}", bg="#f44336", fg="white")
        elif s_val == "PASS" or dk_val == "PASS":
            self.auth_banner.config(text=f"✅ 認証成功: {status_text}", bg="#4caf50", fg="white")
        else:
            self.auth_banner.config(text=f"⚠️ 認証不完全: {status_text}", bg="#ffc107", fg="#333333")

    def schedule_to_google_calendar(self):
        """選択中のメールからGoogleカレンダーの登録画面を開く"""
        selected = self.msg_list.selection()
        if not selected:
            messagebox.showinfo("情報", "カレンダーに登録するメールを選択してください。")
            return

        msg_id = selected[0]
        cursor = self.conn.cursor()
        cursor.execute("SELECT subject, body, date FROM messages WHERE id = ?", (msg_id,))
        subject, body, msg_date = cursor.fetchone()

        # 日時の抽出試行 (例: 4月21日 15:00)
        # 年はメールの送信年から推測、取得できない場合は実行時の年を使用
        year = msg_date.split()[3] if len(msg_date.split()) > 3 else str(datetime.datetime.now().year)
        date_match = self.RE_DATE_TIME.search(body)
        
        start_dt = ""
        if date_match:
            m, d, h, minute = [x.zfill(2) for x in date_match.groups()]
            # Google Calendarのdatesパラメータ形式: YYYYMMDDTHHmmSSZ (UTC)
            # 簡易的にJSTとして扱うためZは付けず、開始・終了1時間をセット
            start_dt = f"{year}{m}{d}T{h}{minute}00"
            end_h = str(int(h) + 1).zfill(2)
            end_dt = f"{year}{m}{d}T{end_h}{minute}00"
            dates = f"{start_dt}/{end_dt}"
        else:
            dates = ""

        # Google Meet URLの抽出
        meet_match = self.RE_MEET_URL.search(body)
        location = meet_match.group(0) if meet_match else ""

        # URLパラメータの構築
        params = {
            "action": "TEMPLATE",
            "text": subject,
            "details": body[:1000] + "...", # 本文を説明欄に
            "location": location,
        }
        if dates:
            params["dates"] = dates

        url = "https://www.google.com/calendar/render?" + urllib.parse.urlencode(params)
        webbrowser.open(url)

    def setup_home_view(self):
        """ホーム画面の構築"""
        self.home_inner = tk.Frame(self.home_frame, bg='white')
        self.home_inner.pack(fill=tk.BOTH, expand=True)
        header = tk.Label(self.home_inner, text="🏠 Petal Home", font=("Arial", 24, "bold"), 
                          bg='white', fg='#0052a3', pady=30)
        header.pack()

        stats_frame = ttk.LabelFrame(self.home_inner, text="ステータス")
        stats_frame.pack(fill=tk.X, padx=50, pady=10)
        
        self.lbl_stats = tk.Label(stats_frame, text="読込中...", font=("MS UI Gothic", 12), pady=10)
        self.lbl_stats.pack()

        recent_frame = ttk.LabelFrame(self.home_inner, text="最近の受信メール")
        recent_frame.pack(fill=tk.BOTH, expand=True, padx=50, pady=10)

        self.recent_container = tk.Frame(recent_frame, bg='white')
        self.recent_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        btn_refresh = ttk.Button(self.home_inner, text="今すぐ受信チェック", command=self.receive_mail)
        btn_refresh.pack(pady=20)

    def update_home_view(self):
        """ホーム画面の情報を最新に更新"""
        cursor = self.conn.cursor()
        
        # 統計情報
        cursor.execute("SELECT COUNT(*) FROM messages")
        total = cursor.fetchone()[0]
        cursor.execute("SELECT email FROM accounts LIMIT 1")
        acc = cursor.fetchone()
        acc_name = acc[0] if acc else "未設定"
        self.lbl_stats.config(text=f"アカウント: {acc_name}  /  総メッセージ数: {total} 件")

        # 最近のメール3件 (クリック可能なリンクとして構築)
        for widget in self.recent_container.winfo_children():
            widget.destroy()

        # アカウントIDも取得するように変更
        cursor.execute("SELECT id, subject, sender, date, folder, account_id FROM messages ORDER BY id DESC LIMIT 3")
        recent = cursor.fetchall()
        if recent:
            for r_id, subj, snd, dt, fld, acc_id in recent:
                link_text = f"【{dt}】 {subj} - {snd}"
                # フォルダの内部ID (FLD_{acc_id}_{folder}) を生成してジャンプ先に指定
                target_iid = f"FLD_{acc_id}_{fld}" if acc_id and fld else "UNIFIED_INBOX"
                lbl = tk.Label(self.recent_container, text=link_text, fg="#0052a3", cursor="hand2", 
                               bg="white", anchor="w", font=("MS UI Gothic", 10, "underline"))
                lbl.pack(fill=tk.X, pady=2)
                lbl.bind("<Button-1>", lambda e, mid=r_id, tid=target_iid: self.jump_to_message(mid, tid))
        else:
            tk.Label(self.recent_container, text="メッセージはありません。", bg="white").pack()

    def jump_to_message(self, msg_id, target_iid):
        """特定のフォルダに切り替え、指定したメッセージを選択する"""
        self.folder_tree.selection_set(target_iid)
        self.folder_tree.see(target_iid)
        # フォルダ選択イベントによりmsg_listが更新されるのを待ってから選択
        self.root.after(10, lambda: self._select_specific_message(msg_id))

    def _select_specific_message(self, msg_id):
        if self.msg_list.exists(msg_id):
            self.msg_list.selection_set(msg_id)
            self.msg_list.see(msg_id)
            self.msg_list.focus(msg_id)

    def select_home(self):
        """起動時にホームを選択する"""
        self.folder_tree.selection_set("HOME_NODE")

    def bind_context_menu(self, widget):
        """ウィジェットに右クリックメニューをバインドする"""
        widget.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        """右クリックメニューを表示"""
        try:
            # フォーカスを右クリックされたウィジェットに移す
            event.widget.focus_set()
            self.context_menu.post(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def strip_html_tags(self, html_content):
        """HTMLタグを高度に除外してテキスト化"""
        if not html_content: return ""
        
        # 巨大すぎるHTMLの正規表現解析によるフリーズを防止 (1MB制限)
        if len(html_content) > 1000000:
            html_content = html_content[:1000000] + "\n...(truncated)"
            
        t = self.RE_HTML_META.sub('', html_content)
        t = self.RE_HTML_NL.sub('\n', t)
        t = self.RE_HTML_TAG.sub('', t)
        return html.unescape(t).strip()

    def reply_mail(self):
        """選択中のメールに返信する"""
        selected = self.msg_list.selection()
        if not selected: return
        
        cursor = self.conn.cursor()
        cursor.execute("SELECT subject, sender, date, body FROM messages WHERE id = ?", (selected[0],))
        subj, sender, date, body = cursor.fetchone()
        
        re_subject = f"Re: {subj}" if not subj.lower().startswith("re:") else subj
        re_body = f"\n\n--- Original Message ---\nFrom: {sender}\nDate: {date}\nSubject: {subj}\n\n{body or ''}"
        
        self.open_compose_window(to=sender, subject=re_subject, body=re_body, from_account_id=self.current_message_account_id)

    def forward_mail(self):
        """選択中のメールを転送する"""
        selected = self.msg_list.selection()
        if not selected: return
        
        cursor = self.conn.cursor()
        cursor.execute("SELECT subject, sender, date, body FROM messages WHERE id = ?", (selected[0],))
        subj, sender, date, body = cursor.fetchone()
        
        fw_subject = f"Fw: {subj}" if not subj.lower().startswith("fw:") else subj
        fw_body = f"\n\n--- Forwarded Message ---\nFrom: {sender}\nDate: {date}\nSubject: {subj}\n\n{body or ''}"
        
        self.open_compose_window(to="", subject=fw_subject, body=fw_body, from_account_id=self.current_message_account_id)

    def delete_selected_mail(self):
        """選択中のメールをゴミ箱に移動、または削除"""
        selected = self.msg_list.selection()
        if not selected: return
        
        msg_id = selected[0]
        if messagebox.askyesno("確認", "選択したメールを削除（ゴミ箱へ移動）しますか？"):
            # folder を 'Trash' に更新する（簡易削除）
            self.conn.execute("UPDATE messages SET folder = 'Trash' WHERE id = ?", (msg_id,))
            self.conn.commit()
            self.on_folder_select(None) # 一覧を更新

    def show_msg_list_context_menu(self, event):
        """メール一覧用の右クリックメニュー"""
        item = self.msg_list.identify_row(event.y)
        if item:
            self.msg_list.selection_set(item)
            cursor = self.conn.cursor()
            cursor.execute("SELECT folder FROM messages WHERE id = ?", (item,))
            folder = cursor.fetchone()[0]
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="↩️ 返信", command=self.reply_mail)
            menu.add_command(label="↪️ 転送", command=self.forward_mail)
            if folder == "Drafts":
                menu.add_command(label="✏️ 下書きを編集", command=lambda: self._edit_draft(item))
            menu.add_separator()
            menu.add_command(label="🗑️ 削除", command=self.delete_selected_mail)
            menu.post(event.x_root, event.y_root)

    def open_compose_window(self, to="", subject="", body="", from_account_id=None, draft_id=None):
        """新規メール作成ウィンドウを開く"""
        # 既に作成画面が開いている場合は、そのウィンドウを前面に出す
        if self.compose_win and self.compose_win.winfo_exists():
            self.compose_win.deiconify()
            self.compose_win.lift()
            self.compose_win.focus_set()
            return

        self.compose_win = self._setup_sub_window("ComposeWindow", "新規作成", 600, 500)
        self.compose_win.focus_set()

        # 差出人アカウント選択
        ttk.Label(self.compose_win, text="差出人:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        from_accounts = [] 
        cursor = self.conn.cursor()
        cursor.execute("SELECT id, email, signature FROM accounts")
        for acc_id, email_addr, sig in cursor.fetchall():
            from_accounts.append({"id": acc_id, "email": email_addr, "sig": sig})

        from_account_var = tk.StringVar()
        acc_emails = [a["email"] for a in from_accounts]
        from_combo = ttk.Combobox(self.compose_win, textvariable=from_account_var, values=acc_emails, state="readonly", width=45)
        from_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # 初期選択アカウントの決定
        if from_account_id:
            for a in from_accounts:
                if a["id"] == from_account_id:
                    from_combo.set(a["email"])
                    break
        elif acc_emails:
            from_combo.set(acc_emails[0])

        # 署名挿入ロジック
        def apply_signature(*args):
            email_val = from_account_var.get()
            signature = next((a["sig"] for a in from_accounts if a["email"] == email_val), "")
            if signature:
                # 既存の署名がある場合は消してから入れる等の処理が必要だが、簡略化のため末尾に追加
                if not body_txt.get("1.0", tk.END).strip():
                    body_txt.insert(tk.END, f"\n\n--\n{signature}")

        from_account_var.trace_add("write", apply_signature)
        
        # 初期表示時に本文が空なら署名をセット
        if not body and from_combo.get():
            sig = next((a["sig"] for a in from_accounts if a["email"] == from_combo.get()), "")
            if sig: body = f"\n\n--\n{sig}"

        # 宛先
        ttk.Label(self.compose_win, text="宛先:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        to_frame = ttk.Frame(self.compose_win)
        to_frame.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        to_entry = ttk.Entry(to_frame, width=50)
        to_entry.pack(side=tk.LEFT)
        to_entry.insert(0, to)
        self.bind_context_menu(to_entry)
        btn_addr = ttk.Button(to_frame, text="📖", width=3, command=lambda: self.open_address_picker(to_entry))
        btn_addr.pack(side=tk.LEFT, padx=2)

        # 件名
        ttk.Label(self.compose_win, text="件名:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        subject_entry = ttk.Entry(self.compose_win, width=60)
        subject_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        subject_entry.insert(0, subject)
        self.bind_context_menu(subject_entry)

        # 添付ファイルリスト
        attached_files = []
        ttk.Label(self.compose_win, text="添付:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        attach_frame = ttk.Frame(self.compose_win)
        attach_frame.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        
        attach_list_lbl = ttk.Label(attach_frame, text="(なし)", foreground="gray")
        attach_list_lbl.pack(side=tk.LEFT)
        
        def add_attachment():
            paths = filedialog.askopenfilenames(parent=self.compose_win, title="添付ファイルを選択")
            if paths:
                attached_files.extend(paths)
                names = [os.path.basename(p) for p in attached_files]
                attach_list_lbl.config(text=", ".join(names), foreground="black")
            self.compose_win.lift()

        btn_attach = ttk.Button(attach_frame, text="ファイルを添付...", command=add_attachment)
        btn_attach.pack(side=tk.LEFT, padx=5)

        # 本文
        body_txt = tk.Text(self.compose_win, font=("MS Gothic", 11), height=20)
        body_txt.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        body_txt.insert("1.0", body)
        if to: body_txt.mark_set("insert", "1.0") # 返信時は先頭にカーソル
        
        self.bind_context_menu(body_txt)
        
        self.compose_win.rowconfigure(4, weight=1)
        self.compose_win.columnconfigure(1, weight=1)

        def save_draft():
            """現在入力を下書きとして保存"""
            sel_email = from_account_var.get()
            sel_acc_id = next((a["id"] for a in from_accounts if a["email"] == sel_email), None)
            if draft_id:
                self.conn.execute(
                    "UPDATE messages SET subject=?, sender=?, recipient=?, body=?, account_id=? WHERE id=?",
                    (subject_entry.get(), sel_email, to_entry.get(), body_txt.get("1.0", tk.END), sel_acc_id, draft_id)
                )
            else:
                self.conn.execute(
                    "INSERT INTO messages (subject, sender, recipient, date, body, folder, account_id) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (subject_entry.get(), sel_email, to_entry.get(), email.utils.formatdate(localtime=True), body_txt.get("1.0", tk.END), "Drafts", sel_acc_id)
                )
            self.conn.commit()
            self.refresh_folders()
            messagebox.showinfo("下書き", "下書きを保存しました。", parent=self.compose_win)
            self.compose_win.destroy()

        def send():
            to_addr = to_entry.get()
            subject_val = subject_entry.get()
            body_val = body_txt.get("1.0", tk.END)
            
            if not to_addr:
                messagebox.showwarning("警告", "宛先を入力してください。", parent=self.compose_win)
                return

            btn_send.config(state=tk.DISABLED, text="送信中...")
            self.compose_win.config(cursor="watch")
            self.compose_win.update()
            
            # 現在選択されているアカウントIDを取得
            sel_email = from_account_var.get()
            sel_acc_id = next((a["id"] for a in from_accounts if a["email"] == sel_email), None)

            res = importer.send_email(self.conn, to_addr, subject_val, body_val, sel_acc_id, attachment_paths=attached_files)

            if res == "success":
                self.compose_win.config(cursor="")
                from_email = sel_email
                sent_headers = f"To: {to_addr}\nFrom: {from_email}\nSubject: {subject_val}\nDate: {email.utils.formatdate(localtime=True)}\nX-Account-ID: {sel_acc_id}"
                cursor.execute(
                    "INSERT INTO messages (subject, sender, recipient, date, body, folder, headers, account_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    (subject_val, from_email, to_addr, email.utils.formatdate(localtime=True), body_val, "Sent", sent_headers, sel_acc_id)
                )
                sent_msg_id = cursor.lastrowid
                
                # 元が下書きだった場合は削除
                if draft_id:
                    self.conn.execute("DELETE FROM messages WHERE id = ?", (draft_id,))

                # 送信済みメールに添付ファイル情報を紐付け
                for p in attached_files:
                    with open(p, 'rb') as f:
                        cursor.execute(
                            "INSERT INTO attachments (message_id, filename, data) VALUES (?, ?, ?)",
                            (sent_msg_id, os.path.basename(p), sqlite3.Binary(f.read()))
                        )
                
                self.conn.commit()
                
                messagebox.showinfo("送信", "メールを送信しました。", parent=self.compose_win)
                self.refresh_folders()
                self.compose_win.destroy()
            else:
                self.compose_win.config(cursor="")
                btn_send.config(state=tk.NORMAL, text="送信")
                messagebox.showerror("エラー", f"送信に失敗しました:\n{res}", parent=self.compose_win)

        btn_frame = ttk.Frame(self.compose_win)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        # 送信ボタン
        btn_send = ttk.Button(btn_frame, text="送信", command=send)
        btn_send.pack(side=tk.LEFT, padx=5)
        
        # 下書き保存ボタン
        ttk.Button(btn_frame, text="下書き保存", command=save_draft).pack(side=tk.LEFT, padx=5)

    def _edit_draft(self, msg_id):
        """既存の下書きを編集"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT subject, sender, recipient, body, account_id FROM messages WHERE id = ?", (msg_id,))
        subj, sender, recp, body, acc_id = cursor.fetchone()
        self.open_compose_window(to=recp, subject=subj, body=body, from_account_id=acc_id, draft_id=msg_id)

    def open_address_book(self):
        """アドレス帳管理ウィンドウ"""
        addr_win = self._setup_sub_window("AddressBookWindow", "アドレス帳", 500, 400)

        # 入力エリア
        input_frame = ttk.LabelFrame(addr_win, text="新規登録/編集")
        input_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(input_frame, text="ﾆｯｸﾈｰﾑ:").grid(row=0, column=0, padx=5, pady=5)
        nick_ent = ttk.Entry(input_frame, width=15)
        nick_ent.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(input_frame, text="名前:").grid(row=0, column=2, padx=5, pady=5)
        name_ent = ttk.Entry(input_frame)
        name_ent.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(input_frame, text="メール:").grid(row=1, column=0, padx=5, pady=5)
        mail_ent = ttk.Entry(input_frame)
        mail_ent.grid(row=1, column=1, columnspan=3, padx=5, pady=5, sticky="ew")

        # リスト表示
        tree = ttk.Treeview(addr_win, columns=("nick", "name", "email"), show="headings")
        tree.heading("nick", text="ニックネーム")
        tree.heading("name", text="名前")
        tree.heading("email", text="メールアドレス")
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        def refresh_list():
            tree.delete(*tree.get_children())
            for row in self.conn.execute("SELECT id, nickname, name, email FROM address_book ORDER BY nickname, name"):
                tree.insert("", tk.END, iid=row[0], values=(row[1] or "", row[2], row[3]))

        def add_addr():
            if name_ent.get() and mail_ent.get():
                self.conn.execute("INSERT INTO address_book (nickname, name, email) VALUES (?, ?, ?)", (nick_ent.get(), name_ent.get(), mail_ent.get()))
                self.conn.commit()
                nick_ent.delete(0, tk.END); name_ent.delete(0, tk.END); mail_ent.delete(0, tk.END)
                refresh_list()

        def delete_addr():
            sel = tree.selection()
            if sel:
                self.conn.execute("DELETE FROM address_book WHERE id = ?", (sel[0],))
                self.conn.commit()
                refresh_list()

        btn_frame = ttk.Frame(addr_win)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Button(btn_frame, text="追加", command=add_addr).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="削除", command=delete_addr).pack(side=tk.LEFT, padx=5)
        
        refresh_list()

    def open_address_picker(self, target_entry):
        """宛先選択用ダイアログ"""
        picker = self._setup_sub_window("AddressPickerWindow", "アドレス選択", 400, 300)

        tree = ttk.Treeview(picker, columns=("nick", "name", "email"), show="headings", selectmode="browse")
        tree.heading("nick", text="見出し")
        tree.heading("name", text="名前")
        tree.heading("email", text="メールアドレス")
        tree.column("nick", width=80)
        tree.pack(fill=tk.BOTH, expand=True)

        for row in self.conn.execute("SELECT nickname, name, email FROM address_book ORDER BY nickname, name"):
            tree.insert("", tk.END, values=(row[0] or "", row[1], row[2]))

        def select():
            sel = tree.selection()
            if sel:
                item = tree.item(sel[0], "values") # (nick, name, email)
                # AL-Mail風に "名前 <email>" 形式でセット
                target_entry.delete(0, tk.END)
                target_entry.insert(0, f"{item[1]} <{item[2]}>")
                picker.destroy()

        ttk.Button(picker, text="決定", command=select).pack(pady=10)
        tree.bind("<Double-1>", lambda e: select())

    def receive_mail(self, silent=False):
        """新着メールチェックの実行"""
        # 現在の選択状態を確認（受信箱を開いているかどうか）
        selected = self.folder_tree.selection()
        current_iid = selected[0] if selected else ""
        is_inbox = (current_iid == "UNIFIED_INBOX") or (current_iid.startswith("FLD_") and current_iid.endswith("_Inbox"))

        if not silent:
            self.root.config(cursor="watch")
            self.root.update()
            
        res = importer.fetch_emails(self.conn)
        
        # 新着メールがあれば通知（簡易的に最新1件を取得して通知）
        if res == "success":
            cursor = self.conn.cursor()
            cursor.execute("SELECT subject, sender FROM messages ORDER BY id DESC LIMIT 1")
            new_mail = cursor.fetchone()
            if new_mail and not silent:
                 self.notify_new_mail(new_mail[0], new_mail[1])
        
        if not silent:
            self.root.config(cursor="")
            if res == "success":
                self.refresh_folders()
                self.update_home_view()
                
                if is_inbox:
                    # 受信箱を開いている場合は、その場所を維持するが
                    # 最新メールへの自動フォーカスをスキップするフラグを立てる
                    if self.folder_tree.exists(current_iid):
                        self._skip_auto_select_first = True
                        self.folder_tree.selection_set(current_iid)
                        self.folder_tree.see(current_iid)
                        self._skip_auto_select_first = False
                else:
                    # 受信箱以外なら統合受信箱へジャンプ（自動フォーカス有効）
                    if self.folder_tree.exists("UNIFIED_INBOX"):
                        self.folder_tree.selection_set("UNIFIED_INBOX")
                        self.folder_tree.see("UNIFIED_INBOX")

                messagebox.showinfo("受信", "受信が完了しました。")
            elif res == "settings_missing":
                messagebox.showwarning("設定", "アカウント設定（IMAPサーバー等）を先に行ってください。")
            else:
                messagebox.showerror("エラー", f"受信中にエラーが発生しました:\n{res}")
        else:
            # サイレントモード（自動受信）時は成功時のみリフレッシュ
            if res == "success":
                self.refresh_folders()
                self.update_home_view()

    def init_auto_receive(self):
        """起動時に自動受信が有効ならタイマーを開始"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT auto_receive_enabled, auto_receive_interval FROM accounts LIMIT 1")
        row = cursor.fetchone()
        if row and row[0]:
            self.schedule_auto_receive(row[1])

    def schedule_auto_receive(self, interval_mins):
        """指定分後に自動受信を実行するように予約"""
        if self.auto_receive_job:
            self.root.after_cancel(self.auto_receive_job)
        
        ms = int(interval_mins) * 60 * 1000
        self.auto_receive_job = self.root.after(ms, self.run_auto_receive_cycle)

    def run_auto_receive_cycle(self):
        """自動受信の実行と次回の予約"""
        self.receive_mail(silent=True)
        # 次回予約のために設定を再確認
        cursor = self.conn.cursor()
        cursor.execute("SELECT auto_receive_enabled, auto_receive_interval FROM accounts LIMIT 1")
        row = cursor.fetchone()
        if row and row[0]:
            self.schedule_auto_receive(row[1])

    def open_settings(self):
        """アカウント管理ダイアログ（マルチアカウント対応版）"""
        settings_win = self._setup_sub_window("SettingsWindow", "Petal 設定", 650, 450, modal=True)
        
        nb = ttk.Notebook(settings_win)
        nb.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- タブ1: アカウント管理 ---
        acc_manage_frame = ttk.Frame(nb)
        nb.add(acc_manage_frame, text=" アカウント管理 ")

        list_frame = ttk.Frame(acc_manage_frame)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        columns = ("email", "server")
        tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse")
        tree.heading("email", text="メールアドレス")
        tree.heading("server", text="受信(IMAP)サーバー")
        tree.pack(fill=tk.BOTH, expand=True)

        def refresh_account_list():
            tree.delete(*tree.get_children())
            for row in self.conn.execute("SELECT id, email, imap_server FROM accounts"):
                tree.insert("", tk.END, iid=row[0], values=(row[1], row[2]))

        # 右側のボタン操作
        btn_frame = ttk.Frame(acc_manage_frame)
        btn_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0), pady=10)
        
        def on_add(): self.open_account_edit_dialog(None, refresh_account_list)
        def on_edit():
            sel = tree.selection()
            if sel: self.open_account_edit_dialog(sel[0], refresh_account_list)

        # ダブルクリックで編集を開くようにバインド
        tree.bind("<Double-1>", lambda e: on_edit())

        def on_delete():
            sel = tree.selection()
            if sel and messagebox.askyesno("削除確認", "このアカウント設定を削除しますか？\n(メールデータは保持されます)"):
                self.conn.execute("DELETE FROM accounts WHERE id = ?", (sel[0],))
                self.conn.commit()
                refresh_account_list()
                self.refresh_folders()

        ttk.Button(btn_frame, text="新規追加", command=on_add).pack(fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="編集...", command=on_edit).pack(fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="削除", command=on_delete).pack(fill=tk.X, pady=2)
        
        refresh_account_list()

        # --- タブ2: 全般・詳細設定 ---
        # マルチアカウント化に伴い、利便性のために共通的な設定をここに表示します。
        # 変更は全てのアカウントに一括適用される仕様とします。
        gen_frame = ttk.Frame(nb)
        nb.add(gen_frame, text=" 全般設定 ")

        # 現在の設定値を取得（最初の1件を代表とする）
        cursor = self.conn.cursor()
        cursor.execute("SELECT display_html_as_text, minimize_to_tray, notify_new_mail, search_almail_at_startup FROM accounts LIMIT 1")
        row = cursor.fetchone() or (0, 0, 0, 0)

        html_as_text_var = tk.BooleanVar(value=bool(row[0]))
        tray_var = tk.BooleanVar(value=bool(row[1]))
        notify_var = tk.BooleanVar(value=bool(row[2]))
        search_startup_var = tk.BooleanVar(value=bool(row[3]))

        container = ttk.Frame(gen_frame)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        ttk.Checkbutton(container, text="HTMLメールをテキストで表示する", variable=html_as_text_var).pack(pady=10, anchor="w")
        ttk.Checkbutton(container, text="最小化時はタスクアイコン（トレイ）に入れる", variable=tray_var).pack(pady=10, anchor="w")
        ttk.Checkbutton(container, text="受信メールをWindowsから通知する", variable=notify_var).pack(pady=10, anchor="w")
        ttk.Checkbutton(container, text="起動時に AL-Mail フォルダを探索する", variable=search_startup_var).pack(pady=10, anchor="w")
        
        def save_common_settings():
            try:
                # 全てのアカウントの設定を一括更新
                self.conn.execute("""
                    UPDATE accounts 
                    SET display_html_as_text=?, minimize_to_tray=?, notify_new_mail=?, search_almail_at_startup=?
                """, (int(html_as_text_var.get()), int(tray_var.get()), int(notify_var.get()), int(search_startup_var.get())))
                self.conn.commit()
                messagebox.showinfo("保存", "全般設定を保存しました。")
            except Exception as e:
                messagebox.showerror("エラー", f"保存に失敗しました:\n{e}")

        ttk.Button(container, text="設定を全てのカウントに適用して保存", command=save_common_settings).pack(pady=20, anchor="w")
        
        ttk.Label(container, text="※サーバー設定や署名は「アカウント管理」から各アカウントを編集してください。", 
                  foreground="gray").pack(side=tk.BOTTOM, pady=10)

    def open_account_edit_dialog(self, account_id, callback):
        """個別アカウントの詳細設定画面"""
        edit_win = tk.Toplevel(self.root)
        edit_win.title("アカウント設定")
        edit_win.geometry("480x520")
        edit_win.grab_set()

        nb = ttk.Notebook(edit_win)
        nb.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        acc_frame = ttk.Frame(nb); nb.add(acc_frame, text="アカウント")
        adv_frame = ttk.Frame(nb); nb.add(adv_frame, text="詳細設定")
        auto_rx_frame = ttk.Frame(nb); nb.add(auto_rx_frame, text="自動受信")
        sig_frame = ttk.Frame(nb); nb.add(sig_frame, text="署名")

        # プロトコル選択
        ttk.Label(acc_frame, text="受信プロトコル").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        protocol_var = tk.StringVar(value="IMAP")
        protocol_combo = ttk.Combobox(acc_frame, textvariable=protocol_var, values=("IMAP", "POP3"), state="readonly", width=27)
        protocol_combo.grid(row=0, column=1, padx=10, pady=5)

        fields = [
            ("メールアドレス", "email"),
            ("SMTPサーバー", "smtp_server"),
            ("SMTPポート", "smtp_port"),
            ("IMAPサーバー", "imap_server"),
            ("IMAPポート", "imap_port"),
            ("ユーザー名", "username"),
            ("パスワード", "password")
        ]
        
        entries = {}
        for i, (label, key) in enumerate(fields, start=1):
            ttk.Label(acc_frame, text=label).grid(row=i, column=0, padx=10, pady=5, sticky="e")
            entry = ttk.Entry(acc_frame, width=30)
            if key == "password": entry.config(show="*")
            entry.grid(row=i, column=1, padx=10, pady=5)
            entries[key] = entry
            self.bind_context_menu(entry)

        # 詳細設定タブの内容
        html_as_text_var = tk.BooleanVar()
        tray_var = tk.BooleanVar()
        notify_var = tk.BooleanVar()
        
        ttk.Checkbutton(adv_frame, text="HTMLメールをテキストで表示する", variable=html_as_text_var).pack(padx=20, pady=10, anchor="w")
        ttk.Checkbutton(adv_frame, text="最小化時はタスクアイコン（トレイ）に入れる", variable=tray_var).pack(padx=20, pady=10, anchor="w")
        ttk.Checkbutton(adv_frame, text="受信メールをWindowsから通知する", variable=notify_var).pack(padx=20, pady=10, anchor="w")

        # 自動受信タブの内容
        auto_rx_enabled_var = tk.BooleanVar()
        
        def toggle_interval_state():
            state = tk.NORMAL if auto_rx_enabled_var.get() else tk.DISABLED
            interval_ent.config(state=state)

        ttk.Checkbutton(auto_rx_frame, text="自動受信する", variable=auto_rx_enabled_var, command=toggle_interval_state).pack(padx=20, pady=20, anchor="w")
        
        int_frame = ttk.Frame(auto_rx_frame)
        int_frame.pack(padx=40, anchor="w")
        interval_ent = ttk.Entry(int_frame, width=5)
        interval_ent.pack(side=tk.LEFT)
        ttk.Label(int_frame, text=" 分おきに自動受信").pack(side=tk.LEFT)

        # 署名設定タブ
        sig_text = tk.Text(sig_frame, height=10)
        sig_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.bind_context_menu(sig_text)

        # 既存データの読み込み (編集時)
        if account_id:
            cursor = self.conn.cursor()
            cursor.execute("SELECT email, smtp_server, smtp_port, imap_server, imap_port, username, password, display_html_as_text, minimize_to_tray, notify_new_mail, auto_receive_enabled, auto_receive_interval, signature, protocol FROM accounts WHERE id = ?", (account_id,))
            row = cursor.fetchone()
            if row:
                for entry, value in zip(entries.values(), row[:7]): entry.insert(0, str(value))
                html_as_text_var.set(bool(row[7]))
                tray_var.set(bool(row[8]))
                notify_var.set(bool(row[9]))
                auto_rx_enabled_var.set(bool(row[10]))
                interval_ent.insert(0, str(row[11]))
                if row[12]: sig_text.insert("1.0", row[12])
                if len(row) > 13 and row[13]: protocol_var.set(row[13])
            toggle_interval_state()

        def test_conn():
            edit_win.config(cursor="watch")
            edit_win.update()
            res = importer.test_connection(
                entries['imap_server'].get(), entries['imap_port'].get(),
                entries['smtp_server'].get(), entries['smtp_port'].get(),
                entries['username'].get(), entries['password'].get()
            )
            edit_win.config(cursor="")
            messagebox.showinfo("接続テスト結果", res, parent=edit_win)

        def save():
            data = {k: v.get() for k, v in entries.items()}
            params = (data['email'], data['smtp_server'], data['smtp_port'], data['imap_server'], data['imap_port'], data['username'], data['password'], 
                      int(html_as_text_var.get()), int(tray_var.get()), int(notify_var.get()), int(auto_rx_enabled_var.get()), int(interval_ent.get() or 10), sig_text.get("1.0", tk.END), protocol_var.get())
            
            if account_id:
                self.conn.execute("""UPDATE accounts SET email=?, smtp_server=?, smtp_port=?, imap_server=?, imap_port=?, username=?, password=?, 
                                     display_html_as_text=?, minimize_to_tray=?, notify_new_mail=?, auto_receive_enabled=?, auto_receive_interval=?, signature=?, protocol=? WHERE id=?""", params + (account_id,))
            else:
                self.conn.execute("""INSERT INTO accounts (email, smtp_server, smtp_port, imap_server, imap_port, username, password, 
                                     display_html_as_text, minimize_to_tray, notify_new_mail, auto_receive_enabled, auto_receive_interval, signature, protocol) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", params)
            
            self.conn.commit()
            callback()
            self.refresh_folders()
            edit_win.destroy()

        btn_frame = ttk.Frame(edit_win)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="接続テスト", command=test_conn).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="保存", command=save).pack(side=tk.LEFT, padx=5)

    def import_dialog(self):
        """AL-Mailフォルダを特定のアカウントへインポート"""
        # インポート先のアカウントを選択させる
        cursor = self.conn.cursor()
        cursor.execute("SELECT id, email FROM accounts")
        accounts = cursor.fetchall()
        
        if not accounts:
            if messagebox.askyesno("確認", "インポート先のアカウントが登録されていません。\nインポート専用の「ローカル用アカウント」を作成して続行しますか？"):
                self.conn.execute("INSERT INTO accounts (email, protocol) VALUES (?, ?)", ("local_archive@petal", "POP3"))
                self.conn.commit()
                self.refresh_folders()
                # 作成したアカウントをリストに反映
                cursor.execute("SELECT id, email FROM accounts")
                accounts = cursor.fetchall()
            else:
                return

        # 簡易的なアカウント選択ダイアログ
        select_win = tk.Toplevel(self.root)
        select_win.title("インポート先アカウントの選択")
        select_win.geometry("400x150")
        select_win.grab_set()

        ttk.Label(select_win, text="インポートしたメールを紐付けるアカウントを選択してください:").pack(pady=10)
        acc_var = tk.StringVar()
        acc_combo = ttk.Combobox(select_win, textvariable=acc_var, values=[a[1] for a in accounts], state="readonly", width=40)
        acc_combo.pack(pady=5)
        acc_combo.current(0)

        def proceed_import():
            email_val = acc_var.get()
            acc_id = next(a[0] for a in accounts if a[1] == email_val)
            select_win.destroy()

            folder = filedialog.askdirectory(title="AL-Mailのメールボックスフォルダを選択")
            if folder:
                importer.import_from_almail(folder, self.conn, acc_id)
                self.refresh_folders()
                messagebox.showinfo("完了", "インポートが完了しました。")

        ttk.Button(select_win, text="次へ", command=proceed_import).pack(pady=10)

    def check_almail_on_startup(self):
        """設定が有効な場合、AL-Mailのデフォルトフォルダを探索してインポートを促す"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT search_almail_at_startup FROM accounts LIMIT 1")
        row = cursor.fetchone()
        if not (row and row[0]): return

        # AL-Mail の標準的なデータパスを探索
        potential_dirs = [r"C:\ALMail", r"C:\Program Files (x86)\ALMail"]
        found_path = None
        for d in potential_dirs:
            if os.path.exists(os.path.join(d, "Mail")):
                found_path = d
                break
        
        if found_path:
            if messagebox.askyesno("AL-Mail 探索", f"AL-Mail のインストール環境が見つかりました:\n{found_path}\n\n設定、アドレス帳、メールをすべてインポートしますか？"):
                # 1. アカウント設定のインポート
                ini_path = os.path.join(found_path, "ALMAIL.INI")
                acc_id = importer.import_almail_settings(ini_path, self.conn)
                
                # 2. アドレス帳のインポート
                adr_path = os.path.join(found_path, "ALMAIL.ADR")
                importer.import_address_book(adr_path, self.conn)
                
                # アカウントが特定できない場合は既存の1つ目を使用
                if not acc_id:
                    cursor.execute("SELECT id FROM accounts LIMIT 1")
                    acc = cursor.fetchone()
                    if acc:
                        acc_id = acc[0]
                    else:
                        # まったくアカウントがない場合、受け皿としてローカルアカウントを作成
                        cursor.execute("INSERT INTO accounts (email, protocol) VALUES (?, ?)", ("local_archive@petal", "POP3"))
                        acc_id = cursor.lastrowid

                if acc_id:
                    # 3. メールデータのインポート
                    mail_root = os.path.join(found_path, "Mail")
                    for sub in os.listdir(mail_root):
                        sub_path = os.path.join(mail_root, sub)
                        if os.path.isdir(sub_path):
                            # AL-Mailのフォルダ名をPetalのフォルダ名にマップ
                            folder_map = {"Inbox": "Inbox", "Sent": "Sent", "Draft": "Drafts", "Trash": "Trash"}
                            target_fld = folder_map.get(sub, sub)
                            importer.import_from_almail(sub_path, self.conn, acc_id)
                            if target_fld != sub:
                                self.conn.execute("UPDATE messages SET folder=? WHERE folder=? AND account_id=?", (target_fld, sub, acc_id))
                    
                    self.conn.commit()
                    self.refresh_folders()
                    messagebox.showinfo("完了", "AL-Mail からのインポートがすべて完了しました。")
        
        # 処理が終わったら（見つからなかった場合も含め）設定をオフにする
        self.conn.execute("UPDATE accounts SET search_almail_at_startup = 0")
        self.conn.commit()

    def run_export(self, tables):
        """指定されたテーブルのデータをJSON形式でエクスポート"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            title="エクスポート先の保存"
        )
        if not file_path: return

        data = {}
        cursor = self.conn.cursor()
        for table in tables:
            cursor.execute(f"SELECT * FROM {table}")
            cols = [column[0] for column in cursor.description]
            rows = cursor.fetchall()
            data[table] = [dict(zip(cols, row)) for row in rows]
        
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", "エクスポートが完了しました。")
        except Exception as e:
            messagebox.showerror("エラー", f"エクスポートに失敗しました:\n{e}")

    def run_import(self, filter_tables):
        """JSON形式のファイルから指定されたテーブルのデータをインポート"""
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")],
            title="インポートするファイルを選択"
        )
        if not file_path: return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            cursor = self.conn.cursor()
            for table, rows in data.items():
                # 選択されたメニューのカテゴリに含まれないテーブルはスキップ
                if table not in filter_tables: continue
                if not rows: continue
                
                # 現在のテーブルの列名を取得（互換性チェック用）
                cursor.execute(f"PRAGMA table_info({table})")
                current_cols = [col[1] for col in cursor.fetchall()]
                
                if table == 'accounts':
                    cursor.execute("DELETE FROM accounts") # アカウントは上書き
                
                for row in rows:
                    if 'id' in row: del row['id'] # 自動採番に任せる
                    # DBに存在する列のみ抽出（カラム不一致によるエラー防止）
                    row_data = {k: v for k, v in row.items() if k in current_cols}
                    if not row_data: continue

                    # メールの重複チェック
                    if table == 'messages' and row_data.get('message_id'):
                        cursor.execute("SELECT id FROM messages WHERE message_id = ?", (row_data['message_id'],))
                        if cursor.fetchone(): continue
                    
                    cols = list(row_data.keys())
                    placeholders = ", ".join(["?"] * len(cols))
                    col_names = ", ".join(cols)
                    cursor.execute(f"INSERT INTO {table} ({col_names}) VALUES ({placeholders})", tuple(row_data.values()))
            
            self.conn.commit()
            self.refresh_folders()
            self.update_home_view()
            messagebox.showinfo("成功", "インポートが完了しました。")
        except Exception as e:
            messagebox.showerror("エラー", f"インポートに失敗しました:\n{e}")

    def refresh_folders(self):
        """フォルダツリーの更新（ホームと受信箱のアイコン対応）"""
        self.folder_tree.delete(*self.folder_tree.get_children())
        
        # ホームを固定で追加
        self.folder_tree.insert("", "end", iid="HOME_NODE", text="🏠 ホーム")
        self.folder_tree.insert("", "end", iid="UNIFIED_INBOX", text="📥 すべての受信箱")

        cursor = self.conn.cursor()
        cursor.execute("SELECT id, email FROM accounts")
        for acc_id, email_addr in cursor.fetchall():
            acc_node = self.folder_tree.insert("", "end", iid=f"ACC_{acc_id}", text=f"👤 {email_addr}", open=True)
            self.folder_tree.insert(acc_node, "end", iid=f"FLD_{acc_id}_Inbox", text="  📥 受信箱")
            self.folder_tree.insert(acc_node, "end", iid=f"FLD_{acc_id}_Drafts", text="  📝 下書き")
            self.folder_tree.insert(acc_node, "end", iid=f"FLD_{acc_id}_Sent", text="  📤 送信済み")
            self.folder_tree.insert(acc_node, "end", iid=f"FLD_{acc_id}_Trash", text="  🗑️ ゴミ箱")

    def on_folder_select(self, event):
        selected = self.folder_tree.selection()
        if not selected: return
        iid = selected[0]

        if iid == "HOME_NODE":
            self.v_pane.pack_forget()
            self.home_frame.pack(fill=tk.BOTH, expand=True)
            self.update_home_view()
            return

        self.home_frame.pack_forget()
        self.v_pane.pack(fill=tk.BOTH, expand=True)
        self.msg_list.delete(*self.msg_list.get_children())
        cursor = self.conn.cursor()

        if iid == "UNIFIED_INBOX":
            # 全アカウントの受信箱を統合して表示（アカウント名も取得）
            cursor.execute("""
                SELECT m.id, m.subject, m.sender, m.date, a.email 
                FROM messages m LEFT JOIN accounts a ON m.account_id = a.id 
                WHERE m.folder = 'Inbox' ORDER BY m.date DESC""")
        elif iid.startswith("FLD_"):
            # 個別アカウントの特定フォルダ
            # iid format: FLD_{acc_id}_{FolderName}
            parts = iid.split("_")
            if len(parts) >= 3:
                acc_id, folder = parts[1], parts[2]
                cursor.execute("""
                    SELECT m.id, m.subject, m.sender, m.date, a.email 
                    FROM messages m LEFT JOIN accounts a ON m.account_id = a.id 
                    WHERE m.account_id = ? AND m.folder = ? ORDER BY m.date DESC""", (int(acc_id), folder))
            else: return
        else: return

        for row in cursor.fetchall():
            self.msg_list.insert("", "end", iid=row[0], values=(row[1], row[2], row[3], row[4] or ""))
        
        children = self.msg_list.get_children()
        if children and not self._skip_auto_select_first:
            self.msg_list.selection_set(children[0])
            self.msg_list.focus(children[0])

        # フォルダ選択後、メッセージはまだ選択されていない状態
        if not children:
            self.auth_banner.config(text="🛡️ メールがありません", bg="#9e9e9e", fg="white")

    def on_message_select(self, event):
        selected = self.msg_list.selection()
        if not selected: return
        msg_id = selected[0]


        self.root.config(cursor="watch")
        self.root.update_idletasks()

        try:
            # メール情報の取得
            cursor = self.conn.cursor()
            cursor.execute("SELECT body, body_html, sender, recipient, folder, account_id FROM messages WHERE id = ?", (msg_id,))
            res = cursor.fetchone()
            if not res: return
            body, body_html, sender, recipient, folder, account_id = res
            self.current_message_account_id = account_id
            
            # 認証情報の判定（送信済みフォルダ以外のみ表示）
            if folder and folder.lower() == "sent":
                self.auth_banner.pack_forget()
                self.lbl_from_to.config(text=f"📤 送信先: {recipient or '???'}")
            else:
                self.auth_banner.pack(side=tk.TOP, fill=tk.X) # 非送信済みメールの場合は再表示
                self.check_message_security(msg_id)
                self.lbl_from_to.config(text=f"From: {sender}   |   To: {recipient or '???'}")

            # 添付ファイルの更新
            cursor.execute("SELECT id, filename, data FROM attachments WHERE message_id = ?", (msg_id,))
            attachments_full = cursor.fetchall()
            self.update_attachment_bar([(a[0], a[1]) for a in attachments_full])

            cursor.execute("SELECT display_html_as_text FROM accounts WHERE id = ?", (account_id,))
            row = cursor.fetchone()
            html_as_text = row[0] if row else False
            
            # --- 巨大HTMLフリーズ対策 ---
            # 150KBを超えるHTMLはレンダリングエンジンが固まる可能性が高いため強制的にテキストモードへ
            is_complex_html = body_html and len(body_html) > 150000
            if is_complex_html and not html_as_text:
                html_as_text = True
                self.auth_banner.config(
                    text="⚠️ メールサイズが大きいため、表示速度を優先しテキストモードで表示しました", 
                    bg="#ff9800", fg="white"
                )

            # --- Google Security Alert 特殊表示 (Intelligent Parsing) ---
            if sender and "no-reply@accounts.google.com" in sender:
                if self._handle_google_security_alert(body, body_html, attachments_full):
                    return

            # 表示エリアのリセット
            self.body_frame.pack_forget()
            self.html_view.pack_forget()

            if body_html and not html_as_text:
                # HTML表示
                self.html_view.pack(fill=tk.BOTH, expand=True)
                styled_html = (
                    f"<html><head><style>"
                    f"body {{ font-family: 'Segoe UI', 'Meiryo', sans-serif; line-height: 1.6; "
                    f"color: #2c3e50; background-color: #ffffff; padding: 25px; margin: 0; }}"
                    f"a {{ color: #3498db; text-decoration: none; }}"
                    f"hr {{ border: 0; border-top: 1px solid #eee; margin: 20px 0; }}"
                    f"</style></head><body>{body_html}</body></html>"
                )
                
                self.html_view.load_html(styled_html)
            else:
                # テキスト表示
                self.body_frame.pack(fill=tk.BOTH, expand=True)
                self.body_text.config(state=tk.NORMAL)
                self.body_text.delete("1.0", tk.END)
                
                # プレーンテキストがあれば優先、なければHTMLから変換（判定を軽量化）
                if body and body.strip():
                    display_content = body
                elif body_html:
                    display_content = self.strip_html_tags(body_html)
                else:
                    display_content = "(本文なし)"

                if html_as_text and body_html:
                    stripped = self.strip_html_tags(body_html)
                    if stripped and stripped != display_content:
                        display_content = f"{display_content}\n\n--- [HTML版のテキスト抽出] ---\n{stripped}"
                
                self.body_text.insert(tk.END, display_content)
                self.body_text.config(state=tk.DISABLED)

            # 最後にプレビューを生成（body_textの状態が確定してから）
            self.show_image_previews(attachments_full)
        
        finally:
            # 処理が終わったら必ずカーソルを元に戻す
            self.root.config(cursor="")

    def _handle_google_security_alert(self, body, body_html, attachments_full):
        """Google Security Alert 特殊表示のパースと描画"""
        content_for_parse = body if (body and body.strip()) else self.strip_html_tags(body_html)
        
        dev = re.search(r'デバイス:\s*(.*)', content_for_parse)
        loc = re.search(r'場所:\s*(.*)', content_for_parse)
        tim = re.search(r'時間:\s*(.*)', content_for_parse)
        
        if not (dev or loc or tim):
            return False

        petal_google_html = f"""
        <html><head><style>
            body {{ font-family: 'Segoe UI', 'Meiryo', sans-serif; background-color: #fefefe; padding: 40px; margin: 0; }}
            .petal-card {{ 
                background: #ffffff; border-radius: 16px; border: 1px solid #ff80ab; 
                box-shadow: 0 10px 25px rgba(216, 27, 96, 0.1); padding: 30px; 
                max-width: 480px; margin: 0 auto; 
            }}
            .petal-header {{ 
                color: #d81b60; font-size: 20px; font-weight: bold; margin-bottom: 25px; 
                text-align: center; border-bottom: 1px dashed #ff80ab; padding-bottom: 15px;
            }}
            .petal-row {{ margin-bottom: 20px; }}
            .petal-label {{ font-size: 11px; color: #999; text-transform: uppercase; font-weight: bold; margin-bottom: 4px; }}
            .petal-value {{ font-size: 16px; color: #333; font-weight: 500; }}
            .petal-footer {{ margin-top: 30px; font-size: 11px; color: #ccc; text-align: center; font-style: italic; }}
        </style></head><body>
        <div class="petal-card">
            <div class="petal-header">🛡️ Google アカウントの保護</div>
            <div class="petal-row">
                <div class="petal-label">デバイス</div>
                <div class="petal-value">💻 {html.escape(dev.group(1).strip() if dev else "不明")}</div>
            </div>
            <div class="petal-row">
                <div class="petal-label">場所</div>
                <div class="petal-value">📍 {html.escape(loc.group(1).strip() if loc else "不明")}</div>
            </div>
            <div class="petal-row">
                <div class="petal-label">時間</div>
                <div class="petal-value">🕒 {html.escape(tim.group(1).strip() if tim else "不明")}</div>
            </div>
            <div class="petal-footer">Parsed by Petal Intelligent Engine</div>
        </div>
        </body></html>
        """
        self.body_frame.pack_forget()
        self.html_view.pack(fill=tk.BOTH, expand=True)
        self.html_view.load_html(petal_google_html)
        self.show_image_previews(attachments_full)
        return True

    def update_attachment_bar(self, attachments):
        """添付ファイルリストを更新（フローレイアウト対応）"""
        self.attachment_bar.config(state=tk.NORMAL)
        self.attachment_bar.delete("1.0", tk.END)
        
        if not attachments:
            self.att_frame.pack_forget()
            return

        self.att_frame.pack(side=tk.TOP, fill=tk.X)
        self.attachment_bar.insert(tk.END, "📎 添付: ")
        
        for att_id, filename in attachments:
            btn = tk.Button(self.attachment_bar, text=filename, font=("MS UI Gothic", 9, "underline"), 
                            fg="blue", bg="#f0f0f0", relief=tk.FLAT, cursor="hand2",
                            command=lambda aid=att_id, fn=filename: self.download_attachment(aid, fn))
            self.attachment_bar.window_create(tk.END, window=btn)
            self.attachment_bar.insert(tk.END, "  ")

        # 行数に合わせて高さを調整（最大3行）
        self.attachment_bar.update_idletasks()
        line_count = int(self.attachment_bar.index('end-1c').split('.')[0])
        new_height = min(line_count, 3)
        self.attachment_bar.config(height=new_height, state=tk.DISABLED)
        if line_count > 3:
            self.att_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        else:
            self.att_scroll.pack_forget()

    def download_attachment(self, att_id, filename):
        """添付ファイルを保存する"""
        save_path = filedialog.asksaveasfilename(initialfile=filename, title="ファイルを保存")
        if not save_path: return
        
        cursor = self.conn.cursor()
        cursor.execute("SELECT data FROM attachments WHERE id = ?", (att_id,))
        data = cursor.fetchone()[0]
        
        with open(save_path, 'wb') as f:
            f.write(data)
        messagebox.showinfo("成功", f"ファイルを保存しました:\n{save_path}")

    def show_image_previews(self, attachments):
        """画像添付ファイルがある場合に本文の下にプレビューを表示"""
        for widget in self.preview_frame.winfo_children():
            widget.destroy()
        self.body_preview_pane.forget(self.preview_container)
        self.preview_images = []

        # 対応形式（Pillowを使用するため大幅に増加）
        image_exts = ('.png', '.gif', '.jpg', '.jpeg', '.bmp', '.webp')
        
        has_previews = False
        for _, filename, data in attachments:
            if filename.lower().endswith(image_exts):
                try:
                    # PILを使用して高品質に縮小 (thumbnail)
                    img_pil = Image.open(io.BytesIO(data))
                    # ペインの幅に合わせてリサイズ（最大幅800pxに拡大）
                    img_pil.thumbnail((800, 800), Image.Resampling.LANCZOS)
                    img = ImageTk.PhotoImage(img_pil)
                    
                    if self.body_frame.winfo_viewable():
                        # テキスト表示中の場合は、テキストの末尾に画像を埋め込む（これでスクロール可能になる）
                        self.body_text.config(state=tk.NORMAL)
                        self.body_text.insert(tk.END, "\n\n")
                        self.body_text.image_create(tk.END, image=img)
                        self.body_text.insert(tk.END, f"\n ({filename})\n")
                        self.body_text.config(state=tk.DISABLED)
                    else:
                        # HTML表示時などは従来通り下のフレームに表示
                        lbl = tk.Label(self.preview_frame, image=img, bg="white", pady=10)
                        lbl.pack()
                    
                    self.preview_images.append(img)
                    has_previews = True
                except Exception as e:
                    print(f"Preview failed for {filename}: {e}")
                    continue
        
        # HTML表示時かつ画像がある場合のみプレビューフレームを表示
        if has_previews and not self.body_frame.winfo_viewable():
            self.body_preview_pane.add(self.preview_container, weight=1)
            self.body_preview_pane.sashpos(0, self.content_container.winfo_height() - 250)

    def open_attachments_manager(self):
        """全添付ファイルの一覧管理ウィンドウ"""
        win = self._setup_sub_window("AttachmentManager", "添付ファイル一覧", 700, 450)
        
        columns = ("filename", "sender", "date", "subject")
        tree = ttk.Treeview(win, columns=columns, show="headings")
        tree.heading("filename", text="ファイル名")
        tree.heading("sender", text="送信者")
        tree.heading("date", text="受信日時")
        tree.heading("subject", text="メール件名")
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT a.id, a.filename, m.sender, m.date, m.subject, m.id, m.folder
            FROM attachments a 
            JOIN messages m ON a.message_id = m.id 
            ORDER BY m.id DESC
        """)
        
        for row in cursor.fetchall():
            tree.insert("", tk.END, iid=row[0], values=(row[1], row[2], row[3], row[4]), tags=(row[5], row[6]))

        def on_double_click(event):
            sel = tree.selection()
            if not sel: return
            item = tree.item(sel[0])
            # ダブルクリックで該当のメールへジャンプ
            msg_id, folder = item['tags']
            win.destroy()
            self.jump_to_message(msg_id, folder)

        tree.bind("<Double-1>", on_double_click)
        ttk.Label(win, text="※ダブルクリックすると該当のメールを表示します", foreground="gray").pack(pady=5)

    def view_headers(self):
        """メールのヘッダー情報を別ウィンドウで表示"""
        selected = self.msg_list.selection()
        if not selected: return
        
        cursor = self.conn.cursor()
        cursor.execute("SELECT headers FROM messages WHERE id = ?", (selected[0],))
        headers = cursor.fetchone()[0]
        
        header_win = self._setup_sub_window("HeaderWindow", "メールヘッダー", 600, 400)
        txt = tk.Text(header_win, font=("Consolas", 9), padx=10, pady=10)
        txt.pack(fill=tk.BOTH, expand=True)
        txt.insert("1.0", headers or "ヘッダー情報がありません（古いメール、または送信済みメール）")
        txt.config(state=tk.DISABLED)

    def view_in_browser(self):
        """現在のメールをデフォルトブラウザで開く"""
        selected = self.msg_list.selection()
        if not selected: return
        
        cursor = self.conn.cursor()
        cursor.execute("SELECT body, body_html, subject FROM messages WHERE id = ?", (selected[0],))
        body, body_html, subject = cursor.fetchone()
        
        # 文字化け防止のためメタタグを追加し、タイトルをセット
        head = f"<head><meta charset='utf-8'><title>{html.escape(subject or 'No Subject')}</title></head>"
        if body_html:
            content = f"<html>{head}<body>{body_html}</body></html>"
        else:
            content = f"<html>{head}<body><pre style='white-space: pre-wrap;'>{html.escape(body or '')}</pre></body></html>"
            
        # 一時ファイルを作成してブラウザで開く
        with tempfile.NamedTemporaryFile('w', delete=False, suffix='.html', encoding='utf-8') as f:
            f.write(content)
            temp_path = os.path.abspath(f.name)
        
        webbrowser.open(f"file://{temp_path}")

    def open_manual(self):
        """マニュアルウィンドウを表示"""
        win = self._setup_sub_window("ManualWindow", "Petal マニュアル", 600, 450)
        txt = tk.Text(win, font=("MS UI Gothic", 10), padx=15, pady=15, wrap=tk.WORD)
        txt.pack(fill=tk.BOTH, expand=True)
        
        content = """【Petal マニュアル】

■ アプリの概要
Petalは、AL-Mailの使い勝手を踏襲しつつ、現代的なインターフェースと機能を備えたマルチアカウント対応のメールクライアントです。

■ 主な機能
1. マルチアカウント管理
   設定画面から複数のメールアカウントを追加・管理できます。
2. 統合受信箱
   すべてのアカウントの受信メールを一覧で確認できます。
3. 添付ファイル管理
   メール内の画像を直接プレビューしたり、すべての添付ファイルを一括管理したりできます。
4. AL-Mailインポート
   既存のAL-Mailのメールボックスからデータを簡単に移行できます。
5. セキュリティチェック
   SPF/DKIM/DMARCなどの認証結果を分かりやすく表示し、なりすましを警告します。

■ 基本的な使い方
・メール受信: ツールバーの「受信」ボタン、またはメニューの「ファイル」→「メール受信」を選択します。
・新規作成: ツールバーの「新規」ボタンを選択します。差出人アカウントはコンボボックスで切り替え可能です。
・返信/転送: メール詳細画面の上部にあるボタンを使用します。
・削除: 不要なメールを選択して「削除」ボタンを押すと、ゴミ箱へ移動します。
・アドレス帳: 頻繁に送信する相手を登録して、新規作成時に呼び出すことができます。
"""
        txt.insert("1.0", content)
        txt.config(state=tk.DISABLED)

    def check_for_updates(self):
        """GitHub上のversion.jsonを確認してアップデートを行う"""
        def _check():
            import ssl
            try:
                # GitHub APIやRawファイルへのアクセスにはUser-Agentの設定が推奨される
                req = urllib.request.Request(self.UPDATE_URL, headers={'User-Agent': 'Petal-Update-Checker'})
                # Windows環境での証明書エラーを回避するためのコンテキスト作成
                context = ssl._create_unverified_context()
                with urllib.request.urlopen(req, context=context) as response:
                    data = json.loads(response.read().decode('utf-8'))
                
                remote_version = data.get("version", "1.0.0")
                update_info = data.get("info", "新しいバージョンが利用可能です。")
                installer_name = data.get("filename", "Petal_Setup.exe")

                if remote_version > self.APP_VERSION:
                    if messagebox.askyesno("アップデート", f"新しいバージョン ({remote_version}) が見つかりました。\n\n【更新内容】\n{update_info}\n\n今すぐアップデートしますか？"):
                        self._perform_update(installer_name)
                else:
                    # メインスレッド以外でGUIを出す場合はafterを使うのが安全
                    self.root.after(0, lambda: messagebox.showinfo("アップデート", f"Petal は最新の状態です。\n現在のバージョン: {self.APP_VERSION}"))
            except urllib.error.HTTPError as e:
                self.root.after(0, lambda: messagebox.showerror("エラー", f"アップデート情報が見つかりません(404)。\nURLを確認してください。\n{self.UPDATE_URL}"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("エラー", f"アップデートの確認に失敗しました:\n{e}"))

        threading.Thread(target=_check, daemon=True).start()

    def _perform_update(self, filename):
        """インストーラーをダウンロードして実行し、アプリを終了する"""
        import ssl
        try:
            # 進捗表示用の小さなウィンドウ
            prog_win = tk.Toplevel(self.root)
            prog_win.title("ダウンロード中...")
            prog_win.geometry("300x100")
            self._center_window_on_parent(prog_win, self.root, 300, 100)
            tk.Label(prog_win, text="最新のインストーラーを取得しています...", pady=20).pack()
            prog_win.update()

            # ダウンロード先 (一時フォルダー)
            download_url = self.DOWNLOAD_BASE_URL + filename
            temp_dir = tempfile.gettempdir()
            dest_path = os.path.join(temp_dir, filename)

            # SSL証明書検証を回避してダウンロード
            context = ssl._create_unverified_context()
            opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=context))
            urllib.request.install_opener(opener)
            urllib.request.urlretrieve(download_url, dest_path)

            # インストーラーを起動
            if dest_path.endswith(".zip"):
                # ZIPの場合はフォルダを開く
                os.startfile(os.path.dirname(dest_path))
                messagebox.showinfo("更新", "ダウンロードが完了しました。ZIPを展開して更新してください。")
            else:
                # EXE (Inno Setup) の場合は実行
                subprocess.Popen([dest_path, "/SILENT"], shell=True)
                # インストーラーが起動する時間を考慮してアプリを終了
                self.root.after(1000, self.root.quit)

        except Exception as e:
            messagebox.showerror("エラー", f"ダウンロードまたは起動に失敗しました:\n{e}")
        finally:
            if 'prog_win' in locals(): prog_win.destroy()

    def open_about(self):
        """Aboutウィンドウを表示"""
        win = tk.Toplevel(self.root)
        win.title("About Petal")
        win.geometry("350x180")
        # 親ウィンドウの中央に表示
        geom = self._center_window_on_parent(win, self.root, 350, 180)
        win.geometry(geom)
        win.resizable(False, False)
        win.grab_set()
        
        tk.Label(win, text="Petal Email Client", font=("Arial", 16, "bold"), fg="#d81b60", pady=10).pack()
        tk.Label(win, text=f"Version {self.APP_VERSION} (Multi-Account Ready)", font=("Arial", 10)).pack()
        tk.Label(win, text="Contact / Developer Support:", font=("MS UI Gothic", 9, "bold"), pady=5).pack()
        
        email_lbl = tk.Label(win, text="t.muguruma.work@gmail.com", fg="#0052a3", cursor="hand2", font=("Consolas", 10, "underline"))
        email_lbl.pack()
        email_lbl.bind("<Button-1>", lambda e: webbrowser.open("mailto:t.muguruma.work@gmail.com"))
        
        ttk.Button(win, text="閉じる", command=win.destroy).pack(pady=15)

if __name__ == "__main__":
    root = tk.Tk()
    app = ModernALMail(root)
    root.mainloop()