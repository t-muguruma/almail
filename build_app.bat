@echo off
setlocal
chcp 65001 >nul

echo ==================================================
echo  Petal 一括ビルド＆デプロイ
echo ==================================================

rem スクリプトのあるディレクトリに移動
cd /d "%~dp0"

echo [準備] 現在のバージョン情報を取得中...
set "CUR_VER=0.0.0"
for /f "tokens=2 delims==" %%I in ('findstr /C:"APP_VERSION =" modern_almail.py') do set "RAW_V=%%I"
if defined RAW_V (
    for /f "tokens=1 delims=#" %%A in ("%RAW_V%") do set "RAW_V=%%A"
    set "RAW_V=%RAW_V: =%"
    set "CUR_VER=%RAW_V:"=%"
)

echo [準備] 現在のバージョン: %CUR_VER%
set /p APP_VER="新しいバージョンを入力 (そのままEnterで %CUR_VER%): "
if "%APP_VER%"=="" set "APP_VER=%CUR_VER%"

echo [準備] 更新内容を入力してください
set /p APP_INFO="Update Info: "

echo [準備] modern_almail.py の APP_VERSION を更新中...

set "PYTHON_EXE="
rem python.exe の場所を特定（1つ目に見つかったものを採用）
for /f "tokens=*" %%i in ('where python 2^>nul') do if not defined PYTHON_EXE set "PYTHON_EXE=%%i"

if not defined PYTHON_EXE (
    echo [エラー] Python が見つかりません。PATHの設定を確認してください。
    pause
    exit /b 1
)

echo [情報] 使用する Python: "%PYTHON_EXE%"

rem modern_almail.py のバージョンを更新（常に UTF-8 で上書き）
"%PYTHON_EXE%" -c "import re, io; p='modern_almail.py'; data=io.open(p, 'r', encoding='utf-8').read(); data=re.sub(r'APP_VERSION = \".*\"', 'APP_VERSION = \"%APP_VER%\"', data); io.open(p, 'w', encoding='utf-8', newline='\n').write(data)"

echo [準備] version.json を生成中...
rem version.json も Python を使って UTF-8 で確実に出力
"%PYTHON_EXE%" -c "import json, sys, io; d={'version': sys.argv[1], 'info': sys.argv[2], 'filename': 'Petal_Setup.exe'}; io.open('version.json', 'w', encoding='utf-8').write(json.dumps(d, ensure_ascii=False, indent=2))" "%APP_VER%" "%APP_INFO%"

echo [準備] 実行環境の確認中...
if not exist "modern_almail.spec" (
    echo [エラー] modern_almail.spec が見つかりません。
    echo 実行ディレクトリ: %CD%
    pause
    exit /b 1
)

echo [準備] 実行中の Petal を終了させています...
rem ロック確認の前にまず終了を試みる
taskkill /F /IM Petal.exe /T >nul 2>&1
timeout /t 1 >nul

echo [準備] 書き込み権限（ファイルロック）の確認中...
if exist "W:\myProjects\almail\Petal\Petal.exe" (
    ren "W:\myProjects\almail\Petal\Petal.exe" "Petal.exe.test" >nul 2>&1
    if errorlevel 1 goto LOCK_ERROR
    ren "W:\myProjects\almail\Petal\Petal.exe.test" "Petal.exe" >nul 2>&1
)

rem 依存ライブラリのチェック（出力を一時ファイルに書き込み、失敗時のみ表示）
"%PYTHON_EXE%" -c "import PIL, PyInstaller; print('OK')" > python_check.tmp 2>&1
if errorlevel 1 (
    echo [エラー] 必要なライブラリが不足しています。
    type python_check.tmp
    del python_check.tmp
    pause
    exit /b 1
)
del python_check.tmp

if exist python_import_check.log del python_import_check.log

echo [準備] アイコンを ICO 形式に強制変換中...
if exist "Petal_icon.ico" del "Petal_icon.ico"
if not exist "Petal_icon.png" (
    echo [エラー] アイコンファイル 'Petal_icon.png' が見つかりません。
    echo スクリプトと同じディレクトリに配置してください。
    pause
    exit /b 1
)
"%PYTHON_EXE%" -c "from PIL import Image; img = Image.open('Petal_icon.png'); img.save('Petal_icon.ico', format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])"
if errorlevel 1 goto PILLOW_ERROR

echo [進捗 1/3] PyInstallerでビルド中...
rem --clean を付けてキャッシュをクリアし、確実に最新の状態をビルドします
pyinstaller --clean modern_almail.spec
if errorlevel 1 goto BUILD_ERROR

echo [準備] アンチウイルスソフトのロック解除を数秒待ちます...
timeout /t 3 >nul

echo.
echo [進捗 2/3] move_exe.bat を実行してファイルを配置中...
rem call を使うことで、move_exe.bat 終了後にこのバッチに戻ってこれます
call move_exe.bat

rem robocopy の戻り値は 8 未満であれば成功（または警告レベル）とみなせます
if %ERRORLEVEL% GEQ 8 goto MOVE_ERROR

echo.
echo [進捗 3/3] インストーラーを作成中...
rem Inno Setup がインストールされている標準的なパスを指定します
set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" goto INNO_MISSING

"%ISCC%" petal_installer.iss
if errorlevel 1 goto INNO_ERROR

goto SUCCESS

:INNO_MISSING
echo [警告] Inno Setup (ISCC.exe) が見つからないためインストーラー作成をスキップします。
pause
goto SUCCESS

:SUCCESS
echo.
echo ==================================================
echo  すべての工程が正常に完了しました
echo ==================================================
pause
exit /b 0

:LOCK_ERROR
echo.
echo [致命的エラー] Petal.exe がロックされています
echo 別のPCで起動していないか、エクスプローラーで開いていないか確認してください。
pause
exit /b 1

:PILLOW_ERROR
echo [エラー] ICO ファイルの生成に失敗しました。
echo 解決策: 'Petal_icon.png' が存在するか、Pillow が正しくインストールされているか ('pip install Pillow') 確認してください。
pause
exit /b 1

:BUILD_ERROR
echo.
echo [エラー] PyInstallerでのビルドに失敗しました。
pause
exit /b 1

:MOVE_ERROR
echo.
echo [エラー] move_exe.bat の実行中に問題が発生しました。
echo ログを確認してください。
pause
exit /b 1

:INNO_ERROR
echo [エラー] インストーラーの作成に失敗しました。
pause
exit /b 1