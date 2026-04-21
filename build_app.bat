@echo off
setlocal
chcp 65001 >nul

echo ==================================================
echo  Petal 一括ビルド＆デプロイ
echo ==================================================

echo [準備] 書き込み権限（ファイルロック）の確認中...
rem 別のPCやバックグラウンドで起動していると、ここでリネームに失敗します
if exist "W:\myProjects\almail\Petal\Petal.exe" (
    ren "W:\myProjects\almail\Petal\Petal.exe" "Petal.exe.test" >nul 2>&1
    if errorlevel 1 goto LOCK_ERROR
    ren "W:\myProjects\almail\Petal\Petal.exe.test" "Petal.exe" >nul 2>&1
)

echo [準備] 実行中の Petal を終了させています...
rem 実行中でない場合はエラーになりますが、>nul 2>&1 で無視します
taskkill /F /IM Petal.exe /T >nul 2>&1
timeout /t 1 >nul

echo [準備] アイコンを ICO 形式に強制変換中...
if exist "Petal_icon.ico" del "Petal_icon.ico"
python -c "from PIL import Image; img = Image.open('Petal_icon.png'); img.save('Petal_icon.ico', format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])"
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
echo [エラー] Pillow が未インストールのため、ICO ファイルの生成に失敗しました。
echo 解決策: 'pip install Pillow' を実行してください。
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