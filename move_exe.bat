@echo off
setlocal
chcp 65001 >nul
set "SRC=W:\myProjects\almail\dist\Petal"
set "DEST=W:\myProjects\almail\Petal"

echo [配置 1/3] ビルド成果物の確認中...
if not exist "%SRC%" (
    echo [エラー] ソースフォルダが見つかりません: %SRC%
    pause
    exit /b
)

if not exist "%DEST%" mkdir "%DEST%"

echo [配置 2/3] ファイルを移動中...
echo --------------------------------------------------
rem /R:3 /W:2 を追加して、失敗時の再試行を3回(各2秒待機)に制限します
robocopy "%SRC%" "%DEST%" /E /IS /MOVE /TEE /R:3 /W:2
set "ROBO_RES=%ERRORLEVEL%"
echo --------------------------------------------------

echo [配置 3/3] 不要な空フォルダを削除中...
if exist "%SRC%" rd /s /q "%SRC%" >nul 2>&1

echo.
echo ==================================================
echo  完了しました。 mailbox.db を残したまま更新しました。
echo ==================================================
exit /b %ROBO_RES%
