@echo off
setlocal
rem ================================================================
rem  OBPM 工数自動入力 (デスクトップ側・PowerShell 版)
rem    obpm_input.bat / obpm_input.ps1 / common.ps1 / obpm_input.ini を同じフォルダに置く
rem    実行する前に OBPM の「工数入力」画面を開いておくこと (自動ログインを登録していれば不要)
rem ================================================================
pushd "%~dp0"
if not exist "%~dp0obpm_input.ps1" (
    echo [エラー] obpm_input.ps1 が見つかりません: %~dp0
    goto :end
)
set "PS=powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0obpm_input.ps1""

echo.
echo   1. 今すぐ入力する (ini の mode に従う)
echo   2. タスク一覧の候補を作る (読み取りのみ。1～2 分かかる)
echo   3. 待機モードを今すぐ起動する (本日終了の依頼で自動入力)
echo   4. 待機モードを PC 起動時に自動で起動するよう登録する
echo   5. 4 の登録を解除する
echo   6. 自動ログインの情報 (アカウント・パスワード) を登録する
echo   7. 自動ログインの情報を削除する
echo   8. 何もしないで終わる
echo.
choice /c 12345678 /n /m "番号を選んでください: "
set "NO=%errorlevel%"
echo.
if "%NO%"=="1" %PS%
if "%NO%"=="2" %PS% -MakeTasks
if "%NO%"=="3" (
    start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0obpm_input.ps1" -Watch
    echo 待機モードを起動しました。動作状況は共有フォルダの obpm_input_log.txt で確認できます。
)
if "%NO%"=="4" %PS% -InstallStartup
if "%NO%"=="5" %PS% -UninstallStartup
if "%NO%"=="6" %PS% -SetLogin
if "%NO%"=="7" %PS% -ClearLogin

:end
popd
echo.
pause
