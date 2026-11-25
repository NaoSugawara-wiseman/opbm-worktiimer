@echo off
setlocal
rem ==========================================================================
rem  作業タイマー (ノート PC 用) の最新版を GitHub から取得して、この bat と同じフォルダに配置する
rem    使い方: この bat をノート PC のツールのフォルダ (例: C:\tools\worktimer) に置いて実行する
rem    必要なもの: Git for Windows、このリポジトリを見られる GitHub アカウント
rem    プログラム (.bat / .ps1) は上書きする。設定 (ini と tasks.csv) は、まだないときだけ見本から作る
rem ==========================================================================
set "REPO=https://github.com/NaoSugawara-wiseman/opbm-worktiimer.git"
set "CACHE=%LOCALAPPDATA%\opbm-worktiimer\repo"

where git >nul 2>nul
if errorlevel 1 (
    echo Git が見つかりません。Git for Windows をインストールしてから、もう一度実行してください。
    echo   https://gitforwindows.org/
    pause
    exit /b 1
)

echo GitHub から最新版を取得しています...
if exist "%CACHE%\.git" (
    git -C "%CACHE%" fetch --quiet origin main
    if errorlevel 1 goto :fail
    git -C "%CACHE%" reset --quiet --hard origin/main
    if errorlevel 1 goto :fail
) else (
    if exist "%CACHE%" rmdir /s /q "%CACHE%"
    git clone --quiet -c core.autocrlf=false "%REPO%" "%CACHE%"
    if errorlevel 1 goto :fail
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%CACHE%\tools\fetch.ps1" -Target laptop -Dest "%~dp0."
if errorlevel 1 (
    echo.
    echo 配置に失敗しました。上のメッセージを確認してください。
)
echo.
pause
exit /b 0

:fail
echo.
echo 取得に失敗しました。次を確認してください。
echo   - ネットワークにつながっているか
echo   - GitHub にサインインしているか (初回は画面が出るのでサインインする)
echo   - このリポジトリを見る権限があるか
pause
exit /b 1
