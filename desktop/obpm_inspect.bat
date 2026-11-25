@echo off
rem ================================================================
rem  OBPM 画面の調査 (読み取り専用)
rem    obpm_inspect.bat と obpm_inspect.ps1 を同じフォルダに置く
rem    次の 3 つの状態で 1 回ずつ実行する
rem      1. ログイン画面を表示した状態
rem      2. ログイン後のメインメニューを表示した状態
rem      3. メインメニューから工数入力画面を開いた状態
rem    結果は同じフォルダの obpm_inspect_日時.txt に出る
rem ================================================================
if not exist "%~dp0obpm_inspect.ps1" (
    echo [エラー] obpm_inspect.ps1 が見つかりません: %~dp0
    pause
    exit /b 1
)
echo OBPM の画面を調べています。終わるまでそのままお待ちください...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0obpm_inspect.ps1"
echo.
pause
