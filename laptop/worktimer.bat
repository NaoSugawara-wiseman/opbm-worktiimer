@echo off
rem ================================================================
rem  作業タイマー 起動バッチ (PowerShell 版)
rem    ・ダブルクリックで起動 (黒い画面は出ず、タイマー画面だけが開く)
rem    ・PC 起動時に自動で起動するには、このバッチのショートカットを
rem      スタートアップフォルダ (Win+R → shell:startup) に入れる
rem    ・worktimer.bat / worktimer.ps1 / common.ps1 / worktimer.ini / tasks.csv を同じフォルダに置く
rem ================================================================
if not exist "%~dp0worktimer.ps1" (
    echo [エラー] worktimer.ps1 が見つかりません: %~dp0
    pause
    exit /b 1
)
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0worktimer.ps1"
exit /b 0
