<#
  配布用の zip を作る
    使い方 (リポジトリのフォルダで):  powershell -ExecutionPolicy Bypass -File tools\build.ps1
    できるもの: dist\worktimer_powershell_v<VERSION>.zip
      worktimer_powershell\
        README.md
        ノートPC用\        worktimer.bat / worktimer.ps1 / common.ps1 / worktimer.ini / tasks.csv
        デスクトップPC用\  obpm_input.bat / obpm_input.ps1 / common.ps1 / obpm_input.ini / obpm_inspect.bat / obpm_inspect.ps1
  ファイルはバイトのままコピーする (.bat の Shift_JIS、.ps1 の BOM 付き UTF-8、CRLF を保つ)
#>
param([string]$Version = '')

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
if (-not $Version) { $Version = (Get-Content -LiteralPath (Join-Path $root 'VERSION') -TotalCount 1).Trim() }

$dist = Join-Path $root 'dist'
$stage = Join-Path $dist 'worktimer_powershell'
if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
$note = Join-Path $stage 'ノートPC用'
$desk = Join-Path $stage 'デスクトップPC用'
New-Item -ItemType Directory -Path $note, $desk -Force | Out-Null

function Copy-To([string]$From, [string]$To) {
    $src = Join-Path $root $From
    if (-not (Test-Path -LiteralPath $src)) { throw "ファイルがありません: $From" }
    Copy-Item -LiteralPath $src -Destination $To -Force
}

# ノート PC 用
Copy-To 'laptop\worktimer.bat' $note
Copy-To 'laptop\worktimer.ps1' $note
Copy-To 'common\common.ps1' $note
Copy-To 'config\worktimer.ini.sample' (Join-Path $note 'worktimer.ini')
Copy-To 'config\tasks.csv.sample' (Join-Path $note 'tasks.csv')

# デスクトップ PC 用
foreach ($f in 'obpm_input.bat', 'obpm_input.ps1', 'obpm_inspect.bat', 'obpm_inspect.ps1') { Copy-To "desktop\$f" $desk }
Copy-To 'common\common.ps1' $desk
Copy-To 'config\obpm_input.ini.sample' (Join-Path $desk 'obpm_input.ini')

# 説明書
Copy-To 'docs\README_利用者向け.md' (Join-Path $stage 'README.md')

# 改行が CRLF になっているか確かめる (LF のままだと .bat が正しく動かない)
foreach ($f in Get-ChildItem -LiteralPath $stage -Recurse -File) {
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        if ($bytes[$i] -eq 10 -and ($i -eq 0 -or $bytes[$i - 1] -ne 13)) {
            throw "改行が CRLF になっていません: $($f.Name) (git の設定 core.autocrlf と .gitattributes を確認)"
        }
    }
}

$zip = Join-Path $dist "worktimer_powershell_v$Version.zip"
if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
Compress-Archive -Path $stage -DestinationPath $zip
Write-Host "作成しました: $zip"
