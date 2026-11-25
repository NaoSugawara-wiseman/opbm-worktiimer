<#
  GitHub から取ってきた最新版を、使う PC のフォルダに配置する
    get_laptop.bat / get_desktop.bat から呼ばれる (直接実行しなくてよい)
    -Target laptop  : 作業タイマー (ノート PC 用)
    -Target desktop : 自動入力 (デスクトップ PC 用)
    -Dest           : 配置先のフォルダ (bat を置いたフォルダ)
  プログラム (.bat / .ps1) は上書きする。設定 (ini・tasks.csv) は、まだないときだけ見本から作る (自分の設定は消さない)
  ファイルはバイトのままコピーする (.bat の Shift_JIS、.ps1 の BOM 付き UTF-8、CRLF を保つ)
#>
param(
    [Parameter(Mandatory = $true)][ValidateSet('laptop', 'desktop')][string]$Target,
    [Parameter(Mandatory = $true)][string]$Dest
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$Dest = [System.IO.Path]::GetFullPath($Dest)
$version = (Get-Content -LiteralPath (Join-Path $root 'VERSION') -TotalCount 1).Trim()

if ($Target -eq 'laptop') {
    $label = 'ノート PC 用 (作業タイマー)'
    $programs = @('laptop\worktimer.bat', 'laptop\worktimer.ps1', 'common\common.ps1')
    $configs = @(@('config\worktimer.ini.sample', 'worktimer.ini', 'data_dir などを設定してください'), @('config\tasks.csv.sample', 'tasks.csv', 'タスク一覧を入れてください'))
} else {
    $label = 'デスクトップ PC 用 (自動入力)'
    $programs = @('desktop\obpm_input.bat', 'desktop\obpm_input.ps1', 'desktop\obpm_inspect.bat', 'desktop\obpm_inspect.ps1', 'common\common.ps1')
    $configs = @(, @('config\obpm_input.ini.sample', 'obpm_input.ini', 'data_dir などを設定してください'))
}

Write-Host ''
Write-Host "$label  v$version を配置します"
Write-Host "  配置先: $Dest"
Write-Host ''
if (-not (Test-Path -LiteralPath $Dest)) { New-Item -ItemType Directory -Path $Dest -Force | Out-Null }

function Test-SameFile([string]$A, [string]$B) {
    if (-not (Test-Path -LiteralPath $B)) { return $false }
    $x = [System.IO.File]::ReadAllBytes($A); $y = [System.IO.File]::ReadAllBytes($B)
    if ($x.Length -ne $y.Length) { return $false }
    for ($i = 0; $i -lt $x.Length; $i++) { if ($x[$i] -ne $y[$i]) { return $false } }
    return $true
}

# 改行が CRLF でないと .bat が正しく動かないので、配置する前に確かめる
foreach ($p in $programs) {
    $bytes = [System.IO.File]::ReadAllBytes((Join-Path $root $p))
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        if ($bytes[$i] -eq 10 -and ($i -eq 0 -or $bytes[$i - 1] -ne 13)) {
            throw "取得したファイルの改行が CRLF になっていません: $p (取得用のフォルダを削除して、もう一度実行してください: $root)"
        }
    }
}

$changed = 0
foreach ($p in $programs) {
    $src = Join-Path $root $p
    $name = Split-Path -Leaf $p
    $dst = Join-Path $Dest $name
    if (Test-SameFile $src $dst) { Write-Host "  変更なし  $name"; continue }
    $isNew = -not (Test-Path -LiteralPath $dst)
    Copy-Item -LiteralPath $src -Destination $dst -Force
    $changed++
    if ($isNew) { Write-Host "  追加      $name" } else { Write-Host "  更新      $name" }
}
foreach ($c in $configs) {
    $dst = Join-Path $Dest $c[1]
    if (Test-Path -LiteralPath $dst) { Write-Host "  そのまま  $($c[1]) (今の設定を使います)"; continue }
    Copy-Item -LiteralPath (Join-Path $root $c[0]) -Destination $dst
    Write-Host "  作成      $($c[1]) (見本から作りました。$($c[2]))"
}
$readme = Join-Path $root 'docs\README_利用者向け.md'
if (Test-Path -LiteralPath $readme) { Copy-Item -LiteralPath $readme -Destination (Join-Path $Dest 'README.md') -Force }

Write-Host ''
if ($changed -eq 0) { Write-Host 'すでに最新版です。' }
else {
    Write-Host "$changed 個のファイルを更新しました。"
    if ($Target -eq 'laptop') {
        Write-Host '※ 作業タイマーが起動中なら、いったん終了して起動し直すと新しい版になります。'
    } else {
        # 待機モードが動いていれば知らせる (古い版のまま動き続けるため)
        $running = @()
        try { $running = @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object { $_.CommandLine -like '*obpm_input.ps1*-Watch*' }) } catch { }
        if ($running.Count -gt 0) {
            Write-Host '※ 待機モードが古い版のまま動いています。PC を再起動するか、obpm_input.bat の 3 で起動し直してください。' -ForegroundColor Yellow
            Write-Host '  (3 を選ぶ前に、タスクマネージャーで「Windows PowerShell」を終了してください)'
        }
    }
}
