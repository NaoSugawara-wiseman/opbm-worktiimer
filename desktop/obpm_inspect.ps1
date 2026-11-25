# -*- coding: utf-8 -*-
<#
obpm_inspect.ps1  OBPM の画面を、プログラムからどう見えるか調べる (読み取り専用・PowerShell 版)

調べること (確認事項 B-7〜9)
  B-7 メインメニューの「工数入力」リンクを、プログラムから見つけて押せるか
  B-8 ログイン画面を見分けられるか
  B-9 工数入力画面は別ウィンドウか、メインメニューと同じウィンドウ内で切り替わるか

使い方
  obpm_inspect.bat をダブルクリック (OBPM の画面を表示した状態で)
  ログイン画面 / メインメニュー / 工数入力画面 を開いた状態で 1 回ずつ実行する
  結果は同じフォルダの obpm_inspect_日時.txt に出る (実行するたびに別ファイル)

安全性
  ・画面の部品の種類・名前・位置を読むだけ。クリック・入力・更新は一切しない
  ・入力欄の中身 (ログインアカウント・パスワードなど) は読まない
  ・結果にはプロジェクト名などが含まれることがある。他の人に渡す前に内容を確認すること
※ このファイルは「UTF-8 (BOM 付き)」で保存すること
#>
param(
    [string]$Keyword = 'SI Object Browser',   # OBPM のウィンドウのタイトルに含まれる文字
    [int]$MaxNodes = 3000,                    # 読み取る部品数の上限 (1 ウィンドウあたり)
    [int]$MaxDepth = 15
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$Walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker

$out = New-Object System.Collections.Generic.List[string]
function W([string]$S) { $out.Add($S) }

$stamp = Get-Date
$outPath = Join-Path $PSScriptRoot ('obpm_inspect_{0}.txt' -f $stamp.ToString('yyyyMMdd_HHmmss'))

function Save-Result {
    [System.IO.File]::WriteAllText($outPath, (($out -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
    Write-Host "結果を書き出しました: $outPath"
    Write-Host '※ プロジェクト名などが含まれることがあります。他の人に渡す前に内容を確認してください。'
}

function Get-Rect($El) {
    try { $r = $El.Current.BoundingRectangle; return '({0},{1})-({2},{3})' -f [int]$r.Left, [int]$r.Top, [int]$r.Right, [int]$r.Bottom } catch { return '(位置不明)' }
}

function Get-Patterns($El) {
    try {
        $names = foreach ($p in $El.GetSupportedPatterns()) { ($p.ProgrammaticName -replace 'PatternIdentifiers\.Pattern$', '' -replace 'Identifiers\.Pattern$', '') }
        return (@($names) -join ',')
    } catch { return '' }
}

function Get-TypeName($El) {
    try { return ($El.Current.ControlType.ProgrammaticName -replace '^ControlType\.', '') } catch { return '?' }
}

W "obpm_inspect 結果  $($stamp.ToString('yyyy-MM-dd HH:mm:ss'))"
W '(読み取り専用。入力欄の中身は読んでいません)'
W ('=' * 78)

# --------------------------------------------------------------------------
# OBPM のウィンドウを探す (タイトルで見つけ、同じプロセスのウィンドウをすべて対象にする)
# --------------------------------------------------------------------------
$tops = @($AE::RootElement.FindAll($TS::Children, [System.Windows.Automation.Condition]::TrueCondition))
$pids = @{}
foreach ($w in $tops) {
    try {
        $n = $w.Current.Name
        if ($n -like "*$Keyword*" -or $n -like '*工数入力*') { $pids[$w.Current.ProcessId] = $true }
    } catch { }
}
if ($pids.Count -eq 0) {
    W "タイトルに「$Keyword」を含むウィンドウが見つかりませんでした。OBPM を起動してから実行してください。"
    W ''
    W '開いているウィンドウのタイトル一覧:'
    foreach ($w in $tops) { try { if ($w.Current.Name) { W "  - $($w.Current.Name)" } } catch { } }
    Save-Result
    exit 0
}

$targets = @(foreach ($w in $tops) { try { if ($pids.ContainsKey($w.Current.ProcessId)) { $w } } catch { } })

# --------------------------------------------------------------------------
# [1] ウィンドウの一覧 (B-9)
# --------------------------------------------------------------------------
W ''
W '[1] OBPM のウィンドウ一覧 (B-9: 工数入力画面が別ウィンドウかどうか)'
W ('-' * 78)
foreach ($w in $targets) {
    $pname = ''
    try { $pname = (Get-Process -Id $w.Current.ProcessId -ErrorAction Stop).ProcessName } catch { }
    W ("  タイトル='{0}'  クラス={1}  プロセス={2} (ID {3})  位置={4}" -f $w.Current.Name, $w.Current.ClassName, $pname, $w.Current.ProcessId, (Get-Rect $w))
}
$hasKosu = @($targets | Where-Object { $_.Current.Name -like '*工数入力*' }).Count -gt 0
$hasOther = @($targets | Where-Object { $_.Current.Name -notlike '*工数入力*' }).Count -gt 0
if ($hasKosu -and $hasOther) { W '  → 工数入力画面とそれ以外の画面が、別々のウィンドウとして開いています' }
elseif ($hasKosu) { W '  → 開いているのは工数入力画面のウィンドウだけです (メインメニューと同じウィンドウ内で切り替わっている可能性)' }
else { W '  → 工数入力画面のウィンドウはありません (ログイン画面またはメインメニューの状態)' }

# --------------------------------------------------------------------------
# [2] 部品の一覧
# --------------------------------------------------------------------------
$script:Hits = New-Object System.Collections.Generic.List[string]       # 名前に「工数入力」を含む部品
$script:LoginHints = New-Object System.Collections.Generic.List[string] # ログイン画面らしさ
$script:Count = 0

function Read-Tree($El, [int]$Depth, [string]$WinTitle) {
    if ($script:Count -ge $MaxNodes -or $Depth -gt $MaxDepth) { return }
    $child = $null
    try { $child = $Walker.GetFirstChild($El) } catch { return }
    while ($null -ne $child) {
        if ($script:Count -ge $MaxNodes) { return }
        $script:Count++
        try {
            $c = $child.Current
            $type = Get-TypeName $child
            $name = [string]$c.Name
            if ($name.Length -gt 60) { $name = $name.Substring(0, 60) + '…' }
            $pat = Get-Patterns $child
            $pw = ''
            try { if ($c.IsPassword) { $pw = '  [パスワード欄]' } } catch { }
            W ("{0}- {1,-10} class={2} name='{3}' id='{4}' {5} 操作={6}{7}" -f ('  ' * $Depth), $type, $c.ClassName, $name, $c.AutomationId, (Get-Rect $child), $pat, $pw)
            if ($name -like '*工数入力*') {
                $script:Hits.Add(("画面='{0}'  種類={1}  name='{2}'  id='{3}'  位置={4}  操作={5}" -f $WinTitle, $type, $name, $c.AutomationId, (Get-Rect $child), $pat))
            }
            if ($pw) { $script:LoginHints.Add("パスワード欄があります (画面='$WinTitle', id='$($c.AutomationId)')") }
            if ($name -like '*ログインアカウント*' -or $name -eq 'パスワード' -or $name -eq 'ログイン') {
                $script:LoginHints.Add("「$name」という $type があります (画面='$WinTitle', id='$($c.AutomationId)')")
            }
        } catch {
            W ("{0}- (読み取り失敗: {1})" -f ('  ' * $Depth), $_.Exception.Message)
        }
        Read-Tree $child ($Depth + 1) $WinTitle
        try { $child = $Walker.GetNextSibling($child) } catch { $child = $null }
    }
}

foreach ($w in $targets) {
    W ''
    W "[2] 部品の一覧: '$($w.Current.Name)'  (種類 / クラス / 名前 / ID / 位置 / プログラムからできる操作)"
    W ('-' * 78)
    $script:Count = 0
    $t0 = Get-Date
    Read-Tree $w 0 $w.Current.Name
    $note = ''
    if ($script:Count -ge $MaxNodes) { $note = '  ※上限に達したため途中で打ち切り' }
    W ("  読み取った部品数: {0} ({1:0.0} 秒){2}" -f $script:Count, ((Get-Date) - $t0).TotalSeconds, $note)
}

# --------------------------------------------------------------------------
# [3] 判定のヒント
# --------------------------------------------------------------------------
W ''
W '[3] 判定のヒント'
W ('-' * 78)
W 'B-7 名前に「工数入力」を含む部品:'
if ($script:Hits.Count -eq 0) { W '  見つかりませんでした (メインメニューを表示した状態で実行すると、リンクが見つかるはず)' }
foreach ($h in $script:Hits) { W "  $h" }
W '  ※ 操作に Invoke または LegacyIAccessible があれば、プログラムから押せる可能性が高い。'
W '     どちらもなければ、位置をクリックする方法になる。'
W ''
W 'B-8 ログイン画面らしさ:'
if ($script:LoginHints.Count -eq 0) { W '  ログイン画面の特徴は見つかりませんでした' }
foreach ($h in $script:LoginHints) { W "  $h" }
W ''
W 'B-9 [1] のウィンドウ一覧を参照'

Save-Result
