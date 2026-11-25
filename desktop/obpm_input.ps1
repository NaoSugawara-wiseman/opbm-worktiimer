# -*- coding: utf-8 -*-
<#
obpm_input.ps1  作業タイマーの記録を SI Object Browser PM (OBPM) の工数入力画面に自動で入力する
(デスクトップ PC 側。Windows PowerShell 5.1 で動作)

動き方
  ・共有フォルダの records.csv (ノートの作業タイマーの記録) を読み、日付×プロジェクト×WBS ごとの工数 (丸め後) を計算
  ・前回入力した値 (obpm_input_state.json) と違うものだけを入力 (その日のそのタスクの合計で上書き)
  ・プロジェクトごとに「プロジェクトコードの「-」より前 (6 桁) で検索 → 目的のプロジェクト行を探す → 行と日付の列を特定 → 入力 → 読み返して確認」
  ・mode = live のときは、全件の確認が取れたプロジェクトだけ「更新」を押す
  ・画面がロックされているときは何もせず、ロック解除後に入力する

起動方法 (通常は bat から)
  obpm_input.ps1                 今すぐ 1 回実行 (ini の mode に従う)
  obpm_input.ps1 -Watch          待機モード: 指定時刻と「本日終了」の依頼で自動実行 (PC 起動時に起動しておく)
  obpm_input.ps1 -MakeTasks      全プロジェクトの WBS を読み、tasks.csv の候補を作る (読み取りのみ)
  obpm_input.ps1 -InstallStartup 待機モードをスタートアップに登録 / -UninstallStartup で解除
  ※ -Config 別名.ini で設定ファイルを指定できる
※ このファイルは「UTF-8 (BOM 付き)」で保存すること
#>
param(
    [string]$Config = '',
    [switch]$Watch,
    [switch]$MakeTasks,
    [switch]$InstallStartup,
    [switch]$UninstallStartup,
    [switch]$SetLogin,
    [switch]$ClearLogin
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'common.ps1')

$script:ProjectRow = '^\d{6}-\d{2}[\u3000 ]'
$script:OffsetCols = 5        # WBS 列から 1 日の列までの間にある列数 (計画工数・予定 2 列・実績 2 列)
$script:SameRunLimit = 6      # 同じ名前がこの回数続いたら表の終わりとみなす
$script:MaxRows = 1000
$script:SyncRange = 4         # 行の位置がずれたときに前後何行まで探して補正するか

$script:ConfigTemplate = @'
; obpm_input 設定ファイル (デスクトップ側。; で始まる行はコメント。パスは " で囲まずそのまま書く)

[files]
; 共有フォルダ上の作業タイマーの data フォルダ (ノートの worktimer.ini の data_dir と同じ場所)
data_dir = \\fileserver\share\worktimer\data

[run]
; check  = 読み取りのみ。入力先のセルを探して現在の値を読み、入力予定と並べてログに出す (最初はこれで確認)
; manual = 入力して読み返すまで。更新は押さずに止まるので、画面を見て自分で「更新」を押す (1 プロジェクトずつ)
; live   = 入力・読み返し・更新まで自動
mode = check
; 工数の丸め (ノートの worktimer.ini と同じにする)
round_unit_min = 15
round_mode = nearest

[daily_total]
; 終わった日 (前日以前・「本日終了」を押した日) の合計を社内ルールに合わせる (true / false)
;   合計 = 定時 + 残業 (残業は下の単位未満を切り捨て)。ずれは一番長いタスクで調整。定時未満の日は調整しない
adjust = true
; 定時の工数 (時間)
base_hours = 7.75
; 残業の計上単位 (時間)
overtime_step_hours = 0.5

[schedule]
; 自動で入力する時刻 (カンマ区切り。平日のみ)。空なら「本日終了」の依頼のときだけ入力する
times =
; 「本日終了」の依頼があるかを確認する間隔 (分)
request_check_min = 5

[obpm]
; 工数入力画面のタイトルに含まれる文字
title_keyword = 工数入力
; OBPM の実行ファイルのパス (任意)。空ならスタートメニューのショートカットなどから自動で探す
;   (スタートメニューの OBPM のアイコンを右クリック →「ファイルの場所を開く」で確認)
obpm_exe =
; 入力欄・ボタンの操作方法: mouse = マウスでクリック / message = マウスを使わない (リモートデスクトップを最小化して使うとき)
input_method = mouse
; セルに数値を打った後に押す確定キー ({ENTER} / {TAB} など)
commit_key = {ENTER}
; 確定キーを押した後に選択セルが下へ動く場合は down、動かない場合は none、右へ動く場合は right
after_commit = none
; キー操作の間隔 (秒)。入力が不安定なら大きくする
key_wait = 0.15
; 検索ボタンを押してから表が更新されるまで待つ秒数
search_wait = 2.0
'@

class AbortException : System.Exception {
    AbortException([string]$Message) : base($Message) { }
}
function Stop-Run([string]$Msg) { throw [AbortException]::new($Msg) }
function Get-AbortMessage($Ex) {
    # ドライバーのメソッド内で投げた例外は PowerShell に包まれるので、中まで見て探す
    while ($null -ne $Ex) {
        if ($Ex -is [AbortException]) { return $Ex.Message }
        $Ex = $Ex.InnerException
    }
    return $null
}

# ==========================================================================
# 設定・ログ・状態
# ==========================================================================
function Get-InputConfig([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-TextFile $Path ($script:ConfigTemplate -replace "`r?`n", "`r`n")
        throw "設定ファイルがないため雛形を作成しました: $Path`n編集してから再実行してください。"
    }
    $ini = Read-Ini $Path
    $base = Split-Path -Parent (Resolve-Path -LiteralPath $Path)
    $cfg = @{
        DataDir         = Resolve-RelPath $base (Get-IniValue $ini 'files' 'data_dir' 'data')
        Mode            = (Get-IniValue $ini 'run' 'mode' 'check').ToLower()
        RoundUnit       = [Math]::Max(1, [int](Get-IniValue $ini 'run' 'round_unit_min' '15'))
        RoundMode       = (Get-IniValue $ini 'run' 'round_mode' 'nearest').ToLower()
        Daily           = @{
            Adjust    = ((Get-IniValue $ini 'daily_total' 'adjust' 'true').ToLower() -eq 'true')
            BaseHours = [double]::Parse((Get-IniValue $ini 'daily_total' 'base_hours' '7.75'), $script:Inv)
            StepHours = [double]::Parse((Get-IniValue $ini 'daily_total' 'overtime_step_hours' '0.5'), $script:Inv)
        }
        Times           = @((Get-IniValue $ini 'schedule' 'times' '') -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        RequestCheckMin = [Math]::Max(1, [int](Get-IniValue $ini 'schedule' 'request_check_min' '5'))
        TitleKeyword    = Get-IniValue $ini 'obpm' 'title_keyword' '工数入力'
        CommitKey       = Get-IniValue $ini 'obpm' 'commit_key' '{ENTER}'
        InputMethod     = (Get-IniValue $ini 'obpm' 'input_method' 'mouse').ToLower()
        ObpmExe         = Get-IniValue $ini 'obpm' 'obpm_exe' ''
        AfterCommit     = (Get-IniValue $ini 'obpm' 'after_commit' 'none').ToLower()
        KeyWait         = [double]::Parse((Get-IniValue $ini 'obpm' 'key_wait' '0.15'), $script:Inv)
        SearchWait      = [double]::Parse((Get-IniValue $ini 'obpm' 'search_wait' '2.0'), $script:Inv)
    }
    if ($cfg.Mode -notin @('check', 'manual', 'live')) { throw "[run] mode は check / manual / live で指定してください (現在: $($cfg.Mode))" }
    if ($cfg.InputMethod -notin @('mouse', 'message')) { throw "[obpm] input_method は mouse / message で指定してください (現在: $($cfg.InputMethod))" }
    if ($cfg.AfterCommit -notin @('none', 'down', 'right')) { throw "[obpm] after_commit は none / down / right で指定してください (現在: $($cfg.AfterCommit))" }
    foreach ($t in $cfg.Times) { if ($t -notmatch '^\d{1,2}:\d{2}$') { throw "[schedule] times の時刻の形式が正しくありません: $t (例: 18:00)" } }
    return $cfg
}

function New-Logger([string]$DataDir, [switch]$Quiet) {
    $script:LogPath = Join-Path $DataDir 'obpm_input_log.txt'
    $script:LogQuiet = [bool]$Quiet
    return {
        param([string]$Msg)
        $line = "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))  $Msg"
        if (-not $script:LogQuiet) { Write-Host $line }
        try { Write-TextFile $script:LogPath ($line + "`r`n") -Append } catch { }
    }
}

function Write-Status([string]$DataDir, [string]$Result, [string]$Message, [int]$Count = 0) {
    try {
        Write-JsonFile (Join-Path $DataDir 'obpm_input_status.json') ([ordered]@{
                time = (Get-Date).ToString('yyyy-MM-dd HH:mm'); result = $Result; message = $Message; count = $Count })
    } catch { }
}

function Get-State([string]$Path) {
    $j = Read-JsonFile $Path
    if ($null -eq $j -or -not $j.PSObject.Properties['entered']) { return @{} }
    return ConvertTo-Hashtable $j.entered
}
function Save-State([string]$Path, $Entered) { Write-JsonFile $Path @{ entered = $Entered } }

# ==========================================================================
# 入力すべき内容 (記録 → 丸め → 前回との差分)
# ==========================================================================
function Get-ClosedDays([string]$DataDir) {
    # 「本日終了」を押した日 (ノートの作業タイマーが closed_days.json に書く)
    $set = New-Object System.Collections.Generic.HashSet[string]
    $j = Read-JsonFile (Join-Path $DataDir 'closed_days.json')
    if ($null -ne $j) { foreach ($p in $j.PSObject.Properties) { [void]$set.Add($p.Name) } }
    return , $set
}

function Get-DesiredHours([string]$RecordsCsv, [datetime]$Today, [int]$Unit, [string]$Mode, $Daily = $null, $ClosedDays = $null) {
    # 終わった日 (今日より前の日・「本日終了」を押した日) は合計を社内ルールに合わせる。今日の途中は丸めただけの値
    $warns = New-Object System.Collections.Generic.List[string]
    $out = @{}; $dayNotes = @{}
    if (-not (Test-Path -LiteralPath $RecordsCsv)) { $warns.Add("作業記録がありません: $RecordsCsv"); return @{ Hours = $out; Warnings = $warns; DayNotes = $dayNotes } }
    $monthStart = (Get-Date -Year $Today.Year -Month $Today.Month -Day 1).Date
    $mins = @{}
    foreach ($r in (Read-Records $RecordsCsv)) {
        $d = [datetime]::MinValue
        if (-not [datetime]::TryParseExact($r.date, 'yyyy-MM-dd', $script:Inv, 'None', [ref]$d)) { continue }
        if ($d -gt $Today.Date -or $d -lt $monthStart.AddDays(-7)) { continue }
        if (-not $r.project -or -not $r.wbs) { $warns.Add("プロジェクトコードか WBS が空のため入力できません: $($r.date) $($r.name)"); continue }
        $k = "$($r.date)|$($r.project)|$($r.wbs)"
        if (-not $mins.ContainsKey($k)) { $mins[$k] = 0.0 }
        $mins[$k] += $r.minutes
    }
    $byDay = [ordered]@{}
    foreach ($k in $mins.Keys) {
        $date = $k.Substring(0, 10)
        if ([string]::CompareOrdinal($date, $monthStart.ToString('yyyy-MM-dd')) -lt 0) {
            $warns.Add("前月分のため自動入力できません (手入力してください): $($k.Replace('|', ' / '))"); continue
        }
        if (-not $byDay.Contains($date)) { $byDay[$date] = [ordered]@{} }
        $byDay[$date][$k] = $mins[$k]
    }
    $todayS = $Today.ToString('yyyy-MM-dd')
    foreach ($date in $byDay.Keys) {
        $closed = ([string]::CompareOrdinal($date, $todayS) -lt 0) -or ($null -ne $ClosedDays -and $ClosedDays.Contains($date))
        $r = Get-DayHours $byDay[$date] $Unit $Mode $Daily $closed
        foreach ($k in $r.Hours.Keys) { $out[$k] = $r.Hours[$k] }
        if ($r.Notes.Count -gt 0) { $dayNotes[$date] = $r.Notes }
    }
    return @{ Hours = $out; Warnings = @($warns | Sort-Object -Unique); DayNotes = $dayNotes }
}

function Get-PendingItems($Desired, $Entered, [datetime]$Today) {
    $items = New-Object System.Collections.Generic.List[object]
    $ms = (Get-Date -Year $Today.Year -Month $Today.Month -Day 1).ToString('yyyy-MM-dd')
    $keys = @($Desired.Keys) + @($Entered.Keys) | Sort-Object -Unique
    foreach ($k in $keys) {
        $p = $k.Split('|', 3)
        if ([string]::CompareOrdinal($p[0], $ms) -lt 0) { continue }
        $want = 0.0; if ($Desired.ContainsKey($k)) { $want = [double]$Desired[$k] }
        $have = 0.0; if ($Entered.ContainsKey($k)) { $have = [double]$Entered[$k] }
        if ([Math]::Abs($want - $have) -gt 1e-9) {
            $items.Add([pscustomobject]@{ Key = $k; Date = $p[0]; Project = $p[1]; Wbs = $p[2]; Hours = $want })
        }
    }
    return , $items
}

# ==========================================================================
# 行の特定・タスク候補
# ==========================================================================
function Find-Row($Rows, [string]$Wbs) {
    # Rows: プロジェクト内の WBS 名 (Rows[0] はプロジェクト行)
    # Wbs : 「親 > 子 > 孫」 (親から順にたどる。途中の階層は省略可) / 末尾に #k で k 番目の出現
    # 同じ名前が続く場合は一番下 (一番深い) 行を返す。見つからなければ -1
    $parts = @($Wbs -split '>' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($parts.Count -eq 0) { return -1 }
    $occ = 1
    $m = [regex]::Match($parts[-1], '^(.*?)#(\d+)$')
    if ($m.Success) { $parts[-1] = $m.Groups[1].Value.Trim(); $occ = [int]$m.Groups[2].Value }
    $names = @(foreach ($r in $Rows) { Get-Norm $r })
    $idx = 1; $found = -1
    for ($i = 0; $i -lt $parts.Count; $i++) {
        $p = Get-Norm $parts[$i]
        $need = 1; if ($i -eq $parts.Count - 1) { $need = $occ }
        $found = -1
        $j = $idx
        while ($j -lt $names.Count) {
            if ($names[$j] -ceq $p) {
                $need--
                if ($need -eq 0) { $found = $j; break }
                while ($j + 1 -lt $names.Count -and $names[$j + 1] -ceq $p) { $j++ }   # 同名の連続は 1 回と数える
            }
            $j++
        }
        if ($found -lt 0) { return -1 }
        $idx = $found + 1
    }
    $j = $found
    while ($j + 1 -lt $names.Count -and $names[$j + 1] -ceq $names[$found]) { $j++ }
    return $j
}

function Get-TaskCandidates($AllRows) {
    # 全行の WBS 名から tasks.csv の候補 (タスク名, プロジェクトコード, WBS) を作る
    $out = New-Object System.Collections.Generic.List[object]
    $projects = New-Object System.Collections.Generic.List[object]
    foreach ($r in $AllRows) {
        if ($r -match $script:ProjectRow) { $projects.Add(@{ Code = ($r -split '[\u3000 ]')[0]; Rows = (New-Object System.Collections.Generic.List[string]) }); $projects[-1].Rows.Add($r) }
        elseif ($projects.Count -gt 0) { $projects[-1].Rows.Add($r) }
    }
    foreach ($pj in $projects) {
        $rows = $pj.Rows
        $names = @(foreach ($x in $rows) { Get-Norm $x })
        $runs = New-Object System.Collections.Generic.List[object]       # 同名の連続は 1 つにまとめる
        $i = 1
        while ($i -lt $rows.Count) {
            $j = $i
            while ($j + 1 -lt $rows.Count -and $names[$j + 1] -ceq $names[$i]) { $j++ }
            $runs.Add(@($i, $rows[$i])); $i = $j + 1
        }
        $counts = @{}
        foreach ($rn in $runs) { $k = Get-Norm $rn[1]; if (-not $counts.ContainsKey($k)) { $counts[$k] = 0 }; $counts[$k]++ }
        $seen = @{}; $prevPos = @{}
        foreach ($rn in $runs) {
            $pos = $rn[0]; $name = $rn[1]; $k = Get-Norm $name
            if (-not $seen.ContainsKey($k)) { $seen[$k] = 0 }
            $seen[$k]++
            if ($seen[$k] -eq 1) { $wbs = $name }
            else {
                # 前の出現より後ろにある、プロジェクト内で 1 回しか出てこない行を親の目印にする
                $anchor = $null
                for ($q = $pos - 1; $q -gt $prevPos[$k]; $q--) {
                    if ($counts.ContainsKey($names[$q]) -and $counts[$names[$q]] -eq 1 -and $names[$q] -cne $k) { $anchor = $rows[$q]; break }
                }
                if ($anchor) { $wbs = "$anchor > $name" } else { $wbs = "$name#$($seen[$k])" }
            }
            $prevPos[$k] = $pos
            $out.Add(@($name, $pj.Code, $wbs))
        }
    }
    return , $out
}

# ==========================================================================
# OBPM の画面操作 (UI Automation・キー送信・クリップボード)
# ==========================================================================
# 調査 (2026-11-18) で分かったこと:
#   PowerShell (System.Windows.Automation) からは部品の ID が起動のたびに変わる数字になり、
#   「押す」機能 (InvokePattern) も見えない。そのため部品は「表示名」と「位置」で探してクリックする。
#   ・ログイン画面 / メインメニュー / 工数入力画面 は同じウィンドウ内で切り替わる (タイトルで見分ける)
#   ・入力欄は見出し (「プロジェクトコード」など) のすぐ右にある。中身は Ctrl+A → Ctrl+C で読む
$script:FieldLabels = @{ txtPjCode = 'プロジェクトコード'; txtYm = '表示年月'; txtAccount = 'ログインアカウント'; txtPassword = 'パスワード' }
$script:ButtonNames = @{ btnSearch = '検索'; btnUpdate = '更新'; btnCancel = 'キャンセル'; btnLogin = 'ログイン'; btnLogout = 'ログアウト' }

function Initialize-Uia {
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Windows.Forms
    [WtNative]::DpiAware()
}

function Get-UiaChildren($El, [string]$Scope = 'Descendants') {
    $ts = [System.Windows.Automation.TreeScope]::$Scope
    return @($El.FindAll($ts, [System.Windows.Automation.Condition]::TrueCondition))
}

function Find-UiaByName($El, [string]$Name, [string]$ClassLike = '*') {
    $c = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $Name)
    # 呼ぶ側は @( ) で囲んで使う (0 件・1 件でも .Count が使えるように)
    return @(foreach ($e in $El.FindAll([System.Windows.Automation.TreeScope]::Descendants, $c)) {
            if ($e.Current.ClassName -like $ClassLike) { $e } })
}

function Get-Center($El) {
    $r = $El.Current.BoundingRectangle
    return @([int]($r.Left + $r.Width / 2), [int]($r.Top + $r.Height / 2))
}

function Get-ObpmWindows($Cfg) {
    # OBPM のウィンドウ (タイトルで見つけ、同じプロセスのウィンドウをすべて返す)。1 つ目がメインのウィンドウ
    $AE = [System.Windows.Automation.AutomationElement]
    $tops = Get-UiaChildren $AE::RootElement 'Children'
    $pidSet = @{}
    $main = $null
    foreach ($w in $tops) {
        try {
            $n = $w.Current.Name
            if ($n -like '*SI Object Browser*' -or $n -like "*$($Cfg.TitleKeyword)*") {
                $pidSet[$w.Current.ProcessId] = $true
                if ($null -eq $main -or $n -like "*$($Cfg.TitleKeyword)*") { $main = $w }
            }
        } catch { }
    }
    if ($null -eq $main) { return , @() }
    $others = @(foreach ($w in $tops) { try { if ($pidSet.ContainsKey($w.Current.ProcessId) -and $w -ne $main) { $w } } catch { } })
    return , (@($main) + $others)
}

function Get-ObpmState($Cfg) {
    # 戻り値: @{ State = kosu / menu / login / none / unknown; Win = メインのウィンドウ }
    $wins = Get-ObpmWindows $Cfg
    if ($wins.Count -eq 0) { return @{ State = 'none'; Win = $null } }
    $main = $wins[0]
    if ($main.Current.Name -like "*$($Cfg.TitleKeyword)*") { return @{ State = 'kosu'; Win = $main } }
    foreach ($w in $wins) {
        if (@(Find-UiaByName $w 'ログインアカウント').Count -gt 0 -and @(Find-UiaByName $w 'ログイン' '*BUTTON*').Count -gt 0) {
            return @{ State = 'login'; Win = $w }
        }
    }
    if (@(Find-UiaByName $main $Cfg.TitleKeyword '*STATIC*').Count -gt 0) { return @{ State = 'menu'; Win = $main } }
    return @{ State = 'unknown'; Win = $main }
}

function Open-KosuFromMenu($Cfg, $Win, $Log = $null) {
    # メインメニューの「工数入力」をクリックし、工数入力画面に切り替わるまで待つ
    #   ログイン直後はメニューの準備ができていなかったり、画面を開くのに時間がかかったりするため、
    #   リンクが出るまで待ち、最大 30 秒待っても切り替わらなければ 1 回だけクリックし直す
    $link = @()
    for ($i = 0; $i -lt 20; $i++) {
        $st0 = Get-ObpmState $Cfg
        if ($st0.State -eq 'kosu') { return }
        if ($null -ne $st0.Win) { $Win = $st0.Win }
        $link = @(Find-UiaByName $Win $Cfg.TitleKeyword '*STATIC*' | Where-Object { try { -not $_.Current.IsOffscreen -and $_.Current.BoundingRectangle.Width -gt 0 } catch { $false } })
        if ($link.Count -gt 0) { break }
        Start-Sleep -Milliseconds 500
    }
    if ($link.Count -eq 0) { Stop-Run 'メインメニューに「工数入力」が見つかりません。' }
    Start-Sleep -Seconds 1                     # メニューの表示が落ち着くのを待つ
    $hwnd = [IntPtr]$Win.Current.NativeWindowHandle
    # 押し方を順に試す: ① マウスを外から乗せてクリック ② 「工数入力」の枠 (文字の外) をクリック ③ ダブルクリック ④ フォーカス + Enter
    $how = @('', 'クリック', '枠のクリック', 'ダブルクリック', 'Enter キー')
    for ($try = 1; $try -le 4; $try++) {
        if ($try -ge 2) {
            # 押し直すのは、まだメインメニューのままで「工数入力」が見えているときだけ (読み込み中の別の画面を押さないように)
            $st1 = Get-ObpmState $Cfg
            if ($st1.State -ne 'menu') { break }
            $link = @(Find-UiaByName $st1.Win $Cfg.TitleKeyword '*STATIC*' | Where-Object { try { -not $_.Current.IsOffscreen -and $_.Current.BoundingRectangle.Width -gt 0 } catch { $false } })
            if ($link.Count -eq 0) { break }
            $Win = $st1.Win; $hwnd = [IntPtr]$Win.Current.NativeWindowHandle
        }
        [WtNative]::Activate($hwnd); Start-Sleep -Milliseconds 500
        if (-not [WtNative]::IsForeground($hwnd)) { Stop-Run "OBPM の画面を手前に出せませんでした。($(Get-ForegroundTitle))" }
        $xy = Get-Center $link[0]
        $lr = $link[0].Current.BoundingRectangle
        switch ($try) {
            1 {
                # 枠の外 → 枠の中 → 文字の上 とマウスを動かしてからクリック (マウスが乗ったことに反応する部品のため)
                [WtNative]::MoveTo([int]$lr.Left - 40, [int]$lr.Top - 30); Start-Sleep -Milliseconds 200
                [WtNative]::MoveTo([int]$lr.Left - 2, $xy[1]); Start-Sleep -Milliseconds 200
                [WtNative]::Click($xy[0], $xy[1])
            }
            2 {
                $tile = $null
                try { $tile = [System.Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($link[0]) } catch { }
                if ($null -ne $tile) {
                    $tr = $tile.Current.BoundingRectangle
                    $xy = @([int]($tr.Right - 12), [int]($tr.Top + $tr.Height / 2))
                }
                [WtNative]::Click($xy[0], $xy[1])
            }
            3 { [WtNative]::DoubleClick($xy[0], $xy[1]) }
            4 { try { $link[0].SetFocus() } catch { }; Start-Sleep -Milliseconds 300; [System.Windows.Forms.SendKeys]::SendWait('{ENTER}') }
        }
        $waitHalfSec = 20; if ($try -eq 1) { $waitHalfSec = 30 }
        for ($i = 0; $i -lt $waitHalfSec; $i++) {
            Start-Sleep -Milliseconds 500
            if ((Get-ObpmState $Cfg).State -eq 'kosu') {
                if ($try -gt 1 -and $null -ne $Log) { & $Log "  「工数入力」は $($how[$try]) で開けました" }
                Start-Sleep -Milliseconds ([int]($Cfg.SearchWait * 1000)); return
            }
        }
    }
    # 原因を調べるため、そのときの状態をログに残す
    $st = Get-ObpmState $Cfg
    $titles = @(Get-ObpmWindows $Cfg | ForEach-Object { try { "「$($_.Current.Name)」" } catch { } }) -join ' '
    $r = New-Object System.Windows.Rect; try { $r = $link[0].Current.BoundingRectangle } catch { }
    Stop-Run "「工数入力」をクリック・枠のクリック・ダブルクリック・Enter で押しましたが、工数入力画面に切り替わりませんでした (画面の状態: $($st.State) / OBPM のウィンドウ: $titles / クリックした位置: $([int]$xy[0]),$([int]$xy[1]) 部品の範囲: $([int]$r.Left),$([int]$r.Top)-$([int]$r.Right),$([int]$r.Bottom))"
}

function Get-ForegroundTitle {
    try {
        $h = [WtNative]::GetForegroundWindow()
        return "手前のウィンドウ: 「$([System.Windows.Automation.AutomationElement]::FromHandle($h).Current.Name)」"
    } catch { return '' }
}

function New-ObpmDriver($Cfg, $Win) {
    $d = [pscustomobject]@{ Win = $Win; Hwnd = [IntPtr]$Win.Current.NativeWindowHandle; Cfg = $Cfg; Pid = $Win.Current.ProcessId }
    $d | Add-Member ScriptMethod ForegroundInfo {
        # 手前にあるウィンドウの名前 (原因を調べるため)
        try {
            $h = [WtNative]::GetForegroundWindow()
            if ($h -eq [IntPtr]::Zero) { return '手前のウィンドウなし' }
            $el = [System.Windows.Automation.AutomationElement]::FromHandle($h)
            $who = 'ほかのアプリ'; if ($el.Current.ProcessId -eq $this.Pid) { $who = 'OBPM のメッセージ画面など' }
            $texts = @(@(foreach ($t in (Get-UiaChildren $el)) { try { $n = $t.Current.Name; if ($n) { $n } } catch { } }) | Select-Object -First 6)
            return "手前のウィンドウ: 「$($el.Current.Name)」($who) $($texts -join ' / ')"
        } catch { return '手前のウィンドウ: 不明' } }
    $d | Add-Member ScriptMethod CheckForeground {
        if (-not [WtNative]::IsForeground($this.Hwnd)) { Stop-Run "OBPM 以外のウィンドウが手前にあるため中断しました (誤入力防止)。$($this.ForegroundInfo())" } }
    $d | Add-Member ScriptMethod Activate {
        [WtNative]::Activate($this.Hwnd); Start-Sleep -Milliseconds 400; $this.CheckForeground() }
    $d | Add-Member ScriptMethod Key { param($K)
        $this.CheckForeground(); [System.Windows.Forms.SendKeys]::SendWait($K); Start-Sleep -Milliseconds ([int]($this.Cfg.KeyWait * 1000)) }
    $d | Add-Member ScriptMethod TypeText { param($S)
        $esc = [regex]::Replace($S, '[+^%~(){}\[\]]', '{$0}')
        $this.CheckForeground(); [System.Windows.Forms.SendKeys]::SendWait($esc); Start-Sleep -Milliseconds ([int]($this.Cfg.KeyWait * 1000)) }
    $d | Add-Member ScriptMethod ClickAt { param($El)
        $xy = Get-Center $El; $this.CheckForeground(); [WtNative]::Click($xy[0], $xy[1]); Start-Sleep -Milliseconds 250 }
    $d | Add-Member ScriptMethod FocusEl { param($El)
        # 入力欄にフォーカスを移す。input_method = message ならマウスを使わない (リモートデスクトップ最小化中の対策)
        if ($this.Cfg.InputMethod -eq 'message') {
            try { $this.CheckForeground(); $El.SetFocus(); Start-Sleep -Milliseconds 200; return } catch { }
        }
        $this.ClickAt($El) }
    $d | Add-Member ScriptMethod PressEl { param($El)
        # ボタンを押す。input_method = message ならマウスを使わずボタンにクリックのメッセージを送る
        if ($this.Cfg.InputMethod -eq 'message') {
            $h = [IntPtr]$El.Current.NativeWindowHandle
            if ($h -ne [IntPtr]::Zero) { $this.CheckForeground(); if ([WtNative]::PostButtonClick($h)) { Start-Sleep -Milliseconds 300; return } }
        }
        $this.ClickAt($El) }
    $d | Add-Member ScriptMethod InputBox { param($Id)
        # 見出しの右隣にある入力欄 (いちばん内側の部品) を探す
        $label = $script:FieldLabels[$Id]
        $lab = @(Find-UiaByName $this.Win $label '*STATIC*')
        if ($lab.Count -eq 0) { Stop-Run "見出し「$label」が見つかりません" }
        $lr = $lab[0].Current.BoundingRectangle
        $best = $null; $bestArea = [double]::MaxValue
        foreach ($e in (Get-UiaChildren $this.Win)) {
            try {
                if ($e.Current.ClassName -notlike '*Window.8*') { continue }
                $r = $e.Current.BoundingRectangle
                if ($r.Left -ge $lr.Right - 2 -and $r.Left -le $lr.Right + 10 -and $r.Top -ge $lr.Top - 2 -and $r.Bottom -le $lr.Bottom + 2 -and $r.Width -gt 10) {
                    $area = $r.Width * $r.Height
                    if ($area -lt $bestArea) { $best = $e; $bestArea = $area }
                }
            } catch { }
        }
        if ($null -eq $best) { Stop-Run "「$label」の入力欄が見つかりません" }
        return $best }
    $d | Add-Member ScriptMethod GetField { param($Id)
        $this.FocusEl($this.InputBox($Id)); $this.Key('^a')
        $v = $this.Copy(); if ($null -eq $v) { return '' } else { return $v } }
    $d | Add-Member ScriptMethod SetField { param($Id, $Value)
        $this.FocusEl($this.InputBox($Id))
        $this.Key('^a'); $this.Key('{DEL}')
        if ($Value) { $this.TypeText($Value) }
        $this.Key('{TAB}'); Start-Sleep -Milliseconds 300
        $got = $this.GetField($Id)
        $this.Key('{TAB}')
        if ((Get-Norm $got) -cne (Get-Norm $Value)) { Stop-Run "入力欄「$($script:FieldLabels[$Id])」に「$Value」を設定できませんでした (現在: 「$got」)" } }
    $d | Add-Member ScriptMethod Press { param($Id)
        # 同じ名前のボタンが複数あるとき (更新・キャンセル) は、画面下部のボタンを使う
        $name = $script:ButtonNames[$Id]
        $btns = @(Find-UiaByName $this.Win $name '*BUTTON*')
        if ($btns.Count -eq 0) { Stop-Run "「$name」ボタンが見つかりません" }
        $b = $btns | Sort-Object { $_.Current.BoundingRectangle.Top } -Descending | Select-Object -First 1
        $this.PressEl($b) }
    $d | Add-Member ScriptMethod Search {
        $this.Press('btnSearch'); Start-Sleep -Milliseconds ([int]($this.Cfg.SearchWait * 1000)) }
    $d | Add-Member ScriptMethod FocusGrid {
        # 工数の表 = 中に部品を持たない (スクロールバーだけは持ってよい) 部品のうち、いちばん大きいもの
        #   行が多いとき (検索語なしで全件を出したときなど) は表の中にスクロールバーの部品ができるため、それは数えない
        # 表の読み込み中はフォーカスを受け取れないことがあるので、最大 15 秒ほど待ちながら繰り返す
        $last = ''
        for ($i = 0; $i -lt 30; $i++) {
            # 検索の後に OBPM のメッセージ画面 (確認・警告・処理中など) が出ていると、表はフォーカスを受け取れない
            $dl = $this.Dialogs(); if ($dl.Count -gt 0) { Stop-Run "検索後にメッセージが出ています: $($dl -join ' | ')" }
            $cands = New-Object System.Collections.Generic.List[object]
            foreach ($e in (Get-UiaChildren $this.Win)) {
                try {
                    if ($e.Current.ClassName -notlike '*Window.8*') { continue }
                    $other = $false
                    $c = [System.Windows.Automation.TreeWalker]::ControlViewWalker.GetFirstChild($e)
                    while ($null -ne $c) {
                        $isBar = ($c.Current.ClassName -like '*SCROLLBAR*') -or ($c.Current.ControlType -eq [System.Windows.Automation.ControlType]::ScrollBar)
                        if (-not $isBar) { $other = $true; break }
                        $c = [System.Windows.Automation.TreeWalker]::ControlViewWalker.GetNextSibling($c)
                    }
                    if ($other) { continue }
                    $r = $e.Current.BoundingRectangle
                    $cands.Add([pscustomobject]@{ El = $e; Area = $r.Width * $r.Height })
                } catch { }
            }
            if ($cands.Count -eq 0) { $last = '工数の表が見つかりません' }
            else {
                $best = ($cands | Sort-Object Area -Descending | Select-Object -First 1).El
                try { $best.SetFocus(); Start-Sleep -Milliseconds 400; return } catch { $last = $_.Exception.Message }
                # SetFocus を受け付けない表の場合: Tab キーで表までフォーカスを送る (Tab はセルの選択位置を変えない)
                if ($i % 10 -eq 0 -and $this.TabToGrid()) { return }
            }
            Start-Sleep -Milliseconds 500
        }
        Stop-Run "工数の表にフォーカスを移せませんでした ($last)。表の読み込みに時間がかかっている場合は、ini の search_wait を大きくしてください`r`n$($this.GridDiag())" }
    $d | Add-Member ScriptMethod TabToGrid {
        # フォーカスのある部品が表 (OBPM の中の、幅 500・高さ 200 以上の部品) になるまで Tab を押す (最大 20 回)
        for ($n = 0; $n -lt 20; $n++) {
            try {
                $f = [System.Windows.Automation.AutomationElement]::FocusedElement
                if ($null -ne $f -and $f.Current.ProcessId -eq $this.Pid) {
                    $r = $f.Current.BoundingRectangle
                    if ($r.Width -ge 500 -and $r.Height -ge 200) { Start-Sleep -Milliseconds 300; return $true }
                }
            } catch { }
            $this.Key('{TAB}')
        }
        return $false }
    $d | Add-Member ScriptMethod GridDiag {
        # 原因調査用: 大きな部品の一覧 (位置・中の部品・フォーカスを受け取れるか)
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add('    [調査用] 画面内の大きな部品:')
        foreach ($e in (Get-UiaChildren $this.Win)) {
            try {
                $r = $e.Current.BoundingRectangle
                if ($r.Width * $r.Height -lt 40000) { continue }
                $kids = New-Object System.Collections.Generic.List[string]
                $c = [System.Windows.Automation.TreeWalker]::ControlViewWalker.GetFirstChild($e)
                while ($null -ne $c -and $kids.Count -lt 6) {
                    $cr = $c.Current.BoundingRectangle
                    $kids.Add("$($c.Current.ControlType.ProgrammaticName -replace 'ControlType\.','')/$($c.Current.ClassName -replace '^WindowsForms10\.|\.app\..*$','') ($([int]$cr.Left),$([int]$cr.Top))-($([int]$cr.Right),$([int]$cr.Bottom))")
                    $c = [System.Windows.Automation.TreeWalker]::ControlViewWalker.GetNextSibling($c)
                }
                $cls = $e.Current.ClassName -replace '^WindowsForms10\.|\.app\..*$', ''
                $lines.Add("    - $($e.Current.ControlType.ProgrammaticName -replace 'ControlType\.','')/$cls ($([int]$r.Left),$([int]$r.Top))-($([int]$r.Right),$([int]$r.Bottom)) 有効=$($e.Current.IsEnabled) フォーカス可=$($e.Current.IsKeyboardFocusable) フォーカス中=$($e.Current.HasKeyboardFocus) 中の部品=[$($kids -join ', ')]")
            } catch { }
        }
        return ($lines -join "`r`n") }
    $d | Add-Member ScriptMethod FindDialogs {
        # OBPM のメッセージ画面 (メイン以外のウィンドウ) を返す
        #   画面の一覧 (デスクトップ直下) にあるものに加え、メイン画面の中にぶら下がって見えるもの
        #   (メイン画面が持ち主のメッセージ画面は、UI Automation ではメイン画面の子として見えることがある: 実機確認) も探す
        $seen = New-Object System.Collections.Generic.HashSet[long]
        $out = New-Object System.Collections.Generic.List[object]
        $cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)
        $inner = @(); try { $inner = @($this.Win.FindAll([System.Windows.Automation.TreeScope]::Descendants, $cond)) } catch { }
        foreach ($w in @(@(Get-UiaChildren ([System.Windows.Automation.AutomationElement]::RootElement) 'Children') + $inner)) {
            try {
                if ($w.Current.ProcessId -ne $this.Pid) { continue }
                $h = [long]$w.Current.NativeWindowHandle
                if ($h -eq 0 -or $h -eq [long]$this.Hwnd) { continue }
                if ($w.Current.ControlType -ne [System.Windows.Automation.ControlType]::Window) { continue }
                if ($seen.Add($h)) { $out.Add($w) }
            } catch { }
        }
        return , $out }
    $d | Add-Member ScriptMethod Dialogs {
        # メッセージ画面の文言 (タイトル / 中の文字)
        return , @(foreach ($w in $this.FindDialogs()) {
                try { ((@($w.Current.Name) + @($this.DialogText($w))) | Where-Object { $_ }) -join ' / ' } catch { } }) }
    $d | Add-Member ScriptMethod DialogText { param($W)
        return ((@(foreach ($t in (Get-UiaChildren $W)) { $n = $t.Current.Name; if ($n) { $n } })) -join ' / ') }
    $d | Add-Member ScriptMethod ClickDialogButton { param($W, $Like)
        $b = @(foreach ($e in (Get-UiaChildren $W)) { if ($e.Current.ClassName -like '*BUTTON*' -or $e.Current.ControlType.ProgrammaticName -like '*Button') { if ($e.Current.Name -like $Like) { $e } } })
        if ($b.Count -eq 0) { return $false }
        $h = [IntPtr]$W.Current.NativeWindowHandle
        if (-not [WtNative]::IsForeground($h)) { [WtNative]::Activate($h); Start-Sleep -Milliseconds 300 }
        if (-not [WtNative]::IsForeground($h)) { return $false }
        $bh = [IntPtr]$b[0].Current.NativeWindowHandle
        if ($this.Cfg.InputMethod -eq 'message' -and $bh -ne [IntPtr]::Zero -and [WtNative]::PostButtonClick($bh)) { Start-Sleep -Milliseconds 300; return $true }
        $xy = Get-Center $b[0]; [WtNative]::Click($xy[0], $xy[1]); Start-Sleep -Milliseconds 300
        return $true }
    $d | Add-Member ScriptMethod ConfirmUpdate {
        # 「更新」を押した後の確認画面「更新します。よろしいですか？」で「はい」を押す (調査 B-11)
        # 戻り値: ok / ok:<完了メッセージ> / nodialog (確認画面が出ない = 更新されていない。調査 B-14) / unexpected:<文言>
        $dlg = $null
        for ($i = 0; $i -lt 20 -and $null -eq $dlg; $i++) { Start-Sleep -Milliseconds 500; $ds = $this.FindDialogs(); if ($ds.Count -gt 0) { $dlg = $ds[0] } }
        if ($null -eq $dlg) { return 'nodialog' }
        $text = $this.DialogText($dlg)
        if ($text -notmatch '更新します') { return "unexpected:$text" }
        foreach ($cb in (Find-UiaByName $dlg 'プロジェクト別工数表')) {   # 「プロジェクト別工数表」にチェックが付いていたら押さない
            try { if ($cb.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern).Current.ToggleState -eq 'On') { return "unexpected:「プロジェクト別工数表」にチェックが付いています" } } catch { }
        }
        if (-not $this.ClickDialogButton($dlg, 'はい*')) { return "unexpected:確認画面の「はい」を押せませんでした ($text)" }
        # 押せたか確認 (確認画面が残っていれば、もう一度だけ押す)
        Start-Sleep -Milliseconds 500
        if (@($this.FindDialogs() | Where-Object { $this.DialogText($_) -match '更新します' }).Count -gt 0) {
            if (-not $this.ClickDialogButton($dlg, 'はい*')) { return "unexpected:確認画面の「はい」を押せませんでした ($text)" }
        }
        Start-Sleep -Milliseconds ([int]($this.Cfg.SearchWait * 1000))
        # 更新後にメッセージが出た場合: 完了の知らせなら閉じる。それ以外は止める
        $ds = $this.FindDialogs()
        if ($ds.Count -eq 0) { return 'ok' }
        $t2 = $this.DialogText($ds[0])
        if ($t2 -match '更新しました|完了') {
            if ($this.ClickDialogButton($ds[0], '*')) { return "ok:$t2" }
        }
        return "unexpected:$t2" }
    $d | Add-Member ScriptMethod Copy {
        foreach ($wait in 100, 300, 600) {
            for ($t = 0; $t -lt 5; $t++) { try { [System.Windows.Forms.Clipboard]::Clear(); break } catch { Start-Sleep -Milliseconds 50 } }
            $this.Key('^c'); Start-Sleep -Milliseconds $wait
            for ($t = 0; $t -lt 5; $t++) {
                try { if ([System.Windows.Forms.Clipboard]::ContainsText()) { return ([System.Windows.Forms.Clipboard]::GetText()).TrimEnd("`r", "`n") } else { break } }
                catch { Start-Sleep -Milliseconds 50 }
            }
        }
        return $null }
    return $d
}

# ==========================================================================
# 自動ログイン
#   ログイン情報は、このフォルダの obpm_login.dat に暗号化して保存する (Windows の DPAPI)
#   暗号は「この PC のこの Windows ユーザー」でしか解けない。共有フォルダや ini には保存しない
#   ログインに失敗したら、アカウントのロックを防ぐため、登録し直すまで自動ではログインしない
# ==========================================================================
function Get-LoginPath { return (Join-Path $PSScriptRoot 'obpm_login.dat') }
function Get-LoginFailPath { return (Join-Path $PSScriptRoot 'obpm_login_failed.json') }

function Set-LoginInfo {
    # メニューから実行: ログインアカウントとパスワードを登録する
    Write-Host 'OBPM の自動ログインに使う情報を登録します。'
    Write-Host '（パスワードは暗号化して、この PC のこのフォルダに保存します。この PC のこのユーザーでしか読めません）'
    $acc = Read-Host 'ログインアカウント'
    if (-not $acc) { Write-Host '中止しました。'; return }
    $p1 = Read-Host 'パスワード' -AsSecureString
    $p2 = Read-Host 'パスワード (確認のためもう一度)' -AsSecureString
    $b1 = ConvertTo-PlainText $p1; $b2 = ConvertTo-PlainText $p2
    $same = ($b1 -ceq $b2); $empty = (-not $b1)
    $b1 = $null; $b2 = $null
    if ($empty) { Write-Host '中止しました (パスワードが空です)。'; return }
    if (-not $same) { Write-Host 'パスワードが一致しません。もう一度登録してください。'; return }
    Write-JsonFile (Get-LoginPath) ([ordered]@{ account = $acc; password = (ConvertFrom-SecureString $p1); saved_at = (Get-Date).ToString($script:TimeFmt) })
    Remove-Item -LiteralPath (Get-LoginFailPath) -ErrorAction SilentlyContinue
    Write-Host "登録しました: $(Get-LoginPath)"
    try {
        $cfg0 = Get-InputConfig $(if ($Config) { $Config } else { Join-Path $PSScriptRoot 'obpm_input.ini' })
        $l = Find-ObpmLauncher $cfg0
        if ($l) { Write-Host "OBPM の場所: $l (OBPM が起動していないときは、ここから起動します)" }
        else { Write-Host '※ OBPM の場所が見つかりませんでした。obpm_input.ini の obpm_exe に OBPM の実行ファイルの場所を書いてください。' -ForegroundColor Yellow }
    } catch { }
    Write-Host '※ 待機モードが動いている場合も、次の入力からこの情報でログインします。'
}

function Clear-LoginInfo {
    foreach ($p in @((Get-LoginPath), (Get-LoginFailPath))) { Remove-Item -LiteralPath $p -ErrorAction SilentlyContinue }
    Write-Host '自動ログインの情報を削除しました。'
}

function ConvertTo-PlainText([System.Security.SecureString]$Sec) {
    $b = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Sec)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($b) } finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b) }
}

function Get-LoginInfo {
    # 戻り値: @{ Account; Password (SecureString) } / 未登録・前回失敗なら $null と理由
    $p = Get-LoginPath
    if (-not (Test-Path -LiteralPath $p)) { return @{ Info = $null; Why = '自動ログインが登録されていません (obpm_input.bat の 6 で登録できます)' } }
    if (Test-Path -LiteralPath (Get-LoginFailPath)) {
        return @{ Info = $null; Why = '前回の自動ログインに失敗したため、自動ではログインしません (アカウントのロック防止)。パスワードを確認して obpm_input.bat の 6 で登録し直してください' }
    }
    try {
        $j = Read-JsonFile $p
        return @{ Info = @{ Account = [string]$j.account; Password = (ConvertTo-SecureString ([string]$j.password)) }; Why = '' }
    } catch {
        return @{ Info = $null; Why = 'ログイン情報を読めませんでした (別の PC・別のユーザーで登録した情報は使えません)。obpm_input.bat の 6 で登録し直してください' }
    }
}

function Test-AutoLoginReady { return ($null -ne (Get-LoginInfo).Info) }

function Wait-ObpmState($Cfg, [string[]]$Want, [int]$Seconds) {
    for ($i = 0; $i -lt $Seconds * 2; $i++) {
        Start-Sleep -Milliseconds 500
        $st = Get-ObpmState $Cfg
        if ($st.State -in $Want) { return $st }
    }
    return (Get-ObpmState $Cfg)
}

function Invoke-ObpmLogin($Cfg, $Win, $Log) {
    # ログイン画面にアカウントとパスワードを入れて「ログイン」を押す。パスワードはログに出さない
    $li = Get-LoginInfo
    if ($null -eq $li.Info) { & $Log "OBPM がログイン画面です。$($li.Why)"; return $false }
    & $Log 'OBPM に自動でログインします。'
    $drv = New-ObpmDriver $Cfg $Win
    $drv.Activate()
    $drv.SetField('txtAccount', $li.Info.Account)
    $drv.FocusEl($drv.InputBox('txtPassword'))
    $drv.Key('^a'); $drv.Key('{DEL}')
    $pw = ConvertTo-PlainText $li.Info.Password
    try { $drv.TypeText($pw) } finally { $pw = $null }
    $drv.Press('btnLogin')
    $st = Wait-ObpmState $Cfg @('menu', 'kosu') 30
    if ($st.State -in @('menu', 'kosu')) { & $Log 'ログインしました。'; return $true }
    # 失敗: メッセージを読んで閉じ、登録し直すまで自動ログインを止める
    $msg = ''
    try {
        $dl = $drv.FindDialogs()
        if ($dl.Count -gt 0) { $msg = $drv.DialogText($dl[0]); [void]$drv.ClickDialogButton($dl[0], 'OK*') }
    } catch { }
    Write-JsonFile (Get-LoginFailPath) ([ordered]@{ failed_at = (Get-Date).ToString($script:TimeFmt); message = $msg })
    Stop-Run "自動ログインに失敗しました ($msg)。アカウントのロックを防ぐため、obpm_input.bat の 6 で登録し直すまで自動ではログインしません"
}

function Get-ObpmDate($Win) {
    # 工数入力画面の右上に出ている日付 (OBPM にログインした日。日付が変わっても更新されない)
    foreach ($e in @(Get-UiaChildren $Win)) {
        try {
            if ($e.Current.ClassName -notlike '*STATIC*') { continue }
            $n = $e.Current.Name
            if ($n -match '^\d{4}/\d{1,2}/\d{1,2}$') { return [datetime]::ParseExact($n, 'yyyy/M/d', $script:Inv) }
        } catch { }
    }
    return $null
}

function Get-LauncherPath { return (Join-Path $PSScriptRoot 'obpm_launcher.txt') }

function Save-ObpmLauncher($Win) {
    # 動いている OBPM の実行ファイルの場所を覚えておく (次に OBPM が起動していないときに使う)
    try {
        $exe = (Get-Process -Id $Win.Current.ProcessId -ErrorAction Stop).Path
        if ($exe -and (Test-Path -LiteralPath $exe)) {
            $p = Get-LauncherPath
            if (-not (Test-Path -LiteralPath $p) -or (Read-TextAuto $p).Trim() -ne $exe) { Write-TextFile $p $exe }
        }
    } catch { }
}

function Find-ObpmLauncher($Cfg) {
    # OBPM を起動するファイル (実行ファイルまたはショートカット) を探す。見つからなければ ''
    #   ① ini の obpm_exe → ② 前に動いていた OBPM の場所 → ③ スタートメニュー・デスクトップのショートカット → ④ Program Files
    if ($Cfg.ObpmExe -and (Test-Path -LiteralPath $Cfg.ObpmExe)) { return $Cfg.ObpmExe }
    $p = Get-LauncherPath
    if (Test-Path -LiteralPath $p) { $x = (Read-TextAuto $p).Trim(); if ($x -and (Test-Path -LiteralPath $x)) { return $x } }
    $dirs = @('StartMenu', 'CommonStartMenu', 'Desktop', 'CommonDesktopDirectory') | ForEach-Object { [Environment]::GetFolderPath($_) } | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
    $ws = New-Object -ComObject WScript.Shell
    foreach ($d in $dirs) {
        foreach ($f in @(Get-ChildItem -LiteralPath $d -Recurse -Include '*.lnk', '*.appref-ms' -ErrorAction SilentlyContinue)) {
            if ($f.BaseName -notlike '*Object Browser PM*' -and $f.BaseName -notlike '*OBPM*') { continue }
            if ($f.Extension -eq '.appref-ms') { return $f.FullName }
            try { $t = $ws.CreateShortcut($f.FullName).TargetPath; if ($t -and $t -like '*.exe' -and (Test-Path -LiteralPath $t)) { return $t } } catch { }
        }
    }
    foreach ($root in @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Where-Object { $_ }) {
        $hit = Get-ChildItem -LiteralPath $root -Filter 'obpm.exe' -Recurse -Depth 4 -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($hit) { return $hit.FullName }
    }
    return ''
}

function Start-ObpmApp($Cfg, $Log, [string]$Exe = '') {
    # OBPM を起動してログイン画面が出るまで待つ
    if (-not $Exe) { $Exe = Find-ObpmLauncher $Cfg }
    if (-not $Exe) { & $Log 'OBPM の場所が見つかりません。obpm_input.ini の obpm_exe に OBPM の実行ファイルの場所を書いてください。'; return $false }
    & $Log "OBPM を起動します ($Exe)"
    $dir = Split-Path -Parent $Exe
    Start-Process -FilePath $Exe -WorkingDirectory $dir
    $st = Wait-ObpmState $Cfg @('login', 'menu', 'kosu') 90
    if ($st.State -eq 'none') { & $Log 'OBPM を起動しましたが、90 秒たっても画面が出ませんでした。'; return $false }
    return $true
}

function Invoke-Relogin($Cfg, $St, $Log) {
    # 日付を今日にするため、OBPM を × で閉じて起動し直し、ログインする (ログアウトでは日付が変わらない)
    & $Log 'OBPM の日付が今日ではないため、OBPM を閉じて起動し直します。'
    $proc = Get-Process -Id $St.Win.Current.ProcessId -ErrorAction SilentlyContinue
    $exe = ''; try { $exe = $proc.Path } catch { }
    if (-not $exe) { $exe = Find-ObpmLauncher $Cfg }
    if (-not $exe) { Stop-Run 'OBPM の実行ファイルの場所が分からないため、起動し直せません。obpm_input.ini の obpm_exe に OBPM の場所を書いてください' }
    # × で閉じる (まだ何も入力していない)
    [void]$proc.CloseMainWindow()
    for ($i = 0; $i -lt 40 -and -not $proc.HasExited; $i++) {
        Start-Sleep -Milliseconds 500
        if ($i -eq 4) {
            # 「終了しますか？」などが出ていれば「はい」を押す
            try {
                $drv = New-ObpmDriver $Cfg $St.Win
                $dl = $drv.FindDialogs()
                if ($dl.Count -gt 0) { & $Log "  閉じるときのメッセージ: $($drv.DialogText($dl[0]))"; [void]$drv.ClickDialogButton($dl[0], 'はい*') }
            } catch { }
        }
        $proc.Refresh()
    }
    if (-not $proc.HasExited) { Stop-Run 'OBPM を閉じられませんでした。OBPM を手で閉じてから、もう一度実行してください' }
    Start-Sleep -Seconds 2
    if (-not (Start-ObpmApp $Cfg $Log $exe)) { Stop-Run "OBPM を起動できませんでした ($exe)" }
    $st2 = Get-ObpmState $Cfg
    if ($st2.State -eq 'login') { [void](Invoke-ObpmLogin $Cfg $st2.Win $Log) }
}

function Get-ReadyDriver($Cfg, $Log) {
    # OBPM の状態を見て、今日の日付の工数入力画面のドライバーを返す。ログインできないときは $null
    Initialize-Uia
    $today = (Get-Date).Date
    for ($round = 1; $round -le 2; $round++) {
        $st = Get-ObpmState $Cfg
        if ($st.State -eq 'none') {
            if (-not (Test-AutoLoginReady)) { & $Log "OBPM が起動していません。起動してログインし、工数入力画面を開いてください ($((Get-LoginInfo).Why))"; return $null }
            if (-not (Start-ObpmApp $Cfg $Log)) { return $null }
            $st = Get-ObpmState $Cfg
        }
        if ($null -ne $st.Win) { Save-ObpmLauncher $st.Win }
        if ($st.State -eq 'login') {
            if (-not (Invoke-ObpmLogin $Cfg $st.Win $Log)) { return $null }
            $st = Get-ObpmState $Cfg
        }
        if ($st.State -eq 'menu') { & $Log 'メインメニューから「工数入力」を開きます。'; Open-KosuFromMenu $Cfg $st.Win $Log; $st = Get-ObpmState $Cfg }
        if ($st.State -eq 'unknown') { Stop-Run "OBPM の画面の状態を判定できません (タイトル: $($st.Win.Current.Name))。工数入力画面を開いてください。" }
        if ($st.State -ne 'kosu') { Stop-Run '工数入力画面を開けませんでした。' }
        # 日付の確認 (OBPM は開いたときの日付のままになり、今日の列に入力できないため)
        $od = Get-ObpmDate $st.Win
        if ($null -eq $od -or $od -eq $today) { return (New-ObpmDriver $Cfg $st.Win) }
        if ($round -ge 2) { Stop-Run "OBPM を起動し直しても日付が今日になりません ($($od.ToString('yyyy/MM/dd')))。OBPM を終了して起動し直してください" }
        if (-not (Test-AutoLoginReady)) {
            & $Log "OBPM の日付が $($od.ToString('yyyy/MM/dd')) のままです。OBPM を閉じて開き直し、ログインしてください ($((Get-LoginInfo).Why))"
            return $null
        }
        Invoke-Relogin $Cfg $st $Log
    }
}

# ==========================================================================
# 表の読み取りと入力 (画面操作はドライバー経由。テストでは模擬の表を使う)
# ==========================================================================
function Move-ToWbsColumn($Drv, [datetime]$Today) {
    # 基準点 (1 行目・今日の列) から WBS 列へ移動し、1 行目の値を返す
    for ($i = 0; $i -lt $script:OffsetCols + $Today.Day; $i++) { $Drv.Key('{LEFT}') }
    $v = $Drv.Copy()
    if (-not $v -or $v -notmatch $script:ProjectRow) {
        Stop-Run "表の位置を確認できませんでした (1 行目の WBS 列の値: '$v')。表示年月が今月か確認してください。"
    }
    return $v
}

function Open-Grid($Drv, [string]$Key, [datetime]$Today, $Log) {
    # 検索して表の 1 行目・WBS 列へ移動し、1 行目の値を返す
    # 途中で OBPM からフォーカスが外れると選択セルの位置が分からなくなるため、失敗したら検索をやり直して 1 回だけ再試行する
    for ($try = 1; ; $try++) {
        try {
            $Drv.Activate()
            $Drv.SetField('txtPjCode', $Key)
            $Drv.Search()
            $dl = $Drv.Dialogs(); if ($dl.Count -gt 0) { Stop-Run "検索後にメッセージが出ています: $($dl -join ' | ')" }
            $Drv.FocusGrid()
            return (Move-ToWbsColumn $Drv $Today)
        } catch {
            $am = Get-AbortMessage $_.Exception
            if ($null -eq $am -or $try -ge 2 -or $am -like '検索後にメッセージ*') { throw }
            & $Log "  表の位置を確認できなかったため、検索をやり直します ($am)"
            Start-Sleep -Seconds 1
        }
    }
}

function Read-GridRows($Drv, [string]$First, [bool]$StopAtNextProject) {
    # WBS 列を下へ読む。戻り値: Rows (名前の一覧) / Cur (現在の行番号) / HasMore (下にまだ行があるか)
    $rows = New-Object System.Collections.Generic.List[string]
    $rows.Add($First)
    $run = 1; $cur = 0
    for ($n = 0; $n -lt $script:MaxRows; $n++) {
        $Drv.Key('{DOWN}')
        $v = $Drv.Copy()
        # 読めない = 表の最後で ↓ が効かなかった、または ↓ が取りこぼされて同じセルを読んだ (実機で発生)。
        # もう 2 回まで ↓ を押し直し、それでも読めなければ表の最後とみなす
        for ($retry = 0; -not $v -and $retry -lt 2; $retry++) { Start-Sleep -Milliseconds 300; $Drv.Key('{DOWN}'); $v = $Drv.Copy() }
        if (-not $v) { break }
        $cur++
        if ($StopAtNextProject -and $v -match $script:ProjectRow) { return @{ Rows = $rows; Cur = $cur; HasMore = $true; Next = $v } }
        if ((Get-Norm $v) -ceq (Get-Norm $rows[$rows.Count - 1])) { $run++ } else { $run = 1 }
        $rows.Add($v)
        if ($run -ge $script:SameRunLimit) {
            # 最終行で ↓ が効かず同じ値が続いた → 重複分を除く
            $rows.RemoveRange($rows.Count - ($run - 1), $run - 1)
            return @{ Rows = $rows; Cur = $rows.Count - 1; HasMore = $false; Next = $null }
        }
    }
    return @{ Rows = $rows; Cur = $cur; HasMore = $false; Next = $null }
}

function Get-SearchKey([string]$Project) {
    # OBPM の検索は「-」付きのコードでは見つからない (実機確認) ため、「-」より前 (6 桁の数字) で検索する
    return ($Project -split '-')[0].Trim()
}

function Get-RowCode([string]$Row) { return ($Row -split '[\u3000 ]')[0] }

function Move-Rows($Drv, [int]$Cur, [int]$Target) {
    # 表のスクロールが追いつかずにキーが取りこぼされないよう、10 行ごとに少し待つ
    $k = '{DOWN}'; if ($Target -lt $Cur) { $k = '{UP}' }
    for ($i = 0; $i -lt [Math]::Abs($Target - $Cur); $i++) {
        $Drv.Key($k)
        if ($i % 10 -eq 9) { Start-Sleep -Milliseconds 300 }
    }
}

function Sync-Row($Drv, $Rows, [int]$Target, $Log) {
    # 移動後の行が目的の行か確かめる。ずれていたら前後 $script:SyncRange 行を探して補正する
    #   行数の多い表では、↓ を押しても 1 行進まない (または 2 行進む) ことがあり、数えた行番号と実際の行がずれる (実機確認)
    #   目的の行 = 名前が一致し、すぐ下の行が同じ名前ではない行 (同じ名前が続く場合は一番下の行に入力するため)
    $want = Get-Norm $Rows[$Target]
    $pos = 0                                   # 移動後の位置からの相対位置
    foreach ($d in (Get-SyncOffsets)) {
        $k = '{DOWN}'; if ($d -lt $pos) { $k = '{UP}' }
        for ($i = 0; $i -lt [Math]::Abs($d - $pos); $i++) { $Drv.Key($k) }
        $pos = $d
        $v = $Drv.Copy()
        if ($null -eq $v -or (Get-Norm $v) -cne $want) { continue }
        # すぐ下の行が同じ名前なら、ここは一番下ではない (表の最終行では ↓ が効かずコピーに失敗する → 一番下とみなす)
        $Drv.Key('{DOWN}'); $below = $Drv.Copy()
        if ($null -ne $below) { $Drv.Key('{UP}') }
        if ($null -ne $below -and (Get-Norm $below) -ceq $want) { continue }
        if ($d -ne 0) { & $Log "  [補正] 行の位置が $d 行ずれていたため補正しました ($($Rows[$Target]))" }
        return
    }
    Stop-Run "行の位置がずれていて、前後 $($script:SyncRange) 行の中に「$($Rows[$Target])」が見つかりません"
}

function Get-SyncOffsets {
    # 0, -1, +1, -2, +2, ... の順 (近い行から探す)
    $o = New-Object System.Collections.Generic.List[int]
    $o.Add(0)
    for ($i = 1; $i -le $script:SyncRange; $i++) { $o.Add(-$i); $o.Add($i) }
    return , $o
}

function Test-SameValue([string]$A, [double]$B) {
    if (-not $A) { $A = '0' }
    $v = ConvertTo-Double $A
    return ($null -ne $v -and [Math]::Abs($v - $B) -lt 1e-9)
}

function Invoke-Project($Drv, $Cfg, [string]$Project, $Items, [datetime]$Today, $Log) {
    # 戻り値: Done (入力・確認できた項目) / OkAll (更新してよいか) / NotFound (見つからなかった件数)
    #   ・見つからない WBS は何も入力していないので、他の項目の更新は妨げない
    #   ・入力したのに読み返しが合わない項目があれば、更新してはいけない
    $key = Get-SearchKey $Project
    $first = Open-Grid $Drv $key $Today $Log
    if (-not (Get-Norm $first).StartsWith((Get-Norm $key))) { Stop-Run "絞り込み結果の 1 行目が $key で始まっていません: $first" }
    # 6 桁で検索すると同じ番号の別プロジェクト (例: 859314-11 と 859314-12) も出るので、目的のプロジェクト行まで進む
    while ($true) {
        $g = Read-GridRows $Drv $first $true
        if ((Get-Norm (Get-RowCode $first)) -ceq (Get-Norm $Project)) { break }
        if (-not $g.HasMore -or -not (Get-Norm $g.Next).StartsWith((Get-Norm $key))) {
            $cnt = 0
            foreach ($it in $Items) { $cnt++; & $Log "  [見つからない] $($it.Date) プロジェクト $Project が検索結果にありません → tasks.csv のプロジェクトコードを確認してください" }
            $none = New-Object System.Collections.Generic.List[object]
            return @{ Done = $none; OkAll = $true; NotFound = $cnt }
        }
        $first = $g.Next
    }
    $rows = $g.Rows; $cur = $g.Cur
    $lastRow = -1; if (-not $g.HasMore) { $lastRow = $rows.Count - 1 }    # 表の最終行 (↓ が効かない行)
    & $Log "  $first  ($($rows.Count) 行)"

    $done = New-Object System.Collections.Generic.List[object]
    $okAll = $true; $notFound = 0
    foreach ($it in $Items) {
        if ($it.Hours -gt 24) {
            # 24 時間を超える値は OBPM が受け付けない (調査 B-14)。記録の誤りの可能性が高いので入力しない
            & $Log "  [入力しない] $($it.Date) $($it.Wbs): $(Format-Num $it.Hours) 時間は 24 時間を超えています。records.csv を確認してください"
            $notFound++; continue
        }
        $target = Find-Row $rows $it.Wbs
        if ($target -lt 0) { & $Log "  [見つからない] $($it.Date) $($it.Wbs) → tasks.csv の WBS 名を確認してください"; $notFound++; continue }
        Move-Rows $Drv $cur $target
        if ($target -ne $cur) { Sync-Row $Drv $rows $target $Log }
        $cur = $target
        $day = [datetime]::ParseExact($it.Date, 'yyyy-MM-dd', $script:Inv).Day
        for ($i = 0; $i -lt $script:OffsetCols + $day; $i++) { $Drv.Key('{RIGHT}') }
        $value = Format-Num $it.Hours
        if ($Cfg.Mode -eq 'check') {
            $now = $Drv.Copy()
            & $Log "  [確認のみ] $($it.Date) $($rows[$target]): 現在 '$now' → 入力予定 $value"
            $backOk = $true
        } else {
            $Drv.TypeText($value)
            $Drv.Key($Cfg.CommitKey)
            # 確定キーで選択が動く画面なら元のセルへ戻す (表の端では動かないので戻さない)
            if ($Cfg.AfterCommit -eq 'down' -and $target -ne $lastRow) { $Drv.Key('{UP}') }
            elseif ($Cfg.AfterCommit -eq 'right' -and $day -lt 31) { $Drv.Key('{LEFT}') }
            $now = $Drv.Copy()
            $backOk = Test-SameValue $now $it.Hours
            if ($backOk) { & $Log "  [入力OK] $($it.Date) $($rows[$target]): $value" }
            else { & $Log "  [入力NG] $($it.Date) $($rows[$target]): $value (読み返し: '$now'。入力できない行の可能性)" }
        }
        for ($i = 0; $i -lt $script:OffsetCols + $day; $i++) { $Drv.Key('{LEFT}') }
        if ($backOk) { $done.Add($it) } else { $okAll = $false }
    }
    return @{ Done = $done; OkAll = $okAll; NotFound = $notFound }
}

function Invoke-Input($Cfg, $Log, $Drv = $null, $Today = $null) {
    # 戻り値: nochange / locked / needlogin / done / partial / error
    if ($null -eq $Today) { $Today = (Get-Date).Date }
    $dataDir = $Cfg.DataDir
    $statePath = Join-Path $dataDir 'obpm_input_state.json'
    $des = Get-DesiredHours (Join-Path $dataDir 'records.csv') $Today $Cfg.RoundUnit $Cfg.RoundMode $Cfg.Daily (Get-ClosedDays $dataDir)
    foreach ($w in $des.Warnings) { & $Log "[注意] $w" }
    $entered = Get-State $statePath
    $items = Get-PendingItems $des.Hours $entered $Today
    # 月末の営業日以降の日は、本人が OBPM で確認して入力するため自動入力しない
    $extra = Get-ExtraHolidays $dataDir
    $meDates = @($items | Where-Object { Test-MonthEnd ([datetime]::ParseExact($_.Date, 'yyyy-MM-dd', $script:Inv)) $extra } | ForEach-Object { $_.Date } | Sort-Object -Unique)
    if ($meDates.Count -gt 0) {
        foreach ($md in $meDates) { & $Log "[月末] $md は月末の営業日以降のため自動入力しません。OBPM でご自身で確認・入力してください" }
        $keep = New-Object System.Collections.Generic.List[object]
        foreach ($it in $items) { if ($meDates -notcontains $it.Date) { $keep.Add($it) } }
        $items = $keep
    }
    # 合計の調整の説明は、今回入力する日の分だけ出す
    foreach ($date in @($items | ForEach-Object { $_.Date } | Sort-Object -Unique)) {
        if ($des.DayNotes.ContainsKey($date)) { foreach ($n in $des.DayNotes[$date]) { & $Log "[合計] ${date}: $n" } }
    }
    if ($items.Count -eq 0) {
        & $Log '入力が必要な変更はありません。'
        Write-Status $dataDir '変更なし' '入力が必要な変更はありませんでした'
        return 'nochange'
    }
    if ($null -eq $Drv -and (Test-ScreenLocked)) {
        & $Log "画面がロックされているため保留します ($($items.Count) 件)。ロック解除後に入力します。"
        Write-Status $dataDir '保留' "デスクトップがロック中のため $($items.Count) 件を保留" $items.Count
        return 'locked'
    }
    & $Log "開始 (mode=$($Cfg.Mode)): $($items.Count) 件"
    $byProject = [ordered]@{}
    foreach ($it in $items) { if (-not $byProject.Contains($it.Project)) { $byProject[$it.Project] = New-Object System.Collections.Generic.List[object] }; $byProject[$it.Project].Add($it) }

    $totalDone = 0; $failed = $false
    try {
        if ($null -eq $Drv) {
            $Drv = Get-ReadyDriver $Cfg $Log
            if ($null -eq $Drv) {
                Write-Status $dataDir '要ログイン' "OBPM にログインし直して工数入力画面を開いてください ($($items.Count) 件を保留。ログ参照)" $items.Count
                return 'needlogin'
            }
        }
        $Drv.Activate()
        $dl = $Drv.Dialogs(); if ($dl.Count -gt 0) { Stop-Run "OBPM にメッセージが出ています。閉じてから再実行してください: $($dl -join ' | ')" }
        $ym = $Today.ToString('yyyy/MM')
        if ((Get-Norm $Drv.GetField('txtYm')) -cne $ym) { $Drv.SetField('txtYm', $ym) }
        foreach ($project in @($byProject.Keys)) {
            $its = $byProject[$project]
            & $Log "プロジェクト ${project}: $($its.Count) 件"
            $res = Invoke-Project $Drv $Cfg $project $its $Today $Log
            if ($res.NotFound -gt 0) { $failed = $true }
            if ($Cfg.Mode -eq 'check') { continue }
            if ($Cfg.Mode -eq 'manual') {
                $rest = ''; if ($byProject.Count -gt 1) { $rest = ' (残りのプロジェクトは次回)' }
                & $Log "manual モードのため、ここで止めます。画面を確認して「更新」を押してください。$rest"
                $totalDone += $res.Done.Count
                if (-not $res.OkAll) { $failed = $true }
                break
            }
            # live
            if (-not $res.OkAll) {
                & $Log '  確認できない項目があるため、このプロジェクトは更新しません。画面を確認してください。'
                $failed = $true
                break   # 未更新の入力が残っているので、次のプロジェクトの検索はしない
            }
            if ($res.Done.Count -eq 0) { & $Log '  入力した項目がないため、更新は押しません。'; continue }
            $Drv.Press('btnUpdate')
            $cr = $Drv.ConfirmUpdate()
            if ($cr -eq 'nodialog') {
                & $Log '  「更新」を押しても確認画面が出ませんでした。入力値に警告が出ている可能性があります (例: 24 時間を超える値)。画面を確認してください。'
                $failed = $true; break
            }
            if ($cr -like 'unexpected:*') { & $Log "  更新時に想定外のメッセージが出ました。画面を確認してください: $($cr.Substring(11))"; $failed = $true; break }
            if ($cr -like 'ok:*') { & $Log "  更新後のメッセージを閉じました: $($cr.Substring(3))" }
            foreach ($it in $res.Done) { $entered[$it.Key] = $it.Hours }
            Save-State $statePath $entered
            $totalDone += $res.Done.Count
            & $Log "  更新しました ($($res.Done.Count) 件)"
        }
    } catch {
        $am = Get-AbortMessage $_.Exception
        if ($null -eq $am) { throw }
        & $Log "[中断] $am"
        Write-Status $dataDir 'エラー' $am $totalDone
        return 'error'
    }
    if ($Cfg.Mode -eq 'check') { Write-Status $dataDir '確認のみ' "$($items.Count) 件の入力先を確認しました (ログ参照)" $items.Count; return 'done' }
    if ($failed) { Write-Status $dataDir '要確認' "$totalDone 件入力。確認できない項目あり (ログ参照)" $totalDone; return 'partial' }
    $msg = "$totalDone 件入力しました"; if ($Cfg.Mode -eq 'manual') { $msg += '。更新は手動で押してください' }
    Write-Status $dataDir '成功' $msg $totalDone
    return 'done'
}

function Invoke-MakeTasks($Cfg, $Log, $Drv = $null, $Today = $null) {
    if ($null -eq $Today) { $Today = (Get-Date).Date }
    try {
        if ($null -eq $Drv) { $Drv = Get-ReadyDriver $Cfg $Log; if ($null -eq $Drv) { return $null } }
        # 全行の読み取りは 1～2 分かかる。途中でフォーカスが外れた場合も、検索語なしで検索し直して最初から 1 回だけやり直す
        for ($try = 1; ; $try++) {
            $first = Open-Grid $Drv '' $Today $Log
            try {
                $g = Read-GridRows $Drv $first $false
                break
            } catch {
                $am = Get-AbortMessage $_.Exception
                if ($null -eq $am -or $try -ge 2) { throw }
                & $Log "  読み取りの途中で中断したため、検索をやり直して最初から読み直します ($am)"
                Start-Sleep -Seconds 1
            }
        }
    } catch {
        $am = Get-AbortMessage $_.Exception
        if ($null -eq $am) { throw }
        & $Log "[中断] $am"; return $null
    }
    $out = Join-Path $Cfg.DataDir 'tasks_candidates.csv'
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add((ConvertTo-CsvLine @('タスク名', 'プロジェクトコード', 'WBS')))
    foreach ($c in (Get-TaskCandidates $g.Rows)) { $lines.Add((ConvertTo-CsvLine $c)) }
    Write-TextFile $out (($lines -join "`r`n") + "`r`n")
    & $Log "$($g.Rows.Count) 行を読み取り、タスクの候補を書き出しました: $out"
    return $out
}

function Invoke-WithRunLock([scriptblock]$Block) {
    # 待機モードと手動実行が同時に OBPM を操作しないようにする
    $m = New-Object System.Threading.Mutex($false, 'Local\OBPMInputRun')
    $got = $false
    try { $got = $m.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $got = $true }
    if (-not $got) { $m.Dispose(); return 'busy' }
    try { return (& $Block) } finally { $m.ReleaseMutex(); $m.Dispose() }
}

# ==========================================================================
# 待機モード (指定時刻・「本日終了」の依頼で実行)
# ==========================================================================
function Get-DueSlots($Times, [datetime]$Now, $DoneSlots) {
    # 今日の予定時刻のうち、時刻を過ぎていてまだ実行していないもの (平日のみ)
    if ($Now.DayOfWeek -in @('Saturday', 'Sunday')) { return @() }
    $due = @()
    foreach ($t in $Times) {
        $p = $t.Split(':'); $at = $Now.Date.AddHours([int]$p[0]).AddMinutes([int]$p[1])
        $key = "$($Now.ToString('yyyy-MM-dd')) $t"
        if ($Now -ge $at -and -not $DoneSlots.Contains($key)) { $due += $key }
    }
    return $due
}

function Start-Watch($Cfg, $Log) {
    $mutex = New-Object System.Threading.Mutex($false, 'Local\OBPMInputWatch')
    if (-not $mutex.WaitOne(0)) { Write-Host '待機モードはすでに起動しています。'; return }
    $request = Join-Path $Cfg.DataDir 'input_request.json'
    $doneSlots = New-Object System.Collections.Generic.HashSet[string]
    $pending = ''                # 保留の理由: '' / 'locked' (画面ロック) / 'needlogin' (要ログイン)
    $retryAt = [datetime]::MinValue
    $lastReqCheck = [datetime]::MinValue
    $first = $true
    $plan = 'なし (「本日終了」の依頼のときだけ入力)'; if ($Cfg.Times.Count -gt 0) { $plan = $Cfg.Times -join ', ' }
    & $Log "待機モードを開始しました (予定時刻: $plan / 依頼の確認: $($Cfg.RequestCheckMin) 分ごと)"
    while ($true) {
        try {
            $now = Get-Date
            $due = @(Get-DueSlots $Cfg.Times $now $doneSlots)
            if ($first -and $due.Count -gt 1) { foreach ($s in $due[0..($due.Count - 2)]) { [void]$doneSlots.Add($s) }; $due = @($due[-1]) }  # 起動前に過ぎた分は 1 回にまとめる
            $first = $false
            $byRequest = $false
            if (($now - $lastReqCheck).TotalMinutes -ge $Cfg.RequestCheckMin) {
                $lastReqCheck = $now
                # 保留中は依頼ファイルが残っているので、新しい依頼としては数えない (保留の再確認で扱う)
                $byRequest = ($pending -eq '') -and (Test-Path -LiteralPath $request)
            }
            $retry = ($pending -ne '' -and $now -ge $retryAt)
            if ($due.Count -gt 0 -or $byRequest -or $retry) {
                foreach ($s in $due) { [void]$doneSlots.Add($s) }       # 予定時刻の分は、以後は保留としてまとめて扱う
                $skip = $false
                if (Test-ScreenLocked) {
                    if ($pending -ne 'locked') { Write-Status $Cfg.DataDir '保留' 'デスクトップがロック中のため、解除後に入力します' }
                    $pending = 'locked'; $retryAt = $now               # ロックの確認は軽いので 30 秒ごと
                    $skip = $true
                } elseif ($pending -eq 'needlogin' -and $due.Count -eq 0 -and -not $byRequest) {
                    # 要ログインで保留中の再確認は、画面の状態だけを静かに確かめる
                    Initialize-Uia
                    if (-not (Test-AutoLoginReady)) {
                        $ost = Get-ObpmState $Cfg
                        $stale = $false
                        if ($ost.State -eq 'kosu') { $od = Get-ObpmDate $ost.Win; $stale = ($null -ne $od -and $od -ne $now.Date) }
                        if ($ost.State -in @('none', 'login') -or $stale) { $retryAt = $now.AddMinutes($Cfg.RequestCheckMin); $skip = $true }
                    }
                }
                if (-not $skip) {
                    if ($byRequest) { & $Log '「本日終了」の依頼を受け付けました。' }
                    elseif ($due.Count -gt 0) { & $Log "予定時刻 ($($due -join ', ')) の実行" }
                    elseif ($pending -eq 'locked') { & $Log 'ロック解除を確認したため、保留していた入力を実行します。' }
                    else { & $Log 'OBPM の画面が使えるようになったため、保留していた入力を実行します。' }
                    $result = Invoke-WithRunLock { Invoke-Input $Cfg $Log }
                    switch ($result) {
                        'busy' { & $Log '手動の実行中のため、次の確認時に実行します。'; $pending = 'locked'; $retryAt = $now.AddMinutes(1) }
                        'locked' { $pending = 'locked'; $retryAt = $now }
                        'needlogin' { $pending = 'needlogin'; $retryAt = $now.AddMinutes($Cfg.RequestCheckMin) }   # 依頼は残す
                        default {
                            $pending = ''
                            if (Test-Path -LiteralPath $request) { Remove-Item -LiteralPath $request -Force -ErrorAction SilentlyContinue }
                        }
                    }
                }
            }
        } catch {
            & $Log "[予期しないエラー] $($_.Exception.Message)`r`n$($_.ScriptStackTrace)"
            Write-Status $Cfg.DataDir 'エラー' '予期しないエラー (ログ参照)'
        }
        Start-Sleep -Seconds 30
    }
}

# ==========================================================================
# メイン
# ==========================================================================
function Invoke-Main {
    if (-not $Config) { $Config = Join-Path $PSScriptRoot 'obpm_input.ini' }
    $startupLnk = Join-Path ([Environment]::GetFolderPath('Startup')) 'OBPM工数自動入力 (待機).lnk'
    if ($UninstallStartup) {
        if (Test-Path -LiteralPath $startupLnk) { Remove-Item -LiteralPath $startupLnk -Force; Write-Host "スタートアップから削除しました: $startupLnk" }
        else { Write-Host 'スタートアップに登録されていません。' }
        Write-Host '※ 起動中の待機モードを止めるには、PC を再起動するか、タスクマネージャーで該当の PowerShell を終了してください。'
        return
    }
    if ($SetLogin) { Set-LoginInfo; return }
    if ($ClearLogin) { Clear-LoginInfo; return }
    $cfg = Get-InputConfig $Config
    if ($InstallStartup) {
        $ws = New-Object -ComObject WScript.Shell
        $lnk = $ws.CreateShortcut($startupLnk)
        $lnk.TargetPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $lnk.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`" -Config `"$((Resolve-Path -LiteralPath $Config).Path)`" -Watch"
        $lnk.WorkingDirectory = $PSScriptRoot
        $lnk.WindowStyle = 7
        $lnk.Save()
        Write-Host "スタートアップに登録しました: $startupLnk"
        Write-Host '次回のサインインから自動で待機モードが起動します。今すぐ起動する場合は obpm_input.bat の 3 を選んでください。'
        return
    }
    if (-not (Test-Path -LiteralPath $cfg.DataDir)) { throw "共有フォルダが見つかりません: $($cfg.DataDir)" }
    if ($Watch) { Start-Watch $cfg (New-Logger $cfg.DataDir -Quiet); return }

    $log = New-Logger $cfg.DataDir
    $r = Invoke-WithRunLock {
        if ($MakeTasks) { $o = Invoke-MakeTasks $cfg $log; if ($o) { Write-Host "候補: $o" } }
        else { [void](Invoke-Input $cfg $log) }
    }
    if ($r -eq 'busy') { Write-Host '別の自動入力が実行中のため終了します。' }
}

if ($MyInvocation.InvocationName -ne '.') {
    try { Invoke-Main } catch {
        Write-Host "エラー: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}
