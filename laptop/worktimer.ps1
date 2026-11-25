# -*- coding: utf-8 -*-
<#
worktimer.ps1  作業タイマー (ノート PC 側。Windows PowerShell 5.1 で動作)

  ・タスク一覧 (tasks.csv) からタスクを選び、「作業開始」「作業終了」で作業時間を記録
  ・作業中に別のタスクを開始すると、前のタスクは自動で終了
  ・一定時間 PC 操作がなければ、最後に操作した時刻で自動停止
  ・シャットダウンなどで止め損ねた記録は、次回起動時に最後の確認時刻で締める
  ・日をまたいだ作業は日付ごとに分けて記録
  ・「集計」で日付×タスクごとの工数 (丸め後) を表示・CSV 保存
  ・「本日終了」で作業を止め、デスクトップに工数の入力を依頼
  ・デスクトップの自動入力 (obpm_input.ps1) の前回の結果を表示

起動: worktimer.bat をダブルクリック (同じフォルダの worktimer.ini を使用)
※ このファイルは「UTF-8 (BOM 付き)」で保存すること
#>
param([string]$Config = '')

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'common.ps1')
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$AppName = '作業タイマー'

$ConfigTemplate = @'
; worktimer 設定ファイル (; で始まる行はコメント。パスは " で囲まずそのまま書く)

[files]
; タスク一覧の CSV (相対パスはこの ini のフォルダ基準)
tasks_csv = tasks.csv
; 作業記録・集計結果の保存先フォルダ (相対パスはこの ini のフォルダ基準)
; デスクトップで自動入力する場合は、両方の PC から見えるファイルサーバーのフォルダにする
;   例: data_dir = \\fileserver\share\worktimer\data
data_dir = data

[timer]
; この分数だけ PC 操作がなければ、最後に操作した時刻で自動停止 (0 = 自動停止しない)
idle_stop_min = 10
; true = タイマー画面を常に手前に表示
always_on_top = true
; true = 画面の右下に小さい画面 (タスク選択・開始・停止) を常に表示する。大きい画面は小さい画面の「画面」ボタンで開閉
mini_window = true
; 小さい画面にマウスが乗っていないときの不透明度 (0.2～1.0。1.0 = 透けない)
mini_opacity = 0.5

[summary]
; 工数の丸め単位 (分)。OBPM の 0.25 時間刻みに合わせるなら 15
round_unit_min = 15
; 丸め方: nearest = 四捨五入 / up = 切り上げ / down = 切り捨て
round_mode = nearest

[daily]
; 1 日の定時の工数 (時間)。「本日終了」のとき、作業時間の合計がこれに足りなければ確認する
base_hours = 7.75

[obpm]
; true = タイマー起動時に SI Object Browser PM が起動していなければ起動する
launch_obpm = false
; SI Object Browser PM の実行ファイルのパス (スタートメニューのアイコンを右クリック →「ファイルの場所を開く」で確認)
obpm_exe =
'@

$TasksTemplate = @'
タスク名,プロジェクトコード,WBS
(例) 総合事業マスター検索 調査,859314-12,Ver21.0.4 総合事業マスター検索が意図通りに動かない
(例) 介給 ID:0526 PG・単体テスト,859314-12,【障害】PG・単体テスト > Ver21.0.4 介給 ID：0526 Ora-00600
(例) 在宅勤務朝礼,869114-13,在宅勤務朝礼
'@

function Show-Error([string]$Msg) {
    [void][System.Windows.Forms.MessageBox]::Show($Msg, $AppName, 'OK', 'Error')
}

# --------------------------------------------------------------------------
# 設定・タスク一覧
# --------------------------------------------------------------------------
if (-not $Config) { $Config = Join-Path $PSScriptRoot 'worktimer.ini' }
try {
    if (-not (Test-Path -LiteralPath $Config)) { Write-TextFile $Config ($ConfigTemplate -replace "`r?`n", "`r`n") }
    $ini = Read-Ini $Config
    $base = Split-Path -Parent (Resolve-Path -LiteralPath $Config)
    $Cfg = @{
        TasksCsv     = Resolve-RelPath $base (Get-IniValue $ini 'files' 'tasks_csv' 'tasks.csv')
        DataDir      = Resolve-RelPath $base (Get-IniValue $ini 'files' 'data_dir' 'data')
        IdleStopMin  = [int](Get-IniValue $ini 'timer' 'idle_stop_min' '10')
        AlwaysOnTop  = (Get-IniValue $ini 'timer' 'always_on_top' 'true') -match '^(true|1|yes|on)$'
        MiniWindow   = (Get-IniValue $ini 'timer' 'mini_window' 'true') -match '^(true|1|yes|on)$'
        MiniOpacity  = [Math]::Min(1.0, [Math]::Max(0.2, [double]::Parse((Get-IniValue $ini 'timer' 'mini_opacity' '0.5'), [System.Globalization.CultureInfo]::InvariantCulture)))
        RoundUnit    = [Math]::Max(1, [int](Get-IniValue $ini 'summary' 'round_unit_min' '15'))
        RoundMode    = (Get-IniValue $ini 'summary' 'round_mode' 'nearest').ToLower()
        BaseHours    = [double]::Parse((Get-IniValue $ini 'daily' 'base_hours' '7.75'), [System.Globalization.CultureInfo]::InvariantCulture)
        LaunchObpm   = (Get-IniValue $ini 'obpm' 'launch_obpm' 'false') -match '^(true|1|yes|on)$'
        ObpmExe      = Get-IniValue $ini 'obpm' 'obpm_exe' ''
    }
    if ($Cfg.RoundMode -notin @('nearest', 'up', 'down')) { throw "[summary] round_mode は nearest / up / down で指定してください (現在: $($Cfg.RoundMode))" }
} catch {
    Write-ErrorLog '' $_.ToString()
    Show-Error "設定ファイルを読み込めません:`n$($_.Exception.Message)"
    exit 1
}

function Get-Tasks {
    if (-not (Test-Path -LiteralPath $Cfg.TasksCsv)) { Write-TextFile $Cfg.TasksCsv (($TasksTemplate -replace "`r?`n", "`r`n") + "`r`n") }
    $list = New-Object System.Collections.Generic.List[object]
    $rows = Read-CsvRows $Cfg.TasksCsv
    for ($i = 1; $i -lt $rows.Count; $i++) {
        $r = @($rows[$i]) + @('', '', '')
        if ($r[0].Trim()) { $list.Add(@{ name = $r[0].Trim(); project = $r[1].Trim(); wbs = $r[2].Trim() }) }
    }
    return , $list
}

# --------------------------------------------------------------------------
# 画面
# --------------------------------------------------------------------------
$script:Rec = New-Recorder $Cfg.DataDir
$script:Tasks = New-Object System.Collections.Generic.List[object]
$script:IdleNotice = $null
$script:LastHeartbeat = Get-Date
$script:LastMinute = ''

$font = New-Object System.Drawing.Font('Yu Gothic UI', 9)
$form = New-Object System.Windows.Forms.Form
$form.Text = $AppName
$form.Font = $font
$form.AutoScaleMode = 'Dpi'
$form.ClientSize = New-Object System.Drawing.Size(450, 390)
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.TopMost = $Cfg.AlwaysOnTop
$form.StartPosition = 'Manual'
$form.Location = New-Object System.Drawing.Point(([System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Right - 480), 40)

function New-Ctl($Type, $X, $Y, $W, $H, $Text = '') {
    $c = New-Object $Type
    $c.Location = New-Object System.Drawing.Point($X, $Y)
    $c.Size = New-Object System.Drawing.Size($W, $H)
    if ($Text) { $c.Text = $Text }
    $form.Controls.Add($c)
    return $c
}

[void](New-Ctl System.Windows.Forms.Label 10 13 40 20 'タスク')
$combo = New-Ctl System.Windows.Forms.ComboBox 52 9 312 24
$combo.DropDownStyle = 'DropDownList'
$btnReload = New-Ctl System.Windows.Forms.Button 370 8 70 26 '再読込'
$lblNow = New-Ctl System.Windows.Forms.Label 10 42 430 20 '停止中'
$lblElapsed = New-Ctl System.Windows.Forms.Label 10 62 430 36 '0:00:00'
$lblElapsed.Font = New-Object System.Drawing.Font('Yu Gothic UI', 18, [System.Drawing.FontStyle]::Bold)
$btnStart = New-Ctl System.Windows.Forms.Button 10 102 140 30 '作業開始'
$btnStop = New-Ctl System.Windows.Forms.Button 155 102 140 30 '作業終了'
$btnSum = New-Ctl System.Windows.Forms.Button 300 102 68 30 '集計'
$btnManual = New-Ctl System.Windows.Forms.Button 372 102 68 30 '記録修正'
$btnDayEnd = New-Ctl System.Windows.Forms.Button 10 137 430 28 '本日終了（作業を止めて、デスクトップに工数の入力を依頼）'
[void](New-Ctl System.Windows.Forms.Label 10 172 200 18 '今日の作業')
$list = New-Ctl System.Windows.Forms.ListView 10 191 430 112
$list.View = 'Details'; $list.FullRowSelect = $true; $list.HeaderStyle = 'Nonclickable'
[void]$list.Columns.Add('タスク', 340)
[void]$list.Columns.Add('時間', 80, 'Right')
$lblStatus = New-Ctl System.Windows.Forms.Label 10 308 430 46 ''
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(170, 50, 50)
$lblAuto = New-Ctl System.Windows.Forms.Label 10 356 430 30 ''
$lblAuto.ForeColor = [System.Drawing.Color]::DimGray

function Set-Status([string]$Msg) { $lblStatus.Text = $(if ($Msg) { $Msg } else { '' }) }

function Invoke-Safe([scriptblock]$Block) {
    try { & $Block } catch {
        Write-ErrorLog $Cfg.DataDir ($_.ToString() + "`r`n" + $_.ScriptStackTrace)
        Set-Status "エラーが発生しました (data フォルダの error.log を確認): $($_.Exception.Message)"
    }
}

function Get-TaskKey($Name, $Project, $Wbs) {
    # タスク名を変えても並び順が引き継がれるよう、プロジェクトコード + WBS で見分ける (どちらかが空ならタスク名)
    if ($Project -and $Wbs) { return "$(Get-Norm $Project)|$(Get-Norm $Wbs)" }
    return "名前:$(Get-Norm $Name)"
}

function Get-LastUsed {
    # 作業記録 (records.csv) から、タスクごとに最後に作業を開始した日時を集める
    $last = @{}
    if (-not (Test-Path -LiteralPath $script:Rec.Records)) { return $last }
    $rows = Read-CsvRows $script:Rec.Records
    for ($i = 1; $i -lt $rows.Count; $i++) {
        $r = @($rows[$i])
        if ($r.Count -lt 7) { continue }
        $t = [datetime]::MinValue
        if (-not [datetime]::TryParse($r[1], [ref]$t) -and -not [datetime]::TryParse($r[0], [ref]$t)) { continue }   # Excel で編集された形式も読む
        $k = Get-TaskKey $r[4] $r[5].Trim() $r[6].Trim()
        if (-not $last.ContainsKey($k) -or $t -gt $last[$k]) { $last[$k] = $t }
    }
    $cur = $script:Rec.Current
    if ($null -ne $cur) { $last[(Get-TaskKey $cur.task.name $cur.task.project $cur.task.wbs)] = Get-Date }
    return $last
}

function Update-TaskList([string]$Select = '') {
    # タスクは最近作業した順に並べる (一度も作業していないタスクは tasks.csv の順で後ろに)
    try { $loaded = Get-Tasks } catch {
        $loaded = New-Object System.Collections.Generic.List[object]
        Set-Status "タスク一覧を読み込めません: $($_.Exception.Message)"
    }
    $last = @{}
    try { $last = Get-LastUsed } catch { Write-ErrorLog $Cfg.DataDir ("並べ替え用の作業記録を読めません: " + $_.ToString()) }
    $keyed = for ($i = 0; $i -lt $loaded.Count; $i++) {
        $t = $loaded[$i]; $k = Get-TaskKey $t.name $t.project $t.wbs
        $when = [datetime]::MinValue; if ($last.ContainsKey($k)) { $when = $last[$k] }
        [pscustomobject]@{ When = $when; Index = $i; Task = $t }
    }
    $script:Tasks = New-Object System.Collections.Generic.List[object]
    if ($null -ne $keyed) { foreach ($x in @($keyed | Sort-Object @{ Expression = 'When'; Descending = $true }, @{ Expression = 'Index'; Ascending = $true })) { $script:Tasks.Add($x.Task) } }
    $combo.BeginUpdate(); $combo.Items.Clear()
    foreach ($t in $script:Tasks) { [void]$combo.Items.Add($t.name) }
    $combo.EndUpdate()
    if (-not $Select -and $null -ne $script:Rec.Current) { $Select = $script:Rec.Current.task.name }
    if ($Select) { $combo.SelectedIndex = $combo.Items.IndexOf($Select) }
    if (Get-Command Update-MiniView -ErrorAction SilentlyContinue) { Update-MiniView }
    if ($script:Tasks.Count -eq 0) { Set-Status "タスク一覧が空です。$($Cfg.TasksCsv) を編集して「再読込」を押してください" }
}

function Update-Today {
    $today = (Get-Date).ToString('yyyy-MM-dd')
    $acc = @{}
    foreach ($r in (Read-Records $script:Rec.Records)) {
        if ($r.date -eq $today) { if (-not $acc.ContainsKey($r.name)) { $acc[$r.name] = 0.0 }; $acc[$r.name] += $r.minutes }
    }
    $cur = $script:Rec.Current
    if ($null -ne $cur) {
        $start = $cur.start; if ($start -lt (Get-Date).Date) { $start = (Get-Date).Date }
        $n = $cur.task.name; if (-not $acc.ContainsKey($n)) { $acc[$n] = 0.0 }
        $acc[$n] += ((Get-Date) - $start).TotalMinutes
    }
    $list.BeginUpdate(); $list.Items.Clear()
    foreach ($e in ($acc.GetEnumerator() | Sort-Object Value -Descending)) {
        $item = New-Object System.Windows.Forms.ListViewItem($e.Key)
        [void]$item.SubItems.Add((Format-HM $e.Value)); [void]$list.Items.Add($item)
    }
    $list.EndUpdate()
    # デスクトップの自動入力の前回の結果
    $st = Read-JsonFile (Join-Path $Cfg.DataDir 'obpm_input_status.json')
    $lblAuto.Text = ''
    if ($null -ne $st -and $st.PSObject.Properties['time'] -and ([string]$st.time).Length -ge 6) {
        $lblAuto.Text = "自動入力 $(([string]$st.time).Substring(5)): $($st.result) - $($st.message)"
    }
}

function Reset-Display {
    $lblNow.Text = '停止中'; $lblElapsed.Text = '0:00:00'; Update-Today
}

$btnReload.Add_Click({ Invoke-Safe { Update-TaskList } })

function Invoke-StartWork {
    $i = $combo.SelectedIndex
    if ($i -lt 0 -or $i -ge $script:Tasks.Count) { Set-Status 'タスクを選んでください'; return }
    $task = $script:Tasks[$i]
    $msg = Start-Work $script:Rec $task (Get-Date)
    Update-TaskList $task.name    # 開始したタスクを一番上へ
    $lblNow.Text = "作業中: $($task.name)  ($($script:Rec.Current.start.ToString('HH:mm')) 開始)"
    Set-Status $msg
    Update-Today
}

function Invoke-StopWork {
    if ($null -eq $script:Rec.Current) { Set-Status '作業中のタスクはありません'; return }
    $name = $script:Rec.Current.task.name
    Stop-Work $script:Rec (Get-Date) '作業終了'
    Set-Status "「$name」を終了しました"
    Reset-Display
}

$btnStart.Add_Click({ Invoke-Safe { Invoke-StartWork } })
$btnStop.Add_Click({ Invoke-Safe { Invoke-StopWork } })

# --------------------------------------------------------------------------
# 本日終了・入力の依頼
# --------------------------------------------------------------------------
function Get-DayRecords {
    # 作業記録を 日付 → @{ Min = 合計分; LastEnd = 最後の終了時刻 } にまとめる (作業中のタスクも含む)
    $out = @{}
    if (Test-Path -LiteralPath $script:Rec.Records) {
        $rows = Read-CsvRows $script:Rec.Records
        for ($i = 1; $i -lt $rows.Count; $i++) {
            $r = @($rows[$i]); if ($r.Count -lt 4) { continue }
            $d = [datetime]::MinValue
            if (-not [datetime]::TryParse($r[0], [ref]$d)) { continue }
            $k = $d.ToString('yyyy-MM-dd')
            $m = 0.0; [void][double]::TryParse($r[3], [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$m)
            $e = [datetime]::MinValue; [void][datetime]::TryParse($r[2], [ref]$e)
            if (-not $out.ContainsKey($k)) { $out[$k] = @{ Min = 0.0; LastEnd = [datetime]::MinValue } }
            $out[$k].Min += $m
            if ($e -gt $out[$k].LastEnd) { $out[$k].LastEnd = $e }
        }
    }
    $cur = $script:Rec.Current
    if ($null -ne $cur) {
        $now = Get-Date; $k = $now.ToString('yyyy-MM-dd')
        if (-not $out.ContainsKey($k)) { $out[$k] = @{ Min = 0.0; LastEnd = [datetime]::MinValue } }
        $out[$k].Min += ($now - [datetime]$cur.start).TotalMinutes; $out[$k].LastEnd = $now
    }
    return $out
}

function Get-ClosedDays {
    $cd = [ordered]@{}
    $old = Read-JsonFile (Join-Path $Cfg.DataDir 'closed_days.json')
    if ($null -ne $old) { foreach ($p in @($old.PSObject.Properties | Sort-Object Name | Select-Object -Last 60)) { $cd[$p.Name] = [string]$p.Value } }
    return $cd
}

function Get-UnclosedDays([datetime]$From, [datetime]$To) {
    # 記録があるのに「本日終了」が押されていない日 (押した後にまた作業した日も含む)
    $days = Get-DayRecords; $cd = Get-ClosedDays
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($k in ($days.Keys | Sort-Object)) {
        $d = [datetime]::ParseExact($k, 'yyyy-MM-dd', $null)
        if ($d -lt $From.Date -or $d -gt $To.Date -or $days[$k].Min -lt 1) { continue }
        $closedAt = [datetime]::MinValue
        if ($cd.Contains($k)) { [void][datetime]::TryParse($cd[$k], [ref]$closedAt) }
        $running = ($null -ne $script:Rec.Current -and $k -eq (Get-Date).ToString('yyyy-MM-dd'))   # 作業中なら本日終了はまだ
        if ($running -or $days[$k].LastEnd -gt $closedAt.AddMinutes(1)) { $out.Add($k) }
    }
    return , $out
}

function Send-InputRequest([string[]]$CloseDates) {
    # デスクトップに入力を依頼し、指定の日を「終わった日」として記録する (デスクトップはその日の合計を社内ルールに合わせて入力する)
    $nowS = (Get-Date).ToString($script:TimeFmt)
    $cd = Get-ClosedDays
    foreach ($d in $CloseDates) { $cd[$d] = $nowS }
    Write-JsonFile (Join-Path $Cfg.DataDir 'closed_days.json') $cd
    Write-JsonFile (Join-Path $Cfg.DataDir 'input_request.json') ([ordered]@{ requested_at = $nowS })
}

function Invoke-DayEnd([bool]$Ask) {
    # 戻り値: $true = 本日終了した / $false = やめた
    $today = (Get-Date).Date
    $min = 0.0; $days = Get-DayRecords; $k = $today.ToString('yyyy-MM-dd'); if ($days.ContainsKey($k)) { $min = $days[$k].Min }
    $q = '本日の作業を終了し、デスクトップに工数の入力を依頼しますか？'
    $icon = 'Question'
    if ($min -lt $Cfg.BaseHours * 60 - 1e-6) {
        $q = "本日の作業時間の合計は $(Format-HM $min) です。$($Cfg.BaseHours)H に足りませんが、本日終了しますか？`n（早退などで問題なければ「はい」）"
        $icon = 'Warning'
    }
    if ($Ask -or $icon -eq 'Warning') {
        if ([System.Windows.Forms.MessageBox]::Show($q, $AppName, 'YesNo', $icon) -ne 'Yes') { return $false }
    }
    if ($null -ne $script:Rec.Current) { Stop-Work $script:Rec (Get-Date) '本日終了'; Reset-Display }
    $msg = 'デスクトップに入力を依頼しました。数分以内に入力されます（結果は下に表示）'
    try {
        if (Test-MonthEnd $today (Get-ExtraHolidays $Cfg.DataDir)) {
            [void][System.Windows.Forms.MessageBox]::Show("本日は月末（月の最後の営業日以降）のため、本日分は自動入力しません。`nOBPM でご自身で確認・入力してください。`n（前日までの未入力分は自動で入力します）", $AppName, 'OK', 'Information')
            $msg = '本日は月末のため本日分は自動入力しません。前日までの未入力分の入力を依頼しました'
        }
        Send-InputRequest @($k)
        Set-Status $msg
    } catch {
        Set-Status "入力の依頼を書き込めませんでした (共有フォルダを確認): $($_.Exception.Message)"
    }
    return $true
}

$btnDayEnd.Add_Click({ Invoke-Safe { [void](Invoke-DayEnd $true) } })

function Show-StartupNotices {
    # 起動時: ① 本日終了が押されていない日 ② 前回の自動入力が完了していない場合 に確認する
    $today = (Get-Date).Date
    $first = Get-Date -Year $today.Year -Month $today.Month -Day 1
    $un = Get-UnclosedDays $first.Date $today.AddDays(-1)
    $requested = $false
    if ($un.Count -gt 0) {
        $list = ($un | ForEach-Object { ([datetime]::ParseExact($_, 'yyyy-MM-dd', $null)).ToString('M/d (ddd)') }) -join ', '
        $ans = [System.Windows.Forms.MessageBox]::Show("次の日は「本日終了」が押されていません:`n$list`n`nデスクトップに工数の入力を依頼しますか？`n（いいえ の場合も、次に「本日終了」を押したときにまとめて入力されます）", $AppName, 'YesNo', 'Warning')
        if ($ans -eq 'Yes') { Send-InputRequest $un; $requested = $true; Set-Status 'デスクトップに入力を依頼しました' }
        else {
            # 次回から聞かないよう「終わった日」として記録だけしておく
            $cd = Get-ClosedDays; foreach ($d in $un) { $cd[$d] = (Get-Date).ToString($script:TimeFmt) }
            Write-JsonFile (Join-Path $Cfg.DataDir 'closed_days.json') $cd
        }
    }
    # 前回の自動入力の結果 (成功・変更なし・確認のみ 以外で、まだ確認していないもの)
    $st = Read-JsonFile (Join-Path $Cfg.DataDir 'obpm_input_status.json')
    if ($requested -or $null -eq $st -or -not $st.PSObject.Properties['result']) { return }
    if ([string]$st.result -in @('成功', '変更なし', '確認のみ')) { return }
    if (Test-Path -LiteralPath (Join-Path $Cfg.DataDir 'input_request.json')) { return }   # まだ依頼が残っている (デスクトップが処理待ち)
    $ackPath = Join-Path $Cfg.DataDir 'timer_notice.json'
    $ack = Read-JsonFile $ackPath
    if ($null -ne $ack -and $ack.PSObject.Properties['status_time'] -and [string]$ack.status_time -eq [string]$st.time) { return }
    $ans = [System.Windows.Forms.MessageBox]::Show("前回のデスクトップの自動入力が完了していません。`n（$($st.time)  $($st.result): $($st.message)）`n`nもう一度入力を依頼しますか？", $AppName, 'YesNo', 'Warning')
    Write-JsonFile $ackPath ([ordered]@{ status_time = [string]$st.time })
    if ($ans -eq 'Yes') { Send-InputRequest @(); Set-Status 'デスクトップに再度入力を依頼しました' }
}

function Test-AppExit($e) {
    # アプリを終了してよいか確かめる (だめなら $e.Cancel = $true)
    try {
        # シャットダウン・サインアウト時は止めない (確認を出すと電源が切れなくなる)。本日終了の押し忘れは次回起動時に知らせる
        if ($e.CloseReason -in @('WindowsShutDown', 'TaskManagerClosing')) {
            if ($null -ne $script:Rec.Current) { Update-Heartbeat $script:Rec (Get-Date) }
            return
        }
        $today = (Get-Date).Date
        if ((Get-UnclosedDays $today $today).Count -gt 0) {
            $ans = [System.Windows.Forms.MessageBox]::Show("「本日終了」が押されていません。`n`nはい = 本日終了してから閉じる`nいいえ = そのまま閉じる（次に本日終了を押したときにまとめて入力されます）`nキャンセル = 閉じない", $AppName, 'YesNoCancel', 'Warning')
            if ($ans -eq 'Cancel') { $e.Cancel = $true; return }
            if ($ans -eq 'Yes') { if (-not (Invoke-DayEnd $false)) { $e.Cancel = $true }; return }
        }
        if ($null -ne $script:Rec.Current) {
            $ans = [System.Windows.Forms.MessageBox]::Show("作業中のタスクがあります。終了して記録してから閉じますか？`n（いいえ = 記録を続けたまま閉じ、次回起動時に閉じた時刻で締めます）", $AppName, 'YesNoCancel', 'Question')
            if ($ans -eq 'Cancel') { $e.Cancel = $true; return }
            if ($ans -eq 'Yes') { Stop-Work $script:Rec (Get-Date) 'アプリ終了' } else { Update-Heartbeat $script:Rec (Get-Date) }
        }
    } catch { Write-ErrorLog $Cfg.DataDir $_.ToString() }
}

$form.Add_FormClosing({
    param($s, $e)
    # 小さい画面を使うときは、大きい画面の × は「隠す」だけ (終了は小さい画面の ×)
    if ($Cfg.MiniWindow -and $e.CloseReason -eq 'UserClosing' -and -not $script:Exiting) { $e.Cancel = $true; $form.Hide(); return }
    if (-not $Cfg.MiniWindow) { Test-AppExit $e }
})

# --------------------------------------------------------------------------
# 小さい画面 (画面の右下に常に表示。タスク選択・開始・停止)
# --------------------------------------------------------------------------
$script:Exiting = $false
$script:SyncingCombo = $false
$mini = $null
if ($Cfg.MiniWindow) {
    $mini = New-Object System.Windows.Forms.Form
    $mini.Text = '停止中'; $mini.Font = $font; $mini.AutoScaleMode = 'Dpi'
    $mini.FormBorderStyle = 'FixedToolWindow'; $mini.TopMost = $true; $mini.ShowInTaskbar = $true
    $mini.ClientSize = New-Object System.Drawing.Size(300, 66)
    $mini.StartPosition = 'Manual'
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $mini.Location = New-Object System.Drawing.Point(($wa.Right - $mini.Width - 8), ($wa.Bottom - $mini.Height - 8))
    $miniCombo = New-Object System.Windows.Forms.ComboBox
    $miniCombo.DropDownStyle = 'DropDownList'; $miniCombo.Location = New-Object System.Drawing.Point(6, 6); $miniCombo.Size = New-Object System.Drawing.Size(288, 24)
    $miniCombo.DropDownWidth = 480
    $mini.Controls.Add($miniCombo)
    $miniStart = New-Object System.Windows.Forms.Button; $miniStart.Text = '開始'; $miniStart.Location = New-Object System.Drawing.Point(6, 35); $miniStart.Size = New-Object System.Drawing.Size(70, 26)
    $miniStop = New-Object System.Windows.Forms.Button; $miniStop.Text = '停止'; $miniStop.Location = New-Object System.Drawing.Point(80, 35); $miniStop.Size = New-Object System.Drawing.Size(70, 26)
    $miniMain = New-Object System.Windows.Forms.Button; $miniMain.Text = '画面'; $miniMain.Location = New-Object System.Drawing.Point(224, 35); $miniMain.Size = New-Object System.Drawing.Size(70, 26)
    $miniTime = New-Object System.Windows.Forms.Label; $miniTime.Text = ''; $miniTime.Location = New-Object System.Drawing.Point(154, 40); $miniTime.Size = New-Object System.Drawing.Size(68, 20)
    $miniTime.TextAlign = 'MiddleCenter'
    foreach ($c in @($miniStart, $miniStop, $miniMain, $miniTime)) { $mini.Controls.Add($c) }
    # 選択の同期 (小さい画面 ⇔ 大きい画面)
    function Sync-ComboIndex($From, $To) {
        if ($script:SyncingCombo) { return }
        $script:SyncingCombo = $true
        try { $i = $From.SelectedIndex; if ($i -lt $To.Items.Count) { $To.SelectedIndex = $i } } catch { } finally { $script:SyncingCombo = $false }
    }
    $miniCombo.Add_SelectedIndexChanged({ Sync-ComboIndex $miniCombo $combo })
    $combo.Add_SelectedIndexChanged({ Sync-ComboIndex $combo $miniCombo })
    $miniStart.Add_Click({ Invoke-Safe { Invoke-StartWork } })
    $miniStop.Add_Click({ Invoke-Safe { Invoke-StopWork } })
    $miniMain.Add_Click({ if ($form.Visible) { $form.Hide() } else { $form.Show(); $form.Activate() } })
    $mini.Add_FormClosing({
        param($s, $e)
        Test-AppExit $e
        if (-not $e.Cancel) { $script:Exiting = $true; $form.Close() }
    })
    # マウスが乗っていない・操作していないときは半透明にする
    $fade = New-Object System.Windows.Forms.Timer
    $fade.Interval = 250
    $fade.Add_Tick({
        try {
            $over = $mini.Bounds.Contains([System.Windows.Forms.Cursor]::Position) -or $mini.ContainsFocus -or $miniCombo.DroppedDown
            $want = $(if ($over) { 1.0 } else { $Cfg.MiniOpacity })
            if ([Math]::Abs($mini.Opacity - $want) -gt 0.01) { $mini.Opacity = $want }
        } catch { }
    })
}

function Update-MiniView {
    # 小さい画面のタスク一覧・表示を大きい画面に合わせる
    if ($null -eq $mini) { return }
    $script:SyncingCombo = $true
    try {
        $miniCombo.BeginUpdate(); $miniCombo.Items.Clear()
        foreach ($x in $combo.Items) { [void]$miniCombo.Items.Add($x) }
        $miniCombo.EndUpdate()
        if ($combo.SelectedIndex -lt $miniCombo.Items.Count) { $miniCombo.SelectedIndex = $combo.SelectedIndex }
    } catch { } finally { $script:SyncingCombo = $false }
}

# --- 1 秒ごと: 経過時間・アイドル判定・生存記録 ---
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000
$timer.Add_Tick({ Invoke-Safe {
    $now = Get-Date
    $cur = $script:Rec.Current
    if ($null -ne $cur) {
        $el = $now - $cur.start
        $lblElapsed.Text = ('{0}:{1:00}:{2:00}' -f [Math]::Floor($el.TotalHours), $el.Minutes, $el.Seconds)
        $limit = $Cfg.IdleStopMin * 60
        $idle = Get-IdleSeconds
        if ($limit -gt 0 -and $idle -ge $limit) {
            $last = $now.AddSeconds(-$idle)
            $name = $cur.task.name
            Stop-Work $script:Rec $last "$($Cfg.IdleStopMin)分操作なしで自動停止"
            $script:IdleNotice = "操作がなかったため「$name」を $($last.ToString('HH:mm')) で自動停止しました。再開するときは「作業開始」を押してください"
            Reset-Display
        } elseif (($now - $script:LastHeartbeat).TotalSeconds -ge 30) {
            Update-Heartbeat $script:Rec $now
            $script:LastHeartbeat = $now
        }
    }
    if ($null -ne $mini) {
        $cur = $script:Rec.Current
        if ($null -ne $cur) { $miniTime.Text = $lblElapsed.Text; $t = "作業中: $($cur.task.name)" } else { $miniTime.Text = ''; $t = '停止中' }
        if ($mini.Text -ne $t) { $mini.Text = $t }
    }
    if ($script:IdleNotice -and (Get-IdleSeconds) -lt 5) { Set-Status $script:IdleNotice; $script:IdleNotice = $null }
    $minute = $now.ToString('HH:mm')
    if ($minute -ne $script:LastMinute) { $script:LastMinute = $minute; Update-Today }
} })

# --- 集計画面 ---
function New-Pt($X, $Y) { New-Object System.Drawing.Point($X, $Y) }
function New-Sz($W, $H) { New-Object System.Drawing.Size($W, $H) }
function Add-To($Parent, $Ctl, $X, $Y, $W, $H, $Text = '') {
    $Ctl.Location = New-Pt $X $Y; $Ctl.Size = New-Sz $W $H
    if ($Text) { $Ctl.Text = $Text }
    $Parent.Controls.Add($Ctl); return $Ctl
}

function Update-SummaryView {
    $d1 = $script:SumFrom.Value.ToString('yyyy-MM-dd'); $d2 = $script:SumTo.Value.ToString('yyyy-MM-dd')
    $rows = Get-Summary (Read-Records $script:Rec.Records) $d1 $d2 $Cfg.RoundUnit $Cfg.RoundMode
    $script:SumRows = $rows
    $lv = $script:SumList
    $lv.BeginUpdate(); $lv.Items.Clear()
    $daily = [ordered]@{}
    foreach ($r in $rows) {
        $it = New-Object System.Windows.Forms.ListViewItem($r.date)
        foreach ($v in @($r.project, $r.wbs, $r.name, (Format-HM $r.minutes), $r.hours.ToString('0.00', $script:Inv))) { [void]$it.SubItems.Add([string]$v) }
        [void]$lv.Items.Add($it)
        if (-not $daily.Contains($r.date)) { $daily[$r.date] = 0.0 }
        $daily[$r.date] += $r.hours
    }
    $lv.EndUpdate()
    $notes = foreach ($k in $daily.Keys) { "$k : 計 $($daily[$k].ToString('0.00', $script:Inv))h" }
    $note = (@($notes) -join '  /  ')
    if (-not $note) { $note = 'この期間の記録はありません' }
    if ($null -ne $script:Rec.Current) { $note += '（作業中のタスクは含みません）' }
    $script:SumMsg.Text = $note
}

function Save-SummaryCsv {
    if ($null -eq $script:SumRows -or $script:SumRows.Count -eq 0) { $script:SumMsg.Text = '保存する記録がありません (先に「集計」を押してください)'; return }
    $name = 'summary_{0}_{1}.csv' -f $script:SumFrom.Value.ToString('yyyyMMdd'), $script:SumTo.Value.ToString('yyyyMMdd')
    $path = Join-Path $Cfg.DataDir $name
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add((ConvertTo-CsvLine @('日付', 'プロジェクトコード', 'WBS', 'タスク名', '実績(分)', '工数(時間)')))
    foreach ($r in $script:SumRows) {
        $lines.Add((ConvertTo-CsvLine @($r.date, $r.project, $r.wbs, $r.name, $r.minutes.ToString('0.0', $script:Inv), $r.hours.ToString('0.00', $script:Inv))))
    }
    try { Write-TextFile $path (($lines -join "`r`n") + "`r`n") } catch { $script:SumMsg.Text = "保存できません。$name を Excel で開いている場合は閉じてください"; return }
    $script:SumMsg.Text = "保存しました: $path"
    Start-Process -FilePath $path
}

$btnSum.Add_Click({ Invoke-Safe {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = '工数の集計'; $f.Font = $font; $f.AutoScaleMode = 'Dpi'; $f.TopMost = $Cfg.AlwaysOnTop
    $f.ClientSize = New-Sz 820 420; $f.StartPosition = 'CenterScreen'
    [void](Add-To $f (New-Object System.Windows.Forms.Label) 10 14 35 20 '期間')
    $script:SumFrom = Add-To $f (New-Object System.Windows.Forms.DateTimePicker) 48 10 110 24
    [void](Add-To $f (New-Object System.Windows.Forms.Label) 162 14 20 20 '～')
    $script:SumTo = Add-To $f (New-Object System.Windows.Forms.DateTimePicker) 185 10 110 24
    foreach ($d in @($script:SumFrom, $script:SumTo)) { $d.Format = 'Custom'; $d.CustomFormat = 'yyyy-MM-dd' }
    $modeJp = @{ nearest = '四捨五入'; up = '切り上げ'; down = '切り捨て' }[$Cfg.RoundMode]
    [void](Add-To $f (New-Object System.Windows.Forms.Label) 305 14 250 20 "（$($Cfg.RoundUnit)分単位で$modeJp）")
    $script:SumList = Add-To $f (New-Object System.Windows.Forms.ListView) 10 42 800 300
    $script:SumList.View = 'Details'; $script:SumList.FullRowSelect = $true
    foreach ($c in @(@('日付', 90, 'Left'), @('プロジェクトコード', 110, 'Left'), @('WBS', 270, 'Left'), @('タスク名', 200, 'Left'), @('実績', 55, 'Right'), @('工数(h)', 60, 'Right'))) {
        [void]$script:SumList.Columns.Add($c[0], $c[1], $c[2])
    }
    $script:SumMsg = Add-To $f (New-Object System.Windows.Forms.Label) 10 348 800 34
    $b1 = Add-To $f (New-Object System.Windows.Forms.Button) 10 384 90 28 '集計'
    $b2 = Add-To $f (New-Object System.Windows.Forms.Button) 105 384 150 28 'CSV に保存して開く'
    $b3 = Add-To $f (New-Object System.Windows.Forms.Button) 720 384 90 28 '閉じる'
    $b1.Add_Click({ Invoke-Safe { Update-SummaryView } })
    $b2.Add_Click({ Invoke-Safe { Save-SummaryCsv } })
    $b3.Add_Click({ $this.FindForm().Close() })
    $script:SumRows = $null
    Update-SummaryView
    [void]$f.ShowDialog($form)
    $f.Dispose()
} })

# --------------------------------------------------------------------------
# 手入力 (開始の押し忘れなど)
# --------------------------------------------------------------------------
function Get-Overlaps([datetime]$Start, [datetime]$End, [int]$ExceptLine = -1) {
    # 指定の時間帯と重なる記録 (作業中のタスクも含む。$ExceptLine の行は除く)
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($r in (Get-RecordList $script:Rec)) {
        if ($r.Line -eq $ExceptLine) { continue }
        if ($r.Start -lt $End -and $r.End -gt $Start) { $out.Add("$($r.Start.ToString('HH:mm'))～$($r.End.ToString('HH:mm')) $($r.Name)") }
    }
    $cur = $script:Rec.Current
    if ($null -ne $cur -and $cur.start -lt $End -and (Get-Date) -gt $Start) { $out.Add("$($cur.start.ToString('HH:mm'))～ (作業中) $($cur.task.name)") }
    return , $out
}

function Update-ManualList {
    # 選んだ日の作業記録を一覧に出す
    $day = $script:ManDate.Value.ToString('yyyy-MM-dd')
    $all = Get-RecordList $script:Rec
    $script:ManRows = @($all | Where-Object { $_.Date -eq $day } | Sort-Object Start)
    $lv = $script:ManList
    $lv.BeginUpdate(); $lv.Items.Clear()
    $total = 0.0
    foreach ($r in $script:ManRows) {
        $m = ($r.End - $r.Start).TotalMinutes; $total += $m
        $it = New-Object System.Windows.Forms.ListViewItem($r.Start.ToString('HH:mm'))
        foreach ($v in @($r.End.ToString('HH:mm'), (Format-HM $m), $r.Name, $r.Reason)) { [void]$it.SubItems.Add([string]$v) }
        [void]$lv.Items.Add($it)
    }
    $lv.EndUpdate()
    $script:ManTotal.Text = "この日の合計: $(Format-HM $total)"
}

function Get-ManualInput {
    # 入力欄の内容を確かめて @{ Task; Start; End } を返す。問題があれば $null (メッセージを表示)
    $i = $script:ManTask.SelectedIndex
    if ($i -lt 0 -or $i -ge $script:ManTasks.Count) { $script:ManMsg.Text = 'タスクを選んでください'; return $null }
    $day = $script:ManDate.Value.Date
    $start = $day.Add($script:ManFrom.Value.TimeOfDay); $start = $start.AddSeconds(-$start.Second)
    $end = $day.Add($script:ManTo.Value.TimeOfDay); $end = $end.AddSeconds(-$end.Second)
    if ($end -le $start) { $script:ManMsg.Text = '終了は開始より後の時刻にしてください (日をまたぐ作業は日ごとに分けて入力)'; return $null }
    if ($end -gt (Get-Date).AddMinutes(1)) { $script:ManMsg.Text = 'これから先の時刻は入力できません'; return $null }
    return @{ Task = $script:ManTasks[$i]; Start = $start; End = $end }
}

function Confirm-Overlap($In, [int]$ExceptLine) {
    $ov = Get-Overlaps $In.Start $In.End $ExceptLine
    if ($ov.Count -eq 0) { return $true }
    $q = "次の記録と時間が重なっています。重なった時間は二重に計上されます。`n`n$($ov -join "`n")`n`nこのまま保存しますか？"
    return ([System.Windows.Forms.MessageBox]::Show($script:ManForm, $q, $AppName, 'YesNo', 'Warning') -eq 'Yes')
}

function Complete-ManualChange([string]$Msg) {
    $script:ManMsg.Text = $Msg + "`n※ OBPM へは、次に「本日終了」を押したときに入力されます"
    Set-Status $Msg
    Update-ManualList
    Update-TaskList (([string]$combo.SelectedItem))
    Update-Today
}

function Add-ManualRecord {
    $in = Get-ManualInput; if ($null -eq $in) { return }
    if (-not (Confirm-Overlap $in -1)) { return }
    Add-RecordRows $script:Rec $in.Task $in.Start $in.End '手入力'
    Complete-ManualChange "追加しました: $($in.Start.ToString('MM/dd HH:mm'))～$($in.End.ToString('HH:mm')) $($in.Task.name)"
}

function Get-SelectedManualRow {
    if ($script:ManList.SelectedIndices.Count -eq 0) { $script:ManMsg.Text = '一覧から記録を選んでください'; return $null }
    return $script:ManRows[$script:ManList.SelectedIndices[0]]
}

function Edit-ManualRecord {
    $r = Get-SelectedManualRow; if ($null -eq $r) { return }
    $in = Get-ManualInput; if ($null -eq $in) { return }
    if (-not (Confirm-Overlap $in $r.Line)) { return }
    Update-RecordFile $script:Rec $r.Line $in.Task $in.Start $in.End '手入力(修正)'
    Complete-ManualChange "修正しました: $($in.Start.ToString('MM/dd HH:mm'))～$($in.End.ToString('HH:mm')) $($in.Task.name)"
}

function Remove-ManualRecord {
    $r = Get-SelectedManualRow; if ($null -eq $r) { return }
    $q = "次の記録を削除しますか？`n$($r.Start.ToString('MM/dd HH:mm'))～$($r.End.ToString('HH:mm')) $($r.Name)"
    if ([System.Windows.Forms.MessageBox]::Show($script:ManForm, $q, $AppName, 'YesNo', 'Question') -ne 'Yes') { return }
    Update-RecordFile $script:Rec $r.Line $null $r.Start $r.End ''
    Complete-ManualChange "削除しました: $($r.Start.ToString('MM/dd HH:mm'))～$($r.End.ToString('HH:mm')) $($r.Name)"
}

function Select-ManualRow {
    # 一覧で選んだ記録を入力欄に写す
    if ($script:ManList.SelectedIndices.Count -eq 0) { return }
    $r = $script:ManRows[$script:ManList.SelectedIndices[0]]
    $idx = -1
    for ($i = 0; $i -lt $script:ManTasks.Count; $i++) {
        $t = $script:ManTasks[$i]
        if ((Get-TaskKey $t.name $t.project $t.wbs) -eq (Get-TaskKey $r.Name $r.Project $r.Wbs)) { $idx = $i; break }
    }
    $script:ManTask.SelectedIndex = $idx
    if ($idx -lt 0) { $script:ManMsg.Text = "この記録のタスク「$($r.Name)」は今のタスク一覧にありません。修正する場合はタスクを選び直してください" } else { $script:ManMsg.Text = '' }
    $script:ManFrom.Value = $r.Start; $script:ManTo.Value = $r.End
}

$btnManual.Add_Click({ Invoke-Safe {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = '作業記録の追加・修正'; $f.Font = $font; $f.AutoScaleMode = 'Dpi'; $f.TopMost = $Cfg.AlwaysOnTop
    $f.ClientSize = New-Sz 560 430; $f.StartPosition = 'CenterParent'; $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $script:ManForm = $f
    $script:ManTasks = $script:Tasks.ToArray()   # 追加のたびにタイマー画面の一覧は並び替わるので、この画面を開いたときの一覧を使う
    # 日付と、その日の記録の一覧
    [void](Add-To $f (New-Object System.Windows.Forms.Label) 10 14 45 20 '日付')
    $script:ManDate = Add-To $f (New-Object System.Windows.Forms.DateTimePicker) 60 10 130 24
    $script:ManDate.Format = 'Custom'; $script:ManDate.CustomFormat = 'yyyy-MM-dd (ddd)'
    $today = (Get-Date).Date
    $script:ManDate.MinDate = (Get-Date -Year $today.Year -Month $today.Month -Day 1).Date   # 修正できるのは今月分だけ (前月分は自動入力されないため)
    $script:ManDate.MaxDate = $today.AddDays(1).AddSeconds(-1); $script:ManDate.Value = $today
    $script:ManTotal = Add-To $f (New-Object System.Windows.Forms.Label) 200 14 200 20 ''
    $script:ManList = Add-To $f (New-Object System.Windows.Forms.ListView) 10 40 540 180
    $lv = $script:ManList; $lv.View = 'Details'; $lv.FullRowSelect = $true; $lv.HideSelection = $false; $lv.MultiSelect = $false; $lv.HeaderStyle = 'Nonclickable'
    foreach ($c in @(@('開始', 55), @('終了', 55), @('時間', 55), @('タスク', 270), @('記録の種類', 85))) { [void]$lv.Columns.Add($c[0], $c[1]) }
    # 入力欄
    $grp = Add-To $f (New-Object System.Windows.Forms.GroupBox) 10 228 540 100 '選んだ記録の内容 / 追加する記録の内容'
    [void](Add-To $grp (New-Object System.Windows.Forms.Label) 10 26 45 20 'タスク')
    $script:ManTask = Add-To $grp (New-Object System.Windows.Forms.ComboBox) 60 22 470 24
    $script:ManTask.DropDownStyle = 'DropDownList'
    foreach ($t in $script:ManTasks) { [void]$script:ManTask.Items.Add($t.name) }
    if ($combo.SelectedIndex -ge 0) { $script:ManTask.SelectedIndex = $combo.SelectedIndex }
    [void](Add-To $grp (New-Object System.Windows.Forms.Label) 10 62 45 20 '時間')
    $script:ManFrom = Add-To $grp (New-Object System.Windows.Forms.DateTimePicker) 60 58 80 24
    [void](Add-To $grp (New-Object System.Windows.Forms.Label) 145 62 20 20 '～')
    $script:ManTo = Add-To $grp (New-Object System.Windows.Forms.DateTimePicker) 168 58 80 24
    $now = Get-Date; $now = $now.Date.AddMinutes([Math]::Floor($now.TimeOfDay.TotalMinutes / 5) * 5)   # 既定: 30 分前～今 (5 分単位)
    foreach ($p in @(@($script:ManFrom, $now.AddMinutes(-30)), @($script:ManTo, $now))) {
        $p[0].Format = 'Custom'; $p[0].CustomFormat = 'HH:mm'; $p[0].ShowUpDown = $true; $p[0].Value = $p[1]
    }
    [void](Add-To $grp (New-Object System.Windows.Forms.Label) 255 62 150 20 '（24 時間表記）')
    # ボタン・メッセージ
    $bAdd = Add-To $f (New-Object System.Windows.Forms.Button) 10 336 110 30 '追加'
    $bEdit = Add-To $f (New-Object System.Windows.Forms.Button) 126 336 110 30 '選んだ記録を修正'
    $bDel = Add-To $f (New-Object System.Windows.Forms.Button) 242 336 110 30 '選んだ記録を削除'
    $bClose = Add-To $f (New-Object System.Windows.Forms.Button) 460 336 90 30 '閉じる'
    $script:ManMsg = Add-To $f (New-Object System.Windows.Forms.Label) 10 374 540 50
    $script:ManMsg.ForeColor = [System.Drawing.Color]::FromArgb(170, 50, 50)
    $script:ManDate.Add_ValueChanged({ Invoke-Safe { Update-ManualList; $script:ManMsg.Text = '' } })
    $lv.Add_SelectedIndexChanged({ Invoke-Safe { Select-ManualRow } })
    $bAdd.Add_Click({ Invoke-Safe { Add-ManualRecord } })
    $bEdit.Add_Click({ Invoke-Safe { Edit-ManualRecord } })
    $bDel.Add_Click({ Invoke-Safe { Remove-ManualRecord } })
    $bClose.Add_Click({ $this.FindForm().Close() })
    Update-ManualList
    [void]$f.ShowDialog($form)
    $f.Dispose()
} })

# --------------------------------------------------------------------------
# 起動
# --------------------------------------------------------------------------
try {
    Update-TaskList
    $msgs = @()
    $m = Restore-Running $script:Rec; if ($m) { $msgs += $m }
    if ($Cfg.LaunchObpm) {
        if (-not $Cfg.ObpmExe) { $msgs += '[obpm] obpm_exe が未設定のため、SI Object Browser PM は起動しませんでした' }
        elseif (-not (Test-Path -LiteralPath $Cfg.ObpmExe)) { $msgs += "SI Object Browser PM が見つかりません: $($Cfg.ObpmExe)" }
        else {
            $pname = [System.IO.Path]::GetFileNameWithoutExtension($Cfg.ObpmExe)
            if (-not (Get-Process -Name $pname -ErrorAction SilentlyContinue)) {
                Start-Process -FilePath $Cfg.ObpmExe -WorkingDirectory (Split-Path -Parent $Cfg.ObpmExe)
            }
        }
    }
    Set-Status ($msgs -join "`n")
    Update-Today
    $form.Add_Shown({ Invoke-Safe { Show-StartupNotices; Update-Today } })
    $timer.Start()
    if ($null -ne $mini) {
        $fade.Start()
        $form.Show()
        [System.Windows.Forms.Application]::Run($mini)
        $fade.Stop()
    } else {
        [void]$form.ShowDialog()
    }
    $timer.Stop()
} catch {
    Write-ErrorLog $Cfg.DataDir ($_.ToString() + "`r`n" + $_.ScriptStackTrace)
    Show-Error "エラーが発生しました:`n$($_.Exception.Message)"
    exit 1
}
