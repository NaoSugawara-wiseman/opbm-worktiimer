# -*- coding: utf-8 -*-
# common.ps1  作業タイマー / OBPM 自動入力で共通に使う部品 (Windows PowerShell 5.1 対応)
# ※ このファイルは「UTF-8 (BOM 付き)」で保存すること。BOM がないと PowerShell 5.1 で日本語が文字化けする

Set-StrictMode -Version 2.0

$script:Inv = [System.Globalization.CultureInfo]::InvariantCulture
$script:TimeFmt = 'yyyy-MM-dd HH:mm:ss'
$script:Utf8Bom = New-Object System.Text.UTF8Encoding($true)
$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# --------------------------------------------------------------------------
# 文字コード・ファイル
# --------------------------------------------------------------------------
function Get-SjisEncoding {
    try { [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance) } catch { }
    return [System.Text.Encoding]::GetEncoding(932)
}

function Read-TextAuto([string]$Path) {
    # UTF-8 (BOM あり/なし) と Shift_JIS (Excel の通常の CSV 保存) のどちらも読む
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return $script:Utf8NoBom.GetString($bytes, 3, $bytes.Length - 3)
    }
    try {
        $strict = New-Object System.Text.UTF8Encoding($false, $true)
        return $strict.GetString($bytes)
    } catch {
        return (Get-SjisEncoding).GetString($bytes)
    }
}

function Write-TextFile([string]$Path, [string]$Text, [switch]$Append) {
    # 新規作成は BOM 付き UTF-8 (Excel で文字化けしない)、追記は BOM なし
    if ($Append -and (Test-Path -LiteralPath $Path)) {
        [System.IO.File]::AppendAllText($Path, $Text, $script:Utf8NoBom)
    } else {
        [System.IO.File]::WriteAllText($Path, $Text, $script:Utf8Bom)
    }
}

function Write-JsonFile([string]$Path, $Object) {
    $tmp = "$Path.tmp"
    [System.IO.File]::WriteAllText($tmp, ($Object | ConvertTo-Json -Depth 5), $script:Utf8NoBom)
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}

function Read-JsonFile([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try { return (Read-TextAuto $Path) | ConvertFrom-Json } catch { return $null }
}

function ConvertTo-Hashtable($Obj) {
    $h = @{}
    if ($null -ne $Obj) { foreach ($p in $Obj.PSObject.Properties) { $h[$p.Name] = $p.Value } }
    return $h
}

# --------------------------------------------------------------------------
# CSV (ダブルクォート対応)
# --------------------------------------------------------------------------
function ConvertFrom-CsvText([string]$Text) {
    $rows = New-Object System.Collections.Generic.List[object]
    $field = New-Object System.Text.StringBuilder
    $row = New-Object System.Collections.Generic.List[string]
    $inQ = $false
    $i = 0
    $n = $Text.Length
    while ($i -lt $n) {
        $c = $Text[$i]
        if ($inQ) {
            if ($c -eq '"') {
                if ($i + 1 -lt $n -and $Text[$i + 1] -eq '"') { [void]$field.Append('"'); $i++ }
                else { $inQ = $false }
            } else { [void]$field.Append($c) }
        } else {
            if ($c -eq '"') { $inQ = $true }
            elseif ($c -eq ',') { $row.Add($field.ToString()); [void]$field.Clear() }
            elseif ($c -eq "`r") { }
            elseif ($c -eq "`n") {
                $row.Add($field.ToString()); [void]$field.Clear()
                $rows.Add($row.ToArray()); $row = New-Object System.Collections.Generic.List[string]
            } else { [void]$field.Append($c) }
        }
        $i++
    }
    if ($field.Length -gt 0 -or $row.Count -gt 0) { $row.Add($field.ToString()); $rows.Add($row.ToArray()) }
    return , $rows
}

function Read-CsvRows([string]$Path) {
    return , (ConvertFrom-CsvText (Read-TextAuto $Path))
}

function ConvertTo-CsvLine([object[]]$Fields) {
    $out = foreach ($f in $Fields) {
        $s = [string]$f
        if ($s -match '[",\r\n]') { '"' + $s.Replace('"', '""') + '"' } else { $s }
    }
    return ($out -join ',')
}

# --------------------------------------------------------------------------
# 設定ファイル (INI)
# --------------------------------------------------------------------------
function Read-Ini([string]$Path) {
    $ini = @{}
    $sec = ''
    $ini[$sec] = @{}
    foreach ($raw in ((Read-TextAuto $Path) -split "`r?`n")) {
        $line = $raw.Trim()
        if ($line -eq '' -or $line.StartsWith(';') -or $line.StartsWith('#')) { continue }
        if ($line -match '^\[(.+)\]$') { $sec = $Matches[1].Trim(); if (-not $ini.ContainsKey($sec)) { $ini[$sec] = @{} }; continue }
        $eq = $line.IndexOf('=')
        if ($eq -lt 1) { continue }
        $key = $line.Substring(0, $eq).Trim()
        $val = $line.Substring($eq + 1)
        $m = [regex]::Match($val, '\s;')       # 「値 ; コメント」の行末コメント
        if ($m.Success) { $val = $val.Substring(0, $m.Index) }
        $val = $val.Trim()
        if ($val.Length -ge 2 -and $val[0] -eq '"' -and $val[$val.Length - 1] -eq '"') { $val = $val.Substring(1, $val.Length - 2) }
        $ini[$sec][$key] = $val
    }
    return $ini
}

function Get-IniValue($Ini, [string]$Section, [string]$Key, [string]$Default = '') {
    if ($Ini.ContainsKey($Section) -and $Ini[$Section].ContainsKey($Key) -and $Ini[$Section][$Key] -ne '') {
        return [string]$Ini[$Section][$Key]
    }
    return $Default
}

function Resolve-RelPath([string]$Base, [string]$P) {
    if ([System.IO.Path]::IsPathRooted($P)) { return $P }
    return (Join-Path $Base $P)
}

# --------------------------------------------------------------------------
# 文字列・数値
# --------------------------------------------------------------------------
function Get-Norm([string]$S) {
    # 全角/半角や前後の空白の違いを無視して比較するための正規化
    if ($null -eq $S) { return '' }
    return $S.Normalize([System.Text.NormalizationForm]::FormKC).Trim()
}

function Format-Num([double]$V) { return $V.ToString('0.##', $script:Inv) }

function ConvertTo-Double([string]$S) {
    $d = 0.0
    if ([double]::TryParse(($S -replace ',', ''), [System.Globalization.NumberStyles]::Float, $script:Inv, [ref]$d)) { return $d }
    return $null
}

function Get-RoundedHours([double]$Minutes, [int]$Unit, [string]$Mode) {
    $u = [decimal]::Parse($Minutes.ToString('R', $script:Inv), $script:Inv) / [decimal]$Unit
    switch ($Mode) {
        'up'    { $units = [Math]::Ceiling($u) }
        'down'  { $units = [Math]::Floor($u) }
        default { $units = [Math]::Round($u, 0, [MidpointRounding]::AwayFromZero) }
    }
    return [double]($units * [decimal]$Unit / [decimal]60)
}

function Format-HM([double]$Minutes) {
    $m = [int][Math]::Round($Minutes)
    return ('{0}:{1:00}' -f [Math]::Floor($m / 60), ($m % 60))
}

# --------------------------------------------------------------------------
# 作業記録 (開始・終了・日跨ぎ・異常終了からの復旧)
# --------------------------------------------------------------------------
$script:RecordHeader = @('日付', '開始', '終了', '分', 'タスク名', 'プロジェクトコード', 'WBS', '終了理由')

function New-Recorder([string]$DataDir) {
    if (-not (Test-Path -LiteralPath $DataDir)) { New-Item -ItemType Directory -Path $DataDir -Force | Out-Null }
    return @{ Records = (Join-Path $DataDir 'records.csv'); Running = (Join-Path $DataDir 'running.json'); Current = $null }
}

function Split-ByDay([datetime]$Start, [datetime]$End) {
    $parts = New-Object System.Collections.Generic.List[object]
    $cur = $Start
    while ($cur.Date -lt $End.Date) {
        $mid = $cur.Date.AddDays(1)
        $parts.Add(@($cur, $mid)); $cur = $mid
    }
    if ($End -gt $cur) { $parts.Add(@($cur, $End)) }
    return , $parts
}

function Add-RecordRows($Rec, $Task, [datetime]$Start, [datetime]$End, [string]$Reason) {
    $lines = New-Object System.Collections.Generic.List[string]
    $new = -not (Test-Path -LiteralPath $Rec.Records)
    if ($new) { $lines.Add((ConvertTo-CsvLine $script:RecordHeader)) }
    foreach ($p in (Split-ByDay $Start $End)) {
        $min = [Math]::Round(($p[1] - $p[0]).TotalSeconds / 60, 1)
        if ($min -le 0) { continue }
        $lines.Add((ConvertTo-CsvLine @($p[0].ToString('yyyy-MM-dd', $script:Inv), $p[0].ToString($script:TimeFmt, $script:Inv),
                    $p[1].ToString($script:TimeFmt, $script:Inv), $min.ToString('0.0', $script:Inv),
                    $Task.name, $Task.project, $Task.wbs, $Reason)))
    }
    if ($lines.Count -gt 0) { Write-TextFile $Rec.Records (($lines -join "`r`n") + "`r`n") -Append:(-not $new) }
}

function Get-RecordList($Rec) {
    # 作業記録の各行を返す (Line = ファイル内の行番号。修正・削除に使う)
    $out = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $Rec.Records)) { return , $out }
    $rows = Read-CsvRows $Rec.Records
    for ($i = 1; $i -lt $rows.Count; $i++) {
        $r = @($rows[$i]); if ($r.Count -lt 8) { continue }
        $s = [datetime]::MinValue; $e = [datetime]::MinValue
        if (-not [datetime]::TryParse($r[1], [ref]$s) -or -not [datetime]::TryParse($r[2], [ref]$e)) { continue }
        $out.Add([pscustomobject]@{ Line = $i; Date = $s.ToString('yyyy-MM-dd'); Start = $s; End = $e; Name = $r[4]; Project = $r[5]; Wbs = $r[6]; Reason = $r[7] })
    }
    return , $out
}

function Update-RecordFile($Rec, [int]$Line, $Task, [datetime]$Start, [datetime]$End, [string]$Reason) {
    # 作業記録の 1 行を置き換える ($Task が $null なら削除)。書き換え前の内容は records.csv.bak に残す
    $rows = Read-CsvRows $Rec.Records
    if ($Line -lt 1 -or $Line -ge $rows.Count) { throw "作業記録の行が見つかりません (行 $Line)。もう一度開き直してください" }
    Copy-Item -LiteralPath $Rec.Records -Destination ($Rec.Records + '.bak') -Force
    $lines = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $rows.Count; $i++) {
        if ($i -ne $Line) { $lines.Add((ConvertTo-CsvLine @($rows[$i]))); continue }
        if ($null -eq $Task) { continue }
        $min = [Math]::Round(($End - $Start).TotalSeconds / 60, 1)
        $lines.Add((ConvertTo-CsvLine @($Start.ToString('yyyy-MM-dd', $script:Inv), $Start.ToString($script:TimeFmt, $script:Inv),
                    $End.ToString($script:TimeFmt, $script:Inv), $min.ToString('0.0', $script:Inv), $Task.name, $Task.project, $Task.wbs, $Reason)))
    }
    Write-TextFile $Rec.Records (($lines -join "`r`n") + "`r`n")
}

function Save-Running($Rec, [datetime]$Now) {
    if ($null -eq $Rec.Current) {
        if (Test-Path -LiteralPath $Rec.Running) { Remove-Item -LiteralPath $Rec.Running -Force }
        return
    }
    Write-JsonFile $Rec.Running ([ordered]@{ task = $Rec.Current.task; start = $Rec.Current.start.ToString($script:TimeFmt, $script:Inv); heartbeat = $Now.ToString($script:TimeFmt, $script:Inv) })
}

function Test-SameTask($A, $B) {
    return ($A.name -ceq $B.name -and $A.project -ceq $B.project -and $A.wbs -ceq $B.wbs)
}

function Start-Work($Rec, $Task, [datetime]$Now) {
    # 戻り値: 表示するメッセージ (なければ $null)
    $msg = $null
    if ($null -ne $Rec.Current) {
        if (Test-SameTask $Rec.Current.task $Task) { return $null }
        $prev = $Rec.Current.task.name
        Stop-Work $Rec $Now '別タスクに切替'
        $msg = "「$prev」を終了して切り替えました"
    }
    $Rec.Current = @{ task = $Task; start = $Now }
    Save-Running $Rec $Now
    return $msg
}

function Stop-Work($Rec, [datetime]$End, [string]$Reason = '作業終了') {
    if ($null -eq $Rec.Current) { return }
    $s = $Rec.Current.start
    if ($End -lt $s) { $End = $s }
    Add-RecordRows $Rec $Rec.Current.task $s $End $Reason
    $Rec.Current = $null
    Save-Running $Rec $End
}

function Update-Heartbeat($Rec, [datetime]$Now) {
    if ($null -ne $Rec.Current) { Save-Running $Rec $Now }
}

function Restore-Running($Rec) {
    # 前回終了し損ねた記録を、最後の確認時刻で締める。戻り値: メッセージ
    if (-not (Test-Path -LiteralPath $Rec.Running)) { return $null }
    $msg = $null
    try {
        $st = Read-JsonFile $Rec.Running
        $task = @{ name = [string]$st.task.name; project = [string]$st.task.project; wbs = [string]$st.task.wbs }
        $start = [datetime]::ParseExact($st.start, $script:TimeFmt, $script:Inv)
        $end = [datetime]::ParseExact($st.heartbeat, $script:TimeFmt, $script:Inv)
        if ($end -lt $start) { $end = $start }
        Add-RecordRows $Rec $task $start $end '前回終了し損ねたため最終確認時刻で締め'
        $msg = "前回の「$($task.name)」を $($end.ToString('MM/dd HH:mm')) で締めました"
    } catch {
        $msg = "前回の記録を復旧できませんでした: $($_.Exception.Message)"
    }
    Remove-Item -LiteralPath $Rec.Running -Force -ErrorAction SilentlyContinue
    return $msg
}

function Read-Records([string]$Path) {
    $out = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $Path)) { return , $out }
    $rows = Read-CsvRows $Path
    for ($i = 1; $i -lt $rows.Count; $i++) {
        $r = $rows[$i]
        if ($r.Count -lt 7) { continue }
        $m = ConvertTo-Double $r[3]
        if ($null -eq $m) { continue }
        $out.Add([pscustomobject]@{ date = $r[0]; minutes = $m; name = $r[4]; project = $r[5].Trim(); wbs = $r[6].Trim() })
    }
    return , $out
}

# --------------------------------------------------------------------------
# 祝日・月末 (月末の営業日以降は自動入力しない)
# --------------------------------------------------------------------------
function Get-NthMonday([int]$Year, [int]$Month, [int]$N) {
    $d = Get-Date -Year $Year -Month $Month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    while ($d.DayOfWeek -ne 'Monday') { $d = $d.AddDays(1) }
    return $d.AddDays(7 * ($N - 1))
}

function Get-JpHolidays([int]$Year) {
    # 日本の祝日 (2020 年以降の制度。振替休日・国民の休日を含む) を 'yyyy-MM-dd' の集合で返す
    $base = New-Object System.Collections.Generic.List[datetime]
    foreach ($md in '1-1', '2-11', '2-23', '4-29', '5-3', '5-4', '5-5', '8-11', '11-3', '11-23') {
        $m, $dd = $md -split '-'; $base.Add((Get-Date -Year $Year -Month ([int]$m) -Day ([int]$dd)).Date)
    }
    $base.Add((Get-NthMonday $Year 1 2)); $base.Add((Get-NthMonday $Year 7 3)); $base.Add((Get-NthMonday $Year 9 3)); $base.Add((Get-NthMonday $Year 10 2))
    $y = $Year - 1980
    $base.Add((Get-Date -Year $Year -Month 3 -Day ([int][Math]::Floor(20.8431 + 0.242194 * $y - [Math]::Floor($y / 4)))).Date)   # 春分の日
    $base.Add((Get-Date -Year $Year -Month 9 -Day ([int][Math]::Floor(23.2488 + 0.242194 * $y - [Math]::Floor($y / 4)))).Date)   # 秋分の日
    $set = New-Object System.Collections.Generic.HashSet[string]
    foreach ($d in $base) { [void]$set.Add($d.ToString('yyyy-MM-dd')) }
    # 振替休日: 祝日が日曜なら、その後の最初の祝日でない日
    foreach ($d in $base) {
        if ($d.DayOfWeek -eq 'Sunday') {
            $x = $d.AddDays(1); while ($set.Contains($x.ToString('yyyy-MM-dd'))) { $x = $x.AddDays(1) }
            [void]$set.Add($x.ToString('yyyy-MM-dd'))
        }
    }
    # 国民の休日: 祝日に挟まれた平日
    foreach ($d in $base) {
        $mid = $d.AddDays(1)
        if ($set.Contains($d.AddDays(2).ToString('yyyy-MM-dd')) -and -not $set.Contains($mid.ToString('yyyy-MM-dd')) -and $mid.DayOfWeek -ne 'Sunday') { [void]$set.Add($mid.ToString('yyyy-MM-dd')) }
    }
    return , $set
}

function Get-ExtraHolidays([string]$DataDir) {
    # 会社独自の休日: 共有フォルダの holidays.csv (1 列目に日付。2 列目以降はメモ。# で始まる行は無視)
    $set = New-Object System.Collections.Generic.HashSet[string]
    $p = Join-Path $DataDir 'holidays.csv'
    if (-not (Test-Path -LiteralPath $p)) { return , $set }
    foreach ($line in ((Read-TextAuto $p) -split "`r?`n")) {
        $c = ($line -split ',')[0].Trim().Trim('"')
        if (-not $c -or $c.StartsWith('#')) { continue }
        $d = [datetime]::MinValue
        if ([datetime]::TryParse($c, [ref]$d)) { [void]$set.Add($d.ToString('yyyy-MM-dd')) }
    }
    return , $set
}

function Get-LastBusinessDay([datetime]$Day, $Extra) {
    # その月の最後の営業日 (土日・祝日・holidays.csv の日を除く)
    $d = (Get-Date -Year $Day.Year -Month $Day.Month -Day 1).Date.AddMonths(1).AddDays(-1)
    $hol = Get-JpHolidays $d.Year
    while ($d.DayOfWeek -in 'Saturday', 'Sunday' -or $hol.Contains($d.ToString('yyyy-MM-dd')) -or ($null -ne $Extra -and $Extra.Contains($d.ToString('yyyy-MM-dd')))) { $d = $d.AddDays(-1) }
    return $d
}

function Test-MonthEnd([datetime]$Day, $Extra) {
    # 月末の営業日、またはそれ以降の同じ月の日なら $true (自動入力しない日)
    return ($Day.Date -ge (Get-LastBusinessDay $Day $Extra))
}

function Get-DayHours($Mins, [int]$Unit, [string]$Mode, $Daily, [bool]$Closed) {
    # 1 日分のタスクの工数を決める。$Mins: キー → 分 (その日のタスクごとの合計)
    #   ・各タスクを丸める。丸めて 0 になる短いタスクも 1 単位 (0.25H) として計上する
    #   ・終わった日 ($Closed) は、合計を社内ルール (定時 + 残業は step 単位で切り捨て) に合わせ、
    #     ずれは一番長く作業したタスクで調整する
    # 戻り値: Hours (キー → 時間) / Notes (ログに出す説明)
    $q = [ordered]@{}; $total = 0.0
    foreach ($k in $Mins.Keys) {
        $m = [double]$Mins[$k]; $total += $m
        $rm = [int][Math]::Round((Get-RoundedHours $m $Unit $Mode) * 60)
        if ($rm -eq 0 -and $m -gt 0) { $rm = $Unit }
        $q[$k] = $rm
    }
    $notes = New-Object System.Collections.Generic.List[string]
    if ($null -ne $Daily -and $Daily.Adjust -and $Closed -and $q.Count -gt 0) {
        $base = [int][Math]::Round($Daily.BaseHours * 60); $step = [int][Math]::Round($Daily.StepHours * 60)
        $sum = 0; foreach ($v in $q.Values) { $sum += $v }
        $tt = Format-HM $total
        if ($total -lt $base - 1e-6) {
            $notes.Add("記録の合計 $tt が $($Daily.BaseHours)H 未満のため、合計の調整はしていません (入力合計 $($sum / 60)H)。記録漏れがないか確認してください")
        } else {
            $target = $base + [int][Math]::Floor(($total - $base) / $step + 1e-6) * $step
            $diff = $target - $sum
            # 長く作業した順 (同じ長さなら記録の順)
            $i = 0
            $order = @(@(foreach ($k in $Mins.Keys) { [pscustomobject]@{ Key = $k; Min = [double]$Mins[$k]; Index = $i }; $i++ }) |
                Sort-Object @{ Expression = 'Min'; Descending = $true }, @{ Expression = 'Index'; Ascending = $true })
            $changes = New-Object System.Collections.Generic.List[string]
            if ($diff -gt 0) {
                $k = $order[0].Key; $q[$k] += $diff; $changes.Add("$($k.Split('|')[-1]) +$($diff / 60)H")
            } elseif ($diff -lt 0) {
                $need = -$diff
                $backup = [ordered]@{}; foreach ($k in $q.Keys) { $backup[$k] = $q[$k] }
                foreach ($o in $order) {
                    $take = [Math]::Min($need, $q[$o.Key] - $Unit)
                    if ($take -gt 0) { $q[$o.Key] -= $take; $need -= $take; $changes.Add("$($o.Key.Split('|')[-1]) -$($take / 60)H") }
                    if ($need -eq 0) { break }
                }
                if ($need -gt 0) { $q = $backup; $changes.Clear(); $notes.Add("短いタスクが多く、合計を $($target / 60)H に合わせられませんでした (入力合計 $($sum / 60)H)。records.csv を確認してください") }
            }
            if ($changes.Count -gt 0) { $notes.Add("記録の合計 $tt → 入力合計 $($target / 60)H に調整 ($($changes -join ', '))") }
        }
    }
    $hours = @{}; foreach ($k in $q.Keys) { $hours[$k] = $q[$k] / 60.0 }
    return @{ Hours = $hours; Notes = $notes }
}

function Get-Summary($Records, [string]$From, [string]$To, [int]$Unit, [string]$Mode) {
    # 日付×タスクごとに合計し、丸めた工数 (時間) を付ける
    $acc = [ordered]@{}
    foreach ($r in $Records) {
        if ([string]::CompareOrdinal($r.date, $From) -lt 0 -or [string]::CompareOrdinal($r.date, $To) -gt 0) { continue }
        $k = "$($r.date)`t$($r.project)`t$($r.wbs)`t$($r.name)"
        if (-not $acc.Contains($k)) { $acc[$k] = 0.0 }
        $acc[$k] += $r.minutes
    }
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($k in ($acc.Keys | Sort-Object)) {
        $p = $k -split "`t"
        $out.Add([pscustomobject]@{ date = $p[0]; project = $p[1]; wbs = $p[2]; name = $p[3]
                minutes = [Math]::Round($acc[$k], 1); hours = (Get-RoundedHours $acc[$k] $Unit $Mode) })
    }
    return , $out
}

# --------------------------------------------------------------------------
# Windows の機能 (アイドル時間・画面ロック・前面ウィンドウ・クリック)
# --------------------------------------------------------------------------
if ($env:OS -eq 'Windows_NT' -and -not ('WtNative' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class WtNative {
    [StructLayout(LayoutKind.Sequential)] struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }
    [DllImport("user32.dll")] static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
    [DllImport("kernel32.dll")] static extern uint GetTickCount();
    [DllImport("user32.dll")] static extern IntPtr OpenInputDesktop(uint f, bool inherit, uint access);
    [DllImport("user32.dll")] static extern bool SwitchDesktop(IntPtr h);
    [DllImport("user32.dll")] static extern bool CloseDesktop(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint flags);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);
    [DllImport("user32.dll")] static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, IntPtr pid);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] static extern bool AttachThreadInput(uint a, uint b, bool attach);
    [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr h);

    public static bool PostButtonClick(IntPtr h) { return PostMessage(h, 0x00F5, IntPtr.Zero, IntPtr.Zero); }  // BM_CLICK (マウスを使わずにボタンを押す)

    public static void DpiAware() { try { SetProcessDPIAware(); } catch { } }  // 画面の拡大率に関係なく正しい位置をクリックする

    public static double IdleSeconds() {
        LASTINPUTINFO li = new LASTINPUTINFO(); li.cbSize = (uint)Marshal.SizeOf(li);
        if (!GetLastInputInfo(ref li)) return 0;
        return unchecked(GetTickCount() - li.dwTime) / 1000.0;
    }
    public static bool IsLocked() {
        IntPtr h = OpenInputDesktop(0, false, 0x0100);
        if (h == IntPtr.Zero) return true;
        bool ok = SwitchDesktop(h); CloseDesktop(h); return !ok;
    }
    public static bool IsForeground(IntPtr hwnd) {
        IntPtr fg = GetForegroundWindow();
        return fg == hwnd || GetAncestor(fg, 2) == hwnd;
    }
    public static void Activate(IntPtr hwnd) {
        if (IsIconic(hwnd)) ShowWindow(hwnd, 9);
        if (IsForeground(hwnd)) return;   // すでに手前なら何もしない (Alt を押すとメニュー操作の状態になり、次のクリックが無視されるため)
        // ① 手前のウィンドウの入力に一時的につないで前面化する (キーを押さない方法)
        IntPtr fg = GetForegroundWindow();
        uint fgThread = GetWindowThreadProcessId(fg, IntPtr.Zero), me = GetCurrentThreadId();
        bool attached = (fg != IntPtr.Zero && fgThread != me) && AttachThreadInput(me, fgThread, true);
        SetForegroundWindow(hwnd); BringWindowToTop(hwnd);
        if (attached) AttachThreadInput(me, fgThread, false);
        if (IsForeground(hwnd)) return;
        // ② それでもだめなら Alt を押して前面化の制限を外す (最後の手段。この後のクリックが 1 回無視されることがある)
        keybd_event(0x12, 0, 0, UIntPtr.Zero); SetForegroundWindow(hwnd); keybd_event(0x12, 0, 2, UIntPtr.Zero);
    }
    public static void Click(int x, int y) {
        // 人の操作に近づける: 移動 → 少し待つ (部品がマウスの乗ったことに気づく) → 押す → 少し待つ → 離す
        SetCursorPos(x, y);
        System.Threading.Thread.Sleep(60);
        mouse_event(0x01, 0, 0, 0, UIntPtr.Zero);   // MOUSEEVENTF_MOVE (その場で動かす)
        System.Threading.Thread.Sleep(120);
        mouse_event(0x02, 0, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(80);
        mouse_event(0x04, 0, 0, 0, UIntPtr.Zero);
    }
    public static void MoveTo(int x, int y) { SetCursorPos(x, y); mouse_event(0x01, 0, 0, 0, UIntPtr.Zero); }
    public static void DoubleClick(int x, int y) {
        Click(x, y);
        System.Threading.Thread.Sleep(90);
        mouse_event(0x02, 0, 0, 0, UIntPtr.Zero); mouse_event(0x04, 0, 0, 0, UIntPtr.Zero);
    }
}
'@
}

function Get-IdleSeconds { if ('WtNative' -as [type]) { return [WtNative]::IdleSeconds() } else { return 0.0 } }
function Test-ScreenLocked { if ('WtNative' -as [type]) { return [WtNative]::IsLocked() } else { return $false } }

function Write-ErrorLog([string]$DataDir, [string]$Text) {
    try {
        if (-not $DataDir) { $DataDir = $PSScriptRoot }
        if (-not (Test-Path -LiteralPath $DataDir)) { New-Item -ItemType Directory -Path $DataDir -Force | Out-Null }
        Write-TextFile (Join-Path $DataDir 'error.log') ("--- $((Get-Date).ToString($script:TimeFmt))`r`n$Text`r`n") -Append
    } catch { }
}
