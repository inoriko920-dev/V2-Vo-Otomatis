param(
    [string]$InputPath = "",
    [ValidateRange(1,5)]
    [int]$Batch = 1
)

$ErrorActionPreference = 'Stop'
$TargetUrl = 'https://aistudio.google.com/generate-speech?model=gemini-2.5-pro-preview-tts'
$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$LegacyScript = Join-Path $BaseDir 'ChromeVO_AIStudio_Zoom67_V14.ps1'
$LogDir = Join-Path $BaseDir 'LOG'
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.IO.Compression.FileSystem

Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class AutoVOWin32 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
}
"@

function Write-Status([string]$Message) {
    $stamp = Get-Date -Format 'HH:mm:ss'
    Write-Host "[$stamp] $Message"
}

function Get-ChromeWindows {
    $items = New-Object 'System.Collections.Generic.List[System.IntPtr]'
    $cb = [AutoVOWin32+EnumWindowsProc]{
        param([IntPtr]$hWnd, [IntPtr]$lParam)
        if ([AutoVOWin32]::IsWindowVisible($hWnd)) {
            $sb = New-Object System.Text.StringBuilder 256
            [void][AutoVOWin32]::GetClassName($hWnd, $sb, $sb.Capacity)
            if ($sb.ToString() -eq 'Chrome_WidgetWin_1') { $items.Add($hWnd) }
        }
        return $true
    }
    [AutoVOWin32]::EnumWindows($cb, [IntPtr]::Zero) | Out-Null
    return @($items)
}

function Wait-NewChromeWindow($Before, [int]$TimeoutSec = 30) {
    $beforeSet = @{}
    foreach ($h in $Before) { $beforeSet[$h.ToInt64()] = $true }
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 250
        foreach ($h in @(Get-ChromeWindows)) {
            if (-not $beforeSet.ContainsKey($h.ToInt64())) { return $h }
        }
    }
    return [IntPtr]::Zero
}

function Get-PartNumberFromName([string]$Name) {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    $m = [regex]::Match($base, '(?i)(?:^|[^a-z0-9])part\s*[-_ ]?\s*0*([1-9][0-9]{0,2})(?:[^0-9]|$)')
    if ($m.Success) { return [int]$m.Groups[1].Value }
    $m = [regex]::Match($base, '(?i)^p\s*0*([1-9][0-9]{0,2})(?:[^0-9]|$)')
    if ($m.Success) { return [int]$m.Groups[1].Value }
    return 0
}

function Read-DocxText([string]$Path) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry('word/document.xml')
        if (-not $entry) { throw "word/document.xml tidak ditemukan pada DOCX." }
        $stream = $entry.Open()
        try {
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
            try { $xmlText = $reader.ReadToEnd() } finally { $reader.Dispose() }
        } finally { $stream.Dispose() }
    } finally { $zip.Dispose() }

    [xml]$xml = $xmlText
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $paras = $xml.SelectNodes('//w:body/w:p', $ns)
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($p in $paras) {
        $sb = New-Object System.Text.StringBuilder
        foreach ($node in $p.SelectNodes('.//w:t|.//w:tab|.//w:br|.//w:cr', $ns)) {
            if ($node.LocalName -eq 't') { [void]$sb.Append($node.InnerText) }
            elseif ($node.LocalName -eq 'tab') { [void]$sb.Append("`t") }
            else { [void]$sb.Append("`n") }
        }
        $lines.Add($sb.ToString())
    }
    return (($lines -join "`r`n") -replace "`r?`n{3,}", "`r`n`r`n").Trim()
}

function Read-PartText([string]$Path) {
    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    switch ($ext) {
        '.txt' { return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8).Trim() }
        '.md'  { return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8).Trim() }
        '.docx' { return (Read-DocxText $Path).Trim() }
        default { throw "Format tidak didukung: $ext. Gunakan TXT, MD, atau DOCX." }
    }
}

function Resolve-InputRoot([string]$RequestedPath) {
    if ([string]::IsNullOrWhiteSpace($RequestedPath)) {
        $RequestedPath = Join-Path $BaseDir 'INPUT PART'
    }
    $resolved = (Resolve-Path -LiteralPath $RequestedPath).Path
    if (Test-Path -LiteralPath $resolved -PathType Container) {
        return [pscustomobject]@{ Root = $resolved; Temp = $null }
    }
    $ext = [System.IO.Path]::GetExtension($resolved).ToLowerInvariant()
    if ($ext -eq '.zip') {
        $temp = Join-Path $env:TEMP ("AutoPartVO_" + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $temp -Force | Out-Null
        [System.IO.Compression.ZipFile]::ExtractToDirectory($resolved, $temp)
        return [pscustomobject]@{ Root = $temp; Temp = $temp }
    }
    if ($ext -in @('.txt','.md','.docx')) {
        return [pscustomobject]@{ Root = (Split-Path -Parent $resolved); Temp = $null; SingleFile = $resolved }
    }
    throw 'Input harus berupa folder, ZIP, TXT, MD, atau DOCX.'
}

function Find-PartFiles($InputInfo) {
    $files = if ($InputInfo.SingleFile) {
        @(Get-Item -LiteralPath $InputInfo.SingleFile)
    } else {
        @(Get-ChildItem -LiteralPath $InputInfo.Root -File -Recurse | Where-Object { $_.Extension.ToLowerInvariant() -in @('.txt','.md','.docx') })
    }

    $map = @{}
    foreach ($f in $files) {
        $n = Get-PartNumberFromName $f.Name
        if ($n -lt 1) { continue }
        if ($map.ContainsKey($n)) {
            throw "Ada lebih dari satu file untuk Part ${n}:`r`n$($map[$n].FullName)`r`n$($f.FullName)"
        }
        $map[$n] = $f
    }
    return $map
}

. (Join-Path $BaseDir 'AIStudioUI_V14.ps1')

function Focus-NarrationWindow([IntPtr]$Hwnd) {
    [AutoVOWin32]::ShowWindow($Hwnd, 9) | Out-Null
    [AutoVOWin32]::BringWindowToTop($Hwnd) | Out-Null
    [AutoVOWin32]::SetForegroundWindow($Hwnd) | Out-Null
    Start-Sleep -Milliseconds 350
    if ([AutoVOWin32]::GetForegroundWindow() -ne $Hwnd) {
        throw 'Jendela VO tidak mendapat fokus. Tidak mengirim shortcut ke jendela lain.'
    }
}

function Fill-Narration([IntPtr]$Hwnd, [string]$Text, [string]$VoLabel) {
    Focus-NarrationWindow $Hwnd
    $editor = Ensure-ComposerEditor $Hwnd
    if (-not $editor.Ok) { return $editor }
    Set-AIStudioZoom67 $Hwnd { param($h) Focus-NarrationWindow $h }
    if (-not (Set-VerifiedNarration $Hwnd $Text { param($h) Focus-NarrationWindow $h })) {
        return [pscustomobject]@{ Ok=$false; Message='Teks Part tidak terverifikasi utuh pada Speech block text. Periksa jendela VO; status tidak dianggap TERISI.' }
    }
    if (-not (Test-AIStudioZoom67 $Hwnd)) {
        return [pscustomobject]@{ Ok=$false; Message='Narasi terisi tetapi zoom akhir bukan 67%.' }
    }
    return [pscustomobject]@{ Ok=$true; Message="${VoLabel}: narasi terverifikasi utuh; Composer aktif; zoom 67% terverifikasi. Tidak memilih voice dan tidak menekan Run/Generate." }
}
function Start-VOGroup([int]$StartVO, [int]$Count) {
    if (-not (Test-Path -LiteralPath $LegacyScript)) { throw "Script pembuka VO tidak ditemukan: $LegacyScript" }
    if ($Count -lt 1 -or $Count -gt 6) { throw "Jumlah window harus 1-6." }
    if (($StartVO + $Count - 1) -gt 30) { throw "Range VO melebihi VO30." }

    $handleFile = Join-Path $env:TEMP ("AutoVO_Group_" + [guid]::NewGuid().ToString('N') + '.txt')
    try {
        $qScript = '"' + $LegacyScript.Replace('"','\"') + '"'
        $qHandle = '"' + $handleFile.Replace('"','\"') + '"'
        $args = @(
            '-NoProfile','-STA','-ExecutionPolicy','Bypass','-WindowStyle','Hidden',
            '-File', $qScript,
            '-OpenAll',
            '-OpenStartVO', $StartVO,
            '-OpenCount', $Count,
            '-HandleOutFile', $qHandle,
            '-SuppressErrorDialog'
        )

        $endVO = $StartVO + $Count - 1
        Write-Status "Membuka VO$StartVO sampai VO$endVO sekaligus..."
        $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $args -WindowStyle Hidden -PassThru
        $deadline = (Get-Date).AddSeconds(600)
        $payload = ''
        while ((Get-Date) -lt $deadline) {
            if (Test-Path -LiteralPath $handleFile) {
                $payload = (Get-Content -LiteralPath $handleFile -Raw -ErrorAction SilentlyContinue).Trim()
                if ($payload) { break }
            }
            if ($proc.HasExited -and -not (Test-Path -LiteralPath $handleFile)) { break }
            Start-Sleep -Milliseconds 250
        }

        if (-not $payload -and (Test-Path -LiteralPath $handleFile)) {
            $payload = (Get-Content -LiteralPath $handleFile -Raw -ErrorAction SilentlyContinue).Trim()
        }
        if ($payload -like 'ERROR:*') { throw ($payload.Substring(6)) }
        if ([string]::IsNullOrWhiteSpace($payload)) {
            if ($proc.HasExited -and $proc.ExitCode -ne 0) {
                throw "Pembuka VO berhenti dengan kode $($proc.ExitCode), tetapi detail error tidak tersedia."
            }
            throw "VO$StartVO-VO$endVO belum mengirim identitas jendela setelah 600 detik."
        }

        $handles = @{}
        foreach ($line in ($payload -split '\r?\n')) {
            $m = [regex]::Match($line.Trim(), '^VO(30|2[0-9]|1[0-9]|[1-9])=(\d+)$')
            if (-not $m.Success) { continue }
            $vo = [int]$m.Groups[1].Value
            [long]$raw = 0
            if ([long]::TryParse($m.Groups[2].Value, [ref]$raw) -and $raw -gt 0) {
                $handles[$vo] = [IntPtr]::new($raw)
            }
        }

        $needed = @($StartVO..$endVO)
        $missing = @($needed | Where-Object { -not $handles.ContainsKey($_) })
        if ($missing.Count -gt 0) {
            throw ('Handle jendela belum lengkap. Tidak terbaca: ' + (($missing | ForEach-Object { 'VO'+$_ }) -join ', '))
        }
        Write-Status "VO$StartVO-VO$endVO sudah terbuka."
        return $handles
    }
    finally {
        Remove-Item -LiteralPath $handleFile -Force -ErrorAction SilentlyContinue
    }
}

$StatePath = Join-Path $BaseDir '.V14_BATCH_STATE.json'

function Save-BatchState([int]$BatchNumber, [hashtable]$Handles, [string]$SourcePath, [int[]]$OrderedParts) {
    $items = @()
    foreach ($key in @($Handles.Keys | Sort-Object)) {
        $items += [pscustomobject]@{
            VO = [int]$key
            Hwnd = [long]$Handles[$key].ToInt64()
        }
    }
    $obj = [pscustomobject]@{
        version = 'V14'
        batch = $BatchNumber
        sourcePath = $SourcePath
        orderedParts = @($OrderedParts)
        savedAt = (Get-Date).ToString('o')
        windows = $items
    }
    [System.IO.File]::WriteAllText($StatePath, ($obj | ConvertTo-Json -Depth 6), [System.Text.UTF8Encoding]::new($false))
}

function Close-PreviousBatchWindows {
    if (-not (Test-Path -LiteralPath $StatePath)) {
        Write-Status 'Tidak ada state batch lama yang perlu ditutup.'
        return 0
    }

    try {
        $state = Get-Content -LiteralPath $StatePath -Raw -Encoding UTF8 | ConvertFrom-Json
        $windows = @($state.windows)
    } catch {
        Write-Status "State lama tidak dapat dibaca: $($_.Exception.Message)"
        Remove-Item -LiteralPath $StatePath -Force -ErrorAction SilentlyContinue
        return 0
    }

    $sent = 0
    foreach ($w in $windows) {
        [long]$raw = 0
        if (-not [long]::TryParse([string]$w.Hwnd, [ref]$raw)) { continue }
        $h = [IntPtr]::new($raw)
        if ($raw -gt 0 -and [AutoVOWin32]::IsWindow($h)) {
            [void][AutoVOWin32]::PostMessage($h, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
            $sent += 1
        }
    }

    if ($sent -gt 0) {
        Write-Status "Menutup $sent window batch sebelumnya sekaligus..."
        $deadline = (Get-Date).AddSeconds(15)
        while ((Get-Date) -lt $deadline) {
            $left = 0
            foreach ($w in $windows) {
                [long]$raw = 0
                if ([long]::TryParse([string]$w.Hwnd, [ref]$raw) -and $raw -gt 0) {
                    if ([AutoVOWin32]::IsWindow([IntPtr]::new($raw))) { $left += 1 }
                }
            }
            if ($left -eq 0) { break }
            Start-Sleep -Milliseconds 300
        }
    }

    Remove-Item -LiteralPath $StatePath -Force -ErrorAction SilentlyContinue
    return $sent
}

function Get-OrderedAvailableParts($PartMap) {
    return @($PartMap.Keys | ForEach-Object { [int]$_ } | Sort-Object)
}

function Get-BatchPlan([int]$BatchNumber, [int[]]$OrderedParts, $PartMap) {
    if ($BatchNumber -lt 1 -or $BatchNumber -gt 5) { throw 'Nomor batch harus 1-5.' }

    $startIndex = ($BatchNumber - 1) * 3
    if ($startIndex -ge $OrderedParts.Count) {
        return $null
    }

    # VO1-VO30 = kapasitas maksimal 15 Part (2 VO per Part).
    if ($startIndex -ge 15) { return $null }

    $remainingCapacity = 15 - $startIndex
    $take = [math]::Min(3, [math]::Min($remainingCapacity, $OrderedParts.Count - $startIndex))
    $parts = @($OrderedParts[$startIndex..($startIndex + $take - 1)])
    $voStart = 1 + (($BatchNumber - 1) * 6)
    $openCount = $parts.Count * 2

    return [pscustomobject]@{
        Batch = $BatchNumber
        StartIndex = $startIndex
        Parts = $parts
        VOStart = $voStart
        OpenCount = $openCount
        VOEnd = $voStart + $openCount - 1
    }
}

function Get-SourcePathForRun([string]$RequestedPath) {
    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        return (Resolve-Path -LiteralPath $RequestedPath).Path
    }
    return (Resolve-Path -LiteralPath (Join-Path $BaseDir 'INPUT PART')).Path
}

$inputInfo = $null
$results = New-Object System.Collections.Generic.List[object]
$plan = $null
$batchNumber = $Batch
$sourcePath = $null
$orderedParts = @()

try {
    $sourcePath = Get-SourcePathForRun $InputPath
    $inputInfo = Resolve-InputRoot $sourcePath
    $partMap = Find-PartFiles $inputInfo
    $orderedParts = @(Get-OrderedAvailableParts $partMap)

    if ($orderedParts.Count -eq 0) {
        throw 'Tidak ditemukan file Part. Nama file harus mengandung Part1, Part2, Part7, dst.'
    }

    if ($orderedParts.Count -gt 15) {
        Write-Status "Ditemukan $($orderedParts.Count) Part. V14 VO1-VO30 hanya memproses 15 Part pertama yang tersedia."
        $orderedParts = @($orderedParts | Select-Object -First 15)
    }

    $plan = Get-BatchPlan $batchNumber $orderedParts $partMap

    # Setiap tombol V14 selalu menutup batch V14 yang sedang terbuka terlebih dahulu.
    [void](Close-PreviousBatchWindows)
    Start-Sleep -Milliseconds 700

    if (-not $plan) {
        $rangeStart = (($batchNumber - 1) * 3) + 1
        $rangeEnd = $rangeStart + 2
        $voStartEmpty = 1 + (($batchNumber - 1) * 6)
        $voEndEmpty = $voStartEmpty + 5
        [System.Windows.Forms.MessageBox]::Show(
            "Batch $batchNumber adalah urutan narasi ke-$rangeStart sampai ke-$rangeEnd.`r`n`r`nPart tersedia hanya $($orderedParts.Count), jadi tidak ada narasi untuk batch ini.`r`n`r`nWindow V14 yang sebelumnya terbuka sudah ditutup. VO$voStartEmpty-VO$voEndEmpty tidak dibuka.",
            'AUDIT V14 - Tidak Ada Part Untuk Batch Ini',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        exit 0
    }

    $partTexts = @{}
    foreach ($part in $plan.Parts) {
        $file = $partMap[$part]
        $text = Read-PartText $file.FullName
        if ([string]::IsNullOrWhiteSpace($text)) { throw "Part $part kosong: $($file.Name)" }
        $partTexts[$part] = $text
    }

    $allNames = ($orderedParts | ForEach-Object { 'Part'+$_ }) -join ', '
    $batchNames = ($plan.Parts | ForEach-Object { 'Part'+$_ }) -join ', '
    Write-Status "Part tersedia (urut): $allNames"
    Write-Status "Batch $batchNumber memakai: $batchNames"
    Write-Status 'Aturan V14: nomor VO mengikuti URUTAN Part yang tersedia, bukan angka Part.'
    Write-Status 'Mode: buka VO yang diperlukan, isi narasi, lalu berhenti. Tidak memilih voice dan tidak Generate.'

    for ($i = 0; $i -lt $plan.Parts.Count; $i++) {
        $part = [int]$plan.Parts[$i]
        $voA = $plan.VOStart + ($i * 2)
        $voB = $voA + 1
        $file = $partMap[$part]
        Write-Status "Part $part -> VO$voA + VO$voB | $($file.Name) | $($partTexts[$part].Length) karakter"
    }

    $handles = Start-VOGroup $plan.VOStart $plan.OpenCount
    Save-BatchState $batchNumber $handles $sourcePath $orderedParts

    for ($i = 0; $i -lt $plan.Parts.Count; $i++) {
        $part = [int]$plan.Parts[$i]
        $text = $partTexts[$part]
        $voA = $plan.VOStart + ($i * 2)
        $voB = $voA + 1

        foreach ($vo in @($voA,$voB)) {
            try {
                Write-Status "Mengisi Part $part ke VO$vo..."
                $r = Fill-Narration $handles[$vo] $text "VO$vo"
                if ($r.Ok) {
                    $results.Add([pscustomobject]@{Part=$part; VO="VO$vo"; Status='TERISI'; Detail=$r.Message})
                    Write-Status "VO${vo}: Part $part berhasil dimasukkan."
                } else {
                    $results.Add([pscustomobject]@{Part=$part; VO="VO$vo"; Status='GAGAL'; Detail=$r.Message})
                    Write-Status "VO${vo} GAGAL: $($r.Message)"
                }
            } catch {
                $results.Add([pscustomobject]@{Part=$part; VO="VO$vo"; Status='GAGAL'; Detail=$_.Exception.Message})
                Write-Status "VO${vo} GAGAL: $($_.Exception.Message)"
            }
        }
    }
}
finally {
    if ($inputInfo -and $inputInfo.Temp -and (Test-Path -LiteralPath $inputInfo.Temp)) {
        Remove-Item -LiteralPath $inputInfo.Temp -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$logPath = Join-Path $LogDir ("AUTO ISI NARASI BATCH $batchNumber " + (Get-Date -Format 'yyyy-MM-dd HH-mm-ss') + '.txt')
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('AUDIT V14 - 5 BATCH VO1-VO30 - PART BERURUTAN OTOMATIS')
$lines.Add('')
$lines.Add('Aturan: VO mengikuti urutan Part yang TERSEDIA, bukan nomor Part.')
$lines.Add('Part tersedia: ' + (($orderedParts | ForEach-Object { 'Part'+$_ }) -join ', '))
if ($plan) {
    $lines.Add("Batch: $batchNumber | VO$($plan.VOStart)-VO$($plan.VOEnd)")
    for ($i = 0; $i -lt $plan.Parts.Count; $i++) {
        $part = [int]$plan.Parts[$i]
        $voA = $plan.VOStart + ($i * 2)
        $voB = $voA + 1
        $lines.Add("Mapping: Part$part -> VO$voA + VO$voB")
    }
}
$lines.Add('Target: Gemini 2.5 Pro Preview TTS')
$lines.Add('Berhenti setelah narasi dimasukkan. Tidak memilih voice, tidak Run/Generate, tidak cek durasi, tidak download.')
$lines.Add('')
foreach ($r in $results) { $lines.Add(("Part {0} | {1} | {2} | {3}" -f $r.Part,$r.VO,$r.Status,$r.Detail)) }
[System.IO.File]::WriteAllLines($logPath, $lines, [System.Text.UTF8Encoding]::new($true))

$ok = @($results | Where-Object {$_.Status -eq 'TERISI'}).Count
$fail = @($results | Where-Object {$_.Status -eq 'GAGAL'}).Count
Write-Status "Selesai batch $batchNumber. TERISI=$ok | GAGAL=$fail"
Write-Status "Log: $logPath"

$rangeStart = (($batchNumber - 1) * 3) + 1
$rangeEnd = $rangeStart + 2
$buttonText = if ($batchNumber -lt 5) {
    "Untuk kelompok berikutnya, klik tombol " + ($batchNumber + 1) + ". Saat diklik, batch ini akan ditutup otomatis."
} else {
    "Ini batch 5 / VO25-VO30. Tidak ada kelompok VO setelah ini."
}

[System.Windows.Forms.MessageBox]::Show(
    "Isi narasi batch $batchNumber selesai.`r`n`r`nNarasi terisi: $ok`r`nGagal: $fail`r`n`r`nBatch $batchNumber = urutan Part tersedia ke-$rangeStart sampai ke-$rangeEnd.`r`n$buttonText`r`n`r`nTidak memilih voice dan tidak menekan Run/Generate.`r`n`r`nLog:`r`n$logPath",
    'AUDIT V14 - Batch Part Tersedia',
    [System.Windows.Forms.MessageBoxButtons]::OK,
    $(if ($fail -gt 0) {[System.Windows.Forms.MessageBoxIcon]::Warning} else {[System.Windows.Forms.MessageBoxIcon]::Information})
) | Out-Null
