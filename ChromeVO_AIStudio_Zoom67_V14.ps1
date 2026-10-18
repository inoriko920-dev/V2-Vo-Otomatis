param(
    [string]$ProfileName = "",
    [ValidateRange(1,6)]
    [int]$Slot = 1,
    [switch]$OpenAll,
    [ValidateRange(1,30)]
    [int]$OpenStartVO = 1,
    [ValidateRange(1,6)]
    [int]$OpenCount = 6,
    [switch]$AutoEmptySlot,
    [string]$HandleOutFile = "",
    [switch]$SuppressErrorDialog
)

$ErrorActionPreference = "Stop"
$TargetUrl = "https://aistudio.google.com/generate-speech?model=gemini-2.5-pro-preview-tts"

Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class VOChromeWin32 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool MoveWindow(
        IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint
    );

    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr lpdwProcessId);

    [DllImport("kernel32.dll")]
    public static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
}
"@

[VOChromeWin32]::SetProcessDPIAware() | Out-Null
Add-Type -AssemblyName System.Windows.Forms
. (Join-Path $PSScriptRoot 'AIStudioUI_V14.ps1')

function Show-Error([string]$Message) {
    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        "Chrome VO Layout V14",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Get-ChromePath {
    $candidatePaths = @(
        (Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe")
        $(if (${env:ProgramFiles(x86)}) { Join-Path ${env:ProgramFiles(x86)} "Google\Chrome\Application\chrome.exe" })
        (Join-Path $env:LOCALAPPDATA "Google\Chrome\Application\chrome.exe")
    )

    $found = @(
        $candidatePaths | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
    )

    if ($found.Count -eq 0) {
        throw "Google Chrome tidak ditemukan."
    }

    return [string]$found[0]
}

function Get-LocalState {
    $path = Join-Path $env:LOCALAPPDATA "Google\Chrome\User Data\Local State"
    if (-not (Test-Path -LiteralPath $path)) {
        throw "File profil Chrome tidak ditemukan:`r`n$path"
    }
    return (Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json)
}

function Find-ProfileDirectory($State, [string]$VisibleName) {
    foreach ($p in $State.profile.info_cache.PSObject.Properties) {
        if ([string]$p.Value.name -eq $VisibleName) {
            return [string]$p.Name
        }
    }

    $available = @(
        $State.profile.info_cache.PSObject.Properties |
        ForEach-Object { "$($_.Name) = $([string]$_.Value.name)" }
    ) -join "`r`n"

    throw "Profil '$VisibleName' tidak ditemukan.`r`n`r`nProfil yang terbaca:`r`n$available"
}

function Get-ChromeWindowHandles {
    $handles = New-Object 'System.Collections.Generic.List[System.IntPtr]'

    $callback = [VOChromeWin32+EnumWindowsProc]{
        param([IntPtr]$hWnd, [IntPtr]$lParam)

        if ([VOChromeWin32]::IsWindowVisible($hWnd)) {
            $sb = New-Object System.Text.StringBuilder 256
            [void][VOChromeWin32]::GetClassName($hWnd, $sb, $sb.Capacity)

            if ($sb.ToString() -eq "Chrome_WidgetWin_1") {
                $handles.Add($hWnd)
            }
        }
        return $true
    }

    [VOChromeWin32]::EnumWindows($callback, [IntPtr]::Zero) | Out-Null
    return @($handles)
}

function Get-SlotRects {
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $gap = 6

    $usableW = $wa.Width - (2 * $gap)
    $baseW = [math]::Floor($usableW / 3)

    $usableH = $wa.Height - $gap
    $baseH = [math]::Floor($usableH / 2)

    $slots = @()

    # Nomor slot disusun per pasangan vertikal:
    # Slot 1 = kiri atas,   Slot 2 = kiri bawah
    # Slot 3 = tengah atas, Slot 4 = tengah bawah
    # Slot 5 = kanan atas,  Slot 6 = kanan bawah
    # Dengan begitu VO1/VO2, VO3/VO4, VO5/VO6 selalu atas-bawah, bukan kiri-kanan.
    for ($slot = 1; $slot -le 6; $slot++) {
        $col = [math]::Floor(($slot - 1) / 2)
        $row = ($slot - 1) % 2

        $x = $wa.Left + ($col * ($baseW + $gap))
        $y = $wa.Top  + ($row * ($baseH + $gap))

        if ($col -eq 2) { $w = $wa.Right - $x } else { $w = $baseW }
        if ($row -eq 1) { $h = $wa.Bottom - $y } else { $h = $baseH }

        $slots += [pscustomobject]@{
            Slot   = $slot
            Left   = [int]$x
            Top    = [int]$y
            Right  = [int]($x + $w)
            Bottom = [int]($y + $h)
            Width  = [int]$w
            Height = [int]$h
        }
    }

    return $slots
}

function Get-OverlapRatio($WindowRect, $SlotRect) {
    $left   = [math]::Max($WindowRect.Left,   $SlotRect.Left)
    $top    = [math]::Max($WindowRect.Top,    $SlotRect.Top)
    $right  = [math]::Min($WindowRect.Right,  $SlotRect.Right)
    $bottom = [math]::Min($WindowRect.Bottom, $SlotRect.Bottom)

    $iw = [math]::Max(0, $right - $left)
    $ih = [math]::Max(0, $bottom - $top)

    $intersection = $iw * $ih
    $slotArea = $SlotRect.Width * $SlotRect.Height

    if ($slotArea -le 0) { return 0.0 }
    return [double]$intersection / [double]$slotArea
}

function Find-FirstEmptySlot {
    $slots = @(Get-SlotRects)

    $occupied = @{}
    for ($i = 1; $i -le 6; $i++) {
        $occupied[$i] = $false
    }

    foreach ($h in @(Get-ChromeWindowHandles)) {
        $r = New-Object VOChromeWin32+RECT
        if (-not [VOChromeWin32]::GetWindowRect($h, [ref]$r)) {
            continue
        }

        # Abaikan window Chrome yang terlalu kecil/minimized.
        $ww = $r.Right - $r.Left
        $wh = $r.Bottom - $r.Top
        if ($ww -lt 250 -or $wh -lt 200) {
            continue
        }

        # Cari slot dengan overlap terbesar.
        $bestSlot = 0
        $bestRatio = 0.0

        foreach ($s in $slots) {
            $ratio = Get-OverlapRatio $r $s
            if ($ratio -gt $bestRatio) {
                $bestRatio = $ratio
                $bestSlot = $s.Slot
            }
        }

        # Window hasil layout biasanya overlap ~100%.
        # Ambang 45% cukup toleran jika user sedikit resize/geser.
        if ($bestSlot -gt 0 -and $bestRatio -ge 0.45) {
            $occupied[$bestSlot] = $true
        }
    }

    for ($i = 1; $i -le 6; $i++) {
        if (-not $occupied[$i]) {
            return $i
        }
    }

    return 0
}

function Force-Foreground([IntPtr]$hWnd) {
    [VOChromeWin32]::ShowWindow($hWnd, 9) | Out-Null
    [VOChromeWin32]::BringWindowToTop($hWnd) | Out-Null

    $foreground = [VOChromeWin32]::GetForegroundWindow()
    $currentThread = [VOChromeWin32]::GetCurrentThreadId()
    $foregroundThread = 0

    if ($foreground -ne [IntPtr]::Zero) {
        $foregroundThread = [VOChromeWin32]::GetWindowThreadProcessId($foreground, [IntPtr]::Zero)
    }

    $attached = $false
    try {
        if ($foregroundThread -ne 0 -and $foregroundThread -ne $currentThread) {
            $attached = [VOChromeWin32]::AttachThreadInput($currentThread, $foregroundThread, $true)
        }

        [VOChromeWin32]::BringWindowToTop($hWnd) | Out-Null
        [VOChromeWin32]::SetForegroundWindow($hWnd) | Out-Null
    }
    finally {
        if ($attached) {
            [VOChromeWin32]::AttachThreadInput($currentThread, $foregroundThread, $false) | Out-Null
        }
    }

    Start-Sleep -Milliseconds 350
    if ([VOChromeWin32]::GetForegroundWindow() -ne $hWnd) {
        throw "Jendela VO tidak mendapat fokus. Shortcut zoom tidak dikirim."
    }
}

function Move-ToSlot([IntPtr]$hWnd, [int]$TargetSlot) {
    $slotRects = @(Get-SlotRects)
    $s = $slotRects | Where-Object { $_.Slot -eq $TargetSlot } | Select-Object -First 1

    [VOChromeWin32]::ShowWindow($hWnd, 9) | Out-Null
    Start-Sleep -Milliseconds 120

    [VOChromeWin32]::MoveWindow(
        $hWnd,
        [int]$s.Left,
        [int]$s.Top,
        [int]$s.Width,
        [int]$s.Height,
        $true
    ) | Out-Null

    Start-Sleep -Milliseconds 300

    [VOChromeWin32]::MoveWindow(
        $hWnd,
        [int]$s.Left,
        [int]$s.Top,
        [int]$s.Width,
        [int]$s.Height,
        $true
    ) | Out-Null
}

function Set-Zoom67([IntPtr]$hWnd) {
    Set-AIStudioZoom67 $hWnd { param($h) Force-Foreground $h }
}
function Open-OneVO(
    [string]$VisibleProfileName,
    [int]$TargetSlot,
    [string]$ChromePath,
    $State
) {
    $profileDir = Find-ProfileDirectory $State $VisibleProfileName

    $before = @{}
    foreach ($h in @(Get-ChromeWindowHandles)) {
        $before[$h.ToInt64()] = $true
    }

    $argProfile = '--profile-directory="{0}"' -f $profileDir

    Start-Process -FilePath $ChromePath -ArgumentList @(
        $argProfile,
        "--new-window",
        "--force-renderer-accessibility",
        $TargetUrl
    ) | Out-Null

    $newHwnd = [IntPtr]::Zero
    $deadline = (Get-Date).AddSeconds(25)

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 150

        foreach ($h in @(Get-ChromeWindowHandles)) {
            if (-not $before.ContainsKey($h.ToInt64())) {
                $newHwnd = $h
                break
            }
        }

        if ($newHwnd -ne [IntPtr]::Zero) {
            break
        }
    }

    if ($newHwnd -eq [IntPtr]::Zero) {
        throw "Window baru untuk '$VisibleProfileName' tidak berhasil dideteksi."
    }

    Start-Sleep -Milliseconds 550
    Move-ToSlot $newHwnd $TargetSlot

    Start-Sleep -Milliseconds 450
    $editor = Ensure-ComposerEditor $newHwnd
    if (-not $editor.Ok) { throw "${VisibleProfileName}: $($editor.Message)" }
    Set-Zoom67 $newHwnd

    Move-ToSlot $newHwnd $TargetSlot
    return $newHwnd
}

try {
    $chromePath = Get-ChromePath
    $state = Get-LocalState

    if ($OpenAll) {
        if (($OpenStartVO + $OpenCount - 1) -gt 30) {
            throw "Range VO melebihi VO30: mulai VO$OpenStartVO sebanyak $OpenCount window."
        }
        $opened = New-Object System.Collections.Generic.List[object]
        for ($slotIndex = 1; $slotIndex -le $OpenCount; $slotIndex++) {
            $voNumber = $OpenStartVO + $slotIndex - 1
            $h = Open-OneVO "$voNumber. VO" $slotIndex $chromePath $state
            $opened.Add([pscustomobject]@{ VO=$voNumber; Hwnd=$h.ToInt64(); Slot=$slotIndex })
            Start-Sleep -Milliseconds 250
        }
        if (-not [string]::IsNullOrWhiteSpace($HandleOutFile)) {
            $payload = ($opened | ForEach-Object { "VO$($_.VO)=$($_.Hwnd)" }) -join "`r`n"
            [System.IO.File]::WriteAllText($HandleOutFile, $payload, [System.Text.UTF8Encoding]::new($false))
        }
    }
    else {
        if ([string]::IsNullOrWhiteSpace($ProfileName)) {
            throw "ProfileName kosong."
        }

        $target = $Slot

        if ($AutoEmptySlot) {
            # Cari slot kosong SEBELUM window baru dibuka.
            $target = Find-FirstEmptySlot

            if ($target -eq 0) {
                throw "Semua 6 posisi layout sedang terisi. Tutup salah satu window VO terlebih dahulu, lalu buka kembali '$ProfileName'."
            }
        }

        $openedHwnd = Open-OneVO $ProfileName $target $chromePath $state
        if (-not [string]::IsNullOrWhiteSpace($HandleOutFile)) {
            [System.IO.File]::WriteAllText($HandleOutFile, ([string]$openedHwnd.ToInt64()), [System.Text.UTF8Encoding]::new($false))
        }
    }
}
catch {
    $msg = $_.Exception.Message
    if (-not [string]::IsNullOrWhiteSpace($HandleOutFile)) {
        try { [System.IO.File]::WriteAllText($HandleOutFile, ("ERROR:" + $msg), [System.Text.UTF8Encoding]::new($false)) } catch {}
    }
    if (-not $SuppressErrorDialog) { Show-Error $msg }
    exit 1
}
