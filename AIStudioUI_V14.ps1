# Kontrol AI Studio yang diperiksa langsung pada UI Composer, 18 Oktober 2026.
# Tidak memilih voice, memainkan preview, atau menjalankan generasi audio.
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms

function Get-UiaRoot([IntPtr]$Hwnd) {
    return [System.Windows.Automation.AutomationElement]::FromHandle($Hwnd)
}

function Get-UiaElements($Root) {
    if (-not $Root) { return @() }
    return @($Root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition))
}

function Find-UiaByNameRegex($Elements, [string]$Pattern, [string[]]$ControlTypes = @()) {
    foreach ($el in $Elements) {
        try {
            if ([string]$el.Current.Name -notmatch $Pattern) { continue }
            if ($ControlTypes.Count -gt 0) {
                $type = $el.Current.ControlType.ProgrammaticName.Split('.')[-1]
                if ($type -notin $ControlTypes) { continue }
            }
            return $el
        } catch {}
    }
    return $null
}

function Invoke-Uia($Element) {
    if (-not $Element) { return $false }
    try {
        if (-not $Element.Current.IsEnabled) { return $false }
        $p = $Element.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
        if ($p) {
            if (-not $p.Current.IsSelected) { $p.Select() }
            return $true
        }
    } catch {}
    try {
        $p = $Element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
        if ($p) { $p.Invoke(); return $true }
    } catch {}
    # Jangan kirim Enter ke kontrol yang tidak mendukung Invoke/Select.
    return $false
}

function Get-UiaText($Element) {
    try {
        $p = $Element.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
        if ($p) { return [pscustomobject]@{ Readable=$true; Text=[string]$p.Current.Value } }
    } catch {}
    try {
        $p = $Element.GetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern)
        if ($p) { return [pscustomobject]@{ Readable=$true; Text=[string]$p.DocumentRange.GetText(-1) } }
    } catch {}
    return [pscustomobject]@{ Readable=$false; Text='' }
}

function Normalize-Narration([string]$Text) {
    return $Text.Replace("`r`n", "`n").Replace("`r", "`n").TrimEnd([char[]]@("`n"))
}

function Test-AIStudioZoom67([IntPtr]$Hwnd) {
    $els = Get-UiaElements (Get-UiaRoot $Hwnd)
    return $null -ne (Find-UiaByNameRegex $els '^(Zoom|Zoom level|Pembesaran|Perbesaran)\s*:\s*67\s*%$' @('Button'))
}

function Set-AIStudioZoom67([IntPtr]$Hwnd, [scriptblock]$Activate) {
    # Jalankan setelah halaman AI Studio siap, bukan ketika tab masih about:blank.
    if (Test-AIStudioZoom67 $Hwnd) { return }
    for ($attempt = 0; $attempt -lt 2; $attempt++) {
        & $Activate $Hwnd
        [System.Windows.Forms.SendKeys]::SendWait('^0')
        Start-Sleep -Milliseconds 250
        # Tangga zoom Chrome: 100 -> 90 -> 80 -> 75 -> 67.
        for ($i = 0; $i -lt 4; $i++) {
            & $Activate $Hwnd
            [System.Windows.Forms.SendKeys]::SendWait('^-')
            Start-Sleep -Milliseconds 180
        }
        $deadline = (Get-Date).AddSeconds(4)
        do {
            if (Test-AIStudioZoom67 $Hwnd) { return }
            Start-Sleep -Milliseconds 250
        } while ((Get-Date) -lt $deadline)
    }
    throw 'Zoom Chrome belum terverifikasi 67%. Pengisian dihentikan agar tidak memakai tampilan yang salah.'
}

function Ensure-ComposerEditor([IntPtr]$Hwnd, [int]$TimeoutSec = 60) {
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $composerRequested = $false
    $dialogRequested = $false
    while ((Get-Date) -lt $deadline) {
        try {
            $root = Get-UiaRoot $Hwnd
            $els = Get-UiaElements $root
            $login = Find-UiaByNameRegex $els '^(Sign in|Login|Log in)$' @('Button','Hyperlink')
            if ($login) {
                return [pscustomobject]@{ Ok=$false; NeedsLogin=$true; Message='Akun belum login Google. Login manual pada profil VO ini, lalu jalankan batch lagi.' }
            }

            # Pada UI baru, kartu contoh menjalankan preview. Pakai dialog kosong.
            $create = Find-UiaByNameRegex $els '^Create new dialog$' @('Button')
            if ($create -and -not $dialogRequested) {
                if (Invoke-Uia $create) { $dialogRequested = $true }
                Start-Sleep -Milliseconds 650
                continue # Selalu baca ulang setelah navigasi.
            }

            $prompt = Find-UiaByNameRegex $els '^Speech block text$' @('Edit','Document')
            $scene = Find-UiaByNameRegex $els '^Scene$' @('Edit')
            $context = Find-UiaByNameRegex $els '^Sample Context$' @('Edit')
            if ($prompt -and $scene -and $context) {
                # Tutup panel hanya bila benar-benar tampak; jangan toggle panel tertutup.
                $closeNav = Find-UiaByNameRegex $els '^Close navigation menu$' @('Button')
                if ($closeNav -and (Invoke-Uia $closeNav)) {
                    Start-Sleep -Milliseconds 350
                    continue
                }
                $settings = Find-UiaByNameRegex $els '^Speaker settings$' @('Group')
                if ($settings -and -not $settings.Current.IsOffscreen -and
                    $settings.Current.BoundingRectangle.Width -gt 0 -and
                    $settings.Current.BoundingRectangle.Height -gt 0) {
                    $children = Get-UiaElements $settings
                    $toggle = Find-UiaByNameRegex $els '^Toggle run settings panel$' @('Button')
                    if ($children.Count -gt 0 -and $toggle -and (Invoke-Uia $toggle)) {
                        Start-Sleep -Milliseconds 350
                        continue
                    }
                }
                return [pscustomobject]@{ Ok=$true; Root=$root; Elements=$els; Prompt=$prompt }
            }

            $composer = Find-UiaByNameRegex $els '^Composer$' @('RadioButton','Button','TabItem')
            if ($composer -and -not $composerRequested) {
                if (Invoke-Uia $composer) { $composerRequested = $true }
                Start-Sleep -Milliseconds 500
                continue
            }
        } catch {
            # ElementNotAvailable saat Angular mengganti halaman: tunggu dan ambil ulang.
        }
        Start-Sleep -Milliseconds 500
    }
    return [pscustomobject]@{ Ok=$false; Message='Composer belum siap: Create new dialog / Scene / Sample Context / Speech block text tidak lengkap. Tidak memilih kartu contoh atau kolom lain.' }
}

function Test-NarrationValue([IntPtr]$Hwnd, [string]$Text) {
    $expected = Normalize-Narration $Text
    for ($i = 0; $i -lt 10; $i++) {
        $els = Get-UiaElements (Get-UiaRoot $Hwnd)
        $prompt = Find-UiaByNameRegex $els '^Speech block text$' @('Edit','Document')
        if ($prompt) {
            $actual = Get-UiaText $prompt
            if ($actual.Readable -and (Normalize-Narration $actual.Text) -ceq $expected) { return $true }
        }
        Start-Sleep -Milliseconds 200
    }
    return $false
}

function Set-VerifiedNarration([IntPtr]$Hwnd, [string]$Text, [scriptblock]$Activate) {
    $editor = Ensure-ComposerEditor $Hwnd
    if (-not $editor.Ok) { return $false }
    try {
        $vp = $editor.Prompt.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
        if ($vp -and -not $vp.Current.IsReadOnly) {
            $vp.SetValue($Text)
            if (Test-NarrationValue $Hwnd $Text) { return $true }
        }
    } catch {}

    $clipboard = $null
    $clipboardSaved = $false
    try {
        $editor = Ensure-ComposerEditor $Hwnd 10
        if (-not $editor.Ok) { return $false }
        & $Activate $Hwnd
        $editor.Prompt.SetFocus()
        Start-Sleep -Milliseconds 200
        $focused = [System.Windows.Automation.AutomationElement]::FocusedElement
        if (-not $focused -or -not $focused.Equals($editor.Prompt)) { return $false }
        $clipboard = [System.Windows.Forms.Clipboard]::GetDataObject()
        $clipboardSaved = $true
        [System.Windows.Forms.Clipboard]::SetText($Text)
        [System.Windows.Forms.SendKeys]::SendWait('^a')
        [System.Windows.Forms.SendKeys]::SendWait('^v')
        Start-Sleep -Milliseconds 350
        return (Test-NarrationValue $Hwnd $Text)
    } catch { return $false }
    finally {
        if ($clipboardSaved) {
            try {
                if ($clipboard) { [System.Windows.Forms.Clipboard]::SetDataObject($clipboard, $true) }
                else { [System.Windows.Forms.Clipboard]::Clear() }
            } catch {}
        }
    }
}
