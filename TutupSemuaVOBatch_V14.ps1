﻿$ErrorActionPreference = 'SilentlyContinue'
$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$StatePath = Join-Path $BaseDir '.V14_BATCH_STATE.json'
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class V14CloseWin32 {
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
}
"@

if (-not (Test-Path -LiteralPath $StatePath)) {
    Write-Host 'Tidak ada batch V14 aktif yang tercatat.'
    Start-Sleep -Seconds 2
    exit 0
}

$state = Get-Content -LiteralPath $StatePath -Raw -Encoding UTF8 | ConvertFrom-Json
$count = 0
foreach ($w in @($state.windows)) {
    [long]$raw = 0
    if ([long]::TryParse([string]$w.Hwnd, [ref]$raw) -and $raw -gt 0) {
        $h = [IntPtr]::new($raw)
        if ([V14CloseWin32]::IsWindow($h)) {
            [void][V14CloseWin32]::PostMessage($h, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
            $count++
        }
    }
}
Remove-Item -LiteralPath $StatePath -Force -ErrorAction SilentlyContinue
Write-Host "Menutup $count window VO batch V14."
Start-Sleep -Seconds 2
