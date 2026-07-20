#!/usr/bin/env pwsh
<#
.SYNOPSIS
Takes a screenshot of the desktop, minimizing PowerShell windows first.
Useful for capturing Excel state when macros hang or fail.

.PARAMETER OutputPath
The file path where the screenshot should be saved.

.EXAMPLE
./scripts/take-screenshot-on-failure.ps1 -OutputPath ./screenshots/failure.png
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$OutputPath
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Win32 API for window manipulation
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindowVisible(IntPtr hWnd);
}
"@

# Minimize all PowerShell windows to see Excel clearly
Write-Host "Minimizing PowerShell windows..." -ForegroundColor Cyan
$processes = Get-Process | Where-Object { $_.MainWindowHandle -ne 0 }
$minimized = 0

foreach ($proc in $processes) {
    if ($proc.ProcessName -like "*pwsh*" -or $proc.ProcessName -like "*powershell*") {
        try {
            if ([Win32]::IsWindowVisible($proc.MainWindowHandle)) {
                [Win32]::ShowWindow($proc.MainWindowHandle, 6) # SW_MINIMIZE = 6
                $minimized++
                Write-Host "Minimized: $($proc.ProcessName)" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "Could not minimize $($proc.ProcessName): $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

Write-Host "Minimized $minimized window(s). Waiting 1 second..." -ForegroundColor Cyan
Start-Sleep -Seconds 1

# Take screenshot
$bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$screenshot = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
$graphics = [System.Drawing.Graphics]::FromImage($screenshot)
$graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)

# Ensure output directory exists
$outputDir = Split-Path -Parent $OutputPath
if (!(Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$screenshot.Save($OutputPath)
$graphics.Dispose()
$screenshot.Dispose()

Write-Host "Screenshot saved to: $OutputPath" -ForegroundColor Green

