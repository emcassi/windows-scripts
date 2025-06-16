<#
.SYNOPSIS
Toggles or sets the Windows theme between light and dark modes.

.DESCRIPTION
This script changes the Windows system theme either by:
    - Toggling between light/dark modes automatically
    - Setting a specific mode (light/dark).
    - Applying a custom .theme file.
Allows forcing UI to update (restart explorer).

.PARAMETER Theme
Specify a custom theme file to apply. Including the .theme extension is optional.
This parameter overrides the Mode parameter.

.PARAMETER Mode
Specify whether to switch to light or dark mode.
Valid values: 'light', 'dark'

.PARAMETER Force
Specify whether to force the UI to update by restarting explorer. No value necessary.

.EXAMPLE
PS> .\ChangeTheme.ps1
Toggles between light and dark mode automatically.

.EXAMPLE
PS> .\ChangeTheme.ps1 -Mode dark
Sets the theme to dark mode.

.EXAMPLE
PS> .\ChangeTheme.ps1 -m dark
Sets the theme to dark mode.

.EXAMPLE
PS> .\ChangeTheme.ps1 -Theme "MyCustomTheme"
Applies the specified custom theme.

.EXAMPLE
PS> .\ChangeTheme.ps1 -t "MyCustomTheme"
Applies the specified custom theme.

.EXAMPLE
PS> .\ChangeTheme.ps1 -Mode dark -Force
Sets dark mode and restarts Explorer to ensure changes apply immediately

.EXAMPLE
PS> .\ChangeTheme.ps1 -Help
Displays this help message.

.EXAMPLE
PS> .\ChangeTheme.ps1 -h
Displays this help message.

.EXAMPLE
PS> .\ChangeTheme.ps1 -?
Displays this help message.

.EXAMPLE
PS> .\ChangeTheme.ps1 -Mode light -WhatIf
Shows what would happen without applying the theme.

.NOTES
File Name: ChangeTheme.ps1
Author: Alex Sigalos
Date Created: June 16, 2025
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
    [SupportsWildcards()]
    [Alias('t')]
    [string]$Theme,
    [Alias('m')]
    [ValidateSet("light", "dark")]
    [string]$Mode,
    [switch]$Force,
    [Alias('?', 'h')]
    [switch]$Help

)

if ($Help) {
    Write-Host "=== THEME SWITCHER HELP ===" -ForegroundColor Cyan
    Get-Help $MyInvocation.MyCommand.Path -Full | Out-String | Write-Host
    exit
}

$entryPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\SystemUsesLightTheme"
$currentValue = (New-Object -ComObject WScript.Shell).RegRead($entryPath)
$themesPath = "C:\Windows\Resources\Themes"
$lightTheme = "themeC.theme"
$darkTheme = "themeA.theme"

$themeApplied = ""

if (-not [string]::IsNullOrEmpty($Theme)) {
    if ($Theme.Trim().EndsWith(".theme")) {
        $themeFile = $Theme.Trim()
    }
    else {
        $themeFile = "$($Theme.Trim()).theme"
    }

    # If path isn't fully qualified, assume it's in the default themes directory.
    if (-not [System.IO.Path]::IsPathRooted($themeFile)) {
        $themeFile = Join-Path $themesPath $themeFile
    }

    # Exit if the path to the theme is invalid.
    if (-not (Test-Path $themeFile)) {
        Write-Error "Theme file not found: $themeFile"
        exit 1 
    }

    $themeApplied = $themeFile
}
else {
    if ($Mode -eq "light") {
        $themeFile = "$themesPath\$lightTheme"
        $themeApplied = "light"
    }
    elseif ($Mode -eq "dark") {
        $themeFile = "$themesPath\$darkTheme"
        $themeApplied = "dark"
    }
    elseif ($currentValue -eq 0) {
        $themeFile = "$themesPath\$lightTheme"
        $themeApplied = "light"
    }
    else {
        $themeFile = "$themesPath\$darkTheme"
        $themeApplied = "dark"
    }
}

if ($PSCmdlet.ShouldProcess("$themeFile", "Apply theme")) {
    Start-Process $themeFile
    if (-not [string]::IsNullOrEmpty($Theme)) {
        Write-Host "Successfully applied theme: $themeFile" -ForegroundColor Green
    }
    else {
        Write-Host "Successfully applied $themeApplied theme" -ForegroundColor Green
    }

    if ($Force) {
        if ($PSCmdlet.ShouldProcess("Explorer Process", "Restart")) {
            Write-Verbose "Restarting Explorer process..." -Verbose
            Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
            Start-Process explorer
            Start-Sleep -Seconds 1
            Write-Host "Explorer restarted to apply theme changes" -ForegroundColor Yellow
        }
    }
}

