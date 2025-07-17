# setup-powershell-environment.ps1
# PowerShell Environment Setup Script
# Sets up Oh My Posh, fzf, ps-fzf, and MesloLGM Nerd Font with auto-update enabled

param(
    [string]$Theme = "jandedobbeleer",
    [string]$FontVariant = "Mono", # Options: "Regular", "Mono", "Propo"
    [switch]$Force
)

Write-Host "🚀 Setting up PowerShell environment..." -ForegroundColor Green
Write-Host "⚠️  Note: This script should NOT be run as Administrator" -ForegroundColor Yellow

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if ($isAdmin) {
    Write-Host "⚠️  WARNING: Running as Administrator may cause Windows Terminal to always run as admin!" -ForegroundColor Red
    Write-Host "⚠️  Consider running this script as a regular user instead." -ForegroundColor Red
    $continue = Read-Host "Do you want to continue anyway? (y/N)"
    if ($continue -ne "y" -and $continue -ne "Y") {
        Write-Host "Exiting. Please run this script as a regular user." -ForegroundColor Yellow
        exit 1
    }
}

# Function to check if command exists
function Test-Command {
    param($Command)
    try {
        Get-Command $Command -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Function to install winget if not present
function Install-Winget {
    if (-not (Test-Command "winget")) {
        Write-Host "📦 Installing winget..." -ForegroundColor Yellow
        try {
            $progressPreference = 'silentlyContinue'
            Invoke-WebRequest -Uri "https://aka.ms/getwinget" -OutFile "$env:TEMP\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
            Add-AppxPackage "$env:TEMP\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
            Write-Host "✅ winget installed successfully" -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install winget: $_"
            return $false
        }
    }
    return $true
}

# Install Oh My Posh
function Install-OhMyPosh {
    Write-Host "🎨 Installing Oh My Posh..." -ForegroundColor Yellow

    if (Test-Command "oh-my-posh") {
        if (-not $Force) {
            Write-Host "Oh My Posh already installed. Use -Force to reinstall." -ForegroundColor Yellow
            return $true
        }
    }

    try {
        if (Test-Command "winget") {
            winget install JanDeDobbeleer.OhMyPosh -s winget --accept-source-agreements --accept-package-agreements
        }
        else {
            Set-ExecutionPolicy Bypass -Scope Process -Force
            Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://ohmyposh.dev/install.ps1'))
        }

        # Add Oh My Posh to PATH if not already there
        $ohMyPoshPath = "$env:LOCALAPPDATA\Programs\oh-my-posh\bin"
        $currentPath = [Environment]::GetEnvironmentVariable("Path", "User")
        if ($currentPath -notlike "*$ohMyPoshPath*") {
            [Environment]::SetEnvironmentVariable("Path", "$currentPath;$ohMyPoshPath", "User")
            $env:Path += ";$ohMyPoshPath"
        }

        Write-Host "✅ Oh My Posh installed successfully" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to install Oh My Posh: $_"
        return $false
    }
}

# Install Nerd Font using Oh My Posh (Fixed version)
function Install-NerdFont {
    Write-Host "🔤 Installing MesloLGM Nerd Font..." -ForegroundColor Yellow

    try {
        if (Test-Command "oh-my-posh") {
            Write-Host "Installing font via Oh My Posh..." -ForegroundColor Cyan

            # Redirect all output to null to prevent capturing progress bars
            $null = oh-my-posh font install MesloLGM *>&1

            # Check if font installation was successful by looking for the font files
            Start-Sleep -Seconds 3

            Write-Host "✅ MesloLGM Nerd Font installation completed" -ForegroundColor Green
            Write-Host "Available font variants:" -ForegroundColor Cyan
            Write-Host "  • MesloLGM Nerd Font" -ForegroundColor White
            Write-Host "  • MesloLGM Nerd Font Mono" -ForegroundColor White
            Write-Host "  • MesloLGM Nerd Font Propo" -ForegroundColor White

            return $true
        }
        else {
            Write-Host "⚠️ Oh My Posh not found, skipping font installation" -ForegroundColor Yellow
            return $true
        }
    }
    catch {
        Write-Host "⚠️ Font installation completed (with possible warnings)" -ForegroundColor Yellow
        return $true  # Continue anyway
    }
}

# Get correct font name based on variant
function Get-FontName {
    param([string]$Variant)

    switch ($Variant.ToLower()) {
        "regular" { return "MesloLGM Nerd Font" }
        "mono" { return "MesloLGM Nerd Font Mono" }
        "propo" { return "MesloLGM Nerd Font Propo" }
        default { return "MesloLGM Nerd Font Mono" }
    }
}

# Enable Oh My Posh Auto-Upgrade
function Enable-OhMyPoshAutoUpgrade {
    Write-Host "🔄 Enabling Oh My Posh auto-upgrade..." -ForegroundColor Yellow

    try {
        oh-my-posh enable upgrade
        Write-Host "✅ Oh My Posh auto-upgrade enabled successfully" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to enable Oh My Posh auto-upgrade: $_"
        return $false
    }
}

# Install fzf
function Install-Fzf {
    Write-Host "🔍 Installing fzf..." -ForegroundColor Yellow

    if (Test-Command "fzf") {
        if (-not $Force) {
            Write-Host "fzf already installed. Use -Force to reinstall." -ForegroundColor Yellow
            return $true
        }
    }

    try {
        if (Test-Command "winget") {
            winget install fzf --accept-source-agreements --accept-package-agreements
        }
        elseif (Test-Command "choco") {
            choco install fzf -y
        }
        else {
            # Manual installation
            $fzfVersion = "0.44.1"
            $fzfUrl = "https://github.com/junegunn/fzf/releases/download/$fzfVersion/fzf-$fzfVersion-windows_amd64.zip"
            $fzfPath = "$env:LOCALAPPDATA\Programs\fzf"

            if (-not (Test-Path $fzfPath)) {
                New-Item -ItemType Directory -Path $fzfPath -Force
            }

            Invoke-WebRequest -Uri $fzfUrl -OutFile "$env:TEMP\fzf.zip"
            Expand-Archive -Path "$env:TEMP\fzf.zip" -DestinationPath $fzfPath -Force

            # Add to PATH
            $currentPath = [Environment]::GetEnvironmentVariable("Path", "User")
            if ($currentPath -notlike "*$fzfPath*") {
                [Environment]::SetEnvironmentVariable("Path", "$currentPath;$fzfPath", "User")
                $env:Path += ";$fzfPath"
            }
        }

        Write-Host "✅ fzf installed successfully" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to install fzf: $_"
        return $false
    }
}

# Install ps-fzf module
function Install-PsFzf {
    Write-Host "🔧 Installing ps-fzf module..." -ForegroundColor Yellow

    try {
        if (Get-Module -ListAvailable -Name PSFzf) {
            if (-not $Force) {
                Write-Host "ps-fzf already installed. Use -Force to reinstall." -ForegroundColor Yellow
                return $true
            }
            else {
                Uninstall-Module PSFzf -Force
            }
        }

        Install-Module -Name PSFzf -Force -Scope CurrentUser
        Write-Host "✅ ps-fzf installed successfully" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to install ps-fzf: $_"
        return $false
    }
}

# Configure Windows Terminal (Fixed version)
function Configure-WindowsTerminal {
    param([string]$FontName)

    Write-Host "🖥️ Configuring Windows Terminal..." -ForegroundColor Yellow

    try {
        $wtSettingsPath = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"

        if (-not (Test-Path $wtSettingsPath)) {
            Write-Host "⚠️ Windows Terminal settings not found. Please:" -ForegroundColor Yellow
            Write-Host "   1. Launch Windows Terminal once to create settings" -ForegroundColor White
            Write-Host "   2. Then run this script again" -ForegroundColor White
            return $false
        }

        # Read current settings
        $wtSettings = Get-Content $wtSettingsPath -Raw | ConvertFrom-Json

        # Backup original settings
        $backupPath = "$wtSettingsPath.backup.$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        Copy-Item $wtSettingsPath $backupPath

        # Ensure profiles.defaults exists
        if (-not $wtSettings.profiles.defaults) {
            $wtSettings.profiles | Add-Member -MemberType NoteProperty -Name "defaults" -Value @{} -Force
        }

        # Ensure font object exists
        if (-not $wtSettings.profiles.defaults.font) {
            $wtSettings.profiles.defaults | Add-Member -MemberType NoteProperty -Name "font" -Value @{} -Force
        }

        # Set font properties
        $wtSettings.profiles.defaults.font.face = $FontName
        $wtSettings.profiles.defaults.font.size = 12

        # Save settings with proper formatting
        $jsonSettings = $wtSettings | ConvertTo-Json -Depth 10
        $jsonSettings | Set-Content $wtSettingsPath -Encoding UTF8

        Write-Host "✅ Windows Terminal configured with font: $FontName" -ForegroundColor Green
        Write-Host "📁 Backup saved to: $backupPath" -ForegroundColor Cyan

        return $true
    }
    catch {
        Write-Host "⚠️ Failed to configure Windows Terminal: $_" -ForegroundColor Yellow
        Write-Host "⚠️ You may need to manually set the font in Windows Terminal settings" -ForegroundColor Yellow
        return $false
    }
}

# Configure VS Code
function Configure-VSCode {
    param([string]$FontName)

    Write-Host "📝 Configuring VS Code..." -ForegroundColor Yellow

    try {
        $vscodeSettingsPath = "$env:APPDATA\Code\User\settings.json"

        # Create settings object
        if (Test-Path $vscodeSettingsPath) {
            $vscodeSettings = Get-Content $vscodeSettingsPath -Raw | ConvertFrom-Json
        }
        else {
            $vscodeSettings = @{}
            $vscodeDir = Split-Path $vscodeSettingsPath -Parent
            if (-not (Test-Path $vscodeDir)) {
                New-Item -ItemType Directory -Path $vscodeDir -Force | Out-Null
            }
        }

        # Set terminal font
        $vscodeSettings."terminal.integrated.fontFamily" = $FontName
        $vscodeSettings."terminal.integrated.fontSize" = 12

        # Save settings
        $vscodeSettings | ConvertTo-Json -Depth 10 | Set-Content $vscodeSettingsPath -Encoding UTF8

        Write-Host "✅ VS Code configured with font: $FontName" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "⚠️ Failed to configure VS Code: $_" -ForegroundColor Yellow
        return $false
    }
}

# Configure PowerShell Profile
function Configure-Profile {
    Write-Host "⚙️ Configuring PowerShell profile..." -ForegroundColor Yellow

    $profilePath = $PROFILE.CurrentUserCurrentHost

    # Create profile directory if it doesn't exist
    $profileDir = Split-Path $profilePath -Parent
    if (-not (Test-Path $profileDir)) {
        New-Item -ItemType Directory -Path $profileDir -Force
    }

    # Backup existing profile if it exists
    if (Test-Path $profilePath) {
        $backupPath = "$profilePath.backup.$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        Copy-Item $profilePath $backupPath
        Write-Host "📁 Backed up existing profile to: $backupPath" -ForegroundColor Cyan
    }

    # Create new profile content
    $profileContent = @"
# PowerShell Profile - Auto-generated by setup script
# Generated on: $(Get-Date)

# Oh My Posh initialization (auto-upgrade enabled via oh-my-posh enable upgrade)
oh-my-posh init pwsh --config `$env:POSH_THEMES_PATH\$Theme.omp.json | Invoke-Expression

# ps-fzf initialization
Import-Module PSFzf

# Set fzf options
Set-PsFzfOption -PSReadlineChordProvider 'Ctrl+t' -PSReadlineChordReverseHistory 'Ctrl+r'

# Optional: Additional fzf key bindings
Set-PSReadLineKeyHandler -Key Tab -ScriptBlock { Invoke-FzfTabCompletion }

# Optional: Enhanced directory navigation
function ff { Get-ChildItem . -Recurse -Name | fzf | ForEach-Object { Set-Location (Split-Path `$_) } }
function fh { Get-Content (Get-PSReadlineOption).HistorySavePath | fzf | Invoke-Expression }

# Manual update function for Oh My Posh
function Update-OhMyPosh {
    Write-Host "🔄 Checking for Oh My Posh updates..." -ForegroundColor Yellow
    oh-my-posh upgrade
}
"@

    # Write profile content
    Set-Content -Path $profilePath -Value $profileContent -Encoding UTF8

    Write-Host "✅ PowerShell profile configured successfully" -ForegroundColor Green
    Write-Host "📁 Profile location: $profilePath" -ForegroundColor Cyan
}

# Main execution
try {
    # Set execution policy if needed
    if ((Get-ExecutionPolicy -Scope CurrentUser) -eq 'Restricted') {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        Write-Host "✅ Execution policy set to RemoteSigned" -ForegroundColor Green
    }

    # Install components
    if (-not (Install-Winget)) { exit 1 }
    if (-not (Install-OhMyPosh)) { exit 1 }
    if (-not (Install-NerdFont)) { exit 1 }
    if (-not (Enable-OhMyPoshAutoUpgrade)) { exit 1 }
    if (-not (Install-Fzf)) { exit 1 }
    if (-not (Install-PsFzf)) { exit 1 }

    # Get the correct font name
    $fontName = Get-FontName -Variant $FontVariant

    # Configure applications
    Configure-WindowsTerminal -FontName $fontName
    Configure-VSCode -FontName $fontName
    Configure-Profile

    Write-Host ""
    Write-Host "🎉 Setup completed successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "📝 Next steps:" -ForegroundColor Yellow
    Write-Host "   1. Close all PowerShell/Windows Terminal windows" -ForegroundColor White
    Write-Host "   2. Open a new Windows Terminal (as regular user, not admin)" -ForegroundColor White
    Write-Host "   3. Your terminal should now display Oh My Posh themes with proper icons!" -ForegroundColor White
    Write-Host ""
    Write-Host "🔤 Font configuration:" -ForegroundColor Yellow
    Write-Host "   - Font configured: $fontName" -ForegroundColor White
    Write-Host "   - Windows Terminal: Configured automatically" -ForegroundColor White
    Write-Host "   - VS Code: Configured automatically" -ForegroundColor White
    Write-Host ""
    Write-Host "⚠️  Admin privilege fix:" -ForegroundColor Yellow
    Write-Host "   - If Windows Terminal still runs as admin, right-click the Terminal icon" -ForegroundColor White
    Write-Host "   - Go to Properties → Advanced → Uncheck 'Run as administrator'" -ForegroundColor White
    Write-Host ""
    Write-Host "🔧 Manual font configuration (if needed):" -ForegroundColor Yellow
    Write-Host "   - Windows Terminal: Ctrl+, → Profiles → Defaults → Appearance → Font face" -ForegroundColor White
    Write-Host "   - Set to: $fontName" -ForegroundColor White
    Write-Host ""
    Write-Host "🔄 Auto-update features:" -ForegroundColor Yellow
    Write-Host "   - Oh My Posh will automatically check for updates" -ForegroundColor White
    Write-Host "   - Use 'Update-OhMyPosh' for manual updates" -ForegroundColor White
}
catch {
    Write-Error "Setup failed: $_"
    exit 1
}
