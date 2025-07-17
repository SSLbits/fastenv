# setup-powershell-environment.ps1
# PowerShell Environment Setup Script
# Sets up Oh My Posh, fzf, and ps-fzf for PowerShell 7 with auto-update enabled

param(
    [string]$Theme = "jandedobbeleer",
    [switch]$Force
)

# Ensure running as Administrator for some installations
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

Write-Host "🚀 Setting up PowerShell environment..." -ForegroundColor Green
Write-Host "Administrator privileges: $isAdmin" -ForegroundColor Yellow

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
            # Download and install App Installer from Microsoft Store
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
            # Fallback to manual installation
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

# Enable Oh My Posh Auto-Upgrade
function Enable-OhMyPoshAutoUpgrade {
    Write-Host "🔄 Enabling Oh My Posh auto-upgrade..." -ForegroundColor Yellow

    try {
        # Use the built-in command to enable auto-upgrade
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

# Configure PowerShell Profile (No welcome message)
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
        Write-Host "Backed up existing profile to: $backupPath" -ForegroundColor Yellow
    }

    # Create new profile content (without welcome message)
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
    Write-Host "Profile location: $profilePath" -ForegroundColor Cyan
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
    if (-not (Enable-OhMyPoshAutoUpgrade)) { exit 1 }
    if (-not (Install-Fzf)) { exit 1 }
    if (-not (Install-PsFzf)) { exit 1 }

    # Configure profile
    Configure-Profile

    Write-Host ""
    Write-Host "🎉 Setup completed successfully!" -ForegroundColor Green
    Write-Host "📝 Next steps:" -ForegroundColor Yellow
    Write-Host "   1. Close and reopen your PowerShell terminal" -ForegroundColor White
    Write-Host "   2. Enjoy your new Oh My Posh prompt with auto-update and fzf functionality!" -ForegroundColor White
    Write-Host ""
    Write-Host "🔄 Auto-update features:" -ForegroundColor Yellow
    Write-Host "   - Oh My Posh will automatically check for and install updates" -ForegroundColor White
    Write-Host "   - Enabled via 'oh-my-posh enable upgrade' command" -ForegroundColor White
    Write-Host "   - Use 'Update-OhMyPosh' for manual update checks" -ForegroundColor White
    Write-Host ""
    Write-Host "🔧 Available fzf commands:" -ForegroundColor Yellow
    Write-Host "   - Ctrl+T: File finder" -ForegroundColor White
    Write-Host "   - Ctrl+R: History search" -ForegroundColor White
    Write-Host "   - Tab: Enhanced tab completion" -ForegroundColor White
    Write-Host "   - ff: Fuzzy find and navigate to directory" -ForegroundColor White
    Write-Host "   - fh: Fuzzy search command history" -ForegroundColor White
    Write-Host ""
    Write-Host "🎨 To change Oh My Posh theme later:" -ForegroundColor Yellow
    Write-Host "   oh-my-posh init pwsh --config `$env:POSH_THEMES_PATH\[theme-name].omp.json | Invoke-Expression" -ForegroundColor White
}
catch {
    Write-Error "Setup failed: $_"
    exit 1
}
