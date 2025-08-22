# PowerShell Environment Setup Script with Theme Support  
# Author: Enhanced version with ultra-minimal profile
# Date: 2025-08-22

param(
    [switch]$Force,
    [string]$Theme = "quick-term"
)

# Check for environment variable override with path extraction
if ($env:POSH_THEME) {
    $envTheme = $env:POSH_THEME
    if ($envTheme -match '\\([^\\]+)\.omp\.json$') {
        $Theme = $matches[1]
        Write-Host "🎨 Using theme from environment variable (extracted): $Theme" -ForegroundColor Cyan
    } elseif ($envTheme -match '^[^\\/:\]+$') {
        $Theme = $envTheme
        Write-Host "🎨 Using theme from environment variable: $Theme" -ForegroundColor Cyan
    } else {
        Write-Warning "Invalid theme format in environment variable: $envTheme"
        Write-Info "Using default theme: quick-term"
        $Theme = "quick-term"
    }
}

# Color functions
function Write-Success { param($Message) Write-Host "✅ $Message" -ForegroundColor Green }
function Write-Warning { param($Message) Write-Host "⚠️ $Message" -ForegroundColor Yellow }
function Write-Info { param($Message) Write-Host "🔧 $Message" -ForegroundColor Cyan }
function Write-Error { param($Message) Write-Host "❌ $Message" -ForegroundColor Red }

# Enhanced PATH refresh function
function Update-SessionPath {
    Write-Info "Refreshing environment PATH..."
    $machinePath = [System.Environment]::GetEnvironmentVariable("PATH", "Machine")
    $userPath = [System.Environment]::GetEnvironmentVariable("PATH", "User")
    $env:PATH = "$machinePath;$userPath"
    Write-Info "PATH updated for current session"
}

# Theme validation function  
function Test-OhMyPoshTheme {
    param([string]$ThemeName)
    $validThemes = @(
        "1_shell", "aliens", "amro", "atomic", "avit", "blueish", "bubbles", "bubblesextra", 
        "bubblesline", "capr4n", "catppuccin", "catppuccin_frappe", "catppuccin_latte", 
        "catppuccin_macchiato", "catppuccin_mocha", "cert", "chips", "clean-detailed", 
        "cleanandcolorful", "cloud-context", "cloud-native-azure", "craver", "darkblood", 
        "dracula", "easy-term", "emodipt", "fish", "free-ukraine", "gmay", "gruvbox", 
        "half-life", "hunk", "huvix", "iterm2", "jandedobbeleer", "jblab_2021", "json", 
        "jtracey93", "kali", "kushal", "lambda", "larserikfinholt", "marcduiker", "markbull", 
        "material", "microverse-power", "minimal", "montys", "multiverse-neon", "negligible", 
        "night-owl", "nordtron", "nu4a", "paradox", "pararussel", "patriksvensson", "peru", 
        "pixelrobots", "plague", "powerlevel10k_classic", "powerlevel10k_lean", 
        "powerlevel10k_modern", "powerlevel10k_rainbow", "powerline", "probua", "pure", 
        "quick-term", "remk", "robbyrussel", "rudolfs-dark", "rudolfs-light", "sim-web", 
        "slim", "smoothie", "sonicboom_dark", "sonicboom_light", "space", "spaceship", 
        "star", "stelbent", "takuya", "thecyberden", "tiwahu", "tokyo", "tokyonight_storm", 
        "unicorn", "velvet", "wopian", "ys", "zash"
    )
    return $validThemes -contains $ThemeName
}

# Validate theme
if (-not (Test-OhMyPoshTheme -ThemeName $Theme)) {
    Write-Warning "Invalid theme: $Theme"
    Write-Info "Valid themes can be found at: https://ohmyposh.dev/docs/themes"
    Write-Info "Using default theme: quick-term"
    $Theme = "quick-term"
}

Write-Host "🚀 Setting up PowerShell environment with theme: $Theme" -ForegroundColor Magenta

# Check if running as Administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Admin check with warning
if (Test-Administrator) {
    Write-Warning "Note: This script should NOT be run as Administrator"
    Write-Warning "WARNING: Running as Administrator may cause Windows Terminal to always run as admin!"
    Write-Warning "Consider running this script as a regular user instead."
    $continue = Read-Host "Do you want to continue anyway? (y/N)"
    if ($continue -ne 'y' -and $continue -ne 'Y') {
        Write-Host "Script cancelled. Please run as regular user." -ForegroundColor Yellow
        exit 1
    }
}

# Install Oh My Posh
Write-Info "Installing Oh My Posh..."
try {
    if ($Force -or -not (Get-Command oh-my-posh -ErrorAction SilentlyContinue)) {
        winget install JanDeDobbeleer.OhMyPosh -s winget --accept-package-agreements --accept-source-agreements
        Write-Success "Oh My Posh installed successfully"
        Update-SessionPath
    } else {
        Write-Warning "Oh My Posh already installed. Use -Force to reinstall."
    }
} catch {
    Write-Error "Failed to install Oh My Posh: $($_.Exception.Message)"
}

# Install MesloLGM Nerd Font
Write-Info "Installing MesloLGM Nerd Font..."
try {
    Write-Info "Installing font via Oh My Posh..."
    $fontProcess = Start-Process -FilePath "oh-my-posh" -ArgumentList @("font", "install", "meslo") -Wait -PassThru -NoNewWindow

    if ($fontProcess.ExitCode -eq 0) {
        Write-Success "MesloLGM Nerd Font installation completed"
    } else {
        Write-Warning "Font installation failed with exit code: $($fontProcess.ExitCode)"
        Write-Info "Trying alternative font installation method..."
        try {
            winget install --id=DEVCOM.MesloLGSNerdFont --accept-package-agreements --accept-source-agreements
            Write-Success "Font installed via winget as alternative method"
        } catch {
            Write-Warning "Alternative font installation also failed. You may need to install manually."
            Write-Info "Download from: https://github.com/ryanoasis/nerd-fonts/releases/download/v3.0.2/Meslo.zip"
        }
    }
} catch {
    Write-Error "Failed to install MesloLGM Nerd Font: $($_.Exception.Message)"
    Write-Info "Manual installation: Download from https://github.com/ryanoasis/nerd-fonts/releases/download/v3.0.2/Meslo.zip"
}

# Enable Oh My Posh auto-upgrade
Write-Info "Enabling Oh My Posh auto-upgrade..."
try {
    oh-my-posh upgrade
    Write-Success "Oh My Posh auto-upgrade completed successfully"
} catch {
    Write-Warning "Failed to enable Oh My Posh auto-upgrade: $($_.Exception.Message)"
}

# Install fzf
Write-Info "Installing fzf..."
try {
    if ($Force -or -not (Get-Command fzf -ErrorAction SilentlyContinue)) {
        winget install junegunn.fzf -s winget --accept-package-agreements --accept-source-agreements
        Write-Success "fzf installed successfully"
        Update-SessionPath
    } else {
        Write-Warning "fzf already installed. Use -Force to reinstall."
    }
} catch {
    Write-Error "Failed to install fzf: $($_.Exception.Message)"
}

# Install PSFzf module
Write-Info "Installing PSFzf module..."
try {
    if ($Force -or -not (Get-Module -ListAvailable -Name PSFzf)) {
        Install-Module -Name PSFzf -Force -Scope CurrentUser -AllowClobber
        Write-Success "PSFzf module installed successfully"
    } else {
        Write-Warning "PSFzf already installed. Use -Force to reinstall."
    }
} catch {
    Write-Error "Failed to install PSFzf module: $($_.Exception.Message)"
}

# Configure Windows Terminal
Write-Info "Configuring Windows Terminal..."
try {
    $wtSettingsPath = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"
    if (Test-Path $wtSettingsPath) {
        $settingsContent = Get-Content $wtSettingsPath -Raw
        $settings = $settingsContent | ConvertFrom-Json

        # Ensure correct JSON structure
        if (-not $settings.profiles) { 
            $settings | Add-Member -NotePropertyName "profiles" -NotePropertyValue @{} -Force
        }
        if (-not $settings.profiles.defaults) { 
            $settings.profiles | Add-Member -NotePropertyName "defaults" -NotePropertyValue @{} -Force
        }

        # Configure font
        $settings.profiles.defaults | Add-Member -NotePropertyName "font" -NotePropertyValue @{
            "face" = "MesloLGM Nerd Font Mono"
            "size" = 12
        } -Force

        # Save settings
        $jsonOutput = $settings | ConvertTo-Json -Depth 10
        $jsonOutput | Set-Content $wtSettingsPath -Encoding UTF8
        Write-Success "Windows Terminal configured successfully"
    } else {
        Write-Warning "Windows Terminal settings file not found"
    }
} catch {
    Write-Warning "Failed to configure Windows Terminal: $($_.Exception.Message)"
}

# Configure VS Code and Cursor AI
Write-Info "Configuring VS Code and Cursor AI..."
try {
    $editorPaths = @(
        @{ Path = "$env:APPDATA\Code\User\settings.json"; Name = "VS Code" },
        @{ Path = "$env:APPDATA\Cursor\User\settings.json"; Name = "Cursor AI" }
    )

    foreach ($editorInfo in $editorPaths) {
        $editorPath = $editorInfo.Path
        $editorName = $editorInfo.Name
        $editorDir = Split-Path $editorPath -Parent

        if (Test-Path $editorDir -PathType Container) {
            Write-Info "Configuring $editorName"

            $editorSettings = @{}
            if (Test-Path $editorPath) {
                try {
                    $existingContent = Get-Content $editorPath -Raw
                    if ($existingContent.Trim()) {
                        $editorSettings = $existingContent | ConvertFrom-Json -AsHashtable
                    }
                } catch {
                    Write-Warning "Could not parse existing $editorName settings"
                    $editorSettings = @{}
                }
            } else {
                New-Item -ItemType Directory -Path $editorDir -Force | Out-Null
            }

            # Configure terminal font
            $editorSettings["terminal.integrated.fontFamily"] = "MesloLGM Nerd Font Mono"
            $editorSettings["terminal.integrated.fontSize"] = 12

            # Save settings
            $jsonContent = $editorSettings | ConvertTo-Json -Depth 10
            $jsonContent | Set-Content $editorPath -Encoding UTF8
            Write-Success "$editorName configured successfully"
        }
    }
} catch {
    Write-Warning "Failed to configure editors: $($_.Exception.Message)"
}

# **ULTRA-MINIMAL: Configure PowerShell profile (silent & fast)**
Write-Info "Configuring PowerShell profile..."
try {
    $profilePath = $PROFILE
    $profileDir = Split-Path $profilePath -Parent

    # Create profile directory if it doesn't exist
    if (-not (Test-Path $profileDir)) {
        New-Item -ItemType Directory -Path $profileDir -Force | Out-Null
    }

    # Backup existing profile
    if (Test-Path $profilePath) {
        $backupPath = "$profilePath.backup.$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        Copy-Item $profilePath $backupPath
        Write-Info "Backed up existing profile to: $backupPath"
    }

    # **ULTRA-MINIMAL: Silent profile content (no messages, no try-catch)**
    $profileContent = @"
# Oh My Posh initialization
oh-my-posh init pwsh --config "`$env:POSH_THEMES_PATH\$Theme.omp.json" | Invoke-Expression

# PSFzf configuration
Import-Module PSFzf
Remove-PSReadLineKeyHandler -Key 'Ctrl+r'
Remove-PSReadLineKeyHandler -Key 'Ctrl+t'
Set-PsFzfOption -PSReadlineChordProvider 'Ctrl+t' -PSReadlineChordReverseHistory 'Ctrl+r'
Set-PsFzfOption -TabExpansion
"@

    # Write ultra-minimal profile
    $profileContent | Set-Content $profilePath -Encoding UTF8
    Write-Success "PowerShell profile configured successfully with $Theme theme"
    Write-Info "Profile location: $profilePath"
    Write-Info "⚡ Ultra-minimal profile for fastest loading"
} catch {
    Write-Error "Failed to configure PowerShell profile: $($_.Exception.Message)"
}

# Display completion message
Write-Host "`n🎉 Setup completed successfully!" -ForegroundColor Green
Write-Host "🎨 Theme configured: $Theme" -ForegroundColor Magenta

Write-Host "`n📝 Next steps:" -ForegroundColor Yellow
Write-Host "   1. Close ALL PowerShell/Windows Terminal windows completely" -ForegroundColor Gray
Write-Host "   2. Wait 5 seconds" -ForegroundColor Gray  
Write-Host "   3. Open a NEW Windows Terminal" -ForegroundColor Gray
Write-Host "   4. Test fzf functionality:" -ForegroundColor Gray
Write-Host "      • Press Ctrl+T for fuzzy file search" -ForegroundColor Gray
Write-Host "      • Press Ctrl+R for fuzzy history search" -ForegroundColor Gray

Write-Host "`n⚡ Ultra-minimal profile optimizations:" -ForegroundColor Yellow
Write-Host "   • No startup messages" -ForegroundColor Gray
Write-Host "   • No try-catch error handling (faster loading)" -ForegroundColor Gray
Write-Host "   • Essential Oh My Posh + fzf configuration only" -ForegroundColor Gray
Write-Host "   • Should reduce profile loading time significantly" -ForegroundColor Gray

Write-Host "`n🎨 Theme selection options:" -ForegroundColor Yellow
Write-Host "   - Environment variable: `$env:POSH_THEME = 'atomic'" -ForegroundColor Gray
Write-Host "   - Download script: .\setup.ps1 -Theme 'atomic'" -ForegroundColor Gray
Write-Host "   - Default theme: quick-term" -ForegroundColor Gray
