# PowerShell Environment Setup Script with Theme Support
# Author: Enhanced version with theme selection
# Date: 2025-07-17

param(
    [switch]$Force,
    [string]$Theme = "quick-term"
)

# Check for environment variable override with path extraction
if ($env:POSH_THEME) {
    $envTheme = $env:POSH_THEME

    # If it's a full path, extract just the theme name
    if ($envTheme -match '\\([^\\]+)\.omp\.json$') {
        $Theme = $matches[1]
        Write-Host "🎨 Using theme from environment variable (extracted): $Theme" -ForegroundColor Cyan
    } elseif ($envTheme -match '^[^\\/:]+$') {
        # It's just a theme name (no path separators)
        $Theme = $envTheme
        Write-Host "🎨 Using theme from environment variable: $Theme" -ForegroundColor Cyan
    } else {
        Write-Warning "Invalid theme format in environment variable: $envTheme"
        Write-Info "Expected theme name (e.g., 'quick-term') or valid path"
        Write-Info "Using default theme: quick-term"
        $Theme = "quick-term"
    }
}

# Color functions
function Write-Success { param($Message) Write-Host "✅ $Message" -ForegroundColor Green }
function Write-Warning { param($Message) Write-Host "⚠️ $Message" -ForegroundColor Yellow }
function Write-Info { param($Message) Write-Host "🔧 $Message" -ForegroundColor Cyan }
function Write-Error { param($Message) Write-Host "❌ $Message" -ForegroundColor Red }

# Theme validation function
function Test-OhMyPoshTheme {
    param([string]$ThemeName)

    # List of valid themes from https://ohmyposh.dev/docs/themes
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
        "pixelrobots", "plague", "powerlevel10k_classic", "powerlevel10k_lean", "powerlevel10k_modern",
        "powerlevel10k_rainbow", "powerline", "probua", "pure", "quick-term", "remk", "robbyrussel",
        "rudolfs-dark", "rudolfs-light", "sim-web", "slim", "smoothie", "sonicboom_dark",
        "sonicboom_light", "space", "spaceship", "star", "stelbent", "takuya", "thecyberden",
        "tiwahu", "tokyo", "tokyonight_storm", "unicorn", "velvet", "wopian", "ys", "zash"
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
    } else {
        Write-Warning "Oh My Posh already installed. Use -Force to reinstall."
    }
} catch {
    Write-Error "Failed to install Oh My Posh: $($_.Exception.Message)"
}

# Install MesloLGM Nerd Font with proper error handling
Write-Info "Installing MesloLGM Nerd Font..."
try {
    # Check if font is already installed
    $fontInstalled = $false
    try {
        $fontCheck = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" -ErrorAction SilentlyContinue | 
                     Where-Object { $_.PSChildName -like "*MesloLGM*" -or $_.PSPropertyName -like "*MesloLGM*" }
        if ($fontCheck) {
            $fontInstalled = $true
        }
    } catch {
        # Font registry check failed, assume not installed
    }

    if (-not $fontInstalled -or $Force) {
        Write-Info "Installing font via Oh My Posh..."

        # Install font with all output suppressed
        $processArgs = @{
            FilePath = "oh-my-posh"
            ArgumentList = @("font", "install", "MesloLGM")
            Wait = $true
            NoNewWindow = $true
            RedirectStandardOutput = [System.IO.Path]::GetTempFileName()
            RedirectStandardError = [System.IO.Path]::GetTempFileName()
        }

        $process = Start-Process @processArgs -PassThru

        # Clean up temp files
        if (Test-Path $processArgs.RedirectStandardOutput) { Remove-Item $processArgs.RedirectStandardOutput -Force }
        if (Test-Path $processArgs.RedirectStandardError) { Remove-Item $processArgs.RedirectStandardError -Force }

        if ($process.ExitCode -eq 0) {
            Write-Success "MesloLGM Nerd Font installation completed"
        } else {
            Write-Warning "Font installation may have failed, but continuing..."
        }
    } else {
        Write-Warning "MesloLGM Nerd Font already installed. Use -Force to reinstall."
    }

} catch {
    Write-Error "Failed to install MesloLGM Nerd Font: $($_.Exception.Message)"
}

# Enable Oh My Posh auto-upgrade
Write-Info "Enabling Oh My Posh auto-upgrade..."
try {
    oh-my-posh enable upgrade
    Write-Success "Oh My Posh auto-upgrade enabled successfully"
} catch {
    Write-Warning "Failed to enable Oh My Posh auto-upgrade: $($_.Exception.Message)"
}

# Install fzf
Write-Info "Installing fzf..."
try {
    if ($Force -or -not (Get-Command fzf -ErrorAction SilentlyContinue)) {
        winget install junegunn.fzf -s winget --accept-package-agreements --accept-source-agreements
        Write-Success "fzf installed successfully"
    } else {
        Write-Warning "fzf already installed. Use -Force to reinstall."
    }
} catch {
    Write-Error "Failed to install fzf: $($_.Exception.Message)"
}

# Install ps-fzf module
Write-Info "Installing ps-fzf module..."
try {
    if ($Force -or -not (Get-Module -ListAvailable -Name PSFzf)) {
        Install-Module -Name PSFzf -Force -Scope CurrentUser
        Write-Success "ps-fzf module installed successfully"
    } else {
        Write-Warning "ps-fzf already installed. Use -Force to reinstall."
    }
} catch {
    Write-Error "Failed to install ps-fzf module: $($_.Exception.Message)"
}

# Configure Windows Terminal
Write-Info "Configuring Windows Terminal..."
try {
    $wtSettingsPath = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"

    if (Test-Path $wtSettingsPath) {
        # Read current settings
        $settings = Get-Content $wtSettingsPath -Raw | ConvertFrom-Json

        # Ensure defaults exist
        if (-not $settings.profiles) {
            $settings.profiles = @{}
        }
        if (-not $settings.profiles.defaults) {
            $settings.profiles.defaults = @{}
        }

        # Set font configuration
        $settings.profiles.defaults.font = @{
            face = "MesloLGM Nerd Font Mono"
            size = 12
        }

        # Save settings with proper formatting
        $settings | ConvertTo-Json -Depth 10 | Set-Content $wtSettingsPath -Encoding UTF8
        Write-Success "Windows Terminal configured successfully"
    } else {
        Write-Warning "Windows Terminal settings file not found at expected location"
        Write-Info "You may need to manually set the font in Windows Terminal settings"
    }
} catch {
    Write-Warning "Failed to configure Windows Terminal: $($_.Exception.Message)"
    Write-Info "You may need to manually set the font in Windows Terminal settings"
}

# Configure VS Code - Enhanced version
Write-Info "Configuring VS Code..."
try {
    # Check for VS Code installation
    $vsCodePaths = @(
        "$env:APPDATA\Code\User\settings.json",           # VS Code
        "$env:APPDATA\Code - Insiders\User\settings.json" # VS Code Insiders
    )

    $vsCodeFound = $false

    foreach ($vsCodeSettingsPath in $vsCodePaths) {
        $vsCodeDir = Split-Path $vsCodeSettingsPath -Parent

        # Check if VS Code directory exists (indicates VS Code is installed)
        if (Test-Path $vsCodeDir -PathType Container) {
            $vsCodeFound = $true
            Write-Info "Found VS Code at: $vsCodeDir"

            # Initialize settings object
            $vsCodeSettings = @{}

            # Read existing settings if file exists
            if (Test-Path $vsCodeSettingsPath) {
                try {
                    $existingContent = Get-Content $vsCodeSettingsPath -Raw
                    if ($existingContent.Trim()) {
                        $vsCodeSettings = $existingContent | ConvertFrom-Json -AsHashtable
                    }
                } catch {
                    Write-Warning "Could not parse existing VS Code settings, creating new ones"
                    $vsCodeSettings = @{}
                }
            }

            # Set terminal font settings
            $vsCodeSettings["terminal.integrated.fontFamily"] = "MesloLGM Nerd Font Mono"
            $vsCodeSettings["terminal.integrated.fontSize"] = 12

            # Convert back to JSON and save
            $jsonContent = $vsCodeSettings | ConvertTo-Json -Depth 10
            $jsonContent | Set-Content $vsCodeSettingsPath -Encoding UTF8

            Write-Success "VS Code configured successfully: $(Split-Path $vsCodeSettingsPath -Leaf)"
        }
    }

    if (-not $vsCodeFound) {
        Write-Warning "VS Code not found or not installed"
        Write-Info "If you have VS Code installed, you can manually set the font:"
        Write-Info "  1. Open VS Code"
        Write-Info "  2. Press Ctrl+, to open settings"
        Write-Info "  3. Search for 'terminal.integrated.fontFamily'"
        Write-Info "  4. Set to: MesloLGM Nerd Font Mono"
    }

} catch {
    Write-Warning "Failed to configure VS Code: $($_.Exception.Message)"
    Write-Info "Manual configuration steps:"
    Write-Info "  1. Open VS Code"
    Write-Info "  2. Press Ctrl+, to open settings"
    Write-Info "  3. Search for 'terminal.integrated.fontFamily'"
    Write-Info "  4. Set to: MesloLGM Nerd Font Mono"
}

# Configure PowerShell profile
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

    # Create new profile content with selected theme
    $profileContent = @"
# Oh My Posh initialization with $Theme theme
oh-my-posh init pwsh --config `"`$env:POSH_THEMES_PATH\$Theme.omp.json`" | Invoke-Expression

# PSFzf configuration
Import-Module PSFzf
Set-PsFzfOption -PSReadlineChordProvider 'Ctrl+t' -PSReadlineChordReverseHistory 'Ctrl+r'

# Custom aliases and functions can be added below
"@

    # Write profile
    $profileContent | Set-Content $profilePath -Encoding UTF8
    Write-Success "PowerShell profile configured successfully with $Theme theme"
    Write-Info "Profile location: $profilePath"

} catch {
    Write-Error "Failed to configure PowerShell profile: $($_.Exception.Message)"
}

# Display completion message
Write-Host "`n🎉 Setup completed successfully!" -ForegroundColor Green
Write-Host "🎨 Theme configured: $Theme" -ForegroundColor Magenta

Write-Host "`n📝 Next steps:" -ForegroundColor Yellow
Write-Host "   1. Close all PowerShell/Windows Terminal windows" -ForegroundColor Gray
Write-Host "   2. Open a new Windows Terminal (as regular user, not admin)" -ForegroundColor Gray
Write-Host "   3. Your terminal should now display the $Theme theme with proper icons!" -ForegroundColor Gray

Write-Host "`n🎨 Theme usage:" -ForegroundColor Yellow
Write-Host "   - Current theme: $Theme" -ForegroundColor Gray
Write-Host "   - Browse themes: https://ohmyposh.dev/docs/themes" -ForegroundColor Gray
Write-Host "   - Change theme: Edit your PowerShell profile or re-run script" -ForegroundColor Gray

Write-Host "`n🔧 Theme selection options:" -ForegroundColor Yellow
Write-Host "   - Environment variable: `$env:POSH_THEME = 'atomic'" -ForegroundColor Gray
Write-Host "   - Download script: .\setup.ps1 -Theme 'atomic'" -ForegroundColor Gray
Write-Host "   - Default theme: quick-term" -ForegroundColor Gray
