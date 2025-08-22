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
    } elseif ($envTheme -match '^[^\\/:\]+$') {
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

        # Refresh PATH for current session
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH","User")
        Write-Info "Environment PATH refreshed"

    } else {
        Write-Warning "Oh My Posh already installed. Use -Force to reinstall."
    }
} catch {
    Write-Error "Failed to install Oh My Posh: $($_.Exception.Message)"
}

# **UPDATED: Install MesloLGM Nerd Font with better error handling and verification**
Write-Info "Installing MesloLGM Nerd Font..."
try {
    Write-Info "Installing font via Oh My Posh..."
    # Use the correct font name that oh-my-posh expects
    $fontProcess = Start-Process -FilePath "oh-my-posh" -ArgumentList @("font", "install", "meslo") -Wait -PassThru -NoNewWindow

    if ($fontProcess.ExitCode -eq 0) {
        Write-Success "MesloLGM Nerd Font installation completed"

        # **ADDED: Verify font installation**
        Start-Sleep -Seconds 2
        $fontCheck = Get-ChildItem -Path "$env:WINDIR\Fonts" -Filter "*Meslo*" -ErrorAction SilentlyContinue
        if ($fontCheck) {
            Write-Success "Font files found in system fonts directory"
        } else {
            Write-Warning "Font files not found in expected location, but installation reported success"
        }
    } else {
        Write-Warning "Font installation process returned exit code: $($fontProcess.ExitCode)"
        Write-Info "Trying alternative font installation method..."

        # **ADDED: Alternative font installation using winget**
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

        # Refresh PATH again after fzf installation
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH","User")
        Write-Info "Environment PATH refreshed after fzf installation"

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

# **UPDATED: Configure Windows Terminal with correct JSON structure**
Write-Info "Configuring Windows Terminal..."
try {
    $wtSettingsPath = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"
    if (Test-Path $wtSettingsPath) {
        # Read current settings
        $settingsContent = Get-Content $wtSettingsPath -Raw
        $settings = $settingsContent | ConvertFrom-Json

        # **FIXED: Ensure correct JSON structure for modern Windows Terminal**
        if (-not $settings.profiles) { 
            $settings | Add-Member -NotePropertyName "profiles" -NotePropertyValue @{} -Force
        }
        if (-not $settings.profiles.defaults) { 
            $settings.profiles | Add-Member -NotePropertyName "defaults" -NotePropertyValue @{} -Force
        }

        # **FIXED: Use correct property structure for font**
        $settings.profiles.defaults | Add-Member -NotePropertyName "font" -NotePropertyValue @{
            "face" = "MesloLGM Nerd Font Mono"
            "size" = 12
        } -Force

        # Save settings with proper formatting
        $jsonOutput = $settings | ConvertTo-Json -Depth 10
        $jsonOutput | Set-Content $wtSettingsPath -Encoding UTF8
        Write-Success "Windows Terminal configured successfully"
    } else {
        Write-Warning "Windows Terminal settings file not found at expected location"
        Write-Info "Windows Terminal may not be installed or may be using a different settings location"
        Write-Info "You can manually set the font in Windows Terminal: Settings > Profiles > Defaults > Appearance > Font face"
    }
} catch {
    Write-Warning "Failed to configure Windows Terminal: $($_.Exception.Message)"
    Write-Info "Manual setup: Settings > Profiles > Defaults > Appearance > Font face > MesloLGM Nerd Font Mono"
}

# **UPDATED: Enhanced VS Code detection with multiple installation paths**
Write-Info "Configuring VS Code..."
try {
    # **EXPANDED: Check multiple possible VS Code installation locations**
    $vsCodePaths = @(
        @{ Path = "$env:APPDATA\Code\User\settings.json"; Name = "VS Code (User Install)" },
        @{ Path = "$env:APPDATA\Code - Insiders\User\settings.json"; Name = "VS Code Insiders (User)" },
        @{ Path = "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Visual Studio Code"; Name = "VS Code (System Install)" }
    )

    # **ADDED: Also check if VS Code is in PATH**
    $vsCodeInPath = Get-Command "code" -ErrorAction SilentlyContinue
    if ($vsCodeInPath) {
        Write-Info "VS Code executable found in PATH: $($vsCodeInPath.Source)"
    }

    $vsCodeFound = $false

    foreach ($vsCodeInfo in $vsCodePaths) {
        $vsCodePath = $vsCodeInfo.Path
        $vsCodeName = $vsCodeInfo.Name
        $vsCodeDir = Split-Path $vsCodePath -Parent

        # **IMPROVED: Better detection logic**
        $pathExists = $false
        if ($vsCodePath.EndsWith("settings.json")) {
            # For settings.json paths, check if the directory exists
            $pathExists = Test-Path $vsCodeDir -PathType Container
        } else {
            # For other paths, check the path directly
            $pathExists = Test-Path $vsCodePath
        }

        if ($pathExists -or $vsCodeInPath) {
            $vsCodeFound = $true
            Write-Info "Found $vsCodeName"

            # Only process settings.json paths
            if ($vsCodePath.EndsWith("settings.json")) {
                # Initialize settings object
                $vsCodeSettings = @{}

                # Read existing settings if file exists
                if (Test-Path $vsCodePath) {
                    try {
                        $existingContent = Get-Content $vsCodePath -Raw
                        if ($existingContent.Trim()) {
                            $vsCodeSettings = $existingContent | ConvertFrom-Json -AsHashtable
                        }
                    } catch {
                        Write-Warning "Could not parse existing VS Code settings, creating new ones"
                        $vsCodeSettings = @{}
                    }
                } else {
                    # Create directory if it doesn't exist
                    if (-not (Test-Path $vsCodeDir)) {
                        New-Item -ItemType Directory -Path $vsCodeDir -Force | Out-Null
                        Write-Info "Created VS Code settings directory: $vsCodeDir"
                    }
                }

                # Set terminal font settings
                $vsCodeSettings["terminal.integrated.fontFamily"] = "MesloLGM Nerd Font Mono"
                $vsCodeSettings["terminal.integrated.fontSize"] = 12

                # Convert back to JSON and save
                $jsonContent = $vsCodeSettings | ConvertTo-Json -Depth 10
                $jsonContent | Set-Content $vsCodePath -Encoding UTF8
                Write-Success "VS Code configured successfully: $vsCodeName"
            }
        }
    }

    if (-not $vsCodeFound) {
        Write-Warning "VS Code not found in standard locations"
        Write-Info "Checked locations:"
        foreach ($location in $vsCodePaths) {
            Write-Info "  - $($location.Path)"
        }
        Write-Info "Manual configuration steps:"
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
oh-my-posh init pwsh --config "`$env:POSH_THEMES_PATH\$Theme.omp.json" | Invoke-Expression

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
