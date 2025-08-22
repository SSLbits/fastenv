# PowerShell Environment Setup Script with Theme Support
# Author: Enhanced version with theme selection, Cursor AI support, and experimental features
# Date: 2025-08-22

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

# **BULLETPROOF: Windows Terminal configuration with reliable JSON handling**
Write-Info "Configuring Windows Terminal..."
try {
    $wtSettingsPath = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"
    if (Test-Path $wtSettingsPath) {
        # Read and parse existing settings
        $settingsContent = Get-Content $wtSettingsPath -Raw
        $existingSettings = $settingsContent | ConvertFrom-Json

        # Create new settings hashtable
        $newSettings = @{}

        # Preserve all existing settings except the ones we want to override
        $existingSettings.PSObject.Properties | ForEach-Object {
            $newSettings[$_.Name] = $_.Value
        }

        # **FORCE SET: Experimental features at root level**
        $newSettings['experimental.repositionCursorWithMouse'] = $true
        $newSettings['experimental.detectURLs'] = $true

        # **ENSURE: Profiles structure exists**
        if (-not $newSettings.profiles) {
            $newSettings.profiles = @{}
        }
        if (-not $newSettings.profiles.defaults) {
            $newSettings.profiles.defaults = @{}
        }

        # **SET: Font configuration**
        $newSettings.profiles.defaults.font = @{
            "face" = "MesloLGM Nerd Font Mono"
            "size" = 12
        }

        # **OPTIONAL: Shell integration (if supported)**
        try {
            if (-not $newSettings.profiles.defaults.shellIntegration) {
                $newSettings.profiles.defaults.shellIntegration = @{
                    "showMarksOnScrollbar" = $true
                }
            }
        } catch {
            Write-Info "Shell integration not supported in this Windows Terminal version (skipped)"
        }

        # Write back to file with proper formatting
        $jsonOutput = $newSettings | ConvertTo-Json -Depth 10
        $jsonOutput | Set-Content $wtSettingsPath -Encoding UTF8

        Write-Success "Windows Terminal configured successfully"
        Write-Info "✓ Cursor repositioning: ENABLED"
        Write-Info "✓ URL detection: ENABLED" 
        Write-Info "✓ Font: MesloLGM Nerd Font Mono"

        # Verify the settings were written correctly
        Write-Info "Verifying configuration..."
        $verifyContent = Get-Content $wtSettingsPath -Raw | ConvertFrom-Json
        if ($verifyContent.'experimental.repositionCursorWithMouse' -eq $true) {
            Write-Success "✓ Cursor repositioning confirmed in settings file"
        } else {
            Write-Warning "⚠ Cursor repositioning may not have been set correctly"
        }

    } else {
        Write-Warning "Windows Terminal settings file not found"
        Write-Info "Manual setup required: Settings → Interaction → Reposition cursor with mouse"
    }
} catch {
    Write-Warning "Configuration failed: $($_.Exception.Message)"
    Write-Info "Manual steps:"
    Write-Info "  1. Open Windows Terminal"
    Write-Info "  2. Press Ctrl+, (Settings)"
    Write-Info "  3. Go to Interaction tab"
    Write-Info "  4. Toggle ON 'Reposition cursor with mouse'"
}

# Configure VS Code and Cursor AI
Write-Info "Configuring VS Code and Cursor AI..."
try {
    # Check multiple possible installation locations including Cursor AI
    $editorPaths = @(
        @{ Path = "$env:APPDATA\Code\User\settings.json"; Name = "VS Code (User Install)" },
        @{ Path = "$env:APPDATA\Code - Insiders\User\settings.json"; Name = "VS Code Insiders (User)" },
        @{ Path = "$env:APPDATA\Cursor\User\settings.json"; Name = "Cursor AI" },
        @{ Path = "$env:PROGRAMDATA\Microsoft\Windows\Start Menu\Programs\Visual Studio Code"; Name = "VS Code (System Install)" }
    )

    # Check if executables are in PATH
    $vsCodeInPath = Get-Command "code" -ErrorAction SilentlyContinue
    $cursorInPath = Get-Command "cursor" -ErrorAction SilentlyContinue

    if ($vsCodeInPath) {
        Write-Info "VS Code executable found in PATH: $($vsCodeInPath.Source)"
    }
    if ($cursorInPath) {
        Write-Info "Cursor AI executable found in PATH: $($cursorInPath.Source)"
    }

    $editorsFound = $false

    foreach ($editorInfo in $editorPaths) {
        $editorPath = $editorInfo.Path
        $editorName = $editorInfo.Name
        $editorDir = Split-Path $editorPath -Parent

        # Better detection logic
        $pathExists = $false
        if ($editorPath.EndsWith("settings.json")) {
            # For settings.json paths, check if the directory exists OR if corresponding executable exists
            $pathExists = (Test-Path $editorDir -PathType Container) -or 
                         ($editorName -like "*VS Code*" -and $vsCodeInPath) -or 
                         ($editorName -like "*Cursor*" -and $cursorInPath)
        } else {
            # For other paths, check the path directly
            $pathExists = Test-Path $editorPath
        }

        if ($pathExists) {
            $editorsFound = $true
            Write-Info "Found $editorName"

            # Only process settings.json paths
            if ($editorPath.EndsWith("settings.json")) {
                # Initialize settings object
                $editorSettings = @{}

                # Read existing settings if file exists
                if (Test-Path $editorPath) {
                    try {
                        $existingContent = Get-Content $editorPath -Raw
                        if ($existingContent.Trim()) {
                            $editorSettings = $existingContent | ConvertFrom-Json -AsHashtable
                        }
                    } catch {
                        Write-Warning "Could not parse existing $editorName settings, creating new ones"
                        $editorSettings = @{}
                    }
                } else {
                    # Create directory if it doesn't exist
                    if (-not (Test-Path $editorDir)) {
                        New-Item -ItemType Directory -Path $editorDir -Force | Out-Null
                        Write-Info "Created $editorName settings directory: $editorDir"
                    }
                }

                # Set terminal font settings
                $editorSettings["terminal.integrated.fontFamily"] = "MesloLGM Nerd Font Mono"
                $editorSettings["terminal.integrated.fontSize"] = 12

                # Cursor AI specific font settings
                if ($editorName -like "*Cursor*") {
                    # Cursor AI may use additional font settings
                    $editorSettings["editor.fontFamily"] = "MesloLGM Nerd Font Mono, 'Courier New', monospace"
                    $editorSettings["editor.fontLigatures"] = $true
                }

                # Convert back to JSON and save
                $jsonContent = $editorSettings | ConvertTo-Json -Depth 10
                $jsonContent | Set-Content $editorPath -Encoding UTF8
                Write-Success "$editorName configured successfully"
            }
        }
    }

    if (-not $editorsFound) {
        Write-Warning "Neither VS Code nor Cursor AI found in standard locations"
        Write-Info "Checked locations:"
        foreach ($location in $editorPaths) {
            Write-Info "  - $($location.Path)"
        }
        Write-Info "Manual configuration steps:"
        Write-Info "  1. Open your editor (VS Code or Cursor AI)"
        Write-Info "  2. Press Ctrl+, to open settings"
        Write-Info "  3. Search for 'terminal.integrated.fontFamily'"
        Write-Info "  4. Set to: MesloLGM Nerd Font Mono"
    }
} catch {
    Write-Warning "Failed to configure editors: $($_.Exception.Message)"
    Write-Info "Manual configuration steps:"
    Write-Info "  1. Open your editor (VS Code or Cursor AI)"
    Write-Info "  2. Press Ctrl+, to open settings"
    Write-Info "  3. Search for 'terminal.integrated.fontFamily'"
    Write-Info "  4. Set to: MesloLGM Nerd Font Mono"
}

# **FIXED: Configure PowerShell profile without invalid experimental features**
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

    # **FIXED: Create profile content without invalid experimental features**
    $profileContent = @"
# Oh My Posh initialization with $Theme theme and shell integration
oh-my-posh init pwsh --config "`$env:POSH_THEMES_PATH\$Theme.omp.json" | Invoke-Expression

# **FIXED: Enhanced PowerShell styling without experimental features**
# Configure file and directory colors (built-in feature in PowerShell 7.2+)
if (`$PSVersionTable.PSVersion.Major -ge 7 -and `$PSVersionTable.PSVersion.Minor -ge 2) {
    `$PSStyle.FileInfo.Directory = "`$(`$PSStyle.Bold)`$(`$PSStyle.Foreground.Blue)"
    `$PSStyle.FileInfo.Executable = "`$(`$PSStyle.Bold)`$(`$PSStyle.Foreground.Green)"
    `$PSStyle.FileInfo.Extension['.ps1'] = "`$(`$PSStyle.Foreground.Cyan)"
}

# PSFzf configuration with enhanced key bindings
Import-Module PSFzf
Set-PsFzfOption -PSReadlineChordProvider 'Ctrl+t' -PSReadlineChordReverseHistory 'Ctrl+r'

# Enhanced shell integration and IntelliSense features
if (`$PSVersionTable.PSVersion.Major -ge 7 -and `$PSVersionTable.PSVersion.Minor -ge 2) {
    # Enable predictive IntelliSense
    Set-PSReadLineOption -PredictionSource History
    Set-PSReadLineOption -PredictionViewStyle ListView

    # Enhanced command completion
    Set-PSReadLineOption -ShowToolTips
    Set-PSReadLineOption -HistorySearchCursorMovesToEnd
}

# Enhanced error handling and command completion
Set-PSReadLineOption -BellStyle None
Set-PSReadLineOption -EditMode Emacs
Set-PSReadLineKeyHandler -Key Tab -Function Complete

# Additional useful key bindings
Set-PSReadLineKeyHandler -Key UpArrow -Function HistorySearchBackward
Set-PSReadLineKeyHandler -Key DownArrow -Function HistorySearchForward

# Custom aliases and functions can be added below
"@

    # Write profile
    $profileContent | Set-Content $profilePath -Encoding UTF8
    Write-Success "PowerShell profile configured successfully with $Theme theme"
    Write-Info "Profile location: $profilePath"
    Write-Info "Shell integration features enabled:"
    Write-Info "  • Enhanced file and directory colors"
    Write-Info "  • Predictive IntelliSense (PowerShell 7.2+)"
    Write-Info "  • Improved command completion"
    Write-Info "  • Enhanced history search"
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
Write-Host "   4. Try right-clicking in the terminal for cursor repositioning" -ForegroundColor Gray

Write-Host "`n🎨 Theme usage:" -ForegroundColor Yellow
Write-Host "   - Current theme: $Theme" -ForegroundColor Gray
Write-Host "   - Browse themes: https://ohmyposh.dev/docs/themes" -ForegroundColor Gray
Write-Host "   - Change theme: Edit your PowerShell profile or re-run script" -ForegroundColor Gray

Write-Host "`n🔧 Features enabled:" -ForegroundColor Yellow
Write-Host "   - Enhanced shell integration with Oh My Posh" -ForegroundColor Gray
Write-Host "   - Cursor repositioning (try right-clicking in terminal)" -ForegroundColor Gray
Write-Host "   - URL detection in terminal output" -ForegroundColor Gray
Write-Host "   - Predictive IntelliSense and enhanced completion" -ForegroundColor Gray
Write-Host "   - Improved command history search" -ForegroundColor Gray

Write-Host "`n🔧 Theme selection options:" -ForegroundColor Yellow
Write-Host "   - Environment variable: `$env:POSH_THEME = 'atomic'" -ForegroundColor Gray
Write-Host "   - Download script: .\setup.ps1 -Theme 'atomic'" -ForegroundColor Gray
Write-Host "   - Default theme: quick-term" -ForegroundColor Gray
