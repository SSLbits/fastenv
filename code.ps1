# PowerShell Environment Setup Script with Theme Support
# Author: Enhanced version with theme selection, Cursor AI support, and improved fzf handling
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

# **NEW: Enhanced PATH refresh function**
function Update-SessionPath {
    Write-Info "Refreshing environment PATH..."
    $machinePath = [System.Environment]::GetEnvironmentVariable("PATH", "Machine")
    $userPath = [System.Environment]::GetEnvironmentVariable("PATH", "User")
    $env:PATH = "$machinePath;$userPath"
    Write-Info "PATH updated for current session"
}

# **NEW: fzf verification function**
function Test-FzfInstallation {
    Write-Info "Verifying fzf installation..."

    # Test if fzf command is available
    $fzfCommand = Get-Command fzf -ErrorAction SilentlyContinue
    if ($fzfCommand) {
        Write-Success "fzf executable found at: $($fzfCommand.Source)"

        # Test fzf version
        try {
            $fzfVersion = & fzf --version
            Write-Info "fzf version: $fzfVersion"
            return $true
        } catch {
            Write-Warning "fzf found but version check failed: $($_.Exception.Message)"
            return $false
        }
    } else {
        Write-Warning "fzf executable not found in PATH"

        # Check common installation locations
        $commonPaths = @(
            "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\junegunn.fzf_Microsoft.Winget.Source_8wekyb3d8bbwe",
            "$env:PROGRAMFILES\fzf",
            "$env:USERPROFILE\.fzf\bin"
        )

        foreach ($path in $commonPaths) {
            if (Test-Path "$path\fzf.exe") {
                Write-Info "Found fzf at: $path\fzf.exe"
                Write-Warning "fzf is installed but not in PATH. This will be fixed after terminal restart."
                return $true
            }
        }

        return $false
    }
}

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

# **ENHANCED: Install fzf with better verification**
Write-Info "Installing fzf..."
try {
    $fzfInstalled = Get-Command fzf -ErrorAction SilentlyContinue

    if ($Force -or -not $fzfInstalled) {
        Write-Info "Installing fzf via winget..."
        winget install junegunn.fzf -s winget --accept-package-agreements --accept-source-agreements
        Write-Success "fzf installation completed"

        # Update PATH after installation
        Update-SessionPath

        # Wait a moment for PATH to propagate
        Start-Sleep -Seconds 2

        # Verify installation
        if (Test-FzfInstallation) {
            Write-Success "fzf verification successful"
        } else {
            Write-Warning "fzf installed but not immediately accessible. This is normal - it will work after terminal restart."
        }
    } else {
        Write-Warning "fzf already installed. Use -Force to reinstall."
        # Still verify it's working
        Test-FzfInstallation | Out-Null
    }
} catch {
    Write-Error "Failed to install fzf: $($_.Exception.Message)"
    Write-Info "Manual installation: Download from https://github.com/junegunn/fzf/releases"
}

# **ENHANCED: Install PSFzf module with better error handling**
Write-Info "Installing PSFzf module..."
try {
    $psfzfInstalled = Get-Module -ListAvailable -Name PSFzf

    if ($Force -or -not $psfzfInstalled) {
        Write-Info "Installing PSFzf PowerShell module..."
        Install-Module -Name PSFzf -Force -Scope CurrentUser -AllowClobber
        Write-Success "PSFzf module installed successfully"
    } else {
        Write-Warning "PSFzf already installed. Use -Force to reinstall."
    }

    # Test if module can be imported
    try {
        Import-Module PSFzf -Force -ErrorAction Stop
        Write-Success "PSFzf module imported successfully"
    } catch {
        Write-Warning "PSFzf module installed but failed to import: $($_.Exception.Message)"
        Write-Info "This may be resolved after terminal restart"
    }
} catch {
    Write-Error "Failed to install PSFzf module: $($_.Exception.Message)"
    Write-Info "Manual installation: Install-Module -Name PSFzf -Scope CurrentUser"
}

# Configure Windows Terminal
Write-Info "Configuring Windows Terminal..."
try {
    $wtSettingsPath = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"
    if (Test-Path $wtSettingsPath) {
        # Read current settings
        $settingsContent = Get-Content $wtSettingsPath -Raw
        $settings = $settingsContent | ConvertFrom-Json

        # Ensure correct JSON structure for modern Windows Terminal
        if (-not $settings.profiles) { 
            $settings | Add-Member -NotePropertyName "profiles" -NotePropertyValue @{} -Force
        }
        if (-not $settings.profiles.defaults) { 
            $settings.profiles | Add-Member -NotePropertyName "defaults" -NotePropertyValue @{} -Force
        }

        # Use correct property structure for font
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

# **ENHANCED: Configure PowerShell profile with robust fzf handling**
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

    # **ENHANCED: Create profile content with robust fzf error handling**
    $profileContent = @"
# Oh My Posh initialization with $Theme theme
oh-my-posh init pwsh --config "`$env:POSH_THEMES_PATH\$Theme.omp.json" | Invoke-Expression

# **ENHANCED: PSFzf configuration with error handling**
try {
    # Import PSFzf module
    Import-Module PSFzf -ErrorAction Stop

    # Configure key bindings
    Set-PsFzfOption -PSReadlineChordProvider 'Ctrl+t' -PSReadlineChordReverseHistory 'Ctrl+r'

    # Additional fzf options for better experience
    Set-PsFzfOption -TabExpansion

    # Verify fzf is working
    if (Get-Command fzf -ErrorAction SilentlyContinue) {
        Write-Host "✅ fzf loaded successfully - Press Ctrl+T for file search, Ctrl+R for history" -ForegroundColor Green
    } else {
        Write-Host "⚠️ fzf executable not found - try restarting terminal" -ForegroundColor Yellow
    }
} catch {
    Write-Host "⚠️ PSFzf failed to load: `$(`$_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "💡 Try: Install-Module PSFzf -Force" -ForegroundColor Cyan
}

# **NEW: Enhanced PowerShell settings for better terminal experience**
# Configure file and directory colors (built-in feature in PowerShell 7.2+)
if (`$PSVersionTable.PSVersion.Major -ge 7 -and `$PSVersionTable.PSVersion.Minor -ge 2) {
    `$PSStyle.FileInfo.Directory = "`$(`$PSStyle.Bold)`$(`$PSStyle.Foreground.Blue)"
    `$PSStyle.FileInfo.Executable = "`$(`$PSStyle.Bold)`$(`$PSStyle.Foreground.Green)"
    `$PSStyle.FileInfo.Extension['.ps1'] = "`$(`$PSStyle.Foreground.Cyan)"

    # Enable predictive IntelliSense
    Set-PSReadLineOption -PredictionSource History -PredictionViewStyle ListView
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
    Write-Info "fzf error handling and verification included in profile"
} catch {
    Write-Error "Failed to configure PowerShell profile: $($_.Exception.Message)"
}

# Display completion message
Write-Host "`n🎉 Setup completed successfully!" -ForegroundColor Green
Write-Host "🎨 Theme configured: $Theme" -ForegroundColor Magenta

Write-Host "`n📝 Next steps:" -ForegroundColor Yellow
Write-Host "   1. Close ALL PowerShell/Windows Terminal windows completely" -ForegroundColor Gray
Write-Host "   2. Wait 5 seconds" -ForegroundColor Gray  
Write-Host "   3. Open a NEW Windows Terminal (as regular user, not admin)" -ForegroundColor Gray
Write-Host "   4. Your terminal should display the $Theme theme with proper icons" -ForegroundColor Gray
Write-Host "   5. **IMPORTANT**: Test fzf functionality:" -ForegroundColor Yellow
Write-Host "      • Press Ctrl+T for fuzzy file search" -ForegroundColor Gray
Write-Host "      • Press Ctrl+R for fuzzy history search" -ForegroundColor Gray
Write-Host "      • If fzf doesn't work immediately, wait 30 seconds and try again" -ForegroundColor Gray

Write-Host "`n🔧 fzf verification commands:" -ForegroundColor Yellow
Write-Host "   - Test fzf: Get-Command fzf" -ForegroundColor Gray
Write-Host "   - Test PSFzf: Get-Module PSFzf -ListAvailable" -ForegroundColor Gray
Write-Host "   - Manual fix: Install-Module PSFzf -Force" -ForegroundColor Gray

Write-Host "`n🎨 Theme usage:" -ForegroundColor Yellow
Write-Host "   - Current theme: $Theme" -ForegroundColor Gray
Write-Host "   - Browse themes: https://ohmyposh.dev/docs/themes" -ForegroundColor Gray
Write-Host "   - Change theme: Edit your PowerShell profile or re-run script" -ForegroundColor Gray

Write-Host "`n🔧 Theme selection options:" -ForegroundColor Yellow
Write-Host "   - Environment variable: `$env:POSH_THEME = 'atomic'" -ForegroundColor Gray
Write-Host "   - Download script: .\setup.ps1 -Theme 'atomic'" -ForegroundColor Gray
Write-Host "   - Default theme: quick-term" -ForegroundColor Gray
