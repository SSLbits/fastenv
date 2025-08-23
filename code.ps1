# PowerShell Environment Setup Script with Manual Font Installation
# Author: Simplified version with manual font installation
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

# **MANUAL FONT INSTALLATION FUNCTION**
function Install-MesloNerdFontManual {
    Write-Host "📦 Starting manual font installation process..." -ForegroundColor Magenta

    # Create temp directory
    $tempDir = Join-Path $env:TEMP "MesloNerdFont_Manual"
    if (Test-Path $tempDir) {
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    # Download latest Meslo.zip from GitHub
    $downloadUrl = "https://github.com/ryanoasis/nerd-fonts/releases/latest/download/Meslo.zip"
    $zipPath = Join-Path $tempDir "Meslo.zip"

    Write-Info "Downloading latest Meslo Nerd Font from GitHub..."
    Write-Info "URL: $downloadUrl"

    try {
        # Download with progress indication
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($downloadUrl, $zipPath)
        $webClient.Dispose()

        $fileSizeMB = [math]::Round((Get-Item $zipPath).Length / 1MB, 2)
        Write-Success "Downloaded: $fileSizeMB MB"
    } catch {
        Write-Error "Download failed: $($_.Exception.Message)"
        return $false
    }

    # Extract the zip file
    Write-Info "Extracting font files..."
    try {
        Expand-Archive -Path $zipPath -DestinationPath $tempDir -Force
        Write-Success "Extraction completed"
    } catch {
        Write-Error "Extraction failed: $($_.Exception.Message)"
        return $false
    }

    # **ENHANCED: Delete zip file immediately after extraction**
    Write-Info "Removing zip file..."
    try {
        Remove-Item $zipPath -Force
        Write-Success "Zip file removed"
    } catch {
        Write-Warning "Could not remove zip file: $($_.Exception.Message)"
    }

    # Clean up unwanted files (README.md, LICENSE)
    Write-Info "Cleaning up documentation files..."
    $unwantedFiles = Get-ChildItem -Path $tempDir -Recurse | Where-Object { 
        $_.Name -like "*README*" -or 
        $_.Name -like "*LICENSE*" -or 
        $_.Name -like "*OFL*" -or
        $_.Extension -eq ".md" -or
        $_.Extension -eq ".txt"
    }

    $cleanedCount = 0
    foreach ($file in $unwantedFiles) {
        try {
            Remove-Item $file.FullName -Force
            Write-Info "Removed: $($file.Name)"
            $cleanedCount++
        } catch {
            Write-Warning "Could not remove: $($file.Name)"
        }
    }
    Write-Success "Cleaned up $cleanedCount documentation files"

    # Count remaining font files
    $fontFiles = Get-ChildItem -Path $tempDir -Recurse -Include "*.ttf", "*.otf"
    Write-Success "Found $($fontFiles.Count) font files ready for installation"

    # Open Explorer for manual installation
    Write-Host "`n🎯 MANUAL INSTALLATION - SIMPLIFIED:" -ForegroundColor Yellow
    Write-Host "   1. Explorer window will open with ONLY font files" -ForegroundColor Gray
    Write-Host "   2. Press Ctrl+A to select ALL files" -ForegroundColor Gray
    Write-Host "   3. Right-click and choose 'Install' or 'Install for all users'" -ForegroundColor Gray
    Write-Host "   4. Wait for installation to complete" -ForegroundColor Gray
    Write-Host "   5. Close the Explorer window" -ForegroundColor Gray
    Write-Host "   6. Return here and press Enter to continue" -ForegroundColor Gray
    Write-Host "   💡 No need to avoid any files - everything shown is a font!" -ForegroundColor Green

    Write-Host "`n🚀 Opening Explorer window..." -ForegroundColor Green
    try {
        # Open the temp directory in Explorer
        Start-Process "explorer.exe" -ArgumentList $tempDir -WindowStyle Normal
        Write-Success "Explorer opened at: $tempDir"
    } catch {
        Write-Error "Failed to open Explorer: $($_.Exception.Message)"
        Write-Info "Manual path: $tempDir"
        return $false
    }

    # Pause script until user presses Enter
    Write-Host "`n⏸️  PAUSED - Waiting for manual font installation..." -ForegroundColor Cyan
    Write-Host "Press Enter when you've finished installing the fonts and closed Explorer..." -ForegroundColor Yellow
    Read-Host

    Write-Success "Resuming script execution..."

    # Clean up temp directory
    Write-Info "Cleaning up temporary files..."
    try {
        Remove-Item -Path $tempDir -Recurse -Force
        Write-Success "Temporary files cleaned up"
    } catch {
        Write-Warning "Could not clean up temp directory: $($_.Exception.Message)"
        Write-Info "Manual cleanup may be needed: $tempDir"
    }

    Write-Success "Manual font installation process completed!"
    return $true
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

# **MAIN SCRIPT EXECUTION**

# Validate theme
if (-not (Test-OhMyPoshTheme -ThemeName $Theme)) {
    Write-Warning "Invalid theme: $Theme"
    Write-Info "Valid themes can be found at: https://ohmyposh.dev/docs/themes"
    Write-Info "Using default theme: quick-term"
    $Theme = "quick-term"
}

Write-Host "🚀 Setting up PowerShell environment with theme: $Theme" -ForegroundColor Magenta

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

# **MANUAL FONT INSTALLATION**
Install-MesloNerdFontManual

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

        # Configure font with multiple fallback options
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

# **ULTRA-CLEAN: Configure PowerShell profile with essential elements only**
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

    # **MINIMAL: Essential profile content only**
    $profileContent = @"
# Oh My Posh initialization
oh-my-posh init pwsh --config "`$env:POSH_THEMES_PATH\$Theme.omp.json" | Invoke-Expression

# PSFzf configuration
Import-Module PSFzf -ErrorAction SilentlyContinue
Remove-PSReadLineKeyHandler -Key 'Ctrl+r' -ErrorAction SilentlyContinue
Remove-PSReadLineKeyHandler -Key 'Ctrl+t' -ErrorAction SilentlyContinue
Set-PsFzfOption -PSReadlineChordProvider 'Ctrl+t' -PSReadlineChordReverseHistory 'Ctrl+r'
Set-PsFzfOption -TabExpansion
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

Write-Host "`n📝 Next Steps:" -ForegroundColor Yellow
Write-Host "   1. **CLOSE THIS TERMINAL COMPLETELY**" -ForegroundColor Red
Write-Host "   2. Wait 5 seconds" -ForegroundColor Gray  
Write-Host "   3. **OPEN A NEW Windows Terminal**" -ForegroundColor Red
Write-Host "   4. Oh My Posh prompt should appear immediately!" -ForegroundColor Green
Write-Host "   5. Test fzf: Ctrl+T (files) and Ctrl+R (history)" -ForegroundColor Gray

Write-Host "`n🚀 Profile Features:" -ForegroundColor Yellow
Write-Host "   • Ultra-clean profile with just the essentials" -ForegroundColor Green
Write-Host "   • Lightning-fast loading" -ForegroundColor Green
Write-Host "   • Oh My Posh + fzf ready to go" -ForegroundColor Green

Write-Host "`n🎯 Font Installation:" -ForegroundColor Yellow
Write-Host "   • Zip automatically deleted after extraction" -ForegroundColor Green
Write-Host "   • Just Ctrl+A and right-click install!" -ForegroundColor Green
Write-Host "   • No files to avoid - everything shown is a font" -ForegroundColor Green

Write-Host "`n🎨 Theme options:" -ForegroundColor Yellow
Write-Host "   - Environment: `$env:POSH_THEME = 'atomic'" -ForegroundColor Gray
Write-Host "   - Script parameter: .\setup.ps1 -Theme 'atomic'" -ForegroundColor Gray
Write-Host "   - Current: $Theme" -ForegroundColor Gray
