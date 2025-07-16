# Download and run directly from GitHub
Invoke-Expression (Invoke-WebRequest -Uri "https://raw.githubusercontent.com/SSLbits/fastenv/refs/heads/main/code.ps1").Content

# Or download first, then run
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/yourusername/yourrepo/main/setup-powershell-environment.ps1" -OutFile "setup-powershell-environment.ps1"
.\setup-powershell-environment.ps1
