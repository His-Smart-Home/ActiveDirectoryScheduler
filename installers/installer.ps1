# Configuration
$DestinationPath = "C:\HSMIT\"
$Repo = "His-Smart-Home/ActiveDirectoryScheduler"
$ApiUrl = "https://api.github.com/repos/$Repo/releases/latest"
$Headers = @{ "User-Agent" = "PowerShell" }

# Get the latest release info
$ReleaseInfo = Invoke-RestMethod -Uri $ApiUrl -Headers $Headers

# Find the zipball URL
$ZipUrl = $ReleaseInfo.zipball_url

# Temp file path
$TempZip = "$env:TEMP\ActiveDirectoryScheduler.zip"

# Download the zip
Invoke-WebRequest -Uri $ZipUrl -OutFile $TempZip -Headers $Headers

# Ensure the destination exists
if (!(Test-Path -Path $DestinationPath)) {
    New-Item -ItemType Directory -Path $DestinationPath | Out-Null
}

# Extract the zip
Expand-Archive -Path $TempZip -DestinationPath $DestinationPath -Force

# Clean up the zip
Remove-Item $TempZip

Write-Host "Latest release downloaded and extracted to: $DestinationPath"
