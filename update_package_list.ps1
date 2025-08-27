# Get current date in YYYY-MM-DD format
$Date = Get-Date -Format "yyyy-MM-dd"

# Define your custom output folder path
$OutputFolder = "C:\Users\EKim\OneDrive - LS Cable\PM - EugeneKim\GITHUB_REPO_CLONE\busway_project_automation"

# Make sure folder exists
if (!(Test-Path -Path $OutputFolder)) {
    New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
}

# Define output file name with date
$OutputFile = "$OutputFolder\python_packages_$Date.txt"

# Export installed Python packages
pip freeze | Out-File -FilePath $OutputFile -Encoding UTF8

Write-Output "✅ Python packages list exported to $OutputFile"
