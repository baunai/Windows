function Write-LogEntry {
    param(
        [parameter(Mandatory = $true, HelpMessage = "Value added to the Install-Chrome.log file")]
        [ValidateNotNullOrEmpty()]
        [string]$Value,

        [parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = "Install-Chrome.log"
    )
    # Determine Log File Location
    $LogFilePath = Join-Path -Path $env:windir -ChildPath "Temp\$($FileName)"

    # Add value to log file
    try {
        Out-File -InputObject $Value -Append -NoClobber -Encoding default -FilePath $LogFilePath -ErrorAction Stop
    }
    catch [System.Exception] {
        Write-Warning -Message "Unable to append log entry to $($FileName) file"
    }
}

# Initial logging
Write-LogEntry -Value "Starting download and install Chrome Enterprise x64 on 64-bit OS"

# Define temorary location to cache Chrome Installer
$tempDirectory = "$env:TEMP\Chrome"

# Silently run the script
$RunScriptSilent = $true

# Set the system architecture as a value
$OSArchitecture = (Get-WmiObject Win32_OperatingSystem).OSArchitecture
Write-LogEntry -Value "OS Architecture is $($OSArchitecture)"

# Exit if the script was not run with Administrator priveleges
$User = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
if (-not $User.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )) {
    Write-Host 'Please run again with Administrator privileges.' -ForegroundColor Red
    Write-LogEntry -Value "Script was not run with Administrator privilleges. Please run again with Administrator privilleges"
    if ($RunScriptSilent -NE $True) {
        Read-Host 'Press [Enter] to exit'
    }
    exit
}

# Enable TLS 1.2 support for downloading modules from PSGallery (Required)
(New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


Function Download-Chrome {
    Write-LogEntry -Value 'Downloading Google Chrome... '

    if ($OSArchitecture -eq "64-Bit") {
        $Link = 'http://dl.google.com/edgedl/chrome/install/GoogleChromeStandaloneEnterprise64.msi'
    }
    ELSE {
        $Link = 'http://dl.google.com/edgedl/chrome/install/GoogleChromeStandaloneEnterprise.msi'
    }
    

    # Download the installer from Google
    try {
        New-Item -ItemType Directory "$TempDirectory" -Force | Out-Null
	        (New-Object System.Net.WebClient).DownloadFile($Link, "$TempDirectory\Chrome.msi")
        Write-Host 'Success downloading msi file!' -ForegroundColor Green
        Write-LogEntry -Value "Success downloading msi file!"
    }
    catch {
        Write-Host 'failed. There was a problem with the download.' -ForegroundColor Red
        Write-LogEntry -Value "Failed. There was a problem with the download"
    }
}

Function Install-Chrome {
    Write-Host 'Installing Chrome... ' -NoNewline
    Write-LogEntry -Value "Installing Chrome x64 Enterprise. This will take several minutes..."

    # Install Chrome
    $ChromeMSI = """$TempDirectory\Chrome.msi"""
    $ExitCode = (Start-Process -filepath msiexec -argumentlist "/i $ChromeMSI /qn /norestart" -Wait -PassThru).ExitCode
    
    if ($ExitCode -eq 0) {
        Write-Host 'Success!' -ForegroundColor Green
        Write-LogEntry -Value "Success!"
    }
    else {
        Write-Host "failed. There was a problem installing Google Chrome. MsiExec returned exit code $ExitCode." -ForegroundColor Red
        Write-LogEntry -Value "Failed. There was a problem installing Google Chrome. MisExec returned exit code $($ExitCode)"
        Clean-Up
    }
}

Function Clean-Up {
    Write-Host 'Removing Chrome installer... ' -NoNewline
    Write-LogEntry -Value "Removing Chrome Installer from $($TempDirectory)"

    try {
        # Remove the installer
        Remove-Item "$TempDirectory\Chrome.msi" -ErrorAction Stop
        Write-Host 'Success!' -ForegroundColor Green
        Write-LogEntry -Value "Success!"
    }
    catch {
        Write-Host "failed. You will have to remove the installer yourself from $TempDirectory\." -ForegroundColor Yellow
        Write-LogEntry -Value "Failed. You will have to remove the installer yourself from $($TempDirectory)"
    }
}

Download-Chrome
Install-Chrome
Clean-Up

Write-Host "Installation completed"
Write-LogEntry -Value "Installation completed!"
