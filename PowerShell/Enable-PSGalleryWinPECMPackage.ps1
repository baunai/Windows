param (
    [string]$ModuleName
)

Function Write-CMLogEntry {	
    param (
        [parameter(Mandatory = $true, HelpMessage = "Value added to the log file.")]
        [ValidateNotNullOrEmpty()]
        [string]$Value,

        [parameter(Mandatory = $true, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("1", "2", "3")]
        [string]$Severity,

        [parameter(Mandatory = $true, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
        [ValidateNotNullOrEmpty()]
        [string]$ScriptLineNumber = "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})",

        [parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = "Invoke-WinPEPSGalleryModules.log"	
    )
		
    # Determine log file location
    $Path = "$($env:windir)\Temp\Logs"
       if (-NOT(Test-Path $Path)) {
            New-Item -ItemType Directory -Path $Path | Out-Null
        }

    $LogFilePath = Join-Path -Path $env:windir -ChildPath "Temp\Logs\$($FileName)"

    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $LogText = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    $LineFormat = $Value, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$ScriptLineNumber", $Severity
    $LogText = $LogText -f $LineFormat

    # Add value to log file
    try {
        Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop 
    }		
    catch [System.Exception] {
        Write-Warning -Message "Unable to append log entry to $($FileName) file. Error message: $($_.Exception.Message)"
    }
}

# Enable TLS 1.2 support for downloading modules from PSGallery (Required)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Write-Host "Enable TLS 1.2 support for downloading modules from PSGallery (Required)" -Verbose
Write-CMLogEntry -Value "TLS 1.2 support for downloading modules from PSGallery (Required)" -Severity 1 -ScriptLineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"

#Setup LOCALAPPDATA Variable
[System.Environment]::SetEnvironmentVariable('LOCALAPPDATA', "$env:SystemDrive\Windows\system32\config\systemprofile\AppData\Local")
Write-Host "Setup LOCALAPPDATA Variable" -Verbose
Write-CMLogEntry -Value "Setup LOCALAPPDATA Variable" -Severity 1 -ScriptLineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"

$WorkingDir = $env:TEMP

#PowerShellGet

if (!(Get-Module -Name PowerShellGet)) {
	Write-Host "PowerShellGet Module does not exist. Trying to download and expand the package" -Verbose
    Write-CMLogEntry -Value "PowerShellGet Module does not exist. Trying to download and expand the package" -Severity 1 -ScriptLineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Copy-Item "$PSScriptRoot\powershellget.2.2.5.nupkg" -Destination "$WorkingDir\powershellget.2.2.5.zip"
    Expand-Archive -Path "$WorkingDir\powershellget.2.2.5.zip" -DestinationPath "$WorkingDir\2.2.5"
    $Null = New-Item -Path "$env:ProgramFiles\WindowsPowerShell\Modules\PowerShellGet" -ItemType Directory -ErrorAction SilentlyContinue
    Move-Item -Path "$WorkingDir\2.2.5" -Destination "$env:ProgramFiles\WindowsPowerShell\Modules\PowerShellGet\2.2.5"
	Write-Host "Move PowerShellGet to its directory" -Verbose
    Write-CMLogEntry -Value "PowerShellGet moved to its directory" -Severity 1 -ScriptLineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
}


#PackageManagement
if (!(Get-Module -Name PackageManagement)) {
	Write-Host "PackageMangement does not exist. Trying to download and expand the package" -Verbose
    Write-CMLogEntry -Value "PackageManagement does not exist. Trying to download and expand the package" -Severity 1 -ScriptLineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Copy-Item "$PSScriptRoot\packagemanagement.1.4.7.nupkg" -Destination "$WorkingDir\packagemanagement.1.4.7.zip"
    Expand-Archive -Path "$WorkingDir\packagemanagement.1.4.7.zip" -DestinationPath "$WorkingDir\1.4.7"
    $Null = New-Item -Path "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement" -ItemType Directory -ErrorAction SilentlyContinue
    Move-Item -Path "$WorkingDir\1.4.7" -Destination "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement\1.4.7"
	Write-Host "Move PackageManagement to its directory" -Verbose
    Write-CMLogEntry -Value "PackageMangement moved to its directory" -Severity 1 -ScriptLineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
}

#Import PowerShellGet
Import-Module PowerShellGet
if (Get-InstalledModule PowerShellGet) {
    Write-Host "PowerShellGet module imported." -Verbose
    Write-CMLogEntry -Value "PowerShellGet module imported." -Severity 1 -ScriptLineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
}


#Install Module from PSGallery
Install-Module -Name $ModuleName -Force -AcceptLicense -SkipPublisherCheck
Import-Module -Name $ModuleName -Force
if (Get-InstalledModule -Name $ModuleName) {
    Write-Host "$($ModuleName) successfully installed and imported." -Verbose
    Write-CMLogEntry -Value "$($ModuleName) successfully installed and imported." -Severity 1 -ScriptLineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
} else {
    Write-Host "$($ModuleName) not found."
    Write-CMLogEntry -Value "$($ModuleName) not found" -Severity 3 -ScriptLineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
}

