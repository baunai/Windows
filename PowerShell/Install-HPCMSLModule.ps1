function Get-TaskSequenceStatus {
    # Determine if a task sequence is currently running
    try {
        $TSEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment
    }
    catch {}
    if ($null -eq $TSEnv) {
        return $false
    }
    else {
        try {
            $SMSTSType = $TSEnv.Value("_SMSTSType")
        }
        catch {}
        if ($null -eq $SMSTSType) {
            return $false
        }
        else {
            return $true
        }
    }
}


function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "Message added to the log file.")]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter(Mandatory = $false, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning, 3 for Error.")]
        [ValidateNotNullOrEmpty()]
        [ValidateRange(1, 3)]
        [int16]$Severity = 1,

        [Parameter(Mandatory = $false, HelpMessage = "Output script run to console host")]
        [ValidateNotNullOrEmpty()]
        [Boolean]$WriteHost = $true,

        [Parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = "Install-HPCMSLModule.log"
    )
    
    if (Get-TaskSequenceStatus) {
        $TSEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment
        $LogDir = $TSEnv.Value("_SMSTSLogPath")
        $LogFilePath = Join-Path -Path $LogDir -ChildPath $FileName
    }
    else {
        $LogDir = Join-Path -Path "${env:SystemRoot}" -ChildPath "Temp"
        $LogFilePath = Join-Path -Path $LogDir -ChildPath $FileName
    }

    $global:ScriptLogFilePath = $LogFilePath
    $VerbosePreference = 'Continue'

    foreach ($msg in $Message) {
        # Create script block for writting log entry to the console
        [scriptblock]$WriteLogLineToHost = {
            Param (
                [string]$lTextLogLine,
                [Int16]$lSeverity
            )
            switch ($lSeverity) {
                3 { $Style = "$($PSStyle.Bold)$($PSStyle.Foreground.BrightRed)"; Write-Host "$($Style)ERROR: $lTextLogLine" }
                2 { Write-Warning $lTextLogLine }
                1 { Write-Verbose $lTextLogLine }
            }
        }
        & $WriteLogLineToHost -lTextLogLine $msg -lSeverity $Severity 
    }

    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$($MyInvocation.ScriptName):$($MyInvocation.ScriptLineNumber)", $Severity
    $Line = $Line -f $LineFormat
    
    try {
        Out-File -InputObject $Line -Append -NoClobber -Encoding Default -FilePath $LogFilePath
    }
    catch [System.Exception] {
        # Exception is stored in the automatic variable _
        Write-Warning -Message "Unable to append log entry to $($LogFilePath) file. Error message: $($_.Exception.Message)"
    }

}

# Leave blank space at top of window to not block output by progress bars
Function AddHeaderSpace {
    Write-Output "This space intentionally left blank..."
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
}


Clear-Host
AddHeaderSpace

$Script_Start_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
Write-Log -Message "================================= Script Start ==================================" -Severity 1
Write-Host "Script log file path [$ScriptLogFilePath]"

#Setup LOCALAPPDATA Variable
[System.Environment]::SetEnvironmentVariable('LOCALAPPDATA', "$env:SystemDrive\Windows\system32\config\systemprofile\AppData\Local")
Write-Log -Message "Setup LOCALAPPDATA Variable" -Severity 1 
# Enable TLS 1.2 support for downloading modules from PSGallery (Required)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Register PSGallery
[system.net.webrequest]::defaultwebproxy = new-object system.net.webproxy('http://proxy.hpdwinad.hpd:8080')
[system.net.webrequest]::defaultwebproxy.credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[system.net.webrequest]::defaultwebproxy.BypassProxyOnLocal = $true
Register-PSRepository -Default -ErrorAction Ignore -Verbose
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue

# Remove old PackageManagement Module
$PkgMgmtPath = "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement\1.0.0.1"
if (!(Test-Path $PkgMgmtPath)) {
    Write-Log -Message "$($PkgMgmtPath) not found" -Severity 1 
}
else {
    Write-Log -Message "$($PkgMgmtPath) does exist. Start Removing" -Severity 1 
    Remove-Item $PkgMgmtPath -Force -Confirm:$false -Recurse
    Write-Log -Message "$($PkgMgmtPath) removed" -Severity 1 
}

$PackageProvider = Get-PackageProvider -Name "NuGet"
if (!($PackageProvider)) {
    try {
        Write-Log -Message "Nuget package provider does not exist. Attempt to install"  -Severity 1
        # Enable TLS 1.2 support for downloading modules from PSGallery (Required)
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Install-PackageProvider -Name "NuGet" -Force -ErrorAction Stop -Verbose
    }
    catch [System.Exception] {
        Write-Log -Message "Unable to install latest NuGet package provider. Error message: $($_.Exception.Message)" -Severity 3 
    }
}
elseif ($PackageProvider.Version.ToString() -le "2.8.5") {
    Write-Log -Message "Nuget Package Provider must be updated" -Severity 2 
    Install-PackageProvider -Name "NuGet" -Force -ErrorAction Stop -Verbose
}
else {
    Write-Log -Message "NuGet Package Provider OK, Checking for modules" -Severity 1 
}


# Install the latest PowershellGet Module
if ($PackageProvider.Version.ToString() -ge "2.8.5") { 
    $PowerShellGetInstalledModule = Get-InstalledModule -Name "PowerShellGet" -ErrorAction SilentlyContinue
    if ($PowerShellGetInstalledModule -ne $null) {
        try {
            # Attempt to locate the latest available version of the PowerShellGet module from repository
            $PowerShellGetLatestModule = Find-Module -Name "PowerShellGet" -ErrorAction Stop -Verbose:$false
            if ($PowerShellGetLatestModule -ne $null) {
                if ($PowerShellGetInstalledModule.Version -lt $PowerShellGetLatestModule.Version) {
                    Write-Log -Message "Attempting to request the latest PowerShellGet module version from repository" -Severity 1 
                    try {
                        # Newer module detected, attempt to update
                        Write-Log -Message "Newer version detected, attempting to update the PowerShellGet module from repository" -Severity 1 
                        Update-Module -Name "PowerShellGet" -Scope "AllUsers" -Force -ErrorAction Stop -Confirm:$false -Verbose
                    }
                    catch [System.Exception] {
                        Write-Log -Message "Failed to update the PowerShellGet module. Error message: $($_.Exception.Message)" -Severity 3 
                    }
                }
            }
            else {
                Write-Log -Message "Location request for the latest available version of the PowerShellGet module failed, can't continue" -Severity 3 
            }
        }
        catch [System.Exception] {
            Write-Log -Message "Failed to retrieve the latest available version of the PowerShellGet module, can't continue. Error message: $($_.Exception.Message)" -Severity 3 
        }
    }
    else {
        try {
            # PowerShellGet module was not found, attempt to install from repository
            Write-Log -Message "PowerShellGet module was not found, attempting to install it including dependencies from repository" -Severity 1 
            Write-Log -Message "Attempting to install PackageManagement module from repository" -Severity 1 
            Install-Module -Name "PackageManagement" -Force -Scope AllUsers -AllowClobber -ErrorAction Stop -Verbose
            Write-Log -Message "Attempting to install PowerShellGet module from repository" -Severity 1 
            Install-Module -Name "PowerShellGet" -Force -Scope AllUsers -AllowClobber -ErrorAction Stop -Verbose
        }
        catch [System.Exception] {
            Write-Log -Message "Unable to install PowerShellGet module from repository. Error message: $($_.Exception.Message)" -Severity 3 
        }
    }
 
    # Check PowerShellGet Module Version
    if ($PowerShellGetInstalledModule.Version -eq $PowerShellGetLatestModule.Version) {
        Write-Log -Message "PowerShellGet Module is up to date. Module version: $($PowerShellGetInstalledModule.Version)" -Severity 1  
    }
} 

# Validate that script is executed on HP hardware
$Manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty Manufacturer).Trim()
switch -Wildcard ($Manufacturer) {
    "*HP*" {
        Write-Log -Message "Validated HP hardware check" -Severity 1  
    }
    "*Hewlett-Packard*" {
        Write-Log -Message "Validated HP hardware check" -Severity 1 
    }
    default {
        Write-Log -Message "Not running on HP Hardware, Script not applicable" -Severity 2 ; exit 0
    }
}

    #Install the latest HPCMSL Module
    $HPInstalledModule = Get-InstalledModule | Where-Object { $_.Name -match "HPCMSL" } -ErrorAction SilentlyContinue -Verbose:$false
    if ($HPInstalledModule -ne $null) {
        $HPGetLatestModule = Find-Module -Name "HPCMSL" -ErrorAction Stop -Verbose:$false
        if ($HPInstalledModule.Version -lt $HPGetLatestModule.Version) {
            Write-Log -Message "Newer HPCMSL version detected, updating from repository" -Severity 1 
                try {
                    # Install HP Client Management Script Library
                    Write-Log -Message "Attempting to install HPCMSL module from repository" -Severity 1 
                    Update-Module -Name "HPCMSL" -Force -ErrorAction Stop -Verbose
                } 
                catch [System.Exception] {
                    Write-Log -Message "Unable to install HPCMSL module from repository. Error message: $($_.Exception.Message)" -Severity 3 
                }
        }
        else {
            Write-Log -Message "HPCMSL Module is up to date. HPCMSL module version: $((Get-InstalledModule -Name "HPCMSL").Version)" -Severity 1 
        }
    }
    else {
        Write-Log -Message "HPCMSL Module is missing, try to install from repository" -Severity 1 
            try {
                # Install HP Client Management Script Library
                Write-Log -Message "Attempting to install HPCMSL module from repository" -Severity 1  
                Install-Module -Name "HPCMSL" -SkipPublisherCheck -AcceptLicense -Scope AllUsers -Force -ErrorAction Stop -Verbose
            } 
            catch [System.Exception] {
                Write-Log -Message "Unable to install HPCMSL module from repository. Error message: $($_.Exception.Message)" -Severity 3 
            }
    # Final check for HPCMSL Module version
    Write-Log -Message "HPCMSL module version: $((Get-InstalledModule -Name "HPCMSL").Version)" -Severity 1 
 } 


$Script_End_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
$Script_Time_Taken = New-TimeSpan -Start $Script_Start_Time -End $Script_End_Time

Write-Log -Message "Script Start: $Script_Start_Time" -Severity 1
Write-Log -Message "Script end: $Script_End_Time" -Severity 1
Write-Log -Message "Execution time: $Script_Time_Taken" -Severity 1
Write-Log -Message "================================ Script End =====================================" -Severity 1
