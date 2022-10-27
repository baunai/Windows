# Create Write-Log function
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
        [string]$FileName = "Update-CMHPDriverPackage.log"
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

    # $global:ScriptLogFilePath = $LogFilePath
    $VerbosePreference = 'Continue'

    if ($WriteHost) {
        foreach ($msg in $Message) {
            # Create script block for writting log entry to the console
            [scriptblock]$WriteLogLineToHost = {
                Param (
                    [string]$lTextLogLine,
                    [Int16]$lSeverity
                )
                switch ($lSeverity) {
                    3 { $Style = "$($PSStyle.Bold)$($PSStyle.Foreground.Red)"; Write-Host "$($Style)$lTextLogLine" }
                    2 { $Style = "$($PSStyle.Bold)$($PSStyle.Foreground.Yellow)"; Write-Host "$($Style)$lTextLogLine" }
                    1 { $Style = "$($PSStyle.Bold)$($PSStyle.Foreground.White)"; Write-Host "$($Style)$lTextLogLine" }
                    #3 { Write-Error $lTextLogLine }
                    #2 { Write-Warning $lTextLogLine }
                    #1 { Write-Verbose $lTextLogLine }
                }
            }
            & $WriteLogLineToHost -lTextLogLine $msg -lSeverity $Severity 
        }
    }

    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$($FileName.Substring(0,$FileName.Length-4)):$($MyInvocation.ScriptLineNumber)", $Severity
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
}

AddHeaderSpace

$Script_Start_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
Write-Log -Message "INFO: Script Start: $Script_Start_Time"

# Validate that script is executed on HP hardware
$Manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty Manufacturer).Trim()
switch -Wildcard ($Manufacturer) {
    "*HP*" {
        Write-Log -Message "Validated HP hardware check"  
    }
    "*Hewlett-Packard*" {
        Write-Log -Message "Validated HP hardware check" 
    }
    default {
        Write-Log -Message "Not running on HP Hardware, Script not applicable" -Severity 2 ; exit 0
    }
}

#Install the latest HPCMSL Module
$HPInstalledModule = Get-Module -Name HPCMSL -ListAvailable | Select-Object -First 1
$HPCMSLInstalledVer = $HPInstalledModule.Version.ToString()
if ($null -ne $HPInstalledModule) {
    $HPGetLatestModule = Find-Module -Name "HPCMSL" -ErrorAction Stop -Verbose:$false
    if ($HPCMSLInstalledVer -lt $HPGetLatestModule.Version) {
        Write-Log -Message "Newer HPCMSL version detected, updating from repository" 
        try {
            # Install HP Client Management Script Library
            Write-Log -Message "Attempting to install HPCMSL module from repository" 
            Update-Module -Name "HPCMSL" -Force -ErrorAction Stop -Verbose
        } 
        catch [System.Exception] {
            Write-Log -Message "Unable to install HPCMSL module from repository. Error message: $($_.Exception.Message)" -Severity 3 
        }
    }
    else {
        Write-Log -Message "HPCMSL Module is up to date. HPCMSL module version: $((Get-InstalledModule -Name "HPCMSL").Version)" 
    }
}
else {
    Write-Log -Message "HPCMSL Module is missing, try to install from repository"
    try {
        # Install HP Client Management Script Library
        Write-Log -Message "Attempting to install HPCMSL module from repository" 
        Install-Module -Name "HPCMSL" -SkipPublisherCheck -AcceptLicense -Scope AllUsers -Force -ErrorAction Stop -Verbose
        Import-Module -Name HPCMSL
    } 
    catch [System.Exception] {
        Write-Log -Message "Unable to install HPCMSL module from repository. Error message: $($_.Exception.Message)" -Severity 3 
    }
    # Final check for HPCMSL Module version
    Write-Log -Message "HPCMSL module version: $((Get-InstalledModule -Name "HPCMSL").Version)" 
} 

#Script Variables
$OS = "Win11"
$OSVER = "21H2"
$DownloadDir = "C:\HPDrivers"
$FileServer = "\\hpdwinad.hpd\departmentFS\Support\SCCMOSDSource\DriverPackages\Hewlett-Packard\Drivers"
$SiteCode = "P01"

#Reset Vars
$DriverPack = ""
$Model = ""

$HPModelsTable = @(
    @{ ProdCode = '80fc'; Model = "Elite x2 1012 G1"; PackageID = "P01006DC" }
    @{ ProdCode = '82ca'; Model = "Elite x2 1012 G2"; PackageID = "P01006DD" }
    @{ ProdCode = '8414'; Model = "Elite x2 1013 G3"; PackageID = "P01006DE" }
    @{ ProdCode = '85B9'; Model = "Elite x2 G4"; PackageID = "P01006DF" }
    @{ ProdCode = '83d5'; Model = "EliteBook 755 G5"; PackageID = "P01006E0" }
    @{ ProdCode = '8079'; Model = "EliteBook 840 G3"; PackageID = "P01006E1" }
    @{ ProdCode = '83B2'; Model = "EliteBook 850 G5"; PackageID = "P01006E2" }
    @{ ProdCode = '1992'; Model = "ProBook 645 G1"; PackageID = "P01006E3" }
    @{ ProdCode = '80fe'; Model = "ProBook 655 G2"; PackageID = "P01006E5" }
    @{ ProdCode = '823a'; Model = "ProBook 655 G3"; PackageID = "P01006E6" }
    @{ ProdCode = '80fd'; Model = "ProBook 640 G2"; PackageID = "P01006E7" }
    @{ ProdCode = '85ad'; Model = "ProBook 455R G6"; PackageID = "P01006E9" }
    @{ ProdCode = '2215'; Model = "EliteDesk 705 G1"; PackageID = "P01006EA" }
    @{ ProdCode = '8265'; Model = "EliteDesk 705 G3"; PackageID = "P01006EB" }
    @{ ProdCode = '805a'; Model = "EliteDesk 705 G2"; PackageID = "P01006ED" }
    @{ ProdCode = '212b'; Model = "Z440 Workstation"; PackageID = "P01006F0" }
    @{ ProdCode = '81c6'; Model = "Z6 G4 Workstation"; PackageID = "P01006F1" }
    @{ ProdCode = '158a'; Model = "Z620 Workstation"; PackageID = "P01006F2" }
    @{ ProdCode = '81c7'; Model = "Z8 G4 Workstation"; PackageID = "P01006F3" }
    @{ ProdCode = '1905'; Model = "Z230 Tower Workstation"; PackageID = "P01006F4" }
    @{ ProdCode = '802f'; Model = "Z240 Tower Workstation"; PackageID = "P01006F5" }
    @{ ProdCode = '1589'; Model = "Z420 Workstation"; PackageID = "P01006F6" }
    @{ ProdCode = '212a'; Model = "Z640 Workstation"; PackageID = "P01006F7" }
    @{ ProdCode = '158b'; Model = "Z820 Workstation"; PackageID = "P01006F8" }
    @{ ProdCode = '2129'; Model = "Z840 Workstation"; PackageID = "P01006F9" }
    @{ ProdCode = '83e7'; Model = "EliteDesk 705 G4"; PackageID = "P01006FA" }
    @{ ProdCode = '1993'; Model = "ProBook 650 G1"; PackageID = "P0100701" }
    @{ ProdCode = '8617'; Model = "EliteDesk 705 G5 SFF"; PackageID = "P0100736" }
    @{ ProdCode = '22da'; Model = "EliteBook Folio 9480m"; PackageID = "P0100739" }
    @{ ProdCode = '8584'; Model = "EliteBook 745 G6"; PackageID = "P010073A" }
    @{ ProdCode = '8053'; Model = "EliteDesk 800 G2"; PackageID = "P010073E" }
    @{ ProdCode = '870D'; Model = "Elite x2 G8"; PackageID = "P01007B8" }
    @{ ProdCode = '856D'; Model = "ProBook 640 G5"; PackageID = "P01007C2" }
    @{ ProdCode = '860F'; Model = "ZBook 15 G6"; PackageID = "P0100833" } 
    @{ ProdCode = '886D'; Model = "ZBook 17 G8"; PackageID = "P01008F9" }
    @{ ProdCode = '872B'; Model = "EliteDesk 805 G6 SFF"; PackageID = "P0100834" }
    @{ ProdCode = '81c5'; Model = "ProBook Z4 G4"; PackageID = "P010084F" }
    @{ ProdCode = '894E'; Model = "Elite SFF 800 G9"; PackageID = "P01008A0" }
)

# Enable TLS 1.2 support for downloading modules from PSGallery (Required)
(New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Log -Message "Download HP Driver Pack for $OS $OSVER - extract and update to MECM Legacy Driver Package."

try {
    Write-Log -Message "Check $FileServer path"
    Test-Path -Path $FileServer
    $DPRootPackage = "$FileServer"
    $DownloadToServer = $true
    Write-Log -Message "Connected to Server $FileServer"
    if (Test-Path 'C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin') {
        Import-Module "C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1"
    }
}
Catch {
    Write-Log -Message "Not connected to $FileServer, exiting" -Severity 3
    $DownloadToServer = $false
}

#Create $DownloadDir
Write-Log -Message "Clean $DownloadDir if exist to start fresh and create if it does not exist"
if (Test-Path -Path $DownloadDir) {
    Write-Log -Message "$DownloadDir existed. Remove for fresh start"
    Remove-Item -Path $DownloadDir -Force -Recurse
    try {
        New-Item -Path $DownloadDir -ItemType Directory -Force
    }
    catch {
        Write-Log -Message "Failed to create the folder: $DownloadDir" -Severity 3
    }
}
elseif ((Test-Path -Path $DownloadDir) -eq $false) {
    Write-Log -Message "The folder $DownloadDir does not exist. Creating it."
    try {
        New-Item -Path $DownloadDir -ItemType Directory -Force
    }
    catch {
        Write-Log -Message "Failed to create the folder: $DownloadDir" -Severity 3
    }
}

foreach ($HPModel in $HPModelsTable) {
    Write-Log "Checking Model $($HPModel.Model) Product Code $($HPModel.ProdCode) for Driver Pack Updates"
    $SoftPaq = Get-SoftpaqList -Platform $HPModel.ProdCode -Os $OS -OsVer $OSVER -ErrorAction SilentlyContinue
    $DriverPack = $SoftPaq | Where-Object { $_.category -eq 'Manageability - Driver Pack' }
    $DriverPack = $DriverPack | Where-Object { $_.Name -notmatch "Windows PE" }
    $DriverPack = $DriverPack | Where-Object { $_.Name -notmatch "WinPE" }
    
    if ($DriverPack) {
        $DPDownloadPath = "$($DownloadDir)\$($HPModel.Model)\$($HPModel.ProdCode)\$($DriverPack.Version)"
        $ServerDPFullPath = "$($DPRootPackage)\$($HPModel.Model)"
        
        #Get Current Driver CMPackage Version from CM
        if ($DownloadToServer -eq $true) {
            Set-Location -Path "$($SiteCode):"
            $PackageInfo = Get-CMPackage -Id $HPModel.PackageID -Fast
            $PackageInfoVersion = $PackageInfo.Version
            Set-Location -Path "C:"
        }
        else {
            $PackageInfoVersion = $null
        }
        if ($PackageInfoVersion -eq $DriverPack.Version) {
            Write-Log -Message "$($HPModel.Model) already current: $PackageInfoVersion HP: $($DriverPack.Version))"
            $AlreadyCurrent = $True
        }
        else {
            Write-Log -Message "$($HPModel.Model) package is version $($PackageInfoVersion), new version is available $($DriverPack.Version)"
            if (!(Test-Path $DPDownloadPath)) {
                Write-Log -Message "$DPDownloadPath does not exist, create it"
                New-Item -Path $DPDownloadPath -ItemType Directory -Force
            }
            New-HPDriverPack -Platform $HPModel.ProdCode -Os $OS -OSVer $OSVER -Path $DPDownloadPath
            if (Test-Path $ServerDPFullPath) { 
                Write-Log -Message "$ServerDPFullPath exist, remove for clean driver"
                Remove-Item -Path $ServerDPFullPath -Recurse -Force 
            }
            Write-Log -Message "Create $ServerDPFullPath directory"
            New-Item $ServerDPFullPath -ItemType Directory -Force
            
            $CopyFromDir = (Get-ChildItem -Path ((Get-ChildItem -Path "$DPDownloadPath\*" -Directory).FullName) -Directory).FullName
            Write-Log -Message "Copy drivers to $ServerDPFullPath"
            
            Copy-Item $CopyFromDir -Destination $ServerDPFullPath -Force -Recurse
            Write-Log -Message "Export DriverPackInfo xml to $ServerDPFullPath"
            Export-Clixml -InputObject $DriverPack -Path "$($ServerDPFullPath)\DriverPackInfo.XML"

            $AlreadyCurrent = $false
        }

        if ($AlreadyCurrent -ne $true) {
            Write-Log -Message "Updating Package Info in ConfigMgr $($PackageInfo.Name) ID: $($HPModel.PackageID)"
            Import-Module 'C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1'
            Set-Location -Path "$($SiteCode):"
            Set-CMPackage -Id $HPModel.PackageID -path "$($ServerDPFullPath)"
            Set-CMPackage -Id $HPModel.PackageID -Version "$($DriverPack.Version)"
            Set-CMPackage -Id $HPModel.PackageID -Description "Version $($DriverPack.Version) Released $($DriverPack.ReleaseDate). Folder = $($DPDownloadPath)"
            Set-CMPackage -Id $HPModel.PackageID -Language $DriverPack.Id
            Set-CMPackage -ID $HPModel.PackageID -Manufacturer "HP"
            $PackageInfo = Get-CMPackage -Id $HPModel.PackageID -Fast
            Update-CMDistributionPoint -PackageId $HPModel.PackageID
            Set-Location -Path "C:"
            Write-Log -Message "Updated Package $($PackageInfo.Name), ID $($HPModel.PackageID) to $($DriverPack.Version) which was released $($DriverPack.ReleaseDate)"
        }
    }
    else {
        Write-Log -Message "No Driver Pack Available for $($HPModel.Model) Product Code $($HPModel.ProdCode) via Internet"
    }

}

$Script_End_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
$Script_Time_Taken = New-TimeSpan -Start $Script_Start_Time -End $Script_End_Time

Write-Log -Message "INFO: Script end: $Script_End_Time"
Write-Log -Message "INFO: Execution time: $Script_Time_Taken"
Write-Log -Message "***************************************************************************"
