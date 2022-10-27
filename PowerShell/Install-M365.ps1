[CmdletBinding(DefaultParameterSetName = 'XMLFile')]
param (
    [Parameter(ParameterSetName = 'XMLFile')]
    [string]$ConfigurationXMLFile,

    [Parameter(ParameterSetName = 'NoXML')]
    [ValidateSet('TRUE', 'FALSE')]$AcceptEULA = 'TRUE',

    [Parameter(ParameterSetName = 'NoXML')]
    [ValidateSet('SemiAnnual', 'SemiAnnualPreview', 'MonthlyEnterprise', 'Current')]$Channel = 'SemiAnnual',

    [Parameter(ParameterSetName = 'NoXML')]
    [Switch]$DisplayInstall = $True,

    [Parameter(ParameterSetName = 'NoXML')]
    [ValidateSet('Groove', 'Outlook', 'OneNote', 'Access', 'OneDrive', 'Publisher', 'Word', 'Excel', 'PowerPoint', 'Teams', 'Lync')]
    [Array]$ExcludeApps,

    [Parameter(ParameterSetName = 'NoXML')]
    [ValidateSet('64', '32')]$OfficeArch = '64',

    [Parameter(ParameterSetName = 'NoXML')]
    [ValidateSet('O365ProPlusRetail', 'O365BusinessRetail')]$OfficeEdition = 'O365ProPlusRetail',

    [Parameter(ParameterSetName = 'NoXML')]
    [ValidateSet(0, 1)]$SharedComputerLicensing = '0',

    [Parameter(ParameterSetName = 'NoXML')]
    [ValidateSet('TRUE', 'FALSE')]$EnableUpdates = 'FALSE',

    [Parameter(ParameterSetName = 'NoXML')]
    [String]$LoggingPath,

    [Parameter(ParameterSetName = 'NoXML')]
    [String]$SourcePath,

    [Parameter(ParameterSetName = 'NoXML')]
    [ValidateSet('TRUE', 'FALSE')]$PinItemsToTaskbar = 'TRUE',

    [Parameter(ParameterSetName = 'NoXML')]
    [Switch]$KeepMSI = $False,

    [String]$OfficeInstallDownloadPath = "$($env:windir)\ccmcache\OfficeInstall",
    [Switch]$CleanUpInstallFiles = $True
)

# Enable TLS 1.2 support for downloading modules from PSGallery (Required)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ModuleName = 'PSADT'

if (Get-Module -Name $ModuleName -ListAvailable) {
    Write-Log -Message "PSADT module found." -Source 'Install-PSADT' -ScriptSection 'Install-PSADT' -LogFileDirectory '$env:windir\Temp' -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$True
    Import-Module -Name $ModuleName -Force
} else {
    # Setup LOCALAPPDATA variable
    [System.Environment]::SetEnvironmentVariable('LOCALAPPDATA', "$env:SystemDrive\Windows\System32\config\systemprofile\AppData\Local")
    $WorkingDir = $env:TEMP
    Write-Host "Seup LOCALAPPDATA Variable" -ForegroundColor Cyan
    Write-Host "Temp Appdata: $($WorkingDir)" -ForegroundColor Cyan

    # Setup PowerShellGet from PSGallery Url
    $PSGPath = "$env:ProgramFiles\WindowsPowerShell\Modules\PowerShellGet\2.2.5"
    if (Test-Path $PSGPath) {
        Write-Host "PowerShellGet found." -ForegroundColor Cyan
    } else {
        if (!(Get-Module -Name PowerShellGet)) {
            Write-Host "PowerShellGet module not found. Start downloading and installing..." -ForegroundColor Cyan
            $PowerShellGetURL = "https://psg-prod-eastus.azureedge.net/packages/powershellget.2.2.5.nupkg"    
            Write-Host "PowerShellGet URL: $($PowerShellGetURL)" -ForegroundColor Cyan
            Invoke-WebRequest -UseBasicParsing -Uri $PowerShellGetURL -OutFile "$WorkingDir\powershellget.2.2.5.zip"
            $Null = New-Item -Path "$WorkingDir\2.2.5" -ItemType Directory -Force
            Expand-Archive -Path "$WorkingDir\powershellget.2.2.5.zip" -DestinationPath "$WorkingDir\2.2.5"
            $Null = New-Item -Path "$env:ProgramFiles\WindowsPowerShell\Modules\PowerShellGet" -ItemType Directory -ErrorAction SilentlyContinue
            Move-Item -Path "$WorkingDir\2.2.5" -Destination "$env:ProgramFiles\WindowsPowerShell\Modules\PowerShellGet\2.2.5"
            Remove-Item "$WorkingDir\powershellget.2.2.5.zip" -Recurse -Force
            Write-Host "PowerShellGet installed." -ForegroundColor Cyan
        }
    }

    # Setup PackageManagement from PSGallery Url
    $PkgMgtPath = "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement\1.4.7"
    if (Test-Path -$PkgMgtPath) {
        Write-Host "PackageManagement installed." -ForegroundColor Cyan
    } else {
        if (!(Get-Module -Name PackageManagement)) {
            Write-Host "PackageManagement not found. Start downloading and installing..." -ForegroundColor Cyan
            $PackageManagementURL = "https://psg-prod-eastus.azureedge.net/packages/packagemanagement.1.4.7.nupkg"
            Invoke-WebRequest -UseBasicParsing -Uri $PackageManagementURL -OutFile "$WorkingDir\packagemanagement.1.4.7.zip"
            $Null = New-Item -Path "$WorkingDir\1.4.7" -ItemType Directory -Force
            Expand-Archive -Path "$WorkingDir\packagemanagement.1.4.7.zip" -DestinationPath "$WorkingDir\1.4.7"
            $Null = New-Item -Path "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement" -ItemType Directory -ErrorAction SilentlyContinue
            Move-Item -Path "$WorkingDir\1.4.7" -Destination "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement\1.4.7"
            Remove-Item "$WorkingDir\packagemanagement.1.4.7.zip" -Recurse -Force
            Write-Host "PackageManagement installed." -ForegroundColor Cyan
    }
}

Import-Module PowerShellGet
try {
    Install-Module -Name $ModuleName -Force -AcceptLicense -SkipPublisherCheck
}
catch [System.Exception] {
    # Exception is stored in the automatic variable Install-Modul
    Write-Warning -Message "Unable to install $ModuleName. Error message: $($_.Exception.Message)"
    }
}

Import-Module -Name $ModuleName -Force

# Start downloading OffScrub files from GitHub for removal of Office Product
function Join-URL {
<#
    .DESCRIPTION
    Join-Path but for URL strings instead
     
    .PARAMETER Path
    Base path string
     
    .PARAMETER ChildPath
    Child path or item name
     
    .EXAMPLE
    Join-Url -Path "https://www.contoso.local" -ChildPath "foo.htm"
    returns "https://www.contoso.local/foo.htm"
#>
    param (
        [parameter(Mandatory = $True, HelpMessage = "Base Path")]
        [ValidateNotNullOrEmpty()]
        [string] $Path,
        [parameter(Mandatory = $True, HelpMessage = "Child Path or Item Name")]
        [ValidateNotNullOrEmpty()]
        [string] $ChildPath
    )
    if ($Path.EndsWith('/')) {
        return "$Path" + "$ChildPath"
    } else {
        return "$Path/$ChildPath"
    }
}

function Remove-MSOffice {
<#
    .DESCRIPTION
    Rip out Office products by the roots
 
    .PARAMETER ScriptSource
    Source URL to the MS Office Scrub scripts (github repo)
 
    .PARAMETER ForceDownload
    Download source scripts even if local copies exist
 
    .EXAMPLE
    Remove-MSOffice -Verbose
 
    .NOTES
    David Stein 08/15/2018
    Modified: The Wiz-09/24/2022
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, HelpMessage = "Source URL")]
        [string]$ScriptSource = "https://raw.githubusercontent.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/master/Office-ProPlus-Deployment/Remove-PreviousOfficeInstalls",
        [Parameter(Mandatory = $false, HelpMessage = "Force new download")]
        [ValidateSet('True','False')]$ForceDownload = 'Tre'
    )
    $Continue = $True
    $files = @("OffScrub03.vbs", "OffScrub07.vbs", "OffScrub10.vbs", "OffScrub_O15msi.vbs", "OffScrub_O16msi.vbs", "OffScrubc2r.vbs", "Remove-PreviousOfficeInstalls.ps1")
    Write-Log -Message "downloading source files from remote repository" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    foreach ($f in $files) {
        $remoteFile = Join-Url -Path $ScriptSource -ChildPath $f
        $localFile = Join-Path -Path $env:windir\ccmcache -ChildPath $f
        if (-not(Test-Path $localFile) -or $ForceDownload) {
            Write-Log -Message "downloading: $remoteFile" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
            try {
                $(New-Object System.Net.WebClient).DownloadFile($remoteFile, $localFile) | Out-Null
            }
            catch {
                Write-Warning $_.Exception.Message
                Write-Log -Message "Error downloading $f. Error message: $_.Exception.Message" -Severity 2 -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
            }
        }
        if (Test-Path $localFile) {
            Write-Log -Message "downloaded successfully to: $localFile" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
        }
        else {
            Write-Warning "error: failed to download"
            Write-Log -Message "Error: Failed to download" -Severity 3 -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
            $continue = $null
        }
    }
    if ($continue) {
        Write-Log -Message "finished downloading source files" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
        Write-Log -Message "saving current working location" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
        $cwd = Get-Location
        Write-Log -Message "changing to temp location" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
        Set-Location -Path "$env:windir\ccmcache"
        Write-Log -Message "invoking script: Remove-PreviousOfficeInstalls.ps1" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
        .\Remove-PreviousOfficeInstalls.ps1 -ProductsToRemove MainOfficeProduct -RemoveClickToRunVersions $true -Remove2016Installs $true -KeepUserSettings $false -KeepLync $false -Force $true -NoReboot $true -Quiet $true
        Write-Log -Message "restoring previous working location" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
        Set-Location -Path $cwd
        Write-Log -Message "finished" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    } else {
        Write-Log -Message "Failed to download source files, skipping execution" -Severity 2 -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    }
}

Remove-MSOffice

# Remove Office 2016 shortcut if existed
$O2016 = "$envCommonStartMenuPrograms\Microsoft Office 2016"
if (Test-Path $O2016) {
    Remove-Folder $O2016 -Verbose
}

function Set-XMLFile {

  if ($ExcludeApps) {
    $ExcludeApps | ForEach-Object {
      $ExcludeAppsString += "<ExcludeApp ID =`"$_`" />"
    }
  }

  if ($OfficeArch) {
    $OfficeArchString = "`"$OfficeArch`""
  }

  if ($KeepMSI) {
    $RemoveMSIString = $Null
  }
  else {
    $RemoveMSIString = '<RemoveMSI />'
  }

  if ($Channel) {
    $ChannelString = "Channel=`"$Channel`""
  }
  else {
    $ChannelString = $Null
  }

  if ($SourcePath) {
    $SourcePathString = "SourcePath=`"$SourcePath`"" 
  }
  else {
    $SourcePathString = $Null
  }

  if ($DisplayInstall) {
    $SilentInstallString = 'Full'
  }
  else {
    $SilentInstallString = 'None'
  }

  if ($LoggingPath) {
    $LoggingString = "<Logging Level=`"Standard`" Path=`"$LoggingPath`" />"
  }
  else {
    $LoggingString = $Null
  }

  $OfficeXML = [XML]@"
  <Configuration>
    <Add OfficeClientEdition=$OfficeArchString $ChannelString $SourcePathString OfficeMgmtCOM="TRUE">
      <Product ID="$OfficeEdition">
        <Language ID="en-us" />
        <ExcludeApp ID="Groove" />
        <ExcludeApp ID="OneDrive" />
        <ExcludeApp ID="Lync" />
        <ExcludeApp ID="Teams" />
        <ExcludeApp ID="Bing" />
      </Product>

    <Product ID="VisioProXVolume" MSICondition="VisPro,VisProR" PIDKEY="69WXN-MBYV6-22PQG-3WGHK-RM6XC">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="OneDrive" />
    </Product>

    <Product ID="ProjectProXVolume" MSICondition="PrjPro,PrjProR" PIDKEY="WGT24-HCNMF-FQ7XH-6M8K7-DRTW9">
      <Language ID="en-us" />
	<ExcludeApp ID="Groove"/>
    </Product>

    </Add>
    <Updates Enabled="$EnableUpdates" />  
    <Property Name="PinIconsToTaskbar" Value="$PinItemsToTaskbar" />
    <Property Name="SharedComputerLicensing" Value="$SharedComputerlicensing" />
    <Property Name="SCLCacheOverride" Value="0" />
    <Property Name="AUTOACTIVATE" Value="0" />
    <Property Name="DeviceBasedLicensing" Value="0" />
    $RemoveMSIString
    $LoggingString
    <AppSettings>
    <Setup Name="Company" Value="Houston Police Department " />
    <User Key="software\microsoft\office\16.0\common\internet" Name="donotuselongfilenames" Value="0" Type="REG_DWORD" App="office16" Id="L_Uselongfilenameswheneverpossible" />
    <User Key="software\microsoft\office\16.0\common\general" Name="shownfirstrunoptin" Value="1" Type="REG_DWORD" App="office16" Id="L_DisableOptinWizard" />
    <User Key="software\microsoft\office\16.0\common" Name="autoorgidgetkey" Value="1" Type="REG_DWORD" App="office16" Id="L_AutoOrgIDGetKey" />
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
  </AppSettings>
    <Display Level="$SilentInstallString" AcceptEULA="$AcceptEULA" />
  </Configuration>
"@

  $OfficeXML.Save("$OfficeInstallDownloadPath\OfficeInstall.xml")
  
    }

function Get-ODTURL {

  [String]$MSWebPage = Invoke-RestMethod 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117'

  #Thank you reddit user, u/sizzlr for this addition.
  $MSWebPage | ForEach-Object {
    if ($_ -match 'url=(https://.*officedeploymenttool.*\.exe)') {
      $matches[1]
    }
  }

}

$VerbosePreference = 'Continue'
$ErrorActionPreference = 'Stop'

$User = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (!($User.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
  Write-Warning 'Script is not running as Administrator'
  Write-Warning 'Please rerun this script as Administrator.'
  Write-Log -Message "Script is not running as Administrator. Please rerun this script as Administrator." -Severity 2 -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
  exit
}

if (Test-Path $OfficeInstallDownloadPath) {
    Write-Log -Message "Deleting $($OfficeInstallDownloadPath).... to create new fresh directory" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    Remove-Item -Path "$OfficeInstallDownloadPath" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
    New-Item -Path $OfficeInstallDownloadPath -ItemType Directory | Out-Null
    Write-Log -Message "New $OfficeInstallDownloadPath created." -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
}

if (-Not(Test-Path $OfficeInstallDownloadPath )) {
  New-Item -Path $OfficeInstallDownloadPath -ItemType Directory | Out-Null
  Write-Log -Message "$OfficeInstallDownloadPath created." -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
}

if (!($ConfigurationXMLFile)) {
  Set-XMLFile
  Write-Log -Message "Create xml file in $($OfficeInstallDownloadPath)." -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
}
else {
  if (!(Test-Path $ConfigurationXMLFile)) {
    Write-Warning 'The configuration XML file is not a valid file'
    Write-Warning 'Please check the path and try again'
    Write-Log -Message "The configuration XML file is not a valid file. Please check the path and try again" -Severity 3 -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    exit
  }
}

$ConfigurationXMLFile = "$OfficeInstallDownloadPath\OfficeInstall.xml"
$ODTInstallLink = Get-ODTURL
Write-Log -Message "Download Office Deployment Tool from $($ODTInstallLink)" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true

#Download the Office Deployment Tool
Write-Log -Message 'Downloading the Office Deployment Tool...' -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
try {
  Invoke-WebRequest -Uri $ODTInstallLink -OutFile "$OfficeInstallDownloadPath\ODTSetup.exe"
}
catch {
  Write-Warning 'There was an error downloading the Office Deployment Tool.'
  Write-Warning 'Please verify the below link is valid:'
  Write-Warning $ODTInstallLink
  Write-Log -Message "Please verify the below link is valid: $ODTInstallLink" -Severity 3 -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
  exit
}

# Run the Office Deployment Tool setup
try {
  Write-Log -Message "Running the Office Deployment Tool..." -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
  Start-Process "$OfficeInstallDownloadPath\ODTSetup.exe" -ArgumentList "/quiet /extract:$OfficeInstallDownloadPath" -Wait
}
catch {
  Write-Warning 'Error running the Office Deployment Tool. The error is below:'
  Write-Warning $_
  Write-Log -Message "Error running the Office Deployment Tool. The error is: $_ " -Severity 2 -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
}

# Create custom Ribbon folder
Write-Log -Message "Start creating Ribbon folder....." -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
$RibbonFolder = "$($env:windir)\ccmcache\Ribbon"

if (Test-Path $RibbonFolder) {
    Write-Log -Message "Deleting $($RibbonFolder).... for creating new files" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    Remove-Item -Path "$RibbonFolder" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
    New-Item -Path $RibbonFolder -ItemType Directory -Force | Out-Null
    Write-Log -Message 'New $RibbonFolder created.' -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
}

If (!(Test-Path $RibbonFolder)) {
    New-Item -Path $RibbonFolder -ItemType Directory -Force | Out-Null
    Write-Log -Message "$($RibbonFolder) created successfully" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
}

# Create custom $XcelInput Office UI
New-Item -Path "$RibbonFolder" -ItemType File -Name "Excel.officeUI"
[xml](Add-Content -Path "$RibbonFolder\Excel.officeUI" -Value '<mso:customUI xmlns:mso="http://schemas.microsoft.com/office/2009/07/customui"><mso:ribbon><mso:qat/><mso:tabs><mso:tab idQ="mso:TabHome"><mso:group id="mso_c1.3AD2682" label="Save" insertBeforeQ="mso:GroupClipboard" autoScale="true"><mso:control idQ="mso:FileSave" visible="true"/><mso:control idQ="mso:FileSaveAs" visible="true"/></mso:group></mso:tab><mso:tab idQ="mso:TabDrawInk" visible="false"/></mso:tabs></mso:ribbon></mso:customUI>
')

# Create custom $WordInput Office UI
New-Item -Path $RibbonFolder -ItemType File -Name "Word.officeUI"
[xml](Add-Content -Path "$RibbonFolder\Word.officeUI" -Value '<mso:customUI xmlns:mso="http://schemas.microsoft.com/office/2009/07/customui"><mso:ribbon><mso:qat><mso:sharedControls><mso:control idQ="mso:AutoSaveSwitch" visible="false"/><mso:control idQ="mso:FileNewDefault" visible="false"/><mso:control idQ="mso:FileOpenUsingBackstage" visible="false"/><mso:control idQ="mso:FileSave" visible="false"/><mso:control idQ="mso:FileSendAsAttachment" visible="false"/><mso:control idQ="mso:FilePrintQuick" visible="false"/><mso:control idQ="mso:PrintPreviewAndPrint" visible="false"/><mso:control idQ="mso:WritingAssistanceCheckDocument" visible="false"/><mso:control idQ="mso:ReadAloud" visible="false"/><mso:control idQ="mso:Undo" visible="true"/><mso:control idQ="mso:RedoOrRepeat" visible="true"/><mso:control idQ="mso:TableDrawTable" visible="false"/><mso:control idQ="mso:PointerModeOptions" visible="false"/></mso:sharedControls></mso:qat><mso:tabs><mso:tab idQ="mso:TabHome"><mso:group id="mso_c1.39687EB" label="Save" insertBeforeQ="mso:GroupClipboard" autoScale="true"><mso:control idQ="mso:FileSave" visible="true"/><mso:control idQ="mso:FileSaveAs" visible="true"/></mso:group></mso:tab><mso:tab idQ="mso:TabDrawInk" visible="false"/></mso:tabs></mso:ribbon></mso:customUI>
')

# Create custom $PwrPtInput Office UI
New-Item -Path $RibbonFolder -ItemType File -Name "PowerPoint.officeUI"
[xml](Add-Content -Path "$RibbonFolder\PowerPoint.officeUI" -Value '<mso:customUI xmlns:mso="http://schemas.microsoft.com/office/2009/07/customui"><mso:ribbon><mso:qat/><mso:tabs><mso:tab idQ="mso:TabHome"><mso:group id="mso_c1.3AC0EF7" label="Save" insertBeforeQ="mso:GroupClipboard" autoScale="true"><mso:control idQ="mso:FileSave" visible="true"/><mso:control idQ="mso:FileSaveAs" visible="true"/></mso:group></mso:tab><mso:tab idQ="mso:TabDrawInk" visible="false"/><mso:tab idQ="mso:TabRecording" visible="false"/></mso:tabs></mso:ribbon></mso:customUI>
')

#Run the O365 install
try {
    Write-Log -Message "Downloading and installing Microsoft 365 Apps for enterprise - en-us..." -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    Start-Process "$OfficeInstallDownloadPath\Setup.exe" -ArgumentList "/configure $ConfigurationXMLFile" -Wait -PassThru -NoNewWindow -Verbose -ErrorAction SilentlyContinue *>&1 | Out-String | Write-Log -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType CMTrace -WriteHost:$True

    # Add Custom Office UI
    $DefLocalFolder = "C:\Users\Default\AppData\Local\Microsoft\Office"
    if (-NOT(Test-Path $DefLocalFolder)) {
        Write-Log -Message "$($DefLocalFolder) does not exist. Start creating" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
        New-Folder -Path $DefLocalFolder
        if (Test-Path $DefLocalFolder) {
            Write-Log -Message "$($DefLocalFolder) created." -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
        }
    }

    $DefRoamingFolder = "C:\Users\Default\AppData\Roaming\Microsoft\Office"
    if (!(Test-Path $DefRoamingFolder)) {
        Write-Log -Message "$($DefRoamingFolder) does not exist. Start creating" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
        New-Folder -Path $DefRoamingFolder
        if (Test-Path $DefRoamingFolder) {
            Write-Log -Message "$($DefRoamingFolder) created." -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
        }
    }

    Write-Log -Message "Copy Custom UI file to $($DefLocalFolder)" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    Copy-File -Path "$RibbonFolder\*.officeUI" -Destination "$DefLocalFolder\"

    Write-Log -Message "Copy Custom UI file to $($DefRoamingFolder)" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    Copy-File -Path "$RibbonFolder\*" -Destination "$DefRoamingFolder\"

    if (Test-Path "$DefLocalFolder\*.officeUI" -PathType Leaf) {
        Write-Log -Message "Office Custom UI files copied" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    }
    else {
        Write-Log -Message "Office Custom UI Files not found" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    }
}

catch {
    Write-Log -Message 'Error running the Office install. The error is below:' -Severity 2 -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    Write-Log -Message "$_" -Severity 2 -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    Write-Log -Message "Error running the Office install. The error is: $_ " -Severity 2 -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
}

# Using PSADT to copy customUI to all user profiles
$ProfilePaths = Get-UserProfiles -ExcludeDefaultUser:$true -ExcludeNTAccount "Administrator","SiteAdmin","Public" | Select-Object -ExpandProperty 'ProfilePath'
foreach ($Profile in $ProfilePaths) {
    Copy-File -Path "$RibbonFolder\*" -Destination "$Profile\AppData\Local\Microsoft\Office\"
    Copy-File -Path "$RibbonFolder\*" -Destination "$Profile\AppData\Roaming\Microsoft\Office\"
}

#Change the SigninOptions to 2 if availble( for activation purpose)       
[scriptblock]$HKCURegistrySettings = {
    Set-RegistryKey -Key 'HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\signin' -Name signinoptions -Value 2 -SID $UserProfile.SID -ContinueOnError:$true
}
Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings

#Check if Office 365 suite was installed correctly.
$RegLocations = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)

$OfficeInstalled = $False
foreach ($Key in (Get-ChildItem $RegLocations) ) {
    if ($Key.GetValue('DisplayName') -like '*Microsoft 365*') {
        $OfficeVersionInstalled = $Key.GetValue('DisplayName')
        $OfficeInstalled = $True
    }
}

if ($OfficeInstalled) {
    Write-Log -Message "$($OfficeVersionInstalled) existed. Start Menu Layout process!" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    # Create custom xml folder
    $xmlDir = "$($env:windir)\ccmcache\xmlDir"

    If (!(Test-Path $xmlDir)) {
        New-Folder -Path $xmlDir
        Write-Log -Message "$($xmlDir) created successfully" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    }

# Create Start Menu Layout xml
Write-Log -Message "Create Start Menu Layout xml for Desktop devices..." -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
[xml]$doc = New-Object System.Xml.XmlDocument
$dec = $doc.CreateXmlDeclaration("1.0","UTF-8",$null)
$doc.AppendChild($dec)
$txtFragment = @'
<LayoutModificationTemplate
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"
    xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout"
    xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout"
    xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"
    Version="1">	
  <LayoutOptions StartTileGroupCellWidth="6" />
  <DefaultLayoutOverride>
    <StartLayoutCollection>
      <defaultlayout:StartLayout GroupCellWidth="6">
        <start:Group Name="Office 365">
          <start:DesktopApplicationTile Size="2x2" Column="2" Row="0" DesktopApplicationID="Microsoft.Office.POWERPNT.EXE.15" />
          <start:DesktopApplicationTile Size="2x2" Column="4" Row="0" DesktopApplicationID="Microsoft.Office.EXCEL.EXE.15" />
          <start:DesktopApplicationTile Size="2x2" Column="0" Row="2" DesktopApplicationID="Microsoft.Office.OUTLOOK.EXE.15" />
          <start:DesktopApplicationTile Size="2x2" Column="4" Row="2" DesktopApplicationID="Microsoft.Office.MSACCESS.EXE.15" />
          <start:DesktopApplicationTile Size="2x2" Column="2" Row="2" DesktopApplicationID="Microsoft.Office.ONENOTE.EXE.15" />
          <start:DesktopApplicationTile Size="2x2" Column="0" Row="0" DesktopApplicationID="Microsoft.Office.WINWORD.EXE.15" />
        </start:Group>
        <start:Group Name="Software">
          <start:DesktopApplicationTile Size="2x2" Column="2" Row="0" DesktopApplicationID="Microsoft.Windows.Explorer" />
          <start:DesktopApplicationTile Size="2x2" Column="0" Row="0" DesktopApplicationID="Microsoft.SoftwareCenter.DesktopToasts" />
          <start:DesktopApplicationTile Size="2x2" Column="4" Row="0" DesktopApplicationID="Microsoft.Windows.ControlPanel" />
        </start:Group>
        <start:Group Name="Tools">
          <start:Tile Size="2x2" Column="2" Row="0" AppUserModelID="Microsoft.WindowsSoundRecorder_8wekyb3d8bbwe!App" />
          <start:Tile Size="2x2" Column="4" Row="0" AppUserModelID="Microsoft.WindowsAlarms_8wekyb3d8bbwe!App" />
          <start:Tile Size="2x2" Column="0" Row="0" AppUserModelID="Microsoft.WindowsCalculator_8wekyb3d8bbwe!App" />
        </start:Group>
        <start:Group Name="Browsers">
          <start:DesktopApplicationTile Size="2x2" Column="0" Row="0" DesktopApplicationID="MSEdge" />
          <start:DesktopApplicationTile Size="2x2" Column="2" Row="0" DesktopApplicationID="Chrome" />
        </start:Group>
        <start:Group Name="Settings">
          <start:Tile Size="2x2" Column="2" Row="0" AppUserModelID="windows.immersivecontrolpanel_cw5n1h2txyewy!microsoft.windows.immersivecontrolpanel" />
          <start:DesktopApplicationTile Size="2x2" Column="0" Row="0" DesktopApplicationID="Microsoft.Windows.Computer" />
          <start:DesktopApplicationTile Size="2x2" Column="4" Row="0" DesktopApplicationID="Microsoft.AutoGenerated.{923DD477-5846-686B-A659-0FCCD73851A8}" />
        </start:Group>
      </defaultlayout:StartLayout>
    </StartLayoutCollection>
  </DefaultLayoutOverride>
	<CustomTaskbarLayoutCollection PinListPlacement="Replace">
    <defaultlayout:TaskbarLayout>
      <taskbar:TaskbarPinList>
        <taskbar:UWA AppUserModelID="Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="%APPDATA%\Microsoft\Windows\Start Menu\Programs\System Tools\File Explorer.lnk" />
		    <taskbar:UWA AppUserModelID="Microsoft.ScreenSketch_8wekyb3d8bbwe!App" />
		    <taskbar:UWA AppUserModelID="Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App" />
		    <taskbar:DesktopApp DesktopApplicationLinkPath="%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Outlook.lnk" />
		    <taskbar:DesktopApp DesktopApplicationLinkPath="%APPDATA%\Microsoft\Windows\Start Menu\Programs\System Tools\Control Panel.lnk" />
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>
'@

$docFrag = $doc.CreateDocumentFragment()
$docFrag.InnerXml = $txtFragment
$doc.AppendChild($docFrag)
$doc.Save("$xmlDir\Layout.xml")

# Copy Start Menu Layout for all users
Write-Log -Message "Installing StartMenu Layout for HPD Desktop Devices" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
Copy-File -Path "$xmlDir\Layout.xml" -Destination "C:\Users\Default\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml"

# Import Start Menu layout
if (Test-Path "$xmlDir\Layout.xml") {
    Write-Log -Message "Layout xml found. Start importing Start Menu Layout..." -Source 'Start-MenuLayout' -ScriptSection 'Start-MenuLayout' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType CMTrace -WriteHost:$true
    Import-StartLayout -LayoutPath "$xmlDir\Layout.xml" -MountPath "C:\" -ErrorAction SilentlyContinue
    Write-Log -Message "Completed importing Start Menu Layout!" -Source 'Start-MenuLayout' -ScriptSection 'Start-MenuLayout' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType CMTrace -WriteHost:$True
}

# Apply LayoutModification xml to all users
$ProfilePaths = Get-UserProfiles -ExcludeDefaultUser:$true | Select-Object -ExpandProperty 'ProfilePath'
Foreach ($Profile in $ProfilePaths) {
    Copy-File -Path "C:\Users\Default\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml" -Destination "$Profile\AppData\Local\Microsoft\Windows\Shell\"
}
Write-Log -Message "Applied LayoutModification.xml" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true

# Remove Store Key from all users to apply new Start Menu Layout (Using these lines since PSADT module does not work as expected!)
$PatternSID = 'S-1-5-21-\d+-\d+\-\d+\-\d+$'
        # Get Username, SID, and location of ntuser.dat for all users
        $ProfileList = gp 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' | Where-Object {$_.PSChildName -match $PatternSID} | 
         Select  @{name="SID";expression={$_.PSChildName}}, 
            @{name="UserHive";expression={"$($_.ProfileImagePath)\ntuser.dat"}}, 
            @{name="Username";expression={$_.ProfileImagePath -replace '^(.*[\\\/])', ''}}
 
        # Get all user SIDs found in HKEY_USERS (ntuder.dat files that are loaded)
        $LoadedHives = gci Registry::HKEY_USERS | ? {$_.PSChildname -match $PatternSID} | Select @{name="SID";expression={$_.PSChildName}} -ErrorAction SilentlyContinue
        $UnloadedHives = Compare-Object $ProfileList.SID $LoadedHives.SID -ErrorAction SilentlyContinue | Select @{name="SID";expression={$_.InputObject}}, UserHive, Username -ErrorAction SilentlyContinue

        Foreach ($item in $ProfileList) {
        # Load User ntuser.dat if it's not already loaded
             IF ($item.SID -in $UnloadedHives.SID) {
             reg load HKU\$($Item.SID) $($Item.UserHive)  | Out-Null
	         if ($ProfileList.Count -le 20) { Start-Sleep -Seconds 2}	
             }
 
         #####################################################################
         # This is where you can read/modify a users portion of the registry 
         $key = "Registry::HKEY_USERS\$($Item.SID)\SOFTWARE\Microsoft\Windows\CurrentVersion\CloudStore\Store" 
         Remove-Item -Path $key -Force -Recurse -ErrorAction SilentlyContinue | Out-Null

         #####################################################################
         # Unload ntuser.dat        
             IF ($item.SID -in $UnloadedHives.SID) {
             ### Garbage collection and closing of ntuser.dat ###
             [gc]::Collect() 
	         if ($ProfileList.Count -le 20) { Start-Sleep -Seconds 1}	
             reg unload HKU\$($Item.SID) | Out-Null
             }
        }

# Cleanup xmlDir Directory
Remove-Folder -Path $xmlDir
Write-Log -Message "$($xmlDir) cleaned up." -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
} else {
Write-Log -Message "Microsoft 365 was not detected. Exit script." -Severity 3 -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
exit 0
}

if ($OfficeInstalled) {
    Write-Log -Message "$($OfficeVersionInstalled) installed successfully!" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
} else {
    Write-Warning 'Microsoft 365 was not detected after the install ran'
    Write-Log -Message "Microsoft 365 was not detected after the install ran" -Severity 2 -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
}

if ($CleanUpInstallFiles) {
    Remove-Folder -Path $OfficeInstallDownloadPath
    Remove-Folder -Path $RibbonFolder
    Write-Log -Message "$OfficeInstallDownloadPath and $RibbonFolder removed." -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
}

# Remove uninstall files
"$envwindir\ccmcache\OffScrub03.vbs",
"$envwindir\ccmcache\OffScrub07.vbs",
"$envwindir\ccmcache\OffScrub10.vbs",
"$envwindir\ccmcache\OffScrub_O15msi.vbs",
"$envwindir\ccmcache\OffScrub_O16msi.vbs",
"$envwindir\ccmcache\OffScrubc2r.vbs",
"$envwindir\ccmcache\Remove-PreviousOfficeInstalls.ps1"`
| ForEach-Object { Remove-File -Path "$_" }

Show-InstallationRestartPrompt -CountdownSeconds 300 -CountdownNoHideSeconds 60
