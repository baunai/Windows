[CmdletBinding(DefaultParameterSetName = 'XMLFile')]
  param(
  [Parameter(ParameterSetName = 'XMLFile')]
  [String]$ConfigurationXMLFile,

  [Parameter(ParameterSetName = 'NoXML')]
  [ValidateSet('TRUE', 'FALSE')]$AcceptEULA = 'TRUE',

  [Parameter(ParameterSetName = 'NoXML')]
  [ValidateSet('SemiAnnual', 'SemiAnnualPreview', 'MonthlyEnterprise', 'Current')]$Channel = 'SemiAnnual',

  [Parameter(ParameterSetName = 'NoXML')]
  [Switch]$DisplayInstall = $False,

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

  [String]$OfficeInstallDownloadPath = "$($env:windir)\Temp\OfficeInstall",
  [Switch]$CleanUpInstallFiles = $True
)

# Enable TLS 1.2 support for downloading modules from PSGallery (Required)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ModuleName = 'PSADT'

#Setup LOCALAPPDATA Variable
[System.Environment]::SetEnvironmentVariable('LOCALAPPDATA', "$env:SystemDrive\Windows\system32\config\systemprofile\AppData\Local")

$WorkingDir = $env:TEMP

Write-Log -Message "Setup LOCALAPPDATA Variable" -Source 'Set-LocalAppData' -LogType 'CMTrace'
Write-Log -Message "Temp Appdata: $($WorkingDir)" -Source 'Set-LocalAppData' -LogType 'CMTrace'

#PowerShellGet from PSGallery URL
$PSGPath = "$env:ProgramFiles\WindowsPowerShell\Modules\PowerShellGet\2.2.5"
if (Test-Path $PSGPath) {
  Write-Log -Message "PowerShellGet installed." -Source 'Install-PowerShellGet' -ScriptSection 'Install-PowerShellGet' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
}
else {
  if (!(Get-Module -Name PowerShellGet)) {   
    Write-Log -Message "PowerShellGet module not found. Start downloading and installing..." -Source 'Install-PowerShellGet' -ScriptSection 'Install-PowerShellGet' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    $PowerShellGetURL = "https://psg-prod-eastus.azureedge.net/packages/powershellget.2.2.5.nupkg"    
    Write-Log -Message "PowerShellGet URL: $($PowerShellGetURL)" -Source 'Install-PowerShellGet' -ScriptSection 'Install-PowerShellGet' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    Invoke-WebRequest -UseBasicParsing -Uri $PowerShellGetURL -OutFile "$WorkingDir\powershellget.2.2.5.zip"
    $Null = New-Item -Path "$WorkingDir\2.2.5" -ItemType Directory -Force
    Expand-Archive -Path "$WorkingDir\powershellget.2.2.5.zip" -DestinationPath "$WorkingDir\2.2.5"
    $Null = New-Item -Path "$env:ProgramFiles\WindowsPowerShell\Modules\PowerShellGet" -ItemType Directory -ErrorAction SilentlyContinue
    Move-Item -Path "$WorkingDir\2.2.5" -Destination "$env:ProgramFiles\WindowsPowerShell\Modules\PowerShellGet\2.2.5"
    Remove-Item "$WorkingDir\powershellget.2.2.5.zip" -Recurse -Force
    Write-Log -Message "PowerShellGet installed." -Source 'Install-PowerShellGet' -ScriptSection 'Install-PowerShellGet' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
  }
}

#PackageManagement from PSGallery URL
$PkgMgtPath = "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement\1.4.7"
if (Test-Path $PkgMgtPath) {
  Write-Log -Message "PackageManagement installed." -Source 'Install-PkgMgmt' -ScriptSection 'Install-PkgMgmt' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
}
else {
  if (!(Get-Module -Name PackageManagement)) {
    Write-Log -Message "PackageManagement not found. Start downloading and installing..." -Source 'Install-PkgMgmt' -ScriptSection 'Install-PkgMgmt' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
    $PackageManagementURL = "https://psg-prod-eastus.azureedge.net/packages/packagemanagement.1.4.7.nupkg"
    Invoke-WebRequest -UseBasicParsing -Uri $PackageManagementURL -OutFile "$WorkingDir\packagemanagement.1.4.7.zip"
    $Null = New-Item -Path "$WorkingDir\1.4.7" -ItemType Directory -Force
    Expand-Archive -Path "$WorkingDir\packagemanagement.1.4.7.zip" -DestinationPath "$WorkingDir\1.4.7"
    $Null = New-Item -Path "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement" -ItemType Directory -ErrorAction SilentlyContinue
    Move-Item -Path "$WorkingDir\1.4.7" -Destination "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement\1.4.7"
    Remove-Item "$WorkingDir\packagemanagement.1.4.7.zip" -Recurse -Force
    Write-Log -Message "PackageManagement installed." -Source 'Install-PkgMgmt' -ScriptSection 'Install-PkgMgmt' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
  }
}

#Import PowerShellGet
Import-Module PowerShellGet
Write-Log -Message "PowerShellGet module imported." -Source 'Install-PowerShellGet' -ScriptSection 'Install-PowerShellGet' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true

#Install Module from PSGallery
Install-Module -Name $ModuleName -Force -AcceptLicense -SkipPublisherCheck
try {
  if (Get-Module -Name $ModuleName -ListAvailable) {
    Write-Log -Message "$ModuleName module installed." -Source 'Install-PSADT' -ScriptSection 'Install-PSADT' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
  }
}
catch [System.Exception] {
  # Exception is stored in the automatic variable _
  Write-Log -Message "Unable to install $ModuleName. Error message: $($_.Exception.Message)" -Severity 3 -Source 'Install-PowerShellGet' -ScriptSection 'Install-PowerShellGet' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
}

Import-Module -Name $ModuleName -Force

# Initial logging
Write-Log -Message "Start creating xml and install Office 365 for HPD Mobility Devices" -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true

$splat = @("MCD", "HPDL", "HPDT")
if ($env:COMPUTERNAME -match ($splat -join '|')) {
  Write-Log -Message "HPD mobile devices detected. Start the installation of M365 for HPD Mobile Device." -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
}
else {
  Write-Log -Message "HPD mobile device not FOUND. Exit Script" -Severity 2 -Source 'Install-M365' -ScriptSection 'Install-M365' -LogFileDirectory "$env:windir\Temp" -LogFileName 'Install-M365.log' -LogType 'CMTrace' -WriteHost:$true
  exit 0
}

Get-Process | Where-Object { $_.Name -eq "Teams" } | Stop-Process -Force -Verbose

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
    </Add>  
    <Property Name="PinIconsToTaskbar" Value="$PinItemsToTaskbar" />
    <Property Name="SharedComputerLicensing" Value="$SharedComputerlicensing" />
    <Property Name="SCLCacheOverride" Value="0" />
    <Property Name="AUTOACTIVATE" Value="0" />
    <Property Name="DeviceBasedLicensing" Value="0" />
    <Display Level="$SilentInstallString" AcceptEULA="$AcceptEULA" />
    <Updates Enabled="$EnableUpdates" />
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
  </Configuration>
"@

  $OfficeXML.Save("$OfficeInstallDownloadPath\OfficeInstall.xml")
  
}


# Enable TLS 1.2 support for downloading modules from PSGallery (Required)
(New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


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
  Write-Log -Level ERROR -Message "Script is not running as Administrator. Please rerun this script as Administrator."
  exit
}


if (Test-Path $OfficeInstallDownloadPath) {
    Write-Verbose "Deleting $($OfficeInstallDownloadPath).... to create new fresh directory"
    Remove-Item -Path "$OfficeInstallDownloadPath" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
}

if (-Not(Test-Path $OfficeInstallDownloadPath )) {
  New-Item -Path $OfficeInstallDownloadPath -ItemType Directory | Out-Null
}

if (!($ConfigurationXMLFile)) {
  Set-XMLFile
  Write-Log -Message "Create xml file in $($OfficeInstallDownloadPath)."
}
else {
  if (!(Test-Path $ConfigurationXMLFile)) {
    Write-Warning 'The configuration XML file is not a valid file'
    Write-Warning 'Please check the path and try again'
    Write-Log -Level WARN -Message "The configuration XML file is not a valid file. Please check the path and try again"
    exit
  }
}

$ConfigurationXMLFile = "$OfficeInstallDownloadPath\OfficeInstall.xml"
$ODTInstallLink = Get-ODTURL
Write-Log -Message "Download Office Deployment Tool from $($ODTInstallLink)"

#Download the Office Deployment Tool
Write-Verbose 'Downloading the Office Deployment Tool...'
try {
  Invoke-WebRequest -Uri $ODTInstallLink -OutFile "$OfficeInstallDownloadPath\ODTSetup.exe"
}
catch {
  Write-Warning 'There was an error downloading the Office Deployment Tool.'
  Write-Warning 'Please verify the below link is valid:'
  Write-Warning $ODTInstallLink
  Write-Log -Message "Please verify the below link is valid: $ODTInstallLink"
  exit
}

#Run the Office Deployment Tool setup
try {
  Write-Verbose 'Running the Office Deployment Tool...'
  Write-Log -Message "Running the Office Deployment Tool..."
  Start-Process "$OfficeInstallDownloadPath\ODTSetup.exe" -ArgumentList "/quiet /extract:$OfficeInstallDownloadPath" -Wait
}
catch {
  Write-Warning 'Error running the Office Deployment Tool. The error is below:'
  Write-Warning $_
  Write-Log -Level ERROR -Message "Error running the Office Deployment Tool. The error is: $_ "
}

# Create custom Ribbon folder
$RibbonFolder = "$($env:windir)\Temp\Ribbon"

if (Test-Path $RibbonFolder) {
    Write-Verbose "Deleting $($RibbonFolder).... for creating new files"
    Remove-Item -Path "$RibbonFolder" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
}

If (!(Test-Path $RibbonFolder)) {
New-Item -Path $RibbonFolder -ItemType Directory -Force | Out-Null
Write-Verbose "$($RibbonFolder) created"
Write-Log -Message "$($RibbonFolder) created successfully"
}

# Create custom Excel Office UI
New-Item -Path "$RibbonFolder" -ItemType File -Name "Excel.officeUI"
$XcelInput = [xml](Add-Content -Path "$RibbonFolder\Excel.officeUI" -Value '<mso:customUI xmlns:mso="http://schemas.microsoft.com/office/2009/07/customui"><mso:ribbon><mso:qat/><mso:tabs><mso:tab idQ="mso:TabHome"><mso:group id="mso_c1.3AD2682" label="Save" insertBeforeQ="mso:GroupClipboard" autoScale="true"><mso:control idQ="mso:FileSave" visible="true"/><mso:control idQ="mso:FileSaveAs" visible="true"/></mso:group></mso:tab><mso:tab idQ="mso:TabDrawInk" visible="false"/></mso:tabs></mso:ribbon></mso:customUI>
')

# Create custom Word Office UI
New-Item -Path $RibbonFolder -ItemType File -Name "Word.officeUI"
$WordInput = [xml](Add-Content -Path "$RibbonFolder\Word.officeUI" -Value '<mso:customUI xmlns:mso="http://schemas.microsoft.com/office/2009/07/customui"><mso:ribbon><mso:qat><mso:sharedControls><mso:control idQ="mso:AutoSaveSwitch" visible="false"/><mso:control idQ="mso:FileNewDefault" visible="false"/><mso:control idQ="mso:FileOpenUsingBackstage" visible="false"/><mso:control idQ="mso:FileSave" visible="false"/><mso:control idQ="mso:FileSendAsAttachment" visible="false"/><mso:control idQ="mso:FilePrintQuick" visible="false"/><mso:control idQ="mso:PrintPreviewAndPrint" visible="false"/><mso:control idQ="mso:WritingAssistanceCheckDocument" visible="false"/><mso:control idQ="mso:ReadAloud" visible="false"/><mso:control idQ="mso:Undo" visible="true"/><mso:control idQ="mso:RedoOrRepeat" visible="true"/><mso:control idQ="mso:TableDrawTable" visible="false"/><mso:control idQ="mso:PointerModeOptions" visible="false"/></mso:sharedControls></mso:qat><mso:tabs><mso:tab idQ="mso:TabHome"><mso:group id="mso_c1.39687EB" label="Save" insertBeforeQ="mso:GroupClipboard" autoScale="true"><mso:control idQ="mso:FileSave" visible="true"/><mso:control idQ="mso:FileSaveAs" visible="true"/></mso:group></mso:tab><mso:tab idQ="mso:TabDrawInk" visible="false"/></mso:tabs></mso:ribbon></mso:customUI>
')

# Create custom PowerPoint Office UI
New-Item -Path $RibbonFolder -ItemType File -Name "PowerPoint.officeUI"
$PwPointInput = [xml](Add-Content -Path "$RibbonFolder\PowerPoint.officeUI" -Value '<mso:customUI xmlns:mso="http://schemas.microsoft.com/office/2009/07/customui"><mso:ribbon><mso:qat/><mso:tabs><mso:tab idQ="mso:TabHome"><mso:group id="mso_c1.3AC0EF7" label="Save" insertBeforeQ="mso:GroupClipboard" autoScale="true"><mso:control idQ="mso:FileSave" visible="true"/><mso:control idQ="mso:FileSaveAs" visible="true"/></mso:group></mso:tab><mso:tab idQ="mso:TabDrawInk" visible="false"/><mso:tab idQ="mso:TabRecording" visible="false"/></mso:tabs></mso:ribbon></mso:customUI>
')

# Create Start Menu Layout xml
New-Item -Path $RibbonFolder -ItemType File -Name "StartMenu-O365.xml"
$StartMenuInput = [xml](Add-Content -Path "$RibbonFolder\StartMenu-O365.xml" -Value '<?xml version="1.0" encoding="utf-8"?>
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
          <start:DesktopApplicationTile Size="2x2" Column="4" Row="0" DesktopApplicationID="Microsoft.InternetExplorer.Default" />
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
		<taskbar:DesktopApp DesktopApplicationLinkPath="%APPDATA%\Microsoft\Windows\Start Menu\Programs\Accessories\Internet Explorer.lnk" />
		<taskbar:UWA AppUserModelID="Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App" />
        <taskbar:DesktopApp DesktopApplicationLinkPath="%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Outlook.lnk" />
		<taskbar:DesktopApp DesktopApplicationLinkPath="%APPDATA%\Microsoft\Windows\Start Menu\Programs\System Tools\Control Panel.lnk" />
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>')


#Run the O365 install
try {
  Write-Verbose 'Downloading and installing Microsoft 365 Apps for enterprise - en-us'
  Write-Log -Message "Downloading and installing Microsoft 365 Apps for enterprise - en-us..."
  $Silent = Start-Process "$OfficeInstallDownloadPath\Setup.exe" -ArgumentList "/configure $ConfigurationXMLFile" -Wait -WindowStyle Hidden -PassThru
 
  # Add Custom Office UI
  $DefLocalFolder = "C:\Users\Default\AppData\Local\Microsoft\Office"
  if (-NOT(Test-Path $DefLocalFolder)) {
    Write-Verbose "$($DefLocalFolder) does not exist. Start creating"
    Write-Log -Message "$($DefLocalFolder) does not exist. Start creating"
    New-Item -Path $DefLocalFolder -ItemType Directory -Force | Out-Null
    if (Test-Path $DefLocalFolder){
    Write-Verbose "$($DefLocalFolder) created."
    Write-Log -Message "$($DefLocalFolder) created."
    Write-Log
    }
  }

  $DefRoamingFolder = "C:\Users\Default\AppData\Roaming\Microsoft\Office"
  if (!(Test-Path $DefRoamingFolder)) {
    Write-Verbose "$($DefRoamingFolder) does not exist. Start creating"
    Write-Log -Message "$($DefRoamingFolder) does not exist. Start creating"
    New-Item -Path $DefRoamingFolder -ItemType Directory -Force | Out-Null
    if (Test-Path $DefRoamingFolder) {
    Write-Verbose "$($DefRoamingFolder) created."
    Write-Log -Message "$($DefRoamingFolder) created."
    Write-Log
    }
  }

   Write-Verbose "Copy Custom UI file to $($DefLocalFolder)"
   Write-Log -Message "Copy Custom UI file to $($DefLocalFolder)"
   Copy-Item -Path "$RibbonFolder\*.officeUI" -Destination "$DefLocalFolder\" -Force -Recurse
    
   Write-Verbose "Copy Custom UI file to $($DefRoamingFolder)"
   Write-Log -Message "Copy Custom UI file to $($DefRoamingFolder)"
   Copy-Item -Path "$RibbonFolder\*" -Destination "$DefRoamingFolder\" -Force -Recurse

   if (Test-Path "$DefLocalFolder\*.officeUI" -PathType Leaf) {
    Write-Verbose "Custom UI files copied"
    Write-Log -Message "Office Custom UI files copied"
   } else {
    Write-Verbose "Custom UI Files not found"
    Write-Log -Message "Office Custom UI Files not found"
   }
}

catch {
  Write-Warning 'Error running the Office install. The error is below:'
  Write-Warning $_
  Write-Log -Level ERROR -Message "Error running the Office install. The error is: $_ "
}

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
    # Set Start Menu Layout for users
    # Set the initial Office folder
    [string]$dirOffice = Join-Path -Path "${env:ProgramFiles(x86)}" -ChildPath "Microsoft Office"
    [string]$dirOfficex64 = Join-Path -Path "${env:ProgramFiles}" -ChildPath "Microsoft Office"
    [string[]]$officeExecutables = 'excel.exe', 'onenote.exe', 'outlook.exe', 'MSACCESS.EXE', 'powerpnt.exe', 'winword.exe'
	
    ForEach ($officeExecutable in $officeExecutables) {
    If (Test-Path -Path (Join-Path -Path $dirOfficeX64 -ChildPath "root\Office16\$officeExecutable") -PathType Leaf) {
    
	# Import Start Menu Layout
	Import-StartLayout -LayoutPath "$RibbonFolder\StartMenu-O365.xml" -MountPath C:\
        Break
    }
    }

    # Apply LayoutModification xml to all users
    $Source = "C:\Users\Default\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml"
    $Destination = "C:\Users\*\AppData\Local\Microsoft\Windows\Shell\"
    Get-ChildItem $Destination | ForEach-Object {Copy-Item -Path $Source -Destination $_ -Recurse -Force}

# Regex pattern for SIDs
$PatternSID = 'S-1-5-21-\d+-\d+\-\d+\-\d+$'
 
# Get Username, SID, and location of ntuser.dat for all users
$ProfileList = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' | Where-Object {$_.PSChildName -match $PatternSID} | 
    Select  @{name="SID";expression={$_.PSChildName}}, 
            @{name="UserHive";expression={"$($_.ProfileImagePath)\ntuser.dat"}}, 
            @{name="Username";expression={$_.ProfileImagePath -replace '^(.*[\\\/])', ''}}
 
# Get all user SIDs found in HKEY_USERS (ntuder.dat files that are loaded)
$LoadedHives = Get-ChildItem Registry::HKEY_USERS | ? {$_.PSChildname -match $PatternSID} | Select-Object @{name="SID";expression={$_.PSChildName}}
Write-Verbose "Loaded user: $($LoadedHives)"
Write-Log -Message "Loaded user: $($LoadedHives)"
 
# Get all users that are not currently logged
$UnloadedHives = Compare-Object $ProfileList.SID $LoadedHives.SID | Select-Object @{name="SID";expression={$_.InputObject}}, UserHive, Username
 
# Loop through each profile on the machine
Foreach ($item in $ProfileList) {
    # Load User ntuser.dat if it's not already loaded
    IF ($item.SID -in $UnloadedHives.SID) {
        reg load HKU\$($Item.SID) $($Item.UserHive) | Out-Null
    }
 
    #####################################################################
    # This is where you can read/modify a users portion of the registry 
    $key = "Registry::HKEY_USERS\$($Item.SID)\SOFTWARE\Microsoft\Windows\CurrentVersion\CloudStore\Store"
    Remove-Item -Path $key -Force -Recurse -ErrorAction SilentlyContinue | Out-Null

    #####################################################################
 
    # Unload ntuser.dat        
    IF ($item.SID -in $UnloadedHives.SID) {
        ### Garbage collection and closing of ntuser.dat ###
        Write-Verbose "Unload user hives"
        Write-Log -Message "Unload user hives"
        [gc]::Collect()
        reg unload HKU\$($Item.SID) | Out-Null
    }
}

    Write-Verbose "$($OfficeVersionInstalled) installed successfully!"
    Write-Log -Message "$($OfficeVersionInstalled) installed successfully!"
}
else {
  Write-Warning 'Microsoft 365 was not detected after the install ran'
  Write-Log -Level ERROR -Message "Microsoft 365 was not detected after the install ran"
}

if ($CleanUpInstallFiles) {
  Remove-Item -Path $OfficeInstallDownloadPath -Force -Recurse
  Remove-Item -Path $RibbonFolder -Force -Recurse
}

# Restart computer to apply the setting
Restart-Computer -Force

