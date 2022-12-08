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
        [string]$FileName = "Start-MenuLayout.log"
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
                    3 { $Style = "$($PSStyle.Bold)$($PSStyle.Foreground.BrightRed)"; Write-Host "$($Style)ERROR: $lTextLogLine" }
                    2 { Write-Warning $lTextLogLine }
                    1 { Write-Verbose $lTextLogLine }
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
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output "This is red because I typed FAILED-------------------------------------"
    Write-Output "This is red because I typed FAILED-------------------------------------"
}

AddHeaderSpace

$Script_Start_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
Write-Log -Message "INFO: Script Start: $Script_Start_Time"

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
    Write-Log -Message "$($OfficeVersionInstalled) existed. Start Menu Layout process!"
    # Create custom xml folder
    $xmlDir = "$($env:windir)\Temp\xmlDir"

    If (!(Test-Path $xmlDir)) {
        New-Item -Path $xmlDir -ItemType Directory -Force | Out-Null
        Write-Log -Message "$($xmlDir) created successfully"
    }
    # Create Start Menu Layout xml
    New-Item -Path $xmlDir -ItemType File -Name "Layout.xml"
    [xml](Add-Content -Path "$xmlDir\Layout.xml" -Value '<?xml version="1.0" encoding="utf-8"?>
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
		<taskbar:DesktopApp DesktopApplicationLinkPath="%APPDATA%\Microsoft\Windows\Start Menu\Programs\System Tools\Control Panel.lnk" />
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>')

# Import Start Menu Layout
if (Test-Path "$xmlDir\Layout.xml") {
    Write-Log -Message "Layout xml created. Start the import"
    Import-StartLayout -LayoutPath "$xmlDir\Layout.xml" -MountPath C:\
     Write-Log -Message "Successfully Import StartMenu Layout!"
}

# Cleanup xmlDir Directory
Remove-Item -Path $xmlDir -Force -Recurse
Write-Log -Message "$($xmlDir) cleaned up."
}
else {
    Write-Log -Message "Microsoft 365 was not detected. Exit script." -Severity 3
    exit 0
}

$Script_End_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
$Script_Time_Taken = New-TimeSpan -Start $Script_Start_Time -End $Script_End_Time

Write-Log -Message "Script Start: $Script_Start_Time"
Write-Log -Message "Script end: $Script_End_Time"
Write-Log -Message "Execution time: $Script_Time_Taken"
Write-Log -Message "================================ Script End ====================================="
AddHeaderSpace
