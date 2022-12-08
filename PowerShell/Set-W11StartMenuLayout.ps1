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
        [string]$FileName = "Create-W11Layout.log"
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

# Check for OS
$OSBuild = (Get-CimInstance -ClassName Win32_OperatingSystem).BuildNumber
Switch -Wildcard ($OSBuild) {
    '19*' {
        $OSVer = "Windows 10"
        Write-Log "This script is intended for use on Windows 11 devices. $($OSVer) was detected...." -Severity 2
        Exit 1
    }
}


# Check if Office 365 installed
$RegLocations = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
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
    Write-Log -Message "$($OfficeVersionInstalled) detected. Configure Start Menu Layout..."

    $ConfigArray = @"
{"pinnedList":[{"desktopAppId":"MSEdge"},{"desktopAppId":"Microsoft.Office.OUTLOOK.EXE.15"},{"desktopAppId":"Microsoft.Office.WINWORD.EXE.15"},{"desktopAppId":"Microsoft.Office.EXCEL.EXE.15"},{"desktopAppId":"Microsoft.Office.MSACCESS.EXE.15"},{"desktopAppId":"Microsoft.Office.POWERPNT.EXE.15"},{"desktopAppId":"Chrome"},{"desktopAppId":"Microsoft.SoftwareCenter.DesktopToasts"},{"desktopAppId":"Microsoft.Windows.ControlPanel"},{"desktopAppId":"Microsoft.Windows.Explorer"},{"packagedAppId":"windows.immersivecontrolpanel_cw5n1h2txyewy!microsoft.windows.immersivecontrolpanel"},{"packagedAppId":"Microsoft.ScreenSketch_8wekyb3d8bbwe!App"},{"packagedAppId":"Microsoft.WindowsStore_8wekyb3d8bbwe!App"},{"packagedAppId":"Microsoft.WindowsCalculator_8wekyb3d8bbwe!App"},{"packagedAppId":"Microsoft.WindowsNotepad_8wekyb3d8bbwe!App"},{"packagedAppId":"Microsoft.WindowsSoundRecorder_8wekyb3d8bbwe!App"},{"packagedAppId":"Microsoft.Paint_8wekyb3d8bbwe!App"}]}
"@

    if ((Test-Path -LiteralPath "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Start") -ne $True) {
        New-Item -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Start" -Force -ErrorAction SilentlyContinue
    }

    New-ItemProperty -LiteralPath "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Start" -Name "ConfigureStartPins" -Value $ConfigArray -PropertyType String -Force -ErrorAction SilentlyContinue
    New-ItemProperty -LiteralPath "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Start" -Name "ConfigureStartPins_ProviderSet" -Value 1 -PropertyType DWord -Force -ErrorAction SilentlyContinue

}

<#
if ($OfficeInstalled) {
    # Create Start Menu Layout Json
    Write-Log -Message "$($OfficeVersionInstalled) detected. Start creating Layout json file...."
    $jsonstring = @"
    {
    "pinnedList":  [
                       {
                           "desktopAppId":"MSEdge"
                       },
                       {
                           ""desktopAppId":"Microsoft.Office.OUTLOOK.EXE.15"
                       },
                       {
                           "desktopAppId":"Microsoft.Office.WINWORD.EXE.15"
                       },
                       {
                           "desktopAppId":"Microsoft.Office.EXCEL.EXE.15"
                       },
                       {
                           "desktopAppId":"Microsoft.Office.MSACCESS.EXE.15"
                       },
                       {
                           "desktopAppId":"Microsoft.Office.POWERPNT.EXE.15"
                       },
                       {
                           "desktopAppId":"Chrome"
                       },
                       {
                           "desktopAppId":"Microsoft.WindowsStore_8wekyb3d8bbwe!App"
                       },
                       {
                           "desktopAppId":"Microsoft.WindowsCalculator_8wekyb3d8bbwe!App"
                       },
                       {
                           "desktopAppId":"Microsoft.WindowsNotepad_8wekyb3d8bbwe!App"
                       },
                       {
                           "desktopAppId":"Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App"
                       },
                       {
                           "desktopAppId":"Microsoft.SoftwareCenter.DesktopToasts"
                       },
                       {
                           "desktopAppId":"Microsoft.Windows.ControlPanel"
                       },
                       {
                           "packagedAppId":"windows.immersivecontrolpanel_cw5n1h2txyewy!microsoft.windows.immersivecontrolpanel"
                       },
                       {
                           "desktopAppId":"Microsoft.Windows.Explorer"
                       },
                       {
                           "desktopAppId":"Microsoft.ScreenSketch_8wekyb3d8bbwe!ClassicApp"
                       },
                       {
                           "desktopAppId":"Microsoft.Paint_8wekyb3d8bbwe!App"
                       }
                       ]
}
"@
    $JsonObject = $jsonstring | ConvertFrom-Json
    $JsonObject | ConvertTo-Json | Out-File "$env:windir\Temp\LayoutModification.json"
}

# Copy LayoutModification.json to Shell folder
Copy-Item -Path $env:windir\Temp'\LayoutModification.json' -Destination $env:SystemDrive'\Users\Default\Appdata\Local\Microsoft\Windows\Shell'
#>

# Start taskbar pinned item xml
$LayoutFile = "$env:windir\Temp\Layout.xml"
[xml]$LayoutXml = @"
<?xml version="1.0" encoding="utf-8"?>
<LayoutModificationTemplate
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"
    xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout"
    xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout"
    xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"
    Version="1">
  <CustomTaskbarLayoutCollection PinListPlacement="Replace">
    <defaultlayout:TaskbarLayout>
      <taskbar:TaskbarPinList>
        <taskbar:DesktopApp DesktopApplicationID="MSEdge"/>
        <taskbar:DesktopApp DesktopApplicationID="Microsoft.Windows.Explorer" />
        <taskbar:UWA AppUserModelID="Microsoft.ScreenSketch_8wekyb3d8bbwe!App" />
        <taskbar:UWA AppUserModelID="Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App" />
        <taskbar:DesktopApp DesktopApplicationID="Microsoft.Office.OUTLOOK.EXE.15" />
        <taskbar:DesktopApp DesktopApplicationID="Microsoft.Windows.ControlPanel" />
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>
"@
if (-NOT[string]::IsNullOrEmpty($LayoutXml)) {
    Write-Log -Message "Layout.xml file not detected. Start creating...."
    $LayoutXml.Save($LayoutFile)
}
if (Test-Path -Path $LayoutFile) {
    try {
        Write-Log -Message "Layout.xml file detected. Importing xml file..."
        Import-StartLayout -LayoutPath $LayoutFile -MountPath $env:SystemDrive\
        Write-Log -Message "LayoutModification.xml imported!"
    }
    catch [System.Exception] {
        Write-Log -Message "Unable to import Layout.xml. Error message: $($_.Exception.Message)" -Severity 3        
    }
}

# Cleanup layout json and xml file
#Remove-Item -Path "$env:windir\Temp\LayoutModification.json" -Force -ErrorAction SilentlyContinue
#Remove-Item $LayoutFile -Force -ErrorAction SilentlyContinue

$Script_End_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
$Script_Time_Taken = New-TimeSpan -Start $Script_Start_Time -End $Script_End_Time

Write-Log -Message "INFO: Script end: $Script_End_Time"
Write-Log -Message "INFO: Execution time: $Script_Time_Taken"
Write-Log -Message "***************************************************************************"
