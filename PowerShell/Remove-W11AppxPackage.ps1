#Log Function
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
        [string]$FileName = "Remove-W11AppxPackages.log"
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

#Fucntion to Remove AppxProvisionedPackage
function Remove-AppxProvisionedPackageCustom {
    #Attemp to remove AppxProvisionedPackage
    if (-NOT([string]::IsNullOrEmpty($BlackListedApp))) {
        try {
    #Get Package Name
            $AppProvisioningPackageName = Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $BlackListedApp} | Select-Object -ExpandProperty PackageName -First 1
            Write-Log -Message "$($BlackListedApp) found. Attempting removal..."

    #Attempt removal
        $RemoveAppx = Remove-AppxProvisionedPackage -PackageName $AppProvisioningPackageName -Online -AllUsers

    #Recheck existence
        $AppProvisioningPackageNameRecheck = Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $BlackListedApp} | Select-Object -ExpandProperty PackageName -First 1

            if ([string]::IsNullOrEmpty($AppProvisioningPackageNameRecheck) -and ($RemoveAppx.Online) -eq $true) {
                Write-Host @CheckIcon
                Write-Host " (Removed)"
                Write-Log -Message "$($BlackListedApp) removed."
            }
        }
        catch [System.Exception] {
            Write-Log -Message "Failed to remove $($BlackListedApps). Error message: $($_.Exception.Message)" -Severity 3               
        }
    }      
}
    Write-Log -Message "##############################"
    Write-Log -Message "Remove-Appx Started"
    Write-Log -Message "##############################"

    #OS Check
    $OS = (Get-CimInstance -ClassName Win32_OperatingSystem).BuildNumber
    switch -Wildcard ($OS) {
        '21*' {
            $OSVer = "Windows 10"
            Write-Log -Message "This script is intend for using on Windows 11 devices. $($OSVer) was detected..."
            Exit 1
            }
    }

    #Black List of Appx Provisioned Packages to Remove for All Users
    $BlackListedApps = $null
    $BlackListedApps = New-Object -TypeName System.Collections.ArrayList
    $BlackListedApps.AddRange(@(
        "Microsoft.BingNews",
        "Microsoft.GetHelp",
        "Microsoft.GetStarted",
        "Microsoft.People",
        "Microsoft.MicrosoftOfficeHub",
        "Microsoft.WindowsCommunicationsApps",
        "Microsoft.WindowsFeedbackHub",
        "Microsoft.GamingApp",
        "Microsoft.MicrosoftSolitaireCollection",
        "Microsoft.WindowsFeedbackHub",
        "Microsoft.Xbox.TCUI",
        "Microsoft.XboxGameOverlay",
        "Microsoft.XboxGamingOverlay",
        "Microsoft.XboxIdentityProvider",
        "Microsoft.XboxSpeechToTextOverlay",
        "Microsoft.YourPhone",
        "Microsoft.ZuneMusic",
        "Microsoft.ZuneVideo",
        "MicrosoftTeams"
        )
    )
    
    #Define Icons
    $CheckIcon = @{
        Object          = [char]8730
        ForegroundColor = 'Green'
        NoNewLine       = $true
    }

    #Define App Count
    [int]$AppCount = 0
    

    if ($($BlackListedApps.Count) -ne 0) {
        Write-Log -Message "The Following $($BlackListedApps.Count) apps were targeted for removal from the device:-"
        Write-Log -Message "Apps marked for removal:$($BlackListedApps)"
        Write-Log -Message ""
        $BlackListedApps

        #Initialize list for apps not targeted
        $AppNotTargetedList = New-Object -TypeName System.Collections.ArrayList

        #Get Appx Provisioned Packages
        Write-Log -Message "Gathering installed Appx Provisioned Packages..."
        Write-Log -Message "....."
        $AppArray = Get-AppxProvisionedPackage -Online | Select-Object -ExpandProperty DisplayName

        #Loop through each Provisioned Package
        foreach ($BlackListedApp in $BlackListedApps) {

            #Function call to Remove Appx Provisioned Packages definced in the Black List
            if ($BlackListedApp -in $AppArray) {
                $AppCount ++
                try {
                    Remove-AppxProvisionedPackageCustom -BlackListedApp $BlackListedApp
                }
                catch [System.Exception] {
                    # Exception is stored in the automatic variable Remove-AppxProvisionedPacak
                    Write-Log -Message "There was error when attempting to remove $($BlackListedApp). Error message: $($_.Exception.Message)"
                }
            } else {
                $AppNotTargetedList.AddRange(@($BlackListedApp))
            }
        }

        #Update Output Information
        if (!([string]::IsNullOrEmpty($AppNotTargetedList))) {
            Write-Log -Message "The following apps were not removed. Either they were already removed or the Package Name is invalid:-"
            Write-Log -Message "$($AppNotTargetedList)"
            Write-Log -Message "....."
            $AppNotTargetedList
        }
        if ($AppCount -eq 0) {
            Write-Log -Message "No apps were removed. Most likely reason is they had been removed previously."
        }
    } else {
        Write-Log -Message "No Black List Apps defined in array"
    }

    #Remove Windows Capabilities
    $WhiteListOnDemand = "NetFX3|DirectX|Tools.DeveloperMode.Core|Language|ContactSupport|OneCoreUAP|WindowsMediaPlayer|Hello.Face|Notepad|MSPaint|App.StepsRecorder|Windows.Kernel.LA57~~~~0.0.1.0|MathRecognizer~~~~0.0.1.0|OpenSSH.Client~~~~0.0.1.0|Microsoft.Windows.WordPad|Print.Fax.Scan|Print.Management.Console|PowerShell.ISE|ShellComponents"
    $OnDemandFeatures = Get-WindowsCapability -Online -LimitAccess -ErrorAction Stop | Where-Object { $_.Name -notmatch $WhiteListOnDemand -and $_.State -like "Installed" } | Select-Object -ExpandProperty Name

    foreach ($Feature in $OnDemandFeatures) {
        try {
            Write-Log -Message "Removing Feature on Demand V2 package: $($Feature)"
            # Handle cmdlet limitations for older OS builds
            Get-WindowsCapability -Online -LimitAccess -ErrorAction Stop | Where-Object { $_.Name -like $Feature } | Remove-WindowsCapability -Online -ErrorAction Stop | Out-Null
        }
        catch [System.Exception] {
            Write-Log -Message "Removing Feature on Demand V2 package failed: $($_.Exception.Message)" -Severity 3
        }
    }

    $Script_End_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
    $Script_Time_Taken = New-TimeSpan -Start $Script_Start_Time -End $Script_End_Time

    Write-Log -Message "INFO: Script end: $Script_End_Time"
    Write-Log -Message "INFO: Execution time: $Script_Time_Taken"
    Write-Log -Message "***************************************************************************"
