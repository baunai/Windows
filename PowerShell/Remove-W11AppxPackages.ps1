Begin {
    #Log Function
    function Get-TaskSequenceStatus {
        #Determine if a task sequence s currently running
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

            [Parameter(Mandatory = $false, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning, 3 for error.")]
            [ValidateNotNullOrEmpty()]
            [ValidateRange(1, 3)]
            [Int16]$Severiy = 1,

            [Parameter(Mandatory = $false, HelpMessage = "Output script run to the console host")]
            [ValidateNotNullOrEmpty()]
            [bool]$WriteHost = $true,

            [Parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will be written to.")]
            [ValidateNotNullOrEmpty()]
            [string]$FileName = "ScriptLogFileName.Log"
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
                #Create script black for writting entry to the console
                [scriptblock]$WriteLogLineToHost = {
                    param (
                        [string]$lTextLogLine,
                        [Int16]$lSeveriy
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

                & $WriteLogLineToHost -lTextLogLine $msg -lSeverity $Severiy

            }
        }

        $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
        $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{}4" thread="" file="">'
        $LineFormat = $Line -f $LineFormat

        try {
            Out-File -InputObject $Line -Append -NoClobber -Encoding Default -FilePath $LogFilePath
        }
        catch [System.Exception] {
            # Exception is stored in the automatic variable _
            Write-Warning -Message "Unable to append log entry to $($LogFilePath) file. Error message: $($_.Exception.Message)"
        }
    }

    #Leave blank space at top of window to not bloack outpur by progress bars
    function AddHeaderSpace {
        Write-Output "This space intentionally left blank....."
        Write-Output ""
        Write-Output ""
        Write-Output ""
        Write-Output ""
        Write-Output ""
    }

    function Remove-AppxProvisionedPackageCustom {
        #Attemp to remove AppxProvisioningPackage
        if (!([string]::IsNullOrEmpty($BlackListedApp))) {
            try {
                #Get Package Name
                $AppProvisioningPackageName = Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $BlackListedApp} | Select-Object -ExpandProperty PackageName -First 1
                Write-Log -Message "$($BlackListedApp) found. Attempting removal..."

                #Attemp removal
                $RemoveAppx = Remove-AppxProvisionedPackage -PackageName $AppProvisioningPackageName -Online -AllUsers

                #Recheck existence
                $AppProvisioningPackageNameRecheck = Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $BlackListedApp} | Select-Object -ExpandProperty PackageName -First 1

                if ([string]::IsNullOrEmpty($AppProvisioningPackageNameRecheck) -and ($RemoveAppx.Online -eq $true)) {
                    Write-Log -Message @CheckIcon " ( Removed)"
                    Write-Log -Message "$($BlackListedApp) rRemoved"
                }
            }
            catch [System.Exception] {
                Write-Log -Message "Failed to remove $($BlackListedApp)" -Severiy 3
            }
        }
    }
    AddHeaderSpace

    $Script_Start_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
    Write-Log -Message "INFO: Script start: $Script_Start_Time"
    Write-Log -Message "Remove-Appx Started..."

    #OS Checkk
    $OS = (Get-CimInstance -ClassName Win32_OperatingSystem).BuildNumber
    switch -Wildcard ($OS) {
        '19*' { 
            $OSVer = "Windows 10"
            Write-Log -Message "This script is intended for use on Windows 11 devices. $($OSVer) was detected..."
            Exit 1
        }
    }

    #Black List of Appx Provisioned Package to remove for All Users
$BlackListedApps = $null
$BlackListedApps = New-Object -TypeName System.Collections.ArrayList
$BlackListedApps.AddRange(@(
        "Microsoft.BingNews",
        "Microsoft.GamingApp",
        "Microsoft.MicrosoftSolitaireCollection",
        "Microsoft.WindowsCommunicationsApps",
        "Microsoft.WindowsFeedbackHub",
        "Microsoft.XboxGameOverlay",
        "Microsoft.XboxGamingOverlay",
        "Microsoft.XboxIdentityProvider",
        "Microsoft.XboxSpeechToTextOverlay",
        "Microsoft.YourPhone",
        "Microsoft.ZuneMusic",
        "Microsoft.ZuneVideo",
        "MicrosoftTeams"
))

#Define Icons
$CheckIcon = @{
    Object          = [char]8730
    ForegroundColor = 'Green'
    NoNewLine       = $true
}

#Define App Count
[int]$AppCount = 0

}

Process {
    if ($($BlackListedApps.Count) -ne 0) {
        Write-Log -Message "The following $($BlackListedApps.Count) apps were target for removal from the device:-"
        Write-Log -Message "Apps marked for removal: $($BlackListedApps)"
        Write-Log -Message ""
        $BlackListedApps

        #Ininitalize list for apps not targeted
        $AppNotTargetedList = New-Object -TypeName System.Collections.ArrayList

        #Get Appx Provisioned Packages
        Write-Log -Message "Gathering installed Appx Provisioned Packages..."
        Write-Log -Message ""
        $AppArray = Get-AppxProvisionedPackage -Online | Select-Object -ExpandProperty DisplayName

        #Loop through each Provisioned Package
        foreach ($BlackListedApp in $BlackListedApps) {
            #Function call to remove Appx Provisioned Packages defined in the Black List
            if (($BlackListedApp -in $AppArray)) {
                $AppCount++
                try {
                    Remove-AppxProvisionedPackageCustom -BlackListedApp $BlackListedApp
                }
                catch [System.Exception] { 
                    Write-Log -Message "There was error while attempting to remove $($BlackListedApp). Error message: $($_.Exception.Message)"
                }
            }
            else {
                $AppNotTargetedList.AddRange(@($BlackListedApp))
            }
        }

        #Update OutputInformation
        if (!([string]::IsNullOrEmpty($AppNotTargetedList))) {
            Write-Log -Message "The following apps were not removed. Either they were already removed or the Package Name is invalid"
            Write-Log -Message "$($AppNotTargetedList)"
            Write-Log -Message ""
            $AppNotTargetedList
        }
        if ($AppCount -eq 0) {
            Write-Log -Message "No apps were removed. Most likely reason is they had been removed previously."
        }
    }
    else {
        Write-Log -Message "No Black List Apps defined in array"
    }
}
end {
    $Script_End_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongDateString()
    $Script_Time_Taken = New-TimeSpan -Start $Script_Start_Time -End $Script_End_Time
    Write-Log -Message "INFO: Script end: $Script_End_Time"
    Write-Log -Message "INFO: Execution time: $Script_Time_Taken"
    Write-Log -Message "***********************************************************************"
}    
