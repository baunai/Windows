function Invoke-HPDriverUpdate {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, HelpMessage = "Specify the HPIA action to perform, e.g. Download or Install")]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Download", "Install")]
        [string]$HPIAAction = "Install"
    )
    
    begin {
        # Enable TLS 1.2 support for downloading modules form PSGallery
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    }
    
    process {
        # Functions
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
                [string]$FileName = "Invoke-HPDriverUpdate.log"
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

        function Set-RegistryValue {
            param (
                [Parameter(Mandatory = $true)]
                [ValidateNotNullOrEmpty()]
                [string]$Path,

                [Parameter(Mandatory = $true)]
                [ValidateNotNullOrEmpty()]
                [string]$Name,

                [Parameter(Mandatory = $true)]
                [ValidateNotNullOrEmpty()]
                [string]$Value
            )
            try {
                $RegistryValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
                if ($null -ne $RegistryValue) {
                    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Force -ErrorAction Stop
                } else {
                    if (-NOT (Test-Path -Path $Path)) {
                        New-Item -Path $Path -ItemType Directory -Force -ErrorAction Stop | Out-Null
                    }
                    New-ItemProperty -Path $Path -Name $Name -PropertyType String -Value $Value -Force -ErrorAction Stop
                }
            }
            catch [System.Exception] {
                Write-Warning -Message "Failed to create or update registry value '$($Name)' in '$($Path)'. Error message: $($_.Exception.Message)"
            }
        }
        function Invoke-Executable {
            param (
                [Parameter(Mandatory = $true)]
                [ValidateNotNullOrEmpty()]
                [string]$FilePath,

                [Parameter(Mandatory = $true)]
                [ValidateNotNullOrEmpty()]
                [string]$Arguments
            )
            
            # Construct a hash table for default parameter splatting
            $SplatArgs = @{
                FilePath = $FilePath
                NoNewWindow = $true
                Passthru = $true
                ErrorAction = 'Stop'
            }
            # Add ArgumentList param if present
            if (! ([System.String]::IsNullOrEmpty($Arguments))) {
                $SplatArgs.Add("ArgumentList", $Arguments)
            }
            # Invoke executable and wait for process to exit
            try {
                $Invocation = Start-Process @SplatArgs
                $Handle = $Invocation.Handle
                $Invocation.WaitForExit()
            }
            catch [System.Exception] {
                Write-Warning -Message $_.Exception.Message; break
            }
            # Handle return value with exitcode from process
            return $Invocation.ExitCode
        }

        # Validate that script is executed on HP hardware
        $Manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty Manufacturer).Trim()
        switch -Wildcard ($Manufacturer) {
            "*HP*" {
                Write-Log -Message "HP hardware validated, allowed to continue"
            }
            "*Hewlett-Packard*" {
                Write-Log -Message "HP hardware validated, allowed to continue"
            }
            Default {
                Write-Log -Message "Unsupported hardware detected, HP hardware is required for this script to operate" -Severity 3; exit 1
            }
        }
        
        # Check HPCMSL module
        $ModuleName = 'HPCMSL'
        if (! (Get-Module -Name $ModuleName -ListAvailable)) {
            function Step-Environment {
                [CmdletBinding()]
                param ()
                if (Get-Item env:LocalAppData -ErrorAction Ignore) {
                    Write-Log -Message 'System Environment Variable LocalAppData is already present in this PowerShell session'
                }
                else {
                    Write-Log -Message 'Set LocalAppData in System Environment'
                    Write-Log -Message 'WinPE does not have the LocalAppData System Environment Variable'
                    Write-Log -Message 'This can be enabled for this Power Session, but it will not persist'
                    Write-Log -Message 'Set System Environment Variable LocalAppData for this PowerShell session'
                    #[System.Environment]::SetEnvironmentVariable('LocalAppData',"$env:UserProfile\AppData\Local")
                    [System.Environment]::SetEnvironmentVariable('APPDATA', "$Env:UserProfile\AppData\Roaming", [System.EnvironmentVariableTarget]::Process)
                    [System.Environment]::SetEnvironmentVariable('HOMEDRIVE', "$Env:SystemDrive", [System.EnvironmentVariableTarget]::Process)
                    [System.Environment]::SetEnvironmentVariable('HOMEPATH', "$Env:UserProfile", [System.EnvironmentVariableTarget]::Process)
                    [System.Environment]::SetEnvironmentVariable('LOCALAPPDATA', "$Env:UserProfile\AppData\Local", [System.EnvironmentVariableTarget]::Process)          
                }
            }
    
            function Step-PowerShellProfile {
                [CmdletBinding()]
                param ()

                if (-not (Test-Path "$env:UserProfile\Documents\WindowsPowerShell")) {
                    $null = New-Item -Path "$env:UserProfile\Documents\WindowsPowerShell" -ItemType Directory -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                }
                Write-Host -ForegroundColor DarkGray 'Set LocalAppData in PowerShell Profile'
                $winpePowerShellProfile | Set-Content -Path "$env:UserProfile\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1" -Force -Encoding Unicode
            }

            function Step-TrustPSGallery {
                [CmdletBinding()]
                param ()
                $PSRepository = Get-PSRepository -Name PSGallery
                if ($PSRepository) {
                    if ($PSRepository.InstallationPolicy -ne 'Trusted') {
                        Write-Log 'Set-PSRepository PSGallery Trusted'
                        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
                    }
                }
            }

            function Step-InstallNuget {
                [CmdletBinding()]
                param ()
                $NuGetClientSourceURL = 'https://nuget.org/nuget.exe'
                $NuGetExeName = 'NuGet.exe'
                Write-Log -Message "Install NuGet from $NugetClientSourceURL"
        
                $PSGetProgramDataPath = Join-Path -Path $env:ProgramData -ChildPath 'Microsoft\Windows\PowerShell\PowerShellGet\'
                $nugetExeBasePath = $PSGetProgramDataPath
                if (-not (Test-Path -Path $nugetExeBasePath)) {
                    $null = New-Item -Path $nugetExeBasePath -ItemType Directory -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                }
                $nugetExeFilePath = Join-Path -Path $nugetExeBasePath -ChildPath $NuGetExeName
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                $null = Invoke-WebRequest -UseBasicParsing -Uri $NuGetClientSourceURL -OutFile $nugetExeFilePath
        
                $PSGetAppLocalPath = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Microsoft\Windows\PowerShell\PowerShellGet\'
                $nugetExeBasePath = $PSGetAppLocalPath
        
                if (-not (Test-Path -Path $nugetExeBasePath)) {
                    $null = New-Item -Path $nugetExeBasePath -ItemType Directory -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                }
                $nugetExeFilePath = Join-Path -Path $nugetExeBasePath -ChildPath $NuGetExeName
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                $null = Invoke-WebRequest -UseBasicParsing -Uri $NuGetClientSourceURL -OutFile $nugetExeFilePath
        
            }
    
            function Step-InstallPackageManagement {
                [CmdletBinding()]
                param ()
                $InstalledModule = Import-Module PackageManagement -PassThru -ErrorAction Ignore
                if (-not $InstalledModule) {
                    Write-Log -Message 'Install PackageManagement'
                    $PackageManagementURL = "https://psg-prod-eastus.azureedge.net/packages/packagemanagement.1.4.7.nupkg"
                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                    Invoke-WebRequest -UseBasicParsing -Uri $PackageManagementURL -OutFile "$env:TEMP\packagemanagement.1.4.7.zip"
                    $null = New-Item -Path "$env:TEMP\1.4.7" -ItemType Directory -Force
                    Expand-Archive -Path "$env:TEMP\packagemanagement.1.4.7.zip" -DestinationPath "$env:TEMP\1.4.7"
                    $null = New-Item -Path "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement" -ItemType Directory -ErrorAction SilentlyContinue
                    Move-Item -Path "$env:TEMP\1.4.7" -Destination "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement\1.4.7"
                    Import-Module PackageManagement -Force -Scope Global
                }
            }
 
            function Step-InstallPowerShellGet {
                [CmdletBinding()]
                param ()
                $InstalledModule = Import-Module PowerShellGet -PassThru -ErrorAction Ignore
                if (-not $InstalledModule) {
                    Write-Log -Message 'Install PowerShellGet'
                    $PowerShellGetURL = "https://psg-prod-eastus.azureedge.net/packages/powershellget.2.2.5.nupkg"
                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                    Invoke-WebRequest -UseBasicParsing -Uri $PowerShellGetURL -OutFile "$env:TEMP\powershellget.2.2.5.zip"
                    $null = New-Item -Path "$env:TEMP\2.2.5" -ItemType Directory -Force
                    Expand-Archive -Path "$env:TEMP\powershellget.2.2.5.zip" -DestinationPath "$env:TEMP\2.2.5"
                    $null = New-Item -Path "$env:ProgramFiles\WindowsPowerShell\Modules\PowerShellGet" -ItemType Directory -ErrorAction SilentlyContinue
                    Move-Item -Path "$env:TEMP\2.2.5" -Destination "$env:ProgramFiles\WindowsPowerShell\Modules\PowerShellGet\2.2.5"
                    Import-Module PowerShellGet -Force -Scope Global
                }
            }  

            Step-Environment
            Step-PowerShellProfile
            Step-TrustPSGallery
            Step-InstallNuget
            Step-InstallPackageManagement
            Step-InstallPowerShellGet

            #Setup LOCALAPPDATA Variable
            [System.Environment]::SetEnvironmentVariable('LOCALAPPDATA', "$env:SystemDrive\Windows\system32\config\systemprofile\AppData\Local")

            $WorkingDir = $env:TEMP

            Write-Log -Message "Setup LOCALAPPDATA Variable"
            Write-Log -Message "Temp Appdata: $($WorkingDir)"

            #PowerShellGet from PSGallery URL
            $PSGPath = "$env:ProgramFiles\WindowsPowerShell\Modules\PowerShellGet\2.2.5"
            if (Test-Path $PSGPath) {
                Write-Log -Message "PowerShellGet installed."
            }
            else {
                if (!(Get-Module -Name PowerShellGet)) {   
                    Write-Log -Message "PowerShellGet module not found. Start downloading and installing..."
                    $PowerShellGetURL = "https://psg-prod-eastus.azureedge.net/packages/powershellget.2.2.5.nupkg"    
                    Write-Log -Message "PowerShellGet URL: $($PowerShellGetURL)"
                    Invoke-WebRequest -UseBasicParsing -Uri $PowerShellGetURL -OutFile "$WorkingDir\powershellget.2.2.5.zip"
                    $Null = New-Item -Path "$WorkingDir\2.2.5" -ItemType Directory -Force
                    Expand-Archive -Path "$WorkingDir\powershellget.2.2.5.zip" -DestinationPath "$WorkingDir\2.2.5"
                    $Null = New-Item -Path "$env:ProgramFiles\WindowsPowerShell\Modules\PowerShellGet" -ItemType Directory -ErrorAction SilentlyContinue
                    Move-Item -Path "$WorkingDir\2.2.5" -Destination "$env:ProgramFiles\WindowsPowerShell\Modules\PowerShellGet\2.2.5"
                    Remove-Item "$WorkingDir\powershellget.2.2.5.zip" -Recurse -Force
                    Write-Log -Message "PowerShellGet installed."
                }
            }

            #PackageManagement from PSGallery URL
            $PkgMgtPath = "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement\1.4.7"
            if (Test-Path $PkgMgtPath) {
                Write-Log -Message "PackageManagement installed."
            }
            else {
                if (!(Get-Module -Name PackageManagement)) {
                    Write-Log -Message "PackageManagement not found. Start downloading and installing..."
                    $PackageManagementURL = "https://psg-prod-eastus.azureedge.net/packages/packagemanagement.1.4.7.nupkg"
                    Invoke-WebRequest -UseBasicParsing -Uri $PackageManagementURL -OutFile "$WorkingDir\packagemanagement.1.4.7.zip"
                    $Null = New-Item -Path "$WorkingDir\1.4.7" -ItemType Directory -Force
                    Expand-Archive -Path "$WorkingDir\packagemanagement.1.4.7.zip" -DestinationPath "$WorkingDir\1.4.7"
                    $Null = New-Item -Path "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement" -ItemType Directory -ErrorAction SilentlyContinue
                    Move-Item -Path "$WorkingDir\1.4.7" -Destination "$env:ProgramFiles\WindowsPowerShell\Modules\PackageManagement\1.4.7"
                    Remove-Item "$WorkingDir\packagemanagement.1.4.7.zip" -Recurse -Force
                    Write-Log -Message "PackageManagement installed."
                }
            }

            #Import PowerShellGet
            Import-Module PowerShellGet
            Write-Log -Message "PowerShellGet module imported."


            #Install Module from PSGallery
            Install-Module -Name $ModuleName -Force -AcceptLicense -SkipPublisherCheck
            try {
                if (Get-Module -Name $ModuleName -ListAvailable) {
                    Write-Log -Message "$ModuleName module installed."
                }
            }
            catch [System.Exception] {
                # Exception is stored in the automatic variable _
                Write-Log -Message "Unable to install $ModuleName. Error message: $($_.Exception.Message)" -Severity 3
            }

            Import-Module -Name $ModuleName -Force
        }

        # Create HPIA directory for HP Image Assistant extraction
        $HPIAExtractPath = Join-Path -Path $env:SystemRoot -ChildPath "Temp\HPIA"
        if (! (Test-Path -Path $HPIAExtractPath)) {
            Write-Log -Message "Creating directory for HP Image Assistant extraction: $($HPIAExtractPath)"
            [void][System.IO.Directory]::CreateDirectory($HPIAExtractPath)
        }
        # Create logs for HPIA
        $HPIAReportPath = Join-Path -Path $env:SystemRoot -ChildPath "Temp\HPIALogs"
        if (! (Test-Path -Path $HPIAReportPath)) {
            Write-Log -Message "Creating directory for HPIA report logs: $($HPIAReportPath)"
            New-Item -Path $HPIAReportPath -ItemType Directory -Force | Out-Null
        }
        # Create HP Drivers directory for driver content
        $SoftpaqDownloadPath = Join-Path -Path $env:SystemRoot -ChildPath "Temp\HPDrivers"
        if (! (Test-Path -Path $SoftpaqDownloadPath)) {
            Write-Log -Message "Creating directory for softpaq downloads: $($SoftpaqDownloadPath)"
            [void][System.IO.Directory]::CreateDirectory($SoftpaqDownloadPath)
        }
        # Set current working directory to HPIA directory
        Write-Log -Message "Switching working directory to: $($env:SystemRoot)\Temp"
        Set-Location -Path (Join-Path -Path $env:SystemRoot -ChildPath "Temp")

        try {
            # Download HPIA softpaq and extract it to Temp directory
            Write-Log -Message "Attempting to download and extract HPIA to: $($HPIAExtractPath)"
            Install-HPImageAssistant -Extract -DestinationPath $HPIAExtractPath -Quiet -ErrorAction Stop

            try {
                # Invoke HPIA to install drivers and driver software
                $HPIAExecutablePath = Join-Path -Path $env:SystemRoot -ChildPath "Temp\HPIA\HPImageAssistant.exe"
                switch ($HPIAAction) {
                    'Download' {
                        Write-Log -Message "Attempting to execute HP Image Assistant to download drivers including driver software, this might take some time"
                        # Prepare arguments for HP Image Assistant download mode
                        $HPIAArguments = "/Operation:Analyze /Action:Download /Noninteractive /Selection:All /Silent /Category:BIOS,Drivers,Software /ReportFolder:$($HPIAReportPath) /LogFolder:$($HPIAReportPath) /SoftpaqDownloadFolder:$($SoftpaqDownloadPath)"
                        # Set HP Image Assistant operational mode in registry
                        Set-RegistryValue -Path "HKLM:\SOFTWARE\HP\ImageAssistant" -Name "OperationalMode" -Value "Download" -ErrorAction Stop
                    }
                    'Install' {
                         Write-Log -Message "Attempting to execute HP Image Assistant to download and install drivers including driver software, this might take some time"
                         # Prepare arguments for HPIA install mode
                         $HPIAArguments = "/Operation:Analyze /Action:Install /Noninteractive /Selection:All /Category:BIOS,Drivers,Software /ReportFolder:$($HPIAReportPath) /LogFolder:$($HPIAReportPath) /SoftpaqDownloadFolder:$($SoftpaqDownloadPath)"

                         # Set HPIA operational mode in registry
                         Set-RegistryValue -Path "HKLM:\SOFTWARE\HP\ImageAssistant" -Name "OperationMode" -Value "Install" -ErrorAction Stop
                        }
                }

                # Invoke HP Image Assistant
                $Invocation = Invoke-Executable -FilePath $HPIAExecutablePath -Arguments $HPIAArguments -ErrorAction Stop

                # Add a registry key for Win32 app detection rule based on HP Image Assistant exit code
                switch ($Invocation) {
                    0 {
                        Write-Log -Message "HP Image Assistant returned successful exit code: $($Invocation)"
                    }
                    256 {
                        # The analysis returned no recommendations
                        Write-Log -Message "HP Image Assistant returned there were no update recommendations for this system, exit code: $($Invocation)"
                    }
                    3010 {
                        # Softpaqs installations are successful, but at least one requires a restart
                        Write-Log -Message "HP Image Assistant returned successful exit code: $($Invocation)"
                    }
                    3020 {
                        # One or more Softpaq's failed to install
                        Write-Log -Message "HP Image Assistant did not install one or more softpaqs successfully, examine the Readme*.html file in: $($HPIAReportPath)" -Severity 2
                        Write-Log -Message "HP Image Assistant returned successful exit code: $($Invocation)"
                        Set-RegistryValue -Path "HKLM:\SOFTWARE\HP\ImageAssistant" -Name "ExecutionResult" -Value "Failed" -ErrorAction Stop
                    }
                    4096 {
                        # This platform is not supported!
                        Write-Log -Message "This platform is not supported!" -Severity 2
                        Write-Log -Message "HP Image Assistant returned successful exit code: $($Invocation)"
                        Set-RegistryValue -Path "HKLM:\SOFTWARE\HP\ImageAssistant" -Name "ExecutionResult" -Value "Failed" -ErrorAction Stop
                    }
                    4097 {
                        # This parameters are invalid
                        Write-Log -Message "This parameters are invalid" -Severity 2
                        Write-Log -Message "HP Image Assistant returned successful exit code: $($Invocation)"
                        Set-RegistryValue -Path "HKLM:\SOFTWARE\HP\ImageAssistant" -Name "ExecutionResult" -Value "Failed" -ErrorAction Stop
                    }
                    4098 {
                        # There is no internet connection.
                        Write-Log -Message "There is no internet connection." -Severity 3
                        Write-Log -Message "HP Image Assistant returned successful exit code: $($Invocation)"
                        Set-RegistryValue -Path "HKLM:\SOFTWARE\HP\ImageAssistant" -Name "ExecutionResult" -Value "Failed" -ErrorAction Stop
                    }
                    8192 {
                        # The operation failed.
                        Write-Log -Message "The operation failed." -Severity 3
                        Write-Log -Message "HP Image Assistant returned successful exit code: $($Invocation)"
                        Set-RegistryValue -Path "HKLM:\SOFTWARE\HP\ImageAssistant" -Name "ExecutionResult" -Value "Failed" -ErrorAction Stop
                    }
                    8196 {
                        # The supported platform list download failed
                        Write-Log -Message "The supported platform download list failed." -Severity 3
                        Write-Log -Message "HP Image Assistant returned successful exit code: $($Invocation)"
                        Set-RegistryValue -Path "HKLM:\SOFTWARE\HP\ImageAssistant" -Name "ExecutionResult" -Value "Failed" -ErrorAction Stop
                    }
                    16384 {
                        # The reference file failed to open
                        Write-Log -Message "The reference file failed to open" -Severity 3
                        Write-Log -Message "HP Image Assistant returned successful exit code: $($Invocation)"
                        Set-RegistryValue -Path "HKLM:\SOFTWARE\HP\ImageAssistant" -Name "ExecutionResult" -Value "Failed" -ErrorAction Stop
                    }
                    16385 {
                        # The reference file failed to open
                        Write-Log -Message "The reference file is invalid." -Severity 2
                        Write-Log -Message "HP Image Assistant returned successful exit code: $($Invocation)"
                        Set-RegistryValue -Path "HKLM:\SOFTWARE\HP\ImageAssistant" -Name "ExecutionResult" -Value "Failed" -ErrorAction Stop
                    }
                    16386 {
                        # The Windows Build is NOT supported by HPIA
                        Write-Log -Message "The Windows Build is NOT supported by HPIA" -Severity 3
                        Write-Log -Message "HP Image Assistant returned successful exit code: $($Invocation)"
                        Set-RegistryValue -Path "HKLM:\SOFTWARE\HP\ImageAssistant" -Name "ExecutionResult" -Value "Failed" -ErrorAction Stop
                    }
                    Default {
                        Write-Log -Message "HP Image Assistant returned unhandled exit code: $($Invocation)"
                        Set-RegistryValue -Path "HKLM:\SOFTWARE\HP\ImageAssistant" -Name "ExecutionResult" -Value "Failed" -ErrorAction Stop
                    }
                }
                if ($HPIAAction -like "Install") {
                    # Cleanup download softpaq executable that was extracted
                    Write-Log -Message "Attempting to cleanup directory for downloaded softpaqs: $($SoftpaqDownloadPath)"
                    Remove-Item -Path $SoftpaqDownloadPath -Force -Recurse -Confirm:$false
                }
                # Cleanup extracted HPIA directory
                Write-Log -Message "Attempting to cleanup extracted HP Image Assistant directory: $($HPIAExtractPath)"
                Remove-Item -Path $HPIAExtractPath -Force -Recurse -Confirm:$false
            }
            catch [System.Exception] {
                Write-Log -Message "Failed to run HP Image Assistant to install drivers and driver software. Error message: $($_.Exception.Message)" -Severity 3; exit 1
            }
        }
        catch [System.Exception] {
            Write-Log -Message "Failed to download and extract HP Image Assistant softpaq. Error message: $($_.Exception.Message)" -Severity 3; exit 1
        }

    }
    
    end {
        
    }
}

Invoke-HPDriverUpdate -HPIAAction Install
