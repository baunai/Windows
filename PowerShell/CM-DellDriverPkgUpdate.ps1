Function Start-Log {
    Param (
        [Parameter(Mandatory = $True)]
        [String]$FilePath,

        [Parameter(Mandatory = $True)]
        [String]$FileName
    )
	
    Try {
        If (!(Test-Path $FilePath)) {
            ## Create the log file
            New-Item -Path "$FilePath" -ItemType "directory" | Out-Null
            New-Item -Path "$FilePath\$FileName" -ItemType "file"
        }
        Else {
            New-Item -Path "$FilePath\$FileName" -ItemType "file"
        }
		
        ## Set the global variable to be used as the FilePath for all subsequent Write-Log calls in this session
        $global:ScriptLogFilePath = "$FilePath\$FileName"
    }
    Catch {
        Write-Error $_.Exception.Message
        Exit
    }
}



Function Write-Log {
    Param (
        [Parameter(Mandatory = $True)]
        [String]$Message,
		
        [Parameter(Mandatory = $False)]
        # 1 == "Informational"
        # 2 == "Warning'
        # 3 == "Error"
        [ValidateSet(1, 2, 3)]
        [Int]$LogLevel = 1,

        [Parameter(Mandatory = $False)]
        [String]$LogFilePath = $ScriptLogFilePath,

        [Parameter(Mandatory = $False)]
        [String]$ScriptLineNumber
    )

    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$ScriptLineNumber", $LogLevel
    $Line = $Line -f $LineFormat

    #Add-Content -Path $LogFilePath -Value $Line
    Out-File -InputObject $Line -Append -NoClobber -Encoding Default -FilePath $ScriptLogFilePath
}



Function Receive-Output {
    Param(
        $Color,
        $BGColor,
        $LogLevel,
        $LogFile,
        $LineNumber
    )

    Process {
        If ($BGColor) {
            Write-Host $_ -ForegroundColor $Color -BackgroundColor $BGColor
        }
        Else {
            Write-Host $_ -ForegroundColor $Color
        }

        If (($LogLevel) -or ($LogFile)) {
            Write-Log -Message $_ -LogLevel $LogLevel -LogFilePath $ScriptLogFilePath -ScriptLineNumber $LineNumber
        }
    }
}




Function AddHeaderSpace
{
    Write-Output "This space intentionally left blank..." | Receive-Output -Color Gray
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
}




###########################
# Begin script processing #
###########################
Clear-Host

# Leave blank space at top of window to not block output by progress bars
AddHeaderSpace

# Get script start time (will be used to determine how long execution takes)
$Script_Start_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
$Now = Get-Date -Format yyyy-MM-dd_HH-mm-ss


# Start logging
$LogFilePath = Join-Path -Path $env:windir -ChildPath "Temp\Logs"
$LogFileName = "Log--DellDriverPackUpdate--$Now.log"
# Get-ChildItem -Path $LogFilePath\* -Include $LogFileName -ErrorAction SilentlyContinue | Remove-Item -Force -Verbose
Start-Log -FilePath $LogFilePath -FileName $LogFileName



if ((Test-Path $LogFilePath)) {
    $logSize = (Get-Item -Path $LogFilePath).Length / 1MB
    $maxLogSize = 5
}
# Check for file size of the log. If greater than 5MB, it will create a new one and delete the old.
if ((Test-Path $LogFilePath) -AND $LogSize -gt $MaxLogSize) {
    Write-Output "$LogFilePath exceeds maximum size. Deleting the $LogFilePath and starting fresh." | Receive-Output -Color Red -LogLevel 3 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Remove-Item $LogFilePath -Force
    # $newLogFile = New-Item $Path -Force -ItemType File
}

Write-Output "Starting Script: =============================================================" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"

Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - Performing Dell Driver Package Update..." | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"

$Source = "http://downloads.dell.com/catalog/DriverPackCatalog.cab"
$CabPath = "C:\Dell" + "\DriverPackCatalog.cab"
$XMLFile = "C:\Dell" + "\DriverPackCatalog.xml"
$Expand = "C:\Dell"
#$DownloadPackRoot = "$Expand"
$FileServerName = "\\hpdwinad.hpd\departmentFS\Support\SCCMOSDSource\DriverPackages\Dell\Drivers"
$SiteCode = "P01"

# Create $Expand directory if not exist
if (!(Test-Path -Path $Expand)) {
    Write-Output "Creating directory $Expand" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    Try {
        New-Item -Path $Expand -ItemType Directory -Force | Out-Null
    }
    Catch {
        Write-Error "$($_.Exception)"
    }
}



function Get-FolderSize {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Path,
        [ValidateSet("KB", "MB", "GB")]
        $Units = "MB"
    )
    if ( (Test-Path $Path) -and (Get-Item $Path).PSIsContainer ) {
        $Measure = Get-ChildItem $Path -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
        $Sum = $Measure.Sum / "1$Units"
        [PSCustomObject]@{
            "Path"         = $Path
            "Size($Units)" = $Sum
        }
    }
}



$DellModelTable = @(

    @{ SystemID = '098D'; Model = "Precision 3640"; PackageID = "P0100148" }
    @{ SystemID = '097D'; Model = "XPS 15 9500"; PackageID = "P0100354" }
    @{ SystemID = '073A'; Model = "Precision 7920"; PackageID = "P010056F" }
    @{ SystemID = '071E'; Model = "Latitude 5414"; PackageID = "P0100756" }
    @{ SystemID = '0738'; Model = "Precision 5820"; PackageID = "P01007BB" }
    @{ SystemID = '0879'; Model = "Latitude 5420"; PackageID = "P01007FB" }
	@{ SystemID = '09A4'; Model = "Optiplex 7080"; PackageID = "P010083A" }
    @{ SystemID = '081C'; Model = "Latitude 7490"; PackageID = "P0100846" }
	@{ SystemID = '093D'; Model = "Latitude 7220"; PackageID = "P01005AF" }
)


(New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Download Dell Cab File (Invoke-WebRequest works with our without Proxy info if you use -Proxy, so no need to change command if you use Proxy.. unlike Bits Transfer)

Write-Output "Starting Download of Dell Catalog" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($Source, $CabPath)

# Extract XML from Cab File
Write-Output "Starting Expand of Dell Catalog" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
expand $CabPath $XMLFile
Write-Output "Successfully Extracted DriverPackCatalog.xml to $XMLFile" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Start-Sleep -Seconds 1 #Wait for extract to finish
[xml]$XML = Get-Content $XMLFile


# Create Array of Downloads
$Downloads = $XML.DriverPackManifest.DriverPackage
ForEach ($DellModel in $DellModelTable) {
    Write-Output "---Starting to Process Model: $($DellModel.Model), SystemID: $($DellModel.SystemID)---" | Receive-Output -Color Cyan -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    #Create Download Link For Model
    #This is where you might need to change some filters to fit your need, like I needed to exclude (-notmatch) 2 in 1 devices.
    $Target = $Downloads | ? { ($_.SupportedSystems.Brand.Model.SystemID -match $($DellModel.SystemID)) -and ($_.SupportedSystems.Brand.Model.Name -match $($DellModel.Model)) -and ($_.SupportedOperatingSystems.OperatingSystem.osCode -eq "Windows10") -and ($_.SupportedOperatingSystems.OperatingSystem.osArch -eq "x64") }
    $Target = $Target | Where-Object { $PSItem.path -notmatch "2-IN-1" }
    $Target = $Target | Where-Object -FilterScript { $PSItem.path -notmatch "AIO" }
    $Target = $Target | Where-Object -FilterScript { $PSItem.path -notmatch "9020M" }
    $CMPackageVersionLabel = "$($Target.vendorVersion), $($Target.dellVersion)"
    $TargetLink = "http://" + $XML.DriverPackManifest.baseLocation + $Target.path
    $TargetLink = "http://" + $XML.DriverPackManifest.baseLocation + "/" + $Target.path
    $TargetFileName = [System.IO.Path]::GetFileName($TargetLink)
    $TargetPath = "$Expand\$($DellModel.Model)\Driver Cab"
    $TargetFilePathName = "$TargetPath" + "\" + $TargetFileName
    $DellModelNumber = $DellModel.Model.Split( )[1]
    $ReleaseDate = Get-Date $Target.dateTime -Format 'yyyy-MM-dd' -ErrorAction SilentlyContinue


    if (Test-Path 'C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin'){
    Import-Module "C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1"
        }
  

    #Get Current Driver CMPackage Version from CM
    Set-Location -Path "$($SiteCode):"
    $PackageInfo = Get-CMPackage -Id $DellModel.PackageID -Fast
    Set-Location -Path "C:"

    #Do the Download
    if ($PackageInfo.Version -eq $Target.dellVersion) {
        Write-Output "CM Package $($PackageInfo.Name) already Current" |  Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
    }
    Else {
        Write-Output "Updated Driver Pack for $($DellModel.Model) available: $TargetFileName" | Receive-Output -Color Cyan -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"



    if (-NOT(Test-Path $TargetPath)) {
        Write-Output "Creating Directory $TargetPath" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
        New-Item -Path $TargetPath -ItemType Directory
    }


if ((Test-NetConnection 10.10.61.65 -Port 8080).TcpTestSucceeded -eq $True) {
    $UseProxy = $true
    Write-Output "Found Proxy Server, using for Downloads" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})" 
    $ProxyServer = "http://10.10.61.65:8080"
    $BitsProxyList = @("10.10.*.*;10.210.*.*;10.15.*.*;*.hpd;txlets.dps.texas.gov;texaslets.dps.texas.gov;webid.houstonpd.tx.us;cch.dao.hctx.net;cohowa.houstontx.gov;cohnpm01.coh.gov;10.128.192.208;cohnpmweb01.houtx.lcl;166.155.20.*;173.11.145.201;10.1.2.140;10.80.60.140;dig.swbamla.com;*.hpddev.tsc;*.dev;*.test;192.168.10.*;192.168.99.*;citypointe.houtx.lcl;ta02219ab281300.houtx.lcl;tf12sqlrpt01.houtx.lcl;itrptserver.houstontx.gov;i247.ip;csmprdprisrsdb1.houtx.lcl;texaslets.dps.texas.gov;hcsomobpsthrgh.hctx;police.mail.hpd;*.hctx;hfsc.lims.hpd;hfs-jtiisvm-02.HPDWINAD.HPD;www.vinelink.com;houstonpolice-my.sharepoint.com;outlook.office365.com;lpr.houstonhidta.net;lims.houstonforensicscience.org")
}
Else {
    $ProxyServer = $null
    Write-Output "No Proxy Server Found, continuing without" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
}


    $wc = New-Object System.Net.WebClient
    $wc.DownloadFile($TargetLink, "$TargetFilePathName")

<#.        
        #BITS Download of Driver Pack with Retry built in (Do loop)
        $BitStartTime = Get-Date
        Import-Module BitsTransfer
        $DownloadAttempts = 0
        if ($UseProxy -eq $true) { Start-BitsTransfer -Source $TargetLink -Destination $TargetFilePathName -ProxyUsage Override -ProxyList $BitsProxyList -DisplayName "$TargetFileName" -Asynchronous }
        else { Start-BitsTransfer -Source $TargetLink -Destination $TargetFilePathName -DisplayName "$TargetFileName" -Asynchronous }
        do {
            $DownloadAttempts++
            Get-BitsTransfer -Name "$TargetFileName" | Resume-BitsTransfer
            
        }
        while
        ((test-path "$TargetFilePathName") -ne $true)

        #Invoke-WebRequest -Uri $TargetLink -OutFile $TargetFilePathName -UseBasicParsing -Verbose -Proxy $ProxyServer -ErrorAction Stop
        $DownloadTime = $((Get-Date).Subtract($BitStartTime).Second)
           
        Write-Output "Download Complete: $TargetFilePathName" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            
        Write-Output "Took $DownloadTime Seconds to Download $TargetFileName with $DownloadAttempts Attempt(s)" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
.#>

            
        if (test-path $TargetFilePathName) {
            
            $ExpandFolder = "$Expand\$($DellModel.Model)\Windows10-$($Target.dellVersion)"
            #$PackageSourceFolder = "$($ExpandFolder)\$($DellModelNumber)\win10\x64"
            Write-Output "Create Source Folder: $ExpandFolder" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            New-Item -Path $ExpandFolder -ItemType Directory -Force
                
            Write-Output "Starting Expand Process for $($DellModel.Model) file $TargetFileName" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            $Expand = expand $TargetFilePathName -F:* $ExpandFolder
            $FolderSize = (Get-FolderSize $ExpandFolder)
            $FolderSize = [math]::Round($FolderSize.'Size(MB)') 
                
            Write-Output "Finished Expand Process to $ExpandFolder, size: $FolderSize MB" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            

            $DriverPackFullPath = "$FileServerName\$($DellModel.Model)"
            if (Test-Path $DriverPackFullPath) {
            Remove-Item -Path $DriverPackFullPath -Recurse -Force
            New-Item $DriverPackFullPath -ItemType Directory -Force
            } else {
            New-Item $DriverPackFullPath -ItemType Directory -Force
            }

            $CopyfromDir = (Get-ChildItem -Path ((Get-ChildItem -Path $ExpandFolder -Directory).FullName) -Directory).FullName
            Write-Output "Copy drivers to $DriverPackFullPath" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
            Copy-Item "$($CopyfromDir)\x64\*" -Destination $DriverPackFullPath -Force -Recurse
            #Write-Output "Export DriverPackInfo xml to $DriverPackFullPath"
            #Export-Clixml -InputObject "$($ExpandFolder)\$($DellModelNumber)" -Path "$($DriverPackFullPath)\Manifest.xml"

            #Update Package to Point to new Source & Trigger Distribution

            Set-Location -Path "$($SiteCode):"
            Set-CMPackage -Id $DellModel.PackageID -Path $DriverPackFullPath
            Set-CMPackage -Id $DellModel.PackageID -Version $Target.dellVersion
            Set-CMPackage -Id $DellModel.PackageID -Manufacturer "Dell"
            Set-CMPackage -Id $DellModel.PackageID -Description "$($DellModel.Model) Driver Cab Version $CMPackageVersionLabel Released $($ReleaseDate)."
            Set-CMPackage -Id $DellModel.PackageID -Language $Target.releaseID
            $PackageInfo = Get-CMPackage -Id $DellModel.PackageID -Fast
            Update-CMDistributionPoint -PackageId $DellModel.PackageID
            Set-Location -Path "C:"
            Write-Output "Updated Package $($PackageInfo.Name), ID $($DellModel.PackageID) to $($PackageInfo.Version) which was released $ReleaseDate" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"

           }
            
        }
    }
          

Write-Output "Script End: =============================================================" | Receive-Output -Color Green -BGColor Black -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"


# Determine ending time
$Script_End_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
$Script_Time_Taken = New-TimeSpan -Start $Script_Start_Time -End $Script_End_Time

# How long did this take?
Write-Output "Script start: $Script_Start_Time" | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output "Script end:   $Script_End_Time" | Receive-Output -Color Gray -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"
Write-Output ""
Write-Output "Execution time: $Script_Time_Taken seconds" | Receive-Output -Color White -LogLevel 1 -LineNumber "$($Invocation.MyCommand.Name):$( & {$MyInvocation.ScriptLineNumber})"

