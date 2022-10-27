function Check-IsInstalled {
#this is funtion for check it the desired application is installed.
param(
    [Parameter(Mandatory=$True)]
    $ApplicationName, 

    [Parameter(Mandatory=$True)]
    $Version
)
    
    $x86 = ((Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall") |
        Where-Object { $_.GetValue( "DisplayName" ) -like "*$ApplicationName*" -and $_.GetValue( "DisplayVersion" ) -like "$version"  } ).Length -gt 0;

    $x64 = ((Get-ChildItem "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall") |
        Where-Object { $_.GetValue( "DisplayName" ) -like "*$ApplicationName*" -and $_.GetValue( "DisplayVersion" ) -like "$version"  }).Length -gt 0;

    return $x86 -or $x64;
}

#Check if notepad++ 7.3.3 installed
$check = Check-IsInstalled -ApplicationName "notepad++" -Version "7.3.3"

#If notepad++ 7.3.3 is not installed
if ($check -ne $true) {
#Check what drive letter is used
$used = Get-PSDrive | Select-Object -Expand Name | Where-Object { $_.Length -eq 1 }

#Set unused drive letter
$unused = 90 .. 65 | ForEach-Object { [string][char]$_ } | Where-Object { $used -notcontains $_ }
$drive = $unused[(Get-Random -Minimum 0 -Maximum $unused.Count)]

#use Azure Storage access key to map Storage as local drive
$acctKey = ConvertTo-SecureString -String "wMLFueEVyu+KlTSzMvlkfy9SmOXVfHK+9gjffENTcQrMS4BXPptImleGdLLdMe74/C4+C+WM9RHPlB10DInEaQ==" -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential -ArgumentList "Azure\smsboot", $acctKey
New-PSDrive -Name $drive -PSProvider FileSystem -Root "\\smsboot.file.core.windows.net\fileshare01" -Credential $credential

#Import BitsTransfer module
try {
    Import-Module BitsTransfer -ErrorAction Stop -PassThru
}
catch {
    throw 'BitsTransfer PowerShell module is not installed!'
}

#Set copy source path
$Source = "$drive" + ":\Notepad_7.3.3"

#Set copy destination path
$Destination = "c:\Temp\Notepad_7.3.3"

#Copy source to local drive with sub folders and files
$folders = Get-ChildItem -Name -Path $source -Directory -Recurse
foreach ($i in $folders)
{
    #Create sub folders
	$exists = Test-Path $Destination\$i
	if ($exists -eq $false)
	{
		New-Item $Destination\$i -ItemType Directory
	}
	
	Start-BitsTransfer -Source $Source\$i\*.* -Destination $Destination\$i -Priority Foreground
	
}

Start-BitsTransfer -Source $Source\*.* -Destination $Destination -Priority Foreground

#Delete mapped drive
Remove-PSDrive $drive

#Set location to Destination folder
Set-Location $Destination

#Install notepad++ by using PSAT tool, must use Silent mode
Write-Output "Start installing notepad"
Start-Process ".\Deploy-Application.exe" -ArgumentList "-DeployMode 'Silent'" -Wait

#Reset localtion to C:\Windows
Set-Location "C:\Windows"

#Try to remove installation media after application is installed, retry for 1 minute.
$timeout = new-timespan -Minutes 1
$sw = [diagnostics.stopwatch]::StartNew()
Do
{
	Start-Sleep 5
	Remove-Item $Destination -Recurse -Force
}
Until (!(Get-Item -Path $Destination -ErrorAction SilentlyContinue) -or ($sw.elapsed -ge $timeout))
}
else {
    #If notepad++ 7.3.3 is installed
    Write-Output "Notepad++ 7.3.3 is already Installed"
}