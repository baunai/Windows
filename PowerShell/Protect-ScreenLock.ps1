param($minutes = 720)

[void] [System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”)

$myshell = New-Object -com “Wscript.Shell”

for ($i = 0; $i -lt $minutes; $i++) {
Start-Sleep -Seconds 60
# $myshell.sendkeys(“.”)
$myshell.sendkeys(“{NUMLOCK}{NUMLOCK}”)
Write-Host "NUMLOCK Key press..." -ForegroundColor Cyan
$Pos = [System.Windows.Forms.Cursor]::Position
[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point((($Pos.X) + 1) , $Pos.Y)
}
Write-Host "Cursor move....." -ForegroundColor Magenta