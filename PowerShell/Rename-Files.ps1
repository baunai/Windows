Get-ChildItem -Path 'D:\MusicFolder' -Filter "1412 *.mp3" -Recurse | Rename-Item -NewName {$_.Name -replace '1412 ',''}
