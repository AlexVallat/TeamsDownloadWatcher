$DownloadLocation = "$env:TEMP\TeamsDownload"

$ErrorActionPreference = 'Stop'
$DownloadLocation = $DownloadLocation -replace '\\', '/' #Electron wants paths with /
$teamsElectronAsar = "$env:LOCALAPPDATA\Microsoft\Teams\current\resources\electron.asar"
$tempFolderPath = Join-Path $env:TEMP $(New-Guid)
New-Item -Type Directory -Path $tempFolderPath | Out-Null
Push-Location $tempFolderPath
npm install asar --no-fund --no-audit --no-progress --loglevel=error
.\node_modules\.bin\asar extract $teamsElectronAsar .\electron
(Get-Content .\electron\browser\init.js) -replace 'app.setAppPath.*', "app.setPath('downloads','$DownloadLocation')`n$&" | Set-Content .\electron\browser\init.js
Rename-Item $teamsElectronAsar -NewName "electron.asar.$(Get-Date -Format FileDateTime).bak"
.\node_modules\.bin\asar pack .\electron $teamsElectronAsar
Pop-Location
Remove-Item $tempFolderPath -Force -Recurse