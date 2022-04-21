Write-Host "  - Copying files ... "
$sourcePath = "$PSScriptRoot\documents\*"
$destPath = "C:\Users\Sophos\Documents\"
Copy-Item -Force -Recurse -Verbose $sourcePath -Destination $destPath

Write-Host " - Installing Python Requirements ... "
Start-Process -FilePath py.exe -ArgumentList "-m pip install -r $PSScriptRoot\python\requirements.txt"

Write-Host " - Starting Office 2016 Installation ... "
Write-Host " - Mounting ISO Image ... "

$imgPath = "C:\office_2016_en_US.iso"
Mount-DiskImage -ImagePath $imgPath

Write-Host " - Starting Office Install ... "
Start-Process -NoNewWindow -FilePath "D:\office\setup64.exe" -ArgumentList "/configure 'c:\windows_poc_config\assets\O365Install.xml'"

Write-Host " - Running Office Python Setup Script ... "
Start-Process -FilePath py.exe -ArgumentList "$PSScriptRoot\python\setupOutlook.py"

