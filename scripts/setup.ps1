Write-Host "  - Copying files ... "
$sourcePath = "..\documents\*"
$destPath = "C:\Users\Sophos\Documents\"
Copy-Item -Force -Recurse -Verbose $sourcePath -Destination $destPath

Write-Host " - Installing Python Requirements ... "
Start-Process -FilePath py.exe -ArgumentList "-m pip install -r ..\python\requirements.txt"

Write-Host " - Starting Office 2016 Installation ... "
Write-Host " - Mounting ISO Image ... "

$imgPath = "C:\office_2016_en_US.iso"
Mount-DiskImage -ImagePath $imgPath

Write-Host " - Starting Office Install ... "
Start-Process -NoNewWindow -FilePath "D:\setup.exe" -ArgumentList "/configure ..\assets\O365Install.xml /quiet /passive"

Write-Host " - Running Office Python Setup Script ... "
Start-Process -FilePath py.exe -ArgumentList "..\python\setupOutlook.py"

