# Run export macro from Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open("C:\Path\To\ExportMacros.xlsm")
$excel.Run("ExportAllVBAModules")
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Change directory to macro repo
Set-Location "C:\Macro_Repository"

# Add all files to staging
git add .

# Commit with timestamp
$timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
git commit -m "Auto-update macros on $timeStamp" 2>$null

# Push to GitHub
git push origin main
