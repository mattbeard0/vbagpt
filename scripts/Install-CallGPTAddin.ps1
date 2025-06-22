param(
    [string]$RepoRoot        = (Split-Path -Parent $MyInvocation.MyCommand.Definition),
    [string]$AddinUrl        = 'https://github.com/<your-org>/excel-callgpt/raw/main/AddIn/CallGPT.xlam'
)

# Paths
$addinsFolder = Join-Path $env:APPDATA 'Microsoft\AddIns'
$addinPath    = Join-Path $addinsFolder 'CallGPT.xlam'

# Ensure AddIns folder exists
if (-not (Test-Path $addinsFolder)) {
    New-Item -Path $addinsFolder -ItemType Directory | Out-Null
}

# Download the add-in
Invoke-WebRequest -Uri $AddinUrl -OutFile $addinPath -UseBasicParsing

# Enable Trust access to VB project object model
$version = (Get-ItemProperty 'HKLM:\Software\Classes\excel.application\CurVer').'(default)' -replace 'Excel\.',''
$regPath = "HKCU:\Software\Microsoft\Office\$version\Excel\Security"
if (-not (Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}
Set-ItemProperty -Path $regPath -Name AccessVBOM -Value 1 -Type DWord

# Launch Excel COM, add references, install add-in
$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$excel.Visible       = $false

# Open the newly downloaded .xlam so we can modify its VBProject
$wb = $excel.Workbooks.Open($addinPath)

# Add Microsoft Scripting Runtime (needed for Dictionary objects and file operations)
try {
    $wb.VBProject.References.AddFromGuid("{420B2830-E718-11CF-893D-00A0C9054228}",1,0)
    Write-Host "Microsoft Scripting Runtime reference added successfully"
} catch {
    Write-Host "Microsoft Scripting Runtime reference may already be present or failed to add: $($_.Exception.Message)"
}

# Add VBA Extensibility (to allow programmatic imports)
try {
    $wb.VBProject.References.AddFromGuid("{0002E157-0000-0000-C000-000000000046}",1,0)
    Write-Host "VBA Extensibility reference added successfully"
} catch {
    Write-Host "VBA Extensibility reference may already be present or failed to add: $($_.Exception.Message)"
}

# Save and close the add-in workbook
$wb.Save()
$wb.Close()

# Register and enable the add-in
$addinObj = $excel.AddIns.Add($addinPath, $true)
$addinObj.Installed = $true

Write-Host "CallGPT add-in installed and enabled successfully!"

# Clean up
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
