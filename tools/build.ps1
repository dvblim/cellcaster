<#
.SYNOPSIS
    Builds CellCaster.xlsm from the modules in src/.

.DESCRIPTION
    Starts Excel in the background, imports every src/*.bas into a new
    workbook, runs InstallMenu to create the Menu sheet and saves the result
    as a macro-enabled workbook in the repository root.

    Excel must allow scripts to touch VBA projects. Enable it once in
    File > Options > Trust Center > Trust Center Settings > Macro Settings >
    "Trust access to the VBA project object model".

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\build.ps1
#>
param([string]$Output = "CellCaster.xlsm")

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$outPath = Join-Path $root $Output
if (Test-Path $outPath) { Remove-Item $outPath }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null

try {
    $workbook = $excel.Workbooks.Add()

    $project = $null
    try { $project = $workbook.VBProject } catch { }
    if ($null -eq $project) {
        throw ("Excel blocked access to the VBA project. Enable 'Trust access to the VBA " +
               "project object model' in File > Options > Trust Center > Trust Center " +
               "Settings > Macro Settings, then run this script again.")
    }

    Get-ChildItem (Join-Path $root "src") -Filter *.bas | Sort-Object Name | ForEach-Object {
        [void]$project.VBComponents.Import($_.FullName)
        Write-Host "imported $($_.Name)"
    }

    [void]$excel.Run("InstallMenu")
    $workbook.SaveAs($outPath, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled
    Write-Host "built $outPath"
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
