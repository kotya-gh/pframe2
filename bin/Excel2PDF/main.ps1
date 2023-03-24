try{ . (Join-Path (Split-Path $PSScriptRoot -Parent) "init.ps1") } catch { exit 1 }

[object]$fileClass=New-Object File
[string]$excelFileDir=$USR_CONF.conf.ExcelDirPath
if($fileClass.IsFile($excelFileDir) -eq $false){
    write-host "Excel file directory is not found."
    exit 1
}
$excelFiles = (Get-Childitem -Recurse -File $excelFileDir)
$excelFiles = $excelFiles.fullname

foreach ($excelFile in $excelFiles) {
    write-host ("Converting file..."+$excelFile)
    $excelFile = Get-Item $excelFile

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $wb = $excel.Workbooks.Open($excelFile)

        $pdfpath = Join-Path $excelFile.DirectoryName ($excelFile.BaseName + ".pdf")
        $wb.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF, $pdfpath)

        $wb.Close()
        $excel.Quit()
    }
    finally {
        $sheet, $wb, $excel | ForEach-Object {
            if ($_ -ne $null) {
                [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
            }
        }
    }
}