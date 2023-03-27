try{ . (Join-Path (Split-Path $PSScriptRoot -Parent) "init.ps1") } catch { exit 1 }

[object]$fileClass=New-Object File
[string]$wordFileDir=$USR_CONF.conf.WordDirPath
if($fileClass.IsFile($wordFileDir) -eq $false){
    write-host "Word file directory is not found."
    exit 1
}
$wordFiles = (Get-Childitem -Recurse -File $wordFileDir)
$wordFiles = $wordFiles.fullname

foreach ($wordFile in $wordFiles) {
    write-host ("Converting file..."+$wordFile)
    $wordFile = Get-Item $wordFile

    try {
        $word = New-Object -ComObject word.application

        $doc = $word.Documents.Open($wordFile)
        $pdfpath = Join-Path $wordFile.DirectoryName ($wordFile.BaseName + ".pdf")
        $doc.SaveAs($pdfpath, [ref]17)

    }
    finally {
        $doc.Close()
        $word.Quit()
    }
}