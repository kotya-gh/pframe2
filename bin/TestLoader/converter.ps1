try{ . (Join-Path (Split-Path $PSScriptRoot -Parent) "init.ps1") } catch { exit 1 }

# Excelファイルとシート名
[object]$fileClass=New-Object File
[string]$excelFile=$USR_CONF.converterConf.excelFile
if($fileClass.IsFile($excelFile) -eq $false){
    $excelFile=(Join-Path -Path $INIT.GetScriptRootPath() -ChildPath "files" | Join-Path -ChildPath $USR_CONF.converterConf.excelFile)
    if($fileClass.IsFile($excelFile) -eq $false){
        write-host "Excel file is not found."
    }
}
$sheetName = $USR_CONF.converterConf.sheetName

# Excelアプリケーションオブジェクト作成
$excel = New-Object -ComObject Excel.Application

# ファイルオープン
$workbook = $excel.Workbooks.Open($excelFile)

# シート選択
$sheet = $workbook.Worksheets.Item($sheetName)

# 表データ取得
$table = $sheet.UsedRange.Value(10) | ForEach-Object -Begin { $i = 0 } -Process {
    if ($i++ % 7 -eq 0) {
        $tmp = @()
    }
    $tmp += $_
    if ($i % 7 -eq 0) {
        ,$tmp
    }
}

# ファイルクローズ
$workbook.Close()

# Excelアプリケーション終了
$excel.Quit()

# COMオブジェクト解放
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# 表データからヘッダー行取得
$headers = $table[0]

# 表データからボディ行取得（ヘッダー行除く）
$body = $table[1..($table.Length - 1)]

# ボディ行からtestIdごとにグループ化
$groups = $body | Group-Object -Property {$_[0]}

# JSON形式用のオブジェクト作成（TestConfigure配列）
$jsonObj = @{"EvidenceHomeDir"=""; "TestConfigure"=@()}

foreach ($group in $groups) {
    # testIdごとにtestItems配列作成
    $testItems = @()
    # testNoごとにグループ化してtestCommands配列作成
    foreach ($subgroup in ($group.Group | Group-Object -Property {$_[1]})) {
        # testNoごとにtestCommands配列作成（order, command, returnCode, returnMsg）
        $testCommands = @()
        foreach ($row in $subgroup.Group) {
            if ([String]::IsNullOrEmpty($row[6])) {
                [string]$rtmsg=""
            }else{
                [string]$rtmsg=$row[6]
            }
            # order, command, returnCode, returnMsg をハッシュテーブルで作成して配列に追加
            $testCommands += @{
                "order"=$row[3];
                "command"=$row[4];
                "returnCode"=$row[5];
                "returnMsg"=$rtmsg
            }
        }
        # testNoごとにhostnameとtestCommands をハッシュテーブルで作成して配列に追加 
        $testItems += @{
            "testNo"=$subgroup.Name;
            "hostname"=$row[2];
            "testCommands"=$testCommands;
        }
    }
    # testIdごとにTestConfigure 配列へ追加 
    $jsonObj.TestConfigure += @{
        "testId"=$group.Name;
        "testItems"=$testItems
    }
}
[string]$DateStr = (Get-Date).ToString("yyyyMMddHHmmss")
[string]$outputpath=(Join-Path -Path $INIT.GetScriptRootPath() -ChildPath "files" | Join-Path -ChildPath ("list.json_"+$DateStr))
Set-Content -Path $outputpath -Value (convertto-json $jsonObj -Depth 6)