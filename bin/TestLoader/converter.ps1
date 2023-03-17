try{ . (Join-Path (Split-Path $PSScriptRoot -Parent) "init.ps1") } catch { exit 1 }

# Excel�t�@�C���ƃV�[�g��
[object]$fileClass=New-Object File
[string]$excelFile=$USR_CONF.converterConf.excelFile
if($fileClass.IsFile($excelFile) -eq $false){
    $excelFile=(Join-Path -Path $INIT.GetScriptRootPath() -ChildPath "files" | Join-Path -ChildPath $USR_CONF.converterConf.excelFile)
    if($fileClass.IsFile($excelFile) -eq $false){
        write-host "Excel file is not found."
    }
}
$sheetName = $USR_CONF.converterConf.sheetName

# Excel�A�v���P�[�V�����I�u�W�F�N�g�쐬
$excel = New-Object -ComObject Excel.Application

# �t�@�C���I�[�v��
$workbook = $excel.Workbooks.Open($excelFile)

# �V�[�g�I��
$sheet = $workbook.Worksheets.Item($sheetName)

# �\�f�[�^�擾
$table = $sheet.UsedRange.Value(10) | ForEach-Object -Begin { $i = 0 } -Process {
    if ($i++ % 7 -eq 0) {
        $tmp = @()
    }
    $tmp += $_
    if ($i % 7 -eq 0) {
        ,$tmp
    }
}

# �t�@�C���N���[�Y
$workbook.Close()

# Excel�A�v���P�[�V�����I��
$excel.Quit()

# COM�I�u�W�F�N�g���
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# �\�f�[�^����w�b�_�[�s�擾
$headers = $table[0]

# �\�f�[�^����{�f�B�s�擾�i�w�b�_�[�s�����j
$body = $table[1..($table.Length - 1)]

# �{�f�B�s����testId���ƂɃO���[�v��
$groups = $body | Group-Object -Property {$_[0]}

# JSON�`���p�̃I�u�W�F�N�g�쐬�iTestConfigure�z��j
$jsonObj = @{"EvidenceHomeDir"=""; "TestConfigure"=@()}

foreach ($group in $groups) {
    # testId���Ƃ�testItems�z��쐬
    $testItems = @()
    # testNo���ƂɃO���[�v������testCommands�z��쐬
    foreach ($subgroup in ($group.Group | Group-Object -Property {$_[1]})) {
        # testNo���Ƃ�testCommands�z��쐬�iorder, command, returnCode, returnMsg�j
        $testCommands = @()
        foreach ($row in $subgroup.Group) {
            if ([String]::IsNullOrEmpty($row[6])) {
                [string]$rtmsg=""
            }else{
                [string]$rtmsg=$row[6]
            }
            # order, command, returnCode, returnMsg ���n�b�V���e�[�u���ō쐬���Ĕz��ɒǉ�
            $testCommands += @{
                "order"=$row[3];
                "command"=$row[4];
                "returnCode"=$row[5];
                "returnMsg"=$rtmsg
            }
        }
        # testNo���Ƃ�hostname��testCommands ���n�b�V���e�[�u���ō쐬���Ĕz��ɒǉ� 
        $testItems += @{
            "testNo"=$subgroup.Name;
            "hostname"=$row[2];
            "testCommands"=$testCommands;
        }
    }
    # testId���Ƃ�TestConfigure �z��֒ǉ� 
    $jsonObj.TestConfigure += @{
        "testId"=$group.Name;
        "testItems"=$testItems
    }
}
[string]$DateStr = (Get-Date).ToString("yyyyMMddHHmmss")
[string]$outputpath=(Join-Path -Path $INIT.GetScriptRootPath() -ChildPath "files" | Join-Path -ChildPath ("list.json_"+$DateStr))
Set-Content -Path $outputpath -Value (convertto-json $jsonObj -Depth 6)