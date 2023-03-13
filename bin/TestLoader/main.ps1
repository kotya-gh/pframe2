try{ . (Join-Path (Split-Path $PSScriptRoot -Parent) "init.ps1") } catch { exit 1 }

[object]$strClass=New-Object Str

# �G�r�f���X�i�[�t�H���_�̍쐬
[object]$fileClass=New-Object File
[string]$EvidenceHomeDir=""
if($USR_CONF.list.EvidenceHomeDir -eq ""){
    $EvidenceHomeDir=(Join-Path $INIT.GetFilesPath() $INIT.GetComponentName())
}elseif($fileClass.IsFile($USR_CONF.list.EvidenceHomeDir) -eq $false){
    $EvidenceHomeDir=$USR_CONF.list.EvidenceHomeDir
}
if($fileClass.IsFile($EvidenceHomeDir) -eq $false){
    $fileClass.MkDir($EvidenceHomeDir)
}

# �z�X�g���̎擾
[object]$envClass=New-Object Env
[string]$hostname=$envClass.GetHostname()

# �e�X�gID�Ń��[�v
foreach($testConfigure in $USR_CONF.list.TestConfigure){
    # �i�[�t�H���_�̍쐬
    [string]$testIdDir=(Join-Path $EvidenceHomeDir $testConfigure.testId)
    if($fileClass.IsFile($testIdDir) -eq $false){
        $fileClass.MkDir($testIdDir)
    }

    # �e�X�g�J�n���b�Z�[�W
    Write-Host $INTL.FormattedMessage("Starting_test_id", @{id=$testConfigure.testId})

    # �e�X�g���ڂŃ��[�v
    foreach($testItem in $testConfigure.testItems){
        # �ȈՃ`�F�b�N�p�ϐ�
        [bool]$CheckResult=$true

        # �G�r�f���X�L�^�p�f�[�^�̃I�u�W�F�N�g��������
        [array]$testResult=@()

        # �z�X�g������v���Ȃ��ꍇ�͏������X�L�b�v
        if($hostname -ne $testItem.hostname){
            continue
        }

        # �e�X�g�p�R�}���h���I�[�_�[���Ƀ\�[�g
        [array]$sortTestCommands=($testItem.testCommands | Sort-Object order)

        # �e�X�g�ԍ���0�l�߂�4���\�L�Ƃ���
        [string]$testNo=$testItem.testNo.ToString().PadLeft(4,"0")

        # �G�r�f���X���L�^����e�L�X�g�t�@�C����JSON�t�@�C���̃p�X�𐶐����閳���֐�
        $GenerateEvidenceFilePath = {
            param($testNo, $hostname, $testIdDir)
            # ���ݓ����𕶎���ɕϊ�
            [string]$startDateStr = (Get-Date).ToString("yyyyMMddHHmmss")
            # �t�@�C����������
            [string]$fileName = "$testNo`_$hostname`_$startDateStr"
            # �t�@�C���p�X��Ԃ�
            return (Join-Path $testIdDir $fileName)
        }
        [string]$evidenceFileName = &$GenerateEvidenceFilePath $testNo $hostname $testIdDir
        [string]$evidenceJsonFileName = &$GenerateEvidenceFilePath $testNo $hostname $testIdDir

        # �R�}���h���s�O���b�Z�[�W
        Write-Host $INTL.FormattedMessage("Starting_test_item_id", @{id=$testNo; hostname=$hostname})

        # �e�X�g�p�R�}���h�����s
        foreach($command in $sortTestCommands){
            # �G�r�f���X�L�^�p�f�[�^�̃I�u�W�F�N�g��������
            [object]$testResultObject=@{}

            $OutputMessage = {
                param($message, $filePath)
                write-host $message
                Add-Content -Path $filePath -Value $message
            }

            $testResultObject.Add("No", $command.order)

            # ���s�R�}���h�̐���
            $testResultObject.Add("Command", $command.command)
            $testResultObject.Add("ExceptedResult", $command.returnMsg)
            $testResultObject.Add("ExceptedReturncode", $command.returnCode)
            &$OutputMessage $INTL.FormattedMessage("Running_command", @{command=$command.command}) $evidenceFileName
            if($command.returnMsg -ne ""){
                &$OutputMessage $INTL.FormattedMessage("Command_result_message", @{returnMsg=$command.returnMsg}) $evidenceFileName
            }
            &$OutputMessage $INTL.FormattedMessage("Command_returncode", @{returncode=$command.returnCode}) $evidenceFileName

            # �R�}���h���s�������擾
            [datetime]$startDate=Get-Date
            &$OutputMessage $INTL.FormattedMessage("Start_datetime", @{datetime=$startDate}) $evidenceFileName
            $testResultObject.Add("StartDatetime", $startDate.ToString("yyyy/MM/dd HH:mm:ss"))

            # �R�}���h���s
            [string]$execCommand=$command.command+";$LastExitCode"
            [array]$result=(Invoke-Expression $execCommand)

            # �R�}���h���s�I���������擾
            [datetime]$endDate=Get-Date
            &$OutputMessage $INTL.FormattedMessage("End_datetime", @{datetime=$endDate}) $evidenceFileName
            $testResultObject.Add("StopDatetime", $endDate.ToString("yyyy/MM/dd HH:mm:ss"))
            
            # �R�}���h���s���ʂ��o��
            &$OutputMessage $INTL.FormattedMessage("Command_result") $evidenceFileName
            [array]$resultOutput=($result | Select-Object -skiplast 1)

            [string]$resultOutputString=($resultOutput -join "`n")
            $testResultObject.Add("CommandResult", $resultOutputString)
            &$OutputMessage $resultOutputString $evidenceFileName

            &$OutputMessage $INTL.FormattedMessage("Command_return_code", @{returncode=$result[-1]}) $evidenceFileName
            $testResultObject.Add("Returncode", $result[-1])

            # �ȈՃ`�F�b�N
            if($result[-1] -ne $command.returnCode){
                $CheckResult=$false
            }
            if(($command.returnMsg -ne "") -and ($strClass.Strpos($resultOutputString, $command.returnMsg) -eq $false)){
                $CheckResult=$false
            }
            $testResultObject.Add("CheckResult", $CheckResult)

            # ���s
            &$OutputMessage "`n" $evidenceFileName
            $testResult+=$testResultObject
        }
        [string]$suffix = if($CheckResult){"_true"}else{"_false"}
        if($fileClass.Move($evidenceFileName, $evidenceFileName+$suffix+".txt")){
            $evidenceJsonFileName=$evidenceJsonFileName+$suffix+".json"
            Add-Content -Path $evidenceJsonFileName -Value (ConvertTo-Json $testResult -Depth 5)      
        }
    }
}