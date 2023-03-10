try{ . (Join-Path (Split-Path $PSScriptRoot -Parent) "init.ps1") } catch { exit 1 }

[object]$strClass=New-Object Str

# エビデンス格納フォルダの作成
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

# ホスト名の取得
[object]$envClass=New-Object Env
[string]$hostname=$envClass.GetHostname()

# テストIDでループ
foreach($testConfigure in $USR_CONF.list.TestConfigure){
    # 格納フォルダの作成
    [string]$testIdDir=(Join-Path $EvidenceHomeDir $testConfigure.testId)
    if($fileClass.IsFile($testIdDir) -eq $false){
        $fileClass.MkDir($testIdDir)
    }

    # テスト開始メッセージ
    Write-Host $INTL.FormattedMessage("Starting_test_id", @{id=$testConfigure.testId})

    # テスト項目でループ
    foreach($testItem in $testConfigure.testItems){
        # 簡易チェック用変数
        [bool]$CheckResult=$true

        # エビデンス記録用データのオブジェクトを初期化
        [array]$testResult=@()

        # ホスト名が一致しない場合は処理をスキップ
        if($hostname -ne $testItem.hostname){
            continue
        }

        # テスト用コマンドをオーダー順にソート
        [array]$sortTestCommands=($testItem.testCommands | Sort-Object order)

        # テスト番号を0詰めで4桁表記とする
        [string]$testNo=$testItem.testNo.ToString().PadLeft(4,"0")

        # エビデンスを記録するテキストファイルとJSONファイルのパスを生成する無名関数
        $GenerateEvidenceFilePath = {
            param($testNo, $hostname, $testIdDir)
            # 現在日時を文字列に変換
            [string]$startDateStr = (Get-Date).ToString("yyyyMMddHHmmss")
            # ファイル名を結合
            [string]$fileName = "$testNo`_$hostname`_$startDateStr"
            # ファイルパスを返す
            return (Join-Path $testIdDir $fileName)
        }
        [string]$evidenceFileName = &$GenerateEvidenceFilePath $testNo $hostname $testIdDir
        [string]$evidenceJsonFileName = &$GenerateEvidenceFilePath $testNo $hostname $testIdDir

        # コマンド実行前メッセージ
        Write-Host $INTL.FormattedMessage("Starting_test_item_id", @{id=$testNo; hostname=$hostname})

        # テスト用コマンドを実行
        foreach($command in $sortTestCommands){
            # エビデンス記録用データのオブジェクトを初期化
            [object]$testResultObject=@{}

            $OutputMessage = {
                param($message, $filePath)
                write-host $message
                Add-Content -Path $filePath -Value $message
            }

            $testResultObject.Add("No", $command.order)

            # 実行コマンドの説明
            $testResultObject.Add("Command", $command.command)
            $testResultObject.Add("ExceptedResult", $command.returnMsg)
            $testResultObject.Add("ExceptedReturncode", $command.returnCode)
            &$OutputMessage $INTL.FormattedMessage("Running_command", @{command=$command.command}) $evidenceFileName
            if($command.returnMsg -ne ""){
                &$OutputMessage $INTL.FormattedMessage("Command_result_message", @{returnMsg=$command.returnMsg}) $evidenceFileName
            }
            &$OutputMessage $INTL.FormattedMessage("Command_returncode", @{returncode=$command.returnCode}) $evidenceFileName

            # コマンド実行時刻を取得
            [datetime]$startDate=Get-Date
            &$OutputMessage $INTL.FormattedMessage("Start_datetime", @{datetime=$startDate}) $evidenceFileName
            $testResultObject.Add("StartDatetime", $startDate.ToString("yyyy/MM/dd HH:mm:ss"))

            # コマンド実行
            [string]$execCommand=$command.command+";$LastExitCode"
            [array]$result=(Invoke-Expression $execCommand)

            # コマンド実行終了時刻を取得
            [datetime]$endDate=Get-Date
            &$OutputMessage $INTL.FormattedMessage("End_datetime", @{datetime=$endDate}) $evidenceFileName
            $testResultObject.Add("StopDatetime", $endDate.ToString("yyyy/MM/dd HH:mm:ss"))
            
            # コマンド実行結果を出力
            &$OutputMessage $INTL.FormattedMessage("Command_result") $evidenceFileName
            [array]$resultOutput=($result | Select-Object -skiplast 1)

            [string]$resultOutputString=($resultOutput -join "`n")
            $testResultObject.Add("CommandResult", $resultOutputString)
            &$OutputMessage $resultOutputString $evidenceFileName

            &$OutputMessage $INTL.FormattedMessage("Command_return_code", @{returncode=$result[-1]}) $evidenceFileName
            $testResultObject.Add("Returncode", $result[-1])

            # 簡易チェック
            [bool]$tmpResult = ($result[-1] -eq $command.returnCode) -and (($command.returnMsg -eq "") -or ($strClass.Strpos($resultOutputString, $command.returnMsg) -ne $false))
            if($tmpResult -eq $false){
                $CheckResult=$tmpResult
            }
            $testResultObject.Add("CheckResult", $tmpResult)

            # 改行
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