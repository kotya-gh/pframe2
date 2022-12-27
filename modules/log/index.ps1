<#
 # [Log]テキストログの出力に関するクラス
 #
 # ログの出力を行うためのクラス。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category システム操作
 # @package なし
 #>
Class Log{
    [object]$log_format=@{
        message="";
        entry_type="Error";
        date="";
        source="App1";
        log_name="Application";
        event_id=65535;
        username="";
        hostname="";
        encoding="UTF8";
    }
    [array]$conf
    [object]$INTL

    # ログの出力形式
    # date <Application.INFORMATION> hostname username: [source="App1" eventid="65535"] message
    Log(){
        $this.conf=$script:COMP_CONF
    
        [object]$env=New-Object Env
        $this.log_format.username=$env.GetUsername()
        $this.log_format.hostname=$env.GetHostname()
        $this.log_format.source=$this.conf.component_name | Select-Object

        $this.INTL=New-Object Intl((Join-Path $PSScriptRoot "locale"))
    }

    <#
     # [Log]テキストログの書き込み
     #
     # ログをログファイルに書き込む。
     # ログ書き込み先のディレクトリが存在しない場合、コンポーネントと同名のディレクトリを作成する。作成失敗時は$falseを返す。
     # ログは日付有りのファイル、および日付なしのファイルに同内容を書き込む。ログファイル名はコンポーネント名と同名となる。
     # ログ書き込み前にログローテート処理を行う。ログローテート失敗時は$falseを返す。
     # ログ書き込み後、Log.log_format.messageを空にする。
     #
     # @access public
     # @param なし
     # @return bool ログ書き込みの成否
     # @see Log.MakeLogText, LogRotation.Rotate
     # @throws ログ書き込みで例外発生時、$falseを返す。
     #>
    [bool]WriteLog(){
        # ログ書き込みディレクトリの作成
        [object]$file=New-Object File
        [string]$logDirPath=(Join-Path ($this.conf.path_log | Select-Object) ($this.conf.component_name | Select-Object))
        if($file.IsFile($logDirPath) -eq $false){
            if($file.MkDir($logDirPath) -eq $false){
                $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_make_log_directory")
            }
        }

        # 書き込み先ログファイルパスの作成
        [string]$rotateLogPath=(($this.conf.component_name | Select-Object)+"_"+$this.GetLogNameDate()+".log")
        $rotateLogPath=(Join-Path $logDirPath $rotateLogPath)
        [string]$currentLogPath=(($this.conf.component_name | Select-Object)+".log")
        $currentLogPath=(Join-Path $logDirPath $currentLogPath)

        # ログローテート処理
        [object]$rotate=New-Object LogRotation
        if($rotate.Rotate($rotateLogPath, $currentLogPath, $this.conf.log.rotation) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_rotate_log")
            return $false
        }

        # ログ書き込み処理
        [string]$log_message=$this.MakeLogText()
        if($log_message -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_make_log_text")
            return $false 
        }
        try{
            Write-Output ($log_message) | Add-Content $rotateLogPath -Encoding $this.log_format.encoding
            Write-Output ($log_message) | Add-Content $currentLogPath -Encoding $this.log_format.encoding
                    
            # ログ書き込み後はメッセージ内容を空にする。
            $this.log_format.message=""
        } catch {
            #Write-Host($_.Exception)
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_write_log")
            return $false 
        }
        return $true
    }

    <#
     # [Log]ログフォーマットに則ったメッセージの作成
     #
     # 出力メッセージを整形する。
     # ログフォーマットに適合しない場合falseを返す。
     # eventId、logName、entryTypeのバリデーションを実施し、適合しない場合は$falseを返す。
     # Log.log_format.messageが空の場合は$falseを返す。
     # Log.log_format.messageに改行が含まれる場合は改行を削除する。
     # 次のフォーマットでログテキストを作成する。
     # date <Application.INFORMATION> hostname username: [source="App1" eventid="65535"] message
     #
     # @access public
     # @param なし
     # @return bool ログメッセージ作成の成否
     # @see LogValidation.eventId, LogValidation.logName, LogValidation.entryType, Log.GetFormattedDate
     # @throws なし
     #>
    [string]MakeLogText(){
        [object]$str=New-Object Str
        if(($this.log_format.message -ne "") -eq $false){
            return $false
        }
        [string]$part=$this.MakeLogTextPart()
        if($part -eq $false){
            return $false
        }
        $this.log_format.date=$this.GetFormattedDate()
        [string]$log_text=(
            [string]$this.log_format.date+" "+$part+" "+
            $str.TrimEol($this.log_format.message))
        return $log_text
    }

    <#
     # [Log]ログフォーマットに則ったメッセージの作成（日付、メッセージ以外）
     #
     # 出力メッセージを整形する。
     # ログフォーマットに適合しない場合falseを返す。
     # eventId、logName、entryTypeのバリデーションを実施し、適合しない場合は$falseを返す。
     # 次のフォーマットでログテキストを作成する。
     # <Application.INFORMATION> hostname username: [source="App1" eventid="65535"]
     #
     # @access public
     # @param なし
     # @return bool ログメッセージ作成の成否
     # @see LogValidation.eventId, LogValidation.logName, LogValidation.entryType, Log.GetFormattedDate
     # @throws なし
     #>
     [string]MakeLogTextPart(){
        [object]$valid=New-Object LogValidation
        if(
            (($valid.eventId($this.log_format.event_id)) -and
            ($valid.logName($this.log_format.log_name)) -and
            ($valid.entryType($this.log_format.entry_type))) -eq $false
        ){
            return $false
        }
        [string]$log_text=(
            "<"+
            $this.log_format.log_name+"."+
            $this.log_format.entry_type+"> "+
            $this.log_format.hostname+" "+
            $this.log_format.username+": [source="""+
            $this.log_format.source+""" eventid="""+
            $this.log_format.event_id+"""] ")
        return $log_text
    }

    <#
     # [Log]ログ出力用日付情報の作成
     #
     # ログ出力用日付情報を返す。
     # 次のフォーマットで日時情報を作成する。
     # yyyy-MM-dd HH:mm:ss
     #
     # @access public
     # @param なし
     # @return string ログ出力用日付情報
     # @see Get-Date
     # @throws なし
     #>
    [string]GetFormattedDate(){
        return [string](Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    }

    <#
     # [Log]ログファイル名用日付情報の作成
     #
     # ログファイル名用日付情報を返す。
     # 次のフォーマットで日付情報を作成する。
     # yyyyMMdd
     #
     # @access public
     # @param なし
     # @return string ログファイル名用日付情報
     # @see Get-Date
     # @throws なし
     #>
    [string]GetLogNameDate(){
        return [string](Get-Date -Format "yyyyMMdd")
    }
    
}

<#
 # [LogRotation]テキストログのローテートに関するクラス
 #
 # ログのローテートを行うためのクラス。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category システム操作
 # @package なし
 #>
Class LogRotation{
    LogRotation(){

    }

    <#
     # [LogRotation]ログファイル名用日付情報の作成
     #
     # ログローテーション処理を実行する。
     # $pathで示すファイルが存在しない場合（ログ書き込み時、当日の最初の書き込みである場合）、
     # ローテート対象であると判定し、カレントログの削除、指定世代以前の日付付きログファイルの削除を実施する。
     # ローテート対象ファイル削除失敗時はfalseを返す。
     #
     # @access public
     # @param string $path 日付付き最新ログファイルのフルパス
     # @param string $currentPath カレントログファイルのフルパス
     # @param int $rotation ローテート世代数
     # @return bool ローテートの成否
     # @see なし
     # @throws なし
     #>
    [bool]Rotate([string]$path, [string]$currentPath, [int]$rotation){
        [object]$file=New-Object File
        if($file.IsFile($path) -eq $true){
            return $true
        }

        # カレントログファイルの削除
        if($file.Rm($currentPath) -eq $false){
            return $false
        }

        # 指定のローテーション以前の世代のファイルを削除
        [array]$filelist=$file.Filelist($file.Dirname($path)).FullName
        # ログファイルリストを降順にソート
        $filelist = $filelist | Sort-Object -Descending
        [int]$i=0
        foreach($log in $filelist){
            if($log -match "^.*_\d{8}\.log$"){
                $i++
                if($i -ge $rotation){
                    if($file.Rm($log) -eq $false){
                        return $false
                    }
                }
            }
        }
        return $true
    }
}

<#
 # [LogValidation]テキストログのバリデーションに関するクラス
 #
 # ログのバリデーションを行うためのクラス。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category システム操作
 # @package なし
 #>
Class LogValidation{
    [array]$define_entry_type=@("N/A", "Information", "Warning", "Error", "SuccessAudit", "FailureAudit")
    [array]$define_all_log_name=@("N/A", "Application", "System", "Security", "Setup")
    [int]$define_max_event_id=65535

    LogValidation(){

    }

    <#
     # [LogValidation]イベントIDのバリデーション
     #
     # $eventIdが一定値以内であることを判定し、boolで返す。
     # $eventIdは0以上LogValidation.define_max_event_id以下とする。
     # 適合しない場合は$falseを返す。
     #
     # @access public
     # @param int $eventId イベントID
     # @return bool バリデーションの結果
     # @see なし
     # @throws なし
     #>
    [bool]eventId([int]$eventId){
        return (($eventId -ge 0) -and ($eventId -le $this.define_max_event_id))
    }

    <#
     # [LogValidation]ログネームのバリデーション
     #
     # $logNameがLogValidation.define_all_log_nameに含まれていることを判定し、boolで返す。
     # 適合しない場合は$falseを返す。
     #
     # @access public
     # @param int $logName ログネーム
     # @return bool バリデーションの結果
     # @see なし
     # @throws なし
     #>
    [bool]logName([string]$logName){
        [object]$arr=New-Object ExArray
        return $arr.ArrayIndexOf($logName, $this.define_all_log_name)
    }

    <#
     # [LogValidation]エントリータイプのバリデーション
     #
     # $entryTypeがLogValidation.define_entry_typeに含まれていることを判定し、boolで返す。
     # 適合しない場合は$falseを返す。
     #
     # @access public
     # @param int $entryType エントリータイプ
     # @return bool バリデーションの結果
     # @see なし
     # @throws なし
     #>
    [bool]entryType([string]$entryType){
        [object]$arr=New-Object ExArray
        return $arr.ArrayIndexOf($entryType, $this.define_entry_type)
    }
}
