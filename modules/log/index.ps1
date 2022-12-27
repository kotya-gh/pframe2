<#
 # [Log]�e�L�X�g���O�̏o�͂Ɋւ���N���X
 #
 # ���O�̏o�͂��s�����߂̃N���X�B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category �V�X�e������
 # @package �Ȃ�
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

    # ���O�̏o�͌`��
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
     # [Log]�e�L�X�g���O�̏�������
     #
     # ���O�����O�t�@�C���ɏ������ށB
     # ���O�������ݐ�̃f�B���N�g�������݂��Ȃ��ꍇ�A�R���|�[�l���g�Ɠ����̃f�B���N�g�����쐬����B�쐬���s����$false��Ԃ��B
     # ���O�͓��t�L��̃t�@�C���A����ѓ��t�Ȃ��̃t�@�C���ɓ����e���������ށB���O�t�@�C�����̓R���|�[�l���g���Ɠ����ƂȂ�B
     # ���O�������ݑO�Ƀ��O���[�e�[�g�������s���B���O���[�e�[�g���s����$false��Ԃ��B
     # ���O�������݌�ALog.log_format.message����ɂ���B
     #
     # @access public
     # @param �Ȃ�
     # @return bool ���O�������݂̐���
     # @see Log.MakeLogText, LogRotation.Rotate
     # @throws ���O�������݂ŗ�O�������A$false��Ԃ��B
     #>
    [bool]WriteLog(){
        # ���O�������݃f�B���N�g���̍쐬
        [object]$file=New-Object File
        [string]$logDirPath=(Join-Path ($this.conf.path_log | Select-Object) ($this.conf.component_name | Select-Object))
        if($file.IsFile($logDirPath) -eq $false){
            if($file.MkDir($logDirPath) -eq $false){
                $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_make_log_directory")
            }
        }

        # �������ݐ惍�O�t�@�C���p�X�̍쐬
        [string]$rotateLogPath=(($this.conf.component_name | Select-Object)+"_"+$this.GetLogNameDate()+".log")
        $rotateLogPath=(Join-Path $logDirPath $rotateLogPath)
        [string]$currentLogPath=(($this.conf.component_name | Select-Object)+".log")
        $currentLogPath=(Join-Path $logDirPath $currentLogPath)

        # ���O���[�e�[�g����
        [object]$rotate=New-Object LogRotation
        if($rotate.Rotate($rotateLogPath, $currentLogPath, $this.conf.log.rotation) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_rotate_log")
            return $false
        }

        # ���O�������ݏ���
        [string]$log_message=$this.MakeLogText()
        if($log_message -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_make_log_text")
            return $false 
        }
        try{
            Write-Output ($log_message) | Add-Content $rotateLogPath -Encoding $this.log_format.encoding
            Write-Output ($log_message) | Add-Content $currentLogPath -Encoding $this.log_format.encoding
                    
            # ���O�������݌�̓��b�Z�[�W���e����ɂ���B
            $this.log_format.message=""
        } catch {
            #Write-Host($_.Exception)
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_write_log")
            return $false 
        }
        return $true
    }

    <#
     # [Log]���O�t�H�[�}�b�g�ɑ��������b�Z�[�W�̍쐬
     #
     # �o�̓��b�Z�[�W�𐮌`����B
     # ���O�t�H�[�}�b�g�ɓK�����Ȃ��ꍇfalse��Ԃ��B
     # eventId�AlogName�AentryType�̃o���f�[�V���������{���A�K�����Ȃ��ꍇ��$false��Ԃ��B
     # Log.log_format.message����̏ꍇ��$false��Ԃ��B
     # Log.log_format.message�ɉ��s���܂܂��ꍇ�͉��s���폜����B
     # ���̃t�H�[�}�b�g�Ń��O�e�L�X�g���쐬����B
     # date <Application.INFORMATION> hostname username: [source="App1" eventid="65535"] message
     #
     # @access public
     # @param �Ȃ�
     # @return bool ���O���b�Z�[�W�쐬�̐���
     # @see LogValidation.eventId, LogValidation.logName, LogValidation.entryType, Log.GetFormattedDate
     # @throws �Ȃ�
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
     # [Log]���O�t�H�[�}�b�g�ɑ��������b�Z�[�W�̍쐬�i���t�A���b�Z�[�W�ȊO�j
     #
     # �o�̓��b�Z�[�W�𐮌`����B
     # ���O�t�H�[�}�b�g�ɓK�����Ȃ��ꍇfalse��Ԃ��B
     # eventId�AlogName�AentryType�̃o���f�[�V���������{���A�K�����Ȃ��ꍇ��$false��Ԃ��B
     # ���̃t�H�[�}�b�g�Ń��O�e�L�X�g���쐬����B
     # <Application.INFORMATION> hostname username: [source="App1" eventid="65535"]
     #
     # @access public
     # @param �Ȃ�
     # @return bool ���O���b�Z�[�W�쐬�̐���
     # @see LogValidation.eventId, LogValidation.logName, LogValidation.entryType, Log.GetFormattedDate
     # @throws �Ȃ�
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
     # [Log]���O�o�͗p���t���̍쐬
     #
     # ���O�o�͗p���t����Ԃ��B
     # ���̃t�H�[�}�b�g�œ��������쐬����B
     # yyyy-MM-dd HH:mm:ss
     #
     # @access public
     # @param �Ȃ�
     # @return string ���O�o�͗p���t���
     # @see Get-Date
     # @throws �Ȃ�
     #>
    [string]GetFormattedDate(){
        return [string](Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    }

    <#
     # [Log]���O�t�@�C�����p���t���̍쐬
     #
     # ���O�t�@�C�����p���t����Ԃ��B
     # ���̃t�H�[�}�b�g�œ��t�����쐬����B
     # yyyyMMdd
     #
     # @access public
     # @param �Ȃ�
     # @return string ���O�t�@�C�����p���t���
     # @see Get-Date
     # @throws �Ȃ�
     #>
    [string]GetLogNameDate(){
        return [string](Get-Date -Format "yyyyMMdd")
    }
    
}

<#
 # [LogRotation]�e�L�X�g���O�̃��[�e�[�g�Ɋւ���N���X
 #
 # ���O�̃��[�e�[�g���s�����߂̃N���X�B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category �V�X�e������
 # @package �Ȃ�
 #>
Class LogRotation{
    LogRotation(){

    }

    <#
     # [LogRotation]���O�t�@�C�����p���t���̍쐬
     #
     # ���O���[�e�[�V�������������s����B
     # $path�Ŏ����t�@�C�������݂��Ȃ��ꍇ�i���O�������ݎ��A�����̍ŏ��̏������݂ł���ꍇ�j�A
     # ���[�e�[�g�Ώۂł���Ɣ��肵�A�J�����g���O�̍폜�A�w�萢��ȑO�̓��t�t�����O�t�@�C���̍폜�����{����B
     # ���[�e�[�g�Ώۃt�@�C���폜���s����false��Ԃ��B
     #
     # @access public
     # @param string $path ���t�t���ŐV���O�t�@�C���̃t���p�X
     # @param string $currentPath �J�����g���O�t�@�C���̃t���p�X
     # @param int $rotation ���[�e�[�g���㐔
     # @return bool ���[�e�[�g�̐���
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]Rotate([string]$path, [string]$currentPath, [int]$rotation){
        [object]$file=New-Object File
        if($file.IsFile($path) -eq $true){
            return $true
        }

        # �J�����g���O�t�@�C���̍폜
        if($file.Rm($currentPath) -eq $false){
            return $false
        }

        # �w��̃��[�e�[�V�����ȑO�̐���̃t�@�C�����폜
        [array]$filelist=$file.Filelist($file.Dirname($path)).FullName
        # ���O�t�@�C�����X�g���~���Ƀ\�[�g
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
 # [LogValidation]�e�L�X�g���O�̃o���f�[�V�����Ɋւ���N���X
 #
 # ���O�̃o���f�[�V�������s�����߂̃N���X�B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category �V�X�e������
 # @package �Ȃ�
 #>
Class LogValidation{
    [array]$define_entry_type=@("N/A", "Information", "Warning", "Error", "SuccessAudit", "FailureAudit")
    [array]$define_all_log_name=@("N/A", "Application", "System", "Security", "Setup")
    [int]$define_max_event_id=65535

    LogValidation(){

    }

    <#
     # [LogValidation]�C�x���gID�̃o���f�[�V����
     #
     # $eventId�����l�ȓ��ł��邱�Ƃ𔻒肵�Abool�ŕԂ��B
     # $eventId��0�ȏ�LogValidation.define_max_event_id�ȉ��Ƃ���B
     # �K�����Ȃ��ꍇ��$false��Ԃ��B
     #
     # @access public
     # @param int $eventId �C�x���gID
     # @return bool �o���f�[�V�����̌���
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]eventId([int]$eventId){
        return (($eventId -ge 0) -and ($eventId -le $this.define_max_event_id))
    }

    <#
     # [LogValidation]���O�l�[���̃o���f�[�V����
     #
     # $logName��LogValidation.define_all_log_name�Ɋ܂܂�Ă��邱�Ƃ𔻒肵�Abool�ŕԂ��B
     # �K�����Ȃ��ꍇ��$false��Ԃ��B
     #
     # @access public
     # @param int $logName ���O�l�[��
     # @return bool �o���f�[�V�����̌���
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]logName([string]$logName){
        [object]$arr=New-Object ExArray
        return $arr.ArrayIndexOf($logName, $this.define_all_log_name)
    }

    <#
     # [LogValidation]�G���g���[�^�C�v�̃o���f�[�V����
     #
     # $entryType��LogValidation.define_entry_type�Ɋ܂܂�Ă��邱�Ƃ𔻒肵�Abool�ŕԂ��B
     # �K�����Ȃ��ꍇ��$false��Ԃ��B
     #
     # @access public
     # @param int $entryType �G���g���[�^�C�v
     # @return bool �o���f�[�V�����̌���
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]entryType([string]$entryType){
        [object]$arr=New-Object ExArray
        return $arr.ArrayIndexOf($entryType, $this.define_entry_type)
    }
}
