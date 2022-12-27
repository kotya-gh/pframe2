<#
 # [Mail]���[�����M�p�N���X
 #
 # ���[�����M�Ɋւ��鏈�����`����B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category ���[�����M�p����
 # @package �Ȃ�
 #>
Class Mail{
    [object]$mailInfo=@{
        from="";
        subject="";
        body="";
        file=@();
        to="";
        mailserveraddr="";
        mailserverport=25;
        user="";
        password=""
        ssl=$false
    }
    [object]$LOCALE

    Mail(){
        $this.LOCALE=New-Object Intl((Join-Path $PSScriptRoot "locale"))
    }

    <#
     # [Mail]���[���̑��M
     #
     # Mail.mailInfo�Ɏw�肷����e�Ń��[���𑗐M����BMail.mailInfo�̃����o�����Ɏ����B
     # from�F���M�����[���A�h���X
     # subject�F���[���^�C�g��
     # body�F���[���{��
     # file�F�Y�t�t�@�C��(�t�@�C���̃t���p�X��z��Ɋi�[����)
     # to�F���M�惁�[���A�h���X
     # mailserveraddr�FSMTP�T�[�oIP�A�h���X
     # mailserverport�FSMTP�T�[�o�|�[�g�ԍ�(�W����25�ԃ|�[�g��ݒ肷��)
     # user�FSTMP�F�ؗp���[�U�A�J�E���g
     # password�FSTMP�F�ؗp�p�X���[�h
     # ssl�FSSL�ʐM�ݒ�($true|$false)
     # ���[�����M�p���̃o���f�[�V�����Ɏ��s�����ꍇ�A�܂��̓��[�����M�Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param �Ȃ�
     # @return bool ���[�����M����
     # @see Net.Mail.SmtpClient, Net.NetworkCredential
     # @throws ���[�����M�ŗ�O�������A$false��Ԃ��B
     #>
    [bool]SendMail(){
        # �o���f�[�V��������(���M�����[���A�h���X)
        if($this.mailInfo.from -eq ""){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Invalid_source_mail_address_format")
            return $false
        }

        # �o���f�[�V��������(���M�����[���A�h���X)
        if($this.mailInfo.to -eq ""){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Invalid_destination_mail_address_format")
            return $false
        }

        # �o���f�[�V��������(���[���^�C�g��)
        if($this.validateSubject($this.mailInfo.subject) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Invalid_subject_format")
            return $false
        }

        # �o���f�[�V��������(���[���{��)
        if($this.mailInfo.body -eq ""){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Mail_body_is_null.")
            return $false
        }

        # �o���f�[�V��������(SMTP�T�[�oIP�A�h���X)
        [object]$hosts=New-Object Hosts
        if(
            ($hosts.ValidateHostname($this.mailInfo.mailserveraddr) -eq $false) -and 
            ($hosts.ValidateIPaddress($this.mailInfo.mailserveraddr) -eq $false))
        {
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Invalid_ipaddress_format")
            return $false
        }

        # �o���f�[�V��������(�|�[�g�ԍ�)
        [object]$netcom=New-Object NetCom
        if($netcom.validatePort($this.mailInfo.mailserverport) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Invalid_port_number")
            return $false
        }

        # �o���f�[�V��������(STMP�F�ؗp���[�U�A�J�E���g)
        if($this.validateUsername($this.mailInfo.user) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Invalid_smtp_user_format")
            return $false
        }

        # �o���f�[�V��������(�Y�t�t�@�C��)
        [object]$file=New-Object File
        if(($this.mailInfo.file.Length -gt 0) -and ($file.IsFile($this.mailInfo.file) -eq $false)){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("No_file_exist")
            return $false
        }

        # �o���f�[�V��������(ssl�ʐM�ۂ̃t���O)
        if($this.validateEnableSsl($this.mailInfo.ssl) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Invalid_ssl_flag")
            return $false
        }

        try{
            # ���M���[���T�[�o�[�̐ݒ�
            [object]$SMTPClient=New-Object Net.Mail.SmtpClient($this.mailInfo.mailserveraddr, $this.mailInfo.mailserverport)
            # SSL�Í����ʐM���Ȃ� $false
            $SMTPClient.EnableSsl=$this.mailInfo.ssl
            $SMTPClient.Credentials=New-Object Net.NetworkCredential($this.mailInfo.user, $this.mailInfo.password)
            
            # ���[�����b�Z�[�W�̍쐬
            [object]$MailMassage=New-Object Net.Mail.MailMessage(
                $this.mailInfo.from,
                $this.mailInfo.to,
                $this.mailInfo.subject,
                $this.mailInfo.body)

            # �Y�t�t�@�C���̏���
            if($this.mailInfo.file.Length -gt 0){
                # �t�@�C������Y�t�t�@�C�����쐬
                [object]$Attachment=New-Object Net.Mail.Attachment($this.mailInfo.file)
                # ���[�����b�Z�[�W�ɓY�t
                $MailMassage.Attachments.Add($Attachment)
            }

            # ���[�����b�Z�[�W�𑗐M
            $SMTPClient.Send($MailMassage)
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $true
    }

    <#
     # [Mail]STMP�F�ؗp���[�U�A�J�E���g�̃o���f�[�V����
     #
     # $username�Ŏw�肷�镶����unix/linux�̃��[�U�A�J�E���g�K��ilibc6�j�ɑ����Ă��邩�m�F����B�K��ɂ��Ĉȉ��Ɏ����B
     # �����A�A���t�@�x�b�g�A�A���_�[�o�[���g�p�\
     # �E�啶���͎g�p�s��
     # �E�擪�����͐����͎g�p�s��
     # �E���͒l����̏ꍇ��$false��Ԃ��B
     # ���[�U�A�J�E���g�̏������s���ł���ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param string $username STMP�F�ؗp���[�U�A�J�E���g
     # @return bool ���[�U�A�J�E���g�̃o���f�[�V��������
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]validateUsername([string]$username){
        if($username -eq ""){
            return $false
        }
        [regex]$reg="(^[a-z_\-][a-z0-9_\-]{0,30}$)"
        return ($username -match $reg)
    }

    <#
     # [Mail]ssl�ʐM�ۂ̃t���O�̃o���f�[�V����
     #
     # $ssl�Ŏw�肷����͒l��bool�ł��邩�ǂ������m�F����
     # ssl�ʐM�ۂ̃t���O�̏������s���ł���ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param bool $ssl ssl�ʐM�ۂ̃t���O
     # @return bool ssl�ʐM�ۂ̃t���O�̃o���f�[�V��������
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]validateEnableSsl([bool]$ssl){
        return ($ssl -is "bool")
    }

    <#
     # [Mail]���[���^�C�g���̃o���f�[�V����
     #
     # $subject�Ŏw�肷�镶����rfc2822�ɏ����������[���^�C�g���i998byte�ȓ��̕�����j�ł��邩�ǂ������m�F����B
     # ���[���^�C�g���̃o���f�[�V�������ʂ��s���ł���ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param string $subject ���[���^�C�g��
     # @return bool ���[���^�C�g���̃o���f�[�V��������
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]validateSubject([string]$subject){
        $int_byte_num = [System.Text.Encoding]::GetEncoding("shift_jis").GetByteCount($subject)
        return (($int_byte_num -ge 0) -and ($int_byte_num -le 998))
    }
}