<#
 # [Mail]メール送信用クラス
 #
 # メール送信に関する処理を定義する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category メール送信用処理
 # @package なし
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
     # [Mail]メールの送信
     #
     # Mail.mailInfoに指定する内容でメールを送信する。Mail.mailInfoのメンバを次に示す。
     # from：送信元メールアドレス
     # subject：メールタイトル
     # body：メール本文
     # file：添付ファイル(ファイルのフルパスを配列に格納する)
     # to：送信先メールアドレス
     # mailserveraddr：SMTPサーバIPアドレス
     # mailserverport：SMTPサーバポート番号(標準で25番ポートを設定する)
     # user：STMP認証用ユーザアカウント
     # password：STMP認証用パスワード
     # ssl：SSL通信設定($true|$false)
     # メール送信用情報のバリデーションに失敗した場合、またはメール送信に失敗した場合、$falseを返す。
     #
     # @access public
     # @param なし
     # @return bool メール送信成否
     # @see Net.Mail.SmtpClient, Net.NetworkCredential
     # @throws メール送信で例外発生時、$falseを返す。
     #>
    [bool]SendMail(){
        # バリデーション処理(送信元メールアドレス)
        if($this.mailInfo.from -eq ""){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Invalid_source_mail_address_format")
            return $false
        }

        # バリデーション処理(送信元メールアドレス)
        if($this.mailInfo.to -eq ""){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Invalid_destination_mail_address_format")
            return $false
        }

        # バリデーション処理(メールタイトル)
        if($this.validateSubject($this.mailInfo.subject) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Invalid_subject_format")
            return $false
        }

        # バリデーション処理(メール本文)
        if($this.mailInfo.body -eq ""){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Mail_body_is_null.")
            return $false
        }

        # バリデーション処理(SMTPサーバIPアドレス)
        [object]$hosts=New-Object Hosts
        if(
            ($hosts.ValidateHostname($this.mailInfo.mailserveraddr) -eq $false) -and 
            ($hosts.ValidateIPaddress($this.mailInfo.mailserveraddr) -eq $false))
        {
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Invalid_ipaddress_format")
            return $false
        }

        # バリデーション処理(ポート番号)
        [object]$netcom=New-Object NetCom
        if($netcom.validatePort($this.mailInfo.mailserverport) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Invalid_port_number")
            return $false
        }

        # バリデーション処理(STMP認証用ユーザアカウント)
        if($this.validateUsername($this.mailInfo.user) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Invalid_smtp_user_format")
            return $false
        }

        # バリデーション処理(添付ファイル)
        [object]$file=New-Object File
        if(($this.mailInfo.file.Length -gt 0) -and ($file.IsFile($this.mailInfo.file) -eq $false)){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("No_file_exist")
            return $false
        }

        # バリデーション処理(ssl通信可否のフラグ)
        if($this.validateEnableSsl($this.mailInfo.ssl) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Invalid_ssl_flag")
            return $false
        }

        try{
            # 送信メールサーバーの設定
            [object]$SMTPClient=New-Object Net.Mail.SmtpClient($this.mailInfo.mailserveraddr, $this.mailInfo.mailserverport)
            # SSL暗号化通信しない $false
            $SMTPClient.EnableSsl=$this.mailInfo.ssl
            $SMTPClient.Credentials=New-Object Net.NetworkCredential($this.mailInfo.user, $this.mailInfo.password)
            
            # メールメッセージの作成
            [object]$MailMassage=New-Object Net.Mail.MailMessage(
                $this.mailInfo.from,
                $this.mailInfo.to,
                $this.mailInfo.subject,
                $this.mailInfo.body)

            # 添付ファイルの処理
            if($this.mailInfo.file.Length -gt 0){
                # ファイルから添付ファイルを作成
                [object]$Attachment=New-Object Net.Mail.Attachment($this.mailInfo.file)
                # メールメッセージに添付
                $MailMassage.Attachments.Add($Attachment)
            }

            # メールメッセージを送信
            $SMTPClient.Send($MailMassage)
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $true
    }

    <#
     # [Mail]STMP認証用ユーザアカウントのバリデーション
     #
     # $usernameで指定する文字列がunix/linuxのユーザアカウント規約（libc6）に則っているか確認する。規約について以下に示す。
     # 数字、アルファベット、アンダーバーを使用可能
     # ・大文字は使用不可
     # ・先頭文字は数字は使用不可
     # ・入力値が空の場合は$falseを返す。
     # ユーザアカウントの書式が不正である場合、$falseを返す。
     #
     # @access public
     # @param string $username STMP認証用ユーザアカウント
     # @return bool ユーザアカウントのバリデーション成否
     # @see なし
     # @throws なし
     #>
    [bool]validateUsername([string]$username){
        if($username -eq ""){
            return $false
        }
        [regex]$reg="(^[a-z_\-][a-z0-9_\-]{0,30}$)"
        return ($username -match $reg)
    }

    <#
     # [Mail]ssl通信可否のフラグのバリデーション
     #
     # $sslで指定する入力値がboolであるかどうかを確認する
     # ssl通信可否のフラグの書式が不正である場合、$falseを返す。
     #
     # @access public
     # @param bool $ssl ssl通信可否のフラグ
     # @return bool ssl通信可否のフラグのバリデーション成否
     # @see なし
     # @throws なし
     #>
    [bool]validateEnableSsl([bool]$ssl){
        return ($ssl -is "bool")
    }

    <#
     # [Mail]メールタイトルのバリデーション
     #
     # $subjectで指定する文字列がrfc2822に準拠したメールタイトル（998byte以内の文字列）であるかどうかを確認する。
     # メールタイトルのバリデーション結果が不正である場合、$falseを返す。
     #
     # @access public
     # @param string $subject メールタイトル
     # @return bool メールタイトルのバリデーション成否
     # @see なし
     # @throws なし
     #>
    [bool]validateSubject([string]$subject){
        $int_byte_num = [System.Text.Encoding]::GetEncoding("shift_jis").GetByteCount($subject)
        return (($int_byte_num -ge 0) -and ($int_byte_num -le 998))
    }
}