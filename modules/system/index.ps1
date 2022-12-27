<#
 # [SystemUtil]システムの操作に関するクラス
 #
 # ホスト名変更やシャットダウン、リブートなどホストの全体的な操作、OS関連の情報取得を行うためのクラス。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category システム操作
 # @package なし
 #>
class SystemUtil{
    SystemUtil(){

    }

    <#
     # [SystemUtil]コンピュータのシャットダウン
     #
     # 自ホストをシャットダウンする。
     #
     # @access public
     # @param なし
     # @return bool シャットダウンの成否
     # @see Stop-Computer
     # @throws シャットダウンで例外発生時、$falseを返す。
     #>
    [bool]ShutdownForce(){
        [object]$env=New-Object Env
        try{
            Stop-Computer -ComputerName $env.GetHostname() -Force
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        return $true
    }

    <#
     # [SystemUtil]コンピュータの再起動
     #
     # 自ホストを再起動する。
     #
     # @access public
     # @param なし
     # @return bool 再起動の成否
     # @see Restart-Computer
     # @throws 再起動で例外発生時、$falseを返す。
     #>
    [bool]RestartForce(){
        [object]$env=New-Object Env
        try{
            Restart-Computer -ComputerName $env.GetHostname() -Force
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        return $true
    }

    <#
     # [SystemUtil]ホスト名の変更
     #
     # 自ホスト名を$new_nameに変更する。
     # ホスト名が規約（RFC2396）に適合していない場合、およびホスト名変更に失敗した場合はfalseを返す。
     #
     # @access public
     # @param string $new_name 新しいホスト名
     # @return bool ホスト名変更の成否
     # @see Rename-Computer, Hosts.ValidateHostname
     # @throws ホスト名変更で例外発生時、$falseを返す。
     #>
    [bool]RenameHostname([string]$new_name){
        $hosts=New-Object Hosts
        if($hosts.ValidateHostname($new_name) -eq $false){
            $script:LAST_ERROR_MESSAGE="Invalid name."
            return $false
        }
        [object]$env=New-Object Env
        try{
            Rename-Computer -NewName $new_name -LocalCredential $env.GetHostname()\administrator -Restart
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        return $true
    }

    <#
     # [SystemUtil]ドメイン追加処理
     #
     # 自ホスト名を$domainで指定するドメインに参加させる。
     # ドメイン参加のために、$domain内でDomainAdmin権限を有する$userを指定する。
     # ドメイン参加後、ホストを再起動する。
     # ドメイン参加に失敗した場合はfalseを返す。
     #
     # @access public
     # @param string $domain 参加するドメイン名
     # @param string $user DomainAdmin権限を有するアカウント名
     # @return bool ドメイン参加の成否
     # @see Add-Computer
     # @throws ドメイン参加で例外発生時、$falseを返す。
     #>
    [bool]AddDomain([string]$domain, [string]$user){
        [object]$env=New-Object Env
        try{
            Add-Computer -ComputerName $env.GetHostname() -DomainName $domain -Credential $domain\$user -Restart
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        return $true
    }

    <#
     # [SystemUtil]GUIDの取得
     #
     # GUIDを取得する。
     #
     # @access public
     # @param なし
     # @return guid 新規発行したGUID
     # @see NewGuid
     # @throws なし
     #>
    [Guid]GetGuid(){
        return [Guid]::NewGuid()
    }

    <#
     # [SystemUtil]GUIDの取得
     #
     # GUIDを取得する。GetGuidのエイリアス。
     #
     # @access public
     # @param なし
     # @return guid 新規発行したGUID
     # @see GetGuid
     # @throws なし
     #>
    [Guid]GetUuid(){
        return $this.GetGuid()
    }

    <#
     # [SystemUtil]ProcessIDの取得
     #
     # 自身のプロセスIDを取得する。
     #
     # @access public
     # @param なし
     # @return int 自身のプロセスID
     # @see Get-WmiObject
     # @throws なし
     #>
    [Int]GetPid(){
        [int]$pid=0
        $processId = Get-WmiObject win32_process -filter processid=$pid | ForEach-Object{$_.parentprocessid;}
        return $processId
    }
}