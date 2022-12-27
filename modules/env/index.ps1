<#
 # [Env]ホストの環境取得用クラス
 #
 # ホスト名やユーザ名等の環境取得用処理を定義する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category 環境情報取得操作
 # @package なし
 #>
class Env{
    Env(){
        
    }

    <#
     # [Env]ホスト名の取得
     #
     # スクリプトを実行しているホスト名を返す。
     # ホスト名取得に失敗した場合$falseを返す。
     #
     # @access public
     # @param なし
     # @return string $hostname ホスト名
     # @see [Net.Dns]::GetHostName
     # @throws ホスト名取得で例外発生時、$falseを返す。
     #>
    [string]GetHostname(){
        try{
            [string]$hostname=([Net.Dns]::GetHostName())
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $hostname
    }

    <#
     # [Env]ユーザ名の取得
     #
     # スクリプトを実行しているユーザ名を返す。
     # ユーザ名取得に失敗した場合$falseを返す。
     #
     # @access public
     # @param なし
     # @return string $username ユーザ名
     # @see [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
     # @throws ユーザ名取得で例外発生時、$falseを返す。
     #>
    [string]GetUsername(){
        try{
            [string]$username=([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $username
    }
}