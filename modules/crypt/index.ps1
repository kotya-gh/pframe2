<#
 # [Crypt]暗号化処理用クラス
 #
 # 文字列、ファイルの暗号化処理を定義する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category 暗号化処理
 # @package なし
 #>
class Crypt{
    Crypt(){

    }

    <#
     # [Crypt]文字列の暗号化
     #
     # $plaintextで指定する文字列をSecureStringで暗号化した文字列を返す。
     # $plaintextが空、まはた暗号化失敗時は$falseを返す。
     #
     # @access public
     # @param string $plaintext 平文
     # @return string $encrypt 暗号化文字列
     # @see ConvertTo-SecureString
     # @throws 文字列暗号化で例外発生時、$falseを返す。
     #>
    [string]EncryptSecureString([string]$plaintext){
        if((($plaintext -as "string") -eq $false) -or ($plaintext -eq "")){
            return $false
        }
        try{
            [System.Security.SecureString]$secure = ConvertTo-SecureString $plaintext -AsPlainText -Force
            [string]$encrypt = ConvertFrom-SecureString -SecureString $secure         
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $encrypt
    }

    <#
     # [Crypt]文字列の復号
     #
     # $encryptで指定する文字列をSecureStringで復号した文字列を返す。
     # $encryptが空、または復号失敗時は$falseを返す。
     #
     # @access public
     # @param string $encrypt 暗号化文字列
     # @return string $plaintext 復号文字列
     # @see ConvertTo-SecureString
     # @throws 文字列復号で例外発生時、$falseを返す。
     #>
    [string]DecryptSecureString([string]$encrypt){
        if((($encrypt -as "string") -eq $false) -or ($encrypt -eq "")){
            return $false
        }
        try{
            [System.Security.SecureString]$decrypt = ConvertTo-SecureString -String $encrypt
            [System.IntPtr]$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($decrypt)
            [string]$plaintext = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)    
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $plaintext
    }

    <#
     # [Crypt]文字列のハッシュ作成
     #
     # $needleで指定する文字列をSHA256でハッシュ化した文字列を返す。
     # $needleが空、またはハッシュ化失敗時は$falseを返す。
     #
     # @access public
     # @param string $needle ハッシュ対象文字列
     # @return string $sha256hash ハッシュ化後文字列
     # @see [security.cryptography.SHA256]::create().computehash
     # @throws ハッシュ化で例外発生時、$falseを返す。
     #>
    [string]textToHashSHA256([string]$needle){
        if((($needle -as "string") -eq $false) -or ($needle -eq "")){
            return $false
        }
        try{
            $sha256hash=[security.cryptography.SHA256]::create().computehash([text.encoding]::ascii.getbytes($needle))|%{$_.tostring("x2")}
            $sha256hash=[string]::concat($sha256hash)
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $sha256hash
    }
}