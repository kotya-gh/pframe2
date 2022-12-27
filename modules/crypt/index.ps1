<#
 # [Crypt]�Í��������p�N���X
 #
 # ������A�t�@�C���̈Í����������`����B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category �Í�������
 # @package �Ȃ�
 #>
class Crypt{
    Crypt(){

    }

    <#
     # [Crypt]������̈Í���
     #
     # $plaintext�Ŏw�肷�镶�����SecureString�ňÍ��������������Ԃ��B
     # $plaintext����A�܂͂��Í������s����$false��Ԃ��B
     #
     # @access public
     # @param string $plaintext ����
     # @return string $encrypt �Í���������
     # @see ConvertTo-SecureString
     # @throws ������Í����ŗ�O�������A$false��Ԃ��B
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
     # [Crypt]������̕���
     #
     # $encrypt�Ŏw�肷�镶�����SecureString�ŕ��������������Ԃ��B
     # $encrypt����A�܂��͕������s����$false��Ԃ��B
     #
     # @access public
     # @param string $encrypt �Í���������
     # @return string $plaintext ����������
     # @see ConvertTo-SecureString
     # @throws �����񕜍��ŗ�O�������A$false��Ԃ��B
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
     # [Crypt]������̃n�b�V���쐬
     #
     # $needle�Ŏw�肷�镶�����SHA256�Ńn�b�V���������������Ԃ��B
     # $needle����A�܂��̓n�b�V�������s����$false��Ԃ��B
     #
     # @access public
     # @param string $needle �n�b�V���Ώە�����
     # @return string $sha256hash �n�b�V�����㕶����
     # @see [security.cryptography.SHA256]::create().computehash
     # @throws �n�b�V�����ŗ�O�������A$false��Ԃ��B
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