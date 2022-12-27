<#
 # [Encode]�����̃G���R�[�h�p�N���X
 #
 # �����̃G���R�[�h�A�f�R�[�h��ϊ��̂��߂̏������`����B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category ������G���R�[�h
 # @package �Ȃ�
 #>
class Encode{
    Encode(){

    }

    <#
     # [Encode]�������URL�G���R�[�h
     #
     # $code�Ŏw�肷�镶���R�[�h��$string�������ǂݍ��݁AURL�G���R�[�h�̌��ʂ�Ԃ��B
     # URL�G���R�[�h�Ɏ��s�����ꍇ$false��Ԃ��B
     #
     # @access public
     # @param string $string �G���R�[�h�O�̕�����
     # @param string $code �����R�[�h��
     # @return string $result URL�G���R�[�h��̕�����
     # @see [System.Web.HttpUtility]::UrlEncode
     # @throws URL�G���R�[�h�ŗ�O�������A$false��Ԃ��B
     #>
    [string]Urlencode([string]$string, [string]$code){
        Add-Type -AssemblyName System.Web
        try{
            [string]$result=([System.Web.HttpUtility]::UrlEncode($string, [Text.Encoding]::GetEncoding($code)))
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        return $result
    }
 
    <#
     # [Encode]�������URL�G���R�[�h
     #
     # $code�Ŏw�肷�镶���R�[�h��$array�z�񂩂�1�v�f���������ǂݍ��݁AURL�G���R�[�h�̌��ʂ�z��ŕԂ��B
     # URL�G���R�[�h�Ɏ��s�����ꍇ$false��Ԃ��B
     #
     # @access public
     # @param array $array �G���R�[�h�O�̕�����z��
     # @param string $code �����R�[�h��
     # @return array $result URL�G���R�[�h��̕�����z��
     # @see Encode.Urlencode
     # @throws URL�G���R�[�h�ŗ�O�������A$false��Ԃ��B
     #>
    [array]Urlencode([array]$array, [string]$code){
        [array]$result=@()
        foreach($value in $array){
            $result+=$this.Urlencode($value, $code)
        }
        return $result
    }

    <#
     # [Encode]�������URL�f�R�[�h
     #
     # $code�Ŏw�肷�镶���R�[�h��$string�������ǂݍ��݁AURL�f�R�[�h�̌��ʂ�Ԃ��B
     # URL�f�R�[�h�Ɏ��s�����ꍇ$false��Ԃ��B
     #
     # @access public
     # @param string $string �f�R�[�h�O�̕�����
     # @param string $code �����R�[�h��
     # @return string $result URL�f�R�[�h��̕�����
     # @see [System.Web.HttpUtility]::UrlDecode
     # @throws URL�f�R�[�h�ŗ�O�������A$false��Ԃ��B
     #>
    [string]Urldecode([string]$string, [string]$code){
        Add-Type -AssemblyName System.Web
        return [System.Web.HttpUtility]::UrlDecode($string, [Text.Encoding]::GetEncoding($code))
    }

    <#
     # [Encode]�������URL�f�R�[�h
     #
     # $code�Ŏw�肷�镶���R�[�h��$array�z�񂩂�1�v�f���������ǂݍ��݁AURL�f�R�[�h�̌��ʂ�z��ŕԂ��B
     # URL�f�R�[�h�Ɏ��s�����ꍇ$false��Ԃ��B
     #
     # @access public
     # @param array $array �f�R�[�h�O�̕�����z��
     # @param string $code �����R�[�h��
     # @return array $result URL�f�R�[�h��̕�����z��
     # @see Encode.Urldecode
     # @throws URL�f�R�[�h�ŗ�O�������A$false��Ԃ��B
     #>
    [array]Urldecode([array]$array, [string]$code){
        [array]$result=@()
        foreach($value in $array){
            $result+=$this.Urldecode($value, $code)
        }
        return $result
    }

    <#
     # [Encode]���蕶����̕ϊ�
     #
     # $encodedString�Ŏw�肷�镶���������%c2%a1�`%c2%bf�A
     # �����%c3%80�`%c3%bf��%3f�ɕϊ����A�ϊ���̕������Ԃ��B
     #
     # @access public
     # @param string $encodedString �ϊ��O�̕�����
     # @return string $encodedString �ϊ���̕�����z��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [string]ConvertSpecialChar([string]$encodedString){
        # sp_char -> "%3f"
        # sp_char(1) ... %c2%a1 - %c2%bf
        # sp_char(2) ... %c3%80 - %c3%bf
        for([int]$i=0xa1; $i -le 0xbf; $i++){
            [string]$hex=("%c2%"+$i.ToString("x2"))
            $encodedString=$encodedString -creplace $hex, "%3f"
        }
        for([int]$i=0x80; $i -le 0xbf; $i++){
            [string]$hex=("%c3%"+$i.ToString("x2"))
            $encodedString=$encodedString -creplace $hex, "%3f"
        }
        return $encodedString
    }

    <#
     # [Encode]���蕶����̕ϊ�
     #
     # $encodedArray�Ŏw�肷�镶����z�񂩂�1�v�f���������ǂݍ��݁A
     # %c2%a1�`%c2%bf�A�����%c3%80�`%c3%bf��%3f�ɕϊ����A�ϊ���̕������Ԃ��B
     #
     # @access public
     # @param array $encodedArray �ϊ��O�̕�����z��
     # @return array $result �ϊ���̕�����z��z��
     # @see Encode.ConvertSpecialChar
     # @throws �Ȃ�
     #>
    [array]ConvertSpecialChar([array]$encodedArray){
        [array]$result=@()
        foreach($value in $encodedArray){
            $result+=$this.ConvertSpecialChar($value)
        }
        return $result
    }

    <#
     # [Encode]�t�@�C����Json�f�R�[�h
     #
     # $file_path�Ŏw�肷��t�@�C����ǂݍ��݁AJson�ϊ���̌��ʂ�z��ŕԂ��B
     # Json�f�R�[�h�Ɏ��s�����ꍇ$false��Ԃ��B
     #
     # @access public
     # @param string $file_path �f�R�[�h����t�@�C���p�X
     # @return array $result Json�f�R�[�h��̔z��
     # @see convertFrom-Json
     # @throws Json�f�R�[�h�ŗ�O�������A$false��Ԃ��B
     #>
    [array]JsonFileDecode([string]$file_path){
        try{
            [array]$decode_array=(Get-Content $file_path -Raw | convertFrom-Json)
        } catch {
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }        
        return $decode_array
    }
}