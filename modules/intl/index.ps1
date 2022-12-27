<#
 # [Intl]����ݒ�p�N���X
 #
 # ������Ή��p�̏������`����B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category ������p����
 # @package �Ȃ�
 #>
class Intl{
    [string]$path
    [object]$define

    Intl([string]$path){
        $this.define=(Get-Content (Join-Path $PSScriptRoot "define.json") -Raw | ConvertFrom-Json)

        # init lang
        $script:LOCALE=$this.defaultLocale($script:LOCALE)
        $this.path=(Join-Path $path ($script:LOCALE+".json"))
    }

    <#
     # [Intl]���P�[��������̎擾
     #
     # $locale�Ŏw�肷�郍�P�[��������`����Ă���ꍇ�A���͒l��Ԃ��B
     # ��`����Ă��Ȃ����P�[���̏ꍇ"ja_JP"��Ԃ��B
     #
     # @access public
     # @param string $locale ���P�[��������
     # @return string ���P�[��������
     # @see
     # @throws �Ȃ�
     #>
    [string]defaultLocale([string]$locale){
        [object]$arr=New-Object ExArray
        if($arr.ArraySearch($locale, $this.define.lang) -eq $true){
            return $locale
        }
        return "ja_JP"
    }

    <#
     # [Intl]ID���當������擾
     #
     # $id�Ŏw�肷��ID�ɑΉ����镶�����Ԃ��B
     # ���݂��Ȃ�$id���w�肵���ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param string $id �����pID
     # @return string ID�ɑΉ����镶����
     # @see Intl.FormattedMessage
     # @throws �Ȃ�
     #>
    [string]FormattedMessage([string]$id){
        return $this.FormattedMessage($id, @{})
    }

    <#
     # [Intl]ID���當������擾
     #
     # $id�Ŏw�肷��ID�ɑΉ����镶�����Ԃ��B
     # ���������$value�̃n�b�V���L�[���u__(�n�b�V���L�[)�v�̌`���ŋL�ڂ��Ă���ꍇ�A�Ή�����n�b�V���l�Ɠ���ւ���B
     # ���݂��Ȃ�$id���w�肵���ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param string $id �����pID
     # @param object $value ���b�Z�[�W��������ϊ��p�n�b�V��
     # @return string $returnString ID�ɑΉ����镶����
     # @see Get-Content
     # @throws ���b�Z�[�W��`�t�@�C���擾�ŗ�O�������A$false��Ԃ��B
     #>
    [string]FormattedMessage([string]$id, [object]$value){
        [object]$file=New-Object File
        if($file.IsFile($this.path) -eq $false){
            $script:LAST_ERROR_MESSAGE="No lang file.($this.path)"            
            return $false
        }
        # import lang file.
        try{
            [object]$lang=(Get-Content $this.path -Raw | ConvertFrom-Json)
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        # search lang key
        if($null -eq $lang.$id){
            $script:LAST_ERROR_MESSAGE="No id is found.($id)"            
            return $false
        }
        [string]$returnString=$lang.$id
        foreach($key in $value.Keys){
            $returnString=$returnString.Replace("__($key)", $value[$key])
        }
        return $returnString
    }
}