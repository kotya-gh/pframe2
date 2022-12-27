<#
 # [Str]�����񑀍�Ɋւ���N���X
 #
 # ������̌����A�ҏW�A�u�����s�����߂̃N���X�B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category �����񑀍�
 # @package �Ȃ�
 #>
class Str{
    [object]$LOCALE
    Str(){
        $this.LOCALE=New-Object Intl((Join-Path $PSScriptRoot "locale"))
    }

    <#
     # [Str]���s�̕ϊ�����
     #
     # $string�Ŏw�肷�镶����̉��s���A�X�y�[�X�ɕϊ��������ʂ�Ԃ��B
     #
     # @access public
     # @param string $string �ϊ��O������
     # @return string $string �ϊ��㕶����
     # @see Replace
     # @throws �Ȃ�
     #>
    [string]TrimEol([string]$string){
        [string]$string=$string.Replace("`r`n", " ").Replace("`r", " ").Replace("`n", " ")
        return $string
    }

    <#
     # [Str]������̍폜����
     #
     # $string�Ŏw�肷�镶����̕�����̍��[����$needle���������A��v�����ꍇ�폜����B
     # ������$string��$needle���Z���ꍇfalse��Ԃ��B
     #
     # @access public
     # @param string $string �ϊ��O������
     # @param string $needle ����������
     # @return string $string �ϊ��㕶����
     # @see Remove
     # @throws �Ȃ�
     #>
    [string]LTrim([string]$string, [string]$needle){
        if($string.Length -lt $needle.Length){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("The_length_of_needle_is_too_long")
            return $false
        }
        if($string.Remove($needle.Length, $string.Length-$needle.Length) -eq $needle){
            return $string.Remove(0, $needle.Length)
        }
        return $string
    }

    <#
     # [Str]������̍폜����
     #
     # $string�Ŏw�肷�镶����̕�����̉E�[����$needle���������A��v�����ꍇ�폜����B
     # ������$string��$needle���Z���ꍇfalse��Ԃ��B
     #
     # @access public
     # @param string $string �ϊ��O������
     # @param string $needle ����������
     # @return string $string �ϊ��㕶����
     # @see Remove
     # @throws �Ȃ�
     #>
    [string]RTrim([string]$string, [string]$needle){
        if($string.Length -lt $needle.Length){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("The_length_of_needle_is_too_long")
            return $false
        }
        if($string.Remove(0, $string.Length-$needle.Length) -eq $needle){
            return $string.Remove($string.Length-$needle.Length, $needle.Length)
        }
        return $string
    }

    <#
     # [Str]������̌����A�u������
     #
     # $subject�Ŏw�肷�镶����𕶎���$search�Ō������A��v�����ꍇ������$replace�ɒu��������B
     #
     # @access public
     # @param string $search ����������
     # @param string $replace �u��������
     # @param string $subject �ϊ��O������
     # @return string �ϊ��㕶����
     # @see creplace
     # @throws �Ȃ�
     #>
    [string]Replace([string]$search, [string]$replace ,[string]$subject){
        return ($subject -creplace $search, $replace)
    }

    <#
     # [Str]������̌�������
     #
     # $haystack�Ŏw�肷�镶����𕶎���$needle�Ō������A��v�����ꍇ$true�A��v���Ȃ��ꍇ$false��Ԃ��B
     #
     # @access public
     # @param string $haystack �����Ώە�����
     # @param string $needle ����������
     # @return bool ��������
     # @see Contains
     # @throws �Ȃ�
     #>
    [bool]Strpos([string]$haystack, [string]$needle){
        return $haystack.Contains($needle)
    }

    <#
     # [Str]������̌�������
     #
     # $haystack�Ŏw�肷�镶�����z��$needle���̕�����Ō������A�P�ł���v�����ꍇ$true�A�S�Ĉ�v���Ȃ��ꍇ$false��Ԃ��B
     #
     # @access public
     # @param string $haystack �����Ώە�����
     # @param array $needle ����������
     # @return bool ��������
     # @see Contains
     # @throws �Ȃ�
     #>
    [bool]Strpos([string]$haystack, [array]$needle){
        foreach($value in $needle){
            if($haystack.Contains($value) -eq $true){
                return $true
            }
        }
        return $false
    }

    <#
     # [Str]������̕�������
     #
     # $string�Ŏw�肷�镶����𕶎���$delimiter�ŕ������A�e�v�f�𕶎���Ƃ��Ċi�[�����z���Ԃ��B
     #
     # @access public
     # @param string $delimiter ����������
     # @param string $string �����Ώە�����
     # @return array ���������v�f���i�[�����z��
     # @see split
     # @throws �Ȃ�
     #>
    [array]Explode([string]$delimiter, [string]$string){
        return ($string -split $delimiter)
    }

    <#
     # [Str]�z��̘A������
     #
     # $pieces�Ŏw�肷��z��v�f�𕶎���$glue�ɂ��A������B
     # ���ׂĂ̔z��v�f�̏�����ς����ɁA�e�v�f�Ԃ�$glue��������͂����1�̕�����ɂ��ĕԂ��B
     #
     # @access public
     # @param string $glue �v�f�Ԃ�A�����镶����
     # @param array $pieces �A��������������̔z��
     # @return string �z��v�f��A������������
     # @see $this.Ltrim
     # @throws �Ȃ�
     #>
    [string]Implode([string]$glue, [array]$pieces){
        if($pieces.Length -eq 0){
            return ""
        }
        [string]$ret=""
        foreach($val in $pieces){
            $ret+=($glue+[string]$val)
        }
        return $this.LTrim($ret, $glue)
    }

    <#
     # [Str]������̕�������
     #
     # $string�Ŏw�肷�镶������z���C�g�X�y�[�X�ŕ������A�e�v�f�𕶎���Ƃ��Ċi�[�����z���Ԃ��B
     # �O��̃z���C�g�X�y�[�X�͍폜����B
     # �z���C�g�X�y�[�X��[���p�X�y�[�X�A\t�A\n�A\r�A\f]�ƒ�`����B
     #
     # @access public
     # @param string $string �����Ώە�����
     # @return array ���������v�f���i�[�����z��
     # @see Trim()
     # @throws �Ȃ�
     #>
    [array]ExplodeWs([string]$string){
        $delimiter="__explodeWs_Delimiter__"
        return $string.Trim() -replace "(\s)+", $delimiter -split $delimiter
    }  
}