<#
 # [ExArray]�z�񑀍�p�N���X
 #
 # �z��̑���A�������̏������`����B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category �z�񑀍쏈��
 # @package �Ȃ�
 #>
class ExArray{

    ExArray(){

    }

    <#
     # [ExArray]�z��̌���
     #
     # $needle�Ŏw�肷�镶����$haystack�Ŏw�肷��z��̗v�f�Ɋ܂܂�邩��bool�ŕԂ��B
     # �����Ɏ��s�����ꍇ�A�܂��͔z����ɑ��݂��Ȃ��ꍇ��$false��Ԃ��B
     #
     # @access public
     # @param string $needle �����p������
     # @param array $haystack �����Ώ۔z��
     # @return bool �z����̕����񌟍�����
     # @see [Array]::IndexOf
     # @throws �z��������ŗ�O�������A$false��Ԃ��B
     #>
    [bool]ArrayIndexOf([string]$needle, [array]$haystack){
        try{
            [int]$index=[Array]::IndexOf($haystack, $needle)
        } catch {
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        if($index -eq -1){
            return $false
        }
        return $true
    }

    <#
     # [ExArray]�z��̌���
     #
     # $needle�Ŏw�肷�镶����$haystack�Ŏw�肷��z��̗v�f�Ɋ܂܂�邩��bool�ŕԂ��B
     # ���������񂪔z����ɑ��݂��Ȃ��ꍇ��$false��Ԃ��B
     #
     # @access public
     # @param string $needle �����p������
     # @param array $haystack �����Ώ۔z��
     # @return bool �z����̕����񌟍�����
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]ArraySearch([string]$needle, [array]$haystack){
        foreach($value in $haystack){
            if($needle -eq $value){
                return $true
            }
        }
        return $false
    }

    <#
     # [ExArray]�n�b�V���̌���
     #
     # $hash.$key����$value�Ŏ����l�����݂��邩��bool�ŕԂ��B
     # ���������񂪔z����ɑ��݂��Ȃ��ꍇ��$false��Ԃ��B
     #
     # @access public
     # @param string $key �����p�L�[
     # @param string $value �����p������
     # @param object $hash �����Ώۃn�b�V��
     # @return bool �n�b�V�����̕����񌟍�����
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]HashValueSearch([string]$key, [string]$value, [object]$hash){
        foreach($record in $hash){
            if($record.$key -eq $value){
                return $true
            }
        }
        return $false
    }

    <#
     # [ExArray]PSCustomObject�̃v���p�e�B����
     #
     # $obj�Ŏ���PSCustomObject����$prop�Ŏ����v���p�e�B�����݂��邩��bool�ŕԂ��B
     # ����������PSCustomObject���ɑ��݂��Ȃ��ꍇ��$false��Ԃ��B
     #
     # @access public
     # @param object $obj PSCustomObject
     # @param string $prop �����p�v���p�e�B
     # @return bool PSCustomObject���̃v���p�e�B��������
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]PsoSearchProp([object]$obj, [string]$prop){
        $list=($obj | Get-Member -MemberType NoteProperty | Select -ExpandProperty Name)
        $Arr=New-object ExArray
        if($Arr.ArraySearch($prop, $list) -eq $true){
            return $true
        }
        return $false
    }

    <#
     # [ExArray]�z��̏d���r��
     #
     # $haystack�Ŏ����z����ɏd�������v�f������ꍇ�A�d����r�������z���Ԃ��B
     # �d���r���Ɏ��s�����ꍇ��$false��Ԃ��B
     #
     # @access public
     # @param array $haystack �d����r������z��
     # @return array �d����r�������z��
     # @see Get-Unique
     # @throws �d���r���ŗ�O�������A$false��Ԃ��B
     #>
    [array]ArrayUnique([array]$haystack){
        try{
            $haystack = $haystack | Sort-Object | Get-Unique
        } catch {
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $haystack
    }

    <#
     # [ExArray]�z��̏d������
     #
     # $haystack�Ŏ����z����ɏd�������v�f�����邩�ǂ�����bool�ŕԂ��B
     # �d�����Ă���ꍇ��$false��Ԃ��B
     #
     # @access public
     # @param array $haystack �d���L���𔻒肷��z��
     # @return bool �d�����茋��
     # @see ExArray.ArrayUnique
     # @throws �Ȃ�
     #>
    [bool]IsDuplicateArrayValue([array]$haystack){
        if( $this.ArrayUnique($haystack).Length -eq $haystack.Length ){
            return $true
        }
        return $false
    }

    <#
     # [ExArray]�z��̍Ō�̗v�f�̍폜
     #
     # $haystack�Ŏ����z�񂩂�Ō�̒l���폜�����z���Ԃ��B�z�񂪋�̏ꍇ�͋�̔z����쐬���ĕԂ��B
     # �d�����Ă���ꍇ��$false��Ԃ��B
     #
     # @access public
     # @param array $haystack �����v�f���폜����z��
     # @return array $haystack �����v�f���폜�����z��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [array]Pop([array]$haystack){
        if($haystack.Length -eq 0){
            return @()
        }
        $haystack=$haystack[0..($haystack.Length-2)]
        return $haystack
    }

    <#
     # [ExArray]�n�b�V���̃\�[�g�i�����j
     #
     # $hash�Ŏ����n�b�V����$key�Ŏ����L�[�ŏ����Ƀ\�[�g�����n�b�V����Ԃ��B
     #
     # @access public
     # @param object $hash �\�[�g����n�b�V��
     # @param string $key �\�[�g����L�[
     # @return object �\�[�g�����n�b�V��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [object]HashSort([object]$hash, [string]$key){
        if($hash.Length -ge 2){
            return ($hash.GetEnumerator() | Sort-Object -Property $key)
        }
        return $hash
    }

    <#
     # [ExArray]�n�b�V���̃\�[�g�i�~���j
     #
     # $hash�Ŏ����n�b�V����$key�Ŏ����L�[�ō~���Ƀ\�[�g�����n�b�V����Ԃ��B
     #
     # @access public
     # @param object $hash �\�[�g����n�b�V��
     # @param string $key �\�[�g����L�[
     # @return object �\�[�g�����n�b�V��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
     [object]HashSortDesc([object]$hash, [string]$key){
        if($hash.Length -ge 2){
            return ($hash.GetEnumerator() | Sort-Object -Property $key -Descending)
        }
        return $hash
    }

    <#
     # [ExArray]�z��̃\�[�g�i�����j
     #
     # $arr�Ŏ����z��������Ƀ\�[�g�����z���Ԃ��B
     #
     # @access public
     # @param array $arr �����Ƀ\�[�g����z��
     # @return array �\�[�g�����z��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [array]ArraySort([array]$arr){
        return ($arr | Sort-Object)
    }

    <#
     # [ExArray]�z��̃\�[�g�i�~���j
     #
     # $arr�Ŏ����z����~���Ƀ\�[�g�����z���Ԃ��B
     #
     # @access public
     # @param array $arr �~���Ƀ\�[�g����z��
     # @return array �\�[�g�����z��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [array]ArraySortDesc([array]$arr){
        return ($arr | Sort-Object -Descending)
    }

    <#
     # [ExArray]�n�b�V���e�[�u���ϊ�
     #
     # $cumtomObject�Ŏ���PSCustomObject���n�b�V���e�[�u���ɕϊ�����
     #
     # @access public
     # @param PSCustomObject $cumtomObject �ϊ�������PSCustomObject
     # @return HashTable �ϊ���̃n�b�V���e�[�u��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [object]ConvertToHashTable([PSCustomObject]$cumtomObject){
        $hashTable=@{}
        foreach($key in $cumtomObject.PSObject.Properties.Name){
            if($cumtomObject.$key -is [System.Management.Automation.PSCustomObject]){
                $hashTable[$key]=$this.ConvertToHashTable($cumtomObject.$key)
            }else{
                $hashTable[$key]=$cumtomObject.$key
            }
        }
        return $hashTable    
    }
}

