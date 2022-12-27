<#
 # [Env]�z�X�g�̊��擾�p�N���X
 #
 # �z�X�g���⃆�[�U�����̊��擾�p�������`����B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category �����擾����
 # @package �Ȃ�
 #>
class Env{
    Env(){
        
    }

    <#
     # [Env]�z�X�g���̎擾
     #
     # �X�N���v�g�����s���Ă���z�X�g����Ԃ��B
     # �z�X�g���擾�Ɏ��s�����ꍇ$false��Ԃ��B
     #
     # @access public
     # @param �Ȃ�
     # @return string $hostname �z�X�g��
     # @see [Net.Dns]::GetHostName
     # @throws �z�X�g���擾�ŗ�O�������A$false��Ԃ��B
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
     # [Env]���[�U���̎擾
     #
     # �X�N���v�g�����s���Ă��郆�[�U����Ԃ��B
     # ���[�U���擾�Ɏ��s�����ꍇ$false��Ԃ��B
     #
     # @access public
     # @param �Ȃ�
     # @return string $username ���[�U��
     # @see [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
     # @throws ���[�U���擾�ŗ�O�������A$false��Ԃ��B
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