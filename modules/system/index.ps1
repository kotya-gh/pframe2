<#
 # [SystemUtil]�V�X�e���̑���Ɋւ���N���X
 #
 # �z�X�g���ύX��V���b�g�_�E���A���u�[�g�Ȃǃz�X�g�̑S�̓I�ȑ���AOS�֘A�̏��擾���s�����߂̃N���X�B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category �V�X�e������
 # @package �Ȃ�
 #>
class SystemUtil{
    SystemUtil(){

    }

    <#
     # [SystemUtil]�R���s���[�^�̃V���b�g�_�E��
     #
     # ���z�X�g���V���b�g�_�E������B
     #
     # @access public
     # @param �Ȃ�
     # @return bool �V���b�g�_�E���̐���
     # @see Stop-Computer
     # @throws �V���b�g�_�E���ŗ�O�������A$false��Ԃ��B
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
     # [SystemUtil]�R���s���[�^�̍ċN��
     #
     # ���z�X�g���ċN������B
     #
     # @access public
     # @param �Ȃ�
     # @return bool �ċN���̐���
     # @see Restart-Computer
     # @throws �ċN���ŗ�O�������A$false��Ԃ��B
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
     # [SystemUtil]�z�X�g���̕ύX
     #
     # ���z�X�g����$new_name�ɕύX����B
     # �z�X�g�����K��iRFC2396�j�ɓK�����Ă��Ȃ��ꍇ�A����уz�X�g���ύX�Ɏ��s�����ꍇ��false��Ԃ��B
     #
     # @access public
     # @param string $new_name �V�����z�X�g��
     # @return bool �z�X�g���ύX�̐���
     # @see Rename-Computer, Hosts.ValidateHostname
     # @throws �z�X�g���ύX�ŗ�O�������A$false��Ԃ��B
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
     # [SystemUtil]�h���C���ǉ�����
     #
     # ���z�X�g����$domain�Ŏw�肷��h���C���ɎQ��������B
     # �h���C���Q���̂��߂ɁA$domain����DomainAdmin������L����$user���w�肷��B
     # �h���C���Q����A�z�X�g���ċN������B
     # �h���C���Q���Ɏ��s�����ꍇ��false��Ԃ��B
     #
     # @access public
     # @param string $domain �Q������h���C����
     # @param string $user DomainAdmin������L����A�J�E���g��
     # @return bool �h���C���Q���̐���
     # @see Add-Computer
     # @throws �h���C���Q���ŗ�O�������A$false��Ԃ��B
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
     # [SystemUtil]GUID�̎擾
     #
     # GUID���擾����B
     #
     # @access public
     # @param �Ȃ�
     # @return guid �V�K���s����GUID
     # @see NewGuid
     # @throws �Ȃ�
     #>
    [Guid]GetGuid(){
        return [Guid]::NewGuid()
    }

    <#
     # [SystemUtil]GUID�̎擾
     #
     # GUID���擾����BGetGuid�̃G�C���A�X�B
     #
     # @access public
     # @param �Ȃ�
     # @return guid �V�K���s����GUID
     # @see GetGuid
     # @throws �Ȃ�
     #>
    [Guid]GetUuid(){
        return $this.GetGuid()
    }

    <#
     # [SystemUtil]ProcessID�̎擾
     #
     # ���g�̃v���Z�XID���擾����B
     #
     # @access public
     # @param �Ȃ�
     # @return int ���g�̃v���Z�XID
     # @see Get-WmiObject
     # @throws �Ȃ�
     #>
    [Int]GetPid(){
        [int]$pid=0
        $processId = Get-WmiObject win32_process -filter processid=$pid | ForEach-Object{$_.parentprocessid;}
        return $processId
    }
}