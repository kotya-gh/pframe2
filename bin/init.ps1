<#
 # [INIT]�����������p�N���X
 #
 # �R���|�[�l���g���s��Require����B
 # ���ʃf�B���N�g���̒�`�A�ݒ�ǂݍ��݁A���W���[���ǂݍ��݂����{����B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category ���������ʏ���
 # @package �Ȃ�
 #>
class Init{
    [string]$path_root
    [string]$path_conf
    [string]$path_modules
    [string]$path_log
    [string]$path_files
    [string]$path_script_root
    [array]$conf_require
    [array]$conflist_user
    [array]$conf_user
    [string]$component_name
    [string]$path_component_conf
    [array]$conf_component

    Init($path_script_root){
        # ���[�g�f�B���N�g���̒�`
        $this.path_root=(Split-Path $PSScriptRoot -Parent)

        # �X�N���v�g�\���f�B���N�g���̕ϐ��i�[����
        $this.path_conf=$this.GetExistDirPath((Join-Path $this.path_root "conf"))
        $this.path_modules=$this.GetExistDirPath((Join-Path $this.path_root "modules"))
        $this.path_log=$this.GetExistDirPath((Join-Path $this.path_root "log"))
        $this.path_files=$this.GetExistDirPath((Join-Path $this.path_root "files"))

        # �X�N���v�g���[�g�f�B���N�g���̐ݒ�
        $this.path_script_root=$path_script_root

        # �R���|�[�l���g�p�O���t�@�C�����̊i�[
        $this.conflist_user=(Get-ChildItem $this.path_script_root -include *.json -Name)

        # �X�N���v�g�ݒ�̊i�[
        $this.conf_require=$this.SetRequiredConfigure()

        # �R���|�[�l���g���̎擾
        $this.component_name=(Split-Path -Leaf $this.GetScriptRootPath())

        # �R���|�[�l���g�̐ݒ�t�@�C�����̊i�[
        $this.path_component_conf=(Join-Path $this.path_conf ($this.component_name+".json"))

        # �R���|�[�l���g�̐ݒ���e�̊i�[
        $this.conf_component=$this.SetComponentConfigure()

        # �R���|�[�l���g�p�O���t�@�C���̃��v���P�[�V��������
        $this.ReplicationUserConfiture()

        # �R���|�[�l���g�p�O���t�@�C���̓��e�i�[
        $this.conf_user=$this.SetUserConfiture()
    }

    [string]GetRootPath(){
        return $this.path_root
    }

    [string]GetConfPath(){
        return $this.path_conf
    }

    [string]GetComponentName(){
        return $this.component_name
    }

    [string]GetModulesPath(){
        return $this.path_modules
    }

    [string]GetLogPath(){
        return $this.path_log
    }

    [string]GetFilesPath(){
        return $this.path_files
    }

    [string]GetScriptRootPath(){
        return $this.path_script_root
    }

    [array]GetRequiredConfigure(){
        return $this.conf_require
    }

    [array]GetUserConfigure(){
        return $this.conf_user
    }

    [array]GetComponentConfigure(){
        return $this.conf_component
    }

    <#
     # [INIT]�f�B���N�g���p�X�̑��݊m�F
     #
     # $path�Ŏw�肷��p�X�̃f�B���N�g���̑��݂��m�F���A���݂���ꍇ�͓��͒l�A���݂��Ȃ��ꍇ�͍쐬��p�X��Ԃ��B
     #
     # @access public
     # @param string $path �f�B���N�g���p�X
     # @return string $path �f�B���N�g���p�X
     # @see New-Item
     # @throws �Ȃ�
     #>
    [string]GetExistDirPath([string]$path){
        if((Test-Path $path) -eq $false){
            New-Item $path -type directory -ErrorAction Stop
        }
        return $path
    }

    <#
     # [INIT]
     #
     # $path�Ŏw�肷��p�X�̃t�@�C���̑��݂��m�F���A���݂���ꍇ�͓��͒l�A
     # ���݂��Ȃ��ꍇ�͋�t�@�C�����쐬��p�X��Ԃ��B
     #
     # @access public
     # @param string $path �t�@�C���p�X
     # @return string $path �t�@�C���p�X
     # @see New-Item
     # @throws �Ȃ�
     #>
    [string]GetExistFilePath([string]$path){
        if((Test-Path $path) -eq $false){
            New-Item $path -type file -ErrorAction Stop
        }
        return $path
    }

    <#
     # [INIT]�X�N���v�g�S�̐ݒ�t�@�C���̃n�b�V���ϊ�
     #
     # �X�N���v�g�S�̐ݒ�t�@�C��"require.json"��ǂݍ��݁AJson�f�R�[�h��̃n�b�V����Ԃ��B
     #
     # @access public
     # @param �Ȃ�
     # @return array Json�f�R�[�h��n�b�V��
     # @see Init.GetExistFilePath, Init.JsonFileDecode
     # @throws Init.JsonFileDecode�ŗ�O�������A$false��Ԃ��B
     #>
    [array]SetRequiredConfigure(){
        [string]$path_conf_require=$this.GetExistFilePath((Join-Path $this.path_conf "require.json"))
        return $this.JsonFileDecode($path_conf_require)
    }

    <#
     # [Crypt]������̕���
     #
     # $encrypt�Ŏw�肷�镶�����SecureString�ŕ��������������Ԃ��B
     # $encrypt����A�܂͂��������s����$false��Ԃ��B
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
     # [INIT]���[�U�ݒ�t�@�C���̃��v���P�[�V����
     #
     # �R���|�[�l���g�f�B���N�g���z����Json�t�@�C���i���[�U�ݒ�t�@�C���j���A
     # �R���|�[�l���g�ݒ�t�@�C���Ŏw�肷��ꏊ���烌�v���P�[�V��������B
     # �R���|�[�l���g�ݒ�t�@�C����exec�t���O��1�̏ꍇ�A����
     # ���v���P�[�V������/��̃t�@�C���̃n�b�V���l���قȂ�ꍇ���v���P�[�V�������������s����B
     #
     # @access public
     # @param �Ȃ�
     # @return bool ���v���P�[�V�����̐���
     # @see Copy-Item, Get-FileHash
     # @throws �t�@�C���R�s�[�ŗ�O�������A$false��Ԃ��B
     #>
    [bool]ReplicationUserConfiture(){
        # exec 1�ȊO��true��Ԃ��݂̂Ƃ���B
        if($this.conf_component.replication.exec -ne 1){
            return $true
        }

        # ���v���P�[�V��������l�b�g���[�N�h���C�u�ɐݒ肷��ꍇ�̏���
        # �l�b�g���[�N�}�E���g�ݒ�Aexec 1�̂Ƃ��̂ݎ��{
        if($this.conf_component.replication.netmount.exec -eq 1){
            $computerName = $this.conf_component.replication.netmount.servername
            $adminPass = $this.conf_component.replication.netmount.password
            $adminUser = $this.conf_component.replication.netmount.user
            
            # �F�؏��̃C���X�^���X�𐶐�����
            $securePass = ConvertTo-SecureString $adminPass -AsPlainText -Force;
            $cred = New-Object System.Management.Automation.PSCredential "$computerName\$adminUser", $securePass;
            
            New-PSDrive -Name $this.conf_component.replication.netmount.mountto -PSProvider FileSystem -Root "\\$computerName\$this.conf_component.replication.netmount.mountfrom" -Credential $cred;
        }

        foreach($filePath in $this.conf_component.replication.files){
            # �t�@�C�������݂��Ȃ��ꍇ�̓X�L�b�v�A�܂���JSON�ȊO�̏ꍇ�̓X�L�b�v
            if(
                ((Test-Path $filePath) -eq $false) -or
                ((Get-ChildItem $filePath).Extension.toLower() -ne ".json")
            ){
                continue
            }
            [string]$filePathBasename=([System.IO.Path]::GetFileName($filePath))
            [string]$dstPath=(Join-Path $this.path_script_root $filePathBasename)
            # conflist_user�Ɋ܂܂�Ȃ��ꍇ�A�V�K�t�@�C���Ɣ��肵�A�`�F�b�N�T����r�Ȃ��ŃR�s�[����B
            if([Array]::IndexOf($this.conflist_user, $filePathBasename) -eq -1){
                try{
                    Copy-Item $filePath $dstPath -ErrorAction Stop
                } catch {
                    Write-Host($_.Exception)
                    return $false 
                }
            }else{
                # �`�F�b�N�T����r��A�t�@�C�����قȂ�ꍇ�̓R�s�[����B
                if(
                    (Get-FileHash $dstPath).Hash.toLower() -ne
                    (Get-FileHash $filePath).Hash.toLower()
                ){
                    try{
                        Copy-Item $filePath $dstPath -ErrorAction Stop
                    } catch {
                        Write-Host($_.Exception)
                        return $false 
                    }
                }
            }
        }
        return $true
    }

    <#
     # [INIT]���[�U�ݒ�t�@�C����Json�f�R�[�h
     #
     # �R���|�[�l���g�f�B���N�g���z����Json�t�@�C���i���[�U�ݒ�t�@�C���j��Json�f�R�[�h���A
     # �t�@�C�����i�g���q�Ȃ��j���L�[�Ƃ���n�b�V���Ɋi�[�������ʂ�Ԃ��B
     #
     # @access public
     # @param �Ȃ�
     # @return array $ret_hash ���[�U�ݒ�t�@�C����Json�f�R�[�h����
     # @see Init.JsonFileDecode
     # @throws Init.JsonFileDecode�ŗ�O�������A$false��Ԃ��B
     #>
    [array]SetUserConfiture(){
        [array]$ret_hash=@{}
        foreach($conf in $this.conflist_user){
            [string]$path_conf_user=(Join-Path $this.path_script_root $conf -Resolve)
            $conf=[System.IO.Path]::GetFileNameWithoutExtension($path_conf_user) 
            $ret_hash+=@{ $conf = $this.JsonFileDecode($path_conf_user) }
        }
        return $ret_hash
    }

    <#
     # [INIT]�R���|�[�l���g�ݒ�t�@�C����Json�f�R�[�h
     #
     # ���ʐݒ�f�B���N�g��(conf)�z����Json�t�@�C���i�R���|�[�l���g�ݒ�t�@�C���j��
     # Json�f�R�[�h�������ʂ�Ԃ��B
     # �R���|�[�l���g�Ɠ����̐ݒ�t�@�C�������݂��Ȃ��ꍇ�Adefault.conf�̃f�R�[�h���ʂ�Ԃ��B
     #
     # @access public
     # @param string Init.path_component_conf
     # @return array $ret_hash ���[�U�ݒ�t�@�C����Json�f�R�[�h����
     # @see Init.GetExistFilePath, Init.JsonFileDecode
     # @throws Init.JsonFileDecode�ŗ�O�������A$false��Ԃ��B
     #>
    [array]SetComponentConfigure(){
        if((Test-Path $this.path_component_conf) -eq $true){
            $conf_path=$this.path_component_conf
        }else{
            $conf_path=$this.GetExistFilePath((Join-Path $this.path_conf "default.json"))
        }
        return $this.JsonFileDecode($conf_path)
    }

    <#
     # [INIT]Json�t�@�C�����I�u�W�F�N�g�ɕϊ�����
     #
     # $file_path�Ŏw��̃t�@�C����Json�f�R�[�h���ʂ�Ԃ��B
     #
     # @access public
     # @param string $file_path
     # @return array $decode_array �t�@�C����Json�f�R�[�h����
     # @see Get-Content, convertFrom-Json
     # @throws Get-Content�A�����convertFrom-Json�ŗ�O�������A$false��Ԃ��B
     #>
    [array]JsonFileDecode([string]$file_path){
        try{
            [array]$decode_array=(Get-Content $file_path | convertFrom-Json)
        } catch {
            Write-Host($_.Exception)
            return $false 
        }
        return $decode_array
    }
}

if([String]::IsNullOrEmpty($PROJECT_NAME)){
    set-variable -name PROJECT_NAME -value "Power Frame 1.0" -option constant -scope global
}

# Version check >= 5
# PowerShell�o�[�W�����̊m�F
[int]$version=$PSVersionTable.PSVersion.Major
if($version -le 4){ 
    Write-Host "PowerShell(>=5.0) is required."
    exit 1
}

# init�N���X�̏�����
$INIT=[Init]::new($MyInvocation.PSScriptRoot)

# include component configure
# �ݒ�t�@�C���̓ǂݍ��݁A�R���|�[�l���g���A���O�i�[�f�B���N�g���̐ݒ�
[array]$script:COMP_CONF=$INIT.GetComponentConfigure()
$script:COMP_CONF+=@{component_name=$INIT.GetComponentName(); path_log=$INIT.GetLogPath()}

# import configure process
# ���W���[���N���X�̓ǂݍ���
foreach($module in $INIT.GetRequiredConfigure().import_module){
    try{
        . (Join-Path $INIT.GetModulesPath() ($module+"\index.ps1") -Resolve)
    } catch { 
        Write-Host($_.Exception)
        exit 1
    }
}

# set user configuration
# ���[�U�ݒ�t�@�C���̓ǂݍ���
[array]$USR_CONF=$INIT.GetUserConfigure()

# set intl
# �o�̓��b�Z�[�W��`�p�����̏�����
$script:LOCALE=$INIT.conf_require.locale
[object]$INTL=New-Object Intl((Join-Path ([System.IO.Path]::GetDirectoryName($script:myInvocation.ScriptName)) "locale"))

# set user log
[object]$USR_LOG=[log]::new()
[object]$USR_MSG=$USR_LOG.log_format