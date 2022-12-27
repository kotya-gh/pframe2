<#
 # [File]�t�@�C���E�f�B���N�g������p�N���X
 #
 # �t�@�C���A�f�B���N�g���̍쐬�A�폜�A���쓙�̏������`����B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category �t�@�C���E�f�B���N�g������
 # @package �Ȃ�
 #>
class File{
    [object]$LOCALE
    [string]$encording="default"

    File(){
        $this.LOCALE=New-Object Intl((Join-Path $PSScriptRoot "locale"))
    }

    <#
     # [File]�t�@�C���E�f�B���N�g���̑��݊m�F
     #
     # $path�Ŏw�肷��p�X�̃t�@�C���A�܂��̓f�B���N�g���̑��݂��m�F���A
     # ���݂���ꍇ��$true�A���݂��Ȃ��ꍇ��$false��Ԃ��B
     #
     # @access public
     # @param string $path �t�@�C���E�f�B���N�g���p�X
     # @return bool �t�@�C���A�܂��̓f�B���N�g���̑��ݗL��
     # @see Test-Path
     # @throws �Ȃ�
     #>
    [bool]IsFile([string]$path){
        if((Test-Path $path) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("File_does_not_exist", @{path=$path})
            return $false
        }
        return $true
    }

    <#
     # [File]�t�@�C���E�f�B���N�g���̑��݊m�F
     #
     # $pathObj�Ŏw�肷��p�X�̃t�@�C���A�܂��̓f�B���N�g���̑��݂��m�F���A
     # ���݂���ꍇ��$true�A�P�ł����݂��Ȃ��ꍇ��$false��Ԃ��B
     # $pathObj�ł́A�t�@�C���p�X�A�܂��̓f�B���N�g���p�X��z��Ŏw�肷��B
     #
     # @access public
     # @param array $pathObj �t�@�C���E�f�B���N�g���p�X
     # @return bool �t�@�C���A�܂��̓f�B���N�g���̑��ݗL��
     # @see Test-Path
     # @throws �Ȃ�
     #>
    [bool]IsFile([array]$pathObj){
        foreach($path in $pathObj){
            if($this.IsFile($path) -eq $false){
                return $false
            }
        }
        return $true
    }

    <#
     # [File]��t�@�C���̍쐬
     #
     # $path�Ŏw�肷��p�X�̋�t�@�C�����쐬����B
     # �d������t�@�C�������݂���ꍇ�A�܂��̓t�@�C���쐬���s�̏ꍇ$false��Ԃ��B
     #
     # @access public
     # @param string $path �t�@�C���p�X
     # @return bool �t�@�C���̍쐬����
     # @see New-Item
     # @throws �t�@�C���쐬�ŗ�O�������A$false��Ԃ��B
     #>
    [bool]MkFile([string]$path){
        if($this.IsFile($path) -eq $true){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Duplicate_files", @{path=$path})
            return $false
        }
        try{
            New-Item $path -type file -ErrorAction Stop
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $true
    }

    <#
     # [File]��t�@�C���̍쐬
     #
     # $pathObj�Ŏw�肷��p�X�̋�t�@�C�����쐬����B
     # �d������t�@�C�������݂���ꍇ�A�܂��͂P�ł��t�@�C���쐬���s�̏ꍇ$false��Ԃ��B
     # $pathObj�ł́A�t�@�C���p�X��z��Ŏw�肷��B
     #
     # @access public
     # @param array $pathObj �t�@�C���p�X
     # @return bool �t�@�C���̍쐬����
     # @see New-Item
     # @throws �Ȃ�
     #>
    [bool]MkFile([array]$pathObj){
        foreach($path in $pathObj){
            if($this.MkFile($path) -eq $false){
                return $false
            }
        }
        return $true
    }

    <#
     # [File]�f�B���N�g���̍쐬
     #
     # $path�Ŏw�肷��p�X�̃f�B���N�g�����쐬����B
     # �d������f�B���N�g�������݂���ꍇ�A�܂��̓f�B���N�g���쐬���s�̏ꍇ$false��Ԃ��B
     #
     # @access public
     # @param string $path �f�B���N�g���p�X
     # @return bool �f�B���N�g���̍쐬����
     # @see New-Item
     # @throws �f�B���N�g���̍쐬�ŗ�O�������A$false��Ԃ��B
     #>
    [bool]MkDir([string]$path){
        if($this.IsFile($path) -eq $true){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Duplicate_files", @{path=$path})
            return $false
        }
        try{
            New-Item $path -type directory -ErrorAction Stop
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $true
    }

    <#
     # [File]�f�B���N�g���̍쐬
     #
     # $pathObj�Ŏw�肷��p�X�̃f�B���N�g�����쐬����B
     # �d������f�B���N�g�������݂���ꍇ�A�܂��͂P�ł��f�B���N�g���쐬���s�̏ꍇ$false��Ԃ��B
     # $pathObj�ł́A�f�B���N�g���p�X��z��Ŏw�肷��B
     #
     # @access public
     # @param array $pathObj �f�B���N�g���p�X
     # @return bool �f�B���N�g���̍쐬����
     # @see New-Item
     # @throws �Ȃ�
     #>
    [bool]MkDir([array]$pathObj){
        foreach($path in $pathObj){
            if($this.MkDir($path) -eq $false){
                return $false
            }
        }
        return $true
    }

    <#
     # [File]��t�@�C���̍쐬�i�����j
     #
     # $path�Ŏw�肷��p�X�̋�t�@�C�����쐬����B
     # �d������t�@�C�������݂���ꍇ�ł������I�Ƀt�@�C�����쐬����B
     # �t�@�C���쐬���s�̏ꍇ$false��Ԃ��B
     #
     # @access public
     # @param string $path �t�@�C���p�X
     # @return bool �t�@�C���̍쐬����
     # @see New-Item
     # @throws �t�@�C���쐬�ŗ�O�������A$false��Ԃ��B
     #>
    [bool]MkFileForce([string]$path){
        try{
            New-Item $path -type file -Force -ErrorAction Stop
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $true
    }

    <#
     # [File]��t�@�C���̍쐬
     #
     # $pathObj�Ŏw�肷��p�X�̋�t�@�C�����쐬����B
     # �d������t�@�C�������݂���ꍇ�ł������I�Ƀt�@�C�����쐬����B
     # �t�@�C���쐬���s�̏ꍇ$false��Ԃ��B
     # $pathObj�ł́A�t�@�C���p�X��z��Ŏw�肷��B
     #
     # @access public
     # @param array $pathObj �t�@�C���p�X
     # @return bool �t�@�C���̍쐬����
     # @see New-Item
     # @throws �Ȃ�
     #>
    [bool]MkFileForce([array]$pathObj){
        foreach($path in $pathObj){
            if($this.MkFileForce($path) -eq $false){
                return $false
            }
        }
        return $true
    }

    <#
     # [File]�f�B���N�g���̍쐬
     #
     # $path�Ŏw�肷��p�X�̃f�B���N�g�����쐬����B
     # �d������f�B���N�g�������݂���ꍇ�ł������I�Ƀf�B���N�g�����쐬����B
     # �f�B���N�g���쐬���s�̏ꍇ$false��Ԃ��B
     #
     # @access public
     # @param string $path �f�B���N�g���p�X
     # @return bool �f�B���N�g���̍쐬����
     # @see New-Item
     # @throws �Ȃ�
     #>
    [bool]MkDirForce([string]$path){
        if($this.Rm($path) -eq $false){
            return $false
        }
        return $this.MkDir($path)
    }

    <#
     # [File]�f�B���N�g���̍쐬
     #
     # $pathObj�Ŏw�肷��p�X�̃f�B���N�g�����쐬����B
     # �d������f�B���N�g�������݂���ꍇ�ł������I�Ƀf�B���N�g�����쐬����B
     # �f�B���N�g���쐬���s�̏ꍇ$false��Ԃ��B
     # $pathObj�ł́A�f�B���N�g���p�X��z��Ŏw�肷��B
     #
     # @access public
     # @param array $pathObj �f�B���N�g���p�X
     # @return bool �f�B���N�g���̍쐬����
     # @see New-Item
     # @throws �Ȃ�
     #>
    [bool]MkDirForce([array]$pathObj){
        foreach($path in $pathObj){
            if($this.MkDirForce($path) -eq $false){
                return $false
            }
        }
        return $true
    }

    <#
     # [File]�t�@�C���E�f�B���N�g���̍폜
     #
     # $path�Ŏw�肷��p�X�̃t�@�C���A�܂��̓f�B���N�g�����폜����B
     # �w��̃t�@�C���E�f�B���N�g�������݂��Ȃ��A�܂��̓t�@�C���E�f�B���N�g���폜���s�̏ꍇ$false��Ԃ��B
     #
     # @access public
     # @param string $path �t�@�C���E�f�B���N�g���p�X
     # @return bool �t�@�C���E�f�B���N�g���̍폜����
     # @see Remove-Item
     # @throws �t�@�C���E�f�B���N�g���폜�ŗ�O�������A$false��Ԃ��B
     #>
    [bool]Rm([string]$path){
        if($this.IsFile($path) -eq $false){
            $script:LAST_ERROR_MESSAGE=""
            return $true
        }
        try{
            Remove-Item $path -Force -Recurse -ErrorAction Stop
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $true
    }

    <#
     # [File]�t�@�C���E�f�B���N�g���̍폜
     #
     # $path�Ŏw�肷��p�X�̃t�@�C���A�܂��̓f�B���N�g�����폜����B
     # �w��̃t�@�C���E�f�B���N�g�������݂��Ȃ��A�܂��̓t�@�C���E�f�B���N�g���폜���s�̏ꍇ$false��Ԃ��B
     # $pathObj�ł́A�t�@�C���E�f�B���N�g���p�X��z��Ŏw�肷��B
     #
     # @access public
     # @param array $pathObj �t�@�C���E�f�B���N�g���p�X
     # @return bool �t�@�C���E�f�B���N�g���̍폜����
     # @see Remove-Item
     # @throws �Ȃ�
     #>
    [bool]Rm([array]$pathObj){
        foreach($path in $pathObj){
            if($this.Rm($path) -eq $false){
                return $false
            }
        }
        return $true
    }

    <#
     # [File]�t�@�C����ZIP���k
     #
     # $src_path�Ŏw�肷��p�X�̃t�@�C���A�܂��̓f�B���N�g����$dst_path�Ŏ����p�X��ZIP���k����B
     # ���̏����̏ꍇ�A$false��Ԃ��B
     # �E$src_path�Ŏw��̃t�@�C���E�f�B���N�g�������݂��Ȃ�
     # �E�t�@�C���E�f�B���N�g���̈��k���s
     # �E$dst_path�����ɑ��݂��Ă���
     # �E$src_path�̃t�@�C���T�C�Y��2GB�ȏ�
     #
     # @access public
     # @param string $src_path ���k�Ώۂ̃t�@�C���E�f�B���N�g���p�X
     # @param string $dst_path ���k��t�@�C���p�X
     # @return bool �t�@�C���E�f�B���N�g���̈��k����
     # @see Compress-Archive, File.FileSize
     # @throws �t�@�C�����k�ŗ�O�������A$false��Ԃ��B
     #>
    [bool]Zip([string]$src_path, [string]$dst_path){
        if($this.IsFile($src_path) -eq $false){
            return $false
        }
        if($this.FileSize($src_path) -ge 2147483648){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("File_size_is_more_than_2GB", @{path=$src_path})
            return $false
        }
        if($this.IsFile($dst_path) -eq $true){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Duplicate_files", @{path=$dst_path})
            return $false
        }
        try{
            Compress-Archive -Path $src_path -DestinationPath $dst_path -ErrorAction Stop
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $true
    }

    <#
     # [File]�t�@�C����ZIP��
     #
     # $src_path�Ŏw�肷��p�X��ZIP�t�@�C����$dst_path�Ŏ����p�X�ɉ𓀂���B
     # ���̏����̏ꍇ�A$false��Ԃ��B
     # �E$src_path�Ŏw��̃t�@�C�������݂��Ȃ�
     # �E�t�@�C���E�f�B���N�g���̉𓀎��s
     # �E$dst_path�����ɑ��݂��Ă���
     #
     # @access public
     # @param string $src_path �𓀑Ώۂ̃t�@�C���p�X
     # @param string $dst_path �𓀌�t�@�C���p�X
     # @return bool �t�@�C���̉𓀐���
     # @see Expand-Archive
     # @throws �t�@�C���𓀂ŗ�O�������A$false��Ԃ��B
     #>
    [bool]UnZip([string]$src_path, [string]$dst_path){
        if(
            ($this.IsFile($src_path) -eq $false) -or
            ($this.IsFile($dst_path) -eq $true)
        ){
            return $false
        }
        try{
            Expand-Archive -Path $src_path -DestinationPath $dst_path -ErrorAction Stop
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $true
    }

    <#
     # [File]�t�@�C���T�C�Y�̎擾
     #
     # $path�Ŏ����t�@�C���̃T�C�Y���o�C�g�P�ʂŕԂ��B
     # $path�����݂��Ȃ��ꍇfalse��Ԃ��B
     #
     # @access public
     # @param string $path �T�C�Y���m�F����t�@�C���̃p�X
     # @return long �t�@�C���T�C�Y�̃o�C�g��
     # @see Expand-Archive
     # @throws �Ȃ�
     #>
    [long]FileSize([string]$path){
        if($this.IsFile($path) -eq $false){
            return $false
        }
        return $(Get-ChildItem $path).Length
    }

    <#
     # [File]���������Ɉ�v����s�̍폜
     #
     # $file�Ŏ����t�@�C�����̍s�̂����A���K�\��$reg�Ɉ�v����s���폜���Č��̃t�@�C���ɏ㏑���ۑ�����B
     # $file�����݂��Ȃ��ꍇ�A�܂��̓t�@�C���̏������ݎ��s���Afalse��Ԃ��B
     #
     # @access public
     # @param string $file ��������t�@�C���̃p�X
     # @param regex $reg �����p���K�\��
     # @return bool �t�@�C���̏������ݐ���
     # @see
     # @throws �t�@�C���������݂ŗ�O�������A$false��Ԃ��B
     #>
    [bool]RemoveLine([string]$file, [regex]$reg){
        if($this.IsFile($file) -eq $false){
            return $false
        }
        [array]$text=(Get-Content $file)
        [int]$row=0

        #$reg.Matches(
        foreach($line in $text){
            $text[$row]=if(($line -match $reg) -eq $true){ $null }
            $row++
        }
        try{
            $text | Out-File $file
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $true
    }

    <#
     # [File]�t�@�C���̌�[�s�̎擾
     #
     # $file�Ŏ����t�@�C���̌�[�s����$row�s���擾����B
     # $file�����݂��Ȃ��ꍇ�A�܂��̓t�@�C���̓ǂݍ��ݎ��s���Afalse��Ԃ��B
     #
     # @access public
     # @param string $file ��������t�@�C���̃p�X
     # @param int $row �擾����s��
     # @return array $result �擾�����s�̓��e
     # @see Get-Content -tail
     # @throws �t�@�C�����e�擾�ŗ�O�������A$false��Ԃ��B
     #>
    [array]Tail([string]$file, [int]$row){
        if($this.IsFile($file) -eq $false){
            return $false
        }
        try{
            [array]$result=(Get-Content -path $file -tail $row)
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $result
    }

    <#
     # [File]�t�@�C���̈ړ�
     #
     # $src_path�Ŏ����t�@�C����$dst_path�Ɉړ�����B
     # $src_path���Ȃ��ꍇ�A�܂��͈ړ����s����false��Ԃ��B
     #
     # @access public
     # @param string $src_path �ړ����t�@�C���p�X
     # @param string $src_path �ړ���t�@�C���p�X
     # @return bool �t�@�C���ړ��̐���
     # @see Move-Item
     # @throws �t�@�C���ړ��ŗ�O�������A$false��Ԃ��B
     #>
    [bool]Move([string]$src_path, [string]$dst_path){
        if($this.IsFile($src_path) -eq $false){
            return $false
        }
        try{
            Move-Item $src_path $dst_path -ErrorAction Stop
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        return $true
    }

    <#
     # [File]�t�@�C���̈ړ��i�����j
     #
     # $src_path�Ŏ����t�@�C����$dst_path�ɋ����I�Ɉړ�����B
     # $src_path���Ȃ��ꍇ�A�܂��͈ړ����s����false��Ԃ��B
     #
     # @access public
     # @param string $src_path �ړ����t�@�C���p�X
     # @param string $src_path �ړ���t�@�C���p�X
     # @return bool �t�@�C���ړ��̐���
     # @see Move-Item
     # @throws �t�@�C���ړ��ŗ�O�������A$false��Ԃ��B
     #>
    [bool]MoveForce([string]$src_path, [string]$dst_path){
        if($this.IsFile($src_path) -eq $false){
            return $false
        }
        try{
            Move-Item $src_path $dst_path -force -ErrorAction Stop
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        return $true
    }

    <#
     # [File]�t�@�C���̃R�s�[
     #
     # $src_path�Ŏ����t�@�C����$dst_path�ɃR�s�[����B
     # $src_path���Ȃ��ꍇ�A�܂��̓R�s�[���s����false��Ԃ��B
     #
     # @access public
     # @param string $src_path �R�s�[���t�@�C���p�X
     # @param string $src_path �R�s�[��t�@�C���p�X
     # @return bool �t�@�C���R�s�[�̐���
     # @see Copy-Item
     # @throws �t�@�C���R�s�[�ŗ�O�������A$false��Ԃ��B
     #>
    [bool]Cp([string]$src_path, [string]$dst_path){
        if($this.IsFile($src_path) -eq $false){
            return $false
        }
        try{
            Copy-Item $src_path $dst_path -ErrorAction Stop
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        return $true
    }

    <#
     # [File]�t�@�C���̃R�s�[�i�����j
     #
     # $src_path�Ŏ����t�@�C����$dst_path�ɋ����I�ɃR�s�[����B
     # $src_path���Ȃ��ꍇ�A�܂��̓R�s�[���s����false��Ԃ��B
     #
     # @access public
     # @param string $src_path �R�s�[���t�@�C���p�X
     # @param string $src_path �R�s�[��t�@�C���p�X
     # @return bool �t�@�C���R�s�[�̐���
     # @see Copy-Item
     # @throws �t�@�C���R�s�[�ŗ�O�������A$false��Ԃ��B
     #>
    [bool]CpForce([string]$src_path, [string]$dst_path){
        if($this.IsFile($src_path) -eq $false){
            return $false
        }
        try{
            Copy-Item $src_path $dst_path -force -ErrorAction Stop
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        return $true
    }

    <#
     # [File]�t�@�C�����e�̎擾
     #
     # $path�Ŏ����p�X�����s�A����сu#�v�Ŏn�܂�s�������s���擾����B
     # $file�����݂��Ȃ��ꍇ�Afalse��Ԃ��B
     #
     # @access public
     # @param string $path �t�@�C���̃p�X
     # @return array $result �擾�����t�@�C�����e
     # @see Get-Content
     # @throws �Ȃ�
     #>
    [array]GetContentsLines([string]$path){
        if($this.IsFile($path) -eq $false){
            return $false
        }
        [array]$result=@()
        $result+=Get-Content -ReadCount 1 $path | ForEach-Object {
            if(($_ -match "(^[\s]*$)|^#") -eq $false){ $_ }
        }
        return $result
    }

    <#
     # [File]�t�@�C�����̎擾
     #
     # $path�Ŏ����p�X����t�@�C�������擾����B
     #
     # @access public
     # @param string $path �t�@�C���̃p�X
     # @return string �擾�����t�@�C����
     # @see [System.IO.Path]::GetFileName
     # @throws �Ȃ�
     #>
    [string]Basename([string]$path){
        return ([System.IO.Path]::GetFileName($path))
    }

    <#
     # [File]�f�B���N�g�����̎擾
     #
     # $path�Ŏ����p�X����f�B���N�g�������擾����B
     #
     # @access public
     # @param string $path �t�@�C���̃p�X
     # @return string �擾�����f�B���N�g����
     # @see [System.IO.Path]::GetDirectoryName
     # @throws �Ȃ�
     #>
    [string]Dirname([string]$path){
        return ([System.IO.Path]::GetDirectoryName($path))
    }

    <#
     # [File]�t�@�C�����X�g�̎擾
     #
     # $dir�Ŏ����f�B���N�g�������̃t�@�C���̃��X�g���擾����B
     # ���X�g�擾���s����$false��Ԃ��B
     #
     # @access public
     # @param string $dir ��������f�B���N�g���̃p�X
     # @return array $list �擾�����t�@�C���̃t���p�X
     # @see Get-ChildItem
     # @throws �t�@�C���ꗗ�擾�ŗ�O�������A$false��Ԃ��B
     #>
    [array]Filelist([string]$dir){
        try{
            [array]$list=(Get-ChildItem $dir -Recurse | select-object fullname)
            return $list
        }catch{
            return $false
        }
    }

    <#
     # [File]CSV�t�@�C�����I�u�W�F�N�g�Ɋi�[
     #
     # $path�Ŏ����t�@�C����CSV�t�@�C���Ƃ��ăI�u�W�F�N�g�ɕϊ�����B�t�@�C����1�s�ڂ̓w�b�_�Ƃ��Ĉ����B
     # �t�@�C�������݂��Ȃ��ꍇ�A�܂��̓I�u�W�F�N�g�ϊ����s����$false��Ԃ��B
     # $this.encording�ŕϊ����̃G���R�[�h���w�肷��B�W���ł́udefault�v�������ݒ肷��B
     #
     # @access public
     # @param string $path CSV�t�@�C���̃p�X
     # @return object $csvObj �ϊ���̃I�u�W�F�N�g
     # @see Import-Csv
     # @throws CSV����̕ϊ��ŗ�O�������A$false��Ԃ��B
     #>
    [object]ImportCSV([string]$path){
        if($this.IsFile($path) -eq $false){
            return $false
        }
        try{
            $csvObj=Import-Csv $path -Encoding $this.encording
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        return $csvObj
    }

    <#
     # [File]CSV�t�@�C�����I�u�W�F�N�g�Ɋi�[
     #
     # $path�Ŏ����t�@�C����CSV�t�@�C���Ƃ��ăI�u�W�F�N�g�ɕϊ�����B$header��CSV�̃w�b�_���w�肷��B
     # �t�@�C�������݂��Ȃ��ꍇ�A�w�b�_�z�񂪋�̏ꍇ�A�܂��̓I�u�W�F�N�g�ϊ����s����$false��Ԃ��B
     # $this.encording�ŕϊ����̃G���R�[�h���w�肷��B�W���ł́udefault�v�������ݒ肷��B
     #
     # @access public
     # @param string $path CSV�t�@�C���̃p�X
     # @param array $header CSV�̃w�b�_
     # @return object $csvObj �ϊ���̃I�u�W�F�N�g
     # @see Import-Csv
     # @throws CSV����̕ϊ��ŗ�O�������A$false��Ԃ��B
     #>
    [object]ImportCSV([string]$path, [array]$header){
        if($this.IsFile($path) -eq $false){
            return $false
        }
        if($header.Length -eq 0){
            return $false
        }
        try{
            $csvObj=Import-Csv $path -Encoding $this.encording -Header $header
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        return $csvObj
    }

    <#
     # [File]�I�u�W�F�N�g��CSV�`���ɕϊ����t�@�C���Ɋi�[
     #
     # $csvObj��CSV�`���ɕϊ����A$path�Ŏ����t�@�C���ɕۑ�����B
     # �t�@�C����ۑ�����f�B���N�g�������݂��Ȃ��ꍇ�A�܂��̓I�u�W�F�N�g�ϊ����s����$false��Ԃ��B
     # $this.encording�ŕϊ����̃G���R�[�h���w�肷��B�W���ł́udefault�v�������ݒ肷��B
     #
     # @access public
     # @param object $csvObj CSV�ϊ��O�̃I�u�W�F�N�g
     # @param string $path �ۑ����CSV�t�@�C���̃t���p�X
     # @return bool CSV�t�@�C���ւ̕ϊ�����
     # @see Export-Csv
     # @throws �I�u�W�F�N�g����CSV�ւ̕ϊ��ŗ�O�������A$false��Ԃ��B
     #>
    [bool]OutputCSV([object]$csvObj, [string]$path){
        $dirname=(Split-Path $path)
        if($this.IsFile($dirname) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("Directory_does_not_exist", @{path=$dirname})
            return $false
        }
        try{
            $csvObj | Export-Csv -NoTypeInformation $path -Encoding $this.encording
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        return $true
    }
}