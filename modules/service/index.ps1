<#
 # [Apps]�T�[�r�X�E�v���Z�X�E�A�v���P�[�V�������擾�p�N���X
 #
 # �z�X�g��̉ғ��T�[�r�X�E�v���Z�X��C���X�g�[���A�v���P�[�V�����̏�񓙂̏������`����B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category �T�[�r�X�E�v���Z�X�E�A�v���P�[�V�������擾
 # @package �Ȃ�
 #>
class Apps{
    [array]$applist    
    [array]$servicelist
    [object]$INTL

    Apps(){
        $this.INTL=New-Object Intl((Join-Path $PSScriptRoot "locale"))
    }

    <#
     # [Apps]�C���X�g�[���A�v���P�[�V�������̎擾
     #
     # �C���X�g�[���A�v���P�[�V�����̃��X�g��z��ŕԂ��B
     #
     # @access public
     # @param �Ȃ�
     # @return array �C���X�g�[���A�v���P�[�V�����̈ꗗ
     # @see Apps.SetInstalledApps
     # @throws �Ȃ�
     #>
    [array]GetInstalledApps(){
        $this.applist=$this.SetInstalledApps()
        return $this.applist
    }

    <#
     # [Apps]�T�[�r�X���̎擾
     #
     # �T�[�r�X�̃��X�g��z��ŕԂ��B
     # Apps.servicelist.all�ɑS�T�[�r�X�̏����i�[����B
     # Apps.servicelist.running�ɉғ��T�[�r�X�̏����i�[����B
     # Apps.servicelist.stopped�ɒ�~�T�[�r�X�̏����i�[����B
     #
     # @access public
     # @param �Ȃ�
     # @return array �����T�[�r�X�̈ꗗ����я��
     # @see Apps.SetService
     # @throws �Ȃ�
     #>
    [array]GetService(){
        $this.servicelist=$this.SetService()
        $this.servicelist+=@{ "all" = $this.SetService() }
        $this.servicelist+=@{ "running" = ($this.SetService()| Where-Object { $_.Status -eq "running" }) }
        $this.servicelist+=@{ "stopped" = ($this.SetService()| Where-Object { $_.Status -eq "stopped" }) }
        return $this.servicelist
    }

    <#
     # [Apps]�T�[�r�X���̎擾
     #
     # $hostname�Ŏw�肷��z�X�g�̃T�[�r�X�̃��X�g��z��ŕԂ��B
     #
     # @access public
     # @param string $hostname �z�X�g���̕�����
     # @return array �����T�[�r�X�̈ꗗ����я��
     # @see Apps.SetService
     # @throws �T�[�r�X���̎擾�ŗ�O�������A$false��Ԃ��B
     #>
    [array]GetService([string]$hostname){
        try{
            [array]$list=(Get-Service -ComputerName $hostname)
        } catch {
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_get_services_list")
            return $false 
        }        
        return $list
    }

    <#
     # [Apps]�C���X�g�[���A�v���P�[�V�������̎擾
     #
     # �C���X�g�[���A�v���P�[�V�����̃��X�g��z��ŕԂ��B
     # �A�v���P�[�V�����̏��́A���O�A�o�[�W�����A�p�u���b�V���[���܂߂�B
     # �A�v���P�[�V�������擾���s����$false��Ԃ��B
     #
     # @access public
     # @param �Ȃ�
     # @return array $list �C���X�g�[���A�v���P�[�V�����̏��
     # @see HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKCU:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
     # @throws �A�v���P�[�V�������擾�ŗ�O�������A$false��Ԃ��B
     #>
    [array]SetInstalledApps(){
        try{
            [array]$list=Get-ChildItem -Path(
                'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                'HKCU:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall') | 
                ForEach-Object { Get-ItemProperty $_.PsPath | Select-Object DisplayName, DisplayVersion, Publisher }
        } catch {
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_get_install_apps_list")
            return $false 
        }
        return $list
    }

    <#
     # [Apps]�T�[�r�X���̎擾
     #
     # �z�X�g��̑S�ẴT�[�r�X��z��ŕԂ��B
     # �T�[�r�X���擾���s����$false��Ԃ��B
     #
     # @access public
     # @param �Ȃ�
     # @return array $list �z�X�g��̃T�[�r�X�̏��
     # @see Get-Service
     # @throws �T�[�r�X���擾�ŗ�O�������A$false��Ԃ��B
     #>
    [array]SetService(){
        try{
            [array]$list=(Get-Service)
        } catch {
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_get_services_list")
            return $false 
        }        
        return $list
    }

    <#
     # [Apps]�A�v���P�[�V�����C���X�g�[���L���̌���
     #
     # $appname�Ŏw�肷��A�v���P�[�V�������C���X�g�[������Ă��邩�ǂ�����bool�ŕԂ��B
     # ���C���X�g�[�����A����уC���X�g�[�����擾���s����$false��Ԃ��B
     #
     # @access public
     # @param string $appname ��������A�v���P�[�V������
     # @return bool �A�v���P�[�V���������L��
     # @see Apps.GetInstalledApps
     # @throws �Ȃ�
     #>
    [bool]AppSearch([string]$appname){
        if($this.GetInstalledApps() -eq $false){
            return $false
        }
        return $this.ArraySearchUtf8($appname, $this.applist.DisplayName)
    }

    <#
     # [Apps]�T�[�r�X�o�^�L���̌���
     #
     # $service�Ŏw�肷��T�[�r�X����������Ă��邩�ǂ�����bool�ŕԂ��B
     # ���������A����ѓ������擾���s����$false��Ԃ��B
     #
     # @access public
     # @param string $service ��������T�[�r�X��
     # @return bool �T�[�r�X�����L��
     # @see Apps.GetService
     # @throws �Ȃ�
     #>
    [bool]ServiceSearch([string]$service){
        if($this.GetService() -eq $false){
            return $false
        }
        return $this.ArraySearchUtf8($service, $this.servicelist.Name)
    }

    <#
     # [Apps]�T�[�r�X�o�^�L���̌���
     #
     # $hostname�Ŏw�肷��z�X�g�ɁA$service�Ŏw�肷��T�[�r�X����������Ă��邩�ǂ�����bool�ŕԂ��B
     # ���������A����ѓ������擾���s����$false��Ԃ��B
     #
     # @access public
     # @param string $service ��������T�[�r�X��
     # @param string $hostname �z�X�g���̕�����
     # @return bool �T�[�r�X�����L��
     # @see Apps.GetService
     # @throws �Ȃ�
     #>
    [bool]ServiceSearch([string]$service, [string]$hostname){
        [array]$list=$this.GetService($hostname)
        if($list -eq $false){
            return $false
        }
        return $this.ArraySearchUtf8($service, $list.Name)
    }

    <#
     # [Apps]�ғ��v���Z�X�̌���
     #
     # $process�Ŏw�肷��v���Z�X���ғ����Ă��邩�ǂ������ғ��v���Z�X���ŕԂ��B
     # ���ғ����A�܂��̓v���Z�X���擾���s����$false��Ԃ��B
     #
     # @access public
     # @param string $process ��������v���Z�X��
     # @return int �ғ��v���Z�X��
     # @see Get-Process
     # @throws �Ȃ�
     #>
    [int]ProcessSearch([string]$process){
        [array]$list=(Get-Process | select-object ProcessName)
        [int]$count=0
        foreach($proc in $list){
            if($proc.ProcessName -eq $process){
                $count++
            }
        }
        return $count
    }

    <#
     # [Apps]�ғ��T�[�r�X�̌���
     #
     # $service�Ŏw�肷��T�[�r�X���ғ����Ă��邩�ǂ�����bool�ŕԂ��B
     # �w��T�[�r�X��Running�̏ꍇ$true��Ԃ��B
     # ���ғ����A����уT�[�r�X���擾���s����$false��Ԃ��B
     #
     # @access public
     # @param string $service ��������T�[�r�X��
     # @return bool �T�[�r�X�ғ���
     # @see Apps.GetService
     # @throws �Ȃ�
     #>
    [bool]ServiceStatus([string]$service){
        if($this.GetService() -eq $false){
            return $false
        }
        return $this.ArraySearchUtf8($service, $this.servicelist.running.Name)
    }

    <#
     # [Apps]�ғ��T�[�r�X�̌���
     #
     # $hostname�Ŏw�肷��z�X�g�́A$service�Ŏw�肷��T�[�r�X���ғ����Ă��邩�ǂ�����bool�ŕԂ��B
     # �w��T�[�r�X��Running�̏ꍇ$true��Ԃ��B
     # ���ғ����A����уT�[�r�X���擾���s����$false��Ԃ��B
     #
     # @access public
     # @param string $service ��������T�[�r�X��
     # @param string $hostname �z�X�g���̕�����
     # @return bool �T�[�r�X�ғ���
     # @see Apps.GetService
     # @throws �Ȃ�
     #>
    [bool]ServiceStatus([string]$service, [string]$hostname){
        [array]$list=$this.GetService($hostname)
        if($list -eq $false){
            return $false
        }
        [array]$running=($list | Where-Object { $_.Status -eq "running" })
        return $this.ArraySearchUtf8($service, $running.Name)
    }

    <#
     # [Apps]��~�T�[�r�X�̌���
     #
     # $service�Ŏw�肷��T�[�r�X����~���Ă��邩�ǂ�����bool�ŕԂ��B
     # �w��T�[�r�X��Stopped�̏ꍇ$true��Ԃ��B
     # �ғ����A����уT�[�r�X���擾���s����$false��Ԃ��B
     #
     # @access public
     # @param string $service ��������T�[�r�X��
     # @return bool �T�[�r�X�ғ���
     # @see Apps.GetService
     # @throws �Ȃ�
     #>
     [bool]ServiceStatusStopped([string]$service){
        if($this.GetService() -eq $false){
            return $false
        }
        return $this.ArraySearchUtf8($service, $this.servicelist.stopped.Name)
    }

    <#
     # [Apps]��~�T�[�r�X�̌���
     #
     # $hostname�Ŏw�肷��z�X�g�́A$service�Ŏw�肷��T�[�r�X����~���Ă��邩�ǂ�����bool�ŕԂ��B
     # �w��T�[�r�X��Stopped�̏ꍇ$true��Ԃ��B
     # �ғ����A����уT�[�r�X���擾���s����$false��Ԃ��B
     #
     # @access public
     # @param string $service ��������T�[�r�X��
     # @param string $hostname �z�X�g���̕�����
     # @return bool �T�[�r�X�ғ���
     # @see Apps.GetService
     # @throws �Ȃ�
     #>
    [bool]ServiceStatusStopped([string]$service, [string]$hostname){
        [array]$list=$this.GetService($hostname)
        if($list -eq $false){
            return $false
        }
        [array]$stopped=($list | Where-Object { $_.Status -eq "stopped" })
        return $this.ArraySearchUtf8($service, $stopped.Name)
    }

    <#
     # [Apps]�T�[�r�X�̋N��
     #
     # $service�Ŏw�肷��T�[�r�X���N������B
     # �w��T�[�r�X��Running�̏ꍇ�A����ыN��������$true��Ԃ��B
     # �w��T�[�r�X�����݂��Ȃ��ꍇ�A����уT�[�r�X�N�����s����$false��Ԃ��B
     #
     # @access public
     # @param string $service �N������T�[�r�X��
     # @return bool �T�[�r�X�N�����ʂ̏��
     # @see Start-Service
     # @throws �T�[�r�X�N���ŗ�O�������A$false��Ԃ��B
     #>
    [bool]ServiceStart([string]$service){
        if($this.ServiceSearch($service) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Service_is_not_found", @{service=$service})
            return $false
        }
        if($this.ServiceStatus($service) -eq $true){
            $script:LAST_ERROR_MESSAGE=""
            return $true
        }
        try{
            Start-Service -Name $service            
        }catch{
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_start_service", @{service=$service})
            return $false 
        }
        return $true
    }

    <#
     # [Apps]�T�[�r�X�̋N��
     #
     # $hostname�Ŏw�肷��z�X�g��$service�Ŏw�肷��T�[�r�X���N������B
     # �w��T�[�r�X��Running�̏ꍇ�A����ыN��������$true��Ԃ��B
     # �w��T�[�r�X�����݂��Ȃ��ꍇ�A����уT�[�r�X�N�����s����$false��Ԃ��B
     #
     # @access public
     # @param string $service �N������T�[�r�X��
     # @param string $hostname �z�X�g���̕�����
     # @return bool �T�[�r�X�N�����ʂ̏��
     # @see Start-Service
     # @throws �T�[�r�X�N���ŗ�O�������A$false��Ԃ��B
     #>
    [bool]ServiceStart([string]$service, [string]$hostname){
        if($this.ServiceSearch($service, $hostname) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Service_is_not_found", @{service=$service})
            return $false
        }
        if($this.ServiceStatus($service, $hostname) -eq $true){
            $script:LAST_ERROR_MESSAGE=""
            return $true
        }
        try{
            Get-Service -Name $service -ComputerName $hostname | Start-Service
        }catch{
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_start_service", @{service=$service})
            return $false 
        }
        return $true
    }

    <#
     # [Apps]�T�[�r�X�̋N��
     #
     # $service�Ŏw�肷��z�񂩂�1�v�f���������ǂݍ��݁A�T�[�r�X���N������B
     # �w��T�[�r�X��Running�̏ꍇ�A����ыN��������$true��Ԃ��B
     # �P�ł��w��T�[�r�X�����݂��Ȃ��ꍇ�A����тP�ł��T�[�r�X�N���Ɏ��s�����ꍇ��$false��Ԃ��B
     #
     # @access public
     # @param array $services �N������T�[�r�X���̔z��
     # @return bool �T�[�r�X�N�����ʂ̏��
     # @see Apps.ServiceStart
     # @throws �T�[�r�X�N���ŗ�O�������A$false��Ԃ��B
     #>
    [bool]ServiceStart([array]$services){
        foreach($service in $services){
            if($this.ServiceStart($service) -eq $false){
                return $false
            }            
        }
        return $true
    }

    <#
     # [Apps]�T�[�r�X�̒�~
     #
     # $service�Ŏw�肷��T�[�r�X���~����B
     # �w��T�[�r�X��Stopped�̏ꍇ�A����ђ�~������$true��Ԃ��B
     # �w��T�[�r�X�����݂��Ȃ��ꍇ�A����уT�[�r�X��~���s����$false��Ԃ��B
     #
     # @access public
     # @param string $service ��~����T�[�r�X��
     # @return bool �T�[�r�X��~���ʂ̏��
     # @see Stop-Service
     # @throws �T�[�r�X��~�ŗ�O�������A$false��Ԃ��B
     #>
    [bool]ServiceStop([string]$service){
        if($this.ServiceSearch($service) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Service_is_not_found", @{service=$service})
            return $false
        }
        if($this.ServiceStatusStopped($service) -eq $true){
            $script:LAST_ERROR_MESSAGE=""
            return $true
        }
        try{
            Stop-Service -Name $service
        }catch{
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_stop_service", @{service=$service})
            return $false 
        }
        return $true
    }

    <#
     # [Apps]�T�[�r�X�̒�~
     #
     # $hostname�Ŏw�肷��z�X�g����$service�Ŏw�肷��T�[�r�X���~����B
     # �w��T�[�r�X��Stopped�̏ꍇ�A����ђ�~������$true��Ԃ��B
     # �w��T�[�r�X�����݂��Ȃ��ꍇ�A����уT�[�r�X��~���s����$false��Ԃ��B
     #
     # @access public
     # @param string $service ��~����T�[�r�X��
     # @param string $hostname �z�X�g���̕�����
     # @return bool �T�[�r�X��~���ʂ̏��
     # @see Stop-Service
     # @throws �T�[�r�X��~�ŗ�O�������A$false��Ԃ��B
     #>
    [bool]ServiceStop([string]$service, [string]$hostname){
        if($this.ServiceSearch($service, $hostname) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Service_is_not_found", @{service=$service})
            return $false
        }
        if($this.ServiceStatusStopped($service, $hostname) -eq $true){
            $script:LAST_ERROR_MESSAGE=""
            return $true
        }
        try{
            Get-Service -Name $service -ComputerName $hostname | Stop-Service -Force
        }catch{
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_stop_service", @{service=$service})
            return $false 
        }
        return $true
    }

    <#
     # [Apps]�T�[�r�X�̒�~
     #
     # $service�Ŏw�肷��z�񂩂�1�v�f���������ǂݍ��݁A�T�[�r�X���~����B
     # �w��T�[�r�X��Stopped�̏ꍇ�A����ђ�~������$true��Ԃ��B
     # �P�ł��w��T�[�r�X�����݂��Ȃ��ꍇ�A����тP�ł��T�[�r�X��~�Ɏ��s�����ꍇ��$false��Ԃ��B
     #
     # @access public
     # @param array $services ��~����T�[�r�X���̔z��
     # @return bool �T�[�r�X��~���ʂ̏��
     # @see Apps.ServiceStop
     # @throws �T�[�r�X��~�ŗ�O�������A$false��Ԃ��B
     #>
    [bool]ServiceStop([array]$services){
        foreach($service in $services){
            if($this.ServiceStop($service) -eq $false){
                return $false
            }            
        }
        return $true
    }

    <#
     # [Apps]�T�[�r�X�̍ċN��
     #
     # $service�Ŏw�肷��T�[�r�X���ċN������B
     # �T�[�r�X�ċN��������$true��Ԃ��B
     # �w��T�[�r�X�����݂��Ȃ��ꍇ�A����уT�[�r�X�ċN�����s����$false��Ԃ��B
     #
     # @access public
     # @param string $service �ċN������T�[�r�X��
     # @return bool �T�[�r�X�ċN�����ʂ̏��
     # @see Restart-Service
     # @throws �T�[�r�X�ċN���ŗ�O�������A$false��Ԃ��B
     #>
    [bool]ServiceRestart([string]$service){
        if($this.ServiceSearch($service) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Service_is_not_found", @{service=$service})
            return $false
        }
        try{
            Restart-Service -Name $service
        }catch{
            $script:LAST_ERROR_MESSAGE=$this.INTL.FormattedMessage("Fail_to_restart_service", @{service=$service})
            return $false 
        }
        return $true
    }

    <#
     # [Apps]�T�[�r�X�̍ċN��
     #
     # $service�Ŏw�肷��z�񂩂�1�v�f���������ǂݍ��݁A�T�[�r�X���ċN������B
     # �T�[�r�X�ċN��������$true��Ԃ��B
     # �P�ł��w��T�[�r�X�����݂��Ȃ��ꍇ�A����тP�ł��T�[�r�X�ċN���Ɏ��s�����ꍇ��$false��Ԃ��B
     #
     # @access public
     # @param array $services �ċN������T�[�r�X���̔z��
     # @return bool �T�[�r�X�ċN�����ʂ̏��
     # @see Apps.ServiceRestart
     # @throws �T�[�r�X�ċN���ŗ�O�������A$false��Ԃ��B
     #>
    [bool]ServiceRestart([array]$services){
        foreach($service in $services){
            if($this.ServiceRestart($service) -eq $false){
                return $false
            }            
        }
        return $true
    }

    <#
     # [Apps]������̌���
     #
     # �z��$haystack�ɕ�����$needle���܂܂�Ă��邩�𔻒肵�A���ʂ�bool�ŕԂ��B
     # �����������utf-8�ɕϊ����ASHIFT-JIS�ɂ�����ꕔ�̋@��ˑ��������%3f�ɕϊ�������Ŕz�������������B
     # �ϊ��Ώە������Encode.ConvertSpecialChar�ɂĒ�`����B
     # �����ň�v���Ȃ��ꍇ�A�����$needle���󕶎���̏ꍇ��$false��Ԃ��B
     #
     # @access public
     # @param string $needle ����������
     # @param array $haystack �����Ώۂ̔z��     
     # @return bool ��������
     # @see Encode.ConvertSpecialChar
     # @throws �Ȃ�
     #>
    [bool]ArraySearchUtf8([string]$needle, [array]$haystack){
        if(($needle -as "bool") -eq $false){
            return $false
        }
        $array=New-Object ExArray
        $encode=New-Object Encode
        [string]$code="utf-8"
        $needle=$encode.Urlencode($needle, $code)
        return $array.ArraySearch($needle, $encode.ConvertSpecialChar($encode.Urlencode($haystack, $code)))
    }
}


