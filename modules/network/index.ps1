. (Join-Path $PSScriptRoot "ipv4calc.ps1")
<#
 # [Hosts]�z�X�g���AIP�A�h���X�̏����Ɋւ���N���X
 #
 # �z�X�g���AIP�A�h���X�̃o���f�[�V�����Ahosts�t�@�C���̑��쓙�̏������`����B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category �z�X�g���AIP�A�h���X����
 # @package �Ȃ�
 #>
class Hosts{
    [array]$define
    [array]$hosts
    
    Hosts(){
        $this.define=(Get-Content (Join-Path $PSScriptRoot "define.json") -Encoding UTF8 -Raw | ConvertFrom-Json)
    }

    <#
     # [Hosts]hosts�t�@�C���̓��e�̎擾
     #
     # $define.hosts_path�Ŏw�肷��p�X��hosts�t�@�C�����e���擾���A�z��ŕԂ��B
     # �擾�Ɏ��s�����ꍇ$false��Ԃ��B
     #
     # @access public
     # @param �Ȃ�
     # @return array Hosts.hosts hosts�t�@�C���̓��e
     # @see Hosts.SetHosts
     # @throws �Ȃ�
     #>
    [array]GetHosts(){
        $this.hosts=$this.SetHosts($this.define.hosts_path)
        return $this.hosts
    }

    <#
     # [Hosts]hosts�t�@�C���̓��e�̎擾
     #
     # $hostsFile�Ŏw�肷��p�X��hosts�t�@�C�����e���擾���A�z��ŕԂ��B
     # �擾�Ɏ��s�����ꍇ$false��Ԃ��B
     #
     # @access public
     # @param string $hostsFile hosts�t�@�C���̃p�X
     # @return array Hosts.hosts hosts�t�@�C���̓��e
     # @see Hosts.SetHosts
     # @throws �Ȃ�
     #>
    [array]GetHosts([string]$hostsFile){
        $this.hosts=$this.SetHosts($hostsFile)
        return $this.hosts
    }

    <#
     # [Hosts]hosts�t�B�A���̓��e�̎擾
     #
     # $hostsFile�Ŏw�肷��p�X��hosts�t�@�C�����e���擾���A�z��ŕԂ��B
     # �R�����g�s�͍폜���AIP�A�h���X�A�z�X�g���̂ݎ擾����BIP�A�h���X�A�z�X�g���͈ȉ��̃����o�ϐ��Ɋi�[���ĕԂ��B
     # IP�A�h���X�F$hostsObj.ipaddress
     # �z�X�g���F$hostsObj.hostname
     # �w���hosts�t�@�C�������݂��Ȃ��ꍇ�A�܂���IP�A�h���X�ƃz�X�g���̃o���f�[�V�����Ɏ��s�����ꍇ$false��Ԃ��B
     #
     # @access public
     # @param string $hostsFile hosts�t�@�C���̃p�X
     # @return array $hostsObj hosts�t�@�C���̓��e
     # @see Hosts.ValidateHosts
     # @throws �Ȃ�
     #>
    [array]SetHosts([string]$hostsFile){
        [array]$tmpArray=@()
        [array]$hostsObj=@{}
        if((Test-Path $hostsFile) -eq $false){
            return $false
        }
        foreach($hosts in (findstr /V /r /c:"^#" $hostsFile)){
            -split $hosts | ForEach-Object{
                $tmpArray+=$_
            }
            # Windows Hosts��1�P��ڂ�IP�A�h���X�A2�P��ڂ��z�X�g��
            # �� �z��[0]�FIP�A�h���X�A�z��[1]�F�z�X�g��
            [string]$ipaddress=$tmpArray[0]
            [string]$hostname=$tmpArray[1]
            [array]$tmpArray=@()
            if(
                ($ipaddress.StartsWith("#") -eq $true) -or
                ((($ipaddress -as "bool") -eq $false) -and (($hostname -as "bool") -eq $false))
            ){
                continue
            }
            if($this.ValidateHosts($ipaddress, $hostname) -eq $false){
                return $false
            }
            $hostsObj+=@{ipaddress = $ipaddress; hostname = $hostname}
        }
        return $hostsObj
    }

    <#
     # [Hosts]IP�A�h���X�ƃz�X�g���̃o���f�[�V����
     #
     # IP�A�h���X�ƃz�X�g���̌`�������킩���肵�Abool�ŕԂ��B�����ꂩ���s���̏ꍇfalse��Ԃ��B
     #
     # @access public
     # @param string $ipaddress IP�A�h���X������
     # @param string $hostname �z�X�g��������
     # @return bool �o���f�[�V�����̐���
     # @see Hosts.ValidateIPaddress, Hosts.ValidateHostname
     # @throws �Ȃ�
     #>
    [bool]ValidateHosts([string]$ipaddress, [string]$hostname){
        if($this.ValidateIPaddress($ipaddress) -and $this.ValidateHostname($hostname)){
            return $true            
        }
        return $false
    }

    <#
     # [Hosts]IP�A�h���X�ƃz�X�g���̃o���f�[�V����
     #
     # IP�A�h���X�ƃz�X�g����hash�z��œn���A�����̐��퐫�𔻒肷��B�����ꂩ���s���̏ꍇfalse��Ԃ��B
     #
     # @access public
     # @param array $hostsObj IP�A�h���X�ƃz�X�g���̃n�b�V���z��
     # @return bool �o���f�[�V�����̐���
     # @see Hosts.ValidateIPaddress, Hosts.ValidateHostname
     # @throws �Ȃ�
     #>
    [bool]ValidateHosts([array]$hostsObj){
        foreach($hosts in $hostsObj){
            if(($null -eq $hosts.ipaddress) -and ($null -eq $hosts.hostname)){
                continue
            }
            if($this.ValidateHosts($hosts.ipaddress, $hosts.hostname) -eq $false){
                return $false
            }
        }
        return $true
    }

    <#
     # [Hosts]�z�X�g���̃o���f�[�V����
     #
     # $hostname�Ŏ���������RFC2396�ɏ��������z�X�g�����ǂ����𔻒肵�Abool�ŕԂ��B
     # RFC2396�����z�X�g���K������Ɏ����B
     # �A���t�@�x�b�g�啶���i�hA�h�`�hZ�h�j
     # �A���t�@�x�b�g�������i�ha�h�`�hz�h�j
     # �����i�h0�h�`�h9�h�j (��1)
     # �n�C�t���i�h-�g�j (��2)
     # �s���I�h�i�h.�h�j (��2)
     # (��1) �Ō�̃s���I�h�̒���ɂ́A�����͎g�p�s�B
     # (��2) �n�C�t������уs���I�h�́A�z�X�g���̐擪�����Ƃ��Ďg�p�s�B
     # �s���ȃz�X�g���̏ꍇfalse�ŕԂ��B
     #
     # @access public
     # @param string $hostname �z�X�g���̕�����
     # @return bool �o���f�[�V�����̐���
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]ValidateHostname([string]$hostname){
        return ($hostname -match "^[a-zA-Z0-9](([a-zA-Z0-9]|[-])*|([a-zA-Z0-9]|[-.])*([.][a-zA-Z]|[-])([a-zA-Z0-9]|[-])*)$")
    }

    <#
     # [Hosts]IP�A�h���X�̃o���f�[�V����
     #
     # ���͒l��IPv4�A�h���X��DDN�t�H�[�}�b�g�ł��邩�𔻒肵�Abool�ŕԂ��B
     # �s����IP�A�h���X�̏ꍇfalse�ŕԂ��B
     #
     # @access public
     # @param string $ipaddress IP�A�h���X�̕�����
     # @return bool �o���f�[�V�����̐���
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]ValidateIPaddress([string]$ipaddress) {
        return ($ipaddress -as [System.Net.IPAddress]).IPAddressToString -eq $ipaddress -and ($null -ne $ipaddress)
    }

    [bool]ValidateNetworkaddress([string]$networkaddress) {
        if($this.ValidateIPaddress($networkaddress) -eq $true){
            # ���͂��ꂽ IP �A�h���X�� IPv4 �l�b�g���[�N �A�h���X�̌`�����ǂ����𔻒肷��
            $ipBytes = [System.Net.IPAddress]::Parse($networkaddress).GetAddressBytes()
            if (($ipBytes[3] -eq 0) -and ($ipBytes[2] -eq 0) -and ($ipBytes[1] -eq 0) -and ($ipBytes[0] -ne 0)) {
                return $true
            }            
        }
        return $false
    }
}

<#
 # [NetCom]�l�b�g���[�N�ʐM�p�N���X
 #
 # �l�b�g���[�N�ʐM�p�����o���f�[�V�������̏������`����B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category �l�b�g���[�N�ʐM�p
 # @package �Ȃ�
 #>
class NetCom{
    NetCom(){

    }

    <#
     # [NetCom]�|�[�g�ԍ��̃o���f�[�V����
     #
     # $port�Ŏw�肷�鐮�����|�[�g�ԍ��Ŏg�p����ԍ��ł��邩���肷��B
     # �|�[�g�ԍ��͈̔͂��s���ł���ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param int $port �|�[�g�ԍ�
     # @return bool �|�[�g�ԍ��̃o���f�[�V��������
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
     [bool]validatePort([int]$port){
        return (($port -ge 0) -and ($port -le 65535))
    }
}

<#
 # [NetIF]�l�b�g���[�N�C���^�[�t�F�[�X����p�N���X
 #
 # �l�b�g���[�N�C���^�[�t�F�[�X�̑����IP�A�h���X�̌v�Z���̏������`����B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category �l�b�g���[�N�C���^�[�t�F�[�X����
 # @package �Ȃ�
 #>
class NetIF{
    [object]$Ipv4Calc
    NetIF(){
        $this.Ipv4Calc=New-Object Ipv4Calc
    }
    
    <#
     # [NetIF]�l�b�g���[�N�C���^�[�t�F�[�X�ݒ���̎擾
     #
     # �z�X�g�̃l�b�g���[�N�C���^�[�t�F�[�X�ݒ�����擾���A�z��ŕԂ��B���̏����擾����
     # Description�F�l�b�g���[�N�C���^�[�t�F�[�X�̐���
     # DefaultIPGateway�F�f�t�H���g�Q�[�g�E�F�C�A�h���X
     # InterfaceIndex�F�l�b�g���[�N�C���^�[�t�F�[�X�̃C���f�b�N�X�ԍ�
     # DNSServer�FDNS�T�[�oIP�A�h���X�iCSV�`���j
     # IPv4Address�FIPv4�A�h���X
     # IPv6Address�FIPv6�A�h���X
     # IPv4Subnet�FIPv4�T�u�l�b�g�}�X�N
     # IPv6Subnet�FIPv6�T�u�l�b�g�}�X�N
     # MACAddress�FMAC�A�h���X
     # ServiceName�F�l�b�g���[�N�C���^�[�t�F�[�X�̃T�[�r�X��
     # �l�b�g���[�N�C���^�[�t�F�[�X�ݒ���̎擾�Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param �Ȃ�
     # @return array $ip �l�b�g���[�N�C���^�[�t�F�[�X�̐ݒ���
     # @see Get-WmiObject
     # @throws �l�b�g���[�N�C���^�[�t�F�[�X�ݒ���擾�ŗ�O�������A$false��Ԃ��B
     #>
    [array]Ip(){
        try{
            [array]$ip=(Get-WmiObject -Computer . `
            -Class Win32_NetworkAdapterConfiguration `
            -Filter "MACAddress Is Not NULL" | `
            Select-Object `
                Description, `
                DefaultIPGateway, `
                InterfaceIndex, `
                @{ Name="DNSServer"; Expression={$_.DNSServerSearchOrder}},
                @{ Name="IPv4Address"; Expression={$_.IPAddress[0]}},
                @{ Name="IPv6Address"; Expression={$_.IPAddress[1]}},
                @{ Name="IPv4Subnet"; Expression={$_.IPSubnet[0]}},
                @{ Name="IPv6Subnet"; Expression={$_.IPSubnet[1]}},
                MACAddress, `
                ServiceName
            )
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false            
        }
        return $ip
    }

    <#
     # [NetIF]IP�A�h���X�̌v�Z
     #
     # IP�A�h���X�ƃl�b�g���[�N�}�X�N����l�b�g���[�N�����v�Z����B���̏���\������B
     # Address�FIP�A�h���X
     # Netmask�F�T�u�l�b�g�}�X�N
     # Wildcard�F���C���h�J�[�h�A�h���X
     # Network�F�l�b�g���[�N�A�h���X
     # Broadcast�F�u���[�h�L���X�g�A�h���X
     # HostMin�F���l�b�g���[�N���Ŏg�p�\��IP�A�h���X�̍ŏ��l
     # HostMax�F���l�b�g���[�N���Ŏg�p�\��IP�A�h���X�̍ő�l
     # Hosts/Net�F�g�p�\IP�A�h���X�c��
     # �s���ȓ��͒l�̏ꍇfalse�ŕԂ��B
     #
     # @access public
     # @param string $ipaddr IP�A�h���X�̕�����
     # @param string $netmask �T�u�l�b�g�}�X�N�̕�����
     # @return array �v�Z���ʂ�z��ŕԂ�
     # @see ipcalc.ps1
     # @throws �Ȃ�
     #>
    [object]IpCalc([string]$ipaddr, [string]$netmask){
        return $this.Ipv4Calc.GetIpv4Calc($ipaddr, $netmask)
    }

    <#
     # [NetIF]IP�A�h���X�̌v�Z
     #
     # IP�A�h���X�ƃl�b�g���[�N�}�X�N����l�b�g���[�N�����v�Z����B
     # IP�A�h���X/�T�u�l�b�g�}�X�N �̌`���������Ƃ���B
     # �s���ȓ��͒l�̏ꍇfalse�ŕԂ��B
     #
     # @access public
     # @param string $ip_mask IP�A�h���X/�T�u�l�b�g�}�X�N�̕�����
     # @return array �v�Z���ʂ�z��ŕԂ�
     # @see ipcalc.ps1
     # @throws �Ȃ�
     #>
    [object]IpCalc([string]$ip_mask){
        [object]$str=New-Object Str
        [array]$netif=$str.Explode("/", $ip_mask)
        if($netif.Length -eq 2){
            if(
                ($netif[1] -lt 0) -or
                ($netif[1] -gt 32)
            ){
                return $false
            }
        }else{
            return $false
        }
        [string]$netmask=$this.Ipv4Calc.GetNetworkAddressFromCidr($netif[1])

        return $this.Ipv4Calc.GetIpv4Calc($netif[0], $netmask)
    }

    <#
     # [NetIF]�l�b�g���[�N�C���^�[�t�F�[�X�̖�����
     #
     # $if_name�Ŏw�肷�镶����̖��O�̃l�b�g���[�N�C���^�[�t�F�[�X�𖳌�������B
     # �l�b�g���[�N�C���^�[�t�F�[�X�̖������Ɏ��s�ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param string $if_name �l�b�g���[�N�C���^�[�t�F�[�X��
     # @return bool �l�b�g���[�N�C���^�[�t�F�[�X����������
     # @see Get-WmiObject
     # @throws �l�b�g���[�N�C���^�[�t�F�[�X�̖������ŗ�O�������A$false��Ԃ��B
     #>
    [bool]DisableNetIF([string]$if_name){
        try{
            (Get-WmiObject -Class Win32_NetworkAdapter -Filter Name = $if_name).disable()
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        return $true
    }

    <#
     # [NetIF]�l�b�g���[�N�C���^�[�t�F�[�X�̗L����
     #
     # $if_name�Ŏw�肷�镶����̖��O�̃l�b�g���[�N�C���^�[�t�F�[�X��L��������B
     # �l�b�g���[�N�C���^�[�t�F�[�X�̗L�����Ɏ��s�ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param string $if_name �l�b�g���[�N�C���^�[�t�F�[�X��
     # @return bool �l�b�g���[�N�C���^�[�t�F�[�X�L��������
     # @see Get-WmiObject
     # @throws �l�b�g���[�N�C���^�[�t�F�[�X�̗L�����ŗ�O�������A$false��Ԃ��B
     #>
    [bool]EnableNetIF([string]$if_name){
        try{
            (Get-WmiObject -Class Win32_NetworkAdapter -Filter Name = $if_name).enable()
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        return $true
    }
}