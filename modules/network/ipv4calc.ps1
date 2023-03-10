<#
 # [Ipv4Calc]IP�A�h���X�v�Z�p�N���X
 #
 # IP�A�h���X�̌v�Z���̏������`����B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category IP�A�h���X�̌v�Z
 # @package �Ȃ�
 #>
 class Ipv4Calc{
    
    Ipv4Calc(){
        
    }

    <#
     # [Ipv4Calc]IP�A�h���X������̕ϊ�
     #
     # IP�A�h���X�����񂩂�h�b�g��������8�r�b�g���Ƃɕ������A�e���l��z��Ɋi�[����B
     # �s����IP�A�h���X�̏ꍇfalse�ŕԂ��B
     #
     # @access public
     # @param string $ipaddress �h�b�g�\�L�iDDN�j��IP�A�h���X�̕�����
     # @return array IP�A�h���X��8�r�b�g���Ƃɕ����������l���i�[�����z��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [array]GetAddressBytes($ipaddress){
        return [System.Net.IPAddress]::Parse($ipaddress).GetAddressBytes()
    }

    <#
     # [Ipv4Calc]IP�A�h���X�z��̕ϊ�
     #
     # IP�A�h���X��8�r�b�g���Ƃɕ��������e���l���i�[�����z����A�h�b�g�\�L�iDDN�j��IP�A�h���X������ɕϊ�����B
     # �s����IP�A�h���X�̏ꍇfalse�ŕԂ��B
     #
     # @access public
     # @param array $ipaddress IP�A�h���X��8�r�b�g���Ƃɕ����������l���i�[�����z��
     # @return string �h�b�g�\�L�iDDN�j��IP�A�h���X�̕�����
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [string]ParseIpByteToString([array]$ipaddress){
        return [System.Net.IPAddress]::Parse(($ipaddress -join ".")).ToString()
    }

    <#
     # [Ipv4Calc]IP�A�h���X�̌v�Z
     #
     # IP�A�h���X�ƃl�b�g���[�N�}�X�N����l�b�g���[�N�����v�Z����B���̏���\������B
     # Address�FIP�A�h���X
     # Netmask�F�T�u�l�b�g�}�X�N
     # Wildcard�F���C���h�J�[�h�A�h���X
     # Network�F�l�b�g���[�N�A�h���X
     # Broadcast�F�u���[�h�L���X�g�A�h���X
     # HostMin�F���l�b�g���[�N���Ŏg�p�\��IP�A�h���X�̍ŏ��l
     # HostMax�F���l�b�g���[�N���Ŏg�p�\��IP�A�h���X�̍ő�l
     # Hosts_Net�F�g�p�\IP�A�h���X�c��
     # �s���ȓ��͒l�̏ꍇfalse�ŕԂ��B
     #
     # @access public
     # @param string $ipaddr IP�A�h���X�̕�����
     # @param string $netmask �T�u�l�b�g�}�X�N�̕�����
     # @return object �v�Z���ʂ�z��ŕԂ�
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [object]GetIpv4Calc([string]$ipaddr, [string]$netmask){
        if($this.ValidateIPaddress($ipaddr) -eq $false){
            return $false
        }
        if($this.ValidateNetworkaddress($netmask) -eq $false){
            return $false
        }
        <#
            Address   : 192.168.64.120
            Netmask   : 255.255.252.0
            Wildcard  : 0.0.3.255
            Network   : 192.168.64.0/22
            Broadcast : 192.168.67.255
            HostMin   : 192.168.64.1
            HostMax   : 192.168.67.254
            Hosts_Net : 1022
        #>
        $broadCast=$this.GetBroadcastAddress($ipaddr, $netmask)
        $network=$this.GetNetworkAddress($ipaddr, $netmask)
        [object]$calc=@{
            Address=$ipaddr;
            Netmask=$netmask;
            Wildcard=$this.GetWildcardAddress($netmask);
            Network=$this.GetNetmask($ipaddr, $netmask);
            Broadcast=$broadCast;
            HostMin=$this.GetMinAddress($ipaddr, $netmask);
            HostMax=$this.GetMaxAddress($ipaddr, $netmask);
            HostsPerNet=$this.HostsPerNet($broadCast, $network)
        }
        return $calc
    }

    <#
     # [Ipv4Calc]�u���[�h�L���X�g�A�h���X�̌v�Z
     #
     # IP�A�h���X�ƃT�u�l�b�g�}�X�N����A�u���[�h�L���X�g�A�h���X���v�Z����B
     #
     # @access public
     # @param string $ipaddr IP�A�h���X�̕�����
     # @param string $netmask �T�u�l�b�g�}�X�N�̕�����
     # @return string �u���[�h�L���X�g�A�h���X�𕶎���ŕԂ�
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [string]GetBroadcastAddress([string]$ipaddr, [string]$netmask){
        # IP �A�h���X���o�C�g�z��ɕϊ�����
        $ipBytes = $this.GetAddressBytes($ipaddr)

        # �l�b�g���[�N �}�X�N���o�C�g�z��ɕϊ�����
        $maskBytes = $this.GetAddressBytes($netmask)

        # �u���[�h�L���X�g �A�h���X���o�͂���
        $broadcast = $this.GetAddressBytes("0.0.0.0")
        for ($i = 0; $i -lt 4; $i++) {
            $broadcast[$i]= [byte]($ipBytes[$i] -bor (255 -bxor $maskBytes[$i]))
        }
        return $this.ParseIpByteToString($broadcast)
    }

    <#
     # [Ipv4Calc]���C���h�J�[�h�A�h���X�̌v�Z
     #
     # �T�u�l�b�g�}�X�N����A���C���h�J�[�h�A�h���X���v�Z����B
     #
     # @access public
     # @param string $netmask �T�u�l�b�g�}�X�N�̕�����
     # @return string ���C���h�J�[�h�A�h���X�𕶎���ŕԂ�
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [string]GetWildcardAddress([string]$netmask){
        # �l�b�g���[�N �}�X�N���o�C�g�z��ɕϊ�����
        $maskBytes = $this.GetAddressBytes($netmask)

        # ���C���h�J�[�h �A�h���X���o�͂���
        # 
        # ���C���h�J�[�h �A�h���X�̓T�u�l�b�g�}�X�N�ƃr�b�g�P�ʂŔ��]�����邱�Ƃŋ��߂���B
        # ��F�T�u�l�b�g�}�X�N��255.255.255.0�Ȃ�A���C���h�J�[�h�A�h���X��0.0.0.255�ƂȂ�B
        # ���̏����ł́A�܂��S��1�̃o�C�g�z��i255.255.255.255�j���쐬���A���ꂩ��T�u�l�b�g�}�X�N�ƈ����Z���Đ�Βl����邱�ƂŔ��]�����Ă���B
        $wildcard = $this.GetAddressBytes("255.255.255.255")
        for ($i = 0; $i -lt 4; $i++) {
            $wildcard[$i] = [byte]([math]::Abs($wildcard[$i] - $maskBytes[$i]))
        }
        # �o�C�g�z�񂩂當����ɕϊ�����
        return $this.ParseIpByteToString($wildcard)
    }

    <#
     # [Ipv4Calc]�l�b�g���[�N�A�h���X�̌v�Z
     #
     # IP�A�h���X�ƃT�u�l�b�g�}�X�N����A�l�b�g���[�N�A�h���X���v�Z����B
     #
     # @access public
     # @param string $ipaddr IP�A�h���X�̕�����
     # @param string $netmask �T�u�l�b�g�}�X�N�̕�����
     # @return string �l�b�g���[�N�A�h���X�𕶎���ŕԂ�
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [string]GetNetworkAddress([string]$ipaddr, [string]$netmask){
        # IP �A�h���X���o�C�g�z��ɕϊ�����
        $ipBytes = $this.GetAddressBytes($ipaddr)

        # �l�b�g���[�N �}�X�N���o�C�g�z��ɕϊ�����
        $maskBytes = $this.GetAddressBytes($netmask)

        # �u���[�h�L���X�g �A�h���X���o�͂���
        $network = $this.GetAddressBytes("0.0.0.0")
        for ($i = 0; $i -lt 4; $i++) {
            $network[$i] = [byte]($ipBytes[$i] -band $maskBytes[$i])
        }
        return $this.ParseIpByteToString($network)
    }

    <#
     # [Ipv4Calc]�l�b�g���[�N�A�h���X�̌v�Z
     #
     # IP�A�h���X�ƃT�u�l�b�g�}�X�N����A�l�b�g���[�N�A�h���X��CIDR���v�Z����B
     #
     # @access public
     # @param string $ipaddr IP�A�h���X�̕�����
     # @param string $netmask �T�u�l�b�g�}�X�N�̕�����
     # @return string �l�b�g���[�N�A�h���X��CIDR���܂߂�������ŕԂ�
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [string]GetNetmask([string]$ipaddr, [string]$netmask){
        # �l�b�g���[�N �}�X�N���o�C�g�z��ɕϊ�����
        $maskBytes = $this.GetAddressBytes($netmask)

        # �l�b�g���[�N�A�h���X���擾����
        $network = $this.GetNetworkAddress($ipaddr, $netmask)
       
        # �l�b�g�}�X�N�A�h���X����o�C�g�z����擾�����]����
        [array]::Reverse($maskBytes)

        # �o�C�g�z�񂩂�r�b�g�z��i0/1�j�ɕϊ�����
        $maskBits = [System.Collections.BitArray]::new($maskBytes)

        # �r�b�g�z�񂩂�v���t�B�b�N�X���i0����n�܂�ŏ���1�܂Łj�����߂�
        $prefix = 0
        foreach ($bit in $maskBits) {
            if ($bit) {
                break
            }
            $prefix++
        }
        # �v���t�B�b�N�X������T�u�l�b�g�}�X�N���i32 - �v���t�B�b�N�X���j�����߂ďo�͂���
        $prefix=$maskBits.Length-$prefix

        # �X���b�V���\�L�̃l�b�g���[�N �A�h���X���o�͂���
        return "$network/$prefix"
    }

    <#
     # [Ipv4Calc]�g�p�\�ȍŏ���IP�A�h���X�̌v�Z����
     #
     # IP�A�h���X�ƃT�u�l�b�g�}�X�N����A�g�p�\�ȍŏ���IP�A�h���X���v�Z����B
     #
     # @access public
     # @param string $ipaddr IP�A�h���X�̕�����
     # @param string $netmask �T�u�l�b�g�}�X�N�̕�����
     # @return string �g�p�\�ȍŏ���IP�A�h���X�𕶎���ŕԂ�
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [string]GetMinAddress([string]$ipaddr, [string]$netmask){
        # �l�b�g���[�N �A�h���X�����߂�
        $network = $this.GetAddressBytes($this.GetNetworkAddress($ipaddr, $netmask))

        # �����l�b�g���[�N���Ŏg�p�\�� IP �A�h���X�̍ŏ��l���o�͂���
        # (IP �A�h���X�Ɠ����ꍇ�� +1 ���ďo�͂���)
        $network[3]++
        return $this.ParseIpByteToString($network)
    }

    <#
     # [Ipv4Calc]�g�p�\�ȍő��IP�A�h���X�̌v�Z����
     #
     # IP�A�h���X�ƃT�u�l�b�g�}�X�N����A�g�p�\�ȍő��IP�A�h���X���v�Z����B
     #
     # @access public
     # @param string $ipaddr IP�A�h���X�̕�����
     # @param string $netmask �T�u�l�b�g�}�X�N�̕�����
     # @return string �g�p�\�ȍő��IP�A�h���X�𕶎���ŕԂ�
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [string]GetMaxAddress([string]$ipaddr, [string]$netmask){
        # �u���[�h�L���X�g �A�h���X�����߂�
        $broadcast = $this.GetAddressBytes($this.GetBroadcastAddress($ipaddr, $netmask))

        # �����l�b�g���[�N���Ŏg�p�\�� IP �A�h���X�̍ő�l���o�͂���
        # (�u���[�h�L���X�g �A�h���X��� -1 ���ďo�͂���)
        $broadcast[3]--
        return $this.ParseIpByteToString($broadcast)
    }

    <#
     # [Ipv4Calc]IP�A�h���X�̃o���f�[�V����
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

    <#
     # [Ipv4Calc]�T�u�l�b�g�}�X�N�̃o���f�[�V����
     #
     # ���͒l��IPv4�A�h���X��DDN�t�H�[�}�b�g�ł��邩�A����ѐ������T�u�l�b�g�}�X�N�𔻒肵�Abool�ŕԂ��B
     # �s���ȃT�u�l�b�g�}�X�N�̏ꍇfalse�ŕԂ��B
     #
     # @access public
     # @param string $networkaddress �T�u�l�b�g�}�X�N�̕�����
     # @return bool �o���f�[�V�����̐���
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]ValidateNetworkaddress([string]$networkaddress) {
        if($this.ValidateIPaddress($networkaddress) -eq $true){
            # ���͂��ꂽ IP �A�h���X�� IPv4 �l�b�g���[�N �A�h���X�̌`�����ǂ����𔻒肷��
            $ipBytes = $this.GetAddressBytes($networkaddress)

            # �e�I�N�e�b�g��2�i���ɕϊ������Ƃ��ɘA������1��0�̃p�^�[���ɂȂ��Ă��邩�ǂ����`�F�b�N
            $Binary = ""
            foreach ($Octet in $ipBytes) {
                # 8���ɂȂ�悤��0��⊮����2�i���ɕϊ�
                $Binary += [Convert]::ToString($Octet, 2).PadLeft(8,"0")
            }
            if ($Binary -match "^1+0+$") {
                # �A������1��0�̃p�^�[���̏ꍇ��True��Ԃ�
                return $True
            }
        }
        return $false
    }

    <#
     # [Ipv4Calc]CIDR�\�L�̃T�u�l�b�g�}�X�N�̕ϊ�
     #
     # CIDR�\�L�̃T�u�l�b�g�}�X�N����A�h�b�g�\�L�iDDN�j�̃T�u�l�b�g�}�X�N������ɕϊ�����B
     #
     # @access public
     # @param int $cidr CIDR�̐��l
     # @return string �h�b�g�\�L�iDDN�j�̃T�u�l�b�g�}�X�N������
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [string]GetNetworkAddressFromCidr([int]$cidr){
        $netmask = $this.GetAddressBytes("255.255.255.255")
        for($j=0; $j -lt $netmask.Length; $j++){
            # CIDR��8�ȏ�Ȃ�8�A0�ȉ��Ȃ�0�A����ȊO�Ȃ炻�̂܂�
            $mask = [Math]::Max([Math]::Min($cidr, 8), 0)
            # �o�C�g�l�����V�t�g���ăl�b�g�}�X�N���쐬
            $netmask[$j] = $netmask[$j] -shl (8 - $mask)
            # CIDR����8������
            $cidr -= 8            
        }
        return $this.ParseIpByteToString($netmask)
    }

    <#
     # [Ipv4Calc]�g�p�\��IP�A�h���X�̐��̌v�Z����
     #
     # �u���[�h�L���X�g�A�h���X�ƃl�b�g���[�N�A�h���X����A�g�p�\��IP�A�h���X�̐����v�Z����B
     #
     # @access public
     # @param string $broadCast �u���[�h�L���X�g�A�h���X�̕�����
     # @param string $network �l�b�g���[�N�A�h���X�̕�����
     # @return Int64 �g�p�\��IP�A�h���X�𐔂𐔒l�ŕԂ�
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [Int64]HostsPerNet($broadCast, $network){
        # �A�h���X����o�C�g�z��i0�`255�j�ɕϊ�����
        $broadCastBytes=$this.GetAddressBytes($broadCast)
        $networkBytes=$this.GetAddressBytes($network)

        # �A�h���X���i4�j
        $addressLength = 4
        # �o�C�g�T�C�Y�i256�j
        $byteSize = 256

        # �o�C�g�z�񂩂�10�i���i0�`4294967295�j�ɕϊ�����
        [Int64]$broadcastNumber=0
        [Int64]$networkNumber=0
        for($i=0; $i -lt $addressLength; $i++){
            $broadcastNumber+=$broadcastBytes[$i]*[Math]::Pow($byteSize, $addressLength-$i-1)
            $networkNumber+=$networkBytes[$i]*[Math]::Pow($byteSize, $addressLength-$i-1)
        }
      
        # 10�i������z�X�g���i�T�u�l�b�g���j�����߂ĕԂ�
        return ($broadcastNumber - $networkNumber - 1 )
   }
}