<#
 # [Ipv4Calc]IPアドレス計算用クラス
 #
 # IPアドレスの計算等の処理を定義する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category IPアドレスの計算
 # @package なし
 #>
 class Ipv4Calc{
    
    Ipv4Calc(){
        
    }

    <#
     # [Ipv4Calc]IPアドレス文字列の変換
     #
     # IPアドレス文字列からドットを除いて8ビットごとに分割し、各数値を配列に格納する。
     # 不正なIPアドレスの場合falseで返す。
     #
     # @access public
     # @param string $ipaddress ドット表記（DDN）のIPアドレスの文字列
     # @return array IPアドレスを8ビットごとに分割した数値を格納した配列
     # @see なし
     # @throws なし
     #>
    [array]GetAddressBytes($ipaddress){
        return [System.Net.IPAddress]::Parse($ipaddress).GetAddressBytes()
    }

    <#
     # [Ipv4Calc]IPアドレス配列の変換
     #
     # IPアドレスを8ビットごとに分割した各数値を格納した配列を、ドット表記（DDN）のIPアドレス文字列に変換する。
     # 不正なIPアドレスの場合falseで返す。
     #
     # @access public
     # @param array $ipaddress IPアドレスを8ビットごとに分割した数値を格納した配列
     # @return string ドット表記（DDN）のIPアドレスの文字列
     # @see なし
     # @throws なし
     #>
    [string]ParseIpByteToString([array]$ipaddress){
        return [System.Net.IPAddress]::Parse(($ipaddress -join ".")).ToString()
    }

    <#
     # [Ipv4Calc]IPアドレスの計算
     #
     # IPアドレスとネットワークマスクからネットワーク情報を計算する。次の情報を表示する。
     # Address：IPアドレス
     # Netmask：サブネットマスク
     # Wildcard：ワイルドカードアドレス
     # Network：ネットワークアドレス
     # Broadcast：ブロードキャストアドレス
     # HostMin：同ネットワーク内で使用可能なIPアドレスの最小値
     # HostMax：同ネットワーク内で使用可能なIPアドレスの最大値
     # Hosts_Net：使用可能IPアドレス残数
     # 不正な入力値の場合falseで返す。
     #
     # @access public
     # @param string $ipaddr IPアドレスの文字列
     # @param string $netmask サブネットマスクの文字列
     # @return object 計算結果を配列で返す
     # @see なし
     # @throws なし
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
     # [Ipv4Calc]ブロードキャストアドレスの計算
     #
     # IPアドレスとサブネットマスクから、ブロードキャストアドレスを計算する。
     #
     # @access public
     # @param string $ipaddr IPアドレスの文字列
     # @param string $netmask サブネットマスクの文字列
     # @return string ブロードキャストアドレスを文字列で返す
     # @see なし
     # @throws なし
     #>
    [string]GetBroadcastAddress([string]$ipaddr, [string]$netmask){
        # IP アドレスをバイト配列に変換する
        $ipBytes = $this.GetAddressBytes($ipaddr)

        # ネットワーク マスクをバイト配列に変換する
        $maskBytes = $this.GetAddressBytes($netmask)

        # ブロードキャスト アドレスを出力する
        $broadcast = $this.GetAddressBytes("0.0.0.0")
        for ($i = 0; $i -lt 4; $i++) {
            $broadcast[$i]= [byte]($ipBytes[$i] -bor (255 -bxor $maskBytes[$i]))
        }
        $broadcast = $this.ParseIpByteToString($broadcast)
        return $broadcast
    }

    <#
     # [Ipv4Calc]ワイルドカードアドレスの計算
     #
     # サブネットマスクから、ワイルドカードアドレスを計算する。
     #
     # @access public
     # @param string $netmask サブネットマスクの文字列
     # @return string ワイルドカードアドレスを文字列で返す
     # @see なし
     # @throws なし
     #>
    [string]GetWildcardAddress([string]$netmask){
        # ネットワーク マスクをバイト配列に変換する
        $maskBytes = $this.GetAddressBytes($netmask)

        # ワイルドカード アドレスを出力する
        $wildcard = $this.GetAddressBytes("255.255.255.255")
        for ($i = 0; $i -lt 4; $i++) {
            $wildcard[$i] = [byte]([math]::Abs($wildcard[$i] - $maskBytes[$i]))
        }
        $wildcard = $this.ParseIpByteToString($wildcard)
        return $wildcard
    }

    <#
     # [Ipv4Calc]ネットワークアドレスの計算
     #
     # IPアドレスとサブネットマスクから、ネットワークアドレスを計算する。
     #
     # @access public
     # @param string $ipaddr IPアドレスの文字列
     # @param string $netmask サブネットマスクの文字列
     # @return string ネットワークアドレスを文字列で返す
     # @see なし
     # @throws なし
     #>
    [string]GetNetworkAddress([string]$ipaddr, [string]$netmask){
        # IP アドレスをバイト配列に変換する
        $ipBytes = $this.GetAddressBytes($ipaddr)

        # ネットワーク マスクをバイト配列に変換する
        $maskBytes = $this.GetAddressBytes($netmask)

        # ブロードキャスト アドレスを出力する
        $network = $this.GetAddressBytes("0.0.0.0")
        for ($i = 0; $i -lt 4; $i++) {
            $network[$i] = [byte]($ipBytes[$i] -band $maskBytes[$i])
        }
        $network = $this.ParseIpByteToString($network)
        return $network
    }

    <#
     # [Ipv4Calc]ネットワークアドレスの計算
     #
     # IPアドレスとサブネットマスクから、ネットワークアドレスとCIDRを計算する。
     #
     # @access public
     # @param string $ipaddr IPアドレスの文字列
     # @param string $netmask サブネットマスクの文字列
     # @return string ネットワークアドレスをCIDRを含めた文字列で返す
     # @see なし
     # @throws なし
     #>
    [string]GetNetmask([string]$ipaddr, [string]$netmask){
        # ネットワーク マスクをバイト配列に変換する
        $maskBytes = $this.GetAddressBytes($netmask)

        $network = $this.GetNetworkAddress($ipaddr, $netmask)
        # ネットワーク マスクのプレフィックス長を取得する
        $prefix = 0
        for ($i = 0; $i -lt 4; $i++) {
            for ($j = 0; $j -lt 8; $j++) {
                if (($maskBytes[$i] -band (1 -shl (7 - $j))) -ne 0) {
                    $prefix++
                } else {
                    break
                }
            }
        }
        # スラッシュ表記のネットワーク アドレスを出力する
        return "$network/$prefix"
    }

    <#
     # [Ipv4Calc]使用可能な最小のIPアドレスの計算する
     #
     # IPアドレスとサブネットマスクから、使用可能な最小のIPアドレスを計算する。
     #
     # @access public
     # @param string $ipaddr IPアドレスの文字列
     # @param string $netmask サブネットマスクの文字列
     # @return string 使用可能な最小のIPアドレスを文字列で返す
     # @see なし
     # @throws なし
     #>
    [string]GetMinAddress([string]$ipaddr, [string]$netmask){
 
        # IP アドレスをバイト配列に変換する
        $ipBytes = $this.GetAddressBytes($ipaddr)

        # ネットワーク アドレスを求める
        $network = $this.GetAddressBytes($this.GetNetworkAddress($ipaddr, $netmask))

        # 同じネットワーク内で使用可能な IP アドレスの最小値を出力する
        # (IP アドレスと同じ場合は +1 して出力する)
        $minIp = $this.ParseIpByteToString($network)
        # IP アドレスを 1 増やす
        $ipBytes = $this.GetAddressBytes($minIp)
        $ipBytes[3]++
        $minIp = $this.ParseIpByteToString($ipBytes)
        return $minIp
    }

    <#
     # [Ipv4Calc]使用可能な最大のIPアドレスの計算する
     #
     # IPアドレスとサブネットマスクから、使用可能な最大のIPアドレスを計算する。
     #
     # @access public
     # @param string $ipaddr IPアドレスの文字列
     # @param string $netmask サブネットマスクの文字列
     # @return string 使用可能な最大のIPアドレスを文字列で返す
     # @see なし
     # @throws なし
     #>
    [string]GetMaxAddress([string]$ipaddr, [string]$netmask){
        # ブロードキャスト アドレスを求める
        $broadcast = $this.GetAddressBytes($this.GetBroadcastAddress($ipaddr, $netmask))

        # 同じネットワーク内で使用可能な IP アドレスの最大値を出力する
        # (ブロードキャスト アドレスより -1 して出力する)
        #$maxIp = $this.ParseIpByteToString($broadcast)
        #$maxIpBytes = $this.GetAddressBytes($maxIp)
        #$maxIpBytes[3]--
        #$maxIp = $this.ParseIpByteToString($maxIpBytes)

        $broadcast[3]--
        $maxIp = $this.ParseIpByteToString($broadcast)
        return $maxIp
    }

    <#
     # [Ipv4Calc]IPアドレスのバリデーション
     #
     # 入力値がIPv4アドレスのDDNフォーマットであるかを判定し、boolで返す。
     # 不正なIPアドレスの場合falseで返す。
     #
     # @access public
     # @param string $ipaddress IPアドレスの文字列
     # @return bool バリデーションの成否
     # @see なし
     # @throws なし
     #>
    [bool]ValidateIPaddress([string]$ipaddress) {
        return ($ipaddress -as [System.Net.IPAddress]).IPAddressToString -eq $ipaddress -and ($null -ne $ipaddress)
    }

    <#
     # [Ipv4Calc]サブネットマスクのバリデーション
     #
     # 入力値がIPv4アドレスのDDNフォーマットであるか、および正しいサブネットマスクを判定し、boolで返す。
     # 不正なサブネットマスクの場合falseで返す。
     #
     # @access public
     # @param string $networkaddress サブネットマスクの文字列
     # @return bool バリデーションの成否
     # @see なし
     # @throws なし
     #>
    [bool]ValidateNetworkaddress([string]$networkaddress) {
        if($this.ValidateIPaddress($networkaddress) -eq $true){
            # 入力された IP アドレスが IPv4 ネットワーク アドレスの形式かどうかを判定する
            $ipBytes = $this.GetAddressBytes($networkaddress)
            $res=$true
            $mid=0
            for($j=0; $j -lt 4; $j++){
                $comp=128
                $ad=$ipBytes[$j]
                for($i=0; $i -lt 8; $i++){
                    $ad=($ad -bxor $comp)
                    if($ad -gt $comp){
                        $res=$false
                        break
                    }elseif($ad -eq $comp){
                        $res=($res -and  $true)
                        break
                    }
                    $comp=($comp -shr 1)
                }
                if(($ipBytes[$j] -ne 0) -and ($ipBytes[$j] -ne 255)){
                    $mid++
                }
            }
            if($mid -gt 1){
                return $false
            }elseif(($res -eq $true) -and ($ipBytes[0] -ge $ipBytes[1]) -and ($ipBytes[1] -ge $ipBytes[2]) -and ($ipBytes[2] -ge $ipBytes[3])){
                return $true
            }       
        }
        return $false
    }

    <#
     # [Ipv4Calc]CIDR表記のサブネットマスクの変換
     #
     # CIDR表記のサブネットマスクから、ドット表記（DDN）のサブネットマスク文字列に変換する。
     #
     # @access public
     # @param int $cidr CIDRの数値
     # @return string ドット表記（DDN）のサブネットマスク文字列
     # @see なし
     # @throws なし
     #>
    [string]GetNetworkAddressFromCidr([int]$cidr){
        $netmask = $this.GetAddressBytes("255.255.255.255")
        for($j=0; $j -lt $netmask.Length; $j++){
            $mask=$cidr
            if($cidr -ge 8){
                $mask=8
            }elseif($cidr -lt 0){
                $mask=0
            }
            $c=255
            for($i=0; $i -lt $mask; $i++){
                $c--
                $c=($c -shr 1)
            }
            $netmask[$j]=$netmask[$j] -bxor $c
            $cidr-=8
        }
        $netmask = $this.ParseIpByteToString($netmask)
        return $netmask 
    }

    <#
     # [Ipv4Calc]使用可能なIPアドレスの数の計算する
     #
     # ブロードキャストアドレスとネットワークアドレスから、使用可能なIPアドレスの数を計算する。
     #
     # @access public
     # @param string $broadCast ブロードキャストアドレスの文字列
     # @param string $network ネットワークアドレスの文字列
     # @return Int64 使用可能なIPアドレスを数を数値で返す
     # @see なし
     # @throws なし
     #>
    [Int64]HostsPerNet($broadCast, $network){
        $broadCastBytes=$this.GetAddressBytes($broadCast)
        $networkBytes=$this.GetAddressBytes($network)

        [Int64]$bnum=0
        [Int64]$nnum=0
        for($i=0; $i -lt 4; $i++){
            $bnum+=$broadCastBytes[$i]*[Math]::Pow(256, 3-$i)
            $nnum+=$networkBytes[$i]*[Math]::Pow(256, 3-$i)
        }
        $Hostspernet = ($bnum - $nnum - 1 )
        return $Hostspernet
   }
}