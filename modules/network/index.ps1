. (Join-Path $PSScriptRoot "ipv4calc.ps1")
<#
 # [Hosts]ホスト名、IPアドレスの書式に関するクラス
 #
 # ホスト名、IPアドレスのバリデーション、hostsファイルの操作等の処理を定義する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category ホスト名、IPアドレス書式
 # @package なし
 #>
class Hosts{
    [array]$define
    [array]$hosts
    
    Hosts(){
        $this.define=(Get-Content (Join-Path $PSScriptRoot "define.json") -Encoding UTF8 -Raw | ConvertFrom-Json)
    }

    <#
     # [Hosts]hostsファイルの内容の取得
     #
     # $define.hosts_pathで指定するパスのhostsファイル内容を取得し、配列で返す。
     # 取得に失敗した場合$falseを返す。
     #
     # @access public
     # @param なし
     # @return array Hosts.hosts hostsファイルの内容
     # @see Hosts.SetHosts
     # @throws なし
     #>
    [array]GetHosts(){
        $this.hosts=$this.SetHosts($this.define.hosts_path)
        return $this.hosts
    }

    <#
     # [Hosts]hostsファイルの内容の取得
     #
     # $hostsFileで指定するパスのhostsファイル内容を取得し、配列で返す。
     # 取得に失敗した場合$falseを返す。
     #
     # @access public
     # @param string $hostsFile hostsファイルのパス
     # @return array Hosts.hosts hostsファイルの内容
     # @see Hosts.SetHosts
     # @throws なし
     #>
    [array]GetHosts([string]$hostsFile){
        $this.hosts=$this.SetHosts($hostsFile)
        return $this.hosts
    }

    <#
     # [Hosts]hostsフィアルの内容の取得
     #
     # $hostsFileで指定するパスのhostsファイル内容を取得し、配列で返す。
     # コメント行は削除し、IPアドレス、ホスト名のみ取得する。IPアドレス、ホスト名は以下のメンバ変数に格納して返す。
     # IPアドレス：$hostsObj.ipaddress
     # ホスト名：$hostsObj.hostname
     # 指定のhostsファイルが存在しない場合、またはIPアドレスとホスト名のバリデーションに失敗した場合$falseを返す。
     #
     # @access public
     # @param string $hostsFile hostsファイルのパス
     # @return array $hostsObj hostsファイルの内容
     # @see Hosts.ValidateHosts
     # @throws なし
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
            # Windows Hostsは1単語目がIPアドレス、2単語目がホスト名
            # ∴ 配列[0]：IPアドレス、配列[1]：ホスト名
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
     # [Hosts]IPアドレスとホスト名のバリデーション
     #
     # IPアドレスとホスト名の形式が正常か判定し、boolで返す。いずれかが不正の場合falseを返す。
     #
     # @access public
     # @param string $ipaddress IPアドレス文字列
     # @param string $hostname ホスト名文字列
     # @return bool バリデーションの成否
     # @see Hosts.ValidateIPaddress, Hosts.ValidateHostname
     # @throws なし
     #>
    [bool]ValidateHosts([string]$ipaddress, [string]$hostname){
        if($this.ValidateIPaddress($ipaddress) -and $this.ValidateHostname($hostname)){
            return $true            
        }
        return $false
    }

    <#
     # [Hosts]IPアドレスとホスト名のバリデーション
     #
     # IPアドレスとホスト名をhash配列で渡し、書式の正常性を判定する。いずれかが不正の場合falseを返す。
     #
     # @access public
     # @param array $hostsObj IPアドレスとホスト名のハッシュ配列
     # @return bool バリデーションの成否
     # @see Hosts.ValidateIPaddress, Hosts.ValidateHostname
     # @throws なし
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
     # [Hosts]ホスト名のバリデーション
     #
     # $hostnameで示す文字列がRFC2396に準拠したホスト名かどうかを判定し、boolで返す。
     # RFC2396準拠ホスト名規約を次に示す。
     # アルファベット大文字（”A”～”Z”）
     # アルファベット小文字（”a”～”z”）
     # 数字（”0”～”9”） (注1)
     # ハイフン（”-“） (注2)
     # ピリオド（”.”） (注2)
     # (注1) 最後のピリオドの直後には、数字は使用不可。
     # (注2) ハイフンおよびピリオドは、ホスト名の先頭文字として使用不可。
     # 不正なホスト名の場合falseで返す。
     #
     # @access public
     # @param string $hostname ホスト名の文字列
     # @return bool バリデーションの成否
     # @see なし
     # @throws なし
     #>
    [bool]ValidateHostname([string]$hostname){
        return ($hostname -match "^[a-zA-Z0-9](([a-zA-Z0-9]|[-])*|([a-zA-Z0-9]|[-.])*([.][a-zA-Z]|[-])([a-zA-Z0-9]|[-])*)$")
    }

    <#
     # [Hosts]IPアドレスのバリデーション
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

    [bool]ValidateNetworkaddress([string]$networkaddress) {
        if($this.ValidateIPaddress($networkaddress) -eq $true){
            # 入力された IP アドレスが IPv4 ネットワーク アドレスの形式かどうかを判定する
            $ipBytes = [System.Net.IPAddress]::Parse($networkaddress).GetAddressBytes()
            if (($ipBytes[3] -eq 0) -and ($ipBytes[2] -eq 0) -and ($ipBytes[1] -eq 0) -and ($ipBytes[0] -ne 0)) {
                return $true
            }            
        }
        return $false
    }
}

<#
 # [NetCom]ネットワーク通信用クラス
 #
 # ネットワーク通信用制御やバリデーション等の処理を定義する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category ネットワーク通信用
 # @package なし
 #>
class NetCom{
    NetCom(){

    }

    <#
     # [NetCom]ポート番号のバリデーション
     #
     # $portで指定する整数がポート番号で使用する番号であるか判定する。
     # ポート番号の範囲が不正である場合、$falseを返す。
     #
     # @access public
     # @param int $port ポート番号
     # @return bool ポート番号のバリデーション成否
     # @see なし
     # @throws なし
     #>
     [bool]validatePort([int]$port){
        return (($port -ge 0) -and ($port -le 65535))
    }
}

<#
 # [NetIF]ネットワークインターフェース操作用クラス
 #
 # ネットワークインターフェースの操作やIPアドレスの計算等の処理を定義する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category ネットワークインターフェース操作
 # @package なし
 #>
class NetIF{
    [object]$Ipv4Calc
    NetIF(){
        $this.Ipv4Calc=New-Object Ipv4Calc
    }
    
    <#
     # [NetIF]ネットワークインターフェース設定情報の取得
     #
     # ホストのネットワークインターフェース設定情報を取得し、配列で返す。次の情報を取得する
     # Description：ネットワークインターフェースの説明
     # DefaultIPGateway：デフォルトゲートウェイアドレス
     # InterfaceIndex：ネットワークインターフェースのインデックス番号
     # DNSServer：DNSサーバIPアドレス（CSV形式）
     # IPv4Address：IPv4アドレス
     # IPv6Address：IPv6アドレス
     # IPv4Subnet：IPv4サブネットマスク
     # IPv6Subnet：IPv6サブネットマスク
     # MACAddress：MACアドレス
     # ServiceName：ネットワークインターフェースのサービス名
     # ネットワークインターフェース設定情報の取得に失敗した場合、$falseを返す。
     #
     # @access public
     # @param なし
     # @return array $ip ネットワークインターフェースの設定情報
     # @see Get-WmiObject
     # @throws ネットワークインターフェース設定情報取得で例外発生時、$falseを返す。
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
     # [NetIF]IPアドレスの計算
     #
     # IPアドレスとネットワークマスクからネットワーク情報を計算する。次の情報を表示する。
     # Address：IPアドレス
     # Netmask：サブネットマスク
     # Wildcard：ワイルドカードアドレス
     # Network：ネットワークアドレス
     # Broadcast：ブロードキャストアドレス
     # HostMin：同ネットワーク内で使用可能なIPアドレスの最小値
     # HostMax：同ネットワーク内で使用可能なIPアドレスの最大値
     # Hosts/Net：使用可能IPアドレス残数
     # 不正な入力値の場合falseで返す。
     #
     # @access public
     # @param string $ipaddr IPアドレスの文字列
     # @param string $netmask サブネットマスクの文字列
     # @return array 計算結果を配列で返す
     # @see ipcalc.ps1
     # @throws なし
     #>
    [object]IpCalc([string]$ipaddr, [string]$netmask){
        return $this.Ipv4Calc.GetIpv4Calc($ipaddr, $netmask)
    }

    <#
     # [NetIF]IPアドレスの計算
     #
     # IPアドレスとネットワークマスクからネットワーク情報を計算する。
     # IPアドレス/サブネットマスク の形式を引数とする。
     # 不正な入力値の場合falseで返す。
     #
     # @access public
     # @param string $ip_mask IPアドレス/サブネットマスクの文字列
     # @return array 計算結果を配列で返す
     # @see ipcalc.ps1
     # @throws なし
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
     # [NetIF]ネットワークインターフェースの無効化
     #
     # $if_nameで指定する文字列の名前のネットワークインターフェースを無効化する。
     # ネットワークインターフェースの無効化に失敗場合、$falseを返す。
     #
     # @access public
     # @param string $if_name ネットワークインターフェース名
     # @return bool ネットワークインターフェース無効化成否
     # @see Get-WmiObject
     # @throws ネットワークインターフェースの無効化で例外発生時、$falseを返す。
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
     # [NetIF]ネットワークインターフェースの有効化
     #
     # $if_nameで指定する文字列の名前のネットワークインターフェースを有効化する。
     # ネットワークインターフェースの有効化に失敗場合、$falseを返す。
     #
     # @access public
     # @param string $if_name ネットワークインターフェース名
     # @return bool ネットワークインターフェース有効化成否
     # @see Get-WmiObject
     # @throws ネットワークインターフェースの有効化で例外発生時、$falseを返す。
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