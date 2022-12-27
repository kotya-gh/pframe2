<#
 # [Apps]サービス・プロセス・アプリケーション情報取得用クラス
 #
 # ホスト上の稼働サービス・プロセスやインストールアプリケーションの情報等の処理を定義する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category サービス・プロセス・アプリケーション情報取得
 # @package なし
 #>
class Apps{
    [array]$applist    
    [array]$servicelist
    [object]$INTL

    Apps(){
        $this.INTL=New-Object Intl((Join-Path $PSScriptRoot "locale"))
    }

    <#
     # [Apps]インストールアプリケーション情報の取得
     #
     # インストールアプリケーションのリストを配列で返す。
     #
     # @access public
     # @param なし
     # @return array インストールアプリケーションの一覧
     # @see Apps.SetInstalledApps
     # @throws なし
     #>
    [array]GetInstalledApps(){
        $this.applist=$this.SetInstalledApps()
        return $this.applist
    }

    <#
     # [Apps]サービス情報の取得
     #
     # サービスのリストを配列で返す。
     # Apps.servicelist.allに全サービスの情報を格納する。
     # Apps.servicelist.runningに稼働サービスの情報を格納する。
     # Apps.servicelist.stoppedに停止サービスの情報を格納する。
     #
     # @access public
     # @param なし
     # @return array 導入サービスの一覧および状態
     # @see Apps.SetService
     # @throws なし
     #>
    [array]GetService(){
        $this.servicelist=$this.SetService()
        $this.servicelist+=@{ "all" = $this.SetService() }
        $this.servicelist+=@{ "running" = ($this.SetService()| Where-Object { $_.Status -eq "running" }) }
        $this.servicelist+=@{ "stopped" = ($this.SetService()| Where-Object { $_.Status -eq "stopped" }) }
        return $this.servicelist
    }

    <#
     # [Apps]サービス情報の取得
     #
     # $hostnameで指定するホストのサービスのリストを配列で返す。
     #
     # @access public
     # @param string $hostname ホスト名の文字列
     # @return array 導入サービスの一覧および状態
     # @see Apps.SetService
     # @throws サービス情報の取得で例外発生時、$falseを返す。
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
     # [Apps]インストールアプリケーション情報の取得
     #
     # インストールアプリケーションのリストを配列で返す。
     # アプリケーションの情報は、名前、バージョン、パブリッシャーを含める。
     # アプリケーション情報取得失敗時は$falseを返す。
     #
     # @access public
     # @param なし
     # @return array $list インストールアプリケーションの情報
     # @see HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKCU:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
     # @throws アプリケーション情報取得で例外発生時、$falseを返す。
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
     # [Apps]サービス情報の取得
     #
     # ホスト上の全てのサービスを配列で返す。
     # サービス情報取得失敗時は$falseを返す。
     #
     # @access public
     # @param なし
     # @return array $list ホスト上のサービスの情報
     # @see Get-Service
     # @throws サービス情報取得で例外発生時、$falseを返す。
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
     # [Apps]アプリケーションインストール有無の検索
     #
     # $appnameで指定するアプリケーションがインストールされているかどうかをboolで返す。
     # 未インストール時、およびインストール情報取得失敗時は$falseを返す。
     #
     # @access public
     # @param string $appname 検索するアプリケーション名
     # @return bool アプリケーション導入有無
     # @see Apps.GetInstalledApps
     # @throws なし
     #>
    [bool]AppSearch([string]$appname){
        if($this.GetInstalledApps() -eq $false){
            return $false
        }
        return $this.ArraySearchUtf8($appname, $this.applist.DisplayName)
    }

    <#
     # [Apps]サービス登録有無の検索
     #
     # $serviceで指定するサービスが導入されているかどうかをboolで返す。
     # 未導入時、および導入情報取得失敗時は$falseを返す。
     #
     # @access public
     # @param string $service 検索するサービス名
     # @return bool サービス導入有無
     # @see Apps.GetService
     # @throws なし
     #>
    [bool]ServiceSearch([string]$service){
        if($this.GetService() -eq $false){
            return $false
        }
        return $this.ArraySearchUtf8($service, $this.servicelist.Name)
    }

    <#
     # [Apps]サービス登録有無の検索
     #
     # $hostnameで指定するホストに、$serviceで指定するサービスが導入されているかどうかをboolで返す。
     # 未導入時、および導入情報取得失敗時は$falseを返す。
     #
     # @access public
     # @param string $service 検索するサービス名
     # @param string $hostname ホスト名の文字列
     # @return bool サービス導入有無
     # @see Apps.GetService
     # @throws なし
     #>
    [bool]ServiceSearch([string]$service, [string]$hostname){
        [array]$list=$this.GetService($hostname)
        if($list -eq $false){
            return $false
        }
        return $this.ArraySearchUtf8($service, $list.Name)
    }

    <#
     # [Apps]稼働プロセスの検索
     #
     # $processで指定するプロセスが稼働しているかどうかを稼働プロセス数で返す。
     # 未稼働時、またはプロセス情報取得失敗時は$falseを返す。
     #
     # @access public
     # @param string $process 検索するプロセス名
     # @return int 稼働プロセス数
     # @see Get-Process
     # @throws なし
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
     # [Apps]稼働サービスの検索
     #
     # $serviceで指定するサービスが稼働しているかどうかをboolで返す。
     # 指定サービスがRunningの場合$trueを返す。
     # 未稼働時、およびサービス情報取得失敗時は$falseを返す。
     #
     # @access public
     # @param string $service 検索するサービス名
     # @return bool サービス稼働状況
     # @see Apps.GetService
     # @throws なし
     #>
    [bool]ServiceStatus([string]$service){
        if($this.GetService() -eq $false){
            return $false
        }
        return $this.ArraySearchUtf8($service, $this.servicelist.running.Name)
    }

    <#
     # [Apps]稼働サービスの検索
     #
     # $hostnameで指定するホストの、$serviceで指定するサービスが稼働しているかどうかをboolで返す。
     # 指定サービスがRunningの場合$trueを返す。
     # 未稼働時、およびサービス情報取得失敗時は$falseを返す。
     #
     # @access public
     # @param string $service 検索するサービス名
     # @param string $hostname ホスト名の文字列
     # @return bool サービス稼働状況
     # @see Apps.GetService
     # @throws なし
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
     # [Apps]停止サービスの検索
     #
     # $serviceで指定するサービスが停止しているかどうかをboolで返す。
     # 指定サービスがStoppedの場合$trueを返す。
     # 稼働時、およびサービス情報取得失敗時は$falseを返す。
     #
     # @access public
     # @param string $service 検索するサービス名
     # @return bool サービス稼働状況
     # @see Apps.GetService
     # @throws なし
     #>
     [bool]ServiceStatusStopped([string]$service){
        if($this.GetService() -eq $false){
            return $false
        }
        return $this.ArraySearchUtf8($service, $this.servicelist.stopped.Name)
    }

    <#
     # [Apps]停止サービスの検索
     #
     # $hostnameで指定するホストの、$serviceで指定するサービスが停止しているかどうかをboolで返す。
     # 指定サービスがStoppedの場合$trueを返す。
     # 稼働時、およびサービス情報取得失敗時は$falseを返す。
     #
     # @access public
     # @param string $service 検索するサービス名
     # @param string $hostname ホスト名の文字列
     # @return bool サービス稼働状況
     # @see Apps.GetService
     # @throws なし
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
     # [Apps]サービスの起動
     #
     # $serviceで指定するサービスを起動する。
     # 指定サービスがRunningの場合、および起動成功時$trueを返す。
     # 指定サービスが存在しない場合、およびサービス起動失敗時は$falseを返す。
     #
     # @access public
     # @param string $service 起動するサービス名
     # @return bool サービス起動結果の状態
     # @see Start-Service
     # @throws サービス起動で例外発生時、$falseを返す。
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
     # [Apps]サービスの起動
     #
     # $hostnameで指定するホストの$serviceで指定するサービスを起動する。
     # 指定サービスがRunningの場合、および起動成功時$trueを返す。
     # 指定サービスが存在しない場合、およびサービス起動失敗時は$falseを返す。
     #
     # @access public
     # @param string $service 起動するサービス名
     # @param string $hostname ホスト名の文字列
     # @return bool サービス起動結果の状態
     # @see Start-Service
     # @throws サービス起動で例外発生時、$falseを返す。
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
     # [Apps]サービスの起動
     #
     # $serviceで指定する配列から1要素ずつ文字列を読み込み、サービスを起動する。
     # 指定サービスがRunningの場合、および起動成功時$trueを返す。
     # １つでも指定サービスが存在しない場合、および１つでもサービス起動に失敗した場合は$falseを返す。
     #
     # @access public
     # @param array $services 起動するサービス名の配列
     # @return bool サービス起動結果の状態
     # @see Apps.ServiceStart
     # @throws サービス起動で例外発生時、$falseを返す。
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
     # [Apps]サービスの停止
     #
     # $serviceで指定するサービスを停止する。
     # 指定サービスがStoppedの場合、および停止成功時$trueを返す。
     # 指定サービスが存在しない場合、およびサービス停止失敗時は$falseを返す。
     #
     # @access public
     # @param string $service 停止するサービス名
     # @return bool サービス停止結果の状態
     # @see Stop-Service
     # @throws サービス停止で例外発生時、$falseを返す。
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
     # [Apps]サービスの停止
     #
     # $hostnameで指定するホスト名の$serviceで指定するサービスを停止する。
     # 指定サービスがStoppedの場合、および停止成功時$trueを返す。
     # 指定サービスが存在しない場合、およびサービス停止失敗時は$falseを返す。
     #
     # @access public
     # @param string $service 停止するサービス名
     # @param string $hostname ホスト名の文字列
     # @return bool サービス停止結果の状態
     # @see Stop-Service
     # @throws サービス停止で例外発生時、$falseを返す。
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
     # [Apps]サービスの停止
     #
     # $serviceで指定する配列から1要素ずつ文字列を読み込み、サービスを停止する。
     # 指定サービスがStoppedの場合、および停止成功時$trueを返す。
     # １つでも指定サービスが存在しない場合、および１つでもサービス停止に失敗した場合は$falseを返す。
     #
     # @access public
     # @param array $services 停止するサービス名の配列
     # @return bool サービス停止結果の状態
     # @see Apps.ServiceStop
     # @throws サービス停止で例外発生時、$falseを返す。
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
     # [Apps]サービスの再起動
     #
     # $serviceで指定するサービスを再起動する。
     # サービス再起動成功時$trueを返す。
     # 指定サービスが存在しない場合、およびサービス再起動失敗時は$falseを返す。
     #
     # @access public
     # @param string $service 再起動するサービス名
     # @return bool サービス再起動結果の状態
     # @see Restart-Service
     # @throws サービス再起動で例外発生時、$falseを返す。
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
     # [Apps]サービスの再起動
     #
     # $serviceで指定する配列から1要素ずつ文字列を読み込み、サービスを再起動する。
     # サービス再起動成功時$trueを返す。
     # １つでも指定サービスが存在しない場合、および１つでもサービス再起動に失敗した場合は$falseを返す。
     #
     # @access public
     # @param array $services 再起動するサービス名の配列
     # @return bool サービス再起動結果の状態
     # @see Apps.ServiceRestart
     # @throws サービス再起動で例外発生時、$falseを返す。
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
     # [Apps]文字列の検索
     #
     # 配列$haystackに文字列$needleが含まれているかを判定し、結果をboolで返す。
     # 検索文字列はutf-8に変換し、SHIFT-JISにおける一部の機種依存文字列を%3fに変換した上で配列内を検索する。
     # 変換対象文字列はEncode.ConvertSpecialCharにて定義する。
     # 検索で一致しない場合、および$needleが空文字列の場合は$falseを返す。
     #
     # @access public
     # @param string $needle 検索文字列
     # @param array $haystack 検索対象の配列     
     # @return bool 検索結果
     # @see Encode.ConvertSpecialChar
     # @throws なし
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


