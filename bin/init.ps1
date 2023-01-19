<#
 # [INIT]初期化処理用クラス
 #
 # コンポーネント実行時Requireする。
 # 共通ディレクトリの定義、設定読み込み、モジュール読み込みを実施する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category 初期化共通処理
 # @package なし
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
        # ルートディレクトリの定義
        $this.path_root=(Split-Path $PSScriptRoot -Parent)

        # スクリプト構成ディレクトリの変数格納処理
        $this.path_conf=$this.GetExistDirPath((Join-Path $this.path_root "conf"))
        $this.path_modules=$this.GetExistDirPath((Join-Path $this.path_root "modules"))
        $this.path_log=$this.GetExistDirPath((Join-Path $this.path_root "log"))
        $this.path_files=$this.GetExistDirPath((Join-Path $this.path_root "files"))

        # スクリプトルートディレクトリの設定
        $this.path_script_root=$path_script_root

        # コンポーネント用外部ファイル名の格納
        $this.conflist_user=(Get-ChildItem $this.path_script_root -include *.json -Name)

        # スクリプト設定の格納
        $this.conf_require=$this.SetRequiredConfigure()

        # コンポーネント名の取得
        $this.component_name=(Split-Path -Leaf $this.GetScriptRootPath())

        # コンポーネントの設定ファイル名の格納
        $this.path_component_conf=(Join-Path $this.path_conf ($this.component_name+".json"))

        # コンポーネントの設定内容の格納
        $this.conf_component=$this.SetComponentConfigure()

        # コンポーネント用外部ファイルのレプリケーション処理
        $this.ReplicationUserConfiture()

        # コンポーネント用外部ファイルの内容格納
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
     # [INIT]ディレクトリパスの存在確認
     #
     # $pathで指定するパスのディレクトリの存在を確認し、存在する場合は入力値、存在しない場合は作成後パスを返す。
     #
     # @access public
     # @param string $path ディレクトリパス
     # @return string $path ディレクトリパス
     # @see New-Item
     # @throws なし
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
     # $pathで指定するパスのファイルの存在を確認し、存在する場合は入力値、
     # 存在しない場合は空ファイルを作成後パスを返す。
     #
     # @access public
     # @param string $path ファイルパス
     # @return string $path ファイルパス
     # @see New-Item
     # @throws なし
     #>
    [string]GetExistFilePath([string]$path){
        if((Test-Path $path) -eq $false){
            New-Item $path -type file -ErrorAction Stop
        }
        return $path
    }

    <#
     # [INIT]スクリプト全体設定ファイルのハッシュ変換
     #
     # スクリプト全体設定ファイル"require.json"を読み込み、Jsonデコード後のハッシュを返す。
     #
     # @access public
     # @param なし
     # @return array Jsonデコード後ハッシュ
     # @see Init.GetExistFilePath, Init.JsonFileDecode
     # @throws Init.JsonFileDecodeで例外発生時、$falseを返す。
     #>
    [array]SetRequiredConfigure(){
        [string]$path_conf_require=$this.GetExistFilePath((Join-Path $this.path_conf "require.json"))
        return $this.JsonFileDecode($path_conf_require)
    }

    <#
     # [Crypt]文字列の復号
     #
     # $encryptで指定する文字列をSecureStringで復号した文字列を返す。
     # $encryptが空、まはた復号失敗時は$falseを返す。
     #
     # @access public
     # @param string $encrypt 暗号化文字列
     # @return string $plaintext 復号文字列
     # @see ConvertTo-SecureString
     # @throws 文字列復号で例外発生時、$falseを返す。
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
     # [INIT]ユーザ設定ファイルのレプリケーション
     #
     # コンポーネントディレクトリ配下のJsonファイル（ユーザ設定ファイル）を、
     # コンポーネント設定ファイルで指定する場所からレプリケーションする。
     # コンポーネント設定ファイルのexecフラグが1の場合、かつ
     # レプリケーション元/先のファイルのハッシュ値が異なる場合レプリケーション処理を実行する。
     #
     # @access public
     # @param なし
     # @return bool レプリケーションの成否
     # @see Copy-Item, Get-FileHash
     # @throws ファイルコピーで例外発生時、$falseを返す。
     #>
    [bool]ReplicationUserConfiture(){
        # exec 1以外はtrueを返すのみとする。
        if($this.conf_component.replication.exec -ne 1){
            return $true
        }

        # レプリケーション先をネットワークドライブに設定する場合の処理
        # ネットワークマウント設定、exec 1のときのみ実施
        if($this.conf_component.replication.netmount.exec -eq 1){
            $computerName = $this.conf_component.replication.netmount.servername
            $adminPass = $this.conf_component.replication.netmount.password
            $adminUser = $this.conf_component.replication.netmount.user
            
            # 認証情報のインスタンスを生成する
            $securePass = ConvertTo-SecureString $adminPass -AsPlainText -Force;
            $cred = New-Object System.Management.Automation.PSCredential "$computerName\$adminUser", $securePass;
            
            New-PSDrive -Name $this.conf_component.replication.netmount.mountto -PSProvider FileSystem -Root "\\$computerName\$this.conf_component.replication.netmount.mountfrom" -Credential $cred;
        }

        foreach($filePath in $this.conf_component.replication.files){
            # ファイルが存在しない場合はスキップ、またはJSON以外の場合はスキップ
            if(
                ((Test-Path $filePath) -eq $false) -or
                ((Get-ChildItem $filePath).Extension.toLower() -ne ".json")
            ){
                continue
            }
            [string]$filePathBasename=([System.IO.Path]::GetFileName($filePath))
            [string]$dstPath=(Join-Path $this.path_script_root $filePathBasename)
            # conflist_userに含まれない場合、新規ファイルと判定し、チェックサム比較なしでコピーする。
            if([Array]::IndexOf($this.conflist_user, $filePathBasename) -eq -1){
                try{
                    Copy-Item $filePath $dstPath -ErrorAction Stop
                } catch {
                    Write-Host($_.Exception)
                    return $false 
                }
            }else{
                # チェックサム比較後、ファイルが異なる場合はコピーする。
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
     # [INIT]ユーザ設定ファイルのJsonデコード
     #
     # コンポーネントディレクトリ配下のJsonファイル（ユーザ設定ファイル）をJsonデコードし、
     # ファイル名（拡張子なし）をキーとするハッシュに格納した結果を返す。
     #
     # @access public
     # @param なし
     # @return array $ret_hash ユーザ設定ファイルのJsonデコード結果
     # @see Init.JsonFileDecode
     # @throws Init.JsonFileDecodeで例外発生時、$falseを返す。
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
     # [INIT]コンポーネント設定ファイルのJsonデコード
     #
     # 共通設定ディレクトリ(conf)配下のJsonファイル（コンポーネント設定ファイル）を
     # Jsonデコードした結果を返す。
     # コンポーネントと同名の設定ファイルが存在しない場合、default.confのデコード結果を返す。
     #
     # @access public
     # @param string Init.path_component_conf
     # @return array $ret_hash ユーザ設定ファイルのJsonデコード結果
     # @see Init.GetExistFilePath, Init.JsonFileDecode
     # @throws Init.JsonFileDecodeで例外発生時、$falseを返す。
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
     # [INIT]Jsonファイルをオブジェクトに変換する
     #
     # $file_pathで指定のファイルのJsonデコード結果を返す。
     #
     # @access public
     # @param string $file_path
     # @return array $decode_array ファイルのJsonデコード結果
     # @see Get-Content, convertFrom-Json
     # @throws Get-Content、およびconvertFrom-Jsonで例外発生時、$falseを返す。
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
# PowerShellバージョンの確認
[int]$version=$PSVersionTable.PSVersion.Major
if($version -le 4){ 
    Write-Host "PowerShell(>=5.0) is required."
    exit 1
}

# initクラスの初期化
$INIT=[Init]::new($MyInvocation.PSScriptRoot)

# include component configure
# 設定ファイルの読み込み、コンポーネント名、ログ格納ディレクトリの設定
[array]$script:COMP_CONF=$INIT.GetComponentConfigure()
$script:COMP_CONF+=@{component_name=$INIT.GetComponentName(); path_log=$INIT.GetLogPath()}

# import configure process
# モジュールクラスの読み込み
foreach($module in $INIT.GetRequiredConfigure().import_module){
    try{
        . (Join-Path $INIT.GetModulesPath() ($module+"\index.ps1") -Resolve)
    } catch { 
        Write-Host($_.Exception)
        exit 1
    }
}

# set user configuration
# ユーザ設定ファイルの読み込み
[array]$USR_CONF=$INIT.GetUserConfigure()

# set intl
# 出力メッセージ定義用処理の初期化
$script:LOCALE=$INIT.conf_require.locale
[object]$INTL=New-Object Intl((Join-Path ([System.IO.Path]::GetDirectoryName($script:myInvocation.ScriptName)) "locale"))

# set user log
[object]$USR_LOG=[log]::new()
[object]$USR_MSG=$USR_LOG.log_format