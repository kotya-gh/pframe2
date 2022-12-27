<#
 # [File]ファイル・ディレクトリ操作用クラス
 #
 # ファイル、ディレクトリの作成、削除、操作等の処理を定義する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category ファイル・ディレクトリ操作
 # @package なし
 #>
class File{
    [object]$LOCALE
    [string]$encording="default"

    File(){
        $this.LOCALE=New-Object Intl((Join-Path $PSScriptRoot "locale"))
    }

    <#
     # [File]ファイル・ディレクトリの存在確認
     #
     # $pathで指定するパスのファイル、またはディレクトリの存在を確認し、
     # 存在する場合は$true、存在しない場合は$falseを返す。
     #
     # @access public
     # @param string $path ファイル・ディレクトリパス
     # @return bool ファイル、またはディレクトリの存在有無
     # @see Test-Path
     # @throws なし
     #>
    [bool]IsFile([string]$path){
        if((Test-Path $path) -eq $false){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("File_does_not_exist", @{path=$path})
            return $false
        }
        return $true
    }

    <#
     # [File]ファイル・ディレクトリの存在確認
     #
     # $pathObjで指定するパスのファイル、またはディレクトリの存在を確認し、
     # 存在する場合は$true、１つでも存在しない場合は$falseを返す。
     # $pathObjでは、ファイルパス、またはディレクトリパスを配列で指定する。
     #
     # @access public
     # @param array $pathObj ファイル・ディレクトリパス
     # @return bool ファイル、またはディレクトリの存在有無
     # @see Test-Path
     # @throws なし
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
     # [File]空ファイルの作成
     #
     # $pathで指定するパスの空ファイルを作成する。
     # 重複するファイルが存在する場合、またはファイル作成失敗の場合$falseを返す。
     #
     # @access public
     # @param string $path ファイルパス
     # @return bool ファイルの作成成否
     # @see New-Item
     # @throws ファイル作成で例外発生時、$falseを返す。
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
     # [File]空ファイルの作成
     #
     # $pathObjで指定するパスの空ファイルを作成する。
     # 重複するファイルが存在する場合、または１つでもファイル作成失敗の場合$falseを返す。
     # $pathObjでは、ファイルパスを配列で指定する。
     #
     # @access public
     # @param array $pathObj ファイルパス
     # @return bool ファイルの作成成否
     # @see New-Item
     # @throws なし
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
     # [File]ディレクトリの作成
     #
     # $pathで指定するパスのディレクトリを作成する。
     # 重複するディレクトリが存在する場合、またはディレクトリ作成失敗の場合$falseを返す。
     #
     # @access public
     # @param string $path ディレクトリパス
     # @return bool ディレクトリの作成成否
     # @see New-Item
     # @throws ディレクトリの作成で例外発生時、$falseを返す。
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
     # [File]ディレクトリの作成
     #
     # $pathObjで指定するパスのディレクトリを作成する。
     # 重複するディレクトリが存在する場合、または１つでもディレクトリ作成失敗の場合$falseを返す。
     # $pathObjでは、ディレクトリパスを配列で指定する。
     #
     # @access public
     # @param array $pathObj ディレクトリパス
     # @return bool ディレクトリの作成成否
     # @see New-Item
     # @throws なし
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
     # [File]空ファイルの作成（強制）
     #
     # $pathで指定するパスの空ファイルを作成する。
     # 重複するファイルが存在する場合でも強制的にファイルを作成する。
     # ファイル作成失敗の場合$falseを返す。
     #
     # @access public
     # @param string $path ファイルパス
     # @return bool ファイルの作成成否
     # @see New-Item
     # @throws ファイル作成で例外発生時、$falseを返す。
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
     # [File]空ファイルの作成
     #
     # $pathObjで指定するパスの空ファイルを作成する。
     # 重複するファイルが存在する場合でも強制的にファイルを作成する。
     # ファイル作成失敗の場合$falseを返す。
     # $pathObjでは、ファイルパスを配列で指定する。
     #
     # @access public
     # @param array $pathObj ファイルパス
     # @return bool ファイルの作成成否
     # @see New-Item
     # @throws なし
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
     # [File]ディレクトリの作成
     #
     # $pathで指定するパスのディレクトリを作成する。
     # 重複するディレクトリが存在する場合でも強制的にディレクトリを作成する。
     # ディレクトリ作成失敗の場合$falseを返す。
     #
     # @access public
     # @param string $path ディレクトリパス
     # @return bool ディレクトリの作成成否
     # @see New-Item
     # @throws なし
     #>
    [bool]MkDirForce([string]$path){
        if($this.Rm($path) -eq $false){
            return $false
        }
        return $this.MkDir($path)
    }

    <#
     # [File]ディレクトリの作成
     #
     # $pathObjで指定するパスのディレクトリを作成する。
     # 重複するディレクトリが存在する場合でも強制的にディレクトリを作成する。
     # ディレクトリ作成失敗の場合$falseを返す。
     # $pathObjでは、ディレクトリパスを配列で指定する。
     #
     # @access public
     # @param array $pathObj ディレクトリパス
     # @return bool ディレクトリの作成成否
     # @see New-Item
     # @throws なし
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
     # [File]ファイル・ディレクトリの削除
     #
     # $pathで指定するパスのファイル、またはディレクトリを削除する。
     # 指定のファイル・ディレクトリが存在しない、またはファイル・ディレクトリ削除失敗の場合$falseを返す。
     #
     # @access public
     # @param string $path ファイル・ディレクトリパス
     # @return bool ファイル・ディレクトリの削除成否
     # @see Remove-Item
     # @throws ファイル・ディレクトリ削除で例外発生時、$falseを返す。
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
     # [File]ファイル・ディレクトリの削除
     #
     # $pathで指定するパスのファイル、またはディレクトリを削除する。
     # 指定のファイル・ディレクトリが存在しない、またはファイル・ディレクトリ削除失敗の場合$falseを返す。
     # $pathObjでは、ファイル・ディレクトリパスを配列で指定する。
     #
     # @access public
     # @param array $pathObj ファイル・ディレクトリパス
     # @return bool ファイル・ディレクトリの削除成否
     # @see Remove-Item
     # @throws なし
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
     # [File]ファイルのZIP圧縮
     #
     # $src_pathで指定するパスのファイル、またはディレクトリを$dst_pathで示すパスにZIP圧縮する。
     # 次の条件の場合、$falseを返す。
     # ・$src_pathで指定のファイル・ディレクトリが存在しない
     # ・ファイル・ディレクトリの圧縮失敗
     # ・$dst_pathが既に存在している
     # ・$src_pathのファイルサイズが2GB以上
     #
     # @access public
     # @param string $src_path 圧縮対象のファイル・ディレクトリパス
     # @param string $dst_path 圧縮後ファイルパス
     # @return bool ファイル・ディレクトリの圧縮成否
     # @see Compress-Archive, File.FileSize
     # @throws ファイル圧縮で例外発生時、$falseを返す。
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
     # [File]ファイルのZIP解凍
     #
     # $src_pathで指定するパスのZIPファイルを$dst_pathで示すパスに解凍する。
     # 次の条件の場合、$falseを返す。
     # ・$src_pathで指定のファイルが存在しない
     # ・ファイル・ディレクトリの解凍失敗
     # ・$dst_pathが既に存在している
     #
     # @access public
     # @param string $src_path 解凍対象のファイルパス
     # @param string $dst_path 解凍後ファイルパス
     # @return bool ファイルの解凍成否
     # @see Expand-Archive
     # @throws ファイル解凍で例外発生時、$falseを返す。
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
     # [File]ファイルサイズの取得
     #
     # $pathで示すファイルのサイズをバイト単位で返す。
     # $pathが存在しない場合falseを返す。
     #
     # @access public
     # @param string $path サイズを確認するファイルのパス
     # @return long ファイルサイズのバイト数
     # @see Expand-Archive
     # @throws なし
     #>
    [long]FileSize([string]$path){
        if($this.IsFile($path) -eq $false){
            return $false
        }
        return $(Get-ChildItem $path).Length
    }

    <#
     # [File]検索条件に一致する行の削除
     #
     # $fileで示すファイル内の行のうち、正規表現$regに一致する行を削除して元のファイルに上書き保存する。
     # $fileが存在しない場合、またはファイルの書き込み失敗時、falseを返す。
     #
     # @access public
     # @param string $file 検索するファイルのパス
     # @param regex $reg 検索用正規表現
     # @return bool ファイルの書き込み成否
     # @see
     # @throws ファイル書き込みで例外発生時、$falseを返す。
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
     # [File]ファイルの後端行の取得
     #
     # $fileで示すファイルの後端行から$row行を取得する。
     # $fileが存在しない場合、またはファイルの読み込み失敗時、falseを返す。
     #
     # @access public
     # @param string $file 検索するファイルのパス
     # @param int $row 取得する行数
     # @return array $result 取得した行の内容
     # @see Get-Content -tail
     # @throws ファイル内容取得で例外発生時、$falseを返す。
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
     # [File]ファイルの移動
     #
     # $src_pathで示すファイルを$dst_pathに移動する。
     # $src_pathがない場合、または移動失敗時はfalseを返す。
     #
     # @access public
     # @param string $src_path 移動元ファイルパス
     # @param string $src_path 移動先ファイルパス
     # @return bool ファイル移動の成否
     # @see Move-Item
     # @throws ファイル移動で例外発生時、$falseを返す。
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
     # [File]ファイルの移動（強制）
     #
     # $src_pathで示すファイルを$dst_pathに強制的に移動する。
     # $src_pathがない場合、または移動失敗時はfalseを返す。
     #
     # @access public
     # @param string $src_path 移動元ファイルパス
     # @param string $src_path 移動先ファイルパス
     # @return bool ファイル移動の成否
     # @see Move-Item
     # @throws ファイル移動で例外発生時、$falseを返す。
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
     # [File]ファイルのコピー
     #
     # $src_pathで示すファイルを$dst_pathにコピーする。
     # $src_pathがない場合、またはコピー失敗時はfalseを返す。
     #
     # @access public
     # @param string $src_path コピー元ファイルパス
     # @param string $src_path コピー先ファイルパス
     # @return bool ファイルコピーの成否
     # @see Copy-Item
     # @throws ファイルコピーで例外発生時、$falseを返す。
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
     # [File]ファイルのコピー（強制）
     #
     # $src_pathで示すファイルを$dst_pathに強制的にコピーする。
     # $src_pathがない場合、またはコピー失敗時はfalseを返す。
     #
     # @access public
     # @param string $src_path コピー元ファイルパス
     # @param string $src_path コピー先ファイルパス
     # @return bool ファイルコピーの成否
     # @see Copy-Item
     # @throws ファイルコピーで例外発生時、$falseを返す。
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
     # [File]ファイル内容の取得
     #
     # $pathで示すパスから空行、および「#」で始まる行を除く行を取得する。
     # $fileが存在しない場合、falseを返す。
     #
     # @access public
     # @param string $path ファイルのパス
     # @return array $result 取得したファイル内容
     # @see Get-Content
     # @throws なし
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
     # [File]ファイル名の取得
     #
     # $pathで示すパスからファイル名を取得する。
     #
     # @access public
     # @param string $path ファイルのパス
     # @return string 取得したファイル名
     # @see [System.IO.Path]::GetFileName
     # @throws なし
     #>
    [string]Basename([string]$path){
        return ([System.IO.Path]::GetFileName($path))
    }

    <#
     # [File]ディレクトリ名の取得
     #
     # $pathで示すパスからディレクトリ名を取得する。
     #
     # @access public
     # @param string $path ファイルのパス
     # @return string 取得したディレクトリ名
     # @see [System.IO.Path]::GetDirectoryName
     # @throws なし
     #>
    [string]Dirname([string]$path){
        return ([System.IO.Path]::GetDirectoryName($path))
    }

    <#
     # [File]ファイルリストの取得
     #
     # $dirで示すディレクトリ直下のファイルのリストを取得する。
     # リスト取得失敗時は$falseを返す。
     #
     # @access public
     # @param string $dir 検索するディレクトリのパス
     # @return array $list 取得したファイルのフルパス
     # @see Get-ChildItem
     # @throws ファイル一覧取得で例外発生時、$falseを返す。
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
     # [File]CSVファイルをオブジェクトに格納
     #
     # $pathで示すファイルをCSVファイルとしてオブジェクトに変換する。ファイルの1行目はヘッダとして扱う。
     # ファイルが存在しない場合、またはオブジェクト変換失敗時は$falseを返す。
     # $this.encordingで変換時のエンコードを指定する。標準では「default」文字列を設定する。
     #
     # @access public
     # @param string $path CSVファイルのパス
     # @return object $csvObj 変換後のオブジェクト
     # @see Import-Csv
     # @throws CSVからの変換で例外発生時、$falseを返す。
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
     # [File]CSVファイルをオブジェクトに格納
     #
     # $pathで示すファイルをCSVファイルとしてオブジェクトに変換する。$headerでCSVのヘッダを指定する。
     # ファイルが存在しない場合、ヘッダ配列が空の場合、またはオブジェクト変換失敗時は$falseを返す。
     # $this.encordingで変換時のエンコードを指定する。標準では「default」文字列を設定する。
     #
     # @access public
     # @param string $path CSVファイルのパス
     # @param array $header CSVのヘッダ
     # @return object $csvObj 変換後のオブジェクト
     # @see Import-Csv
     # @throws CSVからの変換で例外発生時、$falseを返す。
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
     # [File]オブジェクトをCSV形式に変換しファイルに格納
     #
     # $csvObjをCSV形式に変換し、$pathで示すファイルに保存する。
     # ファイルを保存するディレクトリが存在しない場合、またはオブジェクト変換失敗時は$falseを返す。
     # $this.encordingで変換時のエンコードを指定する。標準では「default」文字列を設定する。
     #
     # @access public
     # @param object $csvObj CSV変換前のオブジェクト
     # @param string $path 保存先のCSVファイルのフルパス
     # @return bool CSVファイルへの変換成否
     # @see Export-Csv
     # @throws オブジェクトからCSVへの変換で例外発生時、$falseを返す。
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