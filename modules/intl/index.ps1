<#
 # [Intl]言語設定用クラス
 #
 # 多言語対応用の処理を定義する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category 多言語用処理
 # @package なし
 #>
class Intl{
    [string]$path
    [object]$define

    Intl([string]$path){
        $this.define=(Get-Content (Join-Path $PSScriptRoot "define.json") -Raw | ConvertFrom-Json)

        # init lang
        $script:LOCALE=$this.defaultLocale($script:LOCALE)
        $this.path=(Join-Path $path ($script:LOCALE+".json"))
    }

    <#
     # [Intl]ロケール文字列の取得
     #
     # $localeで指定するロケール名が定義されている場合、入力値を返す。
     # 定義されていないロケールの場合"ja_JP"を返す。
     #
     # @access public
     # @param string $locale ロケール文字列
     # @return string ロケール文字列
     # @see
     # @throws なし
     #>
    [string]defaultLocale([string]$locale){
        [object]$arr=New-Object ExArray
        if($arr.ArraySearch($locale, $this.define.lang) -eq $true){
            return $locale
        }
        return "ja_JP"
    }

    <#
     # [Intl]IDから文字列を取得
     #
     # $idで指定するIDに対応する文字列を返す。
     # 存在しない$idを指定した場合、$falseを返す。
     #
     # @access public
     # @param string $id 検索用ID
     # @return string IDに対応する文字列
     # @see Intl.FormattedMessage
     # @throws なし
     #>
    [string]FormattedMessage([string]$id){
        return $this.FormattedMessage($id, @{})
    }

    <#
     # [Intl]IDから文字列を取得
     #
     # $idで指定するIDに対応する文字列を返す。
     # 文字列内に$valueのハッシュキーを「__(ハッシュキー)」の形式で記載している場合、対応するハッシュ値と入れ替える。
     # 存在しない$idを指定した場合、$falseを返す。
     #
     # @access public
     # @param string $id 検索用ID
     # @param object $value メッセージ内文字列変換用ハッシュ
     # @return string $returnString IDに対応する文字列
     # @see Get-Content
     # @throws メッセージ定義ファイル取得で例外発生時、$falseを返す。
     #>
    [string]FormattedMessage([string]$id, [object]$value){
        [object]$file=New-Object File
        if($file.IsFile($this.path) -eq $false){
            $script:LAST_ERROR_MESSAGE="No lang file.($this.path)"            
            return $false
        }
        # import lang file.
        try{
            [object]$lang=(Get-Content $this.path -Raw | ConvertFrom-Json)
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        # search lang key
        if($null -eq $lang.$id){
            $script:LAST_ERROR_MESSAGE="No id is found.($id)"            
            return $false
        }
        [string]$returnString=$lang.$id
        foreach($key in $value.Keys){
            $returnString=$returnString.Replace("__($key)", $value[$key])
        }
        return $returnString
    }
}