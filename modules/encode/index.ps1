<#
 # [Encode]文字のエンコード用クラス
 #
 # 文字のエンコード、デコードや変換のための処理を定義する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category 文字列エンコード
 # @package なし
 #>
class Encode{
    Encode(){

    }

    <#
     # [Encode]文字列のURLエンコード
     #
     # $codeで指定する文字コードで$string文字列を読み込み、URLエンコードの結果を返す。
     # URLエンコードに失敗した場合$falseを返す。
     #
     # @access public
     # @param string $string エンコード前の文字列
     # @param string $code 文字コード名
     # @return string $result URLエンコード後の文字列
     # @see [System.Web.HttpUtility]::UrlEncode
     # @throws URLエンコードで例外発生時、$falseを返す。
     #>
    [string]Urlencode([string]$string, [string]$code){
        Add-Type -AssemblyName System.Web
        try{
            [string]$result=([System.Web.HttpUtility]::UrlEncode($string, [Text.Encoding]::GetEncoding($code)))
        }catch{
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false
        }
        return $result
    }
 
    <#
     # [Encode]文字列のURLエンコード
     #
     # $codeで指定する文字コードで$array配列から1要素ずつ文字列を読み込み、URLエンコードの結果を配列で返す。
     # URLエンコードに失敗した場合$falseを返す。
     #
     # @access public
     # @param array $array エンコード前の文字列配列
     # @param string $code 文字コード名
     # @return array $result URLエンコード後の文字列配列
     # @see Encode.Urlencode
     # @throws URLエンコードで例外発生時、$falseを返す。
     #>
    [array]Urlencode([array]$array, [string]$code){
        [array]$result=@()
        foreach($value in $array){
            $result+=$this.Urlencode($value, $code)
        }
        return $result
    }

    <#
     # [Encode]文字列のURLデコード
     #
     # $codeで指定する文字コードで$string文字列を読み込み、URLデコードの結果を返す。
     # URLデコードに失敗した場合$falseを返す。
     #
     # @access public
     # @param string $string デコード前の文字列
     # @param string $code 文字コード名
     # @return string $result URLデコード後の文字列
     # @see [System.Web.HttpUtility]::UrlDecode
     # @throws URLデコードで例外発生時、$falseを返す。
     #>
    [string]Urldecode([string]$string, [string]$code){
        Add-Type -AssemblyName System.Web
        return [System.Web.HttpUtility]::UrlDecode($string, [Text.Encoding]::GetEncoding($code))
    }

    <#
     # [Encode]文字列のURLデコード
     #
     # $codeで指定する文字コードで$array配列から1要素ずつ文字列を読み込み、URLデコードの結果を配列で返す。
     # URLデコードに失敗した場合$falseを返す。
     #
     # @access public
     # @param array $array デコード前の文字列配列
     # @param string $code 文字コード名
     # @return array $result URLデコード後の文字列配列
     # @see Encode.Urldecode
     # @throws URLデコードで例外発生時、$falseを返す。
     #>
    [array]Urldecode([array]$array, [string]$code){
        [array]$result=@()
        foreach($value in $array){
            $result+=$this.Urldecode($value, $code)
        }
        return $result
    }

    <#
     # [Encode]特定文字列の変換
     #
     # $encodedStringで指定する文字列内から%c2%a1～%c2%bf、
     # および%c3%80～%c3%bfを%3fに変換し、変換後の文字列を返す。
     #
     # @access public
     # @param string $encodedString 変換前の文字列
     # @return string $encodedString 変換後の文字列配列
     # @see なし
     # @throws なし
     #>
    [string]ConvertSpecialChar([string]$encodedString){
        # sp_char -> "%3f"
        # sp_char(1) ... %c2%a1 - %c2%bf
        # sp_char(2) ... %c3%80 - %c3%bf
        for([int]$i=0xa1; $i -le 0xbf; $i++){
            [string]$hex=("%c2%"+$i.ToString("x2"))
            $encodedString=$encodedString -creplace $hex, "%3f"
        }
        for([int]$i=0x80; $i -le 0xbf; $i++){
            [string]$hex=("%c3%"+$i.ToString("x2"))
            $encodedString=$encodedString -creplace $hex, "%3f"
        }
        return $encodedString
    }

    <#
     # [Encode]特定文字列の変換
     #
     # $encodedArrayで指定する文字列配列から1要素ずつ文字列を読み込み、
     # %c2%a1～%c2%bf、および%c3%80～%c3%bfを%3fに変換し、変換後の文字列を返す。
     #
     # @access public
     # @param array $encodedArray 変換前の文字列配列
     # @return array $result 変換後の文字列配列配列
     # @see Encode.ConvertSpecialChar
     # @throws なし
     #>
    [array]ConvertSpecialChar([array]$encodedArray){
        [array]$result=@()
        foreach($value in $encodedArray){
            $result+=$this.ConvertSpecialChar($value)
        }
        return $result
    }

    <#
     # [Encode]ファイルのJsonデコード
     #
     # $file_pathで指定するファイルを読み込み、Json変換後の結果を配列で返す。
     # Jsonデコードに失敗した場合$falseを返す。
     #
     # @access public
     # @param string $file_path デコードするファイルパス
     # @return array $result Jsonデコード後の配列
     # @see convertFrom-Json
     # @throws Jsonデコードで例外発生時、$falseを返す。
     #>
    [array]JsonFileDecode([string]$file_path){
        try{
            [array]$decode_array=(Get-Content $file_path -Raw | convertFrom-Json)
        } catch {
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }        
        return $decode_array
    }
}