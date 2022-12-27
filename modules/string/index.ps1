<#
 # [Str]文字列操作に関するクラス
 #
 # 文字列の検索、編集、置換を行うためのクラス。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category 文字列操作
 # @package なし
 #>
class Str{
    [object]$LOCALE
    Str(){
        $this.LOCALE=New-Object Intl((Join-Path $PSScriptRoot "locale"))
    }

    <#
     # [Str]改行の変換処理
     #
     # $stringで指定する文字列の改行を、スペースに変換した結果を返す。
     #
     # @access public
     # @param string $string 変換前文字列
     # @return string $string 変換後文字列
     # @see Replace
     # @throws なし
     #>
    [string]TrimEol([string]$string){
        [string]$string=$string.Replace("`r`n", " ").Replace("`r", " ").Replace("`n", " ")
        return $string
    }

    <#
     # [Str]文字列の削除処理
     #
     # $stringで指定する文字列の文字列の左端から$needleを検索し、一致した場合削除する。
     # 文字列$stringが$needleより短い場合falseを返す。
     #
     # @access public
     # @param string $string 変換前文字列
     # @param string $needle 検索文字列
     # @return string $string 変換後文字列
     # @see Remove
     # @throws なし
     #>
    [string]LTrim([string]$string, [string]$needle){
        if($string.Length -lt $needle.Length){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("The_length_of_needle_is_too_long")
            return $false
        }
        if($string.Remove($needle.Length, $string.Length-$needle.Length) -eq $needle){
            return $string.Remove(0, $needle.Length)
        }
        return $string
    }

    <#
     # [Str]文字列の削除処理
     #
     # $stringで指定する文字列の文字列の右端から$needleを検索し、一致した場合削除する。
     # 文字列$stringが$needleより短い場合falseを返す。
     #
     # @access public
     # @param string $string 変換前文字列
     # @param string $needle 検索文字列
     # @return string $string 変換後文字列
     # @see Remove
     # @throws なし
     #>
    [string]RTrim([string]$string, [string]$needle){
        if($string.Length -lt $needle.Length){
            $script:LAST_ERROR_MESSAGE=$this.LOCALE.FormattedMessage("The_length_of_needle_is_too_long")
            return $false
        }
        if($string.Remove(0, $string.Length-$needle.Length) -eq $needle){
            return $string.Remove($string.Length-$needle.Length, $needle.Length)
        }
        return $string
    }

    <#
     # [Str]文字列の検索、置換処理
     #
     # $subjectで指定する文字列を文字列$searchで検索し、一致した場合文字列$replaceに置き換える。
     #
     # @access public
     # @param string $search 検索文字列
     # @param string $replace 置換文字列
     # @param string $subject 変換前文字列
     # @return string 変換後文字列
     # @see creplace
     # @throws なし
     #>
    [string]Replace([string]$search, [string]$replace ,[string]$subject){
        return ($subject -creplace $search, $replace)
    }

    <#
     # [Str]文字列の検索処理
     #
     # $haystackで指定する文字列を文字列$needleで検索し、一致した場合$true、一致しない場合$falseを返す。
     #
     # @access public
     # @param string $haystack 検索対象文字列
     # @param string $needle 検索文字列
     # @return bool 検索結果
     # @see Contains
     # @throws なし
     #>
    [bool]Strpos([string]$haystack, [string]$needle){
        return $haystack.Contains($needle)
    }

    <#
     # [Str]文字列の検索処理
     #
     # $haystackで指定する文字列を配列$needle内の文字列で検索し、１つでも一致した場合$true、全て一致しない場合$falseを返す。
     #
     # @access public
     # @param string $haystack 検索対象文字列
     # @param array $needle 検索文字列
     # @return bool 検索結果
     # @see Contains
     # @throws なし
     #>
    [bool]Strpos([string]$haystack, [array]$needle){
        foreach($value in $needle){
            if($haystack.Contains($value) -eq $true){
                return $true
            }
        }
        return $false
    }

    <#
     # [Str]文字列の分割処理
     #
     # $stringで指定する文字列を文字列$delimiterで分割し、各要素を文字列として格納した配列を返す。
     #
     # @access public
     # @param string $delimiter 分割文字列
     # @param string $string 検索対象文字列
     # @return array 分割した要素を格納した配列
     # @see split
     # @throws なし
     #>
    [array]Explode([string]$delimiter, [string]$string){
        return ($string -split $delimiter)
    }

    <#
     # [Str]配列の連結処理
     #
     # $piecesで指定する配列要素を文字列$glueにより連結する。
     # すべての配列要素の順序を変えずに、各要素間に$glue文字列をはさんで1つの文字列にして返す。
     #
     # @access public
     # @param string $glue 要素間を連結する文字列
     # @param array $pieces 連結したい文字列の配列
     # @return string 配列要素を連結した文字列
     # @see $this.Ltrim
     # @throws なし
     #>
    [string]Implode([string]$glue, [array]$pieces){
        if($pieces.Length -eq 0){
            return ""
        }
        [string]$ret=""
        foreach($val in $pieces){
            $ret+=($glue+[string]$val)
        }
        return $this.LTrim($ret, $glue)
    }

    <#
     # [Str]文字列の分割処理
     #
     # $stringで指定する文字列をホワイトスペースで分割し、各要素を文字列として格納した配列を返す。
     # 前後のホワイトスペースは削除する。
     # ホワイトスペースは[半角スペース、\t、\n、\r、\f]と定義する。
     #
     # @access public
     # @param string $string 分割対象文字列
     # @return array 分割した要素を格納した配列
     # @see Trim()
     # @throws なし
     #>
    [array]ExplodeWs([string]$string){
        $delimiter="__explodeWs_Delimiter__"
        return $string.Trim() -replace "(\s)+", $delimiter -split $delimiter
    }  
}