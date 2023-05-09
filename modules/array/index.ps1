<#
 # [ExArray]配列操作用クラス
 #
 # 配列の操作、検索等の処理を定義する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category 配列操作処理
 # @package なし
 #>
class ExArray{

    ExArray(){

    }

    <#
     # [ExArray]配列の検索
     #
     # $needleで指定する文字列が$haystackで指定する配列の要素に含まれるかをboolで返す。
     # 検索に失敗した場合、または配列内に存在しない場合は$falseを返す。
     #
     # @access public
     # @param string $needle 検索用文字列
     # @param array $haystack 検索対象配列
     # @return bool 配列内の文字列検索結果
     # @see [Array]::IndexOf
     # @throws 配列内検索で例外発生時、$falseを返す。
     #>
    [bool]ArrayIndexOf([string]$needle, [array]$haystack){
        try{
            [int]$index=[Array]::IndexOf($haystack, $needle)
        } catch {
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        if($index -eq -1){
            return $false
        }
        return $true
    }

    <#
     # [ExArray]配列の検索
     #
     # $needleで指定する文字列が$haystackで指定する配列の要素に含まれるかをboolで返す。
     # 検索文字列が配列内に存在しない場合は$falseを返す。
     #
     # @access public
     # @param string $needle 検索用文字列
     # @param array $haystack 検索対象配列
     # @return bool 配列内の文字列検索結果
     # @see なし
     # @throws なし
     #>
    [bool]ArraySearch([string]$needle, [array]$haystack){
        foreach($value in $haystack){
            if($needle -eq $value){
                return $true
            }
        }
        return $false
    }

    <#
     # [ExArray]ハッシュの検索
     #
     # $hash.$key内に$valueで示す値が存在するかをboolで返す。
     # 検索文字列が配列内に存在しない場合は$falseを返す。
     #
     # @access public
     # @param string $key 検索用キー
     # @param string $value 検索用文字列
     # @param object $hash 検索対象ハッシュ
     # @return bool ハッシュ内の文字列検索結果
     # @see なし
     # @throws なし
     #>
    [bool]HashValueSearch([string]$key, [string]$value, [object]$hash){
        foreach($record in $hash){
            if($record.$key -eq $value){
                return $true
            }
        }
        return $false
    }

    <#
     # [ExArray]PSCustomObjectのプロパティ検索
     #
     # $objで示すPSCustomObject内に$propで示すプロパティが存在するかをboolで返す。
     # 検索文字列がPSCustomObject内に存在しない場合は$falseを返す。
     #
     # @access public
     # @param object $obj PSCustomObject
     # @param string $prop 検索用プロパティ
     # @return bool PSCustomObject内のプロパティ検索結果
     # @see なし
     # @throws なし
     #>
    [bool]PsoSearchProp([object]$obj, [string]$prop){
        $list=($obj | Get-Member -MemberType NoteProperty | Select -ExpandProperty Name)
        $Arr=New-object ExArray
        if($Arr.ArraySearch($prop, $list) -eq $true){
            return $true
        }
        return $false
    }

    <#
     # [ExArray]配列の重複排除
     #
     # $haystackで示す配列内に重複した要素がある場合、重複を排除した配列を返す。
     # 重複排除に失敗した場合は$falseを返す。
     #
     # @access public
     # @param array $haystack 重複を排除する配列
     # @return array 重複を排除した配列
     # @see Get-Unique
     # @throws 重複排除で例外発生時、$falseを返す。
     #>
    [array]ArrayUnique([array]$haystack){
        try{
            $haystack = $haystack | Sort-Object | Get-Unique
        } catch {
            $script:LAST_ERROR_MESSAGE=$_.Exception
            return $false 
        }
        return $haystack
    }

    <#
     # [ExArray]配列の重複判定
     #
     # $haystackで示す配列内に重複した要素があるかどうかをboolで返す。
     # 重複している場合は$falseを返す。
     #
     # @access public
     # @param array $haystack 重複有無を判定する配列
     # @return bool 重複判定結果
     # @see ExArray.ArrayUnique
     # @throws なし
     #>
    [bool]IsDuplicateArrayValue([array]$haystack){
        if( $this.ArrayUnique($haystack).Length -eq $haystack.Length ){
            return $true
        }
        return $false
    }

    <#
     # [ExArray]配列の最後の要素の削除
     #
     # $haystackで示す配列から最後の値を削除した配列を返す。配列が空の場合は空の配列を作成して返す。
     # 重複している場合は$falseを返す。
     #
     # @access public
     # @param array $haystack 末尾要素を削除する配列
     # @return array $haystack 末尾要素を削除した配列
     # @see なし
     # @throws なし
     #>
    [array]Pop([array]$haystack){
        if($haystack.Length -eq 0){
            return @()
        }
        $haystack=$haystack[0..($haystack.Length-2)]
        return $haystack
    }

    <#
     # [ExArray]ハッシュのソート（昇順）
     #
     # $hashで示すハッシュの$keyで示すキーで昇順にソートしたハッシュを返す。
     #
     # @access public
     # @param object $hash ソートするハッシュ
     # @param string $key ソートするキー
     # @return object ソートしたハッシュ
     # @see なし
     # @throws なし
     #>
    [object]HashSort([object]$hash, [string]$key){
        if($hash.Length -ge 2){
            return ($hash.GetEnumerator() | Sort-Object -Property $key)
        }
        return $hash
    }

    <#
     # [ExArray]ハッシュのソート（降順）
     #
     # $hashで示すハッシュの$keyで示すキーで降順にソートしたハッシュを返す。
     #
     # @access public
     # @param object $hash ソートするハッシュ
     # @param string $key ソートするキー
     # @return object ソートしたハッシュ
     # @see なし
     # @throws なし
     #>
     [object]HashSortDesc([object]$hash, [string]$key){
        if($hash.Length -ge 2){
            return ($hash.GetEnumerator() | Sort-Object -Property $key -Descending)
        }
        return $hash
    }

    <#
     # [ExArray]配列のソート（昇順）
     #
     # $arrで示す配列を昇順にソートした配列を返す。
     #
     # @access public
     # @param array $arr 昇順にソートする配列
     # @return array ソートした配列
     # @see なし
     # @throws なし
     #>
    [array]ArraySort([array]$arr){
        return ($arr | Sort-Object)
    }

    <#
     # [ExArray]配列のソート（降順）
     #
     # $arrで示す配列を降順にソートした配列を返す。
     #
     # @access public
     # @param array $arr 降順にソートする配列
     # @return array ソートした配列
     # @see なし
     # @throws なし
     #>
    [array]ArraySortDesc([array]$arr){
        return ($arr | Sort-Object -Descending)
    }

    <#
     # [ExArray]ハッシュテーブル変換
     #
     # $cumtomObjectで示すPSCustomObjectをハッシュテーブルに変換する
     #
     # @access public
     # @param PSCustomObject $cumtomObject 変換したいPSCustomObject
     # @return HashTable 変換後のハッシュテーブル
     # @see なし
     # @throws なし
     #>
    [object]ConvertToHashTable([PSCustomObject]$cumtomObject){
        $hashTable=@{}
        foreach($key in $cumtomObject.PSObject.Properties.Name){
            if($cumtomObject.$key -is [System.Management.Automation.PSCustomObject]){
                $hashTable[$key]=$this.ConvertToHashTable($cumtomObject.$key)
            }else{
                $hashTable[$key]=$cumtomObject.$key
            }
        }
        return $hashTable    
    }
}

