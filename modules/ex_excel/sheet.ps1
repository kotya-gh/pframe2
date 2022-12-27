<#
 # [MsExcelSheet]Excelファイル処理用クラス
 #
 # Microsoft Excel 2007タイプのファイルを操作する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category Excelファイルの操作
 # @package なし
 #>
class MsExcelSheet{
    [array]$define
    [MarshalByRefObject]$Book


    MsExcelSheet([MarshalByRefObject]$Book){
        $this.Book=$Book
        [object]$enc=New-Object Encode
        $this.define=$enc.JsonFileDecode((Join-Path $PSScriptRoot "define.json"))
    }

    <#
     # [MsExcelSheet]シートの数を取得する
     #
     # Excelbookのシート数を取得する。
     #
     # @access public
     # @param なし
     # @return int シートの数
     # @see なし
     # @throws なし
     #>
    [int]GetSheetCount(){
        return $this.Book.Sheets.Count
    }

    <#
     # [MsExcelSheet]シート名を取得する
     #
     # $thで指定するシート番号のシート名を取得する。
     # 
     # @access public
     # @param int $th シート番号
     # @return string シート名
     # @see なし
     # @throws なし
     #>
    [string]GetSheetName([int]$th){
        return $this.Book.Sheets($th).Name        
    }

    <#
     # [MsExcelSheet]シート名をすべて取得する
     #
     # シート名をすべて取得する。
     # 
     # @access public
     # @param なし
     # @return array シート名
     # @see なし
     # @throws なし
     #>
    [array]GetSheetNames(){
        [array]$names=@()
        for([int]$i=1; $i -le $this.GetSheetCount(); $i++){
            $names+=$this.Book.Sheets($i).Name
        }
        return $names    
    }

    <#
     # [MsExcelSheet]シート名を記入する
     #
     # $thで指定するシートに$sheet_nameで指定する文字で記入する。
     # シート名の重複、またはシート名の記入に失敗した場合、$falseを返す。
     # 
     # @access public
     # @param int $th 記入するシート番号
     # @param string $sheet_name 記入する文字列
     # @return bool シート名を記入した結果の状態
     # @see なし
     # @throws なし
     #>
    [bool]SetSheetName([int]$th, [string]$sheet_name){
        if($this.ValidateSheetName($sheet_name) -eq $false){
            return $false
        }
        $this.Book.Sheets($th).Name=$sheet_name
        return $true
    }

    <#
     # [MsExcelSheet]有効なシート名であるか検証する
     #
     # $sheet_nameで指定するシート名が有効なシート名であるかを検証する。
     # 有効なシート名であるかの検証に失敗した場合、$falseを返す。
     # 
     # @access public
     # @param string $sheet_name シート名
     # @return bool 有効なシート名であるか検証する結果の状態
     # @see なし
     # @throws なし
     #>
    [bool]ValidateSheetName([string]$sheet_name){
        [object]$str=New-Object Str
        if(
            ($str.Strpos($sheet_name, $this.define.restricted_char) -eq $true) -or
            ($sheet_name.Length -gt $this.define.max_len_of_sheet_name) -or
            (($sheet_name -as "bool") -eq $false)
        ){
            $script:LAST_ERROR_MESSAGE="Invalid Excel sheet name."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelSheet]シートのコピー処理
     #
     # $src_sheet_nameで指定するコピー元シート名を一時的に変更し、
     # $dst_sheet_nameで指定するコピー先シートを作成後、コピー元を元に戻す。
     # シートのコピー処理に失敗した場合、$falseを返す。
     # 
     # @access public
     # @param string $src_sheet_name コピー元シート名
     # @param string $dst_sheet_name コピー先シート名
     # @return bool シート名をコピーした結果の状態
     # @see なし
     # @throws シートのコピー処理で例外発生時、$falseを返す。
     #>
    [bool]CopyWorkSheet([string]$src_sheet_name, [string]$dst_sheet_name){
        # シートのコピー処理：コピー元シート名を一時的に変更し、コピー先シートを作成後、コピー元を元に戻す。
        try{
            [string]$tmp_sheet_name=[string](Get-Date -Format $this.define.dateStringFormat)
            $this.Book.WorkSheets.item($src_sheet_name).name=$tmp_sheet_name
            $this.Book.worksheets.item($tmp_sheet_name).copy($this.Book.worksheets.item($tmp_sheet_name)) 
            $this.Book.worksheets.item($tmp_sheet_name + " (2)").name = $dst_sheet_name
            $this.Book.WorkSheets.item($tmp_sheet_name).name=$src_sheet_name
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to Copy Worksheets."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelSheet]印刷範囲を解除する
     #
     # $excebookで指定する印刷範囲を解除する。
     # 
     # @access public
     # @param object $excebook 印刷範囲を解除するExcelbookのオブジェクト
     # @return bool 印刷範囲を解除する結果の状態
     # @see なし
     # @throws 印刷範囲を解除で例外発生時、$falseを返す。
     #>
    [bool]DeletePrintArea([object]$excebook){
        try {
            $this.Book.ActiveSheet.PageSetup.PrintArea = ""
        }
        catch {
            $script:LAST_ERROR_MESSAGE="Fail to delete print area."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelSheet]印刷範囲を設定する
     #
     # $setrangeで指定する範囲を印刷範囲として設定する。
     # 印刷範囲の設定に失敗した場合、$falseを返す。
     #
     # @access public
     # @param string $setrange 印刷範囲
     # @return bool 印刷範囲を設定する結果の状態
     # @see なし
     # @throws 印刷範囲を設定で例外発生時、$falseを返す。
     #>
    [bool]SetPrintArea([string]$setrange){
        try {
            $this.Book.ActiveSheet.PageSetup.PrintArea = $setrange
        }
        catch {
            $script:LAST_ERROR_MESSAGE="Fail to set print area."
            return $false
        }
        return $true
    }
}