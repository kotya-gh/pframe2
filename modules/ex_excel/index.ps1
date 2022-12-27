. (Join-Path $PSScriptRoot "sheet.ps1")
. (Join-Path $PSScriptRoot "cell.ps1")

<#
 # [MsExcel]Excelファイル処理用クラス
 #
 # Microsoft Excel 2007タイプのファイルを操作する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category Excelファイルの操作
 # @package なし
 #>
class MsExcel{
    [array]$define
    [__ComObject]$Excel
    [MarshalByRefObject]$Book
    [string]$Sheet

    MsExcel(){
        $this.define=(Get-Content (Join-Path $PSScriptRoot "define.json") -Encoding UTF8 -Raw | ConvertFrom-Json)
    }

    <#
     # [MsExcel]Excelを終了する
     #
     # MsExcel Bookを終了後、MsExcelを終了する。
     #
     # @access public
     # @param なし
     # @return なし
     # @see MsExcel.CloseBook MsExcel.CloseExcel
     # @throws なし
     #>
    [void]Quit(){
        $this.CloseBook()
        $this.CloseExcel()
    }

    <#
     # [MsExcel]Excelを起動する
     #
     # Excelを起動する。
     # Excel起動に失敗した場合は、$falseを返す。
     # 
     # @access public
     # @param なし
     # @return bool Excel起動結果の状態
     # @see なし
     # @throws Excel起動時に例外発生時、$falseを返す。
     #>
    [bool]OpenExcel(){
        try{
            [__ComObject] $this.Excel = New-Object -ComObject Excel.Application
            $this.Excel.Visible=$false
            $this.Excel.DisplayAlerts=$false 
        }catch{
            $this.CloseExcel()
            $script:LAST_ERROR_MESSAGE="Fail to open Excel Object."
            return $false
        }
        return $true
    }

    <#
     # [MsExcel]Excelを終了する
     #
     # Excelを終了する。
     # Excel終了失敗時は、$falseを返す。
     #
     # @access public
     # @param なし
     # @return bool Excel終了結果の状態
     # @see なし
     # @throws Excel終了で例外発生時、$falseを返す。
     #>
    [bool]CloseExcel(){
        try{
            $this.Excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.Excel)
            $this.Excel=$null
            [System.GC]::Collect()
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to quit Excel Object."
            return $false
        }
        return $true
    }

    <#
     # [MsExcel]Excelを保存する
     #
     # $file_pathで指定するExcelを保存する。
     # 保存に失敗した場合、$falseを返す。
     #
     # @access public
     # @param string $file_path 保存するExcelのパス
     # @return bool Excel保存結果の状態
     # @see なし
     # @throws Excel保存で例外発生時、$falseを返す。
     #>
    [bool]SaveExcel([string]$file_path){
        [object]$file=New-Object File
        if($file.IsFile($file_path) -eq $true){
            $script:LAST_ERROR_MESSAGE="Exist file."
            return $false            
        }
        try{
            $this.Book.SaveAs($file_path)
        }catch{
            $this.Quit()
            $script:LAST_ERROR_MESSAGE="Fail to save Excel file."
            return $false
        }
        return $true
    }

    <#
     # [MsExcel]上書き保存する
     #
     # Excelの内容を上書き保存する。
     # 上書きに失敗した場合、$falseを返す。
     #
     # @access public
     # @param なし
     # @return bool 上書き保存結果の状態
     # @see なし
     # @throws 上書き保存で例外発生時、$falseを返す。
     #>
    [bool]OverwriteExcel(){
        try{
            $this.Book.Save()
        }catch{
            $this.Quit()
            $script:LAST_ERROR_MESSAGE="Fail to save Excel file."
            return $false
        }
        return $true
    }

    <#
     # [MsExcel]ExcelBookを起動する
     #
     # $file_pathで指定するExcelBookを起動する。
     # ExcelBookの起動に失敗した場合、$falseを返す。
     #
     # @access public
     # @param string $file_path 起動するExcelBookのパス
     # @return bool ExcelBook起動結果の状態
     # @see なし
     # @throws ExcelBook起動で例外発生時、$falseを返す。
     #>
    [bool]OpenBook([string]$file_path){
        try{
            [MarshalByRefObject] $this.Book=$this.Excel.Workbooks.Open($file_path)
        }catch{
            $this.Quit()
            $script:LAST_ERROR_MESSAGE="Fail to open Excel Book."
            return $false
        }
        return $true
    }

    <#
     # [MsExcel]ExcelBookを終了する
     #
     # ExcelBookを終了する。
     # ExcelBookの終了に失敗した場合、$falseを返す。
     #
     # @access public
     # @param なし
     # @return bool ExcelBook終了結果の状態
     # @see なし
     # @throws ExcelBook終了で例外発生時、$falseを返す。
     #>
    [bool]CloseBook(){
        try{
            $this.Book.Close($true)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($this.Book)
            $this.Book=$null
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to close Excel Book."
            return $false
        }
        return $true
    }

    <#
     # [MsExcel]シート名を取得する
     #
     # $sheet_nameで指定するシート名を取得する。
     # シート名取得に失敗した場合、$falseを返す。
     #
     # @access public
     # @param string $sheet_name シート名
     # @return bool シート名取得結果の状態
     # @see なし
     # @throws シート名を取得で例外発生時、$falseを返す。
     #>
    [bool]SetSheet([string]$sheet_name){
        [object]$ExcelSheet=New-Object MsExcelSheet($this.Book)
        if($ExcelSheet.ValidateSheetName($sheet_name) -eq $false){
            return $false
        }
        $this.Sheet=$sheet_name
        return $true
    }

    <#
     # [MsExcel]ExcelSheetの行数を取得する
     #
     # $sheet_numで指定するシートの、使用されているセル範囲の行数を取得する。
     # 
     # @access public
     # @param int $sheet_num 行数を取得するシート番号
     # @return int 使用されているセル範囲の行数
     # @see なし
     # @throws なし
     #>
    [int]GetSheetRows([int]$sheet_num){
        [object]$SheetObj=$this.Excel.Worksheets.Item($sheet_num)
        return $SheetObj.UsedRange.Rows.Count
    }

    <#
     # [MsExcel]ヘッダとフッタのオブジェクトを取得する
     #
     # $sheet_numで指定するヘッダとフッタのオブジェクトを取得する。
     # 
     # @access public
     # @param int $sheet_num シート番号
     # @return object ヘッダとフッタのオブジェクト
     # @see なし
     # @throws なし
     #>
    [object]GetSheetHeaderFooter([int]$sheet_num){
        [object]$SheetObj=$this.Excel.Worksheets.Item($sheet_num)
        return $SheetObj.pageSetup
    }

    <#
     # [MsExcel]シートを削除する
     #
     # $sheet_nameで指定するシートを削除する。
     # シートの削除に失敗した場合、$falseを返す。
     #
     # @access public
     # @param string $sheet_name シート名
     # @return bool シート削除結果の状態
     # @see なし
     # @throws シートを削除で例外発生時、$falseを返す。
     #>
    [bool]DeleteSheet([string]$sheet_name){
        [object]$ExcelSheet=New-Object MsExcelSheet($this.Book)
        # 変更後のシート名のバリデーション
        if($ExcelSheet.ValidateSheetName($sheet_name) -eq $false){
            # エラーメッセージはValidateSheetNameで返す
            return $false
        }
        try{
            $this.Excel.Worksheets.Item($sheet_name).delete()
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to delete sheet."
            return $false            
        }
        return $true
    }

    <#
     # [MsExcel]シートをコピーする
     #
     # $src_sheet_nameで指定するシートを$dst_sheet_nameで指定するシート名に変更してコピーする。
     # シート名の重複があった場合、およびシートコピーに失敗した場合、$falseを返す。
     #
     # @access public
     # @param string $src_sheet_name コピー元シート名
     # @param string $dst_sheet_name コピー先シート名
     # @return bool シートコピー結果の状態
     # @see なし
     # @throws なし
     #>
    [bool]CopySheet([string]$src_sheet_name, [string]$dst_sheet_name){
        [object]$ExcelSheet=New-Object MsExcelSheet($this.Book)
        # 変更後のシート名のバリデーション
        if($ExcelSheet.ValidateSheetName($dst_sheet_name) -eq $false){
            # エラーメッセージはValidateSheetNameで返す
            return $false
        }

        # 変更前・変更後のシート名の重複を確認する。
        if($src_sheet_name -eq $dst_sheet_name){
            $script:LAST_ERROR_MESSAGE="Sheet name is duplicated."
            return $false
        }

        # Book内に存在するシート名と変更後のシート名の重複を確認する。
        [array]$names=$ExcelSheet.GetSheetNames
        [object]$arr=New-Object ExArray
        if($arr.ArraySearch($dst_sheet_name, $names) -eq $true){
            $script:LAST_ERROR_MESSAGE="Sheet name is duplicated."
            return $false
        }
        
        if($ExcelSheet.CopyWorkSheet($src_sheet_name,$dst_sheet_name) -eq $false){
            # エラーメッセージはCopyWorkSheetで返す
            return $false
        }
        return $true
    }
}