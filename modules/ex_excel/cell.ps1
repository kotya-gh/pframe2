<#
 # [MsExcelCell]Excelファイル処理用クラス
 #
 # Microsoft Excel 2007タイプのファイルを操作する。
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category Excelファイルの操作
 # @package なし
 #>
class MsExcelCell{
    [array]$define
    [MarshalByRefObject]$Book
    [object]$Sheet

    MsExcelCell([MarshalByRefObject]$Book, [string]$Sheet){
        $this.Book=$Book
        $this.Sheet=$this.Book.Worksheets.Item($Sheet)
        [object]$enc=New-Object Encode
        $this.define=$enc.JsonFileDecode((Join-Path $PSScriptRoot "define.json"))
    }

    <#
     # [MsExcelCell]行を挿入する
     #
     # $rowで指定する行数を挿入する。
     # 行の挿入で失敗した場合、$falseを返す。
     #
     # @access public
     # @param int $row 挿入する行の数
     # @return bool 行の挿入結果の状態
     # @see なし
     # @throws 行の挿入で例外発生時、$falseを返す。
     #>
    [bool]InsertRow([int]$row){
        try{
            $this.Sheet.rows.item($row).insert()
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to insert row."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]行をコピーする
     #
     # $src_rowで指定する行をコピーし、$dst_rowで指定する行に貼り付ける。
     # 行のコピーで失敗した場合、$falseを返す。
     #
     # @access public
     # @param int $src_row コピー元の行
     # @param int $dst_row コピー先の行
     # @return bool 行のコピー結果の状態
     # @see なし
     # @throws 行コピーで例外発生時、$falseを返す。
     #>
    [bool]CopyRow([int]$src_row, [int]$dst_row){
        try{
            $this.Sheet.rows.item($src_row).copy($this.Sheet.rows.item($dst_row))
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to copy row."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]セルをコピーする
     #
     # $srcで指定するセルの内容を取得し、$dstで指定するセルに貼り付ける。
     # $srcで指定するメンバを次に示す。
     # row：コピー元のセルの行番号
     # col：コピー元のセルの列番号
     # $dstで指定するメンバを次に示す。
     # row：コピー先のセルの行番号
     # col：コピー先のセルの列番号
     # セルをコピー失敗時、$falseを返す。
     #
     # @access public
     # @param array $src コピー元セル
     # @param array $dst コピー先セル
     # @return bool セルのコピー結果の状態
     # @see なし
     # @throws なし
     #>
    [bool]CopyCell([array]$src, [array]$dst){
        if($this.CopyCell($src, $dst, $dst) -eq $false){
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]範囲選択したセルにコピーする
     #
     # $srcで指定するセルの内容を、$dst_originから$dst_endで指定する範囲に貼り付ける。
     # $srcで指定するメンバを次に示す。
     # int row：コピー元のセルの行番号
     # int col：コピー元のセルの列番号
     # $dst_originで指定するメンバを次に示す。
     # int row：コピー先の範囲指定の始まりの行番号
     # int col：コピー先の範囲指定の始まりの列番号
     # $dst_endで指定するメンバを次に示す。
     # int row：コピー先の範囲指定の終わりの行番号
     # int col：コピー先の範囲指定の終わりの列番号
     # セルの範囲選択でコピー失敗時、$falseを返す。
     #
     # @access public
     # @param array $src コピー元セル
     # @param array $dst_origin コピー先セルの始まり
     # @param array $dst_end コピー先セルの終わり
     # @return bool セル範囲コピー結果の状態
     # @see なし
     # @throws セルを範囲選択してコピーで例外発生時、$falseを返す。
     #>
    [bool]CopyCell([array]$src, [array]$dst_origin, [array]$dst_end){
        try{
            $this.Sheet.Cells([int]$src.row, [int]$src.col).copy( `
                $this.Sheet.Range( `
                    $this.Sheet.Cells([int]$dst_origin.row, [int]$dst_origin.col), `
                    $this.Sheet.Cells([int]$dst_end.row, [int]$dst_end.col)
                )
            )
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to copy cells."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]セルの結合
     #
     # $originで指定するセルと、$endで指定するセルを結合する。
     # $originで指定するメンバを次に示す。
     # int row：結合するセルの始めの行番号
     # int col：結合するセルの始めの列番号
     # $endで指定するメンバを次に示す。
     # int row：結合するセルの終わりの行番号
     # int col：結合するセルの終わりの列番号
     # セルの結合に失敗した場合、$falseを返す。
     #
     # @access public
     # @param array $origin 結合するセルの始まり
     # @param array $end 結合するセルの終わり
     # @return bool セルの結合結果の状態
     # @see なし
     # @throws セル結合で例外発生時、$falseを返す。
     #>
    [bool]MergeCells([array]$origin, [array]$end){
        try{
            $this.Sheet.Range( `
                $this.Sheet.Cells([int]$origin.row, [int]$origin.col), 
                $this.Sheet.Cells([int]$end.row, [int]$end.col) `
            ).MergeCells = $true
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to merge cells."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]セルにコメントを追加する
     #
     # $dataで指定するセルにコメントを追加する。$dataで指定するメンバを次に示す。
     # int    row：コメントを追加するセルの行番号
     # int    col：コメントを追加するセルの列番号
     # string comment：追加するコメントの内容
     # セルにコメントを追加で失敗した場合、$falseを返す。
     #
     # @access public
     # @param array $data コメントを追加するセル
     # @return bool コメント追加結果の状態
     # @see なし
     # @throws セルにコメントを追加で例外発生時、$falseを返す。
     #>
    [bool]AddComment([array]$data){
        try{
            $this.Sheet.Cells.Item([int]$data.row, [int]$data.col).AddComment([string]$data.comment)
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to add comment."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]文字列を複数のセルに記入する
     #
     # $data_objectで指定したそれぞれの内容を複数のセルに記入する。$data_objectで指定するメンバを次に示す。
     #  int row：行番号
     #  int col：列番号
     #      value：記入する値
     # セルに記入で失敗した場合、$falseを返す。
     # @access public
     # @param array $data_object 記入する内容
     # @return bool セル記入結果の状態
     # @see なし
     # @throws なし
     #>
    [bool]WriteCells([array]$data_object){
        foreach($data in $data_object){
            if($this.WriteCell($data) -eq $false){
                return $false
            }
        }
        return $true
    }

    <#
     # [MsExcelCell]セルに記入する
     #
     # $dataで指定する内容をセルに記入する。$dataで指定するメンバを次に示す。
     # int row：行番号
     # int col：列番号
     #     value：記入する値
     # セルに記入で失敗した場合、$falseを返す。
     # @access public
     # @param array $data 記入する内容
     # @return bool セル記入結果の状態
     # @see なし
     # @throws セルに記入で例外発生時、$falseを返す。
     #>
    [bool]WriteCell([array]$data){
        try{
            $this.Sheet.Cells.Item([int]$data.row, [int]$data.col)=$data.value
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to write value in Excel sheet."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]セルの内容を取得する
     #
     # $gridで指定するセルの内容を取得する。$gridで指定するメンバを次に示す。
     # int row：内容を取得するセルの行番号
     # int col：内容を取得するセルの列番号
     #
     # @access public
     # @param array $grid 内容を取得するセル
     # @return string セルの内容
     # @see なし
     # @throws なし
     #>
    [string]TextCell([array]$grid){
        return $this.Sheet.Cells([int]$grid.row, [int]$grid.col).Text
    }

    [object]RangeCell([array]$startgrid, [array]$endgrid){
        return $this.sheet.Range(
            $this.sheet.Cells([int]$startgrid.row, [int]$startgrid.col), 
            $this.sheet.Cells([int]$endgrid.row, [int]$endgrid.col))
    }

    [array]RangeCellToArray([array]$startgrid, [array]$endgrid){
        [object]$data=$this.RangeCell($startgrid, $endgrid)
        [array]$output=@()

        [int]$i=1
        [array]$tmp=@()
        $data.Value(10) | ForEach-Object {
            $tmp+=$_
            if($i%($endgrid.col-$startgrid.col+1) -eq 0){
                $output += ,$tmp
                $tmp=@()
            }
            $i++
        }
        return $output
    }

    <#
     # [MsExcelCell]セルに代入されている計算式を取得する
     #
     # $gridで指定するセルの計算式を取得する。$gridで指定するメンバを次に示す。
     # int row：セルの行番号
     # int col：セルの列番号
     # セルの計算式を取得に失敗した場合、$falseを返す。
     #
     # @access public
     # @param array $grid 計算式が代入されているセル
     # @return string 計算式
     # @see なし
     # @throws なし
     #>
    [string]FormulaCell([array]$grid){
        return $this.Sheet.Cells([int]$grid.row, [int]$grid.col).Formula
    }

    <#
     # [MsExcelCell]指定するセル範囲の文字を太字にする
     #
     # $originから$endで指定するセル範囲の文字を太字にする。
     # $originで指定するメンバを次に示す。
     # int row：範囲指定の始まりの行番号
     # int col：範囲指定の始まりの列番号
     # $endで指定するメンバを次に示す。
     # int row：範囲指定の終わりの行番号
     # int col：範囲指定の終わりの列番号
     # セル範囲の文字を太字へ変換に失敗した場合、$falseを返す。
     #
     # @access public
     # @param array $origin セルの始まり
     # @param array $end セルの終わり
     # @return bool 文字を太字にする結果の状態
     # @see なし
     # @throws 文字を太字に変換で例外発生時、$falseを返す。
     #>
    [bool]BoldCell([array]$origin, [array]$end){
        try{
            $this.Sheet.Range( `
                $this.Sheet.Cells([int]$origin.row, [int]$origin.col), 
                $this.Sheet.Cells([int]$end.row, [int]$end.col) `
            ).Font.Bold = $true
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to change cells bold."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]文字を太字にする
     #
     # $gridで指定するセルの文字を太字にする。$gridで指定するメンバを次に示す。
     # row：文字を太字にするセルの行番号
     # col：文字を太字にするセルの列番号
     # 文字を太字へ変換に失敗した場合、$falseを返す。
     #
     # @access public
     # @param array $grid 太字にするセル
     # @return bool 文字を太字にする結果の状態
     # @see なし
     # @throws なし
     #>
    [bool]BoldCell([array]$grid){
        if($this.BoldCell($grid, $grid) -eq $false){
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]セルに色を塗りつぶす
     #
     # $originと$endで指定するセル範囲を、$colorで指定する色に塗りつぶす。
     # $originで指定するメンバを次に示す。
     # int row：範囲指定の始まりの行番号
     # int col：範囲指定の始まりの列番号
     # $endで指定するメンバを次に示す。
     # int row：範囲指定の終わりの行番号
     # int col：範囲指定の終わりの列番号
     # セルに色を塗りつぶしに失敗した場合、$falseを返す。
     #
     # @access public
     # @param int $color ColorIndexの数値
     # @param array $origin セルの始まり
     # @param array $end セルの終わり
     # @return bool セルに色を塗りつぶした結果の状態
     # @see なし
     # @throws セルに色を塗りつぶしで例外発生時、$falseを返す。
     #>
    [bool]ColorCell([int]$color, [array]$origin, [array]$end){
        try{
            $this.Sheet.Range( `
                $this.Sheet.Cells([int]$origin.row, [int]$origin.col), 
                $this.Sheet.Cells([int]$end.row, [int]$end.col) `
            ).interior.ColorIndex = $color
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to change cells color."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]セルに色を塗りつぶす
     #
     # $gridで指定するセルに$colorで指定する色に塗りつぶす。$gridで指定するメンバを次に示す。
     # row：色を塗りつぶすセルの行番号
     # col：色を塗りつぶすセルの列番号
     # セルに色を塗りつぶしに失敗した場合、$falseを返す。
     #
     # @access public
     # @param int $color ColorIndexの数値
     # @param array $grid セルの範囲
     # @return bool セルに色を塗りつぶした結果の状態
     # @see なし
     # @throws なし
     #>
    [bool]ColorCell([int]$color, [array]$grid){
        if($this.ColorCell($color, $grid, $grid) -eq $false){
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]文字の色を変える
     #
     # $originと$endで指定するセル範囲を、$colorで指定する色に変更する。
     # $originで指定するメンバを次に示す。
     # int row：範囲指定の始まりの行番号
     # int col：範囲指定の始まりの列番号
     # $endで指定するメンバを次に示す。
     # int row：範囲指定の終わりの行番号
     # int col：範囲指定の終わりの列番号
     # 文字の色の変更で失敗した場合、$falseを返す。
     #
     # @access public
     # @param int $color ColorIndexの数値
     # @param array $origin セルの始まり
     # @param array $end セルの終わり
     # @return bool 文字の色を変える結果の状態
     # @see なし
     # @throws 文字の色の変更で例外発生時、$falseを返す。
     #>
    [bool]ColorFont([int]$color, [array]$origin, [array]$end){
        try{
            $this.Sheet.Range( `
                $this.Sheet.Cells([int]$origin.row, [int]$origin.col), 
                $this.Sheet.Cells([int]$end.row, [int]$end.col) `
            ).Font.ColorIndex = $color
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to change font color."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]文字の色を変える
     #
     # $gridで指定するセルを、$colorで指定する色に変更する。$gridで指定するメンバを次に示す。
     # row：文字の色を変えるセルの行番号
     # col：文字の色を変えるセルの列番号
     # 文字の色の変更で失敗した場合、$falseを返す。
     #
     # @access public
     # @param int $color ColorIndexの数値
     # @param array $grid 文字の色を変えるセル
     # @return bool 文字の色を変える結果の状態
     # @see なし
     # @throws なし
     #>
    [bool]ColorFont([int]$color, [array]$grid){
        if($this.ColorFont($color, $grid, $grid) -eq $false){
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]セルに罫線を引く
     #
     # $originと$endで指定するセル範囲に、罫線を引く。
     # $originで指定するメンバを次に示す。
     # int row：範囲指定の始まりの行番号
     # int col：範囲指定の始まりの列番号
     # $endで指定するメンバを次に示す。
     # int row：範囲指定の終わりの行番号
     # int col：範囲指定の終わりの列番号
     # 指定するセル範囲に罫線を引くことに失敗した場合、$falseを返す。
     #
     # @access public
     # @param array $origin セルの始まり
     # @param array $end セルの終わり
     # @return bool セルに罫線を引く結果の状態
     # @see なし
     # @throws セル範囲に罫線を引くで例外発生時、$falseを返す。
     #>
    [bool]BoderCell([array]$origin, [array]$end){
        try{
            $this.Sheet.Range( `
                $this.Sheet.Cells([int]$origin.row, [int]$origin.col), 
                $this.Sheet.Cells([int]$end.row, [int]$end.col) `
            ).Borders.LineStyle = 1
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to line cell border."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]セルに罫線を引く
     #
     # $gridで指定するセルに罫線を引く。$gridで指定するメンバを次に示す。
     # row：罫線を引くセルの行番号
     # col：罫線を引くセルの列番号
     # 罫線を引くことに失敗した場合、$falseを返す。
     # 
     # @access public
     # @param array $grid 罫線を引くセル
     # @return bool セルに罫線を引く結果の状態
     # @see なし
     # @throws セルに罫線を引くで例外発生時、$falseを返す。
     #>
    [bool]BoderCell([array]$grid){
        if($this.BoderCell($grid, $grid) -eq $false){
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]行の削除
     #
     # $src_rowで指定する行を削除する。
     # 行の削除に失敗した場合、$falseを返す。
     #
     # @access public
     # @param int $src_row 削除する行番号
     # @return bool 行削除の結果の状態
     # @see なし
     # @throws 行の削除で例外発生時、$falseを返す。
     #>
    [bool]DeleteRow([int]$src_row){
        try{
            $this.Sheet.rows.Item($src_row).Delete()
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to delete row."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]指定する行で改ページ
     #
     # $src_rowで指定する行で改ページする。
     # 指定する行で改ページに失敗した場合、$falseを返す。
     #
     # @access public
     # @param $src_row 改ページする行
     # @return bool 改ページを設定する結果の状態
     # @see なし
     # @throws 指定する行の改ページで例外発生時、$falseを返す。
     #>
    # 名前                  値      説明
    # xlPageBreakAutomatic  -4105   Excel が自動的に改ページを追加します。
    # xlPageBreakManual     -4135   改ページは手動で挿入されます。
    # xlPageBreakNone       -4142   ワークシートに改ページは挿入されません。
    [bool]setpagebreak($src_row){
        try {
            $this.Sheet.Range($src_row).pagebreak = $this.define.xlPageBreakManual
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to set pagebreak."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]指定する行で改ページ
     #
     # $src_rowで指定する行で改ページする。
     # 指定する行で改ページに失敗した場合、$falseを返す。
     #
     # @access public
     # @param $src_row 改ページする行
     # @return bool 改ページを設定する結果の状態
     # @see なし
     # @throws なし
     #>
    [bool]PageBreak($src_row){
        return $this.setpagebreak($src_row)
    }

    <#
     # [MsExcelCell]指定する行の内容に合わせて高さを自動調整する
     #
     # $src_rowで指定する行の内容に合わせて高さを自動調整する。
     # 行の内容に合わせて高さを自動調整に失敗した場合、$falseを返す。
     #
     # @access public
     # @param int $src_row 高さ指定をする行
     # @return bool 指定する行の内容に合わせて高さを自動調整する結果の状態
     # @see なし
     # @throws なし
     #>
    [bool]RowAutofit([int]$src_row){
        try{
            $this.Sheet.rows.item($src_row).Autofit()
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to fit row automatically."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]指定する行の高さを設定する
     #
     # $src_rowで指定する行を、$RowHeightで指定する高さに設定する。
     # 行の高さの設定に失敗した場合、$falseを返す。
     #
     # @access public
     # @param int $src_row 行番号
     # @param int $RowHeight 行の高さ
     # @return bool 行の高さを設定する結果の状態
     # @see なし
     # @throws 指定する行の高さを設定で例外発生時、$falseを返す。
     #>
    [bool]SetRowHeight([int]$src_row,[int]$RowHeight){
        try{
            $this.Sheet.rows.item($src_row).RowHeight = $RowHeight
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to set row height."
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]指定する行の高さを取得する
     #
     # $src_rowで指定する行の高さを取得する。
     #
     # @access public
     # @param int $src_row 取得する行
     # @return int 行の高さ
     # @see なし
     # @throws 指定する行の高さを取得するで例外発生時、$falseを返す。
     #>
    [int]GetRowHeight([int]$src_row){
        try{
            [int]$RowSize = $this.Sheet.rows.item($src_row).RowHeight
        }catch{
            $script:LAST_ERROR_MESSAGE="Fail to get row height."
            return $false
        }
        return $RowSize
    }

    <#
     # [MsExcelCell]指定する行範囲の合計の高さを取得する
     #
     # $start_rowから$end_rowで指定する行範囲の高さの合計を取得する。
     #
     # @access public
     # @param int $start_row 指定開始行
     # @param int $end_row 指定終了行
     # @return int 高さの合計
     # @see なし
     # @throws なし
     #>
    [int]GetRowHeight([int]$start_row, [int]$end_row){
        [int]$total=0
        if($start_row -gt $end_row){
            $script:LAST_ERROR_MESSAGE="Invalid input row."
            return $false
        }
        for([int]$i=$start_row; $i -le $end_row; $i++){
            $total+=$this.GetRowHeight($i)
        }
        return $total
    }
}