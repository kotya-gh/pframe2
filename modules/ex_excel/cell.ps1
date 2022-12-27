<#
 # [MsExcelCell]Excel�t�@�C�������p�N���X
 #
 # Microsoft Excel 2007�^�C�v�̃t�@�C���𑀍삷��B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category Excel�t�@�C���̑���
 # @package �Ȃ�
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
     # [MsExcelCell]�s��}������
     #
     # $row�Ŏw�肷��s����}������B
     # �s�̑}���Ŏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param int $row �}������s�̐�
     # @return bool �s�̑}�����ʂ̏��
     # @see �Ȃ�
     # @throws �s�̑}���ŗ�O�������A$false��Ԃ��B
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
     # [MsExcelCell]�s���R�s�[����
     #
     # $src_row�Ŏw�肷��s���R�s�[���A$dst_row�Ŏw�肷��s�ɓ\��t����B
     # �s�̃R�s�[�Ŏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param int $src_row �R�s�[���̍s
     # @param int $dst_row �R�s�[��̍s
     # @return bool �s�̃R�s�[���ʂ̏��
     # @see �Ȃ�
     # @throws �s�R�s�[�ŗ�O�������A$false��Ԃ��B
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
     # [MsExcelCell]�Z�����R�s�[����
     #
     # $src�Ŏw�肷��Z���̓��e���擾���A$dst�Ŏw�肷��Z���ɓ\��t����B
     # $src�Ŏw�肷�郁���o�����Ɏ����B
     # row�F�R�s�[���̃Z���̍s�ԍ�
     # col�F�R�s�[���̃Z���̗�ԍ�
     # $dst�Ŏw�肷�郁���o�����Ɏ����B
     # row�F�R�s�[��̃Z���̍s�ԍ�
     # col�F�R�s�[��̃Z���̗�ԍ�
     # �Z�����R�s�[���s���A$false��Ԃ��B
     #
     # @access public
     # @param array $src �R�s�[���Z��
     # @param array $dst �R�s�[��Z��
     # @return bool �Z���̃R�s�[���ʂ̏��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]CopyCell([array]$src, [array]$dst){
        if($this.CopyCell($src, $dst, $dst) -eq $false){
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]�͈͑I�������Z���ɃR�s�[����
     #
     # $src�Ŏw�肷��Z���̓��e���A$dst_origin����$dst_end�Ŏw�肷��͈͂ɓ\��t����B
     # $src�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F�R�s�[���̃Z���̍s�ԍ�
     # int col�F�R�s�[���̃Z���̗�ԍ�
     # $dst_origin�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F�R�s�[��͈͎̔w��̎n�܂�̍s�ԍ�
     # int col�F�R�s�[��͈͎̔w��̎n�܂�̗�ԍ�
     # $dst_end�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F�R�s�[��͈͎̔w��̏I���̍s�ԍ�
     # int col�F�R�s�[��͈͎̔w��̏I���̗�ԍ�
     # �Z���͈̔͑I���ŃR�s�[���s���A$false��Ԃ��B
     #
     # @access public
     # @param array $src �R�s�[���Z��
     # @param array $dst_origin �R�s�[��Z���̎n�܂�
     # @param array $dst_end �R�s�[��Z���̏I���
     # @return bool �Z���͈̓R�s�[���ʂ̏��
     # @see �Ȃ�
     # @throws �Z����͈͑I�����ăR�s�[�ŗ�O�������A$false��Ԃ��B
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
     # [MsExcelCell]�Z���̌���
     #
     # $origin�Ŏw�肷��Z���ƁA$end�Ŏw�肷��Z������������B
     # $origin�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F��������Z���̎n�߂̍s�ԍ�
     # int col�F��������Z���̎n�߂̗�ԍ�
     # $end�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F��������Z���̏I���̍s�ԍ�
     # int col�F��������Z���̏I���̗�ԍ�
     # �Z���̌����Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param array $origin ��������Z���̎n�܂�
     # @param array $end ��������Z���̏I���
     # @return bool �Z���̌������ʂ̏��
     # @see �Ȃ�
     # @throws �Z�������ŗ�O�������A$false��Ԃ��B
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
     # [MsExcelCell]�Z���ɃR�����g��ǉ�����
     #
     # $data�Ŏw�肷��Z���ɃR�����g��ǉ�����B$data�Ŏw�肷�郁���o�����Ɏ����B
     # int    row�F�R�����g��ǉ�����Z���̍s�ԍ�
     # int    col�F�R�����g��ǉ�����Z���̗�ԍ�
     # string comment�F�ǉ�����R�����g�̓��e
     # �Z���ɃR�����g��ǉ��Ŏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param array $data �R�����g��ǉ�����Z��
     # @return bool �R�����g�ǉ����ʂ̏��
     # @see �Ȃ�
     # @throws �Z���ɃR�����g��ǉ��ŗ�O�������A$false��Ԃ��B
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
     # [MsExcelCell]������𕡐��̃Z���ɋL������
     #
     # $data_object�Ŏw�肵�����ꂼ��̓��e�𕡐��̃Z���ɋL������B$data_object�Ŏw�肷�郁���o�����Ɏ����B
     #  int row�F�s�ԍ�
     #  int col�F��ԍ�
     #      value�F�L������l
     # �Z���ɋL���Ŏ��s�����ꍇ�A$false��Ԃ��B
     # @access public
     # @param array $data_object �L��������e
     # @return bool �Z���L�����ʂ̏��
     # @see �Ȃ�
     # @throws �Ȃ�
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
     # [MsExcelCell]�Z���ɋL������
     #
     # $data�Ŏw�肷����e���Z���ɋL������B$data�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F�s�ԍ�
     # int col�F��ԍ�
     #     value�F�L������l
     # �Z���ɋL���Ŏ��s�����ꍇ�A$false��Ԃ��B
     # @access public
     # @param array $data �L��������e
     # @return bool �Z���L�����ʂ̏��
     # @see �Ȃ�
     # @throws �Z���ɋL���ŗ�O�������A$false��Ԃ��B
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
     # [MsExcelCell]�Z���̓��e���擾����
     #
     # $grid�Ŏw�肷��Z���̓��e���擾����B$grid�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F���e���擾����Z���̍s�ԍ�
     # int col�F���e���擾����Z���̗�ԍ�
     #
     # @access public
     # @param array $grid ���e���擾����Z��
     # @return string �Z���̓��e
     # @see �Ȃ�
     # @throws �Ȃ�
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
     # [MsExcelCell]�Z���ɑ������Ă���v�Z�����擾����
     #
     # $grid�Ŏw�肷��Z���̌v�Z�����擾����B$grid�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F�Z���̍s�ԍ�
     # int col�F�Z���̗�ԍ�
     # �Z���̌v�Z�����擾�Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param array $grid �v�Z�����������Ă���Z��
     # @return string �v�Z��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [string]FormulaCell([array]$grid){
        return $this.Sheet.Cells([int]$grid.row, [int]$grid.col).Formula
    }

    <#
     # [MsExcelCell]�w�肷��Z���͈͂̕����𑾎��ɂ���
     #
     # $origin����$end�Ŏw�肷��Z���͈͂̕����𑾎��ɂ���B
     # $origin�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F�͈͎w��̎n�܂�̍s�ԍ�
     # int col�F�͈͎w��̎n�܂�̗�ԍ�
     # $end�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F�͈͎w��̏I���̍s�ԍ�
     # int col�F�͈͎w��̏I���̗�ԍ�
     # �Z���͈͂̕����𑾎��֕ϊ��Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param array $origin �Z���̎n�܂�
     # @param array $end �Z���̏I���
     # @return bool �����𑾎��ɂ��錋�ʂ̏��
     # @see �Ȃ�
     # @throws �����𑾎��ɕϊ��ŗ�O�������A$false��Ԃ��B
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
     # [MsExcelCell]�����𑾎��ɂ���
     #
     # $grid�Ŏw�肷��Z���̕����𑾎��ɂ���B$grid�Ŏw�肷�郁���o�����Ɏ����B
     # row�F�����𑾎��ɂ���Z���̍s�ԍ�
     # col�F�����𑾎��ɂ���Z���̗�ԍ�
     # �����𑾎��֕ϊ��Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param array $grid �����ɂ���Z��
     # @return bool �����𑾎��ɂ��錋�ʂ̏��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]BoldCell([array]$grid){
        if($this.BoldCell($grid, $grid) -eq $false){
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]�Z���ɐF��h��Ԃ�
     #
     # $origin��$end�Ŏw�肷��Z���͈͂��A$color�Ŏw�肷��F�ɓh��Ԃ��B
     # $origin�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F�͈͎w��̎n�܂�̍s�ԍ�
     # int col�F�͈͎w��̎n�܂�̗�ԍ�
     # $end�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F�͈͎w��̏I���̍s�ԍ�
     # int col�F�͈͎w��̏I���̗�ԍ�
     # �Z���ɐF��h��Ԃ��Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param int $color ColorIndex�̐��l
     # @param array $origin �Z���̎n�܂�
     # @param array $end �Z���̏I���
     # @return bool �Z���ɐF��h��Ԃ������ʂ̏��
     # @see �Ȃ�
     # @throws �Z���ɐF��h��Ԃ��ŗ�O�������A$false��Ԃ��B
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
     # [MsExcelCell]�Z���ɐF��h��Ԃ�
     #
     # $grid�Ŏw�肷��Z����$color�Ŏw�肷��F�ɓh��Ԃ��B$grid�Ŏw�肷�郁���o�����Ɏ����B
     # row�F�F��h��Ԃ��Z���̍s�ԍ�
     # col�F�F��h��Ԃ��Z���̗�ԍ�
     # �Z���ɐF��h��Ԃ��Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param int $color ColorIndex�̐��l
     # @param array $grid �Z���͈̔�
     # @return bool �Z���ɐF��h��Ԃ������ʂ̏��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]ColorCell([int]$color, [array]$grid){
        if($this.ColorCell($color, $grid, $grid) -eq $false){
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]�����̐F��ς���
     #
     # $origin��$end�Ŏw�肷��Z���͈͂��A$color�Ŏw�肷��F�ɕύX����B
     # $origin�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F�͈͎w��̎n�܂�̍s�ԍ�
     # int col�F�͈͎w��̎n�܂�̗�ԍ�
     # $end�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F�͈͎w��̏I���̍s�ԍ�
     # int col�F�͈͎w��̏I���̗�ԍ�
     # �����̐F�̕ύX�Ŏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param int $color ColorIndex�̐��l
     # @param array $origin �Z���̎n�܂�
     # @param array $end �Z���̏I���
     # @return bool �����̐F��ς��錋�ʂ̏��
     # @see �Ȃ�
     # @throws �����̐F�̕ύX�ŗ�O�������A$false��Ԃ��B
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
     # [MsExcelCell]�����̐F��ς���
     #
     # $grid�Ŏw�肷��Z�����A$color�Ŏw�肷��F�ɕύX����B$grid�Ŏw�肷�郁���o�����Ɏ����B
     # row�F�����̐F��ς���Z���̍s�ԍ�
     # col�F�����̐F��ς���Z���̗�ԍ�
     # �����̐F�̕ύX�Ŏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param int $color ColorIndex�̐��l
     # @param array $grid �����̐F��ς���Z��
     # @return bool �����̐F��ς��錋�ʂ̏��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]ColorFont([int]$color, [array]$grid){
        if($this.ColorFont($color, $grid, $grid) -eq $false){
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]�Z���Ɍr��������
     #
     # $origin��$end�Ŏw�肷��Z���͈͂ɁA�r���������B
     # $origin�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F�͈͎w��̎n�܂�̍s�ԍ�
     # int col�F�͈͎w��̎n�܂�̗�ԍ�
     # $end�Ŏw�肷�郁���o�����Ɏ����B
     # int row�F�͈͎w��̏I���̍s�ԍ�
     # int col�F�͈͎w��̏I���̗�ԍ�
     # �w�肷��Z���͈͂Ɍr�����������ƂɎ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param array $origin �Z���̎n�܂�
     # @param array $end �Z���̏I���
     # @return bool �Z���Ɍr�����������ʂ̏��
     # @see �Ȃ�
     # @throws �Z���͈͂Ɍr���������ŗ�O�������A$false��Ԃ��B
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
     # [MsExcelCell]�Z���Ɍr��������
     #
     # $grid�Ŏw�肷��Z���Ɍr���������B$grid�Ŏw�肷�郁���o�����Ɏ����B
     # row�F�r���������Z���̍s�ԍ�
     # col�F�r���������Z���̗�ԍ�
     # �r�����������ƂɎ��s�����ꍇ�A$false��Ԃ��B
     # 
     # @access public
     # @param array $grid �r���������Z��
     # @return bool �Z���Ɍr�����������ʂ̏��
     # @see �Ȃ�
     # @throws �Z���Ɍr���������ŗ�O�������A$false��Ԃ��B
     #>
    [bool]BoderCell([array]$grid){
        if($this.BoderCell($grid, $grid) -eq $false){
            return $false
        }
        return $true
    }

    <#
     # [MsExcelCell]�s�̍폜
     #
     # $src_row�Ŏw�肷��s���폜����B
     # �s�̍폜�Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param int $src_row �폜����s�ԍ�
     # @return bool �s�폜�̌��ʂ̏��
     # @see �Ȃ�
     # @throws �s�̍폜�ŗ�O�������A$false��Ԃ��B
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
     # [MsExcelCell]�w�肷��s�ŉ��y�[�W
     #
     # $src_row�Ŏw�肷��s�ŉ��y�[�W����B
     # �w�肷��s�ŉ��y�[�W�Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param $src_row ���y�[�W����s
     # @return bool ���y�[�W��ݒ肷�錋�ʂ̏��
     # @see �Ȃ�
     # @throws �w�肷��s�̉��y�[�W�ŗ�O�������A$false��Ԃ��B
     #>
    # ���O                  �l      ����
    # xlPageBreakAutomatic  -4105   Excel �������I�ɉ��y�[�W��ǉ����܂��B
    # xlPageBreakManual     -4135   ���y�[�W�͎蓮�ő}������܂��B
    # xlPageBreakNone       -4142   ���[�N�V�[�g�ɉ��y�[�W�͑}������܂���B
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
     # [MsExcelCell]�w�肷��s�ŉ��y�[�W
     #
     # $src_row�Ŏw�肷��s�ŉ��y�[�W����B
     # �w�肷��s�ŉ��y�[�W�Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param $src_row ���y�[�W����s
     # @return bool ���y�[�W��ݒ肷�錋�ʂ̏��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]PageBreak($src_row){
        return $this.setpagebreak($src_row)
    }

    <#
     # [MsExcelCell]�w�肷��s�̓��e�ɍ��킹�č�����������������
     #
     # $src_row�Ŏw�肷��s�̓��e�ɍ��킹�č�����������������B
     # �s�̓��e�ɍ��킹�č��������������Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param int $src_row �����w�������s
     # @return bool �w�肷��s�̓��e�ɍ��킹�č����������������錋�ʂ̏��
     # @see �Ȃ�
     # @throws �Ȃ�
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
     # [MsExcelCell]�w�肷��s�̍�����ݒ肷��
     #
     # $src_row�Ŏw�肷��s���A$RowHeight�Ŏw�肷�鍂���ɐݒ肷��B
     # �s�̍����̐ݒ�Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param int $src_row �s�ԍ�
     # @param int $RowHeight �s�̍���
     # @return bool �s�̍�����ݒ肷�錋�ʂ̏��
     # @see �Ȃ�
     # @throws �w�肷��s�̍�����ݒ�ŗ�O�������A$false��Ԃ��B
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
     # [MsExcelCell]�w�肷��s�̍������擾����
     #
     # $src_row�Ŏw�肷��s�̍������擾����B
     #
     # @access public
     # @param int $src_row �擾����s
     # @return int �s�̍���
     # @see �Ȃ�
     # @throws �w�肷��s�̍������擾����ŗ�O�������A$false��Ԃ��B
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
     # [MsExcelCell]�w�肷��s�͈͂̍��v�̍������擾����
     #
     # $start_row����$end_row�Ŏw�肷��s�͈͂̍����̍��v���擾����B
     #
     # @access public
     # @param int $start_row �w��J�n�s
     # @param int $end_row �w��I���s
     # @return int �����̍��v
     # @see �Ȃ�
     # @throws �Ȃ�
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