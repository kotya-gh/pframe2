. (Join-Path $PSScriptRoot "sheet.ps1")
. (Join-Path $PSScriptRoot "cell.ps1")

<#
 # [MsExcel]Excel�t�@�C�������p�N���X
 #
 # Microsoft Excel 2007�^�C�v�̃t�@�C���𑀍삷��B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category Excel�t�@�C���̑���
 # @package �Ȃ�
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
     # [MsExcel]Excel���I������
     #
     # MsExcel Book���I����AMsExcel���I������B
     #
     # @access public
     # @param �Ȃ�
     # @return �Ȃ�
     # @see MsExcel.CloseBook MsExcel.CloseExcel
     # @throws �Ȃ�
     #>
    [void]Quit(){
        $this.CloseBook()
        $this.CloseExcel()
    }

    <#
     # [MsExcel]Excel���N������
     #
     # Excel���N������B
     # Excel�N���Ɏ��s�����ꍇ�́A$false��Ԃ��B
     # 
     # @access public
     # @param �Ȃ�
     # @return bool Excel�N�����ʂ̏��
     # @see �Ȃ�
     # @throws Excel�N�����ɗ�O�������A$false��Ԃ��B
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
     # [MsExcel]Excel���I������
     #
     # Excel���I������B
     # Excel�I�����s���́A$false��Ԃ��B
     #
     # @access public
     # @param �Ȃ�
     # @return bool Excel�I�����ʂ̏��
     # @see �Ȃ�
     # @throws Excel�I���ŗ�O�������A$false��Ԃ��B
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
     # [MsExcel]Excel��ۑ�����
     #
     # $file_path�Ŏw�肷��Excel��ۑ�����B
     # �ۑ��Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param string $file_path �ۑ�����Excel�̃p�X
     # @return bool Excel�ۑ����ʂ̏��
     # @see �Ȃ�
     # @throws Excel�ۑ��ŗ�O�������A$false��Ԃ��B
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
     # [MsExcel]�㏑���ۑ�����
     #
     # Excel�̓��e���㏑���ۑ�����B
     # �㏑���Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param �Ȃ�
     # @return bool �㏑���ۑ����ʂ̏��
     # @see �Ȃ�
     # @throws �㏑���ۑ��ŗ�O�������A$false��Ԃ��B
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
     # [MsExcel]ExcelBook���N������
     #
     # $file_path�Ŏw�肷��ExcelBook���N������B
     # ExcelBook�̋N���Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param string $file_path �N������ExcelBook�̃p�X
     # @return bool ExcelBook�N�����ʂ̏��
     # @see �Ȃ�
     # @throws ExcelBook�N���ŗ�O�������A$false��Ԃ��B
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
     # [MsExcel]ExcelBook���I������
     #
     # ExcelBook���I������B
     # ExcelBook�̏I���Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param �Ȃ�
     # @return bool ExcelBook�I�����ʂ̏��
     # @see �Ȃ�
     # @throws ExcelBook�I���ŗ�O�������A$false��Ԃ��B
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
     # [MsExcel]�V�[�g�����擾����
     #
     # $sheet_name�Ŏw�肷��V�[�g�����擾����B
     # �V�[�g���擾�Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param string $sheet_name �V�[�g��
     # @return bool �V�[�g���擾���ʂ̏��
     # @see �Ȃ�
     # @throws �V�[�g�����擾�ŗ�O�������A$false��Ԃ��B
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
     # [MsExcel]ExcelSheet�̍s�����擾����
     #
     # $sheet_num�Ŏw�肷��V�[�g�́A�g�p����Ă���Z���͈͂̍s�����擾����B
     # 
     # @access public
     # @param int $sheet_num �s�����擾����V�[�g�ԍ�
     # @return int �g�p����Ă���Z���͈͂̍s��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [int]GetSheetRows([int]$sheet_num){
        [object]$SheetObj=$this.Excel.Worksheets.Item($sheet_num)
        return $SheetObj.UsedRange.Rows.Count
    }

    <#
     # [MsExcel]�w�b�_�ƃt�b�^�̃I�u�W�F�N�g���擾����
     #
     # $sheet_num�Ŏw�肷��w�b�_�ƃt�b�^�̃I�u�W�F�N�g���擾����B
     # 
     # @access public
     # @param int $sheet_num �V�[�g�ԍ�
     # @return object �w�b�_�ƃt�b�^�̃I�u�W�F�N�g
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [object]GetSheetHeaderFooter([int]$sheet_num){
        [object]$SheetObj=$this.Excel.Worksheets.Item($sheet_num)
        return $SheetObj.pageSetup
    }

    <#
     # [MsExcel]�V�[�g���폜����
     #
     # $sheet_name�Ŏw�肷��V�[�g���폜����B
     # �V�[�g�̍폜�Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param string $sheet_name �V�[�g��
     # @return bool �V�[�g�폜���ʂ̏��
     # @see �Ȃ�
     # @throws �V�[�g���폜�ŗ�O�������A$false��Ԃ��B
     #>
    [bool]DeleteSheet([string]$sheet_name){
        [object]$ExcelSheet=New-Object MsExcelSheet($this.Book)
        # �ύX��̃V�[�g���̃o���f�[�V����
        if($ExcelSheet.ValidateSheetName($sheet_name) -eq $false){
            # �G���[���b�Z�[�W��ValidateSheetName�ŕԂ�
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
     # [MsExcel]�V�[�g���R�s�[����
     #
     # $src_sheet_name�Ŏw�肷��V�[�g��$dst_sheet_name�Ŏw�肷��V�[�g���ɕύX���ăR�s�[����B
     # �V�[�g���̏d�����������ꍇ�A����уV�[�g�R�s�[�Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param string $src_sheet_name �R�s�[���V�[�g��
     # @param string $dst_sheet_name �R�s�[��V�[�g��
     # @return bool �V�[�g�R�s�[���ʂ̏��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]CopySheet([string]$src_sheet_name, [string]$dst_sheet_name){
        [object]$ExcelSheet=New-Object MsExcelSheet($this.Book)
        # �ύX��̃V�[�g���̃o���f�[�V����
        if($ExcelSheet.ValidateSheetName($dst_sheet_name) -eq $false){
            # �G���[���b�Z�[�W��ValidateSheetName�ŕԂ�
            return $false
        }

        # �ύX�O�E�ύX��̃V�[�g���̏d�����m�F����B
        if($src_sheet_name -eq $dst_sheet_name){
            $script:LAST_ERROR_MESSAGE="Sheet name is duplicated."
            return $false
        }

        # Book���ɑ��݂���V�[�g���ƕύX��̃V�[�g���̏d�����m�F����B
        [array]$names=$ExcelSheet.GetSheetNames
        [object]$arr=New-Object ExArray
        if($arr.ArraySearch($dst_sheet_name, $names) -eq $true){
            $script:LAST_ERROR_MESSAGE="Sheet name is duplicated."
            return $false
        }
        
        if($ExcelSheet.CopyWorkSheet($src_sheet_name,$dst_sheet_name) -eq $false){
            # �G���[���b�Z�[�W��CopyWorkSheet�ŕԂ�
            return $false
        }
        return $true
    }
}