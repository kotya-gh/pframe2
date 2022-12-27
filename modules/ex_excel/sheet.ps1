<#
 # [MsExcelSheet]Excel�t�@�C�������p�N���X
 #
 # Microsoft Excel 2007�^�C�v�̃t�@�C���𑀍삷��B
 #
 # @access public
 # @author - <-@->
 # @copyright MIT
 # @category Excel�t�@�C���̑���
 # @package �Ȃ�
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
     # [MsExcelSheet]�V�[�g�̐����擾����
     #
     # Excelbook�̃V�[�g�����擾����B
     #
     # @access public
     # @param �Ȃ�
     # @return int �V�[�g�̐�
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [int]GetSheetCount(){
        return $this.Book.Sheets.Count
    }

    <#
     # [MsExcelSheet]�V�[�g�����擾����
     #
     # $th�Ŏw�肷��V�[�g�ԍ��̃V�[�g�����擾����B
     # 
     # @access public
     # @param int $th �V�[�g�ԍ�
     # @return string �V�[�g��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [string]GetSheetName([int]$th){
        return $this.Book.Sheets($th).Name        
    }

    <#
     # [MsExcelSheet]�V�[�g�������ׂĎ擾����
     #
     # �V�[�g�������ׂĎ擾����B
     # 
     # @access public
     # @param �Ȃ�
     # @return array �V�[�g��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [array]GetSheetNames(){
        [array]$names=@()
        for([int]$i=1; $i -le $this.GetSheetCount(); $i++){
            $names+=$this.Book.Sheets($i).Name
        }
        return $names    
    }

    <#
     # [MsExcelSheet]�V�[�g�����L������
     #
     # $th�Ŏw�肷��V�[�g��$sheet_name�Ŏw�肷�镶���ŋL������B
     # �V�[�g���̏d���A�܂��̓V�[�g���̋L���Ɏ��s�����ꍇ�A$false��Ԃ��B
     # 
     # @access public
     # @param int $th �L������V�[�g�ԍ�
     # @param string $sheet_name �L�����镶����
     # @return bool �V�[�g�����L���������ʂ̏��
     # @see �Ȃ�
     # @throws �Ȃ�
     #>
    [bool]SetSheetName([int]$th, [string]$sheet_name){
        if($this.ValidateSheetName($sheet_name) -eq $false){
            return $false
        }
        $this.Book.Sheets($th).Name=$sheet_name
        return $true
    }

    <#
     # [MsExcelSheet]�L���ȃV�[�g���ł��邩���؂���
     #
     # $sheet_name�Ŏw�肷��V�[�g�����L���ȃV�[�g���ł��邩�����؂���B
     # �L���ȃV�[�g���ł��邩�̌��؂Ɏ��s�����ꍇ�A$false��Ԃ��B
     # 
     # @access public
     # @param string $sheet_name �V�[�g��
     # @return bool �L���ȃV�[�g���ł��邩���؂��錋�ʂ̏��
     # @see �Ȃ�
     # @throws �Ȃ�
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
     # [MsExcelSheet]�V�[�g�̃R�s�[����
     #
     # $src_sheet_name�Ŏw�肷��R�s�[���V�[�g�����ꎞ�I�ɕύX���A
     # $dst_sheet_name�Ŏw�肷��R�s�[��V�[�g���쐬��A�R�s�[�������ɖ߂��B
     # �V�[�g�̃R�s�[�����Ɏ��s�����ꍇ�A$false��Ԃ��B
     # 
     # @access public
     # @param string $src_sheet_name �R�s�[���V�[�g��
     # @param string $dst_sheet_name �R�s�[��V�[�g��
     # @return bool �V�[�g�����R�s�[�������ʂ̏��
     # @see �Ȃ�
     # @throws �V�[�g�̃R�s�[�����ŗ�O�������A$false��Ԃ��B
     #>
    [bool]CopyWorkSheet([string]$src_sheet_name, [string]$dst_sheet_name){
        # �V�[�g�̃R�s�[�����F�R�s�[���V�[�g�����ꎞ�I�ɕύX���A�R�s�[��V�[�g���쐬��A�R�s�[�������ɖ߂��B
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
     # [MsExcelSheet]����͈͂���������
     #
     # $excebook�Ŏw�肷�����͈͂���������B
     # 
     # @access public
     # @param object $excebook ����͈͂���������Excelbook�̃I�u�W�F�N�g
     # @return bool ����͈͂��������錋�ʂ̏��
     # @see �Ȃ�
     # @throws ����͈͂������ŗ�O�������A$false��Ԃ��B
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
     # [MsExcelSheet]����͈͂�ݒ肷��
     #
     # $setrange�Ŏw�肷��͈͂�����͈͂Ƃ��Đݒ肷��B
     # ����͈͂̐ݒ�Ɏ��s�����ꍇ�A$false��Ԃ��B
     #
     # @access public
     # @param string $setrange ����͈�
     # @return bool ����͈͂�ݒ肷�錋�ʂ̏��
     # @see �Ȃ�
     # @throws ����͈͂�ݒ�ŗ�O�������A$false��Ԃ��B
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