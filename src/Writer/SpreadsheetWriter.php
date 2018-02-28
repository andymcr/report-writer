<?php
namespace Uk\Me\AndyCarpenter\Writer;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\Style;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


class SpreadsheetWriter extends Writer
{

    private $excel;
    private $_sheet_used;
    private $rowIndex;
    private $columnIndex;
    private $lastColumnIndex;
    private $tableStart;
    private $tableStyle;
    private $_table_maximum_column;
    private $_merge_start;

    public function __construct()
    {
        $this->spreadsheet = new Spreadsheet();
        $this->setAutoSizeColumns();
        
        $this->addSection();
    }

    public function setDocumentTitle($title)
    {
    	$this->spreadsheet->getProperties()->setTitle($title);
    }

    public function setDocumentSubject($subject)
    {
    	$this->spreadsheet->getProperties()->setTitle($subject);
    }

    public function setDocumentDescription($description)
    {
    	$this->spreadsheet->getProperties()->setDescription($description);
    }

    public function setDocumentCreator($creator)
    {
    	$this->spreadsheet->getProperties()->setCreator($creator);
    }

    public function setDocumentLastModifiedBy($modifiedBy)
    {
    	$this->spreadsheet->getProperties()->setLastModifiedBy($modifiedBy);
    }

    public function setDocumentCompany($company)
    {
    	$this->_word->getDocInfo()->setCompany($company);
    }

    public function getFileExtension()
    {
    	return '.xlsx';
    }

    public function getContentType()
    {
    	return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    }

    public function addSection($title = NULL, $title_level = 1)
    {
        if (!$this->_sheet_used)
            $sheet = $this->spreadsheet->getActiveSheet();
        else
            $sheet = $this->spreadsheet->createSheet($this->spreadsheet->getSheetCount());
        $page_setup = $sheet->getPageSetup();
        $page_setup->setPaperSize(PageSetup::PAPERSIZE_A4);
        
        $this->rowIndex = 1;
        $this->columnIndex = 'A';
        $this->spreadsheet->setActiveSheetIndex($this->spreadsheet->getSheetCount() - 1);
        $this->setSectionTitle($title);
        $this->setAutoSizeColumns();

        $this->spreadsheet->getActiveSheet()->getHeaderFooter()->setOddFooter('&R&D &T');
        $this->spreadsheet->getActiveSheet()->getHeaderFooter()->setEvenFooter('&R&D &T');

        $this->_sheet_used = FALSE;
    }

    public function setSectionTitle($title)
    {
        if (!empty($title))
            $this->spreadsheet->getActiveSheet()->setTitle(
                str_replace(array('*', '/', '\\', '?', '[', ']'), '_',
                    str_replace(':', ' -', $title)));
    }

    public function setSectionOrientation($orientation)
    {
        $pageSetup = $this->spreadsheet->getActiveSheet()->getPageSetup();
        $pageSetup->setOrientation(
            $orientation === Writer::ORIENTATION_PORTRAIT
                ? PageSetup::ORIENTATION_PORTRAIT
                : PageSetup::ORIENTATION_LANDSCAPE);
    }

    public function addPageBreak()
    {
        
    }

    public function writeTitle($title, $level, $style)
    {
        $this->spreadsheet->getActiveSheet()->setCellValue($this->columnIndex . $this->rowIndex, $title);
        $rowRange = 'A' . $this->rowIndex . ':' . 'Z' . $this->rowIndex;
        $this->spreadsheet->getActiveSheet()->duplicateStyle($style, $rowRange);
        $this->mergeCells($rowRange);
        $this->nextRow();

        $this->_sheet_used = TRUE;
    }

    public function writeText($text)
    {
        $this->spreadsheet->getActiveSheet()->setCellValue($this->columnIndex . $this->rowIndex, $text);
        $this->nextColumn();

        $this->_sheet_used = TRUE;
    }

    public function startTable($style, $first_row_style = NULL)
    {
        $this->tableStart = $this->columnIndex  . $this->rowIndex;
        $this->tableStyle = $style;
        $this->_table_maximum_column = 'A';
    }

    public function endTable()
    {
        $this->spreadsheet->getActiveSheet()->duplicateStyle($this->tableStyle, $this->tableStart . ':' . $this->_table_maximum_column . $this->rowIndex);
        $this->nextRow();
    }

    public function startHorizontalMerge()
    {
        $this->_merge_start = $this->columnIndex . $this->rowIndex;
    }

    public function endHorizontalMerge()
    {
        $this->mergeCells($this->_merge_start . ':' . $this->lastColumnIndex . $this->rowIndex);
    }

    private function mergeCells($cell_range)
    {
        $this->spreadsheet->getActiveSheet()->mergeCells($cell_range);
    }

    private function nextColumn()
    {
        $this->lastColumnIndex = $this->columnIndex;
        if ($this->columnIndex > $this->_table_maximum_column)
            $this->_table_maximum_column = $this->columnIndex;
        $this->columnIndex++;
    }

    public function nextRow()
    {
        $this->rowIndex++;
        $this->columnIndex = 'A';
    }

    public function writeCell($content, $border_style = NULL, $font_style = NULL, $paragraph_style = NULL)
    {
        $this->spreadsheet->getActiveSheet()->setCellValue($this->columnIndex . $this->rowIndex, $content);
        $this->nextColumn();

        $this->_sheet_used = TRUE;
    }

    public function save()
    {
        $this->spreadsheet->setActiveSheetIndex(0);
        $filename = tempnam('', '');
        $writer = new Xlsx($this->spreadsheet);
        $writer->save($filename);
        return $filename;
    }

    public function getStyleHeader1()
    {
        $style = new Style();
        $font = new Font();
        $style->setFont($font);
        $font->setSize(15);
        $font->setBold(TRUE);
        $font->setColor(new Color(Color::COLOR_DARKBLUE));
        return $style;
    }

    public function getStyleHeader2()
    {
        $style = new Style();
        $font = new Font();
        $style->setFont($font);
        $font->setSize(13);
        $font->setBold(TRUE);
        $font->setColor(new Color(Color::COLOR_DARKBLUE));
        return $style;
    }

    public function getStyleHeader3()
    {
        $style = new Style();
        $font = new Font();
        $style->setFont($font);
        $font->setSize(11);
        $font->setBold(TRUE);
        $font->setColor(new Color(Color::COLOR_DARKBLUE));
        return $style;
    }

    public function getStyleTable()
    {
        $style = new Style();
        $style->getBorders()->applyFromArray(array('allborders'=>array('style'=>Border::BORDER_THIN)));
        return $style;
    }

    public function getStyleHeaderCellBorder()
    {
        return NULL;
    }

    public function getStyleHeaderCellFont()
    {
        return NULL;
    }

    public function getStyleHeaderCellParagraph()
    {
        return NULL;
    }

    public function getStyleCellBorder()
    {
        return NULL;
    }

    public function getStyleCellFont()
    {
        return NULL;
    }

    public function getStyleCellParagraph()
    {
        return NULL;
    }

    public function getStyleFooter()
    {
        return NULL;
    }
    
    private function setAutoSizeColumns()
    {
        $column = 'A';
        while ($column < 'Z')
            $this->setAutoSizeColumn($column++);
    }

    private function setAutoSizeColumn($column)
    {
        $this->spreadsheet->getActiveSheet()->getColumnDimension($column)->setAutoSize(TRUE);
    }

}