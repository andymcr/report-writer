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

    public function setTitle($title)
    {
    }

    public function addSection($title = null, $title_level = 1)
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

        $this->_sheet_used = false;
    }

    public function setSectionTitle($title, $title_level = 1)
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

        $this->_sheet_used = true;
    }

    public function writeText($text)
    {
        $this->spreadsheet->getActiveSheet()->setCellValue($this->columnIndex . $this->rowIndex, $text);
        $this->nextColumn();

        $this->_sheet_used = true;
    }

    public function startTable($style, $first_row_style = null)
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
        if ($this->_sheet_used) {
            $this->rowIndex++;
        }
        $this->columnIndex = 'A';
    }

    public function writeCell($content, $border_style = null, $font_style = null, $paragraph_style = null)
    {
        $cellAddress = $this->columnIndex . $this->rowIndex;

        $this->spreadsheet->getActiveSheet()->setCellValue($cellAddress, $content);

        if (!is_null($border_style)) {
            $this->spreadsheet->getActiveSheet()->getCell($cellAddress)->getStyle();
        }
        if (!is_null($font_style)) {
            $this->spreadsheet->getActiveSheet()->getCell($cellAddress)->getStyle()
                ->getFont()
                    ->setName($font_style->getFont()->getName())
                    ->setSize($font_style->getFont()->getSize());
        }
        if (!is_null($paragraph_style)) {
            $this->spreadsheet->getActiveSheet()->getCell($cellAddress)->getStyle();
        }
        $this->nextColumn();

        $this->_sheet_used = true;
    }

    public function save()
    {
        $filename = tempnam('', '');
        $this->saveNamed($filename);
        return $filename;
    }

    public function saveNamed($filename)
    {
        $this->spreadsheet->setActiveSheetIndex(0);
        $writer = new Xlsx($this->spreadsheet);
        $writer->save($filename);
    }

    public function getStyleHeader1()
    {
        $style = new Style();
        $font = new Font();
        $style->setFont($font);
        $font->setSize(15);
        $font->setBold(true);
        $font->setColor(new Color(Color::COLOR_DARKBLUE));
        return $style;
    }

    public function getStyleHeader2()
    {
        $style = new Style();
        $font = new Font();
        $style->setFont($font);
        $font->setSize(13);
        $font->setBold(true);
        $font->setColor(new Color(Color::COLOR_DARKBLUE));
        return $style;
    }

    public function getStyleHeader3()
    {
        $style = new Style();
        $font = new Font();
        $style->setFont($font);
        $font->setSize(11);
        $font->setBold(true);
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
        return null;
    }

    public function getStyleHeaderCellFont()
    {
        return null;
    }

    public function getStyleHeaderCellParagraph()
    {
        return null;
    }

    public function getStyleCellBorder()
    {
        return null;
    }

    public function getStyleCellFont()
    {
        return null;
    }

    public function getStyleCellParagraph()
    {
        return null;
    }

    public function getStyleFooter()
    {
        return null;
    }
    
    private function setAutoSizeColumns()
    {
        $column = 'A';
        while ($column < 'Z')
            $this->setAutoSizeColumn($column++);
    }

    private function setAutoSizeColumn($column)
    {
        $this->spreadsheet->getActiveSheet()->getColumnDimension($column)->setAutoSize(true);
    }

}