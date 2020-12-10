<?php
namespace Uk\Me\AndyCarpenter\Writer;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Style\Table;


class WordWriter extends Writer
{
    const FORMAT_DOC = 'Word2007';
    const FORMAT_HTML = 'HTML';
    const FORMAT_ODT = 'ODText';
    const FORMAT_PDF = 'PDF';
    const FORMAT_RTF = 'RTF';

    private $format;
    private $word;
    private $section;
    private $sectionUsed;
    private $table;
    private $_add_table_row;

    public function __construct($format = self::FORMAT_DOC)
    {
        $this->format = $format;
        $this->word = new PhpWord();

        $this->addStyles();

        $this->addSection();
    }

    public function setDocumentTitle($title)
    {
        $this->word->getDocInfo()->setTitle($title);
    }

    public function setDocumentSubject($subject)
    {
        $this->word->getDocInfo()->setTitle($subject);
    }

    public function setDocumentDescription($description)
    {
        $this->word->getDocInfo()->setDescription($description);
    }

    public function setDocumentCreator($creator)
    {
        $this->word->getDocInfo()->setCreator($creator);
    }

    public function setDocumentLastModifiedBy($modifiedBy)
    {
        $this->word->getDocInfo()->setLastModifiedBy($modifiedBy);
    }

    public function setDocumentCompany($company)
    {
        $this->word->getDocInfo()->setCompany($company);
    }
 
    public function getFileExtension()
    {
    	if ($this->format == self::FORMAT_DOC)
    		return '.docx';
        else if ($this->format == self::FORMAT_HTML)
    		return '.html';
    	else if ($this->format == self::FORMAT_ODT)
    		return '.odt';
        else if ($this->format == self::FORMAT_PDF)
    		return '.pdf';
    	else if ($this->format == SELF::FORMAT_RTF)
    		return '.rtf';
    	else
    		return '';
    }
    
    public function getContentType()
    {
    	if ($this->format == self::FORMAT_DOC)
    	    return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    	else if ($this->format == self::FORMAT_HTML)
    		return 'text/html';
    	else if ($this->format == self::FORMAT_ODT)
    		return 'application/vnd.oasis.opendocument.text';
    	else if ($this->format == self::FORMAT_PDF)
    		return 'application/pdf';
    	else if ($this->format == SELF::FORMAT_RTF)
    		return 'application/rtf';
    	else
    		return null;
    }

    public function setTitle($title)
    {
        $this->writeText($title, null, $this->getStyleTitle());
    }

    public function addSection($title = NULL, $titleLevel = 1)
    {
        if (empty($this->section) || $this->sectionUsed) {
            $this->section = $this->word->addSection();
        }
        $this->sectionUsed = false;
        $this->setSectionTitle($title, $titleLevel);

        $footer = $this->section->addFooter()->addTextRun($this->getStyleFooter());
        $footer->addField('DATE', array('dateformat' => 'd MMMM yyyy'));
    }

    public function setSectionTitle($title, $titleLevel = 1)
    {
        if ($title != null) {
            $this->writeTitle($title, $titleLevel);
            $this->sectionUsed = true;
        }
    }

    public function setSectionOrientation($orientation)
    {
    }

    public function addPageBreak(){
        $this->section->addPageBreak();
    }

    public function writeTitle($title, $level, $style = null)
    {
        $this->section->addTitle($title, $level);
        $this->sectionUsed = true;
    }

    public function writeText($text, $fontStyle = null,
        $paragraphStyle = null)
    {
        $this->section->addText($text, $fontStyle, $paragraphStyle);
        $this->sectionUsed = true;
    }

    public function startTable($style, $first_row_style = null)
    {
        $this->table = $this->section->addTable($style);
        $this->table->setWidth(100 * 50);
        $this->nextRow();
        $this->sectionUsed = true;
    }

    public function endTable()
    {
    }

    public function startHorizontalMerge()
    {
    }

    public function endHorizontalMerge()
    {
    }

    public function nextRow()
    {
        $this->_add_table_row = TRUE;
    }

    public function writeCell($content, $border_style = NULL, $font_style = NULL, $paragraph_style = NULL)
    {
        if ($this->_add_table_row)
            $this->table->addRow();
        $this->_add_table_row = FALSE;

        $this->table->addCell(null, $border_style)->addText($content, $font_style, $paragraph_style);
    }

    public function save()
    {
        $filename = tempnam('', '');
        $writer = IOFactory::createWriter($this->word, $this->format);
        $writer->save($filename);
        return $filename;
    }

    public function getStyleTitle()
    {
        return 'Title';
    }

    public function getStyleHeader1()
    {
        return 'Heading_1';
    }

    public function getStyleHeader2()
    {
        return 'Heading_2';
            }

    public function getStyleHeader3()
    {
        return 'Heading_3';
    }

    public function getStyleTable()
    {
        return 'table';
    }

    public function getStyleHeaderCellBorder()
    {
        return array(
            'borderBottomSize' => 2 * 8,
            'borderBottomColor' => '888888',
        );
    }

    public function getStyleHeaderCellFont()
    {
        return 'cell_header_font';
    }

    public function getStyleHeaderCellParagraph()
    {
        return 'table_head';
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
        return 'cell_paragraph';
    }

    public function getStyleFooter()
    {
        return 'footer';
    }

    private function addStyles()
    {
        $header_colour = '000080';
        $table_colour = '888888';

        $this->word->addTitleStyle(1, array(
            'name' => 'Ariel',
            'size' => 20,
            'bold' => TRUE,
            'color' => $header_colour
        ), array(
            'align' => 'left',
            'spaceAfter' => 6 * 20,
            'keepNext' => TRUE,
        ));

        $this->word->addTitleStyle(2, array(
            'name' => 'Ariel',
            'size' => 16,
            'bold' => TRUE,
            'color' => $header_colour
        ), array(
            'align' => 'left',
            'spaceAfter' => '0',
            'keepNext' => TRUE,
        ));

        $this->word->addTitleStyle(3, array(
            'name' => 'Ariel',
            'size' => 14,
            'bold' => TRUE,
            'color' => $header_colour
        ), array(
            'align' => 'left',
            'spaceBefore' => 12 * 20,
            'spaceAfter' => '0',
            'keepNext' => TRUE,
        ));

        $tableStyle = $this->word->addTableStyle($this->getStyleTable(), array(
            'unit' =>  Table::WIDTH_AUTO,
        ), array(
            'borderBottomSize' => 2 * 8,
        ));
        $tableStyle->setBorderSize(2 * 8);
        $tableStyle->setBorderInsideHSize(1 * 8);
        $tableStyle->setBorderInsideVSize(1 * 8);
        $tableStyle->setBorderColor($table_colour);

        $this->word->addFontStyle($this->getStyleHeaderCellFont(), array(
            'bold' => TRUE,
        ));

        $this->word->addParagraphStyle($this->getStyleHeaderCellParagraph(), array(
            'basedOn' => $this->getStyleCellParagraph(),
            'align' => 'center',
        ));

        $this->word->addParagraphStyle($this->getStyleCellParagraph(), array(
            'align' => 'left',
            'spaceAfter' => '0'
        ));

        $this->word->addParagraphStyle($this->getStyleFooter(), array(
            'align' => 'right',
            'spaceAfter' => '0'
        ));
    }

}