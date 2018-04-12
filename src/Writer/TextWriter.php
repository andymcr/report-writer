<?php
namespace Uk\Me\AndyCarpenter\Writer;


class TextWriter extends Writer
{

    private $_filename;
    private $_file;
    private $_first_column = TRUE;

    public function __construct()
    {
        $this->_filename = tempnam('', '');
        $this->_file = fopen($this->_filename, 'w');
    }

    public function setDocumentTitle($title)
    {
    }

    public function setDocumentSubject($subject)
    {
    }

    public function setDocumentDescription($description)
    {
    }

    public function setDocumentCreator($creator)
    {
    }

    public function setDocumentLastModifiedBy($modifiedBy)
    {
    }

    public function setDocumentCompany($company)
    {
    }

    public function getFileExtension()
    {
    	return '.txt';
    }

    public function getContentType()
    {
        return 'txt';
    }

    public function setTitle($title)
    {
    }

    public function addSection($title = NULL, $title_level = 1)
    {
    }

    public function setSectionTitle($title, $title_level = 1)
    {
    }

    public function setSectionOrientation($orientation)
    {
    }

    public function addPageBreak()
    {
    }

    public function writeTitle($title, $level, $style = null)
    {
        $this->writeText($title);
        $this->newLine();
    }

    public function writeText($text)
    {
        fwrite($this->_file, str_replace('"', "'", $text));
    }

    public function startTable($style, $first_row_style = null)
    {
        $this->_first_column= TRUE;
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
        $this->newLine();
        $this->_first_column = TRUE;
    }

    private function newLine()
    {
        fwrite($this->_file, "\r\n");
    }

    public function writeCell($content, $border_style = NULL, $font_style = NULL, $paragraph_style = NULL)
    {
        if (!$this->_first_column)
            fwrite($this->_file, "\t");
        $this->_first_column = FALSE;

        $this->writeText($content);
    }

    public function save()
    {
        fclose($this->_file);
        return $this->_filename;
    }

    public function getStyleHeader1()
    {
        return NULL;
    }

    public function getStyleHeader2()
    {
        return NULL;
    }

    public function getStyleHeader3()
    {
        return NULL;
    }

    public function getStyleTable($style)
    {
        return NULL;
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
    
    }