<?php
namespace Uk\Me\AndyCarpenter\Writer;


abstract class Writer
{
    const ORIENTATION_LANDSCAPE   = 'landscape';
    const ORIENTATION_PORTRAIT    = 'portrait';

    abstract public function setDocumentTitle($title);
    
    abstract public function setDocumentSubject($subject);
    
    abstract public function setDocumentDescription($description);
    
    abstract public function setDocumentCreator($creator);
    
    abstract public function setDocumentLastModifiedBy($modifiedBy);
    
    abstract public function setDocumentCompany($company);

    abstract public function getFileExtension();

    abstract public function getContentType();
    
    abstract public function setTitle($title);

    abstract public function addSection($title = NULL, $title_level = 1);

    abstract public function setSectionTitle($title, $title_level = 1);

    abstract public function setSectionOrientation($orientation);

    abstract public function addPageBreak();

    abstract public function writeTitle($title, $level, $style);

    abstract public function writeText($text);

    abstract public function startTable($style, $first_row_style = NULL);

    abstract public function endTable();

    abstract public function startHorizontalMerge();

    abstract public function endHorizontalMerge();

    abstract public function nextRow();

    abstract public function writeCell($content, $border_style = NULL, $font_style = NULL, $paragraph_style = NULL);

    abstract public function save();

    abstract public function getStyleHeader1();

    abstract public function getStyleHeader2();

    abstract public function getStyleHeader3();

    abstract public function getStyleTable();

    abstract public function getStyleHeaderCellBorder();

    abstract public function getStyleHeaderCellFont();

    abstract public function getStyleHeaderCellParagraph();

    abstract public function getStyleCellBorder();

    abstract public function getStyleCellFont();

    abstract public function getStyleCellParagraph();

    abstract public function getStyleFooter();
}