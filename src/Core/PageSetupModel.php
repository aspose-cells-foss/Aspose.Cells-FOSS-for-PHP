<?php

namespace Aspose\Cells_FOSS\Core;

final class PageSetupModel
{
    public $Margins;
    public $PrintOptions;
    public $HeaderFooter;
    public $PaperSize = 0;
    public $Orientation = 0;
    public $FirstPageNumber = null;
    public $Scale = null;
    public $FitToWidth = null;
    public $FitToHeight = null;
    public $PrintArea;
    public $PrintTitleRows;
    public $PrintTitleColumns;
    public $HorizontalPageBreaks;
    public $VerticalPageBreaks;

    public function __construct()
    {
        $this->Margins = new PageMarginsModel();
        $this->PrintOptions = new PrintOptionsModel();
        $this->HeaderFooter = new HeaderFooterModel();
        $this->HorizontalPageBreaks = array();
        $this->VerticalPageBreaks = array();
    }
}
