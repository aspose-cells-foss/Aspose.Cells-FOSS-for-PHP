<?php

namespace Aspose\Cells_FOSS\CompareOpenXml;

final class OpenXmlCellDescriptor
{
    public function __construct($sheetName, $cellReference, $cell)
    {
        $this->SheetName = $sheetName;
        $this->CellReference = $cellReference;
        $this->Cell = $cell;
    }

    public $SheetName;
    public $CellReference;
    public $Cell;
}
