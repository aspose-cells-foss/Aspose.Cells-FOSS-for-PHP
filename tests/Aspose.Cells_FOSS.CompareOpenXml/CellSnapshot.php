<?php

namespace Aspose\Cells_FOSS\CompareOpenXml;

final class CellSnapshot
{
    public function __construct($sheetName, $cellReference, $cellType, $value, $formula)
    {
        $this->SheetName = $sheetName;
        $this->CellReference = $cellReference;
        $this->CellType = $cellType;
        $this->Value = $value;
        $this->Formula = $formula;
    }

    public $SheetName;
    public $CellReference;
    public $CellType;
    public $Value;
    public $Formula;

    public function ToString()
    {
        return 'Sheet=' . $this->SheetName
            . '; Cell=' . $this->CellReference
            . '; Type=' . $this->CellType
            . '; Formula=' . $this->Formula
            . '; Value=' . $this->Value;
    }

    public function __toString()
    {
        return $this->ToString();
    }
}
