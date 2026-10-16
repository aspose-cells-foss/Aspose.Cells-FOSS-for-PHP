<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\StyleValue;

final class StylesheetLoadContext
{
    public $DefaultCellStyle;
    public $CellFormats;
    public $DifferentialFormats;
    public $DateStyleIndexes;

    public function __construct()
    {
        $this->DefaultCellStyle = StyleValue::Default()->Clone();
        $this->CellFormats = array(StyleValue::Default()->Clone());
        $this->DifferentialFormats = array();
        $this->DateStyleIndexes = array();
    }
}