<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKColor;

final class DataBar
{
    public $Fraction;   // 0..1 of the cell width
    public $Color;
    public $ShowBorder;
    public $Gradient;

    public function __construct()
    {
        $this->Fraction = 0.0;
        $this->Color = new SKColor(0, 0, 0, 0);
        $this->ShowBorder = false;
        $this->Gradient = false;
    }
}
