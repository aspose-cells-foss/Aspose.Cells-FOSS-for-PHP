<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKColor;

final class Icon
{
    public $Color;
    public $BorderColor;
    public $HasBorder;
    public $AngleDeg;    // arrow direction; 90 = up, -90 = down, 0 = right
    public $IsArrow;      // false => draw a filled circle (fallback for non-arrow sets)

    public function __construct()
    {
        $this->Color = new SKColor(0, 0, 0, 0);
        $this->BorderColor = new SKColor(0, 0, 0, 0);
        $this->HasBorder = false;
        $this->AngleDeg = 0.0;
        $this->IsArrow = false;
    }
}
