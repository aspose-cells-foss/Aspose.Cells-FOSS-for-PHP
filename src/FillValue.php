<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ColorValue;
use Aspose\Cells_FOSS\Core\FillPatternKind;

final class FillValue
{
    public $Pattern;
    public $ForegroundColor;
    public $BackgroundColor;

    public function __construct()
    {
        $this->Pattern = FillPatternKind::None;
        $this->ForegroundColor = new ColorValue();
        $this->BackgroundColor = new ColorValue();
    }
}
