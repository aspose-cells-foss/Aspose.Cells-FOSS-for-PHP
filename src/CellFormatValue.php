<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\AlignmentValue;
use Aspose\Cells_FOSS\Core\ProtectionValue;

final class CellFormatValue
{
    public $NumFmtId = 0;
    public $FontId = 0;
    public $FillId = 0;
    public $BorderId = 0;
    public $Alignment;
    public $Protection;

    public function __construct()
    {
        $this->Alignment = new AlignmentValue();
        $this->Protection = new ProtectionValue();
    }
}