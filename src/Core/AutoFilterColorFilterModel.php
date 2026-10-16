<?php

namespace Aspose\Cells_FOSS\Core;

final class AutoFilterColorFilterModel
{
    public $Enabled = false;
    public $DifferentialStyleId = null;
    public $CellColor = false;

    // Clears the current state.
    public function Clear()
    {
        $this->Enabled = false;
        $this->DifferentialStyleId = null;
        $this->CellColor = false;
    }
}
