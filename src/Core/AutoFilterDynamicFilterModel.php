<?php

namespace Aspose\Cells_FOSS\Core;

final class AutoFilterDynamicFilterModel
{
    public $Enabled = false;
    public $Type = '';
    public $Value = null;
    public $MaxValue = null;

    // Clears the current state.
    public function Clear()
    {
        $this->Enabled = false;
        $this->Type = '';
        $this->Value = null;
        $this->MaxValue = null;
    }
}
