<?php

namespace Aspose\Cells_FOSS\Core;

final class AutoFilterTop10Model
{
    public $Enabled = false;
    public $Top = true;
    public $Percent = false;
    public $Value = null;
    public $FilterValue = null;

    // Clears the current state.
    public function Clear()
    {
        $this->Enabled = false;
        $this->Top = true;
        $this->Percent = false;
        $this->Value = null;
        $this->FilterValue = null;
    }
}
