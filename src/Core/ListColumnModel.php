<?php

namespace Aspose\Cells_FOSS\Core;

final class ListColumnModel
{
    public $Id;
    public $Name;
    public $TotalsRowFunction;
    public $TotalsRowLabel;
    public $TotalsRowFormula;

    public function __construct($id, $name)
    {
        $this->Id = $id;
        $this->Name = $name;
        $this->TotalsRowFunction = 'none';
        $this->TotalsRowLabel = '';
        $this->TotalsRowFormula = '';
    }
}
