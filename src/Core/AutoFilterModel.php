<?php

namespace Aspose\Cells_FOSS\Core;

final class AutoFilterModel
{
    public $Range;
    public $FilterColumns;
    public $SortState;

    public function __construct()
    {
        $this->Range = '';
        $this->FilterColumns = array();
        $this->SortState = new AutoFilterSortStateModel();
    }

    // Clears the current state.
    public function Clear()
    {
        $this->Range = '';
        $this->FilterColumns = array();
        $this->SortState->Clear();
    }

    // Performs has stored state.
    public function HasStoredState()
    {
        return !($this->Range === null || $this->Range === '');
    }
}
