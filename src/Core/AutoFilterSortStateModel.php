<?php

namespace Aspose\Cells_FOSS\Core;

final class AutoFilterSortStateModel
{
    public $ColumnSort = false;
    public $CaseSensitive = false;
    public $SortMethod;
    public $Ref;
    public $Conditions;

    public function __construct()
    {
        $this->Ref = '';
        $this->SortMethod = '';
        $this->Conditions = array();
    }

    // Clears the current state.
    public function Clear()
    {
        $this->ColumnSort = false;
        $this->CaseSensitive = false;
        $this->SortMethod = '';
        $this->Ref = '';
        $this->Conditions = array();
    }

    // Performs has stored state.
    public function HasStoredState()
    {
        return !($this->Ref === null || $this->Ref === '')
            || $this->ColumnSort
            || $this->CaseSensitive
            || !($this->SortMethod === null || $this->SortMethod === '')
            || count($this->Conditions) > 0;
    }
}
