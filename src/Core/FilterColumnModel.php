<?php

namespace Aspose\Cells_FOSS\Core;

final class FilterColumnModel
{
    public $ColumnIndex = 0;
    public $HiddenButton = false;
    public $Filters;
    public $CustomFilters;
    public $CustomFiltersAnd = false;
    public $ColorFilter;
    public $DynamicFilter;
    public $Top10;

    public function __construct()
    {
        $this->Filters = array();
        $this->CustomFilters = array();
        $this->ColorFilter = new AutoFilterColorFilterModel();
        $this->DynamicFilter = new AutoFilterDynamicFilterModel();
        $this->Top10 = new AutoFilterTop10Model();
    }

    // Performs clear criteria.
    public function ClearCriteria()
    {
        $this->HiddenButton = false;
        $this->Filters = array();
        $this->CustomFilters = array();
        $this->CustomFiltersAnd = false;
        $this->ColorFilter->Clear();
        $this->DynamicFilter->Clear();
        $this->Top10->Clear();
    }

    // Performs has stored state.
    public function HasStoredState()
    {
        return $this->HiddenButton
            || count($this->Filters) > 0
            || count($this->CustomFilters) > 0
            || $this->ColorFilter->Enabled
            || $this->DynamicFilter->Enabled
            || $this->Top10->Enabled;
    }
}
