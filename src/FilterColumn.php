<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\FilterColumnModel;

final class FilterColumn
{
    private $model;
    public $Filters;
    public $CustomFilters;
    public $ColorFilter;
    public $DynamicFilter;
    public $Top10;

    public function __construct(FilterColumnModel $model)
    {
        $this->model = $model;
        $this->Filters = new FilterValueCollection($model->Filters);
        $this->CustomFilters = new AutoFilterCustomFilterCollection($model->CustomFilters, $model);
        $this->ColorFilter = new AutoFilterColorFilter($model->ColorFilter);
        $this->DynamicFilter = new AutoFilterDynamicFilter($model->DynamicFilter);
        $this->Top10 = new AutoFilterTop10($model->Top10);
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'ColumnIndex') {
            return $this->model->ColumnIndex;
        }
        if ($propertyName === 'HiddenButton') {
            return $this->model->HiddenButton;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'HiddenButton') {
            $this->model->HiddenButton = $value;
            return;
        }
    }

    public function Clear()
    {
        $this->model->ClearCriteria();
    }
}
