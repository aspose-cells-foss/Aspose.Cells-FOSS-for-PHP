<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\AutoFilterModel;

final class AutoFilter
{
    private $model;
    public $FilterColumns;
    public $SortState;

    public function __construct(AutoFilterModel $model)
    {
        $this->model = $model;
        $this->FilterColumns = new FilterColumnCollection($model->FilterColumns);
        $this->SortState = new AutoFilterSortState($model->SortState);
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Range') {
            return $this->model->Range;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Range') {
            $this->model->Range = AutoFilterSupport::NormalizeOptionalRange($value, 'Range');
            return;
        }
    }

    public function Clear()
    {
        $this->model->Clear();
    }
}
