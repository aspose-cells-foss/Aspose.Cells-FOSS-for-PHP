<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\AutoFilterSortStateModel;

final class AutoFilterSortState
{
    private $model;
    public $SortConditions;

    public function __construct(AutoFilterSortStateModel $model)
    {
        $this->model = $model;
        $this->SortConditions = new AutoFilterSortConditionCollection($model->Conditions);
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'ColumnSort') {
            return $this->model->ColumnSort;
        }
        if ($propertyName === 'CaseSensitive') {
            return $this->model->CaseSensitive;
        }
        if ($propertyName === 'SortMethod') {
            return $this->model->SortMethod;
        }
        if ($propertyName === 'Ref') {
            return $this->model->Ref;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'ColumnSort') {
            $this->model->ColumnSort = $value;
            return;
        }

        if ($propertyName === 'CaseSensitive') {
            $this->model->CaseSensitive = $value;
            return;
        }

        if ($propertyName === 'SortMethod') {
            $this->model->SortMethod = AutoFilterSupport::NormalizeOptionalText($value);
            return;
        }

        if ($propertyName === 'Ref') {
            $this->model->Ref = AutoFilterSupport::NormalizeOptionalRange($value, 'Ref');
            return;
        }
    }

    public function Clear()
    {
        $this->model->Clear();
    }
}
