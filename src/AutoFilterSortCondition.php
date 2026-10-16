<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\AutoFilterSortConditionModel;

final class AutoFilterSortCondition
{
    private $model;

    public function __construct(AutoFilterSortConditionModel $model)
    {
        $this->model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Ref') {
            return $this->model->Ref;
        }
        if ($propertyName === 'Descending') {
            return $this->model->Descending;
        }
        if ($propertyName === 'SortBy') {
            return $this->model->SortBy;
        }
        if ($propertyName === 'CustomList') {
            return $this->model->CustomList;
        }
        if ($propertyName === 'DifferentialStyleId') {
            return $this->model->DifferentialStyleId;
        }
        if ($propertyName === 'IconSet') {
            return $this->model->IconSet;
        }
        if ($propertyName === 'IconId') {
            return $this->model->IconId;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Ref') {
            $this->model->Ref = AutoFilterSupport::NormalizeRequiredRange($value, 'Ref');
            return;
        }

        if ($propertyName === 'Descending') {
            $this->model->Descending = $value;
            return;
        }

        if ($propertyName === 'SortBy') {
            $this->model->SortBy = AutoFilterSupport::NormalizeOptionalText($value);
            return;
        }

        if ($propertyName === 'CustomList') {
            $this->model->CustomList = AutoFilterSupport::NormalizeOptionalText($value);
            return;
        }

        if ($propertyName === 'DifferentialStyleId') {
            if ($value !== null && $value < 0) {
                throw new CellsException('Differential style id must be zero or greater.');
            }

            $this->model->DifferentialStyleId = $value;
            return;
        }

        if ($propertyName === 'IconSet') {
            $this->model->IconSet = AutoFilterSupport::NormalizeOptionalText($value);
            return;
        }

        if ($propertyName === 'IconId') {
            if ($value !== null && $value < 0) {
                throw new CellsException('Icon id must be zero or greater.');
            }

            $this->model->IconId = $value;
            return;
        }
    }
}
