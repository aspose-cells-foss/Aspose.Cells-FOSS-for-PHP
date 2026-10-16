<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\AutoFilterColorFilterModel;

final class AutoFilterColorFilter
{
    private $model;

    public function __construct(AutoFilterColorFilterModel $model)
    {
        $this->model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Enabled') {
            return $this->model->Enabled;
        }
        if ($propertyName === 'DifferentialStyleId') {
            return $this->model->DifferentialStyleId;
        }
        if ($propertyName === 'CellColor') {
            return $this->model->CellColor;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Enabled') {
            if (!$value) {
                $this->model->Clear();
                return;
            }

            $this->model->Enabled = true;
            return;
        }

        if ($propertyName === 'DifferentialStyleId') {
            if ($value !== null && $value < 0) {
                throw new CellsException('Differential style id must be zero or greater.');
            }

            $this->model->DifferentialStyleId = $value;
            if ($value !== null) {
                $this->model->Enabled = true;
            }
            return;
        }

        if ($propertyName === 'CellColor') {
            $this->model->CellColor = $value;
            if ($value) {
                $this->model->Enabled = true;
            }
            return;
        }
    }

    public function Clear()
    {
        $this->model->Clear();
    }
}
