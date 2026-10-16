<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\AutoFilterTop10Model;

final class AutoFilterTop10
{
    private $model;

    public function __construct(AutoFilterTop10Model $model)
    {
        $this->model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Enabled') {
            return $this->model->Enabled;
        }
        if ($propertyName === 'Top') {
            return $this->model->Top;
        }
        if ($propertyName === 'Percent') {
            return $this->model->Percent;
        }
        if ($propertyName === 'Value') {
            return $this->model->Value;
        }
        if ($propertyName === 'FilterValue') {
            return $this->model->FilterValue;
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

        if ($propertyName === 'Top') {
            $this->model->Top = $value;
            $this->model->Enabled = true;
            return;
        }

        if ($propertyName === 'Percent') {
            $this->model->Percent = $value;
            $this->model->Enabled = true;
            return;
        }

        if ($propertyName === 'Value') {
            $this->model->Value = $value;
            if ($value !== null) {
                $this->model->Enabled = true;
            }
            return;
        }

        if ($propertyName === 'FilterValue') {
            $this->model->FilterValue = $value;
            if ($value !== null) {
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
