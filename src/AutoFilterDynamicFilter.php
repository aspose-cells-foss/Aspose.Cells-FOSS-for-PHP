<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\AutoFilterDynamicFilterModel;

final class AutoFilterDynamicFilter
{
    private $model;

    public function __construct(AutoFilterDynamicFilterModel $model)
    {
        $this->model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Enabled') {
            return $this->model->Enabled;
        }
        if ($propertyName === 'Type') {
            return $this->model->Type;
        }
        if ($propertyName === 'Value') {
            return $this->model->Value;
        }
        if ($propertyName === 'MaxValue') {
            return $this->model->MaxValue;
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

        if ($propertyName === 'Type') {
            $this->model->Type = AutoFilterSupport::NormalizeOptionalText($value);
            if (strlen($this->model->Type) > 0) {
                $this->model->Enabled = true;
            }
            return;
        }

        if ($propertyName === 'Value') {
            $this->model->Value = $value;
            if ($value !== null) {
                $this->model->Enabled = true;
            }
            return;
        }

        if ($propertyName === 'MaxValue') {
            $this->model->MaxValue = $value;
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
