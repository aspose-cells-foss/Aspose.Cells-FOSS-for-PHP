<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\AutoFilterCustomFilterModel;

final class AutoFilterCustomFilter
{
    private $model;

    public function __construct(AutoFilterCustomFilterModel $model)
    {
        $this->model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Operator') {
            return AutoFilterSupport::ParseOperatorOrDefault($this->model->Operator);
        }
        if ($propertyName === 'Value') {
            return $this->model->Value;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Operator') {
            $name = AutoFilterSupport::ToOperatorName($value);
            $this->model->Operator = $name === null ? '' : $name;
            return;
        }

        if ($propertyName === 'Value') {
            $this->model->Value = AutoFilterSupport::NormalizeText($value, 'Value');
            return;
        }
    }
}
