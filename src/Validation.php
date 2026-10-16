<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ValidationModel;

final class Validation
{
    private $owner;
    private $model;

    public function __construct(&$owner, ValidationModel $model)
    {
        $this->owner = &$owner;
        $this->model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Areas') {
            $areas = array();
            for ($index = 0; $index < count($this->model->Areas); $index++) {
                $areas[] = $this->model->Areas[$index];
            }

            return $areas;
        }

        if ($propertyName === 'Type') {
            return $this->model->Type;
        }
        if ($propertyName === 'AlertStyle') {
            return $this->model->AlertStyle;
        }
        if ($propertyName === 'Operator') {
            return $this->model->Operator;
        }
        if ($propertyName === 'Formula1') {
            return $this->model->Formula1 === null ? '' : $this->model->Formula1;
        }
        if ($propertyName === 'Formula2') {
            return $this->model->Formula2 === null ? '' : $this->model->Formula2;
        }
        if ($propertyName === 'IgnoreBlank') {
            return $this->model->IgnoreBlank;
        }
        if ($propertyName === 'InCellDropDown') {
            return $this->model->InCellDropDown;
        }
        if ($propertyName === 'InputTitle') {
            return $this->model->InputTitle === null ? '' : $this->model->InputTitle;
        }
        if ($propertyName === 'InputMessage') {
            return $this->model->InputMessage === null ? '' : $this->model->InputMessage;
        }
        if ($propertyName === 'ErrorTitle') {
            return $this->model->ErrorTitle === null ? '' : $this->model->ErrorTitle;
        }
        if ($propertyName === 'ErrorMessage') {
            return $this->model->ErrorMessage === null ? '' : $this->model->ErrorMessage;
        }
        if ($propertyName === 'ShowInput') {
            return $this->model->ShowInput;
        }
        if ($propertyName === 'ShowError') {
            return $this->model->ShowError;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Type') {
            $this->model->Type = $value;
            return;
        }

        if ($propertyName === 'AlertStyle') {
            $this->model->AlertStyle = $value;
            return;
        }

        if ($propertyName === 'Operator') {
            $this->model->Operator = $value;
            return;
        }

        if ($propertyName === 'Formula1') {
            $this->model->Formula1 = self::NormalizeFormula($value);
            return;
        }

        if ($propertyName === 'Formula2') {
            $this->model->Formula2 = self::NormalizeFormula($value);
            return;
        }

        if ($propertyName === 'IgnoreBlank') {
            $this->model->IgnoreBlank = $value;
            return;
        }

        if ($propertyName === 'InCellDropDown') {
            $this->model->InCellDropDown = $value;
            return;
        }

        if ($propertyName === 'InputTitle') {
            $this->model->InputTitle = self::NormalizeText($value);
            return;
        }

        if ($propertyName === 'InputMessage') {
            $this->model->InputMessage = self::NormalizeText($value);
            return;
        }

        if ($propertyName === 'ErrorTitle') {
            $this->model->ErrorTitle = self::NormalizeText($value);
            return;
        }

        if ($propertyName === 'ErrorMessage') {
            $this->model->ErrorMessage = self::NormalizeText($value);
            return;
        }

        if ($propertyName === 'ShowInput') {
            $this->model->ShowInput = $value;
            return;
        }

        if ($propertyName === 'ShowError') {
            $this->model->ShowError = $value;
            return;
        }
    }

    public function AddArea(CellArea $area)
    {
        ValidationCollection::AddAreaToValidation($this->owner, $this->model, $area);
    }

    public function RemoveArea(CellArea $area)
    {
        ValidationCollection::RemoveAreaFromValidation($this->owner, $this->model, $area);
    }

    private static function NormalizeFormula($value)
    {
        if ($value === null) {
            return null;
        }

        $trimmed = trim($value);
        if (strlen($trimmed) === 0) {
            return null;
        }

        if ($trimmed[0] === '=') {
            return substr($trimmed, 1);
        }

        return $trimmed;
    }

    private static function NormalizeText($value)
    {
        if ($value === null || $value === '') {
            return null;
        }

        return $value;
    }
}
