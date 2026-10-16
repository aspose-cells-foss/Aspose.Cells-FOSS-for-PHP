<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ConditionalFormattingModel;
use Aspose\Cells_FOSS\Core\FormatConditionModel;
use Aspose\Cells_FOSS\Core\StyleValue;

// Represents format condition.
final class FormatCondition
{
    private $_owner;
    private $_collection;
    private $_model;

    public function __construct(&$owner, ConditionalFormattingModel $collection, FormatConditionModel $model)
    {
        $this->_owner =& $owner;
        $this->_collection = $collection;
        $this->_model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Type') {
            return $this->_model->Type;
        }
        if ($propertyName === 'Operator') {
            return $this->_model->Operator;
        }
        if ($propertyName === 'Formula1') {
            return $this->_model->Formula1 ?? '';
        }
        if ($propertyName === 'Formula2') {
            return $this->_model->Formula2 ?? '';
        }
        if ($propertyName === 'Formula') {
            return $this->_model->Formula1 ?? '';
        }
        if ($propertyName === 'TimePeriod') {
            return $this->_model->TimePeriod ?? '';
        }
        if ($propertyName === 'Duplicate') {
            return $this->_model->Duplicate;
        }
        if ($propertyName === 'Top') {
            return $this->_model->Top;
        }
        if ($propertyName === 'Percent') {
            return $this->_model->Percent;
        }
        if ($propertyName === 'Rank') {
            return $this->_model->Rank;
        }
        if ($propertyName === 'Above') {
            return $this->_model->Above;
        }
        if ($propertyName === 'StandardDeviation') {
            return $this->_model->StandardDeviation;
        }
        if ($propertyName === 'ColorScaleCount') {
            return $this->_model->ColorScaleCount;
        }
        if ($propertyName === 'MinColor') {
            return Color::FromCore($this->_model->MinColor);
        }
        if ($propertyName === 'MidColor') {
            return Color::FromCore($this->_model->MidColor);
        }
        if ($propertyName === 'MaxColor') {
            return Color::FromCore($this->_model->MaxColor);
        }
        if ($propertyName === 'BarColor') {
            return Color::FromCore($this->_model->BarColor);
        }
        if ($propertyName === 'NegativeBarColor') {
            return Color::FromCore($this->_model->NegativeBarColor);
        }
        if ($propertyName === 'ShowBorder') {
            return $this->_model->ShowBorder;
        }
        if ($propertyName === 'Direction') {
            return $this->_model->Direction ?? '';
        }
        if ($propertyName === 'BarLength') {
            return $this->_model->BarLength ?? '';
        }
        if ($propertyName === 'IconSetType') {
            return $this->_model->IconSetType ?? '';
        }
        if ($propertyName === 'ReverseIcons') {
            return $this->_model->ReverseIcons;
        }
        if ($propertyName === 'ShowIconOnly') {
            return $this->_model->ShowIconOnly;
        }
        if ($propertyName === 'Priority') {
            return $this->_model->Priority;
        }
        if ($propertyName === 'StopIfTrue') {
            return $this->_model->StopIfTrue;
        }
        if ($propertyName === 'Style') {
            return Style::FromCore($this->_model->Style)->Clone();
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Type') {
            $this->_model->Type = $value;
            return;
        }
        if ($propertyName === 'Operator') {
            $this->_model->Operator = $value;
            return;
        }
        if ($propertyName === 'Formula1') {
            $this->_model->Formula1 = self::NormalizeFormula($value);
            return;
        }
        if ($propertyName === 'Formula2') {
            $this->_model->Formula2 = self::NormalizeFormula($value);
            return;
        }
        if ($propertyName === 'Formula') {
            $this->_model->Formula1 = self::NormalizeFormula($value);
            return;
        }
        if ($propertyName === 'TimePeriod') {
            $this->_model->TimePeriod = self::NormalizeText($value);
            return;
        }
        if ($propertyName === 'Duplicate') {
            $this->_model->Duplicate = $value;
            return;
        }
        if ($propertyName === 'Top') {
            $this->_model->Top = $value;
            return;
        }
        if ($propertyName === 'Percent') {
            $this->_model->Percent = $value;
            return;
        }
        if ($propertyName === 'Rank') {
            if ($value < 0) {
                throw new CellsException('Conditional formatting rank must be zero or greater.');
            }

            $this->_model->Rank = $value;
            return;
        }
        if ($propertyName === 'Above') {
            $this->_model->Above = $value;
            return;
        }
        if ($propertyName === 'StandardDeviation') {
            if ($value < 0) {
                throw new CellsException('Conditional formatting standard deviation must be zero or greater.');
            }

            $this->_model->StandardDeviation = $value;
            return;
        }
        if ($propertyName === 'ColorScaleCount') {
            if ($value != 2 && $value != 3) {
                throw new CellsException('ColorScaleCount must be 2 or 3.');
            }

            $this->_model->ColorScaleCount = $value;
            return;
        }
        if ($propertyName === 'MinColor') {
            $this->_model->MinColor = $value->ToCore();
            return;
        }
        if ($propertyName === 'MidColor') {
            $this->_model->MidColor = $value->ToCore();
            return;
        }
        if ($propertyName === 'MaxColor') {
            $this->_model->MaxColor = $value->ToCore();
            return;
        }
        if ($propertyName === 'BarColor') {
            $this->_model->BarColor = $value->ToCore();
            return;
        }
        if ($propertyName === 'NegativeBarColor') {
            $this->_model->NegativeBarColor = $value->ToCore();
            return;
        }
        if ($propertyName === 'ShowBorder') {
            $this->_model->ShowBorder = $value;
            return;
        }
        if ($propertyName === 'Direction') {
            $this->_model->Direction = self::NormalizeText($value);
            return;
        }
        if ($propertyName === 'BarLength') {
            $this->_model->BarLength = self::NormalizeText($value);
            return;
        }
        if ($propertyName === 'IconSetType') {
            $this->_model->IconSetType = self::NormalizeText($value);
            return;
        }
        if ($propertyName === 'ReverseIcons') {
            $this->_model->ReverseIcons = $value;
            return;
        }
        if ($propertyName === 'ShowIconOnly') {
            $this->_model->ShowIconOnly = $value;
            return;
        }
        if ($propertyName === 'Priority') {
            if ($value <= 0) {
                throw new CellsException('Conditional formatting priority must be greater than zero.');
            }

            $this->_model->Priority = $value;
            return;
        }
        if ($propertyName === 'StopIfTrue') {
            $this->_model->StopIfTrue = $value;
            return;
        }
        if ($propertyName === 'Style') {
            $this->_model->Style = $value === null ? StyleValue::Default()->Clone() : $value->ToCore();
            return;
        }
    }

    public function Remove()
    {
        (new FormatConditionCollection($this->_owner, $this->_collection))->RemoveCondition($this->_owner, $this->_collection, $this->_model);
    }

    private static function NormalizeFormula($value)
    {
        if ($value === null) {
            return null;
        }

        $trimmed = trim($value);
        if ($trimmed === '') {
            return null;
        }

        if ($trimmed[0] === '=') {
            return substr($trimmed, 1);
        }

        return $trimmed;
    }

    private static function NormalizeText($value)
    {
        if ($value === null) {
            return null;
        }

        $trimmed = trim($value);
        if ($trimmed === '') {
            return null;
        }

        return $trimmed;
    }
}