<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ListColumnModel;

// Represents a single column in an Excel table.
final class ListColumn
{
    private $_model;

    public function __construct(ListColumnModel $model)
    {
        $this->_model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Name') {
            return $this->_model->Name;
        }
        if ($propertyName === 'TotalsCalculation') {
            return self::TotalsCalculationFromString($this->_model->TotalsRowFunction);
        }
        if ($propertyName === 'TotalsRowLabel') {
            return $this->_model->TotalsRowLabel;
        }
        if ($propertyName === 'TotalsRowFormula') {
            return $this->_model->TotalsRowFormula;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Name') {
            if ($value === null || $value === '') {
                throw new CellsException('ListColumn name must be non-empty.');
            }

            $this->_model->Name = $value;
            return;
        }
        if ($propertyName === 'TotalsCalculation') {
            $this->_model->TotalsRowFunction = self::TotalsCalculationToString($value);
            return;
        }
        if ($propertyName === 'TotalsRowLabel') {
            $this->_model->TotalsRowLabel = $value ?? '';
            return;
        }
        if ($propertyName === 'TotalsRowFormula') {
            $this->_model->TotalsRowFormula = $value ?? '';
            return;
        }
    }

    public static function TotalsCalculationFromString($value)
    {
        switch ($value) {
            case 'sum': return TotalsCalculation::Sum;
            case 'count': return TotalsCalculation::Count;
            case 'average': return TotalsCalculation::Average;
            case 'max': return TotalsCalculation::Max;
            case 'min': return TotalsCalculation::Min;
            case 'stdDev': return TotalsCalculation::StdDev;
            case 'var': return TotalsCalculation::Var;
            case 'countNums': return TotalsCalculation::CountNums;
            case 'custom': return TotalsCalculation::Custom;
            default: return TotalsCalculation::None;
        }
    }

    public static function TotalsCalculationToString($value)
    {
        switch ($value) {
            case TotalsCalculation::Sum: return 'sum';
            case TotalsCalculation::Count: return 'count';
            case TotalsCalculation::Average: return 'average';
            case TotalsCalculation::Max: return 'max';
            case TotalsCalculation::Min: return 'min';
            case TotalsCalculation::StdDev: return 'stdDev';
            case TotalsCalculation::Var: return 'var';
            case TotalsCalculation::CountNums: return 'countNums';
            case TotalsCalculation::Custom: return 'custom';
            default: return 'none';
        }
    }
}