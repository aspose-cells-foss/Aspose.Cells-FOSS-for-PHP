<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ChartModel;

// Represents a chart embedded in a worksheet.
final class Chart
{
    private $_model;

    public function __construct(ChartModel $model)
    {
        $this->_model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Name') {
            return $this->_model->Name;
        }
        if ($propertyName === 'ChartType') {
            return $this->_model->ChartType;
        }
        if ($propertyName === 'UpperLeftRow') {
            return $this->_model->UpperLeftRow;
        }
        if ($propertyName === 'UpperLeftColumn') {
            return $this->_model->UpperLeftColumn;
        }
        if ($propertyName === 'LowerRightRow') {
            return $this->_model->LowerRightRow;
        }
        if ($propertyName === 'LowerRightColumn') {
            return $this->_model->LowerRightColumn;
        }
        if ($propertyName === 'ExtentCx') {
            return $this->_model->ExtentCx;
        }
        if ($propertyName === 'ExtentCy') {
            return $this->_model->ExtentCy;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Name') {
            throw new \RuntimeException('Name is read-only.');
        }
        if ($propertyName === 'ChartType') {
            throw new \RuntimeException('ChartType is read-only.');
        }
        if ($propertyName === 'UpperLeftRow') {
            throw new \RuntimeException('UpperLeftRow is read-only.');
        }
        if ($propertyName === 'UpperLeftColumn') {
            throw new \RuntimeException('UpperLeftColumn is read-only.');
        }
        if ($propertyName === 'LowerRightRow') {
            throw new \RuntimeException('LowerRightRow is read-only.');
        }
        if ($propertyName === 'LowerRightColumn') {
            throw new \RuntimeException('LowerRightColumn is read-only.');
        }
        if ($propertyName === 'ExtentCx') {
            throw new \RuntimeException('ExtentCx is read-only.');
        }
        if ($propertyName === 'ExtentCy') {
            throw new \RuntimeException('ExtentCy is read-only.');
        }
    }
}