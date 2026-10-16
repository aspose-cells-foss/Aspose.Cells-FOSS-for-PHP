<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\WorksheetProtectionModel;

final class WorksheetProtection
{
    private $_model;

    public function __construct(WorksheetProtectionModel $model)
    {
        $this->_model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'IsProtected') {
            return $this->_model->IsProtected;
        }
        if ($propertyName === 'Objects') {
            return $this->_model->Objects;
        }
        if ($propertyName === 'Scenarios') {
            return $this->_model->Scenarios;
        }
        if ($propertyName === 'FormatCells') {
            return $this->_model->FormatCells;
        }
        if ($propertyName === 'FormatColumns') {
            return $this->_model->FormatColumns;
        }
        if ($propertyName === 'FormatRows') {
            return $this->_model->FormatRows;
        }
        if ($propertyName === 'InsertColumns') {
            return $this->_model->InsertColumns;
        }
        if ($propertyName === 'InsertRows') {
            return $this->_model->InsertRows;
        }
        if ($propertyName === 'InsertHyperlinks') {
            return $this->_model->InsertHyperlinks;
        }
        if ($propertyName === 'DeleteColumns') {
            return $this->_model->DeleteColumns;
        }
        if ($propertyName === 'DeleteRows') {
            return $this->_model->DeleteRows;
        }
        if ($propertyName === 'SelectLockedCells') {
            return $this->_model->SelectLockedCells;
        }
        if ($propertyName === 'Sort') {
            return $this->_model->Sort;
        }
        if ($propertyName === 'AutoFilter') {
            return $this->_model->AutoFilter;
        }
        if ($propertyName === 'PivotTables') {
            return $this->_model->PivotTables;
        }
        if ($propertyName === 'SelectUnlockedCells') {
            return $this->_model->SelectUnlockedCells;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'IsProtected') {
            $this->_model->IsProtected = $value;
            return;
        }
        if ($propertyName === 'Objects') {
            $this->_model->Objects = $value;
            $this->MarkProtectedWhenEnabled($value);
            return;
        }
        if ($propertyName === 'Scenarios') {
            $this->_model->Scenarios = $value;
            $this->MarkProtectedWhenEnabled($value);
            return;
        }
        if ($propertyName === 'FormatCells') {
            $this->_model->FormatCells = $value;
            $this->MarkProtectedWhenEnabled($value);
            return;
        }
        if ($propertyName === 'FormatColumns') {
            $this->_model->FormatColumns = $value;
            $this->MarkProtectedWhenEnabled($value);
            return;
        }
        if ($propertyName === 'FormatRows') {
            $this->_model->FormatRows = $value;
            $this->MarkProtectedWhenEnabled($value);
            return;
        }
        if ($propertyName === 'InsertColumns') {
            $this->_model->InsertColumns = $value;
            $this->MarkProtectedWhenEnabled($value);
            return;
        }
        if ($propertyName === 'InsertRows') {
            $this->_model->InsertRows = $value;
            $this->MarkProtectedWhenEnabled($value);
            return;
        }
        if ($propertyName === 'InsertHyperlinks') {
            $this->_model->InsertHyperlinks = $value;
            $this->MarkProtectedWhenEnabled($value);
            return;
        }
        if ($propertyName === 'DeleteColumns') {
            $this->_model->DeleteColumns = $value;
            $this->MarkProtectedWhenEnabled($value);
            return;
        }
        if ($propertyName === 'DeleteRows') {
            $this->_model->DeleteRows = $value;
            $this->MarkProtectedWhenEnabled($value);
            return;
        }
        if ($propertyName === 'SelectLockedCells') {
            $this->_model->SelectLockedCells = $value;
            $this->MarkProtectedWhenEnabled($value);
            return;
        }
        if ($propertyName === 'Sort') {
            $this->_model->Sort = $value;
            $this->MarkProtectedWhenEnabled($value);
            return;
        }
        if ($propertyName === 'AutoFilter') {
            $this->_model->AutoFilter = $value;
            $this->MarkProtectedWhenEnabled($value);
            return;
        }
        if ($propertyName === 'PivotTables') {
            $this->_model->PivotTables = $value;
            $this->MarkProtectedWhenEnabled($value);
            return;
        }
        if ($propertyName === 'SelectUnlockedCells') {
            $this->_model->SelectUnlockedCells = $value;
            $this->MarkProtectedWhenEnabled($value);
            return;
        }
    }

    public function Reset()
    {
        $this->_model->Clear();
    }

    private function MarkProtectedWhenEnabled($value)
    {
        if ($value) {
            $this->_model->IsProtected = true;
        }
    }
}
