<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ListObjectModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;

// Represents an Excel table (structured reference / ListObject).
final class ListObject
{
    private $_model;
    private $_worksheetModel;
    private $_owner;
    public $ListColumns;

    public function __construct(ListObjectModel $model, WorksheetModel $worksheetModel, IListObjectOwner $owner)
    {
        $this->_model = $model;
        $this->_worksheetModel = $worksheetModel;
        $this->_owner = $owner;
        $this->ListColumns = new ListColumnCollection($model);
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'DisplayName') {
            return $this->_model->DisplayName;
        }
        if ($propertyName === 'Comment') {
            return $this->_model->Comment;
        }
        if ($propertyName === 'StartRow') {
            return $this->_model->StartRow;
        }
        if ($propertyName === 'StartColumn') {
            return $this->_model->StartColumn;
        }
        if ($propertyName === 'EndRow') {
            return $this->_model->EndRow;
        }
        if ($propertyName === 'EndColumn') {
            return $this->_model->EndColumn;
        }
        if ($propertyName === 'ShowHeaderRow') {
            return $this->_model->ShowHeaderRow;
        }
        if ($propertyName === 'ShowTotals') {
            return $this->_model->ShowTotals;
        }
        if ($propertyName === 'TableStyleType') {
            return ListObjectSupport::TableStyleTypeFromName($this->_model->TableStyleName);
        }
        if ($propertyName === 'TableStyleName') {
            return $this->_model->TableStyleName;
        }
        if ($propertyName === 'ShowTableStyleFirstColumn') {
            return $this->_model->ShowFirstColumn;
        }
        if ($propertyName === 'ShowTableStyleLastColumn') {
            return $this->_model->ShowLastColumn;
        }
        if ($propertyName === 'ShowTableStyleRowStripes') {
            return $this->_model->ShowRowStripes;
        }
        if ($propertyName === 'ShowTableStyleColumnStripes') {
            return $this->_model->ShowColumnStripes;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'DisplayName') {
            if ($value === null || $value === '') {
                throw new CellsException('Table DisplayName must be non-empty.');
            }
            ListObjectSupport::ValidateDisplayName($value);
            $this->_owner->ValidateUniqueDisplayName($value, $this->_model);
            $this->_model->DisplayName = $value;
            $this->_model->Name = $value;
            return true;
        }
        if ($propertyName === 'Comment') {
            $this->_model->Comment = $value ?? '';
            return true;
        }
        if ($propertyName === 'StartRow') {
            throw new \RuntimeException('StartRow is read-only.');
        }
        if ($propertyName === 'StartColumn') {
            throw new \RuntimeException('StartColumn is read-only.');
        }
        if ($propertyName === 'EndRow') {
            throw new \RuntimeException('EndRow is read-only.');
        }
        if ($propertyName === 'EndColumn') {
            throw new \RuntimeException('EndColumn is read-only.');
        }
        if ($propertyName === 'ShowHeaderRow') {
            $this->_model->ShowHeaderRow = $value;
            $this->_model->HasAutoFilter = $value;
            return true;
        }
        if ($propertyName === 'ShowTotals') {
            if ($value === $this->_model->ShowTotals) {
                return true;
            }
            if ($value) {
                $this->_model->EndRow = $this->_model->EndRow + 1;
            } else {
                if ($this->_model->EndRow > $this->_model->StartRow) {
                    $this->_model->EndRow = $this->_model->EndRow - 1;
                }
            }
            $this->_model->ShowTotals = $value;
            return true;
        }
        if ($propertyName === 'TableStyleType') {
            if ($value !== TableStyleType::Custom) {
                $this->_model->TableStyleName = ListObjectSupport::TableStyleTypeToName($value);
            }
            return true;
        }
        if ($propertyName === 'TableStyleName') {
            $this->_model->TableStyleName = $value ?? '';
            return true;
        }
        if ($propertyName === 'ShowTableStyleFirstColumn') {
            $this->_model->ShowFirstColumn = $value;
            return true;
        }
        if ($propertyName === 'ShowTableStyleLastColumn') {
            $this->_model->ShowLastColumn = $value;
            return true;
        }
        if ($propertyName === 'ShowTableStyleRowStripes') {
            $this->_model->ShowRowStripes = $value;
            return true;
        }
        if ($propertyName === 'ShowTableStyleColumnStripes') {
            $this->_model->ShowColumnStripes = $value;
            return true;
        }
        return false;
    }

    public function Resize($startRow, $startColumn, $endRow, $endColumn, $hasHeaders)
    {
        ListObjectSupport::ValidateRange($startRow, $startColumn, $endRow, $endColumn);
        $this->_owner->ValidateNoOverlap($startRow, $startColumn, $endRow, $endColumn, $this->_model);
        $this->_model->StartRow = $startRow;
        $this->_model->StartColumn = $startColumn;
        $this->_model->EndRow = $endRow;
        $this->_model->EndColumn = $endColumn;
        $this->_model->ShowHeaderRow = $hasHeaders;
        $this->_model->HasAutoFilter = $hasHeaders;
        ListObjectSupport::RebuildColumns($this->_model, $this->_worksheetModel);
    }

    public function ShowAutoFilter()
    {
        $this->_model->HasAutoFilter = true;
    }

    public function RemoveAutoFilter()
    {
        $this->_model->HasAutoFilter = false;
    }

    public function ConvertToRange()
    {
        $this->_owner->RemoveTable($this->_model);
    }
}