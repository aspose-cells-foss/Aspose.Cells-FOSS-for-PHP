<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\ListObjectModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;

// Represents the collection of Excel tables on a worksheet.
final class ListObjectCollection implements IListObjectOwner, \ArrayAccess, \Countable
{
    private $_worksheetModel;
    private $_models;

    public function __construct(WorksheetModel $worksheetModel)
    {
        $this->_worksheetModel = $worksheetModel;
        $this->_models =& $worksheetModel->ListObjects;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Count') {
            return count($this->_models);
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Count') {
            throw new \RuntimeException('Count is read-only.');
        }
    }

    public function count(): int
    {
        return count($this->_models);
    }

    public function offsetExists($offset): bool
    {
        if (is_int($offset) || (is_string($offset) && ctype_digit($offset))) {
            $index = (int) $offset;
            return $index >= 0 && $index < count($this->_models);
        }
        return false;
    }

    public function offsetGet($offset): mixed
    {
        if (is_int($offset) || (is_string($offset) && ctype_digit($offset))) {
            $index = (int) $offset;
            if ($index < 0 || $index >= count($this->_models)) {
                throw new CellsException('Table index ' . $index . ' is out of range.');
            }
            return new ListObject($this->_models[$index], $this->_worksheetModel, $this);
        }

        $name = (string) $offset;
        for ($i = 0; $i < count($this->_models); $i++) {
            if (strcasecmp($this->_models[$i]->DisplayName, $name) === 0) {
                return new ListObject($this->_models[$i], $this->_worksheetModel, $this);
            }
        }
        throw new CellsException("No table with display name '" . $name . "' was found.");
    }

    public function offsetSet($offset, $value): void
    {
        throw new \RuntimeException('The list object collection is read-only.');
    }

    public function offsetUnset($offset): void
    {
        throw new \RuntimeException('The list object collection is read-only.');
    }

    // Adds a new table covering the specified zero-based integer range or A1-notation
    // cell range and returns its index. The five-argument form accepts integer
    // coordinates (startRow, startColumn, endRow, endColumn, hasHeaders); the
    // three-argument form accepts (startCellName, endCellName, hasHeaders).
    public function Add($a, $b, $c, $d = null, $e = null)
    {
        if ($d === null) {
            $startCellName = $a;
            $endCellName = $b;
            $hasHeaders = $c;
            try {
                $startCell = CellAddress::Parse($startCellName);
            } catch (\InvalidArgumentException $exception) {
                throw new CellsException("The start cell reference '" . $startCellName . "' is invalid.");
            }
            try {
                $endCell = CellAddress::Parse($endCellName);
            } catch (\InvalidArgumentException $exception) {
                throw new CellsException("The end cell reference '" . $endCellName . "' is invalid.");
            }
            return $this->Add($startCell->RowIndex, $startCell->ColumnIndex, $endCell->RowIndex, $endCell->ColumnIndex, $hasHeaders);
        }

        $startRow = $a;
        $startColumn = $b;
        $endRow = $c;
        $endColumn = $d;
        $hasHeaders = $e;
        ListObjectSupport::ValidateRange($startRow, $startColumn, $endRow, $endColumn);
        ListObjectSupport::ValidateNoOverlap($this->_models, $startRow, $startColumn, $endRow, $endColumn, -1);
        $tableNumber = count($this->_models) + 1;
        $model = ListObjectSupport::CreateModel($this->_worksheetModel, $startRow, $startColumn, $endRow, $endColumn, $hasHeaders, $tableNumber);
        ListObjectSupport::ValidateUniqueDisplayName($this->_models, $model->DisplayName, -1);
        $this->_models[] = $model;
        return count($this->_models) - 1;
    }

    public function RemoveAt($index)
    {
        if ($index < 0 || $index >= count($this->_models)) {
            throw new CellsException('Table index ' . $index . ' is out of range.');
        }
        array_splice($this->_models, $index, 1);
    }

    public function ValidateUniqueDisplayName($displayName, ListObjectModel $skipModel)
    {
        $skipIndex = -1;
        for ($i = 0; $i < count($this->_models); $i++) {
            if ($this->_models[$i] === $skipModel) {
                $skipIndex = $i;
                break;
            }
        }
        ListObjectSupport::ValidateUniqueDisplayName($this->_models, $displayName, $skipIndex);
    }

    public function ValidateNoOverlap($startRow, $startColumn, $endRow, $endColumn, ListObjectModel $skipModel)
    {
        $skipIndex = -1;
        for ($i = 0; $i < count($this->_models); $i++) {
            if ($this->_models[$i] === $skipModel) {
                $skipIndex = $i;
                break;
            }
        }
        ListObjectSupport::ValidateNoOverlap($this->_models, $startRow, $startColumn, $endRow, $endColumn, $skipIndex);
    }

    public function RemoveTable(ListObjectModel $model)
    {
        foreach ($this->_models as $i => $existing) {
            if ($existing === $model) {
                array_splice($this->_models, $i, 1);
                return;
            }
        }
    }
}