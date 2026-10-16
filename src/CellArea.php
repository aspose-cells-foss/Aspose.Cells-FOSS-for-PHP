<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;

class CellArea
{
    public $StartRow;
    public $EndRow;
    public $StartColumn;
    public $EndColumn;

    public function __construct($firstRow, $firstColumn, $totalRows, $totalColumns)
    {
        if ($firstRow < 0) throw new \InvalidArgumentException('firstRow');
        if ($firstColumn < 0) throw new \InvalidArgumentException('firstColumn');
        if ($totalRows <= 0) throw new \InvalidArgumentException('totalRows');
        if ($totalColumns <= 0) throw new \InvalidArgumentException('totalColumns');

        $this->StartRow = $firstRow;
        $this->StartColumn = $firstColumn;
        $this->EndRow = $firstRow + $totalRows - 1;
        $this->EndColumn = $firstColumn + $totalColumns - 1;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'FirstRow') {
            return $this->StartRow;
        }
        if ($propertyName === 'FirstColumn') {
            return $this->StartColumn;
        }
        if ($propertyName === 'TotalRows') {
            return $this->EndRow - $this->StartRow + 1;
        }
        if ($propertyName === 'TotalColumns') {
            return $this->EndColumn - $this->StartColumn + 1;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'FirstRow' || $propertyName === 'FirstColumn' || $propertyName === 'TotalRows' || $propertyName === 'TotalColumns') {
            throw new \RuntimeException($propertyName . ' is read-only.');
        }
    }

    public static function CreateCellArea($startRow, $startColumn, $endRow = null, $endColumn = null)
    {
        if (is_string($startRow) && is_string($startColumn)) {
            $start = CellAddress::Parse($startRow);
            $end = CellAddress::Parse($startColumn);
            return self::CreateCellArea(
                min($start->RowIndex, $end->RowIndex),
                min($start->ColumnIndex, $end->ColumnIndex),
                max($start->RowIndex, $end->RowIndex),
                max($start->ColumnIndex, $end->ColumnIndex));
        }

        if ($startRow < 0) throw new \InvalidArgumentException('startRow');
        if ($startColumn < 0) throw new \InvalidArgumentException('startColumn');
        if ($endRow < $startRow) throw new \InvalidArgumentException('endRow');
        if ($endColumn < $startColumn) throw new \InvalidArgumentException('endColumn');

        return new CellArea($startRow, $startColumn, $endRow - $startRow + 1, $endColumn - $startColumn + 1);
    }

    public function CompareTo($obj)
    {
        if ($obj === null) {
            return 1;
        }

        if (!($obj instanceof CellArea)) {
            throw new \InvalidArgumentException('Object must be a CellArea.', 'obj');
        }

        $other = $obj;
        $rowComparison = $this->StartRow - $other->StartRow;
        if ($rowComparison !== 0) {
            return $rowComparison;
        }

        $columnComparison = $this->StartColumn - $other->StartColumn;
        if ($columnComparison !== 0) {
            return $columnComparison;
        }

        $endRowComparison = $this->EndRow - $other->EndRow;
        if ($endRowComparison !== 0) {
            return $endRowComparison;
        }

        return $this->EndColumn - $other->EndColumn;
    }

    public function ToString()
    {
        return sprintf("R%dC%d:R%dC%d", $this->StartRow, $this->StartColumn, $this->EndRow, $this->EndColumn);
    }
}
