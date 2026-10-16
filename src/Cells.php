<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\MergeRegion;

class Cells implements \ArrayAccess
{
    private $_worksheet;
    public $Rows;
    public $Columns;

    public function __construct($worksheet)
    {
        $this->_worksheet = $worksheet;
        $this->Rows = new RowCollection($worksheet);
        $this->Columns = new ColumnCollection($worksheet);
    }

    #[\ReturnTypeWillChange]
    public function offsetGet($offset)
    {
        if (is_array($offset)) {
            $row = $offset[0];
            $column = $offset[1];
            if ($row < 0 || $column < 0) throw new CellsException("Row and column indices must be non-negative.");
            return new Cell($this->_worksheet, new CellAddress($row, $column));
        }

        try {
            // The public API uses Aspose-style CellsException rather than exposing
            // the lower-level argument validation coming from A1 parsing.
            return new Cell($this->_worksheet, CellAddress::Parse($offset));
        } catch (\InvalidArgumentException $exception) {
            throw new CellsException($exception->getMessage(), $exception);
        }
    }

    #[\ReturnTypeWillChange]
    public function offsetSet($offset, $value)
    {
        throw new \RuntimeException('Cells indexer is read-only.');
    }

    #[\ReturnTypeWillChange]
    public function offsetUnset($offset)
    {
        throw new \RuntimeException('Cells indexer is read-only.');
    }

    #[\ReturnTypeWillChange]
    public function offsetExists($offset)
    {
        try {
            $this->offsetGet($offset);
            return true;
        } catch (\Throwable $exception) {
            return false;
        }
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Style') {
            return Style::FromCore($this->_worksheet->Workbook->Model->DefaultStyle->Clone());
        }

        if ($propertyName === 'MergedCells') {
            $merged = array();
            foreach ($this->_worksheet->Model->MergeRegions as $region) {
                $merged[] = CellArea::CreateCellArea($region->FirstRow, $region->FirstColumn, $region->FirstRow + $region->TotalRows - 1, $region->FirstColumn + $region->TotalColumns - 1);
            }
            return $merged;
        }

        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Style') {
            if ($value === null) throw new \InvalidArgumentException('value');
            $this->_worksheet->Workbook->Model->DefaultStyle = $value->ToCore();
            return;
        }
    }

    public function Merge($firstRow, $firstColumn, $totalRows, $totalColumns)
    {
        if ($firstRow < 0 || $firstColumn < 0) throw new CellsException("Merge origin indices must be non-negative.");
        if ($totalRows <= 0 || $totalColumns <= 0) throw new CellsException("Merge range dimensions must be positive.");

        // Validate overlaps eagerly so the worksheet keeps a deterministic merge set
        // instead of deferring conflict resolution until save time.
        $candidate = new MergeRegion($firstRow, $firstColumn, $totalRows, $totalColumns);
        foreach ($this->_worksheet->Model->MergeRegions as $existing) {
            if (self::Overlaps($existing, $candidate)) {
                throw new CellsException("Merge ranges must not overlap.");
            }
        }

        $this->_worksheet->Model->MergeRegions[] = $candidate;
    }

    private static function Overlaps($left, $right)
    {
        $leftLastRow = $left->FirstRow + $left->TotalRows - 1;
        $leftLastColumn = $left->FirstColumn + $left->TotalColumns - 1;
        $rightLastRow = $right->FirstRow + $right->TotalRows - 1;
        $rightLastColumn = $right->FirstColumn + $right->TotalColumns - 1;

        return $left->FirstRow <= $rightLastRow
            && $right->FirstRow <= $leftLastRow
            && $left->FirstColumn <= $rightLastColumn
            && $right->FirstColumn <= $leftLastColumn;
    }
}
