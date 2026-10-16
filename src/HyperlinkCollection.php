<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\HyperlinkModel;

class HyperlinkCollection implements \ArrayAccess, \Countable
{
    private $hyperlinks;
    private $capacity;

    public function __construct(&$hyperlinks)
    {
        $this->hyperlinks = &$hyperlinks;
        $this->capacity = count($hyperlinks);
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Count') {
            return count($this->hyperlinks);
        }
        if ($propertyName === 'Capacity') {
            return $this->capacity;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Capacity') {
            if ($value < count($this->hyperlinks)) {
                throw new CellsException('Capacity cannot be less than Count.');
            }

            $this->capacity = $value;
            return;
        }
    }

    #[\ReturnTypeWillChange]
    public function offsetExists($offset)
    {
        return $offset >= 0 && $offset < count($this->hyperlinks);
    }

    #[\ReturnTypeWillChange]
    public function offsetGet($offset)
    {
        if ($offset < 0 || $offset >= count($this->hyperlinks)) {
            throw new CellsException('Hyperlink index was out of range.');
        }

        return new Hyperlink($this->hyperlinks, $this->hyperlinks[$offset]);
    }

    #[\ReturnTypeWillChange]
    public function offsetSet($offset, $value)
    {
        if ($offset < 0 || $offset >= count($this->hyperlinks)) {
            throw new CellsException('Hyperlink index was out of range.');
        }

        if ($value === null) {
            throw new \InvalidArgumentException('value');
        }

        if (!$value->BelongsTo($this->hyperlinks)) {
            throw new CellsException('Assigned hyperlink must belong to the same worksheet hyperlink collection.');
        }

        $replacement = self::CloneModel($value->GetModel());
        for ($i = 0; $i < count($this->hyperlinks); $i++) {
            if ($i === $offset) {
                continue;
            }

            if (self::Overlaps($this->hyperlinks[$i], $replacement)) {
                throw new CellsException('Hyperlink ranges must not overlap.');
            }
        }

        $this->hyperlinks[$offset] = $replacement;
    }

    #[\ReturnTypeWillChange]
    public function offsetUnset($offset)
    {
        throw new \RuntimeException('Collection does not support removing by index.');
    }

    #[\ReturnTypeWillChange]
    public function count()
    {
        return count($this->hyperlinks);
    }

    public function Add($arg0 = null, $arg1 = null, $arg2 = null, $arg3 = null, $arg4 = null)
    {
        $arguments = func_get_args();
        if (count($arguments) === 4) {
            $cellName = $arguments[0];
            $totalRows = $arguments[1];
            $totalColumns = $arguments[2];
            $address = $arguments[3];

            if ($address === null) {
                throw new \InvalidArgumentException('address');
            }

            return $this->AddInternal($cellName, $totalRows, $totalColumns, $address);
        }

        if (count($arguments) === 5) {
            if (is_int($arguments[0]) && is_int($arguments[1])) {
                $firstRow = $arguments[0];
                $firstColumn = $arguments[1];
                $totalRows = $arguments[2];
                $totalColumns = $arguments[3];
                $address = $arguments[4];

                if ($firstRow < 0 || $firstColumn < 0) {
                    throw new CellsException('Hyperlink origin indices must be non-negative.');
                }

                if ($address === null) {
                    throw new \InvalidArgumentException('address');
                }

                return $this->AddInternal(null, $totalRows, $totalColumns, $address, $firstRow, $firstColumn);
            }

            $startCellName = $arguments[0];
            $endCellName = $arguments[1];
            $address = $arguments[2];
            $textToDisplay = $arguments[3];
            $screenTip = $arguments[4];

            if ($startCellName === null) {
                throw new \InvalidArgumentException('startCellName');
            }

            if ($endCellName === null) {
                throw new \InvalidArgumentException('endCellName');
            }

            if ($address === null) {
                throw new \InvalidArgumentException('address');
            }

            try {
                $startAddress = CellAddress::Parse($startCellName);
                $endAddress = CellAddress::Parse($endCellName);
            } catch (\InvalidArgumentException $exception) {
                throw new CellsException($exception->getMessage(), $exception);
            }

            $firstRow = min($startAddress->RowIndex, $endAddress->RowIndex);
            $firstColumn = min($startAddress->ColumnIndex, $endAddress->ColumnIndex);
            $lastRow = max($startAddress->RowIndex, $endAddress->RowIndex);
            $lastColumn = max($startAddress->ColumnIndex, $endAddress->ColumnIndex);
            $index = $this->AddInternal(null, $lastRow - $firstRow + 1, $lastColumn - $firstColumn + 1, $address, $firstRow, $firstColumn);
            $hyperlink = $this->hyperlinks[$index];
            $hyperlink->TextToDisplay = self::NormalizeText($textToDisplay);
            $hyperlink->ScreenTip = self::NormalizeText($screenTip);
            return $index;
        }

        throw new \InvalidArgumentException('No matching overload for Add.');
    }

    public function RemoveAt($index)
    {
        if ($index < 0 || $index >= count($this->hyperlinks)) {
            throw new CellsException('Hyperlink index was out of range.');
        }

        array_splice($this->hyperlinks, $index, 1);
    }

    public function Clear()
    {
        $this->hyperlinks = array();
    }

    private function AddInternal($cellName, $totalRows, $totalColumns, $address, $firstRowOverride = null, $firstColumnOverride = null)
    {
        if ($totalRows <= 0 || $totalColumns <= 0) {
            throw new CellsException('Hyperlink range dimensions must be positive.');
        }

        if ($firstRowOverride !== null && $firstColumnOverride !== null) {
            $anchor = new CellAddress($firstRowOverride, $firstColumnOverride);
        } else {
            if ($cellName === null || trim($cellName) === '') {
                throw new CellsException('Hyperlink anchor must be a valid cell reference.');
            }

            try {
                $anchor = CellAddress::Parse($cellName);
            } catch (\InvalidArgumentException $exception) {
                throw new CellsException($exception->getMessage(), $exception);
            }
        }

        $candidate = new HyperlinkModel();
        $candidate->FirstRow = $anchor->RowIndex;
        $candidate->FirstColumn = $anchor->ColumnIndex;
        $candidate->TotalRows = $totalRows;
        $candidate->TotalColumns = $totalColumns;
        self::AssignAddress($candidate, $address);

        for ($index = 0; $index < count($this->hyperlinks); $index++) {
            if (self::Overlaps($this->hyperlinks[$index], $candidate)) {
                throw new CellsException('Hyperlink ranges must not overlap.');
            }
        }

        $this->hyperlinks[] = $candidate;
        return count($this->hyperlinks) - 1;
    }

    private static function AssignAddress(HyperlinkModel $model, $address)
    {
        if ($address === null || trim($address) === '') {
            throw new CellsException('Hyperlink address must be non-empty.');
        }

        $normalized = trim($address);
        if (str_starts_with($normalized, '#')) {
            $model->Address = null;
            $model->SubAddress = substr($normalized, 1);
            return;
        }

        if (strpos($normalized, '!') !== false) {
            $model->Address = null;
            $model->SubAddress = $normalized;
            return;
        }

        $model->Address = $normalized;
        $model->SubAddress = null;
    }

    private static function Overlaps(HyperlinkModel $left, HyperlinkModel $right)
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

    private static function CloneModel(HyperlinkModel $source)
    {
        $model = new HyperlinkModel();
        $model->FirstRow = $source->FirstRow;
        $model->FirstColumn = $source->FirstColumn;
        $model->TotalRows = $source->TotalRows;
        $model->TotalColumns = $source->TotalColumns;
        $model->Address = $source->Address;
        $model->SubAddress = $source->SubAddress;
        $model->TextToDisplay = $source->TextToDisplay;
        $model->ScreenTip = $source->ScreenTip;
        return $model;
    }

    private static function NormalizeText($value)
    {
        if ($value === null || $value === '') {
            return null;
        }

        return $value;
    }
}
