<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ConditionalFormattingModel;

// Represents a collection of conditional formatting objects.
final class ConditionalFormattingCollection implements \ArrayAccess, \Countable
{
    private $_collections;

    public function __construct(&$collections)
    {
        $this->_collections =& $collections;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Count') {
            return count($this->_collections);
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
        return count($this->_collections);
    }

    public function offsetExists($offset): bool
    {
        $index = (int) $offset;
        return $index >= 0 && $index < count($this->_collections);
    }

    public function offsetGet($offset): mixed
    {
        $index = (int) $offset;
        if ($index < 0 || $index >= count($this->_collections)) {
            throw new CellsException('Conditional formatting index was out of range.');
        }

        return new FormatConditionCollection($this->_collections, $this->_collections[$index]);
    }

    public function offsetSet($offset, $value): void
    {
        throw new \RuntimeException('The conditional formatting collection is read-only.');
    }

    public function offsetUnset($offset): void
    {
        throw new \RuntimeException('The conditional formatting collection is read-only.');
    }

    public function Add()
    {
        $this->_collections[] = new ConditionalFormattingModel();
        return count($this->_collections) - 1;
    }

    public function RemoveAt($index)
    {
        if ($index < 0 || $index >= count($this->_collections)) {
            throw new CellsException('Conditional formatting index was out of range.');
        }

        array_splice($this->_collections, $index, 1);
    }

    public function RemoveArea($startRow, $startColumn, $totalRows, $totalColumns)
    {
        $area = CellArea::CreateCellArea($startRow, $startColumn, $startRow + $totalRows - 1, $startColumn + $totalColumns - 1);
        for ($index = count($this->_collections) - 1; $index >= 0; $index--) {
            $collection = new FormatConditionCollection($this->_collections, $this->_collections[$index]);
            $collection->RemoveArea($area);
        }
    }

    public static function GetNextPriority($collections)
    {
        $maxPriority = 0;
        for ($collectionIndex = 0; $collectionIndex < count($collections); $collectionIndex++) {
            $collection = $collections[$collectionIndex];
            for ($conditionIndex = 0; $conditionIndex < count($collection->Conditions); $conditionIndex++) {
                $maxPriority = max($maxPriority, $collection->Conditions[$conditionIndex]->Priority);
            }
        }

        return $maxPriority + 1;
    }
}