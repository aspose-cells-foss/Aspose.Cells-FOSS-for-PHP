<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\AutoFilterSortConditionModel;

final class AutoFilterSortConditionCollection implements \ArrayAccess, \Countable, \IteratorAggregate
{
    private $models;

    public function __construct(&$models)
    {
        $this->models = &$models;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Count') {
            return count($this->models);
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Count') {
            throw new \RuntimeException('Count is read-only.');
        }
    }

    #[\ReturnTypeWillChange]
    public function offsetExists($offset)
    {
        return $offset >= 0 && $offset < count($this->models);
    }

    #[\ReturnTypeWillChange]
    public function offsetGet($offset)
    {
        if ($offset < 0 || $offset >= count($this->models)) {
            throw new CellsException('Sort condition index was out of range.');
        }

        return new AutoFilterSortCondition($this->models[$offset]);
    }

    #[\ReturnTypeWillChange]
    public function offsetSet($offset, $value)
    {
        throw new \RuntimeException('Collection is read-only.');
    }

    #[\ReturnTypeWillChange]
    public function offsetUnset($offset)
    {
        throw new \RuntimeException('Collection is read-only.');
    }

    #[\ReturnTypeWillChange]
    public function count()
    {
        return count($this->models);
    }

    public function Add($reference)
    {
        $model = new AutoFilterSortConditionModel();
        $model->Ref = AutoFilterSupport::NormalizeRequiredRange($reference, 'reference');
        $this->models[] = $model;
        return count($this->models) - 1;
    }

    public function RemoveAt($index)
    {
        if ($index < 0 || $index >= count($this->models)) {
            throw new CellsException('Sort condition index was out of range.');
        }

        array_splice($this->models, $index, 1);
    }

    public function Clear()
    {
        $this->models = array();
    }

    public function GetEnumerator()
    {
        $conditions = array();
        for ($index = 0; $index < count($this->models); $index++) {
            $conditions[] = new AutoFilterSortCondition($this->models[$index]);
        }

        return $conditions;
    }

    #[\ReturnTypeWillChange]
    public function getIterator()
    {
        return new \ArrayIterator($this->GetEnumerator());
    }
}
