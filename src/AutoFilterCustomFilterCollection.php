<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\AutoFilterCustomFilterModel;
use Aspose\Cells_FOSS\Core\FilterColumnModel;

final class AutoFilterCustomFilterCollection implements \ArrayAccess, \Countable, \IteratorAggregate
{
    private $models;
    private $columnModel;

    public function __construct(&$models, FilterColumnModel $columnModel)
    {
        $this->models = &$models;
        $this->columnModel = $columnModel;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'MatchAll') {
            return $this->columnModel->CustomFiltersAnd;
        }
        if ($propertyName === 'Count') {
            return count($this->models);
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'MatchAll') {
            $this->columnModel->CustomFiltersAnd = $value;
            return;
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
            throw new CellsException('Custom filter index was out of range.');
        }

        return new AutoFilterCustomFilter($this->models[$offset]);
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

    public function Add($operatorType, $value)
    {
        if (count($this->models) >= 2) {
            throw new CellsException('Custom filters support at most two filter conditions.');
        }

        $model = new AutoFilterCustomFilterModel();
        $name = AutoFilterSupport::ToOperatorName($operatorType);
        $model->Operator = $name === null ? '' : $name;
        $model->Value = AutoFilterSupport::NormalizeText($value, 'value');
        $this->models[] = $model;
        return count($this->models) - 1;
    }

    public function RemoveAt($index)
    {
        if ($index < 0 || $index >= count($this->models)) {
            throw new CellsException('Custom filter index was out of range.');
        }

        array_splice($this->models, $index, 1);
    }

    public function Clear()
    {
        $this->models = array();
        $this->columnModel->CustomFiltersAnd = false;
    }

    public function GetEnumerator()
    {
        $filters = array();
        for ($index = 0; $index < count($this->models); $index++) {
            $filters[] = new AutoFilterCustomFilter($this->models[$index]);
        }

        return $filters;
    }

    #[\ReturnTypeWillChange]
    public function getIterator()
    {
        return new \ArrayIterator($this->GetEnumerator());
    }
}
