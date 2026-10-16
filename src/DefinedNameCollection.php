<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\DefinedNameModel;

final class DefinedNameCollection implements \ArrayAccess, \Countable, \IteratorAggregate
{
    private $workbook;

    public function __construct(Workbook $workbook)
    {
        $this->workbook = $workbook;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Count') {
            return count($this->workbook->Model->DefinedNames);
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
        return $offset >= 0 && $offset < count($this->workbook->Model->DefinedNames);
    }

    #[\ReturnTypeWillChange]
    public function offsetGet($offset)
    {
        if ($offset < 0 || $offset >= count($this->workbook->Model->DefinedNames)) {
            throw new CellsException('Defined name index was out of range.');
        }

        return new DefinedName($this->workbook, $this->workbook->Model->DefinedNames[$offset]);
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
        return count($this->workbook->Model->DefinedNames);
    }

    public function Add($name, $formula, $localSheetIndex = null)
    {
        $normalizedName = DefinedNameUtility::NormalizeName($name);
        $normalizedFormula = DefinedNameUtility::NormalizeFormula($formula);
        $this->workbook->EnsureValidDefinedNameScope($localSheetIndex);
        $this->workbook->EnsureUniqueDefinedName(null, $normalizedName, $localSheetIndex);

        $model = new DefinedNameModel();
        $model->Name = $normalizedName;
        $model->Formula = $normalizedFormula;
        $model->LocalSheetIndex = $localSheetIndex;

        $this->workbook->Model->DefinedNames[] = $model;
        return count($this->workbook->Model->DefinedNames) - 1;
    }

    public function RemoveAt($index)
    {
        if ($index < 0 || $index >= count($this->workbook->Model->DefinedNames)) {
            throw new CellsException('Defined name index was out of range.');
        }

        array_splice($this->workbook->Model->DefinedNames, $index, 1);
    }

    public function GetEnumerator()
    {
        $names = array();
        for ($index = 0; $index < count($this->workbook->Model->DefinedNames); $index++) {
            $names[] = new DefinedName($this->workbook, $this->workbook->Model->DefinedNames[$index]);
        }

        return $names;
    }

    #[\ReturnTypeWillChange]
    public function getIterator()
    {
        return new \ArrayIterator($this->GetEnumerator());
    }
}
