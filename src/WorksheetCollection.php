<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\WorksheetModel;

class WorksheetCollection implements \ArrayAccess, \IteratorAggregate
{
    private $_workbook;

    public function __construct(Workbook $workbook)
    {
        $this->_workbook = $workbook;
    }

    #[\ReturnTypeWillChange]
    public function offsetExists($offset)
    {
        if (is_int($offset)) {
            return $offset >= 0 && $offset < count($this->_workbook->Model->Worksheets);
        }

        return $this->IndexOf($offset) >= 0;
    }

    #[\ReturnTypeWillChange]
    public function offsetGet($offset)
    {
        if (is_int($offset)) {
            return $this->Wrap($this->_workbook->Model->Worksheets[$offset]);
        }

        $index = $this->IndexOf($offset);
        if ($index < 0) {
            throw new CellsException("Worksheet '" . $offset . "' was not found.");
        }

        return $this->Wrap($this->_workbook->Model->Worksheets[$index]);
    }

    #[\ReturnTypeWillChange]
    public function offsetSet($offset, $value)
    {
        throw new \RuntimeException('WorksheetCollection is read-only.');
    }

    #[\ReturnTypeWillChange]
    public function offsetUnset($offset)
    {
        throw new \RuntimeException('WorksheetCollection is read-only.');
    }

    #[\ReturnTypeWillChange]
    public function getIterator()
    {
        return new \ArrayIterator($this->GetEnumerator());
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Count') {
            return count($this->_workbook->Model->Worksheets);
        }
        if ($propertyName === 'ActiveSheetIndex') {
            return $this->_workbook->Model->ActiveSheetIndex;
        }
        if ($propertyName === 'ActiveSheetName') {
            return $this->_workbook->Model->Worksheets[$this->ActiveSheetIndex]->Name;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'ActiveSheetIndex') {
            if ($value < 0 || $value >= count($this->_workbook->Model->Worksheets)) {
                throw new CellsException('ActiveSheetIndex must refer to an existing worksheet.');
            }

            $this->_workbook->Model->ActiveSheetIndex = $value;
            return;
        }
        if ($propertyName === 'ActiveSheetName') {
            $index = $this->IndexOf($value);
            if ($index < 0) {
                throw new CellsException("Worksheet '" . $value . "' was not found.");
            }

            $this->ActiveSheetIndex = $index;
            return;
        }
    }

    public function Add($sheetName = null)
    {
        if ($sheetName === null) {
            $sheetName = $this->GenerateDefaultSheetName();
        }

        if (trim($sheetName) === '') {
            throw new CellsException('Worksheet name must be non-empty.');
        }
        $this->_workbook->EnsureUniqueSheetName($sheetName);
        $this->_workbook->Model->Worksheets[] = new WorksheetModel($sheetName);
        return count($this->_workbook->Model->Worksheets) - 1;
    }

    public function RemoveAt($sheetNameOrIndex)
    {
        if (is_int($sheetNameOrIndex)) {
            $index = $sheetNameOrIndex;
        } else {
            $index = $this->IndexOf($sheetNameOrIndex);
            if ($index < 0) {
                throw new CellsException("Worksheet '" . $sheetNameOrIndex . "' was not found.");
            }
        }

        if (count($this->_workbook->Model->Worksheets) === 1) {
            throw new CellsException('A workbook must contain at least one worksheet.');
        }

        array_splice($this->_workbook->Model->Worksheets, $index, 1);
        if ($this->_workbook->Model->ActiveSheetIndex > $index) {
            $this->_workbook->Model->ActiveSheetIndex--;
        } elseif ($this->_workbook->Model->ActiveSheetIndex >= count($this->_workbook->Model->Worksheets)) {
            $this->_workbook->Model->ActiveSheetIndex = count($this->_workbook->Model->Worksheets) - 1;
        }

        $firstSheet = $this->_workbook->Model->Properties->View->FirstSheet;
        if ($firstSheet !== null) {
            if (count($this->_workbook->Model->Worksheets) === 0) {
                $this->_workbook->Model->Properties->View->FirstSheet = 0;
            } elseif ($firstSheet >= count($this->_workbook->Model->Worksheets)) {
                $this->_workbook->Model->Properties->View->FirstSheet = count($this->_workbook->Model->Worksheets) - 1;
            }
        }
    }

    public function GetEnumerator()
    {
        $worksheets = array();
        for ($i = 0; $i < count($this->_workbook->Model->Worksheets); $i++) {
            $worksheets[] = $this->Wrap($this->_workbook->Model->Worksheets[$i]);
        }

        return $worksheets;
    }

    private function IndexOf($name)
    {
        if ($name === null) throw new \InvalidArgumentException('name');
        for ($i = 0; $i < count($this->_workbook->Model->Worksheets); $i++) {
            $worksheet = $this->_workbook->Model->Worksheets[$i];
            if (strcasecmp($worksheet->Name, $name) === 0) {
                return $i;
            }
        }

        return -1;
    }

    private function GenerateDefaultSheetName()
    {
        $suffix = 1;
        while (true) {
            $candidate = 'Sheet' . $suffix;
            if ($this->IndexOf($candidate) < 0) {
                return $candidate;
            }

            $suffix++;
        }
    }

    private function Wrap(WorksheetModel $worksheet)
    {
        return new Worksheet($this->_workbook, $worksheet);
    }
}
