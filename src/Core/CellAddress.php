<?php
namespace Aspose\Cells_FOSS\Core;

final class CellAddress
{
    public $RowIndex;
    public $ColumnIndex;

    public function __construct($rowIndex, $columnIndex)
    {
        if ($rowIndex < 0) {
            throw new \InvalidArgumentException("rowIndex");
        }
        if ($columnIndex < 0) {
            throw new \InvalidArgumentException("columnIndex");
        }
        $this->RowIndex = $rowIndex;
        $this->ColumnIndex = $columnIndex;
    }

    public static function Parse($cellReference)
    {
        if ($cellReference === null || trim($cellReference) === "") {
            throw new \InvalidArgumentException("Cell reference must be non-empty.");
        }
        $reference = trim($cellReference);
        $index = 0;
        $column = 0;
        $len = strlen($reference);
        while ($index < $len && ctype_alpha($reference[$index])) {
            $letter = strtoupper($reference[$index]);
            $ord = ord($letter);
            if ($ord < ord("A") || $ord > ord("Z")) {
                throw new \InvalidArgumentException("Cell reference is invalid.");
            }
            $column = ($column * 26) + ($ord - ord("A") + 1);
            $index++;
        }
        if ($column === 0 || $index === $len) {
            throw new \InvalidArgumentException("Cell reference is invalid.");
        }
        $row = 0;
        while ($index < $len && ctype_digit($reference[$index])) {
            $row = ($row * 10) + (ord($reference[$index]) - ord("0"));
            $index++;
        }
        if ($index !== $len || $row <= 0) {
            throw new \InvalidArgumentException("Cell reference is invalid.");
        }
        return new CellAddress($row - 1, $column - 1);
    }

    public function Equals($other)
    {
        if (!($other instanceof CellAddress)) {
            return false;
        }
        return $this->RowIndex === $other->RowIndex && $this->ColumnIndex === $other->ColumnIndex;
    }

    public function EqualsObject($obj)
    {
        if ($obj instanceof CellAddress) {
            return $this->Equals($obj);
        }
        return false;
    }

    public function GetHashCode()
    {
        return ($this->RowIndex * 397) ^ $this->ColumnIndex;
    }

    public function ToString()
    {
        return self::ColumnIndexToName($this->ColumnIndex) . ($this->RowIndex + 1);
    }

    private static function ColumnIndexToName($columnIndex)
    {
        $index = $columnIndex + 1;
        $characters = array();
        while ($index > 0) {
            $index--;
            $characters[] = chr(ord("A") + ($index % 26));
            $index = intdiv($index, 26);
        }
        $characters = array_reverse($characters);
        return implode("", $characters);
    }
}
