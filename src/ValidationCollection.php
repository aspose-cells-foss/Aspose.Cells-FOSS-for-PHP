<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ValidationModel;

final class ValidationCollection implements \ArrayAccess, \Countable
{
    private $validations;

    public function __construct(&$validations)
    {
        $this->validations = &$validations;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Count') {
            return count($this->validations);
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
        return $offset >= 0 && $offset < count($this->validations);
    }

    #[\ReturnTypeWillChange]
    public function offsetGet($offset)
    {
        if ($offset < 0 || $offset >= count($this->validations)) {
            throw new CellsException('Validation index was out of range.');
        }

        return new Validation($this->validations, $this->validations[$offset]);
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
        return count($this->validations);
    }

    public function Add(CellArea $area)
    {
        self::ValidateArea($area);
        self::EnsureNoOverlap($this->validations, null, $area);

        $validation = new ValidationModel();
        $validation->Areas[] = $area;
        $this->validations[] = $validation;
        return count($this->validations) - 1;
    }

    public function GetValidationInCell($row, $column)
    {
        if ($row < 0 || $column < 0) {
            throw new CellsException('Row and column indices must be non-negative.');
        }

        for ($index = 0; $index < count($this->validations); $index++) {
            $validation = $this->validations[$index];
            for ($areaIndex = 0; $areaIndex < count($validation->Areas); $areaIndex++) {
                if (self::Contains($validation->Areas[$areaIndex], $row, $column)) {
                    return new Validation($this->validations, $validation);
                }
            }
        }

        return null;
    }

    public function RemoveACell($row, $column)
    {
        if ($row < 0 || $column < 0) {
            throw new CellsException('Row and column indices must be non-negative.');
        }

        $this->RemoveArea(CellArea::CreateCellArea($row, $column, $row, $column));
    }

    public function RemoveArea(CellArea $cellArea)
    {
        self::ValidateArea($cellArea);

        for ($index = count($this->validations) - 1; $index >= 0; $index--) {
            $validation = $this->validations[$index];
            self::ReplaceAreas($validation, self::SubtractAreas($validation->Areas, $cellArea));
            if (count($validation->Areas) === 0) {
                array_splice($this->validations, $index, 1);
            }
        }
    }

    public static function AddAreaToValidation(&$owner, ValidationModel $validation, CellArea $area)
    {
        self::ValidateArea($area);
        self::EnsureNoOverlap($owner, $validation, $area);

        for ($index = 0; $index < count($validation->Areas); $index++) {
            if (self::AreasOverlap($validation->Areas[$index], $area)) {
                throw new CellsException('Validation areas must not overlap.');
            }
        }

        $validation->Areas[] = $area;
        self::SortAreas($validation->Areas);
    }

    public static function RemoveAreaFromValidation(&$owner, ValidationModel $validation, CellArea $area)
    {
        self::ValidateArea($area);
        self::ReplaceAreas($validation, self::SubtractAreas($validation->Areas, $area));
        if (count($validation->Areas) === 0) {
            for ($index = 0; $index < count($owner); $index++) {
                if ($owner[$index] === $validation) {
                    array_splice($owner, $index, 1);
                    break;
                }
            }
        }
    }

    private static function ReplaceAreas(ValidationModel $validation, $areas)
    {
        $validation->Areas = array();
        for ($index = 0; $index < count($areas); $index++) {
            $validation->Areas[] = $areas[$index];
        }
    }

    private static function ValidateArea(CellArea $area)
    {
        if ($area->FirstRow < 0 || $area->FirstColumn < 0 || $area->TotalRows <= 0 || $area->TotalColumns <= 0) {
            throw new CellsException('Validation area must be a positive cell range.');
        }
    }

    private static function EnsureNoOverlap($validations, $currentValidation, CellArea $candidate)
    {
        for ($validationIndex = 0; $validationIndex < count($validations); $validationIndex++) {
            $validation = $validations[$validationIndex];
            if ($validation === $currentValidation) {
                continue;
            }

            for ($areaIndex = 0; $areaIndex < count($validation->Areas); $areaIndex++) {
                if (self::AreasOverlap($validation->Areas[$areaIndex], $candidate)) {
                    throw new CellsException('Validation areas must not overlap.');
                }
            }
        }
    }

    private static function SubtractAreas($sourceAreas, CellArea $removal)
    {
        $remaining = array();
        for ($index = 0; $index < count($sourceAreas); $index++) {
            self::SubtractArea($sourceAreas[$index], $removal, $remaining);
        }

        self::SortAreas($remaining);
        return $remaining;
    }

    private static function SubtractArea(CellArea $source, CellArea $removal, &$output)
    {
        if (!self::AreasOverlap($source, $removal)) {
            $output[] = $source;
            return;
        }

        $sourceLastRow = $source->FirstRow + $source->TotalRows - 1;
        $sourceLastColumn = $source->FirstColumn + $source->TotalColumns - 1;
        $removalLastRow = $removal->FirstRow + $removal->TotalRows - 1;
        $removalLastColumn = $removal->FirstColumn + $removal->TotalColumns - 1;

        $overlapFirstRow = max($source->FirstRow, $removal->FirstRow);
        $overlapFirstColumn = max($source->FirstColumn, $removal->FirstColumn);
        $overlapLastRow = min($sourceLastRow, $removalLastRow);
        $overlapLastColumn = min($sourceLastColumn, $removalLastColumn);

        self::AddIfNonEmpty($output, $source->FirstRow, $source->FirstColumn, $overlapFirstRow - 1, $sourceLastColumn);
        self::AddIfNonEmpty($output, $overlapLastRow + 1, $source->FirstColumn, $sourceLastRow, $sourceLastColumn);
        self::AddIfNonEmpty($output, $overlapFirstRow, $source->FirstColumn, $overlapLastRow, $overlapFirstColumn - 1);
        self::AddIfNonEmpty($output, $overlapFirstRow, $overlapLastColumn + 1, $overlapLastRow, $sourceLastColumn);
    }

    private static function AddIfNonEmpty(&$areas, $firstRow, $firstColumn, $lastRow, $lastColumn)
    {
        if ($lastRow < $firstRow || $lastColumn < $firstColumn) {
            return;
        }

        $areas[] = CellArea::CreateCellArea($firstRow, $firstColumn, $lastRow, $lastColumn);
    }

    public static function AreasOverlap(CellArea $left, CellArea $right)
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

    private static function Contains(CellArea $area, $row, $column)
    {
        return $row >= $area->FirstRow
            && $row < $area->FirstRow + $area->TotalRows
            && $column >= $area->FirstColumn
            && $column < $area->FirstColumn + $area->TotalColumns;
    }

    public static function SortAreas(&$areas)
    {
        usort($areas, function ($left, $right) {
            return self::CompareAreas($left, $right);
        });
    }

    public static function CompareAreas(CellArea $left, CellArea $right)
    {
        if ($left->FirstRow < $right->FirstRow) {
            return -1;
        }
        if ($left->FirstRow > $right->FirstRow) {
            return 1;
        }

        if ($left->FirstColumn < $right->FirstColumn) {
            return -1;
        }
        if ($left->FirstColumn > $right->FirstColumn) {
            return 1;
        }

        if ($left->TotalRows < $right->TotalRows) {
            return -1;
        }
        if ($left->TotalRows > $right->TotalRows) {
            return 1;
        }

        if ($left->TotalColumns < $right->TotalColumns) {
            return -1;
        }
        if ($left->TotalColumns > $right->TotalColumns) {
            return 1;
        }

        return 0;
    }
}
