<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\CultureInfo;
use Aspose\Cells_FOSS\DisplayTextFormatter;
use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\DateSerialConverter;
use Aspose\Cells_FOSS\Core\StyleValue;

final class ChartWorkbookDataResolver
{
    private function __construct()
    {
    }

    public static function ResolveNumericRange($workbook, $formula, $dateSystem)
    {
        $result = array();
        if (!self::TryParseSingleRangeReference($workbook, $formula, $sheet, $firstRow, $firstColumn, $lastRow, $lastColumn)) {
            return $result;
        }

        if ($firstColumn === $lastColumn) {
            for ($row = $firstRow; $row <= $lastRow; $row++) {
                $result[] = self::ReadNumericCell($sheet, $row, $firstColumn, $dateSystem);
            }

            return $result;
        }

        if ($firstRow === $lastRow) {
            for ($column = $firstColumn; $column <= $lastColumn; $column++) {
                $result[] = self::ReadNumericCell($sheet, $firstRow, $column, $dateSystem);
            }
        }

        return $result;
    }

    public static function ResolveCategoryLabels($workbook, $formula, $formatCode, $dateSystem, $culture)
    {
        $numericValues = self::ResolveNumericRange($workbook, $formula, $dateSystem);
        $result = array();
        for ($i = 0; $i < count($numericValues); $i++) {
            if ($numericValues[$i] !== null) {
                $result[] = ChartXmlParser::FormatLabel($numericValues[$i], $formatCode, $dateSystem, $culture);
            } else {
                $result[] = '';
            }
        }

        return $result;
    }

    public static function ResolveStringRange($workbook, $formula)
    {
        $result = array();
        if (!self::TryParseSingleRangeReference($workbook, $formula, $sheet, $firstRow, $firstColumn, $lastRow, $lastColumn)) {
            return $result;
        }

        if ($firstColumn === $lastColumn) {
            for ($row = $firstRow; $row <= $lastRow; $row++) {
                $result[] = self::ReadStringCell($workbook, $sheet, $row, $firstColumn);
            }

            return $result;
        }

        if ($firstRow === $lastRow) {
            for ($column = $firstColumn; $column <= $lastColumn; $column++) {
                $result[] = self::ReadStringCell($workbook, $sheet, $firstRow, $column);
            }
        }

        return $result;
    }

    public static function ResolveSingleString($workbook, $formula)
    {
        $values = self::ResolveStringRange($workbook, $formula);
        if (count($values) > 0) {
            return $values[0];
        }

        return '';
    }

    private static function ReadNumericCell($sheet, $row, $column, $dateSystem)
    {
        $key = (new CellAddress($row, $column))->ToString();
        if (!array_key_exists($key, $sheet->Cells) || $sheet->Cells[$key] === null || $sheet->Cells[$key]->Value === null) {
            return null;
        }

        $record = $sheet->Cells[$key];
        if ($record->Value instanceof \DateTime) {
            return DateSerialConverter::ToSerial($record->Value, $dateSystem);
        }

        if (is_float($record->Value)) {
            return (float)$record->Value;
        }

        if (is_int($record->Value)) {
            return $record->Value;
        }

        if (!is_bool($record->Value) && is_numeric((string)$record->Value)) {
            return (float)$record->Value;
        }

        return null;
    }

    private static function ReadStringCell($workbook, $sheet, $row, $column)
    {
        $key = (new CellAddress($row, $column))->ToString();
        if (!array_key_exists($key, $sheet->Cells) || $sheet->Cells[$key] === null || $sheet->Cells[$key]->Value === null) {
            return '';
        }

        $record = $sheet->Cells[$key];
        $style = $record->Style !== null ? $record->Style : StyleValue::Default();
        $culture = $workbook !== null && $workbook->Settings !== null
            ? $workbook->Settings->DisplayCulture
            : CultureInfo::GetInvariantCulture();
        return DisplayTextFormatter::FormatDisplayValue($record->Value, $style, $culture,
            $workbook->Settings->DateSystem);
    }

    private static function TryParseSingleRangeReference($workbook, $formula, &$sheet, &$firstRow, &$firstColumn, &$lastRow, &$lastColumn)
    {
        $sheet = null;
        $firstRow = 0;
        $firstColumn = 0;
        $lastRow = 0;
        $lastColumn = 0;

        if ($workbook === null || $formula === null || $formula === '') {
            return false;
        }

        $bang = strpos($formula, '!');
        if ($bang === false || $bang <= 0 || $bang >= strlen($formula) - 1) {
            return false;
        }

        $sheetName = substr($formula, 0, $bang);
        $range = substr($formula, $bang + 1);
        $sheetName = self::UnescapeSheetName($sheetName);

        $sheet = self::FindWorksheet($workbook, $sheetName);
        if ($sheet === null) {
            return false;
        }

        $parts = explode(':', $range);
        if (count($parts) === 1) {
            $parts = array($parts[0], $parts[0]);
        }

        if (!self::TryParseAbsoluteCell($parts[0], $first) || !self::TryParseAbsoluteCell($parts[1], $last)) {
            return false;
        }

        $firstRow = min($first->RowIndex, $last->RowIndex);
        $firstColumn = min($first->ColumnIndex, $last->ColumnIndex);
        $lastRow = max($first->RowIndex, $last->RowIndex);
        $lastColumn = max($first->ColumnIndex, $last->ColumnIndex);
        return true;
    }

    private static function FindWorksheet($workbook, $sheetName)
    {
        for ($i = 0; $i < count($workbook->Worksheets); $i++) {
            if ($workbook->Worksheets[$i]->Name === $sheetName) {
                return $workbook->Worksheets[$i];
            }
        }

        return null;
    }

    private static function UnescapeSheetName($sheetName)
    {
        $trimmed = trim($sheetName);
        if (strlen($trimmed) >= 2 && $trimmed[0] === "'" && $trimmed[strlen($trimmed) - 1] === "'") {
            $trimmed = str_replace("''", "'", substr($trimmed, 1, strlen($trimmed) - 2));
        }

        return $trimmed;
    }

    private static function TryParseAbsoluteCell($value, &$address)
    {
        $normalized = trim(str_replace('$', '', $value));
        try {
            $address = CellAddress::Parse($normalized);
            return true;
        } catch (\Throwable $exception) {
            $address = null;
            return false;
        }
    }
}
