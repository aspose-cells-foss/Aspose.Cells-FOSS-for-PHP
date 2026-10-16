<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\ListColumnModel;
use Aspose\Cells_FOSS\Core\ListObjectModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;

// Internal helpers for ListObject validation and model construction.
final class ListObjectSupport
{
    private function __construct()
    {
    }

    public static function ValidateRange($startRow, $startColumn, $endRow, $endColumn)
    {
        if ($startRow < 0) {
            throw new CellsException('Table startRow must be non-negative.');
        }
        if ($startColumn < 0) {
            throw new CellsException('Table startColumn must be non-negative.');
        }
        if ($endRow < $startRow) {
            throw new CellsException('Table endRow must be greater than or equal to startRow.');
        }
        if ($endColumn < $startColumn) {
            throw new CellsException('Table endColumn must be greater than or equal to startColumn.');
        }
    }

    public static function ValidateDisplayName($displayName)
    {
        if ($displayName === null || $displayName === '') {
            throw new CellsException('Table DisplayName must be non-empty.');
        }

        $first = $displayName[0];
        if (!ctype_alpha($first) && $first !== '_' && $first !== '\\') {
            throw new CellsException("Table DisplayName '" . $displayName . "' must start with a letter or underscore.");
        }
    }

    public static function ValidateUniqueDisplayName($existing, $displayName, $skipIndex)
    {
        for ($i = 0; $i < count($existing); $i++) {
            if ($i === $skipIndex) {
                continue;
            }
            if (strcasecmp($existing[$i]->DisplayName, $displayName) === 0) {
                throw new CellsException("A table with display name '" . $displayName . "' already exists in this worksheet.");
            }
        }
    }

    public static function ValidateNoOverlap($existing, $startRow, $startColumn, $endRow, $endColumn, $skipIndex)
    {
        for ($i = 0; $i < count($existing); $i++) {
            if ($i === $skipIndex) {
                continue;
            }
            $other = $existing[$i];
            if (self::RangesOverlap($startRow, $startColumn, $endRow, $endColumn, $other->StartRow, $other->StartColumn, $other->EndRow, $other->EndColumn)) {
                throw new CellsException("The table range overlaps with existing table '" . $other->DisplayName . "'.");
            }
        }
    }

    public static function RangesOverlap($r1s, $c1s, $r1e, $c1e, $r2s, $c2s, $r2e, $c2e)
    {
        return $r1s <= $r2e && $r2s <= $r1e && $c1s <= $c2e && $c2s <= $c1e;
    }

    public static function CreateModel(WorksheetModel $worksheetModel, $startRow, $startColumn, $endRow, $endColumn, $hasHeaders, $tableNumber)
    {
        $displayName = 'Table' . $tableNumber;
        $model = new ListObjectModel();
        $model->DisplayName = $displayName;
        $model->Name = $displayName;
        $model->StartRow = $startRow;
        $model->StartColumn = $startColumn;
        $model->EndRow = $endRow;
        $model->EndColumn = $endColumn;
        $model->ShowHeaderRow = $hasHeaders;
        $model->HasAutoFilter = $hasHeaders;

        $columnCount = $endColumn - $startColumn + 1;
        for ($c = 0; $c < $columnCount; $c++) {
            $columnName = self::ResolveColumnName($worksheetModel, $startRow, $startColumn + $c, $hasHeaders, $c + 1);
            $model->Columns[] = new ListColumnModel($c + 1, $columnName);
        }

        return $model;
    }

    public static function RebuildColumns(ListObjectModel $model, WorksheetModel $worksheetModel)
    {
        $model->Columns = array();
        $columnCount = $model->EndColumn - $model->StartColumn + 1;
        for ($c = 0; $c < $columnCount; $c++) {
            $columnName = self::ResolveColumnName($worksheetModel, $model->StartRow, $model->StartColumn + $c, $model->ShowHeaderRow, $c + 1);
            $model->Columns[] = new ListColumnModel($c + 1, $columnName);
        }
    }

    private static function ResolveColumnName(WorksheetModel $worksheetModel, $headerRow, $columnIndex, $hasHeaders, $defaultColumnNumber)
    {
        if ($hasHeaders) {
            $address = new CellAddress($headerRow, $columnIndex);
            $key = $address->ToString();
            if (array_key_exists($key, $worksheetModel->Cells) && is_string($worksheetModel->Cells[$key]->Value)) {
                $name = (string) $worksheetModel->Cells[$key]->Value;
                if (!($name === null || $name === '')) {
                    return $name;
                }
            }
        }
        return 'Column' . $defaultColumnNumber;
    }

    public static function TableStyleTypeToName($type)
    {
        if ($type === TableStyleType::None || $type === TableStyleType::Custom) {
            return '';
        }

        $map = self::GetTableStyleNameToValue();
        $names = array_flip($map);
        if (array_key_exists($type, $names)) {
            return $names[$type];
        }
        return (string) $type;
    }

    public static function TableStyleTypeFromName($name)
    {
        if ($name === null || $name === '') {
            return TableStyleType::None;
        }

        $parsed = null;
        if (self::TryParseTableStyleType($name, $parsed)) {
            return $parsed;
        }
        return TableStyleType::Custom;
    }

    private static function TryParseTableStyleType($name, &$result)
    {
        $map = self::GetTableStyleNameToValue();
        if (array_key_exists($name, $map)) {
            $result = $map[$name];
            return $result !== TableStyleType::None && $result !== TableStyleType::Custom;
        }
        $result = TableStyleType::None;
        return false;
    }

    private static function GetTableStyleNameToValue()
    {
        return array(
            'None' => TableStyleType::None,
            'TableStyleLight1' => TableStyleType::TableStyleLight1,
            'TableStyleLight2' => TableStyleType::TableStyleLight2,
            'TableStyleLight3' => TableStyleType::TableStyleLight3,
            'TableStyleLight4' => TableStyleType::TableStyleLight4,
            'TableStyleLight5' => TableStyleType::TableStyleLight5,
            'TableStyleLight6' => TableStyleType::TableStyleLight6,
            'TableStyleLight7' => TableStyleType::TableStyleLight7,
            'TableStyleLight8' => TableStyleType::TableStyleLight8,
            'TableStyleLight9' => TableStyleType::TableStyleLight9,
            'TableStyleLight10' => TableStyleType::TableStyleLight10,
            'TableStyleLight11' => TableStyleType::TableStyleLight11,
            'TableStyleLight12' => TableStyleType::TableStyleLight12,
            'TableStyleLight13' => TableStyleType::TableStyleLight13,
            'TableStyleLight14' => TableStyleType::TableStyleLight14,
            'TableStyleLight15' => TableStyleType::TableStyleLight15,
            'TableStyleLight16' => TableStyleType::TableStyleLight16,
            'TableStyleLight17' => TableStyleType::TableStyleLight17,
            'TableStyleLight18' => TableStyleType::TableStyleLight18,
            'TableStyleLight19' => TableStyleType::TableStyleLight19,
            'TableStyleLight20' => TableStyleType::TableStyleLight20,
            'TableStyleLight21' => TableStyleType::TableStyleLight21,
            'TableStyleMedium1' => TableStyleType::TableStyleMedium1,
            'TableStyleMedium2' => TableStyleType::TableStyleMedium2,
            'TableStyleMedium3' => TableStyleType::TableStyleMedium3,
            'TableStyleMedium4' => TableStyleType::TableStyleMedium4,
            'TableStyleMedium5' => TableStyleType::TableStyleMedium5,
            'TableStyleMedium6' => TableStyleType::TableStyleMedium6,
            'TableStyleMedium7' => TableStyleType::TableStyleMedium7,
            'TableStyleMedium8' => TableStyleType::TableStyleMedium8,
            'TableStyleMedium9' => TableStyleType::TableStyleMedium9,
            'TableStyleMedium10' => TableStyleType::TableStyleMedium10,
            'TableStyleMedium11' => TableStyleType::TableStyleMedium11,
            'TableStyleMedium12' => TableStyleType::TableStyleMedium12,
            'TableStyleMedium13' => TableStyleType::TableStyleMedium13,
            'TableStyleMedium14' => TableStyleType::TableStyleMedium14,
            'TableStyleMedium15' => TableStyleType::TableStyleMedium15,
            'TableStyleMedium16' => TableStyleType::TableStyleMedium16,
            'TableStyleMedium17' => TableStyleType::TableStyleMedium17,
            'TableStyleMedium18' => TableStyleType::TableStyleMedium18,
            'TableStyleMedium19' => TableStyleType::TableStyleMedium19,
            'TableStyleMedium20' => TableStyleType::TableStyleMedium20,
            'TableStyleMedium21' => TableStyleType::TableStyleMedium21,
            'TableStyleMedium22' => TableStyleType::TableStyleMedium22,
            'TableStyleMedium23' => TableStyleType::TableStyleMedium23,
            'TableStyleMedium24' => TableStyleType::TableStyleMedium24,
            'TableStyleMedium25' => TableStyleType::TableStyleMedium25,
            'TableStyleMedium26' => TableStyleType::TableStyleMedium26,
            'TableStyleMedium27' => TableStyleType::TableStyleMedium27,
            'TableStyleMedium28' => TableStyleType::TableStyleMedium28,
            'TableStyleDark1' => TableStyleType::TableStyleDark1,
            'TableStyleDark2' => TableStyleType::TableStyleDark2,
            'TableStyleDark3' => TableStyleType::TableStyleDark3,
            'TableStyleDark4' => TableStyleType::TableStyleDark4,
            'TableStyleDark5' => TableStyleType::TableStyleDark5,
            'TableStyleDark6' => TableStyleType::TableStyleDark6,
            'TableStyleDark7' => TableStyleType::TableStyleDark7,
            'TableStyleDark8' => TableStyleType::TableStyleDark8,
            'TableStyleDark9' => TableStyleType::TableStyleDark9,
            'TableStyleDark10' => TableStyleType::TableStyleDark10,
            'TableStyleDark11' => TableStyleType::TableStyleDark11,
            'Custom' => TableStyleType::Custom,
        );
    }

    public static function SanitizeDisplayName($displayName)
    {
        if ($displayName === null || $displayName === '') {
            return $displayName;
        }
        return str_replace(' ', '_', $displayName);
    }
}