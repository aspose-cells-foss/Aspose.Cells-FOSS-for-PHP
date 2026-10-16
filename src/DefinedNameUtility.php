<?php

namespace Aspose\Cells_FOSS;

final class DefinedNameUtility
{
    const PrintAreaDefinedName = '_xlnm.Print_Area';
    const PrintTitlesDefinedName = '_xlnm.Print_Titles';
    const FilterDatabaseDefinedName = '_xlnm._FilterDatabase';

    private function __construct()
    {
    }

    public static function IsReservedName($name)
    {
        return strcasecmp($name, self::PrintAreaDefinedName) === 0
            || strcasecmp($name, self::PrintTitlesDefinedName) === 0
            || strcasecmp($name, self::FilterDatabaseDefinedName) === 0;
    }

    public static function NormalizeName($name)
    {
        $normalized = trim($name === null ? '' : $name);
        if (strlen($normalized) === 0) {
            throw new CellsException('Defined name must be non-empty.');
        }

        if (self::IsReservedName($normalized)) {
            throw new CellsException('Built-in print defined names must be managed through PageSetup.');
        }

        return $normalized;
    }

    public static function NormalizeFormula($formula)
    {
        $normalized = trim($formula === null ? '' : $formula);
        if (str_starts_with($normalized, '=')) {
            $normalized = trim(substr($normalized, 1));
        }

        if (strlen($normalized) === 0) {
            throw new CellsException('Defined name formula must be non-empty.');
        }

        return $normalized;
    }

    public static function NormalizeComment($comment)
    {
        return trim($comment === null ? '' : $comment);
    }

    public static function SameScope($left, $right)
    {
        if ($left === null && $right === null) {
            return true;
        }

        return $left !== null && $right !== null && $left === $right;
    }
}
