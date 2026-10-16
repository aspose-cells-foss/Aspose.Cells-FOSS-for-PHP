<?php

namespace Aspose\Cells_FOSS;

final class WorkbookPropertySupport
{
    private function __construct()
    {
    }

    public static function NormalizeShowObjects($value)
    {
        return self::NormalizeChoice($value, 'showObjects', 'all', 'placeholders', 'none');
    }

    public static function NormalizeUpdateLinks($value)
    {
        return self::NormalizeChoice($value, 'updateLinks', 'userSet', 'never', 'always');
    }

    public static function NormalizeVisibility($value)
    {
        return self::NormalizeChoice($value, 'visibility', 'visible', 'hidden', 'veryHidden');
    }

    public static function NormalizeCalculationMode($value)
    {
        return self::NormalizeChoice($value, 'calcMode', 'auto', 'manual', 'autoNoTable');
    }

    public static function NormalizeReferenceMode($value)
    {
        return self::NormalizeChoice($value, 'refMode', 'A1', 'R1C1');
    }

    private static function NormalizeChoice($value, $propertyName, ...$allowed)
    {
        if ($value === null || trim($value) === '') {
            return '';
        }

        $trimmed = trim($value);
        for ($index = 0; $index < count($allowed); $index++) {
            if (strcasecmp($allowed[$index], $trimmed) === 0) {
                return $allowed[$index];
            }
        }

        throw new CellsException("Unsupported " . $propertyName . " value '" . $value . "'.");
    }
}
