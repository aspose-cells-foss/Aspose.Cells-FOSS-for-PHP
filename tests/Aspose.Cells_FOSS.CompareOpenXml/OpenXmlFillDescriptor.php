<?php

namespace Aspose\Cells_FOSS\CompareOpenXml;

final class OpenXmlFillDescriptor
{
    private static $default;

    public function __construct($pattern, $foregroundColor, $backgroundColor)
    {
        $this->Pattern = $pattern;
        $this->ForegroundColor = $foregroundColor;
        $this->BackgroundColor = $backgroundColor;
    }

    public static function Default()
    {
        if (self::$default === null) {
            self::$default = new OpenXmlFillDescriptor('None', ComparisonValueHelpers::DefaultColorHex, ComparisonValueHelpers::DefaultColorHex);
        }

        return self::$default;
    }

    public $Pattern;
    public $ForegroundColor;
    public $BackgroundColor;
}
