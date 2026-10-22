<?php

namespace Aspose\Cells_FOSS\CompareOpenXml;

final class OpenXmlBorderSideDescriptor
{
    private static $default;

    public function __construct($style, $color)
    {
        $this->Style = $style;
        $this->Color = $color;
    }

    public static function Default()
    {
        if (self::$default === null) {
            self::$default = new OpenXmlBorderSideDescriptor('None', ComparisonValueHelpers::DefaultColorHex);
        }

        return self::$default;
    }

    public $Style;
    public $Color;
}
