<?php

namespace Aspose\Cells_FOSS\CompareOpenXml;

final class OpenXmlFontDescriptor
{
    private static $default;

    public function __construct($name, $size, $bold, $italic, $underline, $strikeThrough, $color)
    {
        $this->Name = $name;
        $this->Size = $size;
        $this->Bold = $bold;
        $this->Italic = $italic;
        $this->Underline = $underline;
        $this->StrikeThrough = $strikeThrough;
        $this->Color = $color;
    }

    public static function Default()
    {
        if (self::$default === null) {
            self::$default = new OpenXmlFontDescriptor('Calibri', ComparisonValueHelpers::NormalizeDouble(11.0), false, false, false, false, ComparisonValueHelpers::DefaultColorHex);
        }

        return self::$default;
    }

    public $Name;
    public $Size;
    public $Bold;
    public $Italic;
    public $Underline;
    public $StrikeThrough;
    public $Color;
}
