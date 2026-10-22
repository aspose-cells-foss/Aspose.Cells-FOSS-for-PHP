<?php

namespace Aspose\Cells_FOSS\CompareOpenXml;

final class OpenXmlStylesContext
{
    private static $default;

    public function __construct($cellFormats, $dateStyleIndexes)
    {
        $this->CellFormats = $cellFormats;
        $this->DateStyleIndexes = $dateStyleIndexes;
    }

    public static function Default()
    {
        if (self::$default === null) {
            self::$default = new OpenXmlStylesContext(array(StyleSnapshot::Default()), array());
        }

        return self::$default;
    }

    public $CellFormats;
    public $DateStyleIndexes;
}
