<?php

namespace Aspose\Cells_FOSS\CompareOpenXml;

final class OpenXmlBorderDescriptor
{
    private static $default;

    public function __construct($left, $right, $top, $bottom, $diagonal, $diagonalUp, $diagonalDown)
    {
        $this->Left = $left;
        $this->Right = $right;
        $this->Top = $top;
        $this->Bottom = $bottom;
        $this->Diagonal = $diagonal;
        $this->DiagonalUp = $diagonalUp;
        $this->DiagonalDown = $diagonalDown;
    }

    public static function Default()
    {
        if (self::$default === null) {
            self::$default = new OpenXmlBorderDescriptor(OpenXmlBorderSideDescriptor::Default(), OpenXmlBorderSideDescriptor::Default(), OpenXmlBorderSideDescriptor::Default(), OpenXmlBorderSideDescriptor::Default(), OpenXmlBorderSideDescriptor::Default(), false, false);
        }

        return self::$default;
    }

    public $Left;
    public $Right;
    public $Top;
    public $Bottom;
    public $Diagonal;
    public $DiagonalUp;
    public $DiagonalDown;
}
