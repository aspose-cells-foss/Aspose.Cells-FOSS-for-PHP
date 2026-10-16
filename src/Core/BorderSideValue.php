<?php
namespace Aspose\Cells_FOSS\Core;

final class BorderSideValue
{
    public $Style;
    public $Color;

    public function __construct()
    {
        $this->Style = BorderStyle::None;
        $this->Color = new ColorValue();
    }

    public function Clone()
    {
        $clone = new BorderSideValue();
        $clone->Style = $this->Style;
        $clone->Color = ($this->Color === null) ? new ColorValue() : clone $this->Color;
        return $clone;
    }
}
