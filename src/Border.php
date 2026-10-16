<?php

namespace Aspose\Cells_FOSS;

// Represents border.
class Border
{
    public $LineStyle;
    public $Color;

    public function __construct()
    {
        $this->Color = Color::Empty();
    }

    public function Clone()
    {
        $clone = new Border();
        $clone->LineStyle = $this->LineStyle;
        $clone->Color = $this->Color;
        return $clone;
    }
}
