<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKSize
{
    public $Width = 0.0;
    public $Height = 0.0;

    public function __construct($width = 0.0, $height = 0.0)
    {
        $this->Width = (float)$width;
        $this->Height = (float)$height;
    }

    public static function Empty()
    {
        return new SKSize(0.0, 0.0);
    }
}
