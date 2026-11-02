<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKPoint
{
    public $X = 0.0;
    public $Y = 0.0;

    public function __construct($x = 0.0, $y = 0.0)
    {
        $this->X = (float)$x;
        $this->Y = (float)$y;
    }

    public static function Empty()
    {
        return new SKPoint(0.0, 0.0);
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'IsEmpty') {
            return $this->X == 0.0 && $this->Y == 0.0;
        }

        return null;
    }
}
