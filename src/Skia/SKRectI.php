<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKRectI
{
    public $Left = 0;
    public $Top = 0;
    public $Right = 0;
    public $Bottom = 0;

    public function __construct($left = 0, $top = 0, $right = 0, $bottom = 0)
    {
        $this->Left = (int)$left;
        $this->Top = (int)$top;
        $this->Right = (int)$right;
        $this->Bottom = (int)$bottom;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Width') {
            return $this->Right - $this->Left;
        }

        if ($propertyName === 'Height') {
            return $this->Bottom - $this->Top;
        }

        if ($propertyName === 'MidX') {
            return intdiv($this->Left + $this->Right, 2);
        }

        if ($propertyName === 'MidY') {
            return intdiv($this->Top + $this->Bottom, 2);
        }

        if ($propertyName === 'IsEmpty') {
            return $this->Right <= $this->Left || $this->Bottom <= $this->Top;
        }

        return null;
    }

    public function toRect()
    {
        return new SKRect($this->Left, $this->Top, $this->Right, $this->Bottom);
    }
}
