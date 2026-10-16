<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ColorValue;

// Represents color.
class Color
{
    public $A;
    public $R;
    public $G;
    public $B;

    public function __construct($a, $r, $g, $b)
    {
        $this->A = $a;
        $this->R = $r;
        $this->G = $g;
        $this->B = $b;
    }

    public static function Empty()
    {
        return new Color(0, 0, 0, 0);
    }

    public static function FromArgb($a, $r, $g, $b)
    {
        return new Color($a, $r, $g, $b);
    }

    public function Equals($other)
    {
        if (!($other instanceof Color)) {
            return false;
        }

        return $this->A === $other->A && $this->R === $other->R && $this->G === $other->G && $this->B === $other->B;
    }

    public function EqualsObject($obj)
    {
        if ($obj instanceof Color) {
            return $this->Equals($obj);
        }

        return false;
    }

    public function GetHashCode()
    {
        $hash = $this->A;
        $hash = ($hash * 397) ^ $this->R;
        $hash = ($hash * 397) ^ $this->G;
        $hash = ($hash * 397) ^ $this->B;
        return $hash;
    }

    public function ToCore()
    {
        return new ColorValue($this->A, $this->R, $this->G, $this->B);
    }

    public static function FromCore($value)
    {
        if ($value === null) {
            return Color::Empty();
        }

        return new Color($value->A, $value->R, $value->G, $value->B);
    }
}
