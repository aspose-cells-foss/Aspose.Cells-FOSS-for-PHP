<?php
namespace Aspose\Cells_FOSS\Core;

final class ColorValue
{
    public $A;
    public $R;
    public $G;
    public $B;
    public $ThemeIndex;
    public $Tint;
    public $Indexed;

    public function __construct(...$args)
    {
        if (count($args) === 4) {
            $a = $args[0];
            $r = $args[1];
            $g = $args[2];
            $b = $args[3];
            $this->A = $a;
            $this->R = $r;
            $this->G = $g;
            $this->B = $b;
            $this->ThemeIndex = null;
            $this->Tint = null;
            $this->Indexed = null;
            return;
        }
        if (count($args) === 2) {
            $themeIndex = $args[0];
            $tint = $args[1];
            $this->A = 0;
            $this->R = 0;
            $this->G = 0;
            $this->B = 0;
            $this->ThemeIndex = $themeIndex;
            $this->Tint = $tint;
            $this->Indexed = null;
            return;
        }
        if (count($args) === 1) {
            $indexed = $args[0];
            $this->A = 0;
            $this->R = 0;
            $this->G = 0;
            $this->B = 0;
            $this->ThemeIndex = null;
            $this->Tint = null;
            $this->Indexed = $indexed;
            return;
        }
        $this->A = 0;
        $this->R = 0;
        $this->G = 0;
        $this->B = 0;
        $this->ThemeIndex = null;
        $this->Tint = null;
        $this->Indexed = null;
    }

    public function Equals($other)
    {
        if (!($other instanceof ColorValue)) {
            return false;
        }
        return $this->A === $other->A && $this->R === $other->R && $this->G === $other->G && $this->B === $other->B
            && $this->ThemeIndex === $other->ThemeIndex && $this->Tint === $other->Tint && $this->Indexed === $other->Indexed;
    }

    public function EqualsObject($obj)
    {
        if ($obj instanceof ColorValue) {
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
        $themeIndex = ($this->ThemeIndex === null) ? 0 : $this->ThemeIndex;
        $hash = ($hash * 397) ^ $themeIndex;
        if ($this->Tint === null) {
            $hash = ($hash * 397) ^ 0;
        } else {
            $hash = ($hash * 397) ^ $this->Tint;
        }
        $indexed = ($this->Indexed === null) ? 0 : $this->Indexed;
        $hash = ($hash * 397) ^ $indexed;
        return $hash;
    }
}
