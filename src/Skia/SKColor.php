<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKColor
{
    public $Red = 0;
    public $Green = 0;
    public $Blue = 0;
    public $Alpha = 255;

    public function __construct($red = 0, $green = 0, $blue = 0, $alpha = 255)
    {
        $this->Red = self::clamp($red);
        $this->Green = self::clamp($green);
        $this->Blue = self::clamp($blue);
        $this->Alpha = self::clamp($alpha);
    }

    private static function clamp($value)
    {
        $value = (int)$value;
        if ($value < 0) {
            return 0;
        }

        if ($value > 255) {
            return 255;
        }

        return $value;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'R') {
            return $this->Red;
        }

        if ($propertyName === 'G') {
            return $this->Green;
        }

        if ($propertyName === 'B') {
            return $this->Blue;
        }

        if ($propertyName === 'A') {
            return $this->Alpha;
        }

        return null;
    }

    public function toArgb()
    {
        return SkiaLibrary::colorToNative($this);
    }

    public function WithAlpha($alpha)
    {
        return new SKColor($this->Red, $this->Green, $this->Blue, $alpha);
    }

    public function ToHsv(&$hue, &$saturation, &$value)
    {
        $r = $this->Red / 255.0;
        $g = $this->Green / 255.0;
        $b = $this->Blue / 255.0;
        $max = max($r, $g, $b);
        $min = min($r, $g, $b);
        $delta = $max - $min;

        $value = $max * 100.0;
        $saturation = $max <= 0.0 ? 0.0 : ($delta / $max) * 100.0;

        if ($delta <= 0.0) {
            $hue = 0.0;
            return;
        }

        if ($max === $r) {
            $hue = 60.0 * fmod(($g - $b) / $delta, 6.0);
        } elseif ($max === $g) {
            $hue = 60.0 * (((($b - $r) / $delta) + 2.0));
        } else {
            $hue = 60.0 * (((($r - $g) / $delta) + 4.0));
        }

        if ($hue < 0.0) {
            $hue += 360.0;
        }
    }

    public static function FromHsv($hue, $saturation, $value, $alpha = 255)
    {
        $hue = fmod((float)$hue, 360.0);
        if ($hue < 0.0) {
            $hue += 360.0;
        }

        $saturation = (float)$saturation;
        if ($saturation < 0.0) {
            $saturation = 0.0;
        }

        if ($saturation > 100.0) {
            $saturation = 100.0;
        }

        $value = (float)$value;
        if ($value < 0.0) {
            $value = 0.0;
        }

        if ($value > 100.0) {
            $value = 100.0;
        }

        $s = $saturation / 100.0;
        $v = $value / 100.0;
        $h = $hue / 60.0;
        $i = (int)floor($h);
        $f = $h - $i;
        $p = $v * (1.0 - $s);
        $q = $v * (1.0 - ($s * $f));
        $t = $v * (1.0 - ($s * (1.0 - $f)));

        switch ($i % 6) {
            case 0:
                $r = $v;
                $g = $t;
                $b = $p;
                break;
            case 1:
                $r = $q;
                $g = $v;
                $b = $p;
                break;
            case 2:
                $r = $p;
                $g = $v;
                $b = $t;
                break;
            case 3:
                $r = $p;
                $g = $q;
                $b = $v;
                break;
            case 4:
                $r = $t;
                $g = $p;
                $b = $v;
                break;
            default:
                $r = $v;
                $g = $p;
                $b = $q;
                break;
        }

        return new SKColor(
            (int)round($r * 255.0),
            (int)round($g * 255.0),
            (int)round($b * 255.0),
            $alpha);
    }

    public function ToString()
    {
        return '#'
            . str_pad(dechex($this->Red), 2, '0', STR_PAD_LEFT)
            . str_pad(dechex($this->Green), 2, '0', STR_PAD_LEFT)
            . str_pad(dechex($this->Blue), 2, '0', STR_PAD_LEFT);
    }
}
