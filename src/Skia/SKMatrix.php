<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKMatrix
{
    public $ScaleX = 1.0;
    public $SkewX = 0.0;
    public $TransX = 0.0;
    public $SkewY = 0.0;
    public $ScaleY = 1.0;
    public $TransY = 0.0;
    public $Persp0 = 0.0;
    public $Persp1 = 0.0;
    public $Persp2 = 1.0;

    public function __construct($scaleX = 1.0, $skewX = 0.0, $transX = 0.0, $skewY = 0.0, $scaleY = 1.0, $transY = 0.0)
    {
        $this->ScaleX = (float)$scaleX;
        $this->SkewX = (float)$skewX;
        $this->TransX = (float)$transX;
        $this->SkewY = (float)$skewY;
        $this->ScaleY = (float)$scaleY;
        $this->TransY = (float)$transY;
    }

    public static function CreateIdentity()
    {
        return new SKMatrix();
    }

    public static function CreateScale($scaleX, $scaleY = null)
    {
        if ($scaleY === null) {
            $scaleY = $scaleX;
        }

        return new SKMatrix($scaleX, 0.0, 0.0, 0.0, $scaleY, 0.0);
    }

    public static function CreateTranslation($x, $y)
    {
        return new SKMatrix(1.0, 0.0, $x, 0.0, 1.0, $y);
    }

    public static function CreateRotationDegrees($degrees)
    {
        $radians = deg2rad((float)$degrees);
        $c = cos($radians);
        $s = sin($radians);
        return new SKMatrix($c, -$s, 0.0, $s, $c, 0.0);
    }

    public function toNative()
    {
        $ffi = SkiaLibrary::ffi();
        $matrix = $ffi->new('sk_matrix_t');
        $matrix->scaleX = $this->ScaleX;
        $matrix->skewX = $this->SkewX;
        $matrix->transX = $this->TransX;
        $matrix->skewY = $this->SkewY;
        $matrix->scaleY = $this->ScaleY;
        $matrix->transY = $this->TransY;
        $matrix->persp0 = $this->Persp0;
        $matrix->persp1 = $this->Persp1;
        $matrix->persp2 = $this->Persp2;
        return $matrix;
    }
}
