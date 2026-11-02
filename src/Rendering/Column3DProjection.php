<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKMatrix;
use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKRect;
use Aspose\Cells_FOSS\Skia\SKSize;

// Coordinates are category [0,1], value [0,1], and depth [0,1].
// Depth zero is the front plane. All scene vertices use this camera.
final class Column3DProjection
{
    private $sinX;

    private $cosX;

    private $sinY;

    private $cosY;

    private $height;

    private $depth;

    private $perspective;

    private $bounds;

    private $cameraBounds;

    private $unitWidth;

    private $cameraOffsetX;

    private $cameraOffsetY;

    public function __construct($chartArea, $chart, $categoryCount, $labelReserve = 0.0)
    {
        $rx = $chart->RotationXDeg * M_PI / 180.0;
        $ry = $chart->RotationYDeg * M_PI / 180.0;
        $this->sinX = sin($rx);
        $this->cosX = cos($rx);
        $this->sinY = sin($ry);
        $this->cosY = cos($ry);
        $this->height = max(0.05, ($chartArea->Height - $labelReserve) / $chartArea->Width * 0.6);
        $this->depth = max(1, count($chart->Series)) / (float)max(1, $categoryCount)
            * max(0.01, $chart->DepthPercent / 100.0);
        $perspective = $chart->RightAngleAxes ? 0.0 : max(0.0, $chart->PerspectiveDeg);
        // The high-perspective Excel frame is an off-axis view: the camera's
        // vanishing point stays inside the plot instead of the near box corner.
        // Interpolate the framing correction from the default 30-degree view.
        $framing = max(0.0, min(1.0, ($perspective - 30.0) / 70.0));
        $this->perspective = tan(min(120.0, $perspective) * M_PI / 360.0)
            / (0.48 + 0.29 * $framing);
        $this->cameraOffsetX = -0.14 * $framing;
        $this->cameraOffsetY = 0.0925 * $framing;
        $this->unitWidth = $chartArea->Width * 0.62;
        // Keep the entire box in front of the camera even for extreme settings.
        $radius = abs($this->cosX * $this->sinY) * 0.5
            + abs($this->cosX * $this->cosY) * $this->depth * 0.5
            + abs($this->sinX) * $this->height * 0.5;
        $this->perspective = min($this->perspective, 0.9 / max(0.001, $radius));
        $this->bounds = new SKRect($chartArea->Left + $chartArea->Width * (0.158 + (float)$framing * 0.006),
            $chartArea->Top + $chartArea->Height * (0.19 + (float)$framing * 0.011),
            $chartArea->Left + $chartArea->Width * (0.85 - (float)$framing * 0.054),
            $chartArea->Top + $chartArea->Height * 0.752 - $labelReserve * 0.34);

        $minX = PHP_FLOAT_MAX;
        $minY = PHP_FLOAT_MAX;
        $maxX = -PHP_FLOAT_MAX;
        $maxY = -PHP_FLOAT_MAX;
        for ($x = 0; $x <= 1; $x++) {
            for ($y = 0; $y <= 1; $y++) {
                for ($z = 0; $z <= 1; $z++) {
                    $point = $this->CameraPoint($x, $y, $z);
                    $minX = min($minX, $point->X);
                    $minY = min($minY, $point->Y);
                    $maxX = max($maxX, $point->X);
                    $maxY = max($maxY, $point->Y);
                }
            }
        }
        $this->cameraBounds = new SKRect($minX, $minY, $maxX, $maxY);
    }

    public function Project($category, $value, $depth)
    {
        $point = $this->CameraPoint($category, $value, $depth);
        return new SKPoint($this->bounds->Left + ($point->X - $this->cameraBounds->Left)
                / $this->cameraBounds->Width * $this->bounds->Width,
            $this->bounds->Top + ($point->Y - $this->cameraBounds->Top)
                / $this->cameraBounds->Height * $this->bounds->Height);
    }

    public function CameraDepth($category, $value, $depth)
    {
        $yawDepth = -$this->sinY * ($category - 0.5)
            + $this->cosY * ($depth - 0.5) * $this->depth;
        return $this->cosX * $yawDepth - $this->sinX * ($value - 0.5) * $this->height;
    }

    public static function ColumnWidth($chart, $categoryCount)
    {
        return 1.0 / (max(1, $categoryCount)
            * (1.0 + max(0.0, $chart->GapWidthPercent) / 100.0));
    }

    public static function ColumnDepth($chart)
    {
        return 1.0 / (max(1, count($chart->Series))
            * (1.0 + max(0.0, $chart->GapDepthPercent) / 100.0));
    }

    public function WallSize($side)
    {
        return new SKSize((float)($this->unitWidth * ($side ? $this->depth : 1.0)),
            (float)($this->unitWidth * $this->height));
    }

    public function WallTransform($side)
    {
        $a = $side ? $this->Project(0, 1, 0) : $this->Project(0, 1, 1);
        $b = $side ? $this->Project(0, 1, 1) : $this->Project(1, 1, 1);
        $c = $side ? $this->Project(0, 0, 1) : $this->Project(1, 0, 1);
        $d = $side ? $this->Project(0, 0, 0) : $this->Project(0, 0, 1);
        $dx1 = $b->X - $c->X;
        $dx2 = $d->X - $c->X;
        $dy1 = $b->Y - $c->Y;
        $dy2 = $d->Y - $c->Y;
        $dx3 = $a->X - $b->X + $c->X - $d->X;
        $dy3 = $a->Y - $b->Y + $c->Y - $d->Y;
        $determinant = $dx1 * $dy2 - $dx2 * $dy1;
        if (abs($determinant) < 0.00001) {
            return SKMatrix::CreateIdentity();
        }
        $g = ($dx3 * $dy2 - $dx2 * $dy3) / $determinant;
        $h = ($dx1 * $dy3 - $dx3 * $dy1) / $determinant;
        $size = $this->WallSize($side);
        $matrix = new SKMatrix(
            ($b->X - $a->X + $g * $b->X) / $size->Width,
            ($d->X - $a->X + $h * $d->X) / $size->Height,
            $a->X,
            ($b->Y - $a->Y + $g * $b->Y) / $size->Width,
            ($d->Y - $a->Y + $h * $d->Y) / $size->Height,
            $a->Y);
        $matrix->Persp0 = $g / $size->Width;
        $matrix->Persp1 = $h / $size->Height;
        $matrix->Persp2 = 1.0;
        return $matrix;
    }

    public function ColumnCorners($center, $seriesCenter, $width, $depth, $bottom, $top)
    {
        $x0 = $center - $width * 0.5;
        $x1 = $center + $width * 0.5;
        $z0 = $seriesCenter - $depth * 0.5;
        $z1 = $seriesCenter + $depth * 0.5;
        return array(
            $this->Project($x0, $bottom, $z0), $this->Project($x1, $bottom, $z0),
            $this->Project($x1, $bottom, $z1), $this->Project($x0, $bottom, $z1),
            $this->Project($x0, $top, $z0), $this->Project($x1, $top, $z0),
            $this->Project($x1, $top, $z1), $this->Project($x0, $top, $z1),
        );
    }

    private function CameraPoint($category, $value, $depth)
    {
        $x = $category - 0.5;
        $y = ($value - 0.5) * $this->height;
        $z = ($depth - 0.5) * $this->depth;
        $yawDepth = -$this->sinY * $x + $this->cosY * $z;
        $denominator = 1.0 + $this->perspective * $this->CameraDepth($category, $value, $depth);
        return new SKPoint((float)(($this->cosY * $x + $this->sinY * $z + $this->cameraOffsetX) / $denominator),
            (float)(-($this->cosX * $y + $this->sinX * $yawDepth + $this->cameraOffsetY) / $denominator));
    }
}
