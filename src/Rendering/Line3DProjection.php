<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKRect;

// Category, value and depth coordinates all use the same perspective camera.
// Ribbon width is a depth interval, never a screen-space stroke width.
final class Line3DProjection
{
    private $sinX;

    private $cosX;

    private $sinY;

    private $cosY;

    private $height;

    private $depth;

    private $perspective;

    private $cameraOffsetX;

    private $cameraOffsetY;

    private $seriesCount;

    private $ribbonDepth;

    private $bounds;

    private $cameraBounds;

    public function __construct($plot, $chart, $categoryCount)
    {
        $rx = $chart->RotationXDeg * M_PI / 180.0;
        $ry = $chart->RotationYDeg * M_PI / 180.0;
        $this->sinX = sin($rx);
        $this->cosX = cos($rx);
        $this->sinY = sin($ry);
        $this->cosY = cos($ry);
        $this->seriesCount = max(1, count($chart->Series));
        $this->depth = $this->seriesCount / (float)max(1, $categoryCount)
            * max(0.01, $chart->DepthPercent / 100.0);
        // A deep series box must not flatten the value dimension or move
        // its near plane towards the camera singularity. Normalize both
        // height and viewing distance by the horizontal extent of the box.
        $extent = max(1.0, $this->depth);
        $this->height = max(0.05, $plot->Height / max(1.0, $plot->Width) * 0.76)
            * max(1.0, $this->depth * 0.875);
        $this->ribbonDepth = 1.0 / ($this->seriesCount * (1.0 + max(0.0, $chart->GapDepthPercent) / 100.0));
        $framing = $chart->RightAngleAxes ? 0.0
            : max(0.0, min(1.0, ($chart->PerspectiveDeg - 30.0) / 30.0));
        $this->perspective = $chart->RightAngleAxes ? 0.0
            : tan(min(120.0, max(0.0, $chart->PerspectiveDeg)) * M_PI / 360.0)
                / (0.475 * $extent * (1.0 + 0.22 * $framing));
        // Excel's stronger perspective is an off-axis view, with the
        // vanishing point inside the plot rather than above its near edge.
        $this->cameraOffsetX = 0.21 * $framing;
        $this->cameraOffsetY = -0.26 * $framing;
        $radius = abs($this->cosX * $this->sinY) * 0.5
            + abs($this->cosX * $this->cosY) * $this->depth * 0.5
            + abs($this->sinX) * $this->height * 0.5;
        $this->perspective = min($this->perspective, 0.9 / max(0.001, $radius));

        // Reserve the label bands around the complete projected box, not
        // around independently positioned walls and data series.
        $this->bounds = new SKRect($plot->Left + $plot->Width * (0.048 + (float)$framing * 0.109),
            $plot->Top + $plot->Height * (0.045 - (float)$framing * 0.072),
            $plot->Left + $plot->Width * (0.876 - (float)$framing * 0.131),
            $plot->Top + $plot->Height * (0.965 - (float)$framing * 0.05));
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
                / max(0.000001, $this->cameraBounds->Width) * $this->bounds->Width,
            $this->bounds->Top + ($point->Y - $this->cameraBounds->Top)
                / max(0.000001, $this->cameraBounds->Height) * $this->bounds->Height);
    }

    public function SeriesDepth($index)
    {
        return ($index + 0.5) / $this->seriesCount;
    }

    public function RibbonCorners($start, $end, $seriesIndex)
    {
        $center = $this->SeriesDepth($seriesIndex);
        $front = $center - $this->ribbonDepth * 0.5;
        $back = $center + $this->ribbonDepth * 0.5;
        return array(
            $this->Project($start->X, $start->Y, $front), $this->Project($end->X, $end->Y, $front),
            $this->Project($end->X, $end->Y, $back), $this->Project($start->X, $start->Y, $back),
        );
    }

    public function RibbonShade($start, $end, $corners)
    {
        $dx = $end->X - $start->X;
        $dy = ($end->Y - $start->Y) * $this->height;
        $length = sqrt($dx * $dx + $dy * $dy);
        if ($length < 0.000001) {
            return 0.76;
        }

        $edgeX = $corners[1]->X - $corners[0]->X;
        $edgeY = $corners[1]->Y - $corners[0]->Y;
        $depthX = $corners[3]->X - $corners[0]->X;
        $depthY = $corners[3]->Y - $corners[0]->Y;
        // A steep rising segment can expose the underside of the ribbon.
        // The projected winding distinguishes it from an upward-facing face.
        if ($edgeX * $depthY - $edgeY * $depthX > 0.0) {
            return 0.40;
        }

        $normalX = -$dy / $length;
        return max(0.4, min(0.81, 0.76 + 0.25 * $normalX));
    }

    private function CameraPoint($category, $value, $depth)
    {
        $x = $category - 0.5;
        $y = ($value - 0.5) * $this->height;
        $z = ($depth - 0.5) * $this->depth;
        $yawDepth = -$this->sinY * $x + $this->cosY * $z;
        $cameraDepth = $this->cosX * $yawDepth - $this->sinX * $y;
        $divisor = 1.0 + $this->perspective * $cameraDepth;
        return new SKPoint((float)(($this->cosY * $x + $this->sinY * $z + $this->cameraOffsetX) / $divisor),
            (float)((-($this->cosX * $y + $this->sinX * $yawDepth) + $this->cameraOffsetY) / $divisor));
    }
}
