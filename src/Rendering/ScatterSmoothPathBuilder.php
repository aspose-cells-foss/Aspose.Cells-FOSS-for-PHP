<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKPoint;

final class ScatterSmoothPathBuilder
{
    private function __construct()
    {
    }

    public static function Append($path, $points)
    {
        for ($index = 0; $index < count($points) - 1; $index++) {
            $start = $points[$index];
            $end = $points[$index + 1];
            $before = $index > 0 ? $points[$index - 1] : $start;
            $after = $index + 2 < count($points) ? $points[$index + 2] : $end;
            $firstDivisor = $index == 0 ? 3.0 : 6.0;
            $secondDivisor = $index + 2 == count($points) ? 3.0 : 6.0;
            $control1 = new SKPoint(
                $start->X + ($end->X - $before->X) / $firstDivisor,
                $start->Y + ($end->Y - $before->Y) / $firstDivisor);
            $control2 = new SKPoint(
                $end->X - ($after->X - $start->X) / $secondDivisor,
                $end->Y - ($after->Y - $start->Y) / $secondDivisor);
            $path->CubicTo($control1->X, $control1->Y, $control2->X, $control2->Y, $end->X, $end->Y);
        }
    }
}
