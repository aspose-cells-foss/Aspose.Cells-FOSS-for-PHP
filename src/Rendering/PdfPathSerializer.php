<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKPathVerb;
use Aspose\Cells_FOSS\Skia\SKPoint;

final class PdfPathSerializer
{
    private function __construct()
    {
    }

    /**
     * Serializes an SKPath into PDF path-painting operators (m/l/c/v/y/h).
     */
    public static function Serialize($path)
    {
        if ($path === null) {
            return '';
        }

        $builder = '';
        $currentX = 0.0;
        $currentY = 0.0;
        $hasCurrentPoint = false;

        $iterator = $path->CreateIterator(false);
        try {
            while (true) {
                $points = array();
                $verb = $iterator->Next($points);
                if ($verb === SKPathVerb::Done) {
                    break;
                }

                if ($verb === SKPathVerb::Move) {
                    $builder .= PdfNumbers::oneDecimal($points[0]->X) . ' '
                        . PdfNumbers::oneDecimal($points[0]->Y) . " m\n";
                    $currentX = $points[0]->X;
                    $currentY = $points[0]->Y;
                    $hasCurrentPoint = true;
                } elseif ($verb === SKPathVerb::Line) {
                    $builder .= PdfNumbers::oneDecimal($points[1]->X) . ' '
                        . PdfNumbers::oneDecimal($points[1]->Y) . " l\n";
                    $currentX = $points[1]->X;
                    $currentY = $points[1]->Y;
                    $hasCurrentPoint = true;
                } elseif ($verb === SKPathVerb::Quad) {
                    $c1x = $points[0]->X + ($points[1]->X - $points[0]->X) * 2.0 / 3.0;
                    $c1y = $points[0]->Y + ($points[1]->Y - $points[0]->Y) * 2.0 / 3.0;
                    $c2x = $points[2]->X + ($points[1]->X - $points[2]->X) * 2.0 / 3.0;
                    $c2y = $points[2]->Y + ($points[1]->Y - $points[2]->Y) * 2.0 / 3.0;
                    $builder .= self::appendCubic($currentX, $currentY, $hasCurrentPoint, $c1x, $c1y, $c2x, $c2y, $points[2]->X, $points[2]->Y);
                    $currentX = $points[2]->X;
                    $currentY = $points[2]->Y;
                    $hasCurrentPoint = true;
                } elseif ($verb === SKPathVerb::Conic) {
                    // Convert conic to cubic using the standard midpoint subdivision approximation.
                    $w = $iterator->ConicWeight();
                    $c1x = $points[0]->X + ($points[1]->X - $points[0]->X) * $w * 2.0 / 3.0;
                    $c1y = $points[0]->Y + ($points[1]->Y - $points[0]->Y) * $w * 2.0 / 3.0;
                    $c2x = $points[2]->X + ($points[1]->X - $points[2]->X) * $w * 2.0 / 3.0;
                    $c2y = $points[2]->Y + ($points[1]->Y - $points[2]->Y) * $w * 2.0 / 3.0;
                    $builder .= self::appendCubic($currentX, $currentY, $hasCurrentPoint, $c1x, $c1y, $c2x, $c2y, $points[2]->X, $points[2]->Y);
                    $currentX = $points[2]->X;
                    $currentY = $points[2]->Y;
                    $hasCurrentPoint = true;
                } elseif ($verb === SKPathVerb::Cubic) {
                    $builder .= self::appendCubic($currentX, $currentY, $hasCurrentPoint, $points[1]->X, $points[1]->Y, $points[2]->X, $points[2]->Y, $points[3]->X, $points[3]->Y);
                    $currentX = $points[3]->X;
                    $currentY = $points[3]->Y;
                    $hasCurrentPoint = true;
                } elseif ($verb === SKPathVerb::Close) {
                    $builder .= "h\n";
                    $hasCurrentPoint = false;
                }
            }
        } finally {
            $iterator->Dispose();
        }

        return $builder;
    }

    private static function appendCubic($currentX, $currentY, $hasCurrentPoint, $c1x, $c1y, $c2x, $c2y, $x, $y)
    {
        $f = __NAMESPACE__ . '\PdfNumbers::oneDecimal';
        if ($hasCurrentPoint && self::sameCoordinate($c1x, $currentX) && self::sameCoordinate($c1y, $currentY)) {
            return $f($c2x) . ' ' . $f($c2y) . ' ' . $f($x) . ' ' . $f($y) . " v\n";
        }

        if (self::sameCoordinate($c2x, $x) && self::sameCoordinate($c2y, $y)) {
            return $f($c1x) . ' ' . $f($c1y) . ' ' . $f($x) . ' ' . $f($y) . " y\n";
        }

        return $f($c1x) . ' ' . $f($c1y) . ' ' . $f($c2x) . ' ' . $f($c2y) . ' ' . $f($x) . ' ' . $f($y) . " c\n";
    }

    private static function sameCoordinate($left, $right)
    {
        return abs($left - $right) <= 0.05;
    }
}
