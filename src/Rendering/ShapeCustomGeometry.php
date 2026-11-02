<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Internal\Xml\XmlName;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKPoint;

// Numeric freeform paths preserved in drawing XML; guide formulas and arcs remain unsupported.
final class ShapeCustomGeometry
{
    const A = '{http://schemas.openxmlformats.org/drawingml/2006/main}';

    private function __construct()
    {
    }

    public static function Build($geometry, $rect)
    {
        $paths = $geometry->Element(self::A . 'pathLst');
        if ($paths === null) {
            return null;
        }
        $result = new SKPath();
        foreach ($paths->Elements(self::A . 'path') as $path) {
            $width = 0.0;
            $height = 0.0;
            if (!self::TryNumber($path, 'w', $width) || !self::TryNumber($path, 'h', $height) || $width <= 0.0 || $height <= 0.0) {
                $result->Dispose();
                return null;
            }
            foreach ($path->Elements() as $command) {
                $points = array();
                foreach ($command->Elements(self::A . 'pt') as $point) {
                    $x = 0.0;
                    $y = 0.0;
                    if (!self::TryNumber($point, 'x', $x) || !self::TryNumber($point, 'y', $y)) {
                        $result->Dispose();
                        return null;
                    }
                    $points[] = new SKPoint($rect->Left + (float)($x / $width * $rect->Width), $rect->Top + (float)($y / $height * $rect->Height));
                }
                list($namespace, $name) = XmlName::split($command->Name);
                if ($name === 'moveTo' && count($points) === 1) {
                    $result->MoveTo($points[0]);
                } elseif ($name === 'lnTo' && count($points) === 1) {
                    $result->LineTo($points[0]);
                } elseif ($name === 'quadBezTo' && count($points) === 2) {
                    $result->QuadTo($points[0], $points[1]);
                } elseif ($name === 'cubicBezTo' && count($points) === 3) {
                    $result->CubicTo($points[0]->X, $points[0]->Y, $points[1]->X, $points[1]->Y, $points[2]->X, $points[2]->Y);
                } elseif ($name === 'close' && count($points) === 0) {
                    $result->Close();
                } else {
                    $result->Dispose();
                    return null;
                }
            }
        }
        return $result;
    }

    private static function TryNumber($element, $name, &$value)
    {
        $text = (string)$element->Attribute($name);
        if ($text === '' || !is_numeric($text)) {
            $value = 0.0;
            return false;
        }

        $value = (float)$text;
        return !is_nan($value) && !is_infinite($value);
    }
}
