<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\DateSystem;
use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\CultureInfo;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKData;
use Aspose\Cells_FOSS\Skia\SKFilterQuality;
use Aspose\Cells_FOSS\Skia\SKFont;
use Aspose\Cells_FOSS\Skia\SKImage;
use Aspose\Cells_FOSS\Skia\SKMatrix;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKRect;
use Aspose\Cells_FOSS\Skia\SKShaderTileMode;

final class Column3DSceneRenderer
{
    private $context;

    private $dateSystem;

    private $culture;

    public function __construct($context, $dateSystem, $culture)
    {
        $this->context = $context;
        $this->dateSystem = $dateSystem;
        $this->culture = $culture;
    }

    public function Draw($canvas, $area, $chart, $min, $max, $major)
    {
        $count = count($chart->Categories);
        for ($i = 0; $i < count($chart->Series); $i++) {
            $count = max($count, count($chart->Series[$i]->Values));
        }

        if ($count == 0) {
            return;
        }

        $categories = new Column3DCategoryLayout($chart, $count, $this->dateSystem, $this->culture);
        $rotation = ChartDateAxisLayout::Resolve($chart, $this->dateSystem, $area->Width * 0.69)->RotationDeg;
        $labelReserve = abs($rotation) > 45.0 ? $area->Height * 0.25 : 0.0;
        // Empty date slots change the column width, not the series-depth scale.
        $projection = new Column3DProjection($area, $chart, $count, $labelReserve);
        if ($chart->MajorUnit === null && $major > 0.0) {
            $axisHeight = abs($projection->Project(0, 1, 0)->Y
                - $projection->Project(0, 0, 0)->Y);
            while (($max - $min) / $major > 1.0 && $axisHeight * $major / ($max - $min) < 18.0) {
                $major *= 2.0;
            }
        }

        $canvas->Save();
        $canvas->ClipRect($area);
        self::DrawWalls($canvas, $projection, $chart, $min, $max, $major, $categories->SlotCount);
        self::DrawColumns($canvas, $projection, $chart, $min, $max, $categories);
        $this->DrawLabels($canvas, $projection, $chart, $min, $max, $major, $categories, $area->Width);
        $canvas->Restore();
    }

    private static function DrawWalls($canvas, $p, $chart, $min, $max, $major, $count)
    {
        self::DrawWallTexture($canvas, $p, true, $chart->SideWallImageData);
        self::DrawWallTexture($canvas, $p, false, $chart->BackWallImageData);
        $grid = new SKPaint();
        $grid->Color = new SKColor(0xD9, 0xD9, 0xD9);
        $grid->StrokeWidth = 0.6;
        $grid->Style = SKPaintStyle::Stroke;
        $grid->IsAntialias = true;

        if ($chart->HasMinorValueGridlines) {
            $minor = $chart->MinorUnit !== null ? $chart->MinorUnit : $major / 5.0;
            if ($minor > 0.0) {
                $grid->Color = $chart->MinorValueGridlineColor;
                $grid->StrokeWidth = (float)$chart->MinorValueGridlineWidthPt;
                self::DrawGridLevels($canvas, $p, $grid, $min, $max, $minor);
            }
        }

        $grid->Color = new SKColor(0xD9, 0xD9, 0xD9);
        $grid->StrokeWidth = 0.6;
        if ($chart->HasValueGridlines) {
            self::DrawGridLevels($canvas, $p, $grid, $min, $max, $major);
        }

        $canvas->DrawLine($p->Project(0, 0, 1), $p->Project(0, 0, 0), $grid);
        $canvas->DrawLine($p->Project(0, 0, 0), $p->Project(1, 0, 0), $grid);
        for ($i = 0; $i <= $count; $i++) {
            $point = $p->Project($i / $count, 0, 0);
            $canvas->DrawLine($point->X, $point->Y, $point->X, $point->Y + 2.8, $grid);
        }

        if ($chart->HasValueAxisLine) {
            $grid->Color = $chart->ValueAxisLineColor;
            $grid->StrokeWidth = (float)$chart->ValueAxisLineWidthPt;
            $canvas->DrawLine($p->Project(0, 0, 0), $p->Project(0, 1, 0), $grid);
            if ($chart->ValueAxisMajorTickMark === 'cross') {
                for ($value = $min; $value <= $max + $major * 0.001; $value += $major) {
                    $point = $p->Project(0, ($value - $min) / ($max - $min), 0);
                    $canvas->DrawLine($point->X - 3.0, $point->Y, $point->X + 3.0, $point->Y, $grid);
                }
            }
        }

        $grid->Dispose();
    }

    private static function DrawWallTexture($canvas, $p, $side, $imageData)
    {
        if ($imageData === null || strlen((string)$imageData) == 0) {
            return;
        }

        $data = SKData::CreateCopy($imageData);
        $image = SKImage::FromEncodedData($data);
        $paint = new SKPaint();
        $paint->IsAntialias = true;
        $paint->FilterQuality = SKFilterQuality::High;

        if ($image === null) {
            $paint->Dispose();
            $data->Dispose();
            return;
        }

        $size = $p->WallSize($side);
        $matrix = $p->WallTransform($side);
        $canvas->Save();
        $canvas->Concat($matrix);
        $shader = $image->MakeShader(SKShaderTileMode::Repeat,
            SKShaderTileMode::Repeat, SKMatrix::CreateScale(0.75, 0.75));
        $paint->Shader = $shader;
        $canvas->DrawRect(new SKRect(0, 0, $size->Width, $size->Height), $paint);
        $paint->Shader = null;
        $shader->Dispose();
        if ($side) {
            $paint->Color = new SKColor(0, 0, 0, 71);
            $canvas->DrawRect(new SKRect(0, 0, $size->Width, $size->Height), $paint);
        }

        $canvas->Restore();
        $paint->Dispose();
        $image->Dispose();
        $data->Dispose();
    }

    private static function DrawGridLevels($canvas, $p, $paint, $min, $max, $interval)
    {
        if ($interval <= 0.0) {
            return;
        }

        $count = min(1000, (int)round(($max - $min) / $interval));
        for ($i = 0; $i <= $count; $i++) {
            $y = $i * $interval / ($max - $min);
            $corner = $p->Project(0, $y, 1);
            $canvas->DrawLine($p->Project(0, $y, 0), $corner, $paint);
            $canvas->DrawLine($corner, $p->Project(1, $y, 1), $paint);
        }
    }

    private static function DrawColumns($canvas, $p, $chart, $min, $max, $categories)
    {
        $count = count($categories->Centers);
        $width = Column3DProjection::ColumnWidth($chart, $categories->SlotCount);
        $depth = Column3DProjection::ColumnDepth($chart);
        $baseline = max(0.0, min(1.0, -$min / ($max - $min)));
        $order = array();
        for ($s = 0; $s < count($chart->Series); $s++) {
            for ($c = 0; $c < $count; $c++) {
                if ($c < count($chart->Series[$s]->Values) && $chart->Series[$s]->Values[$c] !== null) {
                    $order[] = $s * $count + $c;
                }
            }
        }

        usort($order, function ($a, $b) use ($p, $categories, $count, $chart) {
            $da = $p->CameraDepth($categories->Centers[$a % $count], 0,
                (intdiv($a, $count) + 0.5) / count($chart->Series));
            $db = $p->CameraDepth($categories->Centers[$b % $count], 0,
                (intdiv($b, $count) + 0.5) / count($chart->Series));
            $comparison = $db <=> $da;
            return $comparison != 0 ? $comparison : $a <=> $b;
        });
        for ($i = 0; $i < count($order); $i++) {
            $s = intdiv($order[$i], $count);
            $c = $order[$i] % $count;
            $series = $chart->Series[$s];
            $value = max(0.0, min(1.0,
                ($series->Values[$c] - $min) / ($max - $min)));
            $points = $p->ColumnCorners($categories->Centers[$c],
                ($s + 0.5) / count($chart->Series), $width, $depth,
                min($baseline, $value), max($baseline, $value));
            self::DrawFace($canvas, $points, 0, 3, 7, 4, $series, 0.72);
            self::DrawFace($canvas, $points, 3, 2, 6, 7, $series, 0.8);
            self::DrawFace($canvas, $points, 1, 5, 6, 2, $series, 0.66);
            self::DrawFace($canvas, $points, 4, 7, 6, 5, $series, 0.85);
            self::DrawFace($canvas, $points, 0, 4, 5, 1, $series, 1.0);
        }
    }

    private static function DrawFace($canvas, $points, $a, $b, $c, $d, $series, $brightness)
    {
        // Outward face winding survives the projection; hidden faces are culled.
        $cross = ($points[$b]->X - $points[$a]->X) * ($points[$c]->Y - $points[$a]->Y)
            - ($points[$b]->Y - $points[$a]->Y) * ($points[$c]->X - $points[$a]->X);
        if ($cross <= 0.0) {
            return;
        }

        $path = new SKPath();
        $paint = new SKPaint();
        $paint->IsAntialias = true;
        $paint->Style = SKPaintStyle::Fill;
        $color = $series->Color;
        $paint->Color = new SKColor((int)round($color->Red * $brightness),
            (int)round($color->Green * $brightness),
            (int)round($color->Blue * $brightness), $color->Alpha);
        $path->MoveTo($points[$a]);
        $path->LineTo($points[$b]);
        $path->LineTo($points[$c]);
        $path->LineTo($points[$d]);
        $path->Close();
        $canvas->DrawPath($path, $paint);
        if ($series->HasVisibleLine) {
            $paint->Style = SKPaintStyle::Stroke;
            $paint->Color = $series->LineColor;
            $paint->StrokeWidth = (float)$series->LineWidthPt;
            $canvas->DrawPath($path, $paint);
        }

        $paint->Dispose();
        $path->Dispose();
    }

    private function DrawLabels($canvas, $p, $chart, $min, $max, $major, $categories, $width)
    {
        if ($major > 0.0) {
            for ($value = $min; $value <= $max + $major * 0.001; $value += $major) {
                $point = $p->Project(0, ($value - $min) / ($max - $min), 0);
                $label = ChartXmlParser::FormatLabel($value, $chart->ValueFormatCode,
                    $this->dateSystem, $this->culture);
                $this->DrawText($canvas, rtrim((string)$label), $point->X - 9.0, $point->Y,
                    $chart->AxisLabelColor, 1.0);
            }
        }

        $rotation = ChartDateAxisLayout::Resolve($chart, $this->dateSystem, $width * 0.69)->RotationDeg;
        for ($c = 0; $c < $categories->SlotCount; $c++) {
            $point = $p->Project(($c + 0.5) / $categories->SlotCount, 0, 0);
            if (abs($rotation) > 0.1) {
                $canvas->Save();
                $canvas->Translate($point->X, $point->Y + 7.0);
                $canvas->RotateDegrees($rotation);
                $this->DrawText($canvas, $categories->Labels[$c], 0, 0, $chart->AxisLabelColor,
                    $rotation < 0.0 ? 1.0 : 0.0);
                $canvas->Restore();
            } else {
                $this->DrawText($canvas, $categories->Labels[$c], $point->X, $point->Y + 13.0,
                    $chart->AxisLabelColor, 0.5);
            }
        }

        for ($s = 0; $s < count($chart->Series); $s++) {
            $series = $chart->Series[$s];
            $z = ($s + 0.5) / count($chart->Series);
            if ($chart->HasVisibleSeriesAxis) {
                $point = $p->Project(1, 0, $z);
                $this->DrawText($canvas, $series->Name, $point->X + 8.0, $point->Y,
                    $chart->SeriesAxisLabelColor, 0.0);
            }

            if (!$series->ShowValueDataLabels) {
                continue;
            }

            for ($c = 0; $c < count($series->Values); $c++) {
                if ($series->Values[$c] === null) {
                    continue;
                }

                $value = $series->Values[$c];
                $y = max(0.0, min(1.0, ($value - $min) / ($max - $min)));
                $centered = $series->DataLabelPosition === 'ctr';
                $point = $p->Project($categories->Centers[$c], $centered ? $y * 0.5 : $y, $z);
                $format = ($series->ValueFormatCode === null || $series->ValueFormatCode === '')
                    ? $chart->ValueFormatCode : $series->ValueFormatCode;
                $label = ChartXmlParser::FormatLabel($value, $format, $this->dateSystem, $this->culture);
                $this->DrawText($canvas, rtrim((string)$label), $point->X, $point->Y - ($centered ? 0.0 : 6.0),
                    $series->DataLabelColor, 0.5);
            }
        }
    }

    private function DrawText($canvas, $text, $x, $y, $color, $alignment)
    {
        if ($text === null || $text === '') {
            return;
        }

        $fontValue = new FontValue();
        $fontValue->Name = 'Calibri';
        $fontValue->Size = 9.0;
        $fontContext = $this->context->GetFontContext($fontValue);
        $runs = $fontContext->SplitRuns($text);
        if (count($runs) == 0) {
            return;
        }

        $cursor = $x - $fontContext->Measure($text) * $alignment;
        $metricsFont = new SKFont($runs[0]->Typeface, 9.0);
        $baseline = $y - ($metricsFont->Metrics->Ascent + $metricsFont->Metrics->Descent) * 0.5;
        for ($i = 0; $i < count($runs); $i++) {
            $font = new SKFont($runs[$i]->Typeface, 9.0);
            $font->Subpixel = true;
            PdfTextPathRenderer::DrawText($canvas, $runs[$i]->Text, $cursor, $baseline, $font, $color);
            $font->Dispose();
            $cursor += $runs[$i]->WidthPt;
        }

        $metricsFont->Dispose();
    }
}
