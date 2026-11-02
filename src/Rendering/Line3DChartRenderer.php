<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\DateSystem;
use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\CultureInfo;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKFont;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKRect;

// Draws line3DChart series as filled ribbons on a projected value/category/depth box.
// The noFill line in a 3D series suppresses its outline, not the filled ribbon.
final class Line3DChartRenderer
{
    private const LabelSizePt = 9.0;

    private $context;

    private $dateSystem;

    private $culture;

    public function __construct($context, $dateSystem, $culture)
    {
        $this->context = $context;
        $this->dateSystem = $dateSystem;
        $this->culture = $culture !== null ? $culture : CultureInfo::GetInvariantCulture();
    }

    public function Draw($canvas, $page, $area, $plot, $chart, $min, $max, $major)
    {
        if ($canvas === null || $chart === null || count($chart->Series) === 0 || $max <= $min) {
            return;
        }

        $projection = new Line3DProjection($plot, $chart, self::PointCount($chart));
        $this->DrawBox($canvas, $projection, $chart, $min, $max, $major);
        $this->DrawRibbons($canvas, $projection, $chart, $min, $max);
        $this->DrawValueLabels($canvas, $page, $area, $projection, $chart, $min, $max, $major);
        $this->DrawCategoryLabels($canvas, $page, $area, $projection, $chart);
        $this->DrawSeriesLabels($canvas, $page, $area, $projection, $chart);
    }

    private function DrawBox($canvas, $projection, $chart, $min, $max, $major)
    {
        $grid = new SKPaint();
        $grid->Style = SKPaintStyle::Stroke;
        $grid->Color = $chart->ValueGridlineColor;
        $grid->StrokeWidth = (float)max(0.5, $chart->ValueGridlineWidthPt);
        $grid->IsAntialias = true;
        $axis = new SKPaint();
        $axis->Style = SKPaintStyle::Stroke;
        $axis->Color = $chart->CategoryAxisLineColor;
        $axis->StrokeWidth = (float)max(0.5, $chart->CategoryAxisLineWidthPt);
        $axis->IsAntialias = true;

        if ($chart->HasValueGridlines && $major > 0.0) {
            $steps = (int)round(($max - $min) / $major);
            for ($i = 0; $i <= $steps; $i++) {
                $value = min($max, $min + $i * $major);
                $y = ($value - $min) / ($max - $min);
                $front = $projection->Project(0.0, $y, 0.0);
                $backLeft = $projection->Project(0.0, $y, 1.0);
                $backRight = $projection->Project(1.0, $y, 1.0);
                $canvas->DrawLine($front, $backLeft, $grid);
                $canvas->DrawLine($backLeft, $backRight, $grid);
            }
        }

        $frontLeft = $projection->Project(0.0, 0.0, 0.0);
        $frontRight = $projection->Project(1.0, 0.0, 0.0);
        if ($chart->HasCategoryAxisLine) {
            $canvas->DrawLine($frontLeft, $frontRight, $axis);
        }

        $monthlyTicks = Area3DDateAxisMath::BuildMonthlyTicks($chart, $this->dateSystem, $this->culture);
        $count = count($monthlyTicks) > 0 ? count($monthlyTicks) : self::PointCount($chart);
        $between = !$chart->HasDateCategoryAxis && $chart->CrossBetween;
        for ($i = 0; $i < $count + ($between ? 1 : 0); $i++) {
            $x = count($monthlyTicks) > 0
                ? $monthlyTicks[$i][0]
                : ($between ? (double)$i / max(1, $count)
                : Area3DDateAxisMath::CategoryPosition($chart, $i, $count, $this->dateSystem));
            $tick = $projection->Project($x, 0.0, 0.0);
            $canvas->DrawLine($tick->X, $tick->Y, $tick->X, $tick->Y + 2.5, $axis);
        }

        $axis->Dispose();
        $grid->Dispose();
    }

    private function DrawRibbons($canvas, $projection, $chart, $min, $max)
    {
        $count = self::PointCount($chart);
        for ($seriesIndex = count($chart->Series) - 1; $seriesIndex >= 0; $seriesIndex--) {
            $series = $chart->Series[$seriesIndex];
            $values = ChartStackingMath::BuildDisplayValues($chart, $seriesIndex);
            $points = array();
            for ($i = 0; $i < count($values); $i++) {
                if ($values[$i] === null) {
                    continue;
                }

                $x = Area3DDateAxisMath::CategoryPosition($chart, $i, $count, $this->dateSystem);
                $y = max(0.0, min(1.0, ($values[$i] - $min) / ($max - $min)));
                $points[] = new SKPoint((float)$x, (float)$y);
            }

            if (count($points) < 2) {
                continue;
            }

            for ($i = 1; $i < count($points); $i++) {
                $point = $points[$i];
                $j = $i - 1;
                while ($j >= 0 && $points[$j]->X > $point->X) {
                    $points[$j + 1] = $points[$j];
                    $j--;
                }

                $points[$j + 1] = $point;
            }

            $ribbon = new SKPath();
            $fill = new SKPaint();
            $fill->Style = SKPaintStyle::Fill;
            $fill->IsAntialias = true;

            $faces = array();
            for ($i = 1; $i < count($points); $i++) {
                $faces[] = $projection->RibbonCorners($points[$i - 1], $points[$i], $seriesIndex);
            }

            // Underpaint the continuous silhouette so antialiasing on
            // shared face edges cannot leave a white crack at a bend.
            $ribbon->MoveTo($faces[0][0]);
            for ($i = 0; $i < count($faces); $i++) {
                $ribbon->LineTo($faces[$i][1]);
            }

            for ($i = count($faces) - 1; $i >= 0; $i--) {
                $ribbon->LineTo($faces[$i][2]);
            }

            $ribbon->LineTo($faces[0][3]);
            $ribbon->Close();
            $fill->Color = self::Shade($series->Color, 0.76);
            if ($series->Color->Alpha == 255) {
                $canvas->DrawPath($ribbon, $fill);
            }

            for ($i = 0; $i < count($faces); $i++) {
                $corners = $faces[$i];
                $ribbon->Reset();
                $ribbon->MoveTo($corners[0]);
                $ribbon->LineTo($corners[1]);
                $ribbon->LineTo($corners[2]);
                $ribbon->LineTo($corners[3]);
                $ribbon->Close();
                $fill->Color = self::Shade($series->Color, $projection->RibbonShade($points[$i], $points[$i + 1], $corners));
                $canvas->DrawPath($ribbon, $fill);
            }

            $fill->Dispose();
            $ribbon->Dispose();
        }
    }

    private function DrawValueLabels($canvas, $page, $area, $projection, $chart, $min, $max, $major)
    {
        if ($major <= 0.0) {
            return;
        }

        $steps = (int)round(($max - $min) / $major);
        for ($i = 0; $i <= $steps; $i++) {
            $value = min($max, $min + $i * $major);
            $point = $projection->Project(0.0, ($value - $min) / ($max - $min), 0.0);
            $label = ChartXmlParser::FormatLabel($value, $chart->ValueFormatCode, $this->dateSystem, $this->culture);
            $this->DrawText($canvas, $page, $area, $label !== null ? rtrim($label) : '',
                $point->X - 8.0, $point->Y, $chart->AxisLabelColor, true, false);
        }
    }

    private function DrawCategoryLabels($canvas, $page, $area, $projection, $chart)
    {
        $ticks = Area3DDateAxisMath::BuildMonthlyTicks($chart, $this->dateSystem, $this->culture);
        if (count($ticks) === 0) {
            $count = self::PointCount($chart);
            for ($i = 0; $i < $count; $i++) {
                $label = $i < count($chart->Categories) ? $chart->Categories[$i] : '';
                if (($label === null || $label === '') && $i < count($chart->CategoryValues)
                    && $chart->CategoryValues[$i] !== null) {
                    $format = ($chart->CategoryFormatCode === null || $chart->CategoryFormatCode === '')
                        ? 'mmm-yy' : $chart->CategoryFormatCode;
                    $label = ChartXmlParser::FormatLabel($chart->CategoryValues[$i],
                        $format, $this->dateSystem, $this->culture);
                }

                $x = Area3DDateAxisMath::CategoryPosition($chart, $i, $count, $this->dateSystem);
                $point = $projection->Project($x, 0.0, 0.0);
                $this->DrawText($canvas, $page, $area, $label, $point->X, $point->Y + 7.0,
                    $chart->AxisLabelColor, false, true);
            }

            return;
        }

        for ($i = 0; $i < count($ticks); $i++) {
            $point = $projection->Project($ticks[$i][0], 0.0, 0.0);
            $this->DrawText($canvas, $page, $area, $ticks[$i][1], $point->X - 4.0, $point->Y,
                $chart->AxisLabelColor, true, false,
                (float)$chart->CategoryAxisTextRotationDeg, $point->X, $point->Y);
        }
    }

    private function DrawSeriesLabels($canvas, $page, $area, $projection, $chart)
    {
        if (!$chart->HasVisibleSeriesAxis) {
            return;
        }

        for ($i = count($chart->Series) - 1; $i >= 0; $i--) {
            $depth = $projection->SeriesDepth($i);
            $point = $projection->Project(1.0, 0.0, $depth);
            $this->DrawText($canvas, $page, $area, $chart->Series[$i]->Name, $point->X + 5.0, $point->Y,
                $chart->SeriesAxisLabelColor, false, false);
        }
    }

    private function DrawText($canvas, $page, $clip, $text, $x, $y, $color, $alignRight, $alignTop, $rotationDeg = 0.0, $originX = 0.0, $originY = 0.0)
    {
        if ($text === null || $text === '') {
            return;
        }

        $fontValue = new FontValue();
        $fontValue->Name = 'Calibri';
        $fontValue->Size = self::LabelSizePt;
        $fontContext = $this->context->GetFontContext($fontValue);
        $drawX = $alignRight ? $x - $fontContext->Measure($text)
            : ($alignTop ? $x - $fontContext->Measure($text) * 0.5 : $x);
        $paint = new SKPaint();
        $paint->Typeface = $this->context->Fonts->Resolve($fontValue);
        $paint->TextSize = self::LabelSizePt;
        $paint->IsAntialias = true;

        $metrics = $paint->FontMetrics;
        $baseline = $alignTop ? $y - $metrics->Ascent : $y - ($metrics->Ascent + $metrics->Descent) * 0.5;
        $runs = $fontContext->SplitRuns($text);
        $cursor = $drawX;
        $session = $this->context->PdfTextSession;
        if ($session !== null && $page !== null && strpos($text, "\r") === false && strpos($text, "\n") === false) {
            for ($i = 0; $i < count($runs); $i++) {
                $run = $runs[$i];
                if ($rotationDeg == 0.0) {
                    $session->RecordChartRun($page, $clip, $run->Text, $run->Typeface, $color,
                        self::LabelSizePt, $cursor, $baseline);
                } else {
                    // The PDF session applies the page transform and rotation pivot;
                    // only the glyph position is relative to that pivot.
                    $session->RecordChartRotatedRun($page, $clip, $run->Text, $run->Typeface, $color,
                        self::LabelSizePt, $originX, $originY, $cursor - $originX, $baseline - $originY, $rotationDeg);
                }

                $cursor += $run->WidthPt;
            }

            $paint->Dispose();
            return;
        }

        $canvas->Save();
        $canvas->ClipRect($clip);
        $canvas->RotateDegrees($rotationDeg, $originX, $originY);
        for ($i = 0; $i < count($runs); $i++) {
            $run = $runs[$i];
            $font = new SKFont($run->Typeface, self::LabelSizePt);
            $font->Subpixel = true;
            PdfTextPathRenderer::DrawText($canvas, $run->Text, $cursor, $baseline, $font, $color);
            $font->Dispose();

            $cursor += $run->WidthPt;
        }

        $canvas->Restore();
        $paint->Dispose();
    }

    private static function PointCount($chart)
    {
        $count = max(count($chart->Categories), count($chart->CategoryValues));
        for ($i = 0; $i < count($chart->Series); $i++) {
            $count = max($count, count($chart->Series[$i]->Values));
        }

        return $count;
    }

    private static function Shade($color, $factor)
    {
        return new SKColor((int)round($color->Red * $factor),
            (int)round($color->Green * $factor),
            (int)round($color->Blue * $factor), $color->Alpha);
    }
}
