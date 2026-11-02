<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\DiagnosticEntry;
use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\Skia\SKFont;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKRect;
use Aspose\Cells_FOSS\Skia\SKStrokeCap;
use Aspose\Cells_FOSS\Skia\SKStrokeJoin;
use Aspose\Cells_FOSS\Skia\SKTextAlign;

// Draws a ParsedShape into a rectangle: preset geometries (rect, ellipse,
// triangle, block arrows, plus/cross, ...) with fill/line, straight line connectors with
// arrowheads, and centered text. Unknown geometries fall back to a rounded/plain rectangle.
final class ShapeRenderer
{
    private $context;

    public function __construct($context)
    {
        $this->context = $context;
    }

    public function Draw($canvas, $page = null, $rect = null, $shape = null, $shapeModel = null, $shapes = null)
    {
        // Overload Draw(SKCanvas, SKRect, ParsedShape): the page and rect arguments shift.
        if ($page instanceof SKRect) {
            $shape = $rect;
            $rect = $page;
            $page = null;
        }

        if (self::IsConnector($shape->Geometry)) {
            $this->DrawConnector($canvas, $page, $rect, $shape, $shapeModel, $shapes);
            return;
        }

        $restore = false;
        if (abs($shape->RotationDeg) > 0.01) {
            $canvas->Save();
            $canvas->RotateDegrees((float)$shape->RotationDeg, $rect->MidX, $rect->MidY);
            $restore = true;
        }

        $this->DrawFilledShape($canvas, $page, $rect, $shape);

        if ($restore) {
            $canvas->Restore();
        }
    }

    private static function IsConnector($geometry)
    {
        return $geometry === 'line'
            || str_starts_with($geometry, 'straightConnector')
            || str_starts_with($geometry, 'bentConnector')
            || str_starts_with($geometry, 'curvedConnector');
    }

    // Preserve the existing non-triangle arrow scale. Medium triangle heads use six
    // line widths in length and total width, with a one-point minimum scaling width.
    const MinArrowScalePt = 1.1;

    private function DrawConnector($canvas, $page, $rect, $shape, $shapeModel, $shapes)
    {
        if (!$shape->HasLine) {
            return;
        }

        if (str_starts_with($shape->Geometry, 'curvedConnector')) {
            $this->DrawCurvedConnector($canvas, $page, $rect, $shape, $shapeModel, $shapes);
            return;
        }

        $width = (float)max(0.5, $shape->LineWidthPt);
        $initialStart = self::TransformNormalizedPoint($rect, $shape, 0.0, 0.0);
        $initialEnd = self::TransformNormalizedPoint($rect, $shape, 1.0, 1.0);
        $this->ResolveConnectorEndpoints($page, $shapeModel, $shapes, $initialStart, $initialEnd);
        $points = ShapePolylineGeometry::Build($initialStart, $initialEnd, $shape);

        if ($points === null || count($points) < 2) {
            return;
        }

        $start = $points[0];
        $end = $points[count($points) - 1];
        $dx = $end->X - $start->X;
        $dy = $end->Y - $start->Y;
        $len = (float)sqrt($dx * $dx + $dy * $dy);
        if ($len < 0.01) {
            return;
        }

        $arrowScale = max($width, self::MinArrowScalePt);

        ShapeReflectionRenderer::Draw($canvas, $points, $shape);

        // A filled arrowhead (triangle/stealth/diamond) hides the line end, so stop the line at
        // its base with a flat cap; an open "arrow" (V) sits on top of the line, so the line runs
        // to the tip.
        $drawPoints = self::CopyPoints($points);
        if (self::HasArrow($shape->HeadEnd)) {
            self::ShortenPolylineStart($drawPoints, self::FilledArrow($shape->HeadEnd) ? $arrowScale * self::ArrowLineHideMultiplier($shape->HeadEnd) : 0.0);
        }

        if (self::HasArrow($shape->TailEnd)) {
            self::ShortenPolylineEnd($drawPoints, self::FilledArrow($shape->TailEnd) ? $arrowScale * self::ArrowLineHideMultiplier($shape->TailEnd) : 0.0);
        }

        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = $shape->LineColor;
        $paint->StrokeWidth = $width;
        $paint->IsAntialias = true;
        $paint->StrokeCap = SKStrokeCap::Butt;
        $paint->StrokeJoin = SKStrokeJoin::Miter;
        $path = new SKPath();
        $path->MoveTo($drawPoints[0]);
        for ($i = 1; $i < count($drawPoints); $i++) {
            $path->LineTo($drawPoints[$i]);
        }

        $canvas->DrawPath($path, $paint);
        $path->Dispose();
        $paint->Dispose();

        $headDirection = self::SegmentDirection($points[1], $points[0]);
        $tailDirection = self::SegmentDirection($points[count($points) - 2], $points[count($points) - 1]);
        $this->DrawArrowhead($canvas, $shape->HeadEnd, $start, $headDirection->X, $headDirection->Y, $width, $shape->LineColor);
        $this->DrawArrowhead($canvas, $shape->TailEnd, $end, $tailDirection->X, $tailDirection->Y, $width, $shape->LineColor);
    }

    private function DrawCurvedConnector($canvas, $page, $rect, $shape, $shapeModel, $shapes)
    {
        $width = (float)max(0.5, $shape->LineWidthPt);
        $arrowScale = max($width, self::MinArrowScalePt);

        $start = self::TransformNormalizedPoint($rect, $shape, 0.0, 0.0);
        $end = self::TransformNormalizedPoint($rect, $shape, 1.0, 1.0);
        $this->ResolveConnectorEndpoints($page, $shapeModel, $shapes, $start, $end);

        $curve = ShapeCurveGeometry::Build($start, $end, $shape);
        $control1 = $curve[1];
        $control2 = $curve[count($curve) - 2];

        $endDirection = self::SegmentDirection($control2, $end);
        $headDirection = self::SegmentDirection($control1, $start);

        $visibleStart = $start;
        $visibleEnd = $end;
        if (self::HasArrow($shape->HeadEnd)) {
            $visibleStart = new SKPoint(
                $start->X + $headDirection->X * $arrowScale * self::ArrowLineHideMultiplier($shape->HeadEnd),
                $start->Y + $headDirection->Y * $arrowScale * self::ArrowLineHideMultiplier($shape->HeadEnd));
        }

        if (self::HasArrow($shape->TailEnd)) {
            $visibleEnd = new SKPoint(
                $end->X - $endDirection->X * $arrowScale * self::ArrowLineHideMultiplier($shape->TailEnd),
                $end->Y - $endDirection->Y * $arrowScale * self::ArrowLineHideMultiplier($shape->TailEnd));
        }

        $path = new SKPath();
        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Stroke;
        $paint->Color = $shape->LineColor;
        $paint->StrokeWidth = $width;
        $paint->IsAntialias = true;
        $paint->StrokeCap = SKStrokeCap::Butt;
        $paint->StrokeJoin = SKStrokeJoin::Round;

        $path->MoveTo($visibleStart);
        $curve[count($curve) - 1] = $visibleEnd;
        for ($i = 1; $i < count($curve); $i += 3) {
            $path->CubicTo($curve[$i]->X, $curve[$i]->Y, $curve[$i + 1]->X, $curve[$i + 1]->Y, $curve[$i + 2]->X, $curve[$i + 2]->Y);
        }
        $canvas->DrawPath($path, $paint);
        $path->Dispose();
        $paint->Dispose();

        $this->DrawArrowhead($canvas, $shape->HeadEnd, $start, $headDirection->X, $headDirection->Y, $width, $shape->LineColor);
        $this->DrawArrowhead($canvas, $shape->TailEnd, $end, $endDirection->X, $endDirection->Y, $width, $shape->LineColor);
    }

    private function ResolveConnectorEndpoints($page, $connector, $shapes, &$start, &$end)
    {
        if ($page === null || $connector === null || $shapes === null) {
            return;
        }

        $startShape = self::FindShapeByDrawingObjectId($shapes, $connector->StartConnectionShapeId);
        if ($startShape !== null) {
            $start = self::ConnectionPoint($page, $startShape, $connector->StartConnectionSite);
        }

        $endShape = self::FindShapeByDrawingObjectId($shapes, $connector->EndConnectionShapeId);
        if ($endShape !== null) {
            $end = self::ConnectionPoint($page, $endShape, $connector->EndConnectionSite);
        }
    }

    private static function FindShapeByDrawingObjectId($shapes, $drawingObjectId)
    {
        if ($shapes === null || $drawingObjectId < 0) {
            return null;
        }

        for ($i = 0; $i < count($shapes); $i++) {
            $shape = $shapes[$i];
            if ($shape !== null && $shape->DrawingObjectId === $drawingObjectId) {
                return $shape;
            }
        }

        return null;
    }

    private static function ConnectionPoint($page, $shape, $site)
    {
        $rect = self::ShapeRect($page, $shape);
        if ($shape->GeometryType === 'snip1Rect' || $shape->GeometryType === 'snip2SameRect') {
            // These presets enumerate right, bottom, left, top, unlike rect.
            $sites = array(3, 2, 1, 0);
            if ($site >= 0 && $site < count($sites)) {
                $site = $sites[$site];
            }
        }
        if ($site === 0) {
            return new SKPoint($rect->MidX, $rect->Top);
        }

        if ($site === 1) {
            return new SKPoint($rect->Left, $rect->MidY);
        }

        if ($site === 2) {
            return new SKPoint($rect->MidX, $rect->Bottom);
        }

        if ($site === 3) {
            return new SKPoint($rect->Right, $rect->MidY);
        }

        return new SKPoint($rect->MidX, $rect->MidY);
    }

    private static function ShapeRect($page, $shape)
    {
        $layout = $page->Sheet;
        $originX = $layout->ColumnStartPt[$page->StartColumn];
        $originY = $layout->RowStartPt[$page->StartRow];
        $left = self::ColumnEdgePt($layout, $shape->UpperLeftColumn, $shape->UpperLeftColumnOffset) - $originX;
        $top = self::RowEdgePt($layout, $shape->UpperLeftRow, $shape->UpperLeftRowOffset) - $originY;
        $right = self::ColumnEdgePt($layout, $shape->LowerRightColumn, $shape->LowerRightColumnOffset) - $originX;
        $bottom = self::RowEdgePt($layout, $shape->LowerRightRow, $shape->LowerRightRowOffset) - $originY;
        return new SKRect((float)$left, (float)$top, (float)$right, (float)$bottom);
    }

    private static function ColumnEdgePt($layout, $columnIndex, $offsetEmu)
    {
        return $layout->ColumnStartPt[$columnIndex] + $offsetEmu / 12700.0;
    }

    private static function RowEdgePt($layout, $rowIndex, $offsetEmu)
    {
        return $layout->RowStartPt[$rowIndex] + $offsetEmu / 12700.0;
    }

    private static function HasArrow($type)
    {
        return $type !== null && $type !== '' && $type !== 'none';
    }

    private static function FilledArrow($type)
    {
        return $type === 'triangle' || $type === 'stealth' || $type === 'diamond' || $type === 'oval';
    }

    private function DrawArrowhead($canvas, $type, $tip, $dirX, $dirY, $lineWidth, $color)
    {
        if (!self::HasArrow($type)) {
            return;
        }

        $px = -$dirY;
        $py = $dirX;
        $arrowScale = $type === 'triangle' ? max($lineWidth, 1.0) : max($lineWidth, self::MinArrowScalePt);
        $arrowLen = $arrowScale * self::ArrowLengthMultiplier($type);
        $halfW = $arrowScale * self::ArrowHalfWidthMultiplier($type);
        $baseX = $tip->X - $dirX * $arrowLen;
        $baseY = $tip->Y - $dirY * $arrowLen;

        if (!self::FilledArrow($type)) {
            // Open "arrow": two strokes forming a V at the tip.
            $paint = new SKPaint();
            $paint->Style = SKPaintStyle::Stroke;
            $paint->Color = $color;
            $paint->StrokeWidth = $lineWidth;
            $paint->IsAntialias = true;
            $paint->StrokeCap = SKStrokeCap::Round;
            $canvas->DrawLine($tip->X, $tip->Y, $baseX + $px * $halfW, $baseY + $py * $halfW, $paint);
            $canvas->DrawLine($tip->X, $tip->Y, $baseX - $px * $halfW, $baseY - $py * $halfW, $paint);
            $paint->Dispose();
            return;
        }

        $path = new SKPath();
        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Fill;
        $paint->Color = $color;
        $paint->IsAntialias = true;
        $path->MoveTo($tip->X, $tip->Y);
        $path->LineTo($baseX + $px * $halfW, $baseY + $py * $halfW);
        $path->LineTo($baseX - $px * $halfW, $baseY - $py * $halfW);
        $path->Close();
        $canvas->DrawPath($path, $paint);
        $path->Dispose();
        $paint->Dispose();
    }

    private static function ArrowLengthMultiplier($type)
    {
        if ($type === 'triangle') {
            return 6.0;
        }

        if ($type === 'stealth') {
            return 0.90;
        }

        return 1.05;
    }

    private static function ArrowHalfWidthMultiplier($type)
    {
        if ($type === 'triangle') {
            return 3.0;
        }

        if ($type === 'stealth') {
            return 1.05;
        }

        return 1.15;
    }

    private static function ArrowLineHideMultiplier($type)
    {
        if ($type === 'triangle') {
            // Excel's filled connector arrowheads are compact wedges that sit on top of the
            // line body rather than replacing a long tail segment.
            return 0.00;
        }

        return self::ArrowLengthMultiplier($type);
    }

    private function DrawFilledShape($canvas, $page, $rect, $shape)
    {
        $path = $this->BuildGeometry($shape, $rect);
        if ($shape->HasFill && !ShapePresetDetails::DrawFill($canvas, $rect, $shape)) {
            $fill = new SKPaint();
            $fill->Style = SKPaintStyle::Fill;
            $fill->Color = $shape->FillColor;
            $fill->IsAntialias = true;
            $canvas->DrawPath($path, $fill);
            $fill->Dispose();
            ShapePatternRenderer::Draw($canvas, $rect, $path, $shape);
        }

        if ($shape->HasLine) {
            $stroke = new SKPaint();
            $stroke->Style = SKPaintStyle::Stroke;
            $stroke->Color = $shape->LineColor;
            $stroke->StrokeWidth = max(0.5, $shape->LineWidthPt);
            $stroke->IsAntialias = true;
            $stroke->StrokeJoin = SKStrokeJoin::Miter;
            $canvas->DrawPath($path, $stroke);
            $stroke->Dispose();
        }
        $path->Dispose();

        ShapePresetDetails::DrawLines($canvas, $rect, $shape);
        $this->DrawText($canvas, $page, $rect, $shape);
    }

    private function BuildGeometry($shape, $r)
    {
        $special = ShapeSpecialGeometry::Build($shape, $r);
        if ($special !== null) {
            return $special;
        }
        if ($shape->CustomGeometry !== null) {
            $entry = new DiagnosticEntry();
            $entry->Code = 'PDF_CUSTOM_GEOMETRY_UNSUPPORTED';
            $entry->Message = 'Custom shape path contains unsupported or invalid coordinates or commands; rendered as a rectangle.';
            $entry->DataLossRisk = true;
            $this->context->Workbook->Diagnostics->Add($entry);
        }

        $geometry = $shape->Geometry;
        $path = new SKPath();
        $l = $r->Left;
        $t = $r->Top;
        $right = $r->Right;
        $b = $r->Bottom;
        $w = $r->Width;
        $h = $r->Height;
        $cx = $r->MidX;
        $cy = $r->MidY;

        switch ($geometry) {
            case 'corner':
                self::CornerShape($path, $r, $shape->Adjusts);
                break;
            case 'ellipse':
                $path->AddOval($r);
                break;
            case 'roundRect':
                $path->AddRoundRect($r, min($w, $h) * 0.16, min($w, $h) * 0.16);
                break;
            case 'triangle':
                // The apex x-position is the adjust value (default 0.5 = isosceles; 1.0 puts
                // the apex at the right edge, giving a right triangle).
                $apexAdj = $shape->Adjusts !== null && count($shape->Adjusts) > 0 ? $shape->Adjusts[0] : 0.5;
                $path->MoveTo($l + (float)($w * $apexAdj), $t);
                $path->LineTo($right, $b);
                $path->LineTo($l, $b);
                $path->Close();
                break;
            case 'rtTriangle':
                $path->MoveTo($l, $t);
                $path->LineTo($l, $b);
                $path->LineTo($right, $b);
                $path->Close();
                break;
            case 'diamond':
                $path->MoveTo($cx, $t);
                $path->LineTo($right, $cy);
                $path->LineTo($cx, $b);
                $path->LineTo($l, $cy);
                $path->Close();
                break;
            case 'rightArrow':
                self::BlockArrowRight($path, $r);
                break;
            case 'leftArrow':
                self::BlockArrowLeft($path, $r);
                break;
            case 'upArrow':
                self::BlockArrowUp($path, $r);
                break;
            case 'downArrow':
                self::BlockArrowDown($path, $r);
                break;
            case 'mathPlus':
            case 'plus':
                self::PlusShape($path, $r);
                break;
            default:
                $path->AddRect($r);
                break;
        }

        return $path;
    }

    private static function CornerShape($path, $r, $adj)
    {
        // L-shape: a vertical arm on the left and a horizontal arm on the bottom; the top-right is
        // cut out. Both arm thicknesses are relative to the shape's smaller dimension (ss), so the
        // arms stay equally thick on a non-square box (per the DrawingML "corner" preset).
        $a1 = $adj !== null && count($adj) > 0 ? $adj[0] : 0.5;
        $a2 = $adj !== null && count($adj) > 1 ? $adj[1] : 0.5;
        $ss = min($r->Width, $r->Height);
        $dx1 = (float)($ss * $a2);
        $dy1 = (float)($ss * $a1);
        $l = $r->Left;
        $t = $r->Top;
        $right = $r->Right;
        $b = $r->Bottom;
        $path->MoveTo($l, $t);
        $path->LineTo($l + $dx1, $t);
        $path->LineTo($l + $dx1, $b - $dy1);
        $path->LineTo($right, $b - $dy1);
        $path->LineTo($right, $b);
        $path->LineTo($l, $b);
        $path->Close();
    }

    private static function PlusShape($path, $r)
    {
        // mathPlus preset: arm thickness = ss * adj (default adj 0.2352); the plus arms reach only
        // 73.49% of the box (dx2 = w * 73490/200000), centered - they do NOT fill the box.
        $d = min($r->Width, $r->Height) * 0.2352;
        $half = $d / 2.0;
        $cx = $r->MidX;
        $cy = $r->MidY;
        $hx = $r->Width * 0.367450;   // half horizontal extent of the plus
        $hy = $r->Height * 0.367450;  // half vertical extent of the plus
        $l = $cx - $hx;
        $right = $cx + $hx;
        $t = $cy - $hy;
        $b = $cy + $hy;

        $path->MoveTo($cx - $half, $t);
        $path->LineTo($cx + $half, $t);
        $path->LineTo($cx + $half, $cy - $half);
        $path->LineTo($right, $cy - $half);
        $path->LineTo($right, $cy + $half);
        $path->LineTo($cx + $half, $cy + $half);
        $path->LineTo($cx + $half, $b);
        $path->LineTo($cx - $half, $b);
        $path->LineTo($cx - $half, $cy + $half);
        $path->LineTo($l, $cy + $half);
        $path->LineTo($l, $cy - $half);
        $path->LineTo($cx - $half, $cy - $half);
        $path->Close();
    }

    private static function BlockArrowRight($path, $r)
    {
        $headW = min($r->Width, $r->Height) * 0.5;
        $shaftH = $r->Height * 0.5;
        $headInset = 0.0;
        $sy0 = $r->MidY - $shaftH / 2.0;
        $sy1 = $r->MidY + $shaftH / 2.0;
        $hx = $r->Right - $headW;
        $path->MoveTo($r->Left, $sy0);
        $path->LineTo($hx, $sy0);
        $path->LineTo($hx, $r->Top + $headInset);
        $path->LineTo($r->Right, $r->MidY);
        $path->LineTo($hx, $r->Bottom - $headInset);
        $path->LineTo($hx, $sy1);
        $path->LineTo($r->Left, $sy1);
        $path->Close();
    }

    private static function BlockArrowLeft($path, $r)
    {
        $headW = $r->Width * 0.4;
        $shaftH = $r->Height * 0.5;
        $sy0 = $r->MidY - $shaftH / 2.0;
        $sy1 = $r->MidY + $shaftH / 2.0;
        $hx = $r->Left + $headW;
        $path->MoveTo($r->Right, $sy0);
        $path->LineTo($hx, $sy0);
        $path->LineTo($hx, $r->Top);
        $path->LineTo($r->Left, $r->MidY);
        $path->LineTo($hx, $r->Bottom);
        $path->LineTo($hx, $sy1);
        $path->LineTo($r->Right, $sy1);
        $path->Close();
    }

    private static function BlockArrowUp($path, $r)
    {
        $headH = $r->Height * 0.4;
        $shaftW = $r->Width * 0.5;
        $sx0 = $r->MidX - $shaftW / 2.0;
        $sx1 = $r->MidX + $shaftW / 2.0;
        $hy = $r->Top + $headH;
        $path->MoveTo($sx0, $r->Bottom);
        $path->LineTo($sx0, $hy);
        $path->LineTo($r->Left, $hy);
        $path->LineTo($r->MidX, $r->Top);
        $path->LineTo($r->Right, $hy);
        $path->LineTo($sx1, $hy);
        $path->LineTo($sx1, $r->Bottom);
        $path->Close();
    }

    private static function BlockArrowDown($path, $r)
    {
        $headH = $r->Height * 0.4;
        $shaftW = $r->Width * 0.5;
        $sx0 = $r->MidX - $shaftW / 2.0;
        $sx1 = $r->MidX + $shaftW / 2.0;
        $hy = $r->Bottom - $headH;
        $path->MoveTo($sx0, $r->Top);
        $path->LineTo($sx0, $hy);
        $path->LineTo($r->Left, $hy);
        $path->LineTo($r->MidX, $r->Bottom);
        $path->LineTo($r->Right, $hy);
        $path->LineTo($sx1, $hy);
        $path->LineTo($sx1, $r->Top);
        $path->Close();
    }

    private function DrawText($canvas, $page, $rect, $shape)
    {
        if ($shape->Text === null || $shape->Text === '') {
            return;
        }

        $font = new FontValue();
        $font->Name = 'Calibri';
        $font->Size = $shape->TextSizePt;
        $font->Bold = $shape->TextBold;
        $hAlign = $shape->TextHAlign === 'l' ? SKTextAlign::Left : ($shape->TextHAlign === 'r' ? SKTextAlign::Right : SKTextAlign::Center);
        $fontContext = $this->context->GetFontContext($font);
        $paint = new SKPaint();
        $paint->Color = $shape->TextColor;
        $paint->IsAntialias = true;
        $paint->Typeface = $this->context->Fonts->Resolve($font);
        $paint->TextSize = (float)$shape->TextSizePt;
        $paint->TextAlign = $hAlign;

        $metrics = $paint->FontMetrics;

        $x = 0.0;
        if ($hAlign === SKTextAlign::Left) {
            $x = $rect->Left + (float)$shape->TextInsetLeftPt;
        } elseif ($hAlign === SKTextAlign::Right) {
            $x = $rect->Right - (float)$shape->TextInsetLeftPt;
        } else {
            $x = $rect->MidX;
        }

        $width = $fontContext->Measure($shape->Text);
        $drawX = $hAlign === SKTextAlign::Left ? $x : ($hAlign === SKTextAlign::Center ? $x - $width / 2.0 : $x - $width);

        // Vertical anchor: top / center / bottom.
        if ($shape->TextVAnchor === 't') {
            $baseline = $rect->Top + (float)$shape->TextInsetTopPt - $metrics->Ascent;
        } elseif ($shape->TextVAnchor === 'b') {
            $baseline = $rect->Bottom - (float)$shape->TextInsetTopPt - $metrics->Descent;
        } else {
            $baseline = $rect->MidY - ($metrics->Ascent + $metrics->Descent) / 2.0;
        }

        if (ShapeVerticalTextRenderer::IsVerticalText($shape)) {
            ShapeVerticalTextRenderer::Draw($canvas, $rect, $shape, $fontContext, $paint);
            $paint->Dispose();
            return;
        }

        if ($shape->TextWrap && $fontContext->Measure($shape->Text) > $rect->Width - (float)$shape->TextInsetLeftPt * 2.0) {
            ShapeWrappedTextRenderer::Draw($canvas, $rect, $shape, $fontContext, $paint);
            $paint->Dispose();
            return;
        }

        if (!$this->TryRecordText($page, $rect, $fontContext, $shape->Text, $paint->Color, $paint->TextSize, $drawX, $baseline, $shape->RotationDeg)) {
            $fontObject = new SKFont($paint->Typeface, $paint->TextSize);
            try {
                $fontObject->Subpixel = true;
                PdfTextPathRenderer::DrawText($canvas, $shape->Text, $drawX, $baseline, $fontObject, $paint->Color);
            } finally {
                $fontObject->Dispose();
            }
        }
        $paint->Dispose();
    }

    private function TryRecordText($page, $clipRect, $fontContext, $text, $color, $fontSizePt, $x, $baseline, $rotationDeg)
    {
        $session = $this->context->PdfTextSession;
        if ($session === null || $page === null || $fontContext === null || $text === null || $text === '') {
            return false;
        }

        // Shape text is already laid out in the local drawing canvas. Re-emitting it through the
        // page-level Type3 text rewrite has produced upside-down text in some text boxes
        // (shape1.xlsx "Start"), so keep shape text on the original path-rendering route.
        return false;
    }

    private static function TransformNormalizedPoint($rect, $shape, $x, $y)
    {
        if ($shape->FlipH) {
            $x = 1.0 - $x;
        }

        if ($shape->FlipV) {
            $y = 1.0 - $y;
        }

        $transformed = self::RotateNormalizedPoint($x, $y, $shape->RotationDeg);
        return new SKPoint(
            $rect->Left + $transformed->X * $rect->Width,
            $rect->Top + $transformed->Y * $rect->Height);
    }

    private static function RotateNormalizedPoint($x, $y, $rotationDeg)
    {
        if (abs($rotationDeg) <= 0.01) {
            return new SKPoint($x, $y);
        }

        $radians = $rotationDeg * M_PI / 180.0;
        $dx = $x - 0.5;
        $dy = $y - 0.5;
        $rotatedX = $dx * cos($radians) - $dy * sin($radians);
        $rotatedY = $dx * sin($radians) + $dy * cos($radians);
        return new SKPoint((float)(0.5 + $rotatedX), (float)(0.5 + $rotatedY));
    }

    private static function CopyPoints($points)
    {
        $copy = array();
        for ($i = 0; $i < count($points); $i++) {
            $copy[$i] = $points[$i];
        }

        return $copy;
    }

    private static function ShortenPolylineStart(&$points, $shortenBy)
    {
        if ($shortenBy <= 0.0 || $points === null || count($points) < 2) {
            return;
        }

        $direction = self::SegmentDirection($points[1], $points[0]);
        $points[0] = new SKPoint($points[0]->X + $direction->X * $shortenBy, $points[0]->Y + $direction->Y * $shortenBy);
    }

    private static function ShortenPolylineEnd(&$points, $shortenBy)
    {
        if ($shortenBy <= 0.0 || $points === null || count($points) < 2) {
            return;
        }

        $last = count($points) - 1;
        $direction = self::SegmentDirection($points[$last - 1], $points[$last]);
        $points[$last] = new SKPoint($points[$last]->X - $direction->X * $shortenBy, $points[$last]->Y - $direction->Y * $shortenBy);
    }

    private static function SegmentDirection($from, $to)
    {
        $dx = $to->X - $from->X;
        $dy = $to->Y - $from->Y;
        $length = (float)sqrt($dx * $dx + $dy * $dy);
        if ($length <= 0.01) {
            return new SKPoint(1.0, 0.0);
        }

        return new SKPoint($dx / $length, $dy / $length);
    }
}
