<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\AlignmentValue;
use Aspose\Cells_FOSS\Core\BorderSideValue;
use Aspose\Cells_FOSS\Core\BorderStyle;
use Aspose\Cells_FOSS\Core\BordersValue;
use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\CellRecord;
use Aspose\Cells_FOSS\Core\CellValueKind;
use Aspose\Cells_FOSS\Core\ChartModel;
use Aspose\Cells_FOSS\Core\DateSystem;
use Aspose\Cells_FOSS\Core\FillPatternKind;
use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\Core\HorizontalAlignment;
use Aspose\Cells_FOSS\Core\MergeRegion;
use Aspose\Cells_FOSS\Core\PictureModel;
use Aspose\Cells_FOSS\Core\ShapeModel;
use Aspose\Cells_FOSS\Core\SmartArtModel;
use Aspose\Cells_FOSS\Core\StyleValue;
use Aspose\Cells_FOSS\Core\VerticalAlignment;
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\DisplayTextFormatter;
use Aspose\Cells_FOSS\DisplayTextFormatterSupport;
use Aspose\Cells_FOSS\FontUnderlineType;
use Aspose\Cells_FOSS\NumberFormat;
use Aspose\Cells_FOSS\Skia\SKClipOperation;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;
use Aspose\Cells_FOSS\Skia\SKFilterQuality;
use Aspose\Cells_FOSS\Skia\SKFont;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKPath;
use Aspose\Cells_FOSS\Skia\SKPathEffect;
use Aspose\Cells_FOSS\Skia\SKPoint;
use Aspose\Cells_FOSS\Skia\SKRect;
use Aspose\Cells_FOSS\Skia\SKShader;
use Aspose\Cells_FOSS\Skia\SKShaderTileMode;
use Aspose\Cells_FOSS\Skia\SKStrokeCap;
use Aspose\Cells_FOSS\Skia\SKTypeface;

final class SheetPdfRenderer
{
    private static $GridlineColor = null;

    private $_context;

    private $_chartCache = array();

    private $_mapChartCache = array();

    private $_shapeCache = array();

    private $_smartArtCache = array();

    private $_cfCache = array();

    private $_richTextRenderer;

    private $_tableStyleResolver;

    // English Metric Units per point (914400 EMU per inch / 72 points per inch).
    const EmuPerPoint = 12700.0;

    public function __construct($context)
    {
        if (self::$GridlineColor === null) {
            self::$GridlineColor = new SKColor(0xD0, 0xD0, 0xD0);
        }

        $this->_context = $context;
        $this->_richTextRenderer = new RichTextCellRenderer($context);
        $this->_tableStyleResolver = new TableStyleResolver($context->Workbook);
    }

    public function RenderPage($canvas, $page)
    {
        $layout = $page->Sheet;

        $canvas->Save();
        $canvas->Translate((float)$page->ContentOriginXPt, (float)$page->ContentOriginYPt);
        $canvas->Scale((float)$page->ScaleFactor);

        // Clip to the printable region (in content coordinates) so scaled content never bleeds
        // into the page margins.
        $contentClip = PageContentClip::Resolve($page, false);

        // Cell grid: clip tightly to the content span.
        $canvas->Save();
        $canvas->ClipRect($contentClip);

        $this->DrawFills($canvas, $page);

        if ($layout->Sheet->PageSetup->PrintOptions->GridLines) {
            $this->DrawGridlines($canvas, $page);
        }

        $this->DrawCellBordersAndText($canvas, $page);

        $this->DrawConditionalIcons($canvas, $page);

        // Merged regions render last, as self-contained units, so they draw correctly on every
        // page they intersect - including continuation pages where the merge's origin lies on an
        // earlier page. The page-level clip trims the parts that fall outside this page.
        $this->DrawMerges($canvas, $page);

        $canvas->Restore();

        // Floating objects (pictures, charts, shapes) may slightly overrun the used range (line
        // width, arrowheads) at the sheet's outer edges, where empty page space follows. Allow a
        // small overflow there, but keep the clip tight at internal page-split boundaries so a
        // wide object still splits cleanly across pages.
        // 3D chart series-axis labels can legitimately extend farther beyond the used range than
        // pictures or arrowheads, especially on the sheet's outer right edge. Keep the clip tight
        // at internal page breaks, but allow a larger outer-edge overflow so those labels are not
        // truncated to just their first few characters.
        $objectClip = PageContentClip::Resolve($page, true);

        $canvas->Save();
        $canvas->ClipRect($objectClip);
        $this->DrawPictures($canvas, $page);
        $this->DrawCharts($canvas, $page);
        $this->DrawShapes($canvas, $page);
        $this->DrawSmartArt($canvas, $page);
        $canvas->Restore();

        $canvas->Restore();
    }

    private function ContentX($layout, $startCol, $col)
    {
        return $layout->ColumnStartPt[$col] - $layout->ColumnStartPt[$startCol];
    }

    private function ContentY($layout, $startRow, $row)
    {
        return $layout->RowStartPt[$row] - $layout->RowStartPt[$startRow];
    }

    private function ConditionalFormatFor($layout)
    {
        $key = spl_object_id($layout->Sheet);
        $evaluator = null;
        if (!array_key_exists($key, $this->_cfCache)) {
            $evaluator = new ConditionalFormatEvaluator($layout->Sheet, $this->_context);
            $this->_cfCache[$key] = $evaluator;
        } else {
            $evaluator = $this->_cfCache[$key];
        }

        return $evaluator;
    }

    private function DrawFills($canvas, $page)
    {
        $layout = $page->Sheet;
        $cf = $this->ConditionalFormatFor($layout);
        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Fill;
        $paint->IsAntialias = false;

        for ($row = $page->StartRow; $row <= $page->EndRow; $row++) {
            if ($layout->RowHidden[$row]) continue;

            for ($col = $page->StartColumn; $col <= $page->EndColumn; $col++) {
                if ($layout->ColumnHidden[$col]) continue;
                if (self::IsPartOfMerge($layout, $row, $col)) continue;

                $record = null;
                $addressKey = (new CellAddress($row, $col))->ToString();
                if (array_key_exists($addressKey, $layout->Sheet->Cells)) {
                    $record = $layout->Sheet->Cells[$addressKey];
                }

                $baseStyle = $record !== null && $record->Style !== null ? $record->Style : StyleValue::Default();
                $style = $this->_tableStyleResolver->Resolve($layout->Sheet, $row, $col, $baseStyle);

                // A matching conditional-format fill (dxf or color scale) overrides the cell fill.
                $effect = null;
                if ($cf->TryGet($row, $col, $effect) && $effect->HasFill) {
                    $this->DrawRotatedNeighborUnderlay($canvas, $page, $row, $col, $style);
                    if (self::FillShearFactor($style->Alignment) != 0.0) {
                        $this->DrawSolidCellFill($canvas, $this->CellRect($page, $row, $col), $style->Alignment, $effect->Fill);
                    }
                } else if ($style !== null && $style->Pattern !== FillPatternKind::None) {
                    $this->DrawRotatedNeighborUnderlay($canvas, $page, $row, $col, $style);
                    $this->DrawPatternCellFill($canvas, $this->CellRect($page, $row, $col), $style);
                }
            }
        }

        $paint->Dispose();

        ConditionalFillRunRenderer::Draw($canvas, $page, $cf, $this->_tableStyleResolver);
        $this->DrawConditionalDataBars($canvas, $page, $cf);
    }

    // Draws conditional-formatting data bars over the cell fills and behind the cell text. Each bar
    // grows from the cell's left edge and uses either the default gradient or a solid fill,
    // according to the loaded rule.
    private function DrawConditionalDataBars($canvas, $page, $cf)
    {
        $layout = $page->Sheet;
        for ($row = $page->StartRow; $row <= $page->EndRow; $row++) {
            if ($layout->RowHidden[$row]) continue;

            for ($col = $page->StartColumn; $col <= $page->EndColumn; $col++) {
                if ($layout->ColumnHidden[$col]) continue;
                if (self::IsPartOfMerge($layout, $row, $col)) continue;

                $effect = null;
                if (!$cf->TryGet($row, $col, $effect) || $effect->Bar === null) {
                    continue;
                }

                $bar = $effect->Bar;
                $rect = $this->CellRect($page, $row, $col);
                $inset = 1.9;
                $maxWidth = $rect->Width - 3.0;
                $width = (float)($maxWidth * $bar->Fraction);
                if ($width <= 0.5) {
                    continue;
                }

                $barRect = new SKRect($rect->Left + $inset, $rect->Top + $inset, $rect->Left + $inset + $width, $rect->Bottom - $inset);
                $lighter = new SKColor(
                    (int)($bar->Color->Red + (255 - $bar->Color->Red) * 0.85),
                    (int)($bar->Color->Green + (255 - $bar->Color->Green) * 0.85),
                    (int)($bar->Color->Blue + (255 - $bar->Color->Blue) * 0.85));

                $shader = $bar->Gradient ? SKShader::CreateLinearGradient(
                    new SKPoint($barRect->Left, 0.0), new SKPoint($barRect->Right, 0.0),
                    array($bar->Color, $lighter), null, SKShaderTileMode::Clamp) : null;

                $paint = new SKPaint();
                $paint->Style = SKPaintStyle::Fill;
                $paint->IsAntialias = true;
                $paint->Color = $bar->Color;
                $paint->Shader = $shader;
                $canvas->DrawRect($barRect, $paint);
                $paint->Dispose();

                if ($shader !== null) {
                    $shader->Dispose();
                }

                if ($bar->ShowBorder) {
                    $border = new SKPaint();
                    $border->Style = SKPaintStyle::Stroke;
                    $border->StrokeWidth = 0.6;
                    $border->IsAntialias = true;
                    $border->Color = $bar->Color;
                    $canvas->DrawRect($barRect, $border);
                    $border->Dispose();
                }
            }
        }
    }

    // Draws conditional-formatting icon-set glyphs at the left of each cell. Icon-set cells are
    // numeric and right-aligned, so the left-anchored icon does not collide with the value.
    private function DrawConditionalIcons($canvas, $page)
    {
        $layout = $page->Sheet;
        $cf = $this->ConditionalFormatFor($layout);
        for ($row = $page->StartRow; $row <= $page->EndRow; $row++) {
            if ($layout->RowHidden[$row]) continue;

            for ($col = $page->StartColumn; $col <= $page->EndColumn; $col++) {
                if ($layout->ColumnHidden[$col]) continue;
                if (self::IsPartOfMerge($layout, $row, $col)) continue;

                $effect = null;
                if (!$cf->TryGet($row, $col, $effect) || $effect->IconGlyph === null) {
                    continue;
                }

                $rect = $this->CellRect($page, $row, $col);
                self::DrawIconGlyph($canvas, $rect, $effect->IconGlyph);
            }
        }
    }

    private static function DrawIconGlyph($canvas, $cell, $icon)
    {
        // Fit the glyph to the row height with a little padding, capped at Excel's ~11pt icon size.
        $size = min($cell->Height - 2.0, 11.0);
        if ($size <= 2.0) {
            return;
        }

        $cx = $cell->Left + 2.0 + $size / 2.0;
        $cy = $cell->MidY;
        $half = $size / 2.0;

        $paint = new SKPaint();
        $paint->IsAntialias = true;
        $paint->Color = $icon->Color;

        if (!$icon->IsArrow) {
            $paint->Style = SKPaintStyle::Fill;
            $canvas->DrawCircle($cx, $cy, $half, $paint);
            if ($icon->HasBorder) {
                $paint->Style = SKPaintStyle::Stroke;
                $paint->StrokeWidth = 0.55;
                $paint->Color = $icon->BorderColor;
                $canvas->DrawCircle($cx, $cy, $half - $paint->StrokeWidth / 2.0, $paint);
            }
            $paint->Dispose();
            return;
        }

        // Build an upward arrow (shaft + head) centred on the origin, then rotate it to the
        // glyph's direction. Canvas y is down, so a positive "up" points to -y.
        $canvas->Save();
        $canvas->Translate($cx, $cy);
        $canvas->RotateDegrees(-$icon->AngleDeg + 90.0); // 90 (up) => 0 rotation
        // Excel's arrow glyphs are chunky (roughly 9 wide x 11 tall for an 11pt icon).
        $shaftW = $size * 0.34;
        $headW = $size * 0.82;
        $headLen = $size * 0.5;
        $paint->Style = SKPaintStyle::Fill;
        $path = new SKPath();
        $path->MoveTo(0.0, -$half);                 // tip (up)
        $path->LineTo($headW / 2.0, -$half + $headLen);
        $path->LineTo($shaftW / 2.0, -$half + $headLen);
        $path->LineTo($shaftW / 2.0, $half);
        $path->LineTo(-$shaftW / 2.0, $half);
        $path->LineTo(-$shaftW / 2.0, -$half + $headLen);
        $path->LineTo(-$headW / 2.0, -$half + $headLen);
        $path->Close();
        $canvas->DrawPath($path, $paint);
        $path->Dispose();

        $canvas->Restore();
        $paint->Dispose();
    }

    private function DrawGridlines($canvas, $page)
    {
        $layout = $page->Sheet;
        $paint = new SKPaint();
        $paint->Color = self::$GridlineColor;
        $paint->Style = SKPaintStyle::Stroke;
        $paint->StrokeWidth = 0.5;
        $paint->IsAntialias = false;

        $top = (float)$this->ContentY($layout, $page->StartRow, $page->StartRow);
        $bottom = (float)$this->ContentY($layout, $page->StartRow, $page->EndRow + 1);
        for ($col = $page->StartColumn; $col <= $page->EndColumn + 1; $col++) {
            $x = (float)$this->ContentX($layout, $page->StartColumn, $col);
            $canvas->DrawLine($x, $top, $x, $bottom, $paint);
        }

        $left = (float)$this->ContentX($layout, $page->StartColumn, $page->StartColumn);
        $right = (float)$this->ContentX($layout, $page->StartColumn, $page->EndColumn + 1);
        for ($row = $page->StartRow; $row <= $page->EndRow + 1; $row++) {
            $y = (float)$this->ContentY($layout, $page->StartRow, $row);
            $canvas->DrawLine($left, $y, $right, $y, $paint);
        }

        $paint->Dispose();
    }

    private function DrawCellBordersAndText($canvas, $page)
    {
        $layout = $page->Sheet;

        for ($row = $page->StartRow; $row <= $page->EndRow; $row++) {
            if ($layout->RowHidden[$row]) continue;

            for ($col = $page->StartColumn; $col <= $page->EndColumn; $col++) {
                if ($layout->ColumnHidden[$col]) continue;
                if (self::IsPartOfMerge($layout, $row, $col)) continue;

                $record = null;
                $addressKey = (new CellAddress($row, $col))->ToString();
                if (array_key_exists($addressKey, $layout->Sheet->Cells)) {
                    $record = $layout->Sheet->Cells[$addressKey];
                }

                if ($record === null) {
                    $emptyStyle = $this->_tableStyleResolver->Resolve($layout->Sheet, $row, $col, StyleValue::Default());
                    if ($emptyStyle === null || !self::HasVisibleBorder($emptyStyle->Borders)) {
                        continue;
                    }

                    $this->DrawBorders($canvas, $this->CellRect($page, $row, $col), $emptyStyle->Borders, self::BorderShearFactor($emptyStyle->Alignment));
                    continue;
                }

                $style = $this->_tableStyleResolver->Resolve($layout->Sheet, $row, $col, $record->Style !== null ? $record->Style : StyleValue::Default());
                $rect = $this->CellRect($page, $row, $col);

                $this->DrawBorders($canvas, $rect, $style->Borders, self::BorderShearFactor($style->Alignment));
                $this->DrawText($canvas, $page, $row, $col, $rect, $record, $style);
            }
        }
    }

    private function CellRect($page, $row, $col)
    {
        $layout = $page->Sheet;

        $lastRow = $row;
        $lastCol = $col;
        $merge = null;
        if ($layout->TryGetMergeOrigin($row, $col, $merge)) {
            $lastRow = min($page->EndRow, $row + max(1, $merge->TotalRows) - 1);
            $lastCol = min($page->EndColumn, $col + max(1, $merge->TotalColumns) - 1);
        }

        $left = (float)$this->ContentX($layout, $page->StartColumn, $col);
        $top = (float)$this->ContentY($layout, $page->StartRow, $row);
        $right = (float)$this->ContentX($layout, $page->StartColumn, $lastCol + 1);
        $bottom = (float)$this->ContentY($layout, $page->StartRow, $lastRow + 1);
        return new SKRect($left, $top, $right, $bottom);
    }

    private static function IsPartOfMerge($layout, $row, $col)
    {
        $merge = null;
        return $layout->IsMergeCovered($row, $col) || $layout->TryGetMergeOrigin($row, $col, $merge);
    }

    // Renders every merged region that intersects this page as a single unit (fill, borders,
    // then origin-cell text), using the region's full rectangle so a merge whose origin sits on
    // an earlier page still paints its continuation here. The active page clip trims overflow.
    private function DrawMerges($canvas, $page)
    {
        $layout = $page->Sheet;

        foreach ($layout->Sheet->MergeRegions as $merge) {
            if (!self::MergeIntersectsPage($page, $merge)) {
                continue;
            }

            $record = null;
            $addressKey = (new CellAddress($merge->FirstRow, $merge->FirstColumn))->ToString();
            if (array_key_exists($addressKey, $layout->Sheet->Cells)) {
                $record = $layout->Sheet->Cells[$addressKey];
            }

            $style = $this->_tableStyleResolver->Resolve($layout->Sheet, $merge->FirstRow, $merge->FirstColumn, $record !== null && $record->Style !== null ? $record->Style : StyleValue::Default());
            $rect = $this->FullMergeRect($page, $merge);

            if ($style->Pattern !== FillPatternKind::None) {
                $this->DrawPatternCellFill($canvas, $rect, $style);
            }

            $this->DrawBorders($canvas, $rect, $style->Borders, self::BorderShearFactor($style->Alignment));

            // Text is drawn only on the page that contains the merge's origin cell. Excel clips
            // merged-cell text at the page break rather than repeating its tail on the
            // continuation page, which otherwise shows only the fill/borders.
            if ($record !== null && self::OriginOnPage($page, $merge)) {
                $this->DrawText($canvas, $page, $merge->FirstRow, $merge->FirstColumn, $rect, $record, $style);
            }
        }
    }

    private static function OriginOnPage($page, $merge)
    {
        return $merge->FirstRow >= $page->StartRow && $merge->FirstRow <= $page->EndRow
            && $merge->FirstColumn >= $page->StartColumn && $merge->FirstColumn <= $page->EndColumn;
    }

    private function DrawSolidCellFill($canvas, $rect, $alignment, $color)
    {
        if ($color->Alpha === 0) {
            return;
        }

        $shear = self::FillShearFactor($alignment);
        if ($shear == 0.0) {
            $paint = new SKPaint();
            $paint->Style = SKPaintStyle::Fill;
            $paint->IsAntialias = false;
            $paint->Color = $color;
            $canvas->DrawRect($rect, $paint);
            $paint->Dispose();

            return;
        }

        $canvas->Save();
        $path = self::CreateShearedCellPath($rect, $shear);
        $paint = new SKPaint();
        $paint->Style = SKPaintStyle::Fill;
        $paint->IsAntialias = false;
        $paint->Color = $color;
        $canvas->DrawPath($path, $paint);
        $paint->Dispose();
        $path->Dispose();

        $canvas->Restore();
    }

    private function DrawPatternCellFill($canvas, $rect, $style)
    {
        $shear = self::FillShearFactor($style->Alignment);
        if ($shear == 0.0) {
            PatternFillRenderer::Draw($canvas, $rect, $style, $this->_context->Colors);
            return;
        }

        $canvas->Save();
        $path = self::CreateShearedCellPath($rect, $shear);
        $canvas->ClipPath($path, SKClipOperation::Intersect, true);
        PatternFillRenderer::Draw($canvas, self::ShearedBounds($rect, $shear), $style, $this->_context->Colors);
        $path->Dispose();

        $canvas->Restore();
    }

    private function DrawRotatedNeighborUnderlay($canvas, $page, $row, $col, $style)
    {
        if ($style === null || $style->Alignment === null) {
            return;
        }

        $rotation = $style->Alignment->TextRotation;
        if ($rotation <= 90 || $rotation >= 180 || $col >= $page->Sheet->LastColumn) {
            return;
        }

        $nextStyle = $this->ResolveCellStyle($page->Sheet->Sheet, $row, $col + 1);
        if ($nextStyle === null || $nextStyle->Pattern !== FillPatternKind::Solid) {
            return;
        }

        $fill = $this->_context->Colors->Resolve($nextStyle->ForegroundColor, SKColors::Transparent());
        if ($fill->Alpha === 0) {
            $fill = $this->_context->Colors->Resolve($nextStyle->BackgroundColor, SKColors::Transparent());
        }

        $this->DrawSolidCellFill($canvas, $this->CellRect($page, $row, $col), null, $fill);
    }

    private function ResolveCellStyle($sheet, $row, $col)
    {
        $record = null;
        $addressKey = (new CellAddress($row, $col))->ToString();
        if (array_key_exists($addressKey, $sheet->Cells)) {
            $record = $sheet->Cells[$addressKey];
        }

        $baseStyle = $record !== null && $record->Style !== null ? $record->Style : StyleValue::Default();
        return $this->_tableStyleResolver->Resolve($sheet, $row, $col, $baseStyle);
    }

    private static function ShearedBounds($rect, $shear)
    {
        $dxTop = -$shear * $rect->Height;
        $left = min($rect->Left, $rect->Left + $dxTop);
        $right = max($rect->Right, $rect->Right + $dxTop);
        return new SKRect($left, $rect->Top, $right, $rect->Bottom);
    }

    private static function FillShearFactor($alignment)
    {
        if ($alignment === null) {
            return 0.0;
        }

        $rotation = $alignment->TextRotation;
        if ($rotation <= 90 || $rotation >= 180) {
            return 0.0;
        }

        return self::BorderShearFactor($alignment);
    }

    private static function CreateShearedCellPath($rect, $shear)
    {
        $dxTop = -$shear * $rect->Height;
        $tl = new SKPoint($rect->Left + $dxTop, $rect->Top);
        $tr = new SKPoint($rect->Right + $dxTop, $rect->Top);
        $br = new SKPoint($rect->Right, $rect->Bottom);
        $bl = new SKPoint($rect->Left, $rect->Bottom);
        $path = new SKPath();
        $path->MoveTo($tl);
        $path->LineTo($tr);
        $path->LineTo($br);
        $path->LineTo($bl);
        $path->Close();
        return $path;
    }

    private static function MergeIntersectsPage($page, $merge)
    {
        $lastRow = $merge->FirstRow + max(1, $merge->TotalRows) - 1;
        $lastCol = $merge->FirstColumn + max(1, $merge->TotalColumns) - 1;
        return $merge->FirstRow <= $page->EndRow && $lastRow >= $page->StartRow
            && $merge->FirstColumn <= $page->EndColumn && $lastCol >= $page->StartColumn;
    }

    // The merged region's full rectangle in this page's content coordinates. Columns/rows that
    // precede the page's start yield negative offsets (clipped away by the page clip), which is
    // what lets a continuation page show only the part of the merge that belongs to it.
    private function FullMergeRect($page, $merge)
    {
        $layout = $page->Sheet;
        $firstRow = $merge->FirstRow;
        $firstCol = $merge->FirstColumn;
        $lastRow = min($layout->LastRow, $merge->FirstRow + max(1, $merge->TotalRows) - 1);
        $lastCol = min($layout->LastColumn, $merge->FirstColumn + max(1, $merge->TotalColumns) - 1);

        $left = (float)$this->ContentX($layout, $page->StartColumn, $firstCol);
        $top = (float)$this->ContentY($layout, $page->StartRow, $firstRow);
        $right = (float)$this->ContentX($layout, $page->StartColumn, $lastCol + 1);
        $bottom = (float)$this->ContentY($layout, $page->StartRow, $lastRow + 1);
        return new SKRect($left, $top, $right, $bottom);
    }

    // Draws each anchored picture that intersects this page at its EMU-precise rectangle. Images
    // are decoded on demand; the active page clip trims any part outside the page.
    private function DrawPictures($canvas, $page)
    {
        $layout = $page->Sheet;

        foreach ($layout->Sheet->Pictures as $picture) {
            if ($picture === null || $picture->ImageData === null || strlen($picture->ImageData) === 0) {
                continue;
            }

            $rect = $this->PictureRect($page, $picture);
            if ($rect->Width <= 0.0 || $rect->Height <= 0.0) {
                continue;
            }

            if (!self::PictureRectIntersectsPage($page, $rect)) {
                continue;
            }

            $image = $this->_context->Pictures->GetImage($picture, $rect->Width, $rect->Height);
            if ($image === null) {
                continue;
            }

            $paint = new SKPaint();
            $paint->IsAntialias = true;
            $paint->FilterQuality = self::PictureFilterQuality($picture);
            $canvas->DrawImage($image, $rect, $paint);
            $paint->Dispose();
        }
    }

    private static function PictureFilterQuality($picture)
    {
        if ($picture === null || $picture->ImageExtension === null || $picture->ImageExtension === '') {
            return SKFilterQuality::High;
        }

        $extension = strtolower(trim($picture->ImageExtension));
        if ($extension === 'png' || $extension === 'gif' || $extension === 'bmp') {
            // UI screenshots and text-heavy raster artwork stay sharper with minimal resampling.
            return SKFilterQuality::None;
        }

        return SKFilterQuality::High;
    }

    private static function PictureRectIntersectsPage($page, $rect)
    {
        $layout = $page->Sheet;
        $contentWidth = (float)($layout->ColumnStartPt[$page->EndColumn + 1] - $layout->ColumnStartPt[$page->StartColumn]);
        $contentHeight = (float)($layout->RowStartPt[$page->EndRow + 1] - $layout->RowStartPt[$page->StartRow]);
        return $rect->Right > 0.0 && $rect->Bottom > 0.0 && $rect->Left < $contentWidth && $rect->Top < $contentHeight;
    }

    // Draws each chart intersecting this page at its EMU anchor rectangle. A chart whose XML is
    // unsupported is skipped. The active page clip trims parts outside this page.
    private function DrawCharts($canvas, $page)
    {
        $layout = $page->Sheet;
        if (count($layout->Sheet->Charts) === 0) {
            return;
        }

        $renderer = new ChartRenderer($this->_context);
        foreach ($layout->Sheet->Charts as $chart) {
            if ($chart === null || $chart->RawChartXml === null || $chart->RawChartXml === '') {
                continue;
            }

            if ($chart->UpperLeftRow > $page->EndRow || $chart->LowerRightRow < $page->StartRow
                || $chart->UpperLeftColumn > $page->EndColumn || $chart->LowerRightColumn < $page->StartColumn) {
                continue;
            }

            if ($chart->IsChartEx) {
                $mapChart = null;
                $mapKey = spl_object_id($chart);
                if (!array_key_exists($mapKey, $this->_mapChartCache)) {
                    $mapChart = MapChartXmlParser::Parse($chart->RawChartXml, $this->_context->Workbook);
                    $this->_mapChartCache[$mapKey] = $mapChart;
                } else {
                    $mapChart = $this->_mapChartCache[$mapKey];
                }

                if ($mapChart !== null) {
                    $mapRenderer = new MapChartRenderer($this->_context);
                    $mapRenderer->Draw($canvas, $this->ChartRect($page, $chart), $mapChart, $page);
                }

                continue;
            }

            $parsed = null;
            $chartKey = spl_object_id($chart);
            if (!array_key_exists($chartKey, $this->_chartCache)) {
                $parsed = ChartXmlParser::Parse($chart->RawChartXml, $this->_context->Colors, $this->DateSystemOf(), $this->_context->Culture, $this->_context->Workbook, $chart);
                $this->_chartCache[$chartKey] = $parsed;
            } else {
                $parsed = $this->_chartCache[$chartKey];
            }

            if ($parsed === null) {
                continue;
            }

            $rect = $this->ChartRect($page, $chart);
            if ($rect->Width > 0.0 && $rect->Height > 0.0) {
                // Keep charts vector-backed in the PDF so thin series strokes, gridlines, and
                // axis labels stay crisp instead of being flattened into a raster snapshot.
                $renderer->Draw($canvas, $page, $rect, $parsed, false);
            }
        }
    }

    // Draws each SmartArt diagram intersecting this page. The diagram's pre-laid-out shapes
    // (from its dsp:drawing part) are positioned relative to the diagram frame's anchor and
    // rendered with the shared shape renderer - no SmartArt layout engine required.
    private function DrawSmartArt($canvas, $page)
    {
        $layout = $page->Sheet;
        if (count($layout->Sheet->SmartArts) === 0) {
            return;
        }

        $renderer = new ShapeRenderer($this->_context);
        foreach ($layout->Sheet->SmartArts as $smartArt) {
            if ($smartArt === null || $smartArt->RawDrawingXml === null || $smartArt->RawDrawingXml === '') {
                continue;
            }

            if ($smartArt->UpperLeftRow > $page->EndRow || $smartArt->LowerRightRow < $page->StartRow
                || $smartArt->UpperLeftColumn > $page->EndColumn || $smartArt->LowerRightColumn < $page->StartColumn) {
                continue;
            }

            $shapes = null;
            $smartArtKey = spl_object_id($smartArt);
            if (!array_key_exists($smartArtKey, $this->_smartArtCache)) {
                $shapes = ShapeXmlParser::ParseSmartArtDrawing($smartArt->RawDrawingXml, $this->_context->Colors);
                $this->_smartArtCache[$smartArtKey] = $shapes;
            } else {
                $shapes = $this->_smartArtCache[$smartArtKey];
            }

            if (count($shapes) === 0) {
                continue;
            }

            $originX = $layout->ColumnStartPt[$page->StartColumn];
            $originY = $layout->RowStartPt[$page->StartRow];

            // Diagram origin: the column offset applies to X, but the dsp shape coordinates
            // already bake in the row offset (their vertical origin is the anchor cell's top),
            // so Y uses the cell top without the row offset.
            $diagramLeft = self::ColumnEdgePt($layout, $smartArt->UpperLeftColumn, $smartArt->UpperLeftColumnOffset) - $originX;
            $diagramTop = $layout->RowStartPt[$smartArt->UpperLeftRow] - $originY;
            $frameRight = self::ColumnEdgePt($layout, $smartArt->LowerRightColumn, $smartArt->LowerRightColumnOffset) - $originX;
            $frameWidth = $frameRight - $diagramLeft;

            // Excel scales the laid-out diagram to fit the frame width.
            $minX = 0.0;
            $minY = 0.0;
            $maxX = 0.0;
            $maxY = 0.0;
            self::CanvasBounds($shapes, $minX, $minY, $maxX, $maxY);
            if ($maxX <= 0.0) {
                continue;
            }

            $scale = $frameWidth > 0.0 ? $frameWidth / $maxX : 1.0;

            foreach ($shapes as $s) {
                if ($s->Shape === null) {
                    continue;
                }

                $left = (float)($diagramLeft + $s->XPt * $scale);
                $top = (float)($diagramTop + $s->YPt * $scale);
                $rect = new SKRect($left, $top, $left + (float)($s->WPt * $scale), $top + (float)($s->HPt * $scale));
                if ($rect->Width > 0.0 && $rect->Height > 0.0) {
                    $renderer->Draw($canvas, $page, $rect, $s->Shape, null, null);
                }
            }
        }
    }

    // The bounding box of the laid-out SmartArt shapes, accounting for each shape's rotation
    // (a 90-degree-rotated shape contributes its swapped width/height).
    private static function CanvasBounds($shapes, &$minX, &$minY, &$maxX, &$maxY)
    {
        $minX = PHP_FLOAT_MAX;
        $minY = PHP_FLOAT_MAX;
        $maxX = -PHP_FLOAT_MAX;
        $maxY = -PHP_FLOAT_MAX;

        foreach ($shapes as $s) {
            if ($s->Shape === null) {
                continue;
            }

            $cx = $s->XPt + $s->WPt / 2.0;
            $cy = $s->YPt + $s->HPt / 2.0;
            $rad = $s->Shape->RotationDeg * M_PI / 180.0;
            $cos = abs(cos($rad));
            $sin = abs(sin($rad));
            $halfW = ($s->WPt * $cos + $s->HPt * $sin) / 2.0;
            $halfH = ($s->WPt * $sin + $s->HPt * $cos) / 2.0;

            if ($cx - $halfW < $minX) $minX = $cx - $halfW;
            if ($cy - $halfH < $minY) $minY = $cy - $halfH;
            if ($cx + $halfW > $maxX) $maxX = $cx + $halfW;
            if ($cy + $halfH > $maxY) $maxY = $cy + $halfH;
        }
    }

    private function DateSystemOf()
    {
        return $this->_context->Workbook !== null && $this->_context->Workbook->Settings !== null
            ? $this->_context->Workbook->Settings->DateSystem
            : DateSystem::Windows1900;
    }

    private function ChartRect($page, $chart)
    {
        $layout = $page->Sheet;
        $originX = $layout->ColumnStartPt[$page->StartColumn];
        $originY = $layout->RowStartPt[$page->StartRow];

        $left = self::ColumnEdgePt($layout, $chart->UpperLeftColumn, $chart->UpperLeftColumnOffset) - $originX;
        $top = self::RowEdgePt($layout, $chart->UpperLeftRow, $chart->UpperLeftRowOffset) - $originY;
        $right = self::ColumnEdgePt($layout, $chart->LowerRightColumn, $chart->LowerRightColumnOffset) - $originX;
        $bottom = self::RowEdgePt($layout, $chart->LowerRightRow, $chart->LowerRightRowOffset) - $originY;
        if ($chart->ExtentCx > 0) {
            $right = $left + $chart->ExtentCx / 12700.0;
        }
        if ($chart->ExtentCy > 0) {
            $bottom = $top + $chart->ExtentCy / 12700.0;
        }
        return new SKRect((float)$left, (float)$top, (float)$right, (float)$bottom);
    }

    // Draws each drawing shape (autoShape or connector) intersecting this page at its anchor
    // rectangle. Shapes whose geometry is unknown fall back to a rectangle.
    private function DrawShapes($canvas, $page)
    {
        $layout = $page->Sheet;
        if (count($layout->Sheet->Shapes) === 0) {
            return;
        }

        $renderer = new ShapeRenderer($this->_context);
        foreach ($layout->Sheet->Shapes as $shape) {
            if ($shape === null) {
                continue;
            }

            if ($shape->UpperLeftRow > $page->EndRow || $shape->LowerRightRow < $page->StartRow
                || $shape->UpperLeftColumn > $page->EndColumn || $shape->LowerRightColumn < $page->StartColumn) {
                continue;
            }

            $layoutRef = $page->Sheet;
            $originX = $layoutRef->ColumnStartPt[$page->StartColumn];
            $originY = $layoutRef->RowStartPt[$page->StartRow];
            $left = self::ColumnEdgePt($layoutRef, $shape->UpperLeftColumn, $shape->UpperLeftColumnOffset) - $originX;
            $top = self::RowEdgePt($layoutRef, $shape->UpperLeftRow, $shape->UpperLeftRowOffset) - $originY;
            $right = self::ColumnEdgePt($layoutRef, $shape->LowerRightColumn, $shape->LowerRightColumnOffset) - $originX;
            $bottom = self::RowEdgePt($layoutRef, $shape->LowerRightRow, $shape->LowerRightRowOffset) - $originY;
            $rect = new SKRect((float)$left, (float)$top, (float)$right, (float)$bottom);
            if ($rect->Width > 0.0 && $rect->Height > 0.0) {
                if ((new GroupShapeRenderer($this->_context))->TryDraw($canvas, $page, $rect, $shape->RawElementXml)) {
                    continue;
                }

                $parsed = null;
                $shapeKey = spl_object_id($shape);
                if (!array_key_exists($shapeKey, $this->_shapeCache)) {
                    $parsed = ShapeXmlParser::Parse($shape->RawElementXml, $shape->GeometryType, $this->_context->Colors);
                    $this->_shapeCache[$shapeKey] = $parsed;
                } else {
                    $parsed = $this->_shapeCache[$shapeKey];
                }

                if ($parsed === null) {
                    continue;
                }

                $renderer->Draw($canvas, $page, $rect, $parsed, $shape, $layout->Sheet->Shapes);
            }
        }
    }

    private function PictureRect($page, $picture)
    {
        $layout = $page->Sheet;
        $originX = $layout->ColumnStartPt[$page->StartColumn];
        $originY = $layout->RowStartPt[$page->StartRow];

        $left = self::ColumnEdgePt($layout, $picture->UpperLeftColumn, $picture->UpperLeftColumnOffset) - $originX;
        $top = self::RowEdgePt($layout, $picture->UpperLeftRow, $picture->UpperLeftRowOffset) - $originY;
        $right = 0.0;
        $bottom = 0.0;
        if ($picture->ExtentCx > 0) {
            $right = $left + $picture->ExtentCx / self::EmuPerPoint;
        } else {
            $right = self::ColumnEdgePt($layout, $picture->LowerRightColumn, $picture->LowerRightColumnOffset) - $originX;
        }

        if ($picture->ExtentCy > 0) {
            $bottom = $top + $picture->ExtentCy / self::EmuPerPoint;
        } else {
            $bottom = self::RowEdgePt($layout, $picture->LowerRightRow, $picture->LowerRightRowOffset) - $originY;
        }

        return new SKRect((float)$left, (float)$top, (float)$right, (float)$bottom);
    }

    private static function ColumnEdgePt($layout, $column, $emuOffset)
    {
        if ($column < 0) $column = 0;
        if ($column > $layout->LastColumn + 1) $column = $layout->LastColumn + 1;
        return $layout->ColumnStartPt[$column] + $emuOffset / self::EmuPerPoint;
    }

    private static function RowEdgePt($layout, $row, $emuOffset)
    {
        if ($row < 0) $row = 0;
        if ($row > $layout->LastRow + 1) $row = $layout->LastRow + 1;
        return $layout->RowStartPt[$row] + $emuOffset / self::EmuPerPoint;
    }

    // Extends a cell's text clip rectangle across consecutive empty neighbor columns (within the
    // page) so long non-wrapped text spills over the way Excel renders it. Left/general text
    // extends right, right-aligned extends left, centered extends both ways.
    private function ExpandClipForOverflow($page, $row, $col, $rect, $horizontal)
    {
        $layout = $page->Sheet;
        $left = $rect->Left;
        $right = $rect->Right;

        $extendRight = $horizontal === HorizontalAlignment::Left
            || $horizontal === HorizontalAlignment::General
            || $horizontal === HorizontalAlignment::Fill
            || $horizontal === HorizontalAlignment::Justify
            || $horizontal === HorizontalAlignment::Center
            || $horizontal === HorizontalAlignment::CenterContinuous;

        $extendLeft = $horizontal === HorizontalAlignment::Right
            || $horizontal === HorizontalAlignment::Center
            || $horizontal === HorizontalAlignment::CenterContinuous;

        if ($extendRight) {
            $c = $col + 1;
            while ($c <= $page->EndColumn && $this->IsCellEmptyForOverflow($layout, $row, $c)) {
                $right = (float)$this->ContentX($layout, $page->StartColumn, $c + 1);
                $c++;
            }
        }

        if ($extendLeft) {
            $c = $col - 1;
            while ($c >= $page->StartColumn && $this->IsCellEmptyForOverflow($layout, $row, $c)) {
                $left = (float)$this->ContentX($layout, $page->StartColumn, $c);
                $c--;
            }
        }

        return new SKRect($left, $rect->Top, $right, $rect->Bottom);
    }

    private function IsCellEmptyForOverflow($layout, $row, $col)
    {
        // Hidden columns have zero printed width and must not stop text in the preceding
        // visible cell from flowing into the next empty visible column.
        if ($layout->ColumnHidden[$col]) {
            return true;
        }

        if ($layout->IsMergeCovered($row, $col)) {
            return false;
        }

        $merge = null;
        if ($layout->TryGetMergeOrigin($row, $col, $merge)) {
            return false;
        }

        $record = null;
        $addressKey = (new CellAddress($row, $col))->ToString();
        if (!array_key_exists($addressKey, $layout->Sheet->Cells) || $layout->Sheet->Cells[$addressKey] === null) {
            return true;
        }

        $record = $layout->Sheet->Cells[$addressKey];
        $style = $record->Style !== null ? $record->Style : StyleValue::Default();
        $text = DisplayTextFormatter::FormatDisplayValue($record->Value, $style, $this->_context->Culture,
            $this->DateSystemOf());
        return $text === null || $text === '';
    }

    private function DrawBorders($canvas, $rect, $borders, $shear = 0.0)
    {
        $dxTop = -$shear * $rect->Height;
        $tlx = $rect->Left + $dxTop;
        $trx = $rect->Right + $dxTop;

        $this->DrawBorderSide($canvas, $borders->Top, $tlx, $rect->Top, $trx, $rect->Top);
        $this->DrawBorderSide($canvas, $borders->Bottom, $rect->Left, $rect->Bottom, $rect->Right, $rect->Bottom);
        $this->DrawBorderSide($canvas, $borders->Left, $tlx, $rect->Top, $rect->Left, $rect->Bottom);
        $this->DrawBorderSide($canvas, $borders->Right, $trx, $rect->Top, $rect->Right, $rect->Bottom);
        if ($borders->DiagonalUp) {
            $this->DrawBorderSide($canvas, $borders->Diagonal, $rect->Left, $rect->Bottom, $trx, $rect->Top);
        }

        if ($borders->DiagonalDown) {
            $this->DrawBorderSide($canvas, $borders->Diagonal, $tlx, $rect->Top, $rect->Right, $rect->Bottom);
        }
    }

    // Excel's border shear factor for a cell with rotated text: side edges become parallel to the
    // text baseline (k = -tan(alpha), alpha = signed counter-clockwise angle). Straight and stacked
    // text leave the box rectangular.
    private static function BorderShearFactor($alignment)
    {
        $rotation = $alignment->TextRotation;
        if ($rotation <= 0 || $rotation >= 255) {
            return 0.0;
        }

        $ccwDeg = $rotation <= 90 ? $rotation : -($rotation - 90);
        return (float)(-tan($ccwDeg * M_PI / 180.0));
    }

    private static function HasVisibleBorder($borders)
    {
        return $borders !== null
            && (($borders->Top !== null && $borders->Top->Style !== BorderStyle::None)
                || ($borders->Bottom !== null && $borders->Bottom->Style !== BorderStyle::None)
                || ($borders->Left !== null && $borders->Left->Style !== BorderStyle::None)
                || ($borders->Right !== null && $borders->Right->Style !== BorderStyle::None)
                || (($borders->DiagonalUp || $borders->DiagonalDown) && $borders->Diagonal !== null && $borders->Diagonal->Style !== BorderStyle::None));
    }

    private function DrawBorderSide($canvas, $side, $x0, $y0, $x1, $y1)
    {
        if ($side === null || $side->Style === BorderStyle::None) {
            return;
        }

        $width = self::BorderWidthPt($side->Style);
        $color = $this->_context->Colors->Resolve($side->Color, SKColors::Black());

        $paint = new SKPaint();
        $paint->Color = $color;
        $paint->Style = SKPaintStyle::Stroke;
        $paint->StrokeWidth = $width;
        $paint->IsAntialias = false;
        $paint->StrokeCap = SKStrokeCap::Butt;

        if ($side->Style === BorderStyle::SlantedDashDot && $x0 != $x1 && $y0 != $y1) {
            $paint->IsAntialias = true;
            $paint->StrokeCap = SKStrokeCap::Round;
        }

        $dash = self::DashFor($side->Style, $width);
        if ($dash !== null) {
            $paint->PathEffect = $dash;
        }

        if ($side->Style === BorderStyle::Double) {
            // Two parallel hairlines with a gap approximates a double border.
            $offset = $width;
            if ($y0 == $y1) {
                $canvas->DrawLine($x0, $y0 - $offset, $x1, $y1 - $offset, $paint);
                $canvas->DrawLine($x0, $y0 + $offset, $x1, $y1 + $offset, $paint);
            } else {
                $canvas->DrawLine($x0 - $offset, $y0, $x1 - $offset, $y1, $paint);
                $canvas->DrawLine($x0 + $offset, $y0, $x1 + $offset, $y1, $paint);
            }
        } else {
            $canvas->DrawLine($x0, $y0, $x1, $y1, $paint);
        }

        if ($dash !== null) {
            $dash->Dispose();
            $paint->PathEffect = null;
        }

        $paint->Dispose();
    }

    private static function BorderWidthPt($style)
    {
        switch ($style) {
            case BorderStyle::Hair:
                return 0.5;
            case BorderStyle::Thin:
            case BorderStyle::Dotted:
            case BorderStyle::Dashed:
            case BorderStyle::DashDot:
            case BorderStyle::DashDotDot:
                return 0.75;
            case BorderStyle::Medium:
            case BorderStyle::MediumDashed:
            case BorderStyle::MediumDashDot:
            case BorderStyle::MediumDashDotDot:
            case BorderStyle::SlantedDashDot:
                return 1.75;
            case BorderStyle::Thick:
                return 2.5;
            case BorderStyle::Double:
                return 0.75;
            default:
                return 0.75;
        }
    }

    private static function DashFor($style, $width)
    {
        switch ($style) {
            case BorderStyle::Dotted:
                return SKPathEffect::CreateDash(array(0.9 * $width, 1.6 * $width), 0.0);
            case BorderStyle::Dashed:
            case BorderStyle::MediumDashed:
                return SKPathEffect::CreateDash(array(4.8 * $width, 2.2 * $width), 0.0);
            case BorderStyle::DashDot:
            case BorderStyle::MediumDashDot:
                return SKPathEffect::CreateDash(array(5.4 * $width, 2.3 * $width, 1.1 * $width, 2.3 * $width), 0.0);
            case BorderStyle::DashDotDot:
            case BorderStyle::MediumDashDotDot:
                return SKPathEffect::CreateDash(array(5.2 * $width, 2.1 * $width, $width, 1.9 * $width, $width, 2.1 * $width), 0.0);
            case BorderStyle::SlantedDashDot:
                return SKPathEffect::CreateDash(array(4.1 * $width, 1.55 * $width, 0.95 * $width, 1.55 * $width), 0.0);
            default:
                return null;
        }
    }

    private function DrawText($canvas, $page, $row, $col, $rect, $record, $style)
    {
        $cf = null;
        $hasCf = $this->ConditionalFormatFor($page->Sheet)->TryGet($row, $col, $cf);
        if ($hasCf && $cf->SuppressText) {
            return;
        }

        if ($hasCf && (($cf->HasFontBold && !$style->Font->Bold)
            || ($cf->HasFontItalic && !$style->Font->Italic)
            || ($cf->HasFontUnderline && $style->Font->Underline != $cf->FontUnderline))) {
            $style = $style->Clone();
            if ($cf->HasFontBold) {
                $style->Font->Bold = true;
            }
            if ($cf->HasFontItalic) {
                $style->Font->Italic = true;
            }
            if ($cf->HasFontUnderline) {
                $style->Font->Underline = $cf->FontUnderline;
            }
        }

        $fontContext = $this->_context->GetFontContext($style->Font);
        $pdfFontMetrics = $style->Font !== null && $style->Font->Name !== null
            && (strcasecmp($style->Font->Name, 'SimSun') === 0
                || $style->Font->Name === "\u{5b8b}\u{4f53}");
        $padding = (float)SheetLayout::HorizontalPaddingPt;
        $baseInnerWidth = $rect->Width - 2.0 * $padding;
        if ($baseInnerWidth <= 0.0) {
            return;
        }

        $horizontal = self::ResolveHorizontal($style->Alignment->Horizontal, $record);
        $indentOffset = self::IndentOffsetPt($fontContext, $style->Alignment, $horizontal);
        $innerWidth = max(0.0, $baseInnerWidth - $indentOffset);
        // Distributed indentation is already accounted for by EffectiveWrapWidthPt;
        // use the same unindented input as automatic row-height measurement.
        $wrapWidth = self::EffectiveWrapWidthPt($horizontal === HorizontalAlignment::Distributed ? $baseInnerWidth : $innerWidth, $style->Alignment);
        $implicitDistributedWrap = $horizontal === HorizontalAlignment::Distributed && !$style->Alignment->WrapText;
        if ($implicitDistributedWrap)
            $wrapWidth = DistributedTextLine::UnwrappedWidth($fontContext, $baseInnerWidth, $style->Alignment);
        if ($innerWidth <= 0.0) {
            return;
        }

        // Excel's General number format is width-aware: it rounds a number to as many digits as
        // fit the cell (capped at 11 significant digits) rather than showing full precision and
        // clipping. Explicit number formats keep their fixed display and clip if too wide.
        if (!$style->Alignment->WrapText && self::IsGeneralNumeric($record->Value, $style)) {
            $text = $this->FitGeneralNumber($record->Value, $innerWidth, $fontContext);
        } else {
            $text = DisplayTextFormatter::FormatDisplayValue($record->Value, $style, $this->_context->Culture,
                $this->DateSystemOf());
        }

        if ($text === null || $text === '') {
            return;
        }

        $color = NumberFormatTextColor::Resolve($record->Value, $style,
            $this->_context->Colors->Resolve($style->Font->Color, SKColors::Black()));
        if ($hasCf && $cf->HasFontColor) {
            $color = $cf->FontColor;
        }

        // Rotated / vertically-stacked cell text takes a dedicated path.
        $rotation = $style->Alignment->TextRotation;
        if ($rotation === 255) {
            $this->DrawStackedCellText($canvas, $rect, self::StripLineBreaks($text), $fontContext, $style, $color);
            return;
        }
        if ($rotation !== 0) {
            $this->DrawRotatedCellText($canvas, $rect, self::StripLineBreaks($text), $fontContext, $style, $rotation, $color);
            return;
        }

        // Wrapped text breaks on explicit newlines; non-wrapped text collapses them the way
        // Excel does (a manual line break renders as zero-width, keeping one visual line)
        // rather than drawing a missing-glyph box for the control character.
        $lines = $style->Alignment->WrapText || $implicitDistributedWrap
            ? $fontContext->WrapLines($text, $wrapWidth, true)
            : array(self::StripLineBreaks($text));

        $lineHeight = $style->Alignment->WrapText || $implicitDistributedWrap ? $fontContext->WrappedLineHeightPt : $fontContext->LineHeightPt;
        $blockHeight = count($lines) * $lineHeight;

        $blockTop = 0.0;
        switch ($style->Alignment->Vertical) {
            case VerticalAlignment::Top:
                $blockTop = $rect->Top + (float)SheetLayout::VerticalPaddingPt;
                break;
            case VerticalAlignment::Center:
                $blockTop = $rect->Top + ($rect->Height - $blockHeight) / 2.0;
                break;
            default: // Bottom / Justify / Distributed
                $blockTop = $rect->Bottom - (float)SheetLayout::VerticalPaddingPt - $blockHeight;
                if ($pdfFontMetrics) {
                    // Excel places SimSun's print baseline lower than its nominal Skia
                    // line box. Keep the row height; adjust only bottom-aligned ink.
                    $blockTop += $fontContext->SizePt * (3.2 / 11.0);
                }
                break;
        }

        // Non-wrapped text that is wider than its cell overflows into adjacent EMPTY cells, the
        // way Excel renders it; wrapped text and text with occupied neighbors stay clipped.
        $clip = $rect;
        if (!$style->Alignment->WrapText && !$implicitDistributedWrap && count($lines) === 1) {
            $needed = ($pdfFontMetrics ? $fontContext->MeasurePdfText($lines[0]) : $fontContext->Measure($lines[0]))
                + 2.0 * $padding + $indentOffset;
            if ($needed > $rect->Width) {
                $clip = $this->ExpandClipForOverflow($page, $row, $col, $rect, $horizontal);
            }
        }

        if ($this->_richTextRenderer->CanRender($record, $style, $text)) {
            if ($this->_richTextRenderer->TryDraw($canvas, $page, $rect, $clip, $record, $style, $horizontal, $hasCf && $cf->HasFontColor, $color)) {
                return;
            }
        }

        $canvas->Save();
        $canvas->ClipRect($clip);

        $paint = new SKPaint();
        $paint->Color = $color;
        $paint->IsAntialias = true;

        $metricsFont = new SKFont($this->_context->Fonts->Resolve($style->Font), $fontContext->SizePt);
        $metrics = $metricsFont->Metrics;
        $metricsFont->Dispose();

        $wrappedTextOffset = 0;
        for ($i = 0; $i < count($lines); $i++) {
            $line = $lines[$i];
            $lineTop = $blockTop + $i * $lineHeight;
            $baseline = $lineTop - $metrics->Ascent; // Ascent is negative in Skia.

            $sourceStart = strpos($text, $line, $wrappedTextOffset);
            if ($sourceStart !== false) $wrappedTextOffset = $sourceStart + strlen($line);
            if ($horizontal === HorizontalAlignment::Distributed) {
                $trailingSeparator = $wrappedTextOffset < strlen($text)
                    && ($text[$wrappedTextOffset] === ' ' || $text[$wrappedTextOffset] === "\t");
                $distributed = DistributedTextLine::Build($fontContext, $line, $wrapWidth, $trailingSeparator);
                $left = $rect->MidX - $wrapWidth / 2.0;
                for ($wordIndex = 0; $wordIndex < count($distributed->Words); $wordIndex++) {
                    $word = $distributed->Words[$wordIndex];
                    $wordX = $left + $distributed->Offsets[$wordIndex];
                    if (!$this->TryRecordRuns($page, $clip, $fontContext, $word, $wordX, $baseline, $paint->Color, true))
                        $this->DrawRuns($canvas, $fontContext, $word, $wordX, $baseline, $paint, $style->Font);
                    else
                        $this->DrawDecorations($canvas, $fontContext, $word, $wordX, $baseline, $paint, $style->Font);
                }
                continue;
            }

            $formatCode = NumberFormat::ResolveFormatCode($style->NumberFormat->Number, $style->NumberFormat->Custom);
            $accountingSymbol = '';
            if ($i === 0 && strpos($formatCode, '*') !== false && strlen($line) > 1
                && !ctype_digit($line[0]) && (ctype_digit($line[1]) || $line[1] === '-')) {
                $accountingSymbol = substr($line, 0, 1);
                $line = substr($line, 1);
            }

            $usePdfAdvances = $pdfFontMetrics || $style->Alignment->WrapText;
            $lineWidth = $usePdfAdvances ? $fontContext->MeasurePdfText($line) : $fontContext->Measure($line);
            $lineScale = self::ShrinkScale($style->Alignment, $lineWidth, $innerWidth);
            $renderedWidth = $lineWidth * $lineScale;
            $x = 0.0;
            switch ($horizontal) {
                case HorizontalAlignment::Distributed:
                    $x = $style->Alignment->WrapText && $style->Alignment->IndentLevel > 0
                        ? $rect->Left + ($rect->Width - $renderedWidth) / 2.0
                        : $rect->Left + $padding + $indentOffset;
                    break;
                case HorizontalAlignment::Center:
                case HorizontalAlignment::CenterContinuous:
                    $x = $rect->Left + ($rect->Width - $renderedWidth) / 2.0;
                    break;
                case HorizontalAlignment::Right:
                    $x = $rect->Right - $padding - $indentOffset - $renderedWidth;
                    break;
                default:
                    $x = $rect->Left + $padding + $indentOffset;
                    break;
            }

            if ($lineScale < 0.999) {
                $canvas->Save();
                $canvas->Translate($x, $baseline);
                $canvas->Scale($lineScale, $lineScale);
                $this->DrawRuns($canvas, $fontContext, $line, 0.0, 0.0, $paint, $style->Font);
                $canvas->Restore();
            } else if (!$this->TryRecordRuns($page, $clip, $fontContext, $line, $x, $baseline, $paint->Color, $usePdfAdvances)) {
                $this->DrawRuns($canvas, $fontContext, $line, $x, $baseline, $paint, $style->Font);
            } else {
                $this->DrawDecorations($canvas, $fontContext, $line, $x, $baseline, $paint, $style->Font);
            }

            if (strlen($accountingSymbol) > 0) {
                $symbolX = $rect->Left + $padding + $indentOffset;
                if (!$this->TryRecordRuns($page, $clip, $fontContext, $accountingSymbol, $symbolX, $baseline, $paint->Color, $pdfFontMetrics)) {
                    $this->DrawRuns($canvas, $fontContext, $accountingSymbol, $symbolX, $baseline, $paint, $style->Font);
                } else {
                    $this->DrawDecorations($canvas, $fontContext, $accountingSymbol, $symbolX, $baseline, $paint, $style->Font);
                }
            }
        }

        $paint->Dispose();
        $canvas->Restore();
    }

    private static function EffectiveWrapWidthPt($availableWidth, $alignment)
    {
        if ($availableWidth <= 0.0 || $alignment === null) {
            return $availableWidth;
        }

        if ($alignment->Horizontal === HorizontalAlignment::Distributed) {
            $indentUnits = $alignment->IndentLevel + max(0, $alignment->RelativeIndent);
            if ($indentUnits > 0) {
                return $availableWidth / (1.0 + 2.0 * $indentUnits);
            }
        }

        return $availableWidth;
    }

    private static function ShrinkScale($alignment, $textWidth, $availableWidth)
    {
        if ($alignment === null || !$alignment->ShrinkToFit) {
            return 1.0;
        }

        if ($textWidth <= 0.0 || $availableWidth <= 0.0) {
            return 1.0;
        }

        if ($textWidth <= $availableWidth) {
            return 1.0;
        }

        $scale = $availableWidth / $textWidth;
        if ($scale < 0.1) {
            return 0.1;
        }

        return $scale;
    }

    private static function IndentOffsetPt($fontContext, $alignment, $horizontal)
    {
        if ($fontContext === null || $alignment === null) {
            return 0.0;
        }

        if ($alignment->IndentLevel <= 0) {
            return 0.0;
        }

        if ($horizontal !== HorizontalAlignment::Left
            && $horizontal !== HorizontalAlignment::Right
            && $horizontal !== HorizontalAlignment::Distributed
            && $horizontal !== HorizontalAlignment::Justify) {
            return 0.0;
        }

        $digitWidth = $fontContext->Measure('0');
        $indentStep = max($fontContext->SizePt * 0.68, $digitWidth * 1.15);
        return $alignment->IndentLevel * $indentStep;
    }

    private function TryRecordRuns($page, $clip, $fontContext, $line, $x, $baseline, $color, $pdfFontMetrics)
    {
        $session = $this->_context->PdfTextSession;
        if ($session === null || !$this->_context->EnableWorksheetTextOptimization || $line === null || $line === '') {
            return false;
        }

        $runs = $pdfFontMetrics ? $fontContext->SplitPdfRuns($line) : $fontContext->SplitRuns($line);
        for ($i = 0; $i < count($runs); $i++) {
            if (!self::SupportsType3Run($runs[$i]->Typeface)) {
                return false;
            }
        }

        $cursor = $x;
        for ($i = 0; $i < count($runs); $i++) {
            $run = $runs[$i];
            $session->RecordWorksheetRun($page, $clip, $run->Text, $run->Typeface, $color, $fontContext->SizePt, $cursor, $baseline);
            $cursor += $run->WidthPt;
        }

        return true;
    }

    private static function SupportsType3Run($typeface)
    {
        if ($typeface === null || $typeface->FamilyName === null || $typeface->FamilyName === '') {
            return false;
        }

        // Type3 glyph outlines use the run's coordinate-system matrix, not a
        // font-family-specific orientation. DengXian can share glyphs too.
        return true;
    }

    private function DrawRuns($canvas, $fontContext, $line, $x, $baseline, $paint, $font)
    {
        $runs = $fontContext->SplitRuns($line);
        $cursor = $x;

        for ($i = 0; $i < count($runs); $i++) {
            $run = $runs[$i];
            $runFont = new SKFont($run->Typeface, $fontContext->SizePt);
            $runFont->Subpixel = true;
            PdfTextPathRenderer::DrawText($canvas, $run->Text, $cursor, $baseline, $runFont, $paint->Color);

            if ($font !== null && ($font->Underline !== FontUnderlineType::None)) {
                self::DrawUnderline($canvas, $cursor, $baseline, $cursor + $run->WidthPt, $fontContext->SizePt, $paint->Color);
            }

            if ($font !== null && $font->StrikeThrough) {
                self::DrawStrikeThrough($canvas, $cursor, $baseline, $cursor + $run->WidthPt, $fontContext->SizePt, $paint->Color);
            }

            $runFont->Dispose();
            $cursor += $run->WidthPt;
        }
    }

    private function DrawDecorations($canvas, $fontContext, $line, $x, $baseline, $paint, $font)
    {
        if ($font === null) {
            return;
        }

        if ($font->Underline === FontUnderlineType::None && !$font->StrikeThrough) {
            return;
        }

        $runs = $fontContext->SplitRuns($line);
        $cursor = $x;
        for ($i = 0; $i < count($runs); $i++) {
            $run = $runs[$i];
            if ($font->Underline !== FontUnderlineType::None) {
                self::DrawUnderline($canvas, $cursor, $baseline, $cursor + $run->WidthPt, $fontContext->SizePt, $paint->Color);
            }

            if ($font->StrikeThrough) {
                self::DrawStrikeThrough($canvas, $cursor, $baseline, $cursor + $run->WidthPt, $fontContext->SizePt, $paint->Color);
            }

            $cursor += $run->WidthPt;
        }
    }

    private static function DrawUnderline($canvas, $x0, $baseline, $x1, $fontSizePt, $color)
    {
        $paint = new SKPaint();
        $paint->Color = $color;
        $paint->Style = SKPaintStyle::Stroke;
        $paint->IsAntialias = true;
        $paint->StrokeWidth = max(0.6, $fontSizePt * 0.055);
        $y = $baseline + $fontSizePt * 0.18;
        $canvas->DrawLine($x0, $y, $x1, $y, $paint);
        $paint->Dispose();
    }

    private static function DrawStrikeThrough($canvas, $x0, $baseline, $x1, $fontSizePt, $color)
    {
        $paint = new SKPaint();
        $paint->Color = $color;
        $paint->Style = SKPaintStyle::Stroke;
        $paint->IsAntialias = true;
        $paint->StrokeWidth = max(0.55, $fontSizePt * 0.05);
        $y = $baseline - $fontSizePt * 0.28;
        $canvas->DrawLine($x0, $y, $x1, $y, $paint);
        $paint->Dispose();
    }

    // Draws a single line of cell text rotated by Excel's textRotation value (1-90 =
    // counter-clockwise degrees above the horizon, 91-180 = clockwise degrees below it). The
    // rotated text block is aligned within the cell per the cell's horizontal/vertical alignment.
    private function DrawRotatedCellText($canvas, $rect, $text, $fontContext, $style, $rotation, $color)
    {
        // Excel angle (counter-clockwise positive); canvas rotation is clockwise-positive (y-down).
        $ccwDeg = $rotation <= 90 ? $rotation : -($rotation - 90);
        $canvasDeg = (float)(-$ccwDeg);
        $isClockwiseDiagonal = $rotation > 90 && $rotation < 180;

        $padding = (float)SheetLayout::HorizontalPaddingPt;

        $paint = new SKPaint();
        $paint->Color = $color;
        $paint->IsAntialias = true;

        $metricsFont = new SKFont($this->_context->Fonts->Resolve($style->Font), $fontContext->SizePt);
        $metrics = $metricsFont->Metrics;
        $metricsFont->Dispose();

        $textW = $fontContext->Measure($text);
        $ascent = -$metrics->Ascent;
        $descent = $metrics->Descent;

        $rad = $ccwDeg * M_PI / 180.0;
        $absCos = (float)abs(cos($rad));
        $absSin = (float)abs(sin($rad));

        // A rotated cell is skewed into a parallelogram only when it actually has a border to
        // draw; a borderless rotated cell keeps its upright rectangle (and its text may spill
        // into neighbours), so the shear-based centring and clip apply only when bordered.
        $shear = self::HasVisibleBorder($style->Borders) ? self::BorderShearFactor($style->Alignment) : 0.0;

        // The cell can hold text up to this length along the rotation direction; beyond it the
        // text overflows. Excel centers text that fits and left/bottom-anchors text that does not.
        $capacity = min(
            $absCos > 1e-4 ? ($rect->Width - 2.0 * $padding) / $absCos : PHP_FLOAT_MAX,
            $absSin > 1e-4 ? $rect->Height / $absSin : PHP_FLOAT_MAX);
        if ($isClockwiseDiagonal && $shear != 0.0) {
            // Excel is much less willing to center clockwise diagonal text (91-179): the
            // sheared visible box behaves like a narrower slot, so strings such as "Rot 135"
            // overflow into the reading-start corner instead of staying fully centered.
            $capacity *= 0.55;
        }

        $anchorX = 0.0;
        $anchorY = 0.0;
        $tx = 0.0;
        $baseline = 0.0;
        $allowClockwiseOverflow = $isClockwiseDiagonal && $shear == 0.0;
        $centeredClockwise = $allowClockwiseOverflow
            && ($style->Alignment->Horizontal === HorizontalAlignment::Center
                || $style->Alignment->Horizontal === HorizontalAlignment::CenterContinuous)
            && $style->Alignment->Vertical === VerticalAlignment::Center
            && $textW * $absCos + ($ascent + $descent) * $absSin <= $rect->Width - 2.0 * $padding
            && $textW * $absSin + ($ascent + $descent) * $absCos <= $rect->Height;
        if ($centeredClockwise) {
            // Fitting centered text belongs at the centroid of the same parallelogram
            // used by the clockwise cell fill, not at its top overflow anchor.
            $anchorX = $rect->MidX - self::FillShearFactor($style->Alignment) * $rect->Height / 2.0;
            $anchorY = $rect->MidY;
            $tx = -$textW / 2.0;
            $baseline = ($ascent - $descent) / 2.0;
        } else if ($textW <= $capacity && !$allowClockwiseOverflow) {
            // Fits: center the block. Horizontally, a bordered (sheared) cell places the text a
            // fifth of the shear extent left of the upright centre - between that centre and the
            // parallelogram centroid - so it clears both slanted side borders (matching Excel).
            // Vertically, Excel centres the visible glyphs rather than the nominal line box, which
            // sits half the ascender/descender gap below the row centre for ascender-only text.
            $baseline = ($ascent - $descent) / 2.0;
            $anchorX = $rect->MidX - $shear * $rect->Height / 5.0;
            $anchorY = $rect->MidY + $baseline;
            $tx = -$textW / 2.0;
        } else {
            // Overflows: pin the text's rotated bounding box by its reading-start corner so the
            // start stays visible and the glyphs (whose ascenders lean across the baseline) do not
            // spill past the cell edge. Clockwise text (down-right) starts at the top edge,
            // counter-clockwise (up-right) at the bottom.
            $vpad = (float)SheetLayout::VerticalPaddingPt;
            $cos = (float)cos($canvasDeg * M_PI / 180.0);
            $sin = (float)sin($canvasDeg * M_PI / 180.0);

            // Axis-aligned bounds of the layout box (x in [0, textW], y in [-ascent, descent],
            // baseline at y = 0) after rotation about the origin.
            $minX = 0.0;
            $minY = 0.0;
            $maxX = 0.0;
            $maxY = 0.0;
            $first = true;
            foreach (array(0.0, $textW) as $lx) {
                foreach (array(-$ascent, $descent) as $ly) {
                    $rx = $lx * $cos - $ly * $sin;
                    $ry = $lx * $sin + $ly * $cos;
                    if ($first) {
                        $minX = $rx; $maxX = $rx;
                        $minY = $ry; $maxY = $ry;
                        $first = false;
                    } else {
                        if ($rx < $minX) $minX = $rx;
                        if ($rx > $maxX) $maxX = $rx;
                        if ($ry < $minY) $minY = $ry;
                        if ($ry > $maxY) $maxY = $ry;
                    }
                }
            }

            $startX = 0.0;
            $cellCornerX = 0.0;
            switch ($style->Alignment->Horizontal) {
                case HorizontalAlignment::Center:
                case HorizontalAlignment::CenterContinuous:
                    if ($isClockwiseDiagonal) {
                        $startX = $maxX;
                        $cellCornerX = $rect->Left + $padding + $rect->Width * 0.30;
                    } else {
                        $startX = ($minX + $maxX) / 2.0;
                        $cellCornerX = $rect->MidX;
                    }
                    break;
                case HorizontalAlignment::Right:
                    $startX = $maxX; $cellCornerX = $rect->Right - $padding;
                    break;
                default:
                    $startX = $minX; $cellCornerX = $rect->Left + $padding;
                    break;
            }

            // Horizontally the box is pinned by its edge (so leaning glyphs stay inside the cell),
            // but vertically Excel puts the first character's baseline on the reading edge and lets
            // ascenders/descenders spill past it - bottom edge for counter-clockwise text, top for
            // clockwise.
            $cellCornerY = $ccwDeg >= 0 ? $rect->Bottom - $vpad : $rect->Top + $vpad;

            $anchorX = $cellCornerX - $startX;
            $anchorY = $cellCornerY;
            $tx = 0.0;
            $baseline = 0.0;
        }

        $canvas->Save();
        // Clip to the (possibly sheared) cell outline so rotated text is trimmed to the same
        // parallelogram the border draws, not to the upright grid rectangle.
        if ($shear != 0.0) {
            $dxTop = -$shear * $rect->Height;
            $clip = new SKPath();
            $clip->MoveTo($rect->Left + $dxTop, $rect->Top);
            $clip->LineTo($rect->Right + $dxTop, $rect->Top);
            $clip->LineTo($rect->Right, $rect->Bottom);
            $clip->LineTo($rect->Left, $rect->Bottom);
            $clip->Close();
            $canvas->ClipPath($clip, SKClipOperation::Intersect, true);
            $clip->Dispose();
        }
        $canvas->Translate($anchorX, $anchorY);
        $canvas->RotateDegrees($canvasDeg);
        $this->DrawRuns($canvas, $fontContext, $text, $tx, $baseline, $paint, $style->Font);
        $canvas->Restore();

        $paint->Dispose();
    }

    // Draws vertically-stacked cell text (textRotation 255): each character is upright and placed
    // below the previous one, the block aligned within the cell.
    private function DrawStackedCellText($canvas, $rect, $text, $fontContext, $style, $color)
    {
        if ($text === null || $text === '') {
            return;
        }

        $padding = (float)SheetLayout::HorizontalPaddingPt;
        $glyphAdvance = self::StackedGlyphAdvancePt($fontContext);
        // Walk UTF-8 characters, not bytes, so CJK glyphs stack one per row exactly like the
        // C# char-by-char loop does.
        $chars = preg_split('//u', $text, -1, PREG_SPLIT_NO_EMPTY);
        if ($chars === false) {
            $chars = str_split($text);
        }

        $visibleLength = 0;
        foreach ($chars as $ch) {
            if ($ch !== ' ') {
                $visibleLength++;
            }
        }

        if ($visibleLength === 0) {
            return;
        }

        $availableHeight = $rect->Height - 2.0 * (float)SheetLayout::VerticalPaddingPt;
        $maxRowsPerColumn = $availableHeight > 0.0
            ? max(1, (int)floor(($availableHeight + $glyphAdvance * 0.15) / $glyphAdvance))
            : $visibleLength;
        if ($maxRowsPerColumn < 1) {
            $maxRowsPerColumn = 1;
        }

        $columnCount = intdiv($visibleLength + $maxRowsPerColumn - 1, $maxRowsPerColumn);
        $rowsInFirstColumn = min($visibleLength, $maxRowsPerColumn);
        $blockHeight = $rowsInFirstColumn * $glyphAdvance;
        $columnAdvance = self::StackedColumnAdvancePt($fontContext);

        $top = 0.0;
        switch ($style->Alignment->Vertical) {
            case VerticalAlignment::Top:
                $top = $rect->Top + (float)SheetLayout::VerticalPaddingPt;
                break;
            case VerticalAlignment::Center:
                $top = $rect->Top + ($rect->Height - $blockHeight) / 2.0;
                break;
            default:
                $top = $rect->Bottom - (float)SheetLayout::VerticalPaddingPt - $blockHeight;
                break;
        }

        $blockWidth = $columnCount * $columnAdvance;
        $left = 0.0;
        switch ($style->Alignment->Horizontal) {
            case HorizontalAlignment::Left:
                $left = $rect->Left + $padding;
                break;
            case HorizontalAlignment::Right:
                $left = $rect->Right - $padding - $blockWidth;
                break;
            default:
                $left = $rect->MidX - $blockWidth / 2.0;
                break;
        }

        $paint = new SKPaint();
        $paint->Color = $color;
        $paint->IsAntialias = true;

        $metricsFont = new SKFont($this->_context->Fonts->Resolve($style->Font), $fontContext->SizePt);
        $metrics = $metricsFont->Metrics;
        $metricsFont->Dispose();
        $ascent = -$metrics->Ascent;

        $canvas->Save();
        $canvas->ClipRect($rect);
        $visibleIndex = 0;
        foreach ($chars as $ch) {
            if ($ch === ' ') {
                continue;
            }

            $w = $fontContext->Measure($ch);
            $columnIndex = intdiv($visibleIndex, $maxRowsPerColumn);
            $rowIndex = $visibleIndex % $maxRowsPerColumn;
            $visualColumnIndex = $columnCount - 1 - $columnIndex;
            $centerX = $left + $visualColumnIndex * $columnAdvance + $columnAdvance / 2.0;
            $baseline = $top + $rowIndex * $glyphAdvance + $ascent;
            $this->DrawRuns($canvas, $fontContext, $ch, $centerX - $w / 2.0, $baseline, $paint, $style->Font);
            $visibleIndex++;
        }
        $canvas->Restore();

        $paint->Dispose();
    }

    private static function StackedGlyphAdvancePt($fontContext)
    {
        $textHeight = $fontContext->TextHeightPt;
        if ($textHeight <= 0.0) {
            return $fontContext->LineHeightPt;
        }

        return $textHeight + max(0.6, $fontContext->SizePt * 0.08);
    }

    private static function StackedColumnAdvancePt($fontContext)
    {
        return self::StackedGlyphAdvancePt($fontContext);
    }

    // True when the cell holds a plain number displayed with the General format, which is the
    // only case where Excel adapts the number of shown digits to the column width.
    private static function IsGeneralNumeric($value, $style)
    {
        if (!self::IsPlainNumber($value)) {
            return false;
        }

        $formatCode = NumberFormat::ResolveFormatCode($style->NumberFormat->Number, $style->NumberFormat->Custom);
        return $formatCode === null || trim($formatCode) === '' || $formatCode === 'General';
    }

    private static function IsPlainNumber($value)
    {
        return is_int($value) || is_float($value);
    }

    // Formats a number under Excel's General rules. Excel caps General display at 11 characters
    // (including sign, decimal point, and exponent) and additionally reduces precision to fit
    // the cell width, so precision is dropped until both limits are satisfied.
    private function FitGeneralNumber($value, $maxWidthPt, $font)
    {
        $generalMaxChars = 11;
        $number = self::ToDouble($value);

        for ($significantDigits = 11; $significantDigits >= 1; $significantDigits--) {
            $candidate = sprintf('%.' . $significantDigits . 'G', $number);
            if (strlen($candidate) <= $generalMaxChars && $font->Measure($candidate) <= $maxWidthPt) {
                return $candidate;
            }
        }

        // Nothing fits (extremely narrow cell); the caller clips the least-precise form.
        return sprintf('%.1G', $number);
    }

    private static function ToDouble($value)
    {
        if (is_float($value)) return $value;
        if (is_int($value)) return $value;
        return (float)$value;
    }

    private static function StripLineBreaks($text)
    {
        if (strpos($text, "\n") === false && strpos($text, "\r") === false) {
            return $text;
        }

        return str_replace(array("\r\n", "\r", "\n"), '', $text);
    }

    private static function ResolveHorizontal($declared, $record)
    {
        if ($declared !== HorizontalAlignment::General) {
            return $declared;
        }

        switch ($record->Kind) {
            case CellValueKind::Number:
            case CellValueKind::DateTime:
                return HorizontalAlignment::Right;
            case CellValueKind::Boolean:
            case CellValueKind::Error:
                return HorizontalAlignment::Center;
            default:
                if (is_string($record->Value)) {
                    return HorizontalAlignment::Left;
                }

                // Formula results / untyped numerics: right-align numbers, left-align the rest.
                if (DisplayTextFormatterSupport::IsNumericValue($record->Value) || $record->Value instanceof \DateTime) {
                    return HorizontalAlignment::Right;
                }

                return HorizontalAlignment::Left;
        }
    }
}
