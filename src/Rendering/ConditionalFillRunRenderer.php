<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\MergeRegion;
use Aspose\Cells_FOSS\Core\StyleValue;
use Aspose\Cells_FOSS\Skia\SKCanvas;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKPaintStyle;
use Aspose\Cells_FOSS\Skia\SKRect;

final class ConditionalFillRunRenderer
{
    private function __construct()
    {
    }

    public static function Draw($canvas, $page, $formatting, $tableStyles)
    {
        $layout = $page->Sheet;
        $count = 0;
        $paint = new SKPaint();
        try {
            $paint->Style = SKPaintStyle::Fill;
            $paint->IsAntialias = false;

            for ($column = $page->StartColumn; $column <= $page->EndColumn; $column++) {
                $firstRow = -1;
                $runColor = SKColors::Transparent();
                for ($row = $page->StartRow; $row <= $page->EndRow + 1; $row++) {
                    $color = SKColors::Transparent();
                    $hasFill = $row <= $page->EndRow
                        && self::TryGetFill($layout, $formatting, $tableStyles, $row, $column, $color);
                    if ($firstRow >= 0 && (!$hasFill || !($color->Red === $runColor->Red
                            && $color->Green === $runColor->Green
                            && $color->Blue === $runColor->Blue
                            && $color->Alpha === $runColor->Alpha))) {
                        $left = (float)($layout->ColumnStartPt[$column]
                            - $layout->ColumnStartPt[$page->StartColumn]);
                        $right = (float)($layout->ColumnStartPt[$column + 1]
                            - $layout->ColumnStartPt[$page->StartColumn]);
                        $top = (float)($layout->RowStartPt[$firstRow]
                            - $layout->RowStartPt[$page->StartRow]);
                        $bottom = (float)($layout->RowStartPt[$row]
                            - $layout->RowStartPt[$page->StartRow]);
                        $paint->Color = $runColor;
                        $canvas->DrawRect(new SKRect($left, $top, $right, $bottom), $paint);
                        $count++;
                        $firstRow = -1;
                    }

                    if ($hasFill && $firstRow < 0) {
                        $firstRow = $row;
                        $runColor = $color;
                    }
                }
            }
        } finally {
            $paint->Dispose();
        }

        return $count;
    }

    private static function TryGetFill($layout, $formatting, $tableStyles, $row, $column, &$color)
    {
        $color = SKColors::Transparent();
        if ($layout->RowHidden[$row] || $layout->ColumnHidden[$column]) {
            return false;
        }

        $merge = null;
        if ($layout->IsMergeCovered($row, $column)
            || $layout->TryGetMergeOrigin($row, $column, $merge)) {
            return false;
        }

        $effect = null;
        if (!$formatting->TryGet($row, $column, $effect) || !$effect->HasFill
            || $effect->Fill->Alpha === 0) {
            return false;
        }

        $key = (new CellAddress($row, $column))->ToString();
        $record = array_key_exists($key, $layout->Sheet->Cells) ? $layout->Sheet->Cells[$key] : null;
        $baseStyle = $record !== null && $record->Style !== null ? $record->Style : StyleValue::Default();
        $style = $tableStyles->Resolve($layout->Sheet, $row, $column, $baseStyle);
        if ($style !== null && $style->Alignment !== null
            && $style->Alignment->TextRotation > 90 && $style->Alignment->TextRotation < 180) {
            return false;
        }

        $color = $effect->Fill;
        return true;
    }
}
