<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\MergeRegion;
use Aspose\Cells_FOSS\Core\PageOrientation;
use Aspose\Cells_FOSS\PdfSaveOptions;

final class Paginator
{
    private function __construct()
    {
    }

    public static function Paginate($layout, $sheetIndex, $options)
    {
        $areas = PrintAreaRanges::Read($layout->Sheet->PageSetup->PrintArea);
        if (count($areas) === 0) {
            $areas[] = new MergeRegion($layout->FirstRow, $layout->FirstColumn,
                $layout->LastRow - $layout->FirstRow + 1, $layout->LastColumn - $layout->FirstColumn + 1);
        }
        $pages = array();
        foreach ($areas as $area) {
            foreach (self::PaginateArea($layout, $sheetIndex, $options, $area) as $page) {
                $pages[] = $page;
            }
        }
        return $pages;
    }

    private static function PaginateArea($layout, $sheetIndex, $options, $area)
    {
        $pages = array();
        if ($layout->IsEmpty) {
            return $pages;
        }

        if ($options === null) {
            $options = new PdfSaveOptions();
        }

        $setup = $layout->Sheet->PageSetup;

        $pageWidthPt = 0.0;
        $pageHeightPt = 0.0;
        PaperSizes::GetDimensionsPoints($setup->PaperSize, $pageWidthPt, $pageHeightPt);
        if ($setup->Orientation === PageOrientation::Landscape) {
            $swap = $pageWidthPt; $pageWidthPt = $pageHeightPt; $pageHeightPt = $swap;
        }

        $leftPt = RenderUnits::InchesToPoints($setup->Margins->Left);
        $rightPt = RenderUnits::InchesToPoints($setup->Margins->Right);
        $topPt = RenderUnits::InchesToPoints($setup->Margins->Top);
        $bottomPt = RenderUnits::InchesToPoints($setup->Margins->Bottom);

        $usableWidthPt = max(1.0, $pageWidthPt - $leftPt - $rightPt);
        $usableHeightPt = max(1.0, $pageHeightPt - $topPt - $bottomPt);

        $firstRow = $area->FirstRow;
        $firstColumn = $area->FirstColumn;
        $lastRow = $firstRow + $area->TotalRows - 1;
        $lastColumn = $firstColumn + $area->TotalColumns - 1;
        $totalWidthPt = $layout->ColumnStartPt[$lastColumn + 1] - $layout->ColumnStartPt[$firstColumn];
        $totalHeightPt = $layout->RowStartPt[$lastRow + 1] - $layout->RowStartPt[$firstRow];

        $forceOnePageWide = $options->OnePagePerSheet || $options->AllColumnsInOnePagePerSheet;
        $forceOnePageTall = $options->OnePagePerSheet;

        $scale = self::ResolveScale($setup, $options, $totalWidthPt, $totalHeightPt, $usableWidthPt, $usableHeightPt);

        // Excel's ChartEx region-map print layout uses a slightly reduced worksheet zoom and
        // reserves a larger top band than the pageMargins element alone describes. Apply this
        // to the whole sheet so cells, charts, and shapes keep their shared anchor coordinates.
        if (ChartExRegionMapPrintLayout::Applies($layout->Sheet) && $setup->Scale === null
            && $setup->FitToWidth === null && $setup->FitToHeight === null
            && !$options->OnePagePerSheet && !$options->AllColumnsInOnePagePerSheet) {
            $scale = ChartExRegionMapPrintLayout::Scale;
            $leftPt += ChartExRegionMapPrintLayout::LeftInsetPt;
            $topPt = ChartExRegionMapPrintLayout::TopOrigin($pageHeightPt, $topPt);
            $usableWidthPt += ChartExRegionMapPrintLayout::HorizontalBreakAllowancePt;
            $usableHeightPt = max(1.0, $pageHeightPt - $topPt - $bottomPt);
        }

        $colBreaks = $forceOnePageWide
            ? self::SingleSpan($firstColumn, $lastColumn)
            : self::SplitAxis(
                $firstColumn,
                $lastColumn,
                $usableWidthPt / $scale,
                // Excel places column breaks on its whole-pixel worksheet grid. Keep the
                // fractional vector geometry for drawing, but not for deciding a new page.
                function ($index) use ($layout) {
                    return floor($layout->ColumnWidthPt[$index] / RenderUnits::PixelsToPoints) * RenderUnits::PixelsToPoints;
                },
                $setup->VerticalPageBreaks);

        if (ChartExRegionMapPrintLayout::Applies($layout->Sheet) && !$forceOnePageWide) {
            self::KeepRegionMapChartsTogether($layout, $colBreaks);
        }

        $rowBreaks = $forceOnePageTall
            ? self::SingleSpan($firstRow, $lastRow)
            : self::SplitAxis(
                $firstRow,
                $lastRow,
                $usableHeightPt / $scale,
                function ($index) use ($layout) {
                    return $layout->RowHeightPt[$index];
                },
                $setup->HorizontalPageBreaks);

        $pageCount = count($rowBreaks) * count($colBreaks);
        for ($pageIndex = 0; $pageIndex < $pageCount; $pageIndex++) {
                $rowIndex = $setup->PageOrderOverThenDown ? intdiv($pageIndex, count($colBreaks)) : $pageIndex % count($rowBreaks);
                $colIndex = $setup->PageOrderOverThenDown ? $pageIndex % count($colBreaks) : intdiv($pageIndex, count($rowBreaks));
                $rowSpan = $rowBreaks[$rowIndex];
                $colSpan = $colBreaks[$colIndex];
                $page = new PageLayout();
                $page->SheetIndex = $sheetIndex;
                $page->Sheet = $layout;
                $page->StartRow = $rowSpan->Start;
                $page->EndRow = $rowSpan->End;
                $page->StartColumn = $colSpan->Start;
                $page->EndColumn = $colSpan->End;
                $page->ScaleFactor = $scale;
                $page->PageWidthPt = $pageWidthPt;
                $page->PageHeightPt = $pageHeightPt;
                $page->ContentOriginXPt = $leftPt;
                $page->ContentOriginYPt = $topPt;
                $pages[] = $page;
        }

        return $pages;
    }

    private static function ResolveScale($setup, $options, $totalWidthPt, $totalHeightPt, $usableWidthPt, $usableHeightPt)
    {
        // Save-option fit modes take precedence over the sheet's own page-setup scaling.
        if ($options->OnePagePerSheet) {
            $w = $totalWidthPt > 0.0 ? $usableWidthPt / $totalWidthPt : 1.0;
            $h = $totalHeightPt > 0.0 ? $usableHeightPt / $totalHeightPt : 1.0;
            return min(1.0, min($w, $h));
        }

        if ($options->AllColumnsInOnePagePerSheet) {
            $w = $totalWidthPt > 0.0 ? $usableWidthPt / $totalWidthPt : 1.0;
            return min(1.0, $w);
        }

        $fitWidth = $setup->FitToPage && ($setup->FitToWidth === null || $setup->FitToWidth > 0);
        $fitHeight = $setup->FitToPage && ($setup->FitToHeight === null || $setup->FitToHeight > 0);

        if ($fitWidth || $fitHeight) {
            $scale = PHP_FLOAT_MAX;
            if ($fitWidth && $totalWidthPt > 0.0) {
                $scale = min($scale, ($setup->FitToWidth ?? 1) * $usableWidthPt / $totalWidthPt);
            }

            if ($fitHeight && $totalHeightPt > 0.0) {
                $scale = min($scale, ($setup->FitToHeight ?? 1) * $usableHeightPt / $totalHeightPt);
            }

            if ($scale == PHP_FLOAT_MAX) {
                return 1.0;
            }

            // Fit-to-page only ever shrinks content; it never enlarges past 100%.
            return min(1.0, $scale);
        }

        if ($setup->Scale !== null && $setup->Scale > 0) {
            return $setup->Scale / 100.0;
        }

        return 1.0;
    }

    private static function SingleSpan($first, $last)
    {
        return array(new Span($first, $last));
    }

    private static function KeepRegionMapChartsTogether($layout, $spans)
    {
        for ($spanIndex = 0; $spanIndex + 1 < count($spans); $spanIndex++) {
            $current = $spans[$spanIndex];
            $next = $spans[$spanIndex + 1];
            if ($next->Start !== $current->End + 1 || $next->Start >= $next->End) {
                continue;
            }

            foreach ($layout->Sheet->Charts as $chart) {
                if ($chart !== null && $chart->IsChartEx
                    && $chart->UpperLeftColumn >= $current->Start
                    && $chart->UpperLeftColumn <= $current->End
                    && $chart->LowerRightColumn === $next->Start) {
                    $current->End++;
                    $next->Start++;
                    $spans[$spanIndex] = $current;
                    $spans[$spanIndex + 1] = $next;
                    break;
                }
            }
        }
    }

    // Greedily groups indices 0..last into spans that each fit within $capacity
    // (in unscaled points), honoring manual breaks that fall earlier than the greedy break.
    private static function SplitAxis($first, $last, $capacity, $size, $manualBreaks)
    {
        $breakSet = array();
        if ($manualBreaks !== null) {
            foreach ($manualBreaks as $b) {
                $breakSet[$b] = true;
            }
        }
        $spans = array();
        $start = $first;
        $accumulated = 0.0;

        for ($index = $first; $index <= $last; $index++) {
            $cellSize = $size($index);

            $manualBreakHere = $index > $start && isset($breakSet[$index]);
            $overflow = $index > $start && $accumulated + $cellSize > $capacity;

            if ($manualBreakHere || $overflow) {
                $spans[] = new Span($start, $index - 1);
                $start = $index;
                $accumulated = 0.0;
            }

            $accumulated += $cellSize;
        }

        $spans[] = new Span($start, $last);
        return $spans;
    }
}
