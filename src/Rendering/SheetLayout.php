<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\BorderStyle;
use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\CellValueKind;
use Aspose\Cells_FOSS\Core\FillPatternKind;
use Aspose\Cells_FOSS\Core\HorizontalAlignment;
use Aspose\Cells_FOSS\Core\MergeRegion;
use Aspose\Cells_FOSS\Core\StyleValue;
use Aspose\Cells_FOSS\DisplayTextFormatter;

final class SheetLayout
{
    // Horizontal text padding inside a cell, per side, in points (~2px at 96 DPI).
    const HorizontalPaddingPt = 1.5;

    // Vertical text padding inside a cell, per side, in points.
    const VerticalPaddingPt = 1.0;
    const AutoFitSingleLineVerticalPaddingPt = 0.72;

    private const DefaultColumnWidthChars = 8.43;
    private const DefaultRowHeightPt = 15;

    public $Sheet;

    public $FirstRow;
    public $FirstColumn;
    public $LastRow;
    public $LastColumn;

    public $ColumnWidthPt;   // indexed 0..LastColumn
    public $RowHeightPt;      // indexed 0..LastRow
    public $ColumnHidden;
    public $RowHidden;

    public $ColumnStartPt;    // length LastColumn+2, cumulative left edge (0-based origin)
    public $RowStartPt;       // length LastRow+2, cumulative top edge

    private $_mergeTopLeft;
    private $_mergeCovered;

    public function __get($propertyName)
    {
        if ($propertyName === 'IsEmpty') {
            return $this->LastRow < $this->FirstRow || $this->LastColumn < $this->FirstColumn;
        }

        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'IsEmpty') {
            throw new \RuntimeException('IsEmpty is read-only.');
        }
    }

    public static function Build($context, $sheet)
    {
        $layout = new SheetLayout();
        $layout->Sheet = $sheet;
        $layout->ComputeUsedRange($sheet);
        $layout->BuildMergeLookup($sheet);
        $layout->ComputeColumns($context, $sheet);
        $layout->ComputeRows($context, $sheet);
        $layout->ComputeCumulativeOffsets();
        $layout->TrimFloatingObjectAnchorOvershoot($sheet);
        return $layout;
    }

    private function ComputeUsedRange($sheet)
    {
        $minRow = 2147483647;
        $minCol = 2147483647;
        $maxRow = -1;
        $maxCol = -1;

        foreach ($sheet->Cells as $addressKey => $pairValue) {
            $address = CellAddress::Parse($addressKey);
            if ($pairValue !== null && $pairValue->Kind === CellValueKind::Blank && $pairValue->Value === null
                && self::IsDefaultStyle($pairValue->Style)) {
                continue;
            }

            if ($address->RowIndex < $minRow) $minRow = $address->RowIndex;
            if ($address->ColumnIndex < $minCol) $minCol = $address->ColumnIndex;
            if ($address->RowIndex > $maxRow) $maxRow = $address->RowIndex;
            if ($address->ColumnIndex > $maxCol) $maxCol = $address->ColumnIndex;
        }

        foreach ($sheet->MergeRegions as $merge) {
            if ($merge->FirstRow < $minRow) $minRow = $merge->FirstRow;
            if ($merge->FirstColumn < $minCol) $minCol = $merge->FirstColumn;
            $lastR = $merge->FirstRow + max(1, $merge->TotalRows) - 1;
            $lastC = $merge->FirstColumn + max(1, $merge->TotalColumns) - 1;
            if ($lastR > $maxRow) $maxRow = $lastR;
            if ($lastC > $maxCol) $maxCol = $lastC;
        }

        // Anchored pictures extend the printable area (Excel paginates to include them), so a
        // picture below/right of the data still gets its own page range.
        foreach ($sheet->Pictures as $picture) {
            if ($picture->UpperLeftRow < $minRow) $minRow = $picture->UpperLeftRow;
            if ($picture->UpperLeftColumn < $minCol) $minCol = $picture->UpperLeftColumn;
            if ($picture->LowerRightRow > $maxRow) $maxRow = $picture->LowerRightRow;
            if ($picture->LowerRightColumn > $maxCol) $maxCol = $picture->LowerRightColumn;
        }

        // Charts extend the printable area the same way as pictures.
        foreach ($sheet->Charts as $chart) {
            if ($chart->UpperLeftRow < $minRow) $minRow = $chart->UpperLeftRow;
            if ($chart->UpperLeftColumn < $minCol) $minCol = $chart->UpperLeftColumn;
            if ($chart->LowerRightRow > $maxRow) $maxRow = $chart->LowerRightRow;
            if ($chart->LowerRightColumn > $maxCol) $maxCol = $chart->LowerRightColumn;
        }

        // Shapes (autoShapes, connectors) extend the printable area too.
        foreach ($sheet->Shapes as $shape) {
            if ($shape->UpperLeftRow < $minRow) $minRow = $shape->UpperLeftRow;
            if ($shape->UpperLeftColumn < $minCol) $minCol = $shape->UpperLeftColumn;
            if ($shape->LowerRightRow > $maxRow) $maxRow = $shape->LowerRightRow;
            if ($shape->LowerRightColumn > $maxCol) $maxCol = $shape->LowerRightColumn;
        }

        // SmartArt diagrams extend the printable area too.
        foreach ($sheet->SmartArts as $smartArt) {
            if ($smartArt->UpperLeftRow < $minRow) $minRow = $smartArt->UpperLeftRow;
            if ($smartArt->UpperLeftColumn < $minCol) $minCol = $smartArt->UpperLeftColumn;
            if ($smartArt->LowerRightRow > $maxRow) $maxRow = $smartArt->LowerRightRow;
            if ($smartArt->LowerRightColumn > $maxCol) $maxCol = $smartArt->LowerRightColumn;
        }

        foreach (PrintAreaRanges::Read($sheet->PageSetup->PrintArea) as $area) {
            $maxRow = max($maxRow, $area->FirstRow + $area->TotalRows - 1);
            $maxCol = max($maxCol, $area->FirstColumn + $area->TotalColumns - 1);
        }

        if ($maxRow < 0 || $maxCol < 0) {
            // Nothing to render: keep an empty range anchored at A1.
            $this->FirstRow = 0; $this->FirstColumn = 0; $this->LastRow = -1; $this->LastColumn = -1;
            return;
        }

        // Keep the printable coordinate system anchored at A1, matching Excel's pagination:
        // leading empty rows/columns still consume printable space even when they contain no
        // visible content.
        $this->FirstRow = 0;
        $this->FirstColumn = 0;
        $this->LastRow = $maxRow;
        $this->LastColumn = $maxCol;
    }

    private static function IsDefaultStyle($style)
    {
        return $style === null
            || ($style->Pattern === FillPatternKind::None
                && $style->Borders->Left->Style === BorderStyle::None
                && $style->Borders->Right->Style === BorderStyle::None
                && $style->Borders->Top->Style === BorderStyle::None
                && $style->Borders->Bottom->Style === BorderStyle::None);
    }

    private function BuildMergeLookup($sheet)
    {
        $this->_mergeTopLeft = array();
        $this->_mergeCovered = array();

        foreach ($sheet->MergeRegions as $merge) {
            $rows = max(1, $merge->TotalRows);
            $cols = max(1, $merge->TotalColumns);
            $this->_mergeTopLeft[self::Key($merge->FirstRow, $merge->FirstColumn)] = $merge;

            for ($r = 0; $r < $rows; $r++) {
                for ($c = 0; $c < $cols; $c++) {
                    if ($r === 0 && $c === 0) {
                        continue;
                    }

                    $this->_mergeCovered[self::Key($merge->FirstRow + $r, $merge->FirstColumn + $c)] = true;
                }
            }
        }
    }

    private function ComputeColumns($context, $sheet)
    {
        if ($this->IsEmpty) {
            $this->ColumnWidthPt = array();
            $this->ColumnHidden = array();
            return;
        }

        $count = $this->LastColumn + 1;
        $this->ColumnWidthPt = array_fill(0, $count, 0.0);
        $this->ColumnHidden = array_fill(0, $count, false);

        // An explicit defaultColWidth is a stored (padded) width; the fabricated fallback (8.43) is
        // a display-character count, so each uses the matching conversion to match Excel.
        $defaultPt = $sheet->DefaultColumnWidth !== null && $sheet->DefaultColumnWidth > 0
            ? RenderUnits::ColumnWidthCharsToPoints($sheet->DefaultColumnWidth, $context->MaxDigitWidthPixels)
            : RenderUnits::DefaultColumnWidthCharsToPoints(self::DefaultColumnWidthChars, $context->MaxDigitWidthPixels);
        $normalFont = $context->Workbook !== null && $context->Workbook->DefaultStyle !== null
            ? $context->Workbook->DefaultStyle->Font : null;
        if ($sheet->DefaultColumnWidth === null && $normalFont !== null
            && (strcasecmp((string)$normalFont->Name, 'SimSun') === 0
                || (string)$normalFont->Name === '宋体')) {
            // Excel retains a 66-pixel implicit column width when the normal style uses SimSun.
            // Applying SimSun's wider glyph advance to the unstored 8.43-character fallback
            // makes drawing anchors drift farther right with each intervening column.
            $defaultPt = RenderUnits::PixelsToPointsValue(66.0);
        }

        for ($c = 0; $c < $count; $c++) {
            $this->ColumnWidthPt[$c] = $defaultPt;
        }

        foreach ($sheet->Columns as $range) {
            $min = max(0, $range->MinColumnIndex);
            $max = min($this->LastColumn, $range->MaxColumnIndex);
            for ($c = $min; $c <= $max; $c++) {
                if ($range->Hidden) {
                    $this->ColumnHidden[$c] = true;
                    $this->ColumnWidthPt[$c] = 0.0;
                } elseif ($range->Width !== null && $range->Width > 0) {
                    $this->ColumnWidthPt[$c] = RenderUnits::ColumnWidthCharsToPoints($range->Width, $context->MaxDigitWidthPixels);
                }
            }
        }
    }

    private function ComputeRows($context, $sheet)
    {
        if ($this->IsEmpty) {
            $this->RowHeightPt = array();
            $this->RowHidden = array();
            return;
        }

        $count = $this->LastRow + 1;
        $this->RowHeightPt = array_fill(0, $count, 0.0);
        $this->RowHidden = array_fill(0, $count, false);
        $mapPrintLayout = ChartExRegionMapPrintLayout::Applies($sheet);
        $rowHeightScale = self::ResolveStoredRowHeightScale($context, $sheet);

        $customHeightSet = $sheet->CustomHeight !== null && $sheet->CustomHeight;
        $defaultHeight = 0.0;
        if (($customHeightSet || $mapPrintLayout) && $sheet->DefaultRowHeight !== null && $sheet->DefaultRowHeight > 0) {
            $defaultHeight = $sheet->DefaultRowHeight;
        } elseif ($context->FontDerivedRowHeightPt > 0.0) {
            $defaultHeight = $context->FontDerivedRowHeightPt;
        } else {
            $defaultHeight = $sheet->DefaultRowHeight !== null && $sheet->DefaultRowHeight > 0
                ? $sheet->DefaultRowHeight
                : self::DefaultRowHeightPt;
        }

        for ($r = 0; $r < $count; $r++) {
            $rowModel = array_key_exists($r, $sheet->Rows) ? $sheet->Rows[$r] : null;
            $hasRow = $rowModel !== null;

            if ($hasRow && $rowModel->Hidden) {
                $this->RowHidden[$r] = true;
                $this->RowHeightPt[$r] = 0.0;
                continue;
            }

            if ($hasRow && $rowModel->Height !== null && $rowModel->Height > 0
                && ($rowModel->CustomHeight || $mapPrintLayout || $this->RowUsesCachedAutoHeight($sheet, $r))) {
                $this->RowHeightPt[$r] = $rowModel->Height;
            } else {
                $fittedHeight = max($defaultHeight, $this->AutoRowHeight($context, $sheet, $r));
                if ($hasRow && $rowModel->Height !== null && $rowModel->Height > 0) {
                    $fittedHeight = max($fittedHeight, min($rowModel->Height * $rowHeightScale, $fittedHeight));
                }

                $this->RowHeightPt[$r] = $fittedHeight;
            }
        }
    }

    // Returns the scale for cached row heights after a default-font change.
    private static function ResolveStoredRowHeightScale($context, $sheet)
    {
        if ($sheet === null || $sheet->DefaultRowHeight === null || $sheet->DefaultRowHeight <= 0.0) {
            return 1.0;
        }

        if ($sheet->CustomHeight !== null && $sheet->CustomHeight) {
            return 1.0;
        }

        if ($context === null || $context->FontDerivedRowHeightPt <= 0.0) {
            return 1.0;
        }

        $scale = $context->FontDerivedRowHeightPt / $sheet->DefaultRowHeight;
        if ($scale < 0.75 || $scale > 1.25) {
            return 1.0;
        }

        return $scale;
    }

    // Preserves cached heights for rotated text and empty drawing spacer rows.
    private function RowUsesCachedAutoHeight($sheet, $row)
    {
        $hasContent = false;
        for ($col = 0; $col <= $this->LastColumn; $col++) {
            if ($this->ColumnHidden[$col]) {
                continue;
            }

            $key = (new CellAddress($row, $col))->ToString();
            $record = array_key_exists($key, $sheet->Cells) ? $sheet->Cells[$key] : null;
            if ($record === null) {
                continue;
            }

            if ($this->IsMergeCovered($row, $col) || $this->SpansMultipleRows($row, $col)) {
                continue;
            }

            $style = $record->Style !== null ? $record->Style : StyleValue::Default();
            if ($record->Value !== null && !(is_string($record->Value) && $record->Value === '')) {
                $hasContent = true;
            }
            if ($style->Alignment->TextRotation > 0) {
                return true;
            }
        }

        // Empty rows have no text to re-measure. Their stored height still positions drawings.
        return !$hasContent;
    }

    // Computes the auto-fit height for a row whose height is not fixed.
    private function AutoRowHeight($context, $sheet, $row)
    {
        $needed = 0.0;

        for ($col = 0; $col <= $this->LastColumn; $col++) {
            if ($this->ColumnHidden[$col]) {
                continue;
            }

            $key = (new CellAddress($row, $col))->ToString();
            $record = array_key_exists($key, $sheet->Cells) ? $sheet->Cells[$key] : null;
            if ($record === null) {
                continue;
            }

            $style = $record->Style !== null ? $record->Style : StyleValue::Default();
            $rotation = $style->Alignment->TextRotation;
            $isRotated = $rotation > 0;
            if ($this->IsMergeCovered($row, $col) || $this->SpansMultipleRows($row, $col)) {
                continue;
            }

            $text = DisplayTextFormatter::FormatDisplayValue($record->Value, $style, $context->Culture,
                $context->Workbook->Settings->DateSystem);
            if ($text === null || $text === '') {
                continue;
            }

            $font = $context->GetFontContext($style->Font);

            if ($isRotated) {
                $rotatedHeight = self::RotatedTextHeightPt($font, $text, $rotation);
                if ($rotatedHeight > $needed) {
                    $needed = $rotatedHeight;
                }
                continue;
            }

            $lineCount = 1;
            if ($style->Alignment->WrapText) {
                $availableWidth = $this->MergedWidthPt($row, $col) - 2.0 * self::HorizontalPaddingPt;
                $availableWidth = self::EffectiveWrapWidthPt($availableWidth, $style->Alignment);
                if ($availableWidth <= 0.0) {
                    continue;
                }

                $lines = $font->WrapLines($text, (float)$availableWidth, true);
                $lineCount = max(1, count($lines));
            }

            if (!($style->Alignment->WrapText) && $style->Alignment->Horizontal === HorizontalAlignment::Distributed) {
                $width = DistributedTextLine::UnwrappedWidth($font,
                    (float)($this->MergedWidthPt($row, $col) - 2.0 * self::HorizontalPaddingPt), $style->Alignment);
                $lineCount = max(1, count($font->WrapLines($text, $width, true)));
            }

            $normalFont = $context->Workbook !== null && $context->Workbook->DefaultStyle !== null
                ? $context->Workbook->DefaultStyle->Font
                : null;
            if (!($style->Alignment->WrapText) && $style->Alignment->Horizontal !== HorizontalAlignment::Distributed
                && $normalFont !== null && $style->Font !== null
                && strcasecmp((string)$style->Font->Name, (string)$normalFont->Name) === 0
                && $style->Font->Size <= $normalFont->Size) {
                continue;
            }

            $verticalPadding = $style->Alignment->WrapText
                ? 2.0 * self::VerticalPaddingPt
                : 2.0 * self::AutoFitSingleLineVerticalPaddingPt;
            $lineHeight = $style->Alignment->WrapText ? $font->WrappedLineHeightPt : $font->LineHeightPt;
            if (!($style->Alignment->WrapText) && $style->Alignment->Horizontal === HorizontalAlignment::Distributed) {
                $lineHeight = $font->WrappedLineHeightPt;
                $verticalPadding = 0.0;
            }
            $blockHeight = $lineCount * $lineHeight + $verticalPadding;
            if ($blockHeight > $needed) {
                $needed = $blockHeight;
            }
        }

        return $needed;
    }

    // Returns the unpadded height of rotated or vertically stacked text.
    private static function RotatedTextHeightPt($font, $text, $rotation)
    {
        $flat = str_replace("\r", '', str_replace("\n", '', $text));
        if ($rotation >= 255) {
            return max(1, strlen($flat)) * $font->LineHeightPt;
        }

        $ccwDeg = $rotation <= 90 ? $rotation : $rotation - 90;
        $rad = $ccwDeg * M_PI / 180.0;

        $font->MeasureInkBounds($flat, $inkWidth, $inkHeight);
        return $inkWidth * abs(sin($rad)) + $inkHeight * abs(cos($rad));
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

    // True when any non-merged-continuation cell in the row carries a text-rotation style, so a
    // cached (non-customHeight) row height must be re-fitted rather than trusted.
    private function RowHasRotatedText($sheet, $row)
    {
        for ($col = 0; $col <= $this->LastColumn; $col++) {
            if ($this->ColumnHidden[$col]) {
                continue;
            }

            $key = (new CellAddress($row, $col))->ToString();
            $record = array_key_exists($key, $sheet->Cells) ? $sheet->Cells[$key] : null;
            if ($record === null) {
                continue;
            }

            $style = $record->Style !== null ? $record->Style : StyleValue::Default();
            if ($style->Alignment->TextRotation > 0 && !$this->IsMergeCovered($row, $col) && !$this->SpansMultipleRows($row, $col)) {
                return true;
            }
        }

        return false;
    }

    private function SpansMultipleRows($row, $col)
    {
        $key = self::Key($row, $col);
        if (array_key_exists($key, $this->_mergeTopLeft)) {
            $merge = $this->_mergeTopLeft[$key];
            return max(1, $merge->TotalRows) > 1;
        }

        return false;
    }

    private function MergedWidthPt($row, $col)
    {
        $key = self::Key($row, $col);
        if (array_key_exists($key, $this->_mergeTopLeft)) {
            $merge = $this->_mergeTopLeft[$key];
            $width = 0.0;
            $cols = max(1, $merge->TotalColumns);
            for ($c = 0; $c < $cols && $col + $c <= $this->LastColumn; $c++) {
                $width += $this->ColumnWidthPt[$col + $c];
            }

            return $width;
        }

        return $col <= $this->LastColumn ? $this->ColumnWidthPt[$col] : 0.0;
    }

    private function ComputeCumulativeOffsets()
    {
        if ($this->IsEmpty) {
            $this->ColumnStartPt = array(0.0);
            $this->RowStartPt = array(0.0);
            return;
        }

        $this->ColumnStartPt = array_fill(0, $this->LastColumn + 2, 0.0);
        for ($c = 0; $c <= $this->LastColumn; $c++) {
            $this->ColumnStartPt[$c + 1] = $this->ColumnStartPt[$c] + $this->ColumnWidthPt[$c];
        }

        $this->RowStartPt = array_fill(0, $this->LastRow + 2, 0.0);
        for ($r = 0; $r <= $this->LastRow; $r++) {
            $this->RowStartPt[$r + 1] = $this->RowStartPt[$r] + $this->RowHeightPt[$r];
        }
    }

    // Some real files keep coarse two-cell anchors whose lower-right cell extends beyond the
    // actual object size, while the drawing's EMU extent carries the real width/height. Trim the
    // printable used range back to the actual floating-object bounds so pagination does not keep
    // extra blank trailing columns/rows solely because of that anchor overshoot.
    private function TrimFloatingObjectAnchorOvershoot($sheet)
    {
        if ($this->IsEmpty) {
            return;
        }

        $maxRow = -1;
        $maxCol = -1;

        foreach ($sheet->Cells as $addressKey => $pairValue) {
            $address = CellAddress::Parse($addressKey);
            if ($pairValue !== null && $pairValue->Kind === CellValueKind::Blank && $pairValue->Value === null
                && self::IsDefaultStyle($pairValue->Style)) {
                continue;
            }

            if ($address->RowIndex > $maxRow) {
                $maxRow = $address->RowIndex;
            }

            if ($address->ColumnIndex > $maxCol) {
                $maxCol = $address->ColumnIndex;
            }
        }

        foreach ($sheet->MergeRegions as $merge) {
            $lastRow = $merge->FirstRow + max(1, $merge->TotalRows) - 1;
            $lastCol = $merge->FirstColumn + max(1, $merge->TotalColumns) - 1;
            if ($lastRow > $maxRow) {
                $maxRow = $lastRow;
            }

            if ($lastCol > $maxCol) {
                $maxCol = $lastCol;
            }
        }

        foreach ($sheet->Pictures as $picture) {
            if ($picture === null) {
                continue;
            }

            if ($picture->ExtentCx > 0) {
                $right = $this->ColumnEdgePt($picture->UpperLeftColumn, $picture->UpperLeftColumnOffset) + $picture->ExtentCx / 12700.0;
                $effectiveLastColumn = $this->FindLastCoveredColumn($right);
                if ($effectiveLastColumn > $maxCol) {
                    $maxCol = $effectiveLastColumn;
                }
            }

            if ($picture->ExtentCy > 0) {
                $bottom = $this->RowEdgePt($picture->UpperLeftRow, $picture->UpperLeftRowOffset) + $picture->ExtentCy / 12700.0;
                $effectiveLastRow = $this->FindLastCoveredRow($bottom);
                if ($effectiveLastRow > $maxRow) {
                    $maxRow = $effectiveLastRow;
                }
            }
        }

        foreach ($sheet->Charts as $chart) {
            if ($chart === null) {
                continue;
            }

            if ($chart->LowerRightRow > $maxRow) {
                $maxRow = $chart->LowerRightRow;
            }

            if ($chart->LowerRightColumn > $maxCol) {
                $maxCol = $chart->LowerRightColumn;
            }
        }

        foreach ($sheet->Shapes as $shape) {
            if ($shape === null) {
                continue;
            }

            if ($shape->LowerRightRow > $maxRow) {
                $maxRow = $shape->LowerRightRow;
            }

            if ($shape->LowerRightColumn > $maxCol) {
                $maxCol = $shape->LowerRightColumn;
            }
        }

        foreach ($sheet->SmartArts as $smartArt) {
            if ($smartArt === null) {
                continue;
            }

            if ($smartArt->LowerRightRow > $maxRow) {
                $maxRow = $smartArt->LowerRightRow;
            }

            if ($smartArt->LowerRightColumn > $maxCol) {
                $maxCol = $smartArt->LowerRightColumn;
            }
        }

        if ($maxRow < 0 || $maxCol < 0) {
            return;
        }

        $this->LastRow = $maxRow;
        $this->LastColumn = $maxCol;
    }

    private function ColumnEdgePt($column, $emuOffset)
    {
        if ($column < 0) {
            $column = 0;
        }

        if ($column > count($this->ColumnStartPt) - 1) {
            $column = count($this->ColumnStartPt) - 1;
        }

        return $this->ColumnStartPt[$column] + $emuOffset / 12700.0;
    }

    private function RowEdgePt($row, $emuOffset)
    {
        if ($row < 0) {
            $row = 0;
        }

        if ($row > count($this->RowStartPt) - 1) {
            $row = count($this->RowStartPt) - 1;
        }

        return $this->RowStartPt[$row] + $emuOffset / 12700.0;
    }

    private function FindLastCoveredColumn($rightPt)
    {
        for ($column = 0; $column <= $this->LastColumn; $column++) {
            if ($this->ColumnStartPt[$column + 1] >= $rightPt - 0.01) {
                return $column;
            }
        }

        return $this->LastColumn;
    }

    private function FindLastCoveredRow($bottomPt)
    {
        for ($row = 0; $row <= $this->LastRow; $row++) {
            if ($this->RowStartPt[$row + 1] >= $bottomPt - 0.01) {
                return $row;
            }
        }

        return $this->LastRow;
    }

    public function TryGetPictureColumnSpan($picture, &$startColumn, &$endColumn, &$widthPt)
    {
        $startColumn = 0;
        $endColumn = 0;
        $widthPt = 0.0;

        if ($picture === null || $this->IsEmpty) {
            return false;
        }

        $startColumn = max($this->FirstColumn, $picture->UpperLeftColumn);
        if ($startColumn > $this->LastColumn) {
            $startColumn = $this->LastColumn;
        }

        $left = $this->ColumnEdgePt($picture->UpperLeftColumn, $picture->UpperLeftColumnOffset);
        $right = 0.0;
        if ($picture->ExtentCx > 0) {
            $right = $left + $picture->ExtentCx / 12700.0;
        } else {
            $right = $this->ColumnEdgePt($picture->LowerRightColumn, $picture->LowerRightColumnOffset);
        }

        if ($right <= $left + 0.01) {
            return false;
        }

        $endColumn = $this->FindLastCoveredColumn($right);
        if ($endColumn < $startColumn) {
            $endColumn = $startColumn;
        }

        $widthPt = $right - $left;
        return true;
    }

    public function IsMergeCovered($row, $col)
    {
        return isset($this->_mergeCovered[self::Key($row, $col)]);
    }

    public function TryGetMergeOrigin($row, $col, &$merge)
    {
        $key = self::Key($row, $col);
        if (array_key_exists($key, $this->_mergeTopLeft)) {
            $merge = $this->_mergeTopLeft[$key];
            return true;
        }

        return false;
    }

    private static function Key($row, $col)
    {
        return ($row << 20) | $col;
    }
}
