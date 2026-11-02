<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\MergeRegion;
use Aspose\Cells_FOSS\XlsxWorkbookPageSetup;

final class PrintAreaRanges
{
    private function __construct()
    {
    }

    public static function Read($printArea)
    {
        $ranges = array();
        if ($printArea === null || trim($printArea) === '') {
            return $ranges;
        }
        $normalized = XlsxWorkbookPageSetup::NormalizeLoadedPrintArea($printArea);
        foreach (explode(',', $normalized) as $part) {
            $ends = explode(':', str_replace('$', '', $part));
            $first = CellAddress::Parse($ends[0]);
            $last = CellAddress::Parse($ends[count($ends) - 1]);
            $ranges[] = new MergeRegion($first->RowIndex, $first->ColumnIndex,
                $last->RowIndex - $first->RowIndex + 1, $last->ColumnIndex - $first->ColumnIndex + 1);
        }
        return $ranges;
    }
}
