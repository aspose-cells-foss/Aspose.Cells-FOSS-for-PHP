<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKRect;

final class PageContentClip
{
    private function __construct()
    {
    }

    public static function Resolve($page, $floatingObjects)
    {
        $layout = $page->Sheet;
        $width = (float)($layout->ColumnStartPt[$page->EndColumn + 1]
            - $layout->ColumnStartPt[$page->StartColumn]);
        $height = (float)($layout->RowStartPt[$page->EndRow + 1]
            - $layout->RowStartPt[$page->StartRow]);
        if ($floatingObjects) {
            if ($page->EndColumn >= $layout->LastColumn) {
                $width += 72.0;
            }
            if ($page->EndRow >= $layout->LastRow) {
                $height += 72.0;
            }
        }
        return new SKRect(0.0, 0.0, $width, $height);
    }

    public static function Intersect($page, $clip, $floatingObjects)
    {
        if ($page->Sheet === null) {
            return $clip;
        }
        return SKRect::Intersect($clip, self::Resolve($page, $floatingObjects));
    }
}
