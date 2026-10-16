<?php
namespace Aspose\Cells_FOSS\Core;

final class StyleRepository
{
    public function Normalize($style)
    {
        return $style->Clone();
    }
}
