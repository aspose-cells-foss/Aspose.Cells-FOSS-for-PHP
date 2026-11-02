<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\FontUnderlineType;
use Aspose\Cells_FOSS\Skia\SKColor;

final class CellEffect
{
    public $HasFill;
    public $Fill;
    public $HasFontColor;
    public $FontColor;
    public $HasFontBold;
    public $HasFontItalic;
    public $HasFontUnderline;
    public $FontUnderline;
    public $Bar;
    public $IconGlyph;
    public $SuppressText;   // icon-set rule with showValue="0": draw the icon only

    public function __construct()
    {
        $this->HasFill = false;
        $this->Fill = new SKColor(0, 0, 0, 0);
        $this->HasFontColor = false;
        $this->FontColor = new SKColor(0, 0, 0, 0);
        $this->HasFontBold = false;
        $this->HasFontItalic = false;
        $this->HasFontUnderline = false;
        $this->FontUnderline = FontUnderlineType::None;
        $this->Bar = null;
        $this->IconGlyph = null;
        $this->SuppressText = false;
    }
}
