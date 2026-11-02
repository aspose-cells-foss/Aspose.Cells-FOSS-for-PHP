<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\CultureInfo;

final class RenderContext
{
    private $fontContexts = array();

    public $Workbook;

    public $Fonts;

    public $Pictures;

    public $Charts;

    public $Colors;

    public $Culture;

    public $ConditionalFormattingEvaluationDate;

    public $MaxDigitWidthPixels;

    public $FontDerivedRowHeightPt;

    public $PdfTextSession = null;

    public $EnableWorksheetTextOptimization = false;

    public function __construct($workbook, $defaultFont = null)
    {
        $this->Workbook = $workbook;
        $this->Fonts = new FontRegistry();
        $this->Pictures = new PicturePdfImageCache();
        $this->Charts = new ChartPdfImageCache();
        $this->Fonts->DefaultFontName = $defaultFont;
        $this->Colors = RenderColor::FromWorkbook($workbook);
        $displayCulture = $workbook !== null && $workbook->Settings !== null ? $workbook->Settings->DisplayCulture : null;
        $this->Culture = $displayCulture !== null && !self::equalsInvariant($displayCulture)
            ? $displayCulture
            : CultureInfo::GetInvariantCulture();
        $this->ConditionalFormattingEvaluationDate = new \DateTime('today');
        $normalFont = $workbook !== null && $workbook->DefaultStyle !== null ? $workbook->DefaultStyle->Font : null;
        $this->MaxDigitWidthPixels = $this->Fonts->MeasureMaxDigitWidth($normalFont);
        $this->FontDerivedRowHeightPt = $this->Fonts->MeasureDefaultRowHeightPt($normalFont);
    }

    private static function equalsInvariant($culture)
    {
        $invariant = CultureInfo::GetInvariantCulture();
        return $culture === $invariant || ($culture->Name === $invariant->Name);
    }

    public function GetFontContext($font)
    {
        $name = $font !== null && $font->Name !== null ? $font->Name : 'Calibri';
        $size = $font !== null ? $font->Size : 11.0;
        $bold = $font !== null && $font->Bold;
        $italic = $font !== null && $font->Italic;
        $key = $name . '|' . self::roundTrip($size) . '|' . ($bold ? '1' : '0') . '|' . ($italic ? '1' : '0');

        if (array_key_exists($key, $this->fontContexts)) {
            return $this->fontContexts[$key];
        }

        $context = new CellFontContext($this->Fonts, $font);
        $this->fontContexts[$key] = $context;
        return $context;
    }

    private static function roundTrip($value)
    {
        return (string)(float)$value;
    }

    public function Dispose()
    {
        foreach ($this->fontContexts as $context) {
            $context->Dispose();
        }

        $this->fontContexts = array();
        $this->Charts->Dispose();
        $this->Pictures->Dispose();
        $this->Fonts->Dispose();
    }
}
