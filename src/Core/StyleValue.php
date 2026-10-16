<?php
namespace Aspose\Cells_FOSS\Core;

final class StyleValue
{
    public function __construct()
    {
        $this->Font = new FontValue();
        $this->Pattern = FillPatternKind::None;
        $this->ForegroundColor = new ColorValue();
        $this->BackgroundColor = new ColorValue();
        $this->Borders = new BordersValue();
        $this->Alignment = new AlignmentValue();
        $this->Protection = new ProtectionValue();
        $this->NumberFormat = new NumberFormatValue();
        $this->QuotePrefix = false;
    }

    public static function Default()
    {
        return new StyleValue();
    }

    public $Font;
    public $Pattern;
    public $ForegroundColor;
    public $BackgroundColor;
    public $Borders;
    public $Alignment;
    public $Protection;
    public $NumberFormat;
    public $QuotePrefix;

    public function Clone()
    {
        $clone = new StyleValue();
        $clone->Font = $this->Font->Clone();
        $clone->Pattern = $this->Pattern;
        $clone->ForegroundColor = clone $this->ForegroundColor;
        $clone->BackgroundColor = clone $this->BackgroundColor;
        $clone->Borders = $this->Borders->Clone();
        $clone->Alignment = $this->Alignment->Clone();
        $clone->Protection = $this->Protection->Clone();
        $clone->NumberFormat = $this->NumberFormat->Clone();
        $clone->QuotePrefix = $this->QuotePrefix;
        return $clone;
    }
}
