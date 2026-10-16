<?php

namespace Aspose\Cells_FOSS;

// Represents font.
class Font
{
    public $Name;
    public $Size;
    public $Underline;
    public $IsBold;
    public $IsItalic;
    public $IsStrikeout;
    public $Color;

    public function __construct()
    {
        $this->Name = "Calibri";
        $this->Size = 11.0;
        $this->Color = Color::Empty();
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'DoubleSize') {
            return $this->Size;
        }
        if ($propertyName === 'ArgbColor') {
            return ($this->Color->A << 24) | ($this->Color->R << 16) | ($this->Color->G << 8) | $this->Color->B;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'DoubleSize') {
            $this->Size = $value;
            return;
        }
        if ($propertyName === 'ArgbColor') {
            $a = ($value >> 24) & 0xFF;
            $r = ($value >> 16) & 0xFF;
            $g = ($value >> 8) & 0xFF;
            $b = $value & 0xFF;
            $this->Color = Color::FromArgb($a, $r, $g, $b);
            return;
        }
    }

    public function Equals($font)
    {
        if ($font === null) {
            return false;
        }

        return $this->Name == $font->Name
            && $this->Size == $font->Size
            && $this->IsBold == $font->IsBold
            && $this->IsItalic == $font->IsItalic
            && $this->Underline == $font->Underline
            && $this->IsStrikeout == $font->IsStrikeout
            && $this->Color->Equals($font->Color);
    }

    public function ToString()
    {
        return sprintf(
            "Font(Name=%s, Size=%s, Bold=%s, Italic=%s, Underline=%s, Strikeout=%s, Argb=%s)",
            $this->Name,
            $this->Size,
            $this->IsBold ? 'True' : 'False',
            $this->IsItalic ? 'True' : 'False',
            self::GetUnderlineName($this->Underline),
            $this->IsStrikeout ? 'True' : 'False',
            $this->ArgbColor
        );
    }

    private static function GetUnderlineName($underline)
    {
        switch ($underline) {
            case FontUnderlineType::None:
                return 'None';
            case FontUnderlineType::Single:
                return 'Single';
            case FontUnderlineType::Double:
                return 'Double';
            case FontUnderlineType::Accounting:
                return 'Accounting';
            case FontUnderlineType::DoubleAccounting:
                return 'DoubleAccounting';
            case FontUnderlineType::Dash:
                return 'Dash';
            case FontUnderlineType::DashDotDotHeavy:
                return 'DashDotDotHeavy';
            case FontUnderlineType::DashDotHeavy:
                return 'DashDotHeavy';
            case FontUnderlineType::DashedHeavy:
                return 'DashedHeavy';
            case FontUnderlineType::DashLong:
                return 'DashLong';
            case FontUnderlineType::DashLongHeavy:
                return 'DashLongHeavy';
            case FontUnderlineType::DotDash:
                return 'DotDash';
            case FontUnderlineType::DotDotDash:
                return 'DotDotDash';
            case FontUnderlineType::Dotted:
                return 'Dotted';
            case FontUnderlineType::DottedHeavy:
                return 'DottedHeavy';
            case FontUnderlineType::Heavy:
                return 'Heavy';
            case FontUnderlineType::Wave:
                return 'Wave';
            case FontUnderlineType::WavyDouble:
                return 'WavyDouble';
            case FontUnderlineType::WavyHeavy:
                return 'WavyHeavy';
            case FontUnderlineType::Words:
                return 'Words';
            default:
                return (string) $underline;
        }
    }

    public function Clone()
    {
        $clone = new Font();
        $clone->Name = $this->Name;
        $clone->Size = $this->Size;
        $clone->IsBold = $this->IsBold;
        $clone->IsItalic = $this->IsItalic;
        $clone->Underline = $this->Underline;
        $clone->IsStrikeout = $this->IsStrikeout;
        $clone->Color = $this->Color;
        return $clone;
    }
}
