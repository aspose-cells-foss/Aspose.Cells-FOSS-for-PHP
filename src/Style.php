<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\AlignmentValue;
use Aspose\Cells_FOSS\Core\BorderSideValue;
use Aspose\Cells_FOSS\Core\BordersValue;
use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\Core\NumberFormatValue;
use Aspose\Cells_FOSS\Core\ProtectionValue;
use Aspose\Cells_FOSS\Core\StyleValue;

// Represents a mutable cell style facade that can be applied to one or more cells.
class Style
{
    private $_indentLevel;
    private $_textRotation;
    private $_readingOrder;

    public $Font;
    public $Borders;
    public $Pattern;
    public $ForegroundColor;
    public $BackgroundColor;
    public $Number;
    public $Custom;
    public $HorizontalAlignment;
    public $VerticalAlignment;
    public $WrapText;
    public $ShrinkToFit;
    public $RelativeIndent;
    public $IsLocked;
    public $IsHidden;
    public $QuotePrefix;

    public function __construct()
    {
        $this->Font = new Font();
        $this->Borders = new Borders();
        $this->ForegroundColor = Color::Empty();
        $this->BackgroundColor = Color::Empty();
        $this->VerticalAlignment = VerticalAlignmentType::Bottom;
        $this->IsLocked = true;
    }

    public function Copy($source)
    {
        if ($source === null) throw new \InvalidArgumentException('source');

        $this->Font = $source->Font->Clone();
        $this->Borders = $source->Borders->Clone();
        $this->Pattern = $source->Pattern;
        $this->ForegroundColor = $source->ForegroundColor;
        $this->BackgroundColor = $source->BackgroundColor;
        $this->Number = $source->Number;
        $this->Custom = $source->Custom;
        $this->HorizontalAlignment = $source->HorizontalAlignment;
        $this->VerticalAlignment = $source->VerticalAlignment;
        $this->WrapText = $source->WrapText;
        $this->IndentLevel = $source->IndentLevel;
        $this->TextRotation = $source->TextRotation;
        $this->ShrinkToFit = $source->ShrinkToFit;
        $this->ReadingOrder = $source->ReadingOrder;
        $this->RelativeIndent = $source->RelativeIndent;
        $this->IsLocked = $source->IsLocked;
        $this->IsHidden = $source->IsHidden;
        $this->QuotePrefix = $source->QuotePrefix;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'NumberFormat') {
            return NumberFormat::ResolveFormatCode($this->Number, $this->Custom);
        }
        if ($propertyName === 'IndentLevel') {
            return $this->_indentLevel;
        }
        if ($propertyName === 'TextRotation') {
            return $this->_textRotation;
        }
        if ($propertyName === 'ReadingOrder') {
            return $this->_readingOrder;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'NumberFormat') {
            $builtInId = NumberFormat::GetBuiltInFormatId($value);
            if ($builtInId !== null) {
                $this->Number = $builtInId;
                $this->Custom = null;
                return;
            }

            $this->Number = 0;
            $this->Custom = ($value === null || trim($value) === '') ? null : trim($value);
            return;
        }
        if ($propertyName === 'IndentLevel') {
            if ($value < 0 || $value > 250) {
                throw new CellsException("IndentLevel must be between 0 and 250.");
            }

            $this->_indentLevel = $value;
            return;
        }
        if ($propertyName === 'TextRotation') {
            if ($value != 255 && ($value < 0 || $value > 180)) {
                throw new CellsException("TextRotation must be between 0 and 180, or 255 for vertical text.");
            }

            $this->_textRotation = $value;
            return;
        }
        if ($propertyName === 'ReadingOrder') {
            if ($value < 0 || $value > 2) {
                throw new CellsException("ReadingOrder must be 0, 1, or 2.");
            }

            $this->_readingOrder = $value;
            return;
        }
    }

    public function Clone()
    {
        $clone = new Style();
        $clone->Font = $this->Font->Clone();
        $clone->Borders = $this->Borders->Clone();
        $clone->Pattern = $this->Pattern;
        $clone->ForegroundColor = $this->ForegroundColor;
        $clone->BackgroundColor = $this->BackgroundColor;
        $clone->Number = $this->Number;
        $clone->Custom = $this->Custom;
        $clone->HorizontalAlignment = $this->HorizontalAlignment;
        $clone->VerticalAlignment = $this->VerticalAlignment;
        $clone->WrapText = $this->WrapText;
        $clone->IndentLevel = $this->IndentLevel;
        $clone->TextRotation = $this->TextRotation;
        $clone->ShrinkToFit = $this->ShrinkToFit;
        $clone->ReadingOrder = $this->ReadingOrder;
        $clone->RelativeIndent = $this->RelativeIndent;
        $clone->IsLocked = $this->IsLocked;
        $clone->IsHidden = $this->IsHidden;
        $clone->QuotePrefix = $this->QuotePrefix;
        return $clone;
    }

    public function Equals($obj)
    {
        if (!($obj instanceof Style)) {
            return false;
        }

        $other = $obj;
        return $this->Font->Equals($other->Font)
            && self::BordersEqual($this->Borders, $other->Borders)
            && $this->Pattern == $other->Pattern
            && $this->ForegroundColor->Equals($other->ForegroundColor)
            && $this->BackgroundColor->Equals($other->BackgroundColor)
            && $this->Number == $other->Number
            && $this->Custom === $other->Custom
            && $this->HorizontalAlignment == $other->HorizontalAlignment
            && $this->VerticalAlignment == $other->VerticalAlignment
            && $this->WrapText == $other->WrapText
            && $this->IndentLevel == $other->IndentLevel
            && $this->TextRotation == $other->TextRotation
            && $this->ShrinkToFit == $other->ShrinkToFit
            && $this->ReadingOrder == $other->ReadingOrder
            && $this->RelativeIndent == $other->RelativeIndent
            && $this->IsLocked == $other->IsLocked
            && $this->IsHidden == $other->IsHidden
            && $this->QuotePrefix == $other->QuotePrefix;
    }

    public function GetHashCode()
    {
        $hash = 17;
        $hash = $hash * 31 + spl_object_id($this->Font);
        $hash = $hash * 31 + self::GetBordersHashCode($this->Borders);
        $hash = $hash * 31 + $this->Pattern;
        $hash = $hash * 31 + $this->ForegroundColor->GetHashCode();
        $hash = $hash * 31 + $this->BackgroundColor->GetHashCode();
        $hash = $hash * 31 + $this->Number;
        $hash = $hash * 31 + ($this->Custom === null ? 0 : crc32($this->Custom));
        $hash = $hash * 31 + $this->HorizontalAlignment;
        $hash = $hash * 31 + $this->VerticalAlignment;
        $hash = $hash * 31 + ($this->WrapText ? 1 : 0);
        $hash = $hash * 31 + $this->IndentLevel;
        $hash = $hash * 31 + $this->TextRotation;
        $hash = $hash * 31 + ($this->ShrinkToFit ? 1 : 0);
        $hash = $hash * 31 + $this->ReadingOrder;
        $hash = $hash * 31 + $this->RelativeIndent;
        $hash = $hash * 31 + ($this->IsLocked ? 1 : 0);
        $hash = $hash * 31 + ($this->IsHidden ? 1 : 0);
        $hash = $hash * 31 + ($this->QuotePrefix ? 1 : 0);
        return $hash;
    }

    public function ToCore()
    {
        $font = new FontValue();
        $font->Name = $this->Font->Name;
        $font->Size = $this->Font->Size;
        $font->Bold = $this->Font->IsBold;
        $font->Italic = $this->Font->IsItalic;
        $font->Underline = $this->Font->Underline;
        $font->StrikeThrough = $this->Font->IsStrikeout;
        $font->Color = $this->Font->Color->ToCore();

        $borders = new BordersValue();
        $borders->Left = self::ToCoreBorder($this->Borders->Left);
        $borders->Right = self::ToCoreBorder($this->Borders->Right);
        $borders->Top = self::ToCoreBorder($this->Borders->Top);
        $borders->Bottom = self::ToCoreBorder($this->Borders->Bottom);
        $borders->Diagonal = self::ToCoreBorder($this->Borders->Diagonal);
        $borders->DiagonalUp = $this->Borders->DiagonalUp;
        $borders->DiagonalDown = $this->Borders->DiagonalDown;

        $alignment = new AlignmentValue();
        $alignment->Horizontal = $this->HorizontalAlignment;
        $alignment->Vertical = $this->VerticalAlignment;
        $alignment->WrapText = $this->WrapText;
        $alignment->IndentLevel = $this->IndentLevel;
        $alignment->TextRotation = $this->TextRotation;
        $alignment->ShrinkToFit = $this->ShrinkToFit;
        $alignment->ReadingOrder = $this->ReadingOrder;
        $alignment->RelativeIndent = $this->RelativeIndent;

        $protection = new ProtectionValue();
        $protection->IsLocked = $this->IsLocked;
        $protection->IsHidden = $this->IsHidden;

        $numberFormat = new NumberFormatValue();
        $numberFormat->Number = $this->Number;
        $numberFormat->Custom = $this->Custom;

        $value = new StyleValue();
        $value->Font = $font;
        $value->Pattern = $this->Pattern;
        $value->ForegroundColor = $this->ForegroundColor->ToCore();
        $value->BackgroundColor = $this->BackgroundColor->ToCore();
        $value->Borders = $borders;
        $value->Alignment = $alignment;
        $value->Protection = $protection;
        $value->NumberFormat = $numberFormat;
        $value->QuotePrefix = $this->QuotePrefix;
        return $value;
    }

    public static function FromCore($value)
    {
        if ($value === null) {
            return new Style();
        }

        $font = new Font();
        $font->Name = $value->Font->Name;
        $font->Size = $value->Font->Size;
        $font->IsBold = $value->Font->Bold;
        $font->IsItalic = $value->Font->Italic;
        $font->Underline = $value->Font->Underline;
        $font->IsStrikeout = $value->Font->StrikeThrough;
        $font->Color = Color::FromCore($value->Font->Color);

        $borders = new Borders();
        $borders->Left = self::FromCoreBorder($value->Borders->Left);
        $borders->Right = self::FromCoreBorder($value->Borders->Right);
        $borders->Top = self::FromCoreBorder($value->Borders->Top);
        $borders->Bottom = self::FromCoreBorder($value->Borders->Bottom);
        $borders->Diagonal = self::FromCoreBorder($value->Borders->Diagonal);
        $borders->DiagonalUp = $value->Borders->DiagonalUp;
        $borders->DiagonalDown = $value->Borders->DiagonalDown;

        $style = new Style();
        $style->Font = $font;
        $style->Pattern = $value->Pattern;
        $style->ForegroundColor = Color::FromCore($value->ForegroundColor);
        $style->BackgroundColor = Color::FromCore($value->BackgroundColor);
        $style->Borders = $borders;
        $style->HorizontalAlignment = $value->Alignment->Horizontal;
        $style->VerticalAlignment = $value->Alignment->Vertical;
        $style->WrapText = $value->Alignment->WrapText;
        $style->IndentLevel = $value->Alignment->IndentLevel;
        $style->TextRotation = $value->Alignment->TextRotation;
        $style->ShrinkToFit = $value->Alignment->ShrinkToFit;
        $style->ReadingOrder = $value->Alignment->ReadingOrder;
        $style->RelativeIndent = $value->Alignment->RelativeIndent;
        $style->IsLocked = $value->Protection->IsLocked;
        $style->IsHidden = $value->Protection->IsHidden;
        $style->Number = $value->NumberFormat->Number;
        $style->Custom = $value->NumberFormat->Custom;
        $style->QuotePrefix = $value->QuotePrefix;
        return $style;
    }

    private static function ToCoreBorder($border)
    {
        $value = new BorderSideValue();
        $value->Style = $border->LineStyle;
        $value->Color = $border->Color->ToCore();
        return $value;
    }

    private static function BordersEqual($left, $right)
    {
        if ($left === null || $right === null) {
            return $left === $right;
        }

        return self::BorderEqual($left->Left, $right->Left)
            && self::BorderEqual($left->Right, $right->Right)
            && self::BorderEqual($left->Top, $right->Top)
            && self::BorderEqual($left->Bottom, $right->Bottom)
            && self::BorderEqual($left->Diagonal, $right->Diagonal)
            && $left->DiagonalUp == $right->DiagonalUp
            && $left->DiagonalDown == $right->DiagonalDown;
    }

    private static function BorderEqual($left, $right)
    {
        if ($left === null || $right === null) {
            return $left === $right;
        }

        return $left->LineStyle == $right->LineStyle
            && $left->Color->Equals($right->Color);
    }

    private static function GetBordersHashCode($borders)
    {
        if ($borders === null) {
            return 0;
        }

        $hash = 17;
        $hash = $hash * 31 + self::GetBorderHashCode($borders->Left);
        $hash = $hash * 31 + self::GetBorderHashCode($borders->Right);
        $hash = $hash * 31 + self::GetBorderHashCode($borders->Top);
        $hash = $hash * 31 + self::GetBorderHashCode($borders->Bottom);
        $hash = $hash * 31 + self::GetBorderHashCode($borders->Diagonal);
        $hash = $hash * 31 + ($borders->DiagonalUp ? 1 : 0);
        $hash = $hash * 31 + ($borders->DiagonalDown ? 1 : 0);
        return $hash;
    }

    private static function GetBorderHashCode($border)
    {
        if ($border === null) {
            return 0;
        }

        $hash = 17;
        $hash = $hash * 31 + $border->LineStyle;
        $hash = $hash * 31 + $border->Color->GetHashCode();
        return $hash;
    }

    private static function FromCoreBorder($value)
    {
        $border = new Border();
        if ($value === null) {
            $border->Color = Color::Empty();
            return $border;
        }

        $border->LineStyle = $value->Style;
        $border->Color = Color::FromCore($value->Color);
        return $border;
    }
}
