<?php

namespace Aspose\Cells_FOSS\CompareOpenXml;

final class StyleSnapshot
{
    private static $default;

    public function __construct(
        $sheetName,
        $cellReference,
        $fontName,
        $fontSize,
        $fontBold,
        $fontItalic,
        $fontUnderline,
        $fontStrikeThrough,
        $fontColor,
        $fillPattern,
        $fillForegroundColor,
        $fillBackgroundColor,
        $leftBorderStyle,
        $leftBorderColor,
        $rightBorderStyle,
        $rightBorderColor,
        $topBorderStyle,
        $topBorderColor,
        $bottomBorderStyle,
        $bottomBorderColor,
        $diagonalBorderStyle,
        $diagonalBorderColor,
        $diagonalUp,
        $diagonalDown,
        $horizontalAlignment,
        $verticalAlignment,
        $wrapText,
        $indentLevel,
        $textRotation,
        $shrinkToFit,
        $readingOrder,
        $relativeIndent,
        $numberFormatId,
        $numberFormatCode,
        $isLocked,
        $isHidden
    ) {
        $this->SheetName = $sheetName;
        $this->CellReference = $cellReference;
        $this->FontName = $fontName;
        $this->FontSize = $fontSize;
        $this->FontBold = $fontBold;
        $this->FontItalic = $fontItalic;
        $this->FontUnderline = $fontUnderline;
        $this->FontStrikeThrough = $fontStrikeThrough;
        $this->FontColor = $fontColor;
        $this->FillPattern = $fillPattern;
        $this->FillForegroundColor = $fillForegroundColor;
        $this->FillBackgroundColor = $fillBackgroundColor;
        $this->LeftBorderStyle = $leftBorderStyle;
        $this->LeftBorderColor = $leftBorderColor;
        $this->RightBorderStyle = $rightBorderStyle;
        $this->RightBorderColor = $rightBorderColor;
        $this->TopBorderStyle = $topBorderStyle;
        $this->TopBorderColor = $topBorderColor;
        $this->BottomBorderStyle = $bottomBorderStyle;
        $this->BottomBorderColor = $bottomBorderColor;
        $this->DiagonalBorderStyle = $diagonalBorderStyle;
        $this->DiagonalBorderColor = $diagonalBorderColor;
        $this->DiagonalUp = $diagonalUp;
        $this->DiagonalDown = $diagonalDown;
        $this->HorizontalAlignment = $horizontalAlignment;
        $this->VerticalAlignment = $verticalAlignment;
        $this->WrapText = $wrapText;
        $this->IndentLevel = $indentLevel;
        $this->TextRotation = $textRotation;
        $this->ShrinkToFit = $shrinkToFit;
        $this->ReadingOrder = $readingOrder;
        $this->RelativeIndent = $relativeIndent;
        $this->NumberFormatId = $numberFormatId;
        $this->NumberFormatCode = $numberFormatCode;
        $this->IsLocked = $isLocked;
        $this->IsHidden = $isHidden;
    }

    public static function Default()
    {
        if (self::$default === null) {
            self::$default = new StyleSnapshot(
                '',
                '',
                'Calibri',
                ComparisonValueHelpers::NormalizeDouble(11.0),
                false,
                false,
                false,
                false,
                ComparisonValueHelpers::DefaultColorHex,
                'None',
                ComparisonValueHelpers::DefaultColorHex,
                ComparisonValueHelpers::DefaultColorHex,
                'None',
                ComparisonValueHelpers::DefaultColorHex,
                'None',
                ComparisonValueHelpers::DefaultColorHex,
                'None',
                ComparisonValueHelpers::DefaultColorHex,
                'None',
                ComparisonValueHelpers::DefaultColorHex,
                'None',
                ComparisonValueHelpers::DefaultColorHex,
                false,
                false,
                'General',
                'Bottom',
                false,
                0,
                0,
                false,
                0,
                0,
                0,
                '',
                true,
                false
            );
        }

        return self::$default;
    }

    public function CopyWithLocation($sheetName, $cellReference)
    {
        return new StyleSnapshot(
            $sheetName,
            $cellReference,
            $this->FontName,
            $this->FontSize,
            $this->FontBold,
            $this->FontItalic,
            $this->FontUnderline,
            $this->FontStrikeThrough,
            $this->FontColor,
            $this->FillPattern,
            $this->FillForegroundColor,
            $this->FillBackgroundColor,
            $this->LeftBorderStyle,
            $this->LeftBorderColor,
            $this->RightBorderStyle,
            $this->RightBorderColor,
            $this->TopBorderStyle,
            $this->TopBorderColor,
            $this->BottomBorderStyle,
            $this->BottomBorderColor,
            $this->DiagonalBorderStyle,
            $this->DiagonalBorderColor,
            $this->DiagonalUp,
            $this->DiagonalDown,
            $this->HorizontalAlignment,
            $this->VerticalAlignment,
            $this->WrapText,
            $this->IndentLevel,
            $this->TextRotation,
            $this->ShrinkToFit,
            $this->ReadingOrder,
            $this->RelativeIndent,
            $this->NumberFormatId,
            $this->NumberFormatCode,
            $this->IsLocked,
            $this->IsHidden
        );
    }

    public $SheetName;
    public $CellReference;
    public $FontName;
    public $FontSize;
    public $FontBold;
    public $FontItalic;
    public $FontUnderline;
    public $FontStrikeThrough;
    public $FontColor;
    public $FillPattern;
    public $FillForegroundColor;
    public $FillBackgroundColor;
    public $LeftBorderStyle;
    public $LeftBorderColor;
    public $RightBorderStyle;
    public $RightBorderColor;
    public $TopBorderStyle;
    public $TopBorderColor;
    public $BottomBorderStyle;
    public $BottomBorderColor;
    public $DiagonalBorderStyle;
    public $DiagonalBorderColor;
    public $DiagonalUp;
    public $DiagonalDown;
    public $HorizontalAlignment;
    public $VerticalAlignment;
    public $WrapText;
    public $IndentLevel;
    public $TextRotation;
    public $ShrinkToFit;
    public $ReadingOrder;
    public $RelativeIndent;
    public $NumberFormatId;
    public $NumberFormatCode;
    public $IsLocked;
    public $IsHidden;

    public function Equals($other)
    {
        if ($other === null) {
            return false;
        }

        return $this->SheetName === $other->SheetName
            && $this->CellReference === $other->CellReference
            && $this->FontName === $other->FontName
            && $this->FontSize === $other->FontSize
            && $this->FontBold === $other->FontBold
            && $this->FontItalic === $other->FontItalic
            && $this->FontUnderline === $other->FontUnderline
            && $this->FontStrikeThrough === $other->FontStrikeThrough
            && $this->FontColor === $other->FontColor
            && $this->FillPattern === $other->FillPattern
            && $this->FillForegroundColor === $other->FillForegroundColor
            && $this->FillBackgroundColor === $other->FillBackgroundColor
            && $this->LeftBorderStyle === $other->LeftBorderStyle
            && $this->LeftBorderColor === $other->LeftBorderColor
            && $this->RightBorderStyle === $other->RightBorderStyle
            && $this->RightBorderColor === $other->RightBorderColor
            && $this->TopBorderStyle === $other->TopBorderStyle
            && $this->TopBorderColor === $other->TopBorderColor
            && $this->BottomBorderStyle === $other->BottomBorderStyle
            && $this->BottomBorderColor === $other->BottomBorderColor
            && $this->DiagonalBorderStyle === $other->DiagonalBorderStyle
            && $this->DiagonalBorderColor === $other->DiagonalBorderColor
            && $this->DiagonalUp === $other->DiagonalUp
            && $this->DiagonalDown === $other->DiagonalDown
            && $this->HorizontalAlignment === $other->HorizontalAlignment
            && $this->VerticalAlignment === $other->VerticalAlignment
            && $this->WrapText === $other->WrapText
            && $this->IndentLevel === $other->IndentLevel
            && $this->TextRotation === $other->TextRotation
            && $this->ShrinkToFit === $other->ShrinkToFit
            && $this->ReadingOrder === $other->ReadingOrder
            && $this->RelativeIndent === $other->RelativeIndent
            && $this->NumberFormatId === $other->NumberFormatId
            && $this->NumberFormatCode === $other->NumberFormatCode
            && $this->IsLocked === $other->IsLocked
            && $this->IsHidden === $other->IsHidden;
    }

    public function ToString()
    {
        return 'Sheet=' . $this->SheetName
            . '; Cell=' . $this->CellReference
            . '; Font=' . $this->FontName . '/' . $this->FontSize . '/B=' . self::BoolText($this->FontBold)
            . '/I=' . self::BoolText($this->FontItalic) . '/U=' . self::BoolText($this->FontUnderline)
            . '/S=' . self::BoolText($this->FontStrikeThrough) . '/Color=' . $this->FontColor
            . '; Borders=L(' . $this->LeftBorderStyle . ',' . $this->LeftBorderColor . ')'
            . '/R(' . $this->RightBorderStyle . ',' . $this->RightBorderColor . ')'
            . '/T(' . $this->TopBorderStyle . ',' . $this->TopBorderColor . ')'
            . '/B(' . $this->BottomBorderStyle . ',' . $this->BottomBorderColor . ')'
            . '/D(' . $this->DiagonalBorderStyle . ',' . $this->DiagonalBorderColor . ',Up=' . self::BoolText($this->DiagonalUp)
            . ',Down=' . self::BoolText($this->DiagonalDown) . ')'
            . '; Align=H:' . $this->HorizontalAlignment . '/V:' . $this->VerticalAlignment
            . '/Wrap=' . self::BoolText($this->WrapText) . '/Indent=' . $this->IndentLevel
            . '/Rotation=' . $this->TextRotation . '/Shrink=' . self::BoolText($this->ShrinkToFit)
            . '/Reading=' . $this->ReadingOrder . '/RelativeIndent=' . $this->RelativeIndent
            . '; Fill=' . $this->FillPattern . '/' . $this->FillForegroundColor . '/' . $this->FillBackgroundColor
            . '; Number=' . $this->NumberFormatId . '/' . $this->NumberFormatCode
            . '; Protection=Locked:' . self::BoolText($this->IsLocked) . '/Hidden:' . self::BoolText($this->IsHidden);
    }

    public function __toString()
    {
        return $this->ToString();
    }

    private static function BoolText($value)
    {
        return $value ? 'True' : 'False';
    }
}
