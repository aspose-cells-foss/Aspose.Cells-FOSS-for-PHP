<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\BorderSideValue;
use Aspose\Cells_FOSS\Core\BorderStyle;
use Aspose\Cells_FOSS\Core\BordersValue;
use Aspose\Cells_FOSS\Core\ColorValue;
use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\Core\HorizontalAlignment;
use Aspose\Cells_FOSS\Core\VerticalAlignment;

final class XlsxWorkbookStylesValueHelpers
{
    private function __construct()
    {
    }

    public static function ParseBorderStyle($value)
    {
        if ($value === null || $value === '') {
            return BorderStyle::None;
        }
        $v = strtolower($value);
        switch ($v) {
            case 'thin':
                return BorderStyle::Thin;
            case 'medium':
                return BorderStyle::Medium;
            case 'thick':
                return BorderStyle::Thick;
            case 'dotted':
                return BorderStyle::Dotted;
            case 'dashed':
                return BorderStyle::Dashed;
            case 'double':
                return BorderStyle::Double;
            case 'hair':
                return BorderStyle::Hair;
            case 'mediumdashed':
                return BorderStyle::MediumDashed;
            case 'dashdot':
                return BorderStyle::DashDot;
            case 'mediumdashdot':
                return BorderStyle::MediumDashDot;
            case 'dashdotdot':
                return BorderStyle::DashDotDot;
            case 'mediumdashdotdot':
                return BorderStyle::MediumDashDotDot;
            case 'slantdashdot':
                return BorderStyle::SlantedDashDot;
            default:
                return BorderStyle::None;
        }
    }

    public static function GetBorderStyleName($value)
    {
        switch ($value) {
            case BorderStyle::Thin:
                return 'thin';
            case BorderStyle::Medium:
                return 'medium';
            case BorderStyle::Thick:
                return 'thick';
            case BorderStyle::Dotted:
                return 'dotted';
            case BorderStyle::Dashed:
                return 'dashed';
            case BorderStyle::Double:
                return 'double';
            case BorderStyle::Hair:
                return 'hair';
            case BorderStyle::MediumDashed:
                return 'mediumDashed';
            case BorderStyle::DashDot:
                return 'dashDot';
            case BorderStyle::MediumDashDot:
                return 'mediumDashDot';
            case BorderStyle::DashDotDot:
                return 'dashDotDot';
            case BorderStyle::MediumDashDotDot:
                return 'mediumDashDotDot';
            case BorderStyle::SlantedDashDot:
                return 'slantDashDot';
            default:
                return '';
        }
    }

    public static function GetHorizontalAlignmentName($value)
    {
        switch ($value) {
            case HorizontalAlignment::Left:
                return 'left';
            case HorizontalAlignment::Center:
                return 'center';
            case HorizontalAlignment::Right:
                return 'right';
            case HorizontalAlignment::Fill:
                return 'fill';
            case HorizontalAlignment::Justify:
                return 'justify';
            case HorizontalAlignment::CenterContinuous:
                return 'centerContinuous';
            case HorizontalAlignment::Distributed:
                return 'distributed';
            default:
                return '';
        }
    }

    public static function GetVerticalAlignmentName($value)
    {
        switch ($value) {
            case VerticalAlignment::Center:
                return 'center';
            case VerticalAlignment::Top:
                return 'top';
            case VerticalAlignment::Justify:
                return 'justify';
            case VerticalAlignment::Distributed:
                return 'distributed';
            default:
                return '';
        }
    }

    public static function ToArgbHex(ColorValue $color)
    {
        return sprintf('%02X%02X%02X%02X', $color->A, $color->R, $color->G, $color->B);
    }

    public static function IsEmptyColor(ColorValue $color)
    {
        return ($color->ThemeIndex === null && $color->Indexed === null)
            && $color->A == 0 && $color->R == 0 && $color->G == 0 && $color->B == 0;
    }

    public static function FontEquals(FontValue $left, FontValue $right)
    {
        return $left->Name === $right->Name
            && $left->Size == $right->Size
            && $left->Bold === $right->Bold
            && $left->Italic === $right->Italic
            && $left->Underline === $right->Underline
            && $left->StrikeThrough === $right->StrikeThrough
            && $left->Color->Equals($right->Color);
    }

    public static function BordersEqual(BordersValue $left, BordersValue $right)
    {
        return self::BorderSideEquals($left->Left, $right->Left)
            && self::BorderSideEquals($left->Right, $right->Right)
            && self::BorderSideEquals($left->Top, $right->Top)
            && self::BorderSideEquals($left->Bottom, $right->Bottom)
            && self::BorderSideEquals($left->Diagonal, $right->Diagonal)
            && $left->DiagonalUp === $right->DiagonalUp
            && $left->DiagonalDown === $right->DiagonalDown;
    }

    private static function BorderSideEquals(BorderSideValue $left, BorderSideValue $right)
    {
        return $left->Style === $right->Style && $left->Color->Equals($right->Color);
    }
}
