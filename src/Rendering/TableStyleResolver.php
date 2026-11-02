<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\BorderStyle;
use Aspose\Cells_FOSS\Core\ColorValue;
use Aspose\Cells_FOSS\Core\FillPatternKind;
use Aspose\Cells_FOSS\Core\StyleValue;
use Aspose\Cells_FOSS\Internal\Xml\XElement;
use Aspose\Cells_FOSS\ListObjectSupport;
use Aspose\Cells_FOSS\TableStyleType;

final class TableStyleResolver
{
    private const Main = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';

    private static $HeaderFill = null;
    private static $DataFill = null;
    private static $StripeFill = null;
    private static $BorderFill = null;
    private static $LightText = null;

    private $_workbook;
    private $_customStyles;

    public function __construct($workbook = null)
    {
        self::InitializeColors();
        $this->_workbook = $workbook;
        if ($workbook !== null && !($workbook->RawTableStylesXml === null || $workbook->RawTableStylesXml === '')) {
            try {
                $this->_customStyles = XElement::Parse($workbook->RawTableStylesXml);
            } catch (\Exception $ex) {
                $this->_customStyles = null;
            }
        }
    }

    private static function InitializeColors()
    {
        if (self::$HeaderFill !== null) {
            return;
        }

        self::$HeaderFill = new ColorValue(255, 0, 78, 110);
        self::$DataFill = new ColorValue(255, 255, 255, 255);
        self::$StripeFill = new ColorValue(255, 192, 230, 245);
        self::$BorderFill = new ColorValue(255, 159, 192, 205);
        self::$LightText = new ColorValue(255, 255, 255, 255);
    }

    public function Resolve($sheet, $row, $column, $baseStyle)
    {
        if ($sheet === null || $sheet->ListObjects === null || count($sheet->ListObjects) === 0) {
            return $baseStyle === null ? StyleValue::Default() : $baseStyle;
        }

        $style = $baseStyle === null ? StyleValue::Default() : $baseStyle;
        for ($i = 0; $i < count($sheet->ListObjects); $i++) {
            $table = $sheet->ListObjects[$i];
            if (!self::Contains($table, $row, $column)) {
                continue;
            }

            return $this->ApplyTableStyle($table, $row, $style);
        }

        return $style;
    }

    private static function Contains($table, $row, $column)
    {
        if ($table === null) {
            return false;
        }

        return $row >= $table->StartRow
            && $row <= $table->EndRow
            && $column >= $table->StartColumn
            && $column <= $table->EndColumn;
    }

    private function ApplyTableStyle($table, $row, $baseStyle)
    {
        $type = ListObjectSupport::TableStyleTypeFromName($table->TableStyleName);
        if ($type === TableStyleType::Custom) {
            return $this->ApplyCustomTableStyle($table, $row, $baseStyle);
        }

        if ($type !== TableStyleType::TableStyleMedium2) {
            return $baseStyle;
        }

        if ($table->ShowHeaderRow && $row === $table->StartRow) {
            $styled = $baseStyle->Clone();
            $styled->Pattern = FillPatternKind::Solid;
            $styled->ForegroundColor = clone self::$HeaderFill;
            $styled->BackgroundColor = clone self::$HeaderFill;
            $styled->Font->Bold = true;
            $styled->Font->Color = clone self::$LightText;
            self::ApplyBorders($styled, self::$LightText);
            return $styled;
        }

        if ($table->ShowTotals && $row === $table->EndRow) {
            $totals = $baseStyle->Clone();
            $totals->Pattern = FillPatternKind::Solid;
            $totals->ForegroundColor = clone self::$StripeFill;
            $totals->BackgroundColor = clone self::$StripeFill;
            $totals->Font->Bold = true;
            self::ApplyBorders($totals, self::$BorderFill);
            return $totals;
        }

        $data = $baseStyle->Clone();
        $data->Pattern = FillPatternKind::Solid;
        $data->ForegroundColor = clone self::$DataFill;
        $data->BackgroundColor = clone self::$DataFill;
        self::ApplyBorders($data, self::$BorderFill);

        if ($table->ShowRowStripes) {
            $dataStartRow = $table->StartRow + ($table->ShowHeaderRow ? 1 : 0);
            $dataEndRow = $table->EndRow - ($table->ShowTotals ? 1 : 0);
            if ($row >= $dataStartRow && $row <= $dataEndRow) {
                $stripeIndex = $row - $dataStartRow;
                if (($stripeIndex % 2) === 0) {
                    $striped = $data->Clone();
                    $striped->ForegroundColor = clone self::$StripeFill;
                    $striped->BackgroundColor = clone self::$StripeFill;
                    return $striped;
                }
            }
        }

        return $data;
    }

    private function ApplyCustomTableStyle($table, $row, $baseStyle)
    {
        if ($this->_customStyles === null || $this->_workbook === null) {
            return $baseStyle;
        }

        $definition = null;
        foreach ($this->_customStyles->Elements(self::Main . 'tableStyle') as $candidate) {
            if ((string)$candidate->Attribute('name') === $table->TableStyleName) {
                $definition = $candidate;
                break;
            }
        }

        if ($definition === null) {
            return $baseStyle;
        }

        $result = $baseStyle->Clone();
        $this->ApplyElement($definition, 'wholeTable', $result);
        if ($table->ShowHeaderRow && $row === $table->StartRow) {
            $this->ApplyElement($definition, 'headerRow', $result);
        } else if ($table->ShowTotals && $row === $table->EndRow) {
            $this->ApplyElement($definition, 'totalRow', $result);
        } else if ($table->ShowRowStripes) {
            $firstDataRow = $table->StartRow + ($table->ShowHeaderRow ? 1 : 0);
            if (($row - $firstDataRow) % 2 === 0) {
                $this->ApplyElement($definition, 'firstRowStripe', $result);
            } else {
                $this->ApplyElement($definition, 'secondRowStripe', $result);
            }
        }

        return $result;
    }

    private function ApplyElement($definition, $type, $result)
    {
        foreach ($definition->Elements(self::Main . 'tableStyleElement') as $element) {
            if ((string)$element->Attribute('type') !== $type) {
                continue;
            }

            $index = 0;
            if (!self::TryParseInt((string)$element->Attribute('dxfId'), $index)
                || $index < 0 || $index >= count($this->_workbook->TableDifferentialStyles)
                || $index >= count($this->_workbook->RawTableDifferentialStylesXml)) {
                return;
            }

            $raw = null;
            try {
                $raw = XElement::Parse($this->_workbook->RawTableDifferentialStylesXml[$index]);
            } catch (\Exception $ex) {
                return;
            }

            $value = $this->_workbook->TableDifferentialStyles[$index];
            $font = $raw->Element(self::Main . 'font');
            if ($font !== null) {
                $bold = $font->Element(self::Main . 'b');
                if ($bold !== null) {
                    $result->Font->Bold = (string)$bold->Attribute('val') !== '0';
                }

                if ($font->Element(self::Main . 'color') !== null) {
                    $result->Font->Color = clone $value->Font->Color;
                }
            }

            $fill = $raw->Element(self::Main . 'fill');
            $pattern = $fill !== null ? $fill->Element(self::Main . 'patternFill') : null;
            if ($pattern !== null) {
                $result->Pattern = $value->Pattern;
                $result->ForegroundColor = clone $value->ForegroundColor;
                $result->BackgroundColor = clone $value->BackgroundColor;
            }

            if ($raw->Element(self::Main . 'border') !== null) {
                $result->Borders = $value->Borders->Clone();
            }

            return;
        }
    }

    private static function TryParseInt($text, &$value)
    {
        $value = 0;
        if ($text === null) {
            return false;
        }

        $trimmed = trim($text);
        if (preg_match('/^[+-]?[0-9]+$/', $trimmed) !== 1) {
            return false;
        }

        $value = (int)$trimmed;
        return true;
    }

    private static function ApplyBorders($style, $color)
    {
        $style->Borders->Left->Style = BorderStyle::Thin;
        $style->Borders->Right->Style = BorderStyle::Thin;
        $style->Borders->Top->Style = BorderStyle::Thin;
        $style->Borders->Bottom->Style = BorderStyle::Thin;

        $style->Borders->Left->Color = clone $color;
        $style->Borders->Right->Color = clone $color;
        $style->Borders->Top->Color = clone $color;
        $style->Borders->Bottom->Color = clone $color;
    }
}
