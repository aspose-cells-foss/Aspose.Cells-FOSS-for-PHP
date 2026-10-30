<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\CellRecord;
use Aspose\Cells_FOSS\Core\CellValueKind;

class Cell
{
    private $_worksheet;
    private $_address;

    public function __construct($worksheet, CellAddress $address)
    {
        $this->_worksheet = $worksheet;
        $this->_address = $address;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Value') {
            $record = $this->TryGetRecord();
            return $record === null ? null : $record->Value;
        }

        if ($propertyName === 'StringValue') {
            $record = $this->TryGetRecord();
            return DisplayTextFormatter::FormatStringValue($record === null ? null : $record->Value);
        }

        if ($propertyName === 'DisplayStringValue') {
            $record = $this->TryGetRecord();
            // Display formatting is style-aware while StringValue stays a stable raw
            // representation of the logical cell payload.
            $style = $record === null ? $this->_worksheet->Workbook->Model->DefaultStyle : $record->Style;
            return DisplayTextFormatter::FormatDisplayValue($record === null ? null : $record->Value, $style, $this->_worksheet->Workbook->Model->Settings->DisplayCulture, $this->_worksheet->Workbook->Model->Settings->DateSystem);
        }

        if ($propertyName === 'Formula') {
            $record = $this->TryGetRecord();
            if ($record === null || $record->Formula === null || $record->Formula === '') {
                return '';
            }

            return "=" . $record->Formula;
        }

        if ($propertyName === 'Type') {
            $record = $this->TryGetRecord();
            if ($record === null) {
                return CellValueType::IsNull;
            }

            if (!($record->Formula === null || $record->Formula === '')) {
                return CellValueType::IsUnknown;
            }

            switch ($record->Kind) {
                case CellValueKind::String:
                    return CellValueType::IsString;
                case CellValueKind::Number:
                    return CellValueType::IsNumeric;
                case CellValueKind::Boolean:
                    return CellValueType::IsBool;
                case CellValueKind::DateTime:
                    return CellValueType::IsDateTime;
                case CellValueKind::Formula:
                    return CellValueType::IsUnknown;
                default:
                    return CellValueType::IsNull;
            }
        }

        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Value') {
            $this->AssignValue($value);
            return;
        }

        if ($propertyName === 'Formula') {
            $record = $this->GetOrCreateRecord();
            // Store formulas without a leading '=' so XML persistence and comparisons
            // have one normalized internal representation.
            $record->Formula = self::NormalizeFormula($value);
            if ($record->Formula === null || $record->Formula === '') {
                if ($record->Value === null) {
                    $record->Kind = CellValueKind::Blank;
                }

                return;
            }

            $record->Kind = CellValueKind::Formula;
            return;
        }

        if ($propertyName === 'StringValue' || $propertyName === 'DisplayStringValue' || $propertyName === 'Type') {
            throw new \RuntimeException($propertyName . ' is read-only.');
        }
    }

    public function PutValue($value = null, $isConverted = null, $setStyle = null)
    {
        if ($isConverted !== null) {
            if ($value === null) throw new \InvalidArgumentException('value');
            if ($setStyle === null) $setStyle = true;
            if ($setStyle) {
                // Keep API-compatible signature. Number-format side effects are not
                // applied in the current v0.1 implementation.
            }

            if (!$isConverted) {
                $this->PutValue($value);
                return;
            }

            $trimmed = trim((string)$value);
            if ($trimmed !== '' && preg_match('/^[+-]?[0-9]+$/', $trimmed) === 1) {
                $this->SetScalar((int)$trimmed, CellValueKind::Number);
                return;
            }

            if (is_numeric($value)) {
                $this->SetScalar((float)$value, CellValueKind::Number);
                return;
            }

            if (strcasecmp($value, 'true') === 0) {
                $this->SetScalar(true, CellValueKind::Boolean);
                return;
            }

            if (strcasecmp($value, 'false') === 0) {
                $this->SetScalar(false, CellValueKind::Boolean);
                return;
            }

            $timestamp = strtotime($value);
            if ($timestamp !== false) {
                $this->SetScalar(new \DateTime($value), CellValueKind::DateTime);
                return;
            }

            $this->PutValue($value);
            return;
        }

        if (is_string($value)) {
            if ($value === null) throw new \InvalidArgumentException('value');
            $this->SetScalar($value, CellValueKind::String);
            return;
        }

        if (is_bool($value)) {
            $this->SetScalar($value, CellValueKind::Boolean);
            return;
        }

        if ($value instanceof \DateTime) {
            $this->SetScalar($value, CellValueKind::DateTime);
            return;
        }

        if (is_int($value) || is_float($value)) {
            $this->SetScalar($value, CellValueKind::Number);
            return;
        }

        if ($value === null) {
            $this->ClearValue();
            return;
        }

        $this->AssignValue($value);
    }

    public function GetStyle($checkBorders = false)
    {
        if ($checkBorders) {
            // Kept for Aspose-compatible overload shape. Merge-aware resolution
            // is not differentiated in the current implementation.
        }

        $record = $this->TryGetRecord();
        $style = $record === null ? $this->_worksheet->Workbook->Model->DefaultStyle : $record->Style;
        return Style::FromCore($style->Clone());
    }

    public function SetStyle($style, $flag = null)
    {
        if ($style === null) throw new \InvalidArgumentException('style');

        if ($flag === null) {
            $this->GetOrCreateRecord()->Style = $style->ToCore();
            return;
        }

        if (is_bool($flag)) {
            if ($flag) {
                // Kept for Aspose-compatible overload shape. Explicit-only style
                // application is not differentiated in the current implementation.
            }

            $this->SetStyle($style);
            return;
        }

        if ($flag->All) {
            $this->SetStyle($style);
            return;
        }

        $target = $this->GetStyle();

        if ($flag->NumberFormat) {
            $target->Number = $style->Number;
            $target->Custom = $style->Custom;
        }

        if ($flag->CellShading) {
            $target->Pattern = $style->Pattern;
            $target->ForegroundColor = $style->ForegroundColor;
            $target->BackgroundColor = $style->BackgroundColor;
        }

        if ($flag->Font) {
            $target->Font->Name = $style->Font->Name;
            $target->Font->Size = $style->Font->Size;
            $target->Font->IsBold = $style->Font->IsBold;
            $target->Font->IsItalic = $style->Font->IsItalic;
            $target->Font->Underline = $style->Font->Underline;
            $target->Font->IsStrikeout = $style->Font->IsStrikeout;
            $target->Font->Color = $style->Font->Color;
        } else {
            if ($flag->FontName) {
                $target->Font->Name = $style->Font->Name;
            }

            if ($flag->FontSize) {
                $target->Font->Size = $style->Font->Size;
            }

            if ($flag->FontBold) {
                $target->Font->IsBold = $style->Font->IsBold;
            }

            if ($flag->FontItalic) {
                $target->Font->IsItalic = $style->Font->IsItalic;
            }

            if ($flag->FontUnderline) {
                $target->Font->Underline = $style->Font->Underline;
            }

            if ($flag->FontStrike) {
                $target->Font->IsStrikeout = $style->Font->IsStrikeout;
            }

            if ($flag->FontColor) {
                $target->Font->Color = $style->Font->Color;
            }
        }

        if ($flag->Borders) {
            $target->Borders->Left = $style->Borders->Left->Clone();
            $target->Borders->Right = $style->Borders->Right->Clone();
            $target->Borders->Top = $style->Borders->Top->Clone();
            $target->Borders->Bottom = $style->Borders->Bottom->Clone();
            $target->Borders->Diagonal = $style->Borders->Diagonal->Clone();
            $target->Borders->DiagonalUp = $style->Borders->DiagonalUp;
            $target->Borders->DiagonalDown = $style->Borders->DiagonalDown;
        } else {
            if ($flag->LeftBorder) {
                $target->Borders->Left = $style->Borders->Left->Clone();
            }

            if ($flag->RightBorder) {
                $target->Borders->Right = $style->Borders->Right->Clone();
            }

            if ($flag->TopBorder) {
                $target->Borders->Top = $style->Borders->Top->Clone();
            }

            if ($flag->BottomBorder) {
                $target->Borders->Bottom = $style->Borders->Bottom->Clone();
            }

            if ($flag->DiagonalUpBorder) {
                $target->Borders->Diagonal = $style->Borders->Diagonal->Clone();
                $target->Borders->DiagonalUp = $style->Borders->DiagonalUp;
            }

            if ($flag->DiagonalDownBorder) {
                $target->Borders->Diagonal = $style->Borders->Diagonal->Clone();
                $target->Borders->DiagonalDown = $style->Borders->DiagonalDown;
            }
        }

        if ($flag->Alignments) {
            $target->HorizontalAlignment = $style->HorizontalAlignment;
            $target->VerticalAlignment = $style->VerticalAlignment;
            $target->WrapText = $style->WrapText;
            $target->IndentLevel = $style->IndentLevel;
            $target->TextRotation = $style->TextRotation;
            $target->ReadingOrder = $style->ReadingOrder;
            $target->RelativeIndent = $style->RelativeIndent;
            $target->ShrinkToFit = $style->ShrinkToFit;
        } else {
            if ($flag->HorizontalAlignment) {
                $target->HorizontalAlignment = $style->HorizontalAlignment;
            }

            if ($flag->Indent) {
                $target->IndentLevel = $style->IndentLevel;
            }

            if ($flag->VerticalAlignment) {
                $target->VerticalAlignment = $style->VerticalAlignment;
            }

            if ($flag->WrapText) {
                $target->WrapText = $style->WrapText;
            }

            if ($flag->Rotation) {
                $target->TextRotation = $style->TextRotation;
            }

            if ($flag->TextDirection) {
                $target->ReadingOrder = $style->ReadingOrder;
            }

            if ($flag->ShrinkToFit) {
                $target->ShrinkToFit = $style->ShrinkToFit;
            }
        }

        if ($flag->Locked) {
            $target->IsLocked = $style->IsLocked;
        }

        if ($flag->Hidden) {
            $target->IsHidden = $style->IsHidden;
        }

        if ($flag->QuotePrefix) {
            $target->QuotePrefix = $style->QuotePrefix;
        }

        $this->SetStyle($target);
    }

    private function AssignValue($value)
    {
        if ($value === null) {
            $this->ClearValue();
            return;
        }

        if (is_string($value)) {
            $this->PutValue($value);
            return;
        }
        if (is_bool($value)) {
            $this->PutValue($value);
            return;
        }
        if ($value instanceof \DateTime) {
            $this->PutValue($value);
            return;
        }
        if (is_int($value)) {
            $this->PutValue($value);
            return;
        }
        if (is_float($value)) {
            $this->PutValue($value);
            return;
        }

        $this->PutValue((string)$value);
    }

    private function ClearValue()
    {
        $record = $this->GetOrCreateRecord();
        $record->Value = null;
        $record->Formula = null;
        $record->Kind = CellValueKind::Blank;
    }

    private function TryGetRecord()
    {
        $key = $this->_address->ToString();
        $record = null;
        if (array_key_exists($key, $this->_worksheet->Model->Cells)) {
            $record = $this->_worksheet->Model->Cells[$key];
        }
        return $record;
    }

    private function GetOrCreateRecord()
    {
        $key = $this->_address->ToString();
        if (array_key_exists($key, $this->_worksheet->Model->Cells)) {
            return $this->_worksheet->Model->Cells[$key];
        }

        // Allocate a record even for currently blank cells so style-only and
        // formula-ready cells can participate in later serialization.
        $record = new CellRecord();
        $record->Style = $this->_worksheet->Workbook->Model->DefaultStyle->Clone();
        $this->_worksheet->Model->Cells[$key] = $record;
        return $record;
    }

    private function SetScalar($value, $kind)
    {
        $record = $this->GetOrCreateRecord();
        $record->Value = $value;
        $record->Kind = $kind;
        $record->Formula = null;
    }

    private static function NormalizeFormula($value)
    {
        if ($value === null || trim($value) === '') {
            return '';
        }

        return $value[0] === '=' ? substr($value, 1) : $value;
    }
}
