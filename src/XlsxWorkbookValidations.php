<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\MergeRegion;
use Aspose\Cells_FOSS\Core\ValidationModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

final class XlsxWorkbookValidations
{
    private function __construct()
    {
    }

    public static function BuildDataValidationsElement(WorksheetModel $worksheet)
    {
        $ordered = self::GetOrderedValidations($worksheet->Validations);
        if (count($ordered) === 0) {
            return null;
        }

        $container = new XElement(XlsxWorkbookSerializerCommon::MainNs . "dataValidations");
        $count = 0;

        for ($index = 0; $index < count($ordered); $index++) {
            $validation = $ordered[$index];
            if (count($validation->Areas) === 0) {
                continue;
            }

            $element = new XElement(
                XlsxWorkbookSerializerCommon::MainNs . "dataValidation",
                new XAttribute("sqref", self::BuildSqref($validation->Areas))
            );

            $typeName = self::ToValidationTypeName($validation->Type);
            if (!($typeName === null || $typeName === "")) {
                $element->SetAttributeValue("type", $typeName);
            }

            $operatorName = self::ToOperatorName($validation->Operator);
            if (!($operatorName === null || $operatorName === "")) {
                $element->SetAttributeValue("operator", $operatorName);
            }

            if ($validation->AlertStyle !== ValidationAlertType::Stop) {
                $element->SetAttributeValue("errorStyle", self::ToAlertStyleName($validation->AlertStyle));
            }

            if ($validation->IgnoreBlank) {
                $element->SetAttributeValue("allowBlank", 1);
            }

            if (!$validation->InCellDropDown) {
                $element->SetAttributeValue("showDropDown", 1);
            }

            if ($validation->ShowInput) {
                $element->SetAttributeValue("showInputMessage", 1);
            }

            if ($validation->ShowError) {
                $element->SetAttributeValue("showErrorMessage", 1);
            }

            if (!($validation->ErrorTitle === null || $validation->ErrorTitle === "")) {
                $element->SetAttributeValue("errorTitle", $validation->ErrorTitle);
            }

            if (!($validation->ErrorMessage === null || $validation->ErrorMessage === "")) {
                $element->SetAttributeValue("error", $validation->ErrorMessage);
            }

            if (!($validation->InputTitle === null || $validation->InputTitle === "")) {
                $element->SetAttributeValue("promptTitle", $validation->InputTitle);
            }

            if (!($validation->InputMessage === null || $validation->InputMessage === "")) {
                $element->SetAttributeValue("prompt", $validation->InputMessage);
            }

            if (!($validation->Formula1 === null || $validation->Formula1 === "")) {
                $element->Add(new XElement(XlsxWorkbookSerializerCommon::MainNs . "formula1", $validation->Formula1));
            }

            if (!($validation->Formula2 === null || $validation->Formula2 === "")) {
                $element->Add(new XElement(XlsxWorkbookSerializerCommon::MainNs . "formula2", $validation->Formula2));
            }

            $container->Add($element);
            $count++;
        }

        if ($count === 0) {
            return null;
        }

        $container->SetAttributeValue("count", (string)$count);
        return $container;
    }

    public static function LoadValidations(WorksheetModel $worksheetModel, XElement $worksheetRoot, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $worksheetModel->Validations = array();

        $dataValidationsElement = $worksheetRoot->Element(XlsxWorkbookSerializerCommon::MainNs . "dataValidations");
        $validationElements = array();
        if ($dataValidationsElement !== null) {
            $validationElements = $dataValidationsElement->Elements(XlsxWorkbookSerializerCommon::MainNs . "dataValidation");
        }

        for ($i = 0; $i < count($validationElements); $i++) {
            $validationElement = $validationElements[$i];
            $validation = new ValidationModel();
            $sqrefAttr = $validationElement->Attribute("sqref");
            $sqref = $sqrefAttr === null ? null : (string)$sqrefAttr;
            if (!self::TryLoadAreas($worksheetModel->Validations, $validation, $sqref, $diagnostics, $options, $sheetName)) {
                continue;
            }

            $typeAttr = $validationElement->Attribute("type");
            $validation->Type = self::ParseValidationType($typeAttr === null ? null : (string)$typeAttr, $diagnostics, $options, $sheetName, $sqref);
            $operatorAttr = $validationElement->Attribute("operator");
            $validation->Operator = self::ParseOperatorType($operatorAttr === null ? null : (string)$operatorAttr, $diagnostics, $options, $sheetName, $sqref);
            $alertStyleAttr = $validationElement->Attribute("errorStyle");
            $validation->AlertStyle = self::ParseAlertStyle($alertStyleAttr === null ? null : (string)$alertStyleAttr, $diagnostics, $options, $sheetName, $sqref);
            $validation->IgnoreBlank = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($validationElement->Attribute("allowBlank"));
            $validation->InCellDropDown = !XlsxWorkbookArchiveHelpers::ParseBoolAttribute($validationElement->Attribute("showDropDown"));
            $validation->ShowInput = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($validationElement->Attribute("showInputMessage"));
            $validation->ShowError = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($validationElement->Attribute("showErrorMessage"));
            $errorTitleAttr = $validationElement->Attribute("errorTitle");
            $validation->ErrorTitle = self::NormalizeText($errorTitleAttr === null ? null : (string)$errorTitleAttr);
            $errorAttr = $validationElement->Attribute("error");
            $validation->ErrorMessage = self::NormalizeText($errorAttr === null ? null : (string)$errorAttr);
            $promptTitleAttr = $validationElement->Attribute("promptTitle");
            $validation->InputTitle = self::NormalizeText($promptTitleAttr === null ? null : (string)$promptTitleAttr);
            $promptAttr = $validationElement->Attribute("prompt");
            $validation->InputMessage = self::NormalizeText($promptAttr === null ? null : (string)$promptAttr);
            $formula1Element = $validationElement->Element(XlsxWorkbookSerializerCommon::MainNs . "formula1");
            $validation->Formula1 = self::NormalizeFormulaText($formula1Element === null ? null : $formula1Element->Value);
            $formula2Element = $validationElement->Element(XlsxWorkbookSerializerCommon::MainNs . "formula2");
            $validation->Formula2 = self::NormalizeFormulaText($formula2Element === null ? null : $formula2Element->Value);

            $worksheetModel->Validations[] = $validation;
        }
    }

    private static function TryLoadAreas($existingValidations, ValidationModel $candidate, $sqref, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        if ($sqref === null || trim($sqref) === '') {
            if ($options->StrictMode) {
                throw new InvalidFileFormatException('A data validation is missing sqref.');
            }

            $issue = new LoadIssue('DV-L002', DiagnosticSeverity::LossyRecoverable, 'A data validation without any valid areas was dropped during load.', true, true);
            $issue->SheetName = $sheetName;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return false;
        }

        $tokens = preg_split('/[ \t\r\n]+/', $sqref, -1, PREG_SPLIT_NO_EMPTY);
        $invalidAreaReported = false;
        $overlapReported = false;

        for ($index = 0; $index < count($tokens); $index++) {
            $token = $tokens[$index];
            $region = null;
            if (!XlsxWorkbookSerializerCommon::TryParseMergeReference($token, $region)) {
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException("The data validation sqref '" . $token . "' is invalid.");
                }

                if (!$invalidAreaReported) {
                    $issue = new LoadIssue('DV-L001', DiagnosticSeverity::LossyRecoverable, 'One or more data validation areas were invalid and were dropped during load.', true, true);
                    $issue->SheetName = $sheetName;
                    $issue->CellRef = $token;
                    XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                    $invalidAreaReported = true;
                }
                continue;
            }

            $area = CellArea::CreateCellArea($region->FirstRow, $region->FirstColumn, $region->FirstRow + $region->TotalRows - 1, $region->FirstColumn + $region->TotalColumns - 1);
            if (self::ConflictsWithExisting($existingValidations, $candidate, $area)) {
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException("The data validation sqref '" . $token . "' overlaps an existing validation.");
                }

                if (!$overlapReported) {
                    $issue = new LoadIssue('DV-L003', DiagnosticSeverity::LossyRecoverable, 'Overlapping data validation areas were normalized during load.', true, true);
                    $issue->SheetName = $sheetName;
                    $issue->CellRef = $token;
                    XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                    $overlapReported = true;
                }
                continue;
            }

            $candidate->Areas[] = $area;
        }

        ValidationCollection::SortAreas($candidate->Areas);
        if (count($candidate->Areas) === 0) {
            $issue = new LoadIssue('DV-L002', DiagnosticSeverity::LossyRecoverable, 'A data validation without any valid areas was dropped during load.', true, true);
            $issue->SheetName = $sheetName;
            $issue->CellRef = $sqref;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return false;
        }

        return true;
    }

    private static function ConflictsWithExisting($existingValidations, ValidationModel $candidate, CellArea $area)
    {
        for ($validationIndex = 0; $validationIndex < count($existingValidations); $validationIndex++) {
            $validation = $existingValidations[$validationIndex];
            for ($areaIndex = 0; $areaIndex < count($validation->Areas); $areaIndex++) {
                if (ValidationCollection::AreasOverlap($validation->Areas[$areaIndex], $area)) {
                    return true;
                }
            }
        }

        for ($areaIndex = 0; $areaIndex < count($candidate->Areas); $areaIndex++) {
            if (ValidationCollection::AreasOverlap($candidate->Areas[$areaIndex], $area)) {
                return true;
            }
        }

        return false;
    }

    private static function GetOrderedValidations($validations)
    {
        $ordered = array();
        for ($index = 0; $index < count($validations); $index++) {
            $validation = $validations[$index];
            if (count($validation->Areas) === 0) {
                continue;
            }

            ValidationCollection::SortAreas($validation->Areas);
            $ordered[] = $validation;
        }

        usort($ordered, array(__CLASS__, 'CompareValidations'));
        return $ordered;
    }

    private static function CompareValidations(ValidationModel $left, ValidationModel $right)
    {
        return ValidationCollection::CompareAreas($left->Areas[0], $right->Areas[0]);
    }

    private static function BuildSqref($areas)
    {
        $references = array();
        for ($index = 0; $index < count($areas); $index++) {
            $references[] = self::ToAreaReference($areas[$index]);
        }

        return implode(' ', $references);
    }

    private static function ToAreaReference(CellArea $area)
    {
        $first = (new CellAddress($area->FirstRow, $area->FirstColumn))->ToString();
        if ($area->TotalRows === 1 && $area->TotalColumns === 1) {
            return $first;
        }

        $last = (new CellAddress($area->FirstRow + $area->TotalRows - 1, $area->FirstColumn + $area->TotalColumns - 1))->ToString();
        return $first . ':' . $last;
    }

    private static function ParseValidationType($text, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $cellRef)
    {
        switch (self::NormalizeToken($text)) {
            case null:
                return ValidationType::AnyValue;
            case 'whole':
                return ValidationType::WholeNumber;
            case 'decimal':
                return ValidationType::Decimal;
            case 'list':
                return ValidationType::List;
            case 'date':
                return ValidationType::Date;
            case 'time':
                return ValidationType::Time;
            case 'textlength':
                return ValidationType::TextLength;
            case 'custom':
                return ValidationType::Custom;
            default:
                return self::WarnAndReturnAnyValue('type', $text, $diagnostics, $options, $sheetName, $cellRef);
        }
    }

    private static function ParseOperatorType($text, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $cellRef)
    {
        switch (self::NormalizeToken($text)) {
            case null:
                return OperatorType::None;
            case 'between':
                return OperatorType::Between;
            case 'equal':
                return OperatorType::Equal;
            case 'greaterthan':
                return OperatorType::GreaterThan;
            case 'greaterthanorequal':
                return OperatorType::GreaterOrEqual;
            case 'lessthan':
                return OperatorType::LessThan;
            case 'lessthanorequal':
                return OperatorType::LessOrEqual;
            case 'notbetween':
                return OperatorType::NotBetween;
            case 'notequal':
                return OperatorType::NotEqual;
            default:
                return self::WarnAndReturnNone('operator', $text, $diagnostics, $options, $sheetName, $cellRef);
        }
    }

    private static function ParseAlertStyle($text, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $cellRef)
    {
        switch (self::NormalizeToken($text)) {
            case null:
            case 'stop':
                return ValidationAlertType::Stop;
            case 'warning':
                return ValidationAlertType::Warning;
            case 'information':
                return ValidationAlertType::Information;
            default:
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException("The data validation errorStyle '" . $text . "' is invalid.");
                }

                $issue = new LoadIssue('DV-R001', DiagnosticSeverity::Recoverable, 'An unknown data validation errorStyle was normalized during load.', true);
                $issue->SheetName = $sheetName;
                $issue->CellRef = $cellRef;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                return ValidationAlertType::Stop;
        }
    }

    private static function WarnAndReturnAnyValue($attributeName, $text, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $cellRef)
    {
        if ($options->StrictMode) {
            throw new InvalidFileFormatException("The data validation " . $attributeName . " '" . $text . "' is invalid.");
        }

        $issue = new LoadIssue('DV-R001', DiagnosticSeverity::Recoverable, 'An unknown data validation ' . $attributeName . ' was normalized during load.', true);
        $issue->SheetName = $sheetName;
        $issue->CellRef = $cellRef;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        return ValidationType::AnyValue;
    }

    private static function WarnAndReturnNone($attributeName, $text, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $cellRef)
    {
        if ($options->StrictMode) {
            throw new InvalidFileFormatException("The data validation " . $attributeName . " '" . $text . "' is invalid.");
        }

        $issue = new LoadIssue('DV-R001', DiagnosticSeverity::Recoverable, 'An unknown data validation ' . $attributeName . ' was normalized during load.', true);
        $issue->SheetName = $sheetName;
        $issue->CellRef = $cellRef;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        return OperatorType::None;
    }

    private static function NormalizeToken($value)
    {
        if ($value === null) {
            return null;
        }

        $trimmed = trim($value);
        if (strlen($trimmed) === 0) {
            return null;
        }

        return strtolower($trimmed);
    }

    private static function NormalizeFormulaText($value)
    {
        if ($value === null) {
            return null;
        }

        $trimmed = trim($value);
        if (strlen($trimmed) === 0) {
            return null;
        }

        if ($trimmed[0] === '=') {
            return substr($trimmed, 1);
        }

        return $trimmed;
    }

    private static function NormalizeText($value)
    {
        if ($value === null || $value === '') {
            return null;
        }

        return $value;
    }

    private static function ToValidationTypeName($type)
    {
        switch ($type) {
            case ValidationType::WholeNumber:
                return 'whole';
            case ValidationType::Decimal:
                return 'decimal';
            case ValidationType::List:
                return 'list';
            case ValidationType::Date:
                return 'date';
            case ValidationType::Time:
                return 'time';
            case ValidationType::TextLength:
                return 'textLength';
            case ValidationType::Custom:
                return 'custom';
            default:
                return null;
        }
    }

    private static function ToOperatorName($type)
    {
        switch ($type) {
            case OperatorType::Between:
                return 'between';
            case OperatorType::Equal:
                return 'equal';
            case OperatorType::GreaterThan:
                return 'greaterThan';
            case OperatorType::GreaterOrEqual:
                return 'greaterThanOrEqual';
            case OperatorType::LessThan:
                return 'lessThan';
            case OperatorType::LessOrEqual:
                return 'lessThanOrEqual';
            case OperatorType::NotBetween:
                return 'notBetween';
            case OperatorType::NotEqual:
                return 'notEqual';
            default:
                return null;
        }
    }

    private static function ToAlertStyleName($type)
    {
        switch ($type) {
            case ValidationAlertType::Warning:
                return 'warning';
            case ValidationAlertType::Information:
                return 'information';
            default:
                return 'stop';
        }
    }
}
