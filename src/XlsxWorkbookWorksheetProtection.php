<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

final class XlsxWorkbookWorksheetProtection
{
    public static function BuildSheetProtectionElement(WorksheetModel $worksheet)
    {
        $protection = $worksheet->Protection;
        if (!$protection->HasStoredState()) {
            return null;
        }

        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'sheetProtection');
        $element->SetAttributeValue('sheet', 1);
        self::SetBoolAttributeWhenTrue($element, 'objects', $protection->Objects);
        self::SetBoolAttributeWhenTrue($element, 'scenarios', $protection->Scenarios);
        self::SetBoolAttributeWhenTrue($element, 'formatCells', $protection->FormatCells);
        self::SetBoolAttributeWhenTrue($element, 'formatColumns', $protection->FormatColumns);
        self::SetBoolAttributeWhenTrue($element, 'formatRows', $protection->FormatRows);
        self::SetBoolAttributeWhenTrue($element, 'insertColumns', $protection->InsertColumns);
        self::SetBoolAttributeWhenTrue($element, 'insertRows', $protection->InsertRows);
        self::SetBoolAttributeWhenTrue($element, 'insertHyperlinks', $protection->InsertHyperlinks);
        self::SetBoolAttributeWhenTrue($element, 'deleteColumns', $protection->DeleteColumns);
        self::SetBoolAttributeWhenTrue($element, 'deleteRows', $protection->DeleteRows);
        self::SetBoolAttributeWhenTrue($element, 'selectLockedCells', $protection->SelectLockedCells);
        self::SetBoolAttributeWhenTrue($element, 'sort', $protection->Sort);
        self::SetBoolAttributeWhenTrue($element, 'autoFilter', $protection->AutoFilter);
        self::SetBoolAttributeWhenTrue($element, 'pivotTables', $protection->PivotTables);
        self::SetBoolAttributeWhenTrue($element, 'selectUnlockedCells', $protection->SelectUnlockedCells);
        self::SetStringAttributeWhenPresent($element, 'password', $protection->PasswordHash);
        self::SetStringAttributeWhenPresent($element, 'algorithmName', $protection->AlgorithmName);
        self::SetStringAttributeWhenPresent($element, 'hashValue', $protection->HashValue);
        self::SetStringAttributeWhenPresent($element, 'saltValue', $protection->SaltValue);
        self::SetStringAttributeWhenPresent($element, 'spinCount', $protection->SpinCount);
        return $element;
    }

    public static function LoadWorksheetProtection(WorksheetModel $worksheetModel, XElement $worksheetRoot, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $worksheetModel->Protection->Clear();
        $protectionElement = $worksheetRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'sheetProtection');
        if ($protectionElement === null) {
            return;
        }

        $worksheetModel->Protection->IsProtected = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('sheet'), true, $diagnostics, $options, $sheetName, 'sheet');
        $worksheetModel->Protection->Objects = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('objects'), false, $diagnostics, $options, $sheetName, 'objects');
        $worksheetModel->Protection->Scenarios = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('scenarios'), false, $diagnostics, $options, $sheetName, 'scenarios');
        $worksheetModel->Protection->FormatCells = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('formatCells'), false, $diagnostics, $options, $sheetName, 'formatCells');
        $worksheetModel->Protection->FormatColumns = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('formatColumns'), false, $diagnostics, $options, $sheetName, 'formatColumns');
        $worksheetModel->Protection->FormatRows = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('formatRows'), false, $diagnostics, $options, $sheetName, 'formatRows');
        $worksheetModel->Protection->InsertColumns = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('insertColumns'), false, $diagnostics, $options, $sheetName, 'insertColumns');
        $worksheetModel->Protection->InsertRows = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('insertRows'), false, $diagnostics, $options, $sheetName, 'insertRows');
        $worksheetModel->Protection->InsertHyperlinks = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('insertHyperlinks'), false, $diagnostics, $options, $sheetName, 'insertHyperlinks');
        $worksheetModel->Protection->DeleteColumns = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('deleteColumns'), false, $diagnostics, $options, $sheetName, 'deleteColumns');
        $worksheetModel->Protection->DeleteRows = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('deleteRows'), false, $diagnostics, $options, $sheetName, 'deleteRows');
        $worksheetModel->Protection->SelectLockedCells = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('selectLockedCells'), false, $diagnostics, $options, $sheetName, 'selectLockedCells');
        $worksheetModel->Protection->Sort = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('sort'), false, $diagnostics, $options, $sheetName, 'sort');
        $worksheetModel->Protection->AutoFilter = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('autoFilter'), false, $diagnostics, $options, $sheetName, 'autoFilter');
        $worksheetModel->Protection->PivotTables = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('pivotTables'), false, $diagnostics, $options, $sheetName, 'pivotTables');
        $worksheetModel->Protection->SelectUnlockedCells = self::ParseProtectionBooleanAttribute($protectionElement->Attribute('selectUnlockedCells'), false, $diagnostics, $options, $sheetName, 'selectUnlockedCells');
        $worksheetModel->Protection->PasswordHash = self::ReadOptionalAttribute($protectionElement, 'password');
        $worksheetModel->Protection->AlgorithmName = self::ReadOptionalAttribute($protectionElement, 'algorithmName');
        $worksheetModel->Protection->HashValue = self::ReadOptionalAttribute($protectionElement, 'hashValue');
        $worksheetModel->Protection->SaltValue = self::ReadOptionalAttribute($protectionElement, 'saltValue');
        $worksheetModel->Protection->SpinCount = self::ReadOptionalAttribute($protectionElement, 'spinCount');
    }

    private static function ParseProtectionBooleanAttribute($attribute, $defaultValue, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $attributeName)
    {
        if ($attribute === null) {
            return $defaultValue;
        }

        $rawValue = trim((string)$attribute);
        if ($rawValue === '0' || strcasecmp($rawValue, 'false') === 0) {
            return false;
        }

        if ($rawValue === '1' || strcasecmp($rawValue, 'true') === 0) {
            return true;
        }

        if ($options->StrictMode) {
            throw new InvalidFileFormatException("Worksheet protection attribute '" . $attributeName . "' is invalid.");
        }

        $issue = new LoadIssue('WS-L009', DiagnosticSeverity::Warning, "Worksheet protection attribute '" . $attributeName . "' is invalid and the default value was used.", false, true);
        $issue->SheetName = $sheetName;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        return $defaultValue;
    }

    private static function ReadOptionalAttribute(XElement $element, $attributeName)
    {
        $value = trim((string)$element->Attribute($attributeName));
        if (strlen($value) === 0) {
            return null;
        }

        return $value;
    }

    private static function SetBoolAttributeWhenTrue($element, $attributeName, $value)
    {
        if ($value) {
            $element->SetAttributeValue($attributeName, 1);
        }
    }

    private static function SetStringAttributeWhenPresent($element, $attributeName, $value)
    {
        if (!($value === null || trim($value) === '')) {
            $element->SetAttributeValue($attributeName, $value);
        }
    }
}
