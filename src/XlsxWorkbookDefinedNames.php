<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\DefinedNameModel;
use Aspose\Cells_FOSS\Core\WorkbookModel;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

final class XlsxWorkbookDefinedNames
{
    public static function BuildDefinedNames(WorkbookModel $model)
    {
        $definedNames = array();
        $pageSetupDefinedNames = XlsxWorkbookPageSetup::BuildPageSetupDefinedNames($model);
        if ($pageSetupDefinedNames !== null) {
            $elements = $pageSetupDefinedNames->Elements();
            foreach ($elements as $element) {
                $definedNames[] = $element;
            }
        }

        $autoFilterDefinedNames = self::BuildAutoFilterDefinedNames($model);
        if ($autoFilterDefinedNames !== null) {
            $elements = $autoFilterDefinedNames->Elements();
            foreach ($elements as $element) {
                $definedNames[] = $element;
            }
        }

        for ($index = 0; $index < count($model->DefinedNames); $index++) {
            $definedName = $model->DefinedNames[$index];
            $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'definedName',
                new XAttribute('name', $definedName->Name),
                $definedName->Formula);

            if ($definedName->LocalSheetIndex !== null) {
                $element->SetAttributeValue('localSheetId', (string)$definedName->LocalSheetIndex);
            }

            if ($definedName->Hidden) {
                $element->SetAttributeValue('hidden', 1);
            }

            if (!($definedName->Comment === null || $definedName->Comment === '')) {
                $element->SetAttributeValue('comment', $definedName->Comment);
            }

            $definedNames[] = $element;
        }

        if (count($definedNames) === 0) {
            return null;
        }

        return new XElement(XlsxWorkbookSerializerCommon::MainNs . 'definedNames', $definedNames);
    }

    public static function LoadWorkbookDefinedNames(XElement $workbookRoot, WorkbookModel $workbookModel, $sheetCount, LoadDiagnostics $diagnostics, LoadOptions $options)
    {
        $definedNamesElement = $workbookRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'definedNames');
        $elements = $definedNamesElement === null ? array() : $definedNamesElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'definedName');
        foreach ($elements as $element) {
            $name = trim((string)$element->Attribute('name'));
            if (DefinedNameUtility::IsReservedName($name)) {
                continue;
            }

            if (strlen($name) === 0) {
                self::HandleInvalidDefinedName($options, 'Workbook defined name is missing a valid name.');
                self::AddInvalidDefinedNameIssue($diagnostics, $options, 'Workbook defined name is missing a valid name and was ignored.');
                continue;
            }

            $localSheetIndex = null;
            $localSheetAttribute = $element->Attribute('localSheetId');
            if ($localSheetAttribute !== null) {
                $localSheetIndex = XlsxWorkbookArchiveHelpers::ParseIntAttribute($localSheetAttribute);
                if ($localSheetIndex === null || $localSheetIndex < 0 || $localSheetIndex >= $sheetCount) {
                    self::HandleInvalidDefinedName($options, "Workbook defined name '" . $name . "' has an invalid localSheetId.");
                    self::AddInvalidDefinedNameIssue($diagnostics, $options, "Workbook defined name '" . $name . "' has an invalid localSheetId and was ignored.");
                    continue;
                }
            }

            try {
                $formula = DefinedNameUtility::NormalizeFormula($element->Value);
            } catch (CellsException $exception) {
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException("Workbook defined name '" . $name . "' is invalid.", $exception);
                }

                self::AddInvalidDefinedNameIssue($diagnostics, $options, "Workbook defined name '" . $name . "' has an invalid formula and was ignored.");
                continue;
            }

            if (self::ContainsDuplicate($workbookModel->DefinedNames, $name, $localSheetIndex)) {
                self::HandleInvalidDefinedName($options, "Workbook defined name '" . $name . "' is duplicated in the same scope.");
                self::AddInvalidDefinedNameIssue($diagnostics, $options, "Workbook defined name '" . $name . "' duplicates an existing scope and was ignored.");
                continue;
            }

            $definedName = new DefinedNameModel();
            $definedName->Name = $name;
            $definedName->Formula = $formula;
            $definedName->LocalSheetIndex = $localSheetIndex;
            $definedName->Hidden = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($element->Attribute('hidden'));
            $definedName->Comment = DefinedNameUtility::NormalizeComment((string)$element->Attribute('comment'));
            $workbookModel->DefinedNames[] = $definedName;
        }
    }

    private static function ContainsDuplicate($definedNames, $name, $localSheetIndex)
    {
        for ($index = 0; $index < count($definedNames); $index++) {
            $existing = $definedNames[$index];
            if (strcasecmp($existing->Name, $name) !== 0) {
                continue;
            }

            if (DefinedNameUtility::SameScope($existing->LocalSheetIndex, $localSheetIndex)) {
                return true;
            }
        }

        return false;
    }

    private static function AddInvalidDefinedNameIssue(LoadDiagnostics $diagnostics, LoadOptions $options, $message)
    {
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, new LoadIssue('WB-L002', DiagnosticSeverity::Warning, $message, false, true));
    }

    private static function HandleInvalidDefinedName(LoadOptions $options, $message)
    {
        if ($options->StrictMode) {
            throw new InvalidFileFormatException($message);
        }
    }

    private static function BuildAutoFilterDefinedNames(WorkbookModel $model)
    {
        $definedNames = array();
        for ($sheetIndex = 0; $sheetIndex < count($model->Worksheets); $sheetIndex++) {
            $worksheet = $model->Worksheets[$sheetIndex];
            if ($worksheet->AutoFilter === null || $worksheet->AutoFilter->Range === null || $worksheet->AutoFilter->Range === '') {
                continue;
            }

            $definedNames[] = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'definedName',
                new XAttribute('name', DefinedNameUtility::FilterDatabaseDefinedName),
                new XAttribute('localSheetId', $sheetIndex),
                new XAttribute('hidden', 1),
                self::QualifyAutoFilterRange($worksheet->Name, $worksheet->AutoFilter->Range));
        }

        return count($definedNames) === 0 ? null : new XElement(XlsxWorkbookSerializerCommon::MainNs . 'definedNames', $definedNames);
    }

    private static function QualifyAutoFilterRange($sheetName, $range)
    {
        $parts = explode(':', $range);
        if (count($parts) === 1) {
            return self::QuoteWorksheetName($sheetName) . '!' . self::ToAbsoluteCellReference($parts[0]);
        }

        if (count($parts) !== 2) {
            throw new CellsException('AutoFilter range is invalid.');
        }

        return self::QuoteWorksheetName($sheetName) . '!' . self::ToAbsoluteCellReference($parts[0]) . ':' . self::ToAbsoluteCellReference($parts[1]);
    }

    private static function QuoteWorksheetName($sheetName)
    {
        return "'" . str_replace("'", "''", $sheetName) . "'";
    }

    private static function ToAbsoluteCellReference($value)
    {
        try {
            $address = CellAddress::Parse($value);
        } catch (\InvalidArgumentException $exception) {
            throw new CellsException('AutoFilter range is invalid.', $exception);
        }

        $reference = $address->ToString();
        $splitIndex = 0;
        $len = strlen($reference);
        while ($splitIndex < $len && ctype_alpha($reference[$splitIndex])) {
            $splitIndex++;
        }

        return '$' . substr($reference, 0, $splitIndex) . '$' . substr($reference, $splitIndex);
    }
}
