<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\PageMarginsModel;
use Aspose\Cells_FOSS\Core\PageOrientation;
use Aspose\Cells_FOSS\Core\PageSetupModel;
use Aspose\Cells_FOSS\Core\WorkbookModel;
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

// Serializes and loads worksheet page setup, print settings, and defined names.
final class XlsxWorkbookPageSetup
{
    private const MaxSpreadsheetRow = 1048575;
    private const MaxSpreadsheetColumn = 16383;

    private function __construct()
    {
    }

    public static function BuildPageSetupDefinedNames(WorkbookModel $model)
    {
        $definedNames = array();
        for ($sheetIndex = 0; $sheetIndex < count($model->Worksheets); $sheetIndex++) {
            $worksheet = $model->Worksheets[$sheetIndex];
            $printArea = self::NormalizePrintAreaList($worksheet->PageSetup->PrintArea, $worksheet->Name);
            if (!($printArea === null || $printArea === '')) {
                $definedNames[] = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'definedName',
                    new XAttribute('name', DefinedNameUtility::PrintAreaDefinedName),
                    new XAttribute('localSheetId', $sheetIndex),
                    $printArea);
            }

            $printTitles = self::BuildPrintTitlesDefinedNameText($worksheet->PageSetup, $worksheet->Name);
            if (!($printTitles === null || $printTitles === '')) {
                $definedNames[] = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'definedName',
                    new XAttribute('name', DefinedNameUtility::PrintTitlesDefinedName),
                    new XAttribute('localSheetId', $sheetIndex),
                    $printTitles);
            }
        }

        return count($definedNames) === 0 ? null : new XElement(XlsxWorkbookSerializerCommon::MainNs . 'definedNames', $definedNames);
    }

    public static function LoadWorksheetDefinedNames(XElement $workbookRoot, LoadDiagnostics $diagnostics, LoadOptions $options)
    {
        $states = array();
        $definedNamesElement = $workbookRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'definedNames');
        $definedNameElements = $definedNamesElement === null ? array() : $definedNamesElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'definedName');
        foreach ($definedNameElements as $definedName) {
            $name = trim((string) $definedName->Attribute('name'));
            if (strcasecmp($name, DefinedNameUtility::PrintAreaDefinedName) !== 0
                && strcasecmp($name, DefinedNameUtility::PrintTitlesDefinedName) !== 0) {
                continue;
            }

            $localSheetId = XlsxWorkbookArchiveHelpers::ParseIntAttribute($definedName->Attribute('localSheetId'));
            if ($localSheetId === null || $localSheetId < 0) {
                $issue = new LoadIssue('PG-L004', DiagnosticSeverity::Warning, "Defined name '" . $name . "' is missing a valid localSheetId and was ignored.", false, true);
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                continue;
            }

            if (array_key_exists($localSheetId, $states)) {
                $state = $states[$localSheetId];
            } else {
                $state = new WorksheetDefinedNamesState();
                $states[$localSheetId] = $state;
            }

            $definedNameValue = $definedName->Value;
            $text = $definedNameValue === null ? null : trim($definedNameValue);
            if ($text === null || $text === '') {
                continue;
            }

            try {
                $resolvedState = $state;
                $definedNameText = $text;
                if (strcasecmp($name, DefinedNameUtility::PrintAreaDefinedName) === 0) {
                    $resolvedState->PrintArea = self::NormalizeLoadedPrintArea($definedNameText);
                } else {
                    self::ParseLoadedPrintTitles($definedNameText, $resolvedState);
                }
            } catch (CellsException $exception) {
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException("Defined name '" . $name . "' is invalid.", $exception);
                }

                $issue = new LoadIssue('PG-L004', DiagnosticSeverity::Warning, "Defined name '" . $name . "' is invalid and was ignored.", false, true);
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            }
        }

        return $states;
    }

    public static function ApplyWorksheetDefinedNames(WorksheetModel $worksheetModel, $definedNamesState)
    {
        if ($definedNamesState === null) {
            return;
        }

        $worksheetModel->PageSetup->PrintArea = $definedNamesState->PrintArea;
        $worksheetModel->PageSetup->PrintTitleRows = $definedNamesState->PrintTitleRows;
        $worksheetModel->PageSetup->PrintTitleColumns = $definedNamesState->PrintTitleColumns;
    }

    public static function BuildSheetProperties(PageSetupModel $pageSetup)
    {
        if ($pageSetup->FitToWidth === null && $pageSetup->FitToHeight === null) {
            return null;
        }

        return new XElement(XlsxWorkbookSerializerCommon::MainNs . 'sheetPr',
            new XElement(XlsxWorkbookSerializerCommon::MainNs . 'pageSetUpPr', new XAttribute('fitToPage', 1)));
    }

    public static function BuildPrintOptionsElement(PageSetupModel $pageSetup)
    {
        if (!$pageSetup->PrintOptions->GridLines
            && !$pageSetup->PrintOptions->Headings
            && !$pageSetup->PrintOptions->HorizontalCentered
            && !$pageSetup->PrintOptions->VerticalCentered) {
            return null;
        }

        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'printOptions');
        if ($pageSetup->PrintOptions->Headings) {
            $element->SetAttributeValue('headings', 1);
        }

        if ($pageSetup->PrintOptions->GridLines) {
            $element->SetAttributeValue('gridLines', 1);
            $element->SetAttributeValue('gridLinesSet', 1);
        }

        if ($pageSetup->PrintOptions->HorizontalCentered) {
            $element->SetAttributeValue('horizontalCentered', 1);
        }

        if ($pageSetup->PrintOptions->VerticalCentered) {
            $element->SetAttributeValue('verticalCentered', 1);
        }

        return $element;
    }

    public static function BuildPageMarginsElement(PageSetupModel $pageSetup)
    {
        if (self::MarginsEqual($pageSetup->Margins, new PageMarginsModel())) {
            return null;
        }

        return new XElement(XlsxWorkbookSerializerCommon::MainNs . 'pageMargins',
            new XAttribute('left', (string) $pageSetup->Margins->Left),
            new XAttribute('right', (string) $pageSetup->Margins->Right),
            new XAttribute('top', (string) $pageSetup->Margins->Top),
            new XAttribute('bottom', (string) $pageSetup->Margins->Bottom),
            new XAttribute('header', (string) $pageSetup->Margins->Header),
            new XAttribute('footer', (string) $pageSetup->Margins->Footer));
    }

    public static function BuildPageSetupElement(PageSetupModel $pageSetup)
    {
        if ($pageSetup->PaperSize == 0
            && $pageSetup->Orientation == PageOrientation::Default
            && $pageSetup->FirstPageNumber === null
            && $pageSetup->Scale === null
            && $pageSetup->FitToWidth === null
            && $pageSetup->FitToHeight === null) {
            return null;
        }

        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'pageSetup');
        if ($pageSetup->PaperSize > 0) {
            $element->SetAttributeValue('paperSize', $pageSetup->PaperSize);
        }

        if ($pageSetup->Scale !== null) {
            $element->SetAttributeValue('scale', $pageSetup->Scale);
        }

        if ($pageSetup->FitToWidth !== null) {
            $element->SetAttributeValue('fitToWidth', $pageSetup->FitToWidth);
        }

        if ($pageSetup->FitToHeight !== null) {
            $element->SetAttributeValue('fitToHeight', $pageSetup->FitToHeight);
        }

        if ($pageSetup->FirstPageNumber !== null) {
            $element->SetAttributeValue('firstPageNumber', $pageSetup->FirstPageNumber);
            $element->SetAttributeValue('useFirstPageNumber', 1);
        }

        if ($pageSetup->Orientation == PageOrientation::Portrait) {
            $element->SetAttributeValue('orientation', 'portrait');
        } elseif ($pageSetup->Orientation == PageOrientation::Landscape) {
            $element->SetAttributeValue('orientation', 'landscape');
        }

        return $element;
    }

    public static function BuildHeaderFooterElement(PageSetupModel $pageSetup)
    {
        $hasHeaderFooter = !($pageSetup->HeaderFooter->LeftHeader === null || $pageSetup->HeaderFooter->LeftHeader === '')
            || !($pageSetup->HeaderFooter->CenterHeader === null || $pageSetup->HeaderFooter->CenterHeader === '')
            || !($pageSetup->HeaderFooter->RightHeader === null || $pageSetup->HeaderFooter->RightHeader === '')
            || !($pageSetup->HeaderFooter->LeftFooter === null || $pageSetup->HeaderFooter->LeftFooter === '')
            || !($pageSetup->HeaderFooter->CenterFooter === null || $pageSetup->HeaderFooter->CenterFooter === '')
            || !($pageSetup->HeaderFooter->RightFooter === null || $pageSetup->HeaderFooter->RightFooter === '');
        if (!$hasHeaderFooter) {
            return null;
        }

        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'headerFooter');
        if (!($pageSetup->HeaderFooter->LeftHeader === null || $pageSetup->HeaderFooter->LeftHeader === '') || !($pageSetup->HeaderFooter->CenterHeader === null || $pageSetup->HeaderFooter->CenterHeader === '') || !($pageSetup->HeaderFooter->RightHeader === null || $pageSetup->HeaderFooter->RightHeader === '')) {
            $element->Add(new XElement(XlsxWorkbookSerializerCommon::MainNs . 'oddHeader', self::ComposeHeaderFooterText($pageSetup->HeaderFooter->LeftHeader, $pageSetup->HeaderFooter->CenterHeader, $pageSetup->HeaderFooter->RightHeader)));
        }

        if (!($pageSetup->HeaderFooter->LeftFooter === null || $pageSetup->HeaderFooter->LeftFooter === '') || !($pageSetup->HeaderFooter->CenterFooter === null || $pageSetup->HeaderFooter->CenterFooter === '') || !($pageSetup->HeaderFooter->RightFooter === null || $pageSetup->HeaderFooter->RightFooter === '')) {
            $element->Add(new XElement(XlsxWorkbookSerializerCommon::MainNs . 'oddFooter', self::ComposeHeaderFooterText($pageSetup->HeaderFooter->LeftFooter, $pageSetup->HeaderFooter->CenterFooter, $pageSetup->HeaderFooter->RightFooter)));
        }

        return $element;
    }

    public static function BuildRowBreaksElement(PageSetupModel $pageSetup)
    {
        return self::BuildBreaksElement('rowBreaks', $pageSetup->HorizontalPageBreaks, self::MaxSpreadsheetColumn);
    }

    public static function BuildColumnBreaksElement(PageSetupModel $pageSetup)
    {
        return self::BuildBreaksElement('colBreaks', $pageSetup->VerticalPageBreaks, self::MaxSpreadsheetRow);
    }

    private static function BuildBreaksElement($elementName, $breaks, $maxIndex)
    {
        $distinctBreaks = array();
        foreach ($breaks as $value) {
            $distinctBreaks[$value] = true;
        }

        $orderedBreaks = array_keys($distinctBreaks);
        sort($orderedBreaks);
        if (count($orderedBreaks) === 0) {
            return null;
        }

        $breakElements = array();
        foreach ($orderedBreaks as $value) {
            $breakElements[] = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'brk',
                new XAttribute('id', $value),
                new XAttribute('max', $maxIndex),
                new XAttribute('man', 1));
        }

        return new XElement(XlsxWorkbookSerializerCommon::MainNs . $elementName,
            new XAttribute('count', count($orderedBreaks)),
            new XAttribute('manualBreakCount', count($orderedBreaks)),
            $breakElements);
    }

    public static function LoadWorksheetPageSetup(WorksheetModel $worksheetModel, XElement $worksheetRoot, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $pageSetup = $worksheetModel->PageSetup;
        self::LoadPageMargins($pageSetup, $worksheetRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'pageMargins'), $diagnostics, $options, $sheetName);
        self::LoadPageSetupCore($pageSetup, $worksheetRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'pageSetup'), $diagnostics, $options, $sheetName);
        self::LoadPrintOptions($pageSetup, $worksheetRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'printOptions'));
        self::LoadHeaderFooter($pageSetup, $worksheetRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'headerFooter'));
        self::LoadBreaks($pageSetup->HorizontalPageBreaks, $worksheetRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'rowBreaks'), $diagnostics, $options, $sheetName, 'row');
        self::LoadBreaks($pageSetup->VerticalPageBreaks, $worksheetRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'colBreaks'), $diagnostics, $options, $sheetName, 'column');
    }

    private static function LoadPageMargins(PageSetupModel $pageSetup, $marginsElement, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        if ($marginsElement === null) {
            return;
        }

        $pageSetup->Margins->Left = self::ParseMarginAttribute($marginsElement->Attribute('left'), $pageSetup->Margins->Left, $diagnostics, $options, $sheetName, 'left');
        $pageSetup->Margins->Right = self::ParseMarginAttribute($marginsElement->Attribute('right'), $pageSetup->Margins->Right, $diagnostics, $options, $sheetName, 'right');
        $pageSetup->Margins->Top = self::ParseMarginAttribute($marginsElement->Attribute('top'), $pageSetup->Margins->Top, $diagnostics, $options, $sheetName, 'top');
        $pageSetup->Margins->Bottom = self::ParseMarginAttribute($marginsElement->Attribute('bottom'), $pageSetup->Margins->Bottom, $diagnostics, $options, $sheetName, 'bottom');
        $pageSetup->Margins->Header = self::ParseMarginAttribute($marginsElement->Attribute('header'), $pageSetup->Margins->Header, $diagnostics, $options, $sheetName, 'header');
        $pageSetup->Margins->Footer = self::ParseMarginAttribute($marginsElement->Attribute('footer'), $pageSetup->Margins->Footer, $diagnostics, $options, $sheetName, 'footer');
    }

    private static function LoadPageSetupCore(PageSetupModel $pageSetup, $pageSetupElement, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        if ($pageSetupElement === null) {
            return;
        }

        $paperSize = self::ParsePositiveIntAttribute($pageSetupElement->Attribute('paperSize'), $diagnostics, $options, $sheetName, 'paperSize');
        $pageSetup->PaperSize = $paperSize === null ? 0 : $paperSize;
        $pageSetup->FirstPageNumber = self::ParsePositiveIntAttribute($pageSetupElement->Attribute('firstPageNumber'), $diagnostics, $options, $sheetName, 'firstPageNumber');
        $pageSetup->Scale = self::ParseBoundedIntAttribute($pageSetupElement->Attribute('scale'), 10, 400, $diagnostics, $options, $sheetName, 'scale');
        $pageSetup->FitToWidth = self::ParseNonNegativeIntAttribute($pageSetupElement->Attribute('fitToWidth'), $diagnostics, $options, $sheetName, 'fitToWidth');
        $pageSetup->FitToHeight = self::ParseNonNegativeIntAttribute($pageSetupElement->Attribute('fitToHeight'), $diagnostics, $options, $sheetName, 'fitToHeight');
        $pageSetup->Orientation = self::ParseOrientation((string) $pageSetupElement->Attribute('orientation'), $diagnostics, $options, $sheetName);
    }

    private static function LoadPrintOptions(PageSetupModel $pageSetup, $printOptionsElement)
    {
        if ($printOptionsElement === null) {
            return;
        }

        $pageSetup->PrintOptions->Headings = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($printOptionsElement->Attribute('headings'));
        $pageSetup->PrintOptions->GridLines = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($printOptionsElement->Attribute('gridLines'));
        $pageSetup->PrintOptions->HorizontalCentered = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($printOptionsElement->Attribute('horizontalCentered'));
        $pageSetup->PrintOptions->VerticalCentered = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($printOptionsElement->Attribute('verticalCentered'));
    }

    private static function LoadHeaderFooter(PageSetupModel $pageSetup, $headerFooterElement)
    {
        if ($headerFooterElement === null) {
            return;
        }

        $oddHeader = $headerFooterElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'oddHeader');
        self::ParseHeaderFooterText($oddHeader === null ? null : $oddHeader->Value, $leftHeader, $centerHeader, $rightHeader);
        $pageSetup->HeaderFooter->LeftHeader = $leftHeader;
        $pageSetup->HeaderFooter->CenterHeader = $centerHeader;
        $pageSetup->HeaderFooter->RightHeader = $rightHeader;

        $oddFooter = $headerFooterElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'oddFooter');
        self::ParseHeaderFooterText($oddFooter === null ? null : $oddFooter->Value, $leftFooter, $centerFooter, $rightFooter);
        $pageSetup->HeaderFooter->LeftFooter = $leftFooter;
        $pageSetup->HeaderFooter->CenterFooter = $centerFooter;
        $pageSetup->HeaderFooter->RightFooter = $rightFooter;
    }

    private static function LoadBreaks(&$target, $breaksElement, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $axis)
    {
        $target = array();
        if ($breaksElement === null) {
            return;
        }

        foreach ($breaksElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'brk') as $breakElement) {
            $id = self::ParseNonNegativeIntAttribute($breakElement->Attribute('id'), $diagnostics, $options, $sheetName, $axis . 'Break');
            if ($id === null) {
                continue;
            }

            if (!in_array($id, $target, true)) {
                $target[] = $id;
            }
        }
    }

    private static function ComposeHeaderFooterText($left, $center, $right)
    {
        $builder = '';
        if (!($left === null || $left === '')) {
            $builder .= '&L';
            $builder .= $left;
        }

        if (!($center === null || $center === '')) {
            $builder .= '&C';
            $builder .= $center;
        }

        if (!($right === null || $right === '')) {
            $builder .= '&R';
            $builder .= $right;
        }

        return $builder;
    }

    private static function ParseHeaderFooterText($value, &$left, &$center, &$right)
    {
        $left = null;
        $center = null;
        $right = null;
        if ($value === null || $value === '') {
            return;
        }

        $currentSection = 'C';
        $builder = '';

        $text = $value;
        for ($index = 0; $index < strlen($text); $index++) {
            if ($text[$index] === '&' && $index + 1 < strlen($text)) {
                $marker = strtoupper($text[$index + 1]);
                if ($marker === 'L' || $marker === 'C' || $marker === 'R') {
                    self::AssignHeaderFooterSection($currentSection, $builder, $left, $center, $right);
                    $builder = '';
                    $currentSection = $marker;
                    $index++;
                    continue;
                }
            }

            $builder .= $text[$index];
        }

        self::AssignHeaderFooterSection($currentSection, $builder, $left, $center, $right);
    }

    private static function AssignHeaderFooterSection($section, $value, &$left, &$center, &$right)
    {
        $text = ($value === null || $value === '') ? null : $value;
        if ($section === 'L') {
            $left = $text;
        } elseif ($section === 'R') {
            $right = $text;
        } else {
            $center = $text;
        }
    }

    private static function MarginsEqual(PageMarginsModel $left, PageMarginsModel $right)
    {
        return $left->Left == $right->Left
            && $left->Right == $right->Right
            && $left->Top == $right->Top
            && $left->Bottom == $right->Bottom
            && $left->Header == $right->Header
            && $left->Footer == $right->Footer;
    }

    private static function BuildPrintTitlesDefinedNameText(PageSetupModel $pageSetup, $sheetName)
    {
        $segments = array();
        $normalizedRows = self::NormalizeTitleRows($pageSetup->PrintTitleRows, $sheetName);
        if (!($normalizedRows === null || $normalizedRows === '')) {
            $segments[] = $normalizedRows;
        }

        $normalizedColumns = self::NormalizeTitleColumns($pageSetup->PrintTitleColumns, $sheetName);
        if (!($normalizedColumns === null || $normalizedColumns === '')) {
            $segments[] = $normalizedColumns;
        }

        return implode(',', $segments);
    }

    private static function NormalizePrintAreaList($printArea, $sheetName)
    {
        if ($printArea === null || trim($printArea) === '') {
            return '';
        }

        $segments = array();
        $printAreaText = $printArea;
        foreach (self::SplitReferenceList($printAreaText) as $segment) {
            $segments[] = self::QualifyReference($sheetName, self::NormalizeAreaReference(self::RemoveWorksheetQualifier($segment)));
        }

        return implode(',', $segments);
    }

    private static function NormalizeLoadedPrintArea($value)
    {
        $segments = array();
        foreach (self::SplitReferenceList($value) as $segment) {
            $segments[] = self::NormalizeAreaReference(self::RemoveWorksheetQualifier($segment));
        }

        return implode(',', $segments);
    }

    private static function ParseLoadedPrintTitles($value, WorksheetDefinedNamesState $state)
    {
        foreach (self::SplitReferenceList($value) as $segment) {
            $unqualified = self::RemoveWorksheetQualifier($segment);
            if (self::LooksLikeRowRange($unqualified)) {
                $state->PrintTitleRows = self::NormalizeRowReference($unqualified);
            } elseif (self::LooksLikeColumnRange($unqualified)) {
                $state->PrintTitleColumns = self::NormalizeColumnReference($unqualified);
            } else {
                throw new CellsException('Print title reference is invalid.');
            }
        }
    }

    private static function NormalizeTitleRows($value, $sheetName)
    {
        if ($value === null || trim($value) === '') {
            return '';
        }

        return self::QualifyReference($sheetName, self::NormalizeRowReference(self::RemoveWorksheetQualifier($value)));
    }

    private static function NormalizeTitleColumns($value, $sheetName)
    {
        if ($value === null || trim($value) === '') {
            return '';
        }

        return self::QualifyReference($sheetName, self::NormalizeColumnReference(self::RemoveWorksheetQualifier($value)));
    }

    private static function SplitReferenceList($value)
    {
        $parts = array();
        $builder = '';
        $inQuotes = false;

        $text = $value;
        for ($index = 0; $index < strlen($text); $index++) {
            $character = $value[$index];
            if ($character === "'") {
                $builder .= $character;
                if ($inQuotes && $index + 1 < strlen($value) && $value[$index + 1] === "'") {
                    $builder .= $value[$index + 1];
                    $index++;
                    continue;
                }

                $inQuotes = !$inQuotes;
                continue;
            }

            if ($character === ',' && !$inQuotes) {
                $part = trim($builder);
                if (strlen($part) > 0) {
                    $parts[] = $part;
                }

                $builder = '';
                continue;
            }

            $builder .= $character;
        }

        $lastPart = trim($builder);
        if (strlen($lastPart) > 0) {
            $parts[] = $lastPart;
        }

        return $parts;
    }

    private static function RemoveWorksheetQualifier($value)
    {
        $trimmed = trim($value);
        if (strlen($trimmed) === 0) {
            throw new CellsException('Worksheet reference is invalid.');
        }

        if ($trimmed[0] === "'") {
            $index = 1;
            while ($index < strlen($trimmed)) {
                if ($trimmed[$index] === "'") {
                    if ($index + 1 < strlen($trimmed) && $trimmed[$index + 1] === "'") {
                        $index += 2;
                        continue;
                    }

                    break;
                }

                $index++;
            }

            if ($index + 1 < strlen($trimmed) && $trimmed[$index + 1] === '!') {
                return substr($trimmed, $index + 2);
            }
        }

        $exclamation = strpos($trimmed, '!');
        return $exclamation !== false ? substr($trimmed, $exclamation + 1) : $trimmed;
    }

    private static function QualifyReference($sheetName, $reference)
    {
        return self::QuoteWorksheetName($sheetName) . '!' . $reference;
    }

    private static function QuoteWorksheetName($sheetName)
    {
        return "'" . str_replace("'", "''", $sheetName) . "'";
    }

    private static function NormalizeAreaReference($value)
    {
        $parts = explode(':', $value);
        if (count($parts) === 1) {
            $address = self::ParseAbsoluteCellAddress($parts[0]);
            return self::ToAbsoluteCellReference($address);
        }

        if (count($parts) !== 2) {
            throw new CellsException('PrintArea must be a cell or range reference.');
        }

        $first = self::ParseAbsoluteCellAddress($parts[0]);
        $last = self::ParseAbsoluteCellAddress($parts[1]);
        if ($last->RowIndex < $first->RowIndex || $last->ColumnIndex < $first->ColumnIndex) {
            throw new CellsException('PrintArea range must be ordered from top-left to bottom-right.');
        }

        return self::ToAbsoluteCellReference($first) . ':' . self::ToAbsoluteCellReference($last);
    }

    private static function NormalizeRowReference($value)
    {
        $normalized = trim(str_replace('$', '', $value));
        $parts = explode(':', $normalized);
        if (count($parts) !== 2 || !is_numeric($parts[0]) || !is_numeric($parts[1])) {
            throw new CellsException("PrintTitleRows must be a row span like '1:2'.");
        }

        $first = (int) $parts[0];
        $last = (int) $parts[1];
        if ($first <= 0 || $last < $first) {
            throw new CellsException("PrintTitleRows must be a row span like '1:2'.");
        }

        return '$' . $first . ':$' . $last;
    }

    private static function NormalizeColumnReference($value)
    {
        $normalized = strtoupper(trim(str_replace('$', '', $value)));
        $parts = explode(':', $normalized);
        if (count($parts) !== 2 || !self::IsColumnName($parts[0]) || !self::IsColumnName($parts[1])) {
            throw new CellsException("PrintTitleColumns must be a column span like 'A:B'.");
        }

        return '$' . $parts[0] . ':$' . $parts[1];
    }

    private static function LooksLikeRowRange($value)
    {
        foreach (explode(':', str_replace('$', '', $value)) as $part) {
            if (!is_numeric($part)) {
                return false;
            }
        }
        return true;
    }

    private static function LooksLikeColumnRange($value)
    {
        foreach (explode(':', str_replace('$', '', $value)) as $part) {
            if (!self::IsColumnName($part)) {
                return false;
            }
        }
        return true;
    }

    private static function IsColumnName($value)
    {
        if ($value === null || trim($value) === '') {
            return false;
        }

        $length = strlen($value);
        for ($index = 0; $index < $length; $index++) {
            if (!ctype_alpha($value[$index])) {
                return false;
            }
        }

        return true;
    }

    private static function ParseAbsoluteCellAddress($value)
    {
        $normalized = trim(str_replace('$', '', $value));
        try {
            return CellAddress::Parse($normalized);
        } catch (\InvalidArgumentException $exception) {
            throw new CellsException("Cell reference '" . $value . "' is invalid.", $exception);
        }
    }

    private static function ToAbsoluteCellReference(CellAddress $address)
    {
        $reference = $address->ToString();
        $splitIndex = 0;
        while ($splitIndex < strlen($reference) && ctype_alpha($reference[$splitIndex])) {
            $splitIndex++;
        }

        return '$' . substr($reference, 0, $splitIndex) . '$' . substr($reference, $splitIndex);
    }

    private static function ParseMarginAttribute($attribute, $fallback, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $attributeName)
    {
        if ($attribute === null) {
            return $fallback;
        }

        $value = XlsxWorkbookArchiveHelpers::ParseDoubleAttribute($attribute);
        if ($value !== null && $value >= 0) {
            return $value;
        }

        if ($options->StrictMode) {
            throw new InvalidFileFormatException("The page margin attribute '" . $attributeName . "' is invalid.");
        }

        $issue = new LoadIssue('PG-L001', DiagnosticSeverity::Warning, "Page margin attribute '" . $attributeName . "' is invalid and the default value was used.", false, true);
        $issue->SheetName = $sheetName;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        return $fallback;
    }

    private static function ParsePositiveIntAttribute($attribute, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $attributeName)
    {
        if ($attribute === null) {
            return null;
        }

        $value = XlsxWorkbookArchiveHelpers::ParseIntAttribute($attribute);
        if ($value !== null && $value > 0) {
            return $value;
        }

        if ($options->StrictMode) {
            throw new InvalidFileFormatException("The page setup attribute '" . $attributeName . "' is invalid.");
        }

        $issue = new LoadIssue('PG-L002', DiagnosticSeverity::Warning, "Page setup attribute '" . $attributeName . "' is invalid and was ignored.", false, true);
        $issue->SheetName = $sheetName;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        return null;
    }

    private static function ParseNonNegativeIntAttribute($attribute, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $attributeName)
    {
        if ($attribute === null) {
            return null;
        }

        $value = XlsxWorkbookArchiveHelpers::ParseIntAttribute($attribute);
        if ($value !== null && $value >= 0) {
            return $value;
        }

        if ($options->StrictMode) {
            throw new InvalidFileFormatException("The page setup attribute '" . $attributeName . "' is invalid.");
        }

        $issue = new LoadIssue('PG-L002', DiagnosticSeverity::Warning, "Page setup attribute '" . $attributeName . "' is invalid and was ignored.", false, true);
        $issue->SheetName = $sheetName;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        return null;
    }

    private static function ParseBoundedIntAttribute($attribute, $minimum, $maximum, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $attributeName)
    {
        if ($attribute === null) {
            return null;
        }

        $value = XlsxWorkbookArchiveHelpers::ParseIntAttribute($attribute);
        if ($value !== null && $value >= $minimum && $value <= $maximum) {
            return $value;
        }

        if ($options->StrictMode) {
            throw new InvalidFileFormatException("The page setup attribute '" . $attributeName . "' is invalid.");
        }

        $issue = new LoadIssue('PG-L002', DiagnosticSeverity::Warning, "Page setup attribute '" . $attributeName . "' is invalid and was ignored.", false, true);
        $issue->SheetName = $sheetName;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        return null;
    }

    private static function ParseOrientation($value, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        if ($value === null || trim($value) === '') {
            return PageOrientation::Default;
        }

        if (strcasecmp($value, 'portrait') === 0) {
            return PageOrientation::Portrait;
        }

        if (strcasecmp($value, 'landscape') === 0) {
            return PageOrientation::Landscape;
        }

        if ($options->StrictMode) {
            throw new InvalidFileFormatException('The page setup orientation is invalid.');
        }

        $issue = new LoadIssue('PG-L003', DiagnosticSeverity::Warning, 'Page setup orientation is invalid and the default orientation was used.', false, true);
        $issue->SheetName = $sheetName;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        return PageOrientation::Default;
    }
}
