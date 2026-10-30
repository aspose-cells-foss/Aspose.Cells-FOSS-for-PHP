<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ColorValue;
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

final class XlsxWorkbookWorksheetViews
{
    public static function BuildWorksheetSheetProperties(WorksheetModel $worksheet)
    {
        $pageSetUpProperties = self::BuildPageSetupPropertiesElement($worksheet->PageSetup);
        $tabColorElement = self::BuildTabColorElement($worksheet->TabColor);
        if ($pageSetUpProperties === null && $tabColorElement === null) {
            return null;
        }

        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'sheetPr');
        if ($tabColorElement !== null) {
            $element->Add($tabColorElement);
        }

        if ($pageSetUpProperties !== null) {
            $element->Add($pageSetUpProperties);
        }

        return $element;
    }

    public static function BuildWorksheetViewsElement(WorksheetModel $worksheet)
    {
        if (self::WorksheetViewIsDefault($worksheet->View)) {
            return null;
        }

        $sheetView = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'sheetView',
            new XAttribute('workbookViewId', 0));

        if (!$worksheet->View->ShowGridLines) {
            $sheetView->SetAttributeValue('showGridLines', 0);
        }

        if (!$worksheet->View->ShowRowColumnHeaders) {
            $sheetView->SetAttributeValue('showRowColHeaders', 0);
        }

        if (!$worksheet->View->ShowZeros) {
            $sheetView->SetAttributeValue('showZeros', 0);
        }

        if ($worksheet->View->RightToLeft) {
            $sheetView->SetAttributeValue('rightToLeft', 1);
        }

        if ($worksheet->View->ZoomScale !== 100) {
            $sheetView->SetAttributeValue('zoomScale', (string)$worksheet->View->ZoomScale);
        }

        return new XElement(XlsxWorkbookSerializerCommon::MainNs . 'sheetViews', $sheetView);
    }

    public static function LoadWorksheetViewSettings(WorksheetModel $worksheetModel, XElement $worksheetRoot, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        self::LoadTabColor($worksheetModel, $worksheetRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'sheetPr'), $diagnostics, $options, $sheetName);
        $sheetViewsElement = $worksheetRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'sheetViews');
        $sheetViewElement = $sheetViewsElement === null ? null : $sheetViewsElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'sheetView');
        self::LoadSheetView($worksheetModel, $sheetViewElement, $diagnostics, $options, $sheetName);
    }

    private static function BuildPageSetupPropertiesElement($pageSetup)
    {
        if (!$pageSetup->FitToPage) {
            return null;
        }

        return new XElement(XlsxWorkbookSerializerCommon::MainNs . 'pageSetUpPr', new XAttribute('fitToPage', 1));
    }

    private static function BuildTabColorElement($color)
    {
        if ($color === null) {
            return null;
        }

        return new XElement(XlsxWorkbookSerializerCommon::MainNs . 'tabColor',
            new XAttribute('rgb', self::FormatColor($color)));
    }

    private static function WorksheetViewIsDefault($view)
    {
        return $view->ShowGridLines
            && $view->ShowRowColumnHeaders
            && $view->ShowZeros
            && !$view->RightToLeft
            && $view->ZoomScale === 100;
    }

    private static function LoadTabColor(WorksheetModel $worksheetModel, $sheetPropertiesElement, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $worksheetModel->TabColor = null;
        if ($sheetPropertiesElement === null) {
            return;
        }

        $tabColor = $sheetPropertiesElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'tabColor');
        if ($tabColor === null) {
            return;
        }

        $rgb = trim((string)$tabColor->Attribute('rgb'));
        if (strlen($rgb) === 0) {
            return;
        }

        $color = null;
        if (self::TryParseColor($rgb, $color)) {
            $worksheetModel->TabColor = $color;
            return;
        }

        if ($options->StrictMode) {
            throw new InvalidFileFormatException('Worksheet tab color is invalid.');
        }

        $issue = new LoadIssue('WS-L007', DiagnosticSeverity::Warning, 'Worksheet tab color is invalid and was ignored.', false, true);
        $issue->SheetName = $sheetName;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
    }

    private static function LoadSheetView(WorksheetModel $worksheetModel, $sheetViewElement, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $worksheetModel->View->ShowGridLines = true;
        $worksheetModel->View->ShowRowColumnHeaders = true;
        $worksheetModel->View->ShowZeros = true;
        $worksheetModel->View->RightToLeft = false;
        $worksheetModel->View->ZoomScale = 100;
        if ($sheetViewElement === null) {
            return;
        }

        $worksheetModel->View->ShowGridLines = self::ParseBooleanViewAttribute($sheetViewElement->Attribute('showGridLines'), true, $diagnostics, $options, $sheetName, 'showGridLines');
        $worksheetModel->View->ShowRowColumnHeaders = self::ParseBooleanViewAttribute($sheetViewElement->Attribute('showRowColHeaders'), true, $diagnostics, $options, $sheetName, 'showRowColHeaders');
        $worksheetModel->View->ShowZeros = self::ParseBooleanViewAttribute($sheetViewElement->Attribute('showZeros'), true, $diagnostics, $options, $sheetName, 'showZeros');
        $worksheetModel->View->RightToLeft = self::ParseBooleanViewAttribute($sheetViewElement->Attribute('rightToLeft'), false, $diagnostics, $options, $sheetName, 'rightToLeft');
        $worksheetModel->View->ZoomScale = self::ParseZoomScale($sheetViewElement->Attribute('zoomScale'), $diagnostics, $options, $sheetName);
    }

    private static function ParseBooleanViewAttribute($attribute, $defaultValue, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $attributeName)
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
            throw new InvalidFileFormatException("Worksheet view attribute '" . $attributeName . "' is invalid.");
        }

        $issue = new LoadIssue('WS-L008', DiagnosticSeverity::Warning, "Worksheet view attribute '" . $attributeName . "' is invalid and the default value was used.", false, true);
        $issue->SheetName = $sheetName;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        return $defaultValue;
    }

    private static function ParseZoomScale($attribute, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        if ($attribute === null) {
            return 100;
        }

        $value = XlsxWorkbookArchiveHelpers::ParseIntAttribute($attribute);
        if ($value !== null && $value >= 10 && $value <= 400) {
            return $value;
        }

        if ($options->StrictMode) {
            throw new InvalidFileFormatException('Worksheet zoomScale is invalid.');
        }

        $issue = new LoadIssue('WS-L008', DiagnosticSeverity::Warning, 'Worksheet zoomScale is invalid and the default value was used.', false, true);
        $issue->SheetName = $sheetName;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        return 100;
    }

    private static function FormatColor(ColorValue $color)
    {
        return sprintf('%02X%02X%02X%02X', $color->A, $color->R, $color->G, $color->B);
    }

    private static function TryParseColor($value, &$color)
    {
        $color = null;
        if ($value === null || strlen($value) !== 8 || !ctype_xdigit($value)) {
            return false;
        }

        $a = hexdec(substr($value, 0, 2));
        $r = hexdec(substr($value, 2, 2));
        $g = hexdec(substr($value, 4, 2));
        $b = hexdec(substr($value, 6, 2));

        $color = new ColorValue($a, $r, $g, $b);
        return true;
    }
}
