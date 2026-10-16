<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\DateSystem;
use Aspose\Cells_FOSS\Core\WorkbookModel;
use Aspose\Cells_FOSS\Internal\Xml\XElement;
use Aspose\Cells_FOSS\Internal\Xml\XNamespace;

final class XlsxWorkbookProperties
{
    public static function BuildWorkbookPropertiesElement(WorkbookModel $model)
    {
        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'workbookPr');
        $hasState = false;
        if ($model->Settings->DateSystem === DateSystem::Mac1904) {
            $element->SetAttributeValue('date1904', 1);
            $hasState = true;
        }

        $properties = $model->Properties;
        if (!($properties->CodeName === null || $properties->CodeName === '')) {
            $element->SetAttributeValue('codeName', $properties->CodeName);
            $hasState = true;
        }

        if (!($properties->ShowObjects === null || $properties->ShowObjects === '') && strcasecmp($properties->ShowObjects, 'all') !== 0) {
            $element->SetAttributeValue('showObjects', $properties->ShowObjects);
            $hasState = true;
        }

        if ($properties->FilterPrivacy) {
            $element->SetAttributeValue('filterPrivacy', 1);
            $hasState = true;
        }

        if (!$properties->ShowBorderUnselectedTables) {
            $element->SetAttributeValue('showBorderUnselectedTables', 0);
            $hasState = true;
        }

        if (!$properties->ShowInkAnnotation) {
            $element->SetAttributeValue('showInkAnnotation', 0);
            $hasState = true;
        }

        if ($properties->BackupFile) {
            $element->SetAttributeValue('backupFile', 1);
            $hasState = true;
        }

        if (!$properties->SaveExternalLinkValues) {
            $element->SetAttributeValue('saveExternalLinkValues', 0);
            $hasState = true;
        }

        if (!($properties->UpdateLinks === null || $properties->UpdateLinks === '') && strcasecmp($properties->UpdateLinks, 'userSet') !== 0) {
            $element->SetAttributeValue('updateLinks', $properties->UpdateLinks);
            $hasState = true;
        }

        if ($properties->HidePivotFieldList) {
            $element->SetAttributeValue('hidePivotFieldList', 1);
            $hasState = true;
        }

        if ($properties->DefaultThemeVersion !== null) {
            $element->SetAttributeValue('defaultThemeVersion', (string)$properties->DefaultThemeVersion);
            $hasState = true;
        }

        return $hasState ? $element : null;
    }

    public static function BuildWorkbookProtectionElement(WorkbookModel $model)
    {
        $protection = $model->Properties->Protection;
        if (!$protection->HasStoredState()) {
            return null;
        }

        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'workbookProtection');
        if ($protection->LockStructure) {
            $element->SetAttributeValue('lockStructure', 1);
        }

        if ($protection->LockWindows) {
            $element->SetAttributeValue('lockWindows', 1);
        }

        if ($protection->LockRevision) {
            $element->SetAttributeValue('lockRevision', 1);
        }

        if (!($protection->WorkbookPassword === null || $protection->WorkbookPassword === '')) {
            $element->SetAttributeValue('workbookPassword', $protection->WorkbookPassword);
        }

        if (!($protection->RevisionsPassword === null || $protection->RevisionsPassword === '')) {
            $element->SetAttributeValue('revisionsPassword', $protection->RevisionsPassword);
        }

        return $element;
    }

    public static function BuildBookViewsElement(WorkbookModel $model)
    {
        $view = $model->Properties->View;
        if (!$view->HasStoredState($model->ActiveSheetIndex)) {
            return null;
        }

        $workbookView = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'workbookView');
        if ($view->XWindow !== null) {
            $workbookView->SetAttributeValue('xWindow', (string)$view->XWindow);
        }

        if ($view->YWindow !== null) {
            $workbookView->SetAttributeValue('yWindow', (string)$view->YWindow);
        }

        if ($view->WindowWidth !== null) {
            $workbookView->SetAttributeValue('windowWidth', (string)$view->WindowWidth);
        }

        if ($view->WindowHeight !== null) {
            $workbookView->SetAttributeValue('windowHeight', (string)$view->WindowHeight);
        }

        if ($model->ActiveSheetIndex > 0 && $model->ActiveSheetIndex < count($model->Worksheets)) {
            $workbookView->SetAttributeValue('activeTab', (string)$model->ActiveSheetIndex);
        }

        if ($view->FirstSheet !== null) {
            $firstSheet = $view->FirstSheet;
            if (count($model->Worksheets) > 0 && $firstSheet >= count($model->Worksheets)) {
                $firstSheet = count($model->Worksheets) - 1;
            }

            if ($firstSheet < 0) {
                $firstSheet = 0;
            }

            $workbookView->SetAttributeValue('firstSheet', (string)$firstSheet);
        }

        if ($view->ShowHorizontalScroll !== null) {
            $workbookView->SetAttributeValue('showHorizontalScroll', $view->ShowHorizontalScroll ? 1 : 0);
        }

        if ($view->ShowVerticalScroll !== null) {
            $workbookView->SetAttributeValue('showVerticalScroll', $view->ShowVerticalScroll ? 1 : 0);
        }

        if ($view->ShowSheetTabs !== null) {
            $workbookView->SetAttributeValue('showSheetTabs', $view->ShowSheetTabs ? 1 : 0);
        }

        if ($view->TabRatio !== null) {
            $workbookView->SetAttributeValue('tabRatio', (string)$view->TabRatio);
        }

        if (!($view->Visibility === null || $view->Visibility === '') && strcasecmp($view->Visibility, 'visible') !== 0) {
            $workbookView->SetAttributeValue('visibility', $view->Visibility);
        }

        if ($view->Minimized) {
            $workbookView->SetAttributeValue('minimized', 1);
        }

        if (!$view->AutoFilterDateGrouping) {
            $workbookView->SetAttributeValue('autoFilterDateGrouping', 0);
        }

        return new XElement(XlsxWorkbookSerializerCommon::MainNs . 'bookViews', $workbookView);
    }

    public static function BuildCalculationPropertiesElement(WorkbookModel $model)
    {
        $calculation = $model->Properties->Calculation;
        if (!$calculation->HasStoredState()) {
            return null;
        }

        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'calcPr');
        if ($calculation->CalculationId !== null) {
            $element->SetAttributeValue('calcId', (string)$calculation->CalculationId);
        }

        if (!($calculation->CalculationMode === null || $calculation->CalculationMode === '') && strcasecmp($calculation->CalculationMode, 'auto') !== 0) {
            $element->SetAttributeValue('calcMode', $calculation->CalculationMode);
        }

        if ($calculation->FullCalculationOnLoad) {
            $element->SetAttributeValue('fullCalcOnLoad', 1);
        }

        if (!($calculation->ReferenceMode === null || $calculation->ReferenceMode === '') && strcasecmp($calculation->ReferenceMode, 'A1') !== 0) {
            $element->SetAttributeValue('refMode', $calculation->ReferenceMode);
        }

        if ($calculation->Iterate) {
            $element->SetAttributeValue('iterate', 1);
        }

        if ($calculation->IterateCount !== null) {
            $element->SetAttributeValue('iterateCount', (string)$calculation->IterateCount);
        }

        if ($calculation->IterateDelta !== null) {
            $element->SetAttributeValue('iterateDelta', self::FormatDouble($calculation->IterateDelta));
        }

        if ($calculation->FullPrecision !== null) {
            $element->SetAttributeValue('fullPrecision', $calculation->FullPrecision ? 1 : 0);
        }

        if ($calculation->CalculationCompleted !== null) {
            $element->SetAttributeValue('calcCompleted', $calculation->CalculationCompleted ? 1 : 0);
        }

        if ($calculation->CalculationOnSave !== null) {
            $element->SetAttributeValue('calcOnSave', $calculation->CalculationOnSave ? 1 : 0);
        }

        if ($calculation->ConcurrentCalculation !== null) {
            $element->SetAttributeValue('concurrentCalc', $calculation->ConcurrentCalculation ? 1 : 0);
        }

        if ($calculation->ForceFullCalculation) {
            $element->SetAttributeValue('forceFullCalc', 1);
        }

        return $element;
    }

    public static function LoadWorkbookMetadata(XElement $workbookRoot, WorkbookModel $workbookModel, $sheetCount, LoadDiagnostics $diagnostics, LoadOptions $options)
    {
        $workbookProperties = $workbookModel->Properties;
        $workbookPr = $workbookRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'workbookPr');
        if ($workbookPr === null) {
            $workbookModel->Settings->DateSystem = DateSystem::Windows1900;
        } else {
            $workbookModel->Settings->DateSystem = self::ReadBoolAttribute($workbookPr, 'date1904', $diagnostics, $options, false, '/xl/workbook.xml') ? DateSystem::Mac1904 : DateSystem::Windows1900;
            $workbookProperties->CodeName = self::ReadStringAttribute($workbookPr, 'codeName');
            $workbookProperties->ShowObjects = self::ReadChoiceAttribute($workbookPr, 'showObjects', $diagnostics, $options, '/xl/workbook.xml', array(WorkbookPropertySupport::class, 'NormalizeShowObjects'));
            $workbookProperties->FilterPrivacy = self::ReadBoolAttribute($workbookPr, 'filterPrivacy', $diagnostics, $options, false, '/xl/workbook.xml');
            $workbookProperties->ShowBorderUnselectedTables = self::ReadBoolAttribute($workbookPr, 'showBorderUnselectedTables', $diagnostics, $options, true, '/xl/workbook.xml');
            $workbookProperties->ShowInkAnnotation = self::ReadBoolAttribute($workbookPr, 'showInkAnnotation', $diagnostics, $options, true, '/xl/workbook.xml');
            $workbookProperties->BackupFile = self::ReadBoolAttribute($workbookPr, 'backupFile', $diagnostics, $options, false, '/xl/workbook.xml');
            $workbookProperties->SaveExternalLinkValues = self::ReadBoolAttribute($workbookPr, 'saveExternalLinkValues', $diagnostics, $options, true, '/xl/workbook.xml');
            $workbookProperties->UpdateLinks = self::ReadChoiceAttribute($workbookPr, 'updateLinks', $diagnostics, $options, '/xl/workbook.xml', array(WorkbookPropertySupport::class, 'NormalizeUpdateLinks'));
            $workbookProperties->HidePivotFieldList = self::ReadBoolAttribute($workbookPr, 'hidePivotFieldList', $diagnostics, $options, false, '/xl/workbook.xml');
            $workbookProperties->DefaultThemeVersion = self::ReadNonNegativeIntAttribute($workbookPr, 'defaultThemeVersion', $diagnostics, $options, '/xl/workbook.xml');
        }

        $protection = $workbookRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'workbookProtection');
        if ($protection !== null) {
            $workbookProperties->Protection->LockStructure = self::ReadBoolAttribute($protection, 'lockStructure', $diagnostics, $options, false, '/xl/workbook.xml');
            $workbookProperties->Protection->LockWindows = self::ReadBoolAttribute($protection, 'lockWindows', $diagnostics, $options, false, '/xl/workbook.xml');
            $workbookProperties->Protection->LockRevision = self::ReadBoolAttribute($protection, 'lockRevision', $diagnostics, $options, false, '/xl/workbook.xml');
            $workbookProperties->Protection->WorkbookPassword = self::ReadStringAttribute($protection, 'workbookPassword');
            $workbookProperties->Protection->RevisionsPassword = self::ReadStringAttribute($protection, 'revisionsPassword');
        }

        $bookViewsElement = $workbookRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'bookViews');
        $workbookView = $bookViewsElement === null ? null : $bookViewsElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'workbookView');
        if ($workbookView !== null) {
            $workbookProperties->View->XWindow = self::ReadNonNegativeIntAttribute($workbookView, 'xWindow', $diagnostics, $options, '/xl/workbook.xml');
            $workbookProperties->View->YWindow = self::ReadNonNegativeIntAttribute($workbookView, 'yWindow', $diagnostics, $options, '/xl/workbook.xml');
            $workbookProperties->View->WindowWidth = self::ReadNonNegativeIntAttribute($workbookView, 'windowWidth', $diagnostics, $options, '/xl/workbook.xml');
            $workbookProperties->View->WindowHeight = self::ReadNonNegativeIntAttribute($workbookView, 'windowHeight', $diagnostics, $options, '/xl/workbook.xml');
            $firstSheet = self::ReadNonNegativeIntAttribute($workbookView, 'firstSheet', $diagnostics, $options, '/xl/workbook.xml');
            if ($firstSheet !== null) {
                if ($firstSheet >= $sheetCount) {
                    self::AddWorkbookMetadataIssue($diagnostics, $options, '/xl/workbook.xml', 'Workbook firstSheet exceeded the worksheet count and was clamped.');
                    $firstSheet = $sheetCount > 0 ? $sheetCount - 1 : 0;
                }

                $workbookProperties->View->FirstSheet = $firstSheet;
            }

            $workbookProperties->View->ShowHorizontalScroll = self::ReadNullableBoolAttribute($workbookView, 'showHorizontalScroll', $diagnostics, $options, '/xl/workbook.xml');
            $workbookProperties->View->ShowVerticalScroll = self::ReadNullableBoolAttribute($workbookView, 'showVerticalScroll', $diagnostics, $options, '/xl/workbook.xml');
            $workbookProperties->View->ShowSheetTabs = self::ReadNullableBoolAttribute($workbookView, 'showSheetTabs', $diagnostics, $options, '/xl/workbook.xml');
            $tabRatio = self::ReadNonNegativeIntAttribute($workbookView, 'tabRatio', $diagnostics, $options, '/xl/workbook.xml');
            if ($tabRatio !== null) {
                if ($tabRatio > 1000) {
                    self::AddWorkbookMetadataIssue($diagnostics, $options, '/xl/workbook.xml', 'Workbook tabRatio was out of range and was ignored.');
                } else {
                    $workbookProperties->View->TabRatio = $tabRatio;
                }
            }

            $workbookProperties->View->Visibility = self::ReadChoiceAttribute($workbookView, 'visibility', $diagnostics, $options, '/xl/workbook.xml', array(WorkbookPropertySupport::class, 'NormalizeVisibility'));
            $workbookProperties->View->Minimized = self::ReadBoolAttribute($workbookView, 'minimized', $diagnostics, $options, false, '/xl/workbook.xml');
            $workbookProperties->View->AutoFilterDateGrouping = self::ReadBoolAttribute($workbookView, 'autoFilterDateGrouping', $diagnostics, $options, true, '/xl/workbook.xml');

            $activeTab = self::ReadNonNegativeIntAttribute($workbookView, 'activeTab', $diagnostics, $options, '/xl/workbook.xml');
            if ($activeTab !== null) {
                if ($activeTab >= $sheetCount) {
                    self::AddWorkbookMetadataIssue($diagnostics, $options, '/xl/workbook.xml', 'Workbook activeTab exceeded the worksheet count and was ignored.');
                } else {
                    $workbookModel->ActiveSheetIndex = $activeTab;
                }
            }
        }

        $calcPr = $workbookRoot->Element(XlsxWorkbookSerializerCommon::MainNs . 'calcPr');
        if ($calcPr !== null) {
            $workbookProperties->Calculation->CalculationId = self::ReadNonNegativeIntAttribute($calcPr, 'calcId', $diagnostics, $options, '/xl/workbook.xml');
            $workbookProperties->Calculation->CalculationMode = self::ReadChoiceAttribute($calcPr, 'calcMode', $diagnostics, $options, '/xl/workbook.xml', array(WorkbookPropertySupport::class, 'NormalizeCalculationMode'));
            $workbookProperties->Calculation->FullCalculationOnLoad = self::ReadBoolAttribute($calcPr, 'fullCalcOnLoad', $diagnostics, $options, false, '/xl/workbook.xml');
            $workbookProperties->Calculation->ReferenceMode = self::ReadChoiceAttribute($calcPr, 'refMode', $diagnostics, $options, '/xl/workbook.xml', array(WorkbookPropertySupport::class, 'NormalizeReferenceMode'));
            $workbookProperties->Calculation->Iterate = self::ReadBoolAttribute($calcPr, 'iterate', $diagnostics, $options, false, '/xl/workbook.xml');
            $workbookProperties->Calculation->IterateCount = self::ReadNonNegativeIntAttribute($calcPr, 'iterateCount', $diagnostics, $options, '/xl/workbook.xml');
            $workbookProperties->Calculation->IterateDelta = self::ReadNonNegativeDoubleAttribute($calcPr, 'iterateDelta', $diagnostics, $options, '/xl/workbook.xml');
            $workbookProperties->Calculation->FullPrecision = self::ReadNullableBoolAttribute($calcPr, 'fullPrecision', $diagnostics, $options, '/xl/workbook.xml');
            $workbookProperties->Calculation->CalculationCompleted = self::ReadNullableBoolAttribute($calcPr, 'calcCompleted', $diagnostics, $options, '/xl/workbook.xml');
            $workbookProperties->Calculation->CalculationOnSave = self::ReadNullableBoolAttribute($calcPr, 'calcOnSave', $diagnostics, $options, '/xl/workbook.xml');
            $workbookProperties->Calculation->ConcurrentCalculation = self::ReadNullableBoolAttribute($calcPr, 'concurrentCalc', $diagnostics, $options, '/xl/workbook.xml');
            $workbookProperties->Calculation->ForceFullCalculation = self::ReadBoolAttribute($calcPr, 'forceFullCalc', $diagnostics, $options, false, '/xl/workbook.xml');
        }
    }

    private static function ReadStringAttribute(XElement $element, $name)
    {
        return trim((string)$element->Attribute($name));
    }

    private static function ReadChoiceAttribute(XElement $element, $attributeName, LoadDiagnostics $diagnostics, LoadOptions $options, $partUri, $normalizer)
    {
        $attribute = $element->Attribute($attributeName);
        if ($attribute === null) {
            return '';
        }

        try {
            return call_user_func($normalizer, $attribute->Value);
        } catch (CellsException $exception) {
            self::AddWorkbookMetadataIssue($diagnostics, $options, $partUri, "Workbook metadata attribute '" . $attributeName . "' had an invalid value and was ignored.");
            return '';
        }
    }

    private static function ReadBoolAttribute(XElement $element, $attributeName, LoadDiagnostics $diagnostics, LoadOptions $options, $defaultValue, $partUri)
    {
        $value = self::ReadNullableBoolAttribute($element, $attributeName, $diagnostics, $options, $partUri);
        return $value ?? $defaultValue;
    }

    private static function ReadNullableBoolAttribute(XElement $element, $attributeName, LoadDiagnostics $diagnostics, LoadOptions $options, $partUri)
    {
        $attribute = $element->Attribute($attributeName);
        if ($attribute === null) {
            return null;
        }

        $value = false;
        if (self::TryReadBoolean($attribute->Value, $value)) {
            return $value;
        }

        self::AddWorkbookMetadataIssue($diagnostics, $options, $partUri, "Workbook metadata attribute '" . $attributeName . "' had an invalid Boolean value and was ignored.");
        return null;
    }

    private static function ReadNonNegativeIntAttribute(XElement $element, $attributeName, LoadDiagnostics $diagnostics, LoadOptions $options, $partUri)
    {
        $attribute = $element->Attribute($attributeName);
        if ($attribute === null) {
            return null;
        }

        $rawValue = trim($attribute->Value);
        if (preg_match('/^[+-]?[0-9]+$/', $rawValue) === 1) {
            $value = intval($rawValue);
            if ($value >= 0) {
                return $value;
            }
        }

        self::AddWorkbookMetadataIssue($diagnostics, $options, $partUri, "Workbook metadata attribute '" . $attributeName . "' had an invalid integer value and was ignored.");
        return null;
    }

    private static function ReadNonNegativeDoubleAttribute(XElement $element, $attributeName, LoadDiagnostics $diagnostics, LoadOptions $options, $partUri)
    {
        $attribute = $element->Attribute($attributeName);
        if ($attribute === null) {
            return null;
        }

        $rawValue = trim($attribute->Value);
        if (is_numeric($rawValue)) {
            $value = (float)$rawValue;
            if ($value >= 0) {
                return $value;
            }
        }

        self::AddWorkbookMetadataIssue($diagnostics, $options, $partUri, "Workbook metadata attribute '" . $attributeName . "' had an invalid numeric value and was ignored.");
        return null;
    }

    private static function TryReadBoolean($rawValue, &$value)
    {
        if ($rawValue === '1' || strcasecmp($rawValue, 'true') === 0) {
            $value = true;
            return true;
        }

        if ($rawValue === '0' || strcasecmp($rawValue, 'false') === 0) {
            $value = false;
            return true;
        }

        $value = false;
        return false;
    }

    private static function AddWorkbookMetadataIssue(LoadDiagnostics $diagnostics, LoadOptions $options, $partUri, $message)
    {
        $issue = new LoadIssue('WB-L003', DiagnosticSeverity::Warning, $message);
        $issue->PartUri = $partUri;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
    }

    private static function FormatDouble($value)
    {
        $formatted = rtrim(rtrim(sprintf('%.16F', $value), '0'), '.');
        if ($formatted === '' || $formatted === '-') {
            return '0';
        }

        return $formatted;
    }
}
