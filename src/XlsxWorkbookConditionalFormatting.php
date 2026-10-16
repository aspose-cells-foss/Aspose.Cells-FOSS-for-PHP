<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellAddress;
use Aspose\Cells_FOSS\Core\ColorValue;
use Aspose\Cells_FOSS\Core\ConditionalFormattingModel;
use Aspose\Cells_FOSS\Core\FormatConditionModel;
use Aspose\Cells_FOSS\Core\MergeRegion;
use Aspose\Cells_FOSS\Core\StyleValue;
use Aspose\Cells_FOSS\Core\WorksheetModel;
use Aspose\Cells_FOSS\Internal\Xml\XAttribute;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

final class XlsxWorkbookConditionalFormatting
{
    public static function BuildConditionalFormattingElements(WorksheetModel $worksheet, StylesheetSaveContext $stylesheet)
    {
        $ordered = self::GetOrderedConditionalFormattings($worksheet->ConditionalFormattings);
        $elements = array();

        for ($index = 0; $index < count($ordered); $index++) {
            $formatting = $ordered[$index];
            if (count($formatting->Areas) === 0 || count($formatting->Conditions) === 0) {
                continue;
            }

            $element = new XElement(
                XlsxWorkbookSerializerCommon::MainNs . 'conditionalFormatting',
                new XAttribute('sqref', self::BuildSqref($formatting->Areas))
            );

            for ($conditionIndex = 0; $conditionIndex < count($formatting->Conditions); $conditionIndex++) {
                $element->Add(self::BuildConditionRule($formatting, $formatting->Conditions[$conditionIndex], $conditionIndex, $stylesheet));
            }

            $elements[] = $element;
        }

        return $elements;
    }

    public static function LoadConditionalFormattings(WorksheetModel $worksheetModel, XElement $worksheetRoot, StylesheetLoadContext $stylesheet, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        $worksheetModel->ConditionalFormattings = array();

        $formattingElements = $worksheetRoot->Elements(XlsxWorkbookSerializerCommon::MainNs . 'conditionalFormatting');
        foreach ($formattingElements as $formattingElement) {
            $formatting = new ConditionalFormattingModel();
            $sqref = (string)$formattingElement->Attribute('sqref');
            if (!self::TryLoadAreas($formatting, $sqref, $diagnostics, $options, $sheetName)) {
                continue;
            }

            $ruleElements = $formattingElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'cfRule');
            foreach ($ruleElements as $ruleElement) {
                if (!self::TryLoadCondition($formatting, $ruleElement, $stylesheet, $diagnostics, $options, $sheetName, $sqref)) {
                    continue;
                }
            }

            if (count($formatting->Conditions) === 0) {
                $issue = new LoadIssue('CF-L003', DiagnosticSeverity::LossyRecoverable, 'A conditional formatting entry without any supported rules was dropped during load.', true, true);
                $issue->SheetName = $sheetName;
                $issue->CellRef = $sqref;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                continue;
            }

            $worksheetModel->ConditionalFormattings[] = $formatting;
        }
    }

    private static function BuildConditionRule(ConditionalFormattingModel $formatting, FormatConditionModel $condition, $conditionIndex, StylesheetSaveContext $stylesheet)
    {
        $rule = new XElement(
            XlsxWorkbookSerializerCommon::MainNs . 'cfRule',
            new XAttribute('type', self::ToRuleTypeName($condition)),
            new XAttribute('priority', $condition->Priority > 0 ? $condition->Priority : $conditionIndex + 1)
        );

        $differentialStyleIndex = $stylesheet->GetDifferentialStyleIndex($condition);
        if ($differentialStyleIndex !== null) {
            $rule->SetAttributeValue('dxfId', $differentialStyleIndex);
        }

        if ($condition->StopIfTrue) {
            $rule->SetAttributeValue('stopIfTrue', 1);
        }

        switch ($condition->Type) {
            case FormatConditionType::CellValue:
                $operatorName = self::ToOperatorName($condition->Operator);
                if (!($operatorName === null || $operatorName === '')) {
                    $rule->SetAttributeValue('operator', $operatorName);
                }

                self::AddFormulaElement($rule, $condition->Formula1);
                self::AddFormulaElement($rule, $condition->Formula2);
                break;
            case FormatConditionType::Expression:
                self::AddFormulaElement($rule, $condition->Formula1);
                break;
            case FormatConditionType::ContainsText:
            case FormatConditionType::NotContainsText:
            case FormatConditionType::BeginsWith:
            case FormatConditionType::EndsWith:
                if (!($condition->Formula1 === null || $condition->Formula1 === '')) {
                    $rule->SetAttributeValue('text', $condition->Formula1);
                    self::AddFormulaElement($rule, self::BuildTextRuleFormula($condition->Type, $condition->Formula1, self::GetAnchorCellReference($formatting)));
                }
                break;
            case FormatConditionType::TimePeriod:
                if (!($condition->TimePeriod === null || $condition->TimePeriod === '')) {
                    $rule->SetAttributeValue('timePeriod', $condition->TimePeriod);
                }
                break;
            case FormatConditionType::Top10:
            case FormatConditionType::Bottom10:
                $rule->SetAttributeValue('bottom', $condition->Type === FormatConditionType::Bottom10 || !$condition->Top ? 1 : 0);
                if ($condition->Percent) {
                    $rule->SetAttributeValue('percent', 1);
                }

                if ($condition->Rank > 0) {
                    $rule->SetAttributeValue('rank', $condition->Rank);
                }
                break;
            case FormatConditionType::AboveAverage:
            case FormatConditionType::BelowAverage:
                if ($condition->Type === FormatConditionType::BelowAverage || !$condition->Above) {
                    $rule->SetAttributeValue('aboveAverage', 0);
                }

                if ($condition->StandardDeviation > 0) {
                    $rule->SetAttributeValue('stdDev', $condition->StandardDeviation);
                }
                break;
            case FormatConditionType::ColorScale:
                $rule->Add(self::BuildColorScaleElement($condition));
                break;
            case FormatConditionType::DataBar:
                $rule->Add(self::BuildDataBarElement($condition));
                break;
            case FormatConditionType::IconSet:
                $rule->Add(self::BuildIconSetElement($condition));
                break;
        }

        return $rule;
    }

    private static function TryLoadAreas(ConditionalFormattingModel $formatting, $sqref, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName)
    {
        if ($sqref === null || trim($sqref) === '') {
            if ($options->StrictMode) {
                throw new InvalidFileFormatException('A conditional formatting entry is missing sqref.');
            }

            $issue = new LoadIssue('CF-L003', DiagnosticSeverity::LossyRecoverable, 'A conditional formatting entry without any valid areas was dropped during load.', true, true);
            $issue->SheetName = $sheetName;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return false;
        }

        $tokens = preg_split('/[ \t\r\n]+/', $sqref, -1, PREG_SPLIT_NO_EMPTY);
        $invalidAreaReported = false;
        for ($index = 0; $index < count($tokens); $index++) {
            $token = $tokens[$index];
            $region = null;
            if (!XlsxWorkbookSerializerCommon::TryParseMergeReference($token, $region)) {
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException("The conditional formatting sqref '" . $token . "' is invalid.");
                }

                if (!$invalidAreaReported) {
                    $issue = new LoadIssue('CF-L001', DiagnosticSeverity::LossyRecoverable, 'One or more conditional formatting areas were invalid and were dropped during load.', true, true);
                    $issue->SheetName = $sheetName;
                    $issue->CellRef = $token;
                    XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                    $invalidAreaReported = true;
                }
                continue;
            }

            $formatting->Areas[] = CellArea::CreateCellArea($region->FirstRow, $region->FirstColumn, $region->FirstRow + $region->TotalRows - 1, $region->FirstColumn + $region->TotalColumns - 1);
        }

        FormatConditionCollection::SortAreas($formatting->Areas);
        if (count($formatting->Areas) === 0) {
            $issue = new LoadIssue('CF-L003', DiagnosticSeverity::LossyRecoverable, 'A conditional formatting entry without any valid areas was dropped during load.', true, true);
            $issue->SheetName = $sheetName;
            $issue->CellRef = $sqref;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
            return false;
        }

        return true;
    }

    private static function TryLoadCondition(ConditionalFormattingModel $formatting, XElement $ruleElement, StylesheetLoadContext $stylesheet, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $cellRef)
    {
        $type = null;
        if (!self::TryParseRuleType($ruleElement, $options, $diagnostics, $sheetName, $cellRef, $type)) {
            return false;
        }

        $condition = new FormatConditionModel();
        $condition->Type = $type;
        $condition->Operator = self::ParseOperatorType((string)$ruleElement->Attribute('operator'));
        $condition->StopIfTrue = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($ruleElement->Attribute('stopIfTrue'));
        $condition->Priority = self::ParsePriority($ruleElement->Attribute('priority'), $formatting, $diagnostics, $options, $sheetName, $cellRef);
        $condition->Style = self::ResolveDifferentialStyle($ruleElement->Attribute('dxfId'), $stylesheet, $diagnostics, $options, $sheetName, $cellRef);
        self::ApplyDefaults($condition);

        $formulas = $ruleElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'formula');
        switch ($type) {
            case FormatConditionType::Expression:
                $condition->Formula1 = count($formulas) > 0 ? self::NormalizeFormula($formulas[0]->Value) : null;
                break;
            case FormatConditionType::CellValue:
                $condition->Formula1 = count($formulas) > 0 ? self::NormalizeFormula($formulas[0]->Value) : null;
                $condition->Formula2 = count($formulas) > 1 ? self::NormalizeFormula($formulas[1]->Value) : null;
                break;
            case FormatConditionType::ContainsText:
            case FormatConditionType::NotContainsText:
            case FormatConditionType::BeginsWith:
            case FormatConditionType::EndsWith:
                $condition->Formula1 = self::NormalizeText((string)$ruleElement->Attribute('text'));
                break;
            case FormatConditionType::TimePeriod:
                $condition->TimePeriod = self::NormalizeText((string)$ruleElement->Attribute('timePeriod'));
                break;
            case FormatConditionType::Top10:
            case FormatConditionType::Bottom10:
                $condition->Top = $type === FormatConditionType::Top10;
                $condition->Percent = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($ruleElement->Attribute('percent'));
                $condition->Rank = self::ParseNonNegativeIntAttribute($ruleElement->Attribute('rank'));
                break;
            case FormatConditionType::AboveAverage:
            case FormatConditionType::BelowAverage:
                $condition->Above = $type === FormatConditionType::AboveAverage;
                $condition->StandardDeviation = self::ParseNonNegativeIntAttribute($ruleElement->Attribute('stdDev'));
                break;
            case FormatConditionType::ColorScale:
                self::LoadColorScale($condition, $ruleElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'colorScale'));
                break;
            case FormatConditionType::DataBar:
                self::LoadDataBar($condition, $ruleElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'dataBar'));
                break;
            case FormatConditionType::IconSet:
                self::LoadIconSet($condition, $ruleElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'iconSet'));
                break;
            case FormatConditionType::DuplicateValues:
                $condition->Duplicate = true;
                break;
            case FormatConditionType::UniqueValues:
                $condition->Duplicate = false;
                break;
        }

        $formatting->Conditions[] = $condition;
        return true;
    }

    private static function TryParseRuleType(XElement $ruleElement, LoadOptions $options, LoadDiagnostics $diagnostics, $sheetName, $cellRef, &$type)
    {
        $type = FormatConditionType::CellValue;
        $typeText = self::NormalizeToken((string)$ruleElement->Attribute('type'));
        switch ($typeText) {
            case 'cellis':
                $type = FormatConditionType::CellValue;
                return true;
            case 'expression':
                $type = FormatConditionType::Expression;
                return true;
            case 'containstext':
                $type = FormatConditionType::ContainsText;
                return true;
            case 'notcontainstext':
                $type = FormatConditionType::NotContainsText;
                return true;
            case 'beginswith':
                $type = FormatConditionType::BeginsWith;
                return true;
            case 'endswith':
                $type = FormatConditionType::EndsWith;
                return true;
            case 'timeperiod':
                $type = FormatConditionType::TimePeriod;
                return true;
            case 'duplicatevalues':
                $type = FormatConditionType::DuplicateValues;
                return true;
            case 'uniquevalues':
                $type = FormatConditionType::UniqueValues;
                return true;
            case 'top10':
                $type = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($ruleElement->Attribute('bottom')) ? FormatConditionType::Bottom10 : FormatConditionType::Top10;
                return true;
            case 'aboveaverage':
                $type = $ruleElement->Attribute('aboveAverage') !== null && !XlsxWorkbookArchiveHelpers::ParseBoolAttribute($ruleElement->Attribute('aboveAverage'))
                    ? FormatConditionType::BelowAverage
                    : FormatConditionType::AboveAverage;
                return true;
            case 'colorscale':
                $type = FormatConditionType::ColorScale;
                return true;
            case 'databar':
                $type = FormatConditionType::DataBar;
                return true;
            case 'iconset':
                $type = FormatConditionType::IconSet;
                return true;
            default:
                if ($options->StrictMode) {
                    throw new InvalidFileFormatException("The conditional formatting rule type '" . $typeText . "' is not supported.");
                }

                $issue = new LoadIssue('CF-L002', DiagnosticSeverity::LossyRecoverable, 'An unsupported conditional formatting rule was dropped during load.', true, true);
                $issue->SheetName = $sheetName;
                $issue->CellRef = $cellRef;
                XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
                return false;
        }
    }

    private static function ApplyDefaults(FormatConditionModel $condition)
    {
        switch ($condition->Type) {
            case FormatConditionType::DuplicateValues:
                $condition->Duplicate = true;
                break;
            case FormatConditionType::UniqueValues:
                $condition->Duplicate = false;
                break;
            case FormatConditionType::Top10:
                $condition->Top = true;
                if ($condition->Rank === 0) {
                    $condition->Rank = 10;
                }
                break;
            case FormatConditionType::Bottom10:
                $condition->Top = false;
                if ($condition->Rank === 0) {
                    $condition->Rank = 10;
                }
                break;
            case FormatConditionType::AboveAverage:
                $condition->Above = true;
                break;
            case FormatConditionType::BelowAverage:
                $condition->Above = false;
                break;
            case FormatConditionType::ColorScale:
                if ($condition->ColorScaleCount === 0) {
                    $condition->ColorScaleCount = 2;
                }
                break;
            case FormatConditionType::DataBar:
                if (self::IsEmptyColor($condition->BarColor)) {
                    $condition->BarColor = new ColorValue(255, 99, 142, 198);
                }
                break;
            case FormatConditionType::IconSet:
                if ($condition->IconSetType === null || $condition->IconSetType === '') {
                    $condition->IconSetType = '3TrafficLights1';
                }
                break;
        }
    }

    private static function ParsePriority($attribute, ConditionalFormattingModel $formatting, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $cellRef)
    {
        $parsed = XlsxWorkbookArchiveHelpers::ParseIntAttribute($attribute);
        if ($parsed !== null && $parsed > 0) {
            return $parsed;
        }

        if ($attribute !== null) {
            $issue = new LoadIssue('CF-R001', DiagnosticSeverity::Recoverable, 'A conditional formatting priority was invalid and was normalized during load.', true, false);
            $issue->SheetName = $sheetName;
            $issue->CellRef = $cellRef;
            XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        }

        return count($formatting->Conditions) + 1;
    }

    private static function ResolveDifferentialStyle($attribute, StylesheetLoadContext $stylesheet, LoadDiagnostics $diagnostics, LoadOptions $options, $sheetName, $cellRef)
    {
        $dxfId = XlsxWorkbookArchiveHelpers::ParseIntAttribute($attribute);
        if ($dxfId === null) {
            return StyleValue::Default()->Clone();
        }

        if ($dxfId >= 0 && $dxfId < count($stylesheet->DifferentialFormats)) {
            return $stylesheet->DifferentialFormats[$dxfId]->Clone();
        }

        if ($options->StrictMode) {
            throw new InvalidFileFormatException("The conditional formatting dxfId '" . $dxfId . "' is invalid.");
        }

        $issue = new LoadIssue('CF-R002', DiagnosticSeverity::Recoverable, 'A conditional formatting differential style reference was invalid and the rule style was reset to default during load.', true, false);
        $issue->SheetName = $sheetName;
        $issue->CellRef = $cellRef;
        XlsxWorkbookArchiveHelpers::AddIssue($diagnostics, $options, $issue);
        return StyleValue::Default()->Clone();
    }

    private static function LoadColorScale(FormatConditionModel $condition, $colorScaleElement)
    {
        if ($colorScaleElement === null) {
            return;
        }

        $cfvoCount = 0;
        $cfvoElements = $colorScaleElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'cfvo');
        foreach ($cfvoElements as $cfvoElement) {
            $cfvoCount++;
        }
        $condition->ColorScaleCount = $cfvoCount >= 3 ? 3 : 2;
        $colors = $colorScaleElement->Elements(XlsxWorkbookSerializerCommon::MainNs . 'color');
        if (count($colors) > 0) {
            $condition->MinColor = self::ReadRgbColor($colors[0]);
        }

        if ($condition->ColorScaleCount === 3 && count($colors) > 2) {
            $condition->MidColor = self::ReadRgbColor($colors[1]);
            $condition->MaxColor = self::ReadRgbColor($colors[2]);
            return;
        }

        if (count($colors) > 1) {
            $condition->MaxColor = self::ReadRgbColor($colors[1]);
        }
    }

    private static function LoadDataBar(FormatConditionModel $condition, $dataBarElement)
    {
        if ($dataBarElement === null) {
            return;
        }

        $color = $dataBarElement->Element(XlsxWorkbookSerializerCommon::MainNs . 'color');
        if ($color !== null) {
            $condition->BarColor = self::ReadRgbColor($color);
        }
    }

    private static function LoadIconSet(FormatConditionModel $condition, $iconSetElement)
    {
        if ($iconSetElement === null) {
            return;
        }

        $condition->IconSetType = self::NormalizeText((string)$iconSetElement->Attribute('iconSet')) ?? '3TrafficLights1';
        $condition->ReverseIcons = XlsxWorkbookArchiveHelpers::ParseBoolAttribute($iconSetElement->Attribute('reverse'));
        $condition->ShowIconOnly = $iconSetElement->Attribute('showValue') !== null && !XlsxWorkbookArchiveHelpers::ParseBoolAttribute($iconSetElement->Attribute('showValue'));
    }

    private static function BuildColorScaleElement(FormatConditionModel $condition)
    {
        $element = new XElement(
            XlsxWorkbookSerializerCommon::MainNs . 'colorScale',
            new XElement(XlsxWorkbookSerializerCommon::MainNs . 'cfvo', new XAttribute('type', 'min'))
        );

        if ($condition->ColorScaleCount === 3) {
            $element->Add(new XElement(
                XlsxWorkbookSerializerCommon::MainNs . 'cfvo',
                new XAttribute('type', 'percentile'),
                new XAttribute('val', 50)
            ));
        }

        $element->Add(new XElement(XlsxWorkbookSerializerCommon::MainNs . 'cfvo', new XAttribute('type', 'max')));
        $element->Add(self::BuildColorElement($condition->MinColor, new ColorValue(255, 248, 105, 107)));
        if ($condition->ColorScaleCount === 3) {
            $element->Add(self::BuildColorElement($condition->MidColor, new ColorValue(255, 255, 235, 132)));
        }

        $element->Add(self::BuildColorElement($condition->MaxColor, new ColorValue(255, 99, 190, 123)));
        return $element;
    }

    private static function BuildDataBarElement(FormatConditionModel $condition)
    {
        $element = new XElement(
            XlsxWorkbookSerializerCommon::MainNs . 'dataBar',
            new XElement(XlsxWorkbookSerializerCommon::MainNs . 'cfvo', new XAttribute('type', 'min')),
            new XElement(XlsxWorkbookSerializerCommon::MainNs . 'cfvo', new XAttribute('type', 'max')),
            self::BuildColorElement($condition->BarColor, new ColorValue(255, 99, 142, 198))
        );
        return $element;
    }

    private static function BuildIconSetElement(FormatConditionModel $condition)
    {
        $iconSetType = $condition->IconSetType === null || $condition->IconSetType === '' ? '3TrafficLights1' : $condition->IconSetType;
        $element = new XElement(
            XlsxWorkbookSerializerCommon::MainNs . 'iconSet',
            new XAttribute('iconSet', $iconSetType)
        );
        if ($condition->ReverseIcons) {
            $element->SetAttributeValue('reverse', 1);
        }

        if ($condition->ShowIconOnly) {
            $element->SetAttributeValue('showValue', 0);
        }

        $iconCount = self::GetIconCount($iconSetType);
        for ($index = 0; $index < $iconCount; $index++) {
            $element->Add(new XElement(
                XlsxWorkbookSerializerCommon::MainNs . 'cfvo',
                new XAttribute('type', 'percent'),
                new XAttribute('val', intdiv(100 * $index, $iconCount))
            ));
        }

        return $element;
    }

    private static function BuildColorElement($actual, $fallback)
    {
        $color = self::IsEmptyColor($actual) ? $fallback : $actual;
        return new XElement(XlsxWorkbookSerializerCommon::MainNs . 'color', new XAttribute('rgb', self::ToArgbHex($color)));
    }

    private static function AddFormulaElement(XElement $parent, $formula)
    {
        if (!($formula === null || $formula === '')) {
            $parent->Add(new XElement(XlsxWorkbookSerializerCommon::MainNs . 'formula', $formula));
        }
    }

    private static function BuildTextRuleFormula($type, $text, $firstCell)
    {
        $escapedText = str_replace('"', '""', $text);
        switch ($type) {
            case FormatConditionType::ContainsText:
                return 'NOT(ISERROR(SEARCH("' . $escapedText . '",' . $firstCell . ')))';
            case FormatConditionType::NotContainsText:
                return 'ISERROR(SEARCH("' . $escapedText . '",' . $firstCell . '))';
            case FormatConditionType::BeginsWith:
                return 'LEFT(' . $firstCell . ',LEN("' . $escapedText . '"))="' . $escapedText . '"';
            case FormatConditionType::EndsWith:
                return 'RIGHT(' . $firstCell . ',LEN("' . $escapedText . '"))="' . $escapedText . '"';
            default:
                return '';
        }
    }

    private static function GetAnchorCellReference(ConditionalFormattingModel $formatting)
    {
        if (count($formatting->Areas) === 0) {
            return 'A1';
        }

        $area = $formatting->Areas[0];
        return (new CellAddress($area->FirstRow, $area->FirstColumn))->ToString();
    }

    private static function GetOrderedConditionalFormattings($collections)
    {
        $ordered = array();
        for ($index = 0; $index < count($collections); $index++) {
            $collection = $collections[$index];
            if (count($collection->Areas) === 0 || count($collection->Conditions) === 0) {
                continue;
            }

            FormatConditionCollection::SortAreas($collection->Areas);
            $ordered[] = $collection;
        }

        usort($ordered, array(__CLASS__, 'CompareConditionalFormattings'));
        return $ordered;
    }

    private static function CompareConditionalFormattings(ConditionalFormattingModel $left, ConditionalFormattingModel $right)
    {
        return FormatConditionCollection::CompareAreas($left->Areas[0], $right->Areas[0]);
    }

    private static function BuildSqref($areas)
    {
        $references = array();
        for ($index = 0; $index < count($areas); $index++) {
            $references[] = self::ToAreaReference($areas[$index]);
        }

        return implode(' ', $references);
    }

    private static function ToAreaReference($area)
    {
        $first = (new CellAddress($area->FirstRow, $area->FirstColumn))->ToString();
        if ($area->TotalRows === 1 && $area->TotalColumns === 1) {
            return $first;
        }

        $last = (new CellAddress($area->FirstRow + $area->TotalRows - 1, $area->FirstColumn + $area->TotalColumns - 1))->ToString();
        return $first . ':' . $last;
    }

    private static function ToRuleTypeName(FormatConditionModel $condition)
    {
        switch ($condition->Type) {
            case FormatConditionType::Expression:
                return 'expression';
            case FormatConditionType::ContainsText:
                return 'containsText';
            case FormatConditionType::NotContainsText:
                return 'notContainsText';
            case FormatConditionType::BeginsWith:
                return 'beginsWith';
            case FormatConditionType::EndsWith:
                return 'endsWith';
            case FormatConditionType::TimePeriod:
                return 'timePeriod';
            case FormatConditionType::DuplicateValues:
                return 'duplicateValues';
            case FormatConditionType::UniqueValues:
                return 'uniqueValues';
            case FormatConditionType::Top10:
            case FormatConditionType::Bottom10:
                return 'top10';
            case FormatConditionType::AboveAverage:
            case FormatConditionType::BelowAverage:
                return 'aboveAverage';
            case FormatConditionType::ColorScale:
                return 'colorScale';
            case FormatConditionType::DataBar:
                return 'dataBar';
            case FormatConditionType::IconSet:
                return 'iconSet';
            default:
                return 'cellIs';
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

    private static function ParseOperatorType($text)
    {
        switch (self::NormalizeToken($text)) {
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
                return OperatorType::None;
        }
    }

    private static function ParseNonNegativeIntAttribute($attribute)
    {
        $parsed = XlsxWorkbookArchiveHelpers::ParseIntAttribute($attribute);
        if ($parsed !== null && $parsed >= 0) {
            return $parsed;
        }

        return 0;
    }

    private static function GetIconCount($iconSetType)
    {
        if ($iconSetType === null || $iconSetType === '') {
            return 3;
        }

        if ($iconSetType[0] === '4') {
            return 4;
        }

        if ($iconSetType[0] === '5') {
            return 5;
        }

        return 3;
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

    private static function NormalizeFormula($value)
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
        if ($value === null) {
            return null;
        }

        $trimmed = trim($value);
        if (strlen($trimmed) === 0) {
            return null;
        }

        return $trimmed;
    }

    private static function ReadRgbColor(XElement $element)
    {
        $rgb = self::NormalizeText((string)$element->Attribute('rgb'));
        if ($rgb === null || $rgb === '' || strlen($rgb) !== 8) {
            return new ColorValue();
        }

        return new ColorValue(
            hexdec(substr($rgb, 0, 2)),
            hexdec(substr($rgb, 2, 2)),
            hexdec(substr($rgb, 4, 2)),
            hexdec(substr($rgb, 6, 2))
        );
    }

    private static function ToArgbHex(ColorValue $color)
    {
        return sprintf('%02X%02X%02X%02X', $color->A, $color->R, $color->G, $color->B);
    }

    private static function IsEmptyColor($color)
    {
        if ($color === null) {
            return true;
        }

        return $color->ThemeIndex === null && $color->Indexed === null
            && $color->A === 0 && $color->R === 0 && $color->G === 0 && $color->B === 0;
    }
}
