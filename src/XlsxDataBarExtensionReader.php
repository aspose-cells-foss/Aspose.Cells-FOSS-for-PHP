<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\FormatConditionModel;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

final class XlsxDataBarExtensionReader
{
    const MainNs = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';
    const ExtensionNs = '{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}';

    private function __construct()
    {
    }

    public static function Apply(XElement $worksheetRoot, XElement $ruleElement, FormatConditionModel $condition)
    {
        if ($condition->Type !== FormatConditionType::DataBar) {
            return;
        }

        $ruleId = null;
        foreach ($ruleElement->DescendantsAndSelf() as $idElement) {
            if ($idElement->Name !== self::ExtensionNs . 'id') {
                continue;
            }

            $ruleId = $idElement->Value;
            break;
        }

        $extensions = $worksheetRoot->Element(self::MainNs . 'extLst');
        if ($ruleId === null || $ruleId === '' || $extensions === null) {
            return;
        }

        foreach ($extensions->DescendantsAndSelf() as $extendedRule) {
            if ($extendedRule->Name !== self::ExtensionNs . 'cfRule') {
                continue;
            }

            if (strcasecmp((string) $extendedRule->Attribute('id'), $ruleId) !== 0) {
                continue;
            }

            $dataBar = $extendedRule->Element(self::ExtensionNs . 'dataBar');
            if ($dataBar === null) {
                return;
            }

            $gradient = (string) $dataBar->Attribute('gradient');
            if ($gradient !== null) {
                $condition->DataBarGradient = $gradient !== '0'
                    && strcasecmp($gradient, 'false') !== 0;
            }

            $minLength = 0.0;
            $maxLength = 0.0;
            if (self::TryReadLength($dataBar->Attribute('minLength'), $minLength)) {
                $condition->DataBarMinLength = $minLength;
            }

            if (self::TryReadLength($dataBar->Attribute('maxLength'), $maxLength)) {
                $condition->DataBarMaxLength = $maxLength;
            }

            if ($condition->DataBarMinLength > $condition->DataBarMaxLength) {
                $condition->DataBarMinLength = 0.10;
                $condition->DataBarMaxLength = 0.90;
            }

            foreach ($dataBar->Elements(self::ExtensionNs . 'cfvo') as $threshold) {
                $condition->DataBarAutoMinimum = strcasecmp(
                    (string) $threshold->Attribute('type'), 'autoMin') === 0;
                break;
            }

            return;
        }
    }

    private static function TryReadLength($attribute, &$length)
    {
        $length = 0.0;
        if ($attribute === null) {
            return false;
        }

        $value = trim((string) $attribute);
        if (!is_numeric($value)) {
            return false;
        }

        $percent = (float) $value;
        if ($percent < 0.0 || $percent > 100.0) {
            return false;
        }

        $length = $percent / 100.0;
        return true;
    }
}
