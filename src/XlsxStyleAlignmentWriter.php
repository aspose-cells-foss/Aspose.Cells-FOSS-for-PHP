<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\AlignmentValue;
use Aspose\Cells_FOSS\Internal\Xml\XElement;

final class XlsxStyleAlignmentWriter
{
    private function __construct()
    {
    }

    public static function Build(AlignmentValue $alignment)
    {
        $element = new XElement(XlsxWorkbookSerializerCommon::MainNs . 'alignment');
        $hasValue = false;
        $horizontal = XlsxWorkbookStylesValueHelpers::GetHorizontalAlignmentName($alignment->Horizontal);
        if ($horizontal !== null && $horizontal !== '') {
            $element->SetAttributeValue('horizontal', $horizontal);
            $hasValue = true;
        }

        $vertical = XlsxWorkbookStylesValueHelpers::GetVerticalAlignmentName($alignment->Vertical);
        if ($vertical !== null && $vertical !== '') {
            $element->SetAttributeValue('vertical', $vertical);
            $hasValue = true;
        }

        if ($alignment->WrapText) {
            $element->SetAttributeValue('wrapText', 1);
            $hasValue = true;
        }

        if ($alignment->IndentLevel > 0) {
            $element->SetAttributeValue('indent', $alignment->IndentLevel);
            $hasValue = true;
        }

        if ($alignment->TextRotation !== 0) {
            $element->SetAttributeValue('textRotation', $alignment->TextRotation);
            $hasValue = true;
        }

        if ($alignment->ShrinkToFit) {
            $element->SetAttributeValue('shrinkToFit', 1);
            $hasValue = true;
        }

        if ($alignment->ReadingOrder !== 0) {
            $element->SetAttributeValue('readingOrder', $alignment->ReadingOrder);
            $hasValue = true;
        }

        if ($alignment->RelativeIndent !== 0) {
            $element->SetAttributeValue('relativeIndent', $alignment->RelativeIndent);
            $hasValue = true;
        }

        return $hasValue ? $element : null;
    }
}
