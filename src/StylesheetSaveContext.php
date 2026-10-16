<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CellRecord;
use Aspose\Cells_FOSS\Core\FormatConditionModel;
use Aspose\Cells_FOSS\Core\StyleValue;
use Aspose\Cells_FOSS\Internal\Xml\XDocument;

final class StylesheetSaveContext
{
    private $_styleIndices;
    private $_differentialStyleIndices;
    public $Document;
    public $DifferentialFormatCount;
    public $HasStyles;

    public function __construct(XDocument $document, $styleIndices, $differentialStyleIndices, $differentialFormatCount, $hasStyles)
    {
        $this->Document = $document;
        $this->_styleIndices = $styleIndices;
        $this->_differentialStyleIndices = $differentialStyleIndices;
        $this->DifferentialFormatCount = $differentialFormatCount;
        $this->HasStyles = $hasStyles;
    }

    public function GetStyleIndex(CellRecord $record)
    {
        $style = XlsxWorkbookStyles::GetStyleForSerialization($record);
        $key = XlsxWorkbookStyles::GetStyleKey($style);
        return array_key_exists($key, $this->_styleIndices) ? $this->_styleIndices[$key] : 0;
    }

    public function GetDifferentialStyleIndex(FormatConditionModel $condition)
    {
        if (XlsxWorkbookStyles::StylesEqual($condition->Style, StyleValue::Default())) {
            return null;
        }

        $differentialIndex = 0;
        $key = XlsxWorkbookStyles::GetStyleKey($condition->Style);
        if (array_key_exists($key, $this->_differentialStyleIndices)) {
            $differentialIndex = $this->_differentialStyleIndices[$key];
            return $differentialIndex;
        }
        return null;
    }
}