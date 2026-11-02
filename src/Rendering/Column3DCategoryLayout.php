<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\DateSerialConverter;

final class Column3DCategoryLayout
{
    public $SlotCount;

    public $Centers = array();

    public $Labels = array();

    public function __construct($chart, $count, $dateSystem, $culture)
    {
        $this->SlotCount = $count;
        $min = 0.0;
        $max = 0.0;
        if ($chart->HasDateCategoryAxis
            && $chart->CategoryBaseTimeUnit === 'months'
            && ChartDateAxisMath::TryGetDateAxisRange($chart, $min, $max)) {
            $first = DateSerialConverter::FromSerial($min, $dateSystem);
            $last = DateSerialConverter::FromSerial($max, $dateSystem);
            $this->SlotCount = ((int)$last->format('Y') - (int)$first->format('Y')) * 12 + (int)$last->format('n') - (int)$first->format('n') + 1;
            for ($i = 0; $i < $count; $i++) {
                $slot = $i;
                if ($i < count($chart->CategoryValues) && $chart->CategoryValues[$i] !== null) {
                    $date = DateSerialConverter::FromSerial($chart->CategoryValues[$i], $dateSystem);
                    $slot = ((int)$date->format('Y') - (int)$first->format('Y')) * 12 + (int)$date->format('n') - (int)$first->format('n');
                }
                $this->Centers[] = ($slot + 0.5) / $this->SlotCount;
            }
            $start = ChartDateAxisMath::StartOfMonth($min, $dateSystem);
            for ($i = 0; $i < $this->SlotCount; $i++) {
                $this->Labels[] = ChartXmlParser::FormatLabel(
                    ChartDateAxisMath::AddMonthsSerial($start, $i, $dateSystem),
                    $chart->CategoryFormatCode, $dateSystem, $culture);
            }
            return;
        }
        for ($i = 0; $i < $count; $i++) {
            $this->Centers[] = ($i + 0.5) / $count;
            $this->Labels[] = $i < count($chart->Categories) ? $chart->Categories[$i] : '';
        }
    }
}
