<?php

namespace Aspose\Cells_FOSS\Samples\PdfConversion;

require_once __DIR__ . '/../../autoload.php';

use Aspose\Cells_FOSS\PaperSizeType;
use Aspose\Cells_FOSS\Workbook;

final class PrintOverrides
{
    private static $MarginNames = array(
        '--margin-left-cm',
        '--margin-right-cm',
        '--margin-top-cm',
        '--margin-bottom-cm',
    );

    private $paper = null;
    private $scale = null;
    private $margins = array(null, null, null, null);

    public function Read($args, &$index)
    {
        $name = $args[$index];
        $marginIndex = array_search($name, self::$MarginNames, true);
        if ($name !== '--paper-size' && $name !== '--scale' && $marginIndex === false) {
            return false;
        }

        if ($index + 1 >= count($args)) {
            throw new \InvalidArgumentException($name . ' requires a value.');
        }

        $value = $args[++$index];
        if ($name === '--paper-size') {
            $paper = self::ParsePaperSize($value);
            if ($paper === null) {
                throw new \InvalidArgumentException('Unsupported paper size: ' . $value);
            }

            $this->paper = $paper;
        } elseif ($name === '--scale') {
            $scale = self::ParseInteger($value);
            if ($scale === null || $scale < 10 || $scale > 400) {
                throw new \InvalidArgumentException('--scale must be between 10 and 400.');
            }

            $this->scale = $scale;
        } else {
            $margin = is_numeric(trim($value)) ? (float)$value : NAN;
            if (is_nan($margin) || is_infinite($margin) || $margin < 0) {
                throw new \InvalidArgumentException($name . ' requires a non-negative finite number.');
            }

            $this->margins[$marginIndex] = $margin;
        }

        return true;
    }

    public function Apply(Workbook $workbook)
    {
        for ($index = 0; $index < $workbook->Worksheets->Count; $index++) {
            $setup = $workbook->Worksheets[$index]->PageSetup;
            if ($this->paper !== null) {
                $setup->PaperSize = $this->paper;
            }
            if ($this->scale !== null) {
                $setup->Scale = $this->scale;
            }
            if ($this->margins[0] !== null) {
                $setup->LeftMargin = $this->margins[0];
            }
            if ($this->margins[1] !== null) {
                $setup->RightMargin = $this->margins[1];
            }
            if ($this->margins[2] !== null) {
                $setup->TopMargin = $this->margins[2];
            }
            if ($this->margins[3] !== null) {
                $setup->BottomMargin = $this->margins[3];
            }

            echo 'Print settings: paper=' . self::PaperSizeName($setup->PaperSize)
                . ', scale=' . $setup->Scale
                . ', margins (cm)=' . self::FormatDouble($setup->LeftMargin) . ','
                . self::FormatDouble($setup->RightMargin) . ','
                . self::FormatDouble($setup->TopMargin) . ','
                . self::FormatDouble($setup->BottomMargin) . PHP_EOL;
        }
    }

    private static function ParseInteger($value)
    {
        $trimmed = trim($value);
        if (preg_match('/^[+-]?[0-9]+$/', $trimmed) !== 1) {
            return null;
        }

        return (int)$trimmed;
    }

    private static function ParsePaperSize($value)
    {
        $expectedName = 'Paper' . $value;
        foreach (self::PaperSizeNames() as $name => $paperSize) {
            if (strcasecmp($name, $expectedName) === 0) {
                return $paperSize;
            }
        }

        return null;
    }

    private static function PaperSizeName($value)
    {
        foreach (self::PaperSizeNames() as $name => $paperSize) {
            if ($paperSize === $value) {
                return $name;
            }
        }

        return (string)$value;
    }

    private static function PaperSizeNames()
    {
        return array(
            'Default' => PaperSizeType::Default,
            'PaperLetter' => PaperSizeType::PaperLetter,
            'PaperTabloid' => PaperSizeType::PaperTabloid,
            'PaperLegal' => PaperSizeType::PaperLegal,
            'PaperStatement' => PaperSizeType::PaperStatement,
            'PaperExecutive' => PaperSizeType::PaperExecutive,
            'PaperA3' => PaperSizeType::PaperA3,
            'PaperA4' => PaperSizeType::PaperA4,
            'PaperA5' => PaperSizeType::PaperA5,
            'PaperB4' => PaperSizeType::PaperB4,
            'PaperB5' => PaperSizeType::PaperB5,
        );
    }

    private static function FormatDouble($value)
    {
        return (string)(float)$value;
    }
}
