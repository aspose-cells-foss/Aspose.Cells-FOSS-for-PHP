<?php

namespace Aspose\Cells_FOSS\Rendering;

final class PdfTrueTypePrototypeBackend implements PdfTextFontBackend
{
    private $fallback;
    private $trueTypeUsages = array();

    public function __construct()
    {
        $this->fallback = new PdfType3TextFontBackend();
    }

    public function Initialize($runs)
    {
        $fallbackRuns = array();
        if ($runs !== null) {
            for ($i = 0; $i < count($runs); $i++) {
                $run = $runs[$i];
                if ($this->IsEligibleForPrototype($run)) {
                    $usage = $this->UsageFor($run->Typeface);
                    $usage->AddText($run->Text);
                } else {
                    $fallbackRuns[] = $run;
                }
            }
        }

        $this->fallback->Initialize($fallbackRuns);
    }

    public function BuildObjects(&$nextObjectId)
    {
        $objects = array();
        $builder = new PdfTrueTypePrototypeFontBuilder();
        $trueTypeObjects = $builder->BuildObjects($this->trueTypeUsages, $nextObjectId);
        foreach ($trueTypeObjects as $objectId => $objectBytes) {
            $objects[$objectId] = $objectBytes;
        }

        $fallbackObjects = $this->fallback->BuildObjects($nextObjectId);
        foreach ($fallbackObjects as $objectId => $objectBytes) {
            $objects[$objectId] = $objectBytes;
        }

        return $objects;
    }

    public function BuildFontResourceText()
    {
        $builder = '/Font <<';
        for ($i = 0; $i < count($this->trueTypeUsages); $i++) {
            $usage = $this->trueTypeUsages[$i];
            $builder .= ' /' . $usage->ResourceName . ' ' . $usage->FontObjectId . " 0 R";
        }

        $fallbackText = $this->fallback->BuildFontResourceText();
        $start = strpos($fallbackText, '<<');
        $end = strrpos($fallbackText, '>>');
        if ($start !== false && $end > $start + 2) {
            $builder .= substr($fallbackText, $start + 2, $end - $start - 2);
        }

        $builder .= ' >>';
        return $builder;
    }

    public function EncodeRun($run)
    {
        if (!$this->IsEligibleForPrototype($run)) {
            return $this->fallback->EncodeRun($run);
        }

        $usage = $this->UsageFor($run->Typeface);
        $segment = new PdfEncodedTextSegment();
        $segment->ResourceName = $usage->ResourceName;
        $segment->EncodedBytes = $usage->EncodeText($run->Text);
        $segment->StartXPt = $run->XPt;
        $segment->RequiresTopDownTextFlip = true;
        return array($segment);
    }

    private function IsEligibleForPrototype($run)
    {
        if ($run === null || $run->Typeface === null || $run->Text === null || $run->Text === '') {
            return false;
        }

        if (!$run->UseTopDownCoordinates) {
            return false;
        }

        $len = strlen($run->Text);
        for ($i = 0; $i < $len; $i++) {
            $ch = ord($run->Text[$i]);
            if ($ch < 32 || $ch > 126) {
                return false;
            }
        }

        self::ValidateRequiredTables($run->Typeface);
        return true;
    }

    private function UsageFor($typeface)
    {
        for ($i = 0; $i < count($this->trueTypeUsages); $i++) {
            if ($this->trueTypeUsages[$i]->Typeface() === $typeface) {
                return $this->trueTypeUsages[$i];
            }
        }

        $usage = new PdfTrueTypePrototypeUsage($typeface);
        $this->trueTypeUsages[] = $usage;
        return $usage;
    }

    private static function ValidateRequiredTables($typeface)
    {
        if ($typeface === null) {
            throw new \RuntimeException('TrueType prototype requires a resolved typeface.');
        }

        $requiredTags = array(
            self::Tag('c', 'm', 'a', 'p'),
            self::Tag('h', 'e', 'a', 'd'),
            self::Tag('h', 'h', 'e', 'a'),
            self::Tag('h', 'm', 't', 'x'),
            self::Tag('m', 'a', 'x', 'p'),
            self::Tag('l', 'o', 'c', 'a'),
            self::Tag('g', 'l', 'y', 'f'),
        );

        $availableTags = $typeface->GetTableTags();
        if ($availableTags === null || count($availableTags) === 0) {
            throw new \RuntimeException('TrueType prototype could not enumerate font tables for ' . $typeface->FamilyName . '.');
        }

        for ($i = 0; $i < count($requiredTags); $i++) {
            $required = $requiredTags[$i];
            if (!in_array($required, $availableTags, true)) {
                throw new \RuntimeException('TrueType prototype missing required font table ' . self::TagName($required) . ' for ' . $typeface->FamilyName . '.');
            }

            $tableData = $typeface->GetTableData($required);
            if ($tableData === '' || strlen($tableData) === 0) {
                throw new \RuntimeException('TrueType prototype could not read font table ' . self::TagName($required) . ' for ' . $typeface->FamilyName . '.');
            }
        }
    }

    private static function Tag($a, $b, $c, $d)
    {
        return (ord($a) << 24) | (ord($b) << 16) | (ord($c) << 8) | ord($d);
    }

    private static function TagName($tag)
    {
        return chr(($tag >> 24) & 0xFF) . chr(($tag >> 16) & 0xFF)
            . chr(($tag >> 8) & 0xFF) . chr($tag & 0xFF);
    }

    public function Dispose()
    {
        $this->fallback->Dispose();
        for ($i = 0; $i < count($this->trueTypeUsages); $i++) {
            $this->trueTypeUsages[$i]->Dispose();
        }

        $this->trueTypeUsages = array();
    }
}
