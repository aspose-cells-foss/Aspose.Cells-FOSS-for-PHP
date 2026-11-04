<?php

namespace Aspose\Cells_FOSS\Rendering;

final class PdfTextObjectWriter
{
    private $useTrueTypeSubsetPrototype;

    public function __construct($useTrueTypeSubsetPrototype = false)
    {
        $this->useTrueTypeSubsetPrototype = (bool)$useTrueTypeSubsetPrototype;
    }

    /**
     * @param string $originalPdf
     * @param PdfTextDocumentSession $session
     * @return string
     */
    public function Rewrite($originalPdf, $session)
    {
        if ($originalPdf === null || $session === null || !$session->HasRuns()) {
            return $originalPdf;
        }

        $backend = $this->CreateFontBackend();
        try {
            $runs = &$session->Runs();
            $backend->Initialize($runs);

            $table = PdfObjectTable::Parse($originalPdf);
            $nextObjectId = $table->MaxObjectId() + 1;
            $fontObjects = $backend->BuildObjects($nextObjectId);
            foreach ($fontObjects as $objectId => $objectBytes) {
                $table->Objects[$objectId] = $objectBytes;
            }

            $fontResourceText = $backend->BuildFontResourceText();
            for ($pageIndex = 0; $pageIndex < count($table->PageObjectIds); $pageIndex++) {
                $pageRuns = self::RunsForPage($runs, $pageIndex + 1);
                if (count($pageRuns) === 0) {
                    continue;
                }

                $pageObjectId = $table->PageObjectIds[$pageIndex];
                $pageObjectText = $table->Objects[$pageObjectId];
                $pageHeight = self::ParsePageHeight($pageObjectText);
                $prologueObjectId = $nextObjectId++;
                $contentObjectId = $nextObjectId++;
                $table->Objects[$prologueObjectId] = self::BuildStreamObject($prologueObjectId, "q\n");
                $table->Objects[$contentObjectId] = self::BuildStreamObject(
                    $contentObjectId,
                    self::BuildPageTextContent($pageRuns, $backend, $pageHeight)
                );

                $pageObjectText = self::AddFontResources($pageObjectText, $fontResourceText);
                $pageObjectText = self::AppendContentReferences($pageObjectText, $prologueObjectId, $contentObjectId);
                $table->Objects[$pageObjectId] = $pageObjectText;
            }

            return self::RebuildPdf($table);
        } finally {
            $backend->Dispose();
        }
    }

    private function CreateFontBackend()
    {
        if ($this->useTrueTypeSubsetPrototype) {
            return new PdfTrueTypePrototypeBackend();
        }

        return new PdfType3TextFontBackend();
    }

    private static function RunsForPage($runs, $pageNumber)
    {
        $pageRuns = array();
        for ($i = 0; $i < count($runs); $i++) {
            if ($runs[$i]->PageNumber === $pageNumber) {
                $pageRuns[] = $runs[$i];
            }
        }

        return $pageRuns;
    }

    private static function AddFontResources($pageObjectText, $fontResources)
    {
        if (strpos($pageObjectText, '/Font <<') !== false) {
            return $pageObjectText;
        }

        $marker = '/Resources <<';
        $index = strpos($pageObjectText, $marker);
        if ($index === false) {
            return $pageObjectText;
        }

        return substr_replace($pageObjectText, $fontResources, $index + strlen($marker), 0);
    }

    private static function AppendContentReferences($pageObjectText, $prologueObjectId, $contentObjectId)
    {
        $marker = '/Contents ';
        $index = strpos($pageObjectText, $marker);
        if ($index === false) {
            return $pageObjectText;
        }

        $start = $index + strlen($marker);
        if (isset($pageObjectText[$start]) && $pageObjectText[$start] === '[') {
            $endArray = strpos($pageObjectText, ']', $start);
            if ($endArray === false) {
                return $pageObjectText;
            }

            return substr_replace(
                $pageObjectText,
                ' ' . $prologueObjectId . ' 0 R ' . $contentObjectId . ' 0 R',
                $endArray,
                0
            );
        }

        $end = strpos($pageObjectText, "\n", $start);
        if ($end === false) {
            $end = strlen($pageObjectText);
        }

        $existing = trim(substr($pageObjectText, $start, $end - $start));
        $replacement = '[' . $prologueObjectId . ' 0 R ' . $existing . ' ' . $contentObjectId . ' 0 R]';
        return substr($pageObjectText, 0, $start) . $replacement . substr($pageObjectText, $end);
    }

    private static function BuildPageTextContent($runs, $backend, $pageHeight)
    {
        $builder = "Q\n";

        $start = 0;
        $runCount = count($runs);
        while ($start < $runCount) {
            $end = $start;
            while ($end + 1 < $runCount && self::CanBatchRuns($runs[$end], $runs[$end + 1])) {
                $end++;
            }

            $builder .= self::AppendRunBatch($runs, $start, $end, $backend);
            $start = $end + 1;
        }

        return $builder;
    }

    private static function AppendRunBatch($runs, $start, $end, $backend)
    {
        $f = __NAMESPACE__ . '\PdfNumbers::threeDecimals';
        $firstRun = $runs[$start];
        $builder = "q\n";
        if ($firstRun->UseTopDownCoordinates) {
            $builder .= '1 0 0 -1 0 ' . $f($firstRun->PageHeightPt) . " cm\n";
            $builder .= $f($firstRun->ClipLeftPt) . ' ' . $f($firstRun->ClipTopPt) . ' '
                . $f($firstRun->ClipWidthPt) . ' ' . $f($firstRun->ClipHeightPt) . " re W n\n";
        } else {
            $builder .= $f($firstRun->ClipLeftPt) . ' ' . $f($firstRun->ClipBottomPt) . ' '
                . $f($firstRun->ClipWidthPt) . ' ' . $f($firstRun->ClipHeightPt) . " re W n\n";
        }

        $builder .= "BT\n";
        if ($firstRun->Invisible) {
            $builder .= "3 Tr\n";
        }

        $builder .= self::ColorCommand($firstRun) . "\n";

        $activeResourceName = null;
        $activeFontSize = -1.0;

        for ($i = $start; $i <= $end; $i++) {
            $run = $runs[$i];
            $segments = $backend->EncodeRun($run);
            for ($segmentIndex = 0; $segmentIndex < count($segments); $segmentIndex++) {
                $segment = $segments[$segmentIndex];
                if ($segment === null || $segment->EncodedBytes === '' || strlen($segment->EncodedBytes) === 0) {
                    continue;
                }

                if ($activeResourceName !== $segment->ResourceName
                    || abs($activeFontSize - $run->FontSizePt) > 0.001) {
                    $builder .= '/' . $segment->ResourceName . ' ' . $f($run->FontSizePt) . " Tf\n";
                    $activeResourceName = $segment->ResourceName;
                    $activeFontSize = $run->FontSizePt;
                }

                $hasRotation = abs($run->RotationDeg) > 0.01;
                if ($hasRotation) {
                    $radians = $run->RotationDeg * M_PI / 180.0;
                    $cos = cos($radians);
                    $sin = sin($radians);
                    $e = $run->TransformOriginXPt + $cos * $segment->StartXPt - $sin * $run->BaselineYPt;
                    $fval = $run->TransformOriginYPt + $sin * $segment->StartXPt + $cos * $run->BaselineYPt;
                    $builder .= $f($cos) . ' ' . $f($sin) . ' ';
                    if ($run->UseTopDownCoordinates && $segment->RequiresTopDownTextFlip) {
                        $builder .= $f($sin);
                    } else {
                        $builder .= $f(-$sin);
                    }

                    $builder .= ' ';
                    if ($run->UseTopDownCoordinates && $segment->RequiresTopDownTextFlip) {
                        $builder .= $f(-$cos);
                    } else {
                        $builder .= $f($cos);
                    }

                    $builder .= ' ' . $f($e) . ' ' . $f($fval) . " Tm\n<";
                } else {
                    if ($run->UseTopDownCoordinates && $segment->RequiresTopDownTextFlip) {
                        $builder .= '1 0 0 -1 ';
                    } else {
                        $builder .= '1 0 0 1 ';
                    }

                    $builder .= $f($segment->StartXPt) . ' ' . $f($run->BaselineYPt) . " Tm\n<";
                }

                $byteLen = strlen($segment->EncodedBytes);
                for ($b = 0; $b < $byteLen; $b++) {
                    $builder .= sprintf('%02X', ord($segment->EncodedBytes[$b]));
                }

                $builder .= "> Tj\n";
            }
        }

        $builder .= "ET\nQ\n";
        return $builder;
    }

    private static function CanBatchRuns($left, $right)
    {
        if ($left === null || $right === null) {
            return false;
        }

        if ($left->PageNumber !== $right->PageNumber) {
            return false;
        }

        if ($left->UseTopDownCoordinates !== $right->UseTopDownCoordinates) {
            return false;
        }

        if ($left->Invisible !== $right->Invisible) {
            return false;
        }

        if (abs($left->PageHeightPt - $right->PageHeightPt) > 0.01) {
            return false;
        }

        if (!self::SameFloat($left->ClipLeftPt, $right->ClipLeftPt)
            || !self::SameFloat($left->ClipTopPt, $right->ClipTopPt)
            || !self::SameFloat($left->ClipBottomPt, $right->ClipBottomPt)
            || !self::SameFloat($left->ClipWidthPt, $right->ClipWidthPt)
            || !self::SameFloat($left->ClipHeightPt, $right->ClipHeightPt)) {
            return false;
        }

        if ($left->Color->Red !== $right->Color->Red
            || $left->Color->Green !== $right->Color->Green
            || $left->Color->Blue !== $right->Color->Blue) {
            return false;
        }

        return true;
    }

    private static function SameFloat($left, $right)
    {
        return abs($left - $right) <= 0.01;
    }

    private static function ParsePageHeight($pageObjectText)
    {
        $marker = '/MediaBox [';
        $index = strpos($pageObjectText, $marker);
        if ($index === false) {
            return 842.0;
        }

        $index += strlen($marker);
        $end = strpos($pageObjectText, ']', $index);
        if ($end === false) {
            return 842.0;
        }

        $parts = preg_split('/\s+/', trim(substr($pageObjectText, $index, $end - $index)));
        if (count($parts) !== 4) {
            return 842.0;
        }

        $y0 = (float)$parts[1];
        $y1 = (float)$parts[3];
        return abs($y1 - $y0);
    }

    private static function ColorCommand($run)
    {
        $f = __NAMESPACE__ . '\PdfNumbers::threeDecimals';
        return $f($run->Color->Red / 255.0) . ' '
            . $f($run->Color->Green / 255.0) . ' '
            . $f($run->Color->Blue / 255.0) . ' rg';
    }

    private static function BuildStreamObject($objectId, $streamBytes)
    {
        return self::BuildFlatStreamObject($objectId, $streamBytes);
    }

    private static function BuildFlatStreamObject($objectId, $streamBytes)
    {
        $header = $objectId . " 0 obj\n<< /Length " . strlen($streamBytes) . " >>\nstream\n";
        $footer = "\nendstream\nendobj\n";
        return $header . $streamBytes . $footer;
    }

    private static function RebuildPdf($table)
    {
        $ids = array_keys($table->Objects);
        sort($ids, SORT_NUMERIC);

        $output = $table->HeaderBytes;
        $offsets = array();
        for ($i = 0; $i < count($ids); $i++) {
            $offsets[$ids[$i]] = strlen($output);
            $bytes = $table->Objects[$ids[$i]];
            $output .= $bytes;
            if (strlen($bytes) === 0 || $bytes[strlen($bytes) - 1] !== "\n") {
                $output .= "\n";
            }
        }

        $xrefOffset = strlen($output);
        $maxId = $ids[count($ids) - 1];
        $xref = "xref\n0 " . ($maxId + 1) . "\n0000000000 65535 f \n";
        for ($id = 1; $id <= $maxId; $id++) {
            if (isset($offsets[$id])) {
                $xref .= sprintf('%010d', $offsets[$id]) . " 00000 n \n";
            } else {
                $xref .= "0000000000 00000 f \n";
            }
        }

        $xref .= 'trailer<< /Size ' . ($maxId + 1);
        if ($table->RootReference !== null && $table->RootReference !== '') {
            $xref .= "\n/Root " . $table->RootReference;
        }

        if ($table->InfoReference !== null && $table->InfoReference !== '') {
            $xref .= "\n/Info " . $table->InfoReference;
        }

        $xref .= " >>\nstartxref\n" . $xrefOffset . "\n%%EOF";
        $output .= $xref;
        return $output;
    }
}
