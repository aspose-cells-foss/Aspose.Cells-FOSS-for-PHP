<?php

namespace Aspose\Cells_FOSS\Rendering;

final class PdfType3TextFontBackend implements PdfTextFontBackend
{
    private $registry;

    public function __construct()
    {
        $this->registry = new PdfFontUsageRegistry();
    }

    public function Initialize($runs)
    {
        $this->registry->AddRuns($runs);
        $this->registry->FinalizeSubsets();
    }

    public function BuildObjects(&$nextObjectId)
    {
        $builder = new PdfType3FontBuilder();
        return $builder->BuildObjects($this->registry, $nextObjectId);
    }

    public function BuildFontResourceText()
    {
        $builder = '/Font <<';
        $usages = &$this->registry->Usages();
        for ($i = 0; $i < count($usages); $i++) {
            $subsets = &$usages[$i]->Subsets();
            for ($s = 0; $s < count($subsets); $s++) {
                $subset = $subsets[$s];
                $builder .= ' /' . $subset->ResourceName . ' ' . $subset->FontObjectId . " 0 R";
            }
        }

        $builder .= ' >>';
        return $builder;
    }

    public function EncodeRun($run)
    {
        $segments = array();
        if ($run === null || $run->Typeface === null || $run->Text === null || $run->Text === '') {
            return $segments;
        }

        $usage = $this->registry->UsageFor($run->Typeface, $run->UsePositiveYFontMatrix);
        $cursor = $run->XPt;
        $textLen = mb_strlen($run->Text, 'UTF-8');
        $index = 0;
        while ($index < $textLen) {
            $subset = null;
            $bytes = '';
            $segmentStart = $cursor;

            while ($index < $textLen) {
                $token = mb_substr($run->Text, $index, 1, 'UTF-8');
                $tokenSubset = null;
                $code = 0;
                $width1000 = 0.0;
                if (!$usage->TryResolveToken($token, $tokenSubset, $code, $width1000)) {
                    $index++;
                    continue;
                }

                if ($subset === null) {
                    $subset = $tokenSubset;
                } elseif ($subset !== $tokenSubset) {
                    // Different subset; rewind this token so the next segment picks it up.
                    break;
                }

                $bytes .= chr($code);
                $cursor += $run->FontSizePt * ($width1000 / 1000.0);
                $index++;
            }

            if ($subset !== null && strlen($bytes) > 0) {
                $segment = new PdfEncodedTextSegment();
                $segment->ResourceName = $subset->ResourceName;
                $segment->EncodedBytes = $bytes;
                $segment->StartXPt = $segmentStart;
                $segments[] = $segment;
            }
        }

        return $segments;
    }

    public function Dispose()
    {
        $this->registry->Dispose();
    }
}
