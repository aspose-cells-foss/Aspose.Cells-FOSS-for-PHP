<?php

namespace Aspose\Cells_FOSS\Rendering;

final class PdfFontUsageRegistry
{
    private $usages = array();

    /**
     * @return PdfType3FontUsage[]
     */
    public function &Usages()
    {
        return $this->usages;
    }

    public function AddRuns($runs)
    {
        for ($i = 0; $i < count($runs); $i++) {
            $run = $runs[$i];
            $usage = $this->UsageFor($run->Typeface, $run->UsePositiveYFontMatrix);
            $usage->AddText($run->Text);
        }
    }

    public function FinalizeSubsets()
    {
        $resourceIndex = 1;
        for ($i = 0; $i < count($this->usages); $i++) {
            $this->usages[$i]->FinalizeSubsets($resourceIndex);
        }
    }

    public function UsageFor($typeface, $usePositiveYFontMatrix)
    {
        for ($i = 0; $i < count($this->usages); $i++) {
            if ($this->usages[$i]->Typeface() === $typeface
                && $this->usages[$i]->UsePositiveYFontMatrix() === $usePositiveYFontMatrix) {
                return $this->usages[$i];
            }
        }

        $usage = new PdfType3FontUsage($typeface, $usePositiveYFontMatrix);
        $this->usages[] = $usage;
        return $usage;
    }

    public function Dispose()
    {
        for ($i = 0; $i < count($this->usages); $i++) {
            $this->usages[$i]->Dispose();
        }

        $this->usages = array();
    }
}
