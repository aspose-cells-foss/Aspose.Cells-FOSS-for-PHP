<?php

namespace Aspose\Cells_FOSS\Core;

final class WorkbookViewModel
{
    public $XWindow = null;
    public $YWindow = null;
    public $WindowWidth = null;
    public $WindowHeight = null;
    public $FirstSheet = null;
    public $ShowHorizontalScroll = null;
    public $ShowVerticalScroll = null;
    public $ShowSheetTabs = null;
    public $TabRatio = null;
    public $Visibility = '';
    public $Minimized = false;
    public $AutoFilterDateGrouping = true;

    // Copies values from the specified source.
    public function CopyFrom($source)
    {
        $this->XWindow = $source->XWindow;
        $this->YWindow = $source->YWindow;
        $this->WindowWidth = $source->WindowWidth;
        $this->WindowHeight = $source->WindowHeight;
        $this->FirstSheet = $source->FirstSheet;
        $this->ShowHorizontalScroll = $source->ShowHorizontalScroll;
        $this->ShowVerticalScroll = $source->ShowVerticalScroll;
        $this->ShowSheetTabs = $source->ShowSheetTabs;
        $this->TabRatio = $source->TabRatio;
        $this->Visibility = $source->Visibility;
        $this->Minimized = $source->Minimized;
        $this->AutoFilterDateGrouping = $source->AutoFilterDateGrouping;
    }

    // Performs has stored state.
    public function HasStoredState($activeSheetIndex)
    {
        return $activeSheetIndex > 0
            || $this->XWindow !== null
            || $this->YWindow !== null
            || $this->WindowWidth !== null
            || $this->WindowHeight !== null
            || $this->FirstSheet !== null
            || $this->ShowHorizontalScroll !== null
            || $this->ShowVerticalScroll !== null
            || $this->ShowSheetTabs !== null
            || $this->TabRatio !== null
            || !($this->Visibility === null || $this->Visibility === '')
            || $this->Minimized
            || !$this->AutoFilterDateGrouping;
    }
}