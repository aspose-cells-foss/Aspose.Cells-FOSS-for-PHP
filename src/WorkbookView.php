<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\WorkbookModel;

final class WorkbookView
{
    private $_workbookModel;
    private $_model;

    public function __construct(WorkbookModel $workbookModel)
    {
        $this->_workbookModel = $workbookModel;
        $this->_model = $workbookModel->Properties->View;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'XWindow') {
            return $this->_model->XWindow ?? 0;
        }
        if ($propertyName === 'YWindow') {
            return $this->_model->YWindow ?? 0;
        }
        if ($propertyName === 'WindowWidth') {
            return $this->_model->WindowWidth ?? 0;
        }
        if ($propertyName === 'WindowHeight') {
            return $this->_model->WindowHeight ?? 0;
        }
        if ($propertyName === 'ActiveTab') {
            return $this->_workbookModel->ActiveSheetIndex;
        }
        if ($propertyName === 'FirstSheet') {
            return $this->_model->FirstSheet ?? 0;
        }
        if ($propertyName === 'ShowHorizontalScroll') {
            return $this->_model->ShowHorizontalScroll === null || $this->_model->ShowHorizontalScroll;
        }
        if ($propertyName === 'ShowVerticalScroll') {
            return $this->_model->ShowVerticalScroll === null || $this->_model->ShowVerticalScroll;
        }
        if ($propertyName === 'ShowSheetTabs') {
            return $this->_model->ShowSheetTabs === null || $this->_model->ShowSheetTabs;
        }
        if ($propertyName === 'TabRatio') {
            return $this->_model->TabRatio ?? 600;
        }
        if ($propertyName === 'Visibility') {
            if ($this->_model->Visibility === null || $this->_model->Visibility === '') {
                return 'visible';
            }
            return $this->_model->Visibility;
        }
        if ($propertyName === 'Minimized') {
            return $this->_model->Minimized;
        }
        if ($propertyName === 'AutoFilterDateGrouping') {
            return $this->_model->AutoFilterDateGrouping;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'XWindow') {
            $this->_model->XWindow = $value;
            return;
        }
        if ($propertyName === 'YWindow') {
            $this->_model->YWindow = $value;
            return;
        }
        if ($propertyName === 'WindowWidth') {
            if ($value < 0) {
                throw new CellsException('WindowWidth must be non-negative.');
            }
            $this->_model->WindowWidth = $value;
            return;
        }
        if ($propertyName === 'WindowHeight') {
            if ($value < 0) {
                throw new CellsException('WindowHeight must be non-negative.');
            }
            $this->_model->WindowHeight = $value;
            return;
        }
        if ($propertyName === 'ActiveTab') {
            if ($value < 0 || $value >= count($this->_workbookModel->Worksheets)) {
                throw new CellsException('ActiveTab must refer to an existing worksheet.');
            }
            $this->_workbookModel->ActiveSheetIndex = $value;
            return;
        }
        if ($propertyName === 'FirstSheet') {
            if ($value < 0) {
                throw new CellsException('FirstSheet must be non-negative.');
            }
            $this->_model->FirstSheet = $value;
            return;
        }
        if ($propertyName === 'ShowHorizontalScroll') {
            $this->_model->ShowHorizontalScroll = $value;
            return;
        }
        if ($propertyName === 'ShowVerticalScroll') {
            $this->_model->ShowVerticalScroll = $value;
            return;
        }
        if ($propertyName === 'ShowSheetTabs') {
            $this->_model->ShowSheetTabs = $value;
            return;
        }
        if ($propertyName === 'TabRatio') {
            if ($value < 0 || $value > 1000) {
                throw new CellsException('TabRatio must be between 0 and 1000.');
            }
            $this->_model->TabRatio = $value;
            return;
        }
        if ($propertyName === 'Visibility') {
            $this->_model->Visibility = WorkbookPropertySupport::NormalizeVisibility($value);
            return;
        }
        if ($propertyName === 'Minimized') {
            $this->_model->Minimized = $value;
            return;
        }
        if ($propertyName === 'AutoFilterDateGrouping') {
            $this->_model->AutoFilterDateGrouping = $value;
            return;
        }
    }
}
