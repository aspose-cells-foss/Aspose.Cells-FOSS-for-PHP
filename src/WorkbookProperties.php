<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\WorkbookModel;

final class WorkbookProperties
{
    private $_workbookModel;
    private $_model;
    private $_protection;
    private $_view;
    private $_calculation;

    public function __construct(WorkbookModel $workbookModel)
    {
        $this->_workbookModel = $workbookModel;
        $this->_model = $workbookModel->Properties;
        $this->_protection = new WorkbookProtection($this->_model->Protection);
        $this->_view = new WorkbookView($workbookModel);
        $this->_calculation = new CalculationProperties($this->_model->Calculation);
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'CodeName') {
            return $this->_model->CodeName;
        }
        if ($propertyName === 'ShowObjects') {
            if ($this->_model->ShowObjects === null || $this->_model->ShowObjects === '') {
                return 'all';
            }
            return $this->_model->ShowObjects;
        }
        if ($propertyName === 'FilterPrivacy') {
            return $this->_model->FilterPrivacy;
        }
        if ($propertyName === 'ShowBorderUnselectedTables') {
            return $this->_model->ShowBorderUnselectedTables;
        }
        if ($propertyName === 'ShowInkAnnotation') {
            return $this->_model->ShowInkAnnotation;
        }
        if ($propertyName === 'BackupFile') {
            return $this->_model->BackupFile;
        }
        if ($propertyName === 'SaveExternalLinkValues') {
            return $this->_model->SaveExternalLinkValues;
        }
        if ($propertyName === 'UpdateLinks') {
            if ($this->_model->UpdateLinks === null || $this->_model->UpdateLinks === '') {
                return 'userSet';
            }
            return $this->_model->UpdateLinks;
        }
        if ($propertyName === 'HidePivotFieldList') {
            return $this->_model->HidePivotFieldList;
        }
        if ($propertyName === 'DefaultThemeVersion') {
            return $this->_model->DefaultThemeVersion;
        }
        if ($propertyName === 'Protection') {
            return $this->_protection;
        }
        if ($propertyName === 'View') {
            return $this->_view;
        }
        if ($propertyName === 'Calculation') {
            return $this->_calculation;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'CodeName') {
            $this->_model->CodeName = $value ?? '';
            return;
        }
        if ($propertyName === 'ShowObjects') {
            $this->_model->ShowObjects = WorkbookPropertySupport::NormalizeShowObjects($value);
            return;
        }
        if ($propertyName === 'FilterPrivacy') {
            $this->_model->FilterPrivacy = $value;
            return;
        }
        if ($propertyName === 'ShowBorderUnselectedTables') {
            $this->_model->ShowBorderUnselectedTables = $value;
            return;
        }
        if ($propertyName === 'ShowInkAnnotation') {
            $this->_model->ShowInkAnnotation = $value;
            return;
        }
        if ($propertyName === 'BackupFile') {
            $this->_model->BackupFile = $value;
            return;
        }
        if ($propertyName === 'SaveExternalLinkValues') {
            $this->_model->SaveExternalLinkValues = $value;
            return;
        }
        if ($propertyName === 'UpdateLinks') {
            $this->_model->UpdateLinks = WorkbookPropertySupport::NormalizeUpdateLinks($value);
            return;
        }
        if ($propertyName === 'HidePivotFieldList') {
            $this->_model->HidePivotFieldList = $value;
            return;
        }
        if ($propertyName === 'DefaultThemeVersion') {
            if ($value !== null && $value < 0) {
                throw new CellsException('DefaultThemeVersion must be non-negative.');
            }
            $this->_model->DefaultThemeVersion = $value;
            return;
        }
    }
}
