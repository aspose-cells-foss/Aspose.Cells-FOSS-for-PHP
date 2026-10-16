<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\ExtendedDocumentPropertiesModel;

final class ExtendedDocumentProperties
{
    private $_model;

    public function __construct(ExtendedDocumentPropertiesModel $model)
    {
        $this->_model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Application') {
            return $this->_model->Application;
        }
        if ($propertyName === 'AppVersion') {
            return $this->_model->AppVersion;
        }
        if ($propertyName === 'Company') {
            return $this->_model->Company;
        }
        if ($propertyName === 'Manager') {
            return $this->_model->Manager;
        }
        if ($propertyName === 'DocSecurity') {
            return $this->_model->DocSecurity ?? 0;
        }
        if ($propertyName === 'HyperlinkBase') {
            return $this->_model->HyperlinkBase;
        }
        if ($propertyName === 'ScaleCrop') {
            return $this->_model->ScaleCrop ?? false;
        }
        if ($propertyName === 'LinksUpToDate') {
            return $this->_model->LinksUpToDate ?? false;
        }
        if ($propertyName === 'SharedDoc') {
            return $this->_model->SharedDoc ?? false;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Application') {
            $this->_model->Application = $value ?? '';
            return;
        }
        if ($propertyName === 'AppVersion') {
            $this->_model->AppVersion = $value ?? '';
            return;
        }
        if ($propertyName === 'Company') {
            $this->_model->Company = $value ?? '';
            return;
        }
        if ($propertyName === 'Manager') {
            $this->_model->Manager = $value ?? '';
            return;
        }
        if ($propertyName === 'DocSecurity') {
            if ($value < 0) {
                throw new CellsException('DocSecurity must be non-negative.');
            }
            $this->_model->DocSecurity = $value;
            return;
        }
        if ($propertyName === 'HyperlinkBase') {
            $this->_model->HyperlinkBase = $value ?? '';
            return;
        }
        if ($propertyName === 'ScaleCrop') {
            $this->_model->ScaleCrop = $value;
            return;
        }
        if ($propertyName === 'LinksUpToDate') {
            $this->_model->LinksUpToDate = $value;
            return;
        }
        if ($propertyName === 'SharedDoc') {
            $this->_model->SharedDoc = $value;
            return;
        }
    }
}
