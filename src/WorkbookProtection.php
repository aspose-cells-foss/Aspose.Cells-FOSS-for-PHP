<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\WorkbookProtectionModel;

final class WorkbookProtection
{
    private $_model;

    public function __construct(WorkbookProtectionModel $model)
    {
        $this->_model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'LockStructure') {
            return $this->_model->LockStructure;
        }
        if ($propertyName === 'LockWindows') {
            return $this->_model->LockWindows;
        }
        if ($propertyName === 'LockRevision') {
            return $this->_model->LockRevision;
        }
        if ($propertyName === 'WorkbookPassword') {
            return $this->_model->WorkbookPassword;
        }
        if ($propertyName === 'RevisionsPassword') {
            return $this->_model->RevisionsPassword;
        }
        if ($propertyName === 'IsProtected') {
            return $this->_model->HasStoredState();
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'LockStructure') {
            $this->_model->LockStructure = $value;
            return;
        }
        if ($propertyName === 'LockWindows') {
            $this->_model->LockWindows = $value;
            return;
        }
        if ($propertyName === 'LockRevision') {
            $this->_model->LockRevision = $value;
            return;
        }
        if ($propertyName === 'WorkbookPassword') {
            $this->_model->WorkbookPassword = $value ?? '';
            return;
        }
        if ($propertyName === 'RevisionsPassword') {
            $this->_model->RevisionsPassword = $value ?? '';
            return;
        }
    }
}
