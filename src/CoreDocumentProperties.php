<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\CoreDocumentPropertiesModel;

final class CoreDocumentProperties
{
    private $_model;

    public function __construct(CoreDocumentPropertiesModel $model)
    {
        $this->_model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Title') {
            return $this->_model->Title;
        }
        if ($propertyName === 'Subject') {
            return $this->_model->Subject;
        }
        if ($propertyName === 'Creator') {
            return $this->_model->Creator;
        }
        if ($propertyName === 'Keywords') {
            return $this->_model->Keywords;
        }
        if ($propertyName === 'Description') {
            return $this->_model->Description;
        }
        if ($propertyName === 'LastModifiedBy') {
            return $this->_model->LastModifiedBy;
        }
        if ($propertyName === 'Revision') {
            return $this->_model->Revision;
        }
        if ($propertyName === 'Category') {
            return $this->_model->Category;
        }
        if ($propertyName === 'ContentStatus') {
            return $this->_model->ContentStatus;
        }
        if ($propertyName === 'Created') {
            return $this->_model->Created;
        }
        if ($propertyName === 'Modified') {
            return $this->_model->Modified;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Title') {
            $this->_model->Title = $value ?? '';
            return;
        }
        if ($propertyName === 'Subject') {
            $this->_model->Subject = $value ?? '';
            return;
        }
        if ($propertyName === 'Creator') {
            $this->_model->Creator = $value ?? '';
            return;
        }
        if ($propertyName === 'Keywords') {
            $this->_model->Keywords = $value ?? '';
            return;
        }
        if ($propertyName === 'Description') {
            $this->_model->Description = $value ?? '';
            return;
        }
        if ($propertyName === 'LastModifiedBy') {
            $this->_model->LastModifiedBy = $value ?? '';
            return;
        }
        if ($propertyName === 'Revision') {
            $this->_model->Revision = $value ?? '';
            return;
        }
        if ($propertyName === 'Category') {
            $this->_model->Category = $value ?? '';
            return;
        }
        if ($propertyName === 'ContentStatus') {
            $this->_model->ContentStatus = $value ?? '';
            return;
        }
        if ($propertyName === 'Created') {
            $this->_model->Created = $value;
            return;
        }
        if ($propertyName === 'Modified') {
            $this->_model->Modified = $value;
            return;
        }
    }
}
