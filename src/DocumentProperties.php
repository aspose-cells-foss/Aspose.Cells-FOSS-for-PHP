<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\DocumentPropertiesModel;

final class DocumentProperties
{
    private $_model;
    private $_core;
    private $_extended;

    public function __construct(DocumentPropertiesModel $model)
    {
        $this->_model = $model;
        $this->_core = new CoreDocumentProperties($model->Core);
        $this->_extended = new ExtendedDocumentProperties($model->Extended);
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Core') {
            return $this->_core;
        }
        if ($propertyName === 'Extended') {
            return $this->_extended;
        }
        if ($propertyName === 'Title') {
            return $this->_model->Core->Title;
        }
        if ($propertyName === 'Subject') {
            return $this->_model->Core->Subject;
        }
        if ($propertyName === 'Author') {
            return $this->_model->Core->Creator;
        }
        if ($propertyName === 'Keywords') {
            return $this->_model->Core->Keywords;
        }
        if ($propertyName === 'Comments') {
            return $this->_model->Core->Description;
        }
        if ($propertyName === 'Category') {
            return $this->_model->Core->Category;
        }
        if ($propertyName === 'Company') {
            return $this->_model->Extended->Company;
        }
        if ($propertyName === 'Manager') {
            return $this->_model->Extended->Manager;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Title') {
            $this->_model->Core->Title = $value ?? '';
            return;
        }
        if ($propertyName === 'Subject') {
            $this->_model->Core->Subject = $value ?? '';
            return;
        }
        if ($propertyName === 'Author') {
            $this->_model->Core->Creator = $value ?? '';
            return;
        }
        if ($propertyName === 'Keywords') {
            $this->_model->Core->Keywords = $value ?? '';
            return;
        }
        if ($propertyName === 'Comments') {
            $this->_model->Core->Description = $value ?? '';
            return;
        }
        if ($propertyName === 'Category') {
            $this->_model->Core->Category = $value ?? '';
            return;
        }
        if ($propertyName === 'Company') {
            $this->_model->Extended->Company = $value ?? '';
            return;
        }
        if ($propertyName === 'Manager') {
            $this->_model->Extended->Manager = $value ?? '';
            return;
        }
    }
}
