<?php

namespace Aspose\Cells_FOSS;

use Aspose\Cells_FOSS\Core\DateSystem;
use Aspose\Cells_FOSS\Core\WorkbookSettingsModel;

final class WorkbookSettings
{
    private $_model;

    public function __construct(WorkbookSettingsModel $model)
    {
        $this->_model = $model;
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'Date1904') {
            return $this->_model->DateSystem === DateSystem::Mac1904;
        }
        if ($propertyName === 'Culture') {
            // Expose a defensive copy so callers can mutate the returned CultureInfo
            // without changing workbook formatting rules until they assign it back.
            return $this->_model->DisplayCulture->Clone();
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Date1904') {
            $this->_model->DateSystem = $value ? DateSystem::Mac1904 : DateSystem::Windows1900;
            return;
        }
        if ($propertyName === 'Culture') {
            if ($value === null) {
                throw new \InvalidArgumentException('value');
            }

            $this->_model->DisplayCulture = $value->Clone();
            return;
        }
    }
}
