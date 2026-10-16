<?php
namespace Aspose\Cells_FOSS\Core;

use Aspose\Cells_FOSS\CultureInfo;

final class WorkbookSettingsModel
{
    public $DateSystem;
    public $DisplayCulture;

    public function __construct()
    {
        $this->DateSystem = DateSystem::Windows1900;
        $this->DisplayCulture = CultureInfo::GetInvariantCulture();
    }
}
