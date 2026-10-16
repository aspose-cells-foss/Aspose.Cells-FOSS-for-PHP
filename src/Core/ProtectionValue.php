<?php
namespace Aspose\Cells_FOSS\Core;

final class ProtectionValue
{
    public $IsLocked;
    public $IsHidden;

    public function __construct()
    {
        $this->IsLocked = true;
    }

    public function Clone()
    {
        $clone = new ProtectionValue();
        $clone->IsLocked = $this->IsLocked;
        $clone->IsHidden = $this->IsHidden;
        return $clone;
    }
}
