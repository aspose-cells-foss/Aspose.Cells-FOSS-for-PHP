<?php
namespace Aspose\Cells_FOSS\Core;

final class DiagnosticBag
{
    private $entries;

    public function __construct()
    {
        $this->entries = array();
    }

    public function __get($propertyName)
    {
        if ($propertyName === "Entries") {
            return $this->entries;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === "Entries") {
            throw new \RuntimeException("Entries is read-only.");
        }
    }

    public function Add($entry)
    {
        $this->entries[] = $entry;
    }
}
