<?php
namespace Aspose\Cells_FOSS\Core;

final class SharedStringRepository
{
    private $indices;
    private $values;

    public function __construct()
    {
        $this->indices = array();
        $this->values = array();
    }

    public function __get($propertyName)
    {
        if ($propertyName === "Values") {
            return $this->values;
        }
        return null;
    }

    public function __set($propertyName, $value)
    {
        if ($propertyName === "Values") {
            throw new \RuntimeException("Values is read-only.");
        }
    }

    public function Clear()
    {
        $this->indices = array();
        $this->values = array();
    }

    public function TryGetValue($index, &$value)
    {
        if ($index >= 0 && $index < count($this->values)) {
            $value = $this->values[$index];
            return true;
        }
        $value = "";
        return false;
    }

    public function Intern($value)
    {
        if ($value === null) {
            $value = "";
        }
        $key = $value;
        if (array_key_exists($key, $this->indices)) {
            return $this->indices[$key];
        }
        $index = count($this->values);
        $this->values[] = $value;
        $this->indices[$key] = $index;
        return $index;
    }
}
