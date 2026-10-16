<?php

namespace Aspose\Cells_FOSS\Internal\Xml;

/**
 * Represents an XML attribute in the LINQ-to-XML style shim.
 */
class XAttribute
{
    /**
     * The expanded attribute name of the form "{uri}local".
     *
     * @var string
     */
    public $Name;

    /**
     * The attribute value.
     *
     * @var string
     */
    public $Value;

    /**
     * Initializes a new instance of the {@see XAttribute} class.
     *
     * @param string $name The expanded attribute name.
     * @param mixed $value The attribute value.
     */
    public function __construct($name, $value = null)
    {
        $this->Name = (string) $name;
        $this->Value = self::convertValue($name, $value);
    }

    /**
     * Converts a value to its attribute string representation.
     *
     * @param string $name The attribute name (used to detect namespace declarations).
     * @param mixed $value The value.
     * @return string The string representation.
     */
    private static function convertValue($name, $value)
    {
        if ($value === null) {
            return '';
        }

        if (is_bool($value)) {
            return $value ? 'true' : 'false';
        }

        $text = (string) $value;
        list($namespaceUri) = XmlName::split($name);
        if ($namespaceUri === XNamespace::XmlnsUri) {
            return XmlName::stripBraces($text);
        }

        return $text;
    }

    /**
     * Returns the attribute value.
     *
     * @return string The value.
     */
    public function __toString()
    {
        return $this->Value;
    }
}
