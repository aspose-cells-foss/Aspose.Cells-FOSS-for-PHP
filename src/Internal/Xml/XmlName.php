<?php

namespace Aspose\Cells_FOSS\Internal\Xml;

/**
 * Provides helpers for working with expanded XML names of the form "{uri}local".
 */
final class XmlName
{
    /**
     * Splits an expanded name into its namespace URI and local name.
     *
     * @param string|null $name The expanded name.
     * @return array{0: string, 1: string} A pair of namespace URI and local name.
     */
    public static function split($name)
    {
        if ($name === null || $name === '') {
            return array('', '');
        }

        if ($name[0] === '{') {
            $close = strpos($name, '}');
            if ($close !== false) {
                return array(substr($name, 1, $close - 1), substr($name, $close + 1));
            }
        }

        return array('', $name);
    }

    /**
     * Combines a namespace URI and a local name into an expanded name.
     *
     * @param string $namespaceUri The namespace URI.
     * @param string $localName The local name.
     * @return string The expanded name.
     */
    public static function combine($namespaceUri, $localName)
    {
        if ($namespaceUri === null || $namespaceUri === '') {
            return $localName;
        }

        return '{' . $namespaceUri . '}' . $localName;
    }

    /**
     * Removes surrounding braces from a namespace literal, if present.
     *
     * @param string|null $value The value.
     * @return string The unbraced value.
     */
    public static function stripBraces($value)
    {
        if ($value === null) {
            return '';
        }

        $length = strlen($value);
        if ($length >= 2 && $value[0] === '{' && $value[$length - 1] === '}') {
            return substr($value, 1, $length - 2);
        }

        return $value;
    }

    private function __construct()
    {
    }
}
