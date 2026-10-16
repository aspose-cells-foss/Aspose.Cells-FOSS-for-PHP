<?php

namespace Aspose\Cells_FOSS\Internal\Xml;

/**
 * Represents an XML element in the LINQ-to-XML style shim.
 *
 * Names are stored as expanded names of the form "{uri}local". Namespaces are
 * emitted lazily as the tree is serialized so that the output matches the
 * structural intent of the original C# code.
 */
class XElement
{
    /**
     * The expanded element name of the form "{uri}local".
     *
     * @var string
     */
    public $name;

    /**
     * The ordered attributes.
     *
     * @var XAttribute[]
     */
    public $attributes = array();

    /**
     * The ordered child nodes; each entry is an {@see XElement} or a string.
     *
     * @var array
     */
    public $nodes = array();

    /**
     * The parent element, or null when detached.
     *
     * @var XElement|null
     */
    public $parent = null;

    /**
     * Initializes a new instance of the {@see XElement} class.
     *
     * @param string $name The expanded element name.
     * @param mixed ...$content The element content (attributes, elements, or text).
     */
    public function __construct($name, ...$content)
    {
        $this->name = (string) $name;
        foreach ($content as $item) {
            $this->appendContent($item);
        }
    }

    /**
     * Appends content (attributes, elements, or text) to this element.
     *
     * @param mixed $content The content.
     * @return void
     */
    private function appendContent($content)
    {
        if ($content === null) {
            return;
        }

        if (is_array($content)) {
            foreach ($content as $item) {
                $this->appendContent($item);
            }
            return;
        }

        if ($content instanceof XAttribute) {
            $this->setAttributeObject($content);
            return;
        }

        if ($content instanceof XElement) {
            $content->parent = $this;
            $this->nodes[] = $content;
            return;
        }

        if (is_bool($content)) {
            $this->nodes[] = $content ? 'true' : 'false';
            return;
        }

        $this->nodes[] = (string) $content;
    }

    /**
     * Adds or replaces an attribute object.
     *
     * @param XAttribute $attribute The attribute.
     * @return void
     */
    private function setAttributeObject(XAttribute $attribute)
    {
        for ($i = 0; $i < count($this->attributes); $i++) {
            if ($this->attributes[$i]->Name === $attribute->Name) {
                $this->attributes[$i] = $attribute;
                return;
            }
        }

        $this->attributes[] = $attribute;
    }

    /**
     * Finds the first child element with the specified name.
     *
     * @param string $name The expanded element name.
     * @return XElement|null The element, or null.
     */
    public static function Parse($xml)
    {
        $document = XDocument::Load($xml);
        return $document->Root;
    }

    public function Element($name)
    {
        foreach ($this->nodes as $node) {
            if ($node instanceof XElement && $node->name === $name) {
                return $node;
            }
        }

        return null;
    }

    /**
     * Returns the child elements, optionally filtered by name.
     *
     * @param string|null $name The expanded element name, or null for all.
     * @return XElement[] The matching elements.
     */
    public function Elements($name = null)
    {
        $result = array();
        foreach ($this->nodes as $node) {
            if (!$node instanceof XElement) {
                continue;
            }

            if ($name === null || $node->name === $name) {
                $result[] = $node;
            }
        }

        return $result;
    }

    /**
     * Finds an attribute by name.
     *
     * @param string $name The expanded attribute name.
     * @return XAttribute|null The attribute, or null.
     */
    public function Attribute($name)
    {
        foreach ($this->attributes as $attribute) {
            if ($attribute->Name === $name) {
                return $attribute;
            }
        }

        return null;
    }

    /**
     * Adds or replaces an attribute value. A null value removes the attribute.
     *
     * @param string $name The expanded attribute name.
     * @param mixed $value The value.
     * @return void
     */
    public function SetAttributeValue($name, $value)
    {
        if ($value === null) {
            $this->removeAttribute($name);
            return;
        }

        $this->setAttributeObject(new XAttribute($name, $value));
    }

    /**
     * Removes the attribute with the specified name.
     *
     * @param string $name The expanded attribute name.
     * @return void
     */
    private function removeAttribute($name)
    {
        $kept = array();
        foreach ($this->attributes as $attribute) {
            if ($attribute->Name !== $name) {
                $kept[] = $attribute;
            }
        }

        $this->attributes = $kept;
    }

    /**
     * Adds content to the element.
     *
     * @param mixed $content The content.
     * @return void
     */
    public function Add($content = null)
    {
        if (func_num_args() > 1) {
            foreach (func_get_args() as $item) {
                $this->appendContent($item);
            }
            return;
        }

        $this->appendContent($content);
    }

    /**
     * Detaches this element from its parent.
     *
     * @return void
     */
    public function Remove()
    {
        if ($this->parent === null) {
            return;
        }

        $parent = $this->parent;
        $kept = array();
        foreach ($parent->nodes as $node) {
            if ($node !== $this) {
                $kept[] = $node;
            }
        }

        $parent->nodes = $kept;
        $this->parent = null;
    }

    /**
     * Returns this element and all of its descendants.
     *
     * @return XElement[] The elements.
     */
    public function DescendantsAndSelf()
    {
        $result = array($this);
        foreach ($this->nodes as $node) {
            if ($node instanceof XElement) {
                foreach ($node->DescendantsAndSelf() as $descendant) {
                    $result[] = $descendant;
                }
            }
        }

        return $result;
    }

    /**
     * Returns the concatenated text of this element and all descendants.
     *
     * @return string The text.
     */
    public function GetValue()
    {
        $text = '';
        foreach ($this->nodes as $node) {
            if ($node instanceof XElement) {
                $text .= $node->GetValue();
            } else {
                $text .= $node;
            }
        }

        return $text;
    }

    /**
     * Gets a property value.
     *
     * @param string $propertyName The property name.
     * @return mixed The value.
     */
    public function __get($propertyName)
    {
        if ($propertyName === 'Value') {
            return $this->GetValue();
        }

        if ($propertyName === 'Name') {
            return $this->name;
        }

        if ($propertyName === 'HasElements') {
            foreach ($this->nodes as $node) {
                if ($node instanceof XElement) {
                    return true;
                }
            }

            return false;
        }

        if ($propertyName === 'HasAttributes') {
            return count($this->attributes) > 0;
        }

        return null;
    }

    /**
     * Sets a property value.
     *
     * @param string $propertyName The property name.
     * @param mixed $value The value.
     * @return void
     */
    public function __set($propertyName, $value)
    {
        if ($propertyName === 'Value') {
            $this->nodes = array($value === null ? '' : (string) $value);
            return;
        }

        if ($propertyName === 'Name') {
            $this->name = (string) $value;
            return;
        }

        throw new \RuntimeException('Cannot set property ' . $propertyName . '.');
    }

    /**
     * Returns the serialized XML of this element.
     *
     * @return string The XML.
     */
    public function toString()
    {
        $state = array('counter' => 0);
        return $this->write(array(), '', $state);
    }

    /**
     * Returns the serialized XML of this element.
     *
     * @return string The XML.
     */
    public function __toString()
    {
        return $this->toString();
    }

    /**
     * Returns the preferred namespace prefix for a given namespace URI.
     *
     * @param string $namespaceUri The namespace URI.
     * @return string The preferred prefix (may be empty for a default namespace).
     */
    private static function preferredPrefix($namespaceUri)
    {
        $map = array(
            'http://schemas.openxmlformats.org/spreadsheetml/2006/main' => '',
            'http://schemas.openxmlformats.org/package/2006/relationships' => '',
            'http://schemas.openxmlformats.org/package/2006/content-types' => '',
            'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties' => '',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships' => 'r',
            'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes' => 'vt',
            'http://schemas.openxmlformats.org/package/2006/metadata/core-properties' => 'cp',
            'http://purl.org/dc/elements/1.1/' => 'dc',
            'http://purl.org/dc/terms/' => 'dcterms',
            'http://purl.org/dc/dcmitype/' => 'dcmitype',
            'http://www.w3.org/2001/XMLSchema-instance' => 'xsi',
            'http://schemas.openxmlformats.org/drawingml/2006/main' => 'a',
            'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing' => 'xdr',
            'http://schemas.openxmlformats.org/drawingml/2006/chart' => 'c',
            'http://schemas.openxmlformats.org/drawingml/2006/diagram' => 'dgm',
            'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing' => 'cdr',
            'http://schemas.openxmlformats.org/markup-compatibility/2006' => 'mc',
            'http://schemas.microsoft.com/office/drawing/2014/chartex' => 'cx',
            'http://schemas.microsoft.com/office/drawing/2012/chart' => 'c15',
            'urn:schemas-microsoft-com:vml' => 'v',
            'urn:schemas-microsoft-com:office:office' => 'o',
            'urn:schemas-microsoft-com:office:excel' => 'x',
            'http://schemas.microsoft.com/office/excel/2006/main' => 'x14',
            'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main' => 'x14',
        );

        if (array_key_exists($namespaceUri, $map)) {
            return $map[$namespaceUri];
        }

        return null;
    }

    /**
     * Serializes the element using the supplied namespace scope.
     *
     * @param array<string, string> $scope The in-scope namespace map (uri to prefix).
     * @param string $defaultUri The current default namespace URI.
     * @param array $state The shared serialization state.
     * @return string The XML.
     */
    private function write(array $scope, $defaultUri, array &$state)
    {
        $scope[XNamespace::XmlUri] = 'xml';

        $declarations = array();
        $newScope = $scope;
        $newDefault = $defaultUri;

        list($namespaceUri, $localName) = XmlName::split($this->name);

        if ($namespaceUri === '') {
            if ($defaultUri !== '') {
                $declarations[''] = '';
                $newDefault = '';
            }
            $tag = $localName;
        } elseif ($namespaceUri === $defaultUri) {
            $tag = $localName;
        } elseif (array_key_exists($namespaceUri, $scope)) {
            $tag = $scope[$namespaceUri] . ':' . $localName;
        } else {
            $prefix = self::preferredPrefix($namespaceUri);
            if ($prefix === '') {
                if ($defaultUri === '') {
                    $tag = $localName;
                    $declarations[''] = $namespaceUri;
                    $newDefault = $namespaceUri;
                } else {
                    $prefix = 'ns' . (++$state['counter']);
                    $tag = $prefix . ':' . $localName;
                    $declarations[$prefix] = $namespaceUri;
                    $newScope[$namespaceUri] = $prefix;
                }
            } else {
                $tag = $prefix . ':' . $localName;
                $declarations[$prefix] = $namespaceUri;
                $newScope[$namespaceUri] = $prefix;
            }
        }

        $attributeText = '';
        foreach ($this->attributes as $attribute) {
            list($attributeNamespace, $attributeLocal) = XmlName::split($attribute->Name);

            if ($attributeNamespace === XNamespace::XmlnsUri) {
                $prefix = ($attributeLocal === 'xmlns') ? '' : $attributeLocal;
                $uri = XmlName::stripBraces($attribute->Value);
                $declarations[$prefix] = $uri;
                if ($prefix === '') {
                    $newDefault = $uri;
                } else {
                    $newScope[$uri] = $prefix;
                }
                continue;
            }

            if ($attributeNamespace !== '' && $attributeNamespace !== XNamespace::XmlUri) {
                if (array_key_exists($attributeNamespace, $newScope)) {
                    $prefix = $newScope[$attributeNamespace];
                } else {
                    $prefix = self::preferredPrefix($attributeNamespace);
                    if ($prefix === null || $prefix === '') {
                        $prefix = 'ns' . (++$state['counter']);
                    }
                    $declarations[$prefix] = $attributeNamespace;
                    $newScope[$attributeNamespace] = $prefix;
                }
                $attributeText .= ' ' . $prefix . ':' . $attributeLocal . '="' . self::escapeAttribute($attribute->Value) . '"';
            } elseif ($attributeNamespace === XNamespace::XmlUri) {
                $attributeText .= ' xml:' . $attributeLocal . '="' . self::escapeAttribute($attribute->Value) . '"';
            } else {
                $attributeText .= ' ' . $attributeLocal . '="' . self::escapeAttribute($attribute->Value) . '"';
            }
        }

        $declarationText = '';
        foreach ($declarations as $prefix => $uri) {
            if ($prefix === '') {
                $declarationText .= ' xmlns="' . self::escapeAttribute($uri) . '"';
            } else {
                $declarationText .= ' xmlns:' . $prefix . '="' . self::escapeAttribute($uri) . '"';
            }
        }

        if (count($this->nodes) === 0) {
            return '<' . $tag . $declarationText . $attributeText . ' />';
        }

        $inner = '';
        foreach ($this->nodes as $node) {
            if ($node instanceof XElement) {
                $inner .= $node->write($newScope, $newDefault, $state);
            } else {
                $inner .= self::escapeText($node);
            }
        }

        return '<' . $tag . $declarationText . $attributeText . '>' . $inner . '</' . $tag . '>';
    }

    /**
     * Escapes text content.
     *
     * @param string $value The value.
     * @return string The escaped value.
     */
    private static function escapeText($value)
    {
        return str_replace(array('&', '<', '>'), array('&amp;', '&lt;', '&gt;'), $value);
    }

    /**
     * Escapes an attribute value.
     *
     * @param string $value The value.
     * @return string The escaped value.
     */
    private static function escapeAttribute($value)
    {
        return str_replace(
            array('&', '<', '>', '"', "\r", "\n", "\t"),
            array('&amp;', '&lt;', '&gt;', '&quot;', '&#xD;', '&#xA;', '&#x9;'),
            $value
        );
    }
}
