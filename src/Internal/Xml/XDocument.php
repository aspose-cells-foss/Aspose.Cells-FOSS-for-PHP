<?php

namespace Aspose\Cells_FOSS\Internal\Xml;

/**
 * Represents an XML document in the LINQ-to-XML style shim.
 */
class XDocument
{
    /**
     * The document declaration.
     *
     * @var XDeclaration|null
     */
    public $Declaration;

    /**
     * The document root element.
     *
     * @var XElement|null
     */
    public $Root;

    /**
     * Initializes a new instance of the {@see XDocument} class.
     *
     * @param XDeclaration|XElement|null $declaration The declaration or root element.
     * @param XElement|null $root The root element.
     */
    public function __construct($declaration = null, $root = null)
    {
        if ($declaration instanceof XDeclaration) {
            $this->Declaration = $declaration;
            if ($root instanceof XElement) {
                $this->Root = $root;
                $root->parent = null;
            }
            return;
        }

        if ($declaration instanceof XElement) {
            $this->Declaration = null;
            $this->Root = $declaration;
            $declaration->parent = null;
        }
    }

    /**
     * Loads a document from XML text or a stream resource.
     *
     * @param mixed $source The XML text or stream resource.
     * @return XDocument The document.
     */
    public static function Load($source)
    {
        if (is_resource($source)) {
            $source = stream_get_contents($source);
        }

        $document = new XDocument();

        if ($source === null || $source === '') {
            return $document;
        }

        $dom = new \DOMDocument();
        $dom->preserveWhiteSpace = true;
        $previous = libxml_use_internal_errors(true);
        $loaded = $dom->loadXML($source, LIBXML_NONET);
        libxml_clear_errors();
        libxml_use_internal_errors($previous);

        if (!$loaded || $dom->documentElement === null) {
            throw new \RuntimeException('The XML document could not be parsed.');
        }

        $document->Root = self::fromDom($dom->documentElement);
        return $document;
    }

    /**
     * Converts a DOM element to an {@see XElement}.
     *
     * @param \DOMElement $domElement The DOM element.
     * @return XElement The element.
     */
    private static function fromDom(\DOMElement $domElement)
    {
        $namespaceUri = $domElement->namespaceURI;
        $name = XmlName::combine($namespaceUri === null ? '' : $namespaceUri, $domElement->localName);
        $element = new XElement($name);

        if ($domElement->hasAttributes()) {
            foreach ($domElement->attributes as $attribute) {
                if ($attribute->namespaceURI === XNamespace::XmlnsUri) {
                    continue;
                }

                $attributeName = XmlName::combine(
                    $attribute->namespaceURI === null ? '' : $attribute->namespaceURI,
                    $attribute->localName
                );
                $element->attributes[] = new XAttribute($attributeName, $attribute->value);
            }
        }

        foreach ($domElement->childNodes as $child) {
            if ($child->nodeType === XML_ELEMENT_NODE) {
                $element->Add(self::fromDom($child));
            } elseif ($child->nodeType === XML_TEXT_NODE || $child->nodeType === XML_CDATA_SECTION_NODE) {
                $element->nodes[] = $child->nodeValue;
            }
        }

        return $element;
    }

    /**
     * Serializes the document.
     *
     * @return string The XML.
     */
    public function save()
    {
        $output = '';

        if ($this->Declaration !== null) {
            $output .= '<?xml version="' . $this->Declaration->Version;
            if ($this->Declaration->Encoding !== null && $this->Declaration->Encoding !== '') {
                $output .= '" encoding="' . $this->Declaration->Encoding;
            }
            if ($this->Declaration->Standalone !== null && $this->Declaration->Standalone !== '') {
                $output .= '" standalone="' . $this->Declaration->Standalone;
            }
            $output .= '"?>';
        }

        if ($this->Root !== null) {
            $output .= $this->Root->toString();
        }

        return $output;
    }

    /**
     * Serializes the document.
     *
     * @return string The XML.
     */
    public function toString()
    {
        return $this->save();
    }

    /**
     * Serializes the document.
     *
     * @return string The XML.
     */
    public function __toString()
    {
        return $this->save();
    }
}
