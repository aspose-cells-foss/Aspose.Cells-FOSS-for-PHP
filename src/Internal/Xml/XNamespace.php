<?php

namespace Aspose\Cells_FOSS\Internal\Xml;

/**
 * Provides the well-known XML namespace constants used by the LINQ-to-XML style
 * shim. Namespaces are represented as plain strings of the form "{uri}" so that
 * concatenation with a local name (for example `XNamespace::Xml . "space"`)
 * produces an XML expanded name of the form "{uri}local".
 */
final class XNamespace
{
    /**
     * The XML namespace associated with the "xml" prefix.
     */
    const Xml = '{http://www.w3.org/XML/1998/namespace}';

    /**
     * The XML namespace associated with the "xmlns" prefix.
     */
    const Xmlns = '{http://www.w3.org/2000/xmlns/}';

    /**
     * The URI of the "xml" prefix namespace.
     */
    const XmlUri = 'http://www.w3.org/XML/1998/namespace';

    /**
     * The URI of the "xmlns" prefix namespace.
     */
    const XmlnsUri = 'http://www.w3.org/2000/xmlns/';

    private function __construct()
    {
    }
}
