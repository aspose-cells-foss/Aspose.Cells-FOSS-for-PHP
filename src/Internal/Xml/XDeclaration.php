<?php

namespace Aspose\Cells_FOSS\Internal\Xml;

/**
 * Represents an XML declaration such as `<?xml version="1.0" encoding="utf-8" standalone="yes"?>`.
 */
final class XDeclaration
{
    /**
     * The XML version.
     *
     * @var string|null
     */
    public $Version;

    /**
     * The document encoding.
     *
     * @var string|null
     */
    public $Encoding;

    /**
     * The standalone flag.
     *
     * @var string|null
     */
    public $Standalone;

    /**
     * Initializes a new instance of the {@see XDeclaration} class.
     *
     * @param string|null $version The XML version.
     * @param string|null $encoding The document encoding.
     * @param string|null $standalone The standalone flag.
     */
    public function __construct($version = '1.0', $encoding = 'utf-8', $standalone = null)
    {
        $this->Version = $version;
        $this->Encoding = $encoding;
        $this->Standalone = $standalone;
    }
}
