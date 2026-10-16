<?php

namespace Aspose\Cells_FOSS\Internal\Io;

/**
 * Represents a single entry inside a {@see ZipArchive}.
 */
class ZipEntry
{
    /**
     * The entry name (the part path inside the package).
     *
     * @var string
     */
    public $FullName;

    /**
     * The raw entry content.
     *
     * @var string
     */
    private $content;

    /**
     * Initializes a new instance of the {@see ZipEntry} class.
     *
     * @param string $fullName The entry name.
     * @param string $content The raw content.
     */
    public function __construct($fullName, $content)
    {
        $this->FullName = $fullName;
        $this->content = $content;
    }

    /**
     * Returns the raw entry content.
     *
     * @return string The content.
     */
    public function getContent()
    {
        return $this->content;
    }

    /**
     * Returns the raw entry content as a readable stream resource.
     *
     * @return resource The stream.
     */
    public function open()
    {
        $stream = fopen('php://temp', 'r+b');
        fwrite($stream, $this->content);
        rewind($stream);
        return $stream;
    }
}
