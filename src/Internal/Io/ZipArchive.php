<?php

namespace Aspose\Cells_FOSS\Internal\Io;

/**
 * A thin wrapper over PHP's {@see \ZipArchive} that mirrors the subset of the
 * C# `System.IO.Compression.ZipArchive` API used by the library.
 *
 * A "stream" is either a file path string or a PHP stream resource. When a
 * stream resource is supplied, a temporary file is used as the backing store
 * because PHP's ZIP implementation works with file paths.
 */
class ZipArchive
{
    /**
     * The ZIP mode: "r" for read, "w" for write.
     *
     * @var string
     */
    private $mode;

    /**
     * The original stream (a path string or a resource).
     *
     * @var mixed
     */
    private $target;

    /**
     * The working file path used as the ZIP backing store.
     *
     * @var string
     */
    private $path;

    /**
     * Whether the working path is a temporary file that must be deleted.
     *
     * @var bool
     */
    private $temporary = false;

    /**
     * The underlying PHP ZIP archive.
     *
     * @var \ZipArchive
     */
    private $archive;

    /**
     * Buffered entries for write mode, keyed by entry name.
     *
     * @var array<string, string>
     */
    private $pending = array();

    /**
     * Initializes a new instance of the {@see ZipArchive} class.
     *
     * @param mixed $stream The stream (path string or resource).
     * @param string $mode The mode: "r" or "w".
     */
    public function __construct($stream, $mode = 'r')
    {
        $this->mode = $mode;
        $this->target = $stream;

        if ($mode === 'r') {
            $this->openForRead($stream);
        } else {
            $this->openForWrite();
        }
    }

    /**
     * Opens the archive for reading.
     *
     * @param mixed $stream The stream.
     * @return void
     */
    private function openForRead($stream)
    {
        if (is_string($stream)) {
            $this->path = $stream;
        } else {
            $this->path = self::createTemporaryPath();
            $this->temporary = true;
            file_put_contents($this->path, self::readStream($stream));
        }

        $archive = new \ZipArchive();
        if ($archive->open($this->path) !== true) {
            throw new \RuntimeException('The package is not a valid zip archive.');
        }

        $this->archive = $archive;
    }

    /**
     * Opens the archive for writing.
     *
     * @return void
     */
    private function openForWrite()
    {
        $this->path = self::createTemporaryPath();
        $this->temporary = true;

        $archive = new \ZipArchive();
        if ($archive->open($this->path, \ZipArchive::CREATE | \ZipArchive::OVERWRITE) !== true) {
            throw new \RuntimeException('The output package could not be created.');
        }

        $this->archive = $archive;
    }

    /**
     * Returns an entry by name, or null when it does not exist.
     *
     * @param string $uri The entry name.
     * @return ZipEntry|null The entry.
     */
    public function getEntry($uri)
    {
        $normalized = self::normalizeName($uri);

        if ($this->mode === 'w') {
            if (array_key_exists($normalized, $this->pending)) {
                return new ZipEntry($normalized, $this->pending[$normalized]);
            }

            return null;
        }

        $index = $this->archive->locateName($normalized);
        if ($index === false) {
            return null;
        }

        $content = $this->archive->getFromIndex($index);
        return new ZipEntry($normalized, $content);
    }

    /**
     * Adds (or replaces) an entry with the supplied content.
     *
     * @param string $path The entry name.
     * @param string $content The content.
     * @return ZipEntry The created entry.
     */
    public function addEntry($path, $content)
    {
        $normalized = self::normalizeName($path);
        $this->pending[$normalized] = (string) $content;
        return new ZipEntry($normalized, (string) $content);
    }

    /**
     * Creates an entry. Mirrors the C# `CreateEntry` call; when content is
     * supplied it is stored immediately.
     *
     * @param string $path The entry name.
     * @param string|null $content The content.
     * @return ZipEntry The created entry.
     */
    public function createEntry($path, $content = null)
    {
        return $this->addEntry($path, $content === null ? '' : $content);
    }

    /**
     * Finalizes the archive. In write mode this writes the buffered entries and
     * flushes the result back to the original stream.
     *
     * @return void
     */
    public function close()
    {
        if ($this->mode === 'w') {
            foreach ($this->pending as $name => $content) {
                $this->archive->addFromString($name, $content);
            }

            $this->archive->close();

            if (is_string($this->target)) {
                if ($this->path !== $this->target) {
                    copy($this->path, $this->target);
                }
            } else {
                self::writeStream($this->target, file_get_contents($this->path));
            }

            if ($this->temporary) {
                @unlink($this->path);
            }

            return;
        }

        $this->archive->close();
        if ($this->temporary) {
            @unlink($this->path);
        }
    }

    /**
     * Disposes the archive.
     *
     * @return void
     */
    public function dispose()
    {
        $this->close();
    }

    /**
     * Normalizes an entry name.
     *
     * @param string $name The name.
     * @return string The normalized name.
     */
    private static function normalizeName($name)
    {
        return ltrim(str_replace('\\', '/', $name), '/');
    }

    /**
     * Creates a temporary file path.
     *
     * @return string The path.
     */
    private static function createTemporaryPath()
    {
        $path = tempnam(sys_get_temp_dir(), 'acfxlsx');
        if ($path === false) {
            throw new \RuntimeException('A temporary file could not be created.');
        }

        return $path;
    }

    /**
     * Reads the entire contents of a stream.
     *
     * @param mixed $stream The stream.
     * @return string The content.
     */
    private static function readStream($stream)
    {
        if (is_resource($stream)) {
            $position = @ftell($stream);
            if ($position !== false && $position !== 0) {
                @rewind($stream);
            }

            $content = stream_get_contents($stream);
            return $content === false ? '' : $content;
        }

        return (string) $stream;
    }

    /**
     * Writes content to a stream.
     *
     * @param mixed $stream The stream.
     * @param string $content The content.
     * @return void
     */
    private static function writeStream($stream, $content)
    {
        if (is_resource($stream)) {
            fwrite($stream, $content);
        }
    }
}
