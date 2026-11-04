<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKTypeface
{
    private $handle;
    private $disposed = false;
    private $isDefault = false;
    private static $defaultInstance = null;
    private $fontData = null;
    private $tableDirectory = null;

    private function __construct($handle, $isDefault)
    {
        $this->handle = $handle;
        $this->isDefault = $isDefault;
    }

    public static function Default()
    {
        if (self::$defaultInstance === null) {
            $ffi = SkiaLibrary::ffi();
            self::$defaultInstance = new SKTypeface($ffi->sk_typeface_create_default(), true);
        }

        return self::$defaultInstance;
    }

    public static function FromFamilyName($familyName, $weight = SKFontStyleWeight::Normal, $width = SKFontStyleWidth::Normal, $slant = SKFontStyleSlant::Upright)
    {
        $ffi = SkiaLibrary::ffi();
        $style = $ffi->sk_fontstyle_new((int)$weight, (int)$width, (int)$slant);
        $handle = $ffi->sk_typeface_create_from_name((string)$familyName, $style);
        $ffi->sk_fontstyle_delete($style);

        if ($handle === null) {
            return null;
        }

        return new SKTypeface($handle, false);
    }

    public static function FromFile($path, $index = 0)
    {
        $ffi = SkiaLibrary::ffi();
        $handle = $ffi->sk_typeface_create_from_file((string)$path, (int)$index);
        if ($handle === null) {
            return null;
        }

        return new SKTypeface($handle, false);
    }

    public static function fromHandle($handle)
    {
        if ($handle === null) {
            return null;
        }

        return new SKTypeface($handle, false);
    }

    public function handle()
    {
        return $this->handle;
    }

    public function ContainsGlyph($codepoint)
    {
        if ($this->disposed || $this->handle === null) {
            return false;
        }

        $ffi = SkiaLibrary::ffi();
        return $ffi->sk_typeface_unichar_to_glyph($this->handle, (int)$codepoint) !== 0;
    }

    /**
     * Returns the list of TrueType/OpenType table tags available on this typeface.
     * Each tag is a 32-bit value encoding four ASCII characters.
     */
    public function GetTableTags()
    {
        if ($this->disposed || $this->handle === null) {
            return array();
        }

        $this->ensureTableDirectory();
        if ($this->tableDirectory === null) {
            return array();
        }

        return array_keys($this->tableDirectory);
    }

    /**
     * Returns the size in bytes of the requested font table, or 0 if absent.
     */
    public function GetTableSize($tag)
    {
        if ($this->disposed || $this->handle === null) {
            return 0;
        }

        $this->ensureTableDirectory();
        if ($this->tableDirectory === null || !isset($this->tableDirectory[(int)$tag])) {
            return 0;
        }

        return $this->tableDirectory[(int)$tag]['length'];
    }

    /**
     * Returns the raw binary contents of the requested font table, or an empty
     * string if the table is not available.
     */
    public function GetTableData($tag)
    {
        if ($this->disposed || $this->handle === null) {
            return '';
        }

        $this->ensureTableDirectory();
        if ($this->tableDirectory === null || !isset($this->tableDirectory[(int)$tag])) {
            return '';
        }

        $entry = $this->tableDirectory[(int)$tag];
        return substr($this->fontData, $entry['offset'], $entry['length']);
    }

    /**
     * Opens the font stream and parses the TrueType/OpenType table directory so
     * that individual tables can be extracted. The raw font bytes and the parsed
     * directory are cached on the instance.
     */
    private function ensureTableDirectory()
    {
        if ($this->tableDirectory !== null) {
            return;
        }

        $this->tableDirectory = array();
        $this->fontData = '';

        $ffi = SkiaLibrary::ffi();
        $ttcIndex = $ffi->new('int');
        $stream = $ffi->sk_typeface_open_stream($this->handle, \FFI::addr($ttcIndex));
        if ($stream === null) {
            return;
        }

        $length = (int)$ffi->sk_stream_get_length($stream);
        if ($length <= 0 || $length > 100000000) {
            $ffi->sk_stream_destroy($stream);
            return;
        }

        $buffer = $ffi->new('char[' . $length . ']');
        $read = (int)$ffi->sk_stream_read($stream, $buffer, $length);
        $ffi->sk_stream_destroy($stream);
        if ($read <= 0) {
            return;
        }

        $this->fontData = \FFI::string($buffer, $read);
        if (strlen($this->fontData) < 12) {
            return;
        }

        // sfVersion (4) + numTables (2) + searchRange (2) + entrySelector (2) + rangeShift (2)
        $numTables = (ord($this->fontData[4]) << 8) | ord($this->fontData[5]);
        if ($numTables <= 0 || $numTables > 1000) {
            return;
        }

        $offset = 12;
        for ($i = 0; $i < $numTables; $i++) {
            if ($offset + 16 > strlen($this->fontData)) {
                break;
            }

            $tag = (ord($this->fontData[$offset]) << 24)
                | (ord($this->fontData[$offset + 1]) << 16)
                | (ord($this->fontData[$offset + 2]) << 8)
                | ord($this->fontData[$offset + 3]);
            $tableOffset = (ord($this->fontData[$offset + 8]) << 24)
                | (ord($this->fontData[$offset + 9]) << 16)
                | (ord($this->fontData[$offset + 10]) << 8)
                | ord($this->fontData[$offset + 11]);
            $tableLength = (ord($this->fontData[$offset + 12]) << 24)
                | (ord($this->fontData[$offset + 13]) << 16)
                | (ord($this->fontData[$offset + 14]) << 8)
                | ord($this->fontData[$offset + 15]);

            $this->tableDirectory[$tag] = array(
                'offset' => $tableOffset,
                'length' => $tableLength,
            );

            $offset += 16;
        }
    }

    public function __get($propertyName)
    {
        if ($propertyName === 'FamilyName') {
            if ($this->disposed || $this->handle === null) {
                return '';
            }

            $ffi = SkiaLibrary::ffi();
            $name = $ffi->sk_typeface_get_family_name($this->handle);
            $value = (string)$ffi->sk_string_get_c_str($name);
            $ffi->sk_string_destructor($name);
            return $value;
        }

        return null;
    }

    public function Dispose()
    {
        if ($this->disposed) {
            return;
        }

        $this->disposed = true;
        if ($this->isDefault) {
            return;
        }

        if ($this->handle !== null) {
            $ffi = SkiaLibrary::ffi();
            $ffi->sk_typeface_unref($this->handle);
            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
