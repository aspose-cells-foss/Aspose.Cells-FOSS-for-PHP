<?php

namespace Aspose\Cells_FOSS\Skia;

/**
 * Thin wrapper over sk_fontmgr_t for glyph fallback lookups.
 */
final class SKFontManager
{
    private $handle;
    private $disposed = false;
    private static $defaultInstance = null;

    private function __construct($handle)
    {
        $this->handle = $handle;
    }

    public static function DefaultManager()
    {
        if (self::$defaultInstance === null) {
            self::$defaultInstance = new SKFontManager(SkiaLibrary::ffi()->sk_fontmgr_create_default());
        }

        return self::$defaultInstance;
    }

    public function MatchCharacter($codepoint)
    {
        if ($this->handle === null) {
            return null;
        }

        $handle = SkiaLibrary::ffi()->sk_fontmgr_match_family_style_character(
            $this->handle,
            null,
            null,
            null,
            0,
            (int)$codepoint
        );

        return SKTypeface::fromHandle($handle);
    }

    public function Dispose()
    {
        if ($this->disposed) {
            return;
        }

        $this->disposed = true;
        if ($this->handle !== null) {
            SkiaLibrary::ffi()->sk_fontmgr_unref($this->handle);
            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
