<?php

namespace Aspose\Cells_FOSS\Skia;

final class SKTypeface
{
    private $handle;
    private $disposed = false;
    private $isDefault = false;
    private static $defaultInstance = null;

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
