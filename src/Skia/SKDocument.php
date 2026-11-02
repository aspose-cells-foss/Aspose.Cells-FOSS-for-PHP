<?php

namespace Aspose\Cells_FOSS\Skia;

/**
 * Managed wrapper over sk_document_t (currently used only for PDF output).
 */
final class SKDocument
{
    private $handle;
    private $disposed = false;

    private function __construct($handle)
    {
        $this->handle = $handle;
    }

    public static function CreatePdf($stream)
    {
        $ffi = SkiaLibrary::ffi();
        $handle = $ffi->sk_document_create_pdf_from_stream($ffi->cast('sk_wstream_t*', $stream->handle()));
        return $handle === null ? null : new SKDocument($handle);
    }

    public function BeginPage($width, $height, $contentRect = null)
    {
        $native = $contentRect !== null ? $contentRect->toNative() : null;
        $canvas = SkiaLibrary::ffi()->sk_document_begin_page(
            $this->handle,
            (float)$width,
            (float)$height,
            $native !== null ? \FFI::addr($native) : null
        );

        return SKCanvas::fromHandle($canvas, false);
    }

    public function EndPage()
    {
        SkiaLibrary::ffi()->sk_document_end_page($this->handle);
    }

    public function Close()
    {
        if ($this->disposed || $this->handle === null) {
            return;
        }

        SkiaLibrary::ffi()->sk_document_close($this->handle);
        SkiaLibrary::ffi()->sk_document_unref($this->handle);
        $this->disposed = true;
        $this->handle = null;
    }

    public function Dispose()
    {
        if ($this->disposed) {
            return;
        }

        $this->disposed = true;
        if ($this->handle !== null) {
            SkiaLibrary::ffi()->sk_document_unref($this->handle);
            $this->handle = null;
        }
    }

    public function __destruct()
    {
        $this->Dispose();
    }
}
