<?php

namespace Aspose\Cells_FOSS\Skia;

/**
 * Owns the single FFI binding to the native Skia C API (libSkiaSharp).
 */
final class SkiaLibrary
{
    private static $ffi = null;
    private static $libraryPath = null;

    private function __construct()
    {
    }

    public static function ffi()
    {
        if (self::$ffi === null) {
            $path = self::libraryPath();
            if ($path === null || !is_file($path)) {
                throw new \RuntimeException('libSkiaSharp native library was not found. Set ASPOSE_SKIA_LIBRARY to its full path.');
            }

            self::$ffi = \FFI::cdef(self::definitions(), $path);
        }

        return self::$ffi;
    }

    public static function libraryPath()
    {
        if (self::$libraryPath !== null) {
            return self::$libraryPath;
        }

        $fromEnvironment = getenv('ASPOSE_SKIA_LIBRARY');
        if ($fromEnvironment !== false && $fromEnvironment !== '') {
            self::$libraryPath = $fromEnvironment;
            return self::$libraryPath;
        }

        $root = dirname(__DIR__, 2);
        $candidates = array(
            $root . '/runtime/win-x64/libSkiaSharp.dll',
            $root . '/runtime/win-x64/skia/libSkiaSharp.dll',
            $root . '/runtime/linux-x64/libSkiaSharp.so',
            $root . '/runtime/osx/libSkiaSharp.dylib',
            $root . '/runtime/libSkiaSharp.dll',
            $root . '/runtime/libSkiaSharp.so',
            $root . '/runtime/libSkiaSharp.dylib',
        );

        for ($i = 0; $i < count($candidates); $i++) {
            if (is_file($candidates[$i])) {
                self::$libraryPath = $candidates[$i];
                return self::$libraryPath;
            }
        }

        self::$libraryPath = null;
        return null;
    }

    public static function definitions()
    {
        return <<<'C'
typedef struct sk_canvas_t sk_canvas_t;
typedef struct sk_bitmap_t sk_bitmap_t;
typedef struct sk_paint_t sk_paint_t;
typedef struct sk_path_t sk_path_t;
typedef struct sk_typeface_t sk_typeface_t;
typedef struct sk_stream_t sk_stream_t;
typedef struct sk_font_t sk_font_t;
typedef struct sk_fontstyle_t sk_fontstyle_t;
typedef struct sk_string_t sk_string_t;
typedef struct sk_data_t sk_data_t;
typedef struct sk_image_t sk_image_t;
typedef struct sk_shader_t sk_shader_t;
typedef struct sk_maskfilter_t sk_maskfilter_t;
typedef struct sk_path_effect_t sk_path_effect_t;
typedef struct sk_surface_t sk_surface_t;
typedef struct sk_document_t sk_document_t;
typedef struct sk_wstream_t sk_wstream_t;
typedef struct sk_wstream_dynamicmemorystream_t sk_wstream_dynamicmemorystream_t;
typedef struct sk_surfaceprops_t sk_surfaceprops_t;
typedef struct sk_colorspace_t sk_colorspace_t;
typedef struct sk_fontmgr_t sk_fontmgr_t;
typedef struct sk_pixmap_t sk_pixmap_t;
typedef struct sk_codec_t sk_codec_t;
typedef struct sk_path_iterator_t sk_path_iterator_t;
typedef int sk_codec_result_t;
typedef uint32_t sk_color_t;
typedef int sk_paint_style_t;
typedef int sk_text_encoding_t;
typedef int sk_shader_tilemode_t;
typedef int sk_filter_quality_t;
typedef int sk_stroke_cap_t;
typedef int sk_stroke_join_t;
typedef int sk_clipop_t;
typedef int sk_font_style_slant_t;
typedef int sk_blurstyle_t;
typedef int sk_path_direction_t;
typedef int sk_colortype_t;
typedef int sk_alphatype_t;
typedef int sk_encoded_image_format_t;
struct sk_point_t { float x; float y; };
typedef struct sk_point_t sk_point_t;
struct sk_rect_t { float left; float top; float right; float bottom; };
typedef struct sk_rect_t sk_rect_t;
struct sk_irect_t { int32_t left; int32_t top; int32_t right; int32_t bottom; };
typedef struct sk_irect_t sk_irect_t;
struct sk_isize_t { int32_t w; int32_t h; };
typedef struct sk_isize_t sk_isize_t;
struct sk_matrix_t { float scaleX; float skewX; float transX; float skewY; float scaleY; float transY; float persp0; float persp1; float persp2; };
typedef struct sk_matrix_t sk_matrix_t;
struct sk_fontmetrics_t { float fFlags; float fTop; float fAscent; float fDescent; float fBottom; float fLeading; float fAvgCharWidth; float fMaxCharWidth; float fXMin; float fXMax; float fXHeight; float fCapHeight; float fUnderlineThickness; float fUnderlinePosition; float fStrikeoutThickness; float fStrikeoutPosition; };
typedef struct sk_fontmetrics_t sk_fontmetrics_t;
struct sk_imageinfo_t { sk_colorspace_t* colorspace; int32_t width; int32_t height; sk_colortype_t colorType; sk_alphatype_t alphaType; };
typedef struct sk_imageinfo_t sk_imageinfo_t;

sk_paint_t* sk_paint_new();
sk_paint_t* sk_paint_clone(sk_paint_t*);
void sk_paint_delete(sk_paint_t*);
void sk_paint_set_color(sk_paint_t*, sk_color_t);
void sk_paint_set_style(sk_paint_t*, sk_paint_style_t);
void sk_paint_set_antialias(sk_paint_t*, bool);
void sk_paint_set_stroke_width(sk_paint_t*, float);
void sk_paint_set_stroke_cap(sk_paint_t*, sk_stroke_cap_t);
void sk_paint_set_stroke_join(sk_paint_t*, sk_stroke_join_t);
void sk_paint_set_stroke_miter(sk_paint_t*, float);
void sk_paint_set_shader(sk_paint_t*, sk_shader_t*);
void sk_paint_set_maskfilter(sk_paint_t*, sk_maskfilter_t*);
void sk_paint_set_path_effect(sk_paint_t*, sk_path_effect_t*);
void sk_paint_set_filter_quality(sk_paint_t*, sk_filter_quality_t);

int sk_canvas_save(sk_canvas_t*);
void sk_canvas_restore(sk_canvas_t*);
void sk_canvas_restore_to_count(sk_canvas_t*, int);
int sk_canvas_get_save_count(sk_canvas_t*);
void sk_canvas_translate(sk_canvas_t*, float, float);
void sk_canvas_scale(sk_canvas_t*, float, float);
void sk_canvas_rotate_degrees(sk_canvas_t*, float);
void sk_canvas_concat(sk_canvas_t*, const sk_matrix_t*);
void sk_canvas_clear(sk_canvas_t*, sk_color_t);
void sk_canvas_clip_rect_with_operation(sk_canvas_t*, const sk_rect_t*, sk_clipop_t, bool);
void sk_canvas_clip_path_with_operation(sk_canvas_t*, const sk_path_t*, sk_clipop_t, bool);
void sk_canvas_draw_rect(sk_canvas_t*, const sk_rect_t*, const sk_paint_t*);
void sk_canvas_draw_round_rect(sk_canvas_t*, const sk_rect_t*, float, float, const sk_paint_t*);
void sk_canvas_draw_line(sk_canvas_t*, float, float, float, float, sk_paint_t*);
void sk_canvas_draw_circle(sk_canvas_t*, float, float, float, const sk_paint_t*);
void sk_canvas_draw_oval(sk_canvas_t*, const sk_rect_t*, const sk_paint_t*);
void sk_canvas_draw_path(sk_canvas_t*, const sk_path_t*, const sk_paint_t*);
void sk_canvas_draw_image(sk_canvas_t*, const sk_image_t*, float, float, const sk_paint_t*);
void sk_canvas_draw_image_rect(sk_canvas_t*, const sk_image_t*, const sk_rect_t*, const sk_rect_t*, const sk_paint_t*);
void sk_canvas_draw_simple_text(sk_canvas_t*, const char* text, size_t byte_length, sk_text_encoding_t encoding, float x, float y, const sk_font_t* cfont, const sk_paint_t* cpaint);
int sk_canvas_save_layer(sk_canvas_t*, const sk_rect_t*, const sk_paint_t*);
sk_canvas_t* sk_canvas_new_from_bitmap(const sk_bitmap_t*);
void sk_canvas_destroy(sk_canvas_t*);

sk_bitmap_t* sk_bitmap_new();
void sk_bitmap_destructor(sk_bitmap_t*);
bool sk_bitmap_try_alloc_pixels(sk_bitmap_t*, const sk_imageinfo_t*, size_t);
sk_color_t sk_bitmap_get_pixel_color(const sk_bitmap_t*, int32_t, int32_t);
bool sk_bitmap_extract_subset(const sk_bitmap_t*, sk_bitmap_t*, const sk_irect_t*);
void sk_bitmap_erase_rect(const sk_bitmap_t*, sk_color_t, const sk_irect_t*);
size_t sk_bitmap_get_row_bytes(const sk_bitmap_t*);
void sk_bitmap_get_info(const sk_bitmap_t*, sk_imageinfo_t*);
uint32_t* sk_bitmap_get_addr_32(const sk_bitmap_t*, int32_t, int32_t);
bool sk_bitmap_is_null(const sk_bitmap_t*);
void sk_bitmap_set_immutable(sk_bitmap_t*);
sk_shader_t* sk_bitmap_make_shader(const sk_bitmap_t*, sk_shader_tilemode_t, sk_shader_tilemode_t, const sk_matrix_t*);

sk_path_t* sk_path_new();
sk_path_t* sk_path_clone(const sk_path_t*);
void sk_path_delete(sk_path_t*);
void sk_path_move_to(sk_path_t*, float, float);
void sk_path_line_to(sk_path_t*, float, float);
void sk_path_quad_to(sk_path_t*, float, float, float, float);
void sk_path_cubic_to(sk_path_t*, float, float, float, float, float, float);
void sk_path_conic_to(sk_path_t*, float, float, float, float, float);
void sk_path_arc_to_with_oval(sk_path_t*, const sk_rect_t*, float, float, bool);
void sk_path_close(sk_path_t*);
void sk_path_add_rect(sk_path_t*, const sk_rect_t*, sk_path_direction_t);
void sk_path_add_oval(sk_path_t*, const sk_rect_t*, sk_path_direction_t);
void sk_path_add_rounded_rect(sk_path_t*, const sk_rect_t*, float, float, sk_path_direction_t);
void sk_path_add_circle(sk_path_t*, float, float, float, sk_path_direction_t);
void sk_path_add_path(sk_path_t*, sk_path_t*, int);
void sk_path_rewind(sk_path_t*);
void sk_path_reset(sk_path_t*);
void sk_path_get_bounds(const sk_path_t*, sk_rect_t*);
void sk_path_compute_tight_bounds(const sk_path_t*, sk_rect_t*);
int sk_path_count_verbs(const sk_path_t*);
void sk_path_transform(sk_path_t*, const sk_matrix_t*);
sk_path_iterator_t* sk_path_create_iter(const sk_path_t*, bool);
int sk_path_iter_next(sk_path_iterator_t*, sk_point_t points[4]);
float sk_path_iter_conic_weight(const sk_path_iterator_t*);
bool sk_path_iter_is_close_line(const sk_path_iterator_t*);
bool sk_path_iter_is_closed_contour(const sk_path_iterator_t*);
void sk_path_iter_destroy(sk_path_iterator_t*);

sk_typeface_t* sk_typeface_create_default();
sk_typeface_t* sk_typeface_create_from_name(const char* familyName, const sk_fontstyle_t* style);
sk_typeface_t* sk_typeface_create_from_file(const char* path, int index);
void sk_typeface_unref(sk_typeface_t*);
sk_string_t* sk_typeface_get_family_name(const sk_typeface_t*);
uint16_t sk_typeface_unichar_to_glyph(const sk_typeface_t*, uint32_t);
int sk_typeface_get_table_tags(const sk_typeface_t*, uint32_t tags[], int maxTags);
size_t sk_typeface_get_table_size(const sk_typeface_t*, uint32_t tag);
size_t sk_typeface_get_table_data(const sk_typeface_t*, uint32_t tag, void* data, size_t length);
sk_stream_t* sk_typeface_open_stream(const sk_typeface_t*, int* ttcIndex);
size_t sk_stream_get_length(const sk_stream_t*);
size_t sk_stream_read(const sk_stream_t*, void* buffer, size_t size);
void sk_stream_destroy(sk_stream_t*);
sk_fontstyle_t* sk_fontstyle_new(int weight, int width, sk_font_style_slant_t slant);
void sk_fontstyle_delete(sk_fontstyle_t*);
const char* sk_string_get_c_str(const sk_string_t*);
void sk_string_destructor(const sk_string_t*);

sk_font_t* sk_font_new();
sk_font_t* sk_font_new_with_values(sk_typeface_t*, float size, float scaleX, float skewX);
void sk_font_delete(sk_font_t*);
float sk_font_measure_text(const sk_font_t* font, const char* text, size_t byteLength, sk_text_encoding_t encoding, sk_rect_t* bounds, const sk_paint_t* paint);
float sk_font_get_metrics(const sk_font_t* font, sk_fontmetrics_t* metrics);
void sk_font_set_subpixel(sk_font_t*, bool);
void sk_font_set_hinting(sk_font_t*, int);
void sk_font_set_linear_metrics(sk_font_t*, bool);
void sk_font_set_size(sk_font_t*, float);
float sk_font_get_size(const sk_font_t*);
void sk_font_set_typeface(sk_font_t*, sk_typeface_t*);
sk_typeface_t* sk_font_get_typeface(sk_font_t*);
uint16_t sk_font_unichar_to_glyph(const sk_font_t*, uint32_t);
void sk_text_utils_get_path(const char* text, size_t length, sk_text_encoding_t encoding, float x, float y, const sk_font_t* font, sk_path_t* path);

sk_shader_t* sk_shader_new_color(sk_color_t color);
sk_shader_t* sk_shader_new_linear_gradient(const sk_point_t* points, const sk_color_t* colors, const float* colorPos, int colorCount, sk_shader_tilemode_t tileMode, const sk_matrix_t* localMatrix);
sk_shader_t* sk_shader_new_radial_gradient(const sk_point_t* center, float radius, const sk_color_t* colors, const float* colorPos, int colorCount, sk_shader_tilemode_t tileMode, const sk_matrix_t* localMatrix);
void sk_shader_unref(sk_shader_t*);

sk_maskfilter_t* sk_maskfilter_new_blur(sk_blurstyle_t, float sigma);
void sk_maskfilter_unref(sk_maskfilter_t*);
sk_path_effect_t* sk_path_effect_create_dash(const float* intervals, int count, float phase);
void sk_path_effect_unref(sk_path_effect_t*);

sk_data_t* sk_data_new_with_copy(const char* src, size_t length);
void sk_data_unref(sk_data_t*);
size_t sk_data_get_size(const sk_data_t*);
const void* sk_data_get_bytes(const sk_data_t*);
sk_image_t* sk_image_new_from_encoded(sk_data_t* encoded, const sk_irect_t* subset);
sk_image_t* sk_image_new_from_bitmap(const sk_bitmap_t*);
int sk_image_get_width(const sk_image_t*);
int sk_image_get_height(const sk_image_t*);
void sk_image_unref(const sk_image_t*);
sk_data_t* sk_image_encode_specific(const sk_image_t*, sk_encoded_image_format_t, int);
sk_shader_t* sk_image_make_shader(const sk_image_t*, sk_shader_tilemode_t, sk_shader_tilemode_t, const sk_matrix_t*, sk_filter_quality_t);

sk_surface_t* sk_surface_new_raster(const sk_imageinfo_t*, size_t rowBytes, const sk_surfaceprops_t*);
sk_canvas_t* sk_surface_get_canvas(sk_surface_t*);
sk_image_t* sk_surface_new_image_snapshot(sk_surface_t*);
void sk_surface_unref(sk_surface_t*);

sk_wstream_dynamicmemorystream_t* sk_dynamicmemorywstream_new();
void sk_dynamicmemorywstream_destroy(sk_wstream_dynamicmemorystream_t*);
void sk_dynamicmemorywstream_copy_to(sk_wstream_dynamicmemorystream_t*, char* data);
size_t sk_wstream_bytes_written(sk_wstream_t* cstream);
sk_document_t* sk_document_create_pdf_from_stream(sk_wstream_t* stream);
sk_canvas_t* sk_document_begin_page(sk_document_t* document, float width, float height, const sk_rect_t* content);
void sk_document_end_page(sk_document_t*);
void sk_document_close(sk_document_t*);
void sk_document_unref(sk_document_t*);

sk_fontmgr_t* sk_fontmgr_create_default();
void sk_fontmgr_unref(sk_fontmgr_t*);
sk_typeface_t* sk_fontmgr_match_family_style_character(sk_fontmgr_t*, const char* familyName, sk_fontstyle_t* style, const char** bcp47, int bcp47Count, int32_t character);
C;
    }

    public static function colorToNative($color)
    {
        return (($color->Alpha & 0xFF) << 24) | (($color->Red & 0xFF) << 16) | (($color->Green & 0xFF) << 8) | ($color->Blue & 0xFF);
    }

    public static function colorFromNative($value)
    {
        $value = (int)$value;
        return new SKColor(($value >> 16) & 0xFF, ($value >> 8) & 0xFF, $value & 0xFF, ($value >> 24) & 0xFF);
    }

    public static function toCString($text)
    {
        return is_string($text) ? $text : (string)$text;
    }
}
