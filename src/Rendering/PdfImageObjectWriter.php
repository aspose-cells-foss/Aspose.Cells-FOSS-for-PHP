<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\PictureModel;
use Aspose\Cells_FOSS\Skia\SKBitmap;

final class PdfImageObjectWriter
{
    /**
     * @param string $originalPdf
     * @param \Aspose\Cells_FOSS\Core\WorkbookModel $workbook
     * @return string
     */
    public function Rewrite($originalPdf, $workbook)
    {
        if ($originalPdf === null || $workbook === null) {
            return $originalPdf;
        }

        $replacements = self::CollectLosslessPictureStreams($workbook);
        if (count($replacements) === 0) {
            return $originalPdf;
        }

        $table = PdfObjectTable::Parse($originalPdf);
        $changed = false;
        $objectIds = array_keys($table->Objects);
        for ($i = 0; $i < count($objectIds); $i++) {
            $key = $objectIds[$i];
            $objectBytes = $table->Objects[$key];
            $info = self::TryParseImageObject($objectBytes);
            if ($info === null) {
                continue;
            }

            $replacement = self::TryMatchReplacement($replacements, $info['Width'], $info['Height']);
            if ($replacement === null) {
                continue;
            }

            $table->Objects[$key] = self::BuildImageObject($key, $info['Width'], $info['Height'], $replacement['StreamBytes']);
            $changed = true;
        }

        if (!$changed) {
            return $originalPdf;
        }

        return self::RebuildPdf($table);
    }

    private static function TryMatchReplacement(&$replacements, $width, $height)
    {
        for ($i = 0; $i < count($replacements); $i++) {
            if ($replacements[$i]['Used']) {
                continue;
            }

            if ($replacements[$i]['Width'] === $width && $replacements[$i]['Height'] === $height) {
                $replacements[$i]['Used'] = true;
                return $replacements[$i];
            }
        }

        return null;
    }

    private static function CollectLosslessPictureStreams($workbook)
    {
        $replacements = array();
        if ($workbook === null || $workbook->Worksheets === null) {
            return $replacements;
        }

        for ($sheetIndex = 0; $sheetIndex < count($workbook->Worksheets); $sheetIndex++) {
            $sheet = $workbook->Worksheets[$sheetIndex];
            if ($sheet === null) {
                continue;
            }

            for ($pictureIndex = 0; $sheet->Pictures !== null && $pictureIndex < count($sheet->Pictures); $pictureIndex++) {
                $picture = $sheet->Pictures[$pictureIndex];
                if (!self::ShouldPreserveLosslessly($picture)) {
                    continue;
                }

                $replacement = self::BuildReplacement($picture);
                if ($replacement !== null) {
                    $replacements[] = $replacement;
                }
            }

            for ($imageIndex = 0; $sheet->ShapeImages !== null && $imageIndex < count($sheet->ShapeImages); $imageIndex++) {
                $image = $sheet->ShapeImages[$imageIndex];
                $picture = new PictureModel();
                $picture->ImageData = $image->ImageData;
                $picture->ImageExtension = $image->Extension;
                if (!self::ShouldPreserveLosslessly($picture)) {
                    continue;
                }

                $replacement = self::BuildReplacement($picture);
                if ($replacement !== null) {
                    $replacements[] = $replacement;
                }
            }
        }

        return $replacements;
    }

    private static function ShouldPreserveLosslessly($picture)
    {
        if ($picture === null || $picture->ImageData === null || strlen($picture->ImageData) === 0
            || $picture->ImageExtension === null || $picture->ImageExtension === '') {
            return false;
        }

        $extension = strtolower(trim($picture->ImageExtension));
        return $extension === 'png' || $extension === 'gif' || $extension === 'bmp';
    }

    private static function BuildReplacement($picture)
    {
        $bitmap = SKBitmap::Decode($picture->ImageData);
        if ($bitmap === null || $bitmap->Width <= 0 || $bitmap->Height <= 0) {
            return null;
        }

        $raw = '';
        for ($y = 0; $y < $bitmap->Height; $y++) {
            for ($x = 0; $x < $bitmap->Width; $x++) {
                $color = $bitmap->GetPixel($x, $y);
                if ($color->Alpha < 255) {
                    $alpha = $color->Alpha / 255.0;
                    $raw .= chr(self::BlendOnWhite($color->Red, $alpha));
                    $raw .= chr(self::BlendOnWhite($color->Green, $alpha));
                    $raw .= chr(self::BlendOnWhite($color->Blue, $alpha));
                } else {
                    $raw .= chr($color->Red);
                    $raw .= chr($color->Green);
                    $raw .= chr($color->Blue);
                }
            }
        }

        return array(
            'Width' => $bitmap->Width,
            'Height' => $bitmap->Height,
            'StreamBytes' => gzcompress($raw, 9),
            'Used' => false,
        );
    }

    private static function BlendOnWhite($channel, $alpha)
    {
        $value = $channel * $alpha + 255.0 * (1.0 - $alpha);
        if ($value < 0.0) {
            $value = 0.0;
        }

        if ($value > 255.0) {
            $value = 255.0;
        }

        return (int)round($value);
    }

    private static function TryParseImageObject($objectBytes)
    {
        $text = $objectBytes;
        if (strpos($text, '/Subtype /Image') === false) {
            return null;
        }

        $width = self::ParseIntValue($text, '/Width');
        $height = self::ParseIntValue($text, '/Height');
        if ($width <= 0 || $height <= 0) {
            return null;
        }

        return array('Width' => $width, 'Height' => $height);
    }

    private static function ParseIntValue($text, $key)
    {
        $index = strpos($text, $key);
        if ($index === false) {
            return 0;
        }

        $index += strlen($key);
        while ($index < strlen($text) && ctype_space($text[$index])) {
            $index++;
        }

        $end = $index;
        while ($end < strlen($text) && ctype_digit($text[$end])) {
            $end++;
        }

        if ($end <= $index) {
            return 0;
        }

        return (int)substr($text, $index, $end - $index);
    }

    private static function BuildImageObject($objectId, $width, $height, $streamBytes)
    {
        $header = $objectId . " 0 obj\n<< /Type /XObject\n/Subtype /Image\n/Width "
            . $width . "\n/Height " . $height
            . "\n/ColorSpace /DeviceRGB\n/BitsPerComponent 8\n/Interpolate false\n/Filter /FlateDecode\n/Length "
            . strlen($streamBytes) . " >>\nstream\n";
        $footer = "\nendstream\nendobj\n";
        return $header . $streamBytes . $footer;
    }

    private static function RebuildPdf($table)
    {
        $ids = array_keys($table->Objects);
        sort($ids, SORT_NUMERIC);

        $output = $table->HeaderBytes;
        $offsets = array();
        for ($i = 0; $i < count($ids); $i++) {
            $offsets[$ids[$i]] = strlen($output);
            $bytes = $table->Objects[$ids[$i]];
            $output .= $bytes;
            if (strlen($bytes) === 0 || $bytes[strlen($bytes) - 1] !== "\n") {
                $output .= "\n";
            }
        }

        $xrefOffset = strlen($output);
        $maxId = $ids[count($ids) - 1];
        $xref = "xref\n0 " . ($maxId + 1) . "\n0000000000 65535 f \n";
        for ($id = 1; $id <= $maxId; $id++) {
            if (isset($offsets[$id])) {
                $xref .= sprintf('%010d', $offsets[$id]) . " 00000 n \n";
            } else {
                $xref .= "0000000000 00000 f \n";
            }
        }

        $xref .= 'trailer<< /Size ' . ($maxId + 1);
        if ($table->RootReference !== null && $table->RootReference !== '') {
            $xref .= "\n/Root " . $table->RootReference;
        }

        if ($table->InfoReference !== null && $table->InfoReference !== '') {
            $xref .= "\n/Info " . $table->InfoReference;
        }

        $xref .= " >>\nstartxref\n" . $xrefOffset . "\n%%EOF";
        $output .= $xref;
        return $output;
    }
}
