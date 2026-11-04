<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKTypeface;

/**
 * Reconstructs a valid TrueType/OpenType font file from the table data exposed
 * by an SKTypeface. SkiaSharp's C API does not surface the raw font stream, so
 * the file is rebuilt from individual tables.
 */
final class TrueTypeFontBytesBuilder
{
    private function __construct()
    {
    }

    /**
     * @param SKTypeface $typeface
     * @return string raw font file bytes, or empty string on failure
     */
    public static function Build($typeface)
    {
        if ($typeface === null) {
            return '';
        }

        $tags = $typeface->GetTableTags();
        if (count($tags) === 0) {
            return '';
        }

        $tables = array();
        foreach ($tags as $tag) {
            $data = $typeface->GetTableData($tag);
            if ($data === '' || strlen($data) === 0) {
                // Some tables (e.g. DSIG) may legitimately be empty; skip them.
                continue;
            }

            $tagName = self::tagToString($tag);
            $tables[$tagName] = $data;
        }

        if (count($tables) === 0) {
            return '';
        }

        $tagNames = array_keys($tables);
        sort($tagNames, SORT_STRING);

        $numTables = count($tagNames);
        $entrySelector = 0;
        $searchRange = 1;
        while ($searchRange * 2 <= $numTables) {
            $searchRange *= 2;
            $entrySelector++;
        }

        $searchRange *= 16;
        $rangeShift = $numTables * 16 - $searchRange;

        $isCff = isset($tables['CFF ']) || isset($tables['CFF2']);
        $sfVersion = $isCff ? 0x4F54544F : 0x00010000; // 'OTTO' or 0x00010000

        $output = '';
        $output .= self::writeUInt32($sfVersion);
        $output .= self::writeUInt16($numTables);
        $output .= self::writeUInt16($searchRange);
        $output .= self::writeUInt16($entrySelector);
        $output .= self::writeUInt16($rangeShift);

        $directoryOffset = strlen($output);
        // Reserve space for the table directory.
        $output .= str_repeat("\0", $numTables * 16);

        $records = array();
        for ($i = 0; $i < $numTables; $i++) {
            $tag = $tagNames[$i];
            $data = $tables[$tag];
            while ((strlen($output) & 3) !== 0) {
                $output .= "\0";
            }

            $offset = strlen($output);
            $output .= $data;
            while ((strlen($output) & 3) !== 0) {
                $output .= "\0";
            }

            $records[] = array(
                'Tag' => $tag,
                'Checksum' => self::calculateChecksum($data),
                'Offset' => $offset,
                'Length' => strlen($data),
            );
        }

        // Write the table directory into the reserved space.
        for ($i = 0; $i < count($records); $i++) {
            $entry = $directoryOffset + $i * 16;
            $output = substr_replace($output, $records[$i]['Tag'], $entry, 4);
            $output = substr_replace($output, self::writeUInt32($records[$i]['Checksum']), $entry + 4, 4);
            $output = substr_replace($output, self::writeUInt32($records[$i]['Offset']), $entry + 8, 4);
            $output = substr_replace($output, self::writeUInt32($records[$i]['Length']), $entry + 12, 4);
        }

        // Adjust the head table checksum adjustment.
        $headRecord = null;
        for ($i = 0; $i < count($records); $i++) {
            if ($records[$i]['Tag'] === 'head') {
                $headRecord = $records[$i];
                break;
            }
        }

        if ($headRecord !== null) {
            $output = substr_replace($output, self::writeUInt32(0), $headRecord['Offset'] + 8, 4);
            $checksum = self::calculateChecksum($output);
            $adjustment = 0xB1B0AFBA - $checksum;
            if ($adjustment < 0) {
                $adjustment += 0x100000000;
            }

            $output = substr_replace($output, self::writeUInt32($adjustment), $headRecord['Offset'] + 8, 4);
        }

        return $output;
    }

    private static function tagToString($tag)
    {
        return chr(($tag >> 24) & 0xFF) . chr(($tag >> 16) & 0xFF)
            . chr(($tag >> 8) & 0xFF) . chr($tag & 0xFF);
    }

    private static function writeUInt16($value)
    {
        return chr(($value >> 8) & 0xFF) . chr($value & 0xFF);
    }

    private static function writeUInt32($value)
    {
        return chr(($value >> 24) & 0xFF) . chr(($value >> 16) & 0xFF)
            . chr(($value >> 8) & 0xFF) . chr($value & 0xFF);
    }

    private static function calculateChecksum($bytes)
    {
        $sum = 0;
        $length = strlen($bytes);
        $paddedLength = ($length + 3) & ~3;
        for ($i = 0; $i < $paddedLength; $i += 4) {
            $value = 0;
            if ($i < $length) $value = (ord($bytes[$i]) << 24) & 0xFF000000;
            if ($i + 1 < $length) $value |= (ord($bytes[$i + 1]) << 16) & 0x00FF0000;
            if ($i + 2 < $length) $value |= (ord($bytes[$i + 2]) << 8) & 0x0000FF00;
            if ($i + 3 < $length) $value |= ord($bytes[$i + 3]) & 0x000000FF;
            $sum = ($sum + $value) & 0xFFFFFFFF;
        }

        return $sum;
    }
}
