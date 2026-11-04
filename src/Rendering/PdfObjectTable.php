<?php

namespace Aspose\Cells_FOSS\Rendering;

final class PdfObjectTable
{
    public $Objects = array();
    public $PageObjectIds = array();
    public $HeaderBytes = '';
    public $RootReference = '';
    public $InfoReference = '';

    public function MaxObjectId()
    {
        $max = 0;
        foreach ($this->Objects as $key => $value) {
            if ($key > $max) {
                $max = $key;
            }
        }

        return $max;
    }

    /**
     * @param string $pdf
     * @return PdfObjectTable
     */
    public static function Parse($pdf)
    {
        $text = $pdf;
        $startXrefIndex = strrpos($text, 'startxref');
        if ($startXrefIndex === false) {
            throw new \InvalidArgumentException('PDF startxref not found.');
        }

        $xrefOffset = self::ParseStartXref($text, $startXrefIndex);
        $trailerIndex = strpos($text, 'trailer', $xrefOffset);
        if ($trailerIndex === false) {
            throw new \InvalidArgumentException('PDF trailer not found.');
        }

        $table = new self();
        $offsets = self::ParseXrefOffsets($text, $xrefOffset, $trailerIndex);
        $sorted = array();
        foreach ($offsets as $key => $value) {
            if ($key > 0) {
                $sorted[$key] = $value;
            }
        }

        uasort($sorted, function ($a, $b) {
            if ($a === $b) return 0;
            return $a < $b ? -1 : 1;
        });

        $firstOffset = PHP_INT_MAX;
        foreach ($sorted as $offset) {
            if ($offset < $firstOffset) {
                $firstOffset = $offset;
            }
        }

        $table->HeaderBytes = substr($pdf, 0, $firstOffset);

        $offsetKeys = array_keys($sorted);
        for ($i = 0; $i < count($offsetKeys); $i++) {
            $objectId = $offsetKeys[$i];
            $start = $sorted[$objectId];
            $end = $i + 1 < count($offsetKeys) ? $sorted[$offsetKeys[$i + 1]] : $xrefOffset;
            $length = $end - $start;
            if ($length <= 0) {
                continue;
            }

            $bytes = substr($pdf, $start, $length);
            $table->Objects[$objectId] = $bytes;

            $objectText = $bytes;
            if (strpos($objectText, '/Type /Page') !== false
                && strpos($objectText, '/Type /Pages') === false) {
                $table->PageObjectIds[] = $objectId;
            }
        }

        $trailer = substr($text, $trailerIndex, $startXrefIndex - $trailerIndex);
        $table->RootReference = self::ParseReference($trailer, '/Root');
        $table->InfoReference = self::ParseReference($trailer, '/Info');
        return $table;
    }

    private static function ParseXrefOffsets($text, $xrefOffset, $trailerIndex)
    {
        $offsets = array();
        $xrefText = substr($text, $xrefOffset, $trailerIndex - $xrefOffset);
        $lines = preg_split("/\r\n|\n|\r/", $xrefText);

        $first = isset($lines[0]) ? trim($lines[0]) : '';
        if ($first !== 'xref') {
            throw new \InvalidArgumentException('Unsupported PDF xref table.');
        }

        $lineIndex = 1;
        while ($lineIndex < count($lines)) {
            $line = trim($lines[$lineIndex]);
            if ($line === '') {
                $lineIndex++;
                continue;
            }

            $parts = preg_split('/\s+/', $line);
            if (count($parts) !== 2) {
                break;
            }

            $start = (int)$parts[0];
            $count = (int)$parts[1];
            for ($i = 0; $i < $count; $i++) {
                $lineIndex++;
                if ($lineIndex >= count($lines)) {
                    throw new \InvalidArgumentException('Unexpected end of xref table.');
                }

                $entry = $lines[$lineIndex];
                if (strlen($entry) < 18) {
                    continue;
                }

                if ($entry[17] === 'n') {
                    $offsets[$start + $i] = (int)substr($entry, 0, 10);
                }
            }

            $lineIndex++;
        }

        return $offsets;
    }

    private static function ParseStartXref($text, $startXrefIndex)
    {
        $cursor = $startXrefIndex + strlen('startxref');
        while ($cursor < strlen($text) && ctype_space($text[$cursor])) {
            $cursor++;
        }

        $end = $cursor;
        while ($end < strlen($text) && ctype_digit($text[$end])) {
            $end++;
        }

        return (int)substr($text, $cursor, $end - $cursor);
    }

    private static function ParseReference($trailer, $key)
    {
        $index = strpos($trailer, $key);
        if ($index === false) {
            return null;
        }

        $index += strlen($key);
        while ($index < strlen($trailer) && ctype_space($trailer[$index])) {
            $index++;
        }

        $end = $index;
        while ($end < strlen($trailer)
            && $trailer[$end] !== "\n" && $trailer[$end] !== "\r" && $trailer[$end] !== '>') {
            $end++;
        }

        return trim(substr($trailer, $index, $end - $index));
    }
}
