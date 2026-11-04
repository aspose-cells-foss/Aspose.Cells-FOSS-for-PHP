<?php

namespace Aspose\Cells_FOSS\Rendering;

final class TrueTypeSubsetFontProgram
{
    const CompoundGlyphFlagArg1And2AreWords = 0x0001;
    const CompoundGlyphFlagMoreComponents = 0x0020;
    const CompoundGlyphFlagWeHaveAScale = 0x0008;
    const CompoundGlyphFlagWeHaveAnXAndYScale = 0x0040;
    const CompoundGlyphFlagWeHaveATwoByTwo = 0x0080;
    const CompoundGlyphFlagWeHaveInstructions = 0x0100;

    private $fontBytes;
    private $tables = array();
    private $unitsPerEm;
    private $ascent;
    private $descent;
    private $lineGap;
    private $xMin;
    private $yMin;
    private $xMax;
    private $yMax;
    private $numberOfHMetrics;
    private $numGlyphs;
    private $indexToLocFormat;
    private $glyphOffsets = array();
    private $metrics = array();

    public function __construct($fontBytes)
    {
        if ($fontBytes === null || strlen($fontBytes) < 12) {
            throw new \InvalidArgumentException('TrueType subset source bytes were invalid.');
        }

        $this->fontBytes = $fontBytes;
        $this->tables = self::ParseTableDirectory($fontBytes);
        $head = $this->TableBytes('head');
        $this->unitsPerEm = self::ReadUShort($head, 18);
        $this->xMin = self::ReadShort($head, 36);
        $this->yMin = self::ReadShort($head, 38);
        $this->xMax = self::ReadShort($head, 40);
        $this->yMax = self::ReadShort($head, 42);
        $this->indexToLocFormat = self::ReadShort($head, 50);
        $hhea = $this->TableBytes('hhea');
        $this->ascent = self::ReadShort($hhea, 4);
        $this->descent = self::ReadShort($hhea, 6);
        $this->lineGap = self::ReadShort($hhea, 8);
        $this->numberOfHMetrics = self::ReadUShort($hhea, 34);
        $maxp = $this->TableBytes('maxp');
        $this->numGlyphs = self::ReadUShort($maxp, 4);
        $this->glyphOffsets = $this->ParseGlyphOffsets();
        $this->metrics = $this->ParseHorizontalMetrics();
    }

    public function UnitsPerEm() { return $this->unitsPerEm; }
    public function Ascent() { return $this->ascent; }
    public function Descent() { return $this->descent; }
    public function XMin() { return $this->xMin; }
    public function YMin() { return $this->yMin; }
    public function XMax() { return $this->xMax; }
    public function YMax() { return $this->yMax; }

    /**
     * @param int[] $initialGlyphIds
     * @param array<int,int> $codeToGlyphId  byte code => glyph id
     * @param array<int,int> $glyphIdToUnicode  glyph id => unicode
     * @param string $subsetFontName
     * @return string
     */
    public function BuildSubset($initialGlyphIds, $codeToGlyphId, $glyphIdToUnicode, $subsetFontName)
    {
        if ($initialGlyphIds === null || count($initialGlyphIds) === 0) {
            throw new \InvalidArgumentException('TrueType subset requires at least one glyph.');
        }

        $closure = $this->BuildGlyphClosure($initialGlyphIds);
        $oldToNew = array();
        for ($i = 0; $i < count($closure); $i++) {
            $oldToNew[$closure[$i]] = $i;
        }

        $glyfResult = $this->BuildGlyfTable($closure, $oldToNew);
        $locaTable = $this->BuildLocaTable($glyfResult['LocaOffsets']);
        $hmtxTable = $this->BuildHmtxTable($closure);
        $maxpTable = $this->PatchMaxpTable(count($closure));
        $hheaTable = $this->PatchHheaTable(count($closure));
        $headTable = $this->PatchHeadTable();
        $cmapTable = $this->BuildCmapTable($codeToGlyphId, $glyphIdToUnicode, $oldToNew);

        $outputTables = array();
        $this->CopyTableIfPresent($outputTables, 'gasp');
        $outputTables['OS/2'] = $this->BuildOs2Table();
        $outputTables['cmap'] = $cmapTable;
        $outputTables['glyf'] = $glyfResult['Bytes'];
        $outputTables['head'] = $headTable;
        $outputTables['hhea'] = $hheaTable;
        $outputTables['hmtx'] = $hmtxTable;
        $outputTables['loca'] = $locaTable;
        $outputTables['maxp'] = $maxpTable;
        $outputTables['name'] = $this->BuildNameTable($subsetFontName);
        $outputTables['post'] = $this->BuildPostTable();

        return $this->BuildFontFile($outputTables);
    }

    private function CopyTableIfPresent(&$outputTables, $tag)
    {
        if (isset($this->tables[$tag])) {
            $record = $this->tables[$tag];
            $outputTables[$tag] = substr($this->fontBytes, $record['Offset'], $record['Length']);
        }
    }

    private function BuildGlyphClosure($initialGlyphIds)
    {
        $seen = array();
        $queue = new \SplQueue();
        $queue->enqueue(0);
        for ($i = 0; $i < count($initialGlyphIds); $i++) {
            $queue->enqueue($initialGlyphIds[$i]);
        }

        while (!$queue->isEmpty()) {
            $glyphId = $queue->dequeue();
            if ($glyphId >= $this->numGlyphs || isset($seen[$glyphId])) {
                continue;
            }

            $seen[$glyphId] = true;
            $bytes = $this->GetGlyphBytes($glyphId);
            if ($bytes === null || strlen($bytes) < 10) {
                continue;
            }

            $contourCount = self::ReadShort($bytes, 0);
            if ($contourCount >= 0) {
                continue;
            }

            $offset = 10;
            while ($offset + 4 <= strlen($bytes)) {
                $flags = self::ReadUShort($bytes, $offset);
                $componentGlyphId = self::ReadUShort($bytes, $offset + 2);
                $queue->enqueue($componentGlyphId);
                $offset += 4;
                $offset += ($flags & self::CompoundGlyphFlagArg1And2AreWords) !== 0 ? 4 : 2;
                if (($flags & self::CompoundGlyphFlagWeHaveAScale) !== 0) {
                    $offset += 2;
                } elseif (($flags & self::CompoundGlyphFlagWeHaveAnXAndYScale) !== 0) {
                    $offset += 4;
                } elseif (($flags & self::CompoundGlyphFlagWeHaveATwoByTwo) !== 0) {
                    $offset += 8;
                }

                if (($flags & self::CompoundGlyphFlagMoreComponents) === 0) {
                    break;
                }
            }
        }

        $ordered = array_keys($seen);
        sort($ordered, SORT_NUMERIC);
        if (count($ordered) === 0 || $ordered[0] !== 0) {
            array_unshift($ordered, 0);
        }

        return $ordered;
    }

    private function BuildGlyfTable($closure, $oldToNew)
    {
        $locaOffsets = array();
        $output = '';
        for ($i = 0; $i < count($closure); $i++) {
            $locaOffsets[] = strlen($output);
            $originalGlyphId = $closure[$i];
            $bytes = $this->GetGlyphBytes($originalGlyphId);
            if ($bytes !== null && strlen($bytes) > 0) {
                $bytes = $this->RewriteCompoundGlyphReferences($bytes, $oldToNew);
                $output .= $bytes;
                while ((strlen($output) & 3) !== 0) {
                    $output .= "\0";
                }
            }
        }

        $locaOffsets[] = strlen($output);
        return array('Bytes' => $output, 'LocaOffsets' => $locaOffsets);
    }

    private function RewriteCompoundGlyphReferences($glyphBytes, $oldToNew)
    {
        $copy = $glyphBytes;
        if (strlen($copy) < 10) {
            return $copy;
        }

        $contourCount = self::ReadShort($copy, 0);
        if ($contourCount >= 0) {
            return self::StripSimpleGlyphInstructions($copy, $contourCount);
        }

        $output = substr($copy, 0, 10);
        $offset = 10;
        $len = strlen($copy);
        while ($offset + 4 <= $len) {
            $flags = self::ReadUShort($copy, $offset);
            $flags = $flags & ~self::CompoundGlyphFlagWeHaveInstructions;
            $oldGlyphId = self::ReadUShort($copy, $offset + 2);
            if (!isset($oldToNew[$oldGlyphId])) {
                throw new \RuntimeException('Compound glyph referenced an unmapped component glyph.');
            }

            $newGlyphId = $oldToNew[$oldGlyphId];
            $output .= self::WriteUShortStr($flags);
            $output .= self::WriteUShortStr($newGlyphId);
            $offset += 4;
            $argumentBytes = ($flags & self::CompoundGlyphFlagArg1And2AreWords) !== 0 ? 4 : 2;
            $output .= substr($copy, $offset, $argumentBytes);
            $offset += $argumentBytes;
            if (($flags & self::CompoundGlyphFlagWeHaveAScale) !== 0) {
                $output .= substr($copy, $offset, 2);
                $offset += 2;
            } elseif (($flags & self::CompoundGlyphFlagWeHaveAnXAndYScale) !== 0) {
                $output .= substr($copy, $offset, 4);
                $offset += 4;
            } elseif (($flags & self::CompoundGlyphFlagWeHaveATwoByTwo) !== 0) {
                $output .= substr($copy, $offset, 8);
                $offset += 8;
            }

            if (($flags & self::CompoundGlyphFlagMoreComponents) === 0) {
                break;
            }
        }

        return $output;
    }

    private function BuildLocaTable($offsets)
    {
        $output = '';
        for ($i = 0; $i < count($offsets); $i++) {
            $output .= self::WriteUInt32Str($offsets[$i]);
        }

        return $output;
    }

    private function BuildHmtxTable($closure)
    {
        $output = '';
        for ($i = 0; $i < count($closure); $i++) {
            $metric = $this->MetricForGlyph($closure[$i]);
            $output .= self::WriteUShortStr($metric['AdvanceWidth']);
            $output .= self::WriteShortStr($metric['LeftSideBearing']);
        }

        return $output;
    }

    private function PatchMaxpTable($numGlyphs)
    {
        $bytes = $this->TableBytes('maxp');
        $copy = $bytes;
        self::WriteUShort($copy, 4, $numGlyphs);
        return $copy;
    }

    private function PatchHheaTable($numberOfHMetrics)
    {
        $bytes = $this->TableBytes('hhea');
        $copy = $bytes;
        self::WriteUShort($copy, 34, $numberOfHMetrics);
        return $copy;
    }

    private function PatchHeadTable()
    {
        $bytes = $this->TableBytes('head');
        $copy = $bytes;
        self::WriteUInt32($copy, 8, 0);
        self::WriteShort($copy, 50, 1);
        return $copy;
    }

    private function BuildCmapTable($codeToGlyphId, $glyphIdToUnicode, $oldToNew)
    {
        $mappings = array();
        $seen = array();
        foreach ($codeToGlyphId as $code => $glyphId) {
            $unicode = isset($glyphIdToUnicode[$glyphId]) ? $glyphIdToUnicode[$glyphId] : $code;
            if (!isset($oldToNew[$glyphId])) {
                continue;
            }

            $newGlyphId = $oldToNew[$glyphId];
            if (!isset($seen[$unicode])) {
                $seen[$unicode] = true;
                $mappings[] = array('Unicode' => $unicode, 'GlyphId' => $newGlyphId);
            }
        }

        usort($mappings, function ($a, $b) {
            if ($a['Unicode'] === $b['Unicode']) return 0;
            return $a['Unicode'] < $b['Unicode'] ? -1 : 1;
        });

        return $this->BuildFormat4Cmap($mappings);
    }

    private function BuildFormat4Cmap($mappings)
    {
        $segmentCount = count($mappings) + 1;
        $segCountX2 = $segmentCount * 2;
        $searchRange = 2;
        $entrySelector = 0;
        while ($searchRange * 2 <= $segCountX2) {
            $searchRange *= 2;
            $entrySelector++;
        }

        $rangeShift = $segCountX2 - $searchRange;
        $format4 = '';
        $format4 .= self::WriteUShortStr(4);
        $format4 .= self::WriteUShortStr(0);
        $format4 .= self::WriteUShortStr(0);
        $format4 .= self::WriteUShortStr($segCountX2);
        $format4 .= self::WriteUShortStr($searchRange);
        $format4 .= self::WriteUShortStr($entrySelector);
        $format4 .= self::WriteUShortStr($rangeShift);

        for ($i = 0; $i < count($mappings); $i++) {
            $format4 .= self::WriteUShortStr($mappings[$i]['Unicode']);
        }

        $format4 .= self::WriteUShortStr(0xFFFF);
        $format4 .= self::WriteUShortStr(0);

        for ($i = 0; $i < count($mappings); $i++) {
            $format4 .= self::WriteUShortStr($mappings[$i]['Unicode']);
        }

        $format4 .= self::WriteUShortStr(0xFFFF);
        for ($i = 0; $i < count($mappings); $i++) {
            $delta = $mappings[$i]['GlyphId'] - $mappings[$i]['Unicode'];
            while ($delta < 0) {
                $delta += 65536;
            }

            $format4 .= self::WriteUShortStr($delta);
        }

        $format4 .= self::WriteUShortStr(1);
        for ($i = 0; $i < $segmentCount; $i++) {
            $format4 .= self::WriteUShortStr(0);
        }

        $len = strlen($format4);
        self::WriteUShort($format4, 2, $len);
        return $this->BuildCmapTableHeader($format4);
    }

    private function BuildCmapTableHeader($format4)
    {
        $output = '';
        $output .= self::WriteUShortStr(0);
        $output .= self::WriteUShortStr(1);
        $output .= self::WriteUShortStr(3);
        $output .= self::WriteUShortStr(1);
        $output .= self::WriteUInt32Str(12);
        $output .= $format4;
        return $output;
    }

    private function BuildPostTable()
    {
        $original = $this->TableBytes('post');
        $output = '';
        $output .= self::WriteUInt32Str(0x00030000);
        $output .= self::WriteUInt32Str(self::ReadUInt($original, 4));
        $output .= self::WriteShortStr(self::ReadShort($original, 8));
        $output .= self::WriteShortStr(self::ReadShort($original, 10));
        $output .= self::WriteUInt32Str(self::ReadUInt($original, 12));
        $output .= self::WriteUInt32Str(self::ReadUInt($original, 16));
        $output .= self::WriteUInt32Str(self::ReadUInt($original, 20));
        $output .= self::WriteUInt32Str(self::ReadUInt($original, 24));
        $output .= self::WriteUInt32Str(self::ReadUInt($original, 28));
        return $output;
    }

    private function BuildOs2Table()
    {
        $original = $this->TableBytes('OS/2');
        $copy = $original;
        if (strlen($copy) >= 62) {
            self::WriteUShort($copy, 62, 0);
        }

        return $copy;
    }

    private function BuildNameTable($subsetFontName)
    {
        $records = array();
        $this->AddNameRecord($records, 3, 1, 1033, 1, 'Calibri');
        $this->AddNameRecord($records, 3, 1, 1033, 2, 'Regular');
        $this->AddNameRecord($records, 3, 1, 1033, 4, $subsetFontName);
        $this->AddNameRecord($records, 3, 1, 1033, 6, $subsetFontName);

        $output = '';
        $output .= self::WriteUShortStr(0);
        $output .= self::WriteUShortStr(count($records));
        $output .= self::WriteUShortStr(6 + count($records) * 12);
        $runningOffset = 0;
        for ($i = 0; $i < count($records); $i++) {
            $output .= self::WriteUShortStr($records[$i]['PlatformId']);
            $output .= self::WriteUShortStr($records[$i]['EncodingId']);
            $output .= self::WriteUShortStr($records[$i]['LanguageId']);
            $output .= self::WriteUShortStr($records[$i]['NameId']);
            $output .= self::WriteUShortStr(strlen($records[$i]['Bytes']));
            $output .= self::WriteUShortStr($runningOffset);
            $runningOffset += strlen($records[$i]['Bytes']);
        }

        for ($i = 0; $i < count($records); $i++) {
            $output .= $records[$i]['Bytes'];
        }

        return $output;
    }

    private function AddNameRecord(&$records, $platformId, $encodingId, $languageId, $nameId, $text)
    {
        $records[] = array(
            'PlatformId' => $platformId,
            'EncodingId' => $encodingId,
            'LanguageId' => $languageId,
            'NameId' => $nameId,
            'Bytes' => mb_convert_encoding($text, 'UTF-16BE', 'UTF-8'),
        );
    }

    private function BuildFontFile($tables)
    {
        $tags = array_keys($tables);
        sort($tags, SORT_STRING);
        $numTables = count($tags);
        $entrySelector = 0;
        $searchRange = 1;
        while ($searchRange * 2 <= $numTables) {
            $searchRange *= 2;
            $entrySelector++;
        }

        $searchRange *= 16;
        $rangeShift = $numTables * 16 - $searchRange;

        $output = '';
        $output .= self::WriteUInt32Str(0x00010000);
        $output .= self::WriteUShortStr($numTables);
        $output .= self::WriteUShortStr($searchRange);
        $output .= self::WriteUShortStr($entrySelector);
        $output .= self::WriteUShortStr($rangeShift);

        $recordPosition = strlen($output);
        $output .= str_repeat("\0", $numTables * 16);

        $records = array();
        for ($i = 0; $i < count($tags); $i++) {
            $tag = $tags[$i];
            $bytes = $tables[$tag];
            while ((strlen($output) & 3) !== 0) {
                $output .= "\0";
            }

            $offset = strlen($output);
            $output .= $bytes;
            while ((strlen($output) & 3) !== 0) {
                $output .= "\0";
            }

            $records[] = array(
                'Tag' => $tag,
                'Checksum' => self::CalculateChecksum($bytes),
                'Offset' => $offset,
                'Length' => strlen($bytes),
            );
        }

        $fontBytes = $output;
        for ($i = 0; $i < count($records); $i++) {
            $entryOffset = $recordPosition + $i * 16;
            $fontBytes = substr_replace($fontBytes, $records[$i]['Tag'], $entryOffset, 4);
            self::WriteUInt32($fontBytes, $entryOffset + 4, $records[$i]['Checksum']);
            self::WriteUInt32($fontBytes, $entryOffset + 8, $records[$i]['Offset']);
            self::WriteUInt32($fontBytes, $entryOffset + 12, $records[$i]['Length']);
        }

        $headRecord = null;
        for ($i = 0; $i < count($records); $i++) {
            if ($records[$i]['Tag'] === 'head') {
                $headRecord = $records[$i];
                break;
            }
        }

        self::WriteUInt32($fontBytes, $headRecord['Offset'] + 8, 0);
        $checksum = self::CalculateChecksum($fontBytes);
        $adjustment = 0xB1B0AFBA - $checksum;
        if ($adjustment < 0) {
            $adjustment += 0x100000000;
        }

        self::WriteUInt32($fontBytes, $headRecord['Offset'] + 8, $adjustment);
        return $fontBytes;
    }

    private function TableBytes($tag)
    {
        if (!isset($this->tables[$tag])) {
            throw new \RuntimeException('TrueType subset source was missing required table ' . $tag . '.');
        }

        $record = $this->tables[$tag];
        return substr($this->fontBytes, $record['Offset'], $record['Length']);
    }

    private function MetricForGlyph($glyphId)
    {
        if ($glyphId >= count($this->metrics)) {
            return $this->metrics[count($this->metrics) - 1];
        }

        return $this->metrics[$glyphId];
    }

    private function GetGlyphBytes($glyphId)
    {
        if ($glyphId >= $this->numGlyphs) {
            return null;
        }

        $start = $this->glyphOffsets[$glyphId];
        $end = $this->glyphOffsets[$glyphId + 1];
        if ($end <= $start) {
            return '';
        }

        $glyf = $this->TableBytes('glyf');
        return substr($glyf, $start, $end - $start);
    }

    private function ParseGlyphOffsets()
    {
        $loca = $this->TableBytes('loca');
        $offsets = array();
        if ($this->indexToLocFormat === 0) {
            for ($i = 0; $i <= $this->numGlyphs; $i++) {
                $offsets[$i] = self::ReadUShort($loca, $i * 2) * 2;
            }
        } else {
            for ($i = 0; $i <= $this->numGlyphs; $i++) {
                $offsets[$i] = (int)self::ReadUInt($loca, $i * 4);
            }
        }

        return $offsets;
    }

    private function ParseHorizontalMetrics()
    {
        $bytes = $this->TableBytes('hmtx');
        $metrics = array();
        $lastAdvance = 0;
        $lastLsb = 0;
        for ($i = 0; $i < $this->numGlyphs; $i++) {
            if ($i < $this->numberOfHMetrics) {
                $metricOffset = $i * 4;
                $lastAdvance = self::ReadUShort($bytes, $metricOffset);
                $lastLsb = self::ReadShort($bytes, $metricOffset + 2);
            } else {
                $lsbOffset = $this->numberOfHMetrics * 4 + ($i - $this->numberOfHMetrics) * 2;
                $lastLsb = self::ReadShort($bytes, $lsbOffset);
            }

            $metrics[$i] = array('AdvanceWidth' => $lastAdvance, 'LeftSideBearing' => $lastLsb);
        }

        return $metrics;
    }

    private static function ParseTableDirectory($fontBytes)
    {
        $numTables = self::ReadUShort($fontBytes, 4);
        $tables = array();
        for ($i = 0; $i < $numTables; $i++) {
            $offset = 12 + $i * 16;
            $tag = substr($fontBytes, $offset, 4);
            $tables[$tag] = array(
                'Tag' => $tag,
                'Checksum' => self::ReadUInt($fontBytes, $offset + 4),
                'Offset' => self::ReadUInt($fontBytes, $offset + 8),
                'Length' => self::ReadUInt($fontBytes, $offset + 12),
            );
        }

        return $tables;
    }

    private static function StripSimpleGlyphInstructions($glyphBytes, $contourCount)
    {
        if ($contourCount < 0) {
            return $glyphBytes;
        }

        $contourArrayBytes = $contourCount * 2;
        $instructionLengthOffset = 10 + $contourArrayBytes;
        if ($instructionLengthOffset + 2 > strlen($glyphBytes)) {
            return $glyphBytes;
        }

        $instructionLength = self::ReadUShort($glyphBytes, $instructionLengthOffset);
        $flagsOffset = $instructionLengthOffset + 2 + $instructionLength;
        if ($flagsOffset > strlen($glyphBytes)) {
            return $glyphBytes;
        }

        $output = substr($glyphBytes, 0, $instructionLengthOffset);
        $output .= self::WriteUShortStr(0);
        $output .= substr($glyphBytes, $flagsOffset);
        return $output;
    }

    private static function ReadUShort($bytes, $offset)
    {
        return (ord($bytes[$offset]) << 8) | ord($bytes[$offset + 1]);
    }

    private static function ReadShort($bytes, $offset)
    {
        $val = self::ReadUShort($bytes, $offset);
        if ($val >= 0x8000) {
            $val -= 0x10000;
        }

        return $val;
    }

    private static function ReadUInt($bytes, $offset)
    {
        return ((ord($bytes[$offset]) << 24) | (ord($bytes[$offset + 1]) << 16)
            | (ord($bytes[$offset + 2]) << 8) | ord($bytes[$offset + 3])) & 0xFFFFFFFF;
    }

    private static function WriteUShort(&$bytes, $offset, $value)
    {
        $bytes[$offset] = chr(($value >> 8) & 0xFF);
        $bytes[$offset + 1] = chr($value & 0xFF);
    }

    private static function WriteShort(&$bytes, $offset, $value)
    {
        self::WriteUShort($bytes, $offset, $value & 0xFFFF);
    }

    private static function WriteUInt32(&$bytes, $offset, $value)
    {
        $bytes[$offset] = chr(($value >> 24) & 0xFF);
        $bytes[$offset + 1] = chr(($value >> 16) & 0xFF);
        $bytes[$offset + 2] = chr(($value >> 8) & 0xFF);
        $bytes[$offset + 3] = chr($value & 0xFF);
    }

    private static function WriteUShortStr($value)
    {
        return chr(($value >> 8) & 0xFF) . chr($value & 0xFF);
    }

    private static function WriteShortStr($value)
    {
        return self::WriteUShortStr($value & 0xFFFF);
    }

    private static function WriteUInt32Str($value)
    {
        return chr(($value >> 24) & 0xFF) . chr(($value >> 16) & 0xFF)
            . chr(($value >> 8) & 0xFF) . chr($value & 0xFF);
    }

    private static function CalculateChecksum($bytes)
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
