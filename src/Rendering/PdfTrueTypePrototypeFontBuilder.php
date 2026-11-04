<?php

namespace Aspose\Cells_FOSS\Rendering;

final class PdfTrueTypePrototypeFontBuilder
{
    /**
     * @param PdfTrueTypePrototypeUsage[] $usages
     * @param int &$nextObjectId
     * @return array<int, string>
     */
    public function BuildObjects($usages, &$nextObjectId)
    {
        $objects = array();
        for ($i = 0; $i < count($usages); $i++) {
            $usage = $usages[$i];
            $usage->ResourceName = 'TTF' . ($i + 1);
            $usage->SubsetFontName = self::BuildSubsetFontName($usage, $i);
            $usage->FontFileObjectId = $nextObjectId++;
            $usage->FontDescriptorObjectId = $nextObjectId++;
            $usage->ToUnicodeObjectId = $nextObjectId++;
            $usage->FontObjectId = $nextObjectId++;

            $subsetBytes = $usage->BuildSubsetFontBytes();
            $objects[$usage->FontFileObjectId] = self::BuildFontFileStreamObject($usage->FontFileObjectId, $subsetBytes);
            $objects[$usage->FontDescriptorObjectId] = self::BuildFontDescriptorObject($usage);
            $objects[$usage->ToUnicodeObjectId] = self::BuildStreamObject($usage->ToUnicodeObjectId, self::BuildToUnicodeContent($usage));
            $objects[$usage->FontObjectId] = self::BuildFontObject($usage);
        }

        return $objects;
    }

    private static function BuildFontObject($usage)
    {
        $f = __NAMESPACE__ . '\PdfNumbers::threeDecimals';
        $builder = '';
        $builder .= $usage->FontObjectId . " 0 obj\n<< /Type /Font\n/Subtype /TrueType\n/Name /";
        $builder .= $usage->ResourceName;
        $builder .= "\n/BaseFont /";
        $builder .= $usage->SubsetFontName;
        $builder .= "\n/Encoding /WinAnsiEncoding\n/FontDescriptor ";
        $builder .= $usage->FontDescriptorObjectId . " 0 R\n/FirstChar ";
        $builder .= $usage->FirstChar();
        $builder .= "\n/LastChar ";
        $builder .= $usage->LastChar();
        $builder .= "\n/Widths [";
        $codes = $usage->OrderedCodes();
        $first = $usage->FirstChar();
        $last = $usage->LastChar();
        for ($code = $first; $code <= $last; $code++) {
            if ($code > $first) {
                $builder .= ' ';
            }

            $builder .= $f($usage->WidthForCode($code));
        }

        $builder .= "]\n/ToUnicode ";
        $builder .= $usage->ToUnicodeObjectId . " 0 R\n>>\nendobj\n";
        return $builder;
    }

    private static function BuildFontDescriptorObject($usage)
    {
        $f = __NAMESPACE__ . '\PdfNumbers::threeDecimals';
        $builder = '';
        $builder .= $usage->FontDescriptorObjectId . " 0 obj\n<< /Type /FontDescriptor\n/FontName /";
        $builder .= $usage->SubsetFontName;
        $builder .= "\n/Flags 32\n/ItalicAngle 0\n/Ascent 810\n/Descent -232\n/CapHeight 810\n/StemV 44\n/FontBBox [";
        if ($usage->HasBounds()) {
            $bounds = $usage->FontBounds();
            $builder .= $f($bounds->Left) . ' ' . $f($bounds->Top) . ' '
                . $f($bounds->Right) . ' ' . $f($bounds->Bottom);
        } else {
            $builder .= '0 0 0 0';
        }

        $builder .= "]\n/FontFile2 ";
        $builder .= $usage->FontFileObjectId . " 0 R\n>>\nendobj\n";
        return $builder;
    }

    private static function BuildToUnicodeContent($usage)
    {
        $ordered = $usage->OrderedCodes();
        $builder = '';
        $builder .= "/CIDInit /ProcSet findresource begin\n";
        $builder .= "12 dict begin\nbegincmap\n";
        $builder .= "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n";
        $builder .= "/CMapName /";
        $builder .= $usage->ResourceName;
        $builder .= " def\n/CMapType 2 def\n";
        $builder .= "1 begincodespacerange\n<00> <FF>\nendcodespacerange\n";
        $builder .= count($ordered) . " beginbfchar\n";
        for ($i = 0; $i < count($ordered); $i++) {
            $code = $ordered[$i];
            $builder .= '<' . sprintf('%02X', $code) . '> <00' . sprintf('%02X', $code) . ">\n";
        }

        $builder .= "endbfchar\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n";
        return $builder;
    }

    private static function BuildSubsetFontName($usage, $usageIndex)
    {
        $prefixValue = $usageIndex + 1;
        $prefixChars = array();
        for ($i = 5; $i >= 0; $i--) {
            $prefixChars[$i] = chr(ord('A') + ($prefixValue % 26));
            $prefixValue = intdiv($prefixValue, 26);
        }

        ksort($prefixChars);
        return implode('', $prefixChars) . '+' . $usage->BaseFontName();
    }

    private static function BuildStreamObject($objectId, $streamBytes)
    {
        $compressed = gzcompress($streamBytes, 9);
        $header = $objectId . " 0 obj\n<< /Filter /FlateDecode\n/Length "
            . strlen($compressed) . " >>\nstream\n";
        $footer = "\nendstream\nendobj\n";
        return $header . $compressed . $footer;
    }

    private static function BuildFontFileStreamObject($objectId, $streamBytes)
    {
        $compressed = gzcompress($streamBytes, 9);
        $header = $objectId . " 0 obj\n<< /Filter /FlateDecode\n/Length "
            . strlen($compressed) . "\n/Length1 " . strlen($streamBytes) . " >>\nstream\n";
        $footer = "\nendstream\nendobj\n";
        return $header . $compressed . $footer;
    }
}
