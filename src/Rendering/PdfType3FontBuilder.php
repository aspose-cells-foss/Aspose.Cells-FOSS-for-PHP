<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Skia\SKRect;

final class PdfType3FontBuilder
{
    /**
     * @param PdfFontUsageRegistry $registry
     * @param int &$nextObjectId
     * @return array<int, string>
     */
    public function BuildObjects($registry, &$nextObjectId)
    {
        $objects = array();
        $usages = &$registry->Usages();

        for ($i = 0; $i < count($usages); $i++) {
            $usage = $usages[$i];
            $subsets = &$usage->Subsets();
            for ($s = 0; $s < count($subsets); $s++) {
                $subset = $subsets[$s];
                $glyphObjectIds = array();
                $fontBounds = SKRect::Empty();
                $fontBoundsInitialized = false;

                for ($t = 0; $t < count($subset->Tokens); $t++) {
                    $token = $subset->Tokens[$t];
                    $width = $usage->GetTokenWidth($token);
                    $path = $usage->BuildTokenPath($token);
                    $content = self::BuildCharProcContent($path, $width);
                    $objectId = $nextObjectId++;
                    $glyphObjectIds[$token] = $objectId;
                    $objects[$objectId] = self::BuildStreamObject($objectId, $content);

                    if ($path !== null && !$path->IsEmpty) {
                        $bounds = $path->Bounds;
                        if (!$fontBoundsInitialized) {
                            $fontBounds = $bounds;
                            $fontBoundsInitialized = true;
                        } else {
                            $fontBounds = SKRect::Union($fontBounds, $bounds);
                        }
                    }

                    if ($path !== null) {
                        $path->Dispose();
                    }
                }

                $subset->ToUnicodeObjectId = $nextObjectId++;
                $objects[$subset->ToUnicodeObjectId] = self::BuildStreamObject(
                    $subset->ToUnicodeObjectId,
                    self::BuildToUnicodeContent($subset)
                );

                $subset->UsePositiveYFontMatrix = $usage->UsePositiveYFontMatrix();
                $subset->FontObjectId = $nextObjectId++;
                $objects[$subset->FontObjectId] = self::BuildFontObject($subset, $usage, $glyphObjectIds, $fontBounds, $fontBoundsInitialized);
            }
        }

        return $objects;
    }

    private static function BuildFontObject($subset, $usage, $glyphObjectIds, $fontBounds, $hasBounds)
    {
        $f = __NAMESPACE__ . '\PdfNumbers::threeDecimals';
        $builder = '';
        $builder .= $subset->FontObjectId . " 0 obj\n<< /Type /Font\n/Subtype /Type3\n/Name /";
        $builder .= $subset->ResourceName;
        $builder .= "\n/FontBBox [";
        if ($hasBounds) {
            $builder .= $f($fontBounds->Left) . ' ' . $f($fontBounds->Top) . ' '
                . $f($fontBounds->Right) . ' ' . $f($fontBounds->Bottom);
        } else {
            $builder .= '0 0 0 0';
        }

        $builder .= "]\n/FontMatrix ";
        $builder .= self::BuildFontMatrix($subset->UsePositiveYFontMatrix);
        $builder .= "\n/CharProcs <<";

        for ($i = 0; $i < count($subset->Tokens); $i++) {
            $token = $subset->Tokens[$i];
            $code = $subset->TokenCodes[$token];
            $builder .= " /g" . $code . ' ' . $glyphObjectIds[$token] . " 0 R";
        }

        $builder .= " >>\n/Encoding << /Type /Encoding /Differences [1";
        for ($i = 0; $i < count($subset->Tokens); $i++) {
            $token = $subset->Tokens[$i];
            $code = $subset->TokenCodes[$token];
            $builder .= " /g" . $code;
        }

        $builder .= "] >>\n/FirstChar 1\n/LastChar ";
        $builder .= count($subset->Tokens);
        $builder .= "\n/Widths [";
        for ($i = 0; $i < count($subset->Tokens); $i++) {
            if ($i > 0) {
                $builder .= ' ';
            }

            $builder .= $f($usage->GetTokenWidth($subset->Tokens[$i]));
        }

        $builder .= "]\n/Resources << >>\n/ToUnicode ";
        $builder .= $subset->ToUnicodeObjectId . " 0 R\n>>\nendobj\n";
        return $builder;
    }

    private static function BuildCharProcContent($path, $width)
    {
        $f = __NAMESPACE__ . '\PdfNumbers::threeDecimals';
        $builder = '';
        if ($path === null || $path->IsEmpty) {
            $builder .= $f($width) . " 0 d0\n";
            return $builder;
        }

        $bounds = $path->Bounds;
        $builder .= $f($width) . ' 0 ';
        $builder .= $f($bounds->Left) . ' ' . $f($bounds->Top) . ' '
            . $f($bounds->Right) . ' ' . $f($bounds->Bottom);
        $builder .= " d1\n";
        $builder .= PdfPathSerializer::Serialize($path);
        $builder .= "f\n";
        return $builder;
    }

    private static function BuildToUnicodeContent($subset)
    {
        $ranges = array();
        $singles = array();
        self::BuildToUnicodeMappings($subset, $ranges, $singles);

        $builder = '';
        $builder .= "/CIDInit /ProcSet findresource begin\n";
        $builder .= "12 dict begin\nbegincmap\n";
        $builder .= "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n";
        $builder .= "/CMapName /";
        $builder .= $subset->ResourceName;
        $builder .= " def\n/CMapType 2 def\n";
        $builder .= "1 begincodespacerange\n<00> <FF>\nendcodespacerange\n";

        if (count($ranges) > 0) {
            $builder .= count($ranges) . " beginbfrange\n";
            for ($i = 0; $i < count($ranges); $i++) {
                $builder .= '<' . sprintf('%02X', $ranges[$i]['StartCode']) . '> <'
                    . sprintf('%02X', $ranges[$i]['EndCode']) . '> <'
                    . sprintf('%04X', $ranges[$i]['StartUnicode']) . ">\n";
            }

            $builder .= "endbfrange\n";
        }

        if (count($singles) > 0) {
            $builder .= count($singles) . " beginbfchar\n";
            for ($i = 0; $i < count($singles); $i++) {
                $builder .= '<' . sprintf('%02X', $singles[$i]['Code']) . '> <'
                    . self::ToHexUnicode($singles[$i]['Token']) . ">\n";
            }

            $builder .= "endbfchar\n";
        }

        $builder .= "endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n";
        return $builder;
    }

    private static function BuildToUnicodeMappings($subset, &$ranges, &$singles)
    {
        $index = 0;
        $tokenCount = count($subset->Tokens);
        while ($index < $tokenCount) {
            $token = $subset->Tokens[$index];
            $code = $subset->TokenCodes[$token];
            $unicode = self::SingleUnicodeScalar($token);
            if ($unicode < 0) {
                $singles[] = array('Code' => $code, 'Token' => $token);
                $index++;
                continue;
            }

            $endIndex = $index;
            $endCode = $code;
            $endUnicode = $unicode;
            while ($endIndex + 1 < $tokenCount) {
                $nextToken = $subset->Tokens[$endIndex + 1];
                $nextCode = $subset->TokenCodes[$nextToken];
                $nextUnicode = self::SingleUnicodeScalar($nextToken);
                if ($nextUnicode < 0) {
                    break;
                }

                if ($nextCode !== $endCode + 1 || $nextUnicode !== $endUnicode + 1) {
                    break;
                }

                $endIndex++;
                $endCode = $nextCode;
                $endUnicode = $nextUnicode;
            }

            if ($endIndex > $index) {
                $ranges[] = array(
                    'StartCode' => $code,
                    'EndCode' => $endCode,
                    'StartUnicode' => $unicode,
                );
                $index = $endIndex + 1;
                continue;
            }

            $singles[] = array('Code' => $code, 'Token' => $token);
            $index++;
        }
    }

    private static function SingleUnicodeScalar($token)
    {
        if ($token === null || $token === '') {
            return -1;
        }

        $len = mb_strlen($token, 'UTF-8');
        if ($len === 1) {
            $cp = mb_ord($token, 'UTF-8');
            return $cp !== false ? $cp : -1;
        }

        return -1;
    }

    private static function ToHexUnicode($text)
    {
        $result = '';
        $len = mb_strlen($text, 'UTF-8');
        for ($i = 0; $i < $len; $i++) {
            $ch = mb_substr($text, $i, 1, 'UTF-8');
            $cp = mb_ord($ch, 'UTF-8');
            $result .= sprintf('%04X', $cp);
        }

        return $result;
    }

    private static function BuildFontMatrix($usePositiveYFontMatrix)
    {
        if ($usePositiveYFontMatrix) {
            return '[0.001 0 0 0.001 0 0]';
        }

        return '[0.001 0 0 -0.001 0 0]';
    }

    private static function BuildStreamObject($objectId, $streamBytes)
    {
        $flatObject = self::BuildFlatStreamObject($objectId, $streamBytes);
        $compressedObject = self::BuildCompressedStreamObject($objectId, $streamBytes);
        if (strlen($flatObject) < strlen($compressedObject)) {
            return $flatObject;
        }

        return $compressedObject;
    }

    private static function BuildCompressedStreamObject($objectId, $streamBytes)
    {
        $compressed = gzcompress($streamBytes, 9);
        $header = $objectId . " 0 obj\n<< /Filter /FlateDecode\n/Length "
            . strlen($compressed) . " >>\nstream\n";
        $footer = "\nendstream\nendobj\n";
        return $header . $compressed . $footer;
    }

    private static function BuildFlatStreamObject($objectId, $streamBytes)
    {
        $header = $objectId . " 0 obj\n<< /Length "
            . strlen($streamBytes) . " >>\nstream\n";
        $footer = "\nendstream\nendobj\n";
        return $header . $streamBytes . $footer;
    }
}
