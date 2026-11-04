<?php

namespace Aspose\Cells_FOSS\Rendering;

final class PdfContentStreamOptimizer
{
    private static $decimalPattern = '/(?<![A-Za-z0-9_\/])[-+]?(?:\d+\.\d+|\.\d+)/';

    /**
     * @param string $originalPdf
     * @return string
     */
    public function Rewrite($originalPdf)
    {
        if ($originalPdf === null || strlen($originalPdf) === 0) {
            return $originalPdf;
        }

        $table = PdfObjectTable::Parse($originalPdf);
        $contentObjectIds = self::CollectPageContentObjectIds($table);
        if (count($contentObjectIds) === 0) {
            return $originalPdf;
        }

        $changed = false;
        for ($i = 0; $i < count($contentObjectIds); $i++) {
            $objectId = $contentObjectIds[$i];
            if (!isset($table->Objects[$objectId])) {
                continue;
            }

            $objectBytes = $table->Objects[$objectId];
            $rewritten = self::TryOptimizeStreamObject($objectId, $objectBytes);
            if ($rewritten === null) {
                continue;
            }

            $table->Objects[$objectId] = $rewritten;
            $changed = true;
        }

        if (!$changed) {
            return $originalPdf;
        }

        return self::RebuildPdf($table);
    }

    private static function CollectPageContentObjectIds($table)
    {
        $result = array();
        $seen = array();
        for ($i = 0; $i < count($table->PageObjectIds); $i++) {
            if (!isset($table->Objects[$table->PageObjectIds[$i]])) {
                continue;
            }

            $pageText = $table->Objects[$table->PageObjectIds[$i]];
            $contentIds = self::ParseContentObjectIds($pageText);
            for ($c = 0; $c < count($contentIds); $c++) {
                if (!isset($seen[$contentIds[$c]])) {
                    $seen[$contentIds[$c]] = true;
                    $result[] = $contentIds[$c];
                }
            }
        }

        return $result;
    }

    private static function ParseContentObjectIds($pageObjectText)
    {
        $result = array();
        $marker = '/Contents ';
        $index = strpos($pageObjectText, $marker);
        if ($index === false) {
            return $result;
        }

        $index += strlen($marker);
        if ($index >= strlen($pageObjectText)) {
            return $result;
        }

        if ($pageObjectText[$index] === '[') {
            $endArray = strpos($pageObjectText, ']', $index);
            if ($endArray === false) {
                return $result;
            }

            $items = preg_split('/\s+/', trim(substr($pageObjectText, $index + 1, $endArray - $index - 1)));
            for ($i = 0; $i + 2 < count($items); $i++) {
                if (ctype_digit($items[$i]) && $items[$i + 1] === '0' && $items[$i + 2] === 'R') {
                    $result[] = (int)$items[$i];
                }
            }

            return $result;
        }

        $end = strpos($pageObjectText, "\n", $index);
        if ($end === false) {
            $end = strlen($pageObjectText);
        }

        $single = preg_split('/\s+/', trim(substr($pageObjectText, $index, $end - $index)));
        if (count($single) >= 3) {
            if (ctype_digit($single[0]) && $single[1] === '0' && $single[2] === 'R') {
                $result[] = (int)$single[0];
            }
        }

        return $result;
    }

    private static function TryOptimizeStreamObject($objectId, $objectBytes)
    {
        $objectText = $objectBytes;
        if (strpos($objectText, '/Filter /FlateDecode') === false) {
            return null;
        }

        $streamMarker = "stream\n";
        $streamIndex = strpos($objectText, $streamMarker);
        if ($streamIndex === false) {
            $streamMarker = "stream\r\n";
            $streamIndex = strpos($objectText, $streamMarker);
            if ($streamIndex === false) {
                return null;
            }
        }

        $streamStart = $streamIndex + strlen($streamMarker);
        $endStreamIndex = strpos($objectText, "\nendstream", $streamStart);
        $endStreamMarkerLength = 1;
        if ($endStreamIndex === false) {
            $endStreamIndex = strpos($objectText, "\r\nendstream", $streamStart);
            $endStreamMarkerLength = 2;
            if ($endStreamIndex === false) {
                return null;
            }
        }

        $compressed = substr($objectBytes, $streamStart, $endStreamIndex - $streamStart);

        $decompressed = @gzuncompress($compressed);
        if ($decompressed === false) {
            return null;
        }

        $optimized = self::OptimizeContentText($decompressed);
        if (strlen($optimized) >= strlen($decompressed)) {
            return null;
        }

        $recompressed = gzcompress($optimized, 9);
        $header = substr($objectText, 0, $streamIndex);
        $header = self::ReplaceLength($header, strlen($recompressed));
        $footer = substr($objectText, $endStreamIndex + $endStreamMarkerLength);

        return $header . $streamMarker . $recompressed . $footer;
    }

    private static function ReplaceLength($header, $newLength)
    {
        $marker = '/Length ';
        $index = strpos($header, $marker);
        if ($index === false) {
            return $header;
        }

        $start = $index + strlen($marker);
        $end = $start;
        while ($end < strlen($header) && ctype_digit($header[$end])) {
            $end++;
        }

        return substr($header, 0, $start) . $newLength . substr($header, $end);
    }

    private static function OptimizeContentText($content)
    {
        $compact = preg_replace_callback(self::$decimalPattern, function ($match) {
            $value = (float)$match[0];
            return PdfNumbers::compact($value);
        }, $content);

        $compact = self::CompactWhitespace($compact);
        return self::RemoveRedundantGraphicsStateCommands($compact);
    }

    private static function RemoveRedundantGraphicsStateCommands($content)
    {
        if ($content === null || $content === '') {
            return $content;
        }

        $tokens = self::TokenizeContent($content);
        if (count($tokens) === 0) {
            return $content;
        }

        $output = array();
        $operands = array();
        $strokeRgb = null;
        $fillRgb = null;
        $strokeGray = null;
        $fillGray = null;
        $graphicsState = null;
        $strokeRgbStack = array();
        $fillRgbStack = array();
        $strokeGrayStack = array();
        $fillGrayStack = array();
        $graphicsStateStack = array();

        for ($i = 0; $i < count($tokens); $i++) {
            $token = $tokens[$i];
            if (!self::IsOperatorToken($token)) {
                $operands[] = $token;
                continue;
            }

            if ($token === 'q' && count($operands) === 0) {
                $output[] = $token;
                $strokeRgbStack[] = $strokeRgb;
                $fillRgbStack[] = $fillRgb;
                $strokeGrayStack[] = $strokeGray;
                $fillGrayStack[] = $fillGray;
                $graphicsStateStack[] = $graphicsState;
                continue;
            }

            if ($token === 'Q' && count($operands) === 0) {
                $output[] = $token;
                $lastIndex = count($strokeRgbStack) - 1;
                if ($lastIndex >= 0) {
                    $strokeRgb = $strokeRgbStack[$lastIndex];
                    $fillRgb = $fillRgbStack[$lastIndex];
                    $strokeGray = $strokeGrayStack[$lastIndex];
                    $fillGray = $fillGrayStack[$lastIndex];
                    $graphicsState = $graphicsStateStack[$lastIndex];
                    array_pop($strokeRgbStack);
                    array_pop($fillRgbStack);
                    array_pop($strokeGrayStack);
                    array_pop($fillGrayStack);
                    array_pop($graphicsStateStack);
                } else {
                    $strokeRgb = null;
                    $fillRgb = null;
                    $strokeGray = null;
                    $fillGray = null;
                    $graphicsState = null;
                }

                continue;
            }

            if ($token === 'RG' && count($operands) === 3) {
                $value = implode(' ', $operands);
                if ($strokeRgb !== $value) {
                    self::FlushOperation($output, $operands, $token);
                    $strokeRgb = $value;
                } else {
                    $operands = array();
                }

                continue;
            }

            if ($token === 'rg' && count($operands) === 3) {
                $value = implode(' ', $operands);
                if ($fillRgb !== $value) {
                    self::FlushOperation($output, $operands, $token);
                    $fillRgb = $value;
                } else {
                    $operands = array();
                }

                continue;
            }

            if ($token === 'G' && count($operands) === 1) {
                $value = $operands[0];
                if ($strokeGray !== $value) {
                    self::FlushOperation($output, $operands, $token);
                    $strokeGray = $value;
                } else {
                    $operands = array();
                }

                continue;
            }

            if ($token === 'g' && count($operands) === 1) {
                $value = $operands[0];
                if ($fillGray !== $value) {
                    self::FlushOperation($output, $operands, $token);
                    $fillGray = $value;
                } else {
                    $operands = array();
                }

                continue;
            }

            if ($token === 'gs' && count($operands) === 1) {
                $value = $operands[0];
                if ($graphicsState !== $value) {
                    self::FlushOperation($output, $operands, $token);
                    $graphicsState = $value;
                } else {
                    $operands = array();
                }

                continue;
            }

            self::FlushOperation($output, $operands, $token);
        }

        if (count($operands) > 0) {
            for ($i = 0; $i < count($operands); $i++) {
                $output[] = $operands[$i];
            }
        }

        return implode(' ', $output);
    }

    private static function TokenizeContent($content)
    {
        $tokens = array();
        $builder = '';
        $inLiteralString = false;
        $inHexString = false;
        $literalDepth = 0;
        $escaping = false;
        $len = strlen($content);

        for ($i = 0; $i < $len; $i++) {
            $ch = $content[$i];
            if ($inLiteralString) {
                $builder .= $ch;
                if ($escaping) {
                    $escaping = false;
                } elseif ($ch === '\\') {
                    $escaping = true;
                } elseif ($ch === '(') {
                    $literalDepth++;
                } elseif ($ch === ')') {
                    $literalDepth--;
                    if ($literalDepth <= 0) {
                        $inLiteralString = false;
                        self::AddToken($tokens, $builder);
                        $builder = '';
                    }
                }

                continue;
            }

            if ($inHexString) {
                $builder .= $ch;
                if ($ch === '>') {
                    $inHexString = false;
                    self::AddToken($tokens, $builder);
                    $builder = '';
                }

                continue;
            }

            if (ctype_space($ch)) {
                self::AddToken($tokens, $builder);
                $builder = '';
                continue;
            }

            if ($ch === '(') {
                self::AddToken($tokens, $builder);
                $builder = '';
                $builder .= $ch;
                $inLiteralString = true;
                $literalDepth = 1;
                $escaping = false;
                continue;
            }

            if ($ch === '<' && !self::IsDictionaryStart($content, $i)) {
                self::AddToken($tokens, $builder);
                $builder = '';
                $builder .= $ch;
                $inHexString = true;
                continue;
            }

            $builder .= $ch;
        }

        self::AddToken($tokens, $builder);
        return $tokens;
    }

    private static function AddToken(&$tokens, &$builder)
    {
        if ($builder === '') {
            return;
        }

        $tokens[] = $builder;
        $builder = '';
    }

    private static function IsOperatorToken($token)
    {
        static $operators = array(
            'q' => true, 'Q' => true, 'cm' => true, 'w' => true, 'J' => true, 'j' => true,
            'M' => true, 'd' => true, 'ri' => true, 'i' => true, 'gs' => true,
            'm' => true, 'l' => true, 'c' => true, 'v' => true, 'y' => true, 'h' => true,
            're' => true, 'S' => true, 's' => true, 'f' => true, 'F' => true, 'f*' => true,
            'B' => true, 'B*' => true, 'b' => true, 'b*' => true, 'n' => true,
            'W' => true, 'W*' => true, 'BT' => true, 'ET' => true, 'Tf' => true,
            'Tm' => true, 'Td' => true, 'TD' => true, 'Tj' => true, 'TJ' => true,
            'Tc' => true, 'Tw' => true, 'Tz' => true, 'TL' => true, 'Tr' => true,
            'Ts' => true, 'Do' => true, 'RG' => true, 'rg' => true, 'G' => true,
            'g' => true, 'K' => true, 'k' => true, 'CS' => true, 'cs' => true,
            'SC' => true, 'SCN' => true, 'sc' => true, 'scn' => true, 'sh' => true,
        );

        return isset($operators[$token]);
    }

    private static function FlushOperation(&$output, &$operands, $operatorToken)
    {
        for ($i = 0; $i < count($operands); $i++) {
            $output[] = $operands[$i];
        }

        $operands = array();
        $output[] = $operatorToken;
    }

    private static function CompactWhitespace($content)
    {
        if ($content === null || $content === '') {
            return $content;
        }

        $builder = '';
        $pendingSpace = false;
        $inLiteralString = false;
        $literalDepth = 0;
        $escaping = false;
        $inHexString = false;
        $inComment = false;
        $len = strlen($content);

        for ($i = 0; $i < $len; $i++) {
            $ch = $content[$i];

            if ($inComment) {
                if ($ch === "\r" || $ch === "\n") {
                    $inComment = false;
                    $pendingSpace = strlen($builder) > 0;
                }

                continue;
            }

            if ($inLiteralString) {
                $builder .= $ch;
                if ($escaping) {
                    $escaping = false;
                } elseif ($ch === '\\') {
                    $escaping = true;
                } elseif ($ch === '(') {
                    $literalDepth++;
                } elseif ($ch === ')') {
                    $literalDepth--;
                    if ($literalDepth <= 0) {
                        $inLiteralString = false;
                        $literalDepth = 0;
                    }
                }

                continue;
            }

            if ($inHexString) {
                $builder .= $ch;
                if ($ch === '>') {
                    $inHexString = false;
                }

                continue;
            }

            if (ctype_space($ch)) {
                $pendingSpace = strlen($builder) > 0;
                continue;
            }

            if ($ch === '%') {
                $inComment = true;
                continue;
            }

            if ($pendingSpace && self::ShouldEmitSpace($builder)) {
                $builder .= ' ';
            }

            $pendingSpace = false;
            $builder .= $ch;

            if ($ch === '(') {
                $inLiteralString = true;
                $literalDepth = 1;
                $escaping = false;
            } elseif ($ch === '<' && !self::IsDictionaryStart($content, $i)) {
                $inHexString = true;
            }
        }

        return $builder;
    }

    private static function ShouldEmitSpace($builder)
    {
        if (strlen($builder) === 0) {
            return false;
        }

        $previous = $builder[strlen($builder) - 1];
        if ($previous === '[' || $previous === '<' || $previous === '(' || $previous === ' ') {
            return false;
        }

        return true;
    }

    private static function IsDictionaryStart($content, $index)
    {
        return $index + 1 < strlen($content) && $content[$index + 1] === '<';
    }

    private static function RebuildPdf($table)
    {
        $ids = array_keys($table->Objects);
        sort($ids, SORT_NUMERIC);

        $output = $table->HeaderBytes;
        $offsets = array();
        for ($i = 0; $i < count($ids); $i++) {
            $offsets[$ids[$i]] = strlen($output);
            $output .= $table->Objects[$ids[$i]];
        }

        $xrefStart = strlen($output);
        $maxId = count($ids) > 0 ? $ids[count($ids) - 1] : 0;
        $xref = "xref\n0 " . ($maxId + 1) . "\n0000000000 65535 f \n";
        for ($objectId = 1; $objectId <= $maxId; $objectId++) {
            if (isset($offsets[$objectId])) {
                $xref .= sprintf('%010d', $offsets[$objectId]) . " 00000 n \n";
            } else {
                $xref .= "0000000000 00000 f \n";
            }
        }

        $output .= $xref;

        $trailer = 'trailer<< /Size ' . ($maxId + 1);
        if ($table->RootReference !== null && $table->RootReference !== '') {
            $trailer .= ' /Root ' . $table->RootReference;
        }

        if ($table->InfoReference !== null && $table->InfoReference !== '') {
            $trailer .= ' /Info ' . $table->InfoReference;
        }

        $trailer .= " >>\nstartxref\n" . $xrefStart . "\n%%EOF";
        $output .= $trailer;
        return $output;
    }
}
