<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\FontValue;
use Aspose\Cells_FOSS\Skia\SKBitmap;
use Aspose\Cells_FOSS\Skia\SKColor;
use Aspose\Cells_FOSS\Skia\SKColors;

final class ChartTextStyleResolver
{
    const A = '{http://schemas.openxmlformats.org/drawingml/2006/main}';

    const R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}';

    private function __construct()
    {
    }

    public static function TryResolveFont($defaultRunProperties, $majorFontName, $minorFontName, &$font)
    {
        $font = null;
        if ($defaultRunProperties === null) {
            return false;
        }

        $resolved = new FontValue();
        $found = false;

        if (self::TryResolveSize($defaultRunProperties, $sizePt)) {
            $resolved->Size = $sizePt;
            $found = true;
        }

        if (self::TryResolveBooleanAttr($defaultRunProperties, 'b', $bold)) {
            $resolved->Bold = $bold;
            $found = true;
        }

        if (self::TryResolveBooleanAttr($defaultRunProperties, 'i', $italic)) {
            $resolved->Italic = $italic;
            $found = true;
        }

        if (self::TryResolveTypeface($defaultRunProperties, $majorFontName, $minorFontName, $typeface)) {
            $resolved->Name = $typeface;
            $found = true;
        }

        if (!$found) {
            return false;
        }

        $font = $resolved;
        return true;
    }

    public static function TryResolveTextColor($defaultRunProperties, $colors, $chartModel, &$color)
    {
        $color = SKColors::Transparent();
        if ($defaultRunProperties === null) {
            return false;
        }

        $solidFill = $defaultRunProperties->Element(self::A . 'solidFill');
        $fillSource = ChartXmlParser::FillChild($solidFill);
        if (ChartXmlParser::TryDrawingColor($fillSource, $colors, $color)) {
            return true;
        }

        $gradientFill = $defaultRunProperties->Element(self::A . 'gradFill');
        if (self::TryResolveGradientTextColor($gradientFill, $colors, $color)) {
            return true;
        }

        $blipFill = $defaultRunProperties->Element(self::A . 'blipFill');
        if ($blipFill === null || $chartModel === null) {
            return false;
        }

        $blip = $blipFill->Element(self::A . 'blip');
        $relationshipId = $blip !== null ? self::AttrOf($blip, self::R . 'embed') : null;
        if ($relationshipId === null || $relationshipId === '') {
            $relationshipId = $blip !== null ? self::AttrOf($blip, 'embed') : null;
        }

        if ($relationshipId === null || $relationshipId === '') {
            return false;
        }

        $companion = null;
        for ($i = 0; $i < count($chartModel->CompanionFiles); $i++) {
            $candidate = $chartModel->CompanionFiles[$i];
            $binaryContent = isset($candidate->BinaryContent) ? $candidate->BinaryContent : null;
            if ($candidate !== null
                && $candidate->RelationshipId === $relationshipId
                && $binaryContent !== null
                && strlen((string)$binaryContent) > 0) {
                $companion = $candidate;
                break;
            }
        }

        if ($companion === null) {
            return false;
        }

        return self::TryResolveRepresentativeImageColor($companion->BinaryContent, $color);
    }

    public static function TryResolveTextOutlineColor($runProperties, $colors, &$color)
    {
        $color = SKColors::Transparent();
        $line = $runProperties !== null ? $runProperties->Element(self::A . 'ln') : null;
        if ($line === null) {
            return false;
        }

        if (self::TryResolveGradientTextColor($line->Element(self::A . 'gradFill'), $colors, $color)) {
            return true;
        }

        $solidFill = $line->Element(self::A . 'solidFill');
        return ChartXmlParser::TryDrawingColor(ChartXmlParser::FillChild($solidFill), $colors, $color);
    }

    private static function TryResolveGradientTextColor($gradientFill, $colors, &$color)
    {
        $color = SKColors::Transparent();
        if ($gradientFill === null || $colors === null) {
            return false;
        }

        $gradientStops = $gradientFill->Element(self::A . 'gsLst');
        if ($gradientStops === null) {
            return false;
        }

        $resolvedStops = array();
        foreach ($gradientStops->Elements(self::A . 'gs') as $stop) {
            $raw = self::AttrOf($stop, 'pos');
            $position = 0;
            if ($raw === null || $raw === '' || !is_numeric($raw) || (int)$raw != (float)$raw) {
                continue;
            }
            $position = (int)$raw;

            $source = $stop->Element(self::A . 'srgbClr');
            if ($source === null) {
                $source = $stop->Element(self::A . 'schemeClr');
            }

            if (ChartXmlParser::TryDrawingColor($source, $colors, $stopColor)) {
                $resolvedStops[] = array($position, $stopColor);
            }
        }

        if (count($resolvedStops) === 0) {
            return false;
        }

        if (count($resolvedStops) === 1) {
            $color = $resolvedStops[0][1];
            return true;
        }

        $totalWeight = 0.0;
        $sumR = 0.0;
        $sumG = 0.0;
        $sumB = 0.0;
        $sumA = 0.0;
        for ($i = 0; $i < count($resolvedStops); $i++) {
            $current = $resolvedStops[$i];
            $previousPos = $i === 0 ? 0 : $resolvedStops[$i - 1][0];
            $nextPos = $i === count($resolvedStops) - 1 ? 100000 : $resolvedStops[$i + 1][0];
            $weight = max(1.0, ($nextPos - $previousPos) * 0.5);
            $totalWeight += $weight;
            $sumR += $current[1]->Red * $weight;
            $sumG += $current[1]->Green * $weight;
            $sumB += $current[1]->Blue * $weight;
            $sumA += $current[1]->Alpha * $weight;
        }

        if ($totalWeight <= 0.0) {
            return false;
        }

        $color = new SKColor(
            (int)round($sumR / $totalWeight),
            (int)round($sumG / $totalWeight),
            (int)round($sumB / $totalWeight),
            (int)round($sumA / $totalWeight));
        return true;
    }

    private static function TryResolveRepresentativeImageColor($imageBytes, &$color)
    {
        $color = SKColors::Transparent();
        if ($imageBytes === null || strlen((string)$imageBytes) === 0) {
            return false;
        }

        $bitmap = SKBitmap::Decode($imageBytes);
        if ($bitmap === null) {
            return false;
        }

        try {
            if ($bitmap->Width <= 0 || $bitmap->Height <= 0) {
                return false;
            }

            $sumR = 0;
            $sumG = 0;
            $sumB = 0;
            $count = 0;

            for ($y = 0; $y < $bitmap->Height; $y++) {
                for ($x = 0; $x < $bitmap->Width; $x++) {
                    $pixel = $bitmap->GetPixel($x, $y);
                    if ($pixel->Alpha === 0) {
                        continue;
                    }

                    if ($pixel->Red > 247 && $pixel->Green > 247 && $pixel->Blue > 247) {
                        continue;
                    }

                    $sumR += $pixel->Red;
                    $sumG += $pixel->Green;
                    $sumB += $pixel->Blue;
                    $count++;
                }
            }

            if ($count === 0) {
                for ($y = 0; $y < $bitmap->Height; $y++) {
                    for ($x = 0; $x < $bitmap->Width; $x++) {
                        $pixel = $bitmap->GetPixel($x, $y);
                        if ($pixel->Alpha === 0) {
                            continue;
                        }

                        $sumR += $pixel->Red;
                        $sumG += $pixel->Green;
                        $sumB += $pixel->Blue;
                        $count++;
                    }
                }
            }

            if ($count === 0) {
                return false;
            }

            $color = new SKColor(
                self::ToByte($sumR / (float)$count),
                self::ToByte($sumG / (float)$count),
                self::ToByte($sumB / (float)$count));
            return true;
        } finally {
            $bitmap->Dispose();
        }
    }

    private static function ToByte($value)
    {
        $rounded = round($value);
        if ($rounded < 0.0) {
            return 0;
        }

        if ($rounded > 255.0) {
            return 255;
        }

        return (int)$rounded;
    }

    private static function TryResolveSize($defaultRunProperties, &$sizePt)
    {
        $sizePt = 0.0;
        $raw = self::AttrOf($defaultRunProperties, 'sz');
        if ($raw === null || $raw === '' || !is_numeric($raw)) {
            return false;
        }

        $sizeHundredthPoints = (float)$raw;
        if ($sizeHundredthPoints <= 0.0) {
            return false;
        }

        $sizePt = $sizeHundredthPoints / 100.0;
        return true;
    }

    private static function TryResolveBooleanAttr($element, $name, &$value)
    {
        $value = false;
        $raw = self::AttrOf($element, $name);
        if ($raw === null || $raw === '') {
            return false;
        }

        $value = $raw !== '0' && strcasecmp($raw, 'false') !== 0;
        return true;
    }

    private static function TryResolveTypeface($defaultRunProperties, $majorFontName, $minorFontName, &$typeface)
    {
        $typeface = null;
        $latin = $defaultRunProperties->Element(self::A . 'latin');
        $typeface = self::ResolveTypefaceAttr($latin, $majorFontName, $minorFontName);
        if ($typeface !== null && $typeface !== '') {
            return true;
        }

        $eastAsian = $defaultRunProperties->Element(self::A . 'ea');
        $typeface = self::ResolveTypefaceAttr($eastAsian, $majorFontName, $minorFontName);
        if ($typeface !== null && $typeface !== '') {
            return true;
        }

        $complexScript = $defaultRunProperties->Element(self::A . 'cs');
        $typeface = self::ResolveTypefaceAttr($complexScript, $majorFontName, $minorFontName);
        return $typeface !== null && $typeface !== '';
    }

    private static function ResolveTypefaceAttr($element, $majorFontName, $minorFontName)
    {
        $typeface = self::AttrOf($element, 'typeface');
        if ($typeface === null || $typeface === '') {
            return null;
        }

        if (str_starts_with($typeface, '+mn-')) {
            return $minorFontName;
        }

        if (str_starts_with($typeface, '+mj-')) {
            return $majorFontName;
        }

        if ($typeface[0] === '+') {
            return null;
        }

        return $typeface;
    }

    private static function AttrOf($element, $name)
    {
        if ($element === null) {
            return null;
        }

        $attribute = $element->Attribute($name);
        if ($attribute === null) {
            return null;
        }

        return $attribute->Value;
    }
}
