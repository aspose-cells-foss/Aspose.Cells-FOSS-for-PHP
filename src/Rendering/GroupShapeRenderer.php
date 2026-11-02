<?php

namespace Aspose\Cells_FOSS\Rendering;

use Aspose\Cells_FOSS\Core\PictureModel;
use Aspose\Cells_FOSS\Internal\Xml\XElement;
use Aspose\Cells_FOSS\Skia\SKFilterQuality;
use Aspose\Cells_FOSS\Skia\SKPaint;
use Aspose\Cells_FOSS\Skia\SKRect;

// Positions a DrawingML group's child shapes in its anchor rectangle.
final class GroupShapeRenderer
{
    const Xdr = '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}';
    const A = '{http://schemas.openxmlformats.org/drawingml/2006/main}';
    const R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}';

    private $context;

    public function __construct($context)
    {
        $this->context = $context;
    }

    public function TryDraw($canvas, $page, $frame, $xml)
    {
        if ($xml === null || $xml === '' || strpos($xml, 'grpSp') === false) {
            return false;
        }

        try {
            $root = XElement::Parse($xml);
        } catch (\Throwable $exception) {
            return false;
        }

        if ($root->Name !== self::Xdr . 'grpSp') {
            return false;
        }

        $this->DrawGroup($canvas, $page, $frame, $root);
        return true;
    }

    private function DrawGroup($canvas, $page, $frame, $root)
    {
        $properties = $root->Element(self::Xdr . 'grpSpPr');
        $transform = $properties !== null ? $properties->Element(self::A . 'xfrm') : null;
        $childOffset = $transform !== null ? $transform->Element(self::A . 'chOff') : null;
        $childExtent = $transform !== null ? $transform->Element(self::A . 'chExt') : null;
        if ($childOffset === null || $childExtent === null) {
            return;
        }

        $originX = self::Number($childOffset, 'x');
        $originY = self::Number($childOffset, 'y');
        $extentX = self::Number($childExtent, 'cx');
        $extentY = self::Number($childExtent, 'cy');
        if ($extentX <= 0.0 || $extentY <= 0.0) {
            return;
        }

        $renderer = new ShapeRenderer($this->context);
        foreach ($root->Elements() as $child) {
            if ($child->Name !== self::Xdr . 'sp' && $child->Name !== self::Xdr . 'pic' && $child->Name !== self::Xdr . 'grpSp') {
                continue;
            }

            $shapeProperties = $child->Element($child->Name === self::Xdr . 'grpSp' ? self::Xdr . 'grpSpPr' : self::Xdr . 'spPr');
            $childTransform = $shapeProperties !== null ? $shapeProperties->Element(self::A . 'xfrm') : null;
            $offset = $childTransform !== null ? $childTransform->Element(self::A . 'off') : null;
            $extent = $childTransform !== null ? $childTransform->Element(self::A . 'ext') : null;
            if ($offset === null || $extent === null) {
                continue;
            }

            $left = $frame->Left + (float)((self::Number($offset, 'x') - $originX) / $extentX * $frame->Width);
            $top = $frame->Top + (float)((self::Number($offset, 'y') - $originY) / $extentY * $frame->Height);
            $width = (float)(self::Number($extent, 'cx') / $extentX * $frame->Width);
            $height = (float)(self::Number($extent, 'cy') / $extentY * $frame->Height);
            if ($width <= 0.0 || $height <= 0.0) {
                continue;
            }

            $childFrame = new SKRect($left, $top, $left + $width, $top + $height);
            if ($child->Name === self::Xdr . 'grpSp') {
                $this->DrawGroup($canvas, $page, $childFrame, $child);
            } elseif ($child->Name === self::Xdr . 'pic') {
                $this->DrawPicture($canvas, $page, $childFrame, $child);
            } else {
                $parsed = ShapeXmlParser::Parse($child->toString(), 'rect', $this->context->Colors);
                if ($parsed !== null) {
                    if ($parsed->TextWrap && !($parsed->Text === null || $parsed->Text === '')) {
                        // Excel's grouped text boxes use a wider effective font than the
                        // fallback Calibri metrics, which also changes their line breaks.
                        $parsed->TextSizePt *= 1.18;
                    }

                    $renderer->Draw($canvas, $page, $childFrame, $parsed);
                }
            }
        }
    }

    private function DrawPicture($canvas, $page, $frame, $picture)
    {
        $blipFill = $picture->Element(self::Xdr . 'blipFill');
        $blip = $blipFill !== null ? $blipFill->Element(self::A . 'blip') : null;
        $relationshipId = $blip !== null ? (string)$blip->Attribute(self::R . 'embed') : null;
        if ($relationshipId === null || $relationshipId === '' || $page === null) {
            return;
        }

        $source = null;
        foreach ($page->Sheet->Sheet->ShapeImages as $image) {
            if ($image->OriginalRId === $relationshipId) {
                $source = $image;
                break;
            }
        }

        if ($source === null) {
            return;
        }

        $imageModel = new PictureModel();
        $imageModel->ImageData = $source->ImageData;
        $imageModel->ImageExtension = $source->Extension;
        $decoded = $this->context->Pictures->GetImage($imageModel, $frame->Width, $frame->Height);
        if ($decoded === null) {
            return;
        }

        $imageFrame = $frame;
        $crop = $blipFill->Element(self::A . 'srcRect');
        if ($crop !== null) {
            $left = self::Number($crop, 'l') / 100000.0;
            $top = self::Number($crop, 't') / 100000.0;
            $right = self::Number($crop, 'r') / 100000.0;
            $bottom = self::Number($crop, 'b') / 100000.0;
            if ($left + $right < 1.0 && $top + $bottom < 1.0) {
                $imageFrame = new SKRect(
                    $frame->Left - (float)($frame->Width * $left / (1.0 - $left - $right)),
                    $frame->Top - (float)($frame->Height * $top / (1.0 - $top - $bottom)),
                    $frame->Right + (float)($frame->Width * $right / (1.0 - $left - $right)),
                    $frame->Bottom + (float)($frame->Height * $bottom / (1.0 - $top - $bottom)));
            }
        }

        $paint = new SKPaint();
        $paint->IsAntialias = true;
        $paint->FilterQuality = SKFilterQuality::None;
        $canvas->Save();
        $canvas->ClipRect($frame);
        $canvas->DrawImage($decoded, $imageFrame, $paint);
        $canvas->Restore();
        $paint->Dispose();
    }

    private static function Number($element, $attribute)
    {
        $text = (string)$element->Attribute($attribute);
        if ($text === '' || !is_numeric($text)) {
            return 0.0;
        }

        return (float)$text;
    }
}
