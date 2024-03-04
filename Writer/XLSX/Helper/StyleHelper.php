<?php

namespace Fontai\Spout\Writer\XLSX\Helper;

use Box\Spout\Writer\XLSX\Helper\StyleHelper as OriginalStyleHelper;

class StyleHelper extends OriginalStyleHelper
{
    /**
     * @var array
     */
    protected $registeredNumberFormats = [];
    
    /**
     * @var array [STYLE_ID] => [NUM_FMT_ID] maps a style to a number format declaration
     */
    protected $styleIdToNumFmtMappingTable = [];

    /**
     * @var int The number format index counter for custom number formats.
     */
    protected $numFmtIndex = 164;

    public function registerStyle($style)
    {
        $styles = is_array($style) ? $style : [$style];
        $registeredStyles = [];

        foreach ($styles as $style)
        {
            $serializedStyle = $style->serialize();

            if (!$this->hasStyleAlreadyBeenRegistered($style))
            {
                $nextStyleId = count($this->serializedStyleToStyleIdMappingTable);
                $style->setId($nextStyleId);

                $this->serializedStyleToStyleIdMappingTable[$serializedStyle] = $nextStyleId;
                $this->styleIdToStyleMappingTable[$nextStyleId] = $style;
            }

            $registeredStyle = $this->getStyleFromSerializedStyle($serializedStyle);

            $this->registerFill($registeredStyle);
            $this->registerBorder($registeredStyle);
            $this->registerNumberFormat($registeredStyle);

            $registeredStyles[] = $registeredStyle;
        }
        
        return $registeredStyles;
    }

    /**
     * Register a number format definition
     *
     * @param \Box\Spout\Writer\Style\Style $style
     */
    protected function registerNumberFormat($style)
    {
        $styleId = $style->getId();

        $numberFormat = $style->getNumberFormat();

        if ($numberFormat)
        {
            $isNumberFormatRegistered = isset($this->registeredNumberFormats[$numberFormat]);

            // We need to track the already registered number format definitions
            if ($isNumberFormatRegistered)
            {
                $registeredStyleId                           = $this->registeredNumberFormats[$numberFormat];
                $registeredNumberFormatId                    = $this->styleIdToNumFmtMappingTable[$registeredStyleId];
                $this->styleIdToNumFmtMappingTable[$styleId] = $registeredNumberFormatId;
            }
            else
            {
                $this->registeredNumberFormats[$numberFormat] = $styleId;
                $this->styleIdToNumFmtMappingTable[$styleId]  = $this->numFmtIndex++;
            }

        }
        else
        {
            // The numFmtId maps a style to a number format declaration
            // When there is no number format definition - we default to 0
            $this->styleIdToNumFmtMappingTable[$styleId] = 0;
        }
    }

    public function getStylesXMLFileContent()
    {
        $content = <<<EOD
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
EOD;

        $content .= $this->getNumFmtsSectionContent();
        $content .= $this->getFontsSectionContent();
        $content .= $this->getFillsSectionContent();
        $content .= $this->getBordersSectionContent();
        $content .= $this->getCellStyleXfsSectionContent();
        $content .= $this->getCellXfsSectionContent();
        $content .= $this->getCellStylesSectionContent();

        $content .= <<<EOD
</styleSheet>
EOD;

        return $content;
    }

    /**
     * Returns the content of the "<numFmts>" section.
     *
     * @return string
     */
    protected function getNumFmtsSectionContent()
    {
        $content = '<numFmts count="' . count($this->registeredNumberFormats) . '">';

        foreach ($this->registeredNumberFormats as $styleId)
        {
            $style      = $this->styleIdToStyleMappingTable[$styleId];
            $formatCode = htmlspecialchars($style->getNumberFormat(), ENT_XML1 | ENT_COMPAT, 'UTF-8');
            $numFmtId   = $this->styleIdToNumFmtMappingTable[$styleId];

            $content .= '<numFmt numFmtId="' . $numFmtId . '" formatCode="' . $formatCode . '" />';
        }

        $content .= '</numFmts>';

        return $content;
    }

    protected function getCellXfsSectionContent()
    {
        $registeredStyles = $this->getRegisteredStyles();

        $content = '<cellXfs count="' . count($registeredStyles) . '">';

        foreach ($registeredStyles as $style)
        {
            $styleId  = $style->getId();
            $numFmtId = $this->styleIdToNumFmtMappingTable[$styleId];
            $fillId   = $this->styleIdToFillMappingTable[$styleId];
            $borderId = $this->styleIdToBorderMappingTable[$styleId];

            $content .= '<xf numFmtId="' . $numFmtId . '" fontId="' . $styleId . '" fillId="' . $fillId . '" borderId="' . $borderId . '" xfId="0"';

            if ($style->shouldApplyNumberFormat())
            {
                $content .= ' applyNumberFormat="1"';
            }

            if ($style->shouldApplyFont())
            {
                $content .= ' applyFont="1"';
            }

            $content .= sprintf(' applyBorder="%d"', $style->shouldApplyBorder() ? 1 : 0);

            if ($style->shouldWrapText())
            {
                $content .= ' applyAlignment="1">';
                $content .= '<alignment wrapText="1"/>';
                $content .= '</xf>';
            }
            else
            {
                $content .= '/>';
            }
        }

        $content .= '</cellXfs>';

        return $content;
    }

    protected function applyWrapTextIfCellContainsNewLine($style, $dataRow)
    {
        // if the "wrap text" option is already set, no-op
        if (!is_array($style) && $style->hasSetWrapText())
        {
            return $style;
        }

        foreach ($dataRow as $i => $cell)
        {
            if (is_string($cell) && strpos($cell, "\n") !== false)
            {
                if (!is_array($style))
                {
                    $style->setShouldWrapText();
                }
                else if (!$style[$i]->shouldWrapText())
                {
                    $style[$i]->setShouldWrapText();
                }

                break;
            }
        }

        return $style;
    }
}
