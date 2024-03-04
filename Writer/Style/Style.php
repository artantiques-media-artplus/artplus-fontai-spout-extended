<?php

namespace Fontai\Spout\Writer\Style;

use Box\Spout\Writer\Style\Style as OriginalStyle;

class Style extends OriginalStyle
{
    /** @var string Number format */
    protected $numberFormat = NULL;
    /** @var bool */
    protected $hasSetNumberFormat = false;


    /**
     * Sets the number format
     * @param string $numberFormat format code (@see NumberFormat)
     * @return Style
     */
    public function setNumberFormat($numberFormat)
    {
        $this->hasSetNumberFormat = TRUE;
        $this->numberFormat = $numberFormat;
        return $this;
    }

    /**
     * @return string
     */
    public function getNumberFormat()
    {
        return $this->numberFormat;
    }

    public function getPhpNumberFormat()
    {
        return strtr($this->numberFormat, [
            'mm' => 'i',
            'h' => 'H',
            'yyyy' => 'Y',
            'd' => 'j',
            'm' => 'n'
        ]);
    }

    /**
     * @return bool Whether the number format should be applied
     */
    public function shouldApplyNumberFormat()
    {
        return $this->hasSetNumberFormat;
    }

    /**
     * @param Style $styleToUpdate Style to update (passed as reference)
     * @param Style $baseStyle
     * @return void
     */
    private function mergeCellProperties($styleToUpdate, $baseStyle)
    {
        parent::mergeCellProperties($styleToUpdate, $baseStyle);
        
        if (!$this->hasSetNumberFormat && $baseStyle->shouldApplyNumberFormat()) $styleToUpdate->setNumberFormat($baseStyle->getNumberFormat());
    }
}
