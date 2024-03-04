<?php

namespace Fontai\Spout\Writer\Style;

use Box\Spout\Writer\Style\StyleBuilder as OriginalStyleBuilder;

class StyleBuilder extends OriginalStyleBuilder
{
    public function __construct()
    {
        $this->style = new Style();
    }
    
    /**
     *  Sets a number format
     *
     * @api
     * @param string $numberFormat format code (@see NumberFormat)
     * @return StyleBuilder
     */
    public function setNumberFormat($numberFormat)
    {
        $this->style->setNumberFormat($numberFormat);
        return $this;
    }
}
