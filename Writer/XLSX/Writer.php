<?php

namespace Fontai\Spout\Writer\XLSX;

use Box\Spout\Writer\XLSX\Writer as OriginalWriter;
use Box\Spout\Common\Exception\InvalidArgumentException;
use Fontai\Spout\Writer\Style\Style;
use Fontai\Spout\Writer\Style\StyleBuilder;
use Fontai\Spout\Writer\XLSX\Internal\Workbook;

class Writer extends OriginalWriter
{
    /** @var Style Bold row style */
    protected $boldRowStyle;

    public function __construct()
    {
        $this->boldRowStyle = $this->getBoldRowStyle();

        return parent::__construct();
    }

    protected function openWriter()
    {
        if (!$this->book)
        {
            $tempFolder = $this->tempFolder ? : sys_get_temp_dir();
            $this->book = new Workbook(
                $tempFolder,
                $this->shouldUseInlineStrings,
                $this->shouldCreateNewSheetsAutomatically,
                $this->defaultRowStyle
            );
            $this->book->addNewSheetAndMakeItCurrent();
        }
    }

    protected function getDefaultRowStyle()
    {
        return (new StyleBuilder())
            ->setFontSize(self::DEFAULT_FONT_SIZE)
            ->setFontName(self::DEFAULT_FONT_NAME)
            ->build();
    }

    protected function getBoldRowStyle()
    {
        return (new StyleBuilder())
            ->setFontBold()
            ->build();
    }

    public function addRowWithStyle(array $dataRow, $style)
    {
        $this->setRowStyle($this->convertStyle($style, $dataRow));
        $this->addRow($dataRow);
        $this->resetRowStyleToDefault();

        return $this;
    }

    public function addRowsWithStyle(array $dataRows, $style)
    {
        $this->setRowStyle($this->convertStyle($style, $dataRows[0]));
        $this->addRows($dataRows);
        $this->resetRowStyleToDefault();

        return $this;
    }

    public function addRowWithFontBold(array $dataRow)
    {
        return $this->addRowWithStyle($dataRow, $this->boldRowStyle);
    }

    public function addRowsWithFontBold(array $dataRows)
    {
        return $this->addRowsWithStyle($dataRows, $this->boldRowStyle);
    }

    private function setRowStyle(array $styles)
    {
        $this->rowStyle = [];
        
        foreach ($styles as $style)
        {
            $this->rowStyle[] = $style ? $style->mergeWith($this->defaultRowStyle) : $this->defaultRowStyle;
        }
    }

    private function resetRowStyleToDefault()
    {
        $this->rowStyle = $this->defaultRowStyle;
    }

    protected function convertStyle($style, array $dataRow)
    {
        if (!is_array($style))
        {
            return [$this->parseStyle($style)];
        }
        
        for ($i = 0; $i < count($dataRow); $i++)
        {
            if (!isset($style[$i]))
            {
                $style[$i] = NULL;
            }
            else
            {
                $style[$i] = $this->parseStyle($style[$i]);
            }
        }

        return array_slice($style, 0, $i);
    }

    protected function parseStyle($styleOrNumberFormat)
    {
        if ($styleOrNumberFormat instanceof Style)
        {
            return $styleOrNumberFormat;
        }
        elseif (is_string($styleOrNumberFormat))
        {
            return (new StyleBuilder())
            ->setNumberFormat($styleOrNumberFormat)
            ->build();
        }
    }
}
