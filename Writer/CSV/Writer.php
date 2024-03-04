<?php

namespace Fontai\Spout\Writer\CSV;

use Box\Spout\Writer\CSV\Writer as OriginalWriter;
use Fontai\Spout\Writer\Style\Style;
use Fontai\Spout\Writer\Style\StyleBuilder;

class Writer extends OriginalWriter
{
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

    private function setRowStyle(array $styles)
    {
        $this->rowStyle = [];
        foreach($styles as $style)
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
        if (!is_array($style)) return [$this->parseStyle($style)];
        
        for ($i = 0; $i < count($dataRow); $i++)
        {
            if (!isset($style[$i])) $style[$i] = NULL;
            else $style[$i] = $this->parseStyle($style[$i]);
        }

        return array_slice($style, 0, $i);
    }

    protected function parseStyle($styleOrNumberFormat)
    {
        if ($styleOrNumberFormat instanceof Style) return $styleOrNumberFormat;
        elseif (is_string($styleOrNumberFormat))
        {
            return (new StyleBuilder())
            ->setNumberFormat($styleOrNumberFormat)
            ->build();
        }
    }

    protected function addRowToWriter(array $dataRow, $style)
    {
        foreach ($dataRow as $i => $value)
        {
            if ($value instanceof \DateTimeInterface)
            {
                $dataRow[$i] = $value->format(isset($style[$i]) ? $style[$i]->getPhpNumberFormat() : 'Y-m-d H:i:s');
            }
        }

        $wasWriteSuccessful = $this->globalFunctionsHelper->fputcsv($this->filePointer, $dataRow, $this->fieldDelimiter, $this->fieldEnclosure);
        if ($wasWriteSuccessful === false) {
            throw new IOException('Unable to write data');
        }

        $this->lastWrittenRowIndex++;
        if ($this->lastWrittenRowIndex % self::FLUSH_THRESHOLD === 0) {
            $this->globalFunctionsHelper->fflush($this->filePointer);
        }
    }
}
