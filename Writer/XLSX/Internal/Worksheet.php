<?php

namespace Fontai\Spout\Writer\XLSX\Internal;

use Box\Spout\Writer\XLSX\Internal\Worksheet as OriginalWorksheet;
use Box\Spout\Common\Exception\IOException;
use Box\Spout\Writer\Common\Helper\CellHelper;
use DateTimeInterface;

class Worksheet extends OriginalWorksheet
{
    protected function addNonEmptyRow($dataRow, $style)
    {
        $cellNumber = 0;
        $rowIndex = $this->lastWrittenRowIndex + 1;
        $numCells = count($dataRow);

        $rowXML = '<row r="' . $rowIndex . '" spans="1:' . $numCells . '">';

        if (!is_array($style))
        {
            $style = [$style];
        }

        foreach ($dataRow as $i => $cellValue)
        {
            if ($cellValue instanceof DateTimeInterface)
            {
                $cellValue = 25569 + (($cellValue->format('U') + $cellValue->format('Z')) / 86400);
            }
            
            $styleId = isset($style[$i]) ? $style[$i]->getId() : $style[0]->getId();

            $rowXML .= $this->getCellXML($rowIndex, $cellNumber, $cellValue, $styleId);
            $cellNumber++;
        }

        $rowXML .= '</row>';

        $wasWriteSuccessful = fwrite($this->sheetFilePointer, $rowXML);

        if ($wasWriteSuccessful === false)
        {
            throw new IOException("Unable to write data in {$this->worksheetFilePath}");
        }
    }
}
