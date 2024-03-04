<?php

namespace Fontai\Spout\Writer\XLSX\Internal;

use Box\Spout\Writer\XLSX\Internal\Workbook as OriginalWorkbook;
use Box\Spout\Writer\Common\Sheet;
use Fontai\Spout\Writer\XLSX\Helper\StyleHelper;

class Workbook extends OriginalWorkbook
{
    public function __construct($tempFolder, $shouldUseInlineStrings, $shouldCreateNewSheetsAutomatically, $defaultRowStyle)
    {
        parent::__construct(
            $tempFolder,
            $shouldUseInlineStrings,
            $shouldCreateNewSheetsAutomatically,
            $defaultRowStyle
        );

        $this->styleHelper = new StyleHelper($defaultRowStyle);
    }

    public function addNewSheet()
    {
        $newSheetIndex = count($this->worksheets);
        $sheet = new Sheet($newSheetIndex, $this->internalId);

        $worksheetFilesFolder = $this->fileSystemHelper->getXlWorksheetsFolder();
        $worksheet = new Worksheet($sheet, $worksheetFilesFolder, $this->sharedStringsHelper, $this->styleHelper, $this->shouldUseInlineStrings);
        $this->worksheets[] = $worksheet;

        return $worksheet;
    }
}
