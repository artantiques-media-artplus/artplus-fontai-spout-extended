<?php

namespace Fontai\Spout\Writer;

use Box\Spout\Writer\WriterFactory as OriginalWriterFactory;
use Box\Spout\Common\Helper\GlobalFunctionsHelper;
use Box\Spout\Common\Type;

class WriterFactory extends OriginalWriterFactory
{
    public static function create($writerType)
    {
        if ($writerType == Type::XLSX)
        {
            $writer = new XLSX\Writer();
            $writer->setGlobalFunctionsHelper(new GlobalFunctionsHelper());
            return $writer;
        }

        if ($writerType == Type::CSV)
        {
            $writer = new CSV\Writer();
            $writer->setGlobalFunctionsHelper(new GlobalFunctionsHelper());
            return $writer;
        }

        return parent::create($writerType);
    }
}
