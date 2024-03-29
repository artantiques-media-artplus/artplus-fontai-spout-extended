<?php

namespace Fontai\Spout\Writer\Style;

/**
 * Class NumberFormat
 * @package Fontai\Spout\Writer\Style
 */
class NumberFormat
{
    const FORMAT_GENERAL                 = 'General';
    const FORMAT_TEXT                    = '@';

    const FORMAT_NUMBER                  = '0';
    const FORMAT_NUMBER_00               = '0.00';
    const FORMAT_NUMBER_COMMA_SEPARATED1 = '#,##0.00';
    const FORMAT_NUMBER_COMMA_SEPARATED2 = '#,##0.00_-';

    const FORMAT_PERCENTAGE              = '0%';
    const FORMAT_PERCENTAGE_00           = '0.00%';

    const FORMAT_DATE_YYYYMMDD2          = 'yyyy-mm-dd';
    const FORMAT_DATE_YYYYMMDD           = 'yy-mm-dd';
    const FORMAT_DATE_DDMMYYYY           = 'dd/mm/yy';
    const FORMAT_DATE_DMYSLASH           = 'd/m/y';
    const FORMAT_DATE_DMYMINUS           = 'd-m-y';
    const FORMAT_DATE_DMMINUS            = 'd-m';
    const FORMAT_DATE_MYMINUS            = 'm-y';
    const FORMAT_DATE_XLSX14             = 'mm-dd-yy';
    const FORMAT_DATE_XLSX15             = 'd-mmm-yy';
    const FORMAT_DATE_XLSX16             = 'd-mmm';
    const FORMAT_DATE_XLSX17             = 'mmm-yy';
    const FORMAT_DATE_XLSX22             = 'm/d/yy h:mm';
    const FORMAT_DATE_DATETIME           = 'd/m/y h:mm';
    const FORMAT_DATE_TIME1              = 'h:mm AM/PM';
    const FORMAT_DATE_TIME2              = 'h:mm:ss AM/PM';
    const FORMAT_DATE_TIME3              = 'h:mm';
    const FORMAT_DATE_TIME4              = 'h:mm:ss';
    const FORMAT_DATE_TIME5              = 'mm:ss';
    const FORMAT_DATE_TIME6              = 'h:mm:ss';
    const FORMAT_DATE_TIME7              = 'i:s.S';
    const FORMAT_DATE_TIME8              = 'h:mm:ss;@';
    const FORMAT_DATE_YYYYMMDDSLASH      = 'yy/mm/dd;@';

    const FORMAT_CURRENCY_USD_SIMPLE     = '"$"#,##0.00_-';
    const FORMAT_CURRENCY_USD            = '$#,##0_-';
    const FORMAT_CURRENCY_EUR_SIMPLE     = '[$EUR ]#,##0.00_-';

    public static function getBuiltInId($code)
    {
        $codes = [
            0 => self::FORMAT_GENERAL,
            1 => '0',
            2 => '0.00',
            3 => '#,##0',
            4 => '#,##0.00',
            9 => '0%',
            10 => '0.00%',
            11 => '0.00E+00',
            12 => '# ?/?',
            13 => '# ??/??',
            14 => 'm/d/yyyy',                     // Despite ECMA 'mm-dd-yy'
            15 => 'd-mmm-yy',
            16 => 'd-mmm',
            17 => 'mmm-yy',
            18 => 'h:mm AM/PM',
            19 => 'h:mm:ss AM/PM',
            20 => 'h:mm',
            21 => 'h:mm:ss',
            22 => 'm/d/yyyy h:mm',                // Despite ECMA 'm/d/yy h:mm'
            37 => '#,##0_);(#,##0)',              //  Despite ECMA '#,##0 ;(#,##0)'
            38 => '#,##0_);[Red](#,##0)',         //  Despite ECMA '#,##0 ;[Red](#,##0)'
            39 => '#,##0.00_);(#,##0.00)',        //  Despite ECMA '#,##0.00;(#,##0.00)'
            40 => '#,##0.00_);[Red](#,##0.00)',   //  Despite ECMA '#,##0.00;[Red](#,##0.00)'
            44 => '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)',
            45 => 'mm:ss',
            46 => '[h]:mm:ss',
            47 => 'mm:ss.0',                      //  Despite ECMA 'mmss.0'
            48 => '##0.0E+0',
            49 => '@',
            27 => '[$-404]e/m/d',
            30 => 'm/d/yy',
            36 => '[$-404]e/m/d',
            50 => '[$-404]e/m/d',
            57 => '[$-404]e/m/d',
            59 => 't0',
            60 => 't0.00',
            61 => 't#,##0',
            62 => 't#,##0.00',
            67 => 't0%',
            68 => 't0.00%',
            69 => 't# ?/?',
            70 => 't# ??/??'
        ];

        return array_search($code, $codes);
    }
}