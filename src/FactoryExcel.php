<?php
/**
 * Created by PhpStorm.
 * User: lucky.li
 * Date: 2018/8/17
 * Time: 21:48
 */

namespace Opensite\Pexcel;


use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Writer\Ods;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class FactoryExcel
{
    public static $excel;

    /**
     * @param $factory
     * @return bool|Csv|Ods|Xls|Xlsx
     * @throws \Exception
     */
    public static function factory($factory,Spreadsheet $spreadsheet)
    {
        if (empty($factory)){
            throw new \Exception("请选择导出的类型");
        }
        $factory = strtolower($factory);
        switch ($factory){
            case 'xlsx':
                self::$excel = new Xlsx($spreadsheet);  //导出类型为Xlsx
                break;
            case 'csv':
                self::$excel = new Csv($spreadsheet);   //导出类型为Csv
                break;
            case 'xls':
                self::$excel = new Xls($spreadsheet);   //导出类型为Xls
                break;
            case 'olds':
                self::$excel = new Ods($spreadsheet);   //导出类型为Ods
                break;
            default:
                return false;
        }
        return self::$excel;
    }
}