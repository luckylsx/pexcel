<?php
/**
 * Created by PhpStorm.
 * User: lucky.li
 * Date: 2018/8/17
 * Time: 21:39
 */

namespace Opensite;

use Opensite\FactoryExcel;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Class Excel
 * @package Excel
 */
class Excel
{
    /**
     * @var Spreadsheet object Spreadsheet
     */
    protected $spreadsheet;
    /**
     * @var bool|\PhpOffice\PhpSpreadsheet\Writer\Csv|\PhpOffice\PhpSpreadsheet\Writer\Ods|\PhpOffice\PhpSpreadsheet\Writer\Xls|\PhpOffice\PhpSpreadsheet\Writer\Xlsx
     */
    protected $excel;
    /**
     * @var string 导出类型
     */
    protected $Type;
    /**
     * @var string 文件名
     */
    protected $fileName;
    /**
     * @var array 允许导出的文件类型
     */
    protected $alows = ['csv', 'xls', 'xlsx', 'ods'];
    /**
     * @var array 支持26列及以内导出
     */
    protected $forms;
    /**
     * @var int 导出多个sheet时 支持最大的sheet数
     */
    protected $limitSheet = 5;
    /**
     * @var array 重置配置
     */
    protected $config = [];

    /**
     * Excel constructor.
     * @param $Type string 导出类型
     * @param $fileName string 文件名
     * @throws \Exception
     */
    public function __construct($Type, $fileName, $config = [])
    {
        if (!in_array(strtolower($Type), $this->alows)) {
            throw new \Exception('导出的类型不允许');
        }
        $this->excel = $Type;
        $this->spreadsheet = new Spreadsheet();
        $this->fileName = $fileName . "." . strtolower($Type);
        $this->excel = FactoryExcel::factory($Type, $this->spreadsheet);
        $this->forms = range('A', 'Z');

    }

    /**
     * 设置header头 直接在浏览器上下载
     */
    protected function setHeader()
    {
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $this->fileName . '"');
        header('Cache-Control: max-age=0');
    }

    /**
     * @param $data
     * @param string $sheetName
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * $merge array 需要合并的单元格 如：['A1:B1','B1:D1']
     * @throws \Exception
     */
    public function export($data, $sheetName = "sheet1")
    {
        if (empty($data)) {
            throw new \Exception('导出数据为空！');
        }
        $this->setHeader();
        $sheet = $this->spreadsheet->getActiveSheet();
        //合并单元格
        $config = $this->config;
        if (isset($config['merge']) && !empty($config['merge'])) {
            foreach ($config['merge'] as $value) {
                $sheet->mergeCells($value);
            }
        }
        //设置单元格标题
        $sheet->setTitle($sheetName);
        //填充单元格数据
        for ($i = 0; $i < count($data); $i++) {
            for ($j = 0; $j < count($data[$i]); $j++) {
                $num = $i + 1;
                $cell = $this->forms[$j] . "$num";
                $sheet->setCellValue($cell, $data[$i][$j]);
            }
        }
        $writer = $this->excel;
        $writer->save('php://output');
        //释放内存
        $this->spreadsheet->disconnectWorksheets();
        unset($this->spreadsheet);
    }

    /**
     * 导出多个sheet
     * @param $data array 多个sheet导出的数据 多维数组
     * [
     * [ //每个sheet对应导出数据
     * ['name','age','test'], //sheet里对应每行数据
     * ['lucky',19,'teaue']
     * ],
     * [
     * ['id','name','score'],
     * [1,'lucky',80],
     * [2,'nancy',90]
     * ]
     * ]
     * @param array $sheets 多个sheet时，每个sheet名称
     * @throws \Exception
     */
    public function exportSheets($data, $sheets = [])
    {

        if (empty($data)) {
            throw new \Exception('导出数据为空！');
        }
        $this->setHeader();
        if (is_array($sheets) && !empty($sheets)) {
            foreach ($sheets as $index => $sheetName) {
                $this->setCell($data, $sheetName);
            }
        } else {
            $this->setCell($data);
        }
        $writer = $this->excel;
        $writer->save('php://output');
        $this->spreadsheet->disconnectWorksheets();
        unset($this->spreadsheet);
    }

    /**
     * 重复设置sheet
     * @param $data
     * @param string $sheetName
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function setCell($data, $sheetName = '')
    {
        for ($index = 0; $index < count($data); $index++) {
            if ($index == 0) {
                $sheet = $this->spreadsheet->getActiveSheet();
            } else {
                $sheet = $this->spreadsheet->createSheet();
            }
            if (!empty($sheetName)) {
                $sheet->setTitle($sheetName);
            }
            for ($i = 0; $i < count($data[$index]); $i++) {
                for ($j = 0; $j < count($data[$index][$i]); $j++) {
                    $num = $i + 1;
                    $cell = $this->forms[$j] . "$num";
                    $sheet->setCellValue($cell, $data[$index][$i][$j]);
                }
            }
        }
    }
}