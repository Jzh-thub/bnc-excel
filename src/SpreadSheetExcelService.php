<?php

namespace bnc\excel;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * 电子表格生成
 * Class SpreadSheetExcelService
 * @package bnc\excel
 */
class SpreadSheetExcelService
{
    private static $instance = null;

    /**
     * PHPSpreadsheet 实例化对象
     * @var Spreadsheet
     */
    private static $spreadsheet = null;

    /**
     * sheet 实例化对象
     * @var Worksheet
     */
    private static $sheet = null;

    /**
     * 表头计数
     * @var
     */
    protected static $count;

    /**
     * 表头占行数
     * @var int
     */
    protected static $topNumber = 3;

    /**
     * 表能占据表行的字母对应self::$cellkey
     * @var
     */
    protected static $cells;

    /**
     * 表头数据
     * @var
     */
    protected static $data;

    /**
     * 文件名
     * @var string
     */
    protected static $title = '订单导出';

    /**
     * 行宽
     * @var int
     */
    protected static $width = 20;

    /**
     * 行高
     * @var int
     */
    protected static $height = 50;

    /**
     * 保存文件目录
     * @var string
     */
    protected static $path = "./phpExcel/";

    private static $styleArray = [
        'borders'   => [
            'allBorders' => [
//                PHPExcel_Style_Border里面有很多属性，想要其他的自己去看
//                'style' => Border::BORDER_THICK,//边框是粗的
//                'style' => Border::BORDER_DOUBLE,//双重的
//                'style' => Border::BORDER_HAIR,//虚线
//                'style' => Border::BORDER_MEDIUM,//实粗线
//                'style' => Border::BORDER_MEDIUMDASHDOT,//虚粗线
//                'style' => Border::BORDER_MEDIUMDASHDOTDOT,//点虚粗线
                'style' => Border::BORDER_THIN,//细边框
//                'color' => array('argb' => 'FFFF0000'),
            ]
        ],
        'font'      => [
//            'bold' => true,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ]
    ];

    private function __construct()
    {
    }

    private function __clone()
    {
    }

    /**
     * 初始化
     * @return SpreadSheetExcelService|null
     */
    public static function instance(): ?SpreadSheetExcelService
    {
        if (self::$instance == null) {
            self::$instance    = new self();
            self::$spreadsheet = $spreadsheet = new Spreadsheet();
            self::$sheet       = $spreadsheet->getActiveSheet();
        }
        return self::$instance;
    }

    /**
     * 设置字体格式
     * @param string $title
     * @return false|string
     */
    public static function setUtf8(string $title)
    {
        return iconv('utf-8', 'gb2312', $title);
    }

    /**
     * 设置保存excel目录
     * @return false|string
     */
    public static function savePath()
    {
        if (!is_dir(self::$path)) {
            if (mkdir(self::$path, 0700) == false)
                return false;
        }

        $mont_path = self::$path . date('Ym');
        if (!is_dir($mont_path)) {
            if (mkdir($mont_path, 0700) == false)
                return false;
        }
        $day_path = $mont_path . DIRECTORY_SEPARATOR . date('d');
        if (!is_dir($day_path)) {
            if (mkdir($day_path, 0700) == false)
                return false;
        }
        return $day_path;
    }

    /**
     * 设置表头占行数
     * @param int $topNum
     * @return $this
     */
    public function setExcelTopNum(int $topNum = 3): SpreadSheetExcelService
    {
        self::$topNumber = $topNum;
        return $this;
    }

    /**
     * 设置标题
     * @param        $title || array ['title'=>'','name'=>'','info'=>[]]
     * @param string $name
     * @param        $info
     * @param null   $funCallBack function($style,$A1,$A2) 自定义设置头部样式
     * @return $this
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setExcelTitle($title, string $name = '', $info = null, $funCallBack = null): SpreadSheetExcelService
    {
        if (is_array($title)) {
            if (isset($title['title'])) $title = $info['title'];
            if (isset($title['name'])) $title = $info['name'];
            if (isset($title['info'])) $title = $info['info'];
        }
        if (empty($title))
            $title = self::$title;
        else
            self::$title = $title;
        if (empty($name)) $name = time();
        //设置Excel属性
        self::$spreadsheet->getProperties()
            ->setCreator("Neo")
            ->setLastModifiedBy("Neo")
            ->setTitle(self::setUtf8($title))
            ->setSubject($name)
            ->setDescription("")
            ->setKeywords($name)
            ->setCreator("");
        self::$sheet->setTitle($name);
        self::$sheet->setCellValue("A1", $title);
        self::$sheet->setCellValue("A2", self::setCellInfo($info));
        //文字居中
        self::$sheet->getStyle("A1")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        self::$sheet->getStyle("A2")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        //上下居中
        self::$sheet->getStyle("A1")->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        self::$sheet->getStyle("A2")->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        //合并表头单元格
        self::$sheet->mergeCells('A1:' . self::$cells . '1');
        self::$sheet->mergeCells('A2:' . self::$cells . '2');
        self::$sheet->getRowDimension(1)->setRowHeight(40);
        self::$sheet->getRowDimension(2)->setRowHeight(20);
        //设置表头字体
        if ($funCallBack !== null && is_callable($funCallBack)) {
            //自定义
            $funCallBack(self::$sheet, 'A1', 'A2');
        } else {
            self::$sheet->getStyle('A1')->getFont()->setName('方正小标宋');
            self::$sheet->getStyle('A1')->getFont()->setSize(20);
            self::$sheet->getStyle('A1')->getFont()->setBold(true);
            self::$sheet->getStyle('A2')->getFont()->setName('仿宋');
            self::$sheet->getStyle('A2')->getFont()->setSize(12);
        }
        self::$sheet->getStyle('A3:' . self::$cells . '3')->getFont()->setBold(true);
        return $this;
    }

    /**
     * 设置第二行标题内容
     * @param $info (['name'=>'','site'=>'','phone'=>123456] || ['我是操作者','我是地址','我是电话']) || string 自定义
     * @return string
     */
    private static function setCellInfo($info): string
    {
        $content = ['操作者：', '导出日期：' . date('Y-m-d', time()), '地址：', '电话：'];
        if (is_array($info) && !empty($info)) {
            if (isset($info['name'])) {
                $content[0] .= $info['name'];
            } else {
                $content[0] .= $info[0] ?? '';
            }
            if (isset($info['site'])) {
                $content[2] .= $info['site'];
            } else {
                $content[2] .= $info[1] ?? '';
            }
            if (isset($info['phone'])) {
                $content[3] .= $info['phone'];
            } else {
                $content[3] .= $info[2] ?? '';
            }
            return implode(' ', $content);
        } else if (is_string($info)) {
            return empty($info) ? implode(' ', $content) : $info;
        }
    }

    /**
     * 设置头部信息
     * @param      $data
     * @param null $funCallBack 回调
     * @return SpreadSheetExcelService
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public static function setExcelHeader($data, $funCallBack = null): SpreadSheetExcelService
    {
        $span = 'A';
        if ($funCallBack !== null && is_callable($funCallBack)) {
            $span = $funCallBack(self::$sheet, $data, $span, self::$topNumber);
        } else {
            foreach ($data as $key => $value) {
                self::$sheet->getColumnDimension($span)->setWidth($value['w'] ?? self::$width);
                self::$sheet->setCellValue($span . self::$topNumber, $value['name'] ?? $value);
                self::$sheet->getStyle($span . self::$topNumber)->getFont()->setName('楷体');
                self::$sheet->getStyle($span . self::$topNumber)->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
                self::$sheet->getStyle($span . self::$topNumber)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                self::$sheet->getStyle($span . self::$topNumber)->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
                //设置自动换行
                self::$sheet->getStyle($span . self::$topNumber)->getAlignment()->setWrapText(true);
                $span++;
            }
//        self::$sheet->getRowDimension(3)->setRowHeight(self::$height);
            self::$sheet->getRowDimension(3)->setRowHeight(20);

        }
        self::$cells = chr(ord($span) - 1);
        return new self();
    }

    /**
     * excel 数据导出
     * @param null $data 可以是数组 也可以是匿名函数
     * @return SpreadSheetExcelService
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setExcelContent($data = null): SpreadSheetExcelService
    {
        if ($data !== null && is_callable($data)) {
            $data(self::$sheet, self::$topNumber, self::$styleArray, self::$height);
        } else if ($data !== null && is_array($data) && count($data)) {
            $column = self::$topNumber + 1;
            //行写入
            foreach ($data as $key => $rows) {
                $span = 'A';
                //列写入
                foreach ($rows as $keyName => $value) {
                    self::$sheet->setCellValue($span . $column, $value);
                    $span++;
                }
                $column++;
            }
            $span = chr(ord($span) - 1);
            $column--;
            self::$sheet->getDefaultRowDimension()->setRowHeight(self::$height);
            //设置内容字体样式
            self::$sheet->getStyle('A1:' . $span . $column)->applyFromArray(self::$styleArray);
            //设置边框
            self::$sheet->getStyle('A3:' . $span . $column)->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
            //设置自动换行
            self::$sheet->getStyle('A' . (self::$topNumber + 1) . ':' . $span . $column)->getAlignment()->setWrapText(true);
            //设置字体
            self::$sheet->getStyle('A' . (self::$topNumber + 1) . ':' . $span . $column)->getFont()->setName('仿宋');
        }
        return new self;
    }

    /**
     * 保存表格数据，并下载
     * @param string $fileName 文件名
     * @param string $suffix 文件名后缀
     * @param bool   $is_save 是否保存文件
     * @return string
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function excelSave(string $fileName = '', string $suffix = 'xlsx', bool $is_save = false): string
    {
        if (empty($fileName)) {
            $fileName = date('YmdHis') . time();
        }
        if (empty($suffix)) {
            $suffix = 'xlsx';
        }
        //重命名表格(UTF8编码不需要这一步)
        if (mb_detect_encoding($fileName) != "UTF-8") {
            $fileName = iconv('utf-8', "gbk//IGNORE", $fileName);
        }
        $spreadsheet = self::$spreadsheet;
        $writer      = new Xlsx($spreadsheet);
        if (!$is_save) {
            if ($suffix == 'xlsx') {
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            } elseif ($suffix == 'xls') {
                header('Content-Type: application/vnd.ms-excel');
            }
            //直接下载
            header('Content-Disposition: attachment;filename="' . $fileName . '.' . $suffix . '"');
            header('Cache-Control: max-age=0');
            $writer->save('php://output');
            //删除情况 释放内存
            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
            exit;
        } else {
            $path = self::savePath() . '/' . $fileName . '.' . $suffix;
            $writer->save($path);
            //删除情况 释放内存
            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
            return $path;
        }

    }
}
