<?php

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;

include __DIR__ . '/../vendor/autoload.php';
print_r(test3());
//print_r(1);die;

//初始化格式
function test1()
{
    $export[] = ['test1', '学习委员1', '男', '大专', '职务', '2020-12-01', ''];
    $export[] = ['test2', '学习委员2', '女', '中专', '职务1', '2020-12-11', '辈出'];

    $header   = [
        ['name' => '姓   名', 'w' => 10],
        ['name' => '党内职务', 'w' => 10],
        ['name' => '性别', 'w' => 5],
        ['name' => '年龄', 'w' => 5],
        ['name' => '文化程度', 'w' => 10],
        ['name' => '行政职务', 'w' => 10],
        ['name' => '入党时间', 'w' => 10],
        ['name' => '备注', 'w' => 15],
    ];
    $title    = ['党支部委员(党小组长)登记表', '党支部委员(党小组长)登记表', '生成时间:' . date('Y-m-d H:i:s')];
    $filename = '党支部委员(党小组长)登记表_' . date('YmdHis'
            . time());
    return export($header, $title, $export, $filename, 'xlsx', true);
}

//修改表头
function test2($status = 1)
{

    $export[] = [1, 'xx1', '男', 21, '大专', '4', '23', '23', '23', '23', '23', '23'];
    $export[] = [2, 'xx2', '女', 212, '大专1', '42', '232', '223', '223', '223', '223', '223'];
    if ($status == 1) {
        $title    = ['流出党员名册', '流出党员名册', '生成时间:' . date('Y-m-d H:i:s')];
        $filename = '流出党员名册_' . '_' . date('YmdHis'
                . time());
        $title1   = '何时外出';
        $title2   = '地点';
        $title3   = '外出务工情况';
    } else {
        $title    = ['流入党员名册', '流入党员名册', '生成时间:' . date('Y-m-d H:i:s')];
        $filename = '流入党员名册_   ' . date('YmdHis'
                . time());
        $title1   = '何时流入';
        $title2   = '组织关系所在地';
        $title3   = '流入务工情况';
    }
    $header = [
        ['name' => '序号', 'w' => 3],
        ['name' => '姓  名', 'w' => 10],
        ['name' => '性别', 'w' => 3],
        ['name' => '年龄', 'w' => 4],
        ['name' => '文化程度', 'w' => 9],
        ['name' => '家庭人口', 'w' => 9],
        ['name' => $title1, 'w' => 12],
        ['name' => $title2, 'w' => 12],
        ['name' => '从事职业及年收入', 'w' => 15],
        ['name' => '与党组织联系方式', 'w' => 12],
        ['name' => '思想状况', 'w' => 12],
        ['name' => '备注', 'w' => 12],
    ];
    return export($header, $title, $export, $filename, 'xlsx', false,
        function (
            $sheet,
            $data,
            $span,
            $topNumber//excel,表头数据,单元格
        ) use ($title3) {
            $num = $topNumber + 1;
            $sheet->mergeCells("A{$topNumber}:A{$num}");
            $sheet->mergeCells("B{$topNumber}:B{$num}");
            $sheet->mergeCells("C{$topNumber}:C{$num}");
            $sheet->mergeCells("D{$topNumber}:D{$num}");
            $sheet->mergeCells("E{$topNumber}:E{$num}");
            $sheet->mergeCells("F{$topNumber}:F{$num}");
            $sheet->mergeCells("J{$topNumber}:J{$num}");
            $sheet->mergeCells("K{$topNumber}:K{$num}");
            $sheet->mergeCells("L{$topNumber}:L{$num}");
            $sheet->mergeCells("G{$topNumber}:I{$topNumber}");
            $sheet->setCellValue("A{$topNumber}", '序号');
            $sheet->setCellValue("B{$topNumber}", '姓名');
            $sheet->setCellValue("C{$topNumber}", '性别');
            $sheet->setCellValue("D{$topNumber}", '年龄');
            $sheet->setCellValue("E{$topNumber}", '文化程度');
            $sheet->setCellValue("F{$topNumber}", '家庭人口');
            $sheet->setCellValue("G{$topNumber}", $title3);
            $sheet->setCellValue("J{$topNumber}", '与党组织联系方式');
            $sheet->setCellValue("K{$topNumber}", '思想状况');
            $sheet->setCellValue("L{$topNumber}", '备注');
            $sheet->getStyle("G{$topNumber}")->getFont()->setName('楷体');
            $sheet->getStyle("G{$num}:I{$num}")->getFont()->setName('楷体');
            $sheet->getStyle("G{$num}:I{$num}")->getFont()->setBold(true);
            foreach ($data as $key => $value) {
                $sheet->getColumnDimension($span)
                    ->setWidth(isset($value['w']) ? $value['w'] : 20);
                $sheet->setCellValue($span . $num,
                    isset($value['name']) ? $value['name'] : $value);
                $sheet->getStyle($span . $topNumber)->getFont()
                    ->setName('楷体');
                $span++;
            }
            $cell = chr(ord($span) - 1);
            $sheet->getStyle("A{$topNumber}:{$cell}{$topNumber}")
                ->getAlignment()->setWrapText(true);
            $sheet->getStyle("A{$topNumber}:{$cell}{$topNumber}")
                ->getBorders()->getAllBorders()
                ->setBorderStyle(Border::BORDER_THIN);
            $sheet->getStyle("A{$topNumber}:{$cell}{$topNumber}")
                ->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $sheet->getStyle("A{$topNumber}:{$cell}{$topNumber}")
                ->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
            $sheet->getStyle("A{$num}:{$cell}{$num}")->getAlignment()
                ->setWrapText(true);
            $sheet->getStyle("A{$num}:{$cell}{$num}")->getBorders()
                ->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
            $sheet->getStyle("A{$num}:{$cell}{$num}")->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $sheet->getStyle("A{$num}:{$cell}{$num}")->getAlignment()
                ->setVertical(Alignment::VERTICAL_CENTER);
            $sheet->getRowDimension(3)->setRowHeight(50);
            return $span;
        }, 4);
}

//修改底部
function test3()
{
    $symbol   = [
        0 => ['name' => '', 'title' => ''],   //说了要来 但是并未签到
        1 => ['name' => '√', 'title' => '到会'],//到会
        2 => ['name' => '#', 'title' => '迟到'],//迟到
        3 => ['name' => 'O', 'title' => '未到'],//病假
        4 => ['name' => '★', 'title' => '未到'],//公假
        5 => ['name' => '△', 'title' => '未到'],//事假
        6 => ['name' => '×', 'title' => '未到'] //无故缺席
    ];
    $export[] = [1, 'xx', '2020', '到会', '√'];
    $export[] = [2, 'xx2', '2020', '到会', '★'];
    $export[] = [3, 'xx3', '2020', '到会', '×'];


    $header   = [
        ['name' => '序  号', 'w' => 9],
        ['name' => '姓   名', 'w' => 20],
        ['name' => '签到时间', 'w' => 20],
        ['name' => '签到状态', 'w' => 20],
        ['name' => '备   注', 'w' => 20]
    ];
    $title    = [
        '签到情况表',
        '签到情况表',
        '生成时间:' . date('Y-m-d H:i:s')
    ];
    $filename = '_签到情况表_' . date('YmdHis' . time());
    return export($header, $title,
        function ($sheet, $topNumber, $styleArray, $height) use ($export) {
            $column = $topNumber + 1;
            //行写入
            foreach ($export as $key => $rows) {
                $span = 'A';
                //列写入
                foreach ($rows as $keyName => $value) {
                    $sheet->setCellValue($span . $column, $value);
                    $span++;
                }
                $column++;
            }
            $span = chr(ord($span) - 1);
            //结尾合并
            $sheet->mergeCells('A' . $column . ':' . $span . $column);
            $sheet->setCellValue('A' . $column,
                '考勤符号:√到会;#迟到;O病假;★公假;△事假、探亲、丧假、婚假、休假、产假;×无故缺席');
            $column--;
            $sheet->getDefaultRowDimension()->setRowHeight($height);
            //设置内容字体样式
            $sheet->getStyle('A1:' . $span . $column)
                ->applyFromArray($styleArray);
            //设置边框
            $sheet->getStyle('A3:' . $span . $column)->getBorders()
                ->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
            //设置自动换行
            $sheet->getStyle('A4:' . $span . $column)->getAlignment()
                ->setWrapText(true);
            //设置字体
            $sheet->getStyle('A4:' . $span . $column)->getFont()
                ->setName('仿宋');
        }, $filename, 'xlsx', false);
}

function export(
    $header,
    $title_arr,
    $export = [],
    $fileName = '',
    $suffix = 'xlsx',
    $is_save = false,
    $headerFun = null,
    $topNum = 3
)
{
    $title = isset($title_arr[0]) && !empty($title_arr[0]) ? $title_arr[0]
        : '导出数据';
    $name  = isset($title_arr[1]) && !empty($title_arr[1]) ? $title_arr[1]
        : '导出数据';
    $info  = isset($title_arr[2]) && !empty($title_arr[2]) ? $title_arr[2]
        : date('Y-m-d H:i:s', time());
    $path  = \bnc\excel\SpreadSheetExcelService::instance()
        ->setExcelHeader($header, $headerFun)
        ->setExcelTitle($title, $name, $info)
        ->setExcelTopNum($topNum)
        ->setExcelContent($export)
        ->excelSave($fileName, $suffix, $is_save);
    $path  = siteUrl() . $path;
    return $path;
}

function siteUrl()
{
    $protocol   = (!empty($_SERVER['HTTPS']) && $_SERVER['HTTPS'] !== 'off'
        || $_SERVER['SERVER_PORT'] == 443) ? "https://" : "http://";
    $domainName = $_SERVER['HTTP_HOST'];
    return $protocol . $domainName;
}