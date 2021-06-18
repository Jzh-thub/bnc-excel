<?php
include __DIR__ . '/../vendor/autoload.php';
print_r(test1());
//print_r(1);die;

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
    $title    = ['党支部委员(党小组长)登记表', '党支部委员(党小组长)登记表', '生成时间:'.date('Y-m-d H:i:s')];
    $filename = '党支部委员(党小组长)登记表_' . date('YmdHis'
            . time());
    return export($header, $title, $export, $filename,'xlsx',true);
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