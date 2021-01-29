<?php

include(dirname(__DIR__).DIRECTORY_SEPARATOR.'vendor/autoload.php');


class UserSheet extends \Ms100\ExcelExport\Sheet
{
    protected const LAYOUT
        = [
            'name'   => ['姓名', '无名氏', 40,],
            'age'    => ['年龄', '-', 20,],
            'height' => ['身高', '-', 30,],
        ];

    protected function formatName(\Vtiful\Kernel\Format $format)
    {
        $format->bold()->italic();
    }
}

$data = [
    ['name' => 'Rent', 'age' => 20, 'height' => 170],
    ['name' => 'Gas', 'height' => 160],
    ['name' => 'Food', 'age' => 22,],
    ['age' => 27, 'height' => 155],
];
$excel_export = new \Ms100\ExcelExport\ExcelExport('test.xlsx', __DIR__);
$sheet = new UserSheet(
    ['name', 'height'],
    $data
);

$sheet->bindToExcel($excel_export, '测试1');

$sheet2 = new UserSheet(
    ['name', 'age'],
    $data
);
$sheet2->bindToExcel($excel_export);

$sheet->insertData($data);

$sheet2->insertData($data);
echo $excel_export->export();

//$excel_export->unlink();

