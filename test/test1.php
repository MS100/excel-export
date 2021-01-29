<?php

include(dirname(__DIR__).DIRECTORY_SEPARATOR.'vendor/autoload.php');

$config = [
    'path' => __DIR__,
];
$excel = new \Vtiful\Kernel\Excel($config);
$excel->fileName("free.xlsx");
$format = new \Vtiful\Kernel\Format($excel->getHandle());
$style = $format->number('0')->toResource();

$textFile = $excel->header(['name', 'money'])->setColumn('B:B', 100, $style);

for ($index = 0; $index < 10; $index++) {
    $textFile->insertText($index + 1, 0, 'viest');
    $textFile->insertText($index + 1, 1, '10000000000000000001'); // #,##0 为单元格数据样式
}

$textFile->output();

exit;

