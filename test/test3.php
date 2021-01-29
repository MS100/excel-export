<?php

include(dirname(__DIR__).DIRECTORY_SEPARATOR.'vendor/autoload.php');

#定义一个sheet类继承 \Ms100\ExcelExport\Sheet
class UserSheet extends \Ms100\ExcelExport\Sheet
{
    /**
     * [
     *    字段key => [ '字段表头', '字段为空或不存在时的默认值', '单元格宽度']
     * ]
     */
    protected const LAYOUT
        = [
            'name'   => ['姓名', '无名氏', 40,],
            'age'    => ['年龄', '-', 20,],
            'height' => ['身高', '-', 30,],
            'card'   => ['身份证', '-', 3],
        ];

    /**
     * 用 format大驼峰字段名的方法 单独定义列的单元格格式
     * 可定义样式参照 https://xlswriter-docs.viest.me/zh-cn/yang-shi-lie-biao
     *
     * @param \Vtiful\Kernel\Format $format
     */
    protected function formatName(\Vtiful\Kernel\Format $format)
    {
        //加粗倾斜
        $format->bold()->italic();
    }

    protected function formatHeight(\Vtiful\Kernel\Format $format)
    {
        //设置数值格式
        $format->number('0.00');
    }

    protected function formatCard(\Vtiful\Kernel\Format $format)
    {
        //设置边框
        $format->border(\Vtiful\Kernel\Format::BORDER_THIN)->number('#');
    }
}

//数据
$data = [
    //字段值为文本，则单元格会被设置为文本格式，插入数值再长，列宽再窄，也不会被显示为科学技术法
    ['name' => 'Rent', 'age' => 20, 'height' => 170, 'card' => '1234569123456'],
    ['name' => 'Gas', 'height' => 160, 'card' => '123456'],
    ['name' => 'Food', 'age' => 22, 'card' => 817212313706822311987123],
    ['age' => 27, 'height' => 155, 'card' => '1655231370682194287123'],
];

//创建一个导出类
$excel_export = new \Ms100\ExcelExport\ExcelExport('test.xlsx', __DIR__);

//创建一个sheet，只导出指定字段
$sheet = new UserSheet(
    ['name', 'height', 'card'],
    $data
);

//将sheet绑定到excel
$sheet->bindToExcel($excel_export, '测试1');

//后续可以继续插入
$sheet->insertData($data);

//输出路径
echo $excel_export->export();

//下载
//$excel_export->download();

//删除
//$excel_export->unlink();


