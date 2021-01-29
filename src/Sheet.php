<?php

namespace Ms100\ExcelExport;

use Vtiful\Kernel\Excel;
use Vtiful\Kernel\Format;

/**
 * 用 format大驼峰字段名的方法 单独定义列的单元格格式
 * 可定义样式参照 https://xlswriter-docs.viest.me/zh-cn/yang-shi-lie-biao
 */
abstract class Sheet
{
    /**
     * @var Excel
     */
    private $excel;
    /**
     * @var string
     */
    private $sheet_name = '';

    /**
     * @var array
     */
    private $column_keys = [];

    /**
     * @var array
     */
    private $data = [];

    /**
     * 定义布局
     */
    protected const LAYOUT = [
        //'字段key' => [ '字段表头', '字段为空或不存在时的默认值', '列宽度']
    ];

    protected const FORMAT_METHOD_PREFIX = 'format';

    protected const COLUMN_HEADER_KEY = 0;
    protected const COLUMN_DEFAULT_VALUE_KEY = 1;
    protected const COLUMN_WIDTH_KEY = 2;

    /**
     * Sheet constructor.
     *
     * @param array $column_keys 需要输出的列，传空输出LAYOUT定义的所有列
     * @param array $data        数据
     */
    final public function __construct(
        array $column_keys = [],
        array $data = []
    ) {
        $this->setColumnKeys($column_keys);
        $this->data = $data;
    }

    private function setColumnKeys(array $column_keys)
    {
        if (empty($column_keys)) {
            $this->column_keys = array_keys(static::LAYOUT);
            var_dump($this->column_keys);
        } else {
            $this->column_keys = array_values($column_keys);
        }
    }

    private function setHeader()
    {
        $header = [];
        foreach ($this->column_keys as $key) {
            $header[] = $this->getColumnHeader($key);
        }

        $this->excel->header($header);

        //追加一个占位字符，防止表格切换替换最后一行的bug
        $this->excel->data([[' ']]);
    }

    private function getColumnDefaultValue(string $column_key)
    {
        return static::LAYOUT[$column_key][self::COLUMN_DEFAULT_VALUE_KEY] ?? '';
    }

    private function getColumnHeader(string $column_key)
    {
        return static::LAYOUT[$column_key][self::COLUMN_HEADER_KEY] ?? '';
    }

    private function getColumnWidth(string $column_key)
    {
        return static::LAYOUT[$column_key][self::COLUMN_WIDTH_KEY] ?? 0;
    }

    private function getColumnNo(string $column_key)
    {
        $res = '';
        if (($index = array_search($column_key, $this->column_keys))
            !== false
        ) {
            $index++;
            $conv = base_convert($index, 10, 26);
            for ($i = 0; $i < strlen($conv); $i++) {
                $res .= base_convert(
                    base_convert($conv[$i], 26, 10) + 9,
                    10,
                    36
                );
            }

            $res = strtoupper($res);
        }

        return $res;
    }

    private function getColumnFormat(string $column_key)
    {
        $method_name = self::FORMAT_METHOD_PREFIX
            .$this->toBigCamelCase($column_key);

        if (method_exists($this, $method_name)) {
            $format = new Format($this->excel->getHandle());

            $this->$method_name($format);

            return $format->toResource();
        } else {
            return null;
        }
    }

    private function toBigCamelCase(string $str)
    {
        $str = ucwords(strtolower($str), '_');

        $str = strtr($str, ['_' => '']);

        return $str;
    }

    /**
     * 绑定到ExcelExport对象上
     *
     * @param ExcelExport $export_excel
     * @param string      $sheet_name
     *
     * @return string
     * @throws \ReflectionException
     */
    final public function bindToExcel(
        ExcelExport $export_excel,
        string $sheet_name = ''
    ) {
        return $export_excel->addSheet($this, $sheet_name);
    }

    private function afterBindToExcel(Excel $excel, string $sheet_name)
    {
        $this->excel = $excel;
        $this->sheet_name = $sheet_name;
        $this->setHeader();

        foreach ($this->column_keys as $key) {
            if ($width = $this->getColumnWidth($key)) {
                $column_no = $this->getColumnNo($key);
                if (is_resource($format = $this->getColumnFormat($key))) {
                    $this->excel->setColumn(
                        $column_no.':'.$column_no,
                        $width,
                        $format
                    );
                } else {
                    var_dump($format);

                    $this->excel->setColumn(
                        $column_no.':'.$column_no,
                        $width
                    );
                }
            }
        }

        $this->insertData($this->data);
        $this->data = [];
    }

    /**
     * 插入数据
     *
     * @param $data
     */
    final public function insertData($data)
    {
        if (empty($data)
            || !(is_array($data)
                || $data instanceof \ArrayAccess
                || $data instanceof \Traversable
            )
        ) {
            return;
        }
        $insert = [];

        foreach ($data as $item) {
            $row = [];
            foreach ($this->column_keys as $key) {
                if (isset($item[$key]) && $item[$key] !== '') {
                    $row[] = $item[$key];
                } else {
                    $row[] = $this->getColumnDefaultValue($key);
                }
            }
            $insert[] = $row;
        }

        //追加一个占位字符，防止表格切换替换最后一行的bug
        $insert[] = [' '];

        $this->excel->checkoutSheet($this->sheet_name);
        $this->excel->data($insert);
    }
}

