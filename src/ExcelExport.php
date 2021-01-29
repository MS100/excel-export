<?php

namespace Ms100\ExcelExport;

use Vtiful\Kernel\Excel;

class ExcelExport
{
    /**
     * @var \Vtiful\Kernel\Excel
     */
    private $excel;
    /**
     * @var array
     */
    private $config = [];
    /**
     * @var string
     */
    private $file_name;
    /**
     * @var array
     */
    private $sheets = [];
    /**
     * @var int
     */
    //private $default_font_size = 12;

    /**
     * ExcelExport constructor.
     *
     * @param string      $file_name  文件名
     * @param string|null $export_dir 导出路径
     */
    final public function __construct(
        string $file_name,
        string $export_dir = null
    ) {
        if (empty($export_dir)) {
            $export_dir = sys_get_temp_dir().DIRECTORY_SEPARATOR.'excel-export';
        }
        $export_dir .= DIRECTORY_SEPARATOR.getmypid();

        if (!file_exists($export_dir)) {
            mkdir($export_dir, 0777, true);
        }
        $this->config = [
            'path' => realpath($export_dir),
        ];
        $this->file_name = $file_name;
        //$this->default_font_size = min(100, max(1, $default_font_size));
    }

    final public function addSheet(Sheet $sheet, string $sheet_name = '')
    {
        static $sheet_after_bind_method;

        if ($sheet_name === '') {
            $sheet_name = $this->autoSheetName();
        }
        $this->sheets[$sheet_name] = $sheet;

        if (is_null($this->excel)) {
            $this->excel = new Excel($this->config);
            $this->excel->fileName($this->file_name, $sheet_name);
            //$this->setDefaultFontSize();
        } else {
            $this->excel->addSheet($sheet_name);
        }

        if (is_null($sheet_after_bind_method)) {
            $sheet_reflection = new \ReflectionClass(Sheet::class);
            $sheet_after_bind_method = $sheet_reflection->getMethod(
                'afterBindToExcel'
            );
            $sheet_after_bind_method->setAccessible(true);
        }
        $sheet_after_bind_method->invoke($sheet, $this->excel, $sheet_name);

        return $sheet_name;
    }

    final public function export()
    {
        return $this->excel->output();
    }

    final public function download()
    {
        header(
            'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        header(
            'Content-Disposition: attachment;filename="'.$this->file_name.'"'
        );
        header('Cache-Control: max-age=0');

        copy($this->export(), 'php://output');

        $this->unlink();
    }

    final public function getSheet($sheet_name)
    {
        return $this->sheets[$sheet_name] ?? null;
    }

    private function autoSheetName()
    {
        $number = count($this->sheets) + 1;
        $sheet_name = 'Sheet'.$number;
        while (isset($this->sheets[$sheet_name])) {
            $number--;
            $sheet_name = 'Sheet'.$number;
        }

        return $sheet_name;
    }

    /*private function setDefaultFontSize()
    {
        $format = new Format($this->excel->getHandle());
        $style = $format->fontSize($this->default_font_size)->toResource();

        $this->excel->defaultFormat($style);
    }*/

    public function unlink()
    {
        @unlink($this->config['path'].DIRECTORY_SEPARATOR.$this->file_name);
    }
}