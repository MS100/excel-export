## 说明
基于php扩展 xlswriter，[手册](https://xlswriter-docs.viest.me/) 和 [Pecl](https://pecl.php.net/package/xlswriter)。

提供对excel表格快速导入。

## 安装
```bash
composer require ms100/excel-export
```

## 使用

#### 一个表格一个sheet
[见代码](test/test1.php)

 
#### 一个表格多个sheet

[见代码](test/test2.php)

#### 注意事项 


* 自定义的format方法可以做加粗、倾斜、改字号、设置数值格式(仅当字段值为整型或浮点型时生效)、改颜色，改边框等等**。
* 插入数据时，字段值为文本，则单元格会被设置为文本格式；插入数值再长，列宽再窄，也不会被显示为科学记数法。
