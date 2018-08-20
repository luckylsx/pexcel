## pexcel
Excel library for php

### Requirement

1. PHP >= 5.6
2. **[Composer](https://getcomposer.org/)**

## Installation

```shell
$ composer require opensite/pexcel
```

### Usage

```php
<?php

use pexcel;


/**
 * 导出类型
 * 导出文件名
 * 重设默认配置
 */  
$excel = new Excel("xls","导出数据",$config);

// 发送单个sheet
/**
 * data array 导出单个sheet 的数据
 * $sheet string sheet工作表名称 默认为：sheet1
 */
$excel->export($data);

// 导出多个sheet
/**
 * $data string 为多维数组
 * $param $data:
 * [
 *  	[ //每个sheet对应的导出数据
 *			['name','age','test'],
 *			[''lucky',19,'test']
 *		],
 *		[
 *			['id','name','score'],
 *			[1,'lucky',90],
 *			[2,'nancy',99]
 *		]
 * ]
 * $sheets array 多个sheet时的工作表名称 默认为该类型默认sheet名
 */
$excel->exportSheet($data,$sheets);

```

## Contributors

[Your contributions are always welcome!](https://github.com/openset/http/graphs/contributors)

## LICENSE

Released under [MIT](https://github.com/openset/http/blob/master/LICENSE) LICENSE
