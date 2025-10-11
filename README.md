# EasyExcel

## 项目简介
EasyExcel是一个基于PhpSpreadsheet的PHP库，提供简单易用的Excel导入和导出功能。它封装了PhpSpreadsheet的复杂操作，使开发者能够更加便捷地处理Excel文件。

## 特性
- 简单易用的API接口
- 支持Excel文件的导入和导出
- 支持多种文件格式（Xlsx、Xls、Excel2007、Excel5）
- 支持自定义表头样式、单元格样式
- 支持多工作表操作
- 支持向Excel添加图片
- 优化的性能处理

## 安装
使用Composer安装：

```bash
composer require super-vle/easy-excel
```

## 基本使用

### 初始化

```php
use SuperVle\EasyExcel\Excel;

// 创建实例，可传入配置参数
$excel = new Excel();

// 或者使用自定义配置
$excel = new Excel([
    'autoSizeColumns' => true, // 自动计算列宽
    'defaultFont' => [
        'name' => 'Arial',
        'size' => 12,
        'color' => '000000'
    ],
    'headerStyle' => [
        'font' => [
            'bold' => true,
            'color' => '000000'
        ],
        'fill' => [
            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
            'color' => 'f0f0f0'
        ]
    ]
]);
```

### 导出Excel

```php
// 数据准备
$fields = [
    'username' => [
        'value' => '姓名',   // 表头显示名称
        'width' => '15',     // 列宽
        'type' => 'string'   // 数据类型
    ],
    'phone' => [
        'value' => '电话',
        'width' => '15',
        'type' => 'string'
    ],
    'origOrderId' => [
        'value' => '订单号',
        'width' => '25',
        'type' => 'string'
    ],
    'txnAmt' => [
        'value' => '订单金额',
        'width' => '15',
        'type' => 'string'
    ],
    'status' => [
        'value' => '状态',
        'width' => '10',
        'type' => 'string',
        'color' => 'ff0000' // 字体颜色
    ],
    'create_time' => [
        'value' => '创建时间',
        'width' => '20',
        'type' => 'string'
    ]
];

// 导出的数据
$data = [
    ['username' => '张三', 'phone' => '13800138000', 'origOrderId' => '202305200001', 'txnAmt' => '100.00', 'status' => '已付款', 'create_time' => '2023-05-20 10:00:00'],
    ['username' => '李四', 'phone' => '13800138001', 'origOrderId' => '202305200002', 'txnAmt' => '200.00', 'status' => '已发货', 'create_time' => '2023-05-20 11:00:00'],
    // 更多数据...
];

// 设置文件名
$excel->setFileName('订单列表_' . date('YmdHis'));

// 导出到浏览器（直接下载）
$excel->export($data, '订单列表', $fields, 'Xlsx', Excel::OUTPUT_TO_BROWSER);

// 或者保存到服务器文件
$filePath = $excel->export($data, '订单列表', $fields, 'Xlsx', Excel::OUTPUT_TO_FILE, '/path/to/save');
```

### 导入Excel

```php
// 定义Excel列与数据键的映射关系
$columnMap = [
    'A' => ['key' => 'username', 'title' => '姓名'],
    'B' => ['key' => 'phone', 'title' => '电话'],
    'C' => ['key' => 'origOrderId', 'title' => '订单号'],
    'D' => ['key' => 'txnAmt', 'title' => '订单金额'],
    'E' => ['key' => 'status', 'title' => '状态'],
    'F' => ['key' => 'create_time', 'title' => '创建时间']
];

// 导入Excel文件
$result = $excel->import('/path/to/excel.xlsx', $columnMap, 2); // 2表示数据从第2行开始

// 处理导入结果
if ($result['status']) {
    $importedData = $result['data']; // 导入的数据数组
    // 处理导入的数据
} else {
    echo '导入失败: ' . $result['msg'];
}

// 导入后保留文件（不删除）
$result = $excel->import('/path/to/excel.xlsx', $columnMap, 2, false);
```

### 多工作表操作

```php
// 准备多个工作表数据
$moreSheets = [
    [
        'title' => '用户列表',
        'data' => $userData,
        'field' => $userFields
    ],
    [
        'title' => '产品列表',
        'data' => $productData,
        'field' => $productFields
    ]
];

// 导出多工作表Excel
$excel->export($mainData, '主表', $mainFields, 'Xlsx', Excel::OUTPUT_TO_BROWSER, false, $moreSheets);
```

### 向Excel添加图片

```php
// 先导出Excel，然后添加图片
$excel->export($data, '带图片的报表', $fields, 'Xlsx', Excel::OUTPUT_TO_FILE, '/path/to/save');

// 添加图片
$excel->addImage('带图片的报表', '/path/to/image.jpg', 'A10', 100, 100); // 图片添加到A10单元格，宽100，高100
```

## API 文档

### 构造函数

```php
Excel(array $config = [])
```
- **参数**: `$config` - 配置数组
  - `autoSizeColumns`: 是否自动计算列宽（默认：false）
  - `defaultFont`: 默认字体设置
    - `name`: 字体名称
    - `size`: 字体大小
    - `color`: 字体颜色（RGB值）
  - `headerStyle`: 表头样式设置
    - `font`: 字体设置
    - `fill`: 填充设置

### 设置文件名

```php
setFileName(string $fileName): self
```
- **参数**: `$fileName` - 文件名（不含扩展名）
- **返回值**: 当前实例（支持链式调用）

### 导出Excel

```php
export(array $data, string $title, array $fields, string $fileType = 'Xlsx', string $outputType = 'browser', $filePath = false, $moreSheets = false): string|void
```
- **参数**: 
  - `$data`: 要导出的数据数组
  - `$title`: 工作表标题
  - `$fields`: 字段配置数组
  - `$fileType`: 文件类型（Xlsx、Xls、Excel2007、Excel5）
  - `$outputType`: 输出类型（browser或file）
  - `$filePath`: 保存文件的路径（当outputType为file时有效）
  - `$moreSheets`: 更多工作表数据
- **返回值**: 文件路径或直接输出到浏览器

### 导入Excel

```php
import($excelPath, $param, $startRow, $delFile = true): array
```
- **参数**: 
  - `$excelPath`: Excel文件路径
  - `$param`: 列映射配置
  - `$startRow`: 数据开始的行号
  - `$delFile`: 是否删除文件（默认：true）
- **返回值**: 结果数组
  - `status`: 是否成功
  - `msg`: 消息
  - `data`: 导入的数据

### 添加图片

```php
addImage($sheetName, $imagePath, $coordinate, $width = 100, $height = 100): self
```
- **参数**: 
  - `$sheetName`: 工作表名称
  - `$imagePath`: 图片路径
  - `$coordinate`: 单元格坐标
  - `$width`: 图片宽度
  - `$height`: 图片高度
- **返回值**: 当前实例（支持链式调用）

## 字段配置说明

### 导出字段配置

```php
$fields = [
    'data_key' => [
        'value' => '字段标题',  // 必需：表头显示的标题
        'width' => '15',        // 可选：列宽
        'type' => 'string',     // 可选：数据类型（string、array）
        'max' => 50,            // 可选：最大长度，超过将被截断
        'color' => 'ff0000'     // 可选：字体颜色（RGB值）
    ]
];
```

### 导入列映射配置

```php
$columnMap = [
    'A' => [                  // Excel列字母
        'key' => 'data_key',  // 映射到数据数组的键名
        'title' => '字段标题'   // Excel表头标题（用于验证）
    ]
];
```

## 注意事项

1. 导出时，数据应该是二维数组，每行是一个关联数组，键名对应字段配置中的键。
2. 导入时，确保Excel文件的表头与配置中的title一致，否则会验证失败。
3. PHP版本需要7.4.0或更高。
4. 处理大量数据时，可能需要调整PHP的内存限制。
5. 旧版的`output()`和`input()`方法已被标记为废弃，建议使用`export()`和`import()`方法。

## 兼容性说明

- 兼容PHP 7.4及以上版本
- 基于PhpSpreadsheet ^1.18开发
- 支持Excel 2007及以上版本格式

## 许可证

本项目采用MIT许可证。详情请查看[LICENSE](LICENSE)文件。

## 作者

chenvle <chenvle@cmtime.cn>
