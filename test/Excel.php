<?php
require_once '../vendor/autoload.php';

use SuperVle\EasyExcel\Excel;

$excel = new Excel([
    'autoSizeColumns' => true, // 启用自动列宽调整
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
            'color' => 'e6e6e6'
        ]
    ]
]);

// 设置文件名 - 支持链式调用
$excel->setFileName('测试文件');

/************************
 * 导入功能示例
 ************************/
$importParams = [
    'A' => ['key' => 'username', 'title' => '姓名'],
    'B' => ['key' => 'phone', 'title' => '电话'],
    'C' => ['key' => 'origOrderId', 'title' => '订单号'],
    'D' => ['key' => 'txnAmt', 'title' => '订单金额'],
    'E' => ['key' => 'refundAmt', 'title' => '退款金额'],
    'F' => ['key' => 'remark', 'title' => '备注'],
];

// 导入示例（实际使用时需要提供有效的文件路径）
// $filePath = './uploads/test.xlsx';
// try {
//     $excelData = $excel->import($filePath, $importParams, 2, true);
//     if ($excelData['status']) {
//         // 处理导入的数据
//         print_r($excelData['data']);
//     } else {
//         echo '导入失败: ' . $excelData['msg'];
//     }
// } catch (Exception $e) {
//     echo '导入异常: ' . $e->getMessage();
// }

/************************
 * 导出功能示例
 ************************/
// 字段配置
$fields = [
    'username'    => ['value' => '姓名', 'width' => '15', 'type' => 'string'],
    'school_str'  => ['value' => '学校', 'width' => '15', 'type' => 'string'],
    'grade'       => ['value' => '年级', 'width' => '10', 'type' => 'string'],
    'origOrderId' => ['value' => '订单号', 'width' => '25', 'type' => 'string'],
    'phone'       => ['value' => '电话', 'width' => '15', 'type' => 'string'],
    'txnAmt'      => ['value' => '订单金额', 'width' => '15'],
    'refundAmt'   => ['value' => '退款金额', 'width' => '15'],
    'status_str'  => ['value' => '状态', 'width' => '10', 'type' => 'string', 'color' => 'ff0000'],
    'remark'      => ['value' => '备注', 'width' => '25', 'type' => 'string', 'max' => '50'],
    'create_time' => ['value' => '下单时间', 'width' => '20', 'type' => 'string']
];

// 模拟数据
$title = '订单报表';
$data = [
    [
        'username' => '张三',
        'school_str' => '第一中学',
        'grade' => '高三',
        'origOrderId' => 'ORD20230515001',
        'phone' => '13800138001',
        'txnAmt' => '199.00',
        'refundAmt' => '0.00',
        'status_str' => '已完成',
        'remark' => '正常订单，已发货',
        'create_time' => '2023-05-15 10:30:25'
    ],
    [
        'username' => '李四',
        'school_str' => '第二中学',
        'grade' => '高二',
        'origOrderId' => 'ORD20230515002',
        'phone' => '13800138002',
        'txnAmt' => '299.00',
        'refundAmt' => '299.00',
        'status_str' => '已退款',
        'remark' => '客户取消订单，全额退款',
        'create_time' => '2023-05-15 11:20:13'
    ],
    // 更多数据...
];

/************************
 * 基础导出示例
 ************************/
// 导出到浏览器下载（使用新的API）
// $excel->export($data, $title, $fields);

/************************
 * 保存到服务器示例
 ************************/
// 保存到服务器并获取文件路径
// try {
//     $filePath = $excel->export(
//         $data, 
//         $title, 
//         $fields, 
//         Excel::DEFAULT_FILE_TYPE, 
//         Excel::OUTPUT_TO_FILE, 
//         './exports'
//     );
//     echo '文件已保存到: ' . $filePath;
// } catch (Exception $e) {
//     echo '导出失败: ' . $e->getMessage();
// }

/************************
 * 多工作表导出示例
 ************************/
// 多工作表配置
$moreSheets = [
    [
        'title' => '已完成订单',
        'data' => array_filter($data, function($item) {
            return $item['status_str'] == '已完成';
        }),
        'field' => $fields
    ],
    [
        'title' => '已退款订单',
        'data' => array_filter($data, function($item) {
            return $item['status_str'] == '已退款';
        }),
        'field' => $fields
    ]
];

// 导出多工作表Excel
// $excel->export($data, $title, $fields, Excel::DEFAULT_FILE_TYPE, Excel::OUTPUT_TO_BROWSER, false, $moreSheets);

/************************
 * 添加图片示例
 ************************/
// 先创建一个工作表，然后添加图片
// try {
//     // 首先创建并构建一个工作表
//     $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
//     $sheet = $spreadsheet->getActiveSheet();
//     $sheet->setTitle('图片演示');
//     $sheet->setCellValue('A1', '图片示例');
//     
//     // 将工作表设置到excel实例
//     $excel->{'spreadsheet'} = $spreadsheet; // 注意：这是一个临时解决方案，正式版本可能会提供更好的API
//     
//     // 添加图片
//     $excel->addImage('图片演示', './logo.png', 'B3', 200, 100);
//     
//     // 输出Excel
//     $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
//     $excel->excelBrowserExport('带图片的Excel', 'Xlsx');
//     $writer->save('php://output');
// } catch (Exception $e) {
//     echo '添加图片失败: ' . $e->getMessage();
// }

/************************
 * 兼容旧版API示例
 ************************/
// 使用旧版API导出到浏览器
// $excel->output($data, $title, $fields, 'Xlsx');

// 使用旧版API保存到服务器
// $filePath = $excel->output($data, $title, $fields, 'Xlsx', './uploads', $moreSheets);
// echo '文件已保存到: ' . $filePath;