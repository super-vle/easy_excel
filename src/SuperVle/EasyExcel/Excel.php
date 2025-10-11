<?php

namespace SuperVle\EasyExcel;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use Exception;

class Excel
{
    /**
     * 默认文件类型
     */
    const DEFAULT_FILE_TYPE = 'Xlsx';
    
    /**
     * 输出到浏览器
     */
    const OUTPUT_TO_BROWSER = 'browser';
    
    /**
     * 保存到文件
     */
    const OUTPUT_TO_FILE = 'file';
    
    /**
     * 支持的文件类型
     */
    protected $supportedFileTypes = ['Xlsx', 'Xls', 'Excel2007', 'Excel5'];
    
    /**
     * 文件名
     * @var string
     */
    protected $fileName;
    
    /**
     * 工作表对象
     * @var Spreadsheet
     */
    protected $spreadsheet;
    
    /**
     * 配置项
     * @var array
     */
    protected $config = [
        // 是否自动计算列宽
        'autoSizeColumns' => false,
        // 默认字体
        'defaultFont' => [
            'name' => 'Verdana',
            'size' => 11,
            'color' => '000000'
        ],
        // 默认表头样式
        'headerStyle' => [
            'font' => [
                'bold' => true,
                'color' => '000000'
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'color' => 'f0f0f0'
            ]
        ]
    ];
    
    /**
     * Excel constructor.
     * @param array $config 配置项
     */
    public function __construct(array $config = [])
    {
        $this->fileName = time();
        $this->config = array_merge($this->config, $config);
    }
    
    /**
     * 设置文件名
     * @param string $fileName
     * @return $this
     */
    public function setFileName(string $fileName)
    {
        $this->fileName = $fileName;
        return $this;
    }
    
    /**
     * 获取当前文件名
     * @return string
     */
    public function getFileName()
    {
        return $this->fileName;
    }
    
    /**
     * 导出Excel
     * 
     * @param array $data 数据
     * @param string $title 标题
     * @param array $fields 字段配置
     * @param string $fileType 文件类型
     * @param string $outputType 输出类型
     * @param string|false $filePath 文件路径(当outputType为file时有效)
     * @param array|false $moreSheets 更多工作表
     * @return string|void 文件路径或直接输出
     * @throws Exception
     */
    public function export(array $data, string $title, array $fields, 
                          string $fileType = self::DEFAULT_FILE_TYPE,
                          string $outputType = self::OUTPUT_TO_BROWSER,
                          $filePath = false,
                          $moreSheets = false)
    {
        // 参数验证
        $this->validateExportParams($data, $title, $fields, $fileType, $outputType);
        
        // 创建工作表
        $this->spreadsheet = new Spreadsheet();
        
        // 设置默认字体
        $this->setDefaultFont();
        
        // 处理第一个工作表
        $sheet = $this->spreadsheet->getActiveSheet();
        $this->buildSheet($sheet, $data, $title, $fields);
        
        // 处理更多工作表
        if ($moreSheets) {
            $this->buildMoreSheets($moreSheets);
        }
        
        // 输出
        return $this->outputExcel($fileType, $outputType, $filePath);
    }
    
    /**
     * 兼容旧版output方法
     * 
     * @deprecated 建议使用export方法
     * @param array $data
     * @param string $title
     * @param array $fields
     * @param string $fileType
     * @param string $file_or_url
     * @param array|false $more_table
     * @return string|void
     * @throws Exception
     */
    public function output(array $data, string $title, array $fields, 
                          string $fileType = 'Xlsx',
                          string $file_or_url = 'file',
                          $more_table = false)
    {
        // 转换参数兼容旧版
        $outputType = ($file_or_url === 'file') ? self::OUTPUT_TO_BROWSER : self::OUTPUT_TO_FILE;
        $filePath = ($file_or_url !== 'file') ? $file_or_url : false;
        
        return $this->export($data, $title, $fields, $fileType, $outputType, $filePath, $more_table);
    }
    
    /**
     * 验证导出参数
     * @param array $data
     * @param string $title
     * @param array $fields
     * @param string $fileType
     * @param string $outputType
     * @throws Exception
     */
    protected function validateExportParams(array $data, string $title, array $fields, string $fileType, string $outputType)
    {
        // 验证文件类型
        if (!in_array($fileType, $this->supportedFileTypes)) {
            throw new Exception('不支持的文件类型: ' . $fileType);
        }
        
        // 验证输出类型
        if (!in_array($outputType, [self::OUTPUT_TO_BROWSER, self::OUTPUT_TO_FILE])) {
            throw new Exception('不支持的输出类型: ' . $outputType);
        }
        
        // 验证字段配置
        if (empty($fields)) {
            throw new Exception('字段配置不能为空');
        }
    }
    
    /**
     * 设置默认字体
     */
    protected function setDefaultFont()
    {
        $defaultFont = $this->config['defaultFont'];
        $this->spreadsheet->getDefaultStyle()->getFont()
            ->setName($defaultFont['name'])
            ->setSize($defaultFont['size']);
        
        $this->spreadsheet->getDefaultStyle()->getFont()->getColor()
            ->setRGB($defaultFont['color']);
    }
    
    /**
     * 构建工作表
     * @param $sheet
     * @param array $data
     * @param string $title
     * @param array $fields
     */
    protected function buildSheet($sheet, array $data, string $title, array $fields)
    {
        // 设置工作表标题
        if ($title) {
            $sheet->setTitle($this->truncateSheetTitle($title));
        }
        
        // 获取列字母数组
        $columnLetters = $this->getColumnLetters(count($fields));
        $endColumn = end($columnLetters);
        $endRow = count($data) + 2;
        
        // 标题行
        $titleRow = 1;
        if ($title) {
            $sheet->setCellValue('A1', $title);
            $sheet->getRowDimension('1')->setRowHeight(30);
            $sheet->mergeCells('A1:' . $endColumn . '1');
            
            // 设置标题样式
            $titleStyle = [
                'font' => [
                    'bold' => true,
                    'color' => ['rgb' => '000000'],
                    'size' => 16,
                    'name' => 'Verdana'
                ],
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_CENTER,
                    'vertical' => Alignment::VERTICAL_CENTER
                ]
            ];
            $sheet->getStyle('A1:' . $endColumn . '1')->applyFromArray($titleStyle);
        }
        
        // 表头行
        $headerRow = $title ? 2 : 1;
        $this->buildHeaderRow($sheet, $fields, $columnLetters, $headerRow);
        
        // 数据行 - 使用批量设置提升性能
        $this->buildDataRows($sheet, $data, $fields, $columnLetters, $headerRow);
        
        // 设置整体边框和居中
        $styleArrayBody = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => '666666'],
                ],
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER
            ],
        ];
        
        $startCell = 'A' . $headerRow;
        $endCell = $endColumn . $endRow;
        $sheet->getStyle($startCell . ':' . $endCell)->applyFromArray($styleArrayBody);
        
        // 自动调整列宽
        if ($this->config['autoSizeColumns']) {
            foreach ($columnLetters as $columnLetter) {
                $sheet->getColumnDimension($columnLetter)->setAutoSize(true);
            }
        }
    }
    
    /**
     * 截断工作表标题（Excel工作表名称限制为31个字符）
     * @param string $title
     * @return string
     */
    protected function truncateSheetTitle(string $title)
    {
        return mb_strlen($title) > 31 ? mb_substr($title, 0, 28) . '...' : $title;
    }
    
    /**
     * 获取列字母数组
     * @param int $count
     * @return array
     */
    protected function getColumnLetters(int $count)
    {
        $letters = [];
        $start = ord('A');
        
        for ($i = 0; $i < $count; $i++) {
            $letters[] = chr($start + $i);
        }
        
        return $letters;
    }
    
    /**
     * 构建表头行
     * @param $sheet
     * @param array $fields
     * @param array $columnLetters
     * @param int $headerRow
     */
    protected function buildHeaderRow($sheet, array $fields, array $columnLetters, int $headerRow)
    {
        $headerStyle = $this->config['headerStyle'];
        
        // 使用数值索引遍历，避免使用fields数组的键作为索引
        $fieldKeys = array_keys($fields);
        foreach ($fieldKeys as $index => $fieldKey) {
            $field = $fields[$fieldKey];
            $columnLetter = $columnLetters[$index];
            $cellAddress = $columnLetter . $headerRow;
            
            // 设置表头值
            $sheet->setCellValue($cellAddress, $field['value'] ?? '');
            
            // 设置列宽
            if (isset($field['width']) && is_numeric($field['width'])) {
                $sheet->getColumnDimension($columnLetter)->setWidth((float)$field['width']);
            }
            
            // 应用表头样式
            $style = [];
            if (isset($headerStyle['font'])) {
                $style['font'] = $headerStyle['font'];
                // 检查并修复字体颜色格式
                if (isset($style['font']['color']) && is_string($style['font']['color'])) {
                    $style['font']['color'] = ['rgb' => $style['font']['color']];
                }
            }
            if (isset($headerStyle['fill'])) {
                $style['fill'] = $headerStyle['fill'];
                // 检查并修复填充颜色格式
                if (isset($style['fill']['color']) && is_string($style['fill']['color'])) {
                    $style['fill']['color'] = ['rgb' => $style['fill']['color']];
                }
            }
            
            if (!empty($style)) {
                $sheet->getStyle($cellAddress)->applyFromArray($style);
            }
        }
    }
    
    /**
     * 构建数据行
     * @param $sheet
     * @param array $data
     * @param array $fields
     * @param array $columnLetters
     * @param int $headerRow
     */
    protected function buildDataRows($sheet, array $data, array $fields, array $columnLetters, int $headerRow)
    {
        // 批量设置数据，提升性能
        $batchData = [];
        $styleConfigurations = [];
        
        // 获取所有字段键，避免在循环中重复调用array_keys
        $fieldKeys = array_keys($fields);
        
        foreach ($data as $rowIndex => $rowData) {
            $excelRowIndex = $rowIndex + $headerRow + 1;
            $excelRow = [];
            
            foreach ($fieldKeys as $fieldIndex => $fieldKey) {
                $field = $fields[$fieldKey];
                $value = '';
                
                // 获取单元格值
                if (isset($field['with']) && $field['with'] && isset($rowData[$field['with']][$fieldKey])) {
                    $value = $rowData[$field['with']][$fieldKey];
                } else if (isset($rowData[$fieldKey])) {
                    $value = $rowData[$fieldKey];
                }
                
                // 处理特殊类型
                $value = $this->processCellValue($value, $field);
                
                $excelRow[] = $value;
                
                // 收集样式配置
                $columnLetter = $columnLetters[$fieldIndex];
                $cellAddress = $columnLetter . $excelRowIndex;
                
                if (isset($field['color'])) {
                    $styleConfigurations[$cellAddress] = [
                        'font' => [
                            'color' => ['rgb' => $field['color']],
                        ]
                    ];
                }
            }
            
            $batchData[] = $excelRow;
        }
        
        // 批量设置数据
        if (!empty($batchData)) {
            $startCell = 'A' . ($headerRow + 1);
            $sheet->fromArray($batchData, null, $startCell);
        }
        
        // 应用样式配置
        foreach ($styleConfigurations as $cellAddress => $style) {
            $sheet->getStyle($cellAddress)->applyFromArray($style);
        }
    }
    
    /**
     * 处理单元格值
     * @param mixed $value
     * @param array $field
     * @return mixed
     */
    protected function processCellValue($value, array $field)
    {
        // 字符串类型后面加空格
        if (isset($field['type']) && $field['type'] == 'string') {
            $value = $value . ' ';
        }
        
        // 数组类型用'、'分隔
        if (isset($field['type']) && $field['type'] == 'array' && is_array($value)) {
            $value = implode('、', $value);
        }
        
        // 长数据截取
        if (isset($field['max']) && is_numeric($field['max']) && $field['max'] > 0) {
            if (mb_strlen($value) > $field['max']) {
                $value = mb_substr($value, 0, $field['max']);
            }
        }
        
        return $value;
    }
    
    /**
     * 构建更多工作表
     * @param array $moreSheets
     */
    protected function buildMoreSheets(array $moreSheets)
    {
        // 如果是一维数组，当作单个工作表处理
        if (isset($moreSheets['data']) && isset($moreSheets['field'])) {
            $this->createAdditionalSheet($moreSheets['data'], $moreSheets['title'] ?? '', $moreSheets['field']);
            return;
        }
        
        // 处理多维数组多个工作表
        foreach ($moreSheets as $sheetData) {
            if (isset($sheetData['data']) && isset($sheetData['field'])) {
                $this->createAdditionalSheet($sheetData['data'], $sheetData['title'] ?? '', $sheetData['field']);
            }
        }
    }
    
    /**
     * 创建额外的工作表
     * @param array $data
     * @param string $title
     * @param array $fields
     */
    protected function createAdditionalSheet(array $data, string $title, array $fields)
    {
        $sheetIndex = $this->spreadsheet->createSheet();
        $sheet = $this->spreadsheet->setActiveSheetIndex($sheetIndex);
        $this->buildSheet($sheet, $data, $title, $fields);
    }
    
    /**
     * 输出Excel
     * @param string $fileType
     * @param string $outputType
     * @param string|false $filePath
     * @return string|void
     * @throws Exception
     */
    protected function outputExcel(string $fileType, string $outputType, $filePath)
    {
        $writer = IOFactory::createWriter($this->spreadsheet, $fileType);
        
        // 清理内存
        ob_start();
        
        if ($outputType === self::OUTPUT_TO_BROWSER) {
            $this->excelBrowserExport($this->fileName, $fileType);
            $writer->save('php://output');
        } else {
            if (!$filePath) {
                throw new Exception('保存文件时必须提供文件路径');
            }
            
            $filename = rtrim($filePath, '/') . '/' . $this->fileName . '.' . strtolower($fileType);
            
            // 确保目录存在
            $dir = dirname($filename);
            if (!is_dir($dir) && !mkdir($dir, 0777, true)) {
                throw new Exception('无法创建目录: ' . $dir);
            }
            
            $writer->save($filename);
            return '/' . ltrim($filename, '/');
        }
        
        // 释放资源
        $this->spreadsheet->disconnectWorksheets();
        unset($this->spreadsheet);
        
        ob_end_flush();
        exit;
    }
    
    /**
     * 导入Excel表取出需要的内容
     * @param $excelPath //excel表路径
     * @param $param //每列对应数据键名及标题 ['A' => ['key' => 'A',title => '标题名称']] 标题名为空则不验证
     * @param $startRow //内容开始的行
     * @param bool $delFile //是否删除文件
     * @return array    //返回数据内容 [['A' => 'content']];
     * @throws Exception
     */
    public function import($excelPath, $param, $startRow, $delFile = true)
    {
        try {
            $excelObj = IOFactory::load($excelPath);
            if (!$excelObj) {
                return $this->error('加载Excel表失败，请检查Excel内容');
            }
            
            $excelWorkSheet = $excelObj->getActiveSheet();
            $rowCount = $excelWorkSheet->getHighestRow();
            
            if ($rowCount <= 0) {
                return $this->error('Excel表内容为空。');
            }
            
            // 验证标题
            foreach ($param as $column => $content) {
                $item = $excelWorkSheet->getCell($column . ($startRow - 1))->getCalculatedValue();
                if ($item != $content['title'] && !empty($content['title'])) {
                    return $this->error('请检查模板标题是否正确。');
                }
            }
            
            $excelData = [];
            for ($row = $startRow; $row <= $rowCount; $row++) {
                $rowData = [];
                foreach ($param as $column => $content) {
                    $item = $excelWorkSheet->getCell($column . $row)->getCalculatedValue();
                    $rowData[$content['key']] = $item;
                }
                
                // 删除空行
                if (!implode('', $rowData)) {
                    continue;
                }
                
                $excelData[] = $rowData;
            }
            
            if ($delFile) {
                unlink($excelPath);
            }
            
            return $this->success($excelData);
        } catch (Exception $e) {
            return $this->error('导入Excel失败: ' . $e->getMessage());
        }
    }
    
    /**
     * 兼容旧版input方法
     * 
     * @deprecated 建议使用import方法
     * @param $excelPath
     * @param $param
     * @param $startRow
     * @param bool $del_file
     * @return array
     */
    public function input($excelPath, $param, $startRow, $del_file = true)
    {
        return $this->import($excelPath, $param, $startRow, $del_file);
    }
    
    /**
     * 输出到浏览器(需要设置header头)
     * @param string $fileName 文件名
     * @param string $fileType 文件类型
     */
    protected function excelBrowserExport($fileName, $fileType)
    {
        // 文件名称校验
        if (!$fileName) {
            trigger_error('文件名不能为空', E_USER_ERROR);
        }
        
        // 清除之前的输出
        if (ob_get_length()) {
            ob_clean();
        }
        
        // 设置header
        if ($fileType == 'Excel2007' || $fileType == 'Xlsx') {
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename="' . $fileName . '.xlsx"');
        } else { //Excel5
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="' . $fileName . '.xls"');
        }
        
        header('Cache-Control: max-age=0');
        header('Cache-Control: no-cache');
        header('Pragma: no-cache');
    }
    
    /**
     * 错误响应
     * @param string $info
     * @return array
     */
    protected function error($info)
    {
        return ['msg' => $info, 'status' => false, 'data' => false];
    }
    
    /**
     * 成功响应
     * @param mixed $data
     * @return array
     */
    protected function success($data)
    {
        return ['msg' => 'ok', 'status' => true, 'data' => $data];
    }
    
    /**
     * 添加图片到Excel
     * @param string $sheetName 工作表名称
     * @param string $imagePath 图片路径
     * @param string $coordinate 单元格坐标
     * @param int $width 宽度
     * @param int $height 高度
     * @return $this
     * @throws Exception
     */
    public function addImage($sheetName, $imagePath, $coordinate, $width = 100, $height = 100)
    {
        if (!isset($this->spreadsheet)) {
            $this->spreadsheet = new Spreadsheet();
        }
        
        // 查找工作表
        $sheetFound = false;
        for ($i = 0; $i < $this->spreadsheet->getSheetCount(); $i++) {
            if ($this->spreadsheet->getSheet($i)->getTitle() === $sheetName) {
                $sheet = $this->spreadsheet->getSheet($i);
                $sheetFound = true;
                break;
            }
        }
        
        if (!$sheetFound) {
            throw new Exception('未找到工作表: ' . $sheetName);
        }
        
        if (!file_exists($imagePath)) {
            throw new Exception('图片文件不存在: ' . $imagePath);
        }
        
        $drawing = new Drawing();
        $drawing->setName(basename($imagePath));
        $drawing->setDescription(basename($imagePath));
        $drawing->setPath($imagePath);
        $drawing->setCoordinates($coordinate);
        $drawing->setWidth($width);
        $drawing->setHeight($height);
        $drawing->setWorksheet($sheet);
        
        return $this;
    }
}
