<?php
include_once("xlsxwriter.class.php");
ini_set('display_errors', 0);
ini_set('log_errors', 1);
error_reporting(E_ALL & ~E_NOTICE);

$list = [];
for ($i = 0; $i < 5; $i++) {
    $list[] = ['2015-01-01', '1', '-50.5', '2010-01-01 23:00:00', '2012-12-31 23:00:00', '=D2'];
}

$data = [
    # 文件名称
    'filename'    => '表格文件名', #文件名
    'is_download' => true, # 是否下载
    'is_save'     => false,  # 是否保存文件  同级目录下
    'title'       => '标题',
    'subject'     => '主题',
    'author'      => '作者名字',
    'company'     => '公司名字',
    'keywords'    => '关键字',
    'description' => '描述',
    'temp_dir'    => '临时目录',
    #工作表
    'sheet'       => [
        'sheet1' => [
            # 列样式
            'col_style'   => [
                'widths'       => [20, 30, 20, 40, 40],  # 宽度
                'auto_filter'  => true,  # 筛选
                // 'freeze_rows'=>true,  # 冻结
                // 'freeze_columns'=>true,  # 冻结
                // 'suppress_row' => false,
                'border'       => 'left,right,top,bottom', # 边界 left, right, top, bottom, or multiple ie: 'top,left'
                'border-style' => 'thin',  # 边框样式 thin, medium, thick, dashDot, dashDotDot, dashed, dotted, double, hair, mediumDashDot, mediumDashDotDot, mediumDashed, slantDashDot

            ],
            # 标题
            'table_title' => [
                'title'  => ['这是一个大标题'],
                'length' => 5, # 标题合并的单元格数量
                'style'  => [
                    'font'         => 'Arial',
                    'font-size'    => 18,
                    'font-style'   => 'bold',  #bold, italic, underline, strikethrough or multiple ie: 'bold,italic'
                    'color'        => '#333',
                    'fill'         => '#fff',  # 背景填充
                    'halign'       => 'center',  # 水平位置 general, left, right, justify, center
                    'border'       => 'left,right,top,bottom', # 边界 left, right, top, bottom, or multiple ie: 'top,left'
                    'border-style' => 'medium',  # 边框样式 thin, medium, thick, dashDot, dashDotDot, dashed, dotted, double, hair, mediumDashDot, mediumDashDotDot, mediumDashed, slantDashDot
                    'border-color' => '#FF0016',  # 边框颜色 #RRGGBB, ie: #ff99cc or #f9c
                    'valign'       => 'center', # 垂直位置 bottom, center, distributed

                ],
            ],
            # 表头
            'header'      => [
                # 列的类型
                'col_type' => [
                    'created'     => 'string',
                    'product_id'  => 'date',
                    'quantity'    => '#,##0.00', #价格 #,##0.00表示小数位两个，减少或增加改变长度
                    'amount'      => 'price',
                    'description' => 'string',
                    'tax'         => '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00',
                ],
                'style'    => [
                    // 'suppress_row' => false, # 抑制行
                    'font'         => 'Arial',
                    'font-size'    => 18,
                    'font-style'   => 'bold',  #bold, italic, underline, strikethrough or multiple ie: 'bold,italic'
                    'color'        => '#333',
                    'fill'         => '#fff',  # 背景填充
                    'halign'       => 'center',  # 水平位置 general, left, right, justify, center
                    'border'       => 'left,right,top,bottom', # 边界 left, right, top, bottom, or multiple ie: 'top,left'
                    'border-style' => 'medium',  # 边框样式 thin, medium, thick, dashDot, dashDotDot, dashed, dotted, double, hair, mediumDashDot, mediumDashDotDot, mediumDashed, slantDashDot
                    'border-color' => '#FF0016',  # 边框颜色 #RRGGBB, ie: #ff99cc or #f9c
                    'valign'       => 'center', # 垂直位置 bottom, center, distributed
                ],
            ],
            # 数据行
            'rows'        => [
                'data'  => $list,
                # 样式
                'style' => [
                    'font'         => 'Arial',
                    'font-size'    => 12,
                    'font-style'   => 'bold',  #bold, italic, underline, strikethrough or multiple ie: 'bold,italic'
                    'color'        => '#333',
                    'fill'         => '#fff',  # 背景填充
                    'halign'       => 'center',  # 水平位置 general, left, right, justify, center
                    'border'       => 'left,right,top,bottom', # 边界 left, right, top, bottom, or multiple ie: 'top,left'
                    'border-style' => 'thin',  # 边框样式 thin, medium, thick, dashDot, dashDotDot, dashed, dotted, double, hair, mediumDashDot, mediumDashDotDot, mediumDashed, slantDashDot
                    'border-color' => '#333',  # 边框颜色 #RRGGBB, ie: #ff99cc or #f9c
                    'valign'       => 'center', # 垂直位置 bottom, center, distributed
                    'height'       => 50,  # 行高
                    // 'collapsed'       => true,  # 未知
                    // 'hidden'       => true,  # 隐藏行

                ],
            ],

        ],
    ],

];

$writer = new XLSXWriter($data);
exit(0);

