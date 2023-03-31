<?php
require_once '../vendor/autoload.php';

use Chenvle\EasyExcel\EasyExcel;

$excel = new EasyExcel();


/*导入*/
$var   = [
    'A' => ['key' => 'username', 'title' => '姓名'],
    'B' => ['key' => 'phone', 'title' => '电话'],
    'C' => ['key' => 'origOrderId', 'title' => '订单号'],
    'D' => ['key' => 'txnAmt', 'title' => '订单金额'],
    'E' => ['key' => 'refundAmt', 'title' => '退款金额'],
    'F' => ['key' => 'remark', 'title' => '备注'],
];
$filepath = '';
// 1、excel表路径  2、每列对应数据键名及标题  3、内容开始的行
$excelData  = $excel->input($filepath, $var, 2);


/*导出*/
$field  = [
    'username'    => ['value' => '姓名', 'width' => '15', 'type' => 'string'],
    'school_str'  => ['value' => '学校', 'width' => '15', 'type' => 'string'],
    'grade'       => ['value' => '年级', 'width' => '10', 'type' => 'string'],
    'origOrderId' => ['value' => '订单号', 'width' => '25', 'type' => 'string'],
    'phone'       => ['value' => '电话', 'width' => '15','type'=>'int'],
    'txnAmt'      => ['value' => '订单金额', 'width' => '15'],
    'refundAmt'   => ['value' => '退款金额', 'width' => '15'],
    'status_str'  => ['value' => '状态', 'width' => '10', 'type' => 'string','color'=>'ff0000'],
    'remark'      => ['value' => '备注', 'width' => '15', 'type' => 'string','max'=>'50'],
    'create_time' => ['value' => '下单时间', 'width' => '20', 'type' => 'string']
];
$str    = '待退款表';
//导出的数据，如果是对象需要转换成数组，提高性能
$data = [];

//返回到浏览器，文件流形式
$excel->output($data, $str, $field,'Xlsx');

//保存到服务器返回路径(前面不用'/'，后面也不用"/")
$excel->output($data, $str, $field,'Xlsx','upload');