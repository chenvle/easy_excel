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
    'title'       => ['value' => '项目名称', 'width' => '30', 'type' => 'string', 'with' => 'orderInfo'],
    'phone'       => ['value' => '电话', 'width' => '15'],
    'txnAmt'      => ['value' => '订单金额', 'width' => '15'],
    'refundAmt'   => ['value' => '退款金额', 'width' => '15'],
    'status_str'  => ['value' => '状态', 'width' => '10', 'type' => 'string'],
    'remark'      => ['value' => '备注', 'width' => '15', 'type' => 'string'],
    'create_time' => ['value' => '下单时间', 'width' => '20', 'type' => 'string']
];
$str    = '待退款表';
$data = [];//导出的数据
$excel->output($data, $str, $field);//输出到浏览器
$url = $excel->output($data, $str, $field,'Xlsx','upload');//保存到服务器