简易Excel导入和导出

1、安装
```
composer require chenvle/easy_excel
```

2、使用例子
```
use Chenvle\EasyExcel\EasyExcel;

$excel = new Excel();
```


```
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
```
```
/*导入的格式*/
$var = [
    '列'=>['key'=>'英文命名','title'=>'中文命名'],
    ...
]
注意：
    中文命名必须跟文档中的命名一致，英文命名可自定义
```


```
/*导出*/
$field  = [
    'username'    => ['value' => '姓名', 'width' => '15', 'type' => 'string'],
    'school_str'  => ['value' => '学校', 'width' => '15', 'type' => 'string'],
    'grade'       => ['value' => '年级', 'width' => '10', 'type' => 'string'],
    'origOrderId' => ['value' => '订单号', 'width' => '25', 'type' => 'string'],
    'phone'       => ['value' => '电话', 'width' => '15'],
    'txnAmt'      => ['value' => '订单金额', 'width' => '15'],
    'refundAmt'   => ['value' => '退款金额', 'width' => '15'],
    'status_str'  => ['value' => '状态', 'width' => '10', 'type' => 'string'],
    'remark'      => ['value' => '备注', 'width' => '15', 'type' => 'string'],
    'create_time' => ['value' => '下单时间', 'width' => '20', 'type' => 'string']
];
$str    = '待退款表';
$data = [];//导出的数据
$excel->output($data, $str, $field,'Xlsx');//输出到浏览器
$excel->output($data, $str, $field,'Xlsx','upload');//保存到服务器返回路径(前面不用'/'，后面也不用"/")
```

```
/*导出field参数*/
$field = [
    '字段'=>[
        'value'=>'字段命名',
        'width'=>'宽度',
        'type'=>'string字符串|int整型|array数组',
        'max'=>'长度,超过截取'
        ]
    ...
]
注意：
    字段必须要存在数据中，value字段命名用于导出的命名可自定义，type类型中string会自动在后面增加空格，
    int不做处理，array用'、'分隔，max长度用于截取长字段
```
