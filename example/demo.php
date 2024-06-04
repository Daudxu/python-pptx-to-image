<?php
set_time_limit(3000);

$param1 = "test1.pptx";
$param2 = "output";

// 使用 escapeshellarg 来处理参数，防止命令注入
$escaped_param1 = escapeshellarg($param1);
$escaped_param2 = escapeshellarg($param2);

// 构建命令
$command = "python pptxToImg.py $escaped_param1 $escaped_param2";

// 执行命令并获取输出
shell_exec($command);
// $output = shell_exec($command);

// 输出结果
// echo $output;
echo "操作中请稍后..."
?>
