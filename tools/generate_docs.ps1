param(
    [string]$Root = (Get-Location).Path
)

$ErrorActionPreference = 'Stop'

Set-Location $Root
New-Item -ItemType Directory -Force docs | Out-Null

$source = Get-Content explain.cpp -Raw
$catalogMatch = [regex]::Match(
    $source,
    'std::vector<CommandInfo> buildCommandCatalog\(\) \{\s*return \{(?<body>[\s\S]*?)\n\s*\};\s*\}')
if (-not $catalogMatch.Success) {
    throw '无法从 explain.cpp 中定位 buildCommandCatalog()'
}

$entryPattern = '\{"(?<category>[^"]+)",\s*"(?<name>[^"]+)",\s*"(?<syntax>[^"]+)",\s*"(?<desc>[^"]+)",\s*(?<impl>true|false)\}'
$items = [regex]::Matches($catalogMatch.Groups['body'].Value, $entryPattern) | ForEach-Object {
    [PSCustomObject]@{
        category = $_.Groups['category'].Value
        name = $_.Groups['name'].Value
        syntax = $_.Groups['syntax'].Value
        desc = $_.Groups['desc'].Value
        implemented = ($_.Groups['impl'].Value -eq 'true')
    }
}

$categoryOrder = @('鼠标', '键盘', '系统', '窗口', '文件', '变量', '流程', '交互', '网络', '时间', '图像', '调试')

function Add-UniqueNote {
    param(
        [System.Collections.Generic.List[string]]$List,
        [string]$Text
    )
    if (-not [string]::IsNullOrWhiteSpace($Text) -and -not $List.Contains($Text)) {
        [void]$List.Add($Text)
    }
}

function Get-ParameterNotes {
    param($Item)

    $syntax = $Item.syntax
    $notes = New-Object 'System.Collections.Generic.List[string]'

    if ($syntax -match '左键\|右键') {
        Add-UniqueNote $notes '鼠标按键参数支持 `左键` 或 `右键`。'
    }
    if ($syntax -match '单点\|双击\|多击\|长按') {
        Add-UniqueNote $notes '点击模式参数支持 `单点`、`双击`、`多击`、`长按`。'
    }
    if ($syntax -match '线性\|缓入\|缓出\|缓入缓出') {
        Add-UniqueNote $notes '轨迹参数支持 `线性`、`缓入`、`缓出`、`缓入缓出`。'
    }
    if ($syntax -match '\(x,y\)') {
        Add-UniqueNote $notes '`(x,y)` 表示单个屏幕坐标。'
    }
    if ($syntax -match '\(x1,y1\)->\(x2,y2\)') {
        Add-UniqueNote $notes '`(x1,y1)->(x2,y2)` 表示一个矩形区域或起止点范围。'
    }
    if ($syntax -match '到 变量名|到 名称') {
        Add-UniqueNote $notes '`到 变量名` 或 `到 名称` 表示把执行结果写回变量。'
    }
    if ($syntax -match '秒数') {
        Add-UniqueNote $notes '`秒数` 支持整数或小数。'
    }
    if ($syntax -match '次数') {
        Add-UniqueNote $notes '`次数` 使用整数。'
    }
    if ($syntax -match '频率') {
        Add-UniqueNote $notes '`频率` 单位为 Hz。'
    }
    if ($syntax -match '毫秒') {
        Add-UniqueNote $notes '`毫秒` 单位为毫秒。'
    }
    if ($syntax -match '路径|旧路径|新路径|源|目标|目录|压缩包|文件路径|图片路径') {
        Add-UniqueNote $notes '涉及文件或目录时，支持相对路径和绝对路径。'
    }
    if ($syntax -match '网址') {
        Add-UniqueNote $notes '`网址` 建议写完整的 `http://` 或 `https://` 地址。'
    }
    if ($syntax -match '程序名') {
        Add-UniqueNote $notes '`程序名` 一般填写进程名或可执行文件名。'
    }
    if ($syntax -match '标题') {
        Add-UniqueNote $notes '窗口命令中的 `标题` 使用标题文本匹配窗口。'
    }
    if ($syntax -match '键名') {
        Add-UniqueNote $notes '`键名` 可写为 `A`、`F5`、`Enter`、`Ctrl` 等常见键名。'
    }
    if ($syntax -match '列表') {
        Add-UniqueNote $notes '`列表` 推荐写成 `[甲,乙,丙]` 这样的数组文本。'
    }
    if ($syntax -match '运算符') {
        Add-UniqueNote $notes '`运算符` 支持 `=`、`!=`、`>`、`<`、`>=`、`<=` 等比较符。'
    }
    if ($syntax -match '模板') {
        Add-UniqueNote $notes '`模板` 用于描述时间格式，例如 `yyyy-MM-dd HH:mm:ss`。'
    }
    if ($syntax -match '颜色') {
        Add-UniqueNote $notes '`颜色` 推荐使用十六进制文本，例如 `#FF0000`。'
    }
    if ($syntax -match '提示|消息') {
        Add-UniqueNote $notes '提示或消息中包含空格时，推荐用引号包裹。'
    }
    if ($syntax -match '数据') {
        Add-UniqueNote $notes '`数据` 常用于 `POST` 请求体，可以直接写表单串或 JSON 文本。'
    }
    if ($syntax -match '配置') {
        Add-UniqueNote $notes '`配置` 预留给邮件相关设置，当前命令仍在规划中。'
    }
    if ($syntax -match '\[') {
        Add-UniqueNote $notes '方括号中的内容表示可选参数。'
    }
    if ($notes.Count -eq 0) {
        Add-UniqueNote $notes '按语法顺序依次填写参数即可。'
    }
    return $notes
}

function Get-BehaviorText {
    param($Item)

    switch ($Item.name) {
        '变量' { return '如果省略初始值，解释器会把该变量写成 `0`。' }
        '询问' { return '执行后会把用户输入写入变量。' }
        '确认' { return '执行后会把 `true` 或 `false` 写入变量。' }
        '确定弹窗' { return '按按钮时写入 `true`，关闭弹窗时写入 `false`。' }
        '询问弹窗' { return '按按钮时写入输入文本，关闭弹窗时写入 `false`。' }
        '选择弹窗' { return '点击按钮时写入按钮索引，关闭弹窗时写入 `false`。' }
        '找图' { return '找到时写入首个匹配点左上角坐标，未找到时写入 `false`。' }
        '识字' { return '执行后会把 OCR 识别到的文本写入变量；识别不到时写入空字符串。' }
        '返回' { return '在函数内部结束当前函数，并返回一个可选值。' }
        '跳出' { return '仅在循环块中有效，用来结束当前循环。' }
        '继续' { return '仅在循环块中有效，用来跳过本轮余下语句。' }
        default {
            if ($Item.syntax -match '到 变量名|到 名称') {
                return '执行结果会写入指定变量。'
            }
            if ($Item.name -eq '函数') {
                return '定义函数本身不会立刻执行，只有被调用时才会运行。'
            }
            return '这条命令主要执行动作本身，不强制产生返回值。'
        }
    }
}

function Build-GenericExample {
    param([string]$Syntax)

    $sample = $Syntax
    $replacementPairs = @(
        @('左键\|右键', '左键'),
        @('单点\|双击\|多击\|长按', '单点'),
        @('线性\|缓入\|缓出\|缓入缓出', '线性'),
        @('0\|1', '0'),
        @('GET\|POST', 'GET'),
        @('开始\|停止', '开始'),
        @('开\|关', '开'),
        @('上\|下', '上'),
        @('左值', '${值}'),
        @('右值', '1'),
        @('运算符', '='),
        @('列表', '[甲,乙,丙]'),
        @('变量名', '结果'),
        @('名称', '结果'),
        @('标题', '记事本'),
        @('内容', 'hello'),
        @('提示', '请输入内容'),
        @('消息', '操作完成'),
        @('选项', '[甲,乙,丙]'),
        @('模板', 'yyyy-MM-dd'),
        @('颜色', '#FF0000'),
        @('数据', 'name=demo'),
        @('配置', 'smtp.json'),
        @('程序名', 'notepad.exe'),
        @('网址', 'https://example.com'),
        @('程序', 'notepad.exe'),
        @('文件路径', 'sample.wav'),
        @('图片路径', 'sample.bmp'),
        @('压缩包', 'archive.zip'),
        @('目录', 'data'),
        @('路径', 'demo.txt'),
        @('旧路径', 'old.txt'),
        @('新路径', 'new.txt'),
        @('源', 'a.txt'),
        @('目标', 'b.txt'),
        @('命令文本', 'echo hello'),
        @('主机', '127.0.0.1'),
        @('端口', '80'),
        @('频率', '440'),
        @('毫秒', '200'),
        @('\[次数或秒数\]', '2'),
        @('次数', '3'),
        @('秒数', '1'),
        @('宽', '1280'),
        @('高', '720'),
        @('\(x1,y1\)->\(x2,y2\)', '(100,100)->(500,300)'),
        @('\(x,y\)', '(100,100)'),
        @('\[退出码\]', '0'),
        @('\[路径\]', 'capture.mp4'),
        @('\[数据\]', 'name=demo')
    )

    foreach ($pair in $replacementPairs) {
        $sample = $sample -replace $pair[0], $pair[1]
    }
    return $sample
}

function Get-MinimalExample {
    param($Item)

    switch ($Item.name) {
        '如果' {
            return [string]::Join("`r`n", @(
                '```os',
                '变量 值 = 1',
                '如果 ${值} = 1',
                '    输出 条件成立',
                '```'
            ))
        }
        '否则如果' {
            return [string]::Join("`r`n", @(
                '```os',
                '如果 ${值} = 1',
                '    输出 一',
                '否则如果 ${值} = 2',
                '    输出 二',
                '```'
            ))
        }
        '否则' {
            return [string]::Join("`r`n", @(
                '```os',
                '如果 ${值} = 1',
                '    输出 一',
                '否则',
                '    输出 其他',
                '```'
            ))
        }
        '当' {
            return [string]::Join("`r`n", @(
                '```os',
                '变量 次数 = 0',
                '当 ${次数} < 3',
                '    次数 = ${次数} 加 1',
                '    输出 ${次数}',
                '```'
            ))
        }
        '重复' {
            return [string]::Join("`r`n", @(
                '```os',
                '重复 3',
                '    输出 hello',
                '```'
            ))
        }
        '遍历' {
            return [string]::Join("`r`n", @(
                '```os',
                '遍历 项 在 [甲,乙,丙]',
                '    输出 ${项}',
                '```'
            ))
        }
        '函数' {
            return [string]::Join("`r`n", @(
                '```os',
                '函数 求和(a,b)',
                '    结果 = ${a} 加 ${b}',
                '    返回 ${结果}',
                '```'
            ))
        }
        '返回' {
            return [string]::Join("`r`n", @(
                '```os',
                '函数 求值()',
                '    返回 123',
                '```'
            ))
        }
        '调用' {
            return [string]::Join("`r`n", @(
                '```os',
                '调用 demo.os',
                '```'
            ))
        }
        '标签' {
            return [string]::Join("`r`n", @(
                '```os',
                '标签 开始',
                '输出 hello',
                '```'
            ))
        }
        '去' {
            return [string]::Join("`r`n", @(
                '```os',
                '去 结束点',
                '标签 结束点',
                '```'
            ))
        }
        '变量' {
            return [string]::Join("`r`n", @(
                '```os',
                '变量 计数',
                '变量 名字 = OperationScript',
                '```'
            ))
        }
        '询问' {
            return [string]::Join("`r`n", @(
                '```os',
                '询问 请输入名字 到 名字',
                '```'
            ))
        }
        '确认' {
            return [string]::Join("`r`n", @(
                '```os',
                '确认 是否继续 到 继续执行',
                '```'
            ))
        }
        '确定弹窗' {
            return [string]::Join("`r`n", @(
                '```os',
                '确定弹窗 (400,300) "保存成功" 是否确认[确定]',
                '```'
            ))
        }
        '询问弹窗' {
            return [string]::Join("`r`n", @(
                '```os',
                '询问弹窗 (400,300) "请输入名称" 输入结果[保存]',
                '```'
            ))
        }
        '选择弹窗' {
            return [string]::Join("`r`n", @(
                '```os',
                '选择弹窗 (400,300) "请选择模式" 选择结果[简单,标准,高级]',
                '```'
            ))
        }
        '请求' {
            return [string]::Join("`r`n", @(
                '```os',
                '请求 GET https://example.com 到 响应',
                '```'
            ))
        }
        '找图' {
            return [string]::Join("`r`n", @(
                '```os',
                '找图 button.bmp 到 按钮坐标',
                '```'
            ))
        }
        '识字' {
            return [string]::Join("`r`n", @(
                '```os',
                '识字 (100,100)->(400,180) 到 文本',
                '```'
            ))
        }
        '保存截图' {
            return [string]::Join("`r`n", @(
                '```os',
                '保存截图 screen.bmp',
                '```'
            ))
        }
        '找色' {
            return [string]::Join("`r`n", @(
                '```os',
                '找色 #FF0000 (0,0)->(800,600) 到 命中坐标',
                '```'
            ))
        }
        default {
            return [string]::Join("`r`n", @(
                '```os',
                (Build-GenericExample $Item.syntax),
                '```'
            ))
        }
    }
}

function Get-AdvancedExample {
    param($Item)

    switch ($Item.name) {
        '如果' {
            return [string]::Join("`r`n", @(
                '```os',
                '如果 ${分数} >= 90',
                '    输出 优秀',
                '否则如果 ${分数} >= 60',
                '    输出 及格',
                '否则',
                '    输出 继续努力',
                '```'
            ))
        }
        '当' {
            return [string]::Join("`r`n", @(
                '```os',
                '当 ${在线} = 真',
                '    检测网络 到 在线',
                '    等待 1',
                '```'
            ))
        }
        '遍历' {
            return [string]::Join("`r`n", @(
                '```os',
                '遍历 文件 在 [a.txt,b.txt,c.txt]',
                '    文件存在 ${文件} 到 存在',
                '    输出 ${文件}:${存在}',
                '```'
            ))
        }
        '函数' {
            return [string]::Join("`r`n", @(
                '```os',
                '函数 工具->打印标题(文本)',
                '    输出 =====',
                '    输出 ${文本}',
                '    输出 =====',
                '',
                '工具->打印标题(开始)',
                '```'
            ))
        }
        '请求' {
            return [string]::Join("`r`n", @(
                '```os',
                '请求 POST https://example.com/api {"name":"os"} 到 响应',
                '```'
            ))
        }
        default {
            return ''
        }
    }
}

$implementedCount = ($items | Where-Object implemented).Count
$lines = New-Object 'System.Collections.Generic.List[string]'

[void]$lines.Add('# OperationScript 命令手册')
[void]$lines.Add('')
[void]$lines.Add('这份手册以 `explain.cpp` 中的 `buildCommandCatalog()` 为唯一真实来源整理。')
[void]$lines.Add('')
[void]$lines.Add("当前命令总数：**$($items.Count)**。已实现：**$implementedCount**。未实现：**$($items.Count - $implementedCount)**。")
[void]$lines.Add('')
[void]$lines.Add('## 分类总览')
[void]$lines.Add('')
[void]$lines.Add('| 分类 | 数量 | 已实现 |')
[void]$lines.Add('| --- | ---: | ---: |')

foreach ($category in $categoryOrder) {
    $group = $items | Where-Object category -eq $category
    $groupImplemented = ($group | Where-Object implemented).Count
    [void]$lines.Add("| $category | $($group.Count) | $groupImplemented |")
}

[void]$lines.Add('')
[void]$lines.Add('## 阅读约定')
[void]$lines.Add('')
[void]$lines.Add('- `已实现` 表示当前解释器可以直接执行。')
[void]$lines.Add('- `未实现` 表示语法已保留在语言设计中，示例写法可以作为后续扩展参考。')
[void]$lines.Add('- 变量引用统一推荐使用 `${变量名}`。')
[void]$lines.Add('')

$index = 1
foreach ($category in $categoryOrder) {
    [void]$lines.Add("## $category")
    [void]$lines.Add('')

    foreach ($item in ($items | Where-Object category -eq $category)) {
        $status = if ($item.implemented) { '已实现' } else { '未实现' }
        [void]$lines.Add("### $index. $($item.name)")
        [void]$lines.Add('')
        [void]$lines.Add('- 分类：`' + $item.category + '`')
        [void]$lines.Add('- 实现状态：`' + $status + '`')
        [void]$lines.Add('- 语法：`' + $item.syntax + '`')
        [void]$lines.Add('- 说明：' + $item.desc)
        [void]$lines.Add('- 参数说明：')
        foreach ($note in (Get-ParameterNotes $item)) {
            [void]$lines.Add("  - $note")
        }
        [void]$lines.Add('- 返回/变量回写：' + (Get-BehaviorText $item))
        [void]$lines.Add('- 最小示例：')
        foreach ($line in (Get-MinimalExample $item) -split "`r?`n") {
            [void]$lines.Add($line)
        }

        $advanced = Get-AdvancedExample $item
        if (-not [string]::IsNullOrWhiteSpace($advanced)) {
            [void]$lines.Add('- 进阶示例：')
            foreach ($line in $advanced -split "`r?`n") {
                [void]$lines.Add($line)
            }
        }

        [void]$lines.Add('')
        $index++
    }
}

[System.IO.File]::WriteAllLines(
    (Join-Path (Get-Location) 'docs\commands.md'),
    $lines,
    [System.Text.UTF8Encoding]::new($false))

$stdlibLines = @(
    '# OperationScript 标准库',
    '',
    '标准库脚本位于 `stdlib/` 目录，当前仓库自带 3 个基础模块。',
    '',
    '## 使用方式',
    '',
    '```os',
    '调用 stdlib/console.os',
    '调用 stdlib/math.os',
    '调用 stdlib/files.os',
    '```',
    '',
    '## console.os',
    '',
    '- 文件：`stdlib/console.os`',
    '- 作用：封装控制台标题风格输出和暂停逻辑。',
    '',
    '### 控制台->标题行(文本)',
    '',
    '- 参数：`文本` 为要展示的标题内容。',
    '- 返回：无返回值，直接连续输出三行文本。',
    '- 示例：',
    '```os',
    '调用 stdlib/console.os',
    '控制台->标题行(任务开始)',
    '```',
    '',
    '### 控制台->暂停()',
    '',
    '- 参数：无。',
    '- 返回：无返回值，会提示用户按回车继续。',
    '- 示例：',
    '```os',
    '调用 stdlib/console.os',
    '控制台->暂停()',
    '```',
    '',
    '## math.os',
    '',
    '- 文件：`stdlib/math.os`',
    '- 作用：提供简单数值辅助函数。',
    '',
    '### 数学->自增(值)',
    '',
    '- 参数：`值` 为需要加一的数。',
    '- 返回：返回 `值 + 1`。',
    '- 示例：',
    '```os',
    '调用 stdlib/math.os',
    '结果 = 数学->自增(5)',
    '显示变量 结果',
    '```',
    '',
    '### 数学->相加(a,b)',
    '',
    '- 参数：`a` 和 `b` 为两个加数。',
    '- 返回：返回 `a + b`。',
    '- 示例：',
    '```os',
    '调用 stdlib/math.os',
    '和 = 数学->相加(3,4)',
    '显示变量 和',
    '```',
    '',
    '## files.os',
    '',
    '- 文件：`stdlib/files.os`',
    '- 作用：封装常见日志追加和目录初始化逻辑。',
    '',
    '### 文件->写日志(路径,内容)',
    '',
    '- 参数：`路径` 为日志文件路径，`内容` 为要写入的正文。',
    '- 返回：无返回值，会读取当前时间并把格式化后的日志追加到文件末尾。',
    '- 示例：',
    '```os',
    '调用 stdlib/files.os',
    '文件->写日志(app.log, 启动完成)',
    '```',
    '',
    '### 文件->重建目录(路径)',
    '',
    '- 参数：`路径` 为要重建的目录。',
    '- 返回：无返回值；如果目录已存在，会先删除再创建。',
    '- 示例：',
    '```os',
    '调用 stdlib/files.os',
    '文件->重建目录(temp)',
    '```'
)

[System.IO.File]::WriteAllLines(
    (Join-Path (Get-Location) 'docs\stdlib.md'),
    $stdlibLines,
    [System.Text.UTF8Encoding]::new($false))
