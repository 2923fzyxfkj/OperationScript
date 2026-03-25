# OperationScript

OperationScript，简称 `os`，是一门面向电脑自动化操作的全中文脚本语言。脚本后缀名为 `.os`，可以通过解释器直接运行，也可以通过构建器打包成独立 `exe`。

当前仓库包含：

- `explain.cpp`：`.os` 解释器
- `builder.cpp`：`.os` 到 `exe` 的打包器
- `stdlib/`：标准库脚本
- `docs/commands.md`：完整命令手册
- `docs/stdlib.md`：标准库说明

## 快速开始

### 1. 编译解释器

```bash
g++ explain.cpp -std=c++17 -lshell32 -luser32 -lgdi32 -lgdiplus -o explain.exe
```

### 2. 运行脚本

```bash
explain.exe demo.os
```

### 3. 编译打包器

```bash
g++ builder.cpp -std=c++17 -o os-build.exe
```

### 4. 生成独立程序

```bash
os-build.exe 编译 demo.os demo.exe
os-build.exe 编译 demo.os demo.exe app.ico 1.0.0.0
```

`builder.cpp` 在生成独立程序后，会自动把 `tools/ocr.ps1` 复制到输出目录下的 `tools/`，这样打包后的程序也能直接使用 `识字`。

## 示例

```os
变量 次数
重复 3
    次数 = ${次数} 加 1
    输出 第${次数}次

函数 工具->求和(a,b)
    结果 = ${a} 加 ${b}
    返回 ${结果}

总和 = 工具->求和(3,4)
显示变量 总和

找图 button.bmp 到 按钮位置
识字 (100,100)->(400,180) 到 文本
```

## 当前状态

命令目录以 `buildCommandCatalog()` 为唯一真实来源。当前统计如下：

- 命令总数：`137`
- 已实现：`93`
- 未实现：`44`

已实现的大类包括：

- 鼠标自动化：点击、滑动、拖到、滚轮、鼠标回中
- 键盘基础：按键、回车、退格、删除键、输入文本
- 系统与窗口：执行、打开、等待、暂停、退出脚本、蜂鸣、窗口激活/关闭/最小化/最大化/还原
- 文件：复制、移动、删除、创建、读取、写入、目录操作、文件大小、存在性判断
- 变量与表达式：变量、加减乘除、连接、环境变量、类型转换、赋值表达式
- 流程：如果/否则如果/否则、当、重复、遍历、跳出、继续、标签、去、函数、返回、调用
- 交互：输出、标题、询问、确认、三种弹窗、显示变量、帮助、列出命令
- 网络：下载、读取网页、请求 GET/POST、打开网页、检测网络
- 时间：当前时间、当前日期、格式时间、定时
- 图像：找图、识字、取色、比色、找色、保存截图
- 调试：日志、警告、错误、断点、版本

## 文档入口

- 完整命令手册：[`docs/commands.md`](docs/commands.md)
- 标准库说明：[`docs/stdlib.md`](docs/stdlib.md)

`docs/commands.md` 已覆盖当前目录中的全部 `137` 条命令。每条命令都包含：

- 语法
- 参数说明
- 返回/变量回写行为
- 最小示例
- 需要时补充进阶示例

## 标准库

当前内置 3 个标准库脚本：

- `stdlib/console.os`
- `stdlib/math.os`
- `stdlib/files.os`

使用方式：

```os
调用 stdlib/console.os
调用 stdlib/math.os
调用 stdlib/files.os
```

详情见 [`docs/stdlib.md`](docs/stdlib.md)。

## 已完成的这一轮

这一轮已经完成：

- 把 `README.md` 收成首页入口
- 新增 `docs/commands.md`
- 新增 `docs/stdlib.md`
- `找图`：支持加载 `bmp/png/jpg/jpeg`，在整屏中按像素完全匹配查找首个命中位置
- `识字`：支持 `识字 (x1,y1)->(x2,y2) 到 变量名`
- 新增 `tools/ocr.ps1`，使用 Windows Runtime OCR
- `builder.cpp` 新增 `-lgdiplus` 链接，并自动复制 OCR 运行脚本

## 现在还可以继续做

下一轮更值得继续补的方向：

1. 网络增强：`上传`、`本机IP`、`公网IP`、`端口检测`、`发送邮件`
2. 输入增强：`长按键`、`松开键`、`组合键`、`复制`、`粘贴`、`撤销`、`重做`
3. 窗口增强：`移动窗口`、`调整窗口`、`置顶窗口`、`读取窗口文本`
4. 系统增强：`通知`、`截图`、`录屏`、`设置音量`、`亮度`
5. builder 增强：发布模式、多脚本打包、资源压缩、运行时依赖检查

## 说明

- 当前实现主要面向 Windows。
- `识字` 依赖 Windows 自带 OCR 能力和 `tools/ocr.ps1`。
- 语言里仍保留了一部分“已设计但未实现”的命令，方便继续扩展。
