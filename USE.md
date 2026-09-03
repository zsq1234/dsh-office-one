# dsh-office-one 使用指南

## 这是什么

`dsh-office-one` 是面向 DSH Web 的一体化 Office 插件，集中提供以下三组能力：

1. **对话内 Univer 编辑器**
   - 在当前会话中创建和编辑 Sheet、Doc、Slide。
   - 用户可以手动操作编辑器，也可以让 AI 通过插件工具读取或修改内容。
   - 编辑状态按会话保存；可将结果写入当前 session workspace。

2. **Office 文件导入与导出**
   - 使用 Univer Exchange Node 在 Office 文件和 Univer 数据之间转换。
   - 支持表格、文档和演示文稿的导入、编辑与导出。
   - 转换服务随插件自动启动，默认监听 `127.0.0.1:8787`。

3. **Workspace 文件查看与 Office 编辑**
   - 在 DSH 会话页面中浏览当前 session workspace 文件树。
   - 预览代码和文本文件。
   - 使用 Univer 打开和编辑 workspace 中的 Excel、Word、PowerPoint 文件。

插件包含一个 Host 插件、一个 Web Client 插件和一个 Bundle 配置层，在一个入口内组合对话编辑、文件转换与 workspace 文件查看能力。

## 适用场景

- 让 AI 创建数据表、报告或演示文稿。
- 在对话过程中手动调整 AI 生成的 Office 内容。
- 打开并修改项目 workspace 中已有的 Office 文件。
- 将对话中生成的内容保存为 `.xlsx`、`.docx` 或 `.pptx`。
- 在 DSH 页面内查看 workspace 的代码和文本文件。

## 使用入口

打开 DSH Web 页面并进入一个绑定了 workspace 的会话。会话视图中提供两个入口：

- **Univer**：创建和编辑 Sheet、Doc、Slide。
- **Workspace 文件**：浏览和编辑当前 session workspace 文件。

## 使用 Univer 编辑器

进入会话中的 **Univer** 视图后，可以在顶部切换：

- `Sheet`
- `Doc`
- `Slide`

每种产品都支持两种使用方式：

1. 用户在 Univer 界面中手动编辑。
2. 在对话中告诉 AI 要创建或修改什么，由 AI 调用插件工具操作当前文档。

### 使用 Sheet

可以点击：

- **新建**：创建空工作簿。
- **打开**：从本机选择 `.xlsx` 文件并导入。
- **保存**：保存为 `.xlsx` 到当前 session workspace。

第一次保存时需要填写 workspace 内的相对路径，例如：

```text
reports/sales-report.xlsx
```

目标目录必须已经存在，且路径不能离开当前 session workspace。

可以直接对 AI 说：

```text
创建一个销售统计表，包含产品、1 月销量、2 月销量和合计列，写入三条示例数据，并把表头设置为蓝色粗体。
```

```text
读取当前表格 Sheet1 的 A1:D20，计算合计并把结果写到 D 列。
```

```text
新建一个名为“汇总”的工作表，把当前数据整理成月度汇总。
```

Sheet 工具支持的主要能力包括：

- 新建或重置工作簿。
- 新增、删除、重命名和列出工作表。
- 按 A1 地址读取、写入或清空单元格区域。
- 设置背景色、字体颜色、粗体和字号。
- 写入数字、字符串、布尔值、空值和以 `=` 开头的公式。

### 使用 Doc

可以点击：

- **新建**：创建空文档。
- **打开**：从本机选择 `.docx` 文件并导入。
- **保存**：保存为 `.docx` 到当前 session workspace。

示例对话：

```text
创建一份项目周报，包含本周进展、风险、下周计划三个部分。
```

```text
读取当前文档内容，在末尾追加一个“待确认事项”章节。
```

```text
把当前文档前 20 个字符设置为粗体，字号改成 18。
```

Doc 工具支持的主要能力包括：

- 新建或重置文档。
- 读取和替换全部纯文本。
- 按零起始字符位置插入、追加或删除文本。
- 对指定字符范围设置粗体、斜体、字号和文字颜色。

字符范围采用零起始、左闭右开形式。例如 `start=0, end=5` 表示前 5 个字符。

### 使用 Slide

可以点击：

- **新建**：创建空演示文稿。
- **打开**：从本机选择 `.ppt` 或 `.pptx` 文件并导入。
- **保存**：保存为 `.pptx` 到当前 session workspace。

示例对话：

```text
创建一份 3 页的产品介绍演示文稿，分别是封面、核心功能和下一步计划。
```

```text
列出当前演示文稿的所有幻灯片和元素，然后修改第 2 页的标题。
```

```text
在第 1 页坐标 left=120、top=100 的位置添加标题“季度总结”，字号 36，加粗。
```

Slide 工具支持的主要能力包括：

- 新建或重置演示文稿。
- 列出幻灯片及其元素。
- 添加或删除幻灯片。
- 添加文本元素。
- 根据元素 ID 修改文本、位置、尺寸、字号、颜色和粗体。
- 根据元素 ID 删除元素。

更新或删除元素前，应先让 AI 列出当前幻灯片，取得正确的元素 ID。

## 使用 Workspace 文件视图

进入会话中的 **Workspace 文件** 视图后，左侧显示当前 session workspace 文件树，右侧显示选中文件的内容。

### 浏览文件

- 点击目录可展开或折叠。
- 每次默认加载 100 个目录项，可点击“继续加载”。
- 点击右上角的“刷新”可重新读取文件树。
- `.git`、`node_modules` 和 `.DS_Store` 默认不显示。
- 所有文件访问都限制在当前 session workspace 内，符号链接也不能绕过该限制。

### 查看代码和文本

点击普通代码或文本文件后，可以：

- 查看带语法高亮的内容。
- 复制全部文本。
- 使用 `A−` 和 `A+` 调整代码字号。
- 对大文件使用虚拟化渲染。

### 打开和编辑 Office 文件

Workspace 文件视图支持以下格式：

| 类型 | 可打开 | 可保存 |
| --- | --- | --- |
| Sheet | `.xls`、`.xlsx`、`.csv` | `.xlsx`、`.csv`；`.xls` 将另存为同名 `.xlsx` |
| Doc | `.doc`、`.docx` | `.docx` |
| Slide | `.ppt`、`.pptx` | `.pptx` |

打开 Office 文件后，可以直接在右侧 Univer 编辑器中修改。点击 **保存** 后：

- `.xlsx`、`.csv`、`.docx`、`.pptx` 原位保存。
- `.xls` 不覆盖原文件，而是在同目录生成同名 `.xlsx`。
- 旧版 `.doc` 和 `.ppt` 可以打开，但不能按旧格式原位保存。
- 如果文件在打开后被其他程序修改，插件会拒绝覆盖；刷新并重新打开文件后再保存。

单个读取、转换或保存请求的大小上限为 300 MiB。

## 数据保存说明

### 会话内编辑状态

Univer 视图中的 Sheet、Doc、Slide 快照由 Host 按 session 保存。切换会话标签或重新挂载编辑器时，插件会尝试恢复该会话之前的内容。

Sheet、Doc、Slide 使用独立的存储记录，不会互相覆盖。

### 保存到 workspace

点击 Univer 编辑器中的 **保存** 时，文件会写入当前 session workspace，而不是任意系统目录。

保存路径必须满足：

- 使用 workspace 内的相对路径。
- 后缀与目标格式一致：`.xlsx`、`.docx` 或 `.pptx`。
- 父目录已经存在。
- 不能使用 `..` 或符号链接跳出 workspace。

如果目标文件已经存在，界面会询问是否覆盖。

## 转换服务与环境变量

插件启动时会自动创建 Office 转换子进程，执行内置的：

```text
server/export-file.mjs
```

默认端点：

```text
http://127.0.0.1:8787/api/univer/import-file
http://127.0.0.1:8787/api/univer/export-file
```

支持以下环境变量：

| 环境变量 | 默认值 | 说明 |
| --- | --- | --- |
| `PORT` | `8787` | Office 转换服务监听端口 |
| `CORS_ORIGIN` | 请求 Origin，缺省时为 `http://127.0.0.1:3080` | 转换服务允许的浏览器 Origin |
| `UNIVER_FILE_IMPORT_ENDPOINT` | `http://127.0.0.1:8787/api/univer/import-file` | Host 侧 Workspace Viewer 使用的导入端点 |
| `UNIVER_FILE_EXPORT_ENDPOINT` | `http://127.0.0.1:8787/api/univer/export-file` | Host 侧 Univer Create 和 Workspace Viewer 使用的导出端点 |

例如将转换服务改为 `8790` 端口：

```bash
PORT=8790 \
UNIVER_FILE_IMPORT_ENDPOINT=http://127.0.0.1:8790/api/univer/import-file \
UNIVER_FILE_EXPORT_ENDPOINT=http://127.0.0.1:8790/api/univer/export-file \
dsh web
```

注意：当前 Web Client 的手动“打开/导出”请求默认使用 `http://127.0.0.1:8787`。如果 DSH Web 运行在远程主机、反向代理后面，或者需要修改浏览器侧转换地址，应同步调整客户端配置；只修改 Host 环境变量不能改变浏览器侧已经固定的地址。

## 常见问题

### Office 文件打开或保存时报“无法连接转换服务”

检查：

- `8787` 端口是否被其他程序占用。
- DSH 启动日志中转换子进程是否报错。
- `UNIVER_FILE_IMPORT_ENDPOINT` 和 `UNIVER_FILE_EXPORT_ENDPOINT` 是否与 `PORT` 一致。
- 浏览器是否能访问 `127.0.0.1:8787`。
- `CORS_ORIGIN` 是否与实际 DSH Web Origin 一致。

### 保存时报“目标目录不存在”

插件不会自动创建父目录。先在当前 session workspace 中创建目录，再保存文件。

例如保存到：

```text
reports/weekly.docx
```

需要先确保 `reports/` 已存在。

### 保存时报文件已经改变

Workspace 文件视图会记录文件打开时的修改时间。如果文件随后被其他程序修改，插件会阻止覆盖。点击刷新，重新打开文件，确认内容后再保存。

### `.xls`、`.doc` 或 `.ppt` 为什么不能原格式保存

这些旧版格式可以导入，但当前导出能力以 OOXML 格式为主：

- `.xls` 另存为 `.xlsx`。
- `.doc` 需要保存为 `.docx`。
- `.ppt` 需要保存为 `.pptx`。

