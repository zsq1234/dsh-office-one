# dsh-office-one

面向 DSH Web 的一体化 Office 插件，提供以下功能：

- 在对话中创建和编辑 Sheet、Doc、Slide。
- 导入和导出 Excel、Word、PowerPoint 文件。
- 在 DSH 的 Workspace 文件列表中直接打开、编辑和保存 Office 文件。
- 支持 AI 通过 Univer Facade API 操作文档。

详细的操作说明、格式支持和故障排查请参阅 [`USE.md`](./USE.md)。

## 功能模块

项目由三个模块组成：

- `univer-create`：会话内的 Sheet、Doc、Slide 编辑器及 AI 工具。
- `univer-file-export`：基于 Univer Exchange Node 的 Office 文件转换服务。
- `workspace-file-viewer`：Workspace Office 文件预览和编辑器。

插件以一个 DSH Bundle、一个 Host entry 和一个 Client bundle 安装。

## 支持的文件

| 类型 | 可打开 | 可保存 |
| --- | --- | --- |
| Sheet | `.xls`、`.xlsx`、`.csv` | `.xlsx`、`.csv` |
| Doc | `.doc`、`.docx` | `.docx` |
| Slide | `.ppt`、`.pptx` | `.pptx` |

旧版 `.xls`、`.doc` 和 `.ppt` 文件可以打开，但需保存为对应的新格式。

## 安装

```bash
dsh plugin --profile web add /path/to/dsh-office-one
```

DSH profile 中只会增加 `dsh-office-one` 一个依赖和一个 bundle layer。

## 构建

```bash
pnpm install
pnpm build
```

主要构建产物：

- `lib/index.js`：Host 插件。
- `lib/client.js`：DSH Web Client 插件。
- `lib/runtimes/*.js`：Workspace Office 预览运行时。
- `server/export-file.mjs`：Office 文件转换服务。

## 项目结构

```text
src/
├── index.ts
├── client.ts
└── modules/
    ├── univer-create/
    ├── univer-file-export/
    └── workspace-file-viewer/
server/
└── export-file.mjs
```

## 数据持久化

会话中的 Sheet、Doc、Slide 快照通过 DSH `storageDomain` 保存。编辑器会自动保存，但关闭页面或强制结束服务前的少量修改仍可能丢失；重要内容请及时导出为 Office 文件。

重启 DSH 时应保持相同的 `DSH_HOME` 和存储配置。

## Univer Pro License

构建时按以下顺序读取客户端 License：

1. 环境变量 `UNIVER_CLIENT_LICENSE`。
2. 项目根目录的 `license-univer/license.txt`。

`license-univer/` 已加入 `.gitignore`，请勿提交 License 文件。

## 环境变量

| 变量 | 说明 | 默认值 |
| --- | --- | --- |
| `UNIVER_CLIENT_LICENSE` | Univer 客户端 License | 读取本地 License 文件 |
| `PORT` | Office 转换服务端口 | `8787` |
| `CORS_ORIGIN` | 转换服务允许的浏览器 Origin | DSH Web 地址 |
| `UNIVER_FILE_IMPORT_ENDPOINT` | Workspace Viewer 导入端点 | 本地转换服务 |
| `UNIVER_FILE_EXPORT_ENDPOINT` | Office 文件导出端点 | 本地转换服务 |
