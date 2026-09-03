# dsh-office-one

把以下三个功能作为一个 DSH Bundle、一个 Host entry 和一个 Client bundle 安装：

- `univer-create`：对话中的 Sheet、Doc、Slide 编辑器和模型工具。
- `univer-file-export`：基于 Univer Exchange Node 的 Office 导入/导出服务。
- `workspace-file-viewer`：Workspace 文件树、代码预览和 Office 文件编辑。

业务代码仍按 `src/modules/` 下的三个目录独立维护；`src/index.ts` 和 `src/client.ts` 只是薄组合入口。

安装、界面操作、AI 使用示例、格式支持和故障排查请参阅 [`USE.md`](./USE.md)。

## 目录

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

## 构建

```bash
pnpm install
pnpm build
```

构建产物包括：

- `lib/index.js`：组合后的 Host 插件。
- `lib/client.js`：组合后的 DSH Web Client 插件。
- `lib/runtimes/*.js`：Workspace Office 预览按需加载的 Univer runtime。
- `server/export-file.mjs`：Office 文件转换子进程。

## 单包安装

```bash
dsh plugin --profile web add /home/azeng/WorkSpace/github.com/zsq1234/ai-test/dsh-office-one
```

DSH profile 中只会增加 `dsh-office-one` 一个依赖和一个 bundle layer。插件自身的 `cordis.patch.yml` 只插入一个 Loader entry，该 entry 在内部挂载三个模块。

迁移完成后，可移除旧的三个 profile 依赖：

```bash
dsh plugin --profile web remove \
  dsh-univer-create \
  dsh-univer-file-export \
  dsh-workspace-file-viewer
```

## 环境变量

- `PORT`：Office 转换服务端口，默认 `8787`。
- `CORS_ORIGIN`：转换服务允许的浏览器 Origin。
- `UNIVER_FILE_IMPORT_ENDPOINT`：Workspace Viewer 使用的导入端点。
- `UNIVER_FILE_EXPORT_ENDPOINT`：Univer Create 和 Workspace Viewer 使用的导出端点。
