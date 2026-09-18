import type { DshLanguage } from '../../../dsh-language.js'

type UniverCreateMessages = {
  untitledFile: string
  saveDialogTitle: string
  saveDialogDescription: string
  filePath: string
  example: string
  cancel: string
  save: string
  saving: string
  open: string
  opening: string
  create: string
  close: string
  confirmCloseUnsaved: string
  workbookFiles: string
  switchingWorkbook: string
  chatActivity: string
  closeChatActivity: string
  openChatActivity: string
  hide: string
  show: string
  user: string
  newFileHint: string
  hostRecoveryHint: string
  startSheet: string
  startDoc: string
  startSlide: string
  selectEditor: string
  shellLabel: string
  defaultSheetTitle: string
  defaultDocTitle: string
  defaultSlideTitle: string
  sheetError: string
  docError: string
  slideError: string
  fileStatus: (path: string) => string
  syncedSheetOperations: (count: number) => string
  syncedDocOperations: (count: number) => string
  syncedSlideOperations: (count: number) => string
  savedTo: (path: string) => string
  confirmOverwrite: (path: string) => string
  confirmDiscardChanges: string
  confirmDiscardChangesForNew: string
}

const en: UniverCreateMessages = {
  untitledFile: 'Untitled',
  saveDialogTitle: 'Save to the current session workspace',
  saveDialogDescription: 'Enter a relative path in the workspace. Files are saved in the workspace root by default.',
  filePath: 'File path',
  example: 'Example',
  cancel: 'Cancel',
  save: 'Save',
  saving: 'Saving…',
  open: 'Open',
  opening: 'Opening…',
  create: 'New',
  close: 'Close',
  confirmCloseUnsaved: 'The current file has unsaved changes. Close it and discard those changes?',
  workbookFiles: 'Open files',
  switchingWorkbook: 'Switching…',
  chatActivity: 'Conversation activity',
  closeChatActivity: 'Close conversation activity',
  openChatActivity: 'Open conversation activity',
  hide: 'Hide',
  show: 'Show',
  user: 'You',
  newFileHint: 'New file. Choose a workspace path when saving for the first time.',
  hostRecoveryHint: '. Refresh the page after the Host service recovers.',
  startSheet: 'Start using Univer Sheet',
  startDoc: 'Start using Univer Doc',
  startSlide: 'Start using Univer Slide',
  selectEditor: 'Select a Univer editor',
  shellLabel: 'Univer Sheet, Doc, and Slide',
  defaultSheetTitle: 'Conversation Sheet',
  defaultDocTitle: 'Conversation Document',
  defaultSlideTitle: 'Conversation Presentation',
  sheetError: 'Sheet operation failed: ',
  docError: 'Document operation failed: ',
  slideError: 'Slide operation failed: ',
  fileStatus: (path) => `File: ${path}`,
  syncedSheetOperations: (count) => `Synced ${count} conversation operation${count === 1 ? '' : 's'}`,
  syncedDocOperations: (count) => `Synced ${count} document operation${count === 1 ? '' : 's'}`,
  syncedSlideOperations: (count) => `Synced ${count} slide operation${count === 1 ? '' : 's'}`,
  savedTo: (path) => `Saved to ${path}`,
  confirmOverwrite: (path) => `${path} already exists. Overwrite it?`,
  confirmDiscardChanges: 'The current file has unsaved changes. Open another file and discard these changes?',
  confirmDiscardChangesForNew: 'The current file has unsaved changes. Create a new file and discard these changes?',
}

const zh: UniverCreateMessages = {
  untitledFile: '未命名文件',
  saveDialogTitle: '保存到当前 session workspace',
  saveDialogDescription: '可填写 workspace 内的相对路径；默认保存在 workspace 根目录。',
  filePath: '文件路径',
  example: '例如',
  cancel: '取消',
  save: '保存',
  saving: '保存中…',
  open: '打开',
  opening: '打开中…',
  create: '新建',
  close: '关闭',
  confirmCloseUnsaved: '当前文件有尚未保存的修改，确定关闭并放弃这些修改吗？',
  workbookFiles: '已打开的文件',
  switchingWorkbook: '切换中…',
  chatActivity: '对话动态',
  closeChatActivity: '关闭对话动态',
  openChatActivity: '打开对话动态',
  hide: '隐藏',
  show: '显示',
  user: '你',
  newFileHint: '新文件，首次保存时可选择 workspace 内路径',
  hostRecoveryHint: '。请在 Host 服务恢复后刷新页面重试。',
  startSheet: '开始使用 Univer Sheet',
  startDoc: '开始使用 Univer Doc',
  startSlide: '开始使用 Univer Slide',
  selectEditor: '选择 Univer 编辑器',
  shellLabel: 'Univer Sheet、Doc 与 Slide',
  defaultSheetTitle: '对话表格',
  defaultDocTitle: '对话文档',
  defaultSlideTitle: '对话演示文稿',
  sheetError: '表格操作失败：',
  docError: '文档操作失败：',
  slideError: '幻灯片操作失败：',
  fileStatus: (path) => `文件：${path}`,
  syncedSheetOperations: (count) => `已同步 ${count} 个对话操作`,
  syncedDocOperations: (count) => `已同步 ${count} 个文档操作`,
  syncedSlideOperations: (count) => `已同步 ${count} 个幻灯片操作`,
  savedTo: (path) => `已保存到 ${path}`,
  confirmOverwrite: (path) => `${path} 已存在，是否覆盖？`,
  confirmDiscardChanges: '当前文件有尚未保存的修改。是否打开其他文件并放弃这些修改？',
  confirmDiscardChangesForNew: '当前文件有尚未保存的修改。是否新建文件并放弃这些修改？',
}

export const UNIVER_CREATE_LOCALES: Record<DshLanguage, UniverCreateMessages> = { en, zh }
