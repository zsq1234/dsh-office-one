import React, { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { ClientContext, ConversationNode } from '@deepseek-ai/dsh-client-runtime/client'
import type { ConnectionHandle } from '@deepseek-ai/dsh-client-connection/client'
import type { ConvViewProps } from '@deepseek-ai/dsh-client-ui-conversation/client'
import { getDocsEmptySnapshot, LocaleType, mergeLocales, Univer } from '@univerjs/core'
import { FUniver } from '@univerjs/core/facade'
import { createUniver, defaultTheme } from '@univerjs/presets'
import { UniverDocsCorePreset } from '@univerjs/preset-docs-core'
import UniverPresetDocsCoreZhCN from '@univerjs/preset-docs-core/locales/zh-CN'
import { UniverDocsDrawingPreset } from '@univerjs/preset-docs-drawing'
import UniverPresetDocsDrawingZhCN from '@univerjs/preset-docs-drawing/locales/zh-CN'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import { UniverDocsPlugin } from '@univerjs/docs'
import { UniverDocsUIPlugin } from '@univerjs/docs-ui'
import DocsUIZhCN from '@univerjs/docs-ui/locale/zh-CN'
import { UniverDrawingPlugin } from '@univerjs/drawing'
import { UniverRenderEnginePlugin } from '@univerjs/engine-render'
import { UniverUIPlugin } from '@univerjs/ui'
import UIZhCN from '@univerjs/ui/locale/zh-CN'
import DesignZhCN from '@univerjs/design/locale/zh-CN'
import { UniverLicensePlugin } from '@univerjs-pro/license'
import ShapeEditorUIZhCN from '@univerjs-pro/shape-editor-ui/locale/zh-CN'
import { getSlidesEmptySnapshot, PageElementTypeEnum, UniverSlidesPlugin } from '@univerjs-pro/slides'
import type { ISlideData, ISlideTextElement } from '@univerjs-pro/slides'
import type { FPresentation } from '@univerjs-pro/slides/facade'
import '@univerjs-pro/slides/facade'
import { UniverSlidesUIPlugin } from '@univerjs-pro/slides-ui'
import SlidesUIZhCN from '@univerjs-pro/slides-ui/locale/zh-CN'

import '@univerjs/preset-docs-core/lib/index.css'
import '@univerjs/preset-docs-drawing/lib/index.css'
import '@univerjs/preset-sheets-core/lib/index.css'
import '@univerjs/design/lib/index.css'
import '@univerjs/ui/lib/index.css'
import '@univerjs/docs-ui/lib/index.css'
import '@univerjs-pro/shape-editor-ui/lib/index.css'
import '@univerjs-pro/slides-ui/lib/index.css'
import './styles.css'

type JsonScalar = string | number | boolean | null

type SheetOperation =
  | { seq: number; action: 'new'; title?: string; sheetName?: string; rows?: number; columns?: number }
  | { seq: number; action: 'add-sheet'; name: string; rows?: number; columns?: number }
  | { seq: number; action: 'delete-sheet'; name: string }
  | { seq: number; action: 'rename-sheet'; oldName: string; newName: string }
  | { seq: number; action: 'list-sheets' }
  | { seq: number; action: 'set-range'; range: string; values: JsonScalar[][]; sheetName?: string }
  | { seq: number; action: 'clear-range'; range: string; sheetName?: string }
  | {
      seq: number
      action: 'format-range'
      range: string
      sheetName?: string
      background?: string
      fontColor?: string
      bold?: boolean
      fontSize?: number
    }

type DocOperation =
  | { seq: number; action: 'new-doc'; title?: string; text?: string }
  | { seq: number; action: 'get-doc-text' }
  | { seq: number; action: 'set-doc-text'; text: string }
  | { seq: number; action: 'insert-doc-text'; index: number; text: string }
  | { seq: number; action: 'append-doc-text'; text: string }
  | { seq: number; action: 'delete-doc-range'; start: number; end: number }
  | {
      seq: number
      action: 'format-doc-text'
      start: number
      end: number
      bold?: boolean
      italic?: boolean
      fontSize?: number
      fontColor?: string
    }

type SlideOperation =
  | { seq: number; action: 'new-slide'; title?: string; width?: number; height?: number }
  | { seq: number; action: 'list-slides' }
  | { seq: number; action: 'add-slide'; title?: string; index?: number }
  | { seq: number; action: 'delete-slide'; index: number }
  | {
      seq: number
      action: 'add-slide-text'
      slideIndex: number
      text: string
      left?: number
      top?: number
      width?: number
      height?: number
      fontSize?: number
      fontColor?: string
      bold?: boolean
    }
  | {
      seq: number
      action: 'update-slide-text'
      slideIndex: number
      elementId: string
      text?: string
      left?: number
      top?: number
      width?: number
      height?: number
      fontSize?: number
      fontColor?: string
      bold?: boolean
    }
  | { seq: number; action: 'delete-slide-element'; slideIndex: number; elementId: string }

type UniverUnitType = 'sheet' | 'doc' | 'slide'
type UniverRuntime = ReturnType<typeof createUniver>
type MountedRuntime = UniverRuntime & { mount: HTMLDivElement }
type MountedSlideRuntime = { univer: Univer; univerAPI: FUniver; mount: HTMLDivElement; presentation: FPresentation }

const UNIVER_EXPORT_ENDPOINT = 'http://127.0.0.1:8787/api/univer/export-file'

const SHEET_TOOL_NAMES = new Set([
  'univer_sheet_new',
  'univer_sheet_add',
  'univer_sheet_delete',
  'univer_sheet_rename',
  'univer_sheet_list',
  'univer_sheet_set_range',
  'univer_sheet_clear_range',
  'univer_sheet_format_range',
])

function operationFromNode(node: ConversationNode): SheetOperation | null {
  if (node.kind !== 'tool-result' || node.isError || node.call === null || !SHEET_TOOL_NAMES.has(node.call.name)) return null

  let args: Record<string, unknown>
  try {
    const parsed = JSON.parse(node.call.argsRaw)
    if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) return null
    args = parsed as Record<string, unknown>
  } catch {
    return null
  }

  const sheetName = typeof args.sheetName === 'string' ? args.sheetName : undefined
  if (node.call.name === 'univer_sheet_new') {
    return {
      seq: node.seq,
      action: 'new',
      title: typeof args.title === 'string' ? args.title : undefined,
      sheetName,
      rows: typeof args.rows === 'number' ? args.rows : undefined,
      columns: typeof args.columns === 'number' ? args.columns : undefined,
    }
  }

  if (node.call.name === 'univer_sheet_add') {
    if (typeof args.name !== 'string') return null
    return {
      seq: node.seq,
      action: 'add-sheet',
      name: args.name,
      rows: typeof args.rows === 'number' ? args.rows : undefined,
      columns: typeof args.columns === 'number' ? args.columns : undefined,
    }
  }
  if (node.call.name === 'univer_sheet_delete') {
    return typeof args.name === 'string' ? { seq: node.seq, action: 'delete-sheet', name: args.name } : null
  }
  if (node.call.name === 'univer_sheet_rename') {
    return typeof args.oldName === 'string' && typeof args.newName === 'string'
      ? { seq: node.seq, action: 'rename-sheet', oldName: args.oldName, newName: args.newName }
      : null
  }
  if (node.call.name === 'univer_sheet_list') {
    return { seq: node.seq, action: 'list-sheets' }
  }
  if (typeof args.range !== 'string') return null
  if (node.call.name === 'univer_sheet_set_range') {
    if (!Array.isArray(args.values)) return null
    return { seq: node.seq, action: 'set-range', range: args.range, values: args.values as JsonScalar[][], sheetName }
  }
  if (node.call.name === 'univer_sheet_clear_range') {
    return { seq: node.seq, action: 'clear-range', range: args.range, sheetName }
  }
  return {
    seq: node.seq,
    action: 'format-range',
    range: args.range,
    sheetName,
    background: typeof args.background === 'string' ? args.background : undefined,
    fontColor: typeof args.fontColor === 'string' ? args.fontColor : undefined,
    bold: typeof args.bold === 'boolean' ? args.bold : undefined,
    fontSize: typeof args.fontSize === 'number' ? args.fontSize : undefined,
  }
}

const DOC_TOOL_NAMES = new Set([
  'univer_doc_new',
  'univer_doc_get_text',
  'univer_doc_set_text',
  'univer_doc_insert_text',
  'univer_doc_append_text',
  'univer_doc_delete_range',
  'univer_doc_format_text',
])

function docOperationFromNode(node: ConversationNode): DocOperation | null {
  if (node.kind !== 'tool-result' || node.isError || node.call === null || !DOC_TOOL_NAMES.has(node.call.name)) return null

  let args: Record<string, unknown>
  try {
    const parsed = JSON.parse(node.call.argsRaw)
    if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) return null
    args = parsed as Record<string, unknown>
  } catch {
    return null
  }

  if (node.call.name === 'univer_doc_new') {
    return {
      seq: node.seq,
      action: 'new-doc',
      title: typeof args.title === 'string' ? args.title : undefined,
      text: typeof args.text === 'string' ? args.text : undefined,
    }
  }
  if (node.call.name === 'univer_doc_get_text') return { seq: node.seq, action: 'get-doc-text' }
  if (node.call.name === 'univer_doc_set_text') {
    return typeof args.text === 'string' ? { seq: node.seq, action: 'set-doc-text', text: args.text } : null
  }
  if (node.call.name === 'univer_doc_insert_text') {
    return typeof args.index === 'number' && typeof args.text === 'string'
      ? { seq: node.seq, action: 'insert-doc-text', index: args.index, text: args.text }
      : null
  }
  if (node.call.name === 'univer_doc_append_text') {
    return typeof args.text === 'string' ? { seq: node.seq, action: 'append-doc-text', text: args.text } : null
  }
  if (node.call.name === 'univer_doc_delete_range') {
    return typeof args.start === 'number' && typeof args.end === 'number'
      ? { seq: node.seq, action: 'delete-doc-range', start: args.start, end: args.end }
      : null
  }
  if (typeof args.start !== 'number' || typeof args.end !== 'number') return null
  return {
    seq: node.seq,
    action: 'format-doc-text',
    start: args.start,
    end: args.end,
    bold: typeof args.bold === 'boolean' ? args.bold : undefined,
    italic: typeof args.italic === 'boolean' ? args.italic : undefined,
    fontSize: typeof args.fontSize === 'number' ? args.fontSize : undefined,
    fontColor: typeof args.fontColor === 'string' ? args.fontColor : undefined,
  }
}

const SLIDE_TOOL_NAMES = new Set([
  'univer_slide_new',
  'univer_slide_list',
  'univer_slide_add',
  'univer_slide_delete',
  'univer_slide_add_text',
  'univer_slide_update_text',
  'univer_slide_delete_element',
])

function slideOperationFromNode(node: ConversationNode): SlideOperation | null {
  if (node.kind !== 'tool-result' || node.isError || node.call === null || !SLIDE_TOOL_NAMES.has(node.call.name)) return null

  let args: Record<string, unknown>
  try {
    const parsed = JSON.parse(node.call.argsRaw)
    if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) return null
    args = parsed as Record<string, unknown>
  } catch {
    return null
  }

  const numberArg = (key: string) => typeof args[key] === 'number' ? args[key] : undefined
  if (node.call.name === 'univer_slide_new') {
    return {
      seq: node.seq,
      action: 'new-slide',
      title: typeof args.title === 'string' ? args.title : undefined,
      width: numberArg('width'),
      height: numberArg('height'),
    }
  }
  if (node.call.name === 'univer_slide_list') return { seq: node.seq, action: 'list-slides' }
  if (node.call.name === 'univer_slide_add') {
    return {
      seq: node.seq,
      action: 'add-slide',
      title: typeof args.title === 'string' ? args.title : undefined,
      index: numberArg('index'),
    }
  }
  if (node.call.name === 'univer_slide_delete') {
    return typeof args.index === 'number' ? { seq: node.seq, action: 'delete-slide', index: args.index } : null
  }
  if (typeof args.slideIndex !== 'number') return null
  if (node.call.name === 'univer_slide_delete_element') {
    return typeof args.elementId === 'string'
      ? { seq: node.seq, action: 'delete-slide-element', slideIndex: args.slideIndex, elementId: args.elementId }
      : null
  }
  if (node.call.name === 'univer_slide_add_text') {
    if (typeof args.text !== 'string') return null
    return {
      seq: node.seq,
      action: 'add-slide-text',
      slideIndex: args.slideIndex,
      text: args.text,
      left: numberArg('left'),
      top: numberArg('top'),
      width: numberArg('width'),
      height: numberArg('height'),
      fontSize: numberArg('fontSize'),
      fontColor: typeof args.fontColor === 'string' ? args.fontColor : undefined,
      bold: typeof args.bold === 'boolean' ? args.bold : undefined,
    }
  }
  if (typeof args.elementId !== 'string') return null
  return {
    seq: node.seq,
    action: 'update-slide-text',
    slideIndex: args.slideIndex,
    elementId: args.elementId,
    text: typeof args.text === 'string' ? args.text : undefined,
    left: numberArg('left'),
    top: numberArg('top'),
    width: numberArg('width'),
    height: numberArg('height'),
    fontSize: numberArg('fontSize'),
    fontColor: typeof args.fontColor === 'string' ? args.fontColor : undefined,
    bold: typeof args.bold === 'boolean' ? args.bold : undefined,
  }
}

function normalizeWorkbookSnapshot(value: unknown): any {
  if (value === null || typeof value !== 'object') return value
  const snapshot = JSON.parse(JSON.stringify(value)) as any
  if (snapshot.sheets !== null && typeof snapshot.sheets === 'object') {
    for (const sheet of Object.values(snapshot.sheets) as any[]) {
      if (sheet === null || typeof sheet !== 'object') continue
      if (!Number.isFinite(sheet.rowCount) || sheet.rowCount < 1) sheet.rowCount = 100
      if (!Number.isFinite(sheet.columnCount) || sheet.columnCount < 1) sheet.columnCount = 26
      if (!Number.isFinite(sheet.defaultColumnWidth) || sheet.defaultColumnWidth <= 0) sheet.defaultColumnWidth = 100
      if (sheet.columnHeader !== null && typeof sheet.columnHeader === 'object' && sheet.columnHeader.width <= 0) sheet.columnHeader.width = 100
      if (sheet.columnData !== null && typeof sheet.columnData === 'object') {
        for (const column of Object.values(sheet.columnData) as any[]) {
          if (column !== null && typeof column === 'object' && column.w !== undefined && (!Number.isFinite(column.w) || column.w <= 0)) delete column.w
        }
      }
    }
  }
  return snapshot
}

interface SavePathDialogProps {
  extension: 'xlsx' | 'docx' | 'pptx'
  suggestedName: string
  busy: boolean
  onCancel: () => void
  onSave: (path: string) => void
}

function safeOfficeFilename(title: string, extension: SavePathDialogProps['extension']): string {
  const stem = title.trim().replace(/[\\/:*?"<>|]+/g, '_') || '未命名文件'
  return `${stem}.${extension}`
}

function openedFilePath(file: File): string {
  const nativePath = (file as File & { path?: string }).path
  return typeof nativePath === 'string' && nativePath.length > 0
    ? nativePath
    : file.webkitRelativePath || file.name
}

function SavePathDialog({ extension, suggestedName, busy, onCancel, onSave }: SavePathDialogProps) {
  const [path, setPath] = useState(() => safeOfficeFilename(suggestedName, extension))
  return (
    <div className="dsh-univer-create-save-backdrop" role="presentation" onMouseDown={(event) => {
      if (event.target === event.currentTarget && !busy) onCancel()
    }}>
      <form className="dsh-univer-create-save-dialog" role="dialog" aria-modal="true" aria-labelledby="dsh-univer-save-title" onSubmit={(event) => {
        event.preventDefault()
        const value = path.trim()
        if (value.length > 0) onSave(value)
      }}>
        <h2 id="dsh-univer-save-title">保存到当前 session workspace</h2>
        <p>可填写 workspace 内的相对路径；默认保存在 workspace 根目录。</p>
        <label>
          文件路径
          <input autoFocus value={path} disabled={busy} onChange={(event) => setPath(event.target.value)} placeholder={`例如：${safeOfficeFilename(suggestedName, extension)}`} />
        </label>
        <div className="dsh-univer-create-save-actions">
          <button className="dsh-univer-create-action" type="button" disabled={busy} onClick={onCancel}>取消</button>
          <button className="dsh-univer-create-action dsh-univer-create-action--primary" type="submit" disabled={busy || path.trim().length === 0}>
            {busy ? '保存中…' : '保存'}
          </button>
        </div>
      </form>
    </div>
  )
}

async function loadUnitFilePath(sessionId: string, unitType: UniverUnitType): Promise<string | null> {
  const result = await clientConnection.current?.rpc.call('/dsh-univer-create', 'file-path', { sessionId, unitType })
  return result?.ok === true && typeof result.value === 'string' ? result.value : null
}

async function saveUnitFile(
  sessionId: string,
  unitType: UniverUnitType,
  snapshot: unknown,
  filePath: string,
  overwrite: boolean,
): Promise<string> {
  if (clientConnection.current === null) throw new Error('Host 连接不可用')
  const result = await clientConnection.current.rpc.call('/dsh-univer-create', 'export', {
    sessionId,
    unitType,
    snapshot,
    filePath,
    overwrite,
  })
  if (!result.ok) throw new Error(result.error.message)
  const value = result.value as { path?: unknown }
  if (typeof value.path !== 'string') throw new Error('Host 未返回保存路径')
  return value.path
}

async function exportWorkbookToXlsx(snapshot: unknown, workbookTitle: string): Promise<void> {
  const response = await fetch(UNIVER_EXPORT_ENDPOINT, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ unitType: 'sheet', format: 'xlsx', data: snapshot }),
  })
  if (!response.ok) {
    let message = `导出服务返回 HTTP ${response.status}`
    try {
      const payload = await response.json() as { error?: unknown }
      if (typeof payload.error === 'string' && payload.error.length > 0) message = payload.error
    } catch {
      // Keep the HTTP error if the export plugin did not return JSON.
    }
    throw new Error(message)
  }

  const blob = await response.blob()
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = `${workbookTitle.trim().replace(/[\\\\/:*?\"<>|]+/g, '_') || '对话表格'}.xlsx`
  document.body.appendChild(anchor)
  anchor.click()
  anchor.remove()
  URL.revokeObjectURL(url)
}

async function importWorkbookFromFile(file: File): Promise<any> {
  const params = new URLSearchParams({
    unitType: 'sheet',
    format: 'xlsx',
    fileName: file.name,
  })
  const response = await fetch(`${UNIVER_EXPORT_ENDPOINT.replace('/export-file', '/import-file')}?${params.toString()}`, {
    method: 'POST',
    headers: { 'content-type': 'application/octet-stream' },
    body: file,
  })
  const payload = await response.json() as { data?: unknown; error?: unknown }
  if (!response.ok) {
    throw new Error(typeof payload.error === 'string' ? payload.error : `导入服务返回 HTTP ${response.status}`)
  }
  if (payload.data === null || typeof payload.data !== 'object') throw new Error('导入服务没有返回有效的 workbook')
  return normalizeWorkbookSnapshot(payload.data)
}

async function exportDocumentToDocx(snapshot: unknown, documentTitle: string): Promise<void> {
  const response = await fetch(UNIVER_EXPORT_ENDPOINT, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ unitType: 'doc', format: 'docx', data: snapshot }),
  })
  if (!response.ok) {
    let message = `导出服务返回 HTTP ${response.status}`
    try {
      const payload = await response.json() as { error?: unknown }
      if (typeof payload.error === 'string' && payload.error.length > 0) message = payload.error
    } catch {
      // Keep the HTTP error if the export plugin did not return JSON.
    }
    throw new Error(message)
  }
  const blob = await response.blob()
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = `${documentTitle.trim().replace(/[\\\\/:*?\"<>|]+/g, '_') || '对话文档'}.docx`
  document.body.appendChild(anchor)
  anchor.click()
  anchor.remove()
  URL.revokeObjectURL(url)
}

async function importDocumentFromFile(file: File): Promise<any> {
  const params = new URLSearchParams({
    unitType: 'doc',
    format: 'docx',
    fileName: file.name,
  })
  const response = await fetch(`${UNIVER_EXPORT_ENDPOINT.replace('/export-file', '/import-file')}?${params.toString()}`, {
    method: 'POST',
    headers: { 'content-type': 'application/octet-stream' },
    body: file,
  })
  const payload = await response.json() as { data?: unknown; error?: unknown }
  if (!response.ok) {
    throw new Error(typeof payload.error === 'string' ? payload.error : `导入服务返回 HTTP ${response.status}`)
  }
  if (payload.data === null || typeof payload.data !== 'object') throw new Error('导入服务没有返回有效的 document')
  return payload.data
}

async function exportPresentationToPptx(snapshot: unknown, presentationTitle: string): Promise<void> {
  const response = await fetch(UNIVER_EXPORT_ENDPOINT, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ unitType: 'slide', format: 'pptx', data: snapshot }),
  })
  if (!response.ok) {
    let message = `导出服务返回 HTTP ${response.status}`
    try {
      const payload = await response.json() as { error?: unknown }
      if (typeof payload.error === 'string' && payload.error.length > 0) message = payload.error
    } catch {
      // Keep the HTTP error if the export plugin did not return JSON.
    }
    throw new Error(message)
  }
  const blob = await response.blob()
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = `${presentationTitle.trim().replace(/[\\/:*?\"<>|]+/g, '_') || '对话演示文稿'}.pptx`
  document.body.appendChild(anchor)
  anchor.click()
  anchor.remove()
  URL.revokeObjectURL(url)
}

async function importPresentationFromFile(file: File): Promise<ISlideData> {
  const params = new URLSearchParams({
    unitType: 'slide',
    format: file.name.toLowerCase().endsWith('.ppt') ? 'ppt' : 'pptx',
    fileName: file.name,
  })
  const response = await fetch(`${UNIVER_EXPORT_ENDPOINT.replace('/export-file', '/import-file')}?${params.toString()}`, {
    method: 'POST',
    headers: { 'content-type': 'application/octet-stream' },
    body: file,
  })
  const payload = await response.json() as { data?: unknown; error?: unknown }
  if (!response.ok) throw new Error(typeof payload.error === 'string' ? payload.error : `导入服务返回 HTTP ${response.status}`)
  const candidate = payload.data !== null && typeof payload.data === 'object' && 'slide' in payload.data
    ? (payload.data as { slide?: unknown }).slide
    : payload.data
  if (candidate === null || typeof candidate !== 'object') throw new Error('导入服务没有返回有效的 presentation')
  const slide = candidate as Partial<ISlideData>
  if (slide.slides === undefined || !Array.isArray(slide.slideOrder)) {
    throw new Error('导入服务返回的 Slide snapshot 与当前 Univer Pro Slide 编辑器不兼容')
  }
  return slide as ISlideData
}

function workbookData(operation?: Extract<SheetOperation, { action: 'new' }>) {
  const sheetId = 'sheet-1'
  return {
    id: `dsh-univer-${operation?.seq ?? 'default'}`,
    name: operation?.title ?? '对话表格',
    sheetOrder: [sheetId],
    sheets: {
      [sheetId]: {
        id: sheetId,
        name: operation?.sheetName ?? 'Sheet1',
        rowCount: Math.max(1, operation?.rows ?? 100),
        columnCount: Math.max(1, operation?.columns ?? 26),
      },
    },
  }
}

function documentData(operation?: Extract<DocOperation, { action: 'new-doc' }>) {
  return getDocsEmptySnapshot(
    `dsh-univer-doc-${operation?.seq ?? 'default'}`,
    LocaleType.ZH_CN,
    operation?.title ?? '对话文档',
  )
}

function presentationData(operation?: Extract<SlideOperation, { action: 'new-slide' }>): ISlideData {
  const data = getSlidesEmptySnapshot(
    `dsh-univer-slide-${operation?.seq ?? 'default'}`,
    LocaleType.ZH_CN,
    operation?.title ?? '对话演示文稿',
  )
  data.defaultPageSize = {
    ...data.defaultPageSize,
    width: Math.max(1, operation?.width ?? 960),
    height: Math.max(1, operation?.height ?? 540),
  }
  return data
}

function documentContentLength(document: { getBody(): { dataStream: string } }): number {
  const stream = document.getBody().dataStream
  if (stream.endsWith('\r\n\0')) return Math.max(0, stream.length - 3)
  return Math.max(0, stream.endsWith('\r\n') ? stream.length - 2 : stream.length)
}

function textFromBlocks(blocks: readonly unknown[]): string {
  return blocks.map((block) => {
    if (typeof block === 'string') return block
    if (block === null || typeof block !== 'object') return ''
    const value = block as Record<string, unknown>
    if (typeof value.text === 'string') return value.text
    if (typeof value.content === 'string') return value.content
    return ''
  }).filter(Boolean).join(' ')
}

function conversationLabel(node: ConversationNode): { role: string; text: string } | null {
  if (node.kind === 'user' || node.kind === 'steering') {
    const text = textFromBlocks(node.content)
    return text ? { role: '你', text } : null
  }
  if (node.kind === 'assistant') {
    const text = textFromBlocks(node.blocks)
    return text ? { role: 'AI', text } : null
  }
  return null
}

type SnapshotHook = <Selected>(selector: (value: any) => Selected) => Selected
type ConversationFeedSource = {
  nodes?: readonly ConversationNode[]
  partial?: { blocks?: readonly unknown[] } | null
  views?: { get: (target: string) => unknown }
}
type ChatSnapshotSource = {
  legacy?: {
    nodes?: readonly ConversationNode[]
    partial?: { blocks?: readonly unknown[] } | null
  }
}

const EMPTY_CONVERSATION_NODES: readonly ConversationNode[] = []

/**
 * Read Chat's live conversation projection on current DSH releases. Before the
 * target-neutral Conversation split, the same fields lived on useSession, so
 * keep that source as a compatibility fallback for older hosts.
 */
function useConversationFeed(props: ConvViewProps): { nodes: readonly ConversationNode[]; partialText: string } {
  const useConversation = (props as ConvViewProps & { useConversation?: SnapshotHook }).useConversation
  const useSource = useConversation ?? props.useSession as unknown as SnapshotHook
  const source = useSource((value) => value) as ConversationFeedSource | null | undefined
  const chat = source?.views?.get('chat') as ChatSnapshotSource | undefined
  const legacy = chat?.legacy ?? source
  const nodes = legacy?.nodes ?? EMPTY_CONVERSATION_NODES
  const partialText = legacy?.partial?.blocks == null ? '' : textFromBlocks(legacy.partial.blocks)
  return { nodes, partialText }
}

const clientConnection: { current: ConnectionHandle | null } = { current: null }
// Remember which sessions opened a workbook during this browser lifetime. This
// distinguishes a first visit (welcome screen) from remounting after a tab switch.
const openedWorkbookSessions = new Set<string>()

function SheetProductView(props: ConvViewProps) {
  const { sessionId } = props
  const { nodes, partialText } = useConversationFeed(props)
  const operations = useMemo(() => nodes.map(operationFromNode).filter((item): item is SheetOperation => item !== null), [nodes])
  const chatLines = useMemo(() => nodes
    .map(conversationLabel)
    .filter((item): item is { role: string; text: string } => item !== null)
    .slice(-4), [nodes])
  const containerRef = useRef<HTMLDivElement>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)
  const runtimeRef = useRef<MountedRuntime | null>(null)
  const appliedRef = useRef(new Set<number>())
  const chatStreamRef = useRef<HTMLDivElement>(null)
  const [title, setTitle] = useState('对话表格')
  const [error, setError] = useState<string | null>(null)
  const [importing, setImporting] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [filePath, setFilePath] = useState<string | null>(null)
  const [saveDialogOpen, setSaveDialogOpen] = useState(false)
  const [saveNotice, setSaveNotice] = useState('')
  const [sheetVisible, setSheetVisible] = useState(() => openedWorkbookSessions.has(sessionId))
  // Always let the session initialization effect choose first-visit vs restore
  // before Univer is allowed to mount. Starting as loaded lets the layout effect
  // add the session marker first, so the passive effect misclassifies the same
  // first mount as a restore and races its initial save.
  const [hostLoaded, setHostLoaded] = useState(false)
  const [hostSnapshot, setHostSnapshot] = useState<unknown>(null)
  const hydratedFromHostRef = useRef(false)
  const lastSavedRef = useRef<string | null>(null)
  const savingRef = useRef(false)
  const pendingSaveRef = useRef<Promise<void> | null>(null)
  const [layoutVersion, setLayoutVersion] = useState(0)

  useLayoutEffect(() => {
    const container = containerRef.current
    if (container === null || typeof ResizeObserver === 'undefined') return
    const observer = new ResizeObserver(() => setLayoutVersion((value) => value + 1))
    observer.observe(container)
    return () => observer.disconnect()
  }, [])

  useEffect(() => {
    const stream = chatStreamRef.current
    if (stream !== null) stream.scrollTop = stream.scrollHeight
  }, [chatLines, partialText])

  useEffect(() => {
    let cancelled = false
    runtimeRef.current?.univer.dispose()
    runtimeRef.current?.mount.remove()
    runtimeRef.current = null
    appliedRef.current.clear()
    lastSavedRef.current = null
    setHostSnapshot(null)
    setFilePath(null)
    setSaveDialogOpen(false)
    setSaveNotice('')
    hydratedFromHostRef.current = false

    if (!openedWorkbookSessions.has(sessionId)) {
      // A genuine first visit starts at the welcome screen.
      setHostLoaded(true)
      setSheetVisible(false)
      return () => { cancelled = true }
    }

    // Returning from another tab restores the workbook instead of briefly
    // showing the welcome actions and creating a new blank workbook.
    setHostLoaded(false)
    setSheetVisible(true)
    const restore = async () => {
      const [result, restoredFilePath] = await Promise.all([
        clientConnection.current?.rpc.call('/dsh-univer-create', 'load', { sessionId, unitType: 'sheet' }),
        loadUnitFilePath(sessionId, 'sheet'),
      ])
      if (cancelled) return
      setFilePath(restoredFilePath)
      if (result?.ok === true && result.value !== null && typeof result.value === 'object') {
        setHostSnapshot(result.value)
      } else {
        openedWorkbookSessions.delete(sessionId)
        setSheetVisible(false)
      }
      setHostLoaded(true)
    }
    void restore().catch((reason) => {
      if (cancelled) return
      openedWorkbookSessions.delete(sessionId)
      setSheetVisible(false)
      setHostLoaded(true)
      setError(`无法恢复表格：${reason instanceof Error ? reason.message : String(reason)}`)
    })
    return () => { cancelled = true }
  }, [sessionId])

  useLayoutEffect(() => {
    if (!hostLoaded) return
    const container = containerRef.current
    if (container === null) return
    const bounds = container.getBoundingClientRect()
    // Univer's editor cannot bootstrap while the Sheet view is hidden or its
    // flex parent has not received dimensions yet. ResizeObserver above will
    // rerun this effect once the container becomes usable.
    if (bounds.width <= 0 || bounds.height <= 0) return

    const createRuntime = (
      operation?: Extract<SheetOperation, { action: 'new' }>,
      snapshot?: unknown,
    ) => {
      // Mount each Univer instance in its own DOM island. React owns only the
      // outer host; Univer owns the island and its descendants.
      const previous = runtimeRef.current
      const mount = document.createElement('div')
      mount.className = 'dsh-univer-create-runtime-host'
      container.appendChild(mount)
      const runtime = createUniver({
        locale: LocaleType.ZH_CN,
        locales: {
          [LocaleType.ZH_CN]: mergeLocales(UniverPresetSheetsCoreZhCN),
        },
        theme: defaultTheme,
        presets: [UniverSheetsCorePreset({ container: mount })],
      })
      runtime.univerAPI.createWorkbook(normalizeWorkbookSnapshot(snapshot ?? workbookData(operation)))
      runtimeRef.current = Object.assign(runtime, { mount })
      openedWorkbookSessions.add(sessionId)
      setSheetVisible(true)
      if (previous !== null) {
        previous.univer.dispose()
        previous.mount.remove()
      }
      const snapshotTitle = snapshot !== null && typeof snapshot === 'object' && typeof (snapshot as any).name === 'string'
        ? (snapshot as any).name
        : undefined
      setTitle(operation?.title ?? snapshotTitle ?? '对话表格')
    }

    try {
      if (runtimeRef.current === null && hostSnapshot === null && operations.length === 0) return
      if (runtimeRef.current === null) {
        createRuntime(undefined, hostSnapshot ?? undefined)
        if (hostSnapshot !== null && !hydratedFromHostRef.current) {
          for (const operation of operations) appliedRef.current.add(operation.seq)
          hydratedFromHostRef.current = true
          lastSavedRef.current = JSON.stringify(hostSnapshot)
        }
      }
      for (const operation of operations) {
        if (appliedRef.current.has(operation.seq)) continue
        if (operation.action === 'new') {
          createRuntime(operation)
          setFilePath(null)
          setSaveNotice('')
          const blankSnapshot = normalizeWorkbookSnapshot(runtimeRef.current?.univerAPI.getActiveWorkbook()?.save())
          void clientConnection.current?.rpc.call('/dsh-univer-create', 'save', { sessionId, unitType: 'sheet', snapshot: blankSnapshot, filePath: null })
          appliedRef.current.clear()
          appliedRef.current.add(operation.seq)
          continue
        }

        const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
        if (workbook === null || workbook === undefined) throw new Error('当前没有可操作的工作簿')

        if (operation.action === 'list-sheets') {
          appliedRef.current.add(operation.seq)
          continue
        }

        if (operation.action === 'add-sheet' || operation.action === 'delete-sheet' || operation.action === 'rename-sheet') {
          // Univer's public Facade exposes workbook.save/createWorkbook reliably
          // across the installed beta release. Recreate from that snapshot so
          // the operation remains undo-safe and replayable without mutating
          // internal Workbook/Worksheet models directly.
          const snapshot = workbook.save() as any
          const sheets = Object.values(snapshot.sheets) as Array<any>
          const findSheet = (name: string) => sheets.find((sheet) => sheet.name === name)

          if (operation.action === 'add-sheet') {
            if (findSheet(operation.name) !== undefined) throw new Error(`工作表已存在：${operation.name}`)
            const id = `sheet-${operation.seq}`
            snapshot.sheetOrder.push(id)
            snapshot.sheets[id] = {
              id,
              name: operation.name,
              rowCount: Math.max(1, operation.rows ?? 100),
              columnCount: Math.max(1, operation.columns ?? 26),
            }
          } else if (operation.action === 'delete-sheet') {
            const sheet = findSheet(operation.name)
            if (sheet === undefined) throw new Error(`找不到工作表：${operation.name}`)
            if (snapshot.sheetOrder.length <= 1) throw new Error('至少需要保留一个工作表')
            snapshot.sheetOrder = snapshot.sheetOrder.filter((id: string) => id !== sheet.id)
            delete snapshot.sheets[sheet.id]
          } else {
            const sheet = findSheet(operation.oldName)
            if (sheet === undefined) throw new Error(`找不到工作表：${operation.oldName}`)
            if (findSheet(operation.newName) !== undefined) throw new Error(`工作表已存在：${operation.newName}`)
            sheet.name = operation.newName
          }

          // createRuntime disposes the current Univer instance exactly once
          // before mounting the updated snapshot.
          createRuntime(undefined, snapshot)
          appliedRef.current.add(operation.seq)
          continue
        }

        const worksheet = operation.sheetName
          ? workbook.getSheetByName(operation.sheetName)
          : workbook.getActiveSheet()
        if (worksheet === null || worksheet === undefined) {
          throw new Error(`找不到工作表：${operation.sheetName ?? '(active)'}`)
        }
        const range = worksheet.getRange(operation.range)
        if (operation.action === 'set-range') {
          range.setValues(operation.values.map((row) => row.map((value) => value ?? '')))
        } else if (operation.action === 'clear-range') {
          range.clearContent()
        } else {
          if (operation.background !== undefined) range.setBackground(operation.background)
          if (operation.fontColor !== undefined) range.setFontColor(operation.fontColor)
          if (operation.bold !== undefined) range.setFontWeight(operation.bold ? 'bold' : 'normal')
          if (operation.fontSize !== undefined) range.setFontSize(operation.fontSize)
        }
        appliedRef.current.add(operation.seq)
      }
      setError(null)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
    }

    return undefined
  }, [hostLoaded, hostSnapshot, operations, layoutVersion])

  useEffect(() => {
    const timer = window.setInterval(() => {
      const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
      if (workbook === undefined || workbook === null || clientConnection.current === null || savingRef.current) return
      const nextSnapshot = normalizeWorkbookSnapshot(workbook.save())
      const serialized = JSON.stringify(nextSnapshot)
      if (serialized === lastSavedRef.current) return
      savingRef.current = true
      const save = clientConnection.current.rpc.call(
        '/dsh-univer-create',
        'save',
        { sessionId, unitType: 'sheet', snapshot: nextSnapshot },
      ).then((result) => {
        if (result.ok) lastSavedRef.current = serialized
      }).catch((reason) => {
        setError(`保存到 Host 失败：${reason instanceof Error ? reason.message : String(reason)}`)
      }).finally(() => {
        savingRef.current = false
        pendingSaveRef.current = null
      })
      pendingSaveRef.current = save
    }, 800)
    return () => window.clearInterval(timer)
  }, [sessionId])

  useLayoutEffect(() => () => {
    const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
    if (workbook !== undefined && workbook !== null && clientConnection.current !== null) {
      const connection = clientConnection.current
      const nextSnapshot = normalizeWorkbookSnapshot(workbook.save())
      const finalSave = () => connection.rpc.call('/dsh-univer-create', 'save', { sessionId, unitType: 'sheet', snapshot: nextSnapshot })
      const pending = pendingSaveRef.current
      if (pending === null) void finalSave()
      else void pending.then(finalSave, finalSave)
    }
    runtimeRef.current?.univer.dispose()
    runtimeRef.current?.mount.remove()
    runtimeRef.current = null
    appliedRef.current.clear()
  }, [sessionId])

  const saveAsXlsx = async (chosenPath?: string) => {
    const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
    if (workbook === undefined || workbook === null) {
      setError('当前没有可保存的工作簿')
      return
    }
    const targetPath = chosenPath ?? filePath
    if (targetPath === null) {
      setSaveDialogOpen(true)
      return
    }
    setExporting(true)
    setError(null)
    setSaveNotice('')
    try {
      const nextSnapshot = normalizeWorkbookSnapshot(workbook.save())
      let savedPath: string
      try {
        savedPath = await saveUnitFile(sessionId, 'sheet', nextSnapshot, targetPath, chosenPath === undefined)
      } catch (reason) {
        const message = reason instanceof Error ? reason.message : String(reason)
        if (chosenPath !== undefined && message.includes('目标文件已存在') && window.confirm(`${targetPath} 已存在，是否覆盖？`)) {
          savedPath = await saveUnitFile(sessionId, 'sheet', nextSnapshot, targetPath, true)
        } else {
          throw reason
        }
      }
      setFilePath(savedPath)
      setSaveDialogOpen(false)
      setSaveNotice(`已保存到 ${savedPath}`)
      lastSavedRef.current = JSON.stringify(nextSnapshot)
    } catch (reason) {
      setError(`保存 XLSX 失败：${reason instanceof Error ? reason.message : String(reason)}`)
    } finally {
      setExporting(false)
    }
  }

  const persistHostSnapshot = async (nextSnapshot: unknown, nextFilePath: string | null) => {
    if (clientConnection.current === null) return
    await pendingSaveRef.current
    if (clientConnection.current === null) return
    const serialized = JSON.stringify(nextSnapshot)
    const result = await clientConnection.current.rpc.call(
      '/dsh-univer-create',
      'save',
      { sessionId, unitType: 'sheet', snapshot: nextSnapshot, filePath: nextFilePath },
    )
    if (!result.ok) throw new Error(result.error.message)
    lastSavedRef.current = serialized
  }

  const newWorkbook = async () => {
    const blankSnapshot = normalizeWorkbookSnapshot(workbookData({ action: 'new', seq: Date.now() }))
    runtimeRef.current?.univer.dispose()
    runtimeRef.current?.mount.remove()
    runtimeRef.current = null
    // A manually created workbook replaces both the persisted workbook and any
    // conversation operations that belong to the previous workbook.
    appliedRef.current = new Set(operations.map((operation) => operation.seq))
    hydratedFromHostRef.current = true
    lastSavedRef.current = null
    openedWorkbookSessions.add(sessionId)
    setError(null)
    setTitle('对话表格')
    setFilePath(null)
    setSaveNotice('')
    setHostSnapshot(blankSnapshot)
    try {
      await persistHostSnapshot(blankSnapshot, null)
    } catch (reason) {
      setError(`保存到 Host 失败：${reason instanceof Error ? reason.message : String(reason)}`)
    }
  }

  const importXlsx = async (file: File) => {
    setImporting(true)
    setError(null)
    try {
      const importedSnapshot = await importWorkbookFromFile(file)
      runtimeRef.current?.univer.dispose()
      runtimeRef.current?.mount.remove()
      runtimeRef.current = null
      // An imported file replaces the conversation-generated workbook. Keep the
      // historical tool calls from replaying over the imported workbook.
      appliedRef.current = new Set(operations.map((operation) => operation.seq))
      hydratedFromHostRef.current = true
      lastSavedRef.current = null
      openedWorkbookSessions.add(sessionId)
      setHostSnapshot(importedSnapshot)
      const importedPath = openedFilePath(file)
      setFilePath(importedPath)
      setSaveNotice('')
      setTitle(typeof importedSnapshot.name === 'string' ? importedSnapshot.name : file.name.replace(/\\.[^.]+$/, ''))
      await persistHostSnapshot(importedSnapshot, importedPath)
    } catch (reason) {
      setError(`导入 XLSX 失败：${reason instanceof Error ? reason.message : String(reason)}`)
    } finally {
      setImporting(false)
    }
  }

  const actionButtons = (
    <>
      <button className="dsh-univer-create-action" type="button" disabled={importing || exporting || !hostLoaded} onClick={() => fileInputRef.current?.click()}>
        {importing ? '打开中…' : '打开'}
      </button>
      <button className="dsh-univer-create-action dsh-univer-create-action--primary" type="button" disabled={importing || exporting || !hostLoaded} onClick={newWorkbook}>
        新建
      </button>
    </>
  )

  return (
    <section className="dsh-univer-create-root" aria-label="Univer Sheet">
      <input ref={fileInputRef} className="dsh-univer-create-file-input" type="file" accept=".xlsx" onChange={(event) => {
        const file = event.target.files?.[0]
        event.target.value = ''
        if (file !== undefined) void importXlsx(file)
      }} />
      {sheetVisible && (
        <div className="dsh-univer-create-toolbar">
          <span className="dsh-univer-create-title">{title}</span>
          <span className="dsh-univer-create-status">
            {saveNotice || (filePath !== null ? `文件：${filePath}` : operations.length > 0 ? `已同步 ${operations.length} 个对话操作` : '新文件，首次保存时可选择 workspace 内路径')}
          </span>
          {actionButtons}
          <button
            className="dsh-univer-create-action"
            type="button"
            disabled={importing || exporting || !hostLoaded}
            onClick={() => void saveAsXlsx()}
          >
            {exporting ? '保存中…' : '保存'}
          </button>
        </div>
      )}
      <div ref={containerRef} className="dsh-univer-create-container" />
      {saveDialogOpen && <SavePathDialog extension="xlsx" suggestedName={title} busy={exporting} onCancel={() => setSaveDialogOpen(false)} onSave={(path) => void saveAsXlsx(path)} />}
      {!sheetVisible && (
        <div className="dsh-univer-create-welcome" aria-label="开始使用 Univer Sheet">
          {actionButtons}
        </div>
      )}
      {sheetVisible && (
        <>
          <aside className="dsh-univer-create-chat-overlay" aria-label="当前会话动态">
            <div className="dsh-univer-create-chat-overlay__title">对话动态</div>
            <div ref={chatStreamRef} className="dsh-univer-create-chat-overlay__stream">
              {chatLines.map((line, index) => (
                <div className="dsh-univer-create-chat-overlay__line" key={`${line.role}-${index}-${line.text.slice(0, 16)}`}>
                  <strong>{line.role}</strong><span>{line.text}</span>
                </div>
              ))}
              {partialText && <div className="dsh-univer-create-chat-overlay__line dsh-univer-create-chat-overlay__line--live"><strong>AI</strong><span>{partialText}</span></div>}
            </div>
          </aside>
          {error !== null && <div className="dsh-univer-create-error" role="alert">表格操作失败：{error}</div>}
        </>
      )}
    </section>
  )
}

const openedDocumentSessions = new Set<string>()

function DocProductView(props: ConvViewProps) {
  const { sessionId } = props
  const { nodes, partialText } = useConversationFeed(props)
  const operations = useMemo(() => nodes.map(docOperationFromNode).filter((item): item is DocOperation => item !== null), [nodes])
  const chatLines = useMemo(() => nodes
    .map(conversationLabel)
    .filter((item): item is { role: string; text: string } => item !== null)
    .slice(-4), [nodes])
  const containerRef = useRef<HTMLDivElement>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)
  const runtimeRef = useRef<MountedRuntime | null>(null)
  const appliedRef = useRef(new Set<number>())
  const chatStreamRef = useRef<HTMLDivElement>(null)
  const [title, setTitle] = useState('对话文档')
  const [error, setError] = useState<string | null>(null)
  const [importing, setImporting] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [filePath, setFilePath] = useState<string | null>(null)
  const [saveDialogOpen, setSaveDialogOpen] = useState(false)
  const [saveNotice, setSaveNotice] = useState('')
  const [documentVisible, setDocumentVisible] = useState(() => openedDocumentSessions.has(sessionId))
  const [hostLoaded, setHostLoaded] = useState(false)
  const [hostSnapshot, setHostSnapshot] = useState<unknown>(null)
  const hydratedFromHostRef = useRef(false)
  const lastSavedRef = useRef<string | null>(null)
  const savingRef = useRef(false)
  const pendingSaveRef = useRef<Promise<void> | null>(null)
  const [layoutVersion, setLayoutVersion] = useState(0)

  useLayoutEffect(() => {
    const container = containerRef.current
    if (container === null || typeof ResizeObserver === 'undefined') return
    const observer = new ResizeObserver(() => setLayoutVersion((value) => value + 1))
    observer.observe(container)
    return () => observer.disconnect()
  }, [])

  useEffect(() => {
    const stream = chatStreamRef.current
    if (stream !== null) stream.scrollTop = stream.scrollHeight
  }, [chatLines, partialText])

  useEffect(() => {
    let cancelled = false
    runtimeRef.current?.univer.dispose()
    runtimeRef.current?.mount.remove()
    runtimeRef.current = null
    appliedRef.current.clear()
    lastSavedRef.current = null
    setHostSnapshot(null)
    setFilePath(null)
    setSaveDialogOpen(false)
    setSaveNotice('')
    hydratedFromHostRef.current = false

    if (!openedDocumentSessions.has(sessionId)) {
      setHostLoaded(true)
      setDocumentVisible(false)
      return () => { cancelled = true }
    }

    setHostLoaded(false)
    setDocumentVisible(true)
    const restore = async () => {
      const [result, restoredFilePath] = await Promise.all([
        clientConnection.current?.rpc.call('/dsh-univer-create', 'load', { sessionId, unitType: 'doc' }),
        loadUnitFilePath(sessionId, 'doc'),
      ])
      if (cancelled) return
      setFilePath(restoredFilePath)
      if (result?.ok === true && result.value !== null && typeof result.value === 'object') {
        setHostSnapshot(result.value)
      } else {
        openedDocumentSessions.delete(sessionId)
        setDocumentVisible(false)
      }
      setHostLoaded(true)
    }
    void restore().catch((reason) => {
      if (cancelled) return
      openedDocumentSessions.delete(sessionId)
      setDocumentVisible(false)
      setHostLoaded(true)
      setError(`无法恢复文档：${reason instanceof Error ? reason.message : String(reason)}`)
    })
    return () => { cancelled = true }
  }, [sessionId])

  useLayoutEffect(() => {
    if (!hostLoaded) return
    const container = containerRef.current
    if (container === null) return
    const bounds = container.getBoundingClientRect()
    if (bounds.width <= 0 || bounds.height <= 0) return

    const createRuntime = (
      operation?: Extract<DocOperation, { action: 'new-doc' }>,
      restoredSnapshot?: unknown,
    ) => {
      const previous = runtimeRef.current
      const mount = document.createElement('div')
      mount.className = 'dsh-univer-create-runtime-host'
      container.appendChild(mount)
      const runtime = createUniver({
        locale: LocaleType.ZH_CN,
        locales: {
          [LocaleType.ZH_CN]: mergeLocales(UniverPresetDocsCoreZhCN, UniverPresetDocsDrawingZhCN),
        },
        theme: defaultTheme,
        presets: [
          UniverDocsCorePreset({ container: mount }),
          UniverDocsDrawingPreset(),
        ],
      })
      const fDocument = runtime.univerAPI.createDocument((restoredSnapshot ?? documentData(operation)) as any)
      if (restoredSnapshot === undefined && operation?.text) fDocument.insertText(0, operation.text)
      runtimeRef.current = Object.assign(runtime, { mount })
      openedDocumentSessions.add(sessionId)
      setDocumentVisible(true)
      if (previous !== null) {
        previous.univer.dispose()
        previous.mount.remove()
      }
      const snapshotTitle = restoredSnapshot !== null && typeof restoredSnapshot === 'object' && typeof (restoredSnapshot as any).title === 'string'
        ? (restoredSnapshot as any).title
        : undefined
      setTitle(operation?.title ?? snapshotTitle ?? '对话文档')
    }

    try {
      if (runtimeRef.current === null && hostSnapshot === null && operations.length === 0) return
      if (runtimeRef.current === null) {
        createRuntime(undefined, hostSnapshot ?? undefined)
        if (hostSnapshot !== null && !hydratedFromHostRef.current) {
          for (const operation of operations) appliedRef.current.add(operation.seq)
          hydratedFromHostRef.current = true
          lastSavedRef.current = JSON.stringify(hostSnapshot)
        }
      }

      for (const operation of operations) {
        if (appliedRef.current.has(operation.seq)) continue
        if (operation.action === 'new-doc') {
          createRuntime(operation)
          setFilePath(null)
          setSaveNotice('')
          const blankSnapshot = runtimeRef.current?.univerAPI.getActiveDocument()?.save()
          void clientConnection.current?.rpc.call('/dsh-univer-create', 'save', { sessionId, unitType: 'doc', snapshot: blankSnapshot, filePath: null })
          appliedRef.current.clear()
          appliedRef.current.add(operation.seq)
          continue
        }

        const fDocument = runtimeRef.current?.univerAPI.getActiveDocument()
        if (fDocument === null || fDocument === undefined) throw new Error('当前没有可操作的文档')
        if (operation.action === 'get-doc-text') {
          appliedRef.current.add(operation.seq)
          continue
        }

        const contentLength = documentContentLength(fDocument)
        if (operation.action === 'set-doc-text') {
          if (contentLength > 0 && !fDocument.deleteRange({ startOffset: 0, endOffset: contentLength })) throw new Error('无法清空当前文档')
          if (operation.text.length > 0 && !fDocument.insertText(0, operation.text)) throw new Error('无法写入文档正文')
        } else if (operation.action === 'append-doc-text') {
          if (operation.text.length > 0 && !fDocument.insertText(contentLength, operation.text)) throw new Error('无法追加文档正文')
        } else if (operation.action === 'insert-doc-text') {
          if (!Number.isInteger(operation.index) || operation.index < 0 || operation.index > contentLength) throw new Error(`无效的插入位置：${operation.index}`)
          if (operation.text.length > 0 && !fDocument.insertText(operation.index, operation.text)) throw new Error('无法插入文档文本')
        } else {
          if (!Number.isInteger(operation.start) || !Number.isInteger(operation.end) || operation.start < 0 || operation.end < operation.start || operation.end > contentLength) {
            throw new Error(`无效的文档范围：[${operation.start}, ${operation.end})`)
          }
          if (operation.action === 'delete-doc-range') {
            if (operation.end > operation.start && !fDocument.deleteRange({ startOffset: operation.start, endOffset: operation.end })) throw new Error('无法删除文档文本')
          } else if (operation.end > operation.start) {
            const style: Record<string, unknown> = {}
            if (operation.bold !== undefined) style.bl = operation.bold ? 1 : 0
            if (operation.italic !== undefined) style.it = operation.italic ? 1 : 0
            if (operation.fontSize !== undefined) style.fs = operation.fontSize
            if (operation.fontColor !== undefined) style.cl = { rgb: operation.fontColor }
            if (Object.keys(style).length > 0 && !fDocument.getTextRange(operation.start, operation.end).setTextStyle(style as any)) throw new Error('无法设置文档格式')
          }
        }
        appliedRef.current.add(operation.seq)
      }
      setError(null)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
    }
    return undefined
  }, [hostLoaded, hostSnapshot, operations, layoutVersion])

  useEffect(() => {
    const timer = window.setInterval(() => {
      const fDocument = runtimeRef.current?.univerAPI.getActiveDocument()
      if (fDocument === undefined || fDocument === null || clientConnection.current === null || savingRef.current) return
      const nextSnapshot = fDocument.save()
      const serialized = JSON.stringify(nextSnapshot)
      if (serialized === lastSavedRef.current) return
      savingRef.current = true
      const save = clientConnection.current.rpc.call(
        '/dsh-univer-create',
        'save',
        { sessionId, unitType: 'doc', snapshot: nextSnapshot },
      ).then((result) => {
        if (result.ok) lastSavedRef.current = serialized
      }).catch((reason) => {
        setError(`保存到 Host 失败：${reason instanceof Error ? reason.message : String(reason)}`)
      }).finally(() => {
        savingRef.current = false
        pendingSaveRef.current = null
      })
      pendingSaveRef.current = save
    }, 800)
    return () => window.clearInterval(timer)
  }, [sessionId])

  useLayoutEffect(() => () => {
    const fDocument = runtimeRef.current?.univerAPI.getActiveDocument()
    if (fDocument !== undefined && fDocument !== null && clientConnection.current !== null) {
      const connection = clientConnection.current
      const nextSnapshot = fDocument.save()
      const finalSave = () => connection.rpc.call('/dsh-univer-create', 'save', { sessionId, unitType: 'doc', snapshot: nextSnapshot })
      const pending = pendingSaveRef.current
      if (pending === null) void finalSave()
      else void pending.then(finalSave, finalSave)
    }
    runtimeRef.current?.univer.dispose()
    runtimeRef.current?.mount.remove()
    runtimeRef.current = null
    appliedRef.current.clear()
  }, [sessionId])

  const persistHostSnapshot = async (nextSnapshot: unknown, nextFilePath: string | null) => {
    if (clientConnection.current === null) return
    await pendingSaveRef.current
    if (clientConnection.current === null) return
    const serialized = JSON.stringify(nextSnapshot)
    const result = await clientConnection.current.rpc.call(
      '/dsh-univer-create',
      'save',
      { sessionId, unitType: 'doc', snapshot: nextSnapshot, filePath: nextFilePath },
    )
    if (!result.ok) throw new Error(result.error.message)
    lastSavedRef.current = serialized
  }

  const newDocument = async () => {
    const blankSnapshot = documentData({ action: 'new-doc', seq: Date.now() })
    runtimeRef.current?.univer.dispose()
    runtimeRef.current?.mount.remove()
    runtimeRef.current = null
    appliedRef.current = new Set(operations.map((operation) => operation.seq))
    hydratedFromHostRef.current = true
    lastSavedRef.current = null
    openedDocumentSessions.add(sessionId)
    setError(null)
    setTitle('对话文档')
    setFilePath(null)
    setSaveNotice('')
    setHostSnapshot(blankSnapshot)
    try {
      await persistHostSnapshot(blankSnapshot, null)
    } catch (reason) {
      setError(`保存到 Host 失败：${reason instanceof Error ? reason.message : String(reason)}`)
    }
  }

  const importDocx = async (file: File) => {
    setImporting(true)
    setError(null)
    try {
      const importedSnapshot = await importDocumentFromFile(file)
      runtimeRef.current?.univer.dispose()
      runtimeRef.current?.mount.remove()
      runtimeRef.current = null
      appliedRef.current = new Set(operations.map((operation) => operation.seq))
      hydratedFromHostRef.current = true
      lastSavedRef.current = null
      openedDocumentSessions.add(sessionId)
      setHostSnapshot(importedSnapshot)
      const importedPath = openedFilePath(file)
      setFilePath(importedPath)
      setSaveNotice('')
      setTitle(typeof importedSnapshot.title === 'string' ? importedSnapshot.title : file.name.replace(/\\.[^.]+$/, ''))
      await persistHostSnapshot(importedSnapshot, importedPath)
    } catch (reason) {
      setError(`导入 DOCX 失败：${reason instanceof Error ? reason.message : String(reason)}`)
    } finally {
      setImporting(false)
    }
  }

  const saveAsDocx = async (chosenPath?: string) => {
    const fDocument = runtimeRef.current?.univerAPI.getActiveDocument()
    if (fDocument === undefined || fDocument === null) {
      setError('当前没有可保存的文档')
      return
    }
    const targetPath = chosenPath ?? filePath
    if (targetPath === null) {
      setSaveDialogOpen(true)
      return
    }
    setExporting(true)
    setError(null)
    setSaveNotice('')
    try {
      const nextSnapshot = fDocument.save()
      let savedPath: string
      try {
        savedPath = await saveUnitFile(sessionId, 'doc', nextSnapshot, targetPath, chosenPath === undefined)
      } catch (reason) {
        const message = reason instanceof Error ? reason.message : String(reason)
        if (chosenPath !== undefined && message.includes('目标文件已存在') && window.confirm(`${targetPath} 已存在，是否覆盖？`)) {
          savedPath = await saveUnitFile(sessionId, 'doc', nextSnapshot, targetPath, true)
        } else {
          throw reason
        }
      }
      setFilePath(savedPath)
      setSaveDialogOpen(false)
      setSaveNotice(`已保存到 ${savedPath}`)
      lastSavedRef.current = JSON.stringify(nextSnapshot)
    } catch (reason) {
      setError(`保存 DOCX 失败：${reason instanceof Error ? reason.message : String(reason)}`)
    } finally {
      setExporting(false)
    }
  }

  const actionButtons = (
    <>
      <button className="dsh-univer-create-action" type="button" disabled={importing || exporting || !hostLoaded} onClick={() => fileInputRef.current?.click()}>
        {importing ? '打开中…' : '打开'}
      </button>
      <button className="dsh-univer-create-action dsh-univer-create-action--primary" type="button" disabled={importing || exporting || !hostLoaded} onClick={newDocument}>
        新建
      </button>
    </>
  )

  return (
    <section className="dsh-univer-create-root" aria-label="Univer Doc">
      <input ref={fileInputRef} className="dsh-univer-create-file-input" type="file" accept=".docx" onChange={(event) => {
        const file = event.target.files?.[0]
        event.target.value = ''
        if (file !== undefined) void importDocx(file)
      }} />
      {documentVisible && (
        <div className="dsh-univer-create-toolbar">
          <span className="dsh-univer-create-title">{title}</span>
          <span className="dsh-univer-create-status">
            {saveNotice || (filePath !== null ? `文件：${filePath}` : operations.length > 0 ? `已同步 ${operations.length} 个文档操作` : '新文件，首次保存时可选择 workspace 内路径')}
          </span>
          {actionButtons}
          <button className="dsh-univer-create-action" type="button" disabled={importing || exporting || !hostLoaded} onClick={() => void saveAsDocx()}>
            {exporting ? '保存中…' : '保存'}
          </button>
        </div>
      )}
      <div ref={containerRef} className="dsh-univer-create-container" />
      {saveDialogOpen && <SavePathDialog extension="docx" suggestedName={title} busy={exporting} onCancel={() => setSaveDialogOpen(false)} onSave={(path) => void saveAsDocx(path)} />}
      {!documentVisible && <div className="dsh-univer-create-welcome" aria-label="开始使用 Univer Doc">{actionButtons}</div>}
      {documentVisible && (
        <>
          <aside className="dsh-univer-create-chat-overlay" aria-label="当前会话动态">
            <div className="dsh-univer-create-chat-overlay__title">对话动态</div>
            <div ref={chatStreamRef} className="dsh-univer-create-chat-overlay__stream">
              {chatLines.map((line, index) => (
                <div className="dsh-univer-create-chat-overlay__line" key={`${line.role}-${index}-${line.text.slice(0, 16)}`}>
                  <strong>{line.role}</strong><span>{line.text}</span>
                </div>
              ))}
              {partialText && <div className="dsh-univer-create-chat-overlay__line dsh-univer-create-chat-overlay__line--live"><strong>AI</strong><span>{partialText}</span></div>}
            </div>
          </aside>
          {error !== null && <div className="dsh-univer-create-error" role="alert">文档操作失败：{error}</div>}
        </>
      )}
    </section>
  )
}

const openedPresentationSessions = new Set<string>()

function SlideProductView(props: ConvViewProps) {
  const { sessionId } = props
  const { nodes, partialText } = useConversationFeed(props)
  const operations = useMemo(() => nodes.map(slideOperationFromNode).filter((item): item is SlideOperation => item !== null), [nodes])
  const chatLines = useMemo(() => nodes
    .map(conversationLabel)
    .filter((item): item is { role: string; text: string } => item !== null)
    .slice(-4), [nodes])
  const containerRef = useRef<HTMLDivElement>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)
  const runtimeRef = useRef<MountedSlideRuntime | null>(null)
  const appliedRef = useRef(new Set<number>())
  const chatStreamRef = useRef<HTMLDivElement>(null)
  const [title, setTitle] = useState('对话演示文稿')
  const [error, setError] = useState<string | null>(null)
  const [importing, setImporting] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [filePath, setFilePath] = useState<string | null>(null)
  const [saveDialogOpen, setSaveDialogOpen] = useState(false)
  const [saveNotice, setSaveNotice] = useState('')
  const [presentationVisible, setPresentationVisible] = useState(() => openedPresentationSessions.has(sessionId))
  const [hostLoaded, setHostLoaded] = useState(false)
  const [hostSnapshot, setHostSnapshot] = useState<ISlideData | null>(null)
  const hydratedFromHostRef = useRef(false)
  const lastSavedRef = useRef<string | null>(null)
  const savingRef = useRef(false)
  const pendingSaveRef = useRef<Promise<void> | null>(null)
  const slideResizeTimersRef = useRef<number[]>([])
  const [layoutVersion, setLayoutVersion] = useState(0)

  useLayoutEffect(() => {
    const container = containerRef.current
    if (container === null || typeof ResizeObserver === 'undefined') return
    let frame = 0
    const notifyResize = () => {
      window.cancelAnimationFrame(frame)
      frame = window.requestAnimationFrame(() => {
        setLayoutVersion((value) => value + 1)
        window.dispatchEvent(new Event('resize'))
      })
    }
    const observer = new ResizeObserver(notifyResize)
    observer.observe(container)
    notifyResize()
    return () => {
      observer.disconnect()
      window.cancelAnimationFrame(frame)
    }
  }, [])

  useEffect(() => {
    const stream = chatStreamRef.current
    if (stream !== null) stream.scrollTop = stream.scrollHeight
  }, [chatLines, partialText])

  useEffect(() => {
    let cancelled = false
    runtimeRef.current?.univer.dispose()
    runtimeRef.current?.mount.remove()
    runtimeRef.current = null
    appliedRef.current.clear()
    lastSavedRef.current = null
    setHostSnapshot(null)
    setFilePath(null)
    setSaveDialogOpen(false)
    setSaveNotice('')
    hydratedFromHostRef.current = false

    if (!openedPresentationSessions.has(sessionId)) {
      setHostLoaded(true)
      setPresentationVisible(false)
      return () => { cancelled = true }
    }

    setHostLoaded(false)
    setPresentationVisible(true)
    const restore = async () => {
      const [result, restoredFilePath] = await Promise.all([
        clientConnection.current?.rpc.call('/dsh-univer-create', 'load', { sessionId, unitType: 'slide' }),
        loadUnitFilePath(sessionId, 'slide'),
      ])
      if (cancelled) return
      setFilePath(restoredFilePath)
      if (result?.ok === true && result.value !== null && typeof result.value === 'object') {
        setHostSnapshot(result.value as ISlideData)
      } else {
        openedPresentationSessions.delete(sessionId)
        setPresentationVisible(false)
      }
      setHostLoaded(true)
    }
    void restore().catch((reason) => {
      if (cancelled) return
      openedPresentationSessions.delete(sessionId)
      setPresentationVisible(false)
      setHostLoaded(true)
      setError(`无法恢复演示文稿：${reason instanceof Error ? reason.message : String(reason)}`)
    })
    return () => { cancelled = true }
  }, [sessionId])

  useLayoutEffect(() => {
    if (!hostLoaded) return
    const container = containerRef.current
    if (container === null) return
    const bounds = container.getBoundingClientRect()
    if (bounds.width <= 0 || bounds.height <= 0) return

    const createRuntime = (
      operation?: Extract<SlideOperation, { action: 'new-slide' }>,
      restoredSnapshot?: ISlideData,
    ) => {
      const previous = runtimeRef.current
      slideResizeTimersRef.current.forEach((timer) => window.clearTimeout(timer))
      slideResizeTimersRef.current = []
      const mount = document.createElement('div')
      mount.className = 'dsh-univer-create-runtime-host'
      container.appendChild(mount)

      const univer = new Univer({
        locale: LocaleType.ZH_CN,
        locales: {
          [LocaleType.ZH_CN]: mergeLocales(DesignZhCN, UIZhCN, DocsUIZhCN, ShapeEditorUIZhCN, SlidesUIZhCN),
        },
      })
      univer.registerPlugin(UniverRenderEnginePlugin)
      univer.registerPlugin(UniverUIPlugin, { container: mount })
      univer.registerPlugin(UniverDocsPlugin)
      univer.registerPlugin(UniverDocsUIPlugin)
      univer.registerPlugin(UniverDrawingPlugin)
      univer.registerPlugin(UniverLicensePlugin)
      univer.registerPlugin(UniverSlidesPlugin)
      univer.registerPlugin(UniverSlidesUIPlugin)

      const univerAPI = FUniver.newAPI(univer)
      const presentation = univerAPI.createPresentation(restoredSnapshot ?? presentationData(operation))
      runtimeRef.current = { univer, univerAPI, mount, presentation }
      openedPresentationSessions.add(sessionId)
      setPresentationVisible(true)
      if (previous !== null) {
        previous.univer.dispose()
        previous.mount.remove()
      }
      setTitle(operation?.title ?? restoredSnapshot?.name ?? '对话演示文稿')

      // Slides UI mounts the main render scene and each page thumbnail scene
      // asynchronously. Re-measure after both React and drawing resources settle,
      // otherwise the left page bar can keep zero-sized blank canvases.
      for (const delay of [0, 100, 300, 800]) {
        const timer = window.setTimeout(() => {
          if (runtimeRef.current?.univer === univer) window.dispatchEvent(new Event('resize'))
        }, delay)
        slideResizeTimersRef.current.push(timer)
      }
    }

    try {
      if (runtimeRef.current === null && hostSnapshot === null && operations.length === 0) return
      if (runtimeRef.current === null) {
        createRuntime(undefined, hostSnapshot ?? undefined)
        if (hostSnapshot !== null && !hydratedFromHostRef.current) {
          for (const operation of operations) appliedRef.current.add(operation.seq)
          hydratedFromHostRef.current = true
          lastSavedRef.current = JSON.stringify(hostSnapshot)
        }
      }

      for (const operation of operations) {
        if (appliedRef.current.has(operation.seq)) continue
        if (operation.action === 'new-slide') {
          createRuntime(operation)
          setFilePath(null)
          setSaveNotice('')
          const blankSnapshot = runtimeRef.current?.presentation.save()
          void clientConnection.current?.rpc.call('/dsh-univer-create', 'save', { sessionId, unitType: 'slide', snapshot: blankSnapshot, filePath: null })
          appliedRef.current.clear()
          appliedRef.current.add(operation.seq)
          continue
        }
        if (operation.action === 'list-slides') {
          appliedRef.current.add(operation.seq)
          continue
        }

        const runtime = runtimeRef.current
        if (runtime === null) throw new Error('当前没有可操作的演示文稿')
        const presentation = runtime.presentation
        if (operation.action === 'add-slide') {
          const slides = presentation.getSlides()
          const index = operation.index ?? slides.length
          if (!Number.isInteger(index) || index < 0 || index > slides.length) throw new Error(`无效的插入位置：${index}`)
          const options = { id: `slide-page-${operation.seq}`, name: operation.title ?? `幻灯片 ${index + 1}` }
          if (index === slides.length) presentation.appendSlide(options)
          else presentation.insertSlide(index, options)
        } else if (operation.action === 'delete-slide') {
          const slides = presentation.getSlides()
          if (slides.length <= 1) throw new Error('至少需要保留一张幻灯片')
          const slide = presentation.getSlideByIndex(operation.index)
          if (slide === null) throw new Error(`无效的幻灯片索引：${operation.index}`)
          if (!presentation.deleteSlide(slide)) throw new Error(`无法删除第 ${operation.index + 1} 张幻灯片`)
        } else {
          const slide = presentation.getSlideByIndex(operation.slideIndex)
          if (slide === null) throw new Error(`无效的幻灯片索引：${operation.slideIndex}`)
          if (operation.action === 'delete-slide-element') {
            const element = slide.getElementById(operation.elementId)
            if (element === null) throw new Error(`找不到幻灯片元素：${operation.elementId}`)
            if (!slide.deleteElement(element)) throw new Error(`无法删除幻灯片元素：${operation.elementId}`)
          } else if (operation.action === 'add-slide-text') {
            const element: ISlideTextElement = {
              id: `slide-text-${operation.seq}`,
              type: PageElementTypeEnum.Text,
              transform: {
                left: operation.left ?? 120,
                top: operation.top ?? 100,
                width: Math.max(1, operation.width ?? 720),
                height: Math.max(1, operation.height ?? 80),
              },
              name: 'text',
              visible: true,
              selectable: true,
              text: operation.text,
              textStyle: {
                fontSize: operation.fontSize ?? 30,
                color: operation.fontColor ?? '#333333',
                bold: operation.bold ?? true,
              },
            }
            slide.insertElement(element)
          } else {
            const element = slide.getElementById(operation.elementId)
            if (element === null || !('getType' in element) || element.getType() !== PageElementTypeEnum.Text) {
              throw new Error(`找不到文本元素：${operation.elementId}`)
            }
            const data = element.getData() as ISlideTextElement
            if (operation.text !== undefined || operation.fontSize !== undefined || operation.fontColor !== undefined || operation.bold !== undefined) {
              const richText = runtime.univerAPI.newRichText().span(operation.text ?? data.text ?? '', {
                fontSize: operation.fontSize ?? data.textStyle?.fontSize ?? 30,
                color: operation.fontColor ?? data.textStyle?.color ?? '#333333',
                bold: operation.bold ?? data.textStyle?.bold ?? true,
              })
              element.setRichText(richText)
            }
            const transform = element.getTransform()
            if (operation.left !== undefined || operation.top !== undefined) {
              element.setPosition(operation.left ?? transform.left ?? 0, operation.top ?? transform.top ?? 0)
            }
            if (operation.width !== undefined || operation.height !== undefined) {
              element.setSize(Math.max(1, operation.width ?? transform.width ?? 1), Math.max(1, operation.height ?? transform.height ?? 1))
            }
          }
        }
        appliedRef.current.add(operation.seq)
      }
      setError(null)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
    }
    return undefined
  }, [hostLoaded, hostSnapshot, operations, layoutVersion])

  useEffect(() => {
    const timer = window.setInterval(() => {
      const presentation = runtimeRef.current?.presentation
      if (presentation === undefined || clientConnection.current === null || savingRef.current) return
      const nextSnapshot = presentation.save()
      const serialized = JSON.stringify(nextSnapshot)
      if (serialized === lastSavedRef.current) return
      savingRef.current = true
      const save = clientConnection.current.rpc.call(
        '/dsh-univer-create',
        'save',
        { sessionId, unitType: 'slide', snapshot: nextSnapshot },
      ).then((result) => {
        if (result.ok) lastSavedRef.current = serialized
      }).catch((reason) => {
        setError(`保存到 Host 失败：${reason instanceof Error ? reason.message : String(reason)}`)
      }).finally(() => {
        savingRef.current = false
        pendingSaveRef.current = null
      })
      pendingSaveRef.current = save
    }, 800)
    return () => window.clearInterval(timer)
  }, [sessionId])

  useLayoutEffect(() => () => {
    const presentation = runtimeRef.current?.presentation
    if (presentation !== undefined && clientConnection.current !== null) {
      const connection = clientConnection.current
      const nextSnapshot = presentation.save()
      const finalSave = () => connection.rpc.call('/dsh-univer-create', 'save', { sessionId, unitType: 'slide', snapshot: nextSnapshot })
      const pending = pendingSaveRef.current
      if (pending === null) void finalSave()
      else void pending.then(finalSave, finalSave)
    }
    slideResizeTimersRef.current.forEach((timer) => window.clearTimeout(timer))
    slideResizeTimersRef.current = []
    runtimeRef.current?.univer.dispose()
    runtimeRef.current?.mount.remove()
    runtimeRef.current = null
    appliedRef.current.clear()
  }, [sessionId])

  const persistHostSnapshot = async (nextSnapshot: ISlideData, nextFilePath: string | null) => {
    if (clientConnection.current === null) return
    await pendingSaveRef.current
    if (clientConnection.current === null) return
    const serialized = JSON.stringify(nextSnapshot)
    const result = await clientConnection.current.rpc.call(
      '/dsh-univer-create',
      'save',
      { sessionId, unitType: 'slide', snapshot: nextSnapshot, filePath: nextFilePath },
    )
    if (!result.ok) throw new Error(result.error.message)
    lastSavedRef.current = serialized
  }

  const replacePresentation = async (nextSnapshot: ISlideData, nextTitle: string, nextFilePath: string | null) => {
    slideResizeTimersRef.current.forEach((timer) => window.clearTimeout(timer))
    slideResizeTimersRef.current = []
    runtimeRef.current?.univer.dispose()
    runtimeRef.current?.mount.remove()
    runtimeRef.current = null
    appliedRef.current = new Set(operations.map((operation) => operation.seq))
    hydratedFromHostRef.current = true
    lastSavedRef.current = null
    openedPresentationSessions.add(sessionId)
    setError(null)
    setTitle(nextTitle)
    setFilePath(nextFilePath)
    setSaveNotice('')
    setHostSnapshot(nextSnapshot)
    await persistHostSnapshot(nextSnapshot, nextFilePath)
  }

  const newPresentation = async () => {
    const blankSnapshot = presentationData({ action: 'new-slide', seq: Date.now() })
    try {
      await replacePresentation(blankSnapshot, '对话演示文稿', null)
    } catch (reason) {
      setError(`保存到 Host 失败：${reason instanceof Error ? reason.message : String(reason)}`)
    }
  }

  const importPptx = async (file: File) => {
    setImporting(true)
    setError(null)
    try {
      const importedSnapshot = await importPresentationFromFile(file)
      await replacePresentation(importedSnapshot, importedSnapshot.name || file.name.replace(/\.[^.]+$/, ''), openedFilePath(file))
    } catch (reason) {
      setError(`导入 PPTX 失败：${reason instanceof Error ? reason.message : String(reason)}`)
    } finally {
      setImporting(false)
    }
  }

  const saveAsPptx = async (chosenPath?: string) => {
    const presentation = runtimeRef.current?.presentation
    if (presentation === undefined) {
      setError('当前没有可保存的演示文稿')
      return
    }
    const targetPath = chosenPath ?? filePath
    if (targetPath === null) {
      setSaveDialogOpen(true)
      return
    }
    setExporting(true)
    setError(null)
    setSaveNotice('')
    try {
      const nextSnapshot = presentation.save()
      let savedPath: string
      try {
        savedPath = await saveUnitFile(sessionId, 'slide', nextSnapshot, targetPath, chosenPath === undefined)
      } catch (reason) {
        const message = reason instanceof Error ? reason.message : String(reason)
        if (chosenPath !== undefined && message.includes('目标文件已存在') && window.confirm(`${targetPath} 已存在，是否覆盖？`)) {
          savedPath = await saveUnitFile(sessionId, 'slide', nextSnapshot, targetPath, true)
        } else {
          throw reason
        }
      }
      setFilePath(savedPath)
      setSaveDialogOpen(false)
      setSaveNotice(`已保存到 ${savedPath}`)
      lastSavedRef.current = JSON.stringify(nextSnapshot)
    } catch (reason) {
      setError(`保存 PPTX 失败：${reason instanceof Error ? reason.message : String(reason)}`)
    } finally {
      setExporting(false)
    }
  }

  const actionButtons = (
    <>
      <button className="dsh-univer-create-action" type="button" disabled={importing || exporting || !hostLoaded} onClick={() => fileInputRef.current?.click()}>
        {importing ? '打开中…' : '打开'}
      </button>
      <button className="dsh-univer-create-action dsh-univer-create-action--primary" type="button" disabled={importing || exporting || !hostLoaded} onClick={() => void newPresentation()}>
        新建
      </button>
    </>
  )

  return (
    <section className="dsh-univer-create-root" aria-label="Univer Slide">
      <input ref={fileInputRef} className="dsh-univer-create-file-input" type="file" accept=".ppt,.pptx" onChange={(event) => {
        const file = event.target.files?.[0]
        event.target.value = ''
        if (file !== undefined) void importPptx(file)
      }} />
      {presentationVisible && (
        <div className="dsh-univer-create-toolbar">
          <span className="dsh-univer-create-title">{title}</span>
          <span className="dsh-univer-create-status">
            {saveNotice || (filePath !== null ? `文件：${filePath}` : operations.length > 0 ? `已同步 ${operations.length} 个幻灯片操作` : '新文件，首次保存时可选择 workspace 内路径')}
          </span>
          {actionButtons}
          <button className="dsh-univer-create-action" type="button" disabled={importing || exporting || !hostLoaded} onClick={() => void saveAsPptx()}>
            {exporting ? '保存中…' : '保存'}
          </button>
        </div>
      )}
      <div ref={containerRef} className="dsh-univer-create-container dsh-univer-create-container--slides" />
      {saveDialogOpen && <SavePathDialog extension="pptx" suggestedName={title} busy={exporting} onCancel={() => setSaveDialogOpen(false)} onSave={(path) => void saveAsPptx(path)} />}
      {!presentationVisible && <div className="dsh-univer-create-welcome" aria-label="开始使用 Univer Slide">{actionButtons}</div>}
      {presentationVisible && (
        <>
          <aside className="dsh-univer-create-chat-overlay" aria-label="当前会话动态">
            <div className="dsh-univer-create-chat-overlay__title">对话动态</div>
            <div ref={chatStreamRef} className="dsh-univer-create-chat-overlay__stream">
              {chatLines.map((line, index) => (
                <div className="dsh-univer-create-chat-overlay__line" key={`${line.role}-${index}-${line.text.slice(0, 16)}`}>
                  <strong>{line.role}</strong><span>{line.text}</span>
                </div>
              ))}
              {partialText && <div className="dsh-univer-create-chat-overlay__line dsh-univer-create-chat-overlay__line--live"><strong>AI</strong><span>{partialText}</span></div>}
            </div>
          </aside>
          {error !== null && <div className="dsh-univer-create-error" role="alert">幻灯片操作失败：{error}</div>}
        </>
      )}
    </section>
  )
}

const selectedUnitBySession = new Map<string, UniverUnitType>()

function UniverView(props: ConvViewProps) {
  const { nodes } = useConversationFeed(props)
  const suggestedUnit = useMemo<UniverUnitType>(() => {
    for (let index = nodes.length - 1; index >= 0; index -= 1) {
      const node = nodes[index]!
      if (node.kind !== 'tool-result' || node.call === null) continue
      if (SLIDE_TOOL_NAMES.has(node.call.name)) return 'slide'
      if (DOC_TOOL_NAMES.has(node.call.name)) return 'doc'
      if (SHEET_TOOL_NAMES.has(node.call.name)) return 'sheet'
    }
    return 'sheet'
  }, [nodes])
  const [unitType, setUnitType] = useState<UniverUnitType>(() => selectedUnitBySession.get(props.sessionId) ?? suggestedUnit)

  useEffect(() => {
    const remembered = selectedUnitBySession.get(props.sessionId)
    setUnitType(remembered ?? suggestedUnit)
  }, [props.sessionId, suggestedUnit])

  const selectUnit = (next: UniverUnitType) => {
    selectedUnitBySession.set(props.sessionId, next)
    setUnitType(next)
  }

  return (
    <section className="dsh-univer-create-shell" aria-label="Univer Sheet、Doc 与 Slide">
      <nav className="dsh-univer-create-product-bar" aria-label="选择 Univer 编辑器">
        <button className={`dsh-univer-create-product-tab${unitType === 'sheet' ? ' is-active' : ''}`} type="button" aria-pressed={unitType === 'sheet'} onClick={() => selectUnit('sheet')}>Sheet</button>
        <button className={`dsh-univer-create-product-tab${unitType === 'doc' ? ' is-active' : ''}`} type="button" aria-pressed={unitType === 'doc'} onClick={() => selectUnit('doc')}>Doc</button>
        <button className={`dsh-univer-create-product-tab${unitType === 'slide' ? ' is-active' : ''}`} type="button" aria-pressed={unitType === 'slide'} onClick={() => selectUnit('slide')}>Slide</button>
      </nav>
      {unitType === 'sheet'
        ? <SheetProductView key={`${props.sessionId}:sheet`} {...props} />
        : unitType === 'doc'
          ? <DocProductView key={`${props.sessionId}:doc`} {...props} />
          : <SlideProductView key={`${props.sessionId}:slide`} {...props} />}
    </section>
  )
}

export const inject = ['slots', 'connection']

export function apply(ctx: ClientContext): void {
  clientConnection.current = ctx.get('connection') as unknown as ConnectionHandle
  ctx.effect(() => () => {
    clientConnection.current = null
  }, 'dsh-univer-create: client connection')
  ctx.slots.inject('conversation.view', () => ctx.slots.register({
    name: 'conversation.view',
    id: 'univer-create',
    label: 'Univer',
    order: 20,
  }, UniverView))
}
