import React, { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import type { ClientContext, ConversationNode } from '@deepseek-ai/dsh-client-runtime/client'
import type { ConnectionHandle } from '@deepseek-ai/dsh-client-connection/client'
import type { ConvViewProps } from '@deepseek-ai/dsh-client-ui-conversation/client'
import { getDocsEmptySnapshot, LocaleType, mergeLocales, Univer } from '@univerjs/core'
import { FUniver } from '@univerjs/core/facade'
import { createUniver, defaultTheme } from '@univerjs/presets'
import { UniverDocsCorePreset } from '@univerjs/preset-docs-core'
import UniverPresetDocsCoreEnUS from '@univerjs/preset-docs-core/locales/en-US'
import UniverPresetDocsCoreZhCN from '@univerjs/preset-docs-core/locales/zh-CN'
import { UniverDocsDrawingPreset } from '@univerjs/preset-docs-drawing'
import UniverPresetDocsDrawingEnUS from '@univerjs/preset-docs-drawing/locales/en-US'
import UniverPresetDocsDrawingZhCN from '@univerjs/preset-docs-drawing/locales/zh-CN'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import { UniverDocsPlugin } from '@univerjs/docs'
import { UniverDocsUIPlugin } from '@univerjs/docs-ui'
import DocsUIEnUS from '@univerjs/docs-ui/locale/en-US'
import DocsUIZhCN from '@univerjs/docs-ui/locale/zh-CN'
import { UniverDrawingPlugin } from '@univerjs/drawing'
import { UniverRenderEnginePlugin } from '@univerjs/engine-render'
import { UniverUIPlugin } from '@univerjs/ui'
import UIEnUS from '@univerjs/ui/locale/en-US'
import UIZhCN from '@univerjs/ui/locale/zh-CN'
import DesignEnUS from '@univerjs/design/locale/en-US'
import DesignZhCN from '@univerjs/design/locale/zh-CN'
import ChartUIEnUS from '@univerjs-pro/chart-ui/locale/en-US'
import ChartUIZhCN from '@univerjs-pro/chart-ui/locale/zh-CN'
import { UniverDocsTablePlugin } from '@univerjs-pro/docs-table'
import { UniverLicensePlugin } from '@univerjs-pro/license'
import { UniverSheetsChartPlugin } from '@univerjs-pro/sheets-chart'
import { UniverSheetsChartUIPlugin } from '@univerjs-pro/sheets-chart-ui'
import SheetsChartUIEnUS from '@univerjs-pro/sheets-chart-ui/locale/en-US'
import SheetsChartUIZhCN from '@univerjs-pro/sheets-chart-ui/locale/zh-CN'
import { UNIVER_LICENSE } from 'virtual:dsh-univer-license'
import ShapeEditorUIEnUS from '@univerjs-pro/shape-editor-ui/locale/en-US'
import ShapeEditorUIZhCN from '@univerjs-pro/shape-editor-ui/locale/zh-CN'
import { getSlidesEmptySnapshot, PageElementTypeEnum, UniverSlidesPlugin } from '@univerjs-pro/slides'
import { UniverSlidesChartPlugin } from '@univerjs-pro/slides-chart'
import { UniverSlidesChartUIPlugin } from '@univerjs-pro/slides-chart-ui'
import SlidesChartUIEnUS from '@univerjs-pro/slides-chart-ui/locale/en-US'
import SlidesChartUIZhCN from '@univerjs-pro/slides-chart-ui/locale/zh-CN'
import { installScrollContainment } from './scroll-containment.js'
import type { ISlideData, ISlideTextElement } from '@univerjs-pro/slides'
import type { FPresentation } from '@univerjs-pro/slides/facade'
import { ShapeFillEnum, ShapeLineTypeEnum, ShapeTypeEnum } from '@univerjs-pro/engine-shape'
import '@univerjs-pro/docs-table/facade'
import '@univerjs-pro/engine-chart/facade'
import '@univerjs-pro/chart-ui/facade'
import '@univerjs-pro/sheets-chart/facade'
import '@univerjs-pro/slides/facade'
import '@univerjs-pro/slides-chart/facade'
import { UniverSlidesUIPlugin } from '@univerjs-pro/slides-ui'
import SlidesUIEnUS from '@univerjs-pro/slides-ui/locale/en-US'
import SlidesUIZhCN from '@univerjs-pro/slides-ui/locale/zh-CN'
import { useDshLanguage } from '../../../dsh-language.js'
import { UNIVER_CREATE_LOCALES } from './locales.js'

import '@univerjs/preset-docs-core/lib/index.css'
import '@univerjs/preset-docs-drawing/lib/index.css'
import '@univerjs/preset-sheets-core/lib/index.css'
import '@univerjs-pro/chart-ui/lib/index.css'
import '@univerjs-pro/sheets-chart-ui/lib/index.css'
import '@univerjs-pro/slides-chart-ui/lib/index.css'
import '@univerjs/design/lib/index.css'
import '@univerjs/ui/lib/index.css'
import '@univerjs/docs-ui/lib/index.css'
import '@univerjs-pro/shape-editor-ui/lib/index.css'
import '@univerjs-pro/slides-ui/lib/index.css'
import './styles.css'

function univerLocale(language: ReturnType<typeof useDshLanguage>): LocaleType {
  return language === 'zh' ? LocaleType.ZH_CN : LocaleType.EN_US
}

const SHEET_LOCALES = {
  [LocaleType.EN_US]: mergeLocales(UniverPresetSheetsCoreEnUS, ChartUIEnUS, SheetsChartUIEnUS),
  [LocaleType.ZH_CN]: mergeLocales(UniverPresetSheetsCoreZhCN, ChartUIZhCN, SheetsChartUIZhCN),
}

const DOC_LOCALES = {
  [LocaleType.EN_US]: mergeLocales(UniverPresetDocsCoreEnUS, UniverPresetDocsDrawingEnUS),
  [LocaleType.ZH_CN]: mergeLocales(UniverPresetDocsCoreZhCN, UniverPresetDocsDrawingZhCN),
}

const SLIDE_LOCALES = {
  [LocaleType.EN_US]: mergeLocales(DesignEnUS, UIEnUS, DocsUIEnUS, ShapeEditorUIEnUS, SlidesUIEnUS, ChartUIEnUS, SlidesChartUIEnUS),
  [LocaleType.ZH_CN]: mergeLocales(DesignZhCN, UIZhCN, DocsUIZhCN, ShapeEditorUIZhCN, SlidesUIZhCN, ChartUIZhCN, SlidesChartUIZhCN),
}

type JsonScalar = string | number | boolean | null

type SheetOperationId = number | string

type SheetOperation =
  | { seq: SheetOperationId; action: 'new'; title?: string; sheetName?: string; rows?: number; columns?: number }
  | { seq: SheetOperationId; action: 'add-sheet'; name: string; rows?: number; columns?: number }
  | { seq: SheetOperationId; action: 'delete-sheet'; name: string }
  | { seq: SheetOperationId; action: 'rename-sheet'; oldName: string; newName: string }
  | { seq: SheetOperationId; action: 'list-sheets' }
  | { seq: SheetOperationId; action: 'set-range'; range: string; values: JsonScalar[][]; sheetName?: string }
  | { seq: SheetOperationId; action: 'clear-range'; range: string; sheetName?: string }
  | {
      seq: SheetOperationId
      action: 'format-range'
      range: string
      sheetName?: string
      background?: string
      fontColor?: string
      bold?: boolean
      fontSize?: number
    }

type DocOperation =
  | { seq: SheetOperationId; action: 'new-doc'; title?: string; text?: string }
  | { seq: SheetOperationId; action: 'get-doc-text' }
  | { seq: SheetOperationId; action: 'set-doc-text'; text: string }
  | { seq: SheetOperationId; action: 'insert-doc-text'; index: number; text: string }
  | { seq: SheetOperationId; action: 'append-doc-text'; text: string }
  | { seq: SheetOperationId; action: 'delete-doc-range'; start: number; end: number }
  | {
      seq: SheetOperationId
      action: 'format-doc-text'
      start: number
      end: number
      bold?: boolean
      italic?: boolean
      fontSize?: number
      fontColor?: string
    }

type SlideOperation =
  | { seq: SheetOperationId; action: 'new-slide'; title?: string; width?: number; height?: number }
  | { seq: SheetOperationId; action: 'list-slides' }
  | { seq: SheetOperationId; action: 'add-slide'; title?: string; index?: number }
  | { seq: SheetOperationId; action: 'delete-slide'; index: number }
  | {
      seq: SheetOperationId
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
      seq: SheetOperationId
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
  | { seq: SheetOperationId; action: 'delete-slide-element'; slideIndex: number; elementId: string }
  | {
      seq: SheetOperationId
      action: 'add-slide-shape'
      slideIndex: number
      shapeType: string
      left?: number
      top?: number
      width?: number
      height?: number
      fillColor?: string
      strokeColor?: string
      strokeWidth?: number
      opacity?: number
      selectable?: boolean
    }

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

function sheetOperationFromArgs(name: string, args: Record<string, unknown>, seq: SheetOperationId): SheetOperation | null {
  const sheetName = typeof args.sheetName === 'string' ? args.sheetName : undefined
  if (name === 'univer_sheet_new' || name === 'new') {
    return {
      seq,
      action: 'new',
      title: typeof args.title === 'string' ? args.title : undefined,
      sheetName,
      rows: typeof args.rows === 'number' ? args.rows : undefined,
      columns: typeof args.columns === 'number' ? args.columns : undefined,
    }
  }
  if (name === 'univer_sheet_add' || name === 'add-sheet') {
    if (typeof args.name !== 'string') return null
    return {
      seq,
      action: 'add-sheet',
      name: args.name,
      rows: typeof args.rows === 'number' ? args.rows : undefined,
      columns: typeof args.columns === 'number' ? args.columns : undefined,
    }
  }
  if (name === 'univer_sheet_delete' || name === 'delete-sheet') {
    return typeof args.name === 'string' ? { seq, action: 'delete-sheet', name: args.name } : null
  }
  if (name === 'univer_sheet_rename' || name === 'rename-sheet') {
    return typeof args.oldName === 'string' && typeof args.newName === 'string'
      ? { seq, action: 'rename-sheet', oldName: args.oldName, newName: args.newName }
      : null
  }
  if (name === 'univer_sheet_list' || name === 'list-sheets') return { seq, action: 'list-sheets' }
  if (typeof args.range !== 'string') return null
  if (name === 'univer_sheet_set_range' || name === 'set-range') {
    if (!Array.isArray(args.values)) return null
    return { seq, action: 'set-range', range: args.range, values: args.values as JsonScalar[][], sheetName }
  }
  if (name === 'univer_sheet_clear_range' || name === 'clear-range') {
    return { seq, action: 'clear-range', range: args.range, sheetName }
  }
  if (name !== 'univer_sheet_format_range' && name !== 'format-range') return null
  return {
    seq,
    action: 'format-range',
    range: args.range,
    sheetName,
    background: typeof args.background === 'string' ? args.background : undefined,
    fontColor: typeof args.fontColor === 'string' ? args.fontColor : undefined,
    bold: typeof args.bold === 'boolean' ? args.bold : undefined,
    fontSize: typeof args.fontSize === 'number' ? args.fontSize : undefined,
  }
}

function operationFromNode(node: ConversationNode): SheetOperation | null {
  if (node.kind !== 'tool-result' || node.isError || node.call === null || !SHEET_TOOL_NAMES.has(node.call.name)) return null
  try {
    const parsed = JSON.parse(node.call.argsRaw)
    if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) return null
    return sheetOperationFromArgs(node.call.name, parsed as Record<string, unknown>, node.seq)
  } catch {
    return null
  }
}

interface QueuedUniverOperation {
  id: string
  unitType: UniverUnitType
  operation: Record<string, unknown>
  createdAt?: number
}

interface ProductViewProps extends ConvViewProps {
  queuedOperations: QueuedUniverOperation[]
  codeRequests: UniverCodeRequest[]
  sheetScreenshotRequests: SheetScreenshotRequest[]
  docScreenshotRequests: DocScreenshotRequest[]
  slideScreenshotRequests: SlideScreenshotRequest[]
}

interface SlideScreenshotRequest {
  id: string
  slideIndex: number
  mode: 'slide' | 'editor'
}

interface SheetScreenshotRequest {
  id: string
  mode: 'sheet' | 'editor'
  scroll: 'none' | 'up' | 'down' | 'left' | 'right' | 'top' | 'bottom' | 'start' | 'end'
  amount?: number
}

interface DocScreenshotRequest {
  id: string
  mode: 'document' | 'editor'
  scroll: 'none' | 'up' | 'down' | 'top' | 'bottom'
  amount?: number
}

interface UniverCodeRequest {
  id: string
  unitType: UniverUnitType
  code: string
}

interface UniverTasks {
  operations: QueuedUniverOperation[]
  codeRequests: UniverCodeRequest[]
  sheetScreenshotRequests: SheetScreenshotRequest[]
  docScreenshotRequests: DocScreenshotRequest[]
  slideScreenshotRequests: SlideScreenshotRequest[]
}

const EMPTY_UNIVER_TASKS: UniverTasks = {
  operations: [],
  codeRequests: [],
  sheetScreenshotRequests: [],
  docScreenshotRequests: [],
  slideScreenshotRequests: [],
}

interface UniverCodeRuntime {
  univerAPI: FUniver
  save: () => unknown
}

function parseQueuedUniverOperation(value: unknown): QueuedUniverOperation | null {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return null
  const queued = value as Partial<QueuedUniverOperation>
  if (
    typeof queued.id !== 'string'
    || (queued.unitType !== 'sheet' && queued.unitType !== 'doc' && queued.unitType !== 'slide')
    || queued.operation === null
    || typeof queued.operation !== 'object'
    || Array.isArray(queued.operation)
    || typeof queued.operation.action !== 'string'
  ) return null
  return {
    id: queued.id,
    unitType: queued.unitType,
    operation: queued.operation,
    ...(typeof queued.createdAt === 'number' ? { createdAt: queued.createdAt } : {}),
  }
}

function parseUniverCodeRequest(value: unknown): UniverCodeRequest | null {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return null
  const request = value as Partial<UniverCodeRequest>
  if (
    typeof request.id !== 'string'
    || (request.unitType !== 'sheet' && request.unitType !== 'doc' && request.unitType !== 'slide')
    || typeof request.code !== 'string'
  ) return null
  return { id: request.id, unitType: request.unitType, code: request.code }
}

function parseSheetScreenshotRequest(value: unknown): SheetScreenshotRequest | null {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return null
  const request = value as Partial<SheetScreenshotRequest>
  const validScroll = request.scroll === 'none' || request.scroll === 'up' || request.scroll === 'down' || request.scroll === 'left' || request.scroll === 'right' || request.scroll === 'top' || request.scroll === 'bottom' || request.scroll === 'start' || request.scroll === 'end'
  return typeof request.id === 'string' && (request.mode === 'sheet' || request.mode === 'editor') && validScroll
    ? { id: request.id, mode: request.mode, scroll: request.scroll!, ...(typeof request.amount === 'number' ? { amount: request.amount } : {}) }
    : null
}

function parseDocScreenshotRequest(value: unknown): DocScreenshotRequest | null {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return null
  const request = value as Partial<DocScreenshotRequest>
  const validScroll = request.scroll === 'none' || request.scroll === 'up' || request.scroll === 'down' || request.scroll === 'top' || request.scroll === 'bottom'
  return typeof request.id === 'string' && (request.mode === 'document' || request.mode === 'editor') && validScroll
    ? { id: request.id, mode: request.mode, scroll: request.scroll!, ...(typeof request.amount === 'number' ? { amount: request.amount } : {}) }
    : null
}

function parseSlideScreenshotRequest(value: unknown): SlideScreenshotRequest | null {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return null
  const request = value as Partial<SlideScreenshotRequest>
  return typeof request.id === 'string'
    && typeof request.slideIndex === 'number'
    && (request.mode === 'slide' || request.mode === 'editor')
    ? { id: request.id, slideIndex: request.slideIndex, mode: request.mode }
    : null
}

function queuedSheetOperation(queued: QueuedUniverOperation): SheetOperation | null {
  return queued.unitType === 'sheet'
    ? sheetOperationFromArgs(String(queued.operation.action), queued.operation, queued.id)
    : null
}

const DOC_OPERATION_ACTIONS = new Set(['new-doc', 'set-doc-text', 'insert-doc-text', 'append-doc-text', 'delete-doc-range', 'format-doc-text'])
const SLIDE_OPERATION_ACTIONS = new Set(['new-slide', 'add-slide', 'delete-slide', 'add-slide-text', 'add-slide-shape', 'update-slide-text', 'delete-slide-element'])

function queuedDocOperation(queued: QueuedUniverOperation): DocOperation | null {
  if (queued.unitType !== 'doc' || !DOC_OPERATION_ACTIONS.has(String(queued.operation.action))) return null
  return { ...queued.operation, seq: queued.id } as DocOperation
}

function queuedSlideOperation(queued: QueuedUniverOperation): SlideOperation | null {
  if (queued.unitType !== 'slide' || !SLIDE_OPERATION_ACTIONS.has(String(queued.operation.action))) return null
  return { ...queued.operation, seq: queued.id } as SlideOperation
}

async function commitOperations(
  sessionId: string,
  unitType: UniverUnitType,
  snapshot: unknown,
  operationIds: string[],
): Promise<void> {
  if (operationIds.length === 0) return
  const connection = clientConnection.current
  if (connection === null) throw new Error('Host 连接不可用')
  const result = await connection.rpc.call('/dsh-univer-create', 'commit-operations', {
    sessionId,
    unitType,
    snapshot,
    operationIds,
    clientId: univerCodeClientId,
  })
  if (!result.ok) throw new Error(result.error.message)
}

const DOC_TOOL_NAMES = new Set([
  'univer_doc_new',
  'univer_doc_get_text',
  'univer_doc_set_text',
  'univer_doc_insert_text',
  'univer_doc_append_text',
  'univer_doc_delete_range',
  'univer_doc_format_text',
  'univer_doc_screenshot',
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
  'univer_slide_add_shape',
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
  if (node.call.name === 'univer_slide_add_shape') {
    if (typeof args.shapeType !== 'string') return null
    return {
      seq: node.seq,
      action: 'add-slide-shape',
      slideIndex: args.slideIndex,
      shapeType: args.shapeType,
      left: numberArg('left'),
      top: numberArg('top'),
      width: numberArg('width'),
      height: numberArg('height'),
      fillColor: typeof args.fillColor === 'string' ? args.fillColor : undefined,
      strokeColor: typeof args.strokeColor === 'string' ? args.strokeColor : undefined,
      strokeWidth: numberArg('strokeWidth'),
      opacity: numberArg('opacity'),
      selectable: typeof args.selectable === 'boolean' ? args.selectable : undefined,
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

function safeOfficeFilename(title: string, extension: SavePathDialogProps['extension'], untitled = 'Untitled'): string {
  const stem = title.trim().replace(/[\\/:*?"<>|]+/g, '_') || untitled
  return `${stem}.${extension}`
}

function openedFilePath(file: File): string {
  const nativePath = (file as File & { path?: string }).path
  return typeof nativePath === 'string' && nativePath.length > 0
    ? nativePath
    : file.webkitRelativePath || file.name
}

function SavePathDialog({ extension, suggestedName, busy, onCancel, onSave }: SavePathDialogProps) {
  const language = useDshLanguage()
  const ui = UNIVER_CREATE_LOCALES[language]
  const [path, setPath] = useState(() => safeOfficeFilename(suggestedName, extension, ui.untitledFile))
  return (
    <div className="dsh-univer-create-save-backdrop" role="presentation" onMouseDown={(event) => {
      if (event.target === event.currentTarget && !busy) onCancel()
    }}>
      <form className="dsh-univer-create-save-dialog" role="dialog" aria-modal="true" aria-labelledby="dsh-univer-save-title" onSubmit={(event) => {
        event.preventDefault()
        const value = path.trim()
        if (value.length > 0) onSave(value)
      }}>
        <h2 id="dsh-univer-save-title">{ui.saveDialogTitle}</h2>
        <p>{ui.saveDialogDescription}</p>
        <label>
          {ui.filePath}
          <input autoFocus value={path} disabled={busy} onChange={(event) => setPath(event.target.value)} placeholder={`${ui.example}: ${safeOfficeFilename(suggestedName, extension, ui.untitledFile)}`} />
        </label>
        <div className="dsh-univer-create-save-actions">
          <button className="dsh-univer-create-action" type="button" disabled={busy} onClick={onCancel}>{ui.cancel}</button>
          <button className="dsh-univer-create-action dsh-univer-create-action--primary" type="submit" disabled={busy || path.trim().length === 0}>
            {busy ? ui.saving : ui.save}
          </button>
        </div>
      </form>
    </div>
  )
}

type UnitFileState = { path: string | null; snapshot: unknown | null }

async function loadUnitFileState(sessionId: string, unitType: UniverUnitType): Promise<UnitFileState> {
  const connection = clientConnection.current
  if (connection === null) throw new Error('Host 连接不可用')
  const result = await connection.rpc.call('/dsh-univer-create', 'file-state', { sessionId, unitType })
  if (!result.ok) throw new Error(result.error.message)
  const value = result.value as { path?: unknown; snapshot?: unknown }
  if (value === null || typeof value !== 'object') throw new Error('Host 返回了无效的文件状态')
  return {
    path: typeof value.path === 'string' ? value.path : null,
    snapshot: value.snapshot ?? null,
  }
}

class UniverExportError extends Error {
  constructor(readonly code: string, message: string) {
    super(message)
    this.name = 'UniverExportError'
  }
}

function isWorkspaceFileExistsError(reason: unknown): boolean {
  return reason instanceof UniverExportError
    ? reason.code === 'already-exists'
    : reason instanceof Error && reason.message.includes('目标文件已存在')
}

const UNKNOWN_FILE_SNAPSHOT = '\u0000unknown-file-snapshot'

function confirmOpenWithUnsavedChanges(snapshot: unknown, baseline: string | null, message: string): boolean {
  return JSON.stringify(snapshot) === baseline || window.confirm(message)
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
  if (!result.ok) throw new UniverExportError(result.error.code, result.error.message)
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

function workbookData(operation: Extract<SheetOperation, { action: 'new' }> | undefined, defaultTitle: string) {
  const sheetId = 'sheet-1'
  return {
    id: `dsh-univer-${operation?.seq ?? 'default'}`,
    name: operation?.title ?? defaultTitle,
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

function documentData(operation: Extract<DocOperation, { action: 'new-doc' }> | undefined, locale: LocaleType, defaultTitle: string) {
  return getDocsEmptySnapshot(
    `dsh-univer-doc-${operation?.seq ?? 'default'}`,
    locale,
    operation?.title ?? defaultTitle,
  )
}

function presentationData(operation: Extract<SlideOperation, { action: 'new-slide' }> | undefined, locale: LocaleType, defaultTitle: string): ISlideData {
  const data = getSlidesEmptySnapshot(
    `dsh-univer-slide-${operation?.seq ?? 'default'}`,
    locale,
    operation?.title ?? defaultTitle,
  )
  data.defaultPageSize = {
    ...data.defaultPageSize,
    width: Math.max(1, operation?.width ?? 960),
    height: Math.max(1, operation?.height ?? 540),
  }
  // The Univer empty snapshot defaults to the title-and-body layout. That layout
  // renders untranslated placeholder keys and overlaps AI-positioned text. AI
  // slide tools use absolute model coordinates, so start them from a truly blank
  // layout while leaving imported presentations unchanged.
  const firstSlideId = data.slideOrder[0]
  if (firstSlideId !== undefined && data.slides[firstSlideId] !== undefined) {
    data.slides[firstSlideId].layoutPageId = 'layout-blank'
    data.slides[firstSlideId].showMasterSp = false
  }
  return data
}

function normalizeAiSlideLayouts(data: ISlideData): ISlideData {
  for (const slideId of data.slideOrder) {
    const slide = data.slides[slideId]
    if (slide === undefined) continue
    const elementIds = Object.keys(slide.elements ?? {})
    for (const elementId of elementIds) {
      const element = slide.elements[elementId]
      if (element !== undefined && (elementId.startsWith('slide-text-') || element.name === 'text')) element.selectable = true
    }
    if (elementIds.some((id) => id.startsWith('slide-text-')) && !elementIds.some((id) => id.startsWith('ph-'))) {
      slide.layoutPageId = 'layout-blank'
      slide.showMasterSp = false
    }
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

function conversationLabel(node: ConversationNode, userLabel: string): { role: string; text: string } | null {
  if (node.kind === 'user' || node.kind === 'steering') {
    const text = textFromBlocks(node.content)
    return text ? { role: userLabel, text } : null
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
const univerCodeClientId = crypto.randomUUID()
const UNIVER_CODE_DEBUG_FLAG = '__DSH_UNIVER_CODE_DEBUG__'

type ResidentMountedRuntime = MountedRuntime | MountedSlideRuntime

type ResidentRuntimeEntry = {
  sessionId: string
  unitType: UniverUnitType
  runtime: ResidentMountedRuntime
  save: () => unknown
  savingRef: React.MutableRefObject<boolean>
  pendingSaveRef: React.MutableRefObject<Promise<void> | null>
  lastSavedRef: React.MutableRefObject<string | null>
  lastFileSnapshotRef: React.MutableRefObject<string | null>
}

const residentRuntimes = new Map<string, ResidentRuntimeEntry>()
const residentViewShells = new Map<string, HTMLDivElement>()
let residentRuntimeParkingHost: HTMLDivElement | null = null
let currentResidentSessionId: string | null = null

function residentRuntimeKey(sessionId: string, unitType: UniverUnitType): string {
  return `${sessionId}:${unitType}`
}

function getResidentRuntimeParkingHost(): HTMLDivElement {
  if (residentRuntimeParkingHost?.isConnected) return residentRuntimeParkingHost
  const host = document.createElement('div')
  host.className = 'dsh-univer-create-runtime-parking'
  host.setAttribute('aria-hidden', 'true')
  document.body.appendChild(host)
  residentRuntimeParkingHost = host
  return host
}

function getResidentViewShell(sessionId: string): HTMLDivElement {
  const cached = residentViewShells.get(sessionId)
  if (cached !== undefined) return cached
  const shell = document.createElement('div')
  shell.className = 'dsh-univer-create-resident-shell'
  shell.dataset.sessionId = sessionId
  getResidentRuntimeParkingHost().appendChild(shell)
  residentViewShells.set(sessionId, shell)
  return shell
}

function getResidentRuntime<Runtime extends ResidentMountedRuntime>(sessionId: string, unitType: UniverUnitType): Runtime | null {
  return residentRuntimes.get(residentRuntimeKey(sessionId, unitType))?.runtime as Runtime | undefined ?? null
}

function registerResidentRuntime(
  sessionId: string,
  unitType: UniverUnitType,
  runtime: ResidentMountedRuntime,
  save: () => unknown,
  savingRef: React.MutableRefObject<boolean>,
  pendingSaveRef: React.MutableRefObject<Promise<void> | null>,
  lastSavedRef: React.MutableRefObject<string | null>,
  lastFileSnapshotRef: React.MutableRefObject<string | null>,
): void {
  const key = residentRuntimeKey(sessionId, unitType)
  const previous = residentRuntimes.get(key)
  if (previous !== undefined && previous.runtime !== runtime) {
    previous.runtime.univer.dispose()
    previous.runtime.mount.remove()
  } else if (previous !== undefined) {
    savingRef.current = previous.savingRef.current
    pendingSaveRef.current = previous.pendingSaveRef.current
    lastSavedRef.current = previous.lastSavedRef.current
    lastFileSnapshotRef.current = previous.lastFileSnapshotRef.current
  }
  residentRuntimes.set(key, { sessionId, unitType, runtime, save, savingRef, pendingSaveRef, lastSavedRef, lastFileSnapshotRef })
}

function disposeResidentRuntime(sessionId: string, unitType: UniverUnitType, runtime: ResidentMountedRuntime | null): void {
  if (runtime === null) return
  const key = residentRuntimeKey(sessionId, unitType)
  if (residentRuntimes.get(key)?.runtime === runtime) residentRuntimes.delete(key)
  runtime.univer.dispose()
  runtime.mount.remove()
}

function parkResidentRuntime(sessionId: string, unitType: UniverUnitType, runtime: ResidentMountedRuntime | null): void {
  if (runtime === null || residentRuntimes.get(residentRuntimeKey(sessionId, unitType))?.runtime !== runtime) return
  getResidentRuntimeParkingHost().appendChild(runtime.mount)
}

function debugUniverCode(sessionId: string, unitType: UniverUnitType, requestId: string, code: string): void {
  if ((globalThis as Record<string, unknown>)[UNIVER_CODE_DEBUG_FLAG] !== true) return
  console.groupCollapsed(`[dsh-univer] execute ${unitType} code (${requestId})`)
  console.debug({ sessionId, unitType, requestId })
  console.log(code)
  console.groupEnd()
}

const FORBIDDEN_UNIVER_CODE = [
  { pattern: /\b(?:window|globalThis|self|top|parent|opener|frames|document|location|navigator|localStorage|sessionStorage|indexedDB|caches)\s*(?:\.|\[)/, reason: 'browser globals and network/storage APIs are not allowed' },
  { pattern: /\b(?:fetch|XMLHttpRequest|WebSocket|Worker|SharedWorker|EventSource|setTimeout|setInterval|requestAnimationFrame|alert|confirm|prompt)\s*\(/, reason: 'browser networking, timers, and dialogs are not allowed' },
  { pattern: /\b(?:eval|Function|require)\s*\(/, reason: 'dynamic code loading is not allowed' },
  { pattern: /\bimport\s*(?:\(|["'])/, reason: 'imports are not available in the browser executor' },
  { pattern: /(?:\.|\[\s*["'])(?:constructor|__proto__|prototype)(?:\b|["'])/, reason: 'prototype and constructor access is not allowed' },
  { pattern: /\.(?:getInjector|getUniver|getPresentation|getWorkbook|getDocument)\s*\(/, reason: 'raw Univer internals are not allowed; use Facade methods only' },
] as const

function serializeCodeValue(value: unknown, maxLength = 50_000): string {
  const seen = new WeakSet<object>()
  let remainingNodes = 2_000
  const sanitize = (item: unknown, depth: number): unknown => {
    if (remainingNodes-- <= 0) return '[Truncated]'
    if (item === null || typeof item === 'string' || typeof item === 'number' || typeof item === 'boolean') return item
    if (typeof item === 'undefined') return 'undefined'
    if (typeof item === 'bigint') return `${item}n`
    if (typeof item === 'function') return `[Function ${item.name || 'anonymous'}]`
    if (typeof item === 'symbol') return String(item)
    if (typeof item !== 'object') return String(item)
    if (seen.has(item)) return '[Circular]'
    if (depth >= 8) return `[Object ${(item as { constructor?: { name?: string } }).constructor?.name ?? 'unknown'}]`
    seen.add(item)
    if (Array.isArray(item)) return item.slice(0, 100).map((entry) => sanitize(entry, depth + 1))
    const prototype = Object.getPrototypeOf(item)
    if (prototype !== Object.prototype && prototype !== null) {
      return `[Object ${(item as { constructor?: { name?: string } }).constructor?.name ?? 'unknown'}]`
    }
    const output: Record<string, unknown> = {}
    for (const key of Object.keys(item).slice(0, 100)) output[key] = sanitize((item as Record<string, unknown>)[key], depth + 1)
    return output
  }

  let serialized: string
  try {
    serialized = JSON.stringify(sanitize(value, 0), null, 2) ?? String(value)
  } catch {
    serialized = String(value)
  }
  return serialized.length > maxLength ? `${serialized.slice(0, maxLength)}\n…[truncated]` : serialized
}

async function executeFacadeCode(code: string, univerAPI: FUniver): Promise<{ result?: string; logs: string[]; error?: string }> {
  const blocked = FORBIDDEN_UNIVER_CODE.find(({ pattern }) => pattern.test(code))
  if (blocked !== undefined) return { error: blocked.reason, logs: [] }

  const logs: string[] = []
  const appendLog = (level: string, values: unknown[]) => {
    if (logs.length >= 200) return
    logs.push(`[${level}] ${values.map((value) => serializeCodeValue(value, 2_000)).join(' ')}`)
  }
  const safeConsole = Object.freeze({
    log: (...values: unknown[]) => appendLog('log', values),
    info: (...values: unknown[]) => appendLog('info', values),
    warn: (...values: unknown[]) => appendLog('warn', values),
    error: (...values: unknown[]) => appendLog('error', values),
  })

  try {
    const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor as new (...args: string[]) => (...values: unknown[]) => Promise<unknown>
    const run = new AsyncFunction('univerAPI', 'console', `"use strict";\n${code}\n//# sourceURL=dsh-univer-execute-code.js`)
    let timer = 0
    try {
      const result = await Promise.race([
        run(univerAPI, safeConsole),
        new Promise<never>((_resolve, reject) => {
          timer = window.setTimeout(() => reject(new Error('code exceeded the 10 second asynchronous timeout')), 10_000)
        }),
      ])
      return { result: serializeCodeValue(result), logs }
    } finally {
      window.clearTimeout(timer)
    }
  } catch (reason) {
    return { error: reason instanceof Error ? `${reason.name}: ${reason.message}` : String(reason), logs }
  }
}

const residentCodeRequestsInFlight = new Set<string>()
const residentRuntimeTaskQueues = new Map<string, Promise<void>>()

function disposeInactiveParkedRuntimes(): void {
  for (const entry of [...residentRuntimes.values()]) {
    const key = residentRuntimeKey(entry.sessionId, entry.unitType)
    const inactive = currentResidentSessionId === null || entry.sessionId !== currentResidentSessionId
    if (inactive && entry.runtime.mount.parentElement === residentRuntimeParkingHost && !residentRuntimeTaskQueues.has(key)) {
      disposeResidentRuntime(entry.sessionId, entry.unitType, entry.runtime)
    }
  }
}

function enqueueResidentRuntimeTask(key: string, task: () => Promise<void>): Promise<void> {
  const previous = residentRuntimeTaskQueues.get(key) ?? Promise.resolve()
  const queued = previous.catch(() => {}).then(task)
  residentRuntimeTaskQueues.set(key, queued)
  void queued.finally(() => {
    if (residentRuntimeTaskQueues.get(key) === queued) residentRuntimeTaskQueues.delete(key)
    disposeInactiveParkedRuntimes()
  }).catch(() => {})
  return queued
}

function queueResidentFinalSave(entry: ResidentRuntimeEntry, task: () => Promise<unknown>): void {
  const key = residentRuntimeKey(entry.sessionId, entry.unitType)
  const previousSave = entry.pendingSaveRef.current
  entry.savingRef.current = true
  const queued = enqueueResidentRuntimeTask(key, async () => {
    await previousSave?.catch(() => {})
    await task()
  })
  entry.pendingSaveRef.current = queued
  void queued.finally(() => {
    if (entry.pendingSaveRef.current === queued) {
      entry.pendingSaveRef.current = null
      entry.savingRef.current = false
    }
    const current = residentRuntimes.get(key)
    if (current?.pendingSaveRef.current === queued) {
      current.pendingSaveRef.current = null
      current.savingRef.current = false
    }
  }).catch(() => {})
}

async function executeResidentCodeRequest(sessionId: string, request: UniverCodeRequest): Promise<void> {
  const connection = clientConnection.current
  const key = residentRuntimeKey(sessionId, request.unitType)
  const entry = residentRuntimes.get(key)
  if (connection === null || entry === undefined || residentCodeRequestsInFlight.has(request.id)) return
  residentCodeRequestsInFlight.add(request.id)
  const pendingSave = entry.pendingSaveRef.current
  const queued = enqueueResidentRuntimeTask(key, async () => {
    if (currentResidentSessionId !== sessionId || residentRuntimes.get(key) !== entry) return
    const claimed = await connection.rpc.call('/dsh-univer-create', 'univer-code-claim', {
      sessionId,
      codeRequestId: request.id,
      clientId: univerCodeClientId,
    })
    if (!claimed.ok || claimed.value === null) return
    let execution: { result?: string; logs: string[]; error?: string } = { logs: [] }
    try {
      await pendingSave
      if (residentRuntimes.get(key) !== entry) throw new Error(`当前浏览器没有打开已创建的 Univer ${request.unitType}`)
      entry.savingRef.current = true
      debugUniverCode(sessionId, request.unitType, request.id, request.code)
      execution = await executeFacadeCode(request.code, entry.runtime.univerAPI)
      if (execution.error === undefined) {
        const snapshot = entry.save()
        if (snapshot === undefined || snapshot === null) throw new Error(`当前没有活动的 Univer ${request.unitType}`)
        const saved = await connection.rpc.call('/dsh-univer-create', 'save', { sessionId, unitType: request.unitType, snapshot })
        if (!saved.ok) throw new Error(saved.error.message)
        entry.lastSavedRef.current = JSON.stringify(snapshot)
      }
    } catch (reason) {
      execution = {
        error: `${reason instanceof Error ? reason.message : String(reason)}. Code execution is not transactional; inspect the document for partial changes.`,
        logs: execution.logs,
      }
    } finally {
      entry.savingRef.current = false
    }
    await connection.rpc.call('/dsh-univer-create', 'univer-code-result', {
      sessionId,
      codeRequestId: request.id,
      clientId: univerCodeClientId,
      ...execution,
    }).catch(() => {})
  })
  try {
    await queued
  } finally {
    residentCodeRequestsInFlight.delete(request.id)
  }
}

function UniverSessionRuntimeTracker(props: ConvViewProps) {
  const shell = useMemo(() => getResidentViewShell(props.sessionId), [props.sessionId])
  useLayoutEffect(() => {
    if (!shell.isConnected) getResidentRuntimeParkingHost().appendChild(shell)
  }, [shell])
  useEffect(() => {
    currentResidentSessionId = props.sessionId
    return () => {
      if (currentResidentSessionId === props.sessionId) {
        currentResidentSessionId = null
        disposeInactiveParkedRuntimes()
      }
      window.setTimeout(() => {
        if (residentViewShells.get(props.sessionId) !== shell || currentResidentSessionId === props.sessionId) return
        residentViewShells.delete(props.sessionId)
        shell.remove()
      }, 0)
    }
  }, [props.sessionId, shell])
  return createPortal(<UniverView {...props} />, shell)
}

// Browser-lifetime markers control only the initial loading presentation.
// Persisted Host state, not these markers, decides whether a unit exists.
async function captureRenderedSheet(
  runtime: MountedRuntime,
  request: SheetScreenshotRequest,
): Promise<{ dataUrl: string; width: number; height: number; scrollTop: number; scrollHeight: number; viewportHeight: number; scrollLeft: number; scrollWidth: number; viewportWidth: number }> {
  const elements = [runtime.mount, ...runtime.mount.querySelectorAll<HTMLElement>('*')]
  const verticalScroller = elements.filter((element) => element.clientHeight >= 120 && element.scrollHeight > element.clientHeight + 8)
    .sort((left, right) => (right.scrollHeight - right.clientHeight) - (left.scrollHeight - left.clientHeight))[0] ?? null
  const horizontalScroller = elements.filter((element) => element.clientWidth >= 200 && element.scrollWidth > element.clientWidth + 8)
    .sort((left, right) => (right.scrollWidth - right.clientWidth) - (left.scrollWidth - left.clientWidth))[0] ?? null
  const canvas = [...runtime.mount.querySelectorAll('canvas')]
    .sort((left, right) => right.getBoundingClientRect().width * right.getBoundingClientRect().height - left.getBoundingClientRect().width * left.getBoundingClientRect().height)[0]
  const verticalAmount = Math.max(1, request.amount ?? (verticalScroller?.clientHeight ?? runtime.mount.clientHeight) * 0.8)
  const horizontalAmount = Math.max(1, request.amount ?? (horizontalScroller?.clientWidth ?? runtime.mount.clientWidth) * 0.8)
  let deltaX = 0
  let deltaY = 0
  if (request.scroll === 'up') deltaY = -verticalAmount
  else if (request.scroll === 'down') deltaY = verticalAmount
  else if (request.scroll === 'left') deltaX = -horizontalAmount
  else if (request.scroll === 'right') deltaX = horizontalAmount
  if (request.scroll === 'top') verticalScroller?.scrollTo({ top: 0, behavior: 'auto' })
  else if (request.scroll === 'bottom') verticalScroller?.scrollTo({ top: verticalScroller.scrollHeight, behavior: 'auto' })
  else if (request.scroll === 'start') horizontalScroller?.scrollTo({ left: 0, behavior: 'auto' })
  else if (request.scroll === 'end') horizontalScroller?.scrollTo({ left: horizontalScroller.scrollWidth, behavior: 'auto' })
  else {
    if (deltaY !== 0 && verticalScroller !== null) verticalScroller.scrollBy({ top: deltaY, behavior: 'auto' })
    if (deltaX !== 0 && horizontalScroller !== null) horizontalScroller.scrollBy({ left: deltaX, behavior: 'auto' })
  }
  if ((request.scroll !== 'none') && ((deltaY !== 0 && verticalScroller === null) || (deltaX !== 0 && horizontalScroller === null) || ((request.scroll === 'top' || request.scroll === 'bottom') && verticalScroller === null) || ((request.scroll === 'start' || request.scroll === 'end') && horizontalScroller === null))) {
    canvas?.dispatchEvent(new WheelEvent('wheel', { bubbles: true, cancelable: true, deltaX, deltaY }))
  }
  window.dispatchEvent(new Event('resize'))
  if (document.fonts?.ready !== undefined) await document.fonts.ready
  await new Promise<void>((resolve) => window.setTimeout(resolve, 180))
  await waitForAnimationFrames(3)

  const mountRect = runtime.mount.getBoundingClientRect()
  const layers = [...runtime.mount.querySelectorAll('canvas')].map((layer) => ({ layer, rect: layer.getBoundingClientRect() }))
    .filter(({ layer, rect }) => layer.width >= 20 && layer.height >= 20 && rect.width > 0 && rect.height > 0)
  if (layers.length === 0) throw new Error('没有找到已渲染的工作表画布')
  const outputWidth = Math.max(1, Math.min(1600, Math.round(mountRect.width * Math.min(window.devicePixelRatio || 1, 2))))
  const outputHeight = Math.max(1, Math.round(outputWidth * mountRect.height / mountRect.width))
  const outputCanvas = document.createElement('canvas')
  outputCanvas.width = outputWidth
  outputCanvas.height = outputHeight
  const context = outputCanvas.getContext('2d')
  if (context === null) throw new Error('浏览器无法创建工作表截图画布')
  context.fillStyle = '#ffffff'
  context.fillRect(0, 0, outputWidth, outputHeight)
  const scaleX = outputWidth / mountRect.width
  const scaleY = outputHeight / mountRect.height
  for (const { layer, rect } of layers) {
    context.drawImage(layer, (rect.left - mountRect.left) * scaleX, (rect.top - mountRect.top) * scaleY, rect.width * scaleX, rect.height * scaleY)
  }
  return {
    dataUrl: outputCanvas.toDataURL('image/png'), width: outputWidth, height: outputHeight,
    scrollTop: verticalScroller?.scrollTop ?? 0,
    scrollHeight: verticalScroller?.scrollHeight ?? mountRect.height,
    viewportHeight: verticalScroller?.clientHeight ?? mountRect.height,
    scrollLeft: horizontalScroller?.scrollLeft ?? 0,
    scrollWidth: horizontalScroller?.scrollWidth ?? mountRect.width,
    viewportWidth: horizontalScroller?.clientWidth ?? mountRect.width,
  }
}

const openedWorkbookSessions = new Set<string>()

function SheetProductView(props: ProductViewProps) {
  const { sessionId } = props
  const language = useDshLanguage()
  const locale = univerLocale(language)
  const ui = UNIVER_CREATE_LOCALES[language]
  const { nodes, partialText } = useConversationFeed(props)
  const operations = useMemo(() => props.queuedOperations.map(queuedSheetOperation).filter((item): item is SheetOperation => item !== null), [props.queuedOperations])
  const chatLines = useMemo(() => nodes
    .map((node) => conversationLabel(node, ui.user))
    .filter((item): item is { role: string; text: string } => item !== null)
    .slice(-4), [nodes, ui.user])
  const containerRef = useRef<HTMLDivElement>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)
  const runtimeRef = useRef<MountedRuntime | null>(getResidentRuntime<MountedRuntime>(sessionId, 'sheet'))
  useEffect(() => runtimeRef.current?.univer.setLocale(locale), [locale])
  const appliedRef = useRef(new Set<SheetOperationId>())
  const chatStreamRef = useRef<HTMLDivElement>(null)
  const [title, setTitle] = useState(ui.defaultSheetTitle)
  useEffect(() => setTitle((current) => current === UNIVER_CREATE_LOCALES.en.defaultSheetTitle || current === UNIVER_CREATE_LOCALES.zh.defaultSheetTitle ? ui.defaultSheetTitle : current), [ui.defaultSheetTitle])
  const [error, setError] = useState<string | null>(null)
  const [importing, setImporting] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [filePath, setFilePath] = useState<string | null>(null)
  const [saveDialogOpen, setSaveDialogOpen] = useState(false)
  const [saveNotice, setSaveNotice] = useState('')
  const [sheetVisible, setSheetVisible] = useState(() => openedWorkbookSessions.has(sessionId))
  const [chatOverlayVisible, setChatOverlayVisible] = useState(true)
  // Always finish the Host lookup before mounting or replaying operations.
  // A fresh browser may have persisted data despite having no in-memory marker.
  const [hostLoaded, setHostLoaded] = useState(false)
  const [hostSnapshot, setHostSnapshot] = useState<unknown>(null)
  const hydratedFromHostRef = useRef(false)
  const lastSavedRef = useRef<string | null>(null)
  const lastFileSnapshotRef = useRef<string | null>(null)
  const savingRef = useRef(false)
  const pendingSaveRef = useRef<Promise<void> | null>(null)
  const sheetScreenshotInFlightRef = useRef(new Set<string>())
  const queuedPersistingRef = useRef(false)
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
    appliedRef.current.clear()
    lastSavedRef.current = null
    lastFileSnapshotRef.current = null
    setHostSnapshot(null)
    setFilePath(null)
    setSaveDialogOpen(false)
    setSaveNotice('')
    hydratedFromHostRef.current = false

    // A refresh clears the browser marker, but must not bypass durable storage.
    setHostLoaded(false)
    setSheetVisible(runtimeRef.current !== null || openedWorkbookSessions.has(sessionId))
    const restore = async () => {
      const connection = clientConnection.current
      if (connection === null) throw new Error('Host 连接不可用')
      const [result, restoredFileState] = await Promise.all([
        connection.rpc.call('/dsh-univer-create', 'load', { sessionId, unitType: 'sheet' }),
        loadUnitFileState(sessionId, 'sheet'),
      ])
      if (cancelled) return
      if (!result.ok) throw new Error(result.error.message)
      if (result.value !== null && (typeof result.value !== 'object' || Array.isArray(result.value))) {
        throw new Error('Host 返回了无效的表格快照')
      }
      setFilePath(restoredFileState.path)
      lastFileSnapshotRef.current = restoredFileState.snapshot === null ? UNKNOWN_FILE_SNAPSHOT : JSON.stringify(restoredFileState.snapshot)
      if (result.value !== null) {
        setHostSnapshot(result.value)
      } else {
        openedWorkbookSessions.delete(sessionId)
        setSheetVisible(false)
      }
      setError(null)
      setHostLoaded(true)
    }
    void restore().catch((reason) => {
      if (cancelled) return
      // Fail closed: replaying from blank after a failed load could overwrite data.
      setHostLoaded(false)
      setError(`无法恢复表格：${reason instanceof Error ? reason.message : String(reason)}`)
    })
    return () => { cancelled = true }
  }, [sessionId])

  useLayoutEffect(() => {
    const container = containerRef.current
    if (container === null) return
    const resident = runtimeRef.current
    if (resident !== null && resident.mount.parentElement !== container) {
      container.appendChild(resident.mount)
      registerResidentRuntime(
        sessionId,
        'sheet',
        resident,
        () => normalizeWorkbookSnapshot(resident.univerAPI.getActiveWorkbook()?.save()),
        savingRef,
        pendingSaveRef,
        lastSavedRef,
        lastFileSnapshotRef,
      )
    }
    if (!hostLoaded) return
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
      const mount = document.createElement('div')
      mount.className = 'dsh-univer-create-runtime-host'
      container.appendChild(mount)
      const runtime = createUniver({
        locale,
        locales: SHEET_LOCALES,
        theme: defaultTheme,
        presets: [UniverSheetsCorePreset({ container: mount })],
        plugins: [
          [UniverLicensePlugin, { license: UNIVER_LICENSE }],
          UniverSheetsChartPlugin,
          UniverSheetsChartUIPlugin,
        ],
      })
      runtime.univerAPI.createWorkbook(normalizeWorkbookSnapshot(snapshot ?? workbookData(operation, ui.defaultSheetTitle)))
      if (lastFileSnapshotRef.current === null) {
        lastFileSnapshotRef.current = JSON.stringify(normalizeWorkbookSnapshot(runtime.univerAPI.getActiveWorkbook()?.save()))
      }
      const mounted = Object.assign(runtime, { mount })
      runtimeRef.current = mounted
      registerResidentRuntime(
        sessionId,
        'sheet',
        mounted,
        () => normalizeWorkbookSnapshot(mounted.univerAPI.getActiveWorkbook()?.save()),
        savingRef,
        pendingSaveRef,
        lastSavedRef,
        lastFileSnapshotRef,
      )
      openedWorkbookSessions.add(sessionId)
      setSheetVisible(true)
      const snapshotTitle = snapshot !== null && typeof snapshot === 'object' && typeof (snapshot as any).name === 'string'
        ? (snapshot as any).name
        : undefined
      setTitle(operation?.title ?? snapshotTitle ?? ui.defaultSheetTitle)
    }

    try {
      if (runtimeRef.current === null && hostSnapshot === null && operations.length === 0) return
      if (runtimeRef.current === null) {
        createRuntime(undefined, hostSnapshot ?? undefined)
        if (hostSnapshot !== null && !hydratedFromHostRef.current) {
          // The Host snapshot is the applied baseline; every persisted queue item
          // still needs to be replayed and acknowledged exactly once.
          hydratedFromHostRef.current = true
          lastSavedRef.current = JSON.stringify(hostSnapshot)
        }
      }
      if (savingRef.current) return undefined
      for (const operation of operations) {
        if (appliedRef.current.has(operation.seq)) continue
        if (operation.action === 'new') {
          createRuntime(operation)
          setFilePath(null)
          setSaveNotice('')
          const blankSnapshot = normalizeWorkbookSnapshot(runtimeRef.current?.univerAPI.getActiveWorkbook()?.save())
          void clientConnection.current?.rpc.call('/dsh-univer-create', 'save', { sessionId, unitType: 'sheet', snapshot: blankSnapshot, filePath: null })
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

      const queuedIds = operations
        .map((operation) => operation.seq)
        .filter((id): id is string => typeof id === 'string' && appliedRef.current.has(id))
      if (queuedIds.length > 0 && !queuedPersistingRef.current) {
        queuedPersistingRef.current = true
        savingRef.current = true
        const previousSave = pendingSaveRef.current
        const commit = (async () => {
          await previousSave
          const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
          if (workbook === null || workbook === undefined) return
          const snapshot = normalizeWorkbookSnapshot(workbook.save())
          const serialized = JSON.stringify(snapshot)
          await commitOperations(sessionId, 'sheet', snapshot, queuedIds)
          lastSavedRef.current = serialized
        })()
        pendingSaveRef.current = commit
        void commit.catch((reason) => {
          setError(`提交表格队列操作失败：${reason instanceof Error ? reason.message : String(reason)}`)
        }).finally(() => {
          if (pendingSaveRef.current === commit) pendingSaveRef.current = null
          savingRef.current = false
          queuedPersistingRef.current = false
        })
      }
      setError(null)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
    }

    return undefined
  }, [hostLoaded, hostSnapshot, operations, layoutVersion, sessionId, locale])

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
        if (pendingSaveRef.current === save) {
          savingRef.current = false
          pendingSaveRef.current = null
        }
      })
      pendingSaveRef.current = save
    }, 800)
    return () => window.clearInterval(timer)
  }, [sessionId])

  useEffect(() => {
    if (!hostLoaded || !sheetVisible) return
    const connection = clientConnection.current
    if (connection === null) return
    for (const request of props.sheetScreenshotRequests) {
      if (sheetScreenshotInFlightRef.current.has(request.id)) continue
      sheetScreenshotInFlightRef.current.add(request.id)
      void (async () => {
        try {
          const runtime = runtimeRef.current
          if (runtime === null) throw new Error('当前浏览器没有打开已渲染的 Univer Sheet')
          const captured = await captureRenderedSheet(runtime, request)
          await connection.rpc.call('/dsh-univer-create', 'sheet-screenshot-result', { sessionId, screenshotId: request.id, ...captured })
        } catch (reason) {
          await connection.rpc.call('/dsh-univer-create', 'sheet-screenshot-result', {
            sessionId,
            screenshotId: request.id,
            error: reason instanceof Error ? reason.message : String(reason),
          }).catch(() => {})
        } finally {
          sheetScreenshotInFlightRef.current.delete(request.id)
        }
      })()
    }
  }, [sessionId, hostLoaded, sheetVisible, props.sheetScreenshotRequests])

  useLayoutEffect(() => () => {
    const resident = runtimeRef.current
    const entry = residentRuntimes.get(residentRuntimeKey(sessionId, 'sheet'))
    const connection = clientConnection.current
    if (resident !== null && entry?.runtime === resident && connection !== null) {
      queueResidentFinalSave(entry, async () => {
        const workbook = resident.univerAPI.getActiveWorkbook()
        if (workbook === undefined || workbook === null) return
        const snapshot = normalizeWorkbookSnapshot(workbook.save())
        await connection.rpc.call('/dsh-univer-create', 'save', { sessionId, unitType: 'sheet', snapshot })
      })
    }
    parkResidentRuntime(sessionId, 'sheet', resident)
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
        if (chosenPath !== undefined && isWorkspaceFileExistsError(reason) && window.confirm(ui.confirmOverwrite(targetPath))) {
          savedPath = await saveUnitFile(sessionId, 'sheet', nextSnapshot, targetPath, true)
        } else {
          throw reason
        }
      }
      setFilePath(savedPath)
      setSaveDialogOpen(false)
      setSaveNotice(ui.savedTo(savedPath))
      lastFileSnapshotRef.current = JSON.stringify(nextSnapshot)
      lastSavedRef.current = lastFileSnapshotRef.current
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
    const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
    if (workbook !== undefined && workbook !== null && !confirmOpenWithUnsavedChanges(
      normalizeWorkbookSnapshot(workbook.save()),
      lastFileSnapshotRef.current,
      ui.confirmDiscardChangesForNew,
    )) return
    const blankSnapshot = normalizeWorkbookSnapshot(workbookData({ action: 'new', seq: Date.now() }, ui.defaultSheetTitle))
    disposeResidentRuntime(sessionId, 'sheet', runtimeRef.current)
    runtimeRef.current = null
    // A manually created workbook replaces both the persisted workbook and any
    // conversation operations that belong to the previous workbook.
    appliedRef.current.clear()
    hydratedFromHostRef.current = true
    lastSavedRef.current = null
    lastFileSnapshotRef.current = null
    openedWorkbookSessions.add(sessionId)
    setError(null)
    setTitle(ui.defaultSheetTitle)
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
      disposeResidentRuntime(sessionId, 'sheet', runtimeRef.current)
      runtimeRef.current = null
      // An imported file replaces the conversation-generated workbook. Keep the
      // historical tool calls from replaying over the imported workbook.
      appliedRef.current.clear()
      hydratedFromHostRef.current = true
      lastSavedRef.current = null
      lastFileSnapshotRef.current = null
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

  const openWorkbookFile = () => {
    const workbook = runtimeRef.current?.univerAPI.getActiveWorkbook()
    if (workbook !== undefined && workbook !== null && !confirmOpenWithUnsavedChanges(
      normalizeWorkbookSnapshot(workbook.save()),
      lastFileSnapshotRef.current,
      ui.confirmDiscardChanges,
    )) return
    fileInputRef.current?.click()
  }

  const actionButtons = (
    <>
      <button className="dsh-univer-create-action" type="button" disabled={importing || exporting || !hostLoaded} onClick={openWorkbookFile}>
        {importing ? ui.opening : ui.open}
      </button>
      <button className="dsh-univer-create-action dsh-univer-create-action--primary" type="button" disabled={importing || exporting || !hostLoaded} onClick={newWorkbook}>
        {ui.create}
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
            {saveNotice || (filePath !== null ? ui.fileStatus(filePath) : operations.length > 0 ? ui.syncedSheetOperations(operations.length) : ui.newFileHint)}
          </span>
          {actionButtons}
          <button
            className="dsh-univer-create-action"
            type="button"
            disabled={importing || exporting || !hostLoaded}
            onClick={() => void saveAsXlsx()}
          >
            {exporting ? ui.saving : ui.save}
          </button>
        </div>
      )}
      <div ref={containerRef} className="dsh-univer-create-container" />
      {saveDialogOpen && <SavePathDialog extension="xlsx" suggestedName={title} busy={exporting} onCancel={() => setSaveDialogOpen(false)} onSave={(path) => void saveAsXlsx(path)} />}
      {hostLoaded && !sheetVisible && (
        <div className="dsh-univer-create-welcome" aria-label={ui.startSheet}>
          {actionButtons}
        </div>
      )}
      {sheetVisible && (
        <>
          <aside className="dsh-univer-create-chat-overlay" aria-label={ui.chatActivity}>
            <div className="dsh-univer-create-chat-overlay__header">
              <div className="dsh-univer-create-chat-overlay__title">{ui.chatActivity}</div>
              <button
                className="dsh-univer-create-chat-overlay__toggle"
                type="button"
                aria-label={chatOverlayVisible ? ui.closeChatActivity : ui.openChatActivity}
                aria-expanded={chatOverlayVisible}
                onClick={() => setChatOverlayVisible((visible) => !visible)}
              >
                {chatOverlayVisible ? ui.hide : ui.show}
              </button>
            </div>
            <div ref={chatStreamRef} className="dsh-univer-create-chat-overlay__stream" hidden={!chatOverlayVisible}>
              {chatLines.map((line, index) => (
                <div className="dsh-univer-create-chat-overlay__line" key={`${line.role}-${index}-${line.text.slice(0, 16)}`}>
                  <strong>{line.role}</strong><span>{line.text}</span>
                </div>
              ))}
              {partialText && <div className="dsh-univer-create-chat-overlay__line dsh-univer-create-chat-overlay__line--live"><strong>AI</strong><span>{partialText}</span></div>}
            </div>
          </aside>
        </>
      )}
      {error !== null && <div className="dsh-univer-create-error" role="alert">{ui.sheetError}{error}{!hostLoaded && ui.hostRecoveryHint}</div>}
    </section>
  )
}

function findDocumentScroller(mount: HTMLElement): HTMLElement | null {
  const candidates = [mount, ...mount.querySelectorAll<HTMLElement>('*')].filter((element) => (
    element.clientHeight >= 200 && element.scrollHeight > element.clientHeight + 8
  ))
  return candidates.sort((left, right) => {
    const leftScore = (left.scrollHeight - left.clientHeight) * Math.max(1, left.clientWidth)
    const rightScore = (right.scrollHeight - right.clientHeight) * Math.max(1, right.clientWidth)
    return rightScore - leftScore
  })[0] ?? null
}

async function captureRenderedDocument(
  runtime: MountedRuntime,
  scroll: DocScreenshotRequest['scroll'],
  amount?: number,
): Promise<{ dataUrl: string; width: number; height: number; scrollTop: number; scrollHeight: number; viewportHeight: number }> {
  const scroller = findDocumentScroller(runtime.mount)
  if (scroll !== 'none' && scroller === null) throw new Error('没有找到可滚动的文档区域')
  if (scroller !== null) {
    const distance = Math.max(1, amount ?? scroller.clientHeight * 0.8)
    const target = scroll === 'top'
      ? 0
      : scroll === 'bottom'
        ? scroller.scrollHeight
        : scroll === 'up'
          ? scroller.scrollTop - distance
          : scroll === 'down' ? scroller.scrollTop + distance : scroller.scrollTop
    scroller.scrollTo({ top: target, behavior: 'auto' })
  }
  window.dispatchEvent(new Event('resize'))
  if (document.fonts?.ready !== undefined) await document.fonts.ready
  await new Promise<void>((resolve) => window.setTimeout(resolve, 180))
  await waitForAnimationFrames(3)

  const candidates = [...runtime.mount.querySelectorAll('canvas')].map((canvas) => ({
    canvas,
    rect: canvas.getBoundingClientRect(),
  })).filter(({ canvas, rect }) => canvas.width >= 200 && canvas.height >= 200 && rect.width >= 240 && rect.height >= 240)
  if (candidates.length === 0) throw new Error('没有找到已渲染的文档画布')
  const primary = candidates.sort((left, right) => right.rect.width * right.rect.height - left.rect.width * left.rect.height)[0]!
  const layers = candidates.filter(({ rect }) => (
    Math.abs(rect.left - primary.rect.left) <= 2
    && Math.abs(rect.top - primary.rect.top) <= 2
    && Math.abs(rect.width - primary.rect.width) <= 2
    && Math.abs(rect.height - primary.rect.height) <= 2
  ))
  const outputWidth = Math.max(1, Math.min(1600, Math.round(primary.rect.width * Math.min(window.devicePixelRatio || 1, 2))))
  const outputHeight = Math.max(1, Math.round(outputWidth * primary.rect.height / primary.rect.width))
  const outputCanvas = document.createElement('canvas')
  outputCanvas.width = outputWidth
  outputCanvas.height = outputHeight
  const context = outputCanvas.getContext('2d')
  if (context === null) throw new Error('浏览器无法创建文档截图画布')
  context.fillStyle = '#ffffff'
  context.fillRect(0, 0, outputWidth, outputHeight)
  for (const { canvas } of layers) context.drawImage(canvas, 0, 0, outputWidth, outputHeight)
  return {
    dataUrl: outputCanvas.toDataURL('image/png'),
    width: outputWidth,
    height: outputHeight,
    scrollTop: scroller?.scrollTop ?? 0,
    scrollHeight: scroller?.scrollHeight ?? primary.rect.height,
    viewportHeight: scroller?.clientHeight ?? primary.rect.height,
  }
}

const openedDocumentSessions = new Set<string>()

function DocProductView(props: ProductViewProps) {
  const { sessionId } = props
  const language = useDshLanguage()
  const locale = univerLocale(language)
  const ui = UNIVER_CREATE_LOCALES[language]
  const { nodes, partialText } = useConversationFeed(props)
  const operations = useMemo(() => props.queuedOperations.map(queuedDocOperation).filter((item): item is DocOperation => item !== null), [props.queuedOperations])
  const chatLines = useMemo(() => nodes
    .map((node) => conversationLabel(node, ui.user))
    .filter((item): item is { role: string; text: string } => item !== null)
    .slice(-4), [nodes, ui.user])
  const containerRef = useRef<HTMLDivElement>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)
  const runtimeRef = useRef<MountedRuntime | null>(getResidentRuntime<MountedRuntime>(sessionId, 'doc'))
  useEffect(() => runtimeRef.current?.univer.setLocale(locale), [locale])
  const appliedRef = useRef(new Set<SheetOperationId>())
  const chatStreamRef = useRef<HTMLDivElement>(null)
  const [title, setTitle] = useState(ui.defaultDocTitle)
  useEffect(() => setTitle((current) => current === UNIVER_CREATE_LOCALES.en.defaultDocTitle || current === UNIVER_CREATE_LOCALES.zh.defaultDocTitle ? ui.defaultDocTitle : current), [ui.defaultDocTitle])
  const [error, setError] = useState<string | null>(null)
  const [importing, setImporting] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [filePath, setFilePath] = useState<string | null>(null)
  const [saveDialogOpen, setSaveDialogOpen] = useState(false)
  const [saveNotice, setSaveNotice] = useState('')
  const [documentVisible, setDocumentVisible] = useState(() => openedDocumentSessions.has(sessionId))
  const [chatOverlayVisible, setChatOverlayVisible] = useState(true)
  const [hostLoaded, setHostLoaded] = useState(false)
  const [hostSnapshot, setHostSnapshot] = useState<unknown>(null)
  const hydratedFromHostRef = useRef(false)
  const lastSavedRef = useRef<string | null>(null)
  const lastFileSnapshotRef = useRef<string | null>(null)
  const savingRef = useRef(false)
  const pendingSaveRef = useRef<Promise<void> | null>(null)
  const queuedPersistingRef = useRef(false)
  const docScreenshotInFlightRef = useRef(new Set<string>())
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
    appliedRef.current.clear()
    lastSavedRef.current = null
    lastFileSnapshotRef.current = null
    setHostSnapshot(null)
    setFilePath(null)
    setSaveDialogOpen(false)
    setSaveNotice('')
    hydratedFromHostRef.current = false

    // A refresh clears the browser marker, but must not bypass durable storage.
    setHostLoaded(false)
    setDocumentVisible(runtimeRef.current !== null || openedDocumentSessions.has(sessionId))
    const restore = async () => {
      const connection = clientConnection.current
      if (connection === null) throw new Error('Host 连接不可用')
      const [result, restoredFileState] = await Promise.all([
        connection.rpc.call('/dsh-univer-create', 'load', { sessionId, unitType: 'doc' }),
        loadUnitFileState(sessionId, 'doc'),
      ])
      if (cancelled) return
      if (!result.ok) throw new Error(result.error.message)
      if (result.value !== null && (typeof result.value !== 'object' || Array.isArray(result.value))) {
        throw new Error('Host 返回了无效的文档快照')
      }
      setFilePath(restoredFileState.path)
      lastFileSnapshotRef.current = restoredFileState.snapshot === null ? UNKNOWN_FILE_SNAPSHOT : JSON.stringify(restoredFileState.snapshot)
      if (result.value !== null) {
        setHostSnapshot(result.value)
      } else {
        openedDocumentSessions.delete(sessionId)
        setDocumentVisible(false)
      }
      setError(null)
      setHostLoaded(true)
    }
    void restore().catch((reason) => {
      if (cancelled) return
      // Fail closed: replaying from blank after a failed load could overwrite data.
      setHostLoaded(false)
      setError(`无法恢复文档：${reason instanceof Error ? reason.message : String(reason)}`)
    })
    return () => { cancelled = true }
  }, [sessionId])

  useLayoutEffect(() => {
    const container = containerRef.current
    if (container === null) return
    const resident = runtimeRef.current
    if (resident !== null && resident.mount.parentElement !== container) {
      container.appendChild(resident.mount)
      registerResidentRuntime(
        sessionId,
        'doc',
        resident,
        () => resident.univerAPI.getActiveDocument()?.save(),
        savingRef,
        pendingSaveRef,
        lastSavedRef,
        lastFileSnapshotRef,
      )
    }
    if (!hostLoaded) return
    const bounds = container.getBoundingClientRect()
    if (bounds.width <= 0 || bounds.height <= 0) return

    const createRuntime = (
      operation?: Extract<DocOperation, { action: 'new-doc' }>,
      restoredSnapshot?: unknown,
    ) => {
      const mount = document.createElement('div')
      mount.className = 'dsh-univer-create-runtime-host'
      container.appendChild(mount)
      const runtime = createUniver({
        locale,
        locales: DOC_LOCALES,
        theme: defaultTheme,
        presets: [
          UniverDocsCorePreset({ container: mount }),
          UniverDocsDrawingPreset(),
        ],
        plugins: [
          [UniverLicensePlugin, { license: UNIVER_LICENSE }],
          UniverDocsTablePlugin,
        ],
      })
      const fDocument = runtime.univerAPI.createDocument((restoredSnapshot ?? documentData(operation, locale, ui.defaultDocTitle)) as any)
      if (restoredSnapshot === undefined && operation?.text) fDocument.insertText(0, operation.text)
      if (lastFileSnapshotRef.current === null) lastFileSnapshotRef.current = JSON.stringify(fDocument.save())
      const mounted = Object.assign(runtime, { mount })
      runtimeRef.current = mounted
      registerResidentRuntime(
        sessionId,
        'doc',
        mounted,
        () => mounted.univerAPI.getActiveDocument()?.save(),
        savingRef,
        pendingSaveRef,
        lastSavedRef,
        lastFileSnapshotRef,
      )
      openedDocumentSessions.add(sessionId)
      setDocumentVisible(true)
      const snapshotTitle = restoredSnapshot !== null && typeof restoredSnapshot === 'object' && typeof (restoredSnapshot as any).title === 'string'
        ? (restoredSnapshot as any).title
        : undefined
      setTitle(operation?.title ?? snapshotTitle ?? ui.defaultDocTitle)
    }

    try {
      if (runtimeRef.current === null && hostSnapshot === null && operations.length === 0) return
      if (runtimeRef.current === null) {
        createRuntime(undefined, hostSnapshot ?? undefined)
        if (hostSnapshot !== null && !hydratedFromHostRef.current) {
          hydratedFromHostRef.current = true
          lastSavedRef.current = JSON.stringify(hostSnapshot)
        }
      }

      if (savingRef.current) return undefined
      for (const operation of operations) {
        if (appliedRef.current.has(operation.seq)) continue
        if (operation.action === 'new-doc') {
          createRuntime(operation)
          setFilePath(null)
          setSaveNotice('')
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

      const queuedIds = operations
        .map((operation) => operation.seq)
        .filter((id): id is string => typeof id === 'string' && appliedRef.current.has(id))
      if (queuedIds.length > 0 && !queuedPersistingRef.current) {
        queuedPersistingRef.current = true
        savingRef.current = true
        const previousSave = pendingSaveRef.current
        const commit = (async () => {
          await previousSave
          const fDocument = runtimeRef.current?.univerAPI.getActiveDocument()
          if (fDocument === null || fDocument === undefined) return
          const nextSnapshot = fDocument.save()
          const serialized = JSON.stringify(nextSnapshot)
          await commitOperations(sessionId, 'doc', nextSnapshot, queuedIds)
          lastSavedRef.current = serialized
        })()
        pendingSaveRef.current = commit
        void commit.catch((reason) => {
          setError(`提交文档队列操作失败：${reason instanceof Error ? reason.message : String(reason)}`)
        }).finally(() => {
          if (pendingSaveRef.current === commit) pendingSaveRef.current = null
          savingRef.current = false
          queuedPersistingRef.current = false
        })
      }
      setError(null)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
    }
    return undefined
  }, [hostLoaded, hostSnapshot, operations, layoutVersion, sessionId, locale])

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
        if (pendingSaveRef.current === save) {
          savingRef.current = false
          pendingSaveRef.current = null
        }
      })
      pendingSaveRef.current = save
    }, 800)
    return () => window.clearInterval(timer)
  }, [sessionId])

  useEffect(() => {
    if (!hostLoaded || !documentVisible) return
    const connection = clientConnection.current
    if (connection === null) return
    for (const request of props.docScreenshotRequests) {
      if (docScreenshotInFlightRef.current.has(request.id)) continue
      docScreenshotInFlightRef.current.add(request.id)
      void (async () => {
        try {
          const runtime = runtimeRef.current
          if (runtime === null) throw new Error('当前浏览器没有打开已渲染的 Univer Doc')
          const captured = await captureRenderedDocument(runtime, request.scroll ?? 'none', request.amount)
          await connection.rpc.call('/dsh-univer-create', 'doc-screenshot-result', { sessionId, screenshotId: request.id, ...captured })
        } catch (reason) {
          await connection.rpc.call('/dsh-univer-create', 'doc-screenshot-result', {
            sessionId,
            screenshotId: request.id,
            error: reason instanceof Error ? reason.message : String(reason),
          }).catch(() => {})
        } finally {
          docScreenshotInFlightRef.current.delete(request.id)
        }
      })()
    }
  }, [sessionId, hostLoaded, documentVisible, props.docScreenshotRequests])

  useLayoutEffect(() => () => {
    const resident = runtimeRef.current
    const entry = residentRuntimes.get(residentRuntimeKey(sessionId, 'doc'))
    const connection = clientConnection.current
    if (resident !== null && entry?.runtime === resident && connection !== null) {
      queueResidentFinalSave(entry, async () => {
        const fDocument = resident.univerAPI.getActiveDocument()
        if (fDocument === undefined || fDocument === null) return
        const snapshot = fDocument.save()
        await connection.rpc.call('/dsh-univer-create', 'save', { sessionId, unitType: 'doc', snapshot })
      })
    }
    parkResidentRuntime(sessionId, 'doc', resident)
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
    const document = runtimeRef.current?.univerAPI.getActiveDocument()
    if (document !== undefined && document !== null && !confirmOpenWithUnsavedChanges(
      document.save(),
      lastFileSnapshotRef.current,
      ui.confirmDiscardChangesForNew,
    )) return
    const blankSnapshot = documentData({ action: 'new-doc', seq: Date.now() }, locale, ui.defaultDocTitle)
    disposeResidentRuntime(sessionId, 'doc', runtimeRef.current)
    runtimeRef.current = null
    appliedRef.current.clear()
    hydratedFromHostRef.current = true
    lastSavedRef.current = null
    lastFileSnapshotRef.current = null
    openedDocumentSessions.add(sessionId)
    setError(null)
    setTitle(ui.defaultDocTitle)
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
      disposeResidentRuntime(sessionId, 'doc', runtimeRef.current)
      runtimeRef.current = null
      appliedRef.current.clear()
      hydratedFromHostRef.current = true
      lastSavedRef.current = null
      lastFileSnapshotRef.current = null
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

  const openDocumentFile = () => {
    const document = runtimeRef.current?.univerAPI.getActiveDocument()
    if (document !== undefined && document !== null && !confirmOpenWithUnsavedChanges(
      document.save(),
      lastFileSnapshotRef.current,
      ui.confirmDiscardChanges,
    )) return
    fileInputRef.current?.click()
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
        if (chosenPath !== undefined && isWorkspaceFileExistsError(reason) && window.confirm(ui.confirmOverwrite(targetPath))) {
          savedPath = await saveUnitFile(sessionId, 'doc', nextSnapshot, targetPath, true)
        } else {
          throw reason
        }
      }
      setFilePath(savedPath)
      setSaveDialogOpen(false)
      setSaveNotice(ui.savedTo(savedPath))
      lastFileSnapshotRef.current = JSON.stringify(nextSnapshot)
      lastSavedRef.current = lastFileSnapshotRef.current
    } catch (reason) {
      setError(`保存 DOCX 失败：${reason instanceof Error ? reason.message : String(reason)}`)
    } finally {
      setExporting(false)
    }
  }

  const actionButtons = (
    <>
      <button className="dsh-univer-create-action" type="button" disabled={importing || exporting || !hostLoaded} onClick={openDocumentFile}>
        {importing ? ui.opening : ui.open}
      </button>
      <button className="dsh-univer-create-action dsh-univer-create-action--primary" type="button" disabled={importing || exporting || !hostLoaded} onClick={newDocument}>
        {ui.create}
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
            {saveNotice || (filePath !== null ? ui.fileStatus(filePath) : operations.length > 0 ? ui.syncedDocOperations(operations.length) : ui.newFileHint)}
          </span>
          {actionButtons}
          <button className="dsh-univer-create-action" type="button" disabled={importing || exporting || !hostLoaded} onClick={() => void saveAsDocx()}>
            {exporting ? ui.saving : ui.save}
          </button>
        </div>
      )}
      <div ref={containerRef} className="dsh-univer-create-container" />
      {saveDialogOpen && <SavePathDialog extension="docx" suggestedName={title} busy={exporting} onCancel={() => setSaveDialogOpen(false)} onSave={(path) => void saveAsDocx(path)} />}
      {hostLoaded && !documentVisible && <div className="dsh-univer-create-welcome" aria-label={ui.startDoc}>{actionButtons}</div>}
      {documentVisible && (
        <>
          <aside className="dsh-univer-create-chat-overlay" aria-label={ui.chatActivity}>
            <div className="dsh-univer-create-chat-overlay__header">
              <div className="dsh-univer-create-chat-overlay__title">{ui.chatActivity}</div>
              <button
                className="dsh-univer-create-chat-overlay__toggle"
                type="button"
                aria-label={chatOverlayVisible ? ui.closeChatActivity : ui.openChatActivity}
                aria-expanded={chatOverlayVisible}
                onClick={() => setChatOverlayVisible((visible) => !visible)}
              >
                {chatOverlayVisible ? ui.hide : ui.show}
              </button>
            </div>
            <div ref={chatStreamRef} className="dsh-univer-create-chat-overlay__stream" hidden={!chatOverlayVisible}>
              {chatLines.map((line, index) => (
                <div className="dsh-univer-create-chat-overlay__line" key={`${line.role}-${index}-${line.text.slice(0, 16)}`}>
                  <strong>{line.role}</strong><span>{line.text}</span>
                </div>
              ))}
              {partialText && <div className="dsh-univer-create-chat-overlay__line dsh-univer-create-chat-overlay__line--live"><strong>AI</strong><span>{partialText}</span></div>}
            </div>
          </aside>
        </>
      )}
      {error !== null && <div className="dsh-univer-create-error" role="alert">{ui.docError}{error}{!hostLoaded && ui.hostRecoveryHint}</div>}
    </section>
  )
}

const openedPresentationSessions = new Set<string>()

function waitForAnimationFrames(count = 2): Promise<void> {
  return new Promise((resolve) => {
    const next = (remaining: number) => {
      if (remaining <= 0) resolve()
      else window.requestAnimationFrame(() => next(remaining - 1))
    }
    next(count)
  })
}

async function captureRenderedSlide(
  runtime: MountedSlideRuntime,
  slideIndex: number,
  mode: 'slide' | 'editor',
): Promise<{ dataUrl: string; width: number; height: number }> {
  const slide = runtime.presentation.getSlideByIndex(slideIndex)
  if (slide === null) throw new Error(`找不到第 ${slideIndex + 1} 张幻灯片`)
  runtime.presentation.setActiveSlide(slide)
  window.dispatchEvent(new Event('resize'))
  if (document.fonts?.ready !== undefined) await document.fonts.ready
  await new Promise<void>((resolve) => window.setTimeout(resolve, 180))
  await waitForAnimationFrames(3)

  const pageSize = slide.getPageSize()
  const pageAspect = pageSize.width / pageSize.height
  const candidates = [...runtime.mount.querySelectorAll('canvas')].map((canvas) => ({
    canvas,
    rect: canvas.getBoundingClientRect(),
  })).filter(({ canvas, rect }) => canvas.width >= 200 && canvas.height >= 100 && rect.width >= 240 && rect.height >= 120)
  if (candidates.length === 0) throw new Error('没有找到已渲染的幻灯片画布')

  const primary = candidates.sort((left, right) => {
    const score = ({ rect }: typeof left) => rect.width * rect.height / (1 + Math.abs(Math.log((rect.width / rect.height) / pageAspect)) * 6)
    return score(right) - score(left)
  })[0]!
  const layers = candidates.filter(({ rect }) => (
    Math.abs(rect.left - primary.rect.left) <= 2
    && Math.abs(rect.top - primary.rect.top) <= 2
    && Math.abs(rect.width - primary.rect.width) <= 2
    && Math.abs(rect.height - primary.rect.height) <= 2
  ))

  let cropLeft = 0
  let cropTop = 0
  let cropWidth = primary.rect.width
  let cropHeight = primary.rect.height
  if (mode === 'slide') {
    const currentAspect = cropWidth / cropHeight
    if (currentAspect > pageAspect) {
      cropWidth = cropHeight * pageAspect
      cropLeft = (primary.rect.width - cropWidth) / 2
    } else if (currentAspect < pageAspect) {
      cropHeight = cropWidth / pageAspect
      cropTop = (primary.rect.height - cropHeight) / 2
    }
  }

  const outputWidth = Math.max(1, Math.min(1600, Math.round(cropWidth * Math.min(window.devicePixelRatio || 1, 2))))
  const outputHeight = Math.max(1, Math.round(outputWidth * cropHeight / cropWidth))
  const outputCanvas = document.createElement('canvas')
  outputCanvas.width = outputWidth
  outputCanvas.height = outputHeight
  const context = outputCanvas.getContext('2d')
  if (context === null) throw new Error('浏览器无法创建截图画布')
  context.fillStyle = '#ffffff'
  context.fillRect(0, 0, outputWidth, outputHeight)
  for (const { canvas, rect } of layers) {
    const scaleX = canvas.width / rect.width
    const scaleY = canvas.height / rect.height
    context.drawImage(
      canvas,
      cropLeft * scaleX,
      cropTop * scaleY,
      cropWidth * scaleX,
      cropHeight * scaleY,
      0,
      0,
      outputWidth,
      outputHeight,
    )
  }
  return { dataUrl: outputCanvas.toDataURL('image/png'), width: outputWidth, height: outputHeight }
}

function SlideProductView(props: ProductViewProps) {
  const { sessionId } = props
  const language = useDshLanguage()
  const locale = univerLocale(language)
  const ui = UNIVER_CREATE_LOCALES[language]
  const { nodes, partialText } = useConversationFeed(props)
  const operations = useMemo(() => props.queuedOperations.map(queuedSlideOperation).filter((item): item is SlideOperation => item !== null), [props.queuedOperations])
  const chatLines = useMemo(() => nodes
    .map((node) => conversationLabel(node, ui.user))
    .filter((item): item is { role: string; text: string } => item !== null)
    .slice(-4), [nodes, ui.user])
  const containerRef = useRef<HTMLDivElement>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)
  const runtimeRef = useRef<MountedSlideRuntime | null>(getResidentRuntime<MountedSlideRuntime>(sessionId, 'slide'))
  useEffect(() => runtimeRef.current?.univer.setLocale(locale), [locale])
  const appliedRef = useRef(new Set<SheetOperationId>())
  const chatStreamRef = useRef<HTMLDivElement>(null)
  const [title, setTitle] = useState(ui.defaultSlideTitle)
  useEffect(() => setTitle((current) => current === UNIVER_CREATE_LOCALES.en.defaultSlideTitle || current === UNIVER_CREATE_LOCALES.zh.defaultSlideTitle ? ui.defaultSlideTitle : current), [ui.defaultSlideTitle])
  const [error, setError] = useState<string | null>(null)
  const [importing, setImporting] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [filePath, setFilePath] = useState<string | null>(null)
  const [saveDialogOpen, setSaveDialogOpen] = useState(false)
  const [saveNotice, setSaveNotice] = useState('')
  const [presentationVisible, setPresentationVisible] = useState(() => openedPresentationSessions.has(sessionId))
  const [chatOverlayVisible, setChatOverlayVisible] = useState(true)
  const [hostLoaded, setHostLoaded] = useState(false)
  const [hostSnapshot, setHostSnapshot] = useState<ISlideData | null>(null)
  const hydratedFromHostRef = useRef(false)
  const lastSavedRef = useRef<string | null>(null)
  const lastFileSnapshotRef = useRef<string | null>(null)
  const savingRef = useRef(false)
  const pendingSaveRef = useRef<Promise<void> | null>(null)
  const queuedPersistingRef = useRef(false)
  const screenshotInFlightRef = useRef(new Set<string>())
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

  useLayoutEffect(() => {
    const container = containerRef.current
    if (container === null) return
    // Confine Univer's presentation `scrollIntoView`/`focus` calls to this tab;
    // see scroll-containment.js for the full mechanism.
    return installScrollContainment(container)
  }, [])

  useEffect(() => {
    const stream = chatStreamRef.current
    if (stream !== null) stream.scrollTop = stream.scrollHeight
  }, [chatLines, partialText])

  useEffect(() => {
    let cancelled = false
    appliedRef.current.clear()
    lastSavedRef.current = null
    lastFileSnapshotRef.current = null
    setHostSnapshot(null)
    setFilePath(null)
    setSaveDialogOpen(false)
    setSaveNotice('')
    hydratedFromHostRef.current = false

    // A refresh clears the browser marker, but must not bypass durable storage.
    setHostLoaded(false)
    setPresentationVisible(runtimeRef.current !== null || openedPresentationSessions.has(sessionId))
    const restore = async () => {
      const connection = clientConnection.current
      if (connection === null) throw new Error('Host 连接不可用')
      const [result, restoredFileState] = await Promise.all([
        connection.rpc.call('/dsh-univer-create', 'load', { sessionId, unitType: 'slide' }),
        loadUnitFileState(sessionId, 'slide'),
      ])
      if (cancelled) return
      if (!result.ok) throw new Error(result.error.message)
      if (result.value !== null && (typeof result.value !== 'object' || Array.isArray(result.value))) {
        throw new Error('Host 返回了无效的演示文稿快照')
      }
      setFilePath(restoredFileState.path)
      lastFileSnapshotRef.current = restoredFileState.snapshot === null ? UNKNOWN_FILE_SNAPSHOT : JSON.stringify(restoredFileState.snapshot)
      if (result.value !== null) {
        setHostSnapshot(result.value as ISlideData)
      } else {
        openedPresentationSessions.delete(sessionId)
        setPresentationVisible(false)
      }
      setError(null)
      setHostLoaded(true)
    }
    void restore().catch((reason) => {
      if (cancelled) return
      // Fail closed: replaying from blank after a failed load could overwrite data.
      setHostLoaded(false)
      setError(`无法恢复演示文稿：${reason instanceof Error ? reason.message : String(reason)}`)
    })
    return () => { cancelled = true }
  }, [sessionId])

  useLayoutEffect(() => {
    const container = containerRef.current
    if (container === null) return
    const resident = runtimeRef.current
    if (resident !== null && resident.mount.parentElement !== container) {
      container.appendChild(resident.mount)
      registerResidentRuntime(
        sessionId,
        'slide',
        resident,
        () => resident.presentation.save(),
        savingRef,
        pendingSaveRef,
        lastSavedRef,
        lastFileSnapshotRef,
      )
    }
    if (!hostLoaded) return
    const bounds = container.getBoundingClientRect()
    if (bounds.width <= 0 || bounds.height <= 0) return

    const createRuntime = (
      operation?: Extract<SlideOperation, { action: 'new-slide' }>,
      restoredSnapshot?: ISlideData,
    ) => {
      slideResizeTimersRef.current.forEach((timer) => window.clearTimeout(timer))
      slideResizeTimersRef.current = []
      const mount = document.createElement('div')
      mount.className = 'dsh-univer-create-runtime-host'
      container.appendChild(mount)

      const univer = new Univer({
        locale,
        locales: SLIDE_LOCALES,
      })
      univer.registerPlugin(UniverRenderEnginePlugin)
      univer.registerPlugin(UniverUIPlugin, { container: mount })
      univer.registerPlugin(UniverDocsPlugin)
      univer.registerPlugin(UniverDocsUIPlugin)
      univer.registerPlugin(UniverDrawingPlugin)
      univer.registerPlugin(UniverLicensePlugin, { license: UNIVER_LICENSE })
      univer.registerPlugin(UniverSlidesPlugin)
      univer.registerPlugin(UniverSlidesChartPlugin)
      univer.registerPlugin(UniverSlidesUIPlugin)
      univer.registerPlugin(UniverSlidesChartUIPlugin)

      const univerAPI = FUniver.newAPI(univer)
      const presentation = univerAPI.createPresentation(normalizeAiSlideLayouts(restoredSnapshot ?? presentationData(operation, locale, ui.defaultSlideTitle)))
      if (lastFileSnapshotRef.current === null) lastFileSnapshotRef.current = JSON.stringify(presentation.save())
      const mounted = { univer, univerAPI, mount, presentation }
      runtimeRef.current = mounted
      registerResidentRuntime(
        sessionId,
        'slide',
        mounted,
        () => presentation.save(),
        savingRef,
        pendingSaveRef,
        lastSavedRef,
        lastFileSnapshotRef,
      )
      openedPresentationSessions.add(sessionId)
      setPresentationVisible(true)
      setTitle(operation?.title ?? restoredSnapshot?.name ?? ui.defaultSlideTitle)

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
          hydratedFromHostRef.current = true
          lastSavedRef.current = JSON.stringify(hostSnapshot)
        }
      }

      if (savingRef.current) return undefined
      for (const operation of operations) {
        if (appliedRef.current.has(operation.seq)) continue
        if (operation.action === 'new-slide') {
          createRuntime(operation)
          setFilePath(null)
          setSaveNotice('')
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
          const options = {
            id: `slide-page-${operation.seq}`,
            name: operation.title ?? `幻灯片 ${index + 1}`,
            layoutPageId: 'layout-blank',
            showMasterSp: false,
          }
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
            const element = {
              id: `slide-text-${operation.seq}`,
              type: PageElementTypeEnum.Shape,
              transform: {
                left: operation.left ?? 120,
                top: operation.top ?? 100,
                width: Math.max(1, operation.width ?? 720),
                height: Math.max(1, operation.height ?? 80),
              },
              name: 'text',
              visible: true,
              selectable: true,
              shapeData: {
                shapeType: ShapeTypeEnum.Rect,
                isTextBox: true,
                fill: { fillType: ShapeFillEnum.NoFill },
                stroke: { lineStrokeType: ShapeLineTypeEnum.NoLine },
                shapeText: {
                  isRichText: false,
                  text: operation.text,
                  fontSize: operation.fontSize ?? 30,
                  color: operation.fontColor ?? '#333333',
                  bold: operation.bold ?? true,
                },
              },
            }
            slide.insertElement(element as any)
          } else if (operation.action === 'add-slide-shape') {
            const shapeType = (Object.values(ShapeTypeEnum) as string[]).includes(operation.shapeType)
              ? operation.shapeType as ShapeTypeEnum
              : ShapeTypeEnum.Rect
            const fillColor = operation.fillColor ?? '#4472C4'
            const strokeColor = operation.strokeColor ?? '#44546A'
            slide.insertShape({
              shapeType,
              transform: {
                left: operation.left ?? 120,
                top: operation.top ?? 120,
                width: Math.max(1, operation.width ?? 240),
                height: Math.max(1, operation.height ?? 120),
              },
              visible: true,
              selectable: operation.selectable ?? true,
              shapeData: {
                shapeType,
                fill: {
                  fillType: operation.fillColor === 'transparent' ? ShapeFillEnum.NoFill : ShapeFillEnum.SolidFill,
                  color: fillColor,
                  opacity: operation.opacity ?? 1,
                },
                stroke: {
                  lineStrokeType: operation.strokeColor === 'none' ? ShapeLineTypeEnum.NoLine : ShapeLineTypeEnum.SolidLine,
                  color: strokeColor,
                  width: operation.strokeWidth ?? 1.5,
                  opacity: operation.opacity ?? 1,
                },
              },
            })
          } else {
            const element = slide.getElementById(operation.elementId)
            if (element === null || (!('getText' in element) && (!('getType' in element) || element.getType() !== PageElementTypeEnum.Text))) {
              throw new Error(`找不到文本元素：${operation.elementId}`)
            }
            const editableElement = element as any
            const data = typeof editableElement.getData === 'function' ? editableElement.getData() as ISlideTextElement : {} as ISlideTextElement
            if ('getText' in element) {
              const shapeText = (element as any).getText()
              if (operation.text !== undefined) shapeText.setText(operation.text)
              if (operation.fontSize !== undefined) shapeText.setFontSize(operation.fontSize)
              if (operation.fontColor !== undefined) shapeText.setColor(operation.fontColor)
              if (operation.bold !== undefined) shapeText.setBold(operation.bold)
            } else if (operation.text !== undefined || operation.fontSize !== undefined || operation.fontColor !== undefined || operation.bold !== undefined) {
              const richText = runtime.univerAPI.newRichText().span(operation.text ?? data.text ?? '', {
                fontSize: operation.fontSize ?? data.textStyle?.fontSize ?? 30,
                color: operation.fontColor ?? data.textStyle?.color ?? '#333333',
                bold: operation.bold ?? data.textStyle?.bold ?? true,
              })
              editableElement.setRichText(richText)
            }
            const transform = editableElement.getTransform() ?? {}
            if (operation.left !== undefined || operation.top !== undefined) {
              if ('getText' in element) editableElement.setAbsolutePosition(operation.left ?? transform.left ?? 0, operation.top ?? transform.top ?? 0)
              else editableElement.setPosition(operation.left ?? transform.left ?? 0, operation.top ?? transform.top ?? 0)
            }
            if (operation.width !== undefined || operation.height !== undefined) {
              editableElement.setSize(Math.max(1, operation.width ?? transform.width ?? 1), Math.max(1, operation.height ?? transform.height ?? 1))
            }
          }
        }
        appliedRef.current.add(operation.seq)
      }
      const queuedIds = operations
        .map((operation) => operation.seq)
        .filter((id): id is string => typeof id === 'string' && appliedRef.current.has(id))
      if (queuedIds.length > 0 && !queuedPersistingRef.current) {
        queuedPersistingRef.current = true
        savingRef.current = true
        const previousSave = pendingSaveRef.current
        const commit = (async () => {
          await previousSave
          const presentation = runtimeRef.current?.presentation
          if (presentation === undefined) return
          const nextSnapshot = presentation.save()
          const serialized = JSON.stringify(nextSnapshot)
          await commitOperations(sessionId, 'slide', nextSnapshot, queuedIds)
          lastSavedRef.current = serialized
        })()
        pendingSaveRef.current = commit
        void commit.catch((reason) => {
          setError(`提交幻灯片队列操作失败：${reason instanceof Error ? reason.message : String(reason)}`)
        }).finally(() => {
          if (pendingSaveRef.current === commit) pendingSaveRef.current = null
          savingRef.current = false
          queuedPersistingRef.current = false
        })
      }
      setError(null)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
    }
    return undefined
  }, [hostLoaded, hostSnapshot, operations, layoutVersion, sessionId, locale])

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
        if (pendingSaveRef.current === save) {
          savingRef.current = false
          pendingSaveRef.current = null
        }
      })
      pendingSaveRef.current = save
    }, 800)
    return () => window.clearInterval(timer)
  }, [sessionId])

  useEffect(() => {
    if (!hostLoaded || !presentationVisible) return
    const connection = clientConnection.current
    if (connection === null) return
    for (const request of props.slideScreenshotRequests) {
      if (screenshotInFlightRef.current.has(request.id)) continue
      screenshotInFlightRef.current.add(request.id)
      void (async () => {
        try {
          const runtime = runtimeRef.current
          if (runtime === null) throw new Error('当前浏览器没有打开已渲染的 Univer Slide')
          const captured = await captureRenderedSlide(runtime, request.slideIndex, request.mode)
          await connection.rpc.call('/dsh-univer-create', 'slide-screenshot-result', {
            sessionId,
            screenshotId: request.id,
            ...captured,
          })
        } catch (reason) {
          await connection.rpc.call('/dsh-univer-create', 'slide-screenshot-result', {
            sessionId,
            screenshotId: request.id,
            error: reason instanceof Error ? reason.message : String(reason),
          }).catch(() => {})
        } finally {
          screenshotInFlightRef.current.delete(request.id)
        }
      })()
    }
  }, [sessionId, hostLoaded, presentationVisible, props.slideScreenshotRequests])

  useLayoutEffect(() => () => {
    const resident = runtimeRef.current
    const entry = residentRuntimes.get(residentRuntimeKey(sessionId, 'slide'))
    const connection = clientConnection.current
    if (resident !== null && entry?.runtime === resident && connection !== null) {
      queueResidentFinalSave(entry, async () => {
        const snapshot = resident.presentation.save()
        await connection.rpc.call('/dsh-univer-create', 'save', { sessionId, unitType: 'slide', snapshot })
      })
    }
    slideResizeTimersRef.current.forEach((timer) => window.clearTimeout(timer))
    slideResizeTimersRef.current = []
    parkResidentRuntime(sessionId, 'slide', resident)
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
    disposeResidentRuntime(sessionId, 'slide', runtimeRef.current)
    runtimeRef.current = null
    appliedRef.current.clear()
    hydratedFromHostRef.current = true
    lastSavedRef.current = null
    lastFileSnapshotRef.current = null
    openedPresentationSessions.add(sessionId)
    setError(null)
    setTitle(nextTitle)
    setFilePath(nextFilePath)
    setSaveNotice('')
    setHostSnapshot(nextSnapshot)
    await persistHostSnapshot(nextSnapshot, nextFilePath)
  }

  const newPresentation = async () => {
    const presentation = runtimeRef.current?.presentation
    if (presentation !== undefined && !confirmOpenWithUnsavedChanges(
      presentation.save(),
      lastFileSnapshotRef.current,
      ui.confirmDiscardChangesForNew,
    )) return
    const blankSnapshot = presentationData({ action: 'new-slide', seq: Date.now() }, locale, ui.defaultSlideTitle)
    try {
      await replacePresentation(blankSnapshot, ui.defaultSlideTitle, null)
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

  const openPresentationFile = () => {
    const presentation = runtimeRef.current?.presentation
    if (presentation !== undefined && !confirmOpenWithUnsavedChanges(
      presentation.save(),
      lastFileSnapshotRef.current,
      ui.confirmDiscardChanges,
    )) return
    fileInputRef.current?.click()
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
        if (chosenPath !== undefined && isWorkspaceFileExistsError(reason) && window.confirm(ui.confirmOverwrite(targetPath))) {
          savedPath = await saveUnitFile(sessionId, 'slide', nextSnapshot, targetPath, true)
        } else {
          throw reason
        }
      }
      setFilePath(savedPath)
      setSaveDialogOpen(false)
      setSaveNotice(ui.savedTo(savedPath))
      lastFileSnapshotRef.current = JSON.stringify(nextSnapshot)
      lastSavedRef.current = lastFileSnapshotRef.current
    } catch (reason) {
      setError(`保存 PPTX 失败：${reason instanceof Error ? reason.message : String(reason)}`)
    } finally {
      setExporting(false)
    }
  }

  const actionButtons = (
    <>
      <button className="dsh-univer-create-action" type="button" disabled={importing || exporting || !hostLoaded} onClick={openPresentationFile}>
        {importing ? ui.opening : ui.open}
      </button>
      <button className="dsh-univer-create-action dsh-univer-create-action--primary" type="button" disabled={importing || exporting || !hostLoaded} onClick={() => void newPresentation()}>
        {ui.create}
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
            {saveNotice || (filePath !== null ? ui.fileStatus(filePath) : operations.length > 0 ? ui.syncedSlideOperations(operations.length) : ui.newFileHint)}
          </span>
          {actionButtons}
          <button className="dsh-univer-create-action" type="button" disabled={importing || exporting || !hostLoaded} onClick={() => void saveAsPptx()}>
            {exporting ? ui.saving : ui.save}
          </button>
        </div>
      )}
      <div ref={containerRef} className="dsh-univer-create-container dsh-univer-create-container--slides" />
       
      {presentationVisible && (
         <aside className="dsh-univer-create-chat-overlay" aria-label={ui.chatActivity}>
           <div className="dsh-univer-create-chat-overlay__header">
              <div className="dsh-univer-create-chat-overlay__title">{ui.chatActivity}</div>
              <button
                className="dsh-univer-create-chat-overlay__toggle"
                type="button"
                aria-label={chatOverlayVisible ? ui.closeChatActivity : ui.openChatActivity}
                aria-expanded={chatOverlayVisible}
                onClick={() => setChatOverlayVisible((visible) => !visible)}
              >
                {chatOverlayVisible ? ui.hide : ui.show}
              </button>
            </div>
           <div ref={chatStreamRef} className="dsh-univer-create-chat-overlay__stream" hidden={!chatOverlayVisible}>
             {chatLines.map((line, index) => (
               <div className="dsh-univer-create-chat-overlay__line" key={`${line.role}-${index}-${line.text.slice(0, 16)}`}>
                 <strong>{line.role}</strong><span>{line.text}</span>
               </div>
             ))}
             {partialText && <div className="dsh-univer-create-chat-overlay__line dsh-univer-create-chat-overlay__line--live"><strong>AI</strong><span>{partialText}</span></div>}
           </div>
         </aside>
       )}
       {saveDialogOpen && <SavePathDialog extension="pptx" suggestedName={title} busy={exporting} onCancel={() => setSaveDialogOpen(false)} onSave={(path) => void saveAsPptx(path)} />}
      {hostLoaded && !presentationVisible && <div className="dsh-univer-create-welcome" aria-label={ui.startSlide}>{actionButtons}</div>}
      {error !== null && <div className="dsh-univer-create-error" role="alert">{ui.slideError}{error}{!hostLoaded && ui.hostRecoveryHint}</div>}
    </section>
  )
}

const selectedUnitBySession = new Map<string, UniverUnitType>()

function UniverView(props: ConvViewProps) {
  const language = useDshLanguage()
  const ui = UNIVER_CREATE_LOCALES[language]
  const { nodes } = useConversationFeed(props)
  const sessionRunning = props.useSession((session) => session.running)
  const shellRef = useRef<HTMLElement>(null)
  const [tasks, setTasks] = useState<UniverTasks>(EMPTY_UNIVER_TASKS)
  const [drainingTasks, setDrainingTasks] = useState(true)
  const emptyDrainPollsRef = useRef(0)
  useEffect(() => {
    if (!sessionRunning) return
    emptyDrainPollsRef.current = 0
    setDrainingTasks(true)
  }, [sessionRunning])
  const suggestedUnit = useMemo<UniverUnitType>(() => {
    for (let index = nodes.length - 1; index >= 0; index -= 1) {
      const node = nodes[index]!
      if (node.kind !== 'tool-result' || node.call === null) continue
      if (node.call.name === 'univer_execute_code') {
        try {
          const args = JSON.parse(node.call.argsRaw) as { unitType?: unknown }
          if (args.unitType === 'sheet' || args.unitType === 'doc' || args.unitType === 'slide') return args.unitType
        } catch {}
      }
      if (SLIDE_TOOL_NAMES.has(node.call.name)) return 'slide'
      if (DOC_TOOL_NAMES.has(node.call.name)) return 'doc'
      if (SHEET_TOOL_NAMES.has(node.call.name)) return 'sheet'
    }
    return 'sheet'
  }, [nodes])
  const [unitType, setUnitType] = useState<UniverUnitType>(() => selectedUnitBySession.get(props.sessionId) ?? suggestedUnit)
  const activeUnitRef = useRef(unitType)
  useEffect(() => { activeUnitRef.current = unitType }, [unitType])
  // Apply one durable operation per commit. This keeps every persisted snapshot
  // aligned with exactly one queue transition and prevents a poison item from
  // causing an earlier, successfully-mutated prefix to be replayed.
  const dispatchOperations = useMemo(() => tasks.operations.slice(0, 1), [tasks.operations])

  useEffect(() => {
    if (!sessionRunning && !drainingTasks) {
      setTasks(EMPTY_UNIVER_TASKS)
      return
    }
    let disposed = false
    let polling = false
    const pollTasks = async () => {
      const connection = clientConnection.current
      const shell = shellRef.current
      if (
        connection === null
        || polling
        || document.visibilityState !== 'visible'
        || shell === null
        || shell.getClientRects().length === 0
      ) return
      polling = true
      try {
        const response = await connection.rpc.call('/dsh-univer-create', 'tasks', {
          sessionId: props.sessionId,
          clientId: univerCodeClientId,
          activeUnit: activeUnitRef.current,
        })
        if (!response.ok || disposed || response.value === null || typeof response.value !== 'object') return
        const value = response.value as Record<string, unknown>
        const nextTasks: UniverTasks = {
          operations: Array.isArray(value.operations)
            ? value.operations.map(parseQueuedUniverOperation).filter((item): item is QueuedUniverOperation => item !== null)
            : [],
          codeRequests: Array.isArray(value.codeRequests)
            ? value.codeRequests.map(parseUniverCodeRequest).filter((item): item is UniverCodeRequest => item !== null)
            : [],
          sheetScreenshotRequests: Array.isArray(value.sheetScreenshotRequests)
            ? value.sheetScreenshotRequests.map(parseSheetScreenshotRequest).filter((item): item is SheetScreenshotRequest => item !== null)
            : [],
          docScreenshotRequests: Array.isArray(value.docScreenshotRequests)
            ? value.docScreenshotRequests.map(parseDocScreenshotRequest).filter((item): item is DocScreenshotRequest => item !== null)
            : [],
          slideScreenshotRequests: Array.isArray(value.slideScreenshotRequests)
            ? value.slideScreenshotRequests.map(parseSlideScreenshotRequest).filter((item): item is SlideScreenshotRequest => item !== null)
            : [],
        }
        setTasks(nextTasks)
        if (sessionRunning) {
          emptyDrainPollsRef.current = 0
        } else {
          const remoteEmpty = nextTasks.operations.length === 0
            && nextTasks.codeRequests.length === 0
            && nextTasks.sheetScreenshotRequests.length === 0
            && nextTasks.docScreenshotRequests.length === 0
            && nextTasks.slideScreenshotRequests.length === 0
          const sessionQueuePrefix = `${props.sessionId}:`
          const localBusy = residentCodeRequestsInFlight.size > 0
            || [...residentRuntimeTaskQueues.keys()].some((key) => key.startsWith(sessionQueuePrefix))
          if (remoteEmpty && !localBusy) {
            emptyDrainPollsRef.current += 1
            if (emptyDrainPollsRef.current >= 2) setDrainingTasks(false)
          } else {
            emptyDrainPollsRef.current = 0
            setDrainingTasks(true)
          }
        }
        const target = nextTasks.operations[0]?.unitType
          ?? nextTasks.codeRequests[0]?.unitType
          ?? (nextTasks.sheetScreenshotRequests.length > 0 ? 'sheet' : undefined)
          ?? (nextTasks.docScreenshotRequests.length > 0 ? 'doc' : undefined)
          ?? (nextTasks.slideScreenshotRequests.length > 0 ? 'slide' : undefined)
        if (target !== undefined) {
          selectedUnitBySession.set(props.sessionId, target)
          activeUnitRef.current = target
          setUnitType(target)
        }
      } finally {
        polling = false
      }
    }
    setTasks(EMPTY_UNIVER_TASKS)
    void pollTasks().catch(() => {})
    const timer = window.setInterval(() => void pollTasks().catch(() => {}), 500)
    return () => {
      disposed = true
      window.clearInterval(timer)
      setTasks(EMPTY_UNIVER_TASKS)
    }
  }, [props.sessionId, sessionRunning, drainingTasks])

  useEffect(() => {
    for (const request of tasks.codeRequests) void executeResidentCodeRequest(props.sessionId, request)
  }, [props.sessionId, tasks.codeRequests])

  useEffect(() => {
    const remembered = selectedUnitBySession.get(props.sessionId)
    setUnitType(remembered ?? suggestedUnit)
  }, [props.sessionId, suggestedUnit])

  const selectUnit = (next: UniverUnitType) => {
    selectedUnitBySession.set(props.sessionId, next)
    setUnitType(next)
  }

  return (
    <section ref={shellRef} className="dsh-univer-create-shell" aria-label={ui.shellLabel}>
      <nav className="dsh-univer-create-product-bar" aria-label={ui.selectEditor}>
        <button className={`dsh-univer-create-product-tab${unitType === 'sheet' ? ' is-active' : ''}`} type="button" aria-pressed={unitType === 'sheet'} onClick={() => selectUnit('sheet')}>Sheet</button>
        <button className={`dsh-univer-create-product-tab${unitType === 'doc' ? ' is-active' : ''}`} type="button" aria-pressed={unitType === 'doc'} onClick={() => selectUnit('doc')}>Doc</button>
        <button className={`dsh-univer-create-product-tab${unitType === 'slide' ? ' is-active' : ''}`} type="button" aria-pressed={unitType === 'slide'} onClick={() => selectUnit('slide')}>Slide</button>
      </nav>
      <div className={`dsh-univer-create-product-pane${unitType === 'sheet' ? ' is-active' : ''}`} aria-hidden={unitType !== 'sheet'}>
        <SheetProductView key={`${props.sessionId}:sheet`} {...props} queuedOperations={dispatchOperations} codeRequests={tasks.codeRequests} sheetScreenshotRequests={tasks.sheetScreenshotRequests} docScreenshotRequests={tasks.docScreenshotRequests} slideScreenshotRequests={tasks.slideScreenshotRequests} />
      </div>
      <div className={`dsh-univer-create-product-pane${unitType === 'doc' ? ' is-active' : ''}`} aria-hidden={unitType !== 'doc'}>
        <DocProductView key={`${props.sessionId}:doc`} {...props} queuedOperations={dispatchOperations} codeRequests={tasks.codeRequests} sheetScreenshotRequests={tasks.sheetScreenshotRequests} docScreenshotRequests={tasks.docScreenshotRequests} slideScreenshotRequests={tasks.slideScreenshotRequests} />
      </div>
      <div className={`dsh-univer-create-product-pane${unitType === 'slide' ? ' is-active' : ''}`} aria-hidden={unitType !== 'slide'}>
        <SlideProductView key={`${props.sessionId}:slide`} {...props} queuedOperations={dispatchOperations} codeRequests={tasks.codeRequests} sheetScreenshotRequests={tasks.sheetScreenshotRequests} docScreenshotRequests={tasks.docScreenshotRequests} slideScreenshotRequests={tasks.slideScreenshotRequests} />
      </div>
    </section>
  )
}

function UniverViewHost(props: ConvViewProps) {
  const hostRef = useRef<HTMLElement>(null)
  useLayoutEffect(() => {
    const host = hostRef.current
    if (host === null) return
    const shell = getResidentViewShell(props.sessionId)
    host.appendChild(shell)
    window.dispatchEvent(new Event('resize'))
    return () => {
      if (residentViewShells.get(props.sessionId) === shell) getResidentRuntimeParkingHost().appendChild(shell)
    }
  }, [props.sessionId])
  return <section ref={hostRef} className="dsh-univer-create-view-host" aria-label="Univer" />
}

export const inject = ['slots', 'connection']

export function apply(ctx: ClientContext): void {
  clientConnection.current = ctx.get('connection') as unknown as ConnectionHandle
  ctx.effect(() => () => {
    currentResidentSessionId = null
    residentCodeRequestsInFlight.clear()
    residentRuntimeTaskQueues.clear()
    for (const entry of residentRuntimes.values()) {
      entry.runtime.univer.dispose()
      entry.runtime.mount.remove()
    }
    residentRuntimes.clear()
    for (const shell of residentViewShells.values()) shell.remove()
    residentViewShells.clear()
    residentRuntimeParkingHost?.remove()
    residentRuntimeParkingHost = null
    clientConnection.current = null
  }, 'dsh-univer-create: client connection and resident runtimes')
  ctx.slots.inject('conversation.session.header.utilities', () => ctx.slots.register({
    name: 'conversation.session.header.utilities',
    id: 'univer-runtime-session-tracker',
    order: 10_000,
  }, UniverSessionRuntimeTracker))
  ctx.slots.inject('conversation.view', () => ctx.slots.register({
    name: 'conversation.view',
    id: 'univer-create',
    label: 'Univer',
    order: 20,
  }, UniverViewHost))
}
