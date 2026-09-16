import ChartUIEnUS from '@univerjs-pro/chart-ui/locale/en-US'
import ChartUIZhCN from '@univerjs-pro/chart-ui/locale/zh-CN'
import { UniverLicensePlugin } from '@univerjs-pro/license'
import { UniverSheetsChartPlugin } from '@univerjs-pro/sheets-chart'
import { UniverSheetsChartUIPlugin } from '@univerjs-pro/sheets-chart-ui'
import SheetsChartUIEnUS from '@univerjs-pro/sheets-chart-ui/locale/en-US'
import SheetsChartUIZhCN from '@univerjs-pro/sheets-chart-ui/locale/zh-CN'
import { UniverSheetsPrintPlugin } from '@univerjs-pro/sheets-print'
import SheetsPrintEnUS from '@univerjs-pro/sheets-print/locale/en-US'
import SheetsPrintZhCN from '@univerjs-pro/sheets-print/locale/zh-CN'
import { CommandType, LocaleType, mergeLocales, Univer, UniverInstanceType } from '@univerjs/core'
import { UNIVER_LICENSE } from 'virtual:dsh-univer-license'
import { FUniver } from '@univerjs/core/facade'
import DesignEnUS from '@univerjs/design/locale/en-US'
import DesignZhCN from '@univerjs/design/locale/zh-CN'
import { UniverDocsPlugin } from '@univerjs/docs'
import { UniverDocsUIPlugin } from '@univerjs/docs-ui'
import DocsUIEnUS from '@univerjs/docs-ui/locale/en-US'
import DocsUIZhCN from '@univerjs/docs-ui/locale/zh-CN'
import { UniverDrawingPlugin } from '@univerjs/drawing'
import { UniverDrawingUIPlugin } from '@univerjs/drawing-ui'
import DrawingUIEnUS from '@univerjs/drawing-ui/locale/en-US'
import DrawingUIZhCN from '@univerjs/drawing-ui/locale/zh-CN'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverRenderEnginePlugin } from '@univerjs/engine-render'
import FindReplaceEnUS from '@univerjs/find-replace/locale/en-US'
import FindReplaceZhCN from '@univerjs/find-replace/locale/zh-CN'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import SheetsEnUS from '@univerjs/sheets/locale/en-US'
import SheetsZhCN from '@univerjs/sheets/locale/zh-CN'
import { UniverSheetsConditionalFormattingUIPlugin } from '@univerjs/sheets-conditional-formatting-ui'
import SheetsConditionalFormattingUIEnUS from '@univerjs/sheets-conditional-formatting-ui/locale/en-US'
import SheetsConditionalFormattingUIZhCN from '@univerjs/sheets-conditional-formatting-ui/locale/zh-CN'
import { UniverSheetsCrosshairHighlightPlugin } from '@univerjs/sheets-crosshair-highlight'
import SheetsCrosshairHighlightEnUS from '@univerjs/sheets-crosshair-highlight/locale/en-US'
import SheetsCrosshairHighlightZhCN from '@univerjs/sheets-crosshair-highlight/locale/zh-CN'
import { UniverSheetsDataValidationPlugin } from '@univerjs/sheets-data-validation'
import { UniverSheetsDataValidationUIPlugin } from '@univerjs/sheets-data-validation-ui'
import SheetsDataValidationUIEnUS from '@univerjs/sheets-data-validation-ui/locale/en-US'
import SheetsDataValidationUIZhCN from '@univerjs/sheets-data-validation-ui/locale/zh-CN'
import { UniverSheetsDrawingPlugin } from '@univerjs/sheets-drawing'
import { UniverSheetsDrawingUIPlugin } from '@univerjs/sheets-drawing-ui'
import SheetsDrawingUIEnUS from '@univerjs/sheets-drawing-ui/locale/en-US'
import SheetsDrawingUIZhCN from '@univerjs/sheets-drawing-ui/locale/zh-CN'
import { UniverSheetsFilterUIPlugin } from '@univerjs/sheets-filter-ui'
import SheetsFilterUIEnUS from '@univerjs/sheets-filter-ui/locale/en-US'
import SheetsFilterUIZhCN from '@univerjs/sheets-filter-ui/locale/zh-CN'
import { UniverSheetsFindReplacePlugin } from '@univerjs/sheets-find-replace'
import { UniverSheetsFormulaUIPlugin } from '@univerjs/sheets-formula-ui'
import SheetsFormulaUIEnUS from '@univerjs/sheets-formula-ui/locale/en-US'
import SheetsFormulaUIZhCN from '@univerjs/sheets-formula-ui/locale/zh-CN'
import { UniverSheetsHyperLinkUIPlugin } from '@univerjs/sheets-hyper-link-ui'
import SheetsHyperLinkUIEnUS from '@univerjs/sheets-hyper-link-ui/locale/en-US'
import SheetsHyperLinkUIZhCN from '@univerjs/sheets-hyper-link-ui/locale/zh-CN'
import { UniverSheetsNoteUIPlugin } from '@univerjs/sheets-note-ui'
import SheetsNoteUIEnUS from '@univerjs/sheets-note-ui/locale/en-US'
import SheetsNoteUIZhCN from '@univerjs/sheets-note-ui/locale/zh-CN'
import { UniverSheetsNumfmtUIPlugin } from '@univerjs/sheets-numfmt-ui'
import SheetsNumfmtUIEnUS from '@univerjs/sheets-numfmt-ui/locale/en-US'
import SheetsNumfmtUIZhCN from '@univerjs/sheets-numfmt-ui/locale/zh-CN'
import { UniverSheetsSortUIPlugin } from '@univerjs/sheets-sort-ui'
import SheetsSortUIEnUS from '@univerjs/sheets-sort-ui/locale/en-US'
import SheetsSortUIZhCN from '@univerjs/sheets-sort-ui/locale/zh-CN'
import { UniverSheetsTableUIPlugin } from '@univerjs/sheets-table-ui'
import SheetsTableUIEnUS from '@univerjs/sheets-table-ui/locale/en-US'
import SheetsTableUIZhCN from '@univerjs/sheets-table-ui/locale/zh-CN'
import { UniverSheetsThreadCommentUIPlugin } from '@univerjs/sheets-thread-comment-ui'
import SheetsThreadCommentUIEnUS from '@univerjs/sheets-thread-comment-ui/locale/en-US'
import SheetsThreadCommentUIZhCN from '@univerjs/sheets-thread-comment-ui/locale/zh-CN'
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui'
import SheetsUIEnUS from '@univerjs/sheets-ui/locale/en-US'
import SheetsUIZhCN from '@univerjs/sheets-ui/locale/zh-CN'
import { UniverThreadCommentUIPlugin } from '@univerjs/thread-comment-ui'
import ThreadCommentUIEnUS from '@univerjs/thread-comment-ui/locale/en-US'
import ThreadCommentUIZhCN from '@univerjs/thread-comment-ui/locale/zh-CN'
import { UniverUIPlugin } from '@univerjs/ui'
import UIEnUS from '@univerjs/ui/locale/en-US'
import UIZhCN from '@univerjs/ui/locale/zh-CN'

import '@univerjs/sheets/facade'
import '@univerjs-pro/engine-chart/facade'
import '@univerjs-pro/chart-ui/facade'
import '@univerjs-pro/sheets-chart/facade'
import { isPersistableWorkbookMutation } from '../workbook-drafts.js'
import '../styles.css'

import '@univerjs/design/lib/index.css'
import '@univerjs/ui/lib/index.css'
import '@univerjs/docs-ui/lib/index.css'
import '@univerjs/drawing-ui/lib/index.css'
import '@univerjs/sheets-ui/lib/index.css'
import '@univerjs/sheets-drawing-ui/lib/index.css'
import '@univerjs/sheets-formula-ui/lib/index.css'
import '@univerjs/sheets-numfmt-ui/lib/index.css'
import '@univerjs/sheets-sort-ui/lib/index.css'
import '@univerjs/sheets-filter-ui/lib/index.css'
import '@univerjs/find-replace/lib/index.css'
import '@univerjs/sheets-conditional-formatting-ui/lib/index.css'
import '@univerjs/sheets-data-validation-ui/lib/index.css'
import '@univerjs/sheets-hyper-link-ui/lib/index.css'
import '@univerjs/sheets-note-ui/lib/index.css'
import '@univerjs/sheets-table-ui/lib/index.css'
import '@univerjs/thread-comment-ui/lib/index.css'
import '@univerjs/sheets-crosshair-highlight/lib/index.css'
import '@univerjs-pro/chart-ui/lib/index.css'
import '@univerjs-pro/sheets-chart-ui/lib/index.css'
import '@univerjs-pro/sheets-print/lib/index.css'

export function createSheetRuntime(container: HTMLElement, snapshot: unknown, language: 'en' | 'zh' = 'en') {
  const univer = new Univer({
    locale: language === 'zh' ? LocaleType.ZH_CN : LocaleType.EN_US,
    locales: {
      [LocaleType.EN_US]: mergeLocales(
        DesignEnUS,
        UIEnUS,
        DocsUIEnUS,
        DrawingUIEnUS,
        SheetsEnUS,
        SheetsUIEnUS,
        SheetsDrawingUIEnUS,
        SheetsFormulaUIEnUS,
        SheetsNumfmtUIEnUS,
        SheetsFilterUIEnUS,
        SheetsConditionalFormattingUIEnUS,
        SheetsDataValidationUIEnUS,
        SheetsSortUIEnUS,
        FindReplaceEnUS,
        ThreadCommentUIEnUS,
        SheetsThreadCommentUIEnUS,
        SheetsNoteUIEnUS,
        SheetsHyperLinkUIEnUS,
        SheetsTableUIEnUS,
        SheetsCrosshairHighlightEnUS,
        ChartUIEnUS,
        SheetsChartUIEnUS,
        SheetsPrintEnUS,
      ),
      [LocaleType.ZH_CN]: mergeLocales(
        DesignZhCN,
        UIZhCN,
        DocsUIZhCN,
        DrawingUIZhCN,
        SheetsZhCN,
        SheetsUIZhCN,
        SheetsDrawingUIZhCN,
        SheetsFormulaUIZhCN,
        SheetsNumfmtUIZhCN,
        SheetsFilterUIZhCN,
        SheetsConditionalFormattingUIZhCN,
        SheetsDataValidationUIZhCN,
        SheetsSortUIZhCN,
        FindReplaceZhCN,
        ThreadCommentUIZhCN,
        SheetsThreadCommentUIZhCN,
        SheetsNoteUIZhCN,
        SheetsHyperLinkUIZhCN,
        SheetsTableUIZhCN,
        SheetsCrosshairHighlightZhCN,
        ChartUIZhCN,
        SheetsChartUIZhCN,
        SheetsPrintZhCN,
      ),
    },
  })

  univer.registerPlugin(UniverRenderEnginePlugin)
  univer.registerPlugin(UniverLicensePlugin, { license: UNIVER_LICENSE })
  univer.registerPlugin(UniverFormulaEnginePlugin)
  univer.registerPlugin(UniverUIPlugin, { container })
  univer.registerPlugin(UniverDocsPlugin)
  univer.registerPlugin(UniverDocsUIPlugin)
  univer.registerPlugin(UniverDrawingPlugin)
  univer.registerPlugin(UniverDrawingUIPlugin)
  univer.registerPlugin(UniverSheetsPlugin)
  univer.registerPlugin(UniverSheetsUIPlugin)
  univer.registerPlugin(UniverSheetsDrawingPlugin)
  univer.registerPlugin(UniverSheetsDrawingUIPlugin)
  univer.registerPlugin(UniverSheetsFormulaUIPlugin)
  univer.registerPlugin(UniverSheetsNumfmtUIPlugin)
  univer.registerPlugin(UniverSheetsDataValidationPlugin)
  univer.registerPlugin(UniverSheetsDataValidationUIPlugin)
  univer.registerPlugin(UniverSheetsConditionalFormattingUIPlugin)
  univer.registerPlugin(UniverSheetsFilterUIPlugin)
  univer.registerPlugin(UniverSheetsSortUIPlugin)
  univer.registerPlugin(UniverSheetsFindReplacePlugin)
  univer.registerPlugin(UniverSheetsHyperLinkUIPlugin)
  univer.registerPlugin(UniverThreadCommentUIPlugin)
  univer.registerPlugin(UniverSheetsThreadCommentUIPlugin)
  univer.registerPlugin(UniverSheetsTableUIPlugin)
  univer.registerPlugin(UniverSheetsNoteUIPlugin)
  univer.registerPlugin(UniverSheetsCrosshairHighlightPlugin)
  univer.registerPlugin(UniverSheetsChartPlugin)
  univer.registerPlugin(UniverSheetsChartUIPlugin)
  univer.registerPlugin(UniverSheetsPrintPlugin)

  univer.createUnit(UniverInstanceType.UNIVER_SHEET, snapshot as never)

  // WorkbookDataModel.getSnapshot() only contains the core workbook model. The
  // facade save() path delegates to ResourceLoaderService.saveUnit(), which also
  // serializes plugin resources such as inserted Sheet drawings/images.
  type RuntimeWorksheet = { getSheetId: () => string; getSheetName: () => string }
  type RuntimeWorkbook = {
    save: () => unknown
    getId: () => string
    getSheets: () => RuntimeWorksheet[]
    getSheetBySheetId: (sheetId: string) => RuntimeWorksheet | null
    getSheetByName: (name: string) => RuntimeWorksheet | null
  }
  const univerAPI = FUniver.newAPI(univer) as FUniver & {
    getActiveWorkbook: () => RuntimeWorkbook | null
  }
  const workbook = univerAPI.getActiveWorkbook()
  if (!workbook) {
    univer.dispose()
    throw new Error('failed to create active Sheet workbook')
  }

  type RuntimeMutation = { id: string; params: Record<string, unknown>; subUnitName?: string }

  const remapMutationTargets = (mutation: RuntimeMutation): Record<string, unknown> => {
    const params = structuredClone(mutation.params)
    const currentUnitId = workbook.getId()
    const visited = new Set<object>()
    const remapUnitIds = (value: unknown) => {
      if (value === null || typeof value !== 'object' || visited.has(value)) return
      visited.add(value)
      if (Array.isArray(value)) {
        value.forEach(remapUnitIds)
        return
      }
      for (const [key, child] of Object.entries(value as Record<string, unknown>)) {
        if (key === 'unitId' && typeof child === 'string') (value as Record<string, unknown>)[key] = currentUnitId
        else remapUnitIds(child)
      }
    }
    remapUnitIds(params)

    const oldSheetId = params.subUnitId
    if (typeof oldSheetId === 'string' && workbook.getSheetBySheetId(oldSheetId) === null) {
      const byName = mutation.subUnitName ? workbook.getSheetByName(mutation.subUnitName) : null
      const fallback = workbook.getSheets().length === 1 ? workbook.getSheets()[0] : null
      const sheet = byName ?? fallback
      if (sheet) params.subUnitId = sheet.getSheetId()
    }
    return params
  }

  const replayMutations = async (mutations: RuntimeMutation[]) => {
    for (const mutation of mutations) {
      if (!isPersistableWorkbookMutation(mutation)) continue
      const result = await univerAPI.executeCommand(mutation.id, remapMutationTargets(mutation), {
        onlyLocal: true,
        dshDraftReplay: true,
      })
      if (result === false) throw new Error(`failed to replay workbook mutation: ${mutation.id}`)
    }
  }

  const onMutation = (listener: (mutation: RuntimeMutation) => void) => univerAPI.addEvent(
    univerAPI.Event.CommandExecuted,
    (event) => {
      if (event.type !== CommandType.MUTATION || event.options?.dshDraftReplay === true) return
      const params = (event.params ?? {}) as Record<string, unknown>
      const subUnitId = params.subUnitId
      const mutation: RuntimeMutation = {
        id: event.id,
        params,
        subUnitName: typeof subUnitId === 'string' ? workbook.getSheetBySheetId(subUnitId)?.getSheetName() : undefined,
      }
      if (isPersistableWorkbookMutation(mutation)) listener(mutation)
    },
  )

  return { univer, workbook, replayMutations, onMutation }
}

;(globalThis as any).__DSH_WORKSPACE_FILE_VIEWER__ = {
  ...(globalThis as any).__DSH_WORKSPACE_FILE_VIEWER__,
  createSheetRuntime,
}
