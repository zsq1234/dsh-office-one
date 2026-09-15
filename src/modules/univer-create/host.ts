import type { Context } from '@deepseek-ai/cordis'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { createStandardApiReference } from '@univer-cli/api-reference'
import { defineDomain, domainTable } from '@deepseek-ai/dsh-storage-domain'
import { chmod, lstat, realpath, rename, stat, unlink, writeFile } from 'node:fs/promises'
import { randomUUID } from 'node:crypto'
import { basename, dirname, extname, isAbsolute, relative, resolve, sep } from 'node:path'
import { z } from 'zod'

export const name = 'dsh-univer-create'
export const inject = ['tools', 'connection', 'storageDomain', 'workspaceRegistry', 'attachments']

const queuedSheetOperationSchema = z.object({
  id: z.string(),
  operation: z.record(z.string(), z.unknown()),
})

const workbookRowSchema = z.object({
  snapshot: z.unknown(),
  updatedAt: z.number(),
  filePath: z.string().nullable().optional(),
  queuedSheetOperations: z.array(queuedSheetOperationSchema).optional(),
})

interface SlideScreenshotRequest {
  id: string
  slideIndex: number
  mode: 'slide' | 'editor'
  createdAt: number
}

interface DocScreenshotRequest {
  id: string
  mode: 'document' | 'editor'
  createdAt: number
}

interface SlideScreenshotResult {
  dataUrl?: string
  width?: number
  height?: number
  error?: string
}

interface UniverCodeRequest {
  id: string
  unitType: 'sheet' | 'doc' | 'slide'
  code: string
  createdAt: number
  claimedBy?: string
}

interface UniverCodeResult {
  result?: string
  logs?: string[]
  error?: string
}

const workbookDomainSpec = defineDomain({
  name: 'dsh_univer_sheet',
  version: 1,
  tables: {
    workbooks: domainTable<string, z.infer<typeof workbookRowSchema>>(workbookRowSchema),
  },
})

const successSchema = {
  type: 'object' as const,
  properties: {
    ok: { type: 'boolean', required: true },
    action: { type: 'string', required: true },
    message: { type: 'string', required: true },
  },
  additionalProperties: false,
} as const

const output = {
  schema: successSchema,
  render: (_args: unknown, value: { ok: boolean; action: string; message: string }) => [
    { type: 'text' as const, text: value.message },
  ],
}

const imageValueSchema = {
  type: 'object',
  required: true,
  additionalProperties: false,
  properties: {
    attachmentId: { type: 'string' as const, required: true },
    mediaType: { type: 'string' as const, enum: ['image/png'] as const, required: true },
    bytes: { type: 'integer' as const, required: true },
    width: { type: 'integer' as const, required: true },
    height: { type: 'integer' as const, required: true },
    name: { type: 'string' as const },
  },
} as const

type ScreenshotImageValue = {
  attachmentId: string
  mediaType: 'image/png'
  bytes: number
  width: number
  height: number
  name?: string
}

type CellValue = string | number | boolean | null
type QueuedSheetOperation = z.infer<typeof queuedSheetOperationSchema>

type SheetAgent = {
  id: string
  session?: { header?: { parentSession?: string; origin?: string } }
}

function sheetOwnerSessionId(agent: SheetAgent | undefined): string | undefined {
  if (agent === undefined) return undefined
  return agent.session?.header?.origin === 'subagent' && typeof agent.session.header.parentSession === 'string'
    ? agent.session.header.parentSession
    : agent.id
}

function isSubagent(agent: SheetAgent | undefined): boolean {
  return agent?.session?.header?.origin === 'subagent' && typeof agent.session.header.parentSession === 'string'
}

function columnIndex(label: string): number {
  let index = 0
  for (const character of label.toUpperCase()) index = index * 26 + character.charCodeAt(0) - 64
  return index - 1
}

function parseA1Range(range: string): { startRow: number; endRow: number; startColumn: number; endColumn: number } {
  const match = /^\s*([A-Za-z]+)([1-9]\d*)(?::([A-Za-z]+)([1-9]\d*))?\s*$/.exec(range)
  if (match === null) throw new Error(`无效的 A1 范围：${range}`)
  const startColumn = columnIndex(match[1]!)
  const startRow = Number(match[2]) - 1
  const endColumn = match[3] === undefined ? startColumn : columnIndex(match[3])
  const endRow = match[4] === undefined ? startRow : Number(match[4]) - 1
  if (endRow < startRow || endColumn < startColumn) throw new Error(`无效的 A1 范围：${range}`)
  if ((endRow - startRow + 1) * (endColumn - startColumn + 1) > 10_000) throw new Error('单次最多读取 10000 个单元格')
  return { startRow, endRow, startColumn, endColumn }
}

function readSnapshotRange(snapshot: unknown, range: string, requestedSheetName?: string): { sheetName: string; values: CellValue[][] } {
  if (snapshot === null || typeof snapshot !== 'object') throw new Error('当前会话没有可读取的工作簿')
  const workbook = snapshot as Record<string, any>
  const sheets = workbook.sheets
  if (sheets === null || typeof sheets !== 'object') throw new Error('当前工作簿没有工作表')
  const sheetList = Object.values(sheets) as Array<Record<string, any>>
  const activeId = typeof workbook.activeSheetId === 'string'
    ? workbook.activeSheetId
    : Array.isArray(workbook.sheetOrder) ? workbook.sheetOrder[0] : undefined
  const sheet = requestedSheetName === undefined
    ? (typeof activeId === 'string' ? sheets[activeId] : undefined) ?? sheetList[0]
    : sheetList.find((item) => item.name === requestedSheetName)
  if (sheet === undefined) throw new Error(`找不到工作表：${requestedSheetName ?? '(active)'}`)

  const bounds = parseA1Range(range)
  const cellData = sheet.cellData !== null && typeof sheet.cellData === 'object' ? sheet.cellData : {}
  const values: CellValue[][] = []
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    const outputRow: CellValue[] = []
    const rowData = cellData[row] ?? cellData[String(row)] ?? {}
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      const cell = rowData[column] ?? rowData[String(column)]
      const value = cell !== null && typeof cell === 'object'
        ? (typeof cell.f === 'string' ? cell.f : cell.v)
        : undefined
      outputRow.push(typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' ? value : null)
    }
    values.push(outputRow)
  }
  return { sheetName: typeof sheet.name === 'string' ? sheet.name : 'Sheet1', values }
}

function readDocumentText(snapshot: unknown): { title: string; text: string } {
  if (snapshot === null || typeof snapshot !== 'object') throw new Error('当前会话没有可读取的文档')
  const document = snapshot as Record<string, any>
  const dataStream = document.body?.dataStream
  if (typeof dataStream !== 'string') throw new Error('当前文档没有有效正文')
  // Univer stores paragraph breaks as CR and appends CR/LF document terminators.
  const text = dataStream.replace(/\r\n\0?$/, '').replaceAll('\r', '\n').replace(/\n$/, '')
  return {
    title: typeof document.title === 'string' ? document.title : '对话文档',
    text,
  }
}

interface SlideElementSummary {
  id: string
  type: string
  text?: string
  left?: number
  top?: number
  width?: number
  height?: number
  rotation?: number
  fontSize?: number
  fontColor?: string
  bold?: boolean
  visible?: boolean
}

interface SlideSummary {
  index: number
  id: string
  title: string
  width: number
  height: number
  elements: SlideElementSummary[]
}

function readPresentation(snapshot: unknown): { title: string; slides: SlideSummary[] } {
  if (snapshot === null || typeof snapshot !== 'object') throw new Error('当前会话没有可读取的演示文稿')
  const presentation = snapshot as Record<string, any>
  const pages = presentation.slides
  const pageOrder = presentation.slideOrder
  if (pages === null || typeof pages !== 'object' || !Array.isArray(pageOrder)) throw new Error('当前演示文稿没有有效幻灯片')
  const slides = pageOrder.map((pageId: unknown, index: number) => {
    if (typeof pageId !== 'string') throw new Error('演示文稿的页面顺序无效')
    const page = pages[pageId] as Record<string, any> | undefined
    if (page === undefined) throw new Error(`找不到第 ${index + 1} 张幻灯片`)
    const elementOrder = Array.isArray(page.elementOrder) ? page.elementOrder : Object.keys(page.elements ?? {})
    const rawElements = elementOrder
      .map((elementId: unknown) => typeof elementId === 'string' ? page.elements?.[elementId] : undefined)
      .filter((element: unknown): element is Record<string, any> => element !== null && typeof element === 'object')
    const pageSize = page.pageSize ?? presentation.defaultPageSize ?? {}
    return {
      index,
      id: pageId,
      title: typeof page.name === 'string' && page.name.length > 0 ? page.name : `幻灯片 ${index + 1}`,
      width: typeof pageSize.width === 'number' ? pageSize.width : 960,
      height: typeof pageSize.height === 'number' ? pageSize.height : 540,
      elements: rawElements.map((element) => {
        const dataStream = element.textData?.body?.dataStream
        const richText = typeof dataStream === 'string'
          ? dataStream.replace(/\r\n\0?$/, '').replaceAll('\r', '\n').replace(/\n$/, '')
          : undefined
        const transform = element.transform !== null && typeof element.transform === 'object' ? element.transform : {}
        const style = element.textStyle !== null && typeof element.textStyle === 'object' ? element.textStyle : {}
        return {
          id: typeof element.id === 'string' ? element.id : '',
          type: typeof element.type === 'string' ? element.type : 'other',
          ...(typeof element.text === 'string' ? { text: element.text } : richText !== undefined ? { text: richText } : {}),
          ...(typeof transform.left === 'number' ? { left: transform.left } : {}),
          ...(typeof transform.top === 'number' ? { top: transform.top } : {}),
          ...(typeof transform.width === 'number' ? { width: transform.width } : {}),
          ...(typeof transform.height === 'number' ? { height: transform.height } : {}),
          ...(typeof transform.rotation === 'number' ? { rotation: transform.rotation } : {}),
          ...(typeof style.fontSize === 'number' ? { fontSize: style.fontSize } : {}),
          ...(typeof style.color === 'string' ? { fontColor: style.color } : {}),
          ...(typeof style.bold === 'boolean' ? { bold: style.bold } : {}),
          ...(typeof element.visible === 'boolean' ? { visible: element.visible } : {}),
        }
      }),
    }
  })
  return {
    title: typeof presentation.name === 'string' ? presentation.name : '对话演示文稿',
    slides,
  }
}

const UNIVER_EXPORT_ENDPOINT = process.env.UNIVER_FILE_EXPORT_ENDPOINT ?? 'http://127.0.0.1:8787/api/univer/export-file'
const MAX_EXPORT_BYTES = 300 * 1024 * 1024
const FILE_FORMATS = { sheet: 'xlsx', doc: 'docx', slide: 'pptx' } as const

interface WorkspaceRegistryLike {
  host: { sessionPath(sessionId: string): string | null | undefined }
}

function isInside(root: string, candidate: string): boolean {
  const result = relative(root, candidate)
  return result === '' || (!result.startsWith(`..${sep}`) && result !== '..' && !isAbsolute(result))
}

async function workspaceTarget(root: string, input: string, expectedExtension: string): Promise<string> {
  const normalized = input.trim().replaceAll('\\', '/')
  if (normalized.length === 0 || normalized.endsWith('/')) throw new Error('请输入文件名')
  if (extname(normalized).toLowerCase() !== `.${expectedExtension}`) throw new Error(`文件名必须以 .${expectedExtension} 结尾`)
  const canonicalRoot = await realpath(root)
  const target = resolve(canonicalRoot, normalized)
  if (!isInside(canonicalRoot, target)) throw new Error('保存路径必须位于当前 session workspace 内')
  const canonicalParent = await realpath(dirname(target)).catch(() => null)
  if (canonicalParent === null) throw new Error('目标目录不存在，请先创建目录')
  if (!isInside(canonicalRoot, canonicalParent)) throw new Error('保存路径不能通过符号链接离开 workspace')
  return resolve(canonicalParent, basename(target))
}

async function exportOfficeFile(data: unknown, unitType: keyof typeof FILE_FORMATS): Promise<Buffer> {
  const format = FILE_FORMATS[unitType]
  const response = await fetch(UNIVER_EXPORT_ENDPOINT, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ unitType, format, data }),
    signal: AbortSignal.timeout(300_000),
  }).catch((reason) => {
    throw new Error(`无法连接 dsh-univer-file-export：${reason instanceof Error ? reason.message : String(reason)}`)
  })
  if (!response.ok) {
    let message = `dsh-univer-file-export 返回 HTTP ${response.status}`
    try {
      const payload = await response.json() as { error?: unknown }
      if (typeof payload.error === 'string' && payload.error.length > 0) message = payload.error
    } catch {}
    throw new Error(message)
  }
  const output = Buffer.from(await response.arrayBuffer())
  if (output.byteLength === 0) throw new Error('导出服务返回了空文件')
  if (output.byteLength > MAX_EXPORT_BYTES) throw new Error('导出文件超过 300 MiB')
  if (output[0] !== 0x50 || output[1] !== 0x4b) throw new Error('导出服务返回的 Office 文件无效')
  return output
}

async function saveExportedFile(root: string, filePath: string, unitType: keyof typeof FILE_FORMATS, data: unknown, overwrite: boolean): Promise<{ path: string; size: number }> {
  const target = await workspaceTarget(root, filePath, FILE_FORMATS[unitType])
  const existing = await lstat(target).catch((reason: NodeJS.ErrnoException) => reason.code === 'ENOENT' ? null : Promise.reject(reason))
  if (existing !== null && !existing.isFile()) throw new Error('目标路径不是普通文件')
  if (existing !== null && !overwrite) throw new Error('目标文件已存在')
  const output = await exportOfficeFile(data, unitType)
  const temporary = resolve(dirname(target), `.${basename(target)}.${randomUUID()}.tmp`)
  try {
    await writeFile(temporary, output, { flag: 'wx' })
    if (existing !== null) await chmod(temporary, existing.mode)
    await rename(temporary, target)
  } catch (reason) {
    await unlink(temporary).catch(() => {})
    throw reason
  }
  const info = await stat(target)
  if (!info.isFile()) throw new Error('保存后的目标不是普通文件')
  return { path: relative(await realpath(root), target).split(sep).join('/'), size: output.byteLength }
}

export function apply(ctx: Context): void {
  const services = ctx as Context & { workspaceRegistry: WorkspaceRegistryLike }
  const domainPromise = ctx.storageDomain.open(workbookDomainSpec)
  const workbookUpdateLocks = new Map<string, Promise<void>>()
  const slideScreenshotRequests = new Map<string, Map<string, SlideScreenshotRequest>>()
  const slideScreenshotResults = new Map<string, SlideScreenshotResult>()
  const docScreenshotRequests = new Map<string, Map<string, DocScreenshotRequest>>()
  const docScreenshotResults = new Map<string, SlideScreenshotResult>()
  const univerCodeRequests = new Map<string, Map<string, UniverCodeRequest>>()
  const univerCodeResults = new Map<string, UniverCodeResult>()
  const activeSlideSessions = new Set<string>()
  const apiReference = createStandardApiReference()
  const requireActiveSlide = (sessionId: string) => {
    if (!activeSlideSessions.has(sessionId)) {
      throw new Error('请先打开当前会话的 Univer → Slide 页面，再调用 Slide 工具')
    }
  }

  const updateWorkbookRow = async (
    storageKey: string,
    update: (row: z.infer<typeof workbookRowSchema> | undefined) => z.infer<typeof workbookRowSchema>,
  ): Promise<void> => {
    const previous = workbookUpdateLocks.get(storageKey) ?? Promise.resolve()
    const next = previous.catch(() => {}).then(async () => {
      const table = (await domainPromise).table('workbooks')
      await table.put(storageKey, update(table.get(storageKey)))
    })
    workbookUpdateLocks.set(storageKey, next)
    try {
      await next
    } finally {
      if (workbookUpdateLocks.get(storageKey) === next) workbookUpdateLocks.delete(storageKey)
    }
  }

  const enqueueSubagentSheetOperation = async (
    args: Record<string, unknown>,
    exec: { agent?: SheetAgent; callId: unknown },
    action: string,
  ): Promise<void> => {
    if (!isSubagent(exec.agent)) return
    const ownerSessionId = sheetOwnerSessionId(exec.agent)
    if (ownerSessionId === undefined) throw new Error('无法确定父会话，不能写入表格')
    const queued: QueuedSheetOperation = {
      id: `${exec.agent!.id}:${String(exec.callId)}`,
      operation: { action, ...args },
    }
    await updateWorkbookRow(ownerSessionId, (row) => ({
      snapshot: row?.snapshot ?? null,
      updatedAt: Date.now(),
      filePath: row?.filePath ?? null,
      queuedSheetOperations: [
        ...(row?.queuedSheetOperations ?? []).filter((item) => item.id !== queued.id),
        queued,
      ].slice(-2_000),
    }))
  }

  let disposed = false
  const close = async () => {
    disposed = true
    await (await domainPromise).close()
  }
  // Register the RPC channel directly on the HTTP server instead of through
  // `connection.rpc.handle`. Connection's own route registration reaches
  // `webServer` through the context its methods run on, and the context a
  // plugin effect runs on does not carry the effect owner's inject list, so
  // `webServer` resolves for the plugin context but not inside that effect:
  // the call fails with `cannot get property "webServer" without inject` and
  // the whole plugin tree fails to load. Registering here keeps the wire
  // contract identical — same channel path, same `client-request` /
  // `server-response` envelope the browser-side `connection.rpc.call` speaks.
  let unregister: (() => Promise<void> | void) | undefined
  ctx.effect(() => {
    const channel = '/dsh-univer-create'
    const jsonResponse = (rpcId: string, result: unknown): Response => new Response(JSON.stringify({
      type: 'server-response',
      rpcId,
      result,
    }), { headers: { 'content-type': 'application/json' } })
    const envelopeHandler = async (request: Request): Promise<Response> => {
      if (request.method !== 'POST') return new Response('method not allowed', { status: 405 })
      if ((request.headers.get('content-type') ?? '').split(';', 1)[0]?.trim().toLowerCase() !== 'application/json') {
        return new Response('content type must be application/json', { status: 415 })
      }
      let body: { type?: unknown; rpcId?: unknown; method?: unknown; payload?: unknown }
      try {
        body = await request.json() as typeof body
      } catch {
        return new Response('body is not JSON', { status: 400 })
      }
      const rpcId = typeof body.rpcId === 'string' ? body.rpcId : 'invalid-request'
      const endpoint = new URL(request.url).pathname.slice(channel.length + 1)
      if (body.type !== 'client-request' || typeof body.method !== 'string' || endpoint === '' || endpoint !== body.method) {
        return jsonResponse(rpcId, { ok: false, error: { code: 'gateway/bad-request', message: 'invalid client-request message', details: {} } })
      }
      try {
        return jsonResponse(rpcId, await handler(body.method, body.payload, request.signal))
      } catch (error) {
        return new Response(`handler failure: ${String(error)}`, { status: 500 })
      }
    }
    const handler = async (endpoint: string, payload: unknown, _signal?: AbortSignal): Promise<unknown> => {
      const request = payload as {
        sessionId?: unknown
        unitType?: unknown
        snapshot?: unknown
        filePath?: unknown
        overwrite?: unknown
        operationIds?: unknown
        screenshotId?: unknown
        codeRequestId?: unknown
        clientId?: unknown
        dataUrl?: unknown
        width?: unknown
        height?: unknown
        result?: unknown
        logs?: unknown
        error?: unknown
        active?: unknown
      }
      if (typeof request.sessionId !== 'string' || request.sessionId.length === 0) {
        return { ok: false, error: { code: 'invalid-arguments', message: 'sessionId is required', details: {} } } as any
      }
      if (request.unitType !== undefined && request.unitType !== 'sheet' && request.unitType !== 'doc' && request.unitType !== 'slide') {
        return { ok: false, error: { code: 'invalid-arguments', message: 'unitType must be sheet, doc, or slide', details: {} } } as any
      }
      if (disposed) return { ok: false, error: { code: 'internal', message: 'plugin is disposed', details: {} } } as any
      const table = (await domainPromise).table('workbooks')
      // Keep the legacy Sheet key unchanged and store the other products under
      // namespaced keys, so existing persisted workbooks remain available.
      const storageKey = request.unitType === 'doc'
        ? `${request.sessionId}:doc`
        : request.unitType === 'slide' ? `${request.sessionId}:slide` : request.sessionId
      if (endpoint === 'load') {
        return { ok: true, value: table.get(storageKey)?.snapshot ?? null }
      }
      if (endpoint === 'file-path') {
        return { ok: true, value: table.get(storageKey)?.filePath ?? null }
      }
      if (endpoint === 'sheet-operations') {
        return { ok: true, value: table.get(request.sessionId)?.queuedSheetOperations ?? [] }
      }
      if (endpoint === 'univer-code-target') {
        return { ok: true, value: [...(univerCodeRequests.get(request.sessionId)?.values() ?? [])].map(({ id, unitType }) => ({ id, unitType })) }
      }
      if (endpoint === 'univer-code-requests') {
        const requests = [...(univerCodeRequests.get(request.sessionId)?.values() ?? [])].filter((item) => item.claimedBy === undefined)
        return { ok: true, value: request.unitType === undefined ? requests : requests.filter((item) => item.unitType === request.unitType) }
      }
      if (endpoint === 'univer-code-claim') {
        if (typeof request.codeRequestId !== 'string' || typeof request.clientId !== 'string') {
          return { ok: false, error: { code: 'invalid-arguments', message: 'codeRequestId and clientId are required', details: {} } } as any
        }
        const pending = univerCodeRequests.get(request.sessionId)
        const codeRequest = pending?.get(request.codeRequestId)
        if (codeRequest === undefined || codeRequest.claimedBy !== undefined) return { ok: true, value: null }
        codeRequest.claimedBy = request.clientId
        return { ok: true, value: codeRequest }
      }
      if (endpoint === 'univer-code-result') {
        if (typeof request.codeRequestId !== 'string' || typeof request.clientId !== 'string') {
          return { ok: false, error: { code: 'invalid-arguments', message: 'codeRequestId and clientId are required', details: {} } } as any
        }
        const pending = univerCodeRequests.get(request.sessionId)
        const codeRequest = pending?.get(request.codeRequestId)
        if (pending === undefined || codeRequest === undefined) {
          return { ok: false, error: { code: 'not-found', message: 'code request is no longer pending', details: {} } } as any
        }
        if (codeRequest.claimedBy !== request.clientId) {
          return { ok: false, error: { code: 'forbidden', message: 'code request belongs to another browser client', details: {} } } as any
        }
        const logs = Array.isArray(request.logs)
          ? request.logs.filter((item): item is string => typeof item === 'string').slice(0, 200)
          : []
        const result: UniverCodeResult = typeof request.error === 'string'
          ? { error: request.error, logs }
          : { result: typeof request.result === 'string' ? request.result : 'undefined', logs }
        univerCodeResults.set(request.codeRequestId, result)
        pending.delete(request.codeRequestId)
        if (pending.size === 0) univerCodeRequests.delete(request.sessionId)
        return { ok: true, value: { accepted: true } }
      }
      if (endpoint === 'slide-runtime-heartbeat') {
        if (request.active === true) activeSlideSessions.add(request.sessionId)
        else activeSlideSessions.delete(request.sessionId)
        return { ok: true, value: { active: activeSlideSessions.has(request.sessionId) } }
      }
      if (endpoint === 'doc-screenshot-requests') {
        return { ok: true, value: [...(docScreenshotRequests.get(request.sessionId)?.values() ?? [])] }
      }
      if (endpoint === 'doc-screenshot-result') {
        if (typeof request.screenshotId !== 'string') {
          return { ok: false, error: { code: 'invalid-arguments', message: 'screenshotId is required', details: {} } } as any
        }
        const pending = docScreenshotRequests.get(request.sessionId)
        if (!pending?.has(request.screenshotId)) {
          return { ok: false, error: { code: 'not-found', message: 'document screenshot request is no longer pending', details: {} } } as any
        }
        const result: SlideScreenshotResult = typeof request.error === 'string'
          ? { error: request.error }
          : {
              dataUrl: typeof request.dataUrl === 'string' ? request.dataUrl : undefined,
              width: typeof request.width === 'number' ? request.width : undefined,
              height: typeof request.height === 'number' ? request.height : undefined,
            }
        if (result.error === undefined && (result.dataUrl === undefined || result.width === undefined || result.height === undefined)) {
          return { ok: false, error: { code: 'invalid-arguments', message: 'dataUrl, width, and height are required', details: {} } } as any
        }
        docScreenshotResults.set(request.screenshotId, result)
        pending.delete(request.screenshotId)
        if (pending.size === 0) docScreenshotRequests.delete(request.sessionId)
        return { ok: true, value: { accepted: true } }
      }
      if (endpoint === 'slide-screenshot-requests') {
        return { ok: true, value: [...(slideScreenshotRequests.get(request.sessionId)?.values() ?? [])] }
      }
      if (endpoint === 'slide-screenshot-result') {
        if (typeof request.screenshotId !== 'string') {
          return { ok: false, error: { code: 'invalid-arguments', message: 'screenshotId is required', details: {} } } as any
        }
        const pending = slideScreenshotRequests.get(request.sessionId)
        if (!pending?.has(request.screenshotId)) {
          return { ok: false, error: { code: 'not-found', message: 'screenshot request is no longer pending', details: {} } } as any
        }
        const result: SlideScreenshotResult = typeof request.error === 'string'
          ? { error: request.error }
          : {
              dataUrl: typeof request.dataUrl === 'string' ? request.dataUrl : undefined,
              width: typeof request.width === 'number' ? request.width : undefined,
              height: typeof request.height === 'number' ? request.height : undefined,
            }
        if (result.error === undefined && (result.dataUrl === undefined || result.width === undefined || result.height === undefined)) {
          return { ok: false, error: { code: 'invalid-arguments', message: 'dataUrl, width, and height are required', details: {} } } as any
        }
        slideScreenshotResults.set(request.screenshotId, result)
        pending.delete(request.screenshotId)
        if (pending.size === 0) slideScreenshotRequests.delete(request.sessionId)
        return { ok: true, value: { accepted: true } }
      }
      if (endpoint === 'ack-sheet-operations') {
        if (!Array.isArray(request.operationIds) || !request.operationIds.every((id) => typeof id === 'string')) {
          return { ok: false, error: { code: 'invalid-arguments', message: 'operationIds must be a string array', details: {} } } as any
        }
        const acknowledged = new Set(request.operationIds as string[])
        await updateWorkbookRow(request.sessionId, (row) => ({
          snapshot: row?.snapshot ?? null,
          updatedAt: Date.now(),
          filePath: row?.filePath ?? null,
          queuedSheetOperations: (row?.queuedSheetOperations ?? []).filter((item) => !acknowledged.has(item.id)),
        }))
        return { ok: true, value: { acknowledged: acknowledged.size } }
      }
      if (endpoint === 'save' && request.snapshot !== undefined) {
        let nextFilePath: string | null | undefined
        await updateWorkbookRow(storageKey, (previous) => {
          nextFilePath = request.filePath === null
            ? null
            : typeof request.filePath === 'string' ? request.filePath : previous?.filePath
          return {
            snapshot: request.snapshot,
            updatedAt: Date.now(),
            filePath: nextFilePath,
            queuedSheetOperations: previous?.queuedSheetOperations,
          }
        })
        return { ok: true, value: { saved: true, filePath: nextFilePath ?? null } }
      }
      if (endpoint === 'export' && request.snapshot !== undefined) {
        if (request.unitType !== 'sheet' && request.unitType !== 'doc' && request.unitType !== 'slide') {
          return { ok: false, error: { code: 'invalid-arguments', message: 'unitType is required for export', details: {} } } as any
        }
        if (typeof request.filePath !== 'string') {
          return { ok: false, error: { code: 'invalid-arguments', message: 'filePath is required for export', details: {} } } as any
        }
        const root = services.workspaceRegistry.host.sessionPath(request.sessionId)
        if (typeof root !== 'string' || root.length === 0) throw new Error('无法解析当前 session workspace 目录')
        const saved = await saveExportedFile(root, request.filePath, request.unitType, request.snapshot, request.overwrite === true)
        await updateWorkbookRow(storageKey, (previous) => ({
          snapshot: request.snapshot,
          updatedAt: Date.now(),
          filePath: saved.path,
          queuedSheetOperations: previous?.queuedSheetOperations,
        }))
        return { ok: true, value: saved }
      }
      return { ok: false, error: { code: 'not-found', message: `unknown endpoint: ${endpoint}`, details: {} } } as any
    }
    return (ctx as Context & { webServer: { register: (route: { kind: string; path: string; handler: (req: any, res: any) => void }) => () => void } }).webServer.register({
      kind: 'prefix',
      path: channel,
      handler: (req: any, res: any) => {
        const chunks: Buffer[] = []
        req.on('data', (chunk: Buffer) => chunks.push(chunk))
        req.on('end', () => {
          const request = new Request(new URL(req.url ?? '/', 'http://dsh.internal'), {
            method: req.method ?? 'GET',
            headers: Object.fromEntries(Object.entries(req.headers as Record<string, unknown>).filter(([, value]) => typeof value === 'string')) as Record<string, string>,
            ...(chunks.length > 0 ? { body: Buffer.concat(chunks) } : {}),
          })
          envelopeHandler(request).then(async (response) => {
            res.writeHead(response.status, Object.fromEntries(response.headers))
            res.end(Buffer.from(await response.arrayBuffer()))
          }, (error: unknown) => {
            res.writeHead(500)
            res.end(String(error))
          })
        })
      },
    })
  }, 'dsh-univer-create: rpc channel')
  ctx.effect(() => async () => {
    await unregister?.()
    await close()
  }, 'dsh-univer-create: host persistence')

  ctx.on('tools/pre-execute', async (exec, next) => {
    if (exec.name !== 'univer_execute_code') return next()
    const approval = ctx.get('approval') as undefined | {
      config: { policy?: 'ask' | 'never' }
      overrideOf: (session: unknown) => 'ask' | 'never' | undefined
    }
    const policy = exec.agent === undefined
      ? approval?.config.policy
      : approval?.overrideOf(exec.agent.session) ?? approval?.config.policy
    if (policy === 'never') return next()
    return {
      kind: 'ask',
      reason: 'This runs AI-generated JavaScript in the DSH page realm. It is not a security sandbox: review the code because same-page JavaScript may affect the editor or page data.',
    }
  })

  ctx.tools.register(defineTool({
    name: 'univer_api_reference',
    description: 'Search the version-matched Univer Facade API reference before writing browser code. Use action=find with API-name keywords, then action=show with returned symbols to inspect exact signatures, examples, nullability, and related types.',
    parameters: {
      action: { type: 'string', enum: ['find', 'show'], required: true, description: 'find discovers symbols; show returns exact symbol details.' },
      terms: { type: 'array', items: { type: 'string' }, description: 'Non-empty search terms for action=find.' },
      symbols: { type: 'array', items: { type: 'string' }, description: 'Exact symbols for action=show, for example FRange.setValues.' },
      unit: { type: 'string', enum: ['sheet', 'slide', 'doc', 'base', 'board'], description: 'Optional unit filter for action=find.' },
      limit: { type: 'integer', description: 'Positive result limit per term for action=find. Defaults to 10.' },
    },
    output: {
      schema: {
        type: 'object',
        properties: {
          ok: { type: 'boolean', required: true },
          action: { type: 'string', required: true },
          response: { type: 'string', required: true },
        },
        additionalProperties: false,
      } as const,
      render: (_args: unknown, value: { response: string }) => [{ type: 'text' as const, text: value.response }],
    },
    async execute(args) {
      if (args.action === 'find' && (!Array.isArray(args.terms) || args.terms.length === 0)) throw new Error('action=find requires at least one term')
      if (args.action === 'show' && (!Array.isArray(args.symbols) || args.symbols.length === 0)) throw new Error('action=show requires at least one symbol')
      if ((args.terms?.length ?? 0) > 10 || (args.symbols?.length ?? 0) > 10) throw new Error('a reference query accepts at most 10 terms or symbols')
      if ([...(args.terms ?? []), ...(args.symbols ?? [])].some((item) => item.length > 200)) throw new Error('reference terms and symbols are limited to 200 characters')
      if (args.limit !== undefined && (!Number.isInteger(args.limit) || args.limit <= 0 || args.limit > 50)) throw new Error('limit must be an integer between 1 and 50')
      const response = args.action === 'find'
        ? apiReference.find({
            terms: args.terms!,
            ...(args.unit === undefined ? {} : { unit: args.unit }),
            limit: args.limit ?? 10,
          })
        : apiReference.show(args.symbols!)
      const serialized = JSON.stringify(response, null, 2)
      return {
        ok: true,
        action: args.action,
        response: serialized.length > 100_000 ? `${serialized.slice(0, 100_000)}\n…[truncated; query fewer symbols]` : serialized,
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_execute_code',
    description: 'Execute JavaScript against the active Univer editor in this conversation browser. First query univer_api_reference. The supported contract injects only `univerAPI` (FUniver Facade) and a captured `console`; obtain objects through getActiveWorkbook(), getActiveDocument(), or getActivePresentation(). Do not use imports, DOM/browser globals, raw Univer internals, injector/model access, or TypeScript syntax. Async code and await are supported. Return a JSON-serializable value for inspection. The target Univer tab must be open and the requested unit must already exist. This is trusted same-page execution, not a security sandbox. Under the Ask permission preset it requests approval; under Full Access it executes automatically.',
    parameters: {
      unitType: { type: 'string', enum: ['sheet', 'doc', 'slide'], required: true, description: 'Target active Univer unit.' },
      code: { type: 'string', required: true, description: 'JavaScript function body. `univerAPI` is the only injected Univer binding. Use `return` to send a result back.' },
    },
    output: {
      schema: {
        type: 'object',
        properties: {
          ok: { type: 'boolean', required: true },
          action: { type: 'string', required: true },
          unitType: { type: 'string', required: true },
          result: { type: 'string', required: true },
          logs: { type: 'array', items: { type: 'string' }, required: true },
        },
        additionalProperties: false,
      } as const,
      render: (_args: unknown, value: { unitType: string; result: string; logs: string[] }) => [{
        type: 'text' as const,
        text: `${value.unitType} code result:\n${value.result}${value.logs.length > 0 ? `\nconsole:\n${value.logs.join('\n')}` : ''}`,
      }],
    },
    timeoutMs: 40_000,
    async execute(args, exec) {
      if (isSubagent(exec.agent)) throw new Error('univer_execute_code 只能由有可见 Univer 页面的主会话调用，不能由子代理调用')
      const sessionId = sheetOwnerSessionId(exec.agent)
      if (sessionId === undefined) throw new Error('无法确定当前会话，不能执行 Univer 代码')
      if (args.code.trim().length === 0) throw new Error('code must not be empty')
      if (args.code.length > 50_000) throw new Error('code exceeds the 50000 character limit')
      const request: UniverCodeRequest = {
        id: randomUUID(),
        unitType: args.unitType,
        code: args.code,
        createdAt: Date.now(),
      }
      const pending = univerCodeRequests.get(sessionId) ?? new Map<string, UniverCodeRequest>()
      pending.set(request.id, request)
      univerCodeRequests.set(sessionId, pending)

      const deadline = Date.now() + 30_000
      let result: UniverCodeResult | undefined
      try {
        while (Date.now() < deadline) {
          if (exec.signal.aborted) throw exec.signal.reason ?? new Error('Univer code execution was cancelled')
          result = univerCodeResults.get(request.id)
          if (result !== undefined) break
          await new Promise<void>((resolveWait) => setTimeout(resolveWait, 100))
        }
      } finally {
        univerCodeRequests.get(sessionId)?.delete(request.id)
        if (univerCodeRequests.get(sessionId)?.size === 0) univerCodeRequests.delete(sessionId)
        univerCodeResults.delete(request.id)
      }
      if (result === undefined) throw new Error('浏览器执行超时。请保持当前会话的 Univer 页签打开，并确认目标文档已创建。')
      if (result.error !== undefined) throw new Error(`浏览器执行失败：${result.error}${result.logs?.length ? `\n${result.logs.join('\n')}` : ''}`)
      return {
        ok: true,
        action: 'execute-code',
        unitType: args.unitType,
        result: result.result ?? 'undefined',
        logs: result.logs ?? [],
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_sheet_new',
    description: 'Create or reset the Univer workbook shown in the Sheet tab for this conversation. Call this before editing when the user asks for a new spreadsheet.',
    parameters: {
      title: { type: 'string', description: 'Workbook title. Defaults to 对话表格.' },
      sheetName: { type: 'string', description: 'Initial worksheet name. Defaults to Sheet1.' },
      rows: { type: 'integer', description: 'Initial row count. Defaults to 100.' },
      columns: { type: 'integer', description: 'Initial column count. Defaults to 26.' },
    },
    output,
    async execute(args, exec) {
      await enqueueSubagentSheetOperation(args, exec, 'new')
      return {
        ok: true,
        action: 'new',
        message: `已创建 Univer 表格“${args.title ?? '对话表格'}”，工作表为“${args.sheetName ?? 'Sheet1'}”。`,
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_sheet_add',
    description: 'Add a worksheet to the current Univer workbook. The new worksheet appears as a tab in the current Sheet view.',
    parameters: {
      name: { type: 'string', required: true, description: 'New worksheet name.' },
      rows: { type: 'integer', description: 'Initial row count. Defaults to 100.' },
      columns: { type: 'integer', description: 'Initial column count. Defaults to 26.' },
    },
    output,
    async execute(args, exec) {
      await enqueueSubagentSheetOperation(args, exec, 'add-sheet')
      return { ok: true, action: 'add-sheet', message: `已添加工作表“${args.name}”。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_sheet_delete',
    description: 'Delete a worksheet from the current Univer workbook by name.',
    parameters: {
      name: { type: 'string', required: true, description: 'Worksheet name to delete.' },
    },
    output,
    async execute(args, exec) {
      await enqueueSubagentSheetOperation(args, exec, 'delete-sheet')
      return { ok: true, action: 'delete-sheet', message: `已删除工作表“${args.name}”。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_sheet_rename',
    description: 'Rename a worksheet in the current Univer workbook.',
    parameters: {
      oldName: { type: 'string', required: true, description: 'Current worksheet name.' },
      newName: { type: 'string', required: true, description: 'New worksheet name.' },
    },
    output,
    async execute(args, exec) {
      await enqueueSubagentSheetOperation(args, exec, 'rename-sheet')
      return { ok: true, action: 'rename-sheet', message: `已将工作表“${args.oldName}”重命名为“${args.newName}”。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_sheet_list',
    description: 'List the worksheets in the current Univer workbook.',
    parameters: {},
    output: {
      schema: {
        type: 'object',
        properties: {
          ok: { type: 'boolean', required: true },
          action: { type: 'string', required: true },
          message: { type: 'string', required: true },
          sheets: { type: 'array', items: { type: 'string' }, required: true },
        },
        additionalProperties: false,
      } as const,
      render: (_args: unknown, value: { message: string }) => [
        { type: 'text' as const, text: value.message },
      ],
    },
    async execute() {
      return {
        ok: true,
        action: 'list-sheets',
        sheets: [],
        message: 'Sheet 列表由当前会话中的 Sheet 页面维护；请打开 Sheet 页签查看当前工作表标签。',
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_sheet_get_range',
    description: 'Read values from an A1 range in the current conversation Univer Sheet. Call this before modifying an opened workbook when you need to inspect its existing data.',
    parameters: {
      range: { type: 'string', required: true, description: 'A1 range to read, for example A1:D20.' },
      sheetName: { type: 'string', description: 'Target worksheet name. Defaults to the active sheet.' },
    },
    output: {
      schema: {
        type: 'object',
        properties: {
          ok: { type: 'boolean', required: true },
          action: { type: 'string', required: true },
          message: { type: 'string', required: true },
          sheetName: { type: 'string', required: true },
          range: { type: 'string', required: true },
          values: {
            type: 'array',
            required: true,
            items: {
              type: 'array',
              items: {
                oneOf: [
                  { type: 'string' },
                  { type: 'number' },
                  { type: 'boolean' },
                  { type: 'null' },
                ],
              },
            },
          },
        },
        additionalProperties: false,
      } as const,
      render: (_args: unknown, value: { sheetName: string; range: string; values: CellValue[][] }) => [
        { type: 'text' as const, text: `${value.sheetName}!${value.range}:\n${JSON.stringify(value.values)}` },
      ],
    },
    async execute(args, exec) {
      const sessionId = sheetOwnerSessionId(exec.agent)
      if (sessionId === undefined) throw new Error('无法确定当前会话，不能读取表格')
      const stored = (await domainPromise).table('workbooks').get(sessionId)
      const result = readSnapshotRange(stored?.snapshot, args.range, args.sheetName)
      return {
        ok: true,
        action: 'get-range',
        message: `已读取 ${result.sheetName}!${args.range}。`,
        sheetName: result.sheetName,
        range: args.range,
        values: result.values,
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_sheet_set_range',
    description: 'Write a rectangular 2D array into an A1 range in the conversation Univer Sheet. Values may be strings, numbers, booleans, null, or formulas beginning with =.',
    parameters: {
      range: { type: 'string', required: true, description: 'A1 range, for example A1:D4.' },
      values: {
        type: 'array',
        required: true,
        description: 'Rectangular two-dimensional values array.',
        items: {
          type: 'array',
          items: {
            oneOf: [
              { type: 'string' },
              { type: 'number' },
              { type: 'boolean' },
              { type: 'null' },
            ],
          },
        },
      },
      sheetName: { type: 'string', description: 'Target worksheet name. Defaults to the active sheet.' },
    },
    output,
    async execute(args, exec) {
      await enqueueSubagentSheetOperation(args, exec, 'set-range')
      return { ok: true, action: 'set-range', message: `已写入 ${args.sheetName ? `${args.sheetName}!` : ''}${args.range}。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_sheet_clear_range',
    description: 'Clear cell contents in an A1 range in the conversation Univer Sheet.',
    parameters: {
      range: { type: 'string', required: true, description: 'A1 range to clear.' },
      sheetName: { type: 'string', description: 'Target worksheet name. Defaults to the active sheet.' },
    },
    output,
    async execute(args, exec) {
      await enqueueSubagentSheetOperation(args, exec, 'clear-range')
      return { ok: true, action: 'clear-range', message: `已清空 ${args.sheetName ? `${args.sheetName}!` : ''}${args.range}。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_sheet_format_range',
    description: 'Format an A1 range in the conversation Univer Sheet. Use CSS hex colors such as #2563eb.',
    parameters: {
      range: { type: 'string', required: true, description: 'A1 range to format.' },
      sheetName: { type: 'string', description: 'Target worksheet name. Defaults to the active sheet.' },
      background: { type: 'string', description: 'Cell background hex color.' },
      fontColor: { type: 'string', description: 'Font hex color.' },
      bold: { type: 'boolean', description: 'Whether text is bold.' },
      fontSize: { type: 'number', description: 'Font size in points.' },
    },
    output,
    async execute(args, exec) {
      await enqueueSubagentSheetOperation(args, exec, 'format-range')
      return { ok: true, action: 'format-range', message: `已设置 ${args.sheetName ? `${args.sheetName}!` : ''}${args.range} 的格式。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_doc_new',
    description: 'Create or reset the Univer document shown in the shared Univer tab for this conversation. Call this before editing when the user asks for a new document.',
    parameters: {
      title: { type: 'string', description: 'Document title. Defaults to 对话文档.' },
      text: { type: 'string', description: 'Optional initial plain text.' },
    },
    output,
    async execute(args) {
      return { ok: true, action: 'new-doc', message: `已创建 Univer 文档“${args.title ?? '对话文档'}”。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_doc_get_text',
    description: 'Read the plain text from the current conversation Univer Doc. Call this before modifying an opened document when you need to inspect its existing content.',
    parameters: {},
    output: {
      schema: {
        type: 'object',
        properties: {
          ok: { type: 'boolean', required: true },
          action: { type: 'string', required: true },
          message: { type: 'string', required: true },
          title: { type: 'string', required: true },
          text: { type: 'string', required: true },
        },
        additionalProperties: false,
      } as const,
      render: (_args: unknown, value: { title: string; text: string }) => [
        { type: 'text' as const, text: `${value.title}:\n${value.text}` },
      ],
    },
    async execute(_args, exec) {
      const sessionId = exec.agent?.id
      if (sessionId === undefined) throw new Error('无法确定当前会话，不能读取文档')
      const stored = (await domainPromise).table('workbooks').get(`${sessionId}:doc`)
      const result = readDocumentText(stored?.snapshot)
      return {
        ok: true,
        action: 'get-doc-text',
        message: `已读取文档“${result.title}”。`,
        title: result.title,
        text: result.text,
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_doc_set_text',
    description: 'Replace all plain text in the current conversation Univer Doc.',
    parameters: {
      text: { type: 'string', required: true, description: 'Replacement plain text.' },
    },
    output,
    async execute() {
      return { ok: true, action: 'set-doc-text', message: '已替换 Univer 文档正文。' }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_doc_insert_text',
    description: 'Insert plain text at a zero-based character offset in the current conversation Univer Doc.',
    parameters: {
      index: { type: 'integer', required: true, description: 'Zero-based insertion offset in the document body.' },
      text: { type: 'string', required: true, description: 'Plain text to insert.' },
    },
    output,
    async execute(args) {
      return { ok: true, action: 'insert-doc-text', message: `已在文档位置 ${args.index} 插入文本。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_doc_append_text',
    description: 'Append plain text to the current conversation Univer Doc.',
    parameters: {
      text: { type: 'string', required: true, description: 'Plain text to append.' },
    },
    output,
    async execute() {
      return { ok: true, action: 'append-doc-text', message: '已追加 Univer 文档正文。' }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_doc_delete_range',
    description: 'Delete a zero-based, end-exclusive text range from the current conversation Univer Doc.',
    parameters: {
      start: { type: 'integer', required: true, description: 'Inclusive zero-based start offset.' },
      end: { type: 'integer', required: true, description: 'Exclusive zero-based end offset.' },
    },
    output,
    async execute(args) {
      return { ok: true, action: 'delete-doc-range', message: `已删除文档文本范围 [${args.start}, ${args.end})。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_doc_format_text',
    description: 'Format a zero-based, end-exclusive text range in the current conversation Univer Doc. Use CSS hex colors such as #2563eb.',
    parameters: {
      start: { type: 'integer', required: true, description: 'Inclusive zero-based start offset.' },
      end: { type: 'integer', required: true, description: 'Exclusive zero-based end offset.' },
      bold: { type: 'boolean', description: 'Whether text is bold.' },
      italic: { type: 'boolean', description: 'Whether text is italic.' },
      fontSize: { type: 'number', description: 'Font size in points.' },
      fontColor: { type: 'string', description: 'Text color hex value.' },
    },
    output,
    async execute(args) {
      return { ok: true, action: 'format-doc-text', message: `已设置文档文本范围 [${args.start}, ${args.end}) 的格式。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_doc_screenshot',
    description: 'Capture the currently rendered Univer document viewport as a PNG and return it to the model for visual inspection. Use after editing to check hierarchy, spacing, clipping, and page layout. The Univer Doc tab must be initialized in the browser.',
    parameters: {
      mode: { type: 'string', enum: ['document', 'editor'], description: 'document returns the rendered document canvas; editor is reserved for editor diagnostics. Defaults to document.' },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          ok: { type: 'boolean', required: true },
          action: { type: 'string', required: true },
          message: { type: 'string', required: true },
          image: imageValueSchema,
        },
      } as const,
      render: (_args: unknown, value: { message: string; image: ScreenshotImageValue }) => [{
        type: 'text' as const,
        text: value.message,
      }, {
        type: 'image' as const,
        attachment: {
          attachmentId: value.image.attachmentId,
          mediaType: value.image.mediaType,
          bytes: value.image.bytes,
          width: value.image.width,
          height: value.image.height,
          ...(value.image.name === undefined ? {} : { name: value.image.name }),
        } as any,
      }],
    },
    async execute(args, exec) {
      const sessionId = sheetOwnerSessionId(exec.agent)
      if (sessionId === undefined) throw new Error('无法确定当前会话，不能截取文档')
      const stored = (await domainPromise).table('workbooks').get(`${sessionId}:doc`)
      if (stored?.snapshot === null || stored?.snapshot === undefined) {
        throw new Error('当前没有已创建或已打开的文档；请先在 Univer → Doc 中点击“新建”或“打开”')
      }
      const request: DocScreenshotRequest = { id: randomUUID(), mode: args.mode ?? 'document', createdAt: Date.now() }
      const pending = docScreenshotRequests.get(sessionId) ?? new Map<string, DocScreenshotRequest>()
      pending.set(request.id, request)
      docScreenshotRequests.set(sessionId, pending)

      const deadline = Date.now() + 15_000
      let result: SlideScreenshotResult | undefined
      try {
        while (Date.now() < deadline) {
          if (exec.signal.aborted) throw exec.signal.reason
          result = docScreenshotResults.get(request.id)
          if (result !== undefined) break
          await new Promise<void>((resolveWait) => setTimeout(resolveWait, 100))
        }
      } finally {
        docScreenshotRequests.get(sessionId)?.delete(request.id)
        if (docScreenshotRequests.get(sessionId)?.size === 0) docScreenshotRequests.delete(sessionId)
        docScreenshotResults.delete(request.id)
      }
      if (result === undefined) throw new Error('文档截图超时。请打开当前会话的 Univer → Doc 页签后重试。')
      if (result.error !== undefined) throw new Error(`浏览器文档截图失败：${result.error}`)
      const match = /^data:image\/png;base64,([A-Za-z0-9+/]+={0,2})$/.exec(result.dataUrl ?? '')
      if (match === null) throw new Error('浏览器返回的文档截图不是有效 PNG data URL')
      const data = Buffer.from(match[1]!, 'base64')
      if (data.byteLength === 0 || data.byteLength > 20 * 1024 * 1024) throw new Error('文档截图大小无效或超过 20 MiB')
      const attachments = (ctx as Context & { attachments?: { saveImage: (input: { data: Uint8Array; mediaType: 'image/png'; name?: string }) => Promise<any> } }).attachments
      if (attachments === undefined) throw new Error('当前 DSH 未挂载附件服务，无法把文档截图返回给模型')
      const ref = await attachments.saveImage({ data, mediaType: 'image/png', name: 'document-viewport.png' })
      const image: ScreenshotImageValue = {
        attachmentId: String(ref.attachmentId),
        mediaType: 'image/png',
        bytes: ref.bytes,
        width: ref.width,
        height: ref.height,
        ...(ref.name === undefined ? {} : { name: ref.name }),
      }
      return {
        ok: true,
        action: 'doc-screenshot',
        message: `当前文档可视区域截图（${image.width}×${image.height}）。请检查标题层级、间距、裁切、对齐和页面留白。`,
        image,
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_slide_new',
    description: 'Create or reset the Univer presentation shown in the shared Univer tab for this conversation. Call this before editing when the user asks for a new presentation.',
    parameters: {
      title: { type: 'string', description: 'Presentation title. Defaults to 对话演示文稿.' },
      width: { type: 'number', description: 'Slide width in canvas units. Defaults to 960.' },
      height: { type: 'number', description: 'Slide height in canvas units. Defaults to 540.' },
    },
    output,
    async execute(args, exec) {
      const sessionId = sheetOwnerSessionId(exec.agent)
      if (sessionId === undefined) throw new Error('无法确定当前会话，不能创建演示文稿')
      requireActiveSlide(sessionId)
      return { ok: true, action: 'new-slide', message: `已创建 Univer 演示文稿“${args.title ?? '对话演示文稿'}”。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_slide_list',
    description: 'List slides and their elements in the current conversation Univer presentation. Call this before updating an opened presentation when you need page or element ids.',
    parameters: {},
    output: {
      schema: {
        type: 'object',
        properties: {
          ok: { type: 'boolean', required: true },
          action: { type: 'string', required: true },
          message: { type: 'string', required: true },
          title: { type: 'string', required: true },
          slides: {
            type: 'array',
            required: true,
            items: {
              type: 'object',
              properties: {
                index: { type: 'integer', required: true },
                id: { type: 'string', required: true },
                title: { type: 'string', required: true },
                width: { type: 'number', required: true },
                height: { type: 'number', required: true },
                elements: {
                  type: 'array',
                  required: true,
                  items: {
                    type: 'object',
                    properties: {
                      id: { type: 'string', required: true },
                      type: { type: 'string', required: true },
                      text: { type: 'string' },
                      left: { type: 'number' },
                      top: { type: 'number' },
                      width: { type: 'number' },
                      height: { type: 'number' },
                      rotation: { type: 'number' },
                      fontSize: { type: 'number' },
                      fontColor: { type: 'string' },
                      bold: { type: 'boolean' },
                      visible: { type: 'boolean' },
                    },
                    additionalProperties: false,
                  },
                },
              },
              additionalProperties: false,
            },
          },
        },
        additionalProperties: false,
      } as const,
      render: (_args: unknown, value: { title: string; slides: SlideSummary[] }) => [
        { type: 'text' as const, text: `${value.title}:\n${JSON.stringify(value.slides)}` },
      ],
    },
    async execute(_args, exec) {
      const sessionId = sheetOwnerSessionId(exec.agent)
      if (sessionId === undefined) throw new Error('无法确定当前会话，不能读取演示文稿')
      requireActiveSlide(sessionId)
      const stored = (await domainPromise).table('workbooks').get(`${sessionId}:slide`)
      const result = readPresentation(stored?.snapshot)
      return {
        ok: true,
        action: 'list-slides',
        message: `已读取演示文稿“${result.title}”的 ${result.slides.length} 张幻灯片。`,
        title: result.title,
        slides: result.slides,
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_slide_screenshot',
    description: 'Capture a rendered PNG of one slide and return the image to the model for visual inspection. Use after creating or changing slides to check overlap, clipping, hierarchy, contrast, and whitespace. The Univer Slide tab must be open in the browser.',
    parameters: {
      slideIndex: { type: 'integer', description: 'Zero-based slide index. Defaults to 0.' },
      mode: { type: 'string', enum: ['slide', 'editor'], description: 'slide returns the clean rendered canvas; editor is intended for editor-render diagnostics. Defaults to slide.' },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          ok: { type: 'boolean', required: true },
          action: { type: 'string', required: true },
          message: { type: 'string', required: true },
          slideIndex: { type: 'integer', required: true },
          image: imageValueSchema,
        },
      } as const,
      render: (_args: unknown, value: { message: string; image: ScreenshotImageValue }) => [{
        type: 'text' as const,
        text: value.message,
      }, {
        type: 'image' as const,
        attachment: {
          attachmentId: value.image.attachmentId,
          mediaType: value.image.mediaType,
          bytes: value.image.bytes,
          width: value.image.width,
          height: value.image.height,
          ...(value.image.name === undefined ? {} : { name: value.image.name }),
        } as any,
      }],
    },
    async execute(args, exec) {
      const sessionId = sheetOwnerSessionId(exec.agent)
      if (sessionId === undefined) throw new Error('无法确定当前会话，不能截取幻灯片')
      requireActiveSlide(sessionId)
      const slideIndex = args.slideIndex ?? 0
      if (!Number.isInteger(slideIndex) || slideIndex < 0) throw new Error('slideIndex 必须是非负整数')
      const stored = (await domainPromise).table('workbooks').get(`${sessionId}:slide`)
      if (stored?.snapshot === null || stored?.snapshot === undefined) {
        throw new Error('当前没有已创建或已打开的演示文稿；请先在 Univer → Slide 中点击“新建”或“打开”')
      }
      const mode = args.mode ?? 'slide'
      const request: SlideScreenshotRequest = { id: randomUUID(), slideIndex, mode, createdAt: Date.now() }
      const pending = slideScreenshotRequests.get(sessionId) ?? new Map<string, SlideScreenshotRequest>()
      pending.set(request.id, request)
      slideScreenshotRequests.set(sessionId, pending)

      const deadline = Date.now() + 15_000
      let result: SlideScreenshotResult | undefined
      try {
        while (Date.now() < deadline) {
          if (exec.signal.aborted) throw exec.signal.reason
          result = slideScreenshotResults.get(request.id)
          if (result !== undefined) break
          await new Promise<void>((resolveWait) => setTimeout(resolveWait, 100))
        }
      } finally {
        slideScreenshotRequests.get(sessionId)?.delete(request.id)
        if (slideScreenshotRequests.get(sessionId)?.size === 0) slideScreenshotRequests.delete(sessionId)
        slideScreenshotResults.delete(request.id)
      }
      if (result === undefined) throw new Error('截图超时。请保持 Univer Slide 页签打开并重试。')
      if (result.error !== undefined) throw new Error(`浏览器截图失败：${result.error}`)
      const match = /^data:image\/png;base64,([A-Za-z0-9+/]+={0,2})$/.exec(result.dataUrl ?? '')
      if (match === null) throw new Error('浏览器返回的截图不是有效 PNG data URL')
      const data = Buffer.from(match[1]!, 'base64')
      if (data.byteLength === 0 || data.byteLength > 20 * 1024 * 1024) throw new Error('截图大小无效或超过 20 MiB')
      const attachments = (ctx as Context & { attachments?: { saveImage: (input: { data: Uint8Array; mediaType: 'image/png'; name?: string }) => Promise<any> } }).attachments
      if (attachments === undefined) throw new Error('当前 DSH 未挂载附件服务，无法把截图返回给模型')
      const ref = await attachments.saveImage({ data, mediaType: 'image/png', name: `slide-${slideIndex + 1}.png` })
      const image: ScreenshotImageValue = {
        attachmentId: String(ref.attachmentId),
        mediaType: 'image/png',
        bytes: ref.bytes,
        width: ref.width,
        height: ref.height,
        ...(ref.name === undefined ? {} : { name: ref.name }),
      }
      return {
        ok: true,
        action: 'slide-screenshot',
        message: `第 ${slideIndex + 1} 张幻灯片截图（${image.width}×${image.height}）。请直接检查文字重叠、裁切、层级、对齐、对比度和留白。`,
        slideIndex,
        image,
      }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_slide_add',
    description: 'Add a blank slide to the current Univer presentation.',
    parameters: {
      title: { type: 'string', description: 'Optional slide title.' },
      index: { type: 'integer', description: 'Zero-based insertion index. Defaults to the end.' },
    },
    output,
    async execute(args, exec) {
      const sessionId = sheetOwnerSessionId(exec.agent)
      if (sessionId === undefined) throw new Error('无法确定当前会话，不能添加幻灯片')
      requireActiveSlide(sessionId)
      return { ok: true, action: 'add-slide', message: `已添加${args.title ? `“${args.title}”` : '一张'}幻灯片。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_slide_delete',
    description: 'Delete a slide by its zero-based index from the current Univer presentation.',
    parameters: {
      index: { type: 'integer', required: true, description: 'Zero-based slide index.' },
    },
    output,
    async execute(args, exec) {
      const sessionId = sheetOwnerSessionId(exec.agent)
      if (sessionId === undefined) throw new Error('无法确定当前会话，不能删除幻灯片')
      requireActiveSlide(sessionId)
      return { ok: true, action: 'delete-slide', message: `已删除第 ${args.index + 1} 张幻灯片。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_slide_add_text',
    description: 'Add a text element to a slide in the current Univer presentation. Coordinates and sizes use slide canvas units.',
    parameters: {
      slideIndex: { type: 'integer', required: true, description: 'Zero-based slide index.' },
      text: { type: 'string', required: true, description: 'Text content.' },
      left: { type: 'number', description: 'Left position. Defaults to 120.' },
      top: { type: 'number', description: 'Top position. Defaults to 100.' },
      width: { type: 'number', description: 'Element width. Defaults to 720.' },
      height: { type: 'number', description: 'Element height. Defaults to 80.' },
      fontSize: { type: 'number', description: 'Font size. Defaults to 30.' },
      fontColor: { type: 'string', description: 'Font CSS color. Defaults to #333333.' },
      bold: { type: 'boolean', description: 'Whether text is bold.' },
    },
    output,
    async execute(args, exec) {
      const sessionId = sheetOwnerSessionId(exec.agent)
      if (sessionId === undefined) throw new Error('无法确定当前会话，不能添加文本')
      requireActiveSlide(sessionId)
      return { ok: true, action: 'add-slide-text', message: `已向第 ${args.slideIndex + 1} 张幻灯片添加文本。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_slide_add_shape',
    description: 'Add a real editable Univer Slides shape. Supports rect, roundRect, ellipse, diamond, triangle, parallelogram, hexagon, star5 and other ShapeTypeEnum values. Set fillColor to transparent or strokeColor to none to remove fill or outline.',
    parameters: {
      slideIndex: { type: 'integer', required: true, description: 'Zero-based slide index.' },
      shapeType: { type: 'string', required: true, description: 'Univer ShapeTypeEnum value, for example rect, roundRect, ellipse, diamond, triangle, hexagon, or star5.' },
      left: { type: 'number', description: 'Left position in slide canvas units.' },
      top: { type: 'number', description: 'Top position in slide canvas units.' },
      width: { type: 'number', description: 'Shape width.' },
      height: { type: 'number', description: 'Shape height.' },
      fillColor: { type: 'string', description: 'Fill color such as #2563eb; use transparent for no fill.' },
      strokeColor: { type: 'string', description: 'Outline color; use none for no outline.' },
      strokeWidth: { type: 'number', description: 'Outline width. Defaults to 1.5.' },
      opacity: { type: 'number', description: 'Opacity from 0 to 1.' },
      selectable: { type: 'boolean', description: 'Whether the editor should show/select the shape. Defaults to true so users can select and edit it.' },
    },
    output,
    async execute(args, exec) {
      const sessionId = sheetOwnerSessionId(exec.agent)
      if (sessionId === undefined) throw new Error('无法确定当前会话，不能添加形状')
      requireActiveSlide(sessionId)
      return { ok: true, action: 'add-slide-shape', message: `已向第 ${args.slideIndex + 1} 张幻灯片添加 ${args.shapeType} 形状。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_slide_update_text',
    description: 'Update a text element by id in the current Univer presentation. Use univer_slide_list first to inspect element ids.',
    parameters: {
      slideIndex: { type: 'integer', required: true, description: 'Zero-based slide index.' },
      elementId: { type: 'string', required: true, description: 'Text element id.' },
      text: { type: 'string', description: 'Replacement text.' },
      left: { type: 'number', description: 'New left position.' },
      top: { type: 'number', description: 'New top position.' },
      width: { type: 'number', description: 'New width.' },
      height: { type: 'number', description: 'New height.' },
      fontSize: { type: 'number', description: 'New font size.' },
      fontColor: { type: 'string', description: 'New font CSS color.' },
      bold: { type: 'boolean', description: 'Whether text is bold.' },
    },
    output,
    async execute(args, exec) {
      const sessionId = sheetOwnerSessionId(exec.agent)
      if (sessionId === undefined) throw new Error('无法确定当前会话，不能更新文本')
      requireActiveSlide(sessionId)
      return { ok: true, action: 'update-slide-text', message: `已更新幻灯片文本元素“${args.elementId}”。` }
    },
  }))

  ctx.tools.register(defineTool({
    name: 'univer_slide_delete_element',
    description: 'Delete an element by id from a slide in the current Univer presentation. Use univer_slide_list first to inspect element ids.',
    parameters: {
      slideIndex: { type: 'integer', required: true, description: 'Zero-based slide index.' },
      elementId: { type: 'string', required: true, description: 'Element id.' },
    },
    output,
    async execute(args, exec) {
      const sessionId = sheetOwnerSessionId(exec.agent)
      if (sessionId === undefined) throw new Error('无法确定当前会话，不能删除元素')
      requireActiveSlide(sessionId)
      return { ok: true, action: 'delete-slide-element', message: `已删除幻灯片元素“${args.elementId}”。` }
    },
  }))
}
