import React, { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import Prism from 'prismjs'
import 'prismjs/components/prism-typescript'
import 'prismjs/components/prism-jsx'
import 'prismjs/components/prism-tsx'
import 'prismjs/components/prism-json'
import 'prismjs/components/prism-bash'
import 'prismjs/components/prism-python'
import 'prismjs/components/prism-java'
import 'prismjs/components/prism-c'
import 'prismjs/components/prism-cpp'
import 'prismjs/components/prism-csharp'
import 'prismjs/components/prism-go'
import 'prismjs/components/prism-rust'
import 'prismjs/components/prism-sql'
import 'prismjs/components/prism-yaml'
import 'prismjs/components/prism-markdown'

import { MarkdownPreview, isMarkdown, canPreviewMarkdown } from './markdown.js'
import { type DshLanguage, useDshLanguage } from '../../../dsh-language.js'
import {
  appendWorkbookMutation,
  clearWorkbookDraft,
  commitWorkbookDraft,
  loadWorkbookDraft,
  workbookDraftKey,
  type WorkbookMutation,
} from './workbook-drafts.js'
import './styles.css'

type Entry = {
  path: string
  name: string
  type: 'file' | 'directory'
  size?: number
}

type FileBrowserProps = { sessionId: string }

type DirectoryState = {
  entries: Entry[]
  nextCursor: string | null
  total: number
  loaded: boolean
  loading: boolean
  error: string
}

type SaveableUnit = { save?: () => unknown; getSnapshot?: () => unknown }
type Disposable = { dispose: () => void; setLocale?: (locale: unknown) => void }
type UniverSheetRuntime = {
  univer: Disposable
  workbook: SaveableUnit
  replayMutations: (mutations: WorkbookMutation[]) => Promise<void>
  onMutation: (listener: (mutation: WorkbookMutation) => void) => Disposable
}
type UniverDocumentRuntime = { univer: Disposable; univerAPI: { createDocument: (snapshot: unknown) => SaveableUnit; getActiveDocument?: () => SaveableUnit | null } }
type UniverSlidesRuntime = Disposable & { createUnit: (type: unknown, snapshot: unknown) => SaveableUnit }
type SnapshotProviderRef = React.MutableRefObject<(() => unknown | Promise<unknown>) | null>

type RuntimeFactory = (container: HTMLElement, language: DshLanguage) => UniverDocumentRuntime | UniverSlidesRuntime

function univerLocaleId(language: DshLanguage): 'enUS' | 'zhCN' {
  return language === 'zh' ? 'zhCN' : 'enUS'
}

type RuntimeGlobal = {
  createSheetRuntime?: (container: HTMLElement, snapshot: unknown, language: DshLanguage) => UniverSheetRuntime
  createDocsRuntime?: RuntimeFactory
  createSlidesRuntime?: RuntimeFactory
  createSlide?: (runtime: UniverSlidesRuntime, snapshot: unknown) => unknown
}

// Redi is effectively a singleton inside one JavaScript realm for this Univer
// release. Each cached workbook therefore owns an iframe realm: workbooks can
// remain alive while Workspace tabs switch without sharing (and corrupting) an
// injector. On browsers with Element.moveBefore, the iframe is moved to an
// off-screen parking host without resetting its browsing context; older browsers
// safely dispose/rebuild instead of retaining a stale runtime object.
const MAX_CACHED_SHEET_RUNTIMES = 4
const sheetFrameCache = new Map<string, SheetFrameEntry>()
let sheetFrameInstanceSequence = 0
let sheetFrameParkingHost: HTMLDivElement | null = null

type SheetFrameEntry = {
  key: string
  sourceModified: string
  iframe: HTMLIFrameElement
  runtime: UniverSheetRuntime | null
  replayError: unknown | null
  ready: Promise<UniverSheetRuntime>
  rejectReady: ((reason: unknown) => void) | null
  owner: symbol | null
  onMutation: ((mutation: WorkbookMutation) => void) | null
  disposed: boolean
  lastUsed: number
}

function getSheetFrameParkingHost(): HTMLDivElement {
  if (sheetFrameParkingHost?.isConnected) return sheetFrameParkingHost
  const host = document.createElement('div')
  host.className = 'dsh-wfv-sheet-runtime-parking'
  host.setAttribute('aria-hidden', 'true')
  document.body.appendChild(host)
  sheetFrameParkingHost = host
  return host
}

type StatePreservingParent = HTMLElement & {
  moveBefore?: (node: Node, child: Node | null) => void
}

function supportsStatePreservingMove(parent: HTMLElement): parent is StatePreservingParent & Required<Pick<StatePreservingParent, 'moveBefore'>> {
  return typeof (parent as StatePreservingParent).moveBefore === 'function'
}

function moveSheetFrame(parent: HTMLElement, iframe: HTMLIFrameElement): boolean {
  if (parent.isConnected && iframe.isConnected && supportsStatePreservingMove(parent)) {
    parent.moveBefore(iframe, null)
    return true
  }
  parent.appendChild(iframe)
  return false
}

function disposeSheetFrame(entry: SheetFrameEntry): void {
  if (entry.disposed) return
  entry.disposed = true
  if (sheetFrameCache.get(entry.key) === entry) sheetFrameCache.delete(entry.key)
  entry.rejectReady?.(new Error('sheet runtime was disposed before initialization'))
  entry.rejectReady = null
  entry.runtime?.univer.dispose()
  entry.iframe.remove()
}

function advanceSheetFrameVersion(key: string, previousModified: string, nextModified: string): void {
  const entry = sheetFrameCache.get(key)
  if (entry?.sourceModified === previousModified) entry.sourceModified = nextModified
}

function trimSheetFrameCache(): void {
  if (sheetFrameCache.size <= MAX_CACHED_SHEET_RUNTIMES) return
  const parked = [...sheetFrameCache.values()]
    .filter((entry) => entry.owner === null)
    .sort((left, right) => left.lastUsed - right.lastUsed)
  while (sheetFrameCache.size > MAX_CACHED_SHEET_RUNTIMES && parked.length > 0) {
    disposeSheetFrame(parked.shift()!)
  }
}

type OfficePayload = {
  modified?: string
  workbook?: unknown
  document?: unknown
  presentation?: unknown
}

type OfficePayloadCacheEntry = {
  sourceRef: WeakRef<object> | null
  promise: Promise<OfficePayload>
  value?: OfficePayload
  lastUsed: number
}

const MAX_CACHED_OFFICE_PAYLOADS = 16
const officePayloadCache = new Map<string, OfficePayloadCacheEntry>()

function trimOfficePayloadCache(): void {
  const fulfilled = [...officePayloadCache.entries()]
    .filter(([, entry]) => entry.value !== undefined)
    .sort((left, right) => left[1].lastUsed - right[1].lastUsed)
  while (officePayloadCache.size > MAX_CACHED_OFFICE_PAYLOADS && fulfilled.length > 0) {
    officePayloadCache.delete(fulfilled.shift()![0])
  }
}

function loadOfficePayload(key: string, source: object, request: () => Promise<OfficePayload>): OfficePayloadCacheEntry {
  const cached = officePayloadCache.get(key)
  if (cached && cached.sourceRef?.deref() === source) {
    cached.lastUsed = Date.now()
    return cached
  }
  if (cached) officePayloadCache.delete(key)
  const entry: OfficePayloadCacheEntry = {
    sourceRef: new WeakRef(source),
    promise: Promise.resolve(null as never),
    lastUsed: Date.now(),
  }
  entry.promise = request().then((value) => {
    entry.value = value
    trimOfficePayloadCache()
    return value
  }, (reason) => {
    if (officePayloadCache.get(key) === entry) officePayloadCache.delete(key)
    throw reason
  })
  officePayloadCache.set(key, entry)
  trimOfficePayloadCache()
  return entry
}

function updateOfficePayload(key: string, value: OfficePayload): void {
  const sourceRef = officePayloadCache.get(key)?.sourceRef ?? null
  const entry: OfficePayloadCacheEntry = { sourceRef, value, promise: Promise.resolve(value), lastUsed: Date.now() }
  officePayloadCache.set(key, entry)
  trimOfficePayloadCache()
}

function createSheetFrame(
  key: string,
  sourceModified: string,
  initialParent: HTMLElement,
  snapshot: unknown,
  draftMutations: WorkbookMutation[],
  language: DshLanguage,
): SheetFrameEntry {
  const iframe = document.createElement('iframe')
  iframe.className = 'dsh-wfv-sheet-frame'
  iframe.title = '电子表格 Univer 编辑器'
  iframe.inert = true
  iframe.srcdoc = '<!doctype html><html><head><meta charset="utf-8"><style>html,body,#app{box-sizing:border-box;margin:0;width:100%;height:100%;overflow:hidden}#app{position:relative}</style></head><body><div id="app"></div></body></html>'
  const entry: SheetFrameEntry = {
    key,
    sourceModified,
    iframe,
    runtime: null,
    replayError: null,
    ready: Promise.resolve(null as never),
    rejectReady: null,
    owner: null,
    onMutation: null,
    disposed: false,
    lastUsed: Date.now(),
  }
  entry.ready = new Promise<UniverSheetRuntime>((resolve, reject) => {
    entry.rejectReady = reject
    iframe.addEventListener('load', () => {
      const frameDocument = iframe.contentDocument
      const frameWindow = iframe.contentWindow as (Window & { __DSH_WORKSPACE_FILE_VIEWER__?: RuntimeGlobal }) | null
      const mount = frameDocument?.getElementById('app')
      if (!frameDocument || !frameWindow || !mount) {
        reject(new Error('sheet runtime iframe is unavailable'))
        return
      }
      const script = frameDocument.createElement('script')
      script.src = '/api/workspace-file-viewer/runtime/sheet.js'
      script.onload = () => {
        try {
          const createSheetRuntime = frameWindow.__DSH_WORKSPACE_FILE_VIEWER__?.createSheetRuntime
          if (!createSheetRuntime) throw new Error('sheet runtime factory is unavailable')
          const runtime = createSheetRuntime(mount, snapshot, language)
          entry.runtime = runtime
          // Subscribe before replay so no user mutation can fall into the gap;
          // replay commands are already filtered by dshDraftReplay in runtime.
          runtime.onMutation((mutation) => entry.onMutation?.(mutation))
          void runtime.replayMutations(draftMutations).then(() => {
            if (entry.disposed) return
            entry.rejectReady = null
            iframe.inert = false
            resolve(runtime)
          }, (reason) => {
            entry.replayError = reason
            reject(reason)
          })
        } catch (reason) {
          reject(reason)
        }
      }
      script.onerror = () => reject(new Error('failed to load sheet runtime'))
      frameDocument.head.appendChild(script)
    }, { once: true })
  })
  void entry.ready.catch(() => disposeSheetFrame(entry))
  initialParent.appendChild(iframe)
  sheetFrameCache.set(key, entry)
  return entry
}

type CachedOfficeKind = 'docs' | 'slides'
type CachedOfficeFrameEntry = {
  kind: CachedOfficeKind
  key: string
  resumeKey: string
  sourceModified: string
  iframe: HTMLIFrameElement
  ready: Promise<() => unknown>
  rejectReady: ((reason: unknown) => void) | null
  snapshotProvider: (() => unknown) | null
  disposeRuntime: (() => void) | null
  setLocale: ((locale: 'enUS' | 'zhCN') => void) | null
  owner: symbol | null
  disposed: boolean
  lastUsed: number
}

const OFFICE_RUNTIME_LIMITS: Record<CachedOfficeKind, number> = { docs: 3, slides: 2 }
const officeFrameCaches: Record<CachedOfficeKind, Map<string, CachedOfficeFrameEntry>> = {
  docs: new Map(),
  slides: new Map(),
}
const officeFrameBySnapshotProvider = new WeakMap<() => unknown, CachedOfficeFrameEntry>()
type OfficeResumeSnapshot = { sourceModified: string; snapshot: unknown; lastUsed: number }
const officeResumeLimits: Record<CachedOfficeKind, number> = { docs: 8, slides: 4 }
const officeResumeSnapshots: Record<CachedOfficeKind, Map<string, OfficeResumeSnapshot>> = {
  docs: new Map(),
  slides: new Map(),
}
let officeFrameInstanceSequence = 0
let officeFrameParkingHost: HTMLDivElement | null = null

function getOfficeFrameParkingHost(): HTMLDivElement {
  if (officeFrameParkingHost?.isConnected) return officeFrameParkingHost
  const host = document.createElement('div')
  host.className = 'dsh-wfv-office-runtime-parking'
  host.setAttribute('aria-hidden', 'true')
  document.body.appendChild(host)
  officeFrameParkingHost = host
  return host
}

function rememberOfficeSnapshot(entry: CachedOfficeFrameEntry): void {
  if (!entry.snapshotProvider) return
  try {
    const snapshot = entry.snapshotProvider()
    if (snapshot instanceof Promise) return
    const cache = officeResumeSnapshots[entry.kind]
    cache.set(entry.resumeKey, { sourceModified: entry.sourceModified, snapshot, lastUsed: Date.now() })
    const oldest = [...cache.entries()].sort((left, right) => left[1].lastUsed - right[1].lastUsed)
    while (cache.size > officeResumeLimits[entry.kind] && oldest.length > 0) cache.delete(oldest.shift()![0])
  } catch (reason) {
    console.error(`Failed to preserve ${entry.kind} edits before disposing runtime`, reason)
  }
}

function disposeOfficeFrame(entry: CachedOfficeFrameEntry, preserveSnapshot = true): void {
  if (entry.disposed) return
  entry.disposed = true
  if (preserveSnapshot) rememberOfficeSnapshot(entry)
  const cache = officeFrameCaches[entry.kind]
  if (cache.get(entry.key) === entry) cache.delete(entry.key)
  entry.rejectReady?.(new Error(`${entry.kind} runtime was disposed before initialization`))
  entry.rejectReady = null
  entry.disposeRuntime?.()
  entry.iframe.remove()
}

function trimOfficeFrameCache(kind: CachedOfficeKind): void {
  const cache = officeFrameCaches[kind]
  const parked = [...cache.values()]
    .filter((entry) => entry.owner === null)
    .sort((left, right) => left.lastUsed - right.lastUsed)
  while (cache.size > OFFICE_RUNTIME_LIMITS[kind] && parked.length > 0) disposeOfficeFrame(parked.shift()!)
}

function advanceOfficeProviderVersion(provider: () => unknown, previousModified: string, nextModified: string): void {
  const entry = officeFrameBySnapshotProvider.get(provider)
  if (entry?.sourceModified === previousModified) entry.sourceModified = nextModified
}

function createOfficeFrame(
  kind: CachedOfficeKind,
  key: string,
  resumeKey: string,
  sourceModified: string,
  initialParent: HTMLElement,
  snapshot: unknown,
  language: DshLanguage,
): CachedOfficeFrameEntry {
  const iframe = document.createElement('iframe')
  iframe.className = `dsh-wfv-office-frame dsh-wfv-${kind}-frame`
  iframe.title = kind === 'docs' ? '文档 Univer 编辑器' : '演示文稿 Univer 编辑器'
  iframe.inert = true
  iframe.srcdoc = '<!doctype html><html><head><meta charset="utf-8"><style>html,body,#app{box-sizing:border-box;margin:0;width:100%;height:100%;overflow:hidden}#app{position:relative}#app>div{height:100%!important;min-height:0!important}#app .univer-relative.univer-isolate.univer-h-full{position:absolute!important;inset:0!important;height:100%!important;min-height:0!important}#app canvas{display:block}</style></head><body><div id="app"></div></body></html>'
  const entry: CachedOfficeFrameEntry = {
    kind,
    key,
    resumeKey,
    sourceModified,
    iframe,
    ready: Promise.resolve(null as never),
    rejectReady: null,
    snapshotProvider: null,
    disposeRuntime: null,
    setLocale: null,
    owner: null,
    disposed: false,
    lastUsed: Date.now(),
  }
  entry.ready = new Promise<() => unknown>((resolve, reject) => {
    entry.rejectReady = reject
    iframe.addEventListener('load', () => {
      const frameDocument = iframe.contentDocument
      const frameWindow = iframe.contentWindow as (Window & { __DSH_WORKSPACE_FILE_VIEWER__?: RuntimeGlobal }) | null
      const mount = frameDocument?.getElementById('app')
      if (!frameDocument || !frameWindow || !mount) {
        reject(new Error(`${kind} runtime iframe is unavailable`))
        return
      }
      const script = frameDocument.createElement('script')
      script.src = `/api/workspace-file-viewer/runtime/${kind}.js`
      script.onload = () => {
        try {
          const runtimeGlobal = frameWindow.__DSH_WORKSPACE_FILE_VIEWER__
          let snapshotProvider: () => unknown
          if (kind === 'docs') {
            if (!runtimeGlobal?.createDocsRuntime) throw new Error('docs runtime factory is unavailable')
            const runtime = runtimeGlobal.createDocsRuntime(mount, language) as UniverDocumentRuntime
            const documentUnit = runtime.univerAPI.createDocument(snapshot as any)
            entry.disposeRuntime = () => runtime.univer.dispose()
            entry.setLocale = (locale) => runtime.univer.setLocale?.(locale)
            snapshotProvider = () => unitSnapshot(runtime.univerAPI.getActiveDocument?.() ?? documentUnit)
          } else {
            if (!runtimeGlobal?.createSlidesRuntime || !runtimeGlobal.createSlide) throw new Error('slides runtime factory is unavailable')
            const runtime = runtimeGlobal.createSlidesRuntime(mount, language) as UniverSlidesRuntime
            const presentation = runtimeGlobal.createSlide(runtime, snapshot) as SaveableUnit
            entry.disposeRuntime = () => runtime.dispose()
            entry.setLocale = (locale) => runtime.setLocale?.(locale)
            snapshotProvider = () => unitSnapshot(presentation)
          }
          if (entry.disposed) {
            entry.disposeRuntime()
            return
          }
          entry.rejectReady = null
          entry.snapshotProvider = snapshotProvider
          officeFrameBySnapshotProvider.set(snapshotProvider, entry)
          iframe.inert = false
          resolve(snapshotProvider)
        } catch (reason) {
          reject(reason)
        }
      }
      script.onerror = () => reject(new Error(`failed to load ${kind} runtime`))
      frameDocument.head.appendChild(script)
    }, { once: true })
  })
  void entry.ready.catch(() => disposeOfficeFrame(entry))
  initialParent.appendChild(iframe)
  officeFrameCaches[kind].set(key, entry)
  return entry
}

const CODE_LANGUAGES: Record<string, string> = {
  js: 'javascript', mjs: 'javascript', cjs: 'javascript', jsx: 'jsx',
  ts: 'typescript', mts: 'typescript', cts: 'typescript', tsx: 'tsx',
  html: 'markup', htm: 'markup', xml: 'markup', svg: 'markup', vue: 'markup',
  css: 'css', scss: 'css', less: 'css', json: 'json', jsonl: 'json',
  sh: 'bash', bash: 'bash', zsh: 'bash', py: 'python', java: 'java',
  c: 'c', h: 'c', cc: 'cpp', cpp: 'cpp', cxx: 'cpp', hpp: 'cpp',
  cs: 'csharp', go: 'go', rs: 'rust', sql: 'sql', yaml: 'yaml', yml: 'yaml',
  md: 'markdown', mdx: 'markdown',
}

function codeLanguage(path: string): string | null {
  const name = path.split('/').pop()?.toLowerCase() ?? ''
  if (name === 'dockerfile' || name === 'makefile') return 'bash'
  const extension = name.includes('.') ? name.split('.').pop() ?? '' : ''
  return CODE_LANGUAGES[extension] ?? null
}

type OfficeSaveTarget = { unitType: 'sheet' | 'doc' | 'slide'; format: 'xlsx' | 'csv' | 'docx' | 'pptx' }

function workbookFormat(path: string): 'xls' | 'xlsx' | 'csv' | null {
  const match = path.match(/\.(xls|xlsx|csv)$/i)
  return match ? match[1]!.toLowerCase() as 'xls' | 'xlsx' | 'csv' : null
}

function officeSaveTarget(path: string): OfficeSaveTarget | null {
  const sheetFormat = workbookFormat(path)
  if (sheetFormat === 'xls') return { unitType: 'sheet', format: 'xlsx' }
  if (sheetFormat) return { unitType: 'sheet', format: sheetFormat }
  if (/\.docx$/i.test(path)) return { unitType: 'doc', format: 'docx' }
  if (/\.pptx$/i.test(path)) return { unitType: 'slide', format: 'pptx' }
  return null
}

// Prism tokenizes the whole document into a complete HTML string with one DOM
// node per token, and a plain <pre> still forces the browser to build a line box
// per line. Beyond these sizes both are too slow for the main thread, so
// generated/bundled files (e.g. lib/runtimes/*.js, which can exceed 10 MB) fall
// back to plain text. Very large files additionally render through a virtualized
// window: only the visible slice of lines exists in the DOM, so opening and
// scrolling a 200k-line bundle stays responsive.
const HIGHLIGHT_MAX_BYTES = 256 * 1024
const HIGHLIGHT_MAX_LINES = 4000
const VIRTUALIZE_MIN_LINES = 2000
const VIRTUALIZE_MIN_BYTES = 1024 * 1024
// Extra rows rendered above/below the viewport so fast scrolling never flashes.
const VIRTUAL_OVERSCAN = 20
// Must match the `line-height: 1.55` of .dsh-wfv-pre / .dsh-wfv-vpre.
const CODE_LINE_HEIGHT = 1.55

function VirtualizedCode({ path, lines, fontSize }: { path: string; lines: string[]; fontSize: number }) {
  const containerRef = useRef<HTMLDivElement>(null)
  const frameRef = useRef(0)
  const [scrollTop, setScrollTop] = useState(0)
  const [viewportHeight, setViewportHeight] = useState(0)

  useLayoutEffect(() => {
    const container = containerRef.current
    if (container === null || typeof ResizeObserver === 'undefined') return
    setViewportHeight(container.clientHeight)
    const observer = new ResizeObserver(() => setViewportHeight(container.clientHeight))
    observer.observe(container)
    return () => observer.disconnect()
  }, [])

  useEffect(() => () => cancelAnimationFrame(frameRef.current), [])

  const language = codeLanguage(path)
  const lineHeight = fontSize * CODE_LINE_HEIGHT
  const totalHeight = Math.ceil(lines.length * lineHeight)
  const start = Math.max(0, Math.floor(scrollTop / lineHeight) - VIRTUAL_OVERSCAN)
  const end = Math.min(lines.length, Math.ceil((scrollTop + Math.max(viewportHeight, 1)) / lineHeight) + VIRTUAL_OVERSCAN)
  const visible = lines.slice(start, end)
  const style = { '--dsh-wfv-code-font-size': `${fontSize}px` } as React.CSSProperties

  return (
    <div className="dsh-wfv-code dsh-wfv-vcode" style={style}>
      <div className="dsh-wfv-code-language">{language ?? '纯文本'} · {lines.length.toLocaleString()} 行 · 大文件已虚拟化渲染</div>
      <div
        ref={containerRef}
        className="dsh-wfv-vscroll"
        onScroll={() => {
          const container = containerRef.current
          if (container === null || frameRef.current !== 0) return
          frameRef.current = requestAnimationFrame(() => {
            frameRef.current = 0
            setScrollTop(container.scrollTop)
          })
        }}
      >
        <div className="dsh-wfv-vspacer" style={{ height: totalHeight, paddingTop: start * lineHeight }}>
          <pre className="dsh-wfv-pre dsh-wfv-vpre"><code>{visible.join('\n')}</code></pre>
        </div>
      </div>
    </div>
  )
}

function CodePreview({ path, content, fontSize }: { path: string; content: string; fontSize: number }) {
  const language = codeLanguage(path)
  const style = { '--dsh-wfv-code-font-size': `${fontSize}px` } as React.CSSProperties

  // Split once per file (not per render): the previous version re-split and
  // re-tokenized the whole document on every font-size change or tree toggle.
  const normalized = useMemo(() => content.replace(/\r\n/g, '\n'), [content])
  const lines = useMemo(() => normalized.split('\n'), [normalized])
  const byteCount = content.length
  const lineCount = lines.length

  const canHighlight = language !== null
    && Prism.languages[language] !== undefined
    && byteCount <= HIGHLIGHT_MAX_BYTES
    && lineCount <= HIGHLIGHT_MAX_LINES
  const shouldVirtualize = lineCount >= VIRTUALIZE_MIN_LINES || byteCount >= VIRTUALIZE_MIN_BYTES

  const highlightedHtml = useMemo(() => {
    if (!canHighlight || language === null) return null
    const grammar = Prism.languages[language]
    if (grammar === undefined) return null
    return Prism.highlight(content, grammar, language)
  }, [canHighlight, content, language])

  if (shouldVirtualize) {
    return <VirtualizedCode key={path} path={path} lines={lines} fontSize={fontSize} />
  }
  if (highlightedHtml === null) {
    return (
      <div className="dsh-wfv-code dsh-wfv-code-large" style={style}>
        <div className="dsh-wfv-code-language">纯文本</div>
        <pre className="dsh-wfv-pre dsh-wfv-code-plain"><code>{normalized}</code></pre>
      </div>
    )
  }
  return (
    <div className="dsh-wfv-code" style={style}>
      <div className="dsh-wfv-code-language">{language}</div>
      <pre className={`dsh-wfv-pre language-${language}`}><code dangerouslySetInnerHTML={{ __html: highlightedHtml }} /></pre>
    </div>
  )
}

// These normalizers run once, at load time, on the fresh fetch payload — never
// on React state, and never through JSON.parse(JSON.stringify(...)). The old
// deep clone materialized a single string of the whole document/unit, which
// exceeds the engine's string/allocation limit on large files and aborts the
// preview with "allocation size overflow" / "Invalid string length".
function normalizeWorkbookSnapshot(value: unknown): any {
  if (value === null || typeof value !== 'object') return value
  const snapshot = value as any
  if (snapshot.sheets !== null && typeof snapshot.sheets === 'object') {
    for (const sheet of Object.values(snapshot.sheets) as any[]) {
      if (sheet === null || typeof sheet !== 'object') continue
      if (!Number.isFinite(sheet.rowCount) || sheet.rowCount < 1) sheet.rowCount = 100
      if (!Number.isFinite(sheet.columnCount) || sheet.columnCount < 1) sheet.columnCount = 26
      if (!Number.isFinite(sheet.defaultColumnWidth) || sheet.defaultColumnWidth <= 0) sheet.defaultColumnWidth = 100
      if (sheet.columnHeader !== null && typeof sheet.columnHeader === 'object' && sheet.columnHeader.width <= 0) {
        sheet.columnHeader.width = 100
      }
      if (sheet.columnData !== null && typeof sheet.columnData === 'object') {
        for (const column of Object.values(sheet.columnData) as any[]) {
          if (column !== null && typeof column === 'object' && column.w !== undefined && (!Number.isFinite(column.w) || column.w <= 0)) {
            delete column.w
          }
        }
      }
    }
  }
  return snapshot
}

function unitSnapshot(unit: SaveableUnit | null | undefined): unknown {
  if (typeof unit?.save === 'function') return unit.save()
  if (typeof unit?.getSnapshot === 'function') return unit.getSnapshot()
  throw new Error('当前 Univer 运行时无法生成可保存的数据')
}

type DraftIdentity = {
  fileKey: string
  baseModified: string
  lastSequence: number
}

function useWorkbookDraftManager() {
  const [mutations, setMutations] = useState<WorkbookMutation[]>([])
  const [status, setStatus] = useState('')
  const identityRef = useRef<DraftIdentity | null>(null)
  const generationRef = useRef(0)
  const writeChainRef = useRef<Promise<void>>(Promise.resolve())
  const pendingWritesRef = useRef(0)
  const hasWriteFailureRef = useRef(false)
  const shouldWarnBeforeUnloadRef = useRef(false)
  const releaseLockRef = useRef<(() => void) | null>(null)
  const saveLeasesRef = useRef(0)
  const releasePendingRef = useRef(false)

  useEffect(() => {
    const beforeUnload = (event: BeforeUnloadEvent) => {
      if (!shouldWarnBeforeUnloadRef.current) return
      event.preventDefault()
      event.returnValue = ''
    }
    window.addEventListener('beforeunload', beforeUnload)
    return () => {
      generationRef.current += 1
      window.removeEventListener('beforeunload', beforeUnload)
      if (saveLeasesRef.current > 0) releasePendingRef.current = true
      else {
        releaseLockRef.current?.()
        releaseLockRef.current = null
      }
    }
  }, [])

  const releaseLockAfterWrites = useCallback(() => {
    if (saveLeasesRef.current > 0) {
      releasePendingRef.current = true
      return
    }
    const release = releaseLockRef.current
    releaseLockRef.current = null
    releasePendingRef.current = false
    if (release) writeChainRef.current = writeChainRef.current.then(release, release)
  }, [])

  const reset = useCallback(() => {
    generationRef.current += 1
    releaseLockAfterWrites()
    identityRef.current = null
    shouldWarnBeforeUnloadRef.current = pendingWritesRef.current > 0 || hasWriteFailureRef.current
    setMutations([])
    setStatus('')
  }, [releaseLockAfterWrites])

  const restore = useCallback(async (sessionId: string, path: string, baseModified: string): Promise<WorkbookMutation[] | null> => {
    const generation = ++generationRef.current
    const fileKey = workbookDraftKey(sessionId, path)
    setStatus('正在检查本地草稿…')
    await writeChainRef.current
    if (generationRef.current !== generation) return null

    if (navigator.locks) {
      const lockHandle: { release?: () => void } = {}
      let resolveAcquired!: (acquired: boolean) => void
      const acquired = new Promise<boolean>((resolve) => { resolveAcquired = resolve })
      void navigator.locks.request(`dsh-office-one:${fileKey}`, { mode: 'exclusive', ifAvailable: true }, async (lock) => {
        resolveAcquired(lock !== null)
        if (lock !== null) await new Promise<void>((resolve) => { lockHandle.release = resolve })
      }).catch(() => resolveAcquired(false))
      if (!await acquired) {
        if (generationRef.current === generation) setStatus('此表格已在另一个页面中编辑，请关闭另一页面后重试')
        return null
      }
      if (generationRef.current !== generation) {
        lockHandle.release?.()
        return null
      }
      releaseLockRef.current = () => lockHandle.release?.()
    }

    try {
      const loaded = await loadWorkbookDraft(fileKey, baseModified)
      if (generationRef.current !== generation) return []
      if (loaded.conflictModified !== undefined) {
        const discard = window.confirm('检测到本地草稿，但 Workspace 文件已经变化。\n\n选择“确定”将放弃旧草稿并打开最新文件；选择“取消”会保留草稿且停止打开。')
        if (!discard) {
          setStatus('已保留冲突草稿；请先处理文件版本冲突')
          releaseLockRef.current?.()
          releaseLockRef.current = null
          return null
        }
        await clearWorkbookDraft(fileKey)
        if (generationRef.current !== generation) return null
        identityRef.current = { fileKey, baseModified, lastSequence: 0 }
        setMutations([])
        setStatus('旧草稿已放弃，已打开文件最新版本')
        return []
      }
      identityRef.current = { fileKey, baseModified, lastSequence: loaded.lastSequence }
      setMutations(loaded.mutations)
      if (loaded.mutations.length > 0) {
        const ignored = loaded.skippedMutations > 0 ? `，已忽略 ${loaded.skippedMutations} 条临时编辑器操作` : ''
        setStatus(`正在恢复 ${loaded.mutations.length} 条本地修改${ignored}…`)
      } else {
        setStatus(loaded.skippedMutations > 0 ? `已忽略 ${loaded.skippedMutations} 条无须恢复的临时编辑器操作` : '')
      }
      return loaded.mutations
    } catch (reason) {
      if (generationRef.current !== generation) return []
      identityRef.current = { fileKey, baseModified, lastSequence: 0 }
      setMutations([])
      setStatus(`草稿读取失败：${reason instanceof Error ? reason.message : String(reason)}`)
      return []
    }
  }, [])

  const record = useCallback((mutation: WorkbookMutation) => {
    const identity = identityRef.current
    if (!identity) return
    const generation = generationRef.current
    const fileKey = identity.fileKey
    pendingWritesRef.current += 1
    shouldWarnBeforeUnloadRef.current = true
    setStatus('正在保存本地草稿…')

    writeChainRef.current = writeChainRef.current
      .then(async () => {
        const sequence = await appendWorkbookMutation(fileKey, identity.baseModified, mutation)
        if (generationRef.current === generation && identityRef.current?.fileKey === fileKey) {
          identityRef.current.lastSequence = Math.max(identityRef.current.lastSequence, sequence)
        }
      })
      .then(() => {
        pendingWritesRef.current = Math.max(0, pendingWritesRef.current - 1)
        if (pendingWritesRef.current === 0 && !hasWriteFailureRef.current) shouldWarnBeforeUnloadRef.current = false
        if (generationRef.current === generation && identityRef.current?.fileKey === fileKey && pendingWritesRef.current === 0) {
          setStatus('草稿已保存')
        }
      }, (reason) => {
        pendingWritesRef.current = Math.max(0, pendingWritesRef.current - 1)
        hasWriteFailureRef.current = true
        shouldWarnBeforeUnloadRef.current = true
        if (generationRef.current === generation && identityRef.current?.fileKey === fileKey) {
          setStatus(`草稿保存失败：${reason instanceof Error ? reason.message : String(reason)}`)
        }
      })
  }, [])

  const prepareSave = useCallback(async () => {
    await writeChainRef.current
    const identity = identityRef.current
    if (!identity) return null
    saveLeasesRef.current += 1
    return { ...identity }
  }, [])

  const finishSave = useCallback(() => {
    saveLeasesRef.current = Math.max(0, saveLeasesRef.current - 1)
    if (saveLeasesRef.current === 0 && releasePendingRef.current) releaseLockAfterWrites()
  }, [releaseLockAfterWrites])

  const commitSaved = useCallback(async (prepared: DraftIdentity | null, nextBaseModified: string) => {
    const identity = identityRef.current
    if (!prepared || !identity || identity.fileKey !== prepared.fileKey) return
    const generation = generationRef.current
    const commitTask = writeChainRef.current.then(async () => {
      const remaining = await commitWorkbookDraft(
        prepared.fileKey,
        prepared.baseModified,
        nextBaseModified,
        prepared.lastSequence,
      )
      identity.baseModified = nextBaseModified
      if (generationRef.current !== generation || identityRef.current !== identity) return
      if (!remaining) identity.lastSequence = 0
      hasWriteFailureRef.current = false
      shouldWarnBeforeUnloadRef.current = pendingWritesRef.current > 0
      setMutations([])
      setStatus(remaining ? '文件已保存；之后的修改已存为草稿' : '')
    })
    writeChainRef.current = commitTask.catch(() => undefined)
    try {
      await commitTask
    } catch (reason) {
      shouldWarnBeforeUnloadRef.current = true
      if (generationRef.current === generation) {
        setStatus(`文件已写入，但草稿整理失败：${reason instanceof Error ? reason.message : String(reason)}`)
      }
      throw reason
    }
  }, [])

  const markRestored = useCallback(() => {
    setStatus((current) => current.startsWith('正在恢复') ? '已恢复本地草稿，尚未写回文件' : current)
  }, [])

  const markRestoreFailed = useCallback((reason: unknown) => {
    shouldWarnBeforeUnloadRef.current = true
    setStatus(`草稿恢复失败：${reason instanceof Error ? reason.message : String(reason)}`)
  }, [])

  return { mutations, status, reset, restore, record, prepareSave, finishSave, commitSaved, markRestored, markRestoreFailed }
}

type DocumentPointMetadata = { startIndex: number; [key: string]: unknown }
type DocumentBodySnapshot = {
  dataStream?: unknown
  paragraphs?: DocumentPointMetadata[]
  sectionBreaks?: DocumentPointMetadata[]
  [key: string]: unknown
}

function tokenPositions(dataStream: string, token: string): number[] {
  const positions: number[] = []
  for (let index = 0; index < dataStream.length; index += 1) {
    if (dataStream[index] === token) positions.push(index)
  }
  return positions
}

function normalizePointMetadata(
  metadata: unknown,
  positions: number[],
  createMissing: (startIndex: number, ordinal: number) => DocumentPointMetadata,
): DocumentPointMetadata[] {
  // Already aligned with the current dataStream (the common case when the DOCX
  // converter and the runtime are from the same Univer release): return the
  // original array untouched so large documents are not copied at all.
  const source = Array.isArray(metadata) ? metadata as DocumentPointMetadata[] : null
  if (source !== null && source.length === positions.length) {
    let aligned = true
    for (let index = 0; index < positions.length; index += 1) {
      const item = source[index]
      if (item === null || typeof item !== 'object' || item.startIndex !== positions[index]) {
        aligned = false
        break
      }
    }
    if (aligned) return source
  }

  // DOCX converters from a different Univer release can return point metadata
  // using the previous token offsets. Re-associate it by document order: paragraph
  // and section metadata describe point tokens and therefore have the same order as
  // their corresponding CR/LF sentinels in dataStream. Only this rebuild allocates,
  // and only the touched arrays/items are copied — everything else stays shared.
  const result: DocumentPointMetadata[] = new Array(positions.length)
  for (let ordinal = 0; ordinal < positions.length; ordinal += 1) {
    const current = source?.[ordinal]
    result[ordinal] = current !== null && typeof current === 'object' && Number.isInteger((current as any).startIndex)
      ? { ...(current as object), startIndex: positions[ordinal]! } as DocumentPointMetadata
      : createMissing(positions[ordinal]!, ordinal)
  }
  return result
}

function normalizeDocumentBody(body: unknown, segment: string): void {
  if (body === null || typeof body !== 'object') return
  const target = body as DocumentBodySnapshot
  if (typeof target.dataStream !== 'string') return
  const paragraphPositions = tokenPositions(target.dataStream, '\r')
  const sectionBreakPositions = tokenPositions(target.dataStream, '\n')
  target.paragraphs = normalizePointMetadata(target.paragraphs, paragraphPositions, (startIndex, ordinal) => ({
    startIndex,
    paragraphId: `dsh-paragraph-${segment}-${ordinal}`,
  }))
  target.sectionBreaks = normalizePointMetadata(target.sectionBreaks, sectionBreakPositions, (startIndex, ordinal) => ({
    startIndex,
    sectionId: `dsh-section-${segment}-${ordinal}`,
  }))
}

function normalizeDocumentSnapshot(value: unknown): unknown {
  if (value === null || typeof value !== 'object') return value
  const snapshot = value as any
  normalizeDocumentBody(snapshot.body, 'body')
  for (const [id, header] of Object.entries(snapshot.headers ?? {}) as Array<[string, any]>) normalizeDocumentBody(header?.body, `header-${id}`)
  for (const [id, footer] of Object.entries(snapshot.footers ?? {}) as Array<[string, any]>) normalizeDocumentBody(footer?.body, `footer-${id}`)
  return snapshot
}

function UniverWorkbook({
  runtimeKey,
  sourceModified,
  snapshot,
  snapshotProviderRef,
  draftMutations = [],
  onDraftMutation,
  onDraftRestored,
  onDraftRestoreFailed,
  editingDisabled = false,
}: {
  runtimeKey: string
  sourceModified: string
  snapshot: unknown
  snapshotProviderRef: SnapshotProviderRef
  draftMutations?: WorkbookMutation[]
  onDraftMutation?: (mutation: WorkbookMutation) => void
  onDraftRestored?: () => void
  onDraftRestoreFailed?: (reason: unknown) => void
  editingDisabled?: boolean
}) {
  const language = useDshLanguage()
  const containerRef = useRef<HTMLDivElement>(null)
  const onDraftMutationRef = useRef(onDraftMutation)
  const onDraftRestoredRef = useRef(onDraftRestored)
  const onDraftRestoreFailedRef = useRef(onDraftRestoreFailed)
  const [layoutVersion, setLayoutVersion] = useState(0)
  onDraftMutationRef.current = onDraftMutation
  onDraftRestoredRef.current = onDraftRestored
  onDraftRestoreFailedRef.current = onDraftRestoreFailed

  useEffect(() => {
    if (containerRef.current) containerRef.current.inert = editingDisabled
  }, [editingDisabled])

  useLayoutEffect(() => {
    const container = containerRef.current
    if (container === null || typeof ResizeObserver === 'undefined') return
    const observer = new ResizeObserver(() => {
      const bounds = container.getBoundingClientRect()
      if (bounds.width > 0 && bounds.height > 0 && !container.querySelector('.dsh-wfv-sheet-frame')) {
        setLayoutVersion((value) => value + 1)
      }
    })
    observer.observe(container)
    return () => observer.disconnect()
  }, [])

  useLayoutEffect(() => {
    const container = containerRef.current
    if (container === null) return
    const bounds = container.getBoundingClientRect()
    if (bounds.width <= 0 || bounds.height <= 0) return

    const owner = Symbol(runtimeKey)
    let cached = sheetFrameCache.get(runtimeKey)
    const occupiedByAnotherView = cached?.owner !== null && cached?.owner !== undefined
    if (occupiedByAnotherView) cached = undefined
    // Never let a live workbook cross a server version boundary: otherwise a
    // later save could overwrite a newer file with the stale cached runtime.
    if (cached && cached.sourceModified !== sourceModified) {
      disposeSheetFrame(cached)
      cached = undefined
    }
    // appendChild/reparenting reloads an iframe in older browsers. Fall back to
    // a clean runtime rebuild there rather than retaining a stale runtime object.
    if (cached && !supportsStatePreservingMove(container)) {
      disposeSheetFrame(cached)
      cached = undefined
    }
    const entry = cached ?? createSheetFrame(
      occupiedByAnotherView ? `${runtimeKey}\u0000view-${++sheetFrameInstanceSequence}` : runtimeKey,
      sourceModified,
      container,
      snapshot,
      draftMutations,
      language,
    )
    const reused = cached !== undefined
    entry.owner = owner
    entry.onMutation = (mutation) => onDraftMutationRef.current?.(mutation)
    entry.lastUsed = Date.now()
    if (reused) moveSheetFrame(container, entry.iframe)
    trimSheetFrameCache()

    let cancelled = false
    void entry.ready.then((runtime) => {
      if (cancelled || entry.owner !== owner) return
      runtime.univer.setLocale?.(univerLocaleId(language))
      snapshotProviderRef.current = () => unitSnapshot(runtime.workbook)
      if (entry.replayError !== null) onDraftRestoreFailedRef.current?.(entry.replayError)
      else if (draftMutations.length > 0) onDraftRestoredRef.current?.()
      const frameWindow = entry.iframe.contentWindow
      if (frameWindow) {
        const FrameEvent = (frameWindow as unknown as { Event: typeof Event }).Event
        frameWindow.dispatchEvent(new FrameEvent('resize'))
      }
    }).catch((reason) => {
      if (!cancelled && entry.owner === owner) {
        console.error('Failed to load Sheet preview', reason)
        onDraftRestoreFailedRef.current?.(reason)
        disposeSheetFrame(entry)
      }
    })

    return () => {
      cancelled = true
      if (entry.owner !== owner) return
      snapshotProviderRef.current = null
      entry.onMutation = null
      entry.owner = null
      entry.lastUsed = Date.now()
      if (occupiedByAnotherView) {
        disposeSheetFrame(entry)
        return
      }
      const parkingHost = getSheetFrameParkingHost()
      if (parkingHost.isConnected && entry.iframe.isConnected && supportsStatePreservingMove(parkingHost)) {
        moveSheetFrame(parkingHost, entry.iframe)
      } else {
        disposeSheetFrame(entry)
      }
      trimSheetFrameCache()
    }
  }, [runtimeKey, sourceModified, snapshot, layoutVersion, snapshotProviderRef, language])

  return <div ref={containerRef} className="dsh-wfv-univer" aria-label="电子表格 Univer 预览" />
}

function CachedOfficeUnit({
  kind,
  runtimeKey,
  sourceModified,
  snapshot,
  snapshotProviderRef,
  editingDisabled = false,
}: {
  kind: CachedOfficeKind
  runtimeKey: string
  sourceModified: string
  snapshot: unknown
  snapshotProviderRef: SnapshotProviderRef
  editingDisabled?: boolean
}) {
  const language = useDshLanguage()
  const containerRef = useRef<HTMLDivElement>(null)
  const retryCountRef = useRef(0)
  const [layoutVersion, setLayoutVersion] = useState(0)

  useEffect(() => {
    if (containerRef.current) containerRef.current.inert = editingDisabled
  }, [editingDisabled])

  useLayoutEffect(() => {
    const container = containerRef.current
    if (container === null || typeof ResizeObserver === 'undefined') return
    let frame = 0
    const notifyResize = () => {
      cancelAnimationFrame(frame)
      frame = requestAnimationFrame(() => {
        const iframe = container.querySelector<HTMLIFrameElement>('.dsh-wfv-office-frame')
        if (!iframe) {
          const bounds = container.getBoundingClientRect()
          if (bounds.width > 0 && bounds.height > 0) setLayoutVersion((value) => value + 1)
          return
        }
        const frameWindow = iframe.contentWindow
        if (frameWindow) {
          const FrameEvent = (frameWindow as unknown as { Event: typeof Event }).Event
          frameWindow.dispatchEvent(new FrameEvent('resize'))
        }
      })
    }
    const observer = new ResizeObserver(notifyResize)
    observer.observe(container)
    notifyResize()
    return () => {
      observer.disconnect()
      cancelAnimationFrame(frame)
    }
  }, [])

  useLayoutEffect(() => {
    const container = containerRef.current
    if (container === null) return
    const bounds = container.getBoundingClientRect()
    if (bounds.width <= 0 || bounds.height <= 0) return

    const owner = Symbol(`${kind}:${runtimeKey}`)
    const cache = officeFrameCaches[kind]
    let cached = cache.get(runtimeKey)
    const occupiedByAnotherView = cached?.owner !== null && cached?.owner !== undefined
    if (occupiedByAnotherView) cached = undefined
    if (cached && cached.sourceModified !== sourceModified) {
      disposeOfficeFrame(cached)
      cached = undefined
    }
    if (cached && !supportsStatePreservingMove(container)) {
      disposeOfficeFrame(cached)
      cached = undefined
    }
    const resumeCache = officeResumeSnapshots[kind]
    const resume = resumeCache.get(runtimeKey)
    const initialSnapshot = !cached && resume?.sourceModified === sourceModified ? resume.snapshot : snapshot
    if (!cached && resume) resumeCache.delete(runtimeKey)
    const entry = cached ?? createOfficeFrame(
      kind,
      occupiedByAnotherView ? `${runtimeKey}\u0000view-${++officeFrameInstanceSequence}` : runtimeKey,
      runtimeKey,
      sourceModified,
      container,
      initialSnapshot,
      language,
    )
    const reused = cached !== undefined
    entry.owner = owner
    entry.lastUsed = Date.now()
    entry.iframe.inert = true
    if (reused) moveSheetFrame(container, entry.iframe)
    trimOfficeFrameCache(kind)

    let cancelled = false
    const resizeTimers: number[] = []
    void entry.ready.then((getSnapshot) => {
      if (cancelled || entry.owner !== owner) return
      entry.setLocale?.(univerLocaleId(language))
      snapshotProviderRef.current = getSnapshot
      retryCountRef.current = 0
      entry.iframe.inert = false
      const notifyResize = (delay: number) => {
        resizeTimers.push(window.setTimeout(() => {
          if (cancelled || entry.owner !== owner) return
          const frameWindow = entry.iframe.contentWindow
          if (!frameWindow) return
          const FrameEvent = (frameWindow as unknown as { Event: typeof Event }).Event
          frameWindow.dispatchEvent(new FrameEvent('resize'))
        }, delay))
      }
      notifyResize(0)
      if (kind === 'slides') {
        notifyResize(100)
        notifyResize(300)
        notifyResize(800)
      }
    }).catch((reason) => {
      if (cancelled || entry.owner !== owner) return
      console.error(`Failed to load ${kind} preview`, reason)
      if (retryCountRef.current < 1) {
        retryCountRef.current += 1
        setLayoutVersion((value) => value + 1)
      }
    })

    return () => {
      cancelled = true
      resizeTimers.forEach((timer) => window.clearTimeout(timer))
      if (entry.owner !== owner) return
      snapshotProviderRef.current = null
      entry.owner = null
      entry.lastUsed = Date.now()
      entry.iframe.inert = true
      if (occupiedByAnotherView) {
        disposeOfficeFrame(entry)
        return
      }
      const parkingHost = getOfficeFrameParkingHost()
      if (parkingHost.isConnected && entry.iframe.isConnected && supportsStatePreservingMove(parkingHost)) {
        moveSheetFrame(parkingHost, entry.iframe)
      } else {
        disposeOfficeFrame(entry)
      }
      trimOfficeFrameCache(kind)
    }
  }, [kind, runtimeKey, sourceModified, snapshot, layoutVersion, snapshotProviderRef, language])

  const className = kind === 'docs' ? 'dsh-wfv-univer-doc' : 'dsh-wfv-univer-slides'
  const label = kind === 'docs' ? 'DOC Univer 预览' : 'PPT Univer 预览'
  return <div ref={containerRef} className={`dsh-wfv-univer ${className}`} aria-label={label} />
}

function UniverDocument(props: Omit<React.ComponentProps<typeof CachedOfficeUnit>, 'kind'>) {
  return <CachedOfficeUnit kind="docs" {...props} />
}

function UniverSlides(props: Omit<React.ComponentProps<typeof CachedOfficeUnit>, 'kind'>) {
  return <CachedOfficeUnit kind="slides" {...props} />
}

/*
  const iframeRef = useRef<HTMLIFrameElement>(null)
  useEffect(() => {
    const iframe = iframeRef.current
    if (!iframe) return
    const send = () => {
      console.info('[workspace-file-viewer] Slides snapshot', Object.keys(((snapshot as any)?.slides ?? {})).length + ' slides')
      iframe.contentWindow?.postMessage({ type: 'dsh-wfv-slide', snapshot }, window.location.origin)
    }
    const report = (event: MessageEvent) => { if (event.data?.type === 'dsh-wfv-slide-log') console.info('[workspace-file-viewer] Slides iframe:', event.data.message, event.data.details ?? '') }
    iframe.addEventListener('load', send)
    window.addEventListener('message', report)
    if (iframe.contentDocument?.readyState === 'complete') send()
    return () => { iframe.removeEventListener('load', send); window.removeEventListener('message', report) }
  }, [snapshot])
  const scriptEnd = '<' + '/script>'
  const srcDoc = `<!doctype html><html><head><style>html,body{margin:0;width:100%;height:100%;overflow:hidden}body{position:relative}#app{position:absolute;inset:0;width:auto;height:auto;overflow:hidden}#app>div{height:100%!important;min-height:0!important}#app .univer-relative.univer-isolate.univer-h-full{position:absolute!important;inset:0!important;height:100%!important;min-height:0!important}#app canvas{display:block}</style></head><body><div id="app"></div><script src="/api/workspace-file-viewer/runtime/slides.js">${scriptEnd}<script>function log(m,d){parent.postMessage({type:'dsh-wfv-slide-log',message:m,details:d},'*')}window.addEventListener('error',function(e){log('error',e.message)});window.addEventListener('message',function(e){if(e.data?.type!=='dsh-wfv-slide')return;try{var g=window.__DSH_WORKSPACE_FILE_VIEWER__;log('factories',Object.keys(g||{}));if(!g?.createSlidesRuntime||!g?.createSlide)throw new Error('Slides runtime factory unavailable');var r=g.createSlidesRuntime(document.getElementById('app'));log('runtime created',Object.keys(r||{}));var p=g.createSlide(r,e.data.snapshot);var app=document.getElementById('app');function repair(){var canv=app.querySelectorAll('canvas');Array.from(canv).forEach(function(c){var n=c.parentElement;while(n&&n!==app){if(n.getBoundingClientRect().height===0){n.style.setProperty('height',app.getBoundingClientRect().height+'px','important');n.style.setProperty('min-height','0','important');n.style.setProperty('flex','1 1 auto','important')}n=n.parentElement}});window.dispatchEvent(new Event('resize'))}repair();var repairTimer=setInterval(repair,250);setTimeout(function(){clearInterval(repairTimer);repair();var first=p.getSlideByIndex(0);if(first)p.setActiveSlide(first);window.dispatchEvent(new Event('resize'));},1000);var slides=p.getSlides();var rect=app.getBoundingClientRect();setTimeout(function(){var els=app.querySelectorAll('*');var canv=app.querySelectorAll('canvas');var visible=Array.from(els).filter(function(x){var q=x.getBoundingClientRect();return q.width>0&&q.height>0}).length;log('presentation created',JSON.stringify({slides:slides.length,active:e.data.snapshot.activeSlideId,app:[rect.width,rect.height],elements:els.length,visible:visible,canvas:canv.length,canvasSizes:Array.from(canv).map(function(x){var a=[];var n=x;for(var i=0;n&&i<5;i++,n=n.parentElement)a.push([n.tagName,n.className,n.getBoundingClientRect().width,n.getBoundingClientRect().height]);return [x.width,x.height,x.getBoundingClientRect().width,x.getBoundingClientRect().height,a]})}));},500);}catch(x){log('create failed',String(x&&x.stack||x))}})${scriptEnd}</body></html>`
  return <iframe ref={iframeRef} className="dsh-wfv-univer dsh-wfv-univer-slides" style={{ width: '100%', height: '100%', border: 0, display: 'block' }} srcDoc={srcDoc} title="PPT Univer 预览" />
*/

function DirectoryLoadMore({ depth, onLoad }: { depth: number; onLoad: () => void }) {
  const ref = useRef<HTMLButtonElement>(null)
  const triggeredRef = useRef(false)

  const trigger = () => {
    if (triggeredRef.current) return
    triggeredRef.current = true
    onLoad()
  }

  useEffect(() => {
    const element = ref.current
    if (element === null || typeof IntersectionObserver === 'undefined') return
    const observer = new IntersectionObserver((records) => {
      if (records.some((record) => record.isIntersecting)) trigger()
    }, { rootMargin: '120px 0px' })
    observer.observe(element)
    return () => observer.disconnect()
  }, [onLoad])

  return (
    <button ref={ref} className="dsh-wfv-load-more" style={{ paddingLeft: 22 + depth * 14 }} onClick={trigger}>
      继续加载…
    </button>
  )
}

function FileBrowser({ sessionId }: FileBrowserProps) {
  const [directories, setDirectories] = useState<Record<string, DirectoryState>>({})
  const [expanded, setExpanded] = useState<Set<string>>(new Set())
  const [selected, setSelected] = useState('')
  const [content, setContent] = useState<string | null>(null)
  const [workbook, setWorkbook] = useState<unknown>(null)
  const [documentData, setDocumentData] = useState<unknown>(null)
  const [slidesData, setSlidesData] = useState<unknown>(null)
  const [loading, setLoading] = useState(true)
  const [fileLoading, setFileLoading] = useState(false)
  const [codeFontSize, setCodeFontSize] = useState(14)
  const [markdownSource, setMarkdownSource] = useState(false)
  const markdownFile = isMarkdown(selected)
  const markdownAllowed = useMemo(() => content !== null && canPreviewMarkdown(content), [content])
  const [error, setError] = useState('')
  const [saving, setSaving] = useState(false)
  const [saveStatus, setSaveStatus] = useState('')
  const [loadedModified, setLoadedModified] = useState('')
  const directoryRequestsRef = useRef<Map<string, AbortController>>(new Map())
  const fileRequestRef = useRef<AbortController | null>(null)
  const snapshotProviderRef = useRef<(() => unknown) | null>(null)
  const saveRequestVersionRef = useRef(0)
  const workbookDraft = useWorkbookDraftManager()

  const loadDirectory = async (path: string, cursor = '') => {
    directoryRequestsRef.current.get(path)?.abort()
    const controller = new AbortController()
    directoryRequestsRef.current.set(path, controller)
    setDirectories((current) => ({
      ...current,
      [path]: {
        entries: cursor ? current[path]?.entries ?? [] : [],
        nextCursor: current[path]?.nextCursor ?? null,
        total: current[path]?.total ?? 0,
        loaded: cursor ? current[path]?.loaded ?? false : false,
        loading: true,
        error: '',
      },
    }))
    if (path === '' && cursor === '') setLoading(true)
    try {
      const query = new URLSearchParams({ sessionId, path, limit: '100' })
      if (cursor) query.set('cursor', cursor)
      const response = await fetch(`/api/workspace-files?${query}`, { signal: controller.signal })
      const payload = await response.json()
      if (!response.ok) throw new Error(payload.error || '加载目录失败')
      if (directoryRequestsRef.current.get(path) !== controller) return
      const pageEntries: Entry[] = Array.isArray(payload.entries)
        ? payload.entries.filter((entry: unknown): entry is Entry => {
          if (entry === null || typeof entry !== 'object') return false
          const candidate = entry as Partial<Entry>
          return typeof candidate.path === 'string' && typeof candidate.name === 'string' && (candidate.type === 'file' || candidate.type === 'directory')
        })
        : []
      setDirectories((current) => {
        const combined = cursor ? [...(current[path]?.entries ?? []), ...pageEntries] : pageEntries
        const entries = [...new Map(combined.map((entry) => [entry.path, entry])).values()]
        return {
          ...current,
          [path]: {
            entries,
            nextCursor: typeof payload.nextCursor === 'string' ? payload.nextCursor : null,
            total: typeof payload.total === 'number' ? payload.total : entries.length,
            loaded: true,
            loading: false,
            error: '',
          },
        }
      })
    } catch (reason) {
      if (directoryRequestsRef.current.get(path) !== controller) return
      const message = reason instanceof Error ? reason.message : String(reason)
      setDirectories((current) => ({
        ...current,
        [path]: {
          entries: current[path]?.entries ?? [],
          nextCursor: current[path]?.nextCursor ?? null,
          total: current[path]?.total ?? 0,
          loaded: current[path]?.loaded ?? false,
          loading: false,
          error: message,
        },
      }))
    } finally {
      if (directoryRequestsRef.current.get(path) === controller) {
        directoryRequestsRef.current.delete(path)
        if (path === '' && cursor === '') setLoading(false)
      }
    }
  }

  const refreshTree = () => {
    for (const controller of directoryRequestsRef.current.values()) controller.abort()
    directoryRequestsRef.current.clear()
    setDirectories({})
    setExpanded(new Set())
    setLoading(Boolean(sessionId))
    if (sessionId) void loadDirectory('')
  }

  useEffect(() => {
    fileRequestRef.current?.abort()
    setSelected('')
    setContent(null)
    setWorkbook(null)
    setDocumentData(null)
    setSlidesData(null)
    setSaveStatus('')
    setLoadedModified('')
    snapshotProviderRef.current = null
    saveRequestVersionRef.current += 1
    setSaving(false)
    workbookDraft.reset()
    refreshTree()
    return () => {
      for (const controller of directoryRequestsRef.current.values()) controller.abort()
      directoryRequestsRef.current.clear()
      fileRequestRef.current?.abort()
    }
  }, [sessionId])

  const openFile = async (path: string) => {
    fileRequestRef.current?.abort()
    workbookDraft.reset()
    const controller = new AbortController()
    fileRequestRef.current = controller
    setSelected(path)
    setMarkdownSource(false)
    setContent(null)
    setWorkbook(null)
    setDocumentData(null)
    setSlidesData(null)
    setError('')
    setSaveStatus('')
    setLoadedModified('')
    snapshotProviderRef.current = null
    saveRequestVersionRef.current += 1
    setSaving(false)
    setFileLoading(true)
    try {
      const isWorkbook = workbookFormat(path) !== null
      const isDocument = /\.(doc|docx)$/i.test(path)
      const isSlides = /\.(ppt|pptx)$/i.test(path)
      const endpoint = isWorkbook ? 'workspace-xlsx' : isDocument ? 'workspace-doc' : isSlides ? 'workspace-slides' : 'workspace-file'
      const response = await fetch(`/api/${endpoint}?sessionId=${encodeURIComponent(sessionId)}&path=${encodeURIComponent(path)}`, { signal: controller.signal })
      const payload = await response.json()
      if (!response.ok) throw new Error(payload.error || '读取文件失败')
      if (controller.signal.aborted) return
      if (isWorkbook || isDocument || isSlides) setLoadedModified(payload.modified || '')
      // Normalize once, in place, on the ephemeral fetch payload before it
      // becomes React state. Re-serializing the whole unit (as the old
      // JSON.parse(JSON.stringify(...)) clone did) overflows the engine's
      // string allocation limit on large files.
      if (isWorkbook) {
        const restored = await workbookDraft.restore(sessionId, path, payload.modified || '')
        if (controller.signal.aborted) return
        if (restored === null) throw new Error('表格当前无法安全打开，请查看草稿状态提示')
        setWorkbook(normalizeWorkbookSnapshot(payload.workbook))
      } else if (isDocument) setDocumentData(normalizeDocumentSnapshot(payload.document))
      else if (isSlides) setSlidesData(payload.presentation)
      else setContent(payload.content)
    } catch (reason) {
      if (!controller.signal.aborted) setError(reason instanceof Error ? reason.message : String(reason))
    } finally {
      if (fileRequestRef.current === controller) {
        fileRequestRef.current = null
        setFileLoading(false)
      }
    }
  }

  const saveOfficeFile = async () => {
    const target = officeSaveTarget(selected)
    if (!target) {
      setSaveStatus('当前格式不支持保存；XLS 将另存为 XLSX，XLSX、CSV、DOCX 和 PPTX 可原位保存')
      return
    }
    const getSnapshot = snapshotProviderRef.current
    if (!getSnapshot) {
      setSaveStatus('编辑器仍在加载，请稍后再试')
      return
    }

    const path = selected
    const activeSessionId = sessionId
    const requestVersion = ++saveRequestVersionRef.current
    setSaving(true)
    setSaveStatus('')
    let sheetSavePrepared = false
    try {
      const preparedDraft = target.unitType === 'sheet' ? await workbookDraft.prepareSave() : null
      sheetSavePrepared = preparedDraft !== null
      const snapshot = await getSnapshot()
      const response = await fetch(`/api/workspace-office-save?sessionId=${encodeURIComponent(activeSessionId)}`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({
          path,
          unitType: target.unitType,
          format: target.format,
          expectedModified: loadedModified,
          data: snapshot,
        }),
      })
      const payload = await response.json()
      if (!response.ok) throw new Error(payload.error || '保存失败')
      const savedAsNewFile = payload.path !== path
      const runtimeKey = `${activeSessionId}\u0000${path}`
      if (target.unitType === 'sheet') {
        await workbookDraft.commitSaved(preparedDraft, payload.modified || '')
        if (!savedAsNewFile) advanceSheetFrameVersion(runtimeKey, loadedModified, payload.modified || '')
      } else {
        const kind: CachedOfficeKind = target.unitType === 'doc' ? 'docs' : 'slides'
        if (!savedAsNewFile) advanceOfficeProviderVersion(getSnapshot, loadedModified, payload.modified || '')
        officeResumeSnapshots[kind].delete(runtimeKey)
      }
      // This legacy surface can save a file also opened by DocumentPreview;
      // invalidate address-keyed conversion payloads across surfaces.
      officePayloadCache.clear()
      if (savedAsNewFile && target.unitType === 'sheet') {
        const oldRuntime = sheetFrameCache.get(runtimeKey)
        if (oldRuntime) disposeSheetFrame(oldRuntime)
        workbookDraft.reset()
      }
      if (saveRequestVersionRef.current === requestVersion) {
        setSaveStatus(`${savedAsNewFile ? '已另存为' : '已保存'} ${payload.path}`)
        setLoadedModified(payload.modified || '')
        if (savedAsNewFile) {
          const rebound = await workbookDraft.restore(activeSessionId, payload.path, payload.modified || '')
          if (rebound === null) throw new Error('文件已另存，但无法取得新文件的草稿编辑锁；请重新打开该文件')
          setSelected(payload.path)
          setWorkbook(normalizeWorkbookSnapshot(snapshot))
          void loadDirectory('')
        }
      }
    } catch (reason) {
      if (saveRequestVersionRef.current === requestVersion) setSaveStatus(reason instanceof Error ? reason.message : String(reason))
    } finally {
      if (sheetSavePrepared) workbookDraft.finishSave()
      if (saveRequestVersionRef.current === requestVersion) setSaving(false)
    }
  }

  const toggleDirectory = (path: string) => {
    const open = expanded.has(path)
    setExpanded((current) => {
      const next = new Set(current)
      if (open) next.delete(path)
      else next.add(path)
      return next
    })
    const state = directories[path]
    if (!open && !state?.loaded && !state?.loading) void loadDirectory(path)
  }

  const renderDirectoryEntries = (path: string, depth: number): React.ReactNode => {
    const state = directories[path]
    if (!state) return null
    return (
      <>
        {state.entries.map((entry) => {
          if (entry.type === 'directory') {
            const open = expanded.has(entry.path)
            const childState = directories[entry.path]
            return (
              <React.Fragment key={entry.path}>
                <button className="dsh-wfv-dir" style={{ paddingLeft: 8 + depth * 14 }} onClick={() => toggleDirectory(entry.path)}>
                  {open ? '▾' : '▸'} {entry.name}
                </button>
                {open && <>
                  {childState?.loading && !childState.loaded && <div className="dsh-wfv-tree-status" style={{ paddingLeft: 22 + (depth + 1) * 14 }}>加载中…</div>}
                  {childState?.error && <button className="dsh-wfv-tree-error" style={{ paddingLeft: 22 + (depth + 1) * 14 }} onClick={() => void loadDirectory(entry.path)}>加载失败，点击重试</button>}
                  {renderDirectoryEntries(entry.path, depth + 1)}
                </>}
              </React.Fragment>
            )
          }
          return (
            <button key={entry.path} className={`dsh-wfv-file${selected === entry.path ? ' active' : ''}`} style={{ paddingLeft: 22 + depth * 14 }} onClick={() => void openFile(entry.path)}>
              {entry.name}
            </button>
          )
        })}
        {state.loading && state.loaded && <div className="dsh-wfv-tree-status" style={{ paddingLeft: 22 + depth * 14 }}>加载更多…</div>}
        {!state.loading && state.nextCursor !== null && (
          <DirectoryLoadMore key={`${path}:${state.nextCursor}`} depth={depth} onLoad={() => void loadDirectory(path, state.nextCursor ?? '')} />
        )}
      </>
    )
  }

  return (
    <section className="dsh-wfv">
      <header className="dsh-wfv-head">
        <span className="dsh-wfv-title">Workspace 文件</span>
        <div className="dsh-wfv-actions">
          {selected && content !== null && <>
            {markdownFile && <button className="dsh-wfv-action" disabled={!markdownAllowed} aria-pressed={!markdownSource && markdownAllowed} onClick={() => setMarkdownSource((value) => !value)} title={markdownAllowed ? '切换 Markdown 预览与源码' : '文件过大，已使用源码模式'}>{markdownSource || !markdownAllowed ? '预览' : '源码'}</button>}
            <button className="dsh-wfv-action" onClick={() => navigator.clipboard?.writeText(content)}>复制</button>
            <button className="dsh-wfv-action" onClick={() => setCodeFontSize((size) => Math.max(12, size - 1))} disabled={codeFontSize <= 12} title="缩小代码字体">A−</button>
            <span className="dsh-wfv-font-size" title="当前代码字体大小">{codeFontSize}px</span>
            <button className="dsh-wfv-action" onClick={() => setCodeFontSize((size) => Math.min(20, size + 1))} disabled={codeFontSize >= 20} title="放大代码字体">A+</button>
          </>}
          {selected && (workbook !== null || documentData !== null || slidesData !== null) && (
            <button
              className="dsh-wfv-action dsh-wfv-save"
              onClick={() => void saveOfficeFile()}
              disabled={saving || officeSaveTarget(selected) === null}
              title={workbookFormat(selected) === 'xls' ? '官方 SDK 不支持导出 XLS；将保存为同目录同名 XLSX 文件' : officeSaveTarget(selected) ? '按原格式导出当前编辑内容并覆盖 workspace 中的原文件' : '旧版 DOC/PPT 格式不支持原位导出，请先转换为 DOCX/PPTX'}
            >
              {saving ? '保存中…' : '保存'}
            </button>
          )}
          <button className="dsh-wfv-action" onClick={refreshTree}>↻ 刷新</button>
        </div>
      </header>
      <div className="dsh-wfv-body">
        <nav className="dsh-wfv-tree">
          {loading && !directories['']?.loaded ? <div className="dsh-wfv-muted">加载中…</div> : renderDirectoryEntries('', 0)}
          {!loading && directories['']?.error && <button className="dsh-wfv-tree-error" onClick={() => void loadDirectory('')}>加载失败，点击重试</button>}
          {!loading && directories['']?.loaded && directories[''].entries.length === 0 && <div className="dsh-wfv-muted">目录为空</div>}
        </nav>
        <main className={`dsh-wfv-content${workbook !== null || documentData !== null || slidesData !== null ? ' dsh-wfv-content-univer' : ''}`}>
          {selected && <div className="dsh-wfv-path"><span>{selected}</span>{(saveStatus || workbookDraft.status) && <span className="dsh-wfv-save-status" role="status">{saveStatus || workbookDraft.status}</span>}</div>}
          {fileLoading && <div className="dsh-wfv-muted">正在通过 dsh-univer-file-export 转换并加载…</div>}
          {!fileLoading && error && <div className="dsh-wfv-error">{error}</div>}
          {!fileLoading && workbook !== null && (
            <UniverWorkbook
              key={selected}
              runtimeKey={`${sessionId}\u0000${selected}`}
              sourceModified={loadedModified}
              snapshot={workbook}
              snapshotProviderRef={snapshotProviderRef}
              draftMutations={workbookDraft.mutations}
              onDraftMutation={(mutation) => {
                setSaveStatus('')
                workbookDraft.record(mutation)
              }}
              onDraftRestored={workbookDraft.markRestored}
              onDraftRestoreFailed={workbookDraft.markRestoreFailed}
              editingDisabled={saving}
            />
          )}
          {!fileLoading && documentData !== null && <UniverDocument key={selected} runtimeKey={`${sessionId}\u0000${selected}`} sourceModified={loadedModified} snapshot={documentData} snapshotProviderRef={snapshotProviderRef} editingDisabled={saving} />}
          {!fileLoading && slidesData !== null && <UniverSlides key={selected} runtimeKey={`${sessionId}\u0000${selected}`} sourceModified={loadedModified} snapshot={slidesData} snapshotProviderRef={snapshotProviderRef} editingDisabled={saving} />}
          {!fileLoading && content !== null && <>
            {markdownFile && !markdownAllowed && <div className="dsh-wfv-muted" role="status">Markdown 文件较大，已切换为源码显示以保持流畅。</div>}
            {markdownFile && markdownAllowed && !markdownSource
              ? <MarkdownPreview key={selected} path={selected} content={content} fontSize={codeFontSize} onOpenFile={openFile} />
              : <CodePreview key={selected} path={selected} content={content} fontSize={codeFontSize} />}
          </>}
          {!fileLoading && !selected && !error && <div className="dsh-wfv-muted">选择一个文件进行预览；XLS、XLSX、CSV、DOC/DOCX 和 PPT/PPTX 将转换为 Univer UnitData。</div>}
        </main>
      </div>
    </section>
  )
}

type DocumentPreviewContent =
  | { readonly kind: 'text'; readonly text: string }
  | { readonly kind: 'bytes'; readonly data: Uint8Array<ArrayBuffer> }

type OfficePreviewProps = {
  readonly resourceAddress: string
  readonly content: DocumentPreviewContent
  readonly scrollportRef: React.RefCallback<HTMLElement>
}

type SessionOfficeFile = { sessionId: string; path: string }

const OFFICE_PREVIEW_ID = 'dsh-office-one/office'
const OFFICE_EXTENSIONS = ['xls', 'xlsx', 'csv', 'doc', 'docx', 'ppt', 'pptx'] as const

function sessionOfficeFile(address: string): SessionOfficeFile | null {
  const prefix = 'dsh-resource://file/session/'
  if (!address.startsWith(prefix)) return null
  try {
    const suffixIndex = address.search(/[?#]/)
    const encoded = address.slice(prefix.length, suffixIndex === -1 ? undefined : suffixIndex)
    const [sessionId, ...path] = encoded.split('/')
    if (!sessionId || path.length === 0) return null
    return {
      sessionId: decodeURIComponent(sessionId),
      path: path.map(decodeURIComponent).join('/'),
    }
  } catch {
    return null
  }
}

/** Office renderer selected by DSH's built-in workspace file list/document preview. */
function OfficePreview({ resourceAddress, content, scrollportRef }: OfficePreviewProps) {
  const file = useMemo(() => sessionOfficeFile(resourceAddress), [resourceAddress])
  const [workbook, setWorkbook] = useState<unknown>(null)
  const [documentData, setDocumentData] = useState<unknown>(null)
  const [slidesData, setSlidesData] = useState<unknown>(null)
  const [loadedModified, setLoadedModified] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [saving, setSaving] = useState(false)
  const [saveStatus, setSaveStatus] = useState('')
  const snapshotProviderRef = useRef<(() => unknown) | null>(null)
  const requestVersionRef = useRef(0)
  const workbookDraft = useWorkbookDraftManager()

  useEffect(() => {
    workbookDraft.reset()
    setWorkbook(null)
    setDocumentData(null)
    setSlidesData(null)
    setLoadedModified('')
    setError('')
    setSaveStatus('')
    setSaving(false)
    snapshotProviderRef.current = null
    const requestVersion = ++requestVersionRef.current
    if (content.kind !== 'bytes' || file === null) return

    const controller = new AbortController()
    const isWorkbook = workbookFormat(file.path) !== null
    const isDocument = /\.(doc|docx)$/i.test(file.path)
    const endpoint = isWorkbook ? 'workspace-xlsx' : isDocument ? 'workspace-doc' : 'workspace-slides'
    const previousPayload = officePayloadCache.get(resourceAddress)
    if (previousPayload && previousPayload.sourceRef?.deref() !== content.data) {
      const runtimeKey = `${file.sessionId}\u0000${file.path}`
      if (isWorkbook) {
        const staleRuntime = sheetFrameCache.get(runtimeKey)
        if (staleRuntime?.owner === null) disposeSheetFrame(staleRuntime)
      } else {
        const kind: CachedOfficeKind = isDocument ? 'docs' : 'slides'
        const staleRuntime = officeFrameCaches[kind].get(runtimeKey)
        if (staleRuntime?.owner === null) disposeOfficeFrame(staleRuntime, false)
        officeResumeSnapshots[kind].delete(runtimeKey)
      }
    }
    const payloadEntry = loadOfficePayload(resourceAddress, content.data, async () => {
      const response = await fetch(`/api/${endpoint}?sessionId=${encodeURIComponent(file.sessionId)}&path=${encodeURIComponent(file.path)}`)
      const payload = await response.json()
      if (!response.ok) throw new Error(payload.error || '读取 Office 文件失败')
      return payload as OfficePayload
    })
    setLoading(payloadEntry.value === undefined)
    void (async () => {
      try {
        const payload = payloadEntry.value ?? await payloadEntry.promise
        if (controller.signal.aborted || requestVersionRef.current !== requestVersion) return
        setLoadedModified(payload.modified || '')
        if (isWorkbook) {
          const restored = await workbookDraft.restore(file.sessionId, file.path, payload.modified || '')
          if (controller.signal.aborted || requestVersionRef.current !== requestVersion) return
          if (restored === null) throw new Error('表格当前无法安全打开，请查看草稿状态提示')
          setWorkbook(normalizeWorkbookSnapshot(payload.workbook))
        } else if (isDocument) setDocumentData(normalizeDocumentSnapshot(payload.document))
        else setSlidesData(payload.presentation)
      } catch (reason) {
        if (!controller.signal.aborted && requestVersionRef.current === requestVersion) {
          setError(reason instanceof Error ? reason.message : String(reason))
        }
      } finally {
        if (requestVersionRef.current === requestVersion) setLoading(false)
      }
    })()
    return () => controller.abort()
  }, [content, file])

  const saveOfficeFile = async () => {
    if (file === null) return
    const target = officeSaveTarget(file.path)
    if (!target) {
      setSaveStatus('当前格式不支持保存；XLS 将另存为 XLSX，DOC/PPT 请先转换为 DOCX/PPTX')
      return
    }
    const getSnapshot = snapshotProviderRef.current
    if (!getSnapshot) {
      setSaveStatus('编辑器仍在加载，请稍后再试')
      return
    }

    const requestVersion = ++requestVersionRef.current
    setSaving(true)
    setSaveStatus('')
    let sheetSavePrepared = false
    try {
      const preparedDraft = target.unitType === 'sheet' ? await workbookDraft.prepareSave() : null
      sheetSavePrepared = preparedDraft !== null
      const snapshot = await getSnapshot()
      const response = await fetch(`/api/workspace-office-save?sessionId=${encodeURIComponent(file.sessionId)}`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({
          path: file.path,
          unitType: target.unitType,
          format: target.format,
          expectedModified: loadedModified,
          data: snapshot,
        }),
      })
      const payload = await response.json()
      if (!response.ok) throw new Error(payload.error || '保存失败')
      const savedAsNewFile = payload.path !== file.path
      const runtimeKey = `${file.sessionId}\u0000${file.path}`
      if (target.unitType === 'sheet') {
        await workbookDraft.commitSaved(preparedDraft, payload.modified || '')
        if (!savedAsNewFile) advanceSheetFrameVersion(runtimeKey, loadedModified, payload.modified || '')
      } else {
        const kind: CachedOfficeKind = target.unitType === 'doc' ? 'docs' : 'slides'
        if (!savedAsNewFile) advanceOfficeProviderVersion(getSnapshot, loadedModified, payload.modified || '')
        officeResumeSnapshots[kind].delete(runtimeKey)
      }
      if (!savedAsNewFile) {
        updateOfficePayload(resourceAddress, {
          modified: payload.modified || '',
          ...(target.unitType === 'sheet' ? { workbook: snapshot } : target.unitType === 'doc' ? { document: snapshot } : { presentation: snapshot }),
        })
      } else if (target.unitType === 'sheet') {
        const oldRuntime = sheetFrameCache.get(runtimeKey)
        if (oldRuntime) disposeSheetFrame(oldRuntime)
        workbookDraft.reset()
      } else {
        const kind: CachedOfficeKind = target.unitType === 'doc' ? 'docs' : 'slides'
        const oldRuntime = officeFrameBySnapshotProvider.get(getSnapshot)
        if (oldRuntime) disposeOfficeFrame(oldRuntime, false)
      }
      if (requestVersionRef.current === requestVersion) {
        if (savedAsNewFile) {
          setWorkbook(null)
          setSaveStatus(`已另存为 ${payload.path}；请从 Workspace 打开新文件继续编辑`)
        } else {
          setSaveStatus(`已保存 ${payload.path}`)
          setLoadedModified(payload.modified || '')
        }
      }
    } catch (reason) {
      if (requestVersionRef.current === requestVersion) setSaveStatus(reason instanceof Error ? reason.message : String(reason))
    } finally {
      if (sheetSavePrepared) workbookDraft.finishSave()
      if (requestVersionRef.current === requestVersion) setSaving(false)
    }
  }

  if (file === null) {
    return <div className="dsh-wfv-office-status">Office 预览仅支持当前会话 Workspace 中的文件。</div>
  }
  const loaded = workbook !== null || documentData !== null || slidesData !== null
  const target = officeSaveTarget(file.path)
  return (
    <section ref={scrollportRef} className="dsh-wfv-office" data-office-preview={file.path}>
      <div className="dsh-wfv-office-toolbar">
        <span className="dsh-wfv-office-status" role="status">
          {loading ? '正在通过 dsh-univer-file-export 转换并加载…' : error || saveStatus || workbookDraft.status}
        </span>
        {loaded && (
          <button
            type="button"
            className="dsh-wfv-office-save"
            disabled={saving || target === null}
            title={workbookFormat(file.path) === 'xls' ? '将保存为同目录同名 XLSX 文件' : target ? '按原格式保存到 Workspace' : '旧版 DOC/PPT 格式不支持原位保存'}
            onClick={() => void saveOfficeFile()}
          >
            {saving ? '保存中…' : '保存'}
          </button>
        )}
      </div>
      <div className="dsh-wfv-office-editor">
        {!loading && !error && workbook !== null && (
          <UniverWorkbook
            runtimeKey={`${file.sessionId}\u0000${file.path}`}
            sourceModified={loadedModified}
            snapshot={workbook}
            snapshotProviderRef={snapshotProviderRef}
            draftMutations={workbookDraft.mutations}
            onDraftMutation={(mutation) => {
              setSaveStatus('')
              workbookDraft.record(mutation)
            }}
            onDraftRestored={workbookDraft.markRestored}
            onDraftRestoreFailed={workbookDraft.markRestoreFailed}
            editingDisabled={saving}
          />
        )}
        {!loading && !error && documentData !== null && <UniverDocument runtimeKey={`${file.sessionId}\u0000${file.path}`} sourceModified={loadedModified} snapshot={documentData} snapshotProviderRef={snapshotProviderRef} editingDisabled={saving} />}
        {!loading && !error && slidesData !== null && <UniverSlides runtimeKey={`${file.sessionId}\u0000${file.path}`} sourceModified={loadedModified} snapshot={slidesData} snapshotProviderRef={snapshotProviderRef} editingDisabled={saving} />}
      </div>
    </section>
  )
}

export const inject = ['slots', 'documentPreviews']

export function apply(ctx: any): void {
  ctx.effect(() => ctx.documentPreviews.register({
    id: OFFICE_PREVIEW_ID,
    extensions: OFFICE_EXTENSIONS,
    priority: 'extension',
    title: () => 'Office (Univer)',
    loading: 'bytes-complete',
    wrap: false,
  }), 'dsh-office-one: Office preview metadata')
  ctx.effect(() => ctx.slots.inject('sidebar.right.tab.document', () => ctx.slots.register({
    name: 'sidebar.right.tab.document',
    key: OFFICE_PREVIEW_ID,
  }, OfficePreview)), 'dsh-office-one: Office preview body')
}
