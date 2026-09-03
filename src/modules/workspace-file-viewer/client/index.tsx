import React, { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
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
type UniverSheetRuntime = { univer: { dispose: () => void }; workbook: SaveableUnit }
type UniverDocumentRuntime = { univer: { dispose: () => void }; univerAPI: { createDocument: (snapshot: unknown) => SaveableUnit; getActiveDocument?: () => SaveableUnit | null } }
type UniverSlidesRuntime = { dispose: () => void; createUnit: (type: unknown, snapshot: unknown) => SaveableUnit }
type SnapshotProviderRef = React.MutableRefObject<(() => unknown) | null>

type RuntimeFactory = (container: HTMLElement) => UniverDocumentRuntime | UniverSlidesRuntime

type RuntimeGlobal = {
  createSheetRuntime?: (container: HTMLElement, snapshot: unknown) => UniverSheetRuntime
  createDocsRuntime?: RuntimeFactory
  createSlidesRuntime?: RuntimeFactory
  createSlide?: (runtime: UniverSlidesRuntime, snapshot: unknown) => unknown
}

const runtimePromises: Record<string, Promise<RuntimeGlobal>> = {}
let runtimeLifecycle: Promise<void> = Promise.resolve()

function afterRuntimeLifecycle<T>(task: () => T | Promise<T>): Promise<T> {
  return runtimeLifecycle.then(task)
}

function queueRuntimeDispose(dispose: () => void): void {
  runtimeLifecycle = runtimeLifecycle.then(() => new Promise<void>((resolve) => {
    window.setTimeout(() => {
      try { dispose() } finally { resolve() }
    }, 0)
  }))
}

function loadRuntime(kind: 'sheet' | 'docs' | 'slides'): Promise<RuntimeGlobal> {
  const existing = runtimePromises[kind]
  if (existing) return existing
  const url = `/api/workspace-file-viewer/runtime/${kind}.js`
  const globalKey = kind === 'sheet' ? 'createSheetRuntime' : kind === 'docs' ? 'createDocsRuntime' : 'createSlidesRuntime'
  runtimePromises[kind] = new Promise((resolve, reject) => {
    const script = document.createElement('script')
    script.src = url
    script.async = true
    script.onload = () => {
      const runtime = (window as Window & { __DSH_WORKSPACE_FILE_VIEWER__?: RuntimeGlobal }).__DSH_WORKSPACE_FILE_VIEWER__
      if (!runtime?.[globalKey]) reject(new Error(`runtime ${kind} did not expose ${globalKey}`))
      else resolve(runtime)
    }
    script.onerror = () => reject(new Error(`failed to load ${url}`))
    document.head.appendChild(script)
  })
  return runtimePromises[kind]
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

function UniverWorkbook({ snapshot, snapshotProviderRef }: { snapshot: unknown; snapshotProviderRef: SnapshotProviderRef }) {
  const containerRef = useRef<HTMLDivElement>(null)
  const runtimeRef = useRef<UniverSheetRuntime | null>(null)
  const [layoutVersion, setLayoutVersion] = useState(0)

  useLayoutEffect(() => {
    const container = containerRef.current
    if (container === null || typeof ResizeObserver === 'undefined') return
    const observer = new ResizeObserver(() => setLayoutVersion((value) => value + 1))
    observer.observe(container)
    return () => observer.disconnect()
  }, [])

  useLayoutEffect(() => {
    const container = containerRef.current
    if (container === null || runtimeRef.current !== null) return
    const bounds = container.getBoundingClientRect()
    if (bounds.width <= 0 || bounds.height <= 0) return

    const mount = document.createElement('div')
    mount.className = 'dsh-wfv-univer-runtime'
    container.appendChild(mount)
    let cancelled = false
    void afterRuntimeLifecycle(async () => {
      const { createSheetRuntime } = await loadRuntime('sheet')
      if (cancelled) return
      if (!createSheetRuntime) throw new Error('sheet runtime factory is unavailable')
      const runtime = createSheetRuntime(mount, snapshot)
      if (cancelled) queueRuntimeDispose(() => runtime.univer.dispose())
      else {
        runtimeRef.current = runtime
        snapshotProviderRef.current = () => unitSnapshot(runtime.workbook)
      }
    }).catch((reason) => { if (!cancelled) console.error('Failed to load Sheet preview', reason) })
    return () => {
      cancelled = true
      snapshotProviderRef.current = null
      const runtime = runtimeRef.current
      if (runtime) {
        runtimeRef.current = null
        queueRuntimeDispose(() => runtime.univer.dispose())
      }
      mount.remove()
    }
  }, [snapshot, layoutVersion, snapshotProviderRef])

  return <div ref={containerRef} className="dsh-wfv-univer" aria-label="电子表格 Univer 预览" />
}

function UniverDocument({ snapshot, snapshotProviderRef }: { snapshot: unknown; snapshotProviderRef: SnapshotProviderRef }) {
  const containerRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const container = containerRef.current
    if (container === null) return
    let runtime: UniverDocumentRuntime | null = null
    let cancelled = false
    void afterRuntimeLifecycle(async () => {
      const { createDocsRuntime } = await loadRuntime('docs')
      if (cancelled) return
      if (!createDocsRuntime) throw new Error('docs runtime factory is unavailable')
      runtime = createDocsRuntime(container) as UniverDocumentRuntime
      const documentUnit = runtime.univerAPI.createDocument(snapshot as any)
      if (cancelled) queueRuntimeDispose(() => runtime?.univer.dispose())
      else snapshotProviderRef.current = () => unitSnapshot(runtime?.univerAPI.getActiveDocument?.() ?? documentUnit)
    }).catch((reason) => { if (!cancelled) console.error('Failed to load Docs preview', reason) })
    return () => {
      cancelled = true
      snapshotProviderRef.current = null
      const previous = runtime
      runtime = null
      // Queue teardown before the next file's runtime is created. Univer's
      // nested React root must be detached first, but the global lifecycle
      // container must also be free before another Unit runtime starts.
      if (previous) queueRuntimeDispose(() => previous.univer.dispose())
    }
  }, [snapshot, snapshotProviderRef])

  return <div ref={containerRef} className="dsh-wfv-univer dsh-wfv-univer-doc" aria-label="DOC Univer 预览" />
}

function UniverSlides({ snapshot, snapshotProviderRef }: { snapshot: unknown; snapshotProviderRef: SnapshotProviderRef }) {
  const containerRef = useRef<HTMLDivElement>(null)
  const runtimeRef = useRef<UniverSlidesRuntime | null>(null)

  useLayoutEffect(() => {
    const container = containerRef.current
    if (container === null || typeof ResizeObserver === 'undefined') return
    let frame = 0
    const notifyResize = () => {
      cancelAnimationFrame(frame)
      frame = requestAnimationFrame(() => window.dispatchEvent(new Event('resize')))
    }
    const observer = new ResizeObserver(notifyResize)
    observer.observe(container)
    notifyResize()
    return () => {
      observer.disconnect()
      cancelAnimationFrame(frame)
    }
  }, [])

  useEffect(() => {
    const container = containerRef.current
    if (container === null) return
    let cancelled = false
    const resizeTimers: number[] = []
    const notifyResize = (delay: number) => {
      resizeTimers.push(window.setTimeout(() => {
        if (!cancelled) window.dispatchEvent(new Event('resize'))
      }, delay))
    }
    void afterRuntimeLifecycle(async () => {
      const runtimeGlobal = await loadRuntime('slides')
      if (cancelled) return
      if (!runtimeGlobal.createSlidesRuntime || !runtimeGlobal.createSlide) throw new Error('slides runtime factory is unavailable')
      const runtime = runtimeGlobal.createSlidesRuntime(container) as UniverSlidesRuntime
      const presentation = runtimeGlobal.createSlide(runtime, snapshot) as SaveableUnit
      if (cancelled) queueRuntimeDispose(() => runtime.dispose())
      else {
        runtimeRef.current = runtime
        snapshotProviderRef.current = () => unitSnapshot(presentation)
        // Slides UI creates the main scene and thumbnail scenes asynchronously.
        // Re-measure after both the React tree and image resources have mounted.
        notifyResize(0)
        notifyResize(100)
        notifyResize(300)
        notifyResize(800)
      }
    }).catch((reason) => { if (!cancelled) console.error('Failed to load Slides preview', reason) })
    return () => {
      cancelled = true
      snapshotProviderRef.current = null
      resizeTimers.forEach((timer) => window.clearTimeout(timer))
      const runtime = runtimeRef.current
      runtimeRef.current = null
      if (runtime) queueRuntimeDispose(() => runtime.dispose())
    }
  }, [snapshot, snapshotProviderRef])

  return <div ref={containerRef} className="dsh-wfv-univer dsh-wfv-univer-slides" aria-label="PPT Univer 预览" />
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
  const [error, setError] = useState('')
  const [saving, setSaving] = useState(false)
  const [saveStatus, setSaveStatus] = useState('')
  const [loadedModified, setLoadedModified] = useState('')
  const directoryRequestsRef = useRef<Map<string, AbortController>>(new Map())
  const fileRequestRef = useRef<AbortController | null>(null)
  const snapshotProviderRef = useRef<(() => unknown) | null>(null)
  const saveRequestVersionRef = useRef(0)

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
    refreshTree()
    return () => {
      for (const controller of directoryRequestsRef.current.values()) controller.abort()
      directoryRequestsRef.current.clear()
      fileRequestRef.current?.abort()
    }
  }, [sessionId])

  const openFile = async (path: string) => {
    fileRequestRef.current?.abort()
    const controller = new AbortController()
    fileRequestRef.current = controller
    setSelected(path)
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
      if (isWorkbook) setWorkbook(normalizeWorkbookSnapshot(payload.workbook))
      else if (isDocument) setDocumentData(normalizeDocumentSnapshot(payload.document))
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
    try {
      const snapshot = getSnapshot()
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
      if (saveRequestVersionRef.current === requestVersion) {
        const savedAsNewFile = payload.path !== path
        setSaveStatus(`${savedAsNewFile ? '已另存为' : '已保存'} ${payload.path}`)
        setLoadedModified(payload.modified || '')
        if (savedAsNewFile) {
          setSelected(payload.path)
          setWorkbook(normalizeWorkbookSnapshot(snapshot))
          void loadDirectory('')
        }
      }
    } catch (reason) {
      if (saveRequestVersionRef.current === requestVersion) setSaveStatus(reason instanceof Error ? reason.message : String(reason))
    } finally {
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
          {selected && <div className="dsh-wfv-path"><span>{selected}</span>{saveStatus && <span className="dsh-wfv-save-status" role="status">{saveStatus}</span>}</div>}
          {fileLoading && <div className="dsh-wfv-muted">正在通过 dsh-univer-file-export 转换并加载…</div>}
          {!fileLoading && error && <div className="dsh-wfv-error">{error}</div>}
          {!fileLoading && workbook !== null && <UniverWorkbook key={selected} snapshot={workbook} snapshotProviderRef={snapshotProviderRef} />}
          {!fileLoading && documentData !== null && <UniverDocument key={selected} snapshot={documentData} snapshotProviderRef={snapshotProviderRef} />}
          {!fileLoading && slidesData !== null && <UniverSlides key={selected} snapshot={slidesData} snapshotProviderRef={snapshotProviderRef} />}
          {!fileLoading && content !== null && <CodePreview path={selected} content={content} fontSize={codeFontSize} />}
          {!fileLoading && !selected && !error && <div className="dsh-wfv-muted">选择一个文件进行预览；XLS、XLSX、CSV、DOC/DOCX 和 PPT/PPTX 将转换为 Univer UnitData。</div>}
        </main>
      </div>
    </section>
  )
}

export const inject = ['slots']

export function apply(ctx: any): void {
  ctx.slots.inject('conversation.view', () => ctx.slots.register({
    name: 'conversation.view',
    id: 'workspace-file-viewer',
    order: 100,
    label: 'Workspace 文件',
  }, FileBrowser))
}
