import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'

// Execute the actual initialization effects without booting a browser/Univer.
// Only the Slide snapshot type assertion is removed; no restore logic is copied.
const source = await readFile(new URL('../src/modules/univer-create/client/index.tsx', import.meta.url), 'utf8')
const products = [
  ['sheet', 'SheetProductView', 'openedWorkbookSessions', 'setSheetVisible'],
  ['doc', 'DocProductView', 'openedDocumentSessions', 'setDocumentVisible'],
  ['slide', 'SlideProductView', 'openedPresentationSessions', 'setPresentationVisible'],
]
const settle = () => new Promise((resolve) => setImmediate(resolve))

function mountRestore(product, options = {}) {
  const [unitType, component, markerName, visibleSetter] = product
  const componentStart = source.indexOf(`function ${component}(`)
  const start = source.indexOf('  useEffect(() => {\n    let cancelled = false\n    appliedRef.current.clear()', componentStart)
  assert.ok(start > componentStart, `${component} initialization effect exists`)
  const end = source.indexOf('  }, [sessionId])', start)
  assert.ok(end > start)
  const effect = source.slice(start + '  useEffect(() => {'.length, end).replaceAll(' as ISlideData', '')
  const markers = new Set(options.opened ? ['session'] : [])
  const state = { loaded: false, snapshot: null, visible: options.opened ?? false, error: null }
  const calls = []
  const connection = options.disconnected ? null : {
    rpc: {
      async call(namespace, method, args) {
        calls.push({ namespace, method, args })
        if (method === 'load') {
          if (options.reject) throw new Error('offline')
          return options.load ?? { ok: true, value: { id: 'persisted', body: 'manual edit' } }
        }
        if (method === 'file-state') return options.fileState ?? { ok: true, value: { path: 'saved.office', snapshot: { id: 'persisted', body: 'saved content' } } }
        if (method === 'unit-state') {
          if (options.reject) throw new Error('offline')
          const loadResult = options.load === undefined ? { ok: true, value: { id: 'persisted', body: 'manual edit' } } : await options.load
          if (loadResult?.ok === false) return loadResult
          if (options.fileState?.ok === false) return options.fileState
          return options.unitState ?? { ok: true, value: { snapshot: loadResult?.value, path: options.fileState ? options.fileState.value?.path ?? null : 'saved.office', fileSnapshot: options.fileState ? options.fileState.value?.snapshot ?? null : { id: 'persisted', body: 'saved content' }, activeFileId: 'default', files: [{ id: 'default', title: 'Persisted', filePath: 'saved.office', updatedAt: 1 }], revision: 1 } }
        }
        if (method === 'sheet-files' || method === 'unit-files') return options.sheetFiles ?? { ok: true, value: { activeFileId: 'default', files: [{ id: 'default', title: 'Persisted', filePath: 'saved.office', updatedAt: 1 }] } }
        throw new Error(`Unexpected RPC: ${method}`)
      },
    },
  }
  const clientConnection = { current: connection }
  // Exercise the production metadata loader as well, including RPC errors.
  const helperStart = source.indexOf('async function loadUnitFileState(')
  const helperBodyStart = source.indexOf('{', helperStart)
  const helperEnd = source.indexOf('\n}', helperBodyStart)
  const helperBody = source.slice(helperBodyStart, helperEnd + 2)
    .replace(' as { path?: unknown; snapshot?: unknown }', '')
  const loadUnitFileState = new Function('clientConnection', `return async function(sessionId, unitType) ${helperBody}`)(clientConnection)
  const loadUnitState = async (sessionId, unitType) => {
    const result = await clientConnection.current.rpc.call('/dsh-univer-create', 'unit-state', { sessionId, unitType })
    if (!result.ok) throw new Error(result.error.message)
    const value = result.value
    return { unitSnapshot: value.snapshot, path: value.path ?? null, snapshot: value.fileSnapshot ?? null, activeFileId: value.activeFileId ?? null, files: value.files, revision: value.revision }
  }
  const loadUnitFiles = async (sessionId, unitType) => {
    const result = await clientConnection.current.rpc.call('/dsh-univer-create', 'unit-files', { sessionId, unitType })
    if (!result.ok) throw new Error(result.error.message)
    return result.value
  }
  const loadSheetFiles = async (sessionId) => {
    const result = await clientConnection.current.rpc.call('/dsh-univer-create', 'sheet-files', { sessionId })
    if (!result.ok) throw new Error(result.error.message)
    return result.value
  }
  const lastFileSnapshotRef = { current: 'old-file' }
  const env = {
    sessionId: 'session',
    runtimeRef: { current: null },
    residentRuntimes: new Map(),
    residentRuntimeKey: (sessionId, unitType) => `${sessionId}:${unitType}`,
    disposeResidentRuntime: () => {},
    appliedRef: { current: new Set([123]) },
    lastSavedRef: { current: 'old' },
    lastFileSnapshotRef,
    activeFileIdRef: { current: null },
    nextRuntimeTitleRef: { current: null },
    hydratedFromHostRef: { current: true },
    [markerName]: markers,
    clientConnection,
    props: { onRuntimePresenceChange: () => {} },
    loadUnitFileState,
    loadUnitState,
    loadUnitFiles,
    loadSheetFiles,
    UNKNOWN_FILE_SNAPSHOT: '\u0000unknown-file-snapshot',
    setHostLoaded: (value) => { state.loaded = value },
    setHostSnapshot: (value) => { state.snapshot = value },
    setFilePath: (value) => { state.filePath = value },
    setActiveFileId: (value) => { state.activeFileId = value },
    setSheetFiles: (value) => { state.sheetFiles = value },
    setUnitFiles: (value) => { state.unitFiles = value },
    setSaveDialogOpen: () => {},
    setSaveNotice: () => {},
    [visibleSetter]: (value) => { state.visible = value },
    setError: (value) => { state.error = value },
  }
  const cleanup = new Function(...Object.keys(env), effect)(...Object.values(env))
  return { state, calls, markers, cleanup, unitType, lastFileSnapshotRef }
}

for (const product of products) {
  const [unitType] = product
  test(`${unitType}: fresh page loads persisted snapshot without an opened marker`, async () => {
    const { state, calls, markers, lastFileSnapshotRef } = mountRestore(product)
    assert.equal(state.loaded, false, 'replay blocked until restore finishes')
    assert.equal(markers.size, 0)
    await settle()
    assert.deepEqual(state.snapshot, { id: 'persisted', body: 'manual edit' })
    assert.equal(state.filePath, 'saved.office')
    assert.equal(lastFileSnapshotRef.current, JSON.stringify({ id: 'persisted', body: 'saved content' }))
    assert.equal(state.loaded, true)
    assert.equal(state.error, null)
    assert.deepEqual(calls.map(({ method }) => method), ['unit-state'])
    assert.deepEqual(calls[0].args, { sessionId: 'session', unitType })
  })

  test(`${unitType}: legacy restored data without a file baseline remains dirty`, async () => {
    const { lastFileSnapshotRef } = mountRestore(product, {
      fileState: { ok: true, value: { path: null, snapshot: null } },
    })
    await settle()
    assert.equal(lastFileSnapshotRef.current, '\u0000unknown-file-snapshot')
  })

  test(`${unitType}: successful null permits welcome and clears stale marker`, async () => {
    const { state, markers } = mountRestore(product, { opened: true, load: { ok: true, value: null } })
    await settle()
    assert.equal(state.loaded, true)
    assert.equal(state.snapshot, null)
    assert.equal(state.visible, false)
    assert.equal(markers.size, 0)
    assert.equal(state.error, null)
  })

  for (const [name, options] of [
    ['load RPC error', { load: { ok: false, error: { message: 'storage failed' } } }],
    ['rejected RPC', { reject: true }],
    ['missing connection', { disconnected: true }],
    ['file-state RPC error', { fileState: { ok: false, error: { message: 'metadata failed' } } }],
    ['undefined snapshot', { load: { ok: true, value: undefined } }],
    ['array snapshot', { load: { ok: true, value: [] } }],
  ]) {
    test(`${unitType}: ${name} fails closed without forgetting opened state`, async () => {
      const { state, markers, calls } = mountRestore(product, { ...options, opened: true })
      await settle()
      assert.equal(state.loaded, false, 'blank replay remains blocked')
      assert.equal(state.snapshot, null)
      assert.equal(markers.has('session'), true)
      assert.equal(typeof state.error, 'string')
      assert.ok(state.error.length > 0)
      assert.ok(calls.every(({ method }) => method !== 'save'))
    })
  }

  for (const rejected of [false, true]) {
    test(`${unitType}: unmounted restore ignores late ${rejected ? 'failure' : 'success'}`, async () => {
      let resolveLoad
      let rejectLoad
      const load = new Promise((resolve, reject) => { resolveLoad = resolve; rejectLoad = reject })
      const { state, cleanup, markers } = mountRestore(product, { load, opened: true })
      cleanup()
      const before = { ...state }
      if (rejected) rejectLoad(new Error('late disconnect'))
      else resolveLoad({ ok: true, value: null })
      await settle()
      assert.deepEqual(state, before)
      assert.equal(markers.has('session'), true)
    })
  }
}
