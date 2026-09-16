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
        if (method === 'file-path') return options.filePath ?? { ok: true, value: 'saved.office' }
        throw new Error(`Unexpected RPC: ${method}`)
      },
    },
  }
  const clientConnection = { current: connection }
  // Exercise the production metadata loader as well, including RPC errors.
  const helperStart = source.indexOf('async function loadUnitFilePath(')
  const helperBodyStart = source.indexOf('{', helperStart)
  const helperEnd = source.indexOf('\n}', helperBodyStart)
  const loadUnitFilePath = new Function('clientConnection', `return async function(sessionId, unitType) ${source.slice(helperBodyStart, helperEnd + 2)}`)(clientConnection)
  const env = {
    sessionId: 'session',
    runtimeRef: { current: null },
    appliedRef: { current: new Set([123]) },
    lastSavedRef: { current: 'old' },
    hydratedFromHostRef: { current: true },
    [markerName]: markers,
    clientConnection,
    props: { onRuntimePresenceChange: () => {} },
    loadUnitFilePath,
    setHostLoaded: (value) => { state.loaded = value },
    setHostSnapshot: (value) => { state.snapshot = value },
    setFilePath: (value) => { state.filePath = value },
    setSaveDialogOpen: () => {},
    setSaveNotice: () => {},
    [visibleSetter]: (value) => { state.visible = value },
    setError: (value) => { state.error = value },
  }
  const cleanup = new Function(...Object.keys(env), effect)(...Object.values(env))
  return { state, calls, markers, cleanup, unitType }
}

for (const product of products) {
  const [unitType] = product
  test(`${unitType}: fresh page loads persisted snapshot without an opened marker`, async () => {
    const { state, calls, markers } = mountRestore(product)
    assert.equal(state.loaded, false, 'replay blocked until restore finishes')
    assert.equal(markers.size, 0)
    await settle()
    assert.deepEqual(state.snapshot, { id: 'persisted', body: 'manual edit' })
    assert.equal(state.filePath, 'saved.office')
    assert.equal(state.loaded, true)
    assert.equal(state.error, null)
    assert.deepEqual(calls.map(({ method }) => method), ['load', 'file-path'])
    for (const call of calls) assert.deepEqual(call.args, { sessionId: 'session', unitType })
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
    ['file-path RPC error', { filePath: { ok: false, error: { message: 'metadata failed' } } }],
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
