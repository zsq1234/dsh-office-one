import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'
import ts from 'typescript'

const client = await readFile(new URL('../src/modules/workspace-file-viewer/client/index.tsx', import.meta.url), 'utf8')
const hook = client.slice(client.indexOf('function useFrameSourceBoundary('), client.indexOf('\nfunction UniverWorkbook('))
const js = ts.transpileModule(hook, { compilerOptions: { target: ts.ScriptTarget.ES2022 } }).outputText

function harness() {
  const ref = { current: null }
  let previousDeps
  let memo
  const useMemo = (factory, deps) => {
    if (!previousDeps || deps.some((value, i) => value !== previousDeps[i])) {
      memo = factory()
      previousDeps = deps
    }
    return memo
  }
  return new Function('useRef', 'useMemo', `${js}; return useFrameSourceBoundary`)(() => ref, useMemo)
}

test('in-place save preserves frame mount identity, including repeated saves', () => {
  const render = harness()
  const first = render('session/file', 'v1')
  const entry = { sourceModified: 'v1', disposed: false }
  first.mounted.current = { runtimeKey: 'session/file', boundary: first.boundary, entry }
  for (const version of ['v2', 'v3', 'v4']) {
    entry.sourceModified = version // advanced by save success before the state update
    assert.equal(render('session/file', version).boundary, first.boundary)
  }
})

test('external updates, file switches, and disposed frames invalidate mount identity', () => {
  for (const mode of ['external', 'switch', 'disposed']) {
    const render = harness()
    const first = render('session/file', 'v1')
    first.mounted.current = {
      runtimeKey: 'session/file', boundary: first.boundary,
      entry: { sourceModified: mode === 'disposed' ? 'v2' : 'v1', disposed: mode === 'disposed' },
    }
    assert.notEqual(render(mode === 'switch' ? 'session/other' : 'session/file', 'v2').boundary, first.boundary)
  }
})

test('both sheet and document/slide mount effects depend on the source boundary instead of save mtime', () => {
  assert.match(client, /\[runtimeKey, boundary, snapshot, layoutVersion, snapshotProviderRef, language\]/)
  assert.match(client, /\[kind, runtimeKey, boundary, snapshot, layoutVersion, snapshotProviderRef, language\]/)
  assert.equal(client.match(/mounted\.current = \{ runtimeKey, boundary, entry \}/g)?.length, 2)
})
