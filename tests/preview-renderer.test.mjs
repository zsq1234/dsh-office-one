import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'
import ts from 'typescript'

const client = await readFile(new URL('../src/modules/workspace-file-viewer/client/index.tsx', import.meta.url), 'utf8')
const host = await readFile(new URL('../src/modules/workspace-file-viewer/host/index.js', import.meta.url), 'utf8')

const start = client.indexOf('type OfficeRefreshDecision')
const end = client.indexOf('async function fetchOfficeFileMeta')
assert.ok(start >= 0 && end > start)
const helper = client.slice(start, end)
const js = ts.transpileModule(helper, { compilerOptions: { target: ts.ScriptTarget.ES2022 } }).outputText
const decide = new Function(`${js}; return decideOfficeRefresh`)()

test('a save whose on-disk version is already on screen keeps the mounted editor', () => {
  assert.equal(decide({
    alreadyMounted: true,
    loadedVersion: 'v2',
    loadedModified: 'm2',
    meta: { version: 'v2', modified: 'm2' },
    cachedVersion: 'v2',
  }), 'keep')
  // Older Host replies may not carry a version; the mtime still identifies a self-save.
  assert.equal(decide({
    alreadyMounted: true,
    loadedVersion: '',
    loadedModified: 'm2',
    meta: { version: 'v2', modified: 'm2' },
  }), 'keep')
})

test('an external change reloads instead of being mistaken for a self-save', () => {
  assert.equal(decide({
    alreadyMounted: true,
    loadedVersion: 'v1',
    loadedModified: 'm1',
    meta: { version: 'v2', modified: 'm2' },
    cachedVersion: 'v1',
  }), 'reload')
  assert.equal(decide({ alreadyMounted: true, loadedVersion: 'v1', loadedModified: 'm1', meta: null }), 'reload')
})

test('a remount reuses a conversion for the same version and otherwise reloads', () => {
  assert.equal(decide({
    alreadyMounted: false,
    loadedVersion: '',
    loadedModified: '',
    meta: { version: 'v2', modified: 'm2' },
    cachedVersion: 'v2',
  }), 'reuse')
  assert.equal(decide({
    alreadyMounted: false,
    loadedVersion: '',
    loadedModified: '',
    meta: { version: 'v3', modified: 'm3' },
    cachedVersion: 'v2',
  }), 'reload')
})

test('the Office preview owns content loading so DSH never unmounts it during a save refresh', () => {
  assert.match(client, /loading: 'renderer'/)
  assert.doesNotMatch(client, /loading: 'bytes-complete'/)
  assert.match(client, /\}, \[file, rendererRevision, revisionTick\]\)/)
  assert.match(client, /expectedModified: loadedModifiedRef\.current/)
  assert.match(client, /syncRendererVersion\(typeof payload\.version === 'string' \? payload\.version : ''\)/)
  assert.match(client, /savingRef\.current > 0/)
})

test('the Host exposes the version DSH watches for saves and metadata', () => {
  assert.match(host, /\/api\/workspace-office-meta/)
  assert.match(host, /version: await fsVersion\(fs, root, target\)/)
  assert.match(host, /version: await fsVersion\(fs, root, absolute\)/)
})
