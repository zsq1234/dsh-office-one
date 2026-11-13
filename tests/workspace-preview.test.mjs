import assert from 'node:assert/strict'
import { mkdtemp, writeFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import test from 'node:test'
import { apply, inject } from '../src/modules/workspace-file-viewer/host/index.js'

function fakeFs(version = 'dev:1:16:100:200') {
  return {
    resolve: async (path) => path,
    contains: () => true,
    stat: async () => ({ type: 'file', version, size: 16 }),
  }
}

function request(persistence, path, fs = fakeFs()) {
  const routes = new Map()
  apply({
    sessionPersistence: persistence,
    fs,
    webServer: { register: ({ path, handler }) => { routes.set(path, handler); return () => {} } },
    effect: (setup) => setup(),
  })
  let status
  let payload
  const url = new URL(path, 'http://localhost')
  return routes.get(url.pathname)({ url: path, method: 'GET' }, {
    writeHead: (value) => { status = value },
    end: (body) => { payload = JSON.parse(body) },
  }).then(() => ({ status, payload }))
}

test('workspace XLSX route resolves the session through public persistence and imports the file', async (t) => {
  assert.ok(inject.includes('sessionPersistence'))
  assert.ok(inject.includes('fs'))
  assert.ok(!inject.includes('workspaceRegistry'))
  const root = await mkdtemp(join(tmpdir(), 'office-preview-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  await writeFile(join(root, 'aws-money.xlsx'), 'test workbook')
  const workbook = { id: 'test-book', sheets: {} }
  const fetchMock = t.mock.method(globalThis, 'fetch', async (url, options) => {
    assert.equal(new URL(url).searchParams.get('fileName'), 'aws-money.xlsx')
    assert.equal(options.body.toString(), 'test workbook')
    return Response.json({ data: workbook })
  })
  const result = await request({ stat: async (id) => {
    assert.equal(id, 'session-a')
    return { header: { cwd: root } }
  } }, '/api/workspace-xlsx?sessionId=session-a&path=aws-money.xlsx')
  assert.equal(result.status, 200)
  assert.deepEqual(result.payload.workbook, workbook)
  assert.equal(result.payload.version, 'dev:1:16:100:200')
  assert.equal(fetchMock.mock.callCount(), 1)
})

test('office metadata exposes the same freshness token DSH uses for auto-refresh', async (t) => {
  const root = await mkdtemp(join(tmpdir(), 'office-meta-'))
  t.after(() => rm(root, { recursive: true, force: true }))
  await writeFile(join(root, 'aws-money.xlsx'), 'test workbook')
  const result = await request({ stat: async () => ({ header: { cwd: root } }) }, '/api/workspace-office-meta?sessionId=a&path=aws-money.xlsx')
  assert.equal(result.status, 200)
  assert.match(result.payload.modified, /^\d{4}-\d{2}-\d{2}T/)
  assert.equal(result.payload.version, 'dev:1:16:100:200')
  assert.equal(result.payload.size, 13)
})

test('preview rejects missing session ids and missing or relative directories', async () => {
  for (const snapshot of [undefined, { header: {} }, { header: { cwd: 'relative' } }]) {
    const result = await request({ stat: async () => snapshot }, '/api/workspace-xlsx?sessionId=a&path=a.xlsx')
    assert.equal(result.status, 400)
    assert.match(result.payload.error, /cannot resolve/)
  }
  const result = await request({ stat: async () => { throw new Error('must not be called') } }, '/api/workspace-xlsx?path=a.xlsx')
  assert.equal(result.status, 400)
  assert.equal(result.payload.error, 'sessionId is required')
})
