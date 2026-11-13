import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { isAbsolute } from 'node:path'
import test from 'node:test'

const host = await readFile(new URL('../src/modules/univer-create/host.ts', import.meta.url), 'utf8')
const wrapper = await readFile(new URL('../src/index.ts', import.meta.url), 'utf8')
const resolution = host.slice(host.indexOf('          const root = (await services.sessionPersistence.stat(sessionId))'), host.indexOf('          saved = await saveExportedFile(root,'))
assert.ok(resolution.length > 0)
const resolveRoot = new Function('services', 'sessionId', 'isAbsolute', `return (async () => { ${resolution}; return root })()`)

test('export uses the public session persistence service, not workspace internals', async () => {
  assert.doesNotMatch(host, /workspaceRegistry\.host|sessionPath\(/)
  assert.match(host, /export const inject = .*'sessionPersistence'/)
  assert.match(wrapper, /'sessionPersistence'/)
  let requested
  const root = await resolveRoot({ sessionPersistence: { stat: async (id) => {
    requested = id
    return { header: { cwd: '/workspace/session-a' } }
  } } }, 'session-a', isAbsolute)
  assert.equal(requested, 'session-a')
  assert.equal(root, '/workspace/session-a')
})

test('missing or invalid session directories fail closed without a process cwd fallback', async () => {
  for (const value of [undefined, { header: {} }, { header: { cwd: '' } }, { header: { cwd: 'relative' } }]) {
    await assert.rejects(resolveRoot({ sessionPersistence: { stat: async () => value } }, 'a', isAbsolute), /无法解析/)
  }
})

test('storage errors propagate and directory lookup stays inside export error handling', async () => {
  await assert.rejects(resolveRoot({ sessionPersistence: { stat: async () => { throw new Error('storage unavailable') } } }, 'a', isAbsolute), /storage unavailable/)
  assert.match(host, /try \{\s+const root = \(await services\.sessionPersistence\.stat/)
})
