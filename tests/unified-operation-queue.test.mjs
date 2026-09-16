import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'

const client = await readFile(new URL('../src/modules/univer-create/client/index.tsx', import.meta.url), 'utf8')
const host = await readFile(new URL('../src/modules/univer-create/host.ts', import.meta.url), 'utf8')

test('Host registers only the 12 core Univer tools', () => {
  const names = [...host.matchAll(/ctx\.tools\.register\(defineTool\(\{\s*name: '([^']+)'/g)].map((match) => match[1])
  assert.deepEqual(names, [
    'univer_api_reference',
    'univer_execute_code',
    'univer_sheet_new',
    'univer_sheet_list',
    'univer_sheet_get_range',
    'univer_sheet_screenshot',
    'univer_doc_new',
    'univer_doc_get_text',
    'univer_doc_screenshot',
    'univer_slide_new',
    'univer_slide_list',
    'univer_slide_screenshot',
  ])
})

test('Host persists one target-aware operation queue for all Univer products', () => {
  assert.match(host, /const queuedOperationSchema = z\.object\(/)
  assert.match(host, /unitType: z\.enum\(\['sheet', 'doc', 'slide'\]\)/)
  assert.match(host, /queuedOperations: z\.array\(queuedOperationSchema\)\.optional\(\)/)
  assert.match(host, /endpoint === 'operations'/)
  assert.match(host, /endpoint === 'tasks'/)
  assert.match(host, /endpoint === 'ack-operations'/)
  assert.match(host, /endpoint === 'commit-operations'/)
  assert.match(host, /claimedBy: z\.string\(\)\.optional\(\)/)
  assert.match(host, /revision: z\.number\(\)\.int\(\)\.optional\(\)/)
  assert.match(host, /MAX_PENDING_OPERATIONS/)
  assert.doesNotMatch(host, /\.slice\(-2_000\)/)
  for (const unitType of ['sheet', 'doc', 'slide']) {
    assert.ok(host.includes(`enqueueOperation('${unitType}'`), `${unitType} writes enter the unified queue`)
  }
})

test('resident Univer shell uses one aggregated task poll and atomic operation commits', () => {
  assert.match(client, /function UniverView\(props: ConvViewProps\)/)
  assert.match(client, /createPortal\(<UniverView/)
  assert.match(client, /'tasks'/)
  assert.doesNotMatch(client, /'univer-code-requests'/)
  assert.match(client, /tasks\.operations\.slice\(0, 1\)/)
  assert.match(client, /'commit-operations'/)
  assert.match(client, /clientId: univerCodeClientId/)
  for (const legacyPoll of ['sheet-operations', 'operations', 'univer-code-target', 'doc-screenshot-requests', 'slide-screenshot-requests', 'slide-runtime-heartbeat']) {
    assert.ok(!client.includes(`'${legacyPoll}'`), `${legacyPoll} is not polled by the client`)
  }
  assert.doesNotMatch(client, /acknowledgeOperations/)
})

test('Slide screenshots can queue without the Slide tab being active', () => {
  const calls = host.match(/requireActiveSlide\(sessionId\)/g) ?? []
  assert.equal(calls.length, 0)
  const screenshotStart = host.indexOf("name: 'univer_slide_screenshot'")
  assert.ok(screenshotStart >= 0)
  assert.doesNotMatch(host.slice(screenshotStart), /requireActiveSlide\(sessionId\)/)
})

test('tool results distinguish queued, applied, and rendered states', () => {
  assert.match(host, /status: 'queued' as const/)
  assert.match(host, /status: 'applied' as const/)
  assert.match(host, /status: 'rendered' as const/)
})

test('applied reads and rendered screenshots wait for pending writes', () => {
  assert.ok((host.match(/waitForOperationQueueEmpty\(sessionId, 'sheet'/g) ?? []).length >= 2)
  assert.ok((host.match(/waitForOperationQueueEmpty\(sessionId, 'doc'/g) ?? []).length >= 2)
  assert.ok((host.match(/waitForOperationQueueEmpty\(sessionId, 'slide'/g) ?? []).length >= 2)
})

test('manual replacement never pre-acknowledges pending AI operations', () => {
  assert.doesNotMatch(client, /new Set\(operations\.map/)
})

test('queue commits participate in the per-product save mutex', () => {
  assert.equal((client.match(/savingRef\.current = true\n        const previousSave = pendingSaveRef\.current/g) ?? []).length, 3)
  assert.equal((client.match(/if \(pendingSaveRef\.current === save\)/g) ?? []).length, 3)
})
