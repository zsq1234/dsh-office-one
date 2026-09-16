import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'

const client = await readFile(new URL('../src/modules/univer-create/client/index.tsx', import.meta.url), 'utf8')
const host = await readFile(new URL('../src/modules/univer-create/host.ts', import.meta.url), 'utf8')
const workspaceDocsRuntime = await readFile(new URL('../src/modules/workspace-file-viewer/client/runtimes/docs.ts', import.meta.url), 'utf8')
const workspaceSheetRuntime = await readFile(new URL('../src/modules/workspace-file-viewer/client/runtimes/sheet.ts', import.meta.url), 'utf8')
const manifest = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'))

test('host exposes versioned Facade reference and permission-aware browser code tools', () => {
  assert.equal(manifest.dependencies['@univer-cli/api-reference'], '1.0.0-rc.0')
  assert.match(host, /name: 'univer_api_reference'/)
  assert.match(host, /createStandardApiReference\(\)/)
  assert.match(host, /name: 'univer_execute_code'/)
  assert.match(host, /exec\.name !== 'univer_execute_code'/)
  assert.match(host, /approval\?\.overrideOf\(exec\.agent\.session\)/)
  assert.match(host, /if \(policy === 'never'\) return next\(\)/)
  assert.match(host, /kind: 'ask'/)
})

test('browser executor injects only Facade root and controlled console', () => {
  assert.match(client, /new AsyncFunction\('univerAPI', 'console'/)
  assert.match(client, /executeFacadeCode\(request\.code, runtime\.univerAPI\)/)
  assert.match(client, /FORBIDDEN_UNIVER_CODE/)
  assert.match(client, /__DSH_UNIVER_CODE_DEBUG__/)
  assert.match(client, /debugUniverCode\(sessionId, unitType, request\.id, request\.code\)/)
  assert.match(client, /'tasks'/)
  assert.doesNotMatch(client, /'univer-code-requests'/)
  assert.match(client, /'univer-code-claim'/)
  assert.match(client, /'univer-code-result'/)
})

test('Doc runtimes register the enhanced table plugin and Facade mixin', () => {
  assert.equal(manifest.dependencies['@univerjs-pro/docs-table'], '1.0.0-rc.0')
  for (const source of [client, workspaceDocsRuntime]) {
    assert.match(source, /UniverDocsTablePlugin/)
    assert.match(source, /@univerjs-pro\/docs-table\/facade/)
  }
})

test('Sheet runtimes register Chart plugins, Facade mixins, and UI resources', () => {
  for (const dependency of ['@univerjs-pro/engine-chart', '@univerjs-pro/chart-ui', '@univerjs-pro/sheets-chart', '@univerjs-pro/sheets-chart-ui']) {
    assert.equal(manifest.dependencies[dependency], '1.0.0-rc.0')
  }
  for (const source of [client, workspaceSheetRuntime]) {
    assert.match(source, /UniverSheetsChartPlugin/)
    assert.match(source, /UniverSheetsChartUIPlugin/)
    assert.match(source, /@univerjs-pro\/engine-chart\/facade/)
    assert.match(source, /@univerjs-pro\/chart-ui\/facade/)
    assert.match(source, /@univerjs-pro\/sheets-chart\/facade/)
    assert.match(source, /ChartUIZhCN/)
    assert.match(source, /SheetsChartUIZhCN/)
  }
})

test('successful browser code execution persists each supported unit snapshot', () => {
  assert.match(client, /getActiveWorkbook\(\)\?\.save\(\)/)
  assert.match(client, /getActiveDocument\(\)\?\.save\(\)/)
  assert.match(client, /presentation\.save\(\)/)
  assert.match(client, /'save', \{ sessionId, unitType, snapshot \}/)
  assert.match(client, /useUniverCodeExecutor\(sessionId, 'sheet', props\.codeRequests, hostLoaded/)
  assert.match(client, /useUniverCodeExecutor\(sessionId, 'doc', props\.codeRequests, hostLoaded/)
  assert.match(client, /useUniverCodeExecutor\(sessionId, 'slide', props\.codeRequests, hostLoaded/)
})
