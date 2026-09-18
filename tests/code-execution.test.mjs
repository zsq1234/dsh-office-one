import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'

const client = await readFile(new URL('../src/modules/univer-create/client/index.tsx', import.meta.url), 'utf8')
const host = await readFile(new URL('../src/modules/univer-create/host.ts', import.meta.url), 'utf8')
const workspaceDocsRuntime = await readFile(new URL('../src/modules/workspace-file-viewer/client/runtimes/docs.ts', import.meta.url), 'utf8')
const workspaceSheetRuntime = await readFile(new URL('../src/modules/workspace-file-viewer/client/runtimes/sheet.ts', import.meta.url), 'utf8')
const workspaceSlidesRuntime = await readFile(new URL('../src/modules/workspace-file-viewer/client/runtimes/slides.ts', import.meta.url), 'utf8')
const styles = await readFile(new URL('../src/modules/univer-create/client/styles.css', import.meta.url), 'utf8')
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
  assert.match(client, /executeFacadeCode\(request\.code, entry\.runtime\.univerAPI\)/)
  assert.match(client, /FORBIDDEN_UNIVER_CODE/)
  assert.match(client, /__DSH_UNIVER_CODE_DEBUG__/)
  assert.match(client, /debugUniverCode\(sessionId, request\.unitType, request\.id, request\.code\)/)
  assert.match(client, /'tasks'/)
  assert.doesNotMatch(client, /'univer-code-requests'/)
  assert.match(client, /'univer-code-claim'/)
  assert.match(client, /'univer-code-result'/)
})

test('resident tasks portal survives conversation tab hiding and follows the active session', () => {
  assert.match(client, /currentResidentSessionId = props\.sessionId/)
  assert.match(client, /createPortal\(<UniverView \{\.\.\.props\} \/>, shell\)/)
  assert.match(client, /getResidentViewShell\(props\.sessionId\)/)
  assert.match(client, /host\.appendChild\(shell\)/)
  assert.match(client, /getResidentRuntimeParkingHost\(\)\.appendChild\(shell\)/)
  assert.match(client, /conversation\.session\.header\.utilities/)
  assert.match(client, /\}, UniverViewHost\)\)/)
  assert.match(client, /props\.useSession\(\(session\) => session\.running\)/)
  assert.match(client, /if \(!sessionRunning && !drainingTasks\)/)
  assert.match(client, /emptyDrainPollsRef\.current >= 2/)
  assert.match(client, /remoteEmpty && !localBusy/)
  assert.match(client, /residentRuntimeTaskQueues/)
  assert.match(client, /enqueueResidentRuntimeTask\(key/)
  assert.match(client, /queueResidentFinalSave\(entry/)
  assert.match(client, /!residentRuntimeTaskQueues\.has\(key\)/)
  assert.match(styles, /\.dsh-univer-create-runtime-parking/)
  assert.match(styles, /\.dsh-univer-create-resident-shell/)
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

test('Slide runtimes register Chart model, UI, locales, and Facade mixins', () => {
  assert.equal(manifest.dependencies['@univerjs-pro/slides-chart'], '1.0.0-rc.0')
  assert.equal(manifest.dependencies['@univerjs-pro/slides-chart-ui'], '1.0.0-rc.0')
  for (const source of [client, workspaceSlidesRuntime]) {
    assert.match(source, /UniverSlidesChartPlugin/)
    assert.match(source, /UniverSlidesChartUIPlugin/)
    assert.match(source, /SlidesChartUIZhCN/)
    assert.match(source, /ChartUIZhCN/)
    assert.match(source, /@univerjs-pro\/engine-chart\/facade/)
    assert.match(source, /@univerjs-pro\/slides-chart\/facade/)
  }
})

test('successful resident code execution persists each supported unit snapshot', () => {
  assert.match(client, /normalizeWorkbookSnapshot\(mounted\.univerAPI\.getActiveWorkbook\(\)\?\.save\(\)\)/)
  assert.match(client, /mounted\.univerAPI\.getActiveDocument\(\)\?\.save\(\)/)
  assert.match(client, /\(\) => presentation\.save\(\)/)
  assert.match(client, /unitType: request\.unitType,[\s\S]*?snapshot,/)
  assert.match(client, /entry\.pendingSaveRef\.current = queued/)
  assert.match(client, /fileId: targetFileId/)
  assert.match(client, /entry\.activeFileIdRef\?\.current !== targetFileId/)
  for (const unitType of ['sheet', 'doc', 'slide']) {
    assert.match(client, new RegExp(`registerResidentRuntime\\(\\s*sessionId,\\s*'${unitType}'`))
  }
})
