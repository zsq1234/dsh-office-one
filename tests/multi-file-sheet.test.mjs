import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'

const client = await readFile(new URL('../src/modules/univer-create/client/index.tsx', import.meta.url), 'utf8')
const host = await readFile(new URL('../src/modules/univer-create/host.ts', import.meta.url), 'utf8')
const styles = await readFile(new URL('../src/modules/univer-create/client/styles.css', import.meta.url), 'utf8')
const locales = await readFile(new URL('../src/modules/univer-create/client/locales.ts', import.meta.url), 'utf8')

test('Sheet storage keeps a backward-compatible active workbook and archived snapshots', () => {
  assert.match(host, /activeSheetFileId: z\.string\(\)\.optional\(\)/)
  assert.match(host, /sheetFiles: z\.record\(z\.string\(\), sheetFileSchema\)\.optional\(\)/)
  assert.match(host, /const LEGACY_SHEET_FILE_ID = 'default'/)
  assert.match(host, /function activeSheetFile\(/)
  assert.match(host, /await table\.put\(storageKey, \{ \.\.\.current, \.\.\.update\(current\) \}\)/)
})

test('switch RPC atomically archives the outgoing snapshot and restores the target', () => {
  const start = host.indexOf("if (endpoint === 'switch-sheet-file' || endpoint === 'switch-unit-file')")
  const end = host.indexOf("if (endpoint === 'operations'", start)
  assert.ok(start >= 0 && end > start)
  const block = host.slice(start, end)
  assert.match(block, /withSessionQueueLock\(sessionId/)
  assert.match(block, /typeof sourceFileId !== 'string'/)
  assert.match(block, /currentSourceFileId !== sourceFileId/)
  assert.match(block, /const isTargetSelection = queuedAction === 'select-file'/)
  assert.match(block, /const expectedCreateAction = targetUnit === 'sheet' \? 'new' : targetUnit === 'doc' \? 'new-doc' : 'new-slide'/)
  assert.match(block, /const isSiblingCreation = targetSnapshot !== undefined/)
  assert.match(block, /snapshot: currentSnapshot/)
  assert.match(block, /selected = files\[fileId\]/)
  assert.match(block, /snapshot: selected\.snapshot/)
  assert.match(block, /activeSheetFileId: fileId/)
  assert.match(block, /activeDocFileId: fileId/)
  assert.match(block, /activeSlideFileId: fileId/)
  assert.match(block, /queuedOperations: pending\.filter/)
})

test('Sheet uses one resident runtime and serializes saves before switching', () => {
  assert.match(client, /function residentRuntimeKey\(sessionId: string, unitType: UniverUnitType\)/)
  assert.doesNotMatch(client, /function residentRuntimeKey\([^)]*fileId/)
  const start = client.indexOf('const activateSheetFile = async')
  const end = client.indexOf('const saveAsXlsx', start)
  assert.ok(start >= 0 && end > start)
  const block = client.slice(start, end)
  assert.match(block, /const previousSave = pendingSaveRef\.current/)
  assert.match(block, /await previousSave/)
  assert.match(block, /normalizeWorkbookSnapshot\(currentWorkbook\.save\(\)\)/)
  assert.match(block, /disposeResidentRuntime\(sessionId, 'sheet'/)
  assert.match(block, /setHostSnapshot\(switched\.workbookSnapshot\)/)
})

test('AI can list and queue selection of an opened Sheet workbook', () => {
  assert.match(host, /name: 'univer_sheet_file_list'/)
  assert.match(host, /name: 'univer_sheet_file_select'/)
  assert.match(host, /enqueueOperation\('sheet', args, exec, 'select-file'\)/)
  assert.match(client, /operation\.action === 'select-file'/)
  assert.match(client, /operationIds: \[String\(operation\.seq\)\]/)
})

test('stale saves and exports are bound to the active Sheet file', () => {
  assert.match(host, /const assertActiveSheetFile = /)
  assert.ok((host.match(/assertActiveSheetFile\(previous\)/g) ?? []).length >= 2)
  assert.match(client, /fileId: activeFileIdRef\.current/)
  assert.match(client, /const savingFileId = activeFileIdRef\.current/)
  assert.match(client, /pendingSaveRef\.current = save/)
})

test('Doc and Slide expose independent durable file registries and AI tools', () => {
  for (const [unit, title] of [['doc', 'Doc'], ['slide', 'Slide']]) {
    assert.match(host, new RegExp(`active${title}FileId: z\\.string\\(\\)\\.optional\\(\\)`))
    assert.match(host, new RegExp(`${unit}Files: z\\.record`))
    assert.match(host, new RegExp(`name: 'univer_${unit}_file_list'`))
    assert.match(host, new RegExp(`name: 'univer_${unit}_file_select'`))
    assert.match(client, new RegExp(`switchUnitFile\\(sessionId, '${unit}'`))
  }
  assert.match(client, /const activateDocumentFile = async/)
  assert.match(client, /const activatePresentationFile = async/)
  assert.match(host, /endpoint === 'unit-state'/)
  assert.match(client, /residentEntry\?\.fileId !== restored\.activeFileId/)
  assert.ok((client.match(/const saveGate = new Promise<void>/g) ?? []).length >= 2)
})

test('all Univer products expose a localized file selector and input blocker', () => {
  assert.match(client, /className="dsh-univer-create-file-selector"/)
  assert.match(client, /aria-label=\{ui\.workbookFiles\}/)
  assert.match(client, /sheetFiles\.map\(\(file\) => <option key=\{file\.id\}/)
  assert.match(styles, /\.dsh-univer-create-file-selector/)
  assert.match(client, /className="dsh-univer-create-switch-blocker"/)
  assert.match(styles, /\.dsh-univer-create-switch-blocker/)
  assert.match(locales, /workbookFiles: 'Open files'/)
  assert.match(locales, /workbookFiles: '已打开的文件'/)
})
