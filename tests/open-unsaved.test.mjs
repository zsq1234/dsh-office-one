import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'

const client = await readFile(new URL('../src/modules/univer-create/client/index.tsx', import.meta.url), 'utf8')
const host = await readFile(new URL('../src/modules/univer-create/host.ts', import.meta.url), 'utf8')
const locales = await readFile(new URL('../src/modules/univer-create/client/locales.ts', import.meta.url), 'utf8')

test('Sheet preserves the current draft when opening or creating another workbook', () => {
  assert.match(client, /const activateSheetFile = async/)
  assert.match(client, /const currentSnapshot = normalizeWorkbookSnapshot\(currentWorkbook\.save\(\)\)/)
  assert.match(client, /await switchSheetFile\(sessionId, targetFileId, currentSnapshot/)
  assert.match(client, /await activateSheetFile\(crypto\.randomUUID\(\), \{ snapshot: importedSnapshot/)
  assert.match(client, /await activateSheetFile\(crypto\.randomUUID\(\), \{ snapshot: blankSnapshot/)
  assert.match(client, /onClick=\{openWorkbookFile\}/)
})

test('Doc and Slide also preserve drafts through sibling-file switching', () => {
  assert.match(client, /const activateDocumentFile = async/)
  assert.match(client, /const activatePresentationFile = async/)
  assert.match(client, /switchUnitFile\(sessionId, 'doc'/)
  assert.match(client, /switchUnitFile\(sessionId, 'slide'/)
  assert.match(client, /onClick=\{openDocumentFile\}/)
  assert.match(client, /onClick=\{openPresentationFile\}/)
})

test('AI new operations create sibling files instead of replacing active runtimes', () => {
  assert.match(client, /operation\.action === 'new'[\s\S]*?switchSheetFile\(sessionId, crypto\.randomUUID\(\)/)
  assert.match(client, /operation\.action === 'new-doc'[\s\S]*?switchUnitFile\(sessionId, 'doc', crypto\.randomUUID\(\)/)
  assert.match(client, /operation\.action === 'new-slide'[\s\S]*?switchUnitFile\(sessionId, 'slide', crypto\.randomUUID\(\)/)
  assert.match(client, /await commitOperations\(sessionId, 'doc', finalSnapshot, \[String\(operation\.seq\)\]\)/)
})

test('the file baseline survives a page refresh', () => {
  assert.match(host, /fileSnapshot: z\.unknown\(\)\.optional\(\)/)
  assert.match(host, /endpoint === 'file-state'/)
  assert.match(host, /fileSnapshot: request\.snapshot/)
  assert.match(client, /restored\.snapshot === null \? UNKNOWN_FILE_SNAPSHOT/)
})

test('successful file saves refresh the dirty-check baseline', () => {
  assert.equal(client.match(/lastFileSnapshotRef\.current = JSON\.stringify\(nextSnapshot\)/g)?.length, 3)
  assert.match(client, /lastFileSnapshotRef: React\.MutableRefObject<string \| null>/)
})

test('unsaved-change confirmation is localized', () => {
  assert.equal(locales.match(/The current file has unsaved changes/g)?.length, 3)
  assert.equal(locales.match(/当前文件有尚未保存的修改/g)?.length, 3)
})
