import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'

const client = await readFile(new URL('../src/modules/univer-create/client/index.tsx', import.meta.url), 'utf8')
const host = await readFile(new URL('../src/modules/univer-create/host.ts', import.meta.url), 'utf8')
const locales = await readFile(new URL('../src/modules/univer-create/client/locales.ts', import.meta.url), 'utf8')

test('opening another file confirms when the current Univer unit is dirty', () => {
  assert.match(client, /function confirmOpenWithUnsavedChanges\(/)
  assert.match(client, /const openWorkbookFile = \(\) =>/)
  assert.match(client, /const openDocumentFile = \(\) =>/)
  assert.match(client, /const openPresentationFile = \(\) =>/)
  assert.equal(client.match(/ui\.confirmDiscardChanges,/g)?.length, 3)
  assert.match(client, /onClick=\{openWorkbookFile\}/)
  assert.match(client, /onClick=\{openDocumentFile\}/)
  assert.match(client, /onClick=\{openPresentationFile\}/)
})

test('creating a new Univer unit confirms when the current unit is dirty', () => {
  assert.match(client, /const newWorkbook = async \(\) => \{[\s\S]*?ui\.confirmDiscardChangesForNew,/)
  assert.match(client, /const newDocument = async \(\) => \{[\s\S]*?ui\.confirmDiscardChangesForNew,/)
  assert.match(client, /const newPresentation = async \(\) => \{[\s\S]*?ui\.confirmDiscardChangesForNew,/)
  assert.equal(client.match(/ui\.confirmDiscardChangesForNew,/g)?.length, 3)
})

test('the file baseline survives a page refresh', () => {
  assert.match(host, /fileSnapshot: z\.unknown\(\)\.optional\(\)/)
  assert.match(host, /endpoint === 'file-state'/)
  assert.match(host, /fileSnapshot: request\.snapshot/)
  assert.match(client, /restoredFileState\.snapshot === null \? UNKNOWN_FILE_SNAPSHOT/)
})

test('successful file saves refresh the dirty-check baseline', () => {
  assert.equal(client.match(/lastFileSnapshotRef\.current = JSON\.stringify\(nextSnapshot\)/g)?.length, 3)
  assert.match(client, /lastFileSnapshotRef: React\.MutableRefObject<string \| null>/)
})

test('unsaved-change confirmation is localized', () => {
  assert.equal(locales.match(/The current file has unsaved changes/g)?.length, 2)
  assert.equal(locales.match(/当前文件有尚未保存的修改/g)?.length, 2)
})
