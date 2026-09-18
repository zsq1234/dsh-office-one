import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'

const client = await readFile(new URL('../src/modules/univer-create/client/index.tsx', import.meta.url), 'utf8')
const host = await readFile(new URL('../src/modules/univer-create/host.ts', import.meta.url), 'utf8')
const locales = await readFile(new URL('../src/modules/univer-create/client/locales.ts', import.meta.url), 'utf8')

test('Host atomically closes only the active file and activates a remaining file', () => {
  const start = host.indexOf("if (endpoint === 'close-unit-file')")
  const end = host.indexOf("if (endpoint === 'operations' || endpoint === 'tasks')", start)
  assert.ok(start >= 0 && end > start)
  const block = host.slice(start, end)
  assert.match(block, /withSessionQueueLock\(sessionId/)
  assert.match(block, /state\.activeFileId !== sourceFileId/)
  assert.match(block, /delete files\[sourceFileId\]/)
  assert.match(block, /right\.updatedAt - left\.updatedAt/)
  assert.match(block, /activeSheetFileId: activeFileId/)
  assert.match(block, /activeDocFileId: activeFileId/)
  assert.match(block, /activeSlideFileId: activeFileId/)
  assert.match(block, /snapshot: selected\?\.snapshot \?\? null/)
})

test('Sheet, Doc, and Slide expose close actions with unsaved confirmation', () => {
  assert.match(client, /function confirmCloseFile\(/)
  assert.match(client, /filePath !== null && JSON\.stringify\(snapshot\) === baseline/)
  assert.match(client, /const closeCurrentWorkbook = async/)
  assert.match(client, /const closeCurrentDocument = async/)
  assert.match(client, /const closeCurrentPresentation = async/)
  assert.equal(client.match(/await closeUnitFile\(sessionId, '(?:sheet|doc|slide)', fileId\)/g)?.length, 3)
  assert.equal(client.match(/\{ui\.close\}/g)?.length, 3)
  assert.match(locales, /confirmCloseUnsaved: 'The current file has unsaved changes/)
  assert.match(locales, /confirmCloseUnsaved: '当前文件有尚未保存的修改/)
})

test('closing the last file returns each product to its welcome state', () => {
  assert.match(client, /openedWorkbookSessions\.delete\(sessionId\)[\s\S]*?setSheetVisible\(false\)/)
  assert.match(client, /openedDocumentSessions\.delete\(sessionId\)[\s\S]*?setDocumentVisible\(false\)/)
  assert.match(client, /openedPresentationSessions\.delete\(sessionId\)[\s\S]*?setPresentationVisible\(false\)/)
})
