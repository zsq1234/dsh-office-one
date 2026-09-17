import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'

const client = await readFile(new URL('../src/modules/univer-create/client/index.tsx', import.meta.url), 'utf8')
const host = await readFile(new URL('../src/modules/univer-create/host.ts', import.meta.url), 'utf8')

test('workspace save conflicts cross RPC as a typed error', () => {
  assert.match(host, /class WorkspaceFileExistsError extends Error/)
  assert.match(host, /code = 'already-exists'/)
  assert.match(host, /error instanceof WorkspaceFileExistsError \? error\.code : 'export-failed'/)
  assert.match(client, /throw new UniverExportError\(result\.error\.code, result\.error\.message\)/)
})

test('all Univer editors ask before overwriting an existing workspace file', () => {
  assert.match(client, /reason\.code === 'already-exists'/)
  assert.equal(client.match(/window\.confirm\(ui\.confirmOverwrite\(targetPath\)\)/g)?.length, 3)
})
