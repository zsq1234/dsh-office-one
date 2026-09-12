import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'

const client = await readFile(new URL('../src/modules/univer-create/client/index.tsx', import.meta.url), 'utf8')
const host = await readFile(new URL('../src/modules/univer-create/host.ts', import.meta.url), 'utf8')

test('AI-created slides use the blank layout and suppress inherited placeholders', () => {
  assert.match(client, /layoutPageId = 'layout-blank'/)
  assert.match(client, /showMasterSp = false/)
  assert.match(client, /layoutPageId: 'layout-blank'/)
  assert.match(client, /id\.startsWith\('slide-text-'\)/)
})

test('slide screenshots use rendered canvases while keeping the React chat overlay', () => {
  assert.match(client, /querySelectorAll\('canvas'\)/)
  assert.match(client, /toDataURL\('image\/png'\)/)
  assert.match(client, /'slide-screenshot-requests'/)
  assert.match(client, /'slide-screenshot-result'/)

  const slideStart = client.indexOf('function SlideProductView(')
  const slideEnd = client.indexOf('const selectedUnitBySession', slideStart)
  const slideView = client.slice(slideStart, slideEnd)
  assert.match(slideView, /dsh-univer-create-chat-overlay/)
})

test('host exposes visual screenshot tool and richer layout metadata', () => {
  assert.match(host, /name: 'univer_slide_screenshot'/)
  assert.match(host, /type: 'image' as const/)
  assert.match(host, /attachments\.saveImage/)
  for (const field of ['left', 'top', 'width', 'height', 'fontSize', 'fontColor']) {
    assert.ok(host.includes(`${field}: { type: 'number'`) || host.includes(`${field}: { type: 'string'`), `${field} metadata is exposed`)
  }
})
