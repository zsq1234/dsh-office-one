import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'

const client = await readFile(new URL('../src/modules/univer-create/client/index.tsx', import.meta.url), 'utf8')
const styles = await readFile(new URL('../src/modules/univer-create/client/styles.css', import.meta.url), 'utf8')
const containment = await readFile(new URL('../src/modules/univer-create/client/scroll-containment.js', import.meta.url), 'utf8')
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

test('slide presentation navigation cannot scroll the surrounding DSH page', () => {
  const hostRule = styles.match(/\.dsh-univer-create-container--slides\s*\{([^}]*)\}/)?.[1] ?? ''
  assert.match(hostRule, /overflow:\s*hidden/)
  assert.match(hostRule, /overscroll-behavior:\s*contain/)

  // Univer's presentation dock calls scrollIntoView({ inline: 'center' }) on the
  // active thumbnail every page turn; the tab must contain that scroll.
  const slideStart = client.indexOf('function SlideProductView(')
  const slideEnd = client.indexOf('const selectedUnitBySession', slideStart)
  const slideView = client.slice(slideStart, slideEnd)
  assert.match(slideView, /installScrollContainment\(container\)/)
  assert.match(client, /from '\.\/scroll-containment\.js'/)
})

test('scroll containment confines Univer scrolling to the mount point', () => {
  // The shipped module must intercept both side effects Univer performs while
  // presenting, and must stop walking at the boundary instead of scrolling it.
  assert.match(containment, /Element\.prototype\.scrollIntoView = function/)
  assert.match(containment, /HTMLElement\.prototype\.focus = function/)
  assert.match(containment, /node !== null && node !== boundary/)
  assert.match(containment, /preventScroll: true/)
  assert.match(containment, /\.univer-fixed\.univer-inset-0/)
  assert.match(containment, /container\.contains\(this\)/)
})

test('document screenshots use the rendered Univer canvas and attachment output', () => {
  assert.match(client, /function captureRenderedDocument/)
  assert.match(client, /'doc-screenshot-requests'/)
  assert.match(client, /'doc-screenshot-result'/)
  assert.match(host, /name: 'univer_doc_screenshot'/)
  assert.match(host, /action: 'doc-screenshot'/)
})

test('host exposes visual screenshot tool and richer layout metadata', () => {
  assert.match(host, /name: 'univer_slide_screenshot'/)
  assert.match(host, /type: 'image' as const/)
  assert.match(host, /attachments\.saveImage/)
  for (const field of ['left', 'top', 'width', 'height', 'fontSize', 'fontColor']) {
    assert.ok(host.includes(`${field}: { type: 'number'`) || host.includes(`${field}: { type: 'string'`), `${field} metadata is exposed`)
  }
})
