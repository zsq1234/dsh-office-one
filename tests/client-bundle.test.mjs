import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'

test('client bundle requires only DSH-provided React modules', async () => {
  const code = await readFile(new URL('../lib/client.js', import.meta.url), 'utf8')
  const imports = [...code.matchAll(/(?<![\w$.])require\(\s*["']([^"']+)["']\s*\)/g)].map((match) => match[1])
  const allowed = new Set(['react', 'react/jsx-runtime', 'react/jsx-dev-runtime', 'react-dom', 'react-dom/client'])
  assert.deepEqual([...new Set(imports)].filter((id) => !allowed.has(id)), [])
  for (const name of ['minproc', 'minpath', 'minurl']) {
    assert.ok(code.includes(`${name}.browser.js`), `vfile must use its ${name} browser adapter`)
  }
})

test('workspace office runtimes contain no unresolved CommonJS imports', async () => {
  for (const runtime of ['sheet', 'docs', 'slides']) {
    const code = await readFile(new URL(`../lib/runtimes/${runtime}.js`, import.meta.url), 'utf8')
    const imports = [...code.matchAll(/(?:^|[^\w$.])require\(\s*["']([^"']+)["']\s*\)/gm)].map((match) => match[1])
    assert.deepEqual([...new Set(imports)], [], `${runtime} runtime must be directly executable in a browser`)
  }
})

test('client extends the built-in document preview without its own Workspace tab', async () => {
  const code = await readFile(new URL('../lib/client.js', import.meta.url), 'utf8')
  const manifest = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'))

  assert.ok(code.includes('dsh-office-one/office'))
  assert.ok(code.includes('sidebar.right.tab.document'))
  assert.ok(!code.includes('id: "workspace-file-viewer"'))
  assert.ok(!code.includes('label: "Workspace 文件"'))
  assert.ok(manifest.dsh.client.inject.includes('@deepseek-ai/dsh-client-ui-sidebar-documentpreview'))
})

test('Univer view reads the target-neutral Chat projection', async () => {
  const code = await readFile(new URL('../lib/client.js', import.meta.url), 'utf8')

  assert.ok(code.includes('useConversation'))
  assert.match(code, /views\?\.get\(["']chat["']\)/)
  assert.match(code, /get\(["']chat["']\)\)\?\.legacy \?\? source/)
})

test('Univer products wait for session initialization before mounting', async () => {
  const code = await readFile(new URL('../lib/client.js', import.meta.url), 'utf8')
  const initializers = code.match(/\[hostLoaded, setHostLoaded\] = .*?useState\)\(false\)/g) ?? []

  assert.equal(initializers.length, 3)
})
