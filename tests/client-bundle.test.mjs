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

test('workspace workbooks use isolated bounded keep-alive runtimes', async () => {
  const code = await readFile(new URL('../lib/client.js', import.meta.url), 'utf8')

  assert.ok(code.includes('dsh-wfv-sheet-runtime-parking'))
  assert.ok(code.includes('dsh-wfv-sheet-frame'))
  assert.ok(code.includes('MAX_CACHED_SHEET_RUNTIMES'))
  assert.ok(code.includes('MAX_CACHED_OFFICE_PAYLOADS'))
  assert.ok(code.includes('OFFICE_RUNTIME_LIMITS'))
  assert.ok(code.includes('officeResumeSnapshots'))
  assert.ok(code.includes('dsh-wfv-office-runtime-parking'))
  assert.match(code, /runtimeKey:/)
  assert.match(code, /sourceModified/)
  assert.match(code, /moveBefore/)
  assert.match(code, /saveLeasesRef/)
  assert.match(code, /contentDocument/)
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

test('Univer follows DSH zh/en language and falls back to English', async () => {
  const manifest = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'))
  const client = await readFile(new URL('../lib/client.js', import.meta.url), 'utf8')
  const runtimes = await Promise.all(['sheet', 'docs', 'slides'].map((name) => readFile(new URL(`../lib/runtimes/${name}.js`, import.meta.url), 'utf8')))

  assert.ok(manifest.dsh.client.inject.includes('@deepseek-ai/dsh-client-locale'))
  assert.ok(client.includes('dsh-office-one: DSH locale bridge'))
  assert.match(client, /value === ["']zh["'] \? ["']zh["'] : ["']en["']/)
  for (const code of runtimes) {
    assert.ok(code.includes('enUS'))
    assert.ok(code.includes('zhCN'))
  }
})

test('Univer Tab custom chrome ships matching English and Chinese packs', async () => {
  const client = await readFile(new URL('../lib/client.js', import.meta.url), 'utf8')
  for (const text of [
    'Save to the current session workspace',
    '保存到当前 session workspace',
    'Conversation activity',
    '对话动态',
    'New file. Choose a workspace path when saving for the first time.',
    '新文件，首次保存时可选择 workspace 内路径',
  ]) assert.ok(client.includes(text), `missing translated Univer Tab copy: ${text}`)
})
