import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { createRequire } from 'node:module'
import { pathToFileURL } from 'node:url'
import test from 'node:test'
import ts from 'typescript'
import React from 'react'
import { renderToStaticMarkup } from 'react-dom/server'

// Compile just this TSX component in memory, using the project's installed libraries.
const require = createRequire(import.meta.url)
const source = await readFile(new URL('../src/modules/workspace-file-viewer/client/markdown.tsx', import.meta.url), 'utf8')
const { outputText } = ts.transpileModule(source, { compilerOptions: { jsx: ts.JsxEmit.React, module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2022 } })
const code = outputText.replace(/from '([^']+)'/g, (_, specifier) => `from '${pathToFileURL(require.resolve(specifier)).href}'`)
const { MarkdownPreview, isMarkdown, canPreviewMarkdown } = await import(`data:text/javascript;base64,${Buffer.from(code).toString('base64')}`)
const render = (content) => renderToStaticMarkup(React.createElement(MarkdownPreview, { path: 'docs/README.md', content, fontSize: 14, onOpenFile() {} }))

test('recognizes Markdown extensions, but does not execute MDX', () => {
  for (const path of ['README.md', 'a.MARKDOWN', 'a.mdown', 'a.mkd']) assert.ok(isMarkdown(path))
  assert.equal(isMarkdown('a.mdx'), false)
  assert.equal(isMarkdown('a.txt'), false)
})

test('renders GFM and highlighted code', () => {
  const html = render('# 标题\n\n| A | B |\n| - | - |\n| 1 | 2 |\n\n- [x] 完成\n\n~~删除~~\n\n```javascript\nconst x = 1\n```')
  for (const expected of ['<h1>标题</h1>', '<table>', 'type="checkbox"', '<del>删除</del>', 'token keyword']) assert.ok(html.includes(expected), expected)
  assert.ok(render('```unknown\n<a>\n```').includes('&lt;a&gt;'))
})

test('blocks raw HTML and dangerous URLs', () => {
  const html = render('<script>alert(1)</script>\n\n<img src=x onerror=alert(1)>\n\n[x](javascript:alert%281%29)\n\n![x](data:text/html,hello)')
  assert.doesNotMatch(html, /<script|onerror=|href="javascript:|src="data:/)
})

test('handles external and relative resources safely', () => {
  assert.match(render('[site](https://example.com)'), /rel="noopener noreferrer"/)
  assert.match(render('[file](..\/README.md)'), /href="..\/README.md"/)
  assert.doesNotMatch(render('[outside](..\/..\/secret)'), /href=/)
  assert.doesNotMatch(render('![local](image.png)'), /<img/)
  assert.match(render('![remote](https://example.com/a.png)'), /referrerPolicy="no-referrer"/i)
})

test('guards large Markdown before parsing', () => {
  assert.ok(canPreviewMarkdown(''))
  assert.ok(canPreviewMarkdown('x'.repeat(256 * 1024)))
  assert.equal(canPreviewMarkdown('x'.repeat(256 * 1024 + 1)), false)
  assert.equal(canPreviewMarkdown('\n'.repeat(4000)), false)
})
