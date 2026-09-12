import { readFile } from 'node:fs/promises'
import { createRequire } from 'node:module'
import { dirname, isAbsolute, resolve, sep } from 'node:path'
import { defineConfig } from 'tsdown'
import { transform } from 'lightningcss'

const PACKAGE_ID = 'dsh-office-one'
const nodeRequire = createRequire(import.meta.url)
const CSS_PREFIX = '\0dsh-office-one-css:'
const CSS_SUFFIX = '.mjs'
const UNIVER_LICENSE_MODULE = 'virtual:dsh-univer-license'
const UNIVER_LICENSE_ID = '\0dsh-office-one-univer-license'
const UNIVER_LICENSE_FILE = resolve('license-univer/license.txt')

function cssTagId(path: string): string {
  const nodeModulesMarker = `${sep}node_modules${sep}`
  const markerIndex = path.lastIndexOf(nodeModulesMarker)
  const shortPath = markerIndex >= 0 ? path.slice(markerIndex + nodeModulesMarker.length) : path
  return `${PACKAGE_ID}/${shortPath.replace(/[^a-zA-Z0-9._@-]+/g, '-')}`
}

function dedupeRediPlugin() {
  const redi = nodeRequire.resolve('@wendellhu/redi').replace('/cjs/', '/esm/')
  const bindings = nodeRequire.resolve('@wendellhu/redi/react-bindings').replace('/cjs/', '/esm/')
  return {
    name: 'dsh-office-one-dedupe-redi',
    resolveId(source: string) {
      if (source === '@wendellhu/redi') return redi
      if (source === '@wendellhu/redi/react-bindings') return bindings
      return null
    },
  }
}

function runtimeDependencyPlugin() {
  const paths = new Map([
    ['react', nodeRequire.resolve('react')],
    ['react/jsx-runtime', nodeRequire.resolve('react/jsx-runtime')],
    ['react-dom', nodeRequire.resolve('react-dom')],
    ['react-dom/client', nodeRequire.resolve('react-dom/client')],
    ['rxjs', nodeRequire.resolve('rxjs')],
    ['rxjs/operators', nodeRequire.resolve('rxjs/operators')],
    ['@univerjs-pro/slides', nodeRequire.resolve('@univerjs-pro/slides/lib/es/index.js')],
    ['@univerjs-pro/slides/facade', nodeRequire.resolve('@univerjs-pro/slides/lib/es/facade.js')],
    ['@univerjs-pro/slides-ui', nodeRequire.resolve('@univerjs-pro/slides-ui/lib/es/index.js')],
    ['@univerjs-pro/engine-shape', nodeRequire.resolve('@univerjs-pro/engine-shape/lib/es/index.js')],
    ['@univerjs-pro/engine-formula', nodeRequire.resolve('@univerjs-pro/engine-formula/lib/es/index.js')],
    ['@univerjs/rpc', nodeRequire.resolve('@univerjs/rpc/lib/es/index.js')],
    ['@univerjs-pro/shape-editor-ui/lib/index.css', CSS_PREFIX + nodeRequire.resolve('@univerjs-pro/shape-editor-ui/lib/index.css') + CSS_SUFFIX],
    ['@univerjs-pro/slides-ui/lib/index.css', CSS_PREFIX + nodeRequire.resolve('@univerjs-pro/slides-ui/lib/index.css') + CSS_SUFFIX],
    ['virtual:dsh-shape-editor-ui-css', CSS_PREFIX + nodeRequire.resolve('@univerjs-pro/shape-editor-ui/lib/index.css') + CSS_SUFFIX],
    ['virtual:dsh-slides-css', CSS_PREFIX + nodeRequire.resolve('@univerjs-pro/slides-ui/lib/index.css') + CSS_SUFFIX],
  ])
  return {
    name: 'dsh-office-one-runtime-dependencies',
    resolveId(source: string) {
      return paths.get(source) ?? null
    },
  }
}

function univerLicensePlugin() {
  return {
    name: 'dsh-office-one-univer-license',
    resolveId(source: string) {
      return source === UNIVER_LICENSE_MODULE ? UNIVER_LICENSE_ID : null
    },
    async load(id: string) {
      if (id !== UNIVER_LICENSE_ID) return null
      const environmentLicense = process.env.UNIVER_CLIENT_LICENSE?.trim()
      if (environmentLicense) return `export const UNIVER_LICENSE = ${JSON.stringify(environmentLicense)};`

      try {
        this.addWatchFile(UNIVER_LICENSE_FILE)
        const license = (await readFile(UNIVER_LICENSE_FILE, 'utf8')).trim()
        return `export const UNIVER_LICENSE = ${JSON.stringify(license)};`
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
        return 'export const UNIVER_LICENSE = "";'
      }
    },
  }
}

function cssPlugin() {
  return {
    name: 'dsh-office-one-css-inline',
    resolveId(source: string, importer?: string) {
      if (!source.endsWith('.css')) return null
      const path = isAbsolute(source)
        ? source
        : source.startsWith('.') && importer !== undefined
          ? resolve(dirname(importer), source)
          : nodeRequire.resolve(source, { paths: importer === undefined ? undefined : [dirname(importer)] })
      return CSS_PREFIX + path + CSS_SUFFIX
    },
    async load(id: string) {
      if (!id.startsWith(CSS_PREFIX)) return null
      const path = id.slice(CSS_PREFIX.length, -CSS_SUFFIX.length)
      this.addWatchFile(path)
      const css = transform({ filename: path, code: await readFile(path), minify: true }).code.toString()
      const tagId = cssTagId(path)
      return [
        `const css = ${JSON.stringify(css)};`,
        `const tagId = ${JSON.stringify(tagId)};`,
        'if (typeof document !== "undefined" && document.querySelector(`style[data-plugin-css="${tagId}"]`) === null) {',
        '  const tag = document.createElement("style");',
        `  tag.dataset.plugin = ${JSON.stringify(PACKAGE_ID)};`,
        '  tag.dataset.pluginCss = tagId;',
        '  tag.textContent = css;',
        '  document.head.appendChild(tag);',
        '}',
        'export {};',
      ].join('\n')
    },
  }
}

// vfile uses package-private #imports with node/default conditions. Resolve
// its shipped browser adapters explicitly: this toolchain otherwise selects
// the Node adapters even for a browser build. Do not polyfill Node globally.
function vfileBrowserPlugin() {
  return {
    name: 'dsh-office-one-vfile-browser',
    resolveId(source: string, importer?: string) {
      if (!/^#min(path|proc|url)$/.test(source) || !importer) return null
      const normalized = importer.replaceAll('\\\\', '/')
      if (!normalized.includes('/node_modules/vfile/lib/')) return null
      return resolve(dirname(importer), `${source.slice(1)}.browser.js`)
    },
  }
}

const browserShared = {
  outDir: 'lib',
  format: 'cjs' as const,
  platform: 'browser' as const,
  target: 'es2022',
  dts: false,
  clean: false,
  sourcemap: true,
  define: {
    'process.env.NODE_ENV': JSON.stringify(process.env.NODE_ENV ?? 'production'),
    'import.meta.env.MODE': JSON.stringify(process.env.NODE_ENV ?? 'production'),
    'import.meta.env': JSON.stringify({ MODE: process.env.NODE_ENV ?? 'production' }),
  },
  plugins: [vfileBrowserPlugin(), dedupeRediPlugin(), univerLicensePlugin(), cssPlugin()],
}

export default defineConfig([
  {
    name: `${PACKAGE_ID}/host`,
    entry: { index: 'src/index.ts' },
    outDir: 'lib',
    format: 'esm',
    platform: 'node',
    target: 'es2024',
    dts: false,
    clean: true,
    external: [/^(?:node:|@deepseek-ai\/|zod$)/],
  },
  {
    ...browserShared,
    name: `${PACKAGE_ID}/client`,
    entry: { client: 'src/client.ts' },
    external: [/^react(?:\/|$)/, /^react-dom(?:\/|$)/],
    noExternal: (specifier: string) => !(
      specifier === 'react'
      || specifier === 'react-dom'
      || specifier.startsWith('react/')
      || specifier.startsWith('react-dom/')
    ),
    outputOptions: {
      entryFileNames: 'client.js',
      inlineDynamicImports: true,
      externalLiveBindings: false,
      banner: `window.__ModuleLoader__.load({ id: ${JSON.stringify(PACKAGE_ID)}, factory: (require) => {`,
      footer: 'return module.exports; } });',
      intro: 'var module = { exports: {} }; var exports = module.exports;',
    },
  },
  {
    ...browserShared,
    name: `${PACKAGE_ID}/sheet-runtime`,
    entry: { sheet: 'src/modules/workspace-file-viewer/client/runtimes/sheet.ts' },
    external: [/^@deepseek-ai\//],
    noExternal: ['react', 'react-dom', 'react/jsx-runtime', 'react-dom/client', '@wendellhu/redi', '@wendellhu/redi/react-bindings', 'rxjs', 'rxjs/operators', '@univerjs-pro/license', '@univerjs-pro/slides', '@univerjs-pro/slides-ui', /^@univerjs\//],
    plugins: [dedupeRediPlugin(), runtimeDependencyPlugin(), univerLicensePlugin(), cssPlugin()],
    outputOptions: {
      entryFileNames: 'runtimes/sheet.js',
      inlineDynamicImports: true,
      externalLiveBindings: false,
      banner: '(function(){',
      footer: '})();',
      intro: 'var module = { exports: {} }; var exports = module.exports;',
    },
  },
  {
    ...browserShared,
    name: `${PACKAGE_ID}/docs-runtime`,
    entry: { docs: 'src/modules/workspace-file-viewer/client/runtimes/docs.ts' },
    external: [/^@deepseek-ai\//],
    noExternal: ['react', 'react-dom', 'react/jsx-runtime', 'react-dom/client', '@wendellhu/redi', '@wendellhu/redi/react-bindings', 'rxjs', 'rxjs/operators', '@univerjs-pro/license', '@univerjs-pro/slides', '@univerjs-pro/slides-ui', /^@univerjs\//],
    plugins: [dedupeRediPlugin(), runtimeDependencyPlugin(), univerLicensePlugin(), cssPlugin()],
    outputOptions: {
      entryFileNames: 'runtimes/docs.js',
      inlineDynamicImports: true,
      externalLiveBindings: false,
      banner: '(function(){',
      footer: '})();',
      intro: 'var module = { exports: {} }; var exports = module.exports;',
    },
  },
  {
    ...browserShared,
    name: `${PACKAGE_ID}/slides-runtime`,
    entry: { slides: 'src/modules/workspace-file-viewer/client/runtimes/slides.ts' },
    external: [/^@deepseek-ai\//],
    noExternal: ['react', 'react-dom', 'react/jsx-runtime', 'react-dom/client', '@wendellhu/redi', '@wendellhu/redi/react-bindings', 'rxjs', 'rxjs/operators', '@univerjs-pro/license', '@univerjs-pro/shape-editor-ui', '@univerjs-pro/shape-editor-ui/locale/zh-CN', '@univerjs-pro/slides', '@univerjs-pro/slides/facade', '@univerjs-pro/slides-ui', '@univerjs-pro/slides-ui/locale/zh-CN', '@univerjs-pro/engine-shape', '@univerjs-pro/engine-formula', /^@univerjs\//],
    plugins: [dedupeRediPlugin(), runtimeDependencyPlugin(), univerLicensePlugin(), cssPlugin()],
    outputOptions: {
      entryFileNames: 'runtimes/slides.js',
      inlineDynamicImports: true,
      externalLiveBindings: false,
      banner: '(function(){',
      footer: '})();',
      intro: 'var module = { exports: {} }; var exports = module.exports;',
    },
  },
])
