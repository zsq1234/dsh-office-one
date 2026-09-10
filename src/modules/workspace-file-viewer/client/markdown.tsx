import React, { useMemo } from 'react'
import Markdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import Prism from 'prismjs'

export function isMarkdown(path: string): boolean {
  return /\.(md|markdown|mdown|mkd)$/i.test(path)
}

export function canPreviewMarkdown(content: string): boolean {
  return content.length <= 256 * 1024 && content.split('\n', 4001).length <= 4000
}

function MarkdownCode({ className, children }: { className?: string; children?: React.ReactNode }) {
  const content = String(children ?? '')
  const language = /(?:^|\s)language-([\w-]+)/.exec(className ?? '')?.[1]
  const html = useMemo(() => {
    const grammar = language ? Prism.languages[language] : undefined
    return grammar && content.length <= 64 * 1024 ? Prism.highlight(content, grammar, language!) : null
  }, [content, language])
  return html === null
    ? <code className={className}>{children}</code>
    : <code className={className} dangerouslySetInnerHTML={{ __html: html }} />
}

// Resolve document-relative links without allowing traversal outside the workspace.
function relativeFile(path: string, href: string): string | null {
  if (!href || href.startsWith('#') || href.startsWith('/') || /^[a-z][a-z\d+.-]*:/i.test(href)) return null
  try {
    const parts = path.split('/').slice(0, -1)
    for (const part of decodeURIComponent(href.split(/[?#]/)[0]!).split('/')) {
      if (!part || part === '.') continue
      if (part === '..') { if (!parts.length) return null; parts.pop() }
      else if (part.includes('\\') || part.includes('\0')) return null
      else parts.push(part)
    }
    return parts.join('/') || null
  } catch { return null }
}

export const MarkdownPreview = React.memo(function MarkdownPreview({ path, content, fontSize, onOpenFile }: {
  path: string
  content: string
  fontSize: number
  onOpenFile: (path: string) => void
}) {
  return (
    <article className="dsh-wfv-markdown dsh-wfv-code" aria-label="Markdown 预览" style={{ fontSize }}>
      <Markdown
        remarkPlugins={[remarkGfm]}
        skipHtml
        components={{
          code: MarkdownCode,
          a: ({ href, children }) => {
            const file = relativeFile(path, href ?? '')
            if (file) return <a href={href} onClick={(event) => { event.preventDefault(); onOpenFile(file) }}>{children}</a>
            if (href && /^(https?:|mailto:)/i.test(href)) return <a href={href} target="_blank" rel="noopener noreferrer">{children}</a>
            return <span title="暂不支持此链接目标">{children}</span>
          },
          // Never interpret workspace-relative images as GUI routes or permit data URLs.
          img: ({ src, alt, title }) => src && /^https?:\/\//i.test(src)
            ? <img src={src} alt={alt ?? ''} title={title} loading="lazy" referrerPolicy="no-referrer" />
            : <span className="dsh-wfv-markdown-image">[图片：{alt || '本地图片暂不支持预览'}]</span>,
          table: ({ children }) => <div className="dsh-wfv-markdown-table"><table>{children}</table></div>,
        }}
      >{content}</Markdown>
    </article>
  )
})
