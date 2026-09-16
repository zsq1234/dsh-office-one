import type { ClientContext } from '@deepseek-ai/dsh-client-runtime/client'

import { attachDshLocale } from './dsh-language.js'
import { apply as applyUniverCreate } from './modules/univer-create/client/index.js'
import { apply as applyWorkspaceFileViewer } from './modules/workspace-file-viewer/client/index.js'

export const inject = ['slots', 'connection', 'documentPreviews', 'locale']

/** Mount both browser modules inside the single dsh-office-one client fiber. */
export function apply(ctx: ClientContext): void {
  ctx.effect(() => attachDshLocale((ctx as ClientContext & { locale?: Parameters<typeof attachDshLocale>[0] }).locale), 'dsh-office-one: DSH locale bridge')
  applyUniverCreate(ctx)
  applyWorkspaceFileViewer(ctx)
}
