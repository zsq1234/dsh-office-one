import type { ClientContext } from '@deepseek-ai/dsh-client-runtime/client'

import { apply as applyUniverCreate } from './modules/univer-create/client/index.js'
import { apply as applyWorkspaceFileViewer } from './modules/workspace-file-viewer/client/index.js'

export const inject = ['slots', 'connection']

/** Mount both browser modules inside the single dsh-office-one client fiber. */
export function apply(ctx: ClientContext): void {
  applyUniverCreate(ctx)
  applyWorkspaceFileViewer(ctx)
}
