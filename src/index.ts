import type { Context } from '@deepseek-ai/cordis'

import { apply as applyUniverCreate } from './modules/univer-create/host.js'
import { apply as applyUniverFileExport } from './modules/univer-file-export/index.js'
import { apply as applyWorkspaceFileViewer } from './modules/workspace-file-viewer/host/index.js'

export const name = 'dsh-office-one'

// The wrapper owns the complete service contract of its three internal modules.
export const inject = [
  'tools',
  'connection',
  'storageDomain',
  'workspaceRegistry',
  'webServer',
  'attachments',
]

/** Mount the three independently maintained Host modules as one DSH plugin. */
export function apply(ctx: Context): void {
  applyUniverFileExport(ctx)
  applyWorkspaceFileViewer(ctx)
  applyUniverCreate(ctx)
}
