import { UniverLicensePlugin } from '@univerjs-pro/license'
import { UNIVER_LICENSE } from 'virtual:dsh-univer-license'
import ShapeEditorUIZhCN from '@univerjs-pro/shape-editor-ui/locale/zh-CN'
import { UniverSlidesPlugin } from '@univerjs-pro/slides'
import { UniverSlidesUIPlugin } from '@univerjs-pro/slides-ui'
import SlidesUIZhCN from '@univerjs-pro/slides-ui/locale/zh-CN'
import { LocaleType, mergeLocales, Univer } from '@univerjs/core'
import { FUniver } from '@univerjs/core/facade'
import '@univerjs-pro/slides/facade'
import DesignZhCN from '@univerjs/design/locale/zh-CN'
import { UniverDocsPlugin } from '@univerjs/docs'
import { UniverDocsUIPlugin } from '@univerjs/docs-ui'
import DocsUIZhCN from '@univerjs/docs-ui/locale/zh-CN'
import { UniverDrawingPlugin } from '@univerjs/drawing'
import { UniverRenderEnginePlugin } from '@univerjs/engine-render'
import { UniverUIPlugin } from '@univerjs/ui'
import UIZhCN from '@univerjs/ui/locale/zh-CN'

import '../styles.css'
import '@univerjs/design/lib/index.css'
import '@univerjs/ui/lib/index.css'
import '@univerjs/docs-ui/lib/index.css'
import 'virtual:dsh-shape-editor-ui-css'
import 'virtual:dsh-slides-css'

export function createSlidesRuntime(container: HTMLElement) {
  const univer = new Univer({
    locale: LocaleType.ZH_CN,
    locales: {
      [LocaleType.ZH_CN]: mergeLocales(DesignZhCN, UIZhCN, DocsUIZhCN, ShapeEditorUIZhCN, SlidesUIZhCN),
    },
  })

  univer.registerPlugin(UniverRenderEnginePlugin)
  univer.registerPlugin(UniverUIPlugin, { container })
  univer.registerPlugin(UniverDocsPlugin)
  univer.registerPlugin(UniverDocsUIPlugin)
  univer.registerPlugin(UniverDrawingPlugin)
  univer.registerPlugin(UniverLicensePlugin, { license: UNIVER_LICENSE })
  univer.registerPlugin(UniverSlidesPlugin)
  univer.registerPlugin(UniverSlidesUIPlugin)

  return univer
}

export function createSlide(runtime: ReturnType<typeof createSlidesRuntime>, snapshot: unknown) {
  const univerAPI = FUniver.newAPI(runtime) as FUniver & {
    createPresentation: (data: unknown) => { save: () => unknown }
  }
  return univerAPI.createPresentation(snapshot)
}

;(globalThis as any).__DSH_WORKSPACE_FILE_VIEWER__ = {
  ...(globalThis as any).__DSH_WORKSPACE_FILE_VIEWER__,
  createSlidesRuntime,
  createSlide,
}
