import { UniverDocsTablePlugin } from '@univerjs-pro/docs-table'
import { UniverLicensePlugin } from '@univerjs-pro/license'
import { createUniver, defaultTheme, LocaleType, mergeLocales } from '@univerjs/presets'
import { UNIVER_LICENSE } from 'virtual:dsh-univer-license'
import { UniverDocsCorePreset } from '@univerjs/preset-docs-core'
import UniverPresetDocsCoreEnUS from '@univerjs/preset-docs-core/locales/en-US'
import UniverPresetDocsCoreZhCN from '@univerjs/preset-docs-core/locales/zh-CN'
import { UniverDocsDrawingPreset } from '@univerjs/preset-docs-drawing'
import UniverPresetDocsDrawingEnUS from '@univerjs/preset-docs-drawing/locales/en-US'
import UniverPresetDocsDrawingZhCN from '@univerjs/preset-docs-drawing/locales/zh-CN'
import '@univerjs/preset-docs-core/lib/index.css'
import '@univerjs/preset-docs-drawing/lib/index.css'
import '@univerjs-pro/docs-table/facade'

export function createDocsRuntime(container: HTMLElement, language: 'en' | 'zh' = 'en') {
  return createUniver({
    locale: language === 'zh' ? LocaleType.ZH_CN : LocaleType.EN_US,
    locales: {
      [LocaleType.EN_US]: mergeLocales(UniverPresetDocsCoreEnUS, UniverPresetDocsDrawingEnUS),
      [LocaleType.ZH_CN]: mergeLocales(UniverPresetDocsCoreZhCN, UniverPresetDocsDrawingZhCN),
    },
    theme: defaultTheme,
    presets: [UniverDocsCorePreset({ container, disableAutoFocus: true }), UniverDocsDrawingPreset()],
    plugins: [[UniverLicensePlugin, { license: UNIVER_LICENSE }], UniverDocsTablePlugin],
  })
}

;(globalThis as any).__DSH_WORKSPACE_FILE_VIEWER__ = {
  ...(globalThis as any).__DSH_WORKSPACE_FILE_VIEWER__,
  createDocsRuntime,
}
