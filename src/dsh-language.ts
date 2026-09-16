import { useSyncExternalStore } from 'react'

export type DshLanguage = 'en' | 'zh'

type LocaleSnapshot = { active?: unknown }
type LocaleRuntime = {
  getSnapshot: () => LocaleSnapshot
  subscribe: (listener: () => void) => () => void
}

let language: DshLanguage = 'en'
let detachRuntime: (() => void) | null = null
const listeners = new Set<() => void>()

/** Only the two bundled languages are supported. Every other DSH locale uses English. */
export function normalizeDshLanguage(value: unknown): DshLanguage {
  return value === 'zh' ? 'zh' : 'en'
}

function publish(next: DshLanguage): void {
  if (next === language) return
  language = next
  for (const listener of listeners) listener()
}

export function attachDshLocale(runtime: LocaleRuntime | null | undefined): () => void {
  detachRuntime?.()
  detachRuntime = null
  if (runtime === null || runtime === undefined) {
    publish('en')
    return () => {}
  }

  const sync = () => publish(normalizeDshLanguage(runtime.getSnapshot().active))
  sync()
  const unsubscribe = runtime.subscribe(sync)
  detachRuntime = unsubscribe
  return () => {
    if (detachRuntime !== unsubscribe) return
    detachRuntime = null
    unsubscribe()
    publish('en')
  }
}

export function getDshLanguage(): DshLanguage {
  return language
}

export function subscribeDshLanguage(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

export function useDshLanguage(): DshLanguage {
  return useSyncExternalStore(subscribeDshLanguage, getDshLanguage, () => 'en')
}
