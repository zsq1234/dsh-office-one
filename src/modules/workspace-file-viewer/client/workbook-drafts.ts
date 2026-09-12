export type WorkbookMutation = {
  id: string
  params: Record<string, unknown>
  /** Stable worksheet identity used when an XLSX import generates new sheet IDs. */
  subUnitName?: string
}

type DraftMetadata = {
  fileKey: string
  baseModified: string
  nextSequence: number
  updatedAt: number
}

type StoredMutation = WorkbookMutation & {
  fileKey: string
  sequence: number
  createdAt: number
}

export type LoadedWorkbookDraft = {
  mutations: Array<WorkbookMutation & { sequence: number }>
  lastSequence: number
  skippedMutations: number
  conflictModified?: string
}

// Formula calculation/worker messages do not modify the persisted workbook.
// Formula text itself is already contained in sheet.mutation.set-range-values.
const PERSISTENT_FORMULA_MUTATIONS = new Set([
  'formula.mutation.set-defined-name',
  'formula.mutation.remove-defined-name',
  'formula.mutation.set-super-table',
  'formula.mutation.remove-super-table',
  'formula.mutation.set-super-table-option',
])

/** Exclude transient editor units and formula-engine transport notifications. */
export function isPersistableWorkbookMutation(mutation: WorkbookMutation): boolean {
  if (mutation.id.startsWith('formula.mutation.') && !PERSISTENT_FORMULA_MUTATIONS.has(mutation.id)) return false

  const visited = new Set<object>()
  const containsInternalEditorUnit = (value: unknown): boolean => {
    if (typeof value === 'string') return value.startsWith('__INTERNAL_EDITOR__')
    if (value === null || typeof value !== 'object' || visited.has(value)) return false
    visited.add(value)
    if (Array.isArray(value)) return value.some(containsInternalEditorUnit)
    return Object.values(value as Record<string, unknown>).some(containsInternalEditorUnit)
  }
  return !containsInternalEditorUnit(mutation.params)
}

const DATABASE_NAME = 'dsh-office-one-workbook-drafts'
const DATABASE_VERSION = 2
const METADATA_STORE = 'metadata'
const MUTATION_STORE = 'mutations'
const MAX_SEQUENCE = Number.MAX_SAFE_INTEGER

let databasePromise: Promise<IDBDatabase> | null = null

function requestResult<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error ?? new Error('IndexedDB request failed'))
  })
}

function transactionDone(transaction: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = () => resolve()
    transaction.onabort = () => reject(transaction.error ?? new Error('IndexedDB transaction aborted'))
    transaction.onerror = () => reject(transaction.error ?? new Error('IndexedDB transaction failed'))
  })
}

function openDatabase(): Promise<IDBDatabase> {
  if (databasePromise !== null) return databasePromise
  databasePromise = new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') {
      reject(new Error('当前浏览器不支持 IndexedDB'))
      return
    }
    const request = indexedDB.open(DATABASE_NAME, DATABASE_VERSION)
    request.onupgradeneeded = (event) => {
      const database = request.result
      if (!database.objectStoreNames.contains(METADATA_STORE)) {
        database.createObjectStore(METADATA_STORE, { keyPath: 'fileKey' })
      }
      let mutationStore: IDBObjectStore
      if (!database.objectStoreNames.contains(MUTATION_STORE)) {
        mutationStore = database.createObjectStore(MUTATION_STORE, { keyPath: ['fileKey', 'sequence'] })
        mutationStore.createIndex('fileKey', 'fileKey', { unique: false })
      } else {
        mutationStore = request.transaction!.objectStore(MUTATION_STORE)
      }

      // Version 2 removes formula calculation notifications previously written
      // by version 1. The cursor migration cleans every workbook, not only the
      // file currently opened by the user.
      if (event.oldVersion < 2) {
        const cursorRequest = mutationStore.openCursor()
        cursorRequest.onsuccess = () => {
          const cursor = cursorRequest.result
          if (!cursor) return
          if (!isPersistableWorkbookMutation(cursor.value as StoredMutation)) cursor.delete()
          cursor.continue()
        }
      }
    }
    request.onsuccess = () => {
      const database = request.result
      database.onversionchange = () => {
        database.close()
        databasePromise = null
      }
      resolve(database)
    }
    request.onerror = () => {
      databasePromise = null
      reject(request.error ?? new Error('无法打开 IndexedDB'))
    }
    request.onblocked = () => {
      databasePromise = null
      reject(new Error('IndexedDB 升级被其他页面阻塞'))
    }
  })
  return databasePromise
}

function mutationRange(fileKey: string, startSequence = 0, endSequence = MAX_SEQUENCE): IDBKeyRange {
  return IDBKeyRange.bound([fileKey, startSequence], [fileKey, endSequence])
}

export function workbookDraftKey(sessionId: string, path: string): string {
  return `${sessionId}\u0000${path}`
}

export async function loadWorkbookDraft(fileKey: string, baseModified: string): Promise<LoadedWorkbookDraft> {
  const database = await openDatabase()
  const transaction = database.transaction([METADATA_STORE, MUTATION_STORE], 'readwrite')
  const metadataStore = transaction.objectStore(METADATA_STORE)
  const mutationStore = transaction.objectStore(MUTATION_STORE)
  const metadata = await requestResult(metadataStore.get(fileKey)) as DraftMetadata | undefined
  if (!metadata) {
    await transactionDone(transaction)
    return { mutations: [], lastSequence: 0, skippedMutations: 0 }
  }
  if (metadata.baseModified !== baseModified) {
    await transactionDone(transaction)
    return { mutations: [], lastSequence: 0, skippedMutations: 0, conflictModified: metadata.baseModified }
  }
  const stored = await requestResult(mutationStore.getAll(mutationRange(fileKey))) as StoredMutation[]
  stored.sort((left, right) => left.sequence - right.sequence)
  const persistable: StoredMutation[] = []
  const skipped: StoredMutation[] = []
  for (const mutation of stored) {
    (isPersistableWorkbookMutation(mutation) ? persistable : skipped).push(mutation)
  }
  for (const mutation of skipped) mutationStore.delete([fileKey, mutation.sequence])
  if (persistable.length === 0) metadataStore.delete(fileKey)
  await transactionDone(transaction)

  const mutations = persistable.map(({ id, params, subUnitName, sequence }) => ({ id, params, subUnitName, sequence }))
  return {
    mutations,
    lastSequence: persistable.at(-1)?.sequence ?? 0,
    skippedMutations: skipped.length,
  }
}

export async function appendWorkbookMutation(
  fileKey: string,
  baseModified: string,
  mutation: WorkbookMutation,
): Promise<number> {
  const database = await openDatabase()
  const transaction = database.transaction([METADATA_STORE, MUTATION_STORE], 'readwrite')
  const metadataStore = transaction.objectStore(METADATA_STORE)
  const mutationStore = transaction.objectStore(MUTATION_STORE)
  const current = await requestResult(metadataStore.get(fileKey)) as DraftMetadata | undefined
  if (current && current.baseModified !== baseModified) {
    transaction.abort()
    throw new Error('文件版本已变化，无法继续写入旧草稿')
  }
  const sequence = current?.nextSequence ?? 1
  const now = Date.now()
  mutationStore.put({
    fileKey,
    sequence,
    id: mutation.id,
    params: mutation.params,
    subUnitName: mutation.subUnitName,
    createdAt: now,
  } satisfies StoredMutation)
  metadataStore.put({
    fileKey,
    baseModified,
    nextSequence: sequence + 1,
    updatedAt: now,
  } satisfies DraftMetadata)
  await transactionDone(transaction)
  return sequence
}

export async function clearWorkbookDraft(fileKey: string): Promise<void> {
  const database = await openDatabase()
  const transaction = database.transaction([METADATA_STORE, MUTATION_STORE], 'readwrite')
  transaction.objectStore(METADATA_STORE).delete(fileKey)
  transaction.objectStore(MUTATION_STORE).delete(mutationRange(fileKey))
  await transactionDone(transaction)
}

/**
 * Move edits made after a save started onto the newly saved file version, while
 * removing mutations already represented by that saved file.
 */
export async function commitWorkbookDraft(
  fileKey: string,
  previousBaseModified: string,
  nextBaseModified: string,
  throughSequence: number,
): Promise<boolean> {
  const database = await openDatabase()
  const transaction = database.transaction([METADATA_STORE, MUTATION_STORE], 'readwrite')
  const metadataStore = transaction.objectStore(METADATA_STORE)
  const mutationStore = transaction.objectStore(MUTATION_STORE)
  const current = await requestResult(metadataStore.get(fileKey)) as DraftMetadata | undefined
  if (!current) {
    await transactionDone(transaction)
    return false
  }
  if (current.baseModified !== previousBaseModified) {
    transaction.abort()
    throw new Error('草稿基线已变化，无法完成压缩')
  }

  if (throughSequence > 0) mutationStore.delete(mutationRange(fileKey, 0, throughSequence))
  const remaining = await requestResult(mutationStore.count(mutationRange(fileKey, throughSequence + 1)))
  if (remaining === 0) metadataStore.delete(fileKey)
  else metadataStore.put({ ...current, baseModified: nextBaseModified, updatedAt: Date.now() } satisfies DraftMetadata)
  await transactionDone(transaction)
  return remaining > 0
}
