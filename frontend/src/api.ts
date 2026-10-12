import type {
  AffixSuggestionResult,
  FileRecord,
  FilterGroup,
  FilterRun,
  InvalidDetail,
  PageResult,
  SumSource,
  SumRun,
} from './types'

export const API_BASE = import.meta.env.VITE_API_BASE ?? ''

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(`${API_BASE}${path}`, init)
  if (!response.ok) {
    let message = `请求失败（${response.status}）`
    try {
      const body = await response.json()
      message = body.detail ?? message
    } catch {
      // Keep the HTTP fallback message.
    }
    throw new Error(message)
  }
  if (response.status === 204) return undefined as T
  return response.json() as Promise<T>
}

const json = (method: string, body?: unknown): RequestInit => ({
  method,
  headers: { 'Content-Type': 'application/json' },
  body: body === undefined ? undefined : JSON.stringify(body),
})

export async function uploadWorkbook(file: File): Promise<FileRecord> {
  const form = new FormData()
  form.append('file', file)
  const result = await request<{ file: FileRecord }>('/api/files', { method: 'POST', body: form })
  return result.file
}

export async function getWorkbook(fileId: string): Promise<FileRecord> {
  return (await request<{ file: FileRecord }>(`/api/files/${fileId}`)).file
}

export async function getWorkbooks(): Promise<FileRecord[]> {
  return (await request<{ files: FileRecord[] }>('/api/files')).files
}

export async function deleteWorkbook(fileId: string): Promise<void> {
  await request(`/api/files/${fileId}`, { method: 'DELETE' })
}

export async function getSourcePage(fileId: string, sheetId: string, page = 1, pageSize = 100): Promise<PageResult> {
  return request(`/api/files/${fileId}/sheets/${sheetId}/preview?page=${page}&pageSize=${pageSize}`)
}

export async function getAffixSuggestions(fileId: string, options: {
  headerName: string
  occurrence: number
  kind: 'PREFIX' | 'SUFFIX'
  delimiter?: string
  scopeMode: 'ALL' | 'SINGLE'
  sheetId?: string | null
}): Promise<AffixSuggestionResult> {
  const parameters = new URLSearchParams({
    headerName: options.headerName,
    occurrence: String(options.occurrence),
    kind: options.kind,
    delimiter: options.delimiter ?? '-',
    scopeMode: options.scopeMode,
  })
  if (options.sheetId) parameters.set('sheetId', options.sheetId)
  return request(`/api/files/${fileId}/affix-suggestions?${parameters}`)
}

export async function getGroups(fileId: string): Promise<FilterGroup[]> {
  return (await request<{ groups: FilterGroup[] }>(`/api/files/${fileId}/filter-groups`)).groups
}

export async function getSavedGroups(): Promise<FilterGroup[]> {
  return (await request<{ groups: FilterGroup[] }>('/api/filter-groups/saved')).groups
}

export async function createGroup(fileId: string, payload: Partial<FilterGroup>): Promise<FilterGroup> {
  return (await request<{ group: FilterGroup }>(`/api/files/${fileId}/filter-groups`, json('POST', payload))).group
}

export async function createGroupFromSaved(fileId: string, groupId: string): Promise<FilterGroup> {
  return (await request<{ group: FilterGroup }>(`/api/files/${fileId}/filter-groups/from-saved/${groupId}`, { method: 'POST' })).group
}

export async function updateGroup(groupId: string, payload: Partial<FilterGroup>): Promise<FilterGroup> {
  return (await request<{ group: FilterGroup }>(`/api/filter-groups/${groupId}`, json('PATCH', payload))).group
}

export async function deleteGroup(groupId: string): Promise<void> {
  await request(`/api/filter-groups/${groupId}`, { method: 'DELETE' })
}

export async function clearCurrentGroups(fileId: string): Promise<number> {
  return (await request<{ clearedCount: number }>(`/api/files/${fileId}/filter-groups/current`, { method: 'DELETE' })).clearedCount
}

export async function startFilterRun(fileId: string, orderedGroupIds: string[]): Promise<FilterRun> {
  return (await request<{ run: FilterRun }>(`/api/files/${fileId}/filter-runs`, json('POST', { orderedGroupIds }))).run
}

export async function startColumnSplitRun(fileId: string, payload: {
  columnRef: { headerName: string; occurrence: number; fallbackColumnIndex: number }
  sourceScope: { mode: 'ALL' | 'SINGLE'; sheetId?: string | null }
}): Promise<FilterRun> {
  return (await request<{ run: FilterRun }>(`/api/files/${fileId}/column-split-runs`, json('POST', payload))).run
}

export async function startAffixSplitRun(fileId: string, payload: {
  columnRef: { headerName: string; occurrence: number; fallbackColumnIndex: number }
  sourceScope: { mode: 'ALL' | 'SINGLE'; sheetId?: string | null }
  kind: 'PREFIX' | 'SUFFIX'
  delimiter: string
}): Promise<FilterRun> {
  return (await request<{ run: FilterRun }>(`/api/files/${fileId}/affix-split-runs`, json('POST', payload))).run
}

export async function getFilterRun(runId: string): Promise<FilterRun> {
  return (await request<{ run: FilterRun }>(`/api/filter-runs/${runId}`)).run
}

export async function getFilterRuns(fileId: string): Promise<FilterRun[]> {
  return (await request<{ runs: FilterRun[] }>(`/api/files/${fileId}/filter-runs`)).runs
}

export async function getTemporaryFilterRun(fileId: string): Promise<FilterRun | null> {
  return (await request<{ run: FilterRun | null }>(`/api/files/${fileId}/temporary-filter-run`)).run
}

export async function deleteFilterRun(runId: string): Promise<void> {
  await request(`/api/filter-runs/${runId}`, { method: 'DELETE' })
}

export async function saveFilterRun(runId: string): Promise<FilterRun> {
  return (await request<{ run: FilterRun }>(`/api/filter-runs/${runId}/save`, { method: 'POST' })).run
}

export async function getFilterPage(runId: string, sheetKey: string, page = 1, pageSize = 100): Promise<PageResult> {
  return request(`/api/filter-runs/${runId}/sheets/${sheetKey}/rows?page=${page}&pageSize=${pageSize}`)
}

export async function createFilterExport(runId: string): Promise<string> {
  return (await request<{ downloadUrl: string }>(`/api/filter-runs/${runId}/exports`, { method: 'POST' })).downloadUrl
}

export async function getSumSources(): Promise<SumSource[]> {
  return (await request<{ sources: SumSource[] }>('/api/sum-sources')).sources
}

export async function startSumRun(source: Pick<SumSource, 'id' | 'type'>, columnSelector: object): Promise<SumRun> {
  const payload = { sourceRef: { id: source.id, type: source.type }, columnSelector, numberPolicy: {} }
  return (await request<{ run: SumRun }>('/api/sum-runs', json('POST', payload))).run
}

export async function getSumRun(runId: string): Promise<SumRun> {
  return (await request<{ run: SumRun }>(`/api/sum-runs/${runId}`)).run
}

export async function getInvalidDetails(runId: string): Promise<InvalidDetail[]> {
  return (await request<{ items: InvalidDetail[] }>(`/api/sum-runs/${runId}/invalid-details`)).items
}

export async function createSumExport(runId: string): Promise<string> {
  return (await request<{ downloadUrl: string }>(`/api/sum-runs/${runId}/exports`, { method: 'POST' })).downloadUrl
}
