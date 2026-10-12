export type CellValue = string | number | boolean | null

export interface ColumnMeta {
  id: string
  name: string
  displayName: string
  normalizedName: string
  index: number
  occurrence: number
  key: string
}

export interface SheetMeta {
  id: string
  name: string
  rowCount: number
  columnCount: number
  declaredRange: string | null
  actualRange: string
  rangeCorrected: boolean
  columns: ColumnMeta[]
}

export interface UnifiedColumn {
  key: string
  headerName: string
  displayName: string
  normalizedName: string
  occurrence: number
  fallbackColumnIndex: number
  sheetIds: string[]
  coverage: number
  totalSheets: number
}

export interface WorkbookMetadata {
  sheetCount: number
  totalRows: number
  maxColumns: number
  sheets: SheetMeta[]
  unifiedColumns: UnifiedColumn[]
}

export interface FileRecord {
  id: string
  name: string
  size: number
  format: string
  status: string
  metadata: WorkbookMetadata
  createdAt: string
  lastAccessedAt: string
}

export type Operator = 'starts_with' | 'ends_with' | 'contains' | 'not_contains' | 'empty' | 'not_empty' | 'in' | 'not_in' | 'no_prefix' | 'no_suffix'

export interface AffixSuggestion {
  value: string
  count: number
}

export interface AffixSuggestionResult {
  items: AffixSuggestion[]
  delimiter: string
  kind: 'PREFIX' | 'SUFFIX'
  scannedCount: number
  unmarkedCount: number
}

export interface ConditionOptions {
  caseSensitive: boolean
  trimCell: boolean
  trimInput: boolean
}

export interface ConditionNode {
  type: 'condition'
  columnRef: {
    headerName: string
    fallbackColumnIndex?: number
    occurrence: number
  }
  operator: Operator
  value?: string | null
  values: string[]
  options: ConditionOptions
}

export interface ConditionGroupNode {
  type: 'group'
  relation: 'AND' | 'OR'
  children: Array<ConditionNode | ConditionGroupNode>
}

export interface FilterGroup {
  id: string
  fileId: string
  name: string
  priority: number
  sourceScope: { mode: 'ALL' | 'SINGLE'; sheetId?: string | null }
  conditionTree: ConditionGroupNode
  status: string
  saved?: boolean
  active?: boolean
  sourceFileName?: string | null
  savedAt?: string | null
  createdAt: string
  updatedAt: string
}

export interface FilterRunSheet {
  key: string
  name: string
  rowCount: number
  headers: string[]
  isSystem: boolean
}

export interface FilterRun {
  id: string
  status: 'queued' | 'running' | 'succeeded' | 'failed'
  progress: number
  error?: string | null
  summary?: FilterRunSheet[]
  sourceRowCount?: number
  assignedRowCount?: number
  rowCountConserved?: boolean
  createdAt: string
  updatedAt: string
  saved: boolean
}

export interface PageResult {
  key?: string
  name?: string
  sheetId?: string
  sheetName?: string
  headers: string[]
  rows: CellValue[][]
  page: number
  pageSize: number
  total: number
  isSystem?: boolean
}

export interface SumResultItem {
  sheetId: string
  sheetName: string
  columnName: string
  dataRowCount: number
  numericCount: number
  textConvertedCount: number
  emptyCount: number
  nonNumericCount: number
  errorValueCount: number
  sum: string | null
  status: string
  message: string
}

export interface SumSource {
  id: string
  type: 'FILE' | 'FILTER_RESULT'
  name: string
  sourceFileName: string
  fileId: string
  saved: boolean
  createdAt: string
  metadata: WorkbookMetadata
}

export interface SumRun {
  id: string
  status: 'queued' | 'running' | 'succeeded' | 'failed'
  progress: number
  error?: string | null
  results?: SumResultItem[]
  invalidCount?: number
  grandTotal?: string
}

export interface InvalidDetail {
  sheetName: string
  sourceRow: number
  cell: string
  columnIndex: number
  columnName: string
  originalValue: CellValue
  reason: string
}
