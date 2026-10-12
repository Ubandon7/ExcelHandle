import { GitBranch, LoaderCircle, Plus, Trash2 } from 'lucide-react'
import { useEffect, useRef, useState } from 'react'
import { getAffixSuggestions } from '../api'
import type { AffixSuggestion, ConditionGroupNode, ConditionNode, FilterGroup, Operator, UnifiedColumn } from '../types'

const OPERATORS: Array<{ value: Operator; label: string }> = [
  { value: 'starts_with', label: '以前缀开头' },
  { value: 'no_prefix', label: '无前缀' },
  { value: 'ends_with', label: '以后缀结尾' },
  { value: 'no_suffix', label: '无后缀' },
  { value: 'contains', label: '包含' },
  { value: 'not_contains', label: '不包含' },
  { value: 'empty', label: '为空' },
  { value: 'not_empty', label: '非空' },
  { value: 'in', label: '属于' },
  { value: 'not_in', label: '不属于' },
]

const noValueOperators: Operator[] = ['empty', 'not_empty']
const listOperators: Operator[] = ['in', 'not_in']

function makeCondition(columns: UnifiedColumn[]): ConditionNode {
  const column = columns[0]
  return {
    type: 'condition',
    columnRef: {
      headerName: column?.headerName ?? '',
      fallbackColumnIndex: column?.fallbackColumnIndex ?? 1,
      occurrence: column?.occurrence ?? 1,
    },
    operator: 'contains',
    value: '',
    values: [],
    options: { caseSensitive: false, trimCell: false, trimInput: true },
  }
}

export function newConditionGroup(columns: UnifiedColumn[]): ConditionGroupNode {
  return { type: 'group', relation: 'AND', children: [makeCondition(columns)] }
}

interface SharedProps {
  columns: UnifiedColumn[]
  fileId: string
  sourceScope: FilterGroup['sourceScope']
}

interface TreeProps extends SharedProps {
  value: ConditionGroupNode
  onChange: (value: ConditionGroupNode) => void
}

export function ConditionTreeEditor({ value, columns, fileId, sourceScope, onChange }: TreeProps) {
  return <GroupEditor group={value} columns={columns} fileId={fileId} sourceScope={sourceScope} onChange={onChange} depth={0} root />
}

function GroupEditor({ group, columns, fileId, sourceScope, onChange, depth, root, onRemove }: SharedProps & {
  group: ConditionGroupNode
  onChange: (group: ConditionGroupNode) => void
  depth: number
  root?: boolean
  onRemove?: () => void
}) {
  const updateChild = (index: number, child: ConditionNode | ConditionGroupNode) => {
    const children = [...group.children]
    children[index] = child
    onChange({ ...group, children })
  }

  const removeChild = (index: number) => {
    onChange({ ...group, children: group.children.filter((_, itemIndex) => itemIndex !== index) })
  }

  return (
    <div className={`condition-group depth-${Math.min(depth, 3)}`}>
      <div className="group-toolbar">
        <div className="relation-toggle" aria-label="条件关系">
          <button className={group.relation === 'AND' ? 'active' : ''} onClick={() => onChange({ ...group, relation: 'AND' })}>且 AND</button>
          <button className={group.relation === 'OR' ? 'active' : ''} onClick={() => onChange({ ...group, relation: 'OR' })}>或 OR</button>
        </div>
        <div className="group-actions">
          <button className="text-button" title="在当前 AND/OR 层增加一条判断" onClick={() => onChange({ ...group, children: [...group.children, makeCondition(columns)] })}>
            <Plus size={15} /> 添加条件
          </button>
          {depth < 4 && (
            <button className="text-button" title="增加一层可独立设置 AND/OR 的括号逻辑" onClick={() => onChange({ ...group, children: [...group.children, newConditionGroup(columns)] })}>
              <GitBranch size={15} /> 添加条件组
            </button>
          )}
          {!root && <button className="icon-button danger" title="删除条件组" onClick={onRemove}><Trash2 size={16} /></button>}
        </div>
      </div>

      <div className="condition-list">
        {group.children.length === 0 && <div className="empty-inline">还没有条件，请先添加一个筛选条件。</div>}
        {group.children.map((child, index) => child.type === 'condition' ? (
          <ConditionEditor
            key={index}
            condition={child}
            columns={columns}
            fileId={fileId}
            sourceScope={sourceScope}
            onChange={(next) => updateChild(index, next)}
            onRemove={() => removeChild(index)}
          />
        ) : (
          <GroupEditor
            key={index}
            group={child}
            columns={columns}
            fileId={fileId}
            sourceScope={sourceScope}
            depth={depth + 1}
            onChange={(next) => updateChild(index, next)}
            onRemove={() => removeChild(index)}
          />
        ))}
      </div>
    </div>
  )
}

function ConditionEditor({ condition, columns, fileId, sourceScope, onChange, onRemove }: SharedProps & {
  condition: ConditionNode
  onChange: (condition: ConditionNode) => void
  onRemove: () => void
}) {
  const selectedKey = columns.find((column) =>
    column.headerName === condition.columnRef.headerName && column.occurrence === condition.columnRef.occurrence
  )?.key ?? ''
  const needsValue = !noValueOperators.includes(condition.operator)
  const needsList = listOperators.includes(condition.operator)
  const isAffix = ['starts_with', 'ends_with'].includes(condition.operator)
  const isNoAffix = ['no_prefix', 'no_suffix'].includes(condition.operator)

  return (
    <div className="condition-row">
      <select
        aria-label="选择列"
        value={selectedKey}
        onChange={(event) => {
          const column = columns.find((item) => item.key === event.target.value)
          if (!column) return
          onChange({ ...condition, columnRef: { headerName: column.headerName, fallbackColumnIndex: column.fallbackColumnIndex, occurrence: column.occurrence } })
        }}
      >
        {columns.map((column) => <option key={column.key} value={column.key}>{column.displayName}</option>)}
      </select>
      <select aria-label="筛选方式" value={condition.operator} onChange={(event) => {
        const operator = event.target.value as Operator
        onChange({ ...condition, operator, value: ['no_prefix', 'no_suffix'].includes(operator) ? (condition.value || '-') : condition.value })
      }}>
        {OPERATORS.map((operator) => <option key={operator.value} value={operator.value}>{operator.label}</option>)}
      </select>
      {needsValue ? (needsList ? (
        <ListValueInput condition={condition} onChange={onChange} />
      ) : isAffix ? (
        <AffixValueInput condition={condition} fileId={fileId} sourceScope={sourceScope} onChange={onChange} />
      ) : isNoAffix ? (
        <DelimiterInput condition={condition} onChange={onChange} />
      ) : (
        <input aria-label="筛选值" value={condition.value ?? ''} placeholder="输入筛选内容" onChange={(event) => onChange({ ...condition, value: event.target.value })} />
      )) : <div className="value-placeholder">无需输入值</div>}
      <label className="check-label" title="开启后区分英文字母大小写">
        <input type="checkbox" checked={condition.options.caseSensitive} onChange={(event) => onChange({ ...condition, options: { ...condition.options, caseSensitive: event.target.checked } })} />
        区分大小写
      </label>
      <button className="icon-button danger" title="删除条件" onClick={onRemove}><Trash2 size={16} /></button>
    </div>
  )
}

function AffixValueInput({ condition, fileId, sourceScope, onChange }: {
  condition: ConditionNode
  fileId: string
  sourceScope: FilterGroup['sourceScope']
  onChange: (condition: ConditionNode) => void
}) {
  const [items, setItems] = useState<AffixSuggestion[]>([])
  const [open, setOpen] = useState(false)
  const [loading, setLoading] = useState(false)
  const [loadedKey, setLoadedKey] = useState('')
  const kind = condition.operator === 'starts_with' ? 'PREFIX' : 'SUFFIX'
  const query = condition.value ?? ''
  const requestKey = `${fileId}|${sourceScope.mode}|${sourceScope.sheetId ?? ''}|${condition.columnRef.headerName}|${condition.columnRef.occurrence}|${kind}`
  const visibleItems = items.filter((item) => item.value.toLocaleLowerCase().includes(query.toLocaleLowerCase())).slice(0, 100)

  const load = async () => {
    setOpen(true)
    if (loadedKey === requestKey) return
    setLoading(true)
    setItems([])
    try {
      const result = await getAffixSuggestions(fileId, {
        headerName: condition.columnRef.headerName,
        occurrence: condition.columnRef.occurrence,
        kind,
        delimiter: '-',
        scopeMode: sourceScope.mode,
        sheetId: sourceScope.sheetId,
      })
      setItems(result.items)
      setLoadedKey(requestKey)
    } catch {
      setItems([])
      setLoadedKey(requestKey)
    } finally { setLoading(false) }
  }

  return (
    <div className="suggestion-field">
      <input
        aria-label="筛选值"
        value={query}
        placeholder={kind === 'PREFIX' ? '点击选择识别出的前缀' : '点击选择识别出的后缀'}
        onFocus={() => void load()}
        onClick={() => void load()}
        onBlur={() => window.setTimeout(() => setOpen(false), 120)}
        onChange={(event) => { onChange({ ...condition, value: event.target.value }); setOpen(true) }}
      />
      {open && <div className="suggestion-menu">
        {loading ? <div className="suggestion-empty"><LoaderCircle className="spin" size={15} /> 正在识别此列…</div> : visibleItems.length ? visibleItems.map((item) => (
          <button key={item.value} onMouseDown={(event) => event.preventDefault()} onClick={() => { onChange({ ...condition, value: item.value }); setOpen(false) }}><span>{item.value}</span><small>{item.count} 项</small></button>
        )) : <div className="suggestion-empty">此列没有可选的“-”前后缀</div>}
        {!loading && items.length > 100 && <div className="suggestion-tip">共 {items.length} 项，可输入文字继续筛选</div>}
      </div>}
    </div>
  )
}

function DelimiterInput({ condition, onChange }: { condition: ConditionNode; onChange: (condition: ConditionNode) => void }) {
  return (
    <div className="delimiter-field">
      <input list="affix-delimiters" maxLength={5} aria-label="前后缀标识符" value={condition.value ?? '-'} placeholder="选择或输入标识符" onChange={(event) => onChange({ ...condition, value: event.target.value })} />
      <datalist id="affix-delimiters"><option value="-" /><option value="_" /><option value="@" /><option value="/" /><option value="." /></datalist>
      <span>标识符</span>
    </div>
  )
}

function ListValueInput({ condition, onChange }: { condition: ConditionNode; onChange: (condition: ConditionNode) => void }) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [text, setText] = useState(condition.values.join('，'))

  useEffect(() => {
    if (document.activeElement !== inputRef.current) setText(condition.values.join('，'))
  }, [condition.values])

  return (
    <input
      ref={inputRef}
      aria-label="筛选值"
      value={text}
      placeholder="多个值用逗号分隔"
      onChange={(event) => {
        const raw = event.target.value
        setText(raw)
        onChange({ ...condition, values: raw.split(/[,，\n]/).map((item) => item.trim()).filter(Boolean) })
      }}
    />
  )
}
