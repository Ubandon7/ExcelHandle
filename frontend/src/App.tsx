import { useEffect, useRef, useState } from 'react'
import {
  ArrowDown, ArrowUp, Calculator, Check, ChevronLeft, ChevronRight,
  Download, FileSpreadsheet, Filter, Layers3, LayoutDashboard, LoaderCircle, Pencil, Plus,
  Save, Search, Trash2, UploadCloud, X,
} from 'lucide-react'
import {
  API_BASE, clearCurrentGroups, createFilterExport, createGroup, createGroupFromSaved, createSumExport, deleteGroup,
  deleteFilterRun, deleteWorkbook, getFilterPage, getFilterRun, getFilterRuns, getGroups, getInvalidDetails, getSavedGroups,
  getTemporaryFilterRun,
  getSourcePage, getSumRun, getSumSources, getWorkbooks, saveFilterRun, startAffixSplitRun, startColumnSplitRun, startFilterRun, startSumRun,
  updateGroup, uploadWorkbook,
} from './api'
import { ConditionTreeEditor, newConditionGroup } from './components/ConditionTree'
import type {
  CellValue, FileRecord, FilterGroup, FilterRun, InvalidDetail, PageResult,
  SumRun, SumSource, UnifiedColumn,
} from './types'
import './styles.css'

type PageName = 'file' | 'filter' | 'sum'

const sleep = (milliseconds: number) => new Promise((resolve) => window.setTimeout(resolve, milliseconds))
const formatNumber = (value: number) => new Intl.NumberFormat('zh-CN').format(value)
const formatBytes = (bytes: number) => bytes > 1024 * 1024
  ? `${(bytes / 1024 / 1024).toFixed(1)} MB`
  : `${Math.max(1, Math.round(bytes / 1024))} KB`

async function pollRun<T extends { status: string }>(getter: () => Promise<T>, onChange: (run: T) => void): Promise<T> {
  for (let attempt = 0; attempt < 240; attempt += 1) {
    const run = await getter()
    onChange(run)
    if (run.status === 'succeeded') return run
    if (run.status === 'failed') throw new Error(('error' in run && String(run.error)) || '任务执行失败')
    await sleep(500)
  }
  throw new Error('处理超时，请稍后重试')
}

export default function App() {
  const [files, setFiles] = useState<FileRecord[]>([])
  const [file, setFile] = useState<FileRecord | null>(null)
  const [page, setPage] = useState<PageName>('file')
  const [sumRefreshToken, setSumRefreshToken] = useState(0)
  const [loading, setLoading] = useState(true)
  const [message, setMessage] = useState<{ text: string; error?: boolean } | null>(null)
  const [sidebarDragging, setSidebarDragging] = useState(false)
  const [sidebarUploading, setSidebarUploading] = useState(false)

  useEffect(() => {
    const preferredId = localStorage.getItem('excel-handle-file-id')
    getWorkbooks().then((items) => {
      setFiles(items)
      const selected = items.find((item) => item.id === preferredId) ?? items[0] ?? null
      setFile(selected)
      if (selected) localStorage.setItem('excel-handle-file-id', selected.id)
      else localStorage.removeItem('excel-handle-file-id')
    }).catch(() => localStorage.removeItem('excel-handle-file-id')).finally(() => setLoading(false))
  }, [])

  const acceptFile = (next: FileRecord) => {
    setFiles((items) => [next, ...items.filter((item) => item.id !== next.id)])
    setFile(next)
    localStorage.setItem('excel-handle-file-id', next.id)
    setMessage({ text: '文件解析完成，可以开始处理了。' })
  }

  const clearFile = async () => {
    if (!file) return
    await deleteWorkbook(file.id)
    const remaining = files.filter((item) => item.id !== file.id)
    const next = remaining[0] ?? null
    setFiles(remaining)
    setFile(next)
    if (next) localStorage.setItem('excel-handle-file-id', next.id)
    else localStorage.removeItem('excel-handle-file-id')
    setPage('file')
    setMessage({ text: '文件及其处理记录已删除。' })
  }

  const selectFile = (selected: FileRecord) => {
    setFile(selected)
    localStorage.setItem('excel-handle-file-id', selected.id)
    setPage('file')
  }

  const uploadToLibrary = async (selected: FileList) => {
    const items = Array.from(selected)
    if (!items.length || sidebarUploading) return
    setSidebarUploading(true)
    let uploaded = 0
    try {
      for (const item of items) {
        const next = await uploadWorkbook(item)
        acceptFile(next)
        uploaded += 1
      }
      setMessage({ text: `已上传 ${uploaded} 个工作簿。` })
    } catch (error) {
      setMessage({ text: `${uploaded ? `已成功上传 ${uploaded} 个；` : ''}${errorMessage(error)}`, error: true })
    } finally {
      setSidebarUploading(false)
    }
  }

  if (loading) return <div className="center-screen"><LoaderCircle className="spin" /> 正在恢复工作区…</div>

  return (
    <div className="app-shell">
      <header className="topbar">
        <button className="brand" onClick={() => setPage('file')}>
          <span className="brand-mark"><FileSpreadsheet size={23} /></span>
          <span><strong>ExcelHandle</strong><small>表格处理工作台</small></span>
        </button>
        <div className="top-note">文件仅在本机处理 · 结果由你决定保存或清理</div>
      </header>

      <aside className="sidebar">
        <nav>
          <NavButton active={page === 'file'} icon={<LayoutDashboard size={18} />} onClick={() => setPage('file')}>文件概览</NavButton>
          <NavButton disabled={!file} active={page === 'filter'} icon={<Filter size={18} />} onClick={() => setPage('filter')}>筛选处理</NavButton>
          <NavButton disabled={!file} active={page === 'sum'} icon={<Calculator size={18} />} onClick={() => setPage('sum')}>跨 Sheet 求和</NavButton>
        </nav>
        <div
          className={`sidebar-library ${sidebarDragging ? 'dragging' : ''}`}
          onDragOver={(event) => { event.preventDefault(); setSidebarDragging(true) }}
          onDragLeave={(event) => { if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setSidebarDragging(false) }}
          onDrop={(event) => { event.preventDefault(); setSidebarDragging(false); void uploadToLibrary(event.dataTransfer.files) }}
        >
          <div className="sidebar-library-title"><span>工作簿</span><small>{files.length} 个文件</small></div>
          {files.length > 0 && <div className="sidebar-file-list">{files.map((item) => (
            <button className={`sidebar-file ${item.id === file?.id ? 'active' : ''}`} key={item.id} onClick={() => selectFile(item)}>
              <FileSpreadsheet size={20} />
              <span><strong title={item.name}>{item.name}</strong><small>{formatBytes(item.size)} · {item.metadata.sheetCount} 个 Sheet</small></span>
            </button>
          ))}</div>}
          <div className="sidebar-drop-hint">{sidebarUploading ? <><LoaderCircle className="spin" size={15} /> 正在上传…</> : <><UploadCloud size={15} /> 拖动文件到此处上传</>}</div>
        </div>
      </aside>

      <main className="main-content">
        {!file ? <UploadWelcome onUploaded={acceptFile} /> : <>
          <div className="page-pane" hidden={page !== 'file'}><FileOverview file={file} onUploaded={acceptFile} onDelete={clearFile} setMessage={setMessage} /></div>
          <div className="page-pane" hidden={page !== 'filter'}><FilterWorkspace file={file} setMessage={setMessage} onResultChanged={() => setSumRefreshToken((value) => value + 1)} /></div>
          <div className="page-pane" hidden={page !== 'sum'}><SumWorkspace preferredFileId={file.id} active={page === 'sum'} refreshToken={sumRefreshToken} setMessage={setMessage} /></div>
        </>}
      </main>

      {message && <div className={`toast ${message.error ? 'error' : ''}`}><span>{message.error ? <X size={17} /> : <Check size={17} />}{message.text}</span><button onClick={() => setMessage(null)}><X size={15} /></button></div>}
    </div>
  )
}

function NavButton({ children, icon, active, disabled, onClick }: { children: React.ReactNode; icon: React.ReactNode; active?: boolean; disabled?: boolean; onClick: () => void }) {
  return <button className={`nav-button ${active ? 'active' : ''}`} disabled={disabled} onClick={onClick}>{icon}<span>{children}</span></button>
}

function UploadWelcome({ onUploaded }: { onUploaded: (file: FileRecord) => void }) {
  return (
    <div className="welcome-page">
      <div className="welcome-copy"><span className="eyebrow">无需安装 Excel</span><h1>把繁琐的表格处理，<br /><em>变成几次点击。</em></h1><p>上传工作簿，组合任意筛选条件，或对所有 Sheet 的同一列快速求和。</p></div>
      <UploadCard onUploaded={onUploaded} />
      <div className="feature-strip"><span><strong>50,000</strong> 行以内流畅处理</span><span><strong>AND / OR</strong> 多层条件组</span><span><strong>全格式</strong> Excel 文件支持</span></div>
    </div>
  )
}

function UploadCard({ onUploaded, compact }: { onUploaded: (file: FileRecord) => void; compact?: boolean }) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [busy, setBusy] = useState(false)
  const [dragging, setDragging] = useState(false)

  const handleFile = async (selected?: File) => {
    if (!selected) return
    setBusy(true)
    try { onUploaded(await uploadWorkbook(selected)) }
    catch (error) { window.alert(error instanceof Error ? error.message : '上传失败') }
    finally { setBusy(false) }
  }

  return (
    <div className={`upload-card ${compact ? 'compact' : ''} ${dragging ? 'dragging' : ''}`}
      onDragOver={(event) => { event.preventDefault(); setDragging(true) }}
      onDragLeave={() => setDragging(false)}
      onDrop={(event) => { event.preventDefault(); setDragging(false); void handleFile(event.dataTransfer.files[0]) }}>
      <input ref={inputRef} hidden type="file" accept=".xlsx,.xls,.xlsm,.xlsb,.xltx,.xltm,.xlt,.xml,.csv,.tsv" onChange={(event) => void handleFile(event.target.files?.[0])} />
      <span className="upload-icon">{busy ? <LoaderCircle className="spin" /> : <UploadCloud />}</span>
      <div><strong>{busy ? '正在解析文件…' : compact ? '添加新文件' : '拖放 Excel 文件到这里'}</strong><p>{compact ? '保留现有文件并切换到新工作簿' : '或点击选择本地文件，最大 100 MB'}</p></div>
      <button className="primary-button" disabled={busy} onClick={() => inputRef.current?.click()}>{busy ? '处理中' : '选择文件'}</button>
    </div>
  )
}

function FileOverview({ file, onUploaded, onDelete, setMessage }: { file: FileRecord; onUploaded: (file: FileRecord) => void; onDelete: () => Promise<void>; setMessage: MessageSetter }) {
  const [sheetId, setSheetId] = useState(file.metadata.sheets[0]?.id ?? '')
  const [preview, setPreview] = useState<PageResult | null>(null)
  const [busy, setBusy] = useState(false)

  useEffect(() => {
    const effectiveSheetId = file.metadata.sheets.some((sheet) => sheet.id === sheetId)
      ? sheetId
      : file.metadata.sheets[0]?.id ?? ''
    if (effectiveSheetId !== sheetId) {
      setSheetId(effectiveSheetId)
      setPreview(null)
      return
    }
    if (!effectiveSheetId) return
    setBusy(true)
    getSourcePage(file.id, effectiveSheetId).then(setPreview).catch((error) => setMessage({ text: String(error), error: true })).finally(() => setBusy(false))
  }, [file.id, sheetId, setMessage])

  return (
    <div className="page-stack">
      <PageHeader eyebrow="当前工作簿" title="文件概览" description="确认识别结果并预览原始数据。" />
      <div className="stats-grid">
        <StatCard label="工作表" value={file.metadata.sheetCount} suffix="个 Sheet" />
        <StatCard label="数据总量" value={formatNumber(file.metadata.totalRows)} suffix="行" />
        <StatCard label="最大列数" value={file.metadata.maxColumns} suffix="列" />
        <StatCard label="文件大小" value={formatBytes(file.size)} suffix={file.format.toUpperCase()} />
      </div>
      <section className="panel file-summary">
        <div className="file-badge"><FileSpreadsheet size={30} /></div><div className="grow"><h3>{file.name}</h3><p>已成功识别 {file.metadata.sheetCount} 个 Sheet；如源文件维度记录不正确，系统已按实际单元格范围修正。</p></div>
        <UploadCard compact onUploaded={onUploaded} />
        <button className="secondary-button danger-outline" onClick={() => void onDelete()}><Trash2 size={16} /> 删除</button>
      </section>
      <section className="panel">
        <div className="panel-header"><div><h2>原始数据预览</h2><p>一次显示前 100 行，完整数据将在处理时使用。</p></div><select value={sheetId} onChange={(event) => setSheetId(event.target.value)}>{file.metadata.sheets.map((sheet) => <option value={sheet.id} key={sheet.id}>{sheet.name}（{formatNumber(sheet.rowCount)} 行）</option>)}</select></div>
        {busy ? <LoadingBlock /> : preview && <DataTable page={preview} />}
      </section>
    </div>
  )
}

type MessageSetter = (value: { text: string; error?: boolean } | null) => void

function FilterWorkspace({ file, setMessage, onResultChanged }: { file: FileRecord; setMessage: MessageSetter; onResultChanged: () => void }) {
  const [groups, setGroups] = useState<FilterGroup[]>([])
  const [savedGroups, setSavedGroups] = useState<FilterGroup[]>([])
  const [activeId, setActiveId] = useState<string | null>(null)
  const [draft, setDraft] = useState<FilterGroup | null>(null)
  const [dirty, setDirty] = useState(false)
  const [saving, setSaving] = useState(false)
  const [busy, setBusy] = useState(true)
  const [run, setRun] = useState<FilterRun | null>(null)
  const [history, setHistory] = useState<FilterRun[]>([])
  const [selectedResult, setSelectedResult] = useState<string | null>(null)
  const [resultPage, setResultPage] = useState<PageResult | null>(null)
  const [exporting, setExporting] = useState(false)
  const [splitScope, setSplitScope] = useState<string>('ALL')
  const [splitColumnKey, setSplitColumnKey] = useState(file.metadata.unifiedColumns[0]?.key ?? '')
  const [splitMethod, setSplitMethod] = useState<'OVERALL' | 'PREFIX' | 'SUFFIX'>('OVERALL')
  const [affixDelimiter, setAffixDelimiter] = useState('-')

  const selectedSheet = run?.summary?.find((sheet) => sheet.key === selectedResult)
  const splitColumns = splitScope === 'ALL'
    ? file.metadata.unifiedColumns.map((column) => ({
      key: column.key, headerName: column.headerName, displayName: column.displayName,
      occurrence: column.occurrence, fallbackColumnIndex: column.fallbackColumnIndex,
      coverage: column.coverage, totalSheets: column.totalSheets,
    }))
    : (file.metadata.sheets.find((sheet) => sheet.id === splitScope)?.columns ?? []).map((column) => ({
      key: column.key, headerName: column.name, displayName: column.displayName,
      occurrence: column.occurrence, fallbackColumnIndex: column.index,
      coverage: 1, totalSheets: 1,
    }))
  const activeGroups = groups.filter((group) => group.active !== false)

  const reloadGroups = async (preferId?: string) => {
    const [items, savedItems] = await Promise.all([getGroups(file.id), getSavedGroups()])
    setGroups(items)
    setSavedGroups(savedItems)
    const activeItems = items.filter((group) => group.active !== false)
    const nextId = activeItems.find((group) => group.id === preferId)?.id
      ?? activeItems.find((group) => group.id === activeId)?.id
      ?? activeItems[0]?.id
      ?? null
    setActiveId(nextId)
    setDraft(items.find((item) => item.id === nextId) ?? null)
    setDirty(false)
  }

  useEffect(() => {
    let cancelled = false
    setBusy(true)
    setGroups([])
    setSavedGroups([])
    setActiveId(null)
    setDraft(null)
    setDirty(false)
    setRun(null)
    setHistory([])
    setSelectedResult(null)
    setResultPage(null)
    setExporting(false)
    setSplitScope('ALL')
    setSplitColumnKey(file.metadata.unifiedColumns[0]?.key ?? '')
    setSplitMethod('OVERALL')
    setAffixDelimiter('-')
    Promise.all([getGroups(file.id), getSavedGroups(), getFilterRuns(file.id), getTemporaryFilterRun(file.id)]).then(async ([items, savedItems, runs, temporary]) => {
      if (cancelled) return
      setGroups(items)
      setSavedGroups(savedItems)
      const activeItems = items.filter((group) => group.active !== false)
      setActiveId(activeItems[0]?.id ?? null)
      setDraft(activeItems[0] ?? null)
      setDirty(false)
      setHistory(runs)
      let latest = temporary ?? runs.find((item) => item.status === 'succeeded' && item.summary?.length)
      if (!latest) return
      if (['queued', 'running'].includes(latest.status)) {
        latest = await pollRun(() => getFilterRun(latest!.id), (current) => { if (!cancelled) setRun(current) })
      }
      if (latest.status !== 'succeeded' || !latest.summary?.length) return
      setRun(latest)
      const firstKey = latest.summary?.[0]?.key ?? null
      setSelectedResult(firstKey)
      if (firstKey) {
        const page = await getFilterPage(latest.id, firstKey)
        if (!cancelled) setResultPage(page)
      }
    }).catch((error) => setMessage({ text: errorMessage(error), error: true })).finally(() => {
      if (!cancelled) setBusy(false)
    })
    return () => { cancelled = true }
  }, [file.id, setMessage])

  const addGroup = async () => {
    try {
      const savedPrevious = dirty
      if (dirty) await saveDraft()
      const created = await createGroup(file.id, {
        name: nextDefaultSheetName(groups),
        sourceScope: { mode: 'ALL', sheetId: null },
        conditionTree: newConditionGroup(file.metadata.unifiedColumns),
      })
      await reloadGroups(created.id)
      setMessage({ text: savedPrevious ? '上一结果组已自动保存，并已创建新的结果组。' : '已创建新的结果组。' })
    } catch (error) { setMessage({ text: errorMessage(error), error: true }) }
  }

  const saveDraft = async (markSaved = false) => {
    if (!draft) return null
    setSaving(true)
    try {
      const payload: Partial<FilterGroup> = { name: draft.name, sourceScope: draft.sourceScope, conditionTree: draft.conditionTree }
      if (markSaved) payload.saved = true
      const saved = await updateGroup(draft.id, payload)
      setGroups((items) => items.map((item) => item.id === saved.id ? saved : item))
      if (saved.saved) setSavedGroups((items) => [saved, ...items.filter((item) => item.id !== saved.id)])
      setDraft(saved)
      setDirty(false)
      return saved
    } finally { setSaving(false) }
  }

  const selectGroup = async (groupId: string) => {
    if (groupId === activeId) return
    try {
      if (dirty) await saveDraft()
      const selected = groups.find((group) => group.id === groupId) ?? null
      setActiveId(groupId)
      setDraft(selected)
      setDirty(false)
    } catch (error) { setMessage({ text: errorMessage(error), error: true }) }
  }

  const editSavedGroup = async (groupId: string) => {
    const group = savedGroups.find((item) => item.id === groupId)
    if (!group) return
    if (group.fileId !== file.id) {
      try {
        const created = await createGroupFromSaved(file.id, groupId)
        await reloadGroups(created.id)
        setMessage({ text: `已将全局条件“${group.name}”加入当前工作簿。` })
      } catch (error) { setMessage({ text: errorMessage(error), error: true }); return }
    } else if (group.active === false) {
      try {
        await updateGroup(groupId, { active: true, priority: activeGroups.length + 1 })
        await reloadGroups(groupId)
      } catch (error) { setMessage({ text: errorMessage(error), error: true }); return }
    } else {
      await selectGroup(groupId)
    }
    document.getElementById('filter-condition-editor')?.scrollIntoView({ behavior: 'smooth', block: 'start' })
  }

  const removeGroup = async (group: FilterGroup) => {
    const message = group.saved
      ? `从当前结果组移除“${group.name}”？已保存的筛选条件会继续保留。`
      : `删除结果组“${group.name}”？`
    if (!window.confirm(message)) return
    try {
      if (dirty && (group.saved || group.id !== activeId)) await saveDraft()
      if (group.saved) await updateGroup(group.id, { active: false })
      else await deleteGroup(group.id)
      await reloadGroups(group.id === activeId ? undefined : activeId ?? undefined)
    }
    catch (error) { setMessage({ text: errorMessage(error), error: true }) }
  }

  const clearGroups = async () => {
    if (!activeGroups.length || !window.confirm(`清除当前 ${activeGroups.length} 个结果组？已保存的筛选条件会继续保留。`)) return
    try {
      if (dirty && draft?.saved) await saveDraft()
      const count = await clearCurrentGroups(file.id)
      await reloadGroups()
      setMessage({ text: `已清除 ${count} 个当前结果组，已保存的筛选条件不受影响。` })
    } catch (error) { setMessage({ text: errorMessage(error), error: true }) }
  }

  const deleteSavedCondition = async (group: FilterGroup) => {
    if (!window.confirm(`永久删除筛选条件“${group.name}”？此操作只会在你确认后执行。`)) return
    try {
      await deleteGroup(group.id)
      await reloadGroups(group.id === activeId ? undefined : activeId ?? undefined)
      setMessage({ text: '已手动删除筛选条件。' })
    } catch (error) { setMessage({ text: errorMessage(error), error: true }) }
  }

  const moveGroup = async (index: number, direction: -1 | 1) => {
    const target = index + direction
    if (target < 0 || target >= activeGroups.length) return
    try {
      if (dirty) await saveDraft()
      const reordered = [...activeGroups]
      ;[reordered[index], reordered[target]] = [reordered[target], reordered[index]]
      await Promise.all(reordered.map((group, groupIndex) => updateGroup(group.id, { priority: groupIndex + 1 })))
      await reloadGroups(activeId ?? undefined)
    } catch (error) { setMessage({ text: errorMessage(error), error: true }) }
  }

  const execute = async () => {
    if (!activeGroups.length) { setMessage({ text: '请至少创建一个筛选结果组。', error: true }); return }
    setBusy(true)
    try {
      if (dirty) await saveDraft()
      const started = await startFilterRun(file.id, activeGroups.map((group) => group.id))
      setRun(started)
      setResultPage(null)
      const finished = await pollRun(() => getFilterRun(started.id), setRun)
      const firstKey = finished.summary?.[0]?.key ?? null
      setSelectedResult(firstKey)
      if (firstKey) setResultPage(await getFilterPage(finished.id, firstKey))
      onResultChanged()
      setMessage({ text: `筛选完成，共分配 ${formatNumber(finished.assignedRowCount ?? 0)} 行数据。` })
    } catch (error) { setMessage({ text: errorMessage(error), error: true }) }
    finally { setBusy(false) }
  }

  const changeSplitScope = (value: string) => {
    setSplitScope(value)
    const firstColumn = value === 'ALL'
      ? file.metadata.unifiedColumns[0]
      : file.metadata.sheets.find((sheet) => sheet.id === value)?.columns[0]
    setSplitColumnKey(firstColumn?.key ?? '')
  }

  const executeColumnSplit = async () => {
    const column = splitColumns.find((item) => item.key === splitColumnKey)
    if (!column) { setMessage({ text: '请选择需要分类的表头。', error: true }); return }
    setBusy(true)
    try {
      const started = await startColumnSplitRun(file.id, {
        columnRef: {
          headerName: column.headerName,
          occurrence: column.occurrence,
          fallbackColumnIndex: column.fallbackColumnIndex,
        },
        sourceScope: splitScope === 'ALL'
          ? { mode: 'ALL', sheetId: null }
          : { mode: 'SINGLE', sheetId: splitScope },
      })
      setRun(started)
      setResultPage(null)
      const finished = await pollRun(() => getFilterRun(started.id), setRun)
      const firstKey = finished.summary?.[0]?.key ?? null
      setSelectedResult(firstKey)
      if (firstKey) setResultPage(await getFilterPage(finished.id, firstKey))
      onResultChanged()
      const categoryCount = finished.summary?.filter((item) => !item.isSystem).length ?? 0
      setMessage({ text: `一键分类完成，识别出 ${categoryCount} 类数据，共处理 ${formatNumber(finished.assignedRowCount ?? 0)} 行。` })
    } catch (error) { setMessage({ text: errorMessage(error), error: true }) }
    finally { setBusy(false) }
  }

  const executeAffixSplit = async () => {
    if (splitMethod === 'OVERALL') { await executeColumnSplit(); return }
    const column = splitColumns.find((item) => item.key === splitColumnKey)
    if (!column) { setMessage({ text: '请选择需要识别前后缀的表头。', error: true }); return }
    if (!affixDelimiter) { setMessage({ text: '请输入连接符。', error: true }); return }
    setBusy(true)
    try {
      const started = await startAffixSplitRun(file.id, {
        columnRef: {
          headerName: column.headerName,
          occurrence: column.occurrence,
          fallbackColumnIndex: column.fallbackColumnIndex,
        },
        sourceScope: splitScope === 'ALL'
          ? { mode: 'ALL', sheetId: null }
          : { mode: 'SINGLE', sheetId: splitScope },
        kind: splitMethod,
        delimiter: affixDelimiter,
      })
      setRun(started)
      setResultPage(null)
      const finished = await pollRun(() => getFilterRun(started.id), setRun)
      const firstKey = finished.summary?.[0]?.key ?? null
      setSelectedResult(firstKey)
      if (firstKey) setResultPage(await getFilterPage(finished.id, firstKey))
      onResultChanged()
      const categoryCount = finished.summary?.filter((item) => !item.isSystem).length ?? 0
      const kindName = splitMethod === 'PREFIX' ? '前缀' : '后缀'
      setMessage({ text: `${kindName}分类完成，识别出 ${categoryCount} 类数据，共处理 ${formatNumber(finished.assignedRowCount ?? 0)} 行。` })
    } catch (error) { setMessage({ text: errorMessage(error), error: true }) }
    finally { setBusy(false) }
  }

  const restoreRun = async (target: FilterRun) => {
    if (target.status !== 'succeeded' || !target.summary?.length) return
    setRun(target)
    setResultPage(null)
    const firstKey = target.summary[0].key
    setSelectedResult(firstKey)
    try { setResultPage(await getFilterPage(target.id, firstKey)) }
    catch (error) { setMessage({ text: errorMessage(error), error: true }) }
  }

  const removeRun = async (target: FilterRun) => {
    if (!window.confirm(`删除 ${formatRunTime(target.createdAt)} 的筛选记录及导出文件？`)) return
    try {
      await deleteFilterRun(target.id)
      const remaining = history.filter((item) => item.id !== target.id)
      setHistory(remaining)
      if (run?.id === target.id) {
        const latest = remaining.find((item) => item.status === 'succeeded' && item.summary?.length)
        if (latest) await restoreRun(latest)
        else { setRun(null); setResultPage(null); setSelectedResult(null) }
      }
      onResultChanged()
      setMessage({ text: '筛选记录已删除。' })
    } catch (error) { setMessage({ text: errorMessage(error), error: true }) }
  }

  const clearHistory = async () => {
    const removable = history.filter((item) => !['queued', 'running'].includes(item.status))
    if (!removable.length || !window.confirm(`清空全部 ${removable.length} 条已完成的筛选记录？`)) return
    try {
      await Promise.all(removable.map((item) => deleteFilterRun(item.id)))
      const remaining = history.filter((item) => !removable.some((removed) => removed.id === item.id))
      setHistory(remaining)
      if (run && removable.some((item) => item.id === run.id)) {
        setRun(null); setResultPage(null); setSelectedResult(null)
      }
      onResultChanged()
      setMessage({ text: '已清空筛选处理记录。' })
    } catch (error) { setMessage({ text: errorMessage(error), error: true }) }
  }

  const loadResult = async (key: string, page = 1) => {
    if (!run || run.status !== 'succeeded') return
    setSelectedResult(key)
    try { setResultPage(await getFilterPage(run.id, key, page)) }
    catch (error) { setMessage({ text: errorMessage(error), error: true }) }
  }

  const exportResult = async () => {
    if (!run || run.status !== 'succeeded' || exporting) return
    setExporting(true)
    setMessage({ text: '正在生成 Excel，结果 Sheet 较多时可能需要几分钟，请勿重复点击。' })
    try {
      const url = `${API_BASE}${await createFilterExport(run.id)}`
      const link = document.createElement('a')
      link.href = url
      link.download = ''
      document.body.appendChild(link)
      link.click()
      link.remove()
      setMessage({ text: 'Excel 已生成，浏览器正在下载。' })
    } catch (error) { setMessage({ text: `Excel 导出失败：${errorMessage(error)}`, error: true }) }
    finally { setExporting(false) }
  }

  const persistRun = async () => {
    if (!run || run.status !== 'succeeded' || run.saved) return
    try {
      const saved = await saveFilterRun(run.id)
      setRun(saved)
      setHistory((items) => [saved, ...items.filter((item) => item.id !== saved.id)])
      onResultChanged()
      setMessage({ text: '筛选结果已保存到处理记录。' })
    } catch (error) { setMessage({ text: errorMessage(error), error: true }) }
  }

  return (
    <div className="page-stack filter-page">
      <PageHeader eyebrow="多条件分流" title="筛选处理" description="结果组按顺序匹配；一行数据只进入首个命中的 Sheet。">
        <button className="secondary-button" disabled={!draft || busy || saving || (!dirty && Boolean(draft.saved))} onClick={() => void saveDraft(true).then(() => setMessage({ text: draft?.saved ? '筛选条件修改已保存。' : '筛选条件已持久化保存。' })).catch((error) => setMessage({ text: errorMessage(error), error: true }))}>{saving ? <LoaderCircle className="spin" size={16} /> : <Save size={16} />} {saving ? '保存中' : draft?.saved ? '保存修改' : '保存条件'}</button>
        <button className="primary-button" disabled={!activeGroups.length || busy} onClick={() => void execute()}>{busy ? <LoaderCircle className="spin" size={17} /> : <Filter size={17} />} 开始筛选</button>
      </PageHeader>

      {run && run.status !== 'succeeded' && <ProgressBar label={run.status === 'failed' ? '处理失败' : '正在读取并分配数据'} progress={run.progress} />}

      <section className="panel quick-split-panel">
        <div className="panel-header"><div><h2>一键分类处理</h2><p>可按完整数据、前缀或后缀识别类别，每种数据分别生成一个结果 Sheet。</p></div><span className="saved-count">自动分类</span></div>
        <div className="quick-split-form affix-split-form">
          <div className="field quick-split-scope"><label>数据范围</label><select value={splitScope} disabled={busy} onChange={(event) => changeSplitScope(event.target.value)}><option value="ALL">全部 Sheet</option>{file.metadata.sheets.map((sheet) => <option value={sheet.id} key={sheet.id}>{sheet.name}</option>)}</select></div>
          <div className="field grow"><label>识别表头</label><select value={splitColumnKey} disabled={busy || !splitColumns.length} onChange={(event) => setSplitColumnKey(event.target.value)}>{splitColumns.map((column) => <option value={column.key} key={column.key}>{column.displayName}{splitScope === 'ALL' ? ` · 覆盖 ${column.coverage}/${column.totalSheets} Sheet` : ''}</option>)}</select></div>
          <div className="field affix-mode-field"><label>识别方式</label><select value={splitMethod} disabled={busy} onChange={(event) => setSplitMethod(event.target.value as 'OVERALL' | 'PREFIX' | 'SUFFIX')}><option value="OVERALL">按整体</option><option value="PREFIX">按前缀</option><option value="SUFFIX">按后缀</option></select></div>
          <div className="field delimiter-field"><label>连接符</label><input value={affixDelimiter} maxLength={5} disabled={busy || splitMethod === 'OVERALL'} placeholder={splitMethod === 'OVERALL' ? '无需填写' : '例如 -'} onChange={(event) => setAffixDelimiter(event.target.value)} /></div>
          <button className="primary-button large" disabled={busy || !splitColumnKey || (splitMethod !== 'OVERALL' && !affixDelimiter)} onClick={() => void executeAffixSplit()}>{busy ? <LoaderCircle className="spin" size={18} /> : <Filter size={18} />} 一键分类</button>
        </div>
        <div className="quick-split-note">按整体时直接识别该列全部不同数据，空值归入“空白”；连接符为“-”时，<strong>华东-上海-001</strong> 的前缀为“华东-”，后缀为“-001”。</div>
      </section>

      <div className="workspace-grid">
        <section className="panel group-sidebar">
          <div className="panel-header compact"><div><h2>结果组</h2><p>越靠上优先级越高</p></div><div className="group-header-actions"><button className="group-clear-button" disabled={!activeGroups.length} title="清除当前结果组" onClick={() => void clearGroups()}><Trash2 size={14} /> 清除</button><button className="icon-button primary" title="新增结果组" onClick={() => void addGroup()}><Plus size={17} /></button></div></div>
          {busy && !activeGroups.length ? <LoadingBlock /> : activeGroups.length === 0 ? <EmptyState compact title="还没有结果组" text="新增结果组，或从下方已保存条件中重新加入。" /> : (
            <div className="group-list">{activeGroups.map((group, index) => (
              <button key={group.id} className={`group-card ${group.id === activeId ? 'active' : ''}`} onClick={() => void selectGroup(group.id)}>
                <span className="priority">{index + 1}</span><span className="group-card-copy"><strong>{group.name}</strong><small>{group.sourceScope.mode === 'ALL' ? '全部 Sheet' : file.metadata.sheets.find((sheet) => sheet.id === group.sourceScope.sheetId)?.name ?? '单个 Sheet'} · {countConditions(group.conditionTree)} 个条件{group.saved ? ' · 已保存' : ''}</small></span>
                <span className="row-actions"><span role="button" title="上移" onClick={(event) => { event.stopPropagation(); void moveGroup(index, -1) }}><ArrowUp size={14} /></span><span role="button" title="下移" onClick={(event) => { event.stopPropagation(); void moveGroup(index, 1) }}><ArrowDown size={14} /></span><span role="button" title="删除" onClick={(event) => { event.stopPropagation(); void removeGroup(group) }}><Trash2 size={14} /></span></span>
              </button>
            ))}<div className="unmatched-card"><span className="priority">末</span><span><strong>未匹配</strong><small>系统固定 Sheet</small></span></div></div>
          )}
        </section>

        <section className="panel condition-editor-panel" id="filter-condition-editor">
          {!draft ? <EmptyState title="选择或新增结果组" text="在左侧选择一个结果组后编辑筛选规则。" /> : <>
            <div className="editor-heading"><div className="field grow"><label>结果 Sheet 名称</label><input maxLength={31} value={draft.name} onChange={(event) => { setDraft({ ...draft, name: event.target.value }); setDirty(true) }} /></div><div className="field scope-field"><label>处理范围</label><select value={draft.sourceScope.mode === 'ALL' ? 'ALL' : draft.sourceScope.sheetId ?? ''} onChange={(event) => { const value = event.target.value; setDraft({ ...draft, sourceScope: value === 'ALL' ? { mode: 'ALL', sheetId: null } : { mode: 'SINGLE', sheetId: value } }); setDirty(true) }}><option value="ALL">全部 Sheet</option>{file.metadata.sheets.map((sheet) => <option value={sheet.id} key={sheet.id}>{sheet.name}</option>)}</select></div></div>
            <div className="section-label"><span>筛选条件</span><small>支持嵌套 AND / OR，最多 5 层</small></div>
            <div className="condition-help"><span><strong>添加条件</strong>在当前 AND/OR 层增加一条判断</span><span><strong>添加条件组</strong>增加一层可独立设置 AND/OR 的括号逻辑</span></div>
            <ConditionTreeEditor
              value={draft.conditionTree}
              columns={file.metadata.unifiedColumns}
              fileId={file.id}
              sourceScope={draft.sourceScope}
              onChange={(conditionTree) => { setDraft({ ...draft, conditionTree }); setDirty(true) }}
            />
            <div className="logic-note">当前结果组命中后，该行不会再进入后续结果组；未命中的数据统一进入“未匹配”。</div>
          </>}
        </section>
      </div>

      <section className="panel saved-groups-panel">
        <div className="panel-header"><div><h2>全局筛选条件</h2><p>所有上传的工作簿均可加入使用；只有手动删除才会永久移除。</p></div><span className="saved-count">{savedGroups.length} 个方案</span></div>
        {savedGroups.length ? <div className="saved-group-list">{savedGroups.map((group) => (
          <button className={`saved-group-card ${activeId === group.id ? 'active' : ''}`} key={group.id} onClick={() => void editSavedGroup(group.id)}>
            <span className="saved-group-icon"><Layers3 size={18} /></span>
            <span className="saved-group-copy"><strong>{group.name}</strong><small>{countConditions(group.conditionTree)} 个条件 · 来源：{group.sourceFileName ?? '已删除工作簿'} · {formatRunTime(group.savedAt ?? group.updatedAt)}</small></span>
            <span className="saved-edit"><Pencil size={14} /> {group.fileId === file.id && group.active !== false ? '修改' : '加入使用'}</span>
            <span role="button" className="saved-delete" title="手动删除筛选条件" onClick={(event) => { event.stopPropagation(); void deleteSavedCondition(group) }}><Trash2 size={14} /></span>
          </button>
        ))}</div> : <EmptyState compact title="还没有保存的条件" text="在上方配置结果组后点击“保存条件”，方案会显示在这里。" />}
      </section>

      <section className="panel history-panel">
        <div className="panel-header"><div><h2>已保存的处理记录</h2><p>筛选完成后点击“保存处理记录”才会加入这里；记录仅由你手动清理。</p></div><button className="secondary-button danger-outline" disabled={!history.some((item) => !['queued', 'running'].includes(item.status))} onClick={() => void clearHistory()}><Trash2 size={16} /> 清空记录</button></div>
        {history.length === 0 ? <EmptyState compact title="暂无已保存记录" text="筛选完成后，可在结果区域手动保存本次处理记录。" /> : <div className="run-history-list">{history.map((item) => (
          <div className={`run-record ${run?.id === item.id ? 'active' : ''}`} key={item.id}>
            <button className="run-record-main" disabled={item.status !== 'succeeded'} onClick={() => void restoreRun(item)}>
              <span className={`run-status-dot ${item.status}`} />
              <span><strong>{formatRunTime(item.createdAt)}</strong><small>{runSummaryText(item)}</small></span>
              <span className={`status-pill ${item.status}`}>{runStatusText(item.status)}</span>
            </button>
            <button className="icon-button danger" disabled={['queued', 'running'].includes(item.status)} title="删除这条处理记录" onClick={() => void removeRun(item)}><Trash2 size={15} /></button>
          </div>
        ))}</div>}
      </section>

      {run?.status === 'succeeded' && run.summary && (
        <section className="panel results-panel">
          <div className="panel-header"><div><h2>筛选结果</h2><p>{run.rowCountConserved ? '行数校验通过' : '请检查行数'}：源数据 {formatNumber(run.sourceRowCount ?? 0)} 行，输出 {formatNumber(run.assignedRowCount ?? 0)} 行{!run.saved && '。当前为临时结果，请手动保存'}</p></div><div className="header-actions">{!run.saved && <button className="secondary-button" onClick={() => void persistRun()}><Save size={16} /> 保存处理记录</button>}<button className="primary-button" disabled={exporting} onClick={() => void exportResult()}>{exporting ? <LoaderCircle className="spin" size={16} /> : <Download size={16} />} {exporting ? '正在生成 Excel' : '下载 Excel'}</button></div></div>
          <div className="sheet-tabs">{run.summary.map((sheet) => <button className={selectedResult === sheet.key ? 'active' : ''} key={sheet.key} onClick={() => void loadResult(sheet.key)}>{sheet.name}<span>{formatNumber(sheet.rowCount)}</span></button>)}</div>
          {resultPage ? <DataTable page={resultPage} onPage={(page) => selectedResult && void loadResult(selectedResult, page)} /> : selectedSheet && <EmptyState compact title={selectedSheet.name} text="此结果中没有数据。" />}
        </section>
      )}
    </div>
  )
}

function SumWorkspace({ preferredFileId, active, refreshToken, setMessage }: { preferredFileId: string; active: boolean; refreshToken: number; setMessage: MessageSetter }) {
  const [sources, setSources] = useState<SumSource[]>([])
  const [sourceKey, setSourceKey] = useState('')
  const [loadingSources, setLoadingSources] = useState(false)
  const [mode, setMode] = useState<'NORMALIZED_HEADER' | 'FIXED_POSITION'>('NORMALIZED_HEADER')
  const [header, setHeader] = useState('')
  const [columnIndex, setColumnIndex] = useState(1)
  const [run, setRun] = useState<SumRun | null>(null)
  const [busy, setBusy] = useState(false)
  const [details, setDetails] = useState<InvalidDetail[] | null>(null)
  const selectedSource = sources.find((source) => sumSourceKey(source) === sourceKey) ?? null

  useEffect(() => {
    if (!active) return
    let cancelled = false
    setLoadingSources(true)
    getSumSources().then((items) => {
      if (cancelled) return
      setSources(items)
      setSourceKey((current) => {
        const next = items.find((item) => sumSourceKey(item) === current)
          ?? items.find((item) => item.type === 'FILE' && item.fileId === preferredFileId)
          ?? items[0]
        setHeader(next?.metadata.unifiedColumns[0]?.headerName ?? '')
        return next ? sumSourceKey(next) : ''
      })
    }).catch((error) => setMessage({ text: errorMessage(error), error: true })).finally(() => {
      if (!cancelled) setLoadingSources(false)
    })
    return () => { cancelled = true }
  }, [active, preferredFileId, refreshToken, setMessage])

  const selectSource = (source: SumSource) => {
    setSourceKey(sumSourceKey(source))
    setHeader(source.metadata.unifiedColumns[0]?.headerName ?? '')
    setColumnIndex(1)
    setRun(null)
    setDetails(null)
  }

  const execute = async () => {
    if (!selectedSource) { setMessage({ text: '请先选择需要求和的数据表。', error: true }); return }
    setBusy(true)
    setDetails(null)
    try {
      const selector = mode === 'NORMALIZED_HEADER' ? { matchMode: mode, headerName: header } : { matchMode: mode, columnIndex }
      const started = await startSumRun(selectedSource, selector)
      setRun(started)
      const finished = await pollRun(() => getSumRun(started.id), setRun)
      setMessage({ text: `求和完成，工作簿合计 ${finished.grandTotal ?? '0'}。` })
    } catch (error) { setMessage({ text: errorMessage(error), error: true }) }
    finally { setBusy(false) }
  }

  const showDetails = async () => {
    if (!run || run.status !== 'succeeded') return
    try { setDetails(await getInvalidDetails(run.id)) }
    catch (error) { setMessage({ text: errorMessage(error), error: true }) }
  }

  const exportResult = async () => {
    if (!run || run.status !== 'succeeded') return
    try { window.location.href = `${API_BASE}${await createSumExport(run.id)}` }
    catch (error) { setMessage({ text: errorMessage(error), error: true }) }
  }

  return (
    <div className="page-stack">
      <PageHeader eyebrow="全工作簿汇总" title="跨 Sheet 求和" description="先选择一个原始工作簿或筛选结果，再指定需要汇总的列。" />
      <section className="panel sum-source-panel">
        <div className="panel-header"><div><h2>选择数据表</h2><p>这里包含全部上传工作簿，以及已保存和未保存的筛选结果。</p></div><span className="saved-count">{sources.length} 个可用数据表</span></div>
        {loadingSources ? <LoadingBlock /> : sources.length ? <div className="sum-source-list">{sources.map((source) => (
          <button className={`sum-source-card ${sumSourceKey(source) === sourceKey ? 'active' : ''}`} key={sumSourceKey(source)} onClick={() => selectSource(source)}>
            <span className={`sum-source-icon ${source.type === 'FILTER_RESULT' ? 'filtered' : ''}`}>{source.type === 'FILTER_RESULT' ? <Filter size={18} /> : <FileSpreadsheet size={18} />}</span>
            <span className="sum-source-copy"><strong title={source.name}>{source.name}</strong><small>{source.metadata.sheetCount} 个 Sheet · {formatNumber(source.metadata.totalRows)} 行</small></span>
            <span className={`source-kind ${source.type === 'FILTER_RESULT' ? 'filtered' : ''}`}>{source.type === 'FILE' ? '原始工作簿' : source.saved ? '已保存筛选结果' : '临时筛选结果'}</span>
          </button>
        ))}</div> : <EmptyState compact title="暂无可用数据表" text="请先上传工作簿或完成一次筛选处理。" />}
      </section>
      <section className="panel sum-config">
        <div className="panel-header"><div><h2>选择目标列</h2><p>{selectedSource ? `当前数据表：${selectedSource.name}` : '请先在上方选择数据表'}。货币、千分位、百分号和括号会被剥离后解析。</p></div></div>
        <div className="mode-cards">
          <button className={mode === 'NORMALIZED_HEADER' ? 'active' : ''} onClick={() => setMode('NORMALIZED_HEADER')}><span className="radio-dot" /><strong>按表头名称匹配</strong><small>忽略前后空格、全半角与大小写差异</small></button>
          <button className={mode === 'FIXED_POSITION' ? 'active' : ''} onClick={() => setMode('FIXED_POSITION')}><span className="radio-dot" /><strong>按固定列位置匹配</strong><small>每个 Sheet 读取相同的第 N 列</small></button>
        </div>
        <div className="sum-form">
          {mode === 'NORMALIZED_HEADER' ? <div className="field grow"><label>表头名称</label><select value={header} disabled={!selectedSource} onChange={(event) => setHeader(event.target.value)}>{uniqueHeaders(selectedSource?.metadata.unifiedColumns ?? []).map((column) => <option value={column.headerName} key={column.key}>{column.displayName} · 覆盖 {column.coverage}/{column.totalSheets} Sheet</option>)}</select></div> : <div className="field grow"><label>列位置</label><div className="input-with-prefix"><span>第</span><input type="number" min={1} max={16384} disabled={!selectedSource} value={columnIndex} onChange={(event) => setColumnIndex(Math.max(1, Number(event.target.value)))} /><span>列</span></div></div>}
          <button className="primary-button large" disabled={!selectedSource || busy || (mode === 'NORMALIZED_HEADER' && !header)} onClick={() => void execute()}>{busy ? <LoaderCircle className="spin" size={18} /> : <Calculator size={18} />} 开始求和</button>
        </div>
        <div className="parse-examples"><span><code>$1,234.50</code> → 1234.50</span><span><code>12%</code> → 12</span><span><code>(100)</code> → 100</span><span><code>-100</code> → -100</span><span className="invalid"><code>约12元</code> → 不解析</span></div>
      </section>

      {run && run.status !== 'succeeded' && <ProgressBar label={run.status === 'failed' ? '处理失败' : '正在扫描全部 Sheet'} progress={run.progress} />}

      {run?.status === 'succeeded' && run.results && (
        <section className="panel results-panel">
          <div className="panel-header"><div><h2>求和结果</h2><p>共处理 {run.results.length} 个 Sheet，发现 {formatNumber(run.invalidCount ?? 0)} 个非数值单元格。</p></div><div className="header-actions"><button className="secondary-button" disabled={!run.invalidCount} onClick={() => void showDetails()}><Search size={16} /> 查看异常明细</button><button className="primary-button" onClick={() => void exportResult()}><Download size={16} /> 下载 Excel</button></div></div>
          <div className="grand-total"><span>工作簿总计</span><strong>{run.grandTotal ?? '0'}</strong></div>
          <div className="table-scroll"><table><thead><tr><th>Sheet 名称</th><th>匹配列</th><th>数据行</th><th>有效数字</th><th>文本转数字</th><th>空值</th><th>非数值</th><th>求和结果</th><th>状态</th></tr></thead><tbody>{run.results.map((item) => <tr key={item.sheetId}><td className="strong-cell">{item.sheetName}</td><td>{item.columnName || '—'}</td><td>{formatNumber(item.dataRowCount)}</td><td>{formatNumber(item.numericCount)}</td><td>{formatNumber(item.textConvertedCount)}</td><td>{formatNumber(item.emptyCount)}</td><td>{formatNumber(item.nonNumericCount)}</td><td className="number-cell">{item.sum ?? '—'}</td><td><span className={`status-pill ${item.status.toLowerCase()}`}>{item.status.toLowerCase() === 'succeeded' ? '完成' : '跳过'}</span></td></tr>)}</tbody></table></div>
        </section>
      )}

      {details && <Modal title={`异常明细（${details.length}）`} onClose={() => setDetails(null)}><div className="table-scroll modal-table"><table><thead><tr><th>Sheet</th><th>源行号</th><th>单元格</th><th>列名</th><th>原始内容</th><th>原因</th></tr></thead><tbody>{details.map((item, index) => <tr key={`${item.sheetName}-${item.cell}-${index}`}><td>{item.sheetName}</td><td>{item.sourceRow}</td><td>{item.cell}</td><td>{item.columnName}</td><td>{String(item.originalValue ?? '')}</td><td>{item.reason}</td></tr>)}</tbody></table></div></Modal>}
    </div>
  )
}

function PageHeader({ eyebrow, title, description, children }: { eyebrow: string; title: string; description: string; children?: React.ReactNode }) {
  return <div className="page-header"><div><span className="eyebrow">{eyebrow}</span><h1>{title}</h1><p>{description}</p></div>{children && <div className="header-actions">{children}</div>}</div>
}

function StatCard({ label, value, suffix }: { label: string; value: string | number; suffix: string }) {
  return <div className="stat-card"><span>{label}</span><strong>{value}</strong><small>{suffix}</small></div>
}

function DataTable({ page, onPage }: { page: PageResult; onPage?: (page: number) => void }) {
  const pageCount = Math.max(1, Math.ceil(page.total / page.pageSize))
  if (page.rows.length === 0) return <EmptyState compact title="暂无数据" text="当前 Sheet 没有可显示的数据行。" />
  return <><div className="table-scroll"><table><thead><tr><th className="row-number">#</th>{page.headers.map((header, index) => <th key={`${header}-${index}`}>{header}</th>)}</tr></thead><tbody>{page.rows.map((row, rowIndex) => <tr key={rowIndex}><td className="row-number">{(page.page - 1) * page.pageSize + rowIndex + 1}</td>{page.headers.map((_, cellIndex) => <td key={cellIndex} title={displayCell(row[cellIndex])}>{displayCell(row[cellIndex])}</td>)}</tr>)}</tbody></table></div><div className="pagination"><span>共 {formatNumber(page.total)} 行</span><div><button disabled={!onPage || page.page <= 1} onClick={() => onPage?.(page.page - 1)}><ChevronLeft size={15} /></button><span>{page.page} / {pageCount}</span><button disabled={!onPage || page.page >= pageCount} onClick={() => onPage?.(page.page + 1)}><ChevronRight size={15} /></button></div></div></>
}

function ProgressBar({ label, progress }: { label: string; progress: number }) {
  return <div className="progress-card"><div><LoaderCircle className="spin" size={18} /><strong>{label}</strong><span>{progress}%</span></div><div className="progress-track"><span style={{ width: `${progress}%` }} /></div></div>
}

function LoadingBlock() { return <div className="loading-block"><LoaderCircle className="spin" /> 正在加载…</div> }

function EmptyState({ title, text, compact }: { title: string; text: string; compact?: boolean }) {
  return <div className={`empty-state ${compact ? 'compact' : ''}`}><FileSpreadsheet /><strong>{title}</strong><p>{text}</p></div>
}

function Modal({ title, onClose, children }: { title: string; onClose: () => void; children: React.ReactNode }) {
  return <div className="modal-backdrop" onMouseDown={onClose}><div className="modal" onMouseDown={(event) => event.stopPropagation()}><div className="modal-header"><h2>{title}</h2><button className="icon-button" onClick={onClose}><X size={18} /></button></div>{children}</div></div>
}

function displayCell(value: CellValue | undefined) {
  if (value === null || value === undefined || value === '') return '—'
  return String(value)
}

function errorMessage(error: unknown) { return error instanceof Error ? error.message : String(error) }

function formatRunTime(value: string) {
  const date = new Date(value)
  if (Number.isNaN(date.getTime())) return '未知时间'
  return new Intl.DateTimeFormat('zh-CN', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false }).format(date)
}

function runStatusText(status: FilterRun['status']) {
  return { queued: '等待中', running: '处理中', succeeded: '已完成', failed: '失败' }[status]
}

function runSummaryText(run: FilterRun) {
  if (run.status === 'failed') return run.error || '筛选失败'
  if (run.status !== 'succeeded') return `处理进度 ${run.progress}%`
  return `${run.summary?.length ?? 0} 个结果 Sheet · ${formatNumber(run.assignedRowCount ?? 0)} 行`
}

function countConditions(group: FilterGroup['conditionTree']): number {
  return group.children.reduce((count, child) => count + (child.type === 'condition' ? 1 : countConditions(child)), 0)
}

function uniqueHeaders(columns: UnifiedColumn[]) {
  const seen = new Set<string>()
  return columns.filter((column) => {
    if (seen.has(column.normalizedName)) return false
    seen.add(column.normalizedName)
    return true
  })
}

function nextDefaultSheetName(groups: FilterGroup[]) {
  const names = new Set(groups.map((group) => group.name.toLocaleLowerCase()))
  let sequence = 1
  while (names.has(`sheet${sequence}`)) sequence += 1
  return `Sheet${sequence}`
}

function sumSourceKey(source: Pick<SumSource, 'id' | 'type'>) {
  return `${source.type}:${source.id}`
}
