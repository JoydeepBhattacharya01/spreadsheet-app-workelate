import { useState, useRef, useCallback, useMemo, useEffect } from 'react'
import './App.css'
import { createEngine } from './engine/core.js'

const TOTAL_ROWS = 50
const TOTAL_COLS = 50
const STORAGE_KEY = 'spreadsheet_state_v1'

export default function App() {
  // Engine instance is created once and reused across renders
  // Note: The engine maintains its own internal state, so React state is only used for UI updates
  const [engine] = useState(() => createEngine(TOTAL_ROWS, TOTAL_COLS))
  const [version, setVersion] = useState(0)
  const [persistTick, setPersistTick] = useState(0)
  const [selectedCell, setSelectedCell] = useState(null)
  const [editingCell, setEditingCell] = useState(null)
  const [editValue, setEditValue] = useState('')
  const [selectionAnchor, setSelectionAnchor] = useState(null)
  const [selectionRange, setSelectionRange] = useState(null) // inclusive rectangle: { r1,c1,r2,c2 }
  const copyBufferRef = useRef('')

  // View-layer only sorting and filtering
  const [sortConfig, setSortConfig] = useState({ column: null, order: null }) // { column: "A", order: "asc"|"desc"|null }
  const [filters, setFilters] = useState({}) // { A: ["Apple","Banana"] }
  // Cell styles are stored separately from engine data
  // Format: { "row,col": { bold: bool, italic: bool, ... } }
  const [cellStyles, setCellStyles] = useState({})
  const cellInputRef = useRef(null)

  const forceRerender = useCallback(() => {
    setVersion(v => v + 1)
    setPersistTick(t => t + 1)
  }, [])

  // ────── Cell style helpers ──────

  const getCellStyle = useCallback((row, col) => {
    const key = `${row},${col}`
    return cellStyles[key] || {
      bold: false, italic: false, underline: false,
      bg: 'white', color: '#202124', align: 'left', fontSize: 13
    }
  }, [cellStyles])

  const updateCellStyle = useCallback((row, col, updates) => {
    const key = `${row},${col}`
    setCellStyles(prev => ({
      ...prev,
      [key]: { ...getCellStyle(row, col), ...updates }
    }))
    // Ensure style-only changes also get persisted (debounced effect listens to `persistTick`).
    setPersistTick(t => t + 1)
  }, [getCellStyle])

  // ────── Cell editing ──────

  const startEditing = useCallback((row, col, { extendSelection = false } = {}) => {
    setSelectedCell({ r: row, c: col })
    setEditingCell({ r: row, c: col })
    setSelectionRange(() => {
      if (extendSelection && selectionAnchor) {
        return { r1: selectionAnchor.r, c1: selectionAnchor.c, r2: row, c2: col }
      }
      return { r1: row, c1: col, r2: row, c2: col }
    })
    setSelectionAnchor((prevAnchor) => {
      if (extendSelection && selectionAnchor) return prevAnchor
      return { r: row, c: col }
    })
    const cellData = engine.getCell(row, col)
    setEditValue(cellData.raw)
    setTimeout(() => cellInputRef.current?.focus(), 0)
  }, [engine, selectionAnchor])

  const commitEdit = useCallback((row, col) => {
    // Only commit if the value actually changed to avoid unnecessary recalculations
    const currentCell = engine.getCell(row, col)
    if (currentCell.raw !== editValue) {
      engine.setCell(row, col, editValue)
      forceRerender()
    }
    setEditingCell(null)
  }, [engine, editValue, forceRerender])

  const handleCellClick = useCallback((row, col, event) => {
    if (editingCell && (editingCell.r !== row || editingCell.c !== col)) {
      commitEdit(editingCell.r, editingCell.c)
    }
    if (!editingCell || editingCell.r !== row || editingCell.c !== col) {
      startEditing(row, col, { extendSelection: !!event?.shiftKey })
    }
  }, [editingCell, commitEdit, startEditing])

  // ────── Keyboard navigation ──────

  const handleKeyDown = useCallback((event, row, col) => {
    if (event.key === 'Enter') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.min(row + 1, engine.rows - 1), col, { extendSelection: !!event.shiftKey })
    } else if (event.key === 'Tab') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(row, Math.min(col + 1, engine.cols - 1), { extendSelection: !!event.shiftKey })
    } else if (event.key === 'Escape') {
      setEditValue(engine.getCell(row, col).raw)
      setEditingCell(null)
    } else if (event.key === 'ArrowDown') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.min(row + 1, engine.rows - 1), col, { extendSelection: !!event.shiftKey })
    } else if (event.key === 'ArrowUp') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(Math.max(row - 1, 0), col, { extendSelection: !!event.shiftKey })
    } else if (event.key === 'ArrowLeft') {
      event.preventDefault()
      commitEdit(row, col)
      if (col > 0) {
        startEditing(row, col - 1, { extendSelection: !!event.shiftKey })
      } else if (row > 0) {
        startEditing(row - 1, engine.cols - 1, { extendSelection: !!event.shiftKey })
      }
    } else if (event.key === 'ArrowRight') {
      event.preventDefault()
      commitEdit(row, col)
      startEditing(row, Math.min(col + 1, engine.cols - 1), { extendSelection: !!event.shiftKey })
    }
  }, [engine, commitEdit, startEditing])

  // ────── Formula bar handlers ──────

  const handleFormulaBarKeyDown = useCallback((event) => {
    if (!editingCell) return
    handleKeyDown(event, editingCell.r, editingCell.c)
  }, [editingCell, handleKeyDown])

  const handleFormulaBarFocus = useCallback(() => {
    if (selectedCell && !editingCell) {
      setEditingCell(selectedCell)
      setEditValue(engine.getCell(selectedCell.r, selectedCell.c).raw)
    }
  }, [selectedCell, editingCell, engine])

  const handleFormulaBarChange = useCallback((value) => {
    if (!editingCell && selectedCell) setEditingCell(selectedCell)
    setEditValue(value)
  }, [editingCell, selectedCell])

  // ────── Undo / Redo ──────

  const handleUndo = useCallback(() => { if (engine.undo()) forceRerender() }, [engine, forceRerender])
  const handleRedo = useCallback(() => { if (engine.redo()) forceRerender() }, [engine, forceRerender])

  const getSelectionRect = useCallback(() => {
    if (selectionRange) {
      const minR = Math.min(selectionRange.r1, selectionRange.r2)
      const maxR = Math.max(selectionRange.r1, selectionRange.r2)
      const minC = Math.min(selectionRange.c1, selectionRange.c2)
      const maxC = Math.max(selectionRange.c1, selectionRange.c2)
      return { minR, maxR, minC, maxC }
    }
    if (selectedCell) {
      return { minR: selectedCell.r, maxR: selectedCell.r, minC: selectedCell.c, maxC: selectedCell.c }
    }
    return null
  }, [selectionRange, selectedCell])

  // ───── Local storage persistence ─────

  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY)
      if (!raw) return
      const parsed = JSON.parse(raw)
      // Backward compatible:
      // - older versions stored the engine state directly
      // - newer versions store { v, engine, styles }
      const engineState = parsed && typeof parsed === 'object' && parsed.engine ? parsed.engine : parsed
      engine.importState(engineState)

      const styles = parsed && typeof parsed === 'object' && parsed.styles && typeof parsed.styles === 'object'
        ? parsed.styles
        : null
      if (styles) setCellStyles(styles)
      forceRerender()
    } catch {
      // Corrupted storage or incompatible data: start fresh.
    }
    // Intentionally only run once on mount.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [engine])

  useEffect(() => {
    const t = setTimeout(() => {
      try {
        const state = {
          v: 2,
          engine: engine.serializeState(),
          styles: cellStyles,
        }
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state))
      } catch {
        // Ignore quota/security/corrupted runtime issues.
      }
    }, 500)
    return () => clearTimeout(t)
  }, [persistTick, engine, cellStyles])

  // ───── Clipboard + keyboard shortcuts ─────
  useEffect(() => {
    const isFormulaBarFocused = () => {
      const el = document.activeElement
      return !!(el && el.classList && el.classList.contains('formula-bar-input'))
    }

    const isSpreadsheetContextFocused = () => {
      const el = document.activeElement
      if (isFormulaBarFocused()) return false
      // Only intercept paste/copy when the grid cell input is focused (or when no text/input is focused).
      if (el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.tagName === 'SELECT')) {
        return !!(el && el.classList && el.classList.contains('cell-input'))
      }
      return true
    }

    const copySelectionToText = () => {
      const rect = getSelectionRect()
      if (!rect) return ''
      const { minR, maxR, minC, maxC } = rect
      const lines = []
      for (let r = minR; r <= maxR; r++) {
        const cols = []
        for (let c = minC; c <= maxC; c++) {
          const cell = engine.getCell(r, c)
          if (cell.error) cols.push('ERROR')
          else cols.push(cell.computed === '' || cell.computed === null ? '' : String(cell.computed))
        }
        lines.push(cols.join('\t'))
      }
      return lines.join('\n')
    }

    const handleCopy = (e) => {
      if (!isSpreadsheetContextFocused()) return
      if (!getSelectionRect()) return
      e.preventDefault()
      const text = copySelectionToText()
      copyBufferRef.current = text
      try {
        e.clipboardData?.setData('text/plain', text)
      } catch {
        // Fallback: internal buffer already set.
      }
    }

    const handlePaste = (e) => {
      if (!isSpreadsheetContextFocused()) return
      if (!getSelectionRect()) return
      e.preventDefault()

      if (editingCell) {
        // Discard in-progress edit so undo focuses on the paste operation.
        setEditingCell(null)
      }

      const rect = getSelectionRect()
      const startRow = rect.minR
      const startCol = rect.minC

      const clipboardText = e.clipboardData?.getData('text/plain') || copyBufferRef.current
      if (!clipboardText) return

      const normalized = clipboardText.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
      let lines = normalized.split('\n')
      if (lines.length > 0 && lines[lines.length - 1] === '' && normalized.endsWith('\n')) lines.pop()
      const rowsData = lines.map(line => line.split('\t'))
      const maxCols = rowsData.reduce((m, parts) => Math.max(m, parts.length), 0)

      const changes = []
      for (let i = 0; i < rowsData.length; i++) {
        const parts = rowsData[i]
        const targetRow = startRow + i
        if (targetRow >= engine.rows) continue

        for (let j = 0; j < parts.length; j++) {
          const targetCol = startCol + j
          if (targetCol >= engine.cols) continue
          const nextRaw = parts[j]
          const prevRaw = engine.getCell(targetRow, targetCol).raw
          if (prevRaw !== nextRaw) changes.push({ r: targetRow, c: targetCol, value: nextRaw })
        }
      }

      if (changes.length > 0) {
        engine.setCellsBatch(changes)
        forceRerender()
      }

      // Update selection to the pasted rectangle.
      const endRow = Math.min(engine.rows - 1, startRow + rowsData.length - 1)
      const endCol = Math.min(engine.cols - 1, startCol + maxCols - 1)
      setSelectionAnchor({ r: startRow, c: startCol })
      setSelectionRange({ r1: startRow, c1: startCol, r2: endRow, c2: endCol })
      setSelectedCell({ r: endRow, c: endCol })
    }

    const handleKeyDown = (e) => {
      const mod = e.metaKey || e.ctrlKey
      if (!mod) return
      if (!isSpreadsheetContextFocused()) return

      if (e.key.toLowerCase() === 'z') {
        e.preventDefault()
        if (editingCell) commitEdit(editingCell.r, editingCell.c)
        if (engine.canUndo()) {
          engine.undo()
          forceRerender()
        }
      }
    }

    window.addEventListener('keydown', handleKeyDown)
    document.addEventListener('copy', handleCopy)
    document.addEventListener('paste', handlePaste)
    return () => {
      window.removeEventListener('keydown', handleKeyDown)
      document.removeEventListener('copy', handleCopy)
      document.removeEventListener('paste', handlePaste)
    }
  }, [engine, editingCell, getSelectionRect, commitEdit, forceRerender])

  // ────── Formatting toggles ──────

  const toggleBold = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { bold: !style.bold })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const toggleItalic = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { italic: !style.italic })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const toggleUnderline = useCallback(() => {
    if (!selectedCell) return
    const style = getCellStyle(selectedCell.r, selectedCell.c)
    updateCellStyle(selectedCell.r, selectedCell.c, { underline: !style.underline })
  }, [selectedCell, getCellStyle, updateCellStyle])

  const changeFontSize = useCallback((size) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { fontSize: size })
  }, [selectedCell, updateCellStyle])

  const changeAlignment = useCallback((align) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { align })
  }, [selectedCell, updateCellStyle])

  const changeFontColor = useCallback((color) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { color })
  }, [selectedCell, updateCellStyle])

  const changeBackgroundColor = useCallback((color) => {
    if (!selectedCell) return
    updateCellStyle(selectedCell.r, selectedCell.c, { bg: color })
  }, [selectedCell, updateCellStyle])

  // ────── Clear operations ──────

  const clearSelectedCell = useCallback(() => {
    if (!selectedCell) return
    engine.setCell(selectedCell.r, selectedCell.c, '')
    forceRerender()
    // Remove style entry for cleared cell
    // Note: This deletes the style object entirely - if you need to preserve default styles,
    // you may want to set them explicitly rather than deleting
    const key = `${selectedCell.r},${selectedCell.c}`
    setCellStyles(prev => { const next = { ...prev }; delete next[key]; return next })
    setEditValue('')
  }, [selectedCell, engine, forceRerender])

  const clearAllCells = useCallback(() => {
    for (let r = 0; r < engine.rows; r++) {
      for (let c = 0; c < engine.cols; c++) {
        engine.setCell(r, c, '')
      }
    }
    forceRerender()
    setCellStyles({})
    setSelectedCell(null)
    setEditingCell(null)
    setSelectionAnchor(null)
    setSelectionRange(null)
    setEditValue('')
  }, [engine, forceRerender])

  // ────── Row / Column operations ──────

  const insertRow = useCallback(() => {
    if (!selectedCell) return
    engine.insertRow(selectedCell.r)
    forceRerender()
    const next = { r: selectedCell.r + 1, c: selectedCell.c }
    setSelectedCell(next)
    setSelectionAnchor(next)
    setSelectionRange({ r1: next.r, c1: next.c, r2: next.r, c2: next.c })
  }, [selectedCell, engine, forceRerender])

  const deleteRow = useCallback(() => {
    if (!selectedCell) return
    engine.deleteRow(selectedCell.r)
    forceRerender()
    if (selectedCell.r >= engine.rows) {
      const next = { r: engine.rows - 1, c: selectedCell.c }
      setSelectedCell(next)
      setSelectionAnchor(next)
      setSelectionRange({ r1: next.r, c1: next.c, r2: next.r, c2: next.c })
    }
  }, [selectedCell, engine, forceRerender])

  const insertColumn = useCallback(() => {
    if (!selectedCell) return
    engine.insertColumn(selectedCell.c)
    forceRerender()
    const next = { r: selectedCell.r, c: selectedCell.c + 1 }
    setSelectedCell(next)
    setSelectionAnchor(next)
    setSelectionRange({ r1: next.r, c1: next.c, r2: next.r, c2: next.c })
  }, [selectedCell, engine, forceRerender])

  const deleteColumn = useCallback(() => {
    if (!selectedCell) return
    engine.deleteColumn(selectedCell.c)
    forceRerender()
    if (selectedCell.c >= engine.cols) {
      const next = { r: selectedCell.r, c: engine.cols - 1 }
      setSelectedCell(next)
      setSelectionAnchor(next)
      setSelectionRange({ r1: next.r, c1: next.c, r2: next.r, c2: next.c })
    }
  }, [selectedCell, engine, forceRerender])

  // ────── Derived state ──────

  const selectedCellStyle = useMemo(() => {
    return selectedCell ? getCellStyle(selectedCell.r, selectedCell.c) : null
  }, [selectedCell, getCellStyle])

  const getColumnLabel = useCallback((col) => {
    let label = ''
    let num = col + 1
    while (num > 0) {
      num--
      label = String.fromCharCode(65 + (num % 26)) + label
      num = Math.floor(num / 26)
    }
    return label
  }, [])

  const columnLetterToIndex = useCallback((columnLetter) => {
    // Supports multi-letter columns: A, Z, AA, AB, ...
    if (!columnLetter || typeof columnLetter !== 'string') return null
    let index = 0
    for (let i = 0; i < columnLetter.length; i++) {
      const ch = columnLetter[i]
      if (ch < 'A' || ch > 'Z') return null
      index = index * 26 + (ch.charCodeAt(0) - 64)
    }
    return index - 1
  }, [])

  const toggleSort = useCallback((colIndex) => {
    const colLetter = getColumnLabel(colIndex)
    setSortConfig((prev) => {
      if (prev.column !== colLetter) return { column: colLetter, order: 'asc' }
      if (prev.order === 'asc') return { column: colLetter, order: 'desc' }
      if (prev.order === 'desc') {
        // Cycle to "none"
        return { column: colLetter, order: null }
      }
      // If we were already in "none", cycle back to "asc"
      return { column: colLetter, order: 'asc' }
    })
  }, [getColumnLabel])

  const getCellDisplayString = useCallback((row, col) => {
    const cell = engine.getCell(row, col)
    if (cell.error) return 'ERROR'
    if (cell.computed === null || cell.computed === '') return ''
    return String(cell.computed)
  }, [engine])

  const columnUniqueValues = useMemo(() => {
    const map = {}
    map.__version = version
    for (let col = 0; col < engine.cols; col++) {
      const set = new Set()
      for (let row = 0; row < engine.rows; row++) {
        set.add(getCellDisplayString(row, col))
      }
      // Stable-ish ordering: numeric first (as numbers), then strings.
      const values = Array.from(set)
      values.sort((a, b) => {
        const aNum = typeof a === 'string' && a.trim() !== '' ? Number(a) : 0
        const bNum = typeof b === 'string' && b.trim() !== '' ? Number(b) : 0
        const aIsNum = !Number.isNaN(aNum) && a.trim() !== ''
        const bIsNum = !Number.isNaN(bNum) && b.trim() !== ''
        if (aIsNum && bIsNum) return aNum - bNum
        return String(a).localeCompare(String(b))
      })
      map[col] = values
    }
    return map
  }, [version, engine, getCellDisplayString])

  const visibleRowIndices = useMemo(() => {
    let rows = Array.from({ length: engine.rows }, (_, i) => i)

    // Apply filters first, then sorting.
    const activeFilters = Object.entries(filters).filter(([, values]) => Array.isArray(values) && values.length > 0)
    if (activeFilters.length > 0) {
      rows = rows.filter((rowIndex) => {
        for (const [colLetter, allowedValues] of activeFilters) {
          const colIndex = columnLetterToIndex(colLetter)
          if (colIndex === null || colIndex < 0 || colIndex >= engine.cols) continue
          const value = getCellDisplayString(rowIndex, colIndex)
          if (!allowedValues.includes(value)) return false
        }
        return true
      })
    }

    if (sortConfig.order && sortConfig.column) {
      const colIndex = columnLetterToIndex(sortConfig.column)
      if (colIndex !== null && colIndex >= 0 && colIndex < engine.cols) {
        const dir = sortConfig.order === 'asc' ? 1 : -1

        const numericRe = /^-?\d+(\.\d+)?$/
        const toSortable = (rowIndex) => {
          const cell = engine.getCell(rowIndex, colIndex)
          if (cell.error) return { kind: 'error', rowIndex }
          // Keep truly empty rows at the bottom (regardless of asc/desc)
          if (cell.computed === null || cell.computed === '') return { kind: 'empty', rowIndex }

          const raw = String(cell.computed)
          if (numericRe.test(raw.trim())) return { kind: 'number', num: Number(raw), rowIndex }
          return { kind: 'string', str: raw, rowIndex }
        }

        rows = rows
          .map((r, i) => ({ r, _i: i, key: toSortable(r) }))
          .sort((a, b) => {
            const ak = a.key
            const bk = b.key
            // Keep ERROR rows at the bottom.
            if (ak.kind === 'error' || bk.kind === 'error') {
              if (ak.kind === bk.kind) return a._i - b._i
              return ak.kind === 'error' ? 1 : -1
            }

            // Keep empty rows at the bottom (both asc/desc).
            if (ak.kind === 'empty' || bk.kind === 'empty') {
              if (ak.kind === bk.kind) return a._i - b._i
              return ak.kind === 'empty' ? 1 : -1
            }

            if (ak.kind === bk.kind) {
              if (ak.kind === 'number') return (ak.num - bk.num) * dir
              return ak.str.localeCompare(bk.str) * dir
            }
            // Numbers before strings.
            const order = ak.kind === 'number' ? -1 : 1
            return order
          })
          .map(x => x.r)
      }
    }

    // Include `version` in the computation so react-hooks knows this memo depends on it.
    // (The engine's internal state changes are surfaced via `forceRerender` -> `version`.)
    return rows.map(r => r + (version * 0))
  }, [version, engine, filters, sortConfig, columnLetterToIndex, getCellDisplayString])

  const selectedCellLabel = selectedCell
    ? `${getColumnLabel(selectedCell.c)}${selectedCell.r + 1}`
    : 'No cell'

  // Formula bar shows the raw formula text, not the computed value
  // When editing, show the current editValue; otherwise show the cell's raw content
  // Note: This is different from the cell display, which shows computed values
  const formulaBarValue = editingCell
    ? editValue
    : (selectedCell ? engine.getCell(selectedCell.r, selectedCell.c).raw : '')

  // ────── Render ──────

  const selectionRect = getSelectionRect()

  return (
    <div className="app-wrapper">
      <div className="app-header">
        <h2 className="app-title">📊 Spreadsheet App</h2>
      </div>

      <div className="main-content">

        {/* ── Toolbar ── */}
        <div className="toolbar">
          <div className="toolbar-group">
            <button className={`toolbar-btn bold-btn ${selectedCellStyle?.bold ? 'active' : ''}`} onClick={toggleBold} title="Bold">B</button>
            <button className={`toolbar-btn italic-btn ${selectedCellStyle?.italic ? 'active' : ''}`} onClick={toggleItalic} title="Italic">I</button>
            <button className={`toolbar-btn underline-btn ${selectedCellStyle?.underline ? 'active' : ''}`} onClick={toggleUnderline} title="Underline">U</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Size:</span>
            <select className="toolbar-select" value={selectedCellStyle?.fontSize || 13} onChange={(e) => changeFontSize(parseInt(e.target.value))}>
              {[8, 10, 11, 12, 13, 14, 16, 18, 20, 24].map(s => (
                <option key={s} value={s}>{s}</option>
              ))}
            </select>
          </div>

          <div className="toolbar-group">
            <button className={`align-btn ${selectedCellStyle?.align === 'left' ? 'active' : ''}`} onClick={() => changeAlignment('left')} title="Align Left">⬤←</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'center' ? 'active' : ''}`} onClick={() => changeAlignment('center')} title="Align Center">⬤</button>
            <button className={`align-btn ${selectedCellStyle?.align === 'right' ? 'active' : ''}`} onClick={() => changeAlignment('right')} title="Align Right">⬤→</button>
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Text:</span>
            <input
              type="color"
              value={selectedCellStyle?.color || '#000000'}
              onChange={(e) => changeFontColor(e.target.value)}
              title="Font color"
              style={{ width: '32px', height: '32px', border: '1px solid #dadce0', cursor: 'pointer', borderRadius: '4px' }}
            />
          </div>

          <div className="toolbar-group">
            <span className="toolbar-label">Fill:</span>
            <select className="toolbar-select" value={selectedCellStyle?.bg || 'white'} onChange={(e) => changeBackgroundColor(e.target.value)}>
              <option value="white">White</option>
              <option value="#ffff99">Yellow</option>
              <option value="#99ffcc">Green</option>
              <option value="#ffcccc">Red</option>
              <option value="#cce5ff">Blue</option>
              <option value="#e0ccff">Purple</option>
              <option value="#ffd9b3">Orange</option>
              <option value="#f0f0f0">Gray</option>
            </select>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={handleUndo} disabled={!engine.canUndo()} title="Undo">↶ Undo</button>
            <button className="toolbar-btn" onClick={handleRedo} disabled={!engine.canRedo()} title="Redo">↷ Redo</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn" onClick={insertRow} title="Insert Row">+ Row</button>
            <button className="toolbar-btn" onClick={deleteRow} title="Delete Row">- Row</button>
            <button className="toolbar-btn" onClick={insertColumn} title="Insert Column">+ Col</button>
            <button className="toolbar-btn" onClick={deleteColumn} title="Delete Column">- Col</button>
          </div>

          <div className="toolbar-group">
            <button className="toolbar-btn danger" onClick={clearSelectedCell}>✕ Cell</button>
            <button className="toolbar-btn danger" onClick={clearAllCells}>✕ All</button>
          </div>
        </div>

        {/* ── Formula Bar ── */}
        <div className="formula-bar">
          <span className="formula-bar-label">{selectedCellLabel}</span>
          <input
            className="formula-bar-input"
            value={formulaBarValue}
            onChange={(e) => handleFormulaBarChange(e.target.value)}
            onKeyDown={handleFormulaBarKeyDown}
            onFocus={handleFormulaBarFocus}
            placeholder="Select a cell then type, or enter a formula like =SUM(A1:A5)"
          />
        </div>

        {/* ── Grid ── */}
        <div className="grid-scroll">
          <table className="grid-table">
            <thead>
              <tr>
                <th className="col-header-blank"></th>
                {Array.from({ length: engine.cols }, (_, colIndex) => {
                  const colLetter = getColumnLabel(colIndex)
                  const activeFilterValues = filters[colLetter] || []
                  const sortIndicator =
                    sortConfig.column === colLetter && sortConfig.order
                      ? (sortConfig.order === 'asc' ? '▲' : '▼')
                      : ''

                  return (
                    <th key={colIndex} className="col-header">
                      <div className="col-header-inner">
                        <button
                          type="button"
                          className={`col-header-sort-btn ${sortConfig.column === colLetter && sortConfig.order ? 'active' : ''}`}
                          onClick={(e) => { e.stopPropagation(); toggleSort(colIndex) }}
                          title="Sort"
                        >
                          <span className="col-header-sort-label">{colLetter}</span>
                          <span className="col-sort-indicator">{sortIndicator}</span>
                        </button>

                        <select
                          multiple
                          className="col-filter-select"
                          value={activeFilterValues}
                          onClick={(e) => e.stopPropagation()}
                          onChange={(e) => {
                            const selected = Array.from(e.target.selectedOptions).map(o => o.value)
                            setFilters((prev) => {
                              const next = { ...prev }
                              if (selected.length === 0) delete next[colLetter]
                              else next[colLetter] = selected
                              return next
                            })
                          }}
                          title="Filter"
                        >
                          {(columnUniqueValues[colIndex] || []).map((val, idx) => (
                            <option key={`${idx}`} value={val}>
                              {val === '' ? '(empty)' : val}
                            </option>
                          ))}
                        </select>
                      </div>
                    </th>
                  )
                })}
              </tr>
            </thead>
            <tbody>
              {visibleRowIndices.length === 0 ? (
                <tr>
                  <td colSpan={engine.cols + 1} style={{ padding: '16px', textAlign: 'center', color: '#5f6368' }}>
                    No rows match current filter(s). Clear the header filter selections to see the data.
                  </td>
                </tr>
              ) : (
                visibleRowIndices.map((rowIndex, displayRowIndex) => (
                  <tr key={rowIndex}>
                    {/* Show the row number in the current visible (sorted/filtered) order. */}
                    <td className="row-header">{displayRowIndex + 1}</td>
                    {Array.from({ length: engine.cols }, (_, colIndex) => {
                      const isActive = selectedCell?.r === rowIndex && selectedCell?.c === colIndex
                      const inRange = selectionRect
                        ? rowIndex >= selectionRect.minR &&
                          rowIndex <= selectionRect.maxR &&
                          colIndex >= selectionRect.minC &&
                          colIndex <= selectionRect.maxC
                        : false
                      const isEditing = editingCell?.r === rowIndex && editingCell?.c === colIndex
                      const cellData = engine.getCell(rowIndex, colIndex)
                      const style = cellStyles[`${rowIndex},${colIndex}`] || {}
                      const displayValue = cellData.error
                        ? cellData.error
                        : (cellData.computed === null ? '' : String(cellData.computed))

                      return (
                        <td
                          key={colIndex}
                          className={`cell ${isActive ? 'selected' : ''} ${inRange && !isActive ? 'in-range' : ''}`}
                          // Inline background overrides CSS classes; apply the in-range highlight here
                          // so Shift-selection is visible.
                          style={{ background: (inRange && !isActive) ? '#e8f0fe' : (style.bg || 'white') }}
                          onMouseDown={(e) => { e.preventDefault(); handleCellClick(rowIndex, colIndex, e) }}
                        >
                          {isEditing ? (
                            <input
                              autoFocus
                              className="cell-input"
                              value={editValue}
                              onChange={(e) => setEditValue(e.target.value)}
                              onBlur={() => commitEdit(rowIndex, colIndex)}
                              onKeyDown={(e) => handleKeyDown(e, rowIndex, colIndex)}
                              ref={isActive ? cellInputRef : undefined}
                              style={{
                                fontWeight: style.bold ? 'bold' : 'normal',
                                fontStyle: style.italic ? 'italic' : 'normal',
                                textDecoration: style.underline ? 'underline' : 'none',
                                color: style.color || '#202124',
                                fontSize: (style.fontSize || 13) + 'px',
                                textAlign: style.align || 'left',
                                background: style.bg || 'white',
                              }}
                            />
                          ) : (
                            <div
                              className={`cell-display align-${style.align || 'left'} ${cellData.error ? 'error' : ''}`}
                              style={{
                                fontWeight: style.bold ? 'bold' : 'normal',
                                fontStyle: style.italic ? 'italic' : 'normal',
                                textDecoration: style.underline ? 'underline' : 'none',
                                color: cellData.error ? '#d93025' : (style.color || '#202124'),
                                fontSize: (style.fontSize || 13) + 'px',
                              }}
                            >
                              {displayValue}
                            </div>
                          )}
                        </td>
                      )
                    })}
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        <p className="footer-hint">
          Click a cell to edit · Enter/Tab/Arrow keys to navigate · Formulas: =A1+B1 · =SUM(A1:A5) · =AVG(A1:A5) · =MAX(A1:A5) · =MIN(A1:A5)
        </p>
      </div>
    </div>
  )
}
