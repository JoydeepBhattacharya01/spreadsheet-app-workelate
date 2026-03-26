# 📊 Spreadsheet App (WorkElate Assignment)

An Excel-like spreadsheet web application built with React and Vite, featuring formula evaluation, view-layer sorting & filtering, clipboard integration, undo support, and localStorage persistence.

---

## 🚀 Features

- Editable spreadsheet grid (50 × 50)
- Formula support:
  - Arithmetic: `+ - * /`
  - Cell references: `A1`, `B2`
  - Ranges: `A1:A5`
  - Functions: `SUM`, `AVG`, `MIN`, `MAX`
- Column sorting (ascending → descending → none)
- Column filtering with dropdown (multi-select)
- Multi-cell copy & paste (Excel/Google Sheets compatible)
- Undo support (Cmd/Ctrl + Z)
- Auto-save with localStorage (debounced)
- Restore data on page reload
- Cell formatting (bold, italic, colors, alignment, etc.)

---

## 🧱 Architecture

The application is designed with a clear separation of concerns:

### UI Layer (`src/App.jsx`)
- Handles rendering of the spreadsheet grid
- Manages UI state:
  - Selection and editing
  - Sorting and filtering
  - Clipboard interactions
  - Cell styling
- Integrates keyboard shortcuts and user interactions

### Engine Layer (`src/engine/core.js`)
- Manages:
  - Cell values and formulas
  - Formula parsing and evaluation
  - Dependency tracking
  - Recalculation of affected cells
- Handles:
  - Error detection (invalid formulas, circular references)
  - Undo/redo logic
  - State serialization for persistence

---

## 🔑 Key Engineering Decisions

- **View-layer sorting & filtering**  
  Sorting and filtering are applied only at the UI level without mutating underlying data. This ensures formulas continue referencing original cell positions, similar to real spreadsheet behavior.

- **Formula engine with dependency tracking**  
  Implemented a lightweight formula engine supporting arithmetic operations, cell references, ranges, and functions. A dependency graph ensures automatic recalculation when referenced cells change.

- **Copy returns computed values**  
  Clipboard copy operations return computed values instead of formulas to match expected spreadsheet behavior.

- **Undo system using history stack**  
  Implemented undo functionality for edits and paste operations using a state history stack.

- **Debounced localStorage persistence**  
  Used a 500ms debounce to reduce excessive writes and improve performance.

- **Safe state restoration**  
  Corrupted localStorage data is handled gracefully using try-catch to prevent app crashes.

---

## 🔽 Assignment Requirements Coverage

### Task 1 — Column Sort & Filter
- Sorting cycles: `asc → desc → none`
- Sorting uses computed values (formula results)
- Filter dropdown per column (multi-select)
- Filtering hides rows (non-destructive)
- Sorting/filtering reversible
- Implemented at view layer (formulas unaffected)

### Task 2 — Multi-Cell Copy & Paste
- Paste from Excel / Google Sheets (tab + newline parsing)
- Supports multi-row and multi-column data
- Paste is undoable (Cmd/Ctrl + Z)
- Copy returns computed values
- Internal copy-paste supported

### Task 3 — Local Storage Persistence
- Auto-save with debounce (~500ms)
- Restore state on reload
- Persists:
  - Cell values and formulas
  - Grid dimensions
  - Cell styles
- Does NOT persist undo/redo history
- Handles corrupted storage safely

---

## ⚠️ Edge Cases Handled

- Invalid formulas return `ERROR`
- Circular references handled safely
- Empty cells treated as 0 in formulas
- Sorting keeps empty rows at bottom
- Sorting + filtering combined behavior
- Large clipboard paste handled efficiently
- Clipboard fallback if API fails
- Corrupted localStorage ignored safely

---

## 🎥 Demo

- Loom Video: https://www.loom.com/share/789ce901e4d24c0d85365c454ae4161a
- Live Demo: https://spreadsheet-app-workelate.vercel.app/

---

## 🛠️ Tech Stack

- React (UI & state management)
- Vite (development & build tool)
- JavaScript (ES2020+)
- CSS (styling)

---

## ⚙️ Getting Started

### Prerequisites
- Node.js (v18+)
- npm or yarn
- Git

### Installation
```bash
npm install

Run Development Server
npm run dev
Open: http://localhost:5173

Build for Production
npm run build

Preview Production Build
npm run preview

Lint Code
npm run lint
