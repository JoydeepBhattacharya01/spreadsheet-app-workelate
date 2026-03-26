# SpreadsheetApp

A lightweight Excel-like spreadsheet built with React + Vite.

This project was implemented as an assignment with three core deliverables:
- **Column Sort & Filter** (view-layer only, computed-value aware)
- **Multi-cell Copy & Paste** (clipboard integration, undoable)
- **Local Storage Persistence** (debounced autosave, safe restore)

## Prerequisites

Before you begin, ensure you have the following installed on your system:

- **Node.js** (version 18 or higher) - [Download](https://nodejs.org/)
- **npm** (comes with Node.js) or **yarn**
- **Git** - [Download](https://git-scm.com/)

## Getting Started

### 2. Install Dependencies

Install all required project dependencies:

```bash
npm install
```

Or if you prefer yarn:

```bash
yarn install
```

### 3. Run Development Server

Start the development server with hot module replacement (HMR):

```bash
npm run dev
```

The application will be available at `http://localhost:5173`

### 4. Build for Production

Create an optimized production build:

```bash
npm run build
```

The build output will be in the `dist/` directory.

### 5. Preview Production Build

Preview the production build locally:

```bash
npm run preview
```

### 6. Lint Code

Run ESLint to check for code quality issues:

```bash
npm run lint
```



## Browser Support

This application works on all modern browsers that support ES2020+ JavaScript:

- Chrome (latest)
- Firefox (latest)
- Safari (latest)
- Edge (latest)

## Troubleshooting

### Dependencies won't install
- Clear npm cache: `npm cache clean --force`
- Delete `node_modules` and `package-lock.json`, then reinstall: `rm -rf node_modules package-lock.json && npm install`

### Port 5173 already in use
- The dev server will automatically try the next available port
- Or specify a custom port: `npm run dev -- --port 3000`

### Build fails
- Ensure all dependencies are installed: `npm install`
- Clear any build cache: `rm -rf dist`
- Try rebuilding: `npm run build`
