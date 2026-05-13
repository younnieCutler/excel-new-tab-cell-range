# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm run dev        # Start webpack-dev-server on https://localhost:3000 (hot reload)
npm test           # Run focused Node regression tests
npm start          # Start dev server + sideload add-in into Excel Desktop
npm run build      # Production build → dist/
npm run stop       # Unload add-in from Excel
npm run validate   # Validate manifest.xml
```

First-time setup: `npx office-addin-dev-certs install` (generates trusted HTTPS cert for localhost).

For public Marketplace deployment on this repo, use `npm run build:github-pages`; it generates a manifest for `https://younnieCutler.github.io/excel-new-tab-cell-range`. For another public host, use `npm run build:marketplace -- --base-url https://<public-host> --support-url https://<public-host>/support.html`, deploy `dist/`, then run `npm run validate:marketplace`. For customer/internal deployment, use `npm run build:customer -- --base-url https://<customer-host>`.

## Architecture

Single Shared Runtime add-in. All JS runs in one context shared by the taskpane and the function file.

```
src/
├── taskpane/
│   ├── taskpane.js          # Entry point. Wires modules, exposes window.captureSelectedRange
│   ├── taskpane.html        # Shell HTML (Office.js CDN + webpack-injected bundle)
│   ├── taskpane.css         # Excel-like light theme tokens + component styles
│   └── modules/
│       ├── i18n.js          # t(key) — Japanese default, English fallback via Office.context.displayLanguage
│       ├── utils.js         # parseAddress, isAddressWithinRange, buildCellStyle
│       ├── tabManager.js    # Tab CRUD (max 8), active tab state, onChange callbacks
│       ├── syncEngine.js    # captureSelection/captureRange, source cell selection, change listeners
│       └── gridRenderer.js  # <table> renderer: displays captured cells and selects source cells
└── commands/
    ├── commands.js          # openInCellFocus: shows taskpane then calls window.captureSelectedRange
    └── commands.html        # Minimal function file HTML (no visible UI)
src/shortcuts.json           # Ctrl+Shift+F → openInCellFocus mapping
manifest.xml                 # Add-in manifest: Shared Runtime, ContextMenu, ribbon button
```

## Key Behaviors

**Captured values**: Loads `address`, `rowCount`, `columnCount`, `values`, and `text` for the selected range. The taskpane displays Excel's formatted `text` when available and falls back to raw `value`.

**Native Excel operation flow**: Click a `<td>` in the taskpane → `syncEngine.selectSourceCell()` activates the tab's source worksheet and selects the matching cell in Excel. Editing, undo, keyboard movement, fill handle, and formulas remain native Excel behavior.

**Change listeners**: Each captured tab registers a worksheet-level `onChanged` listener against its source sheet. Listener cleanup is keyed by tab ID and removes the handler from the same worksheet it was added to.

**Multiple sheets**: Selection tracking records `worksheetId`, so the taskpane can capture a selection made on another sheet even while the taskpane keeps focus. Source-cell clicks explicitly activate the source worksheet before selecting the cell.

**Deployment**: Public Marketplace/AppSource uses public HTTPS static hosting + Partner Center submission. Customer-owned HTTPS static hosting + Microsoft 365 Admin centralized deployment remains available for private installs.
