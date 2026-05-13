# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm run dev        # Start webpack-dev-server on https://localhost:3000 (hot reload)
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
│   ├── taskpane.css         # Dark theme design tokens + component styles
│   └── modules/
│       ├── i18n.js          # t(key) — Japanese default, English fallback via Office.context.displayLanguage
│       ├── utils.js         # parseAddress, isAddressWithinRange, buildCellStyle
│       ├── tabManager.js    # Tab CRUD (max 8), active tab state, onChange callbacks
│       ├── syncEngine.js    # captureSelection/captureRange, writeBack, change listeners
│       └── gridRenderer.js  # <table> renderer: Excel formatting → inline CSS, span/input edit toggle
└── commands/
    ├── commands.js          # openInCellFocus: shows taskpane then calls window.captureSelectedRange
    └── commands.html        # Minimal function file HTML (no visible UI)
src/shortcuts.json           # Ctrl+Shift+F → openInCellFocus mapping
manifest.xml                 # Add-in manifest: Shared Runtime, ContextMenu, ribbon button
```

## Key Behaviors

**Formatting**: Uses `getCellProperties()` to load `text` (Excel-formatted display string), raw `value`, and `format.*` (fill/font/borders/alignment) per cell. The `text` field is displayed; `value` is used for editing.

**Edit flow**: Click/type on a `<td>` → hides `<span class="cell-display">`, shows `<input class="cell-edit">` with raw value → Enter/Tab/blur triggers `syncEngine.writeBack()` → reloads formatted `text` from Excel.

**Event loop prevention**: `syncEngine.js` uses a module-level `isWritingBack` boolean flag. The `onChanged` listener skips processing while `writeBack` is in progress.

**Merge cells**: `getCellProperties()` returns `isMerged` + `mergeArea.address` per cell. `gridRenderer.js` identifies slave cells (non-top-left in a merge area) and skips `<td>` creation for them; master cells get `rowSpan`/`colSpan`.

**Deployment**: Public Marketplace/AppSource uses public HTTPS static hosting + Partner Center submission. Customer-owned HTTPS static hosting + Microsoft 365 Admin centralized deployment remains available for private installs.
