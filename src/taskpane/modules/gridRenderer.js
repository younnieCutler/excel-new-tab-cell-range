import { parseAddress, buildCellStyle, buildTsvFromSelection, normalizeSelection, parseTsv } from './utils.js';
import { selectSourceRange, writeCell, writeRange } from './syncEngine.js';

export class GridRenderer {
    constructor(container, onSyncStart, onSyncDone, onSyncError, onSelectionChange) {
        this.container = container;
        this.onSyncStart = onSyncStart;
        this.onSyncDone = onSyncDone;
        this.onSyncError = onSyncError;
        this.onSelectionChange = onSelectionChange;
        this.activeTab = null;
        this._selection = null;
        this._dragAnchor = null;
        this._isDragging = false;
        this._editingInput = null;

        this.container.tabIndex = 0;
        this.container.addEventListener('keydown', (event) => this._handleKeyDown(event));
        this.container.addEventListener('pointermove', (event) => this._handlePointerMove(event));
        window.addEventListener('pointerup', () => this._finishDrag());
    }

    render(tab) {
        this.activeTab = tab;
        this._selection = tab.selection ?? this._createSelection(0, 0, 0, 0);
        this.container.innerHTML = '';

        const table = document.createElement('table');
        table.className = 'grid-table';
        table.addEventListener('selectstart', (event) => event.preventDefault());

        const colgroup = document.createElement('colgroup');
        for (let c = 0; c < tab.colCount; c++) {
            const col = document.createElement('col');
            const width = this._columnWidthToPx(tab.columnWidths?.[c]);
            if (width) col.style.width = `${width}px`;
            colgroup.appendChild(col);
        }
        table.appendChild(colgroup);

        const tbody = document.createElement('tbody');

        // Build set of slave cells (merged but not top-left master)
        const slaveSet = this._buildSlaveSet(tab);

        for (let r = 0; r < tab.rowCount; r++) {
            const tr = document.createElement('tr');
            const height = this._rowHeightToPx(tab.rowHeights?.[r]);
            if (height) tr.style.height = `${height}px`;
            for (let c = 0; c < tab.colCount; c++) {
                if (slaveSet.has(`${r},${c}`)) continue;

                const cellProps = tab.cells[r]?.[c] ?? {};
                const td = this._buildCell(tab, r, c, cellProps, slaveSet);
                tr.appendChild(td);
            }
            tbody.appendChild(tr);
        }

        table.appendChild(tbody);
        this.container.appendChild(table);
        this._applySelectionClasses();
    }

    updateCell(tab, row, col, displayText) {
        const td = this.container.querySelector(`[data-row="${row}"][data-col="${col}"]`);
        if (!td) return;
        const span = td.querySelector('.cell-display');
        if (span) span.textContent = displayText;
    }

    _buildSlaveSet(tab) {
        const rangeAddr = parseAddress(tab.address);
        const slaveSet = new Set();

        for (let r = 0; r < tab.rowCount; r++) {
            for (let c = 0; c < tab.colCount; c++) {
                const cell = tab.cells[r]?.[c];
                if (!cell?.isMerged || !cell.mergeArea?.address) continue;

                const merge = parseAddress(cell.mergeArea.address);
                if (!merge) continue;

                const masterRelRow = merge.startRow - (rangeAddr?.startRow ?? 0);
                const masterRelCol = merge.startCol - (rangeAddr?.startCol ?? 0);

                if (masterRelRow !== r || masterRelCol !== c) {
                    slaveSet.add(`${r},${c}`);
                }
            }
        }
        return slaveSet;
    }

    _buildCell(tab, r, c, cellProps, slaveSet) {
        const rangeAddr = parseAddress(tab.address);
        const td = document.createElement('td');
        td.dataset.row = r;
        td.dataset.col = c;
        td.tabIndex = -1;

        // Merge span
        if (cellProps.isMerged && cellProps.mergeArea?.address) {
            const merge = parseAddress(cellProps.mergeArea.address);
            if (merge) {
                const rowspan = Math.min(merge.endRow, (rangeAddr?.endRow ?? merge.endRow)) - Math.max(merge.startRow, (rangeAddr?.startRow ?? merge.startRow)) + 1;
                const colspan = Math.min(merge.endCol, (rangeAddr?.endCol ?? merge.endCol)) - Math.max(merge.startCol, (rangeAddr?.startCol ?? merge.startCol)) + 1;
                if (rowspan > 1) td.rowSpan = rowspan;
                if (colspan > 1) td.colSpan = colspan;
            }
        }

        // Inline style from Excel format
        const styleStr = buildCellStyle(cellProps.format);
        if (styleStr) td.style.cssText = styleStr;

        // Display span (shows formatted text)
        const span = document.createElement('span');
        span.className = 'cell-display';
        span.textContent = cellProps.text ?? cellProps.value ?? '';

        td.appendChild(span);

        this._attachCellEvents(td, tab, r, c, cellProps);
        return td;
    }

    _attachCellEvents(td, tab, r, c, cellProps) {
        td.addEventListener('pointerdown', (event) => {
            if (td.querySelector('.cell-edit')) return;
            event.preventDefault();
            this.container.focus();
            this._isDragging = true;
            this._dragAnchor = this._expandPointForMerge(tab, r, c);
            this._setSelection(this._dragAnchor, false);
        });

        td.addEventListener('pointerup', (event) => {
            if (!this._isDragging) return;
            event.preventDefault();
            this._finishDrag();
        });

        td.addEventListener('dblclick', () => {
            this._startEdit(td, tab, r, c, cellProps);
        });
    }

    _startEdit(td, tab, r, c, cellProps) {
        if (this._editingInput || td.querySelector('.cell-edit')) return;

        this._setSelection(this._createSelection(r, c, r, c), false);
        if (td.querySelector('.cell-edit')) return;

        const span = td.querySelector('.cell-display');
        const originalValue = cellProps.value ?? '';

        const input = document.createElement('input');
        input.className = 'cell-edit';
        input.value = String(originalValue);
        span.replaceWith(input);
        this._editingInput = input;
        input.focus();
        input.select();

        let committed = false;
        const commit = async () => {
            if (committed) return;
            committed = true;
            this._editingInput = null;

            const newValue = input.value;
            const restoredSpan = document.createElement('span');
            restoredSpan.className = 'cell-display';
            restoredSpan.textContent = newValue !== '' ? newValue : (cellProps.text ?? '');
            if (input.parentNode) input.replaceWith(restoredSpan);

            if (newValue !== String(originalValue)) {
                this.onSyncStart?.();
                try {
                    const result = await writeCell(tab, r, c, newValue);
                    restoredSpan.textContent = result.text ?? String(result.value) ?? newValue;
                    this.onSyncDone?.();
                } catch (err) {
                    console.error('Cell write failed:', err);
                    restoredSpan.textContent = cellProps.text ?? cellProps.value ?? '';
                    this.onSyncError?.();
                }
            }
        };

        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); commit(); }
            if (e.key === 'Escape') {
                committed = true;
                this._editingInput = null;
                const restoredSpan = document.createElement('span');
                restoredSpan.className = 'cell-display';
                restoredSpan.textContent = cellProps.text ?? cellProps.value ?? '';
                if (input.parentNode) input.replaceWith(restoredSpan);
            }
        });
        input.addEventListener('blur', commit);
    }

    _setSelection(selection, notify = true) {
        this._selection = normalizeSelection(selection);
        if (this.activeTab) this.activeTab.selection = this._selection;
        this._applySelectionClasses();
        if (notify) this.onSelectionChange?.(this.activeTab?.id, this._selection);
    }

    _applySelectionClasses() {
        if (!this._selection) return;
        const selection = normalizeSelection(this._selection);
        this.container.querySelectorAll('td').forEach((td) => {
            const row = Number(td.dataset.row);
            const col = Number(td.dataset.col);
            const active = row === selection.startRow && col === selection.startCol;
            const selected = row >= selection.startRow && row <= selection.endRow && col >= selection.startCol && col <= selection.endCol;
            td.classList.toggle('source-selected', selected);
            td.classList.toggle('active-cell', active);
        });
    }

    _createSelection(startRow, startCol, endRow, endCol) {
        return { startRow, startCol, endRow, endCol };
    }

    _expandPointForMerge(tab, row, col) {
        const rangeAddr = parseAddress(tab.address);
        const cell = tab.cells[row]?.[col];
        if (!cell?.isMerged || !cell.mergeArea?.address) {
            return this._createSelection(row, col, row, col);
        }

        const merge = parseAddress(cell.mergeArea.address);
        if (!merge) return this._createSelection(row, col, row, col);
        const sourceStartRow = rangeAddr?.startRow ?? 0;
        const sourceStartCol = rangeAddr?.startCol ?? 0;
        return {
            startRow: Math.max(0, merge.startRow - sourceStartRow),
            startCol: Math.max(0, merge.startCol - sourceStartCol),
            endRow: Math.min(tab.rowCount - 1, merge.endRow - sourceStartRow),
            endCol: Math.min(tab.colCount - 1, merge.endCol - sourceStartCol),
        };
    }

    async _syncSelectionToExcel() {
        if (!this.activeTab || !this._selection) return;
        this.onSelectionChange?.(this.activeTab.id, this._selection);
        try {
            await selectSourceRange(this.activeTab, this._selection);
        } catch (err) {
            console.warn('Source selection sync failed:', err);
        }
    }

    _handlePointerMove(event) {
        if (!this._isDragging || !this._dragAnchor || !this.activeTab) return;
        const targetCell = document.elementFromPoint(event.clientX, event.clientY)?.closest?.('td[data-row][data-col]');
        if (!targetCell || !this.container.contains(targetCell)) return;

        const target = this._expandPointForMerge(
            this.activeTab,
            Number(targetCell.dataset.row),
            Number(targetCell.dataset.col),
        );
        this._setSelection({
            startRow: this._dragAnchor.startRow,
            startCol: this._dragAnchor.startCol,
            endRow: target.endRow,
            endCol: target.endCol,
        }, false);
    }

    _finishDrag() {
        if (!this._isDragging) return;
        this._isDragging = false;
        this._dragAnchor = null;
        this._syncSelectionToExcel();
    }

    async _handleKeyDown(event) {
        if (!this.activeTab || !this._selection || this._editingInput) return;

        if (event.key === 'Enter' || event.key === 'F2') {
            event.preventDefault();
            const selection = normalizeSelection(this._selection);
            const td = this.container.querySelector(`[data-row="${selection.startRow}"][data-col="${selection.startCol}"]`);
            const cellProps = this.activeTab.cells[selection.startRow]?.[selection.startCol] ?? {};
            if (td) this._startEdit(td, this.activeTab, selection.startRow, selection.startCol, cellProps);
            return;
        }

        if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'c') {
            event.preventDefault();
            if (navigator.clipboard) {
                await navigator.clipboard.writeText(buildTsvFromSelection(this.activeTab, this._selection));
            }
            return;
        }

        if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'v') {
            event.preventDefault();
            if (!navigator.clipboard) return;
            const text = await navigator.clipboard.readText();
            await this._pasteTsv(text);
        }
    }

    async _pasteTsv(text) {
        const selection = normalizeSelection(this._selection);
        const matrix = parseTsv(text);
        const clippedRows = matrix
            .slice(0, this.activeTab.rowCount - selection.startRow)
            .map((row) => row.slice(0, this.activeTab.colCount - selection.startCol));
        const width = Math.max(...clippedRows.map((row) => row.length));
        const clipped = clippedRows.map((row) => Array.from({ length: width }, (_, index) => row[index] ?? ''));
        if (!clipped.length || !clipped[0]?.length) return;

        this.onSyncStart?.();
        try {
            await writeRange(this.activeTab, selection.startRow, selection.startCol, clipped);
            for (let r = 0; r < clipped.length; r++) {
                for (let c = 0; c < clipped[r].length; c++) {
                    const target = this.activeTab.cells[selection.startRow + r]?.[selection.startCol + c];
                    if (target) {
                        target.value = clipped[r][c];
                        target.text = clipped[r][c];
                    }
                }
            }
            this.render(this.activeTab);
            this.onSyncDone?.();
        } catch (err) {
            console.error('Range paste failed:', err);
            this.onSyncError?.();
        }
    }

    _columnWidthToPx(width) {
        if (!Number.isFinite(width)) return null;
        return Math.max(32, Math.round(width * 7));
    }

    _rowHeightToPx(height) {
        if (!Number.isFinite(height)) return null;
        return Math.max(18, Math.round(height * 1.33));
    }
}
