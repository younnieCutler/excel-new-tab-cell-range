import { parseAddress, buildCellStyle } from './utils.js';
import { writeBack } from './syncEngine.js';

export class GridRenderer {
    constructor(container, onSyncStart, onSyncDone, onSyncError) {
        this.container = container;
        this.onSyncStart = onSyncStart;
        this.onSyncDone = onSyncDone;
        this.onSyncError = onSyncError;
        this.activeTab = null;
        this._focusedCell = null; // { row, col }
    }

    render(tab) {
        this.activeTab = tab;
        this.container.innerHTML = '';

        const table = document.createElement('table');
        table.className = 'grid-table';
        const tbody = document.createElement('tbody');

        // Build set of slave cells (merged but not top-left master)
        const slaveSet = this._buildSlaveSet(tab);

        for (let r = 0; r < tab.rowCount; r++) {
            const tr = document.createElement('tr');
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

        // Merge span
        if (cellProps.isMerged && cellProps.mergeArea?.address) {
            const merge = parseAddress(cellProps.mergeArea.address);
            if (merge) {
                const rowspan = merge.endRow - merge.startRow + 1;
                const colspan = merge.endCol - merge.startCol + 1;
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

        // Edit input (hidden by default)
        const input = document.createElement('input');
        input.className = 'cell-edit';
        input.type = 'text';
        input.value = cellProps.value ?? '';
        input.hidden = true;
        input.setAttribute('aria-label', `Row ${r + 1}, Column ${c + 1}`);

        td.appendChild(span);
        td.appendChild(input);

        this._attachCellEvents(td, input, span, tab, r, c);
        return td;
    }

    _attachCellEvents(td, input, span, tab, r, c) {
        const enterEdit = () => {
            span.hidden = true;
            input.hidden = false;
            input.value = tab.cells[r]?.[c]?.value ?? '';
            input.focus();
            input.select();
        };

        const exitEdit = async (save) => {
            input.hidden = true;
            span.hidden = false;
            if (!save) return;

            const newValue = input.value;
            this.onSyncStart?.();
            try {
                const displayText = await writeBack(tab, r, c, newValue);
                span.textContent = displayText;
                this.onSyncDone?.();
            } catch (err) {
                console.error('writeBack failed', err);
                this.onSyncError?.();
            }
        };

        // Click to edit
        td.addEventListener('click', () => {
            if (!input.hidden) return;
            enterEdit();
        });

        // Typing directly on focused td
        td.addEventListener('keydown', (e) => {
            if (input.hidden && e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
                enterEdit();
                input.value = e.key;
            }
        });
        td.setAttribute('tabindex', '0');

        // Input key handling
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                exitEdit(false);
                td.focus();
            } else if (e.key === 'Enter') {
                e.preventDefault();
                exitEdit(true).then(() => this._moveFocus(r + 1, c));
            } else if (e.key === 'Tab') {
                e.preventDefault();
                exitEdit(true).then(() => this._moveFocus(r, c + (e.shiftKey ? -1 : 1)));
            }
        });

        input.addEventListener('blur', (e) => {
            // Only save on true blur (not when navigating via Tab/Enter handled above)
            if (!input.hidden) exitEdit(true);
        });
    }

    _moveFocus(row, col) {
        const td = this.container.querySelector(`[data-row="${row}"][data-col="${col}"]`);
        if (td) td.focus();
    }
}
