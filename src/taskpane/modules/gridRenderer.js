import { parseAddress, buildCellStyle } from './utils.js';
import { selectSourceCell } from './syncEngine.js';

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

        td.title = 'Selects the source cell in Excel. Edit in Excel to keep native undo and keyboard behavior.';
        td.appendChild(span);

        this._attachCellEvents(td, tab, r, c);
        return td;
    }

    _attachCellEvents(td, tab, r, c) {
        td.addEventListener('click', async () => {
            this.onSyncStart?.();
            try {
                await selectSourceCell(tab, r, c);
                this._markSelectedCell(r, c);
                this.onSyncDone?.();
            } catch (err) {
                console.error('select source cell failed', err);
                this.onSyncError?.();
            }
        });
    }

    _markSelectedCell(row, col) {
        this.container.querySelectorAll('.source-selected').forEach((el) => {
            el.classList.remove('source-selected');
        });
        const td = this.container.querySelector(`[data-row="${row}"][data-col="${col}"]`);
        td?.classList.add('source-selected');
    }
}
