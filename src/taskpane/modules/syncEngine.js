import { isAddressWithinRange } from './utils.js';

// Prevents write-back from triggering the onChanged listener
let isWritingBack = false;

// eventHandlers: tabId → { sheet, handler } for cleanup
const eventHandlers = new Map();

// Capture all cell data for a range using getCellProperties (Office.js 1.9+)
export async function captureRange(sheetName, address) {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(sheetName);
        const range = sheet.getRange(address);
        range.load(['rowCount', 'columnCount']);

        const propsResult = range.getCellProperties({
            address: true,
            value: true,
            text: true,
            isMerged: true,
            mergeArea: { address: true },
            format: {
                fill: { color: true },
                font: {
                    bold: true,
                    italic: true,
                    size: true,
                    name: true,
                    color: true,
                    strikethrough: true,
                    underline: true,
                },
                horizontalAlignment: true,
                verticalAlignment: true,
                wrapText: true,
                borders: {
                    top: { style: true, color: true, weight: true },
                    bottom: { style: true, color: true, weight: true },
                    left: { style: true, color: true, weight: true },
                    right: { style: true, color: true, weight: true },
                },
            },
        });

        await ctx.sync();

        return {
            sheetName,
            address,
            rowCount: range.rowCount,
            colCount: range.columnCount,
            cells: propsResult.value,  // 2D array of CellProperties
        };
    });
}

// Capture only current selection (uses active worksheet)
export async function captureSelection() {
    return Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load(['address', 'worksheet/name']);
        await ctx.sync();

        const sheetName = range.worksheet.name;
        const address = range.address;
        return captureRange(sheetName, address);
    });
}

// Write a new value back to Excel and reload the cell's text (formatted display)
export async function writeBack(tab, row, col, value) {
    return Excel.run(async (ctx) => {
        isWritingBack = true;
        try {
            const sheet = ctx.workbook.worksheets.getItem(tab.sheetName);
            const cell = sheet.getRange(tab.address).getCell(row, col);
            cell.values = [[value]];
            await ctx.sync();

            // Reload formatted display text after write
            cell.load('text');
            await ctx.sync();

            const displayText = cell.text[0][0];
            if (tab.cells[row] && tab.cells[row][col]) {
                tab.cells[row][col].value = value;
                tab.cells[row][col].text = displayText;
            }
            return displayText;
        } finally {
            isWritingBack = false;
        }
    });
}

// Register onChanged listener for a tab; returns cleanup function
export async function registerChangeListener(tab, onCellChanged) {
    const handlerFn = async (args) => {
        if (isWritingBack) return;
        if (isAddressWithinRange(args.address, tab.address)) {
            onCellChanged(args.address);
        }
    };

    await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(tab.sheetName);
        sheet.onChanged.add(handlerFn);
        await ctx.sync();
    });

    eventHandlers.set(tab.id, handlerFn);

    return () => deregisterChangeListener(tab.id);
}

export async function deregisterChangeListener(tabId) {
    const handler = eventHandlers.get(tabId);
    if (!handler) return;
    try {
        await Excel.run(async (ctx) => {
            // Office.js removes by reference
            ctx.workbook.worksheets.onChanged.remove(handler);
            await ctx.sync();
        });
    } catch {
        // Ignore cleanup errors
    }
    eventHandlers.delete(tabId);
}
