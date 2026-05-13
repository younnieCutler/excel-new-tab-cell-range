import { isAddressWithinRange } from './utils.js';

// Prevents write-back from triggering the onChanged listener
let isWritingBack = false;

// eventHandlers: tabId → { sheet, handler } for cleanup
const eventHandlers = new Map();

function stripSheetPrefix(address) {
    const bangIndex = address.indexOf('!');
    return bangIndex === -1 ? address : address.slice(bangIndex + 1);
}

function buildCells(values, text) {
    const rowCount = values.length;
    const colCount = values[0]?.length ?? 0;
    return Array.from({ length: rowCount }, (_, row) => (
        Array.from({ length: colCount }, (_, col) => ({
            value: values[row]?.[col] ?? '',
            text: text[row]?.[col] ?? values[row]?.[col] ?? '',
        }))
    ));
}

async function captureRangeInContext(ctx, sheetName, address) {
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const rangeAddress = stripSheetPrefix(address);
    const range = sheet.getRange(rangeAddress);
    range.load(['address', 'rowCount', 'columnCount', 'values', 'text']);
    await ctx.sync();

    return {
        sheetName,
        address: range.address,
        rowCount: range.rowCount,
        colCount: range.columnCount,
        cells: buildCells(range.values, range.text),
    };
}

// Capture all display text and raw values for a range.
export async function captureRange(sheetName, address) {
    return Excel.run((ctx) => captureRangeInContext(ctx, sheetName, address));
}

// Capture only current selection (uses active worksheet)
export async function captureSelection() {
    return Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load(['address', 'worksheet/name']);
        await ctx.sync();

        const sheetName = range.worksheet.name;
        const address = range.address;
        return captureRangeInContext(ctx, sheetName, address);
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
