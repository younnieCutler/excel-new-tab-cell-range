import { isAddressWithinRange } from './utils.js';

// eventHandlers: tabId → { sheet, handler } for cleanup
const eventHandlers = new Map();
let lastWorksheetSelection = null;

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

export async function registerSelectionTracker() {
    await Excel.run(async (ctx) => {
        ctx.workbook.worksheets.onSelectionChanged.add((event) => {
            lastWorksheetSelection = {
                address: event.address,
                worksheetId: event.worksheetId,
                capturedAt: Date.now(),
            };
        });
        await ctx.sync();
    });
}

async function captureTrackedSelection() {
    if (!lastWorksheetSelection) return null;

    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(lastWorksheetSelection.worksheetId);
        sheet.load('name');
        await ctx.sync();
        return captureRangeInContext(ctx, sheet.name, lastWorksheetSelection.address);
    });
}

// Capture only current selection (uses active worksheet)
export async function captureSelection({ preferTrackedSelection = false } = {}) {
    if (preferTrackedSelection) {
        try {
            const tracked = await captureTrackedSelection();
            if (tracked) return tracked;
        } catch (err) {
            console.warn('Tracked selection capture failed, falling back to current selection:', err);
        }
    }

    return Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load(['address', 'worksheet/name']);
        await ctx.sync();

        const sheetName = range.worksheet.name;
        const address = range.address;
        return captureRangeInContext(ctx, sheetName, address);
    });
}

export async function selectSourceCell(tab, row, col) {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(tab.sheetName);
        const cell = sheet.getRange(tab.address).getCell(row, col);
        cell.select();
        await ctx.sync();
    });
}

// Register onChanged listener for a tab; returns cleanup function
export async function registerChangeListener(tab, onCellChanged) {
    const handlerFn = async (args) => {
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
