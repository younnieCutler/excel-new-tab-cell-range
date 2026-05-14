import { isAddressWithinRange, selectionToSourceAddress } from './utils.js';

// eventHandlers: tabId → { sheetName, handler } for cleanup
const eventHandlers = new Map();
let lastWorksheetSelection = null;

function stripSheetPrefix(address) {
    const bangIndex = address.indexOf('!');
    return bangIndex === -1 ? address : address.slice(bangIndex + 1);
}

const BORDER_SIDES = {
    top: 'EdgeTop',
    bottom: 'EdgeBottom',
    left: 'EdgeLeft',
    right: 'EdgeRight',
};

function normalizeAddress(address) {
    return stripSheetPrefix(address).replace(/\$/g, '');
}

function readBorder(border) {
    return {
        color: border.color,
        style: border.style,
        weight: border.weight,
    };
}

function readCellFormat(format, borders) {
    return {
        fill: { color: format.fill.color },
        font: {
            bold: format.font.bold,
            color: format.font.color,
            italic: format.font.italic,
            name: format.font.name,
            size: format.font.size,
            strikethrough: format.font.strikethrough,
            underline: format.font.underline,
        },
        horizontalAlignment: format.horizontalAlignment,
        verticalAlignment: format.verticalAlignment,
        wrapText: format.wrapText,
        borders,
    };
}

function createCellModel(values, text, formulas, cells, row, col) {
    const cell = cells[row]?.[col];
    return {
        value: values[row]?.[col] ?? '',
        text: text[row]?.[col] ?? values[row]?.[col] ?? '',
        formula: formulas[row]?.[col] ?? null,
        format: cell?.format ?? null,
        isMerged: cell?.isMerged ?? false,
        mergeArea: cell?.mergeArea ?? null,
    };
}

function buildCells(values, text, formulas, cellDetails) {
    const rowCount = values.length;
    const colCount = values[0]?.length ?? 0;
    return Array.from({ length: rowCount }, (_, row) => (
        Array.from({ length: colCount }, (_, col) => createCellModel(values, text, formulas, cellDetails, row, col))
    ));
}

async function captureRangeInContext(ctx, sheetName, address) {
    const sheet = ctx.workbook.worksheets.getItem(sheetName);
    const rangeAddress = stripSheetPrefix(address);
    const range = sheet.getRange(rangeAddress);
    range.load(['address', 'rowCount', 'columnCount', 'values', 'text', 'formulas']);
    await ctx.sync();

    const rowFormats = [];
    const colFormats = [];
    const cellRanges = [];
    const cellBorders = [];
    const mergeAreas = [];

    for (let row = 0; row < range.rowCount; row++) {
        const rowRange = range.getRow(row);
        rowRange.format.load('rowHeight');
        rowFormats.push(rowRange.format);
    }

    for (let col = 0; col < range.columnCount; col++) {
        const colRange = range.getColumn(col);
        colRange.format.load('columnWidth');
        colFormats.push(colRange.format);
    }

    for (let row = 0; row < range.rowCount; row++) {
        cellRanges[row] = [];
        cellBorders[row] = [];
        mergeAreas[row] = [];

        for (let col = 0; col < range.columnCount; col++) {
            const cell = range.getCell(row, col);
            cell.load('isMerged');
            cell.format.load(['horizontalAlignment', 'verticalAlignment', 'wrapText']);
            cell.format.fill.load('color');
            cell.format.font.load(['bold', 'color', 'italic', 'name', 'size', 'strikethrough', 'underline']);

            const borders = {};
            for (const [side, officeSide] of Object.entries(BORDER_SIDES)) {
                const border = cell.format.borders.getItem(officeSide);
                border.load(['color', 'style', 'weight']);
                borders[side] = border;
            }

            cellRanges[row][col] = cell;
            cellBorders[row][col] = borders;
            mergeAreas[row][col] = cell.getMergedAreasOrNullObject();
            mergeAreas[row][col].load(['address', 'isNullObject']);
        }
    }

    await ctx.sync();

    const cellDetails = Array.from({ length: range.rowCount }, (_, row) => (
        Array.from({ length: range.columnCount }, (_, col) => {
            const cell = cellRanges[row][col];
            const mergeArea = mergeAreas[row][col];
            const isMerged = Boolean(cell.isMerged);

            return {
                isMerged,
                mergeArea: isMerged && !mergeArea.isNullObject
                    ? { address: normalizeAddress(mergeArea.address) }
                    : null,
                format: readCellFormat(
                    cell.format,
                    Object.fromEntries(Object.entries(cellBorders[row][col]).map(([side, border]) => [side, readBorder(border)])),
                ),
            };
        })
    ));

    return {
        sheetName,
        address: range.address,
        rowCount: range.rowCount,
        colCount: range.columnCount,
        rowHeights: rowFormats.map((format) => format.rowHeight),
        columnWidths: colFormats.map((format) => format.columnWidth),
        cells: buildCells(range.values, range.text, range.formulas, cellDetails),
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
        sheet.activate();
        const cell = sheet.getRange(tab.address).getCell(row, col);
        cell.select();
        await ctx.sync();
    });
}

export async function selectSourceRange(tab, selection) {
    const sourceAddress = selectionToSourceAddress(tab, selection);
    if (!sourceAddress) return;

    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(tab.sheetName);
        sheet.activate();
        sheet.getRange(stripSheetPrefix(sourceAddress)).select();
        await ctx.sync();
    });
}

export async function writeRange(tab, startRow, startCol, values) {
    const rowCount = values.length;
    const colCount = values[0]?.length ?? 0;
    if (!rowCount || !colCount) return null;

    const sourceAddress = selectionToSourceAddress(tab, {
        startRow,
        startCol,
        endRow: startRow + rowCount - 1,
        endCol: startCol + colCount - 1,
    });
    if (!sourceAddress) return null;

    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(tab.sheetName);
        const range = sheet.getRange(stripSheetPrefix(sourceAddress));
        range.values = values;
        await ctx.sync();
        return { address: sourceAddress, values };
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

    eventHandlers.set(tab.id, { sheetName: tab.sheetName, handler: handlerFn });

    return () => deregisterChangeListener(tab.id);
}

export async function writeCell(tab, row, col, value) {
    return Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(tab.sheetName);
        const cell = sheet.getRange(tab.address).getCell(row, col);
        if (typeof value === 'string' && value.startsWith('=')) {
            cell.formulas = [[value]];
        } else {
            cell.values = [[value]];
        }
        cell.load(['values', 'text']);
        await ctx.sync();
        return { value: cell.values[0][0], text: cell.text[0][0] };
    });
}

export async function deregisterChangeListener(tabId) {
    const entry = eventHandlers.get(tabId);
    if (!entry) return;
    try {
        await Excel.run(async (ctx) => {
            // Office.js removes by reference
            const sheet = ctx.workbook.worksheets.getItem(entry.sheetName);
            sheet.onChanged.remove(entry.handler);
            await ctx.sync();
        });
    } catch {
        // Ignore cleanup errors
    }
    eventHandlers.delete(tabId);
}
