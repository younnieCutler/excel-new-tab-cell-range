import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

async function loadSyncEngineWithExcelMock(excelMock) {
    const context = vm.createContext({
        Excel: excelMock,
        console,
        Date,
    });

    const source = readFileSync(new URL('../src/taskpane/modules/syncEngine.js', import.meta.url), 'utf8');
    const syncEngine = new vm.SourceTextModule(source, {
        context,
        identifier: 'syncEngine.js',
    });
    const utils = new vm.SourceTextModule(
        `
        export function isAddressWithinRange() { return true; }
        function colToName(index) {
            let n = index + 1;
            let name = '';
            while (n > 0) {
                const rem = (n - 1) % 26;
                name = String.fromCharCode(65 + rem) + name;
                n = Math.floor((n - 1) / 26);
            }
            return name;
        }
        function parseAddress(address) {
            const range = address.includes('!') ? address.split('!')[1] : address;
            const match = range.match(/^([A-Z]+)(\\d+):([A-Z]+)(\\d+)$/);
            const colToIdx = (col) => [...col].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
            return { startCol: colToIdx(match[1]), startRow: Number(match[2]) - 1 };
        }
        export function selectionToSourceAddress(tab, selection) {
            const source = parseAddress(tab.address);
            const startRow = source.startRow + Math.min(selection.startRow, selection.endRow);
            const startCol = source.startCol + Math.min(selection.startCol, selection.endCol);
            const endRow = source.startRow + Math.max(selection.startRow, selection.endRow);
            const endCol = source.startCol + Math.max(selection.startCol, selection.endCol);
            return tab.sheetName + '!' + colToName(startCol) + (startRow + 1) + ':' + colToName(endCol) + (endRow + 1);
        }
        `,
        { context, identifier: 'utils.js' },
    );

    await syncEngine.link(async (specifier) => {
        if (specifier === './utils.js') return utils;
        throw new Error(`Unexpected import: ${specifier}`);
    });
    await utils.evaluate();
    await syncEngine.evaluate();
    return syncEngine.namespace;
}

function createExcelMock() {
    const sheets = new Map();

    function getSheet(name) {
        if (!sheets.has(name)) {
            const handlers = [];
            sheets.set(name, {
                handlers,
                activated: 0,
                selectedCells: [],
                selectedRanges: [],
                writtenRanges: [],
                activate() {
                    this.activated += 1;
                },
                getRange(address) {
                    const worksheet = this;
                    return {
                        values: null,
                        select: () => {
                            this.selectedRanges.push(address);
                        },
                        getCell: (row, col) => ({
                            select: () => {
                                this.selectedCells.push({ address, row, col });
                            },
                        }),
                        set values(matrix) {
                            worksheet.writtenRanges.push({ address, values: matrix });
                        },
                        get values() {
                            return null;
                        },
                    };
                },
                onChanged: {
                    add(handler) {
                        handlers.push(handler);
                    },
                    remove(handler) {
                        const idx = handlers.indexOf(handler);
                        if (idx !== -1) handlers.splice(idx, 1);
                    },
                },
            });
        }
        return sheets.get(name);
    }

    return {
        sheets,
        Excel: {
            run: async (callback) => callback({
                workbook: {
                    worksheets: {
                        getItem: getSheet,
                    },
                },
                sync: async () => {},
            }),
        },
    };
}

test('registerChangeListener cleans up the listener for each tab on its source worksheet', async () => {
    const mock = createExcelMock();
    const { registerChangeListener } = await loadSyncEngineWithExcelMock(mock.Excel);

    const cleanupA = await registerChangeListener({ id: 'tab-a', sheetName: 'Sheet1', address: 'Sheet1!A1:B2' }, () => {});
    const cleanupB = await registerChangeListener({ id: 'tab-b', sheetName: 'Sheet2', address: 'Sheet2!C3:D4' }, () => {});

    assert.equal(mock.sheets.get('Sheet1').handlers.length, 1);
    assert.equal(mock.sheets.get('Sheet2').handlers.length, 1);

    await cleanupA();

    assert.equal(mock.sheets.get('Sheet1').handlers.length, 0);
    assert.equal(mock.sheets.get('Sheet2').handlers.length, 1);

    await cleanupB();

    assert.equal(mock.sheets.get('Sheet1').handlers.length, 0);
    assert.equal(mock.sheets.get('Sheet2').handlers.length, 0);
});

test('selectSourceCell activates the tab worksheet before selecting the source cell', async () => {
    const mock = createExcelMock();
    const { selectSourceCell } = await loadSyncEngineWithExcelMock(mock.Excel);

    await selectSourceCell({ sheetName: 'Sheet2', address: 'Sheet2!C3:D4' }, 1, 0);

    assert.equal(mock.sheets.get('Sheet2').activated, 1);
    assert.deepEqual(mock.sheets.get('Sheet2').selectedCells, [
        { address: 'Sheet2!C3:D4', row: 1, col: 0 },
    ]);
});

test('selectSourceRange activates the tab worksheet and selects the matching source range', async () => {
    const mock = createExcelMock();
    const { selectSourceRange } = await loadSyncEngineWithExcelMock(mock.Excel);

    await selectSourceRange(
        { sheetName: 'Sheet2', address: 'Sheet2!B24:D29' },
        { startRow: 1, startCol: 1, endRow: 3, endCol: 2 },
    );

    assert.equal(mock.sheets.get('Sheet2').activated, 1);
    assert.deepEqual(mock.sheets.get('Sheet2').selectedRanges, ['C25:D27']);
});

test('writeRange writes a matrix to the matching source range', async () => {
    const mock = createExcelMock();
    const { writeRange } = await loadSyncEngineWithExcelMock(mock.Excel);

    await writeRange(
        { sheetName: 'Sheet1', address: 'Sheet1!B24:D29' },
        1,
        1,
        [['A', 'B'], ['1', '2']],
    );

    assert.deepEqual(mock.sheets.get('Sheet1').writtenRanges, [
        { address: 'C25:D26', values: [['A', 'B'], ['1', '2']] },
    ]);
});
