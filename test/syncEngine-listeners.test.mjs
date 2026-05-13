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
        'export function isAddressWithinRange() { return true; }',
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
                activate() {
                    this.activated += 1;
                },
                getRange(address) {
                    return {
                        getCell: (row, col) => ({
                            select: () => {
                                this.selectedCells.push({ address, row, col });
                            },
                        }),
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
