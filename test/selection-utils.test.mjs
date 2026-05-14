import assert from 'node:assert/strict';
import test from 'node:test';

import {
    buildTsvFromSelection,
    parseTsv,
    selectionToSourceAddress,
} from '../src/taskpane/modules/utils.js';

test('selectionToSourceAddress maps a relative CellFocus selection back to the Excel source range', () => {
    assert.equal(
        selectionToSourceAddress(
            { sheetName: 'Sheet1', address: 'Sheet1!B24:D29' },
            { startRow: 1, startCol: 1, endRow: 3, endCol: 2 },
        ),
        'Sheet1!C25:D27',
    );
});

test('selectionToSourceAddress normalizes reverse drag selections', () => {
    assert.equal(
        selectionToSourceAddress(
            { sheetName: 'Sheet1', address: 'Sheet1!B24:D29' },
            { startRow: 3, startCol: 2, endRow: 1, endCol: 1 },
        ),
        'Sheet1!C25:D27',
    );
});

test('buildTsvFromSelection copies visible text from the selected rectangle', () => {
    const tab = {
        cells: [
            [{ text: 'A' }, { text: 'B' }, { text: 'C' }],
            [{ text: '1' }, { text: '2' }, { value: 3 }],
        ],
    };

    assert.equal(
        buildTsvFromSelection(tab, { startRow: 0, startCol: 1, endRow: 1, endCol: 2 }),
        'B\tC\n2\t3',
    );
});

test('parseTsv parses spreadsheet clipboard rows and columns', () => {
    assert.deepEqual(parseTsv('A\tB\r\n1\t2\n'), [
        ['A', 'B'],
        ['1', '2'],
    ]);
});
