export function generateId() {
    return typeof crypto !== 'undefined' && crypto.randomUUID
        ? crypto.randomUUID()
        : `${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

// Parse "Sheet1!A1:D10" or "A1:D10" or "$A$1:$D$10"
// Returns { sheet, startRow, startCol, endRow, endCol } (all 0-indexed)
export function parseAddress(address) {
    let sheet = null;
    let rangeStr = address;

    const sheetSep = address.indexOf('!');
    if (sheetSep !== -1) {
        sheet = address.slice(0, sheetSep).replace(/^'|'$/g, '');
        rangeStr = address.slice(sheetSep + 1);
    }

    rangeStr = rangeStr.replace(/\$/g, '');
    const match = rangeStr.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
    if (!match) return null;

    const colToIdx = (col) => {
        let n = 0;
        for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
        return n - 1;
    };

    const startCol = colToIdx(match[1]);
    const startRow = parseInt(match[2]) - 1;
    const endCol = match[3] ? colToIdx(match[3]) : startCol;
    const endRow = match[4] ? parseInt(match[4]) - 1 : startRow;

    return { sheet, startRow, startCol, endRow, endCol };
}

export function columnIndexToName(index) {
    let n = index + 1;
    let name = '';
    while (n > 0) {
        const rem = (n - 1) % 26;
        name = String.fromCharCode(65 + rem) + name;
        n = Math.floor((n - 1) / 26);
    }
    return name;
}

export function normalizeSelection(selection) {
    if (!selection) return null;
    return {
        startRow: Math.min(selection.startRow, selection.endRow),
        startCol: Math.min(selection.startCol, selection.endCol),
        endRow: Math.max(selection.startRow, selection.endRow),
        endCol: Math.max(selection.startCol, selection.endCol),
    };
}

export function selectionToSourceAddress(tab, selection) {
    const source = parseAddress(tab.address);
    const normalized = normalizeSelection(selection);
    if (!source || !normalized) return null;

    const startRow = source.startRow + normalized.startRow;
    const startCol = source.startCol + normalized.startCol;
    const endRow = source.startRow + normalized.endRow;
    const endCol = source.startCol + normalized.endCol;
    const rangeAddress = `${columnIndexToName(startCol)}${startRow + 1}:${columnIndexToName(endCol)}${endRow + 1}`;
    const sheetName = tab.sheetName ?? source.sheet;
    return sheetName ? `${sheetName}!${rangeAddress}` : rangeAddress;
}

export function buildTsvFromSelection(tab, selection) {
    const normalized = normalizeSelection(selection);
    if (!normalized) return '';

    const rows = [];
    for (let row = normalized.startRow; row <= normalized.endRow; row++) {
        const cols = [];
        for (let col = normalized.startCol; col <= normalized.endCol; col++) {
            const cell = tab.cells?.[row]?.[col];
            cols.push(String(cell?.text ?? cell?.value ?? ''));
        }
        rows.push(cols.join('\t'));
    }
    return rows.join('\n');
}

export function parseTsv(text) {
    return String(text ?? '')
        .replace(/\r\n/g, '\n')
        .replace(/\r/g, '\n')
        .replace(/\n$/, '')
        .split('\n')
        .map((row) => row.split('\t'));
}

// Returns true if changedAddress intersects watchAddress
// Both may include sheet prefix (e.g. "Sheet1!B2:C3")
export function isAddressWithinRange(changedAddress, watchAddress) {
    const changed = parseAddress(changedAddress);
    const watch = parseAddress(watchAddress);
    if (!changed || !watch) return false;

    // Sheet filter: if both have sheets, must match
    if (changed.sheet && watch.sheet && changed.sheet !== watch.sheet) return false;

    return (
        changed.startRow <= watch.endRow &&
        changed.endRow >= watch.startRow &&
        changed.startCol <= watch.endCol &&
        changed.endCol >= watch.startCol
    );
}

// Convert Excel border weight string to CSS pixel value
function borderWeightToPx(weight) {
    const map = { Hairline: '0.5px', Thin: '1px', Medium: '2px', Thick: '3px' };
    return map[weight] ?? '1px';
}

// Convert Excel border style string to CSS border-style
function borderStyleToCss(style) {
    if (!style || style === 'None') return 'none';
    const map = {
        Continuous: 'solid',
        Dash: 'dashed',
        DashDot: 'dashed',
        DashDotDot: 'dashed',
        Dot: 'dotted',
        Double: 'double',
        SlantDashDot: 'dashed',
    };
    return map[style] ?? 'solid';
}

// Convert Office.js cell format object to a CSS style string
export function buildCellStyle(format) {
    if (!format) return '';
    const styles = [];

    const fill = format.fill;
    if (fill?.color && fill.color !== 'transparent') {
        styles.push(`background-color:${fill.color}`);
    }

    const font = format.font;
    if (font) {
        if (font.bold) styles.push('font-weight:bold');
        if (font.italic) styles.push('font-style:italic');
        if (font.size) styles.push(`font-size:${font.size}pt`);
        if (font.name) styles.push(`font-family:"${font.name}",sans-serif`);
        if (font.color) styles.push(`color:${font.color}`);
        if (font.strikethrough) styles.push('text-decoration:line-through');
        if (font.underline && font.underline !== 'None') styles.push('text-decoration:underline');
    }

    const hAlign = format.horizontalAlignment;
    if (hAlign) {
        const map = { Left: 'left', Center: 'center', Right: 'right', Fill: 'left', Justify: 'justify' };
        styles.push(`text-align:${map[hAlign] ?? 'left'}`);
    }

    const vAlign = format.verticalAlignment;
    if (vAlign) {
        const map = { Top: 'top', Center: 'middle', Bottom: 'bottom', Justify: 'middle' };
        styles.push(`vertical-align:${map[vAlign] ?? 'middle'}`);
    }

    if (format.wrapText) styles.push('white-space:pre-wrap');

    const borders = format.borders;
    if (borders) {
        const sides = { top: 'border-top', bottom: 'border-bottom', left: 'border-left', right: 'border-right' };
        for (const [key, css] of Object.entries(sides)) {
            const b = borders[key];
            if (b && b.style && b.style !== 'None') {
                styles.push(`${css}:${borderWeightToPx(b.weight)} ${borderStyleToCss(b.style)} ${b.color ?? '#000'}`);
            }
        }
    }

    return styles.join(';');
}
