const LOCALES = {
    'ja-JP': {
        openInCellFocus: 'CellFocusで開く',
        synced: '同期済み',
        syncing: '同期中...',
        syncError: '同期エラー',
        maxTabsReached: 'タブは最大8つまでです',
        closeTab: '閉じる',
        addRange: '範囲を追加',
        noRangeSelected: '範囲が選択されていません',
        emptyStateMessage: '範囲を選択して「CellFocusで開く」を実行してください',
        captureBtn: '選択範囲を取り込む',
        step1Label: 'Excelでセル範囲を選択',
        step2Label: '下のボタンをクリック（右クリック→CellFocusで開く）',
        shortcutHint: 'ショートカット: Ctrl+Shift+F',
        newTab: '新規',
    },
    'en-US': {
        openInCellFocus: 'Open in CellFocus',
        synced: 'Synced',
        syncing: 'Syncing...',
        syncError: 'Sync error',
        maxTabsReached: 'Maximum 8 tabs reached',
        closeTab: 'Close',
        addRange: 'Add Range',
        noRangeSelected: 'No range selected',
        emptyStateMessage: 'Select a range and click "Open in CellFocus"',
        captureBtn: 'Capture Selection',
        step1Label: 'Select a range in Excel',
        step2Label: 'Click the button below (or right-click → Open in CellFocus)',
        shortcutHint: 'Shortcut: Ctrl+Shift+F',
        newTab: 'New Tab',
    },
};

function getLocale() {
    try {
        const lang = Office.context.displayLanguage ?? 'en-US';
        return lang.startsWith('ja') ? 'ja-JP' : 'en-US';
    } catch {
        return 'ja-JP';
    }
}

export function t(key) {
    const locale = getLocale();
    return (LOCALES[locale] ?? LOCALES['en-US'])[key] ?? key;
}
