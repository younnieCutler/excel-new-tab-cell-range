import { generateId } from './utils.js';
import { t } from './i18n.js';

const MAX_TABS = 8;

export class TabManager {
    constructor() {
        this.tabs = new Map();    // id → tabData
        this.activeTabId = null;
        this._listeners = [];     // change callbacks
    }

    addTab(tabData) {
        if (this.tabs.size >= MAX_TABS) {
            alert(t('maxTabsReached'));
            return null;
        }
        const id = generateId();
        this.tabs.set(id, { ...tabData, id });
        this.activeTabId = id;
        this._notify();
        return id;
    }

    removeTab(id) {
        this.tabs.delete(id);
        if (this.activeTabId === id) {
            const remaining = [...this.tabs.keys()];
            this.activeTabId = remaining[remaining.length - 1] ?? null;
        }
        this._notify();
    }

    switchTab(id) {
        if (this.tabs.has(id)) {
            this.activeTabId = id;
            this._notify();
        }
    }

    updateTabCells(id, cells) {
        const tab = this.tabs.get(id);
        if (tab) {
            tab.cells = cells;
            this._notify();
        }
    }

    getActiveTab() {
        return this.activeTabId ? this.tabs.get(this.activeTabId) : null;
    }

    getAllTabs() {
        return [...this.tabs.values()];
    }

    onChange(cb) {
        this._listeners.push(cb);
    }

    _notify() {
        for (const cb of this._listeners) cb();
    }
}
