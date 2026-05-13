import './taskpane.css';
import { t } from './modules/i18n.js';
import { TabManager } from './modules/tabManager.js';
import { GridRenderer } from './modules/gridRenderer.js';
import { captureSelection, captureRange, registerChangeListener, deregisterChangeListener, registerSelectionTracker } from './modules/syncEngine.js';

const tabManager = new TabManager();
let gridRenderer = null;
// Cleanup functions for change listeners: tabId → cleanup fn
const listenerCleanups = new Map();

// Expose for commands.js (Shared Runtime)
window.captureSelectedRange = doCapture;

Office.onReady((info) => {
    if (info.host !== Office.HostType.Excel) return;

    // Init i18n text
    document.getElementById('empty-message').textContent = t('emptyStateMessage');
    document.getElementById('capture-btn').textContent = t('captureBtn');
    document.getElementById('status-text').textContent = t('synced');

    gridRenderer = new GridRenderer(
        document.getElementById('grid-container'),
        () => setSyncState('busy'),
        () => setSyncState('ok'),
        () => setSyncState('error'),
    );

    tabManager.onChange(renderAll);

    registerSelectionTracker().catch((err) => {
        console.warn('Selection tracking unavailable:', err);
    });

    document.getElementById('capture-btn').addEventListener('click', () => doCapture({ preferTrackedSelection: true }));
    document.getElementById('add-tab-btn').addEventListener('click', () => doCapture({ preferTrackedSelection: true }));

    renderAll();
});

async function doCapture(options = {}) {
    setSyncState('busy');
    try {
        const tabData = await captureSelection(options);
        const tabId = tabManager.addTab(tabData);
        if (tabId) {
            const cleanup = await registerChangeListener(tabData, (changedAddress) => {
                handleExcelChange(tabId, changedAddress);
            });
            listenerCleanups.set(tabId, cleanup);
        }
        setSyncState('ok');
    } catch (err) {
        console.error('Capture failed:', err);
        setSyncState('error');
    }
}

async function handleExcelChange(tabId, changedAddress) {
    const tab = tabManager.tabs.get(tabId);
    if (!tab) return;
    try {
        const refreshed = await captureRange(tab.sheetName, tab.address);
        tabManager.updateTabCells(tabId, refreshed.cells);
        if (tabManager.activeTabId === tabId) {
            gridRenderer.render(tabManager.getActiveTab());
        }
    } catch (err) {
        console.error('Refresh failed:', err);
    }
}

function renderAll() {
    renderTabBar();
    const activeTab = tabManager.getActiveTab();
    const emptyState = document.getElementById('empty-state');
    const rangeInfo = document.getElementById('range-info');

    if (activeTab) {
        emptyState.style.display = 'none';
        gridRenderer.render(activeTab);
        rangeInfo.textContent = `${activeTab.sheetName}!${activeTab.address.includes('!') ? activeTab.address.split('!')[1] : activeTab.address}`;
    } else {
        emptyState.style.display = '';
        document.getElementById('grid-container').innerHTML = '';
        document.getElementById('grid-container').appendChild(emptyState);
        rangeInfo.textContent = t('noRangeSelected');
    }
}

function renderTabBar() {
    const tabBar = document.getElementById('tab-bar');
    const addBtn = document.getElementById('add-tab-btn');

    tabBar.querySelectorAll('.tab').forEach(el => el.remove());

    for (const tab of tabManager.getAllTabs()) {
        const label = tab.address.includes('!')
            ? tab.address.split('!')[1]
            : tab.address;

        const tabEl = document.createElement('button');
        tabEl.className = `tab${tab.id === tabManager.activeTabId ? ' active' : ''}`;

        const labelSpan = document.createElement('span');
        labelSpan.className = 'tab-label';
        labelSpan.title = `${tab.sheetName}!${label}`;
        labelSpan.textContent = `${tab.sheetName}!${label}`;

        const closeBtn = document.createElement('button');
        closeBtn.className = 'tab-close';
        closeBtn.title = t('closeTab');
        closeBtn.textContent = '✕';
        closeBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            closeTab(tab.id);
        });

        tabEl.appendChild(labelSpan);
        tabEl.appendChild(closeBtn);
        tabEl.addEventListener('click', () => tabManager.switchTab(tab.id));
        tabBar.insertBefore(tabEl, addBtn);
    }
}

function closeTab(tabId) {
    const cleanup = listenerCleanups.get(tabId);
    if (cleanup) {
        cleanup();
        listenerCleanups.delete(tabId);
    }
    tabManager.removeTab(tabId);
}

function setSyncState(state) {
    const dot = document.getElementById('status-dot');
    const text = document.getElementById('status-text');
    dot.className = 'status-dot' + (state !== 'ok' ? ` ${state}` : '');
    text.textContent = state === 'ok' ? t('synced') : state === 'busy' ? t('syncing') : t('syncError');
}
