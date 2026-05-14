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
    document.getElementById('capture-btn').textContent = t('captureBtn');
    document.getElementById('status-text').textContent = t('synced');
    document.getElementById('step1-label').textContent = t('step1Label');
    document.getElementById('step2-label').textContent = t('step2Label');
    document.getElementById('shortcut-hint').textContent = t('shortcutHint');

    gridRenderer = new GridRenderer(
        document.getElementById('grid-container'),
        () => setSyncState('busy'),
        () => setSyncState('ok'),
        () => { setSyncState('error'); showNotification(t('syncError'), 'error'); },
        (tabId, selection) => tabManager.updateTabSelection(tabId, selection),
    );

    tabManager.onChange(renderAll);

    registerSelectionTracker().catch((err) => {
        console.warn('Selection tracking unavailable:', err);
    });

    document.getElementById('capture-btn').addEventListener('click', () => doCapture({ preferTrackedSelection: false }));
    document.getElementById('add-tab-btn').addEventListener('click', () => {
        const id = tabManager.addEmptyTab();
        if (!id) showNotification(t('maxTabsReached'), 'warning');
    });

    renderAll();
});

async function doCapture(options = {}) {
    setSyncState('busy');
    try {
        const tabData = await captureSelection(options);
        const activeTab = tabManager.getActiveTab();

        if (activeTab && !activeTab.cells) {
            // Fill the existing empty tab
            tabManager.fillTab(activeTab.id, tabData);
            const filledTab = tabManager.getActiveTab();
            const cleanup = await registerChangeListener(filledTab, (changedAddress) => {
                handleExcelChange(activeTab.id, changedAddress);
            });
            listenerCleanups.set(activeTab.id, cleanup);
        } else {
            // Create a new tab
            const tabId = tabManager.addTab(tabData);
            if (!tabId) {
                showNotification(t('maxTabsReached'), 'warning');
                setSyncState('ok');
                return;
            }
            const cleanup = await registerChangeListener({ ...tabData, id: tabId }, (changedAddress) => {
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
        tabManager.updateTabData(tabId, refreshed);
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
    const ctxEl = document.getElementById('active-context');

    if (activeTab && activeTab.cells) {
        const label = activeTab.address.includes('!') ? activeTab.address.split('!')[1] : activeTab.address;
        emptyState.style.display = 'none';
        gridRenderer.render(activeTab);
        rangeInfo.textContent = `${activeTab.sheetName}!${label}`;
        ctxEl.textContent = `${activeTab.sheetName}!${label}`;
    } else {
        // No tabs or active tab is empty — show capture UI
        emptyState.style.display = '';
        document.getElementById('grid-container').innerHTML = '';
        document.getElementById('grid-container').appendChild(emptyState);
        rangeInfo.textContent = t('noRangeSelected');
        ctxEl.textContent = activeTab ? t('newTab') : 'CellFocus';
    }
}

function renderTabBar() {
    const tabBar = document.getElementById('tab-bar');
    const addBtn = document.getElementById('add-tab-btn');

    tabBar.querySelectorAll('.tab').forEach(el => el.remove());

    for (const tab of tabManager.getAllTabs()) {
        const label = tab.address
            ? (tab.address.includes('!') ? tab.address.split('!')[1] : tab.address)
            : t('newTab');
        const fullLabel = tab.sheetName ? `${tab.sheetName}!${label}` : label;

        const tabEl = document.createElement('button');
        tabEl.className = `tab${tab.id === tabManager.activeTabId ? ' active' : ''}`;

        const labelSpan = document.createElement('span');
        labelSpan.className = 'tab-label';
        labelSpan.title = fullLabel;
        labelSpan.textContent = fullLabel;

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

function showNotification(message, type = 'warning') {
    const bar = document.getElementById('notification-bar');
    bar.textContent = message;
    bar.className = `notification-bar visible${type === 'error' ? ' error' : ''}`;
    setTimeout(() => { bar.className = 'notification-bar'; }, 4000);
}
