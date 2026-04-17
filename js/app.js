/**
 * Main application logic for the Accounting Journal Calculator
 */

const App = {
    // ==================== UNDO/REDO SYSTEM ====================
    _undoStack: [],
    _redoStack: [],
    _undoMaxSize: 50,

    /**
     * Push an action onto the undo stack
     * @param {Object} action - { type, label, undo: Function, redo: Function }
     */
    pushUndo(action) {
        this._undoStack.push(action);
        if (this._undoStack.length > this._undoMaxSize) this._undoStack.shift();
        this._redoStack = []; // clear redo on new action
    },

    undo() {
        const action = this._undoStack.pop();
        if (!action) return;
        try {
            action.undo();
            this._redoStack.push(action);
            this._showUndoToast('Undid: ' + action.label, true);
            this.refreshAll();
        } catch (e) {
            console.error('Undo failed:', e);
            UI.showNotification('Undo failed', 'error');
        }
    },

    redo() {
        const action = this._redoStack.pop();
        if (!action) return;
        try {
            action.redo();
            this._undoStack.push(action);
            this._showUndoToast('Redid: ' + action.label, false);
            this.refreshAll();
        } catch (e) {
            console.error('Redo failed:', e);
            UI.showNotification('Redo failed', 'error');
        }
    },

    _showUndoToast(message, showRedo) {
        const existing = document.querySelector('.undo-toast');
        if (existing) existing.remove();

        const toast = document.createElement('div');
        toast.className = 'undo-toast';
        toast.innerHTML = '<span class="undo-toast-msg">' + this._escapeHtml(message) + '</span>' +
            (showRedo ? '<button class="undo-toast-btn" onclick="App.redo()">Redo</button>' : '') +
            (!showRedo ? '<button class="undo-toast-btn" onclick="App.undo()">Undo</button>' : '');

        document.body.appendChild(toast);
        toast.offsetHeight; // reflow
        toast.classList.add('visible');

        setTimeout(() => {
            toast.classList.remove('visible');
            setTimeout(() => toast.remove(), 300);
        }, 4000);
    },

    clearUndoHistory() {
        this._undoStack = [];
        this._redoStack = [];
    },

    deleteTargetId: null,
    deleteCategoryTargetId: null,
    deleteFolderTargetId: null,
    deleteAssetTargetId: null,
    deleteLoanTargetId: null,
    deleteB2BContractTargetId: null,
    selectedB2BContractId: null,
    deleteBudgetExpenseTargetId: null,
    deleteProductTargetId: null,
    selectedAssetId: null,
    selectedLoanId: null,
    selectedBudgetExpenseId: null,
    collapsedBudgetGroups: new Set(),
    folderCreatedFromCategory: false,
    pendingFileLoad: null,
    pendingLoadBuffer: null,
    pendingInlineStatusChange: null, // {id, newStatus, selectElement}
    bulkSelectMode: false,
    bulkSelectDirection: 'to-paid', // 'to-paid' or 'to-pending'
    bulkSelectedIds: new Set(),
    pendingBulkAction: false,
    currentMode: 'work',
    lastWorkTab: 'journal',
    lastAnalyzeTab: 'dashboard',
    currentSortMode: 'entryDate',
    _timeline: null,
    _beChartBreakeven: null,
    _beChartTimeline: null,
    _beChartProgress: null,
    _beProgressState: null,
    _pvmChartUnits: null,
    _pvmChartRevenue: null,
    _syncAutoSaveWrapped: false,
    _rollbackTargetVersion: null,
    calcMode: false,
    calcRefCounter: 0,

    // Theme preset palettes: { c1: sidebar, c2: accent, c3: background, c4: surface, c5: border, c6: text, style?: string }
    themePresets: {
        // Color-only themes
        default:  { c1: '#1e2530', c2: '#3b82f6', c3: '#f3f4f6', c4: '#ffffff', c5: '#e5e7eb', c6: '#1f2937' },
        ocean:    { c1: '#1a2e3a', c2: '#0ea5e9', c3: '#f0f7fa', c4: '#ffffff', c5: '#d1e3eb', c6: '#1e3a4a' },
        forest:   { c1: '#1a2e24', c2: '#10b981', c3: '#f0f7f0', c4: '#ffffff', c5: '#c6dcd0', c6: '#1a3028' },
        sunset:   { c1: '#2e1f1a', c2: '#e76f51', c3: '#fdf8f4', c4: '#ffffff', c5: '#e8d5c8', c6: '#3d2518' },
        midnight: { c1: '#1e1a30', c2: '#6c63ff', c3: '#f5f4ff', c4: '#ffffff', c5: '#d5d0e8', c6: '#2a2440' },
        // Extreme design styles (CSS handles font, radius, shadows via data-theme-style)
        modern:     { c1: '#0f172a', c2: '#3b82f6', c3: '#f8fafc', c4: '#ffffff', c5: '#e2e8f0', c6: '#0f172a', style: 'modern' },
        futuristic: { c1: '#060c18', c2: '#00d4ff', c3: '#080e1a', c4: '#111827', c5: '#1e293b', c6: '#e2e8f0', style: 'futuristic' },
        vintage:    { c1: '#3d2e20', c2: '#a0764e', c3: '#faf6f0', c4: '#fffdf8', c5: '#d6c8b4', c6: '#3d2e20', style: 'vintage' },
        accounting: { c1: '#1e3a5f', c2: '#2563eb', c3: '#ffffff', c4: '#ffffff', c5: '#d1d5db', c6: '#1e3a5f', style: 'accounting' },
    },

    /**
     * Initialize the application
     */
    isViewOnly: false,

    async init() {
        try {
            // Restore sidebar state early to prevent flash
            if (localStorage.getItem('sidebarCollapsed') === 'true') {
                document.querySelector('.app-container').classList.add('sidebar-collapsed');
            }

            document.body.style.opacity = '0.5';

            // Check for share token in URL hash — enter view-only mode
            const shareMatch = window.location.hash.match(/^#share=(.+)$/);
            if (shareMatch) {
                await this.initViewOnlyMode(shareMatch[1]);
                return;
            }

            await Database.init();

            // First-run: prompt user to name the migrated company
            if (CompanyManager.needsNamingPrompt()) {
                await this.promptInitialCompanyName();
            }

            // Render company switcher in header
            CompanyManager.renderSwitcher();

            // Set up initial UI state
            document.getElementById('entryDate').value = Utils.getTodayDate();

            // Populate dropdowns
            UI.populateYearDropdowns();
            UI.populatePaymentForMonthDropdown();

            // Load journal owner and update title
            const owner = Database.getJournalOwner();
            document.getElementById('journalOwner').value = owner;
            UI.updateJournalTitle(owner);

            // Load and apply saved theme
            this.loadAndApplyTheme();

            // Load shipping fee rate
            this.loadShippingFeeRate();

            // Restore tab order, hidden tabs, and set up tab drag-drop
            this.restoreTabOrder();
            this.setupTabDragDrop();
            // Clean stale always-visible tabs from hidden tabs
            const ht = Database.getHiddenTabs();
            if (ht.includes('changelog') || ht.includes('quickguide')) {
                Database.setHiddenTabs(ht.filter(t => t !== 'changelog' && t !== 'quickguide'));
            }
            this.applyHiddenTabs();
            this.setupTabScrollFade();

            // Load and apply timeline
            this.loadAndApplyTimeline();

            // Restore balance sheet date before refreshAll (needs timeline years populated first)
            this.initBalanceSheetDate();

            // Load and render data
            this.refreshAll();

            // Set up event listeners
            this.setupEventListeners();
            this.setupVESalesListeners();

            // Initialize sync if previously configured
            await this.initSync();

            document.body.style.opacity = '1';

            // Initialize tutorial system
            if (typeof Tutorial !== 'undefined') Tutorial.init();

            console.log('Application initialized successfully');
        } catch (error) {
            console.error('Failed to initialize application:', error);
            document.body.style.opacity = '1';
            UI.showNotification('Failed to initialize application. Please refresh the page.', 'error');
        }
    },

    /**
     * Refresh all UI components
     */
    refreshAll() {
        // Skip write-heavy sync operations in view-only mode — they can throw
        // when the database snapshot is missing expected state
        if (!this.isViewOnly) {
            this._syncAllLoanBudgetExpenses();
            this.syncAllBudgetJournalEntries();
            this.syncAllB2BContractEntries();
        }

        // Wrap each refresh in try-catch so one failure doesn't block the rest
        const refreshCalls = [
            () => this.refreshCategories(),
            () => this.refreshTransactions(),
            () => this.refreshSummary(),
            () => this.refreshCashFlow(),
            () => this.refreshPnL(),
            () => this.refreshBalanceSheet(),
            () => this.refreshFixedAssets(),
            () => this.refreshLoans(),
            () => this.refreshBudget(),
            () => this.refreshBreakeven(),
            () => this.refreshProjectedSales(),
            () => this.refreshProducts(),
            () => this.refreshVESales(),
            () => this.refreshB2BContracts(),
        ];
        for (const fn of refreshCalls) {
            try { fn(); } catch (e) { console.error('refreshAll: tab refresh failed:', e); }
        }
    },

    /**
     * Refresh categories in dropdowns
     */
    refreshCategories() {
        const categories = Database.getCategories();
        UI.populateCategoryDropdown(categories);
        UI.populateFilterCategories(categories);
        UI.populateFilterFolders(Database.getFolders());
    },

    /**
     * Refresh transactions list
     */
    refreshTransactions() {
        const filters = this.getActiveFilters();
        const transactions = Database.getTransactions(filters);

        UI.renderTransactions(transactions, this.currentSortMode);

        const allTransactions = Database.getTransactions();
        const months = Utils.getUniqueMonths(allTransactions);
        UI.populateFilterMonths(months);
    },

    /**
     * Refresh summary calculations
     */
    refreshSummary() {
        const summary = Database.calculateSummary();
        UI.updateSummary(summary);
    },

    /**
     * Refresh cash flow spreadsheet tab
     */
    refreshCashFlow() {
        const data = Database.getCashFlowSpreadsheet();
        const timeline = this.getTimeline();
        const currentMonth = Utils.getCurrentMonth();
        const cfOverrides = Database.getAllCashFlowOverrides();

        // Filter months by timeline
        if (timeline.start || timeline.end) {
            data.months = Utils.filterMonthsByTimeline(data.months, timeline.start, timeline.end);
        }

        // Add future months up to timeline end (for projections)
        if (timeline.end && timeline.end > currentMonth) {
            let m = Utils.nextMonth(currentMonth);
            while (m <= timeline.end) {
                if (!data.months.includes(m)) {
                    data.months.push(m);
                }
                m = Utils.nextMonth(m);
            }
            data.months.sort();
        }

        // Get projected sales data for cashflow integration
        let projectedSales = null;
        const psConfig = Database.getProjectedSalesConfig();
        const cfViewToggle = document.getElementById('cfViewToggle');
        if (psConfig.enabled && psConfig.projectionStartMonth) {
            // Ensure continuous month range through at least currentMonth for projections
            const lastNeeded = data.months.length > 0 && data.months[data.months.length - 1] > currentMonth
                ? data.months[data.months.length - 1] : currentMonth;
            let fill = psConfig.projectionStartMonth;
            while (fill <= lastNeeded) {
                if (!data.months.includes(fill)) data.months.push(fill);
                fill = Utils.nextMonth(fill);
            }
            data.months.sort();

            const cfViewModeEl = document.getElementById('cfViewMode');
            const cfAsOfEl = document.getElementById('cfAsOfMonth');
            const psSpreadsheet = Database.getProjectedSalesSpreadsheet(psConfig, data.months);

            // Populate as-of dropdown with timeline months (restore saved value)
            this._populateAsOfSelect(cfAsOfEl, data.months, Database.getAsOfMonth('cf'));

            const asOfVal = cfAsOfEl ? cfAsOfEl.value : 'current';
            projectedSales = {
                enabled: true,
                projectionStartMonth: psConfig.projectionStartMonth,
                byMonth: psSpreadsheet.byMonth,
                channels: psSpreadsheet.channels,
                viewMode: cfViewModeEl ? cfViewModeEl.value : 'projected',
                asOfMonth: asOfVal !== 'current' ? asOfVal : null
            };
            if (cfViewToggle) cfViewToggle.style.display = 'flex';
        } else {
            if (cfViewToggle) cfViewToggle.style.display = 'none';
        }

        UI.renderCashFlowSpreadsheet(data, cfOverrides, currentMonth, projectedSales);
        this.setupCashFlowDragDrop();
        this.setupCashFlowCellEditing();
    },

    /**
     * Set up drag-and-drop for cashflow category rows
     */
    setupCashFlowDragDrop() {
        const container = document.getElementById('cashflowSpreadsheet');
        if (!container || container.dataset.dragSetup) return;
        container.dataset.dragSetup = '1';

        let draggedRow = null;

        container.addEventListener('dragstart', (e) => {
            const row = e.target.closest('tr[draggable="true"]');
            if (!row) return;
            draggedRow = row;
            row.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
            e.dataTransfer.setData('text/plain', row.dataset.categoryId);
        });

        container.addEventListener('dragover', (e) => {
            e.preventDefault();
            const row = e.target.closest('tr[draggable="true"]');
            if (!row || row === draggedRow || !draggedRow) return;
            if (row.dataset.section !== draggedRow.dataset.section) return;
            e.dataTransfer.dropEffect = 'move';
            container.querySelectorAll(`tr[data-section="${draggedRow.dataset.section}"].drag-over`)
                .forEach(r => r.classList.remove('drag-over'));
            row.classList.add('drag-over');
        });

        container.addEventListener('dragleave', (e) => {
            const row = e.target.closest('tr[draggable="true"]');
            if (row) row.classList.remove('drag-over');
        });

        container.addEventListener('drop', (e) => {
            e.preventDefault();
            const targetRow = e.target.closest('tr[draggable="true"]');
            if (!targetRow || !draggedRow || targetRow === draggedRow) return;
            if (targetRow.dataset.section !== draggedRow.dataset.section) return;

            const parent = draggedRow.parentNode;
            parent.insertBefore(draggedRow, targetRow);

            const section = draggedRow.dataset.section;
            const rows = parent.querySelectorAll(`tr[data-section="${section}"]`);
            const orderList = [];
            rows.forEach((row, index) => {
                orderList.push({ id: parseInt(row.dataset.categoryId), sortOrder: index });
            });

            Database.updateCashflowSortOrder(orderList);
            targetRow.classList.remove('drag-over');
        });

        container.addEventListener('dragend', () => {
            if (draggedRow) {
                draggedRow.classList.remove('dragging');
                draggedRow = null;
            }
            container.querySelectorAll('.drag-over').forEach(r => r.classList.remove('drag-over'));
        });
    },

    /**
     * Set up drag and drop for main tab reordering
     */
    setupTabDragDrop() {
        const nav = document.querySelector('.main-tabs');
        if (!nav || nav.dataset.tabDragSetup) return;
        nav.dataset.tabDragSetup = '1';

        let draggedTab = null;

        nav.querySelectorAll('.main-tab[data-tab]').forEach(btn => {
            btn.setAttribute('draggable', 'true');
        });

        nav.addEventListener('dragstart', (e) => {
            const tab = e.target.closest('.main-tab[data-tab]');
            if (!tab) return;
            draggedTab = tab;
            tab.classList.add('tab-dragging');
            e.dataTransfer.effectAllowed = 'move';
            e.dataTransfer.setData('text/plain', tab.dataset.tab);
        });

        nav.addEventListener('dragover', (e) => {
            e.preventDefault();
            const tab = e.target.closest('.main-tab[data-tab]');
            if (!tab || tab === draggedTab || !draggedTab) return;
            e.dataTransfer.dropEffect = 'move';
            nav.querySelectorAll('.main-tab.tab-drag-over').forEach(t => t.classList.remove('tab-drag-over'));
            tab.classList.add('tab-drag-over');
        });

        nav.addEventListener('dragleave', (e) => {
            const tab = e.target.closest('.main-tab');
            if (tab) tab.classList.remove('tab-drag-over');
        });

        nav.addEventListener('drop', (e) => {
            e.preventDefault();
            const targetTab = e.target.closest('.main-tab[data-tab]');
            if (!targetTab || !draggedTab || targetTab === draggedTab) return;

            nav.insertBefore(draggedTab, targetTab);
            targetTab.classList.remove('tab-drag-over');

            // Save new order (keep Info group + changelog + gear at end)
            this._pinInfoGroup(nav);
            const tabs = nav.querySelectorAll('.main-tab[data-tab]');
            const order = Array.from(tabs).map(t => t.dataset.tab);
            Database.setTabOrder(order);
        });

        nav.addEventListener('dragend', () => {
            if (draggedTab) {
                draggedTab.classList.remove('tab-dragging');
                draggedTab = null;
            }
            nav.querySelectorAll('.tab-drag-over').forEach(t => t.classList.remove('tab-drag-over'));
        });
    },

    /**
     * Show fade gradients on tab bar edges when there is more content to scroll
     */
    setupTabScrollFade() {
        const nav = document.getElementById('mainTabs');
        if (!nav) return;
        const wrapper = nav.closest('.main-tabs-wrapper');
        if (!wrapper) return;

        const update = () => {
            const { scrollLeft, scrollWidth, clientWidth } = nav;
            wrapper.classList.toggle('can-scroll-left',  scrollLeft > 1);
            wrapper.classList.toggle('can-scroll-right', scrollLeft + clientWidth < scrollWidth - 1);
        };

        nav.addEventListener('scroll', update, { passive: true });
        // Also update after resize or tab reorder
        new ResizeObserver(update).observe(nav);
        update();
    },

    /**
     * Restore saved tab order from database
     */
    restoreTabOrder() {
        const order = Database.getTabOrder();
        if (!order || !Array.isArray(order)) return;

        const nav = document.querySelector('.main-tabs');
        if (!nav) return;

        order.forEach(tabName => {
            const tab = nav.querySelector(`.main-tab[data-tab="${tabName}"]`);
            if (tab) nav.appendChild(tab);
        });

        // Append any tabs not in the saved order (e.g., newly added tabs)
        nav.querySelectorAll('.main-tab[data-tab]').forEach(tab => {
            if (!order.includes(tab.dataset.tab)) {
                nav.appendChild(tab);
            }
        });

        this._pinInfoGroup(nav);
    },

    _pinInfoGroup(nav) {
        const infoLabel = Array.from(nav.querySelectorAll('.sidebar-nav-group')).find(el => el.textContent.trim() === 'Info');
        const quickguideBtn = nav.querySelector('.main-tab[data-tab="quickguide"]');
        const changelogBtn = nav.querySelector('.main-tab[data-tab="changelog"]');
        const gearBtn = document.getElementById('manageTabsBtn');
        if (infoLabel) nav.appendChild(infoLabel);
        if (quickguideBtn) nav.appendChild(quickguideBtn);
        if (changelogBtn) nav.appendChild(changelogBtn);
        if (gearBtn) nav.appendChild(gearBtn);
    },

    // Tabs that cannot be reset (derived from other data)
    NON_RESETTABLE_TABS: ['balancesheet'],

    TAB_LABELS: {
        journal: 'Journal', cashflow: 'Cash Flow', pnl: 'P&L',
        balancesheet: 'Balance Sheet', assets: 'Assets & Equity',
        loan: 'Loans', budget: 'Budget', breakeven: 'Break-Even',
        projectedsales: 'Projected Sales', products: 'Products', vesales: 'VE Sales',
        b2bcontract: 'B2B Contract',
        dashboard: 'Dashboard', quickguide: 'Quick Guide', changelog: 'Change Log'
    },

    WORK_TABS: ['journal', 'budget', 'products', 'vesales', 'b2bcontract', 'assets', 'loan', 'projectedsales'],
    ANALYZE_TABS: ['dashboard', 'cashflow', 'pnl', 'balancesheet', 'breakeven'],

    applyHiddenTabs() {
        this.applyModeAwareHiddenTabs();
    },

    applyModeAwareHiddenTabs() {
        const hidden = Database.getHiddenTabs();
        const modeTabs = this.currentMode === 'work' ? this.WORK_TABS : this.ANALYZE_TABS;
        const nav = document.querySelector('.main-tabs');

        nav.querySelectorAll('.main-tab[data-tab]').forEach(btn => {
            const tab = btn.dataset.tab;
            if (tab === 'changelog' || tab === 'quickguide') {
                // Always visible — not mode-dependent or hideable
                btn.style.display = '';
                return;
            }
            if (!modeTabs.includes(tab)) {
                btn.style.display = 'none';
            } else if (hidden.includes(tab)) {
                btn.style.display = 'none';
            } else {
                btn.style.display = '';
            }
        });

        // Ensure Info group + changelog are positioned correctly before checking visibility
        this._pinInfoGroup(nav);

        // Hide sidebar group labels that have no visible tabs beneath them
        this.updateSidebarGroupLabels();

        // If active tab is now hidden, switch to first visible
        const activeBtn = nav.querySelector('.main-tab.active');
        if (activeBtn && activeBtn.style.display === 'none') {
            const firstVisible = nav.querySelector('.main-tab[data-tab]:not([style*="display: none"])');
            if (firstVisible) this.switchMainTab(firstVisible.dataset.tab);
        }
    },

    updateSidebarGroupLabels() {
        const nav = document.querySelector('.main-tabs');
        const children = Array.from(nav.children);
        let currentGroup = null;

        for (const child of children) {
            if (child.classList.contains('sidebar-nav-group')) {
                // Finalize previous group
                if (currentGroup) {
                    currentGroup.el.style.display = currentGroup.hasVisible ? '' : 'none';
                }
                currentGroup = { el: child, hasVisible: false };
            } else if (child.classList.contains('main-tab') && child.dataset.tab) {
                if (currentGroup && child.style.display !== 'none') {
                    currentGroup.hasVisible = true;
                }
            }
        }
        // Handle last group
        if (currentGroup) {
            currentGroup.el.style.display = currentGroup.hasVisible ? '' : 'none';
        }
    },

    setupTabContextMenu() {
        const menu = document.getElementById('tabContextMenu');
        let targetTab = null;

        document.querySelector('.main-tabs').addEventListener('contextmenu', (e) => {
            const btn = e.target.closest('.main-tab[data-tab]');
            if (!btn) return;
            e.preventDefault();
            targetTab = btn.dataset.tab;
            menu.style.left = e.clientX + 'px';
            menu.style.top = e.clientY + 'px';
            menu.style.display = 'block';
            document.getElementById('tabCtxReset').style.display =
                this.NON_RESETTABLE_TABS.includes(targetTab) ? 'none' : '';
        });

        document.getElementById('tabCtxHide').addEventListener('click', () => {
            menu.style.display = 'none';
            if (!targetTab || targetTab === 'changelog' || targetTab === 'quickguide') return;
            const hidden = Database.getHiddenTabs();
            if (!hidden.includes(targetTab)) {
                hidden.push(targetTab);
                Database.setHiddenTabs(hidden);
            }
            this.applyHiddenTabs();
        });

        document.getElementById('tabCtxReset').addEventListener('click', () => {
            menu.style.display = 'none';
            if (!targetTab || this.NON_RESETTABLE_TABS.includes(targetTab)) return;
            const label = this.TAB_LABELS[targetTab] || targetTab;
            if (confirm('Reset "' + label + '"? This will delete all data in this tab. This cannot be undone.')) {
                Database.resetTabData(targetTab);
                this.refreshCurrentTab(targetTab);
            }
        });

        document.addEventListener('click', () => { menu.style.display = 'none'; });
    },

    refreshCurrentTab(tab) {
        switch (tab) {
            case 'journal': this.refreshCategories(); this.refreshTransactions(); this.refreshSummary(); break;
            case 'cashflow': this.refreshCashFlow(); break;
            case 'pnl': this.refreshPnL(); break;
            case 'balancesheet': this.refreshBalanceSheet(); break;
            case 'assets': this.refreshFixedAssets(); break;
            case 'loan': this.refreshLoans(); break;
            case 'budget': this.refreshBudget(); break;
            case 'breakeven': this.refreshBreakeven(); break;
            case 'projectedsales': this.refreshProjectedSales(); break;
            case 'products': this.refreshProducts(); break;
            case 'vesales': this.refreshVESales(); break;
            case 'b2bcontract': this.refreshB2BContracts(); break;
            case 'dashboard': this.refreshDashboard(); break;
        }
    },

    openManageTabsModal() {
        const list = document.getElementById('manageTabsList');
        const hidden = Database.getHiddenTabs();
        const modeTabs = this.currentMode === 'work' ? this.WORK_TABS : this.ANALYZE_TABS;
        const allTabs = modeTabs.slice();

        list.innerHTML = allTabs.map(tab => {
            const label = this.TAB_LABELS[tab] || tab;
            const checked = !hidden.includes(tab) ? 'checked' : '';
            return '<label><input type="checkbox" data-tab="' + tab + '" ' + checked + '><span>' + label + '</span></label>';
        }).join('');

        list.querySelectorAll('input[type="checkbox"]').forEach(cb => {
            cb.addEventListener('change', () => {
                let h = Database.getHiddenTabs();
                if (cb.checked) {
                    h = h.filter(t => t !== cb.dataset.tab);
                } else {
                    if (!h.includes(cb.dataset.tab)) h.push(cb.dataset.tab);
                }
                Database.setHiddenTabs(h);
                this.applyHiddenTabs();
            });
        });

        UI.showModal('manageTabsModal');
    },

    /**
     * Set up inline cell editing for Cash Flow projected cells (only binds once)
     */
    setupCashFlowCellEditing() {
        const container = document.getElementById('cashflowSpreadsheet');
        if (!container || container.dataset.cfCellSetup) return;
        container.dataset.cfCellSetup = '1';

        container.addEventListener('click', (e) => {
            if (this.calcMode) return;
            const cell = e.target.closest('.cf-editable');
            if (!cell) return;

            const catId = parseInt(cell.dataset.catId);
            const month = cell.dataset.month;
            const currentText = cell.textContent.replace(/[^0-9.\-]/g, '');

            const input = document.createElement('input');
            input.type = 'number';
            input.step = '0.01';
            input.className = 'pnl-cell-input';
            input.value = currentText || '';
            cell.textContent = '';
            cell.appendChild(input);
            input.focus();
            input.select();

            const save = () => {
                const val = input.value.trim();
                if (val === '') {
                    Database.setCashFlowOverride(catId, month, null);
                } else {
                    Database.setCashFlowOverride(catId, month, parseFloat(val));
                }
                this.refreshCashFlow();
            };

            input.addEventListener('blur', save);
            input.addEventListener('keydown', (ke) => {
                if (ke.key === 'Enter') {
                    ke.preventDefault();
                    input.blur();
                } else if (ke.key === 'Escape') {
                    ke.preventDefault();
                    this.refreshCashFlow();
                }
            });
        });
    },

    /**
     * Get cached timeline or read from DB
     * @returns {Object} {start, end}
     */
    getTimeline() {
        if (!this._timeline) {
            this._timeline = Database.getTimeline();
        }
        return this._timeline;
    },

    /**
     * Invalidate cached timeline
     */
    invalidateTimeline() {
        this._timeline = null;
    },

    /**
     * Load timeline from DB and sync gear popover selects
     */
    loadAndApplyTimeline() {
        const timeline = Database.getTimeline();
        this._timeline = timeline;

        // Populate timeline year selects
        const years = Utils.generateYearOptions();
        ['timelineStartYear', 'timelineEndYear'].forEach(id => {
            const select = document.getElementById(id);
            if (!select) return;
            select.innerHTML = '<option value="">Year...</option>';
            years.forEach(y => {
                const opt = document.createElement('option');
                opt.value = y;
                opt.textContent = y;
                select.appendChild(opt);
            });
        });

        // Sync selects to saved values
        if (timeline.start) {
            const [sy, sm] = timeline.start.split('-');
            document.getElementById('timelineStartMonth').value = sm;
            document.getElementById('timelineStartYear').value = sy;
        } else {
            document.getElementById('timelineStartMonth').value = '';
            document.getElementById('timelineStartYear').value = '';
        }
        if (timeline.end) {
            const [ey, em] = timeline.end.split('-');
            document.getElementById('timelineEndMonth').value = em;
            document.getElementById('timelineEndYear').value = ey;
        } else {
            document.getElementById('timelineEndMonth').value = '';
            document.getElementById('timelineEndYear').value = '';
        }

        this.applyTimelineConstraints(timeline);
    },

    /**
     * Apply timeline constraints to date pickers and year dropdowns
     * @param {Object} timeline - {start, end}
     */
    applyTimelineConstraints(timeline) {
        const dateMin = Utils.timelineToDateMin(timeline.start);
        const dateMax = Utils.timelineToDateMax(timeline.end);

        // Constrain all date inputs
        const dateInputIds = [
            'entryDate', 'dateProcessed', 'assetDate', 'assetDeprStart',
            'loanStartDate', 'seedExpectedDate', 'seedReceivedDate',
            'apicExpectedDate', 'apicReceivedDate', 'bulkEntryDate', 'bulkDateProcessed'
        ];
        dateInputIds.forEach(id => {
            const input = document.getElementById(id);
            if (!input) return;
            input.min = dateMin;
            input.max = dateMax;
        });

        // Constrain year dropdowns using timeline-aware years
        const tlYears = Utils.getYearsInTimeline(timeline.start, timeline.end);
        UI.populateYearDropdowns(timeline);

        // Constrain BS year dropdown
        const bsYearSelect = document.getElementById('bsMonthYear');
        if (bsYearSelect) {
            const currentVal = bsYearSelect.value;
            bsYearSelect.innerHTML = '<option value="">Year...</option>';
            tlYears.forEach(y => {
                const opt = document.createElement('option');
                opt.value = y;
                opt.textContent = y;
                bsYearSelect.appendChild(opt);
            });
            if (currentVal) bsYearSelect.value = currentVal;
        }
    },

    /**
     * Handle timeline select change
     */
    handleTimelineChange() {
        const startMonth = document.getElementById('timelineStartMonth').value;
        const startYear = document.getElementById('timelineStartYear').value;
        const endMonth = document.getElementById('timelineEndMonth').value;
        const endYear = document.getElementById('timelineEndYear').value;

        const start = (startMonth && startYear) ? `${startYear}-${startMonth}` : null;
        const end = (endMonth && endYear) ? `${endYear}-${endMonth}` : null;

        // Save to DB and update cache (don't re-sync selects — user is editing them)
        Database.setTimelineStart(start);
        Database.setTimelineEnd(end);
        this._timeline = { start, end };

        this.applyTimelineConstraints(this._timeline);

        // Only refresh tabs when at least one complete range endpoint exists
        if (start || end) {
            this.refreshAll();
            // Also refresh analytics (dashboard + analyze chart panels)
            if (this.currentMode === 'analyze') {
                const activeTab = document.querySelector('.main-tab.active');
                const tab = activeTab ? activeTab.dataset.tab : 'dashboard';
                if (tab === 'dashboard') this.refreshDashboard();
                else this._updateAnalyzeCharts(tab);
            }
        }
    },

    /**
     * Refresh P&L spreadsheet
     */
    refreshPnL() {
        const plData = Database.getPLSpreadsheet();
        const overrides = Database.getAllPLOverrides();
        const taxMode = Database.getPLTaxMode();
        const timeline = this.getTimeline();
        const currentMonth = Utils.getCurrentMonth();

        // Filter months by timeline
        if (timeline.start || timeline.end) {
            plData.months = Utils.filterMonthsByTimeline(plData.months, timeline.start, timeline.end);
        }

        // Add future months up to timeline end (for projections)
        if (timeline.end && timeline.end > currentMonth) {
            let m = Utils.nextMonth(currentMonth);
            // Only add if beyond existing months
            while (m <= timeline.end) {
                if (!plData.months.includes(m)) {
                    plData.months.push(m);
                }
                m = Utils.nextMonth(m);
            }
            plData.months.sort();
        }

        // Get projected sales data for P&L integration
        let projectedSales = null;
        const psConfig = Database.getProjectedSalesConfig();
        const pnlViewToggle = document.getElementById('pnlViewToggle');
        if (psConfig.enabled && psConfig.projectionStartMonth) {
            // Ensure continuous month range through at least currentMonth for projections
            const lastNeeded = plData.months.length > 0 && plData.months[plData.months.length - 1] > currentMonth
                ? plData.months[plData.months.length - 1] : currentMonth;
            let fill = psConfig.projectionStartMonth;
            while (fill <= lastNeeded) {
                if (!plData.months.includes(fill)) plData.months.push(fill);
                fill = Utils.nextMonth(fill);
            }
            plData.months.sort();

            const pnlViewMode = document.getElementById('pnlViewMode');
            const pnlAsOfEl = document.getElementById('pnlAsOfMonth');
            const psSpreadsheet = Database.getProjectedSalesSpreadsheet(psConfig, plData.months);

            // Populate as-of dropdown with timeline months (restore saved value)
            this._populateAsOfSelect(pnlAsOfEl, plData.months, Database.getAsOfMonth('pnl'));

            const asOfVal = pnlAsOfEl ? pnlAsOfEl.value : 'current';
            projectedSales = {
                enabled: true,
                projectionStartMonth: psConfig.projectionStartMonth,
                byMonth: psSpreadsheet.byMonth,
                channels: psSpreadsheet.channels,
                viewMode: pnlViewMode ? pnlViewMode.value : 'projected',
                asOfMonth: asOfVal !== 'current' ? asOfVal : null
            };
            if (pnlViewToggle) pnlViewToggle.style.display = 'flex';
        } else {
            if (pnlViewToggle) pnlViewToggle.style.display = 'none';
        }

        // Sync dropdown
        const taxModeSelect = document.getElementById('plTaxMode');
        if (taxModeSelect) taxModeSelect.value = taxMode;
        UI.renderProfitLossSpreadsheet(plData, overrides, taxMode, currentMonth, projectedSales);
        this.setupPnLDragDrop();
        this.setupPnLCellEditing();
    },

    /**
     * Set up drag-and-drop for P&L category rows
     */
    setupPnLDragDrop() {
        const container = document.getElementById('pnlSpreadsheet');
        if (!container || container.dataset.pnlDragSetup) return;
        container.dataset.pnlDragSetup = '1';

        let draggedRow = null;

        container.addEventListener('dragstart', (e) => {
            const row = e.target.closest('tr[draggable="true"]');
            if (!row) return;
            draggedRow = row;
            row.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
            e.dataTransfer.setData('text/plain', row.dataset.categoryId);
        });

        container.addEventListener('dragover', (e) => {
            e.preventDefault();
            const row = e.target.closest('tr[draggable="true"]');
            if (!row || row === draggedRow || !draggedRow) return;
            if (row.dataset.section !== draggedRow.dataset.section) return;
            e.dataTransfer.dropEffect = 'move';
            container.querySelectorAll(`tr[data-section="${draggedRow.dataset.section}"].drag-over`)
                .forEach(r => r.classList.remove('drag-over'));
            row.classList.add('drag-over');
        });

        container.addEventListener('dragleave', (e) => {
            const row = e.target.closest('tr[draggable="true"]');
            if (row) row.classList.remove('drag-over');
        });

        container.addEventListener('drop', (e) => {
            e.preventDefault();
            const targetRow = e.target.closest('tr[draggable="true"]');
            if (!targetRow || !draggedRow || targetRow === draggedRow) return;
            if (targetRow.dataset.section !== draggedRow.dataset.section) return;

            const parent = draggedRow.parentNode;
            parent.insertBefore(draggedRow, targetRow);

            const section = draggedRow.dataset.section;
            const rows = parent.querySelectorAll(`tr[data-section="${section}"]`);
            const orderList = [];
            rows.forEach((row, index) => {
                orderList.push({ id: parseInt(row.dataset.categoryId), sortOrder: index });
            });

            Database.updatePLSortOrder(orderList);
            targetRow.classList.remove('drag-over');
        });

        container.addEventListener('dragend', () => {
            if (draggedRow) {
                draggedRow.classList.remove('dragging');
                draggedRow = null;
            }
            container.querySelectorAll('.drag-over').forEach(r => r.classList.remove('drag-over'));
        });
    },

    /**
     * Set up inline cell editing for P&L spreadsheet (only binds once)
     */
    setupPnLCellEditing() {
        const container = document.getElementById('pnlSpreadsheet');
        if (!container || container.dataset.pnlEditing) return;
        container.dataset.pnlEditing = '1';

        container.addEventListener('click', (e) => {
            if (this.calcMode) return;
            const cell = e.target.closest('.pnl-editable');
            if (!cell || cell.querySelector('.pnl-cell-input')) return;

            const catId = parseInt(cell.dataset.catId);
            const month = cell.dataset.month;

            // Get current displayed value (strip currency formatting)
            const currentText = cell.textContent.replace(/[^0-9.\-]/g, '');
            const currentVal = parseFloat(currentText) || 0;

            const input = document.createElement('input');
            input.type = 'number';
            input.className = 'pnl-cell-input';
            input.step = '0.01';
            input.value = currentVal;

            cell.textContent = '';
            cell.appendChild(input);
            input.focus();
            input.select();

            const save = () => {
                const newVal = input.value.trim();
                if (newVal === '' || newVal === currentText) {
                    // No change or cleared - remove override if cleared
                    if (newVal === '') {
                        Database.setPLOverride(catId, month, null);
                    }
                } else {
                    Database.setPLOverride(catId, month, parseFloat(newVal));
                }
                this.refreshPnL();
            };

            input.addEventListener('blur', save);
            input.addEventListener('keydown', (ke) => {
                if (ke.key === 'Enter') {
                    ke.preventDefault();
                    input.blur();
                } else if (ke.key === 'Escape') {
                    ke.preventDefault();
                    // Cancel - just re-render without saving
                    this.refreshPnL();
                }
            });
        });
    },

    /**
     * Switch between main tabs
     * @param {string} tab - 'journal' | 'cashflow' | 'pnl' | 'balancesheet' | 'assets' | 'loan' | 'budget' | 'breakeven'
     */
    switchMainTab(tab) {
        // Close sidebar on mobile after selecting a tab
        if (window.innerWidth <= 1024) {
            document.querySelector('.app-container').classList.remove('sidebar-open');
        }

        // Exit bulk select mode when switching tabs
        if (this.bulkSelectMode) {
            this.exitBulkSelectMode();
        }

        // Show/hide Quick Add button
        this._showQuickAddButton(tab);

        // Hide global summary cards on dashboard tab (KPI cards replace them)
        const summarySection = document.querySelector('.summary-section');
        if (summarySection) {
            summarySection.style.display = tab === 'dashboard' ? 'none' : '';
        }

        document.querySelectorAll('.main-tab').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tab === tab);
        });

        const tabs = ['journalTab', 'cashflowTab', 'pnlTab', 'balancesheetTab', 'assetsTab', 'loanTab', 'budgetTab', 'breakevenTab', 'projectedsalesTab', 'productsTab', 'vesalesTab', 'b2bcontractTab', 'quickguideTab', 'changelogTab', 'dashboardTab'];
        tabs.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.style.display = 'none';
        });

        if (tab === 'cashflow') {
            document.getElementById('cashflowTab').style.display = 'block';
            this.refreshCashFlow();
        } else if (tab === 'pnl') {
            document.getElementById('pnlTab').style.display = 'block';
            this.refreshPnL();
        } else if (tab === 'balancesheet') {
            document.getElementById('balancesheetTab').style.display = 'block';
            this.refreshBalanceSheet();
        } else if (tab === 'assets') {
            document.getElementById('assetsTab').style.display = 'block';
            this.refreshFixedAssets();
        } else if (tab === 'loan') {
            document.getElementById('loanTab').style.display = 'block';
            this.refreshLoans();
        } else if (tab === 'budget') {
            document.getElementById('budgetTab').style.display = 'block';
            this.refreshBudget();
        } else if (tab === 'breakeven') {
            document.getElementById('breakevenTab').style.display = 'block';
            this.refreshBreakeven();
        } else if (tab === 'projectedsales') {
            document.getElementById('projectedsalesTab').style.display = 'block';
            this.refreshProjectedSales();
        } else if (tab === 'products') {
            document.getElementById('productsTab').style.display = 'block';
            this.refreshProducts();
        } else if (tab === 'vesales') {
            document.getElementById('vesalesTab').style.display = 'block';
            this.refreshVESales();
        } else if (tab === 'b2bcontract') {
            document.getElementById('b2bcontractTab').style.display = 'block';
            this.refreshB2BContracts();
        } else if (tab === 'dashboard') {
            document.getElementById('dashboardTab').style.display = 'block';
            this.refreshDashboard();
        } else if (tab === 'quickguide') {
            document.getElementById('quickguideTab').style.display = 'block';
            this.renderQuickGuide();
        } else if (tab === 'changelog') {
            document.getElementById('changelogTab').style.display = 'block';
            this.renderChangelog();
        } else {
            document.getElementById('journalTab').style.display = 'block';
        }

        // Show/hide analyze chart panels based on mode
        this._updateAnalyzeCharts(tab);
    },

    // ==================== MODE TOGGLE (Work / Analyze) ====================

    switchMode(mode) {
        if (mode === this.currentMode) return;

        // Save current tab for the mode we're leaving
        const activeTab = document.querySelector('.main-tab.active');
        if (activeTab) {
            if (this.currentMode === 'work') this.lastWorkTab = activeTab.dataset.tab;
            else this.lastAnalyzeTab = activeTab.dataset.tab;
        }

        this.currentMode = mode;

        // Update toggle buttons
        document.querySelectorAll('.mode-toggle-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.mode === mode);
        });

        // Apply mode-aware hidden tabs (shows/hides sidebar tabs + group labels)
        this.applyModeAwareHiddenTabs();

        // Switch to the remembered tab for this mode
        const targetTab = mode === 'work' ? this.lastWorkTab : this.lastAnalyzeTab;
        this.switchMainTab(targetTab);
    },

    _dashCharts: {},

    _dashSnapshotMonth: null,

    refreshDashboard() {
        // Populate snapshot month dropdown
        const select = document.getElementById('dashSnapshotMonth');
        if (select) {
            const realCurrent = Utils.getCurrentMonth();
            const allMonths = this._getTimelineMonths().filter(m => m <= realCurrent);
            const prevVal = this._dashSnapshotMonth || select.value;
            select.innerHTML = '<option value="">Current Month</option>' +
                allMonths.map(m => `<option value="${m}">${Utils.formatMonthShort(m)}</option>`).join('');
            if (prevVal && allMonths.includes(prevVal)) {
                select.value = prevVal;
                this._dashSnapshotMonth = prevVal;
            } else {
                select.value = '';
                this._dashSnapshotMonth = null;
            }
        }
        const snapshotMonth = this._dashSnapshotMonth || null;
        this._computeKpiData(snapshotMonth);
        this._renderDashboardSections(snapshotMonth);
    },

    // ==================== KPI DATA COMPUTATION (cached for modal use) ====================

    _kpiCache: null,

    _computeKpiData(snapshotMonth) {
        const summary = Database.calculateSummary();
        const currentMonth = snapshotMonth || Utils.getCurrentMonth();

        // Calculate monthly data for sparklines — only up to current month (exclude future)
        const allMonths = this._getTimelineMonths().filter(m => m <= currentMonth);
        const months = allMonths.slice(-6);

        // Get monthly revenue & expense data
        const monthlyRevenue = [];
        const monthlyExpenses = [];
        const monthlyCash = [];
        let runningCash = 0;

        // P&L accrual-basis expense totals (for burn metrics)
        const plOpex = Database.getMonthlyTotalOpex(months);
        const plCogs = {};
        const cogsResult = Database.db.exec(
            "SELECT t.month_due, SUM(t.amount) as total FROM transactions t JOIN categories c ON t.category_id = c.id WHERE t.month_due IN (" + months.map(() => '?').join(',') + ") AND c.is_cogs = 1 AND c.show_on_pl != 1 GROUP BY t.month_due", months);
        if (cogsResult[0]) { for (const row of cogsResult[0].values) { plCogs[row[0]] = row[1]; } }

        for (const month of months) {
            const revResult = Database.db.exec(
                "SELECT COALESCE(SUM(COALESCE(t.pretax_amount, t.amount)),0) FROM transactions t JOIN categories c ON t.category_id = c.id WHERE t.transaction_type='receivable' AND t.month_due=? AND c.is_cogs = 0 AND c.show_on_pl != 1 AND (c.is_b2b = 1 OR c.is_sales = 1) AND COALESCE(t.source_type, '') NOT IN ('loan_receivable', 'loan_payment')", [month]);
            const cashExpResult = Database.db.exec(
                "SELECT COALESCE(SUM(amount),0) FROM transactions WHERE transaction_type='payable' AND status='paid' AND month_paid=?", [month]);
            const rev = Math.round((revResult[0] ? revResult[0].values[0][0] : 0) * 100) / 100;
            const cashExp = Math.round((cashExpResult[0] ? cashExpResult[0].values[0][0] : 0) * 100) / 100;
            const plExp = Math.round(((plOpex[month] || 0) + (plCogs[month] || 0)) * 100) / 100;
            monthlyRevenue.push(rev);
            monthlyExpenses.push(plExp);
            runningCash = Math.round((runningCash + rev - cashExp) * 100) / 100;
            monthlyCash.push(runningCash);
        }

        // Overdue receivables
        const overdueResult = Database.db.exec(
            "SELECT COUNT(*), COALESCE(SUM(amount),0) FROM transactions WHERE transaction_type='receivable' AND status='pending' AND month_due < ?", [currentMonth]);
        const overdueCount = overdueResult[0] ? overdueResult[0].values[0][0] : 0;
        const overdueAmount = Math.round((overdueResult[0] ? overdueResult[0].values[0][1] : 0) * 100) / 100;

        // Burn rate (avg expenses last 6 months)
        const totalExp = monthlyExpenses.reduce((s, v) => s + v, 0);
        const burnRate = Math.round((monthlyExpenses.length > 0 ? totalExp / monthlyExpenses.length : 0) * 100) / 100;

        // Gross burn (total expenses) vs Net burn (expenses minus revenue)
        const totalRev = monthlyRevenue.reduce((s, v) => s + v, 0);
        const avgRevenue = Math.round((monthlyRevenue.length > 0 ? totalRev / monthlyRevenue.length : 0) * 100) / 100;
        const netBurn = Math.round((burnRate - avgRevenue) * 100) / 100;
        const monthlyNetBurn = months.map((m, i) => monthlyExpenses[i] - monthlyRevenue[i]);

        // EBITDA — use P&L data for current timeline
        const plData = Database.getPLSpreadsheet();
        const plOverrides = Database.getAllPLOverrides();
        const isFuture = (m) => m > currentMonth;
        const pastMonths = months.filter(m => !isFuture(m));

        const getPlVal = (catId, month, computed) => {
            const key = `${catId}-${month}`;
            return (key in plOverrides) ? plOverrides[key] : computed;
        };
        const groupPlByCategory = (rows) => {
            const map = {};
            (rows || []).forEach(row => {
                if (!map[row.category_id]) map[row.category_id] = {};
                map[row.category_id][row.month] = (map[row.category_id][row.month] || 0) + row.total;
            });
            return map;
        };

        // Sum revenue through current month
        let ebitdaRevenue = 0;
        const revByCat = groupPlByCategory(plData.revenue);
        Object.entries(revByCat).forEach(([catId, catMonths]) => {
            pastMonths.forEach(m => { ebitdaRevenue += getPlVal(catId, m, catMonths[m] || 0); });
        });

        // Sum COGS + OpEx through current month
        let ebitdaExpenses = 0;
        const cogsByCat = groupPlByCategory(plData.cogs);
        Object.entries(cogsByCat).forEach(([catId, catMonths]) => {
            pastMonths.forEach(m => { ebitdaExpenses += getPlVal(catId, m, catMonths[m] || 0); });
        });
        const opexByCat = groupPlByCategory(plData.opex);
        Object.entries(opexByCat).forEach(([catId, catMonths]) => {
            pastMonths.forEach(m => { ebitdaExpenses += getPlVal(catId, m, catMonths[m] || 0); });
        });

        // EBITDA = Revenue - (COGS + OpEx), excluding depreciation, interest, and tax
        const ebitda = Math.round((ebitdaRevenue - ebitdaExpenses) * 100) / 100;
        const ebitdaMargin = ebitdaRevenue > 0 ? ((ebitda / ebitdaRevenue) * 100).toFixed(1) : '0.0';

        // Revenue trend — current month vs previous month
        const thisMonthRev = monthlyRevenue[monthlyRevenue.length - 1] || 0;
        const lastMonthRev = monthlyRevenue[monthlyRevenue.length - 2] || 0;
        const revTrend = lastMonthRev > 0 ? ((thisMonthRev - lastMonthRev) / lastMonthRev * 100).toFixed(1) : '0';

        // CMGR (Compound Monthly Growth Rate) over available months
        const firstRev = monthlyRevenue.find(v => v > 0) || 0;
        const lastRevVal = monthlyRevenue[monthlyRevenue.length - 1] || 0;
        const revMonthsCount = monthlyRevenue.length;
        let cmgr = 0;
        if (firstRev > 0 && lastRevVal > 0 && revMonthsCount > 1) {
            cmgr = (Math.pow(lastRevVal / firstRev, 1 / (revMonthsCount - 1)) - 1) * 100;
        }
        const cmgrFormatted = cmgr.toFixed(1);

        // Non-B2B CMGR (consumer sales only, excluding B2B categories)
        const monthlyNonB2BRev = [];
        for (const month of months) {
            const nbRevResult = Database.db.exec(
                "SELECT COALESCE(SUM(COALESCE(t.pretax_amount, t.amount)),0) FROM transactions t JOIN categories c ON t.category_id = c.id " +
                "WHERE t.transaction_type='receivable' AND t.month_due=? AND c.is_cogs = 0 AND c.show_on_pl != 1 AND c.is_sales = 1 AND c.is_b2b = 0 " +
                "AND COALESCE(t.source_type, '') NOT IN ('loan_receivable', 'loan_payment')", [month]);
            monthlyNonB2BRev.push(Math.round((nbRevResult[0] ? nbRevResult[0].values[0][0] : 0) * 100) / 100);
        }
        const firstNB = monthlyNonB2BRev.find(v => v > 0) || 0;
        const lastNB = monthlyNonB2BRev[monthlyNonB2BRev.length - 1] || 0;
        const nbMonthsCount = monthlyNonB2BRev.length;
        let cmgrNonB2B = 0;
        if (firstNB > 0 && lastNB > 0 && nbMonthsCount > 1) {
            cmgrNonB2B = (Math.pow(lastNB / firstNB, 1 / (nbMonthsCount - 1)) - 1) * 100;
        }
        const cmgrNonB2BFormatted = cmgrNonB2B.toFixed(1);
        const thisMonthNB = monthlyNonB2BRev[monthlyNonB2BRev.length - 1] || 0;

        // Working Capital trend (current assets - current liabilities per month)
        const monthlyWorkingCap = [];
        for (const month of months) {
            const cashAsOf = Database.getCashAsOf(month);
            const arAsOf = Database.getAccountsReceivableAsOf ? Database.getAccountsReceivableAsOf(month) : 0;
            const apAsOf = Database.getAccountsPayableAsOf ? Database.getAccountsPayableAsOf(month) : 0;
            const stpAsOf = Database.getSalesTaxPayableAsOf ? Database.getSalesTaxPayableAsOf(month) : 0;
            monthlyWorkingCap.push((cashAsOf + arAsOf) - (apAsOf + stpAsOf));
        }
        const currentWC = monthlyWorkingCap[monthlyWorkingCap.length - 1] || 0;
        const prevWC = monthlyWorkingCap.length > 1 ? monthlyWorkingCap[monthlyWorkingCap.length - 2] : 0;
        const wcTrend = prevWC !== 0 ? ((currentWC - prevWC) / Math.abs(prevWC) * 100).toFixed(1) : '0.0';

        // Rule of 40: Revenue Growth % + Profit Margin %
        const revGrowthPct = parseFloat(cmgrFormatted) * 12; // Annualized
        const profitMarginPct = ebitdaRevenue > 0 ? (ebitda / ebitdaRevenue) * 100 : 0;
        const rule40Score = Math.round(revGrowthPct + profitMarginPct);
        const rule40Cls = rule40Score >= 40 ? 'positive' : (rule40Score >= 20 ? 'dash-kpi-warning' : 'negative');

        // Rule of 40 (Non-B2B): uses non-B2B CMGR for growth component
        const nbRevGrowthPct = parseFloat(cmgrNonB2BFormatted) * 12; // Annualized
        const rule40NB = Math.round(nbRevGrowthPct + profitMarginPct);
        const rule40NBCls = rule40NB >= 40 ? 'positive' : (rule40NB >= 20 ? 'dash-kpi-warning' : 'negative');

        // DSCR — Net Operating Income / Total Debt Service
        const loans = Database.getLoans();
        let totalDebtService = 0;
        loans.forEach(loan => {
            const skipped = Database.getSkippedPayments(loan.id);
            const overridesLoan = Database.getLoanPaymentOverrides(loan.id);
            const schedule = Utils.computeAmortizationSchedule({
                principal: loan.principal, annual_rate: loan.annual_rate,
                term_months: loan.term_months, payments_per_year: loan.payments_per_year,
                start_date: loan.start_date, first_payment_date: loan.first_payment_date
            }, skipped, overridesLoan);
            schedule.forEach(pmt => {
                if (pmt.month <= currentMonth && months.includes(pmt.month)) {
                    totalDebtService += pmt.payment;
                }
            });
        });
        const dscr = totalDebtService > 0 ? (ebitda / totalDebtService) : null;
        const dscrDisplay = dscr !== null ? dscr.toFixed(2) + 'x' : 'No debt';
        const dscrCls = dscr === null ? '' : (dscr >= 1.25 ? 'positive' : (dscr >= 1.0 ? 'dash-kpi-warning' : 'negative'));

        this._kpiCache = {
            summary, currentMonth, months, revMonthsCount, nbMonthsCount,
            monthlyRevenue, monthlyExpenses, monthlyCash, monthlyNetBurn, monthlyNonB2BRev, monthlyWorkingCap,
            overdueCount, overdueAmount, burnRate, netBurn,
            ebitda, ebitdaMargin, thisMonthRev, revTrend,
            cmgr, cmgrFormatted, cmgrNonB2B, cmgrNonB2BFormatted, thisMonthNB,
            currentWC, wcTrend,
            rule40Score, rule40Cls, rule40NB, rule40NBCls,
            dscr, dscrDisplay, dscrCls
        };
    },

    // ==================== SECTION → KPI MODAL ====================

    _sectionKpiMap: {
        cashflow:         { title: 'Cash & Survival',        kpis: ['cashposition', 'grossburn', 'netburn'] },
        pnl:              { title: 'Revenue & Profitability', kpis: ['revtrend', 'ebitda', 'cmgr', 'cmgrnonb2b', 'rule40', 'rule40nb'] },
        mom:              { title: 'Growth Metrics',          kpis: ['cmgrnonb2b', 'rule40nb'] },
        balancesheet:     { title: 'Health & Risk',           kpis: ['workingcapital', 'overdue', 'dscr'] },
        revconcentration: { title: 'Revenue Growth',          kpis: ['revtrend', 'cmgr', 'cmgrnonb2b'] }
    },

    _openSectionAnalysis(section) {
        const group = this._sectionKpiMap[section];
        if (!group) {
            this.switchMainTab(section === 'breakeven' ? 'breakeven' : section);
            return;
        }

        this._computeKpiData(this._dashSnapshotMonth || null);
        const d = this._kpiCache;

        const snapshotLabel = this._dashSnapshotMonth ? ' — ' + Utils.formatMonthShort(this._dashSnapshotMonth) : '';
        document.getElementById('analyzeKpiTitle').textContent = group.title + snapshotLabel;
        const body = document.getElementById('analyzeKpiBody');

        let html = '<div class="analyze-kpi-grid">';
        const sparklines = [];

        group.kpis.forEach(type => {
            const card = this._buildKpiCardHtml(type, d, sparklines);
            if (card) html += card;
        });
        html += '</div>';

        body.innerHTML = html;

        body.querySelectorAll('.dash-kpi[data-kpi]').forEach(card => {
            card.addEventListener('click', () => this._openKpiDetail(card.dataset.kpi));
        });

        sparklines.forEach(s => this._renderSparkline(s.id, s.data, s.color));

        UI.showModal('analyzeKpiModal');
    },

    _buildKpiCardHtml(type, d, sparklines) {
        const sid = 'modalSpark_' + type;
        switch (type) {
            case 'cashposition':
                sparklines.push({ id: sid, data: d.monthlyCash, color: 'rgba(59,130,246,0.8)' });
                return '<div class="dash-kpi" data-kpi="cashposition">' +
                    '<span class="dash-kpi-label">Cash Position</span>' +
                    '<span class="dash-kpi-value ' + (d.summary.cashBalance >= 0 ? 'positive' : 'negative') + '">' + Utils.formatCurrency(d.summary.cashBalance) + '</span>' +
                    '<canvas class="dash-sparkline" id="' + sid + '" width="80" height="30"></canvas>' +
                '</div>';
            case 'grossburn':
                sparklines.push({ id: sid, data: d.monthlyExpenses, color: 'rgba(239,68,68,0.8)' });
                return '<div class="dash-kpi" data-kpi="grossburn">' +
                    '<span class="dash-kpi-label">Gross Burn</span>' +
                    '<span class="dash-kpi-value negative">' + Utils.formatCurrency(d.burnRate) + '<small style="font-size:0.6em;font-weight:500;">/mo</small></span>' +
                    '<canvas class="dash-sparkline" id="' + sid + '" width="80" height="30"></canvas>' +
                '</div>';
            case 'netburn':
                sparklines.push({ id: sid, data: d.monthlyNetBurn, color: 'rgba(251,146,60,0.8)' });
                return '<div class="dash-kpi" data-kpi="netburn">' +
                    '<span class="dash-kpi-label">Net Burn</span>' +
                    '<span class="dash-kpi-value ' + (d.netBurn > 0 ? 'negative' : 'positive') + '">' + Utils.formatCurrency(Math.abs(d.netBurn)) + '<small style="font-size:0.6em;font-weight:500;">/mo ' + (d.netBurn <= 0 ? '(net positive)' : '') + '</small></span>' +
                    '<canvas class="dash-sparkline" id="' + sid + '" width="80" height="30"></canvas>' +
                '</div>';
            case 'revtrend':
                sparklines.push({ id: sid, data: d.monthlyRevenue, color: 'rgba(16,185,129,0.8)' });
                return '<div class="dash-kpi" data-kpi="revtrend">' +
                    '<span class="dash-kpi-label">Revenue Trend</span>' +
                    '<span class="dash-kpi-value">' + Utils.formatCurrency(d.thisMonthRev) + ' <small style="font-size:0.7em;color:' + (d.revTrend >= 0 ? 'var(--color-success,#10b981)' : 'var(--color-danger,#ef4444)') + '">' + (d.revTrend >= 0 ? '+' : '') + d.revTrend + '%</small></span>' +
                    '<canvas class="dash-sparkline" id="' + sid + '" width="80" height="30"></canvas>' +
                '</div>';
            case 'cmgr':
                return '<div class="dash-kpi" data-kpi="cmgr">' +
                    '<span class="dash-kpi-label">Growth Rate <small style="font-size:0.85em;text-transform:none;">(CMGR)</small></span>' +
                    '<span class="dash-kpi-value ' + (d.cmgr >= 0 ? 'positive' : 'negative') + '">' + (d.cmgr >= 0 ? '+' : '') + d.cmgrFormatted + '%<small style="font-size:0.6em;font-weight:500;">/mo</small></span>' +
                    '<span class="dash-kpi-sub" style="font-size:0.7rem;color:var(--text-muted);">over ' + d.revMonthsCount + ' months</span>' +
                '</div>';
            case 'cmgrnonb2b':
                sparklines.push({ id: sid, data: d.monthlyNonB2BRev, color: 'rgba(251,146,60,0.8)' });
                return '<div class="dash-kpi" data-kpi="cmgrnonb2b">' +
                    '<span class="dash-kpi-label">Non-B2B CMGR</span>' +
                    '<span class="dash-kpi-value ' + (d.cmgrNonB2B >= 0 ? 'positive' : 'negative') + '">' + (d.cmgrNonB2B >= 0 ? '+' : '') + d.cmgrNonB2BFormatted + '%<small style="font-size:0.6em;font-weight:500;">/mo</small></span>' +
                    '<span class="dash-kpi-sub" style="font-size:0.7rem;color:var(--text-muted);">' + Utils.formatCurrency(d.thisMonthNB) + ' this month</span>' +
                    '<canvas class="dash-sparkline" id="' + sid + '" width="80" height="30"></canvas>' +
                '</div>';
            case 'ebitda':
                return '<div class="dash-kpi" data-kpi="ebitda">' +
                    '<span class="dash-kpi-label">EBITDA</span>' +
                    '<span class="dash-kpi-value ' + (d.ebitda >= 0 ? 'positive' : 'negative') + '">' + Utils.formatCurrency(d.ebitda) + '</span>' +
                    '<span class="dash-kpi-sub" style="font-size:0.7rem;color:var(--text-muted);">margin: ' + d.ebitdaMargin + '%</span>' +
                '</div>';
            case 'rule40':
                return '<div class="dash-kpi" data-kpi="rule40">' +
                    '<span class="dash-kpi-label">Rule of 40</span>' +
                    '<span class="dash-kpi-value ' + d.rule40Cls + '">' + d.rule40Score + ' <small style="font-size:0.6em;font-weight:500;">/ 40</small></span>' +
                    '<span class="dash-kpi-sub" style="font-size:0.7rem;color:var(--text-muted);">growth + margin</span>' +
                '</div>';
            case 'rule40nb':
                return '<div class="dash-kpi" data-kpi="rule40nb">' +
                    '<span class="dash-kpi-label">Rule of 40 <small style="font-size:0.85em;text-transform:none;">(Non-B2B)</small></span>' +
                    '<span class="dash-kpi-value ' + d.rule40NBCls + '">' + d.rule40NB + ' <small style="font-size:0.6em;font-weight:500;">/ 40</small></span>' +
                    '<span class="dash-kpi-sub" style="font-size:0.7rem;color:var(--text-muted);">consumer growth + margin</span>' +
                '</div>';
            case 'workingcapital':
                sparklines.push({ id: sid, data: d.monthlyWorkingCap, color: 'rgba(168,85,247,0.8)' });
                return '<div class="dash-kpi" data-kpi="workingcapital">' +
                    '<span class="dash-kpi-label">Working Capital</span>' +
                    '<span class="dash-kpi-value ' + (d.currentWC >= 0 ? 'positive' : 'negative') + '">' + Utils.formatCurrency(d.currentWC) + ' <small style="font-size:0.7em;color:' + (parseFloat(d.wcTrend) >= 0 ? 'var(--color-success,#10b981)' : 'var(--color-danger,#ef4444)') + '">' + (parseFloat(d.wcTrend) >= 0 ? '+' : '') + d.wcTrend + '%</small></span>' +
                    '<canvas class="dash-sparkline" id="' + sid + '" width="80" height="30"></canvas>' +
                '</div>';
            case 'overdue':
                return '<div class="dash-kpi" data-kpi="overdue">' +
                    '<span class="dash-kpi-label">Overdue Receivables</span>' +
                    '<span class="dash-kpi-value ' + (d.overdueAmount > 0 ? 'negative' : 'positive') + '">' + Utils.formatCurrency(d.overdueAmount) + ' <small style="font-size:0.7em;color:var(--text-muted)">(' + d.overdueCount + ')</small></span>' +
                '</div>';
            case 'dscr':
                return '<div class="dash-kpi" data-kpi="dscr">' +
                    '<span class="dash-kpi-label">DSCR</span>' +
                    '<span class="dash-kpi-value ' + d.dscrCls + '">' + d.dscrDisplay + '</span>' +
                    '<span class="dash-kpi-sub" style="font-size:0.7rem;color:var(--text-muted);">debt service coverage</span>' +
                '</div>';
            default:
                return '';
        }
    },

    // ==================== KPI DETAIL MODAL ====================

    _kpiMeta: {
        cashposition: { title: 'Cash Position — Detail', tab: 'cashflow', tabLabel: 'Cash Flow' },
        grossburn:    { title: 'Gross Burn — Expense Breakdown', tab: 'pnl', tabLabel: 'P&L' },
        netburn:      { title: 'Net Burn — Revenue vs Expenses', tab: 'pnl', tabLabel: 'P&L' },
        revtrend:     { title: 'Revenue Trend — Monthly Detail', tab: 'pnl', tabLabel: 'P&L' },
        cmgr:         { title: 'Growth Rate (CMGR) — Analysis', tab: 'pnl', tabLabel: 'P&L' },
        cmgrnonb2b:   { title: 'Non-B2B CMGR — Consumer Sales Growth', tab: 'pnl', tabLabel: 'P&L' },
        ebitda:       { title: 'EBITDA — Breakdown', tab: 'pnl', tabLabel: 'P&L' },
        overdue:      { title: 'Overdue Receivables — Transactions', tab: 'journal', tabLabel: 'Journal' },
        workingcapital: { title: 'Working Capital — Trend', tab: 'balancesheet', tabLabel: 'Balance Sheet' },
        rule40:       { title: 'Rule of 40 — Breakdown', tab: 'pnl', tabLabel: 'P&L' },
        rule40nb:     { title: 'Rule of 40 (Non-B2B) — Breakdown', tab: 'pnl', tabLabel: 'P&L' },
        dscr:         { title: 'DSCR — Debt Service Coverage', tab: 'loan', tabLabel: 'Loans' }
    },

    _openKpiDetail(kpiType) {
        const meta = this._kpiMeta[kpiType];
        if (!meta) return;

        document.getElementById('kpiDetailTitle').textContent = meta.title;

        const body = document.getElementById('kpiDetailBody');
        const renderer = this['_kpiDetail_' + kpiType];
        body.innerHTML = renderer ? renderer.call(this) : '<p>No detail available.</p>';

        const goBtn = document.getElementById('kpiDetailGoBtn');
        goBtn.textContent = 'Go to ' + meta.tabLabel;
        goBtn.onclick = () => {
            UI.hideModal('kpiDetailModal');
            UI.hideModal('analyzeKpiModal');
            this.switchMainTab(meta.tab);
        };

        UI.showModal('kpiDetailModal');
    },

    _kpiDetailMonthLabel(m) {
        const [y, mo] = m.split('-');
        return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][parseInt(mo)-1] + ' ' + y;
    },

    _kpiDetail_cashposition() {
        const snapshotMonth = this._kpiCache ? this._kpiCache.currentMonth : Utils.getCurrentMonth();
        const months = this._getTimelineMonths().filter(m => m <= snapshotMonth);
        let html = '<div class="kpi-detail-summary">Month-by-month cash flow breakdown</div>';
        html += '<table class="kpi-detail-table"><thead><tr><th>Month</th><th>Money In</th><th>Money Out</th><th>Net</th><th>Running Balance</th></tr></thead><tbody>';

        let running = 0;
        for (const month of months) {
            const revResult = Database.db.exec("SELECT COALESCE(SUM(amount),0) FROM transactions WHERE transaction_type='receivable' AND status='received' AND month_paid=?", [month]);
            const expResult = Database.db.exec("SELECT COALESCE(SUM(amount),0) FROM transactions WHERE transaction_type='payable' AND status='paid' AND month_paid=?", [month]);
            const rev = revResult[0] ? revResult[0].values[0][0] : 0;
            const exp = expResult[0] ? expResult[0].values[0][0] : 0;
            const net = rev - exp;
            running += net;
            html += '<tr><td class="text-cell">' + this._kpiDetailMonthLabel(month) + '</td>' +
                '<td class="positive">' + Utils.formatCurrency(rev) + '</td>' +
                '<td class="negative">' + Utils.formatCurrency(exp) + '</td>' +
                '<td class="' + (net >= 0 ? 'positive' : 'negative') + '">' + Utils.formatCurrency(net) + '</td>' +
                '<td class="' + (running >= 0 ? 'positive' : 'negative') + '">' + Utils.formatCurrency(running) + '</td></tr>';
        }
        html += '</tbody></table>';
        return html;
    },

    _kpiDetail_grossburn() {
        const currentMonth = this._kpiCache ? this._kpiCache.currentMonth : Utils.getCurrentMonth();
        const months = this._getTimelineMonths().filter(m => m <= currentMonth).slice(-6);

        // Get expenses by category for last 6 months (accrual basis, matching P&L)
        // OpEx: payable, not cogs/depreciation/sales-tax/hidden, not loan payments
        const result = Database.db.exec(
            "SELECT c.name, t.month_due, SUM(t.amount) as total " +
            "FROM transactions t JOIN categories c ON t.category_id = c.id " +
            "WHERE t.month_due IN (" + months.map(() => '?').join(',') + ") " +
            "AND t.transaction_type='payable' AND c.is_cogs = 0 AND c.is_depreciation = 0 " +
            "AND c.is_sales_tax = 0 AND c.show_on_pl != 1 " +
            "AND COALESCE(t.source_type, '') NOT IN ('loan_payment', 'loan_receivable') " +
            "GROUP BY c.name, t.month_due ORDER BY total DESC", months);
        // COGS categories (is_cogs=1, any transaction type)
        const cogsResult = Database.db.exec(
            "SELECT c.name, t.month_due, SUM(t.amount) as total " +
            "FROM transactions t JOIN categories c ON t.category_id = c.id " +
            "WHERE t.month_due IN (" + months.map(() => '?').join(',') + ") " +
            "AND c.is_cogs = 1 AND c.show_on_pl != 1 " +
            "GROUP BY c.name, t.month_due ORDER BY total DESC", months);

        const catMap = {};
        if (result[0]) {
            for (const row of result[0].values) {
                if (!catMap[row[0]]) catMap[row[0]] = {};
                catMap[row[0]][row[1]] = (catMap[row[0]][row[1]] || 0) + row[2];
            }
        }
        if (cogsResult[0]) {
            for (const row of cogsResult[0].values) {
                if (!catMap[row[0]]) catMap[row[0]] = {};
                catMap[row[0]][row[1]] = (catMap[row[0]][row[1]] || 0) + row[2];
            }
        }
        // Sort categories by total descending
        const catTotals = Object.entries(catMap).map(([name, mData]) => ({
            name, total: Object.values(mData).reduce((s, v) => s + v, 0), mData
        })).sort((a, b) => b.total - a.total);

        let html = '<div class="kpi-detail-summary">Expense breakdown by category — last ' + months.length + ' months</div>';
        html += '<table class="kpi-detail-table"><thead><tr><th>Category</th>';
        months.forEach(m => { html += '<th>' + this._kpiDetailMonthLabel(m) + '</th>'; });
        html += '<th>Total</th></tr></thead><tbody>';

        const monthTotals = {};
        months.forEach(m => { monthTotals[m] = 0; });
        let grandTotal = 0;

        catTotals.forEach(cat => {
            html += '<tr><td class="text-cell">' + Utils.escapeHtml(cat.name) + '</td>';
            months.forEach(m => {
                const v = cat.mData[m] || 0;
                monthTotals[m] += v;
                html += '<td>' + Utils.formatCurrency(v) + '</td>';
            });
            grandTotal += cat.total;
            html += '<td style="font-weight:600;">' + Utils.formatCurrency(cat.total) + '</td></tr>';
        });

        html += '<tr class="kpi-row-total"><td class="text-cell">Total</td>';
        months.forEach(m => { html += '<td>' + Utils.formatCurrency(monthTotals[m]) + '</td>'; });
        html += '<td>' + Utils.formatCurrency(grandTotal) + '</td></tr>';
        html += '</tbody></table>';
        return html;
    },

    _kpiDetail_netburn() {
        const currentMonth = this._kpiCache ? this._kpiCache.currentMonth : Utils.getCurrentMonth();
        const months = this._getTimelineMonths().filter(m => m <= currentMonth).slice(-6);

        let html = '<div class="kpi-detail-summary">Revenue vs Expenses — last ' + months.length + ' months</div>';
        html += '<table class="kpi-detail-table"><thead><tr><th>Month</th><th>Revenue</th><th>Expenses</th><th>Net Burn</th></tr></thead><tbody>';

        // Use P&L accrual-basis data (matching gross burn)
        const plOpex = Database.getMonthlyTotalOpex(months);
        const plCogsResult = Database.db.exec(
            "SELECT t.month_due, SUM(t.amount) as total FROM transactions t JOIN categories c ON t.category_id = c.id WHERE t.month_due IN (" + months.map(() => '?').join(',') + ") AND c.is_cogs = 1 AND c.show_on_pl != 1 GROUP BY t.month_due", months);
        const plCogs = {};
        if (plCogsResult[0]) { for (const row of plCogsResult[0].values) { plCogs[row[0]] = row[1]; } }

        let totalRev = 0, totalExp = 0;
        for (const month of months) {
            const revResult = Database.db.exec(
                "SELECT COALESCE(SUM(COALESCE(t.pretax_amount, t.amount)),0) FROM transactions t JOIN categories c ON t.category_id = c.id WHERE t.transaction_type='receivable' AND t.month_due=? AND c.is_cogs = 0 AND c.show_on_pl != 1 AND (c.is_b2b = 1 OR c.is_sales = 1) AND COALESCE(t.source_type, '') NOT IN ('loan_receivable', 'loan_payment')", [month]);
            const rev = revResult[0] ? revResult[0].values[0][0] : 0;
            const exp = Math.round(((plOpex[month] || 0) + (plCogs[month] || 0)) * 100) / 100;
            const net = exp - rev;
            totalRev += rev;
            totalExp += exp;
            html += '<tr><td class="text-cell">' + this._kpiDetailMonthLabel(month) + '</td>' +
                '<td class="positive">' + Utils.formatCurrency(rev) + '</td>' +
                '<td class="negative">' + Utils.formatCurrency(exp) + '</td>' +
                '<td class="' + (net > 0 ? 'negative' : 'positive') + '">' + Utils.formatCurrency(Math.abs(net)) + (net <= 0 ? ' (net+)' : '') + '</td></tr>';
        }
        html += '<tr class="kpi-row-total"><td class="text-cell">Total</td>' +
            '<td class="positive">' + Utils.formatCurrency(totalRev) + '</td>' +
            '<td class="negative">' + Utils.formatCurrency(totalExp) + '</td>' +
            '<td class="' + ((totalExp - totalRev) > 0 ? 'negative' : 'positive') + '">' + Utils.formatCurrency(Math.abs(totalExp - totalRev)) + '</td></tr>';
        html += '</tbody></table>';
        return html;
    },

    _kpiDetail_revtrend() {
        const currentMonth = this._kpiCache ? this._kpiCache.currentMonth : Utils.getCurrentMonth();
        const months = this._getTimelineMonths().filter(m => m <= currentMonth);

        let html = '<div class="kpi-detail-summary">Monthly revenue with month-over-month change</div>';
        html += '<table class="kpi-detail-table"><thead><tr><th>Month</th><th>Revenue</th><th>MoM Change</th><th>MoM %</th></tr></thead><tbody>';

        let prevRev = null;
        for (const month of months) {
            const revResult = Database.db.exec(
                "SELECT COALESCE(SUM(COALESCE(t.pretax_amount, t.amount)),0) FROM transactions t JOIN categories c ON t.category_id = c.id WHERE t.transaction_type='receivable' AND t.month_due=? AND c.is_cogs = 0 AND c.show_on_pl != 1 AND (c.is_b2b = 1 OR c.is_sales = 1) AND COALESCE(t.source_type, '') NOT IN ('loan_receivable', 'loan_payment')", [month]);
            const rev = revResult[0] ? revResult[0].values[0][0] : 0;
            const change = prevRev !== null ? rev - prevRev : 0;
            const pct = prevRev && prevRev > 0 ? ((change / prevRev) * 100).toFixed(1) : (prevRev === null ? '—' : '0.0');
            const pctStr = pct === '—' ? '—' : (parseFloat(pct) >= 0 ? '+' + pct + '%' : pct + '%');
            const cls = pct === '—' ? '' : (parseFloat(pct) >= 0 ? 'positive' : 'negative');
            html += '<tr><td class="text-cell">' + this._kpiDetailMonthLabel(month) + '</td>' +
                '<td>' + Utils.formatCurrency(rev) + '</td>' +
                '<td class="' + cls + '">' + (prevRev !== null ? Utils.formatCurrency(change) : '—') + '</td>' +
                '<td class="' + cls + '">' + pctStr + '</td></tr>';
            prevRev = rev;
        }
        html += '</tbody></table>';
        return html;
    },

    _kpiDetail_cmgr() {
        const currentMonth = this._kpiCache ? this._kpiCache.currentMonth : Utils.getCurrentMonth();
        const months = this._getTimelineMonths().filter(m => m <= currentMonth);

        // Get monthly revenue
        const revData = [];
        for (const month of months) {
            const revResult = Database.db.exec(
                "SELECT COALESCE(SUM(COALESCE(t.pretax_amount, t.amount)),0) FROM transactions t JOIN categories c ON t.category_id = c.id WHERE t.transaction_type='receivable' AND t.month_due=? AND c.is_cogs = 0 AND c.show_on_pl != 1 AND (c.is_b2b = 1 OR c.is_sales = 1) AND COALESCE(t.source_type, '') NOT IN ('loan_receivable', 'loan_payment')", [month]);
            revData.push(revResult[0] ? revResult[0].values[0][0] : 0);
        }

        // Compute CMGR over different windows
        const computeCmgr = (data) => {
            const first = data.find(v => v > 0) || 0;
            const last = data[data.length - 1] || 0;
            if (first <= 0 || last <= 0 || data.length < 2) return 0;
            return (Math.pow(last / first, 1 / (data.length - 1)) - 1) * 100;
        };

        const cmgr3 = months.length >= 3 ? computeCmgr(revData.slice(-3)) : null;
        const cmgr6 = months.length >= 6 ? computeCmgr(revData.slice(-6)) : null;
        const cmgrAll = computeCmgr(revData);

        let html = '<div class="kpi-detail-summary">Compound Monthly Growth Rate across periods</div>';

        // Summary cards
        html += '<div style="display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap;">';
        if (cmgr3 !== null) {
            html += '<div style="flex:1;min-width:120px;padding:12px;border-radius:8px;background:var(--c5,var(--border));text-align:center;">' +
                '<div style="font-size:0.75rem;color:var(--text-muted);margin-bottom:4px;">3-Month CMGR</div>' +
                '<div style="font-size:1.2rem;font-weight:700;font-family:DM Mono,monospace;color:' + (cmgr3 >= 0 ? 'var(--color-success)' : 'var(--color-danger)') + ';">' + (cmgr3 >= 0 ? '+' : '') + cmgr3.toFixed(1) + '%</div></div>';
        }
        if (cmgr6 !== null) {
            html += '<div style="flex:1;min-width:120px;padding:12px;border-radius:8px;background:var(--c5,var(--border));text-align:center;">' +
                '<div style="font-size:0.75rem;color:var(--text-muted);margin-bottom:4px;">6-Month CMGR</div>' +
                '<div style="font-size:1.2rem;font-weight:700;font-family:DM Mono,monospace;color:' + (cmgr6 >= 0 ? 'var(--color-success)' : 'var(--color-danger)') + ';">' + (cmgr6 >= 0 ? '+' : '') + cmgr6.toFixed(1) + '%</div></div>';
        }
        html += '<div style="flex:1;min-width:120px;padding:12px;border-radius:8px;background:var(--c5,var(--border));text-align:center;">' +
            '<div style="font-size:0.75rem;color:var(--text-muted);margin-bottom:4px;">Full Period CMGR</div>' +
            '<div style="font-size:1.2rem;font-weight:700;font-family:DM Mono,monospace;color:' + (cmgrAll >= 0 ? 'var(--color-success)' : 'var(--color-danger)') + ';">' + (cmgrAll >= 0 ? '+' : '') + cmgrAll.toFixed(1) + '%</div></div>';
        html += '</div>';

        // Monthly table with MoM growth
        html += '<table class="kpi-detail-table"><thead><tr><th>Month</th><th>Revenue</th><th>MoM Change</th><th>MoM %</th></tr></thead><tbody>';
        months.forEach((m, i) => {
            const prev = i > 0 ? revData[i - 1] : null;
            const change = prev !== null ? revData[i] - prev : null;
            const pct = prev && prev > 0 ? ((change / prev) * 100).toFixed(1) : null;
            const pctStr = pct !== null ? ((parseFloat(pct) >= 0 ? '+' : '') + pct + '%') : '—';
            const changeStr = change !== null ? Utils.formatCurrency(change) : '—';
            const cls = pct === null ? '' : (parseFloat(pct) >= 0 ? 'positive' : 'negative');
            html += '<tr><td class="text-cell">' + this._kpiDetailMonthLabel(m) + '</td>' +
                '<td>' + Utils.formatCurrency(revData[i]) + '</td>' +
                '<td class="' + cls + '">' + changeStr + '</td>' +
                '<td class="' + cls + '">' + pctStr + '</td></tr>';
        });
        html += '</tbody></table>';
        return html;
    },

    _kpiDetail_cmgrnonb2b() {
        const currentMonth = this._kpiCache ? this._kpiCache.currentMonth : Utils.getCurrentMonth();
        const months = this._getTimelineMonths().filter(m => m <= currentMonth);

        const revData = [];
        for (const month of months) {
            const revResult = Database.db.exec(
                "SELECT COALESCE(SUM(COALESCE(t.pretax_amount, t.amount)),0) FROM transactions t JOIN categories c ON t.category_id = c.id " +
                "WHERE t.transaction_type='receivable' AND t.month_due=? AND c.is_cogs = 0 AND c.show_on_pl != 1 AND c.is_sales = 1 AND c.is_b2b = 0 " +
                "AND COALESCE(t.source_type, '') NOT IN ('loan_receivable', 'loan_payment')", [month]);
            revData.push(revResult[0] ? revResult[0].values[0][0] : 0);
        }

        const computeCmgr = (data) => {
            const firstIdx = data.findIndex(v => v > 0);
            if (firstIdx < 0) return 0;
            const first = data[firstIdx];
            const last = data[data.length - 1];
            const n = data.length - 1 - firstIdx;
            if (first <= 0 || last <= 0 || n < 1) return 0;
            return (Math.pow(last / first, 1 / n) - 1) * 100;
        };

        const cmgr3 = months.length >= 3 ? computeCmgr(revData.slice(-3)) : null;
        const cmgr6 = months.length >= 6 ? computeCmgr(revData.slice(-6)) : null;
        const cmgrAll = computeCmgr(revData);

        let html = '<div class="kpi-detail-summary">Compound Monthly Growth Rate — Non-B2B (consumer) sales only</div>';

        html += '<div style="display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap;">';
        if (cmgr3 !== null) {
            html += '<div style="flex:1;min-width:120px;padding:12px;border-radius:8px;background:var(--c5,var(--border));text-align:center;">' +
                '<div style="font-size:0.75rem;color:var(--text-muted);margin-bottom:4px;">3-Month CMGR</div>' +
                '<div style="font-size:1.2rem;font-weight:700;font-family:DM Mono,monospace;color:' + (cmgr3 >= 0 ? 'var(--color-success)' : 'var(--color-danger)') + ';">' + (cmgr3 >= 0 ? '+' : '') + cmgr3.toFixed(1) + '%</div></div>';
        }
        if (cmgr6 !== null) {
            html += '<div style="flex:1;min-width:120px;padding:12px;border-radius:8px;background:var(--c5,var(--border));text-align:center;">' +
                '<div style="font-size:0.75rem;color:var(--text-muted);margin-bottom:4px;">6-Month CMGR</div>' +
                '<div style="font-size:1.2rem;font-weight:700;font-family:DM Mono,monospace;color:' + (cmgr6 >= 0 ? 'var(--color-success)' : 'var(--color-danger)') + ';">' + (cmgr6 >= 0 ? '+' : '') + cmgr6.toFixed(1) + '%</div></div>';
        }
        html += '<div style="flex:1;min-width:120px;padding:12px;border-radius:8px;background:var(--c5,var(--border));text-align:center;">' +
            '<div style="font-size:0.75rem;color:var(--text-muted);margin-bottom:4px;">Full Period CMGR</div>' +
            '<div style="font-size:1.2rem;font-weight:700;font-family:DM Mono,monospace;color:' + (cmgrAll >= 0 ? 'var(--color-success)' : 'var(--color-danger)') + ';">' + (cmgrAll >= 0 ? '+' : '') + cmgrAll.toFixed(1) + '%</div></div>';
        html += '</div>';

        const catResult = Database.db.exec(
            "SELECT c.name, t.month_due, SUM(COALESCE(t.pretax_amount, t.amount)) as total " +
            "FROM transactions t JOIN categories c ON t.category_id = c.id " +
            "WHERE t.transaction_type='receivable' AND c.is_cogs = 0 AND c.show_on_pl != 1 AND c.is_sales = 1 AND c.is_b2b = 0 " +
            "AND COALESCE(t.source_type, '') NOT IN ('loan_receivable', 'loan_payment') " +
            "AND t.month_due IN (" + months.map(() => '?').join(',') + ") " +
            "GROUP BY c.name, t.month_due ORDER BY total DESC", months);

        const catMap = {};
        if (catResult[0]) {
            for (const row of catResult[0].values) {
                if (!catMap[row[0]]) catMap[row[0]] = {};
                catMap[row[0]][row[1]] = row[2];
            }
        }
        const catTotals = Object.entries(catMap).map(([name, mData]) => ({
            name, total: Object.values(mData).reduce((s, v) => s + v, 0), mData
        })).sort((a, b) => b.total - a.total);

        html += '<table class="kpi-detail-table"><thead><tr><th>Month</th><th>Revenue</th><th>MoM Change</th><th>MoM %</th></tr></thead><tbody>';
        months.forEach((m, i) => {
            const prev = i > 0 ? revData[i - 1] : null;
            const change = prev !== null ? revData[i] - prev : null;
            const pct = prev && prev > 0 ? ((change / prev) * 100).toFixed(1) : null;
            const pctStr = pct !== null ? ((parseFloat(pct) >= 0 ? '+' : '') + pct + '%') : '—';
            const changeStr = change !== null ? Utils.formatCurrency(change) : '—';
            const cls = pct === null ? '' : (parseFloat(pct) >= 0 ? 'positive' : 'negative');
            html += '<tr><td class="text-cell">' + this._kpiDetailMonthLabel(m) + '</td>' +
                '<td>' + Utils.formatCurrency(revData[i]) + '</td>' +
                '<td class="' + cls + '">' + changeStr + '</td>' +
                '<td class="' + cls + '">' + pctStr + '</td></tr>';
        });
        html += '</tbody></table>';

        if (catTotals.length > 1) {
            html += '<div style="margin-top:16px;font-weight:600;font-size:0.85rem;margin-bottom:8px;">By Category</div>';
            html += '<table class="kpi-detail-table"><thead><tr><th>Category</th><th>Total</th><th>Share</th></tr></thead><tbody>';
            const grandTotal = catTotals.reduce((s, c) => s + c.total, 0);
            catTotals.forEach(cat => {
                const share = grandTotal > 0 ? ((cat.total / grandTotal) * 100).toFixed(1) : '0.0';
                html += '<tr><td class="text-cell">' + Utils.escapeHtml(cat.name) + '</td>' +
                    '<td>' + Utils.formatCurrency(cat.total) + '</td>' +
                    '<td>' + share + '%</td></tr>';
            });
            html += '</tbody></table>';
        }

        return html;
    },

    _kpiDetail_ebitda() {
        const currentMonth = this._kpiCache ? this._kpiCache.currentMonth : Utils.getCurrentMonth();
        const pastMonths = this._getTimelineMonths().filter(m => m <= currentMonth);

        const plData = Database.getPLSpreadsheet();
        const plOverrides = Database.getAllPLOverrides();

        const getVal = (catId, month, computed) => {
            const key = `${catId}-${month}`;
            return (key in plOverrides) ? plOverrides[key] : computed;
        };
        const groupByCat = (rows) => {
            const map = {};
            (rows || []).forEach(row => {
                if (!map[row.category_id]) map[row.category_id] = {};
                map[row.category_id][row.month] = (map[row.category_id][row.month] || 0) + row.total;
            });
            return map;
        };

        let totalRevenue = 0;
        const revByCat = groupByCat(plData.revenue);
        Object.entries(revByCat).forEach(([catId, catMonths]) => {
            pastMonths.forEach(m => { totalRevenue += getVal(catId, m, catMonths[m] || 0); });
        });

        let totalCogs = 0;
        const cogsByCat = groupByCat(plData.cogs);
        Object.entries(cogsByCat).forEach(([catId, catMonths]) => {
            pastMonths.forEach(m => { totalCogs += getVal(catId, m, catMonths[m] || 0); });
        });

        let totalOpex = 0;
        const opexByCat = groupByCat(plData.opex);
        Object.entries(opexByCat).forEach(([catId, catMonths]) => {
            pastMonths.forEach(m => { totalOpex += getVal(catId, m, catMonths[m] || 0); });
        });

        // Items excluded from EBITDA
        let totalDepr = 0;
        (plData.depreciation || []).forEach(cat => {
            pastMonths.forEach(m => { totalDepr += getVal(cat.category_id, m, 0); });
        });
        const assetDepr = plData.assetDeprByMonth || {};
        pastMonths.forEach(m => { totalDepr += (assetDepr[m] || 0); });

        let totalInterest = 0;
        const loanInt = plData.loanInterestByMonth || {};
        pastMonths.forEach(m => { totalInterest += (loanInt[m] || 0); });

        const grossProfit = totalRevenue - totalCogs;
        const ebitda = totalRevenue - totalCogs - totalOpex;
        const ebitdaMargin = totalRevenue > 0 ? ((ebitda / totalRevenue) * 100).toFixed(1) : '0.0';

        let html = '<div class="kpi-detail-summary">EBITDA computation — cumulative through ' + this._kpiDetailMonthLabel(pastMonths[pastMonths.length - 1] || currentMonth) + '</div>';
        html += '<table class="kpi-detail-table"><tbody>';
        html += '<tr><td class="text-cell">Revenue</td><td class="positive">' + Utils.formatCurrency(totalRevenue) + '</td></tr>';
        html += '<tr><td class="text-cell">Less: Cost of Goods Sold</td><td class="negative">(' + Utils.formatCurrency(totalCogs) + ')</td></tr>';
        html += '<tr class="kpi-row-total"><td class="text-cell">Gross Profit</td><td>' + Utils.formatCurrency(grossProfit) + '</td></tr>';
        html += '<tr><td class="text-cell">Less: Operating Expenses</td><td class="negative">(' + Utils.formatCurrency(totalOpex) + ')</td></tr>';
        html += '<tr class="kpi-row-total"><td class="text-cell" style="font-size:0.95rem;">EBITDA</td><td style="font-size:0.95rem;" class="' + (ebitda >= 0 ? 'positive' : 'negative') + '">' + Utils.formatCurrency(ebitda) + '</td></tr>';
        html += '<tr><td colspan="2" style="padding:8px;"></td></tr>';
        html += '<tr style="opacity:0.6;"><td class="text-cell">Excluded: Depreciation & Amortization</td><td>' + Utils.formatCurrency(totalDepr) + '</td></tr>';
        html += '<tr style="opacity:0.6;"><td class="text-cell">Excluded: Interest</td><td>' + Utils.formatCurrency(totalInterest) + '</td></tr>';
        html += '<tr><td colspan="2" style="padding:4px;"></td></tr>';
        html += '<tr><td class="text-cell" style="font-weight:600;">EBITDA Margin</td><td style="font-weight:600;">' + ebitdaMargin + '%</td></tr>';
        html += '</tbody></table>';
        return html;
    },

    _kpiDetail_overdue() {
        const currentMonth = this._kpiCache ? this._kpiCache.currentMonth : Utils.getCurrentMonth();
        const result = Database.db.exec(
            "SELECT t.item_description, c.name as category_name, t.amount, t.month_due, t.entry_date " +
            "FROM transactions t JOIN categories c ON t.category_id = c.id " +
            "WHERE t.transaction_type='receivable' AND t.status='pending' AND t.month_due < ? " +
            "ORDER BY t.month_due ASC", [currentMonth]);

        if (!result[0] || result[0].values.length === 0) {
            return '<div class="kpi-detail-summary" style="color:var(--color-success);">No overdue receivables — all caught up!</div>';
        }

        const rows = result[0].values;
        let totalAmount = 0;
        let html = '<div class="kpi-detail-summary">' + rows.length + ' overdue receivable' + (rows.length !== 1 ? 's' : '') + ' — sorted oldest first</div>';
        html += '<table class="kpi-detail-table"><thead><tr><th>Description</th><th>Category</th><th>Amount</th><th>Due</th><th>Overdue</th></tr></thead><tbody>';

        for (const row of rows) {
            const desc = row[0] || '(no description)';
            const catName = row[1];
            const amount = row[2];
            const monthDue = row[3];
            totalAmount += amount;

            // Calculate months overdue
            const [dueY, dueM] = monthDue.split('-').map(Number);
            const [curY, curM] = currentMonth.split('-').map(Number);
            const monthsOverdue = (curY - dueY) * 12 + (curM - dueM);
            const overdueLabel = monthsOverdue === 1 ? '1 month' : monthsOverdue + ' months';
            const urgencyCls = monthsOverdue >= 3 ? 'negative' : (monthsOverdue >= 2 ? 'dash-kpi-warning' : '');

            html += '<tr><td class="text-cell">' + Utils.escapeHtml(desc) + '</td>' +
                '<td class="text-cell">' + Utils.escapeHtml(catName) + '</td>' +
                '<td class="negative">' + Utils.formatCurrency(amount) + '</td>' +
                '<td class="text-cell">' + this._kpiDetailMonthLabel(monthDue) + '</td>' +
                '<td class="text-cell ' + urgencyCls + '" style="font-weight:600;">' + overdueLabel + '</td></tr>';
        }

        html += '<tr class="kpi-row-total"><td class="text-cell" colspan="2">Total Overdue</td>' +
            '<td class="negative">' + Utils.formatCurrency(totalAmount) + '</td><td colspan="2"></td></tr>';
        html += '</tbody></table>';
        return html;
    },

    _kpiDetail_workingcapital() {
        const currentMonth = this._kpiCache ? this._kpiCache.currentMonth : Utils.getCurrentMonth();
        const months = this._getTimelineMonths().filter(m => m <= currentMonth);

        let html = '<div class="kpi-detail-summary">Working Capital = Current Assets − Current Liabilities</div>';
        html += '<table class="kpi-detail-table"><thead><tr><th>Month</th><th>Cash</th><th>AR</th><th>Current Assets</th><th>AP + Tax</th><th>Working Capital</th></tr></thead><tbody>';

        let prevWC = null;
        for (const month of months) {
            const cash = Database.getCashAsOf(month);
            const ar = Database.getAccountsReceivableAsOf ? Database.getAccountsReceivableAsOf(month) : 0;
            const ap = Database.getAccountsPayableAsOf ? Database.getAccountsPayableAsOf(month) : 0;
            const stp = Database.getSalesTaxPayableAsOf ? Database.getSalesTaxPayableAsOf(month) : 0;
            const ca = cash + ar;
            const cl = ap + stp;
            const wc = ca - cl;
            const trend = prevWC !== null && prevWC !== 0 ? ((wc - prevWC) / Math.abs(prevWC) * 100).toFixed(1) : null;
            const trendStr = trend !== null ? (' <small style="font-size:0.8em;color:' + (parseFloat(trend) >= 0 ? 'var(--color-success)' : 'var(--color-danger)') + '">' + (parseFloat(trend) >= 0 ? '+' : '') + trend + '%</small>') : '';

            html += '<tr><td class="text-cell">' + this._kpiDetailMonthLabel(month) + '</td>' +
                '<td>' + Utils.formatCurrency(cash) + '</td>' +
                '<td>' + Utils.formatCurrency(ar) + '</td>' +
                '<td style="font-weight:500;">' + Utils.formatCurrency(ca) + '</td>' +
                '<td class="negative">' + Utils.formatCurrency(cl) + '</td>' +
                '<td class="' + (wc >= 0 ? 'positive' : 'negative') + '" style="font-weight:600;">' + Utils.formatCurrency(wc) + trendStr + '</td></tr>';
            prevWC = wc;
        }
        html += '</tbody></table>';
        return html;
    },

    _kpiDetail_rule40() {
        const currentMonth = this._kpiCache ? this._kpiCache.currentMonth : Utils.getCurrentMonth();
        const months = this._getTimelineMonths().filter(m => m <= currentMonth);

        // Compute revenue growth (annualized CMGR)
        const revData = [];
        for (const month of months) {
            const revResult = Database.db.exec(
                "SELECT COALESCE(SUM(COALESCE(t.pretax_amount, t.amount)),0) FROM transactions t JOIN categories c ON t.category_id = c.id WHERE t.transaction_type='receivable' AND t.month_due=? AND c.is_cogs = 0 AND c.show_on_pl != 1 AND (c.is_b2b = 1 OR c.is_sales = 1) AND COALESCE(t.source_type, '') NOT IN ('loan_receivable', 'loan_payment')", [month]);
            revData.push(revResult[0] ? revResult[0].values[0][0] : 0);
        }
        const firstRev = revData.find(v => v > 0) || 0;
        const lastRev = revData[revData.length - 1] || 0;
        let cmgr = 0;
        if (firstRev > 0 && lastRev > 0 && revData.length > 1) {
            cmgr = (Math.pow(lastRev / firstRev, 1 / (revData.length - 1)) - 1) * 100;
        }
        const annualizedGrowth = cmgr * 12;

        // Compute profit margin (EBITDA margin)
        const plData = Database.getPLSpreadsheet();
        const plOverrides = Database.getAllPLOverrides();
        const getVal = (catId, month, computed) => {
            const key = `${catId}-${month}`;
            return (key in plOverrides) ? plOverrides[key] : computed;
        };
        const groupByCat = (rows) => {
            const map = {};
            (rows || []).forEach(row => {
                if (!map[row.category_id]) map[row.category_id] = {};
                map[row.category_id][row.month] = (map[row.category_id][row.month] || 0) + row.total;
            });
            return map;
        };

        let totalRevenue = 0;
        Object.entries(groupByCat(plData.revenue)).forEach(([catId, catMonths]) => {
            months.forEach(m => { totalRevenue += getVal(catId, m, catMonths[m] || 0); });
        });
        let totalExpenses = 0;
        Object.entries(groupByCat(plData.cogs)).forEach(([catId, catMonths]) => {
            months.forEach(m => { totalExpenses += getVal(catId, m, catMonths[m] || 0); });
        });
        Object.entries(groupByCat(plData.opex)).forEach(([catId, catMonths]) => {
            months.forEach(m => { totalExpenses += getVal(catId, m, catMonths[m] || 0); });
        });
        const ebitda = totalRevenue - totalExpenses;
        const profitMargin = totalRevenue > 0 ? (ebitda / totalRevenue) * 100 : 0;
        const score = Math.round(annualizedGrowth + profitMargin);

        let html = '<div class="kpi-detail-summary">Rule of 40 = Annualized Revenue Growth % + EBITDA Margin %</div>';

        // Visual gauge
        const barPct = Math.min(100, Math.max(0, (score / 60) * 100));
        const barColor = score >= 40 ? 'var(--color-success)' : score >= 20 ? 'var(--color-warning)' : 'var(--color-danger)';
        html += '<div style="margin:16px 0 20px;">' +
            '<div style="display:flex;justify-content:space-between;margin-bottom:4px;font-size:0.8rem;color:var(--text-muted);"><span>0</span><span style="color:var(--color-success);font-weight:600;">40 (target)</span><span>60</span></div>' +
            '<div style="height:24px;border-radius:8px;background:var(--c5,var(--border));position:relative;overflow:hidden;">' +
                '<div style="width:' + barPct + '%;height:100%;border-radius:8px;background:' + barColor + ';transition:width 0.3s;"></div>' +
                '<div style="position:absolute;left:' + (40/60*100) + '%;top:0;bottom:0;width:2px;background:var(--text-muted);opacity:0.5;"></div>' +
            '</div>' +
            '<div style="text-align:center;margin-top:8px;font-size:1.3rem;font-weight:700;font-family:DM Mono,monospace;color:' + barColor + ';">' + score + '</div>' +
        '</div>';

        html += '<table class="kpi-detail-table"><tbody>';
        html += '<tr><td class="text-cell">Annualized Revenue Growth (CMGR × 12)</td><td class="' + (annualizedGrowth >= 0 ? 'positive' : 'negative') + '">' + annualizedGrowth.toFixed(1) + '%</td></tr>';
        html += '<tr><td class="text-cell">EBITDA Margin</td><td class="' + (profitMargin >= 0 ? 'positive' : 'negative') + '">' + profitMargin.toFixed(1) + '%</td></tr>';
        html += '<tr class="kpi-row-total"><td class="text-cell">Rule of 40 Score</td><td style="font-size:1rem;" class="' + (score >= 40 ? 'positive' : 'negative') + '">' + score + '</td></tr>';
        html += '</tbody></table>';

        html += '<div style="margin-top:12px;font-size:0.8rem;color:var(--text-muted);">' +
            (score >= 40 ? 'Score ≥ 40 indicates a healthy balance of growth and profitability.' :
            'Score < 40 suggests the company needs stronger growth, better margins, or both.') + '</div>';
        return html;
    },

    _kpiDetail_rule40nb() {
        const currentMonth = this._kpiCache ? this._kpiCache.currentMonth : Utils.getCurrentMonth();
        const months = this._getTimelineMonths().filter(m => m <= currentMonth);

        // Non-B2B revenue growth (CMGR)
        const revData = [];
        for (const month of months) {
            const revResult = Database.db.exec(
                "SELECT COALESCE(SUM(COALESCE(t.pretax_amount, t.amount)),0) FROM transactions t JOIN categories c ON t.category_id = c.id " +
                "WHERE t.transaction_type='receivable' AND t.month_due=? AND c.is_cogs = 0 AND c.show_on_pl != 1 AND c.is_sales = 1 AND c.is_b2b = 0 " +
                "AND COALESCE(t.source_type, '') NOT IN ('loan_receivable', 'loan_payment')", [month]);
            revData.push(revResult[0] ? revResult[0].values[0][0] : 0);
        }
        const firstRev = revData.find(v => v > 0) || 0;
        const lastRev = revData[revData.length - 1] || 0;
        let cmgr = 0;
        if (firstRev > 0 && lastRev > 0 && revData.length > 1) {
            cmgr = (Math.pow(lastRev / firstRev, 1 / (revData.length - 1)) - 1) * 100;
        }
        const annualizedGrowth = cmgr * 12;

        // Profit margin (EBITDA margin — same as overall, since margin is company-wide)
        const plData = Database.getPLSpreadsheet();
        const plOverrides = Database.getAllPLOverrides();
        const getVal = (catId, month, computed) => {
            const key = `${catId}-${month}`;
            return (key in plOverrides) ? plOverrides[key] : computed;
        };
        const groupByCat = (rows) => {
            const map = {};
            (rows || []).forEach(row => {
                if (!map[row.category_id]) map[row.category_id] = {};
                map[row.category_id][row.month] = (map[row.category_id][row.month] || 0) + row.total;
            });
            return map;
        };

        let totalRevenue = 0;
        Object.entries(groupByCat(plData.revenue)).forEach(([catId, catMonths]) => {
            months.forEach(m => { totalRevenue += getVal(catId, m, catMonths[m] || 0); });
        });
        let totalExpenses = 0;
        Object.entries(groupByCat(plData.cogs)).forEach(([catId, catMonths]) => {
            months.forEach(m => { totalExpenses += getVal(catId, m, catMonths[m] || 0); });
        });
        Object.entries(groupByCat(plData.opex)).forEach(([catId, catMonths]) => {
            months.forEach(m => { totalExpenses += getVal(catId, m, catMonths[m] || 0); });
        });
        const ebitda = totalRevenue - totalExpenses;
        const profitMargin = totalRevenue > 0 ? (ebitda / totalRevenue) * 100 : 0;
        const score = Math.round(annualizedGrowth + profitMargin);

        let html = '<div class="kpi-detail-summary">Rule of 40 (Non-B2B) = Annualized Consumer Sales Growth % + EBITDA Margin %</div>';

        // Visual gauge
        const barPct = Math.min(100, Math.max(0, (score / 60) * 100));
        const barColor = score >= 40 ? 'var(--color-success)' : score >= 20 ? 'var(--color-warning)' : 'var(--color-danger)';
        html += '<div style="margin:16px 0 20px;">' +
            '<div style="display:flex;justify-content:space-between;margin-bottom:4px;font-size:0.8rem;color:var(--text-muted);"><span>0</span><span style="color:var(--color-success);font-weight:600;">40 (target)</span><span>60</span></div>' +
            '<div style="height:24px;border-radius:8px;background:var(--c5,var(--border));position:relative;overflow:hidden;">' +
                '<div style="width:' + barPct + '%;height:100%;border-radius:8px;background:' + barColor + ';transition:width 0.3s;"></div>' +
                '<div style="position:absolute;left:' + (40/60*100) + '%;top:0;bottom:0;width:2px;background:var(--text-muted);opacity:0.5;"></div>' +
            '</div>' +
            '<div style="text-align:center;margin-top:8px;font-size:1.3rem;font-weight:700;font-family:DM Mono,monospace;color:' + barColor + ';">' + score + '</div>' +
        '</div>';

        html += '<table class="kpi-detail-table"><tbody>';
        html += '<tr><td class="text-cell">Non-B2B Revenue Growth (CMGR × 12)</td><td class="' + (annualizedGrowth >= 0 ? 'positive' : 'negative') + '">' + annualizedGrowth.toFixed(1) + '%</td></tr>';
        html += '<tr><td class="text-cell">EBITDA Margin</td><td class="' + (profitMargin >= 0 ? 'positive' : 'negative') + '">' + profitMargin.toFixed(1) + '%</td></tr>';
        html += '<tr class="kpi-row-total"><td class="text-cell">Rule of 40 Score</td><td style="font-size:1rem;" class="' + (score >= 40 ? 'positive' : 'negative') + '">' + score + '</td></tr>';
        html += '</tbody></table>';

        html += '<div style="margin-top:12px;font-size:0.8rem;color:var(--text-muted);">' +
            'Uses non-B2B (consumer) sales CMGR for the growth component. EBITDA margin is company-wide. ' +
            (score >= 40 ? 'Score ≥ 40 indicates a healthy balance of growth and profitability.' :
            'Score < 40 suggests the company needs stronger consumer growth, better margins, or both.') + '</div>';
        return html;
    },

    _kpiDetail_dscr() {
        const currentMonth = this._kpiCache ? this._kpiCache.currentMonth : Utils.getCurrentMonth();
        const months = this._getTimelineMonths().filter(m => m <= currentMonth);

        const loans = Database.getLoans();
        if (loans.length === 0) {
            return '<div class="kpi-detail-summary" style="color:var(--color-success);">No active loans — DSCR is not applicable.</div>';
        }

        // Compute EBITDA (NOI)
        const plData = Database.getPLSpreadsheet();
        const plOverrides = Database.getAllPLOverrides();
        const getVal = (catId, month, computed) => {
            const key = `${catId}-${month}`;
            return (key in plOverrides) ? plOverrides[key] : computed;
        };
        const groupByCat = (rows) => {
            const map = {};
            (rows || []).forEach(row => {
                if (!map[row.category_id]) map[row.category_id] = {};
                map[row.category_id][row.month] = (map[row.category_id][row.month] || 0) + row.total;
            });
            return map;
        };

        let totalRevenue = 0;
        Object.entries(groupByCat(plData.revenue)).forEach(([catId, catMonths]) => {
            months.forEach(m => { totalRevenue += getVal(catId, m, catMonths[m] || 0); });
        });
        let totalExpenses = 0;
        Object.entries(groupByCat(plData.cogs)).forEach(([catId, catMonths]) => {
            months.forEach(m => { totalExpenses += getVal(catId, m, catMonths[m] || 0); });
        });
        Object.entries(groupByCat(plData.opex)).forEach(([catId, catMonths]) => {
            months.forEach(m => { totalExpenses += getVal(catId, m, catMonths[m] || 0); });
        });
        const noi = totalRevenue - totalExpenses;

        // Compute debt service per loan
        let totalDebtService = 0;
        let html = '<div class="kpi-detail-summary">DSCR = Net Operating Income / Total Debt Service</div>';
        html += '<table class="kpi-detail-table"><thead><tr><th>Loan</th><th>Principal Paid</th><th>Interest Paid</th><th>Total Service</th></tr></thead><tbody>';

        loans.forEach(loan => {
            const skipped = Database.getSkippedPayments(loan.id);
            const overridesLoan = Database.getLoanPaymentOverrides(loan.id);
            const schedule = Utils.computeAmortizationSchedule({
                principal: loan.principal, annual_rate: loan.annual_rate,
                term_months: loan.term_months, payments_per_year: loan.payments_per_year,
                start_date: loan.start_date, first_payment_date: loan.first_payment_date
            }, skipped, overridesLoan);

            let loanPrincipal = 0, loanInterest = 0;
            schedule.forEach(pmt => {
                if (pmt.month <= currentMonth && months.includes(pmt.month)) {
                    loanPrincipal += pmt.principal_payment;
                    loanInterest += pmt.interest_payment;
                }
            });
            const loanTotal = loanPrincipal + loanInterest;
            totalDebtService += loanTotal;

            html += '<tr><td class="text-cell">' + Utils.escapeHtml(loan.name || 'Loan #' + loan.id) + '</td>' +
                '<td>' + Utils.formatCurrency(loanPrincipal) + '</td>' +
                '<td>' + Utils.formatCurrency(loanInterest) + '</td>' +
                '<td style="font-weight:500;">' + Utils.formatCurrency(loanTotal) + '</td></tr>';
        });

        html += '<tr class="kpi-row-total"><td class="text-cell">Total</td><td colspan="2"></td>' +
            '<td>' + Utils.formatCurrency(totalDebtService) + '</td></tr>';
        html += '</tbody></table>';

        const dscr = totalDebtService > 0 ? noi / totalDebtService : null;
        const dscrStr = dscr !== null ? dscr.toFixed(2) + 'x' : 'N/A';
        const dscrCls = dscr === null ? '' : (dscr >= 1.25 ? 'positive' : (dscr >= 1.0 ? 'dash-kpi-warning' : 'negative'));

        html += '<table class="kpi-detail-table" style="margin-top:16px;"><tbody>';
        html += '<tr><td class="text-cell">Net Operating Income (EBITDA)</td><td class="' + (noi >= 0 ? 'positive' : 'negative') + '">' + Utils.formatCurrency(noi) + '</td></tr>';
        html += '<tr><td class="text-cell">Total Debt Service</td><td class="negative">' + Utils.formatCurrency(totalDebtService) + '</td></tr>';
        html += '<tr class="kpi-row-total"><td class="text-cell" style="font-size:0.95rem;">DSCR</td><td style="font-size:0.95rem;" class="' + dscrCls + '">' + dscrStr + '</td></tr>';
        html += '</tbody></table>';

        html += '<div style="margin-top:12px;font-size:0.8rem;color:var(--text-muted);">' +
            (dscr === null ? '' : dscr >= 1.25 ? 'DSCR ≥ 1.25x — strong ability to service debt.' :
            dscr >= 1.0 ? 'DSCR 1.0-1.25x — just covering debt obligations. Lenders prefer ≥ 1.25x.' :
            'DSCR < 1.0x — not generating enough income to cover debt payments.') + '</div>';
        return html;
    },

    _renderSparkline(canvasId, data, color) {
        const canvas = document.getElementById(canvasId);
        if (!canvas || !data.length) return;
        const ctx = canvas.getContext('2d');

        // Destroy existing chart
        if (this._dashCharts[canvasId]) {
            this._dashCharts[canvasId].destroy();
        }

        this._dashCharts[canvasId] = new Chart(ctx, {
            type: 'line',
            data: {
                labels: data.map(() => ''),
                datasets: [{
                    data: data,
                    borderColor: color,
                    borderWidth: 1.5,
                    fill: true,
                    backgroundColor: color.replace('0.8', '0.1'),
                    pointRadius: 0,
                    tension: 0.4
                }]
            },
            options: {
                responsive: false,
                plugins: { legend: { display: false }, tooltip: { enabled: false } },
                scales: { x: { display: false }, y: { display: false } },
                animation: false
            }
        });
    },

    _renderDashboardSections(snapshotMonth) {
        const container = document.getElementById('dashboardSections');
        if (!container) return;

        // Only rebuild the HTML structure if it's empty
        if (!container.querySelector('.dash-section')) {
            const hasKpis = (section) => !!this._sectionKpiMap[section];
            const kpiHint = '<span class="dash-section-kpi-hint">Click for KPIs</span>';

            container.innerHTML =
                '<div class="dash-section dash-section-clickable" data-section="cashflow">' +
                    '<div class="dash-section-header"><span>Cash Flow Overview</span>' + kpiHint + '</div>' +
                    '<div class="dash-section-body"><canvas id="dashChartCashFlow" height="250"></canvas></div>' +
                '</div>' +
                '<div class="dash-section dash-section-clickable" data-section="pnl">' +
                    '<div class="dash-section-header"><span>P&L Trends</span>' + kpiHint + '</div>' +
                    '<div class="dash-section-body"><canvas id="dashChartPnL" height="250"></canvas></div>' +
                '</div>' +
                '<div class="dash-section dash-section-clickable" data-section="balancesheet">' +
                    '<div class="dash-section-header"><span>Balance Sheet Snapshot</span>' + kpiHint + '</div>' +
                    '<div class="dash-section-body"><canvas id="dashChartBS" height="250"></canvas></div>' +
                '</div>' +
                '<div class="dash-section" data-section="breakeven">' +
                    '<div class="dash-section-header"><span>Break-Even Progress</span></div>' +
                    '<div class="dash-section-body"><div id="dashBreakevenBar" class="dash-be-bar-wrapper"></div></div>' +
                '</div>' +
                '<div class="dash-section dash-section-clickable" data-section="revconcentration">' +
                    '<div class="dash-section-header"><span>Revenue Concentration</span>' + kpiHint + '</div>' +
                    '<div class="dash-section-body"><canvas id="dashChartRevConc" height="250"></canvas></div>' +
                '</div>';

            // Wire up clickable sections to open KPI modals
            container.querySelectorAll('.dash-section-clickable').forEach(section => {
                section.addEventListener('click', () => {
                    this._openSectionAnalysis(section.dataset.section);
                });
            });
        }

        // Render charts
        this._renderDashCashFlowChart(snapshotMonth);
        this._renderDashPnLChart(snapshotMonth);
        this._renderDashBSChart(snapshotMonth);
        this._renderDashBreakevenBar(snapshotMonth);
        this._renderDashRevConcentrationChart(snapshotMonth);
    },

    _getTimelineMonths() {
        const timeline = this.getTimeline();
        const currentMonth = Utils.getCurrentMonth();
        let start, end;

        if (timeline.start && timeline.end) {
            start = timeline.start;
            end = timeline.end;
        } else if (timeline.start) {
            start = timeline.start;
            end = currentMonth;
        } else if (timeline.end) {
            // Default to 12 months before the end
            end = timeline.end;
            let m = end;
            for (let i = 0; i < 11; i++) {
                const [y, mo] = m.split('-').map(Number);
                m = (mo === 1) ? (y - 1) + '-12' : y + '-' + String(mo - 1).padStart(2, '0');
            }
            start = m;
        } else {
            // No timeline set — default to last 12 months
            end = currentMonth;
            let m = currentMonth;
            for (let i = 0; i < 11; i++) {
                const [y, mo] = m.split('-').map(Number);
                m = (mo === 1) ? (y - 1) + '-12' : y + '-' + String(mo - 1).padStart(2, '0');
            }
            start = m;
        }

        // Generate full month range
        const months = [];
        let m = start;
        while (m <= end) {
            months.push(m);
            const [y, mo] = m.split('-').map(Number);
            m = (mo === 12) ? (y + 1) + '-01' : y + '-' + String(mo + 1).padStart(2, '0');
        }
        return months;
    },

    _renderDashCashFlowChart(snapshotMonth) {
        const canvas = document.getElementById('dashChartCashFlow');
        if (!canvas) return;
        if (this._dashCharts.cashflow) this._dashCharts.cashflow.destroy();

        let months = this._getTimelineMonths();
        if (snapshotMonth) months = months.filter(m => m <= snapshotMonth);
        const netData = [];
        const colors = [];

        for (const month of months) {
            const revResult = Database.db.exec(
                "SELECT COALESCE(SUM(amount),0) FROM transactions WHERE transaction_type='receivable' AND status='received' AND month_paid=?", [month]);
            const expResult = Database.db.exec(
                "SELECT COALESCE(SUM(amount),0) FROM transactions WHERE transaction_type='payable' AND status='paid' AND month_paid=?", [month]);
            const rev = revResult[0] ? revResult[0].values[0][0] : 0;
            const exp = expResult[0] ? expResult[0].values[0][0] : 0;
            const net = rev - exp;
            netData.push(net);
            colors.push(net >= 0 ? 'rgba(16,185,129,0.7)' : 'rgba(239,68,68,0.7)');
        }

        this._dashCharts.cashflow = new Chart(canvas, {
            type: 'bar',
            data: {
                labels: months.map(m => { const [y, mo] = m.split('-'); return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][parseInt(mo)-1] + ' ' + y.slice(2); }),
                datasets: [{ label: 'Net Cash Flow', data: netData, backgroundColor: colors, borderRadius: 4 }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: { y: { ticks: { callback: v => Utils.formatCurrency(v) } } }
            }
        });
    },

    _renderDashPnLChart(snapshotMonth) {
        const canvas = document.getElementById('dashChartPnL');
        if (!canvas) return;
        if (this._dashCharts.pnl) this._dashCharts.pnl.destroy();

        let months = this._getTimelineMonths();
        if (snapshotMonth) months = months.filter(m => m <= snapshotMonth);
        const plData = Database.getPLSpreadsheet();
        const overrides = Database.getAllPLOverrides();
        const currentMonth = snapshotMonth || Utils.getCurrentMonth();
        const isFuture = (m) => m > currentMonth;

        const getVal = (catId, month, computed) => {
            const key = `${catId}-${month}`;
            return (key in overrides) ? overrides[key] : computed;
        };

        // Group rows by category
        const groupByCategory = (rows) => {
            const map = {};
            (rows || []).forEach(row => {
                if (!map[row.category_id]) map[row.category_id] = {};
                map[row.category_id][row.month] = (map[row.category_id][row.month] || 0) + row.total;
            });
            return map;
        };

        const computeProjectedAvg = (catMonths) => {
            const pastValues = months.filter(m => !isFuture(m)).map(m => catMonths[m] || 0).filter(v => v > 0);
            return pastValues.length > 0 ? pastValues.reduce((a, b) => a + b, 0) / pastValues.length : 0;
        };

        // Revenue (matches P&L table Total Revenue)
        const revByCat = groupByCategory(plData.revenue);
        const revenueByMonth = {};
        months.forEach(m => { revenueByMonth[m] = 0; });
        Object.entries(revByCat).forEach(([catId, catMonths]) => {
            const projAvg = computeProjectedAvg(catMonths);
            months.forEach(m => {
                const computed = catMonths[m] || 0;
                const fallback = isFuture(m) && computed === 0 ? projAvg : computed;
                revenueByMonth[m] += getVal(catId, m, fallback);
            });
        });

        // Expenses (COGS + Operating Expenses)
        const expensesByMonth = {};
        months.forEach(m => { expensesByMonth[m] = 0; });
        // COGS
        const cogsByCat = groupByCategory(plData.cogs);
        Object.entries(cogsByCat).forEach(([catId, catMonths]) => {
            const projAvg = computeProjectedAvg(catMonths);
            months.forEach(m => {
                const computed = catMonths[m] || 0;
                const fallback = isFuture(m) && computed === 0 ? projAvg : computed;
                expensesByMonth[m] += getVal(catId, m, fallback);
            });
        });
        // OpEx
        const opexByCat = groupByCategory(plData.opex);
        Object.entries(opexByCat).forEach(([catId, catMonths]) => {
            const projAvg = computeProjectedAvg(catMonths);
            months.forEach(m => {
                const computed = catMonths[m] || 0;
                const fallback = isFuture(m) && computed === 0 ? projAvg : computed;
                expensesByMonth[m] += getVal(catId, m, fallback);
            });
        });
        // Depreciation categories (from overrides only)
        (plData.depreciation || []).forEach(cat => {
            months.forEach(m => {
                expensesByMonth[m] += getVal(cat.category_id, m, 0);
            });
        });
        // Asset depreciation
        const assetDepr = plData.assetDeprByMonth || {};
        months.forEach(m => { expensesByMonth[m] += (assetDepr[m] || 0); });
        // Loan interest
        const loanInt = plData.loanInterestByMonth || {};
        months.forEach(m => { expensesByMonth[m] += (loanInt[m] || 0); });

        const revenue = months.map(m => revenueByMonth[m] || 0);
        const expenses = months.map(m => expensesByMonth[m] || 0);

        const labels = months.map(m => { const [y, mo] = m.split('-'); return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][parseInt(mo)-1] + ' ' + y.slice(2); });

        this._dashCharts.pnl = new Chart(canvas, {
            type: 'line',
            data: {
                labels,
                datasets: [
                    { label: 'Revenue', data: revenue, borderColor: 'rgba(16,185,129,0.9)', backgroundColor: 'rgba(16,185,129,0.1)', fill: true, tension: 0.3 },
                    { label: 'Expenses', data: expenses, borderColor: 'rgba(239,68,68,0.9)', backgroundColor: 'rgba(239,68,68,0.1)', fill: true, tension: 0.3 }
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { position: 'top' } },
                scales: { y: { ticks: { callback: v => Utils.formatCurrency(v) } } }
            }
        });
    },


    _renderDashBSChart(snapshotMonth) {
        const canvas = document.getElementById('dashChartBS');
        if (!canvas) return;
        if (this._dashCharts.bs) this._dashCharts.bs.destroy();

        let assets, liabilities, receivables;
        if (snapshotMonth) {
            assets = Math.max(0, Database.getCashAsOf(snapshotMonth));
            receivables = Database.getAccountsReceivableAsOf ? Database.getAccountsReceivableAsOf(snapshotMonth) : 0;
            liabilities = Database.getAccountsPayableAsOf ? Database.getAccountsPayableAsOf(snapshotMonth) : 0;
        } else {
            const summary = Database.calculateSummary();
            assets = summary.cashBalance > 0 ? summary.cashBalance : 0;
            liabilities = summary.accountsPayable;
            receivables = summary.accountsReceivable;
        }

        this._dashCharts.bs = new Chart(canvas, {
            type: 'bar',
            data: {
                labels: ['Assets', 'Liabilities'],
                datasets: [
                    { label: 'Cash', data: [assets, 0], backgroundColor: 'rgba(16,185,129,0.7)', borderRadius: 4 },
                    { label: 'Receivables', data: [receivables, 0], backgroundColor: 'rgba(59,130,246,0.7)', borderRadius: 4 },
                    { label: 'Payables', data: [0, liabilities], backgroundColor: 'rgba(239,68,68,0.7)', borderRadius: 4 }
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { position: 'top' } },
                scales: { x: { stacked: true }, y: { stacked: true, ticks: { callback: v => Utils.formatCurrency(v) } } }
            }
        });
    },

    // ==================== ANALYZE MODE CHART PANELS ====================

    _analyzeCharts: {},

    _updateAnalyzeCharts(tab) {
        const isAnalyze = this.currentMode === 'analyze';
        const panels = ['analyzeCashflowChart', 'analyzePnlChart', 'analyzeBSChart', 'analyzeBEChart'];
        panels.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.style.display = 'none';
        });

        if (!isAnalyze) return;

        if (tab === 'cashflow') {
            document.getElementById('analyzeCashflowChart').style.display = '';
            this._renderAnalyzeCFChart();
        } else if (tab === 'pnl') {
            document.getElementById('analyzePnlChart').style.display = '';
            this._renderAnalyzePnLChart();
        } else if (tab === 'balancesheet') {
            document.getElementById('analyzeBSChart').style.display = '';
            this._renderAnalyzeBSChart();
        } else if (tab === 'breakeven') {
            document.getElementById('analyzeBEChart').style.display = '';
            this._renderAnalyzeBEProgress();
        }
    },

    _renderAnalyzeCFChart() {
        const canvas = document.getElementById('analyzeChartCF');
        if (!canvas) return;
        if (this._analyzeCharts.cf) this._analyzeCharts.cf.destroy();

        const months = this._getTimelineMonths();
        const inData = [];
        const outData = [];
        let running = 0;
        const runningData = [];

        for (const month of months) {
            const revResult = Database.db.exec(
                "SELECT COALESCE(SUM(amount),0) FROM transactions WHERE transaction_type='receivable' AND status='received' AND month_paid=?", [month]);
            const expResult = Database.db.exec(
                "SELECT COALESCE(SUM(amount),0) FROM transactions WHERE transaction_type='payable' AND status='paid' AND month_paid=?", [month]);
            const rev = revResult[0] ? revResult[0].values[0][0] : 0;
            const exp = expResult[0] ? expResult[0].values[0][0] : 0;
            inData.push(rev);
            outData.push(-exp);
            running += rev - exp;
            runningData.push(running);
        }

        const labels = months.map(m => { const [y, mo] = m.split('-'); return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][parseInt(mo)-1] + " '" + y.slice(2); });

        this._analyzeCharts.cf = new Chart(canvas, {
            type: 'bar',
            data: {
                labels,
                datasets: [
                    { label: 'Money In', data: inData, backgroundColor: 'rgba(16,185,129,0.7)', borderRadius: 4, stack: 'flow' },
                    { label: 'Money Out', data: outData, backgroundColor: 'rgba(239,68,68,0.7)', borderRadius: 4, stack: 'flow' },
                    { label: 'Running Balance', data: runningData, type: 'line', borderColor: 'rgba(59,130,246,0.9)', backgroundColor: 'transparent', tension: 0.3, pointRadius: 3, yAxisID: 'y1' }
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { position: 'top' } },
                scales: {
                    y: { stacked: true, ticks: { callback: v => Utils.formatCurrency(v) } },
                    y1: { position: 'right', grid: { display: false }, ticks: { callback: v => Utils.formatCurrency(v) } }
                }
            }
        });
    },

    _renderAnalyzePnLChart() {
        const canvas = document.getElementById('analyzeChartPnL');
        if (!canvas) return;
        if (this._analyzeCharts.pnl) this._analyzeCharts.pnl.destroy();

        const months = this._getTimelineMonths();
        const plData = Database.getPLSpreadsheet();
        const overrides = Database.getAllPLOverrides();
        const currentMonth = Utils.getCurrentMonth();
        const isFuture = (m) => m > currentMonth;

        const getVal = (catId, month, computed) => {
            const key = `${catId}-${month}`;
            return (key in overrides) ? overrides[key] : computed;
        };

        const groupByCategory = (rows) => {
            const map = {};
            (rows || []).forEach(row => {
                if (!map[row.category_id]) map[row.category_id] = {};
                map[row.category_id][row.month] = (map[row.category_id][row.month] || 0) + row.total;
            });
            return map;
        };

        const computeProjectedAvg = (catMonths) => {
            const pastValues = months.filter(m => !isFuture(m)).map(m => catMonths[m] || 0).filter(v => v > 0);
            return pastValues.length > 0 ? pastValues.reduce((a, b) => a + b, 0) / pastValues.length : 0;
        };

        // Revenue (matches P&L table Total Revenue)
        const revByCat = groupByCategory(plData.revenue);
        const revenueByMonth = {};
        months.forEach(m => { revenueByMonth[m] = 0; });
        Object.entries(revByCat).forEach(([catId, catMonths]) => {
            const projAvg = computeProjectedAvg(catMonths);
            months.forEach(m => {
                const computed = catMonths[m] || 0;
                const fallback = isFuture(m) && computed === 0 ? projAvg : computed;
                revenueByMonth[m] += getVal(catId, m, fallback);
            });
        });

        // Expenses (COGS + Operating Expenses)
        const expensesByMonth = {};
        months.forEach(m => { expensesByMonth[m] = 0; });
        // COGS
        const cogsByCat = groupByCategory(plData.cogs);
        Object.entries(cogsByCat).forEach(([catId, catMonths]) => {
            const projAvg = computeProjectedAvg(catMonths);
            months.forEach(m => {
                const computed = catMonths[m] || 0;
                const fallback = isFuture(m) && computed === 0 ? projAvg : computed;
                expensesByMonth[m] += getVal(catId, m, fallback);
            });
        });
        // OpEx
        const opexByCat = groupByCategory(plData.opex);
        Object.entries(opexByCat).forEach(([catId, catMonths]) => {
            const projAvg = computeProjectedAvg(catMonths);
            months.forEach(m => {
                const computed = catMonths[m] || 0;
                const fallback = isFuture(m) && computed === 0 ? projAvg : computed;
                expensesByMonth[m] += getVal(catId, m, fallback);
            });
        });
        (plData.depreciation || []).forEach(cat => {
            months.forEach(m => { expensesByMonth[m] += getVal(cat.category_id, m, 0); });
        });
        const assetDepr = plData.assetDeprByMonth || {};
        months.forEach(m => { expensesByMonth[m] += (assetDepr[m] || 0); });
        const loanInt = plData.loanInterestByMonth || {};
        months.forEach(m => { expensesByMonth[m] += (loanInt[m] || 0); });

        const revenue = months.map(m => revenueByMonth[m] || 0);
        const expenses = months.map(m => expensesByMonth[m] || 0);
        const profit = months.map((m, i) => revenue[i] - expenses[i]);

        const labels = months.map(m => { const [y, mo] = m.split('-'); return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][parseInt(mo)-1] + " '" + y.slice(2); });

        this._analyzeCharts.pnl = new Chart(canvas, {
            type: 'line',
            data: {
                labels,
                datasets: [
                    { label: 'Revenue', data: revenue, borderColor: 'rgba(16,185,129,0.9)', backgroundColor: 'rgba(16,185,129,0.1)', fill: true, tension: 0.3 },
                    { label: 'Expenses', data: expenses, borderColor: 'rgba(239,68,68,0.9)', backgroundColor: 'rgba(239,68,68,0.1)', fill: true, tension: 0.3 },
                    { label: 'Profit/Loss', data: profit, borderColor: 'rgba(59,130,246,0.9)', borderDash: [5, 5], fill: false, tension: 0.3, pointRadius: 3 }
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { position: 'top' } },
                scales: { y: { ticks: { callback: v => Utils.formatCurrency(v) } } }
            }
        });
    },

    _renderAnalyzeBSChart() {
        const canvas = document.getElementById('analyzeChartBS');
        if (!canvas) return;
        if (this._analyzeCharts.bs) this._analyzeCharts.bs.destroy();

        // Use the balance sheet month selector if set, otherwise last timeline month
        const bsMonth = document.getElementById('bsMonthMonth').value;
        const bsYear = document.getElementById('bsMonthYear').value;
        let asOfMonth;
        if (bsMonth && bsYear) {
            asOfMonth = `${bsYear}-${bsMonth}`;
        } else {
            const tlMonths = this._getTimelineMonths();
            asOfMonth = tlMonths.length > 0 ? tlMonths[tlMonths.length - 1] : Utils.getCurrentMonth();
        }

        const cash = Database.getCashAsOf(asOfMonth);
        const receivables = Database.getAccountsReceivableAsOf(asOfMonth);
        const payables = Database.getAccountsPayableAsOf(asOfMonth);
        const salesTaxPayable = Database.getSalesTaxPayableAsOf(asOfMonth);

        // Compute proper loan balances (respecting start dates)
        const round2 = (v) => Math.round(v * 100) / 100;
        const loans = Database.getLoans();
        let totalLoanBalance = 0;
        loans.forEach(loan => {
            const loanStartMonth = loan.start_date ? loan.start_date.substring(0, 7) : '';
            if (loanStartMonth && loanStartMonth > asOfMonth) return;
            const skipped = Database.getSkippedPayments(loan.id);
            const overrides = Database.getLoanPaymentOverrides(loan.id);
            const schedule = Utils.computeAmortizationSchedule({
                principal: loan.principal, annual_rate: loan.annual_rate,
                term_months: loan.term_months, payments_per_year: loan.payments_per_year,
                start_date: loan.start_date, first_payment_date: loan.first_payment_date
            }, skipped, overrides);
            let balance = loan.principal;
            for (let i = schedule.length - 1; i >= 0; i--) {
                if (schedule[i].month <= asOfMonth) { balance = schedule[i].ending_balance; break; }
            }
            if (schedule.length > 0 && schedule[0].month > asOfMonth) balance = loan.principal;
            totalLoanBalance = round2(totalLoanBalance + balance);
        });

        // Fixed assets
        const fixedAssets = Database.getFixedAssets();
        let netFixedAssets = 0;
        fixedAssets.forEach(asset => {
            const deprSchedule = Utils.computeDepreciationSchedule(asset);
            let accumDepr = 0;
            Object.entries(deprSchedule).forEach(([m, amt]) => { if (m <= asOfMonth) accumDepr += amt; });
            netFixedAssets = round2(netFixedAssets + asset.purchase_cost - accumDepr);
        });

        const totalAssets = round2(cash + receivables + netFixedAssets);
        const totalLiabilities = round2(payables + salesTaxPayable + totalLoanBalance);

        // Proper equity from equity config + retained earnings
        const taxMode = Database.getPLTaxMode();
        const equityConfig = Database.getEquityConfig();
        const seedEffective = (equityConfig.seed_received_date || equityConfig.seed_expected_date || '');
        const apicEffective = (equityConfig.apic_expected_date || equityConfig.apic_received_date || '');
        const seedMonth = seedEffective ? seedEffective.substring(0, 7) : '';
        const apicMonth = apicEffective ? apicEffective.substring(0, 7) : '';
        const commonStock = (seedMonth && seedMonth > asOfMonth) ? 0 : round2(equityConfig.common_stock_par * equityConfig.common_stock_shares);
        const apicVal = (apicMonth && apicMonth > asOfMonth) ? 0 : round2(equityConfig.apic || 0);
        const retainedEarnings = round2(Database.getRetainedEarningsAsOf(asOfMonth, taxMode));
        const equity = round2(commonStock + apicVal + retainedEarnings);

        this._analyzeCharts.bs = new Chart(canvas, {
            type: 'bar',
            data: {
                labels: ['Assets', 'Liabilities & Equity'],
                datasets: [
                    { label: 'Cash', data: [cash, 0], backgroundColor: 'rgba(16,185,129,0.7)', borderRadius: 4 },
                    { label: 'Receivables', data: [receivables, 0], backgroundColor: 'rgba(59,130,246,0.7)', borderRadius: 4 },
                    { label: 'Fixed Assets', data: [netFixedAssets, 0], backgroundColor: 'rgba(251,191,36,0.7)', borderRadius: 4 },
                    { label: 'Liabilities', data: [0, totalLiabilities], backgroundColor: 'rgba(239,68,68,0.7)', borderRadius: 4 },
                    { label: 'Equity', data: [0, Math.max(0, equity)], backgroundColor: 'rgba(168,85,247,0.7)', borderRadius: 4 }
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                indexAxis: 'y',
                plugins: { legend: { position: 'top' } },
                scales: { x: { stacked: true, ticks: { callback: v => Utils.formatCurrency(v) } }, y: { stacked: true } }
            }
        });

        // Render gauges
        const gaugesEl = document.getElementById('analyzeGauges');
        if (gaugesEl) {
            const currentAssets = round2(cash + receivables);
            const currentLiabilities = round2(payables + salesTaxPayable);
            const currentRatio = currentLiabilities > 0 ? (currentAssets / currentLiabilities).toFixed(2) : 'N/A';
            const debtToEquity = equity > 0 ? (totalLiabilities / equity).toFixed(2) : 'N/A';

            const makeGauge = (value, max, label, lowerIsBetter) => {
                const numVal = parseFloat(value);
                const pct = isNaN(numVal) ? 0 : Math.min(100, Math.max(0, (numVal / max) * 100));
                let color;
                if (lowerIsBetter) {
                    // For D/E: low = green, high = red
                    color = pct < 40 ? 'var(--color-success, #10b981)' : pct < 70 ? 'var(--color-warning, #f59e0b)' : 'var(--color-danger, #ef4444)';
                } else {
                    // For Current Ratio: high = green, low = red
                    color = pct > 60 ? 'var(--color-success, #10b981)' : pct > 30 ? 'var(--color-warning, #f59e0b)' : 'var(--color-danger, #ef4444)';
                }
                return '<div class="analyze-gauge">' +
                    '<div class="analyze-gauge-ring" style="background: conic-gradient(' + color + ' ' + (pct * 3.6) + 'deg, var(--border, #e5e7eb) ' + (pct * 3.6) + 'deg);">' +
                    '<div class="analyze-gauge-value" style="background:var(--surface,#fff);width:60px;height:60px;border-radius:50%;display:flex;align-items:center;justify-content:center;">' + value + '</div>' +
                    '</div>' +
                    '<div class="analyze-gauge-label">' + label + '</div>' +
                    '</div>';
            };

            gaugesEl.innerHTML = makeGauge(currentRatio, 3, 'Current Ratio', false) + makeGauge(debtToEquity, 3, 'Debt-to-Equity', true);
        }
    },

    _renderAnalyzeBEProgress() {
        const container = document.getElementById('analyzeBEProgress');
        if (!container) return;

        // Use the real current month (not timeline end which may be in the future)
        const currentMonth = Utils.getCurrentMonth();
        // Use actual P&L operating expenses (varies by month) instead of flat budget amounts
        const opexByMonth = Database.getMonthlyTotalOpex([currentMonth]);
        const totalTarget = opexByMonth[currentMonth] || 0;

        // Gross profit directly from P&L (Revenue - COGS, with overrides)
        const gpByMonth = Database.getMonthlyGrossProfit([currentMonth]);
        const currentGP = gpByMonth[currentMonth] || 0;

        const pct = totalTarget > 0 ? Math.min(100, Math.round(currentGP / totalTarget * 100)) : 0;
        const color = pct >= 100 ? 'var(--color-success, #10b981)' : pct >= 60 ? 'var(--color-warning, #f59e0b)' : 'var(--color-danger, #ef4444)';

        const revByMonth = Database.getMonthlyTotalRevenue([currentMonth]);
        const cogsByMonth = Database.getMonthlyTotalCogs([currentMonth]);

        container.innerHTML =
            '<div style="text-align:center;margin-bottom:16px;">' +
                '<div style="font-size:2.5rem;font-weight:700;color:' + color + ';">' + pct + '%</div>' +
                '<div style="font-size:0.85rem;color:var(--text-muted);">Break-Even Progress This Month</div>' +
            '</div>' +
            '<div class="dash-be-info">' +
                '<span>Gross Profit: ' + Utils.formatCurrency(currentGP) + '</span>' +
                '<span>Target: ' + Utils.formatCurrency(totalTarget) + '</span>' +
            '</div>' +
            '<div style="font-size:0.75rem;color:var(--text-muted);display:flex;justify-content:space-between;margin-top:2px;">' +
                '<span>Revenue: ' + Utils.formatCurrency(revByMonth[currentMonth] || 0) + '</span>' +
                '<span>COGS: ' + Utils.formatCurrency(cogsByMonth[currentMonth] || 0) + '</span>' +
            '</div>' +
            '<div class="dash-be-track" style="height:32px;">' +
                '<div class="dash-be-fill" style="width: ' + pct + '%; background: ' + color + ';"></div>' +
            '</div>';
    },

    _renderDashRevConcentrationChart(snapshotMonth) {
        const canvas = document.getElementById('dashChartRevConc');
        if (!canvas) return;
        if (this._dashCharts.revconc) this._dashCharts.revconc.destroy();

        // Get revenue by category (exclude equity, loans, sales tax, non-revenue)
        const monthFilter = snapshotMonth ? " AND t.month_due <= ?" : "";
        const monthParams = snapshotMonth ? [snapshotMonth] : [];
        const result = Database.db.exec(
            "SELECT c.name, SUM(COALESCE(t.pretax_amount, t.amount)) as total " +
            "FROM transactions t JOIN categories c ON t.category_id = c.id " +
            "WHERE t.transaction_type = 'receivable' AND c.is_cogs = 0 " +
            "AND c.show_on_pl != 1 AND c.is_sales_tax = 0 " +
            "AND (c.is_b2b = 1 OR c.is_sales = 1) " +
            "AND COALESCE(t.source_type, '') NOT IN ('loan_receivable', 'loan_payment')" +
            monthFilter +
            " GROUP BY c.name ORDER BY total DESC", monthParams);

        if (!result[0] || result[0].values.length === 0) {
            canvas.parentElement.innerHTML = '<div style="padding:24px;color:var(--text-muted);text-align:center;">No revenue data yet</div>';
            return;
        }

        const rows = result[0].values;
        const totalRevenue = rows.reduce((sum, r) => sum + r[1], 0);
        if (totalRevenue <= 0) {
            canvas.parentElement.innerHTML = '<div style="padding:24px;color:var(--text-muted);text-align:center;">No revenue data yet</div>';
            return;
        }

        // Show top 5, bundle rest as "Other"
        const top = rows.slice(0, 5);
        const otherTotal = rows.slice(5).reduce((sum, r) => sum + r[1], 0);
        const labels = top.map(r => r[0]);
        const data = top.map(r => r[1]);
        if (otherTotal > 0) {
            labels.push('Other');
            data.push(otherTotal);
        }
        const pcts = data.map(v => ((v / totalRevenue) * 100).toFixed(1));

        // Concentration risk indicator — HHI (Herfindahl-Hirschman Index)
        const shares = rows.map(r => (r[1] / totalRevenue) * 100);
        const hhi = Math.round(shares.reduce((sum, s) => sum + s * s, 0));
        // HHI > 2500 = highly concentrated, 1500-2500 = moderate, < 1500 = diversified
        const concLabel = hhi > 2500 ? 'High Concentration' : hhi > 1500 ? 'Moderate' : 'Diversified';
        const concColor = hhi > 2500 ? 'rgba(239,68,68,0.9)' : hhi > 1500 ? 'rgba(251,146,60,0.9)' : 'rgba(16,185,129,0.9)';

        const barColors = [
            'rgba(59,130,246,0.8)', 'rgba(16,185,129,0.7)', 'rgba(251,146,60,0.7)',
            'rgba(168,85,247,0.7)', 'rgba(236,72,153,0.7)', 'rgba(148,163,184,0.5)'
        ];

        this._dashCharts.revconc = new Chart(canvas, {
            type: 'bar',
            data: {
                labels: labels.map((l, i) => l + ' (' + pcts[i] + '%)'),
                datasets: [{
                    label: 'Revenue',
                    data: data,
                    backgroundColor: barColors.slice(0, data.length),
                    borderRadius: 4
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true, maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    title: {
                        display: true,
                        text: 'Risk: ' + concLabel + ' (HHI: ' + hhi + ')',
                        color: concColor,
                        font: { size: 12, weight: '600' },
                        padding: { bottom: 12 }
                    }
                },
                scales: {
                    x: { ticks: { callback: v => Utils.formatCurrency(v) } },
                    y: { ticks: { font: { size: 11 } } }
                }
            }
        });
    },

    _renderDashARAging() {
        const container = document.getElementById('dashARAging');
        if (!container) return;

        const currentMonth = Utils.getCurrentMonth();
        const [curY, curM] = currentMonth.split('-').map(Number);

        // Get all pending receivables
        const result = Database.db.exec(
            "SELECT t.amount, t.month_due FROM transactions t " +
            "WHERE t.transaction_type = 'receivable' AND t.status = 'pending'");

        const buckets = { current: 0, days30: 0, days60: 0, days90: 0, days90plus: 0 };
        const bucketCounts = { current: 0, days30: 0, days60: 0, days90: 0, days90plus: 0 };

        if (result[0]) {
            for (const row of result[0].values) {
                const amount = row[0];
                const monthDue = row[1];
                if (!monthDue) continue;
                const [dueY, dueM] = monthDue.split('-').map(Number);
                const monthsOverdue = (curY - dueY) * 12 + (curM - dueM);

                if (monthsOverdue <= 0) { buckets.current += amount; bucketCounts.current++; }
                else if (monthsOverdue === 1) { buckets.days30 += amount; bucketCounts.days30++; }
                else if (monthsOverdue === 2) { buckets.days60 += amount; bucketCounts.days60++; }
                else if (monthsOverdue === 3) { buckets.days90 += amount; bucketCounts.days90++; }
                else { buckets.days90plus += amount; bucketCounts.days90plus++; }
            }
        }

        const total = Object.values(buckets).reduce((s, v) => s + v, 0);
        const totalCount = Object.values(bucketCounts).reduce((s, v) => s + v, 0);

        if (total <= 0) {
            container.innerHTML = '<div style="padding:24px;color:var(--text-muted);text-align:center;">No pending receivables</div>';
            return;
        }

        const bucketData = [
            { label: 'Current', amount: buckets.current, count: bucketCounts.current, color: 'var(--color-success, #10b981)' },
            { label: '1 month', amount: buckets.days30, count: bucketCounts.days30, color: '#60a5fa' },
            { label: '2 months', amount: buckets.days60, count: bucketCounts.days60, color: 'var(--color-warning, #f59e0b)' },
            { label: '3 months', amount: buckets.days90, count: bucketCounts.days90, color: '#f97316' },
            { label: '3+ months', amount: buckets.days90plus, count: bucketCounts.days90plus, color: 'var(--color-danger, #ef4444)' }
        ];

        let html = '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;">';
        bucketData.forEach(b => {
            const pct = total > 0 ? ((b.amount / total) * 100).toFixed(0) : 0;
            html += '<div style="flex:1;min-width:100px;padding:10px 12px;border-radius:8px;background:var(--c5,var(--border));text-align:center;">' +
                '<div style="font-size:0.7rem;color:var(--text-muted);margin-bottom:2px;">' + b.label + '</div>' +
                '<div style="font-size:1.05rem;font-weight:700;font-family:DM Mono,monospace;">' + Utils.formatCurrency(b.amount) + '</div>' +
                '<div style="font-size:0.7rem;color:var(--text-muted);">' + b.count + ' item' + (b.count !== 1 ? 's' : '') + ' &middot; ' + pct + '%</div>' +
            '</div>';
        });
        html += '</div>';

        // Stacked progress bar
        html += '<div style="display:flex;height:20px;border-radius:6px;overflow:hidden;background:var(--c5,var(--border));">';
        bucketData.forEach(b => {
            const pct = total > 0 ? (b.amount / total) * 100 : 0;
            if (pct > 0) {
                html += '<div style="width:' + pct + '%;background:' + b.color + ';transition:width 0.3s;" title="' + b.label + ': ' + Utils.formatCurrency(b.amount) + '"></div>';
            }
        });
        html += '</div>';
        html += '<div style="display:flex;justify-content:space-between;margin-top:6px;font-size:0.7rem;color:var(--text-muted);">' +
            '<span>Total: ' + Utils.formatCurrency(total) + ' (' + totalCount + ' items)</span>' +
            '<span style="color:' + (buckets.days90plus > 0 ? 'var(--color-danger)' : 'var(--color-success)') + ';font-weight:600;">' + (buckets.days90plus > 0 ? 'Collection risk' : 'Healthy') + '</span>' +
        '</div>';

        container.innerHTML = html;
    },

    _renderDashBreakevenBar(snapshotMonth) {
        const container = document.getElementById('dashBreakevenBar');
        if (!container) return;

        const targetMonth = snapshotMonth || Utils.getCurrentMonth();
        const opexByMonth = Database.getMonthlyTotalOpex([targetMonth]);
        const totalMonthlyExpenses = opexByMonth[targetMonth] || 0;

        const gpByMonth = Database.getMonthlyGrossProfit([targetMonth]);
        const currentGP = gpByMonth[targetMonth] || 0;

        const pct = totalMonthlyExpenses > 0 ? Math.min(100, Math.round(currentGP / totalMonthlyExpenses * 100)) : 0;
        const color = pct >= 100 ? 'var(--color-success, #10b981)' : pct >= 60 ? 'var(--color-warning, #f59e0b)' : 'var(--color-danger, #ef4444)';
        const monthLabel = snapshotMonth ? Utils.formatMonthShort(targetMonth) : 'this month';

        container.innerHTML =
            '<div class="dash-be-info">' +
                '<span>Gross Profit: ' + Utils.formatCurrency(currentGP) + '</span>' +
                '<span>Target: ' + Utils.formatCurrency(totalMonthlyExpenses) + '</span>' +
            '</div>' +
            '<div class="dash-be-track">' +
                '<div class="dash-be-fill" style="width: ' + pct + '%; background: ' + color + ';"></div>' +
            '</div>' +
            '<div class="dash-be-pct">' + pct + '% to break-even ' + monthLabel + '</div>';
    },

    // ==================== QUICK ENTRY (Multi-Line) ====================

    _qeRowCount: 0,

    openQuickEntry(type) {
        if (type === 'budget') {
            this._qeRowCount = 0;
            document.getElementById('quickBudgetRows').innerHTML = '';
            this._addQuickBudgetRow();
            this._addQuickBudgetRow();
            this._addQuickBudgetRow();
            this._updateQuickBudgetSaveCount();
            UI.showModal('quickBudgetModal');
        } else {
            this._qeRowCount = 0;
            document.getElementById('quickEntryRows').innerHTML = '';
            this._addQuickEntryRow();
            this._addQuickEntryRow();
            this._addQuickEntryRow();
            this._updateQuickEntrySaveCount();
            UI.showModal('quickEntryModal');
        }
    },

    _addQuickEntryRow() {
        const tbody = document.getElementById('quickEntryRows');
        const idx = this._qeRowCount++;
        const today = Utils.getTodayDate();
        const categories = Database.getCategories();
        const catOptions = categories.map(c => '<option value="' + c.id + '">' + this._escapeHtml(c.name) + '</option>').join('');
        const currentMonth = Utils.getCurrentMonth();

        const tr = document.createElement('tr');
        tr.dataset.qeIdx = idx;
        tr.innerHTML =
            '<td><input type="date" class="qe-date" value="' + today + '"></td>' +
            '<td><input type="text" class="qe-desc" placeholder="Description"></td>' +
            '<td><select class="qe-cat"><option value="">Select...</option>' + catOptions + '</select></td>' +
            '<td><input type="number" class="qe-amount" step="0.01" min="0" placeholder="0.00"></td>' +
            '<td><select class="qe-type"><option value="receivable">Receivable</option><option value="payable">Payable</option></select></td>' +
            '<td><select class="qe-status"><option value="pending">Pending</option><option value="received">Received</option><option value="paid">Paid</option></select></td>' +
            '<td><input type="month" class="qe-monthdue" value="' + currentMonth + '"></td>' +
            '<td><input type="text" class="qe-notes" placeholder="Notes"></td>' +
            '<td><button class="qe-delete-btn" title="Remove row">&times;</button></td>';

        // Wire delete
        tr.querySelector('.qe-delete-btn').addEventListener('click', () => {
            tr.remove();
            const expandRow = tbody.querySelector('tr.qe-expand-row[data-qe-parent="' + idx + '"]');
            if (expandRow) expandRow.remove();
            this._updateQuickEntrySaveCount();
        });

        // Wire category change for smart expand
        tr.querySelector('.qe-cat').addEventListener('change', (e) => {
            const catId = parseInt(e.target.value);
            const cat = categories.find(c => c.id === catId);
            const existingExpand = tbody.querySelector('tr.qe-expand-row[data-qe-parent="' + idx + '"]');

            if (cat && cat.is_sales) {
                if (!existingExpand) {
                    const expandTr = document.createElement('tr');
                    expandTr.className = 'qe-expand-row';
                    expandTr.dataset.qeParent = idx;
                    expandTr.innerHTML = '<td colspan="9"><div class="qe-expand-fields">' +
                        '<div class="form-group"><label>Pretax Amount</label><input type="number" class="qe-pretax" step="0.01" min="0" placeholder="0.00"></div>' +
                        '<div class="form-group"><label>Inventory Cost</label><input type="number" class="qe-invcost" step="0.01" min="0" placeholder="0.00"></div>' +
                        '<div class="form-group"><label>Sale Start</label><input type="date" class="qe-salestart"></div>' +
                        '<div class="form-group"><label>Sale End</label><input type="date" class="qe-saleend"></div>' +
                        '<div class="form-group"><label>Payment For</label><input type="month" class="qe-paymentfor"></div>' +
                        '</div></td>';
                    tr.after(expandTr);
                }
            } else if (existingExpand) {
                existingExpand.remove();
            }
        });

        // Tab from last cell of last row adds new row
        tr.querySelector('.qe-notes').addEventListener('keydown', (e) => {
            if (e.key === 'Tab' && !e.shiftKey) {
                const allRows = tbody.querySelectorAll('tr:not(.qe-expand-row)');
                if (tr === allRows[allRows.length - 1]) {
                    e.preventDefault();
                    this._addQuickEntryRow();
                    this._updateQuickEntrySaveCount();
                    const newRow = tbody.querySelector('tr:not(.qe-expand-row):last-child');
                    if (newRow) newRow.querySelector('.qe-date').focus();
                }
            }
        });

        tbody.appendChild(tr);
    },

    _addQuickBudgetRow() {
        const tbody = document.getElementById('quickBudgetRows');
        const idx = this._qeRowCount++;
        const groups = Database.getBudgetGroups();
        const groupOptions = groups.map(g => '<option value="' + g.id + '">' + this._escapeHtml(g.name) + '</option>').join('');
        const currentMonth = Utils.getCurrentMonth();

        const tr = document.createElement('tr');
        tr.dataset.qeIdx = idx;
        tr.innerHTML =
            '<td><input type="text" class="qe-name" placeholder="Expense name"></td>' +
            '<td><input type="number" class="qe-amount" step="0.01" min="0" placeholder="0.00"></td>' +
            '<td><select class="qe-group"><option value="">No group</option>' + groupOptions + '</select></td>' +
            '<td><input type="month" class="qe-start" value="' + currentMonth + '"></td>' +
            '<td><input type="month" class="qe-end" placeholder="Ongoing"></td>' +
            '<td><input type="text" class="qe-notes" placeholder="Notes"></td>' +
            '<td><button class="qe-delete-btn" title="Remove row">&times;</button></td>';

        tr.querySelector('.qe-delete-btn').addEventListener('click', () => {
            tr.remove();
            this._updateQuickBudgetSaveCount();
        });

        tr.querySelector('.qe-notes').addEventListener('keydown', (e) => {
            if (e.key === 'Tab' && !e.shiftKey) {
                const allRows = tbody.querySelectorAll('tr');
                if (tr === allRows[allRows.length - 1]) {
                    e.preventDefault();
                    this._addQuickBudgetRow();
                    this._updateQuickBudgetSaveCount();
                    const newRow = tbody.querySelector('tr:last-child');
                    if (newRow) newRow.querySelector('.qe-name').focus();
                }
            }
        });

        tbody.appendChild(tr);
    },

    _updateQuickEntrySaveCount() {
        const rows = document.getElementById('quickEntryRows').querySelectorAll('tr:not(.qe-expand-row)');
        document.getElementById('quickEntrySaveBtn').textContent = 'Save All (' + rows.length + ')';
    },

    _updateQuickBudgetSaveCount() {
        const rows = document.getElementById('quickBudgetRows').querySelectorAll('tr');
        document.getElementById('quickBudgetSaveBtn').textContent = 'Save All (' + rows.length + ')';
    },

    saveQuickEntries() {
        const tbody = document.getElementById('quickEntryRows');
        const rows = tbody.querySelectorAll('tr:not(.qe-expand-row)');
        let errors = 0;
        let saved = 0;

        // Clear previous errors
        tbody.querySelectorAll('.qe-error').forEach(el => el.classList.remove('qe-error'));

        for (const row of rows) {
            const date = row.querySelector('.qe-date').value;
            const catId = row.querySelector('.qe-cat').value;
            const amount = parseFloat(row.querySelector('.qe-amount').value);

            let hasError = false;
            if (!date) { row.querySelector('.qe-date').classList.add('qe-error'); hasError = true; }
            if (!catId) { row.querySelector('.qe-cat').classList.add('qe-error'); hasError = true; }
            if (!amount || isNaN(amount)) { row.querySelector('.qe-amount').classList.add('qe-error'); hasError = true; }

            if (hasError) { errors++; continue; }

            const idx = row.dataset.qeIdx;
            const expandRow = tbody.querySelector('tr.qe-expand-row[data-qe-parent="' + idx + '"]');

            const data = {
                entry_date: date,
                category_id: parseInt(catId),
                item_description: row.querySelector('.qe-desc').value || null,
                amount: amount,
                transaction_type: row.querySelector('.qe-type').value,
                status: row.querySelector('.qe-status').value,
                month_due: row.querySelector('.qe-monthdue').value || null,
                notes: row.querySelector('.qe-notes').value || null,
                pretax_amount: expandRow ? (parseFloat(expandRow.querySelector('.qe-pretax').value) || null) : null,
                inventory_cost: expandRow ? (parseFloat(expandRow.querySelector('.qe-invcost').value) || null) : null,
                sale_date_start: expandRow ? (expandRow.querySelector('.qe-salestart').value || null) : null,
                sale_date_end: expandRow ? (expandRow.querySelector('.qe-saleend').value || null) : null,
                payment_for_month: expandRow ? (expandRow.querySelector('.qe-paymentfor').value || null) : null
            };

            const id = Database.addTransaction(data);
            this._manageSalesTaxEntry(id, data);
            this._manageInventoryCostEntry(id, data);

            saved++;
        }

        const errSpan = document.getElementById('quickEntryErrors');
        if (errors > 0) {
            errSpan.textContent = errors + ' of ' + rows.length + ' rows have errors';
        } else {
            errSpan.textContent = '';
        }

        if (saved > 0) {
            this.pushUndo({
                type: 'quick-entry',
                label: 'Quick add ' + saved + ' transactions',
                undo: () => { /* Complex multi-add undo — not practical for bulk quick entry */ },
                redo: () => {}
            });

            if (errors === 0) {
                UI.hideModal('quickEntryModal');
                UI.showNotification(saved + ' transactions added', 'success');
            } else {
                UI.showNotification(saved + ' saved, ' + errors + ' have errors', 'info');
                // Remove saved rows
                for (const row of Array.from(rows)) {
                    if (!row.querySelector('.qe-error')) {
                        const idx = row.dataset.qeIdx;
                        const expandRow = tbody.querySelector('tr.qe-expand-row[data-qe-parent="' + idx + '"]');
                        if (expandRow) expandRow.remove();
                        row.remove();
                    }
                }
                this._updateQuickEntrySaveCount();
            }
            this.refreshAll();
        }
    },

    saveQuickBudgetEntries() {
        const tbody = document.getElementById('quickBudgetRows');
        const rows = tbody.querySelectorAll('tr');
        let errors = 0;
        let saved = 0;

        tbody.querySelectorAll('.qe-error').forEach(el => el.classList.remove('qe-error'));

        for (const row of rows) {
            const name = row.querySelector('.qe-name').value.trim();
            const amount = parseFloat(row.querySelector('.qe-amount').value);
            const startMonth = row.querySelector('.qe-start').value;

            let hasError = false;
            if (!name) { row.querySelector('.qe-name').classList.add('qe-error'); hasError = true; }
            if (!amount || isNaN(amount)) { row.querySelector('.qe-amount').classList.add('qe-error'); hasError = true; }
            if (!startMonth) { row.querySelector('.qe-start').classList.add('qe-error'); hasError = true; }

            if (hasError) { errors++; continue; }

            const groupId = row.querySelector('.qe-group').value || null;
            const endMonth = row.querySelector('.qe-end').value || null;
            const notes = row.querySelector('.qe-notes').value || null;

            Database.addBudgetExpense(name, amount, startMonth, endMonth, null, notes, groupId ? parseInt(groupId) : null);
            saved++;
        }

        const errSpan = document.getElementById('quickBudgetErrors');
        if (errors > 0) {
            errSpan.textContent = errors + ' of ' + rows.length + ' rows have errors';
        } else {
            errSpan.textContent = '';
        }

        if (saved > 0) {
            if (errors === 0) {
                UI.hideModal('quickBudgetModal');
                UI.showNotification(saved + ' budget items added', 'success');
            } else {
                UI.showNotification(saved + ' saved, ' + errors + ' have errors', 'info');
                for (const row of Array.from(rows)) {
                    if (!row.querySelector('.qe-error')) row.remove();
                }
                this._updateQuickBudgetSaveCount();
            }
            this.refreshBudget();
        }
    },

    _showQuickAddButton(tab) {
        const btn = document.getElementById('quickAddBtn');
        if (!btn) return;
        btn.style.display = (tab === 'journal' || tab === 'budget') ? '' : 'none';
    },

    // ==================== THEME ====================

    /**
     * Load theme settings from DB and apply
     */
    loadAndApplyTheme() {
        const preset = Database.getThemePreset();
        const customColors = Database.getThemeColors();

        // Sync UI controls
        const presetSelect = document.getElementById('themePreset');
        if (presetSelect) presetSelect.value = preset;

        // Show/hide custom picker
        const picker = document.getElementById('customColorPicker');
        if (picker) picker.style.display = preset === 'custom' ? 'grid' : 'none';

        // Sync color inputs if custom
        if (preset === 'custom' && customColors) {
            ['themeC1', 'themeC2', 'themeC3', 'themeC4', 'themeC5', 'themeC6'].forEach((id, i) => {
                const input = document.getElementById(id);
                if (input) input.value = customColors[`c${i + 1}`] || '#000000';
            });
        }

        this.applyTheme(preset, customColors);

        // Load dark mode state
        this.loadDarkMode();
    },

    /**
     * Load shipping fee rate from DB and populate the input
     */
    loadShippingFeeRate() {
        const config = Database.getShippingFeeConfig();
        const input = document.getElementById('shippingFeeRate');
        if (input) {
            input.value = Math.round(config.rate * 1e6) / 1e4;
        }
        const minInput = document.getElementById('shippingFeeMin');
        if (minInput) {
            minInput.value = config.minFee || 0;
        }
    },

    /**
     * Apply theme colors to the document
     * @param {string} preset - Preset name or 'custom'
     * @param {Object|null} customColors - Custom color object {c1, c2, c3, c4}
     * @param {boolean} isDark - Dark mode enabled
     */
    applyTheme(preset, customColors) {
        const colors = preset === 'custom' && customColors
            ? customColors
            : (this.themePresets[preset] || this.themePresets.default);

        const root = document.documentElement;

        // Set or clear design style (drives CSS overrides for font, radius, shadows, etc.)
        root.setAttribute('data-theme-style', colors.style || '');

        // Sidebar (c1) — sidebar background
        root.style.setProperty('--c1', colors.c1);
        root.style.setProperty('--sidebar-bg', colors.c1);

        // Accent (c2) — buttons, active states, focus rings, primary color
        root.style.setProperty('--c2', colors.c2);
        root.style.setProperty('--color-primary', colors.c2);
        root.style.setProperty('--color-primary-hover', Utils.adjustLightness(colors.c2, -12));
        root.style.setProperty('--color-primary-light', Utils.adjustLightness(colors.c2, -6));
        root.style.setProperty('--color-primary-rgb', Utils.hexToRGBString(colors.c2));
        root.style.setProperty('--color-accent', colors.c2);
        root.style.setProperty('--color-accent-bg', Utils.adjustLightness(colors.c2, 35));
        root.style.setProperty('--color-accent-bg-hover', Utils.adjustLightness(colors.c2, 30));
        root.style.setProperty('--color-accent-rgb', Utils.hexToRGBString(colors.c2));

        // Background (c3) — page background
        root.style.setProperty('--c3', colors.c3);
        root.style.setProperty('--color-bg', colors.c3);
        root.style.setProperty('--color-bg-dark', Utils.adjustLightness(colors.c3, -3));

        // Surface (c4) — card/panel backgrounds
        root.style.setProperty('--c4', colors.c4);
        root.style.setProperty('--color-white', colors.c4);
        root.style.setProperty('--color-surface', colors.c4);

        // Border (c5) — borders, dividers, table lines
        root.style.setProperty('--c5', colors.c5 || '#e5e7eb');
        root.style.setProperty('--color-border', colors.c5 || '#e5e7eb');

        // Text (c6) — primary text color
        root.style.setProperty('--c6', colors.c6 || '#1f2937');
        root.style.setProperty('--color-text', colors.c6 || '#1f2937');
        root.style.setProperty('--color-text-secondary', Utils.adjustLightness(colors.c6 || '#1f2937', 25));
        root.style.setProperty('--color-text-muted', Utils.adjustLightness(colors.c6 || '#1f2937', 45));
    },

    /**
     * Load dark mode preference and apply
     */
    loadDarkMode() {
        const saved = localStorage.getItem('darkMode');
        let isDark;
        if (saved !== null) {
            isDark = saved === 'true';
        } else {
            isDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
        }
        this.setDarkMode(isDark);

        // Listen for system preference changes
        window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', (e) => {
            if (localStorage.getItem('darkMode') === null) {
                this.setDarkMode(e.matches);
            }
        });
    },

    /**
     * Set dark mode on/off
     */
    setDarkMode(enabled) {
        const root = document.documentElement;
        root.setAttribute('data-theme', enabled ? 'dark' : 'light');
        const toggle = document.getElementById('darkModeToggle');
        if (toggle) toggle.checked = enabled;

        if (enabled) {
            // Remove inline --c vars for bg/surface/border/text so CSS dark mode overrides take effect
            // Keep --c1 (sidebar) and --c2 (accent) since they're design choices that work in both modes
            ['--c3', '--c4', '--c5', '--c6',
             '--color-bg', '--color-bg-dark', '--color-white', '--color-surface',
             '--color-border', '--color-text', '--color-text-secondary', '--color-text-muted',
             '--color-accent-bg', '--color-accent-bg-hover'
            ].forEach(prop => root.style.removeProperty(prop));
        } else {
            // Re-apply theme to restore inline vars for light mode
            const preset = Database.getThemePreset();
            const customColors = Database.getThemeColors();
            this.applyTheme(preset, customColors);
        }
    },

    /**
     * Get active filter values
     * @returns {Object} Filter object
     */
    getActiveFilters() {
        return {
            type: document.getElementById('filterType').value || null,
            status: document.getElementById('filterStatus').value || null,
            month: document.getElementById('filterMonth').value || null,
            folderId: document.getElementById('filterFolder').value || null,
            categoryId: document.getElementById('filterCategory').value || null
        };
    },

    /**
     * Set up event listeners
     */
    setupEventListeners() {
        // ==================== UNDO/REDO KEYBOARD SHORTCUTS ====================
        document.addEventListener('keydown', (e) => {
            // Don't intercept when typing in inputs/textareas
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.tagName === 'SELECT') return;
            if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
                e.preventDefault();
                this.undo();
            } else if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.key === 'z' && e.shiftKey))) {
                e.preventDefault();
                this.redo();
            }
        });

        // ==================== SIDEBAR TOGGLE ====================
        document.getElementById('sidebarToggleBtn').addEventListener('click', () => {
            const container = document.querySelector('.app-container');
            const isMobile = window.innerWidth <= 1024;
            if (isMobile) {
                // On mobile, toggle sidebar-open (overlay mode)
                container.classList.toggle('sidebar-open');
            } else {
                // On desktop, toggle sidebar-collapsed
                container.classList.toggle('sidebar-collapsed');
                localStorage.setItem('sidebarCollapsed', container.classList.contains('sidebar-collapsed'));
            }
        });

        // Close mobile sidebar when tapping outside (on the backdrop)
        document.querySelector('.app-container').addEventListener('click', (e) => {
            const container = e.currentTarget;
            if (window.innerWidth <= 1024 && container.classList.contains('sidebar-open')) {
                const sidebar = document.querySelector('.app-sidebar');
                if (!sidebar.contains(e.target) && !e.target.closest('.sidebar-toggle-btn')) {
                    container.classList.remove('sidebar-open');
                }
            }
        });

        // ==================== ENTRY FORM ====================

        // Entry form submission
        document.getElementById('entryForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleFormSubmit();
        });

        // New Entry button - open empty form modal
        document.getElementById('newEntryBtn').addEventListener('click', () => {
            UI.resetForm(); // Reset first (this also closes modal)
            const today = Utils.getTodayDate();
            document.getElementById('entryDate').value = today;
            document.getElementById('monthDue').value = today.substring(0, 7);
            this._monthDueManuallySet = false;
            document.getElementById('formTitle').textContent = 'Add New Entry';
            document.getElementById('submitBtn').textContent = 'Add Entry';
            UI.showModal('entryModal');
        });

        // Cancel edit button - close modal
        document.getElementById('cancelEditBtn').addEventListener('click', () => {
            UI.resetForm();
        });

        // Save & New button - save entry and reset form for another
        document.getElementById('saveAndNewBtn').addEventListener('click', () => {
            this.handleFormSubmitAndNew();
        });

        // Transaction type radio change
        document.querySelectorAll('input[name="transactionType"]').forEach(radio => {
            radio.addEventListener('change', (e) => {
                UI.updateStatusOptions(e.target.value);
                // Reset to pending when type changes
                document.getElementById('status').value = 'pending';
                UI.updateFormFieldVisibility('pending');
                // Pretax is now controlled by is_sales, not by type
                const catSelect = document.getElementById('category');
                const selectedOpt = catSelect.options[catSelect.selectedIndex];
                const isSales = selectedOpt && selectedOpt.dataset.isSales === '1';
                if (!isSales) {
                    document.getElementById('pretaxAmountGroup').style.display = 'none';
                    document.getElementById('pretaxAmount').value = '';
                }
            });
        });

        // Status change in form - show/hide dateProcessed and monthPaid, auto-fill date processed
        document.getElementById('status').addEventListener('change', (e) => {
            UI.updateFormFieldVisibility(e.target.value);
            // Auto-fill date processed and month paid when changing to paid/received
            if (e.target.value !== 'pending') {
                const dateProcessed = document.getElementById('dateProcessed');
                if (!dateProcessed.value) {
                    const today = Utils.getTodayDate();
                    dateProcessed.value = today;
                    document.getElementById('monthPaid').value = today.substring(0, 7);
                }
            }
        });

        // Date processed change - auto-fill month paid
        document.getElementById('dateProcessed').addEventListener('change', (e) => {
            if (e.target.value) {
                document.getElementById('monthPaid').value = e.target.value.substring(0, 7);
            }
        });

        // Today button
        document.getElementById('todayBtn').addEventListener('click', () => {
            const today = Utils.getTodayDate();
            document.getElementById('entryDate').value = today;
            if (!this._monthDueManuallySet) {
                document.getElementById('monthDue').value = today.substring(0, 7);
            }
        });

        // Entry date change - auto-fill month due if not manually set
        document.getElementById('entryDate').addEventListener('change', (e) => {
            if (e.target.value && !this._monthDueManuallySet) {
                document.getElementById('monthDue').value = e.target.value.substring(0, 7);
            }
        });

        // Track manual month due changes
        document.getElementById('monthDue').addEventListener('change', () => {
            this._monthDueManuallySet = true;
        });

        // Category change - auto-fill defaults and show/hide payment for month field
        document.getElementById('category').addEventListener('change', (e) => {
            const selectedOption = e.target.options[e.target.selectedIndex];
            if (!selectedOption || !selectedOption.value) return;

            // Auto-fill default amount if set and amount field is empty
            const defaultAmount = selectedOption.dataset.defaultAmount;
            const amountField = document.getElementById('amount');
            if (defaultAmount && (!amountField.value || amountField.value === '0')) {
                amountField.value = defaultAmount;
            }

            // Auto-fill default type if set
            const defaultType = selectedOption.dataset.defaultType;
            if (defaultType) {
                const typeRadio = document.querySelector(`input[name="transactionType"][value="${defaultType}"]`);
                if (typeRadio) {
                    typeRadio.checked = true;
                    UI.updateStatusOptions(defaultType);
                    // Apply default status if set, otherwise pending
                    const defaultStatus = selectedOption.dataset.defaultStatus;
                    const status = defaultStatus || 'pending';
                    document.getElementById('status').value = status;
                    UI.updateFormFieldVisibility(status);
                    // Auto-fill date processed and month paid if status is not pending
                    if (status !== 'pending') {
                        const dateProcessed = document.getElementById('dateProcessed');
                        if (!dateProcessed.value) {
                            const today = Utils.getTodayDate();
                            dateProcessed.value = today;
                            document.getElementById('monthPaid').value = today.substring(0, 7);
                        }
                    }
                }
            }

            // Sales category handling: show/hide pretax + sale date fields, force receivable
            const isSales = selectedOption.dataset.isSales === '1';
            const pretaxGroup = document.getElementById('pretaxAmountGroup');
            const saleDateGroup = document.getElementById('saleDateGroup');

            const inventoryCostGroup = document.getElementById('inventoryCostGroup');

            if (isSales) {
                pretaxGroup.style.display = 'flex';
                inventoryCostGroup.style.display = 'flex';
                saleDateGroup.style.display = 'flex';
                // Force receivable type for sales
                const receivableRadio = document.querySelector('input[name="transactionType"][value="receivable"]');
                if (receivableRadio) {
                    receivableRadio.checked = true;
                    UI.updateStatusOptions('receivable');
                }
            } else {
                pretaxGroup.style.display = 'none';
                document.getElementById('pretaxAmount').value = '';
                inventoryCostGroup.style.display = 'none';
                document.getElementById('inventoryCost').value = '';
                saleDateGroup.style.display = 'none';
                document.getElementById('saleDateStart').value = '';
                document.getElementById('saleDateEnd').value = '';
            }

            // Show/hide payment for month field
            if (selectedOption.dataset.isMonthly === '1') {
                UI.togglePaymentForMonth(true, selectedOption.textContent);
            } else {
                UI.togglePaymentForMonth(false);
            }
        });

        // ==================== JOURNAL OWNER ====================

        // Journal owner name change - save and update title
        const journalOwnerInput = document.getElementById('journalOwner');
        journalOwnerInput.addEventListener('input', Utils.debounce(() => {
            const owner = journalOwnerInput.value.trim();
            Database.setJournalOwner(owner);
            UI.updateJournalTitle(owner);
        }, 500));

        // Auto-size the owner input by measuring actual text width
        const measureSpan = document.createElement('span');
        measureSpan.style.cssText = 'position:absolute;visibility:hidden;white-space:pre;';
        document.body.appendChild(measureSpan);
        const autoSizeOwnerInput = () => {
            const style = getComputedStyle(journalOwnerInput);
            measureSpan.style.font = style.font;
            measureSpan.style.fontWeight = style.fontWeight;
            measureSpan.style.letterSpacing = style.letterSpacing;
            const text = journalOwnerInput.value || journalOwnerInput.placeholder;
            measureSpan.textContent = text;
            journalOwnerInput.style.width = (measureSpan.offsetWidth + 2) + 'px';
        };
        journalOwnerInput.addEventListener('input', autoSizeOwnerInput);
        autoSizeOwnerInput();

        // ==================== CATEGORIES ====================

        // Add category button (from entry form)
        document.getElementById('addCategoryBtn').addEventListener('click', () => {
            this._categoryModalOrigin = 'entry';
            this.openCategoryModal();
        });

        // Add category button (from budget expense form) - auto-create from expense name
        document.getElementById('addCategoryFromBudgetBtn').addEventListener('click', () => {
            // Auto-use the expense name as the category name
            const catName = document.getElementById('budgetExpenseName').value.trim();
            if (!catName) {
                UI.showNotification('Please enter an expense name first', 'error');
                return;
            }

            // Check if a category with this name already exists
            const existingCats = Database.getCategories();
            const duplicate = existingCats.find(c => c.name.toLowerCase() === catName.toLowerCase());
            if (duplicate) {
                UI.showNotification(`A category named "${duplicate.name}" already exists — please select it from the dropdown`, 'error');
                return;
            }

            // Find the Monthly Expenses folder (auto-created payable folder)
            const folders = Database.getFolders();
            const expenseFolder = folders.find(f => f.folder_type === 'payable');
            const folderId = expenseFolder ? expenseFolder.id : null;

            // Use the budget amount as typical price if entered
            const amountVal = document.getElementById('budgetExpenseAmount').value;
            const typicalPrice = amountVal ? parseFloat(amountVal) : null;

            try {
                const newId = Database.addCategory(
                    catName,
                    false,           // isMonthly
                    typicalPrice,    // defaultAmount
                    'payable',       // defaultType
                    folderId,        // folderId
                    false, false, false, false, false, null, false
                );
                this.refreshCategories();

                // Repopulate budget expense category dropdown and select new category
                const catSelect = document.getElementById('budgetExpenseCategory');
                catSelect.innerHTML = '<option value="">None (won\'t record)</option>';
                const cats = Database.getCategories();
                cats.forEach(cat => {
                    const opt = document.createElement('option');
                    opt.value = cat.id;
                    opt.textContent = cat.folder_name ? `${cat.folder_name} / ${cat.name}` : cat.name;
                    catSelect.appendChild(opt);
                });
                catSelect.value = newId;
                UI.showNotification('Category added', 'success');
            } catch (error) {
                console.error('Error adding category from budget:', error);
                UI.showNotification('Failed to add category', 'error');
            }
        });

        // Category form submission (handles both add and edit)
        document.getElementById('categoryForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleSaveCategory();
        });

        // Cancel category button
        document.getElementById('cancelCategoryBtn').addEventListener('click', () => {
            const wasEditing = !!document.getElementById('editingCategoryId').value;
            UI.hideModal('categoryModal');
            this.resetCategoryForm();
            this._categoryModalOrigin = null;
            if (wasEditing) {
                this.openManageCategories();
            }
        });

        // Manage Categories button
        document.getElementById('manageCategoriesBtn').addEventListener('click', () => {
            this.openManageCategories();
        });

        // Close manage categories
        document.getElementById('closeManageCategoriesBtn').addEventListener('click', () => {
            UI.hideModal('manageCategoriesModal');
        });

        // Add new category from manage modal
        document.getElementById('addNewCategoryFromManageBtn').addEventListener('click', () => {
            UI.hideModal('manageCategoriesModal');
            this.openCategoryModal();
        });

        // Add new folder from manage modal
        document.getElementById('addNewFolderBtn').addEventListener('click', () => {
            UI.hideModal('manageCategoriesModal');
            this.openFolderModal();
        });

        // Add folder from category modal (opens folder modal)
        document.getElementById('addFolderFromCategoryBtn').addEventListener('click', () => {
            this.folderCreatedFromCategory = true;
            UI.hideModal('categoryModal');
            this.openFolderModal();
        });

        // Enforce folder type on category default type
        document.getElementById('categoryFolder').addEventListener('change', (e) => {
            const selectedOption = e.target.options[e.target.selectedIndex];
            const folderType = selectedOption ? selectedOption.dataset.folderType : null;
            const defaultTypeSelect = document.getElementById('categoryDefaultType');

            if (folderType && folderType !== 'none') {
                defaultTypeSelect.value = folderType;
                defaultTypeSelect.disabled = true;
            } else {
                defaultTypeSelect.disabled = false;
            }
        });

        // Category edit/delete clicks (delegated) - also handles folder clicks
        document.getElementById('categoriesList').addEventListener('click', (e) => {
            const editBtn = e.target.closest('.edit-category-btn');
            const deleteBtn = e.target.closest('.delete-category-btn');
            const editFolderBtn = e.target.closest('.edit-folder-btn');
            const deleteFolderBtn = e.target.closest('.delete-folder-btn');
            const folderHeader = e.target.closest('.folder-header');

            if (editBtn && !editBtn.disabled) {
                this.handleEditCategory(parseInt(editBtn.dataset.id));
            } else if (deleteBtn && !deleteBtn.disabled) {
                this.handleDeleteCategory(parseInt(deleteBtn.dataset.id));
            } else if (editFolderBtn) {
                e.stopPropagation();
                this.handleEditFolder(parseInt(editFolderBtn.dataset.id));
            } else if (deleteFolderBtn) {
                e.stopPropagation();
                this.handleDeleteFolder(parseInt(deleteFolderBtn.dataset.id));
            } else if (folderHeader && !e.target.closest('.folder-actions')) {
                // Toggle folder collapse
                const folderId = folderHeader.dataset.folderId;
                const children = document.querySelector(`.folder-children[data-folder-id="${folderId}"]`);
                const toggle = folderHeader.querySelector('.folder-toggle');
                if (children) children.classList.toggle('collapsed');
                if (toggle) toggle.classList.toggle('collapsed');
            }
        });

        // Delete category confirmation
        document.getElementById('confirmDeleteCategoryBtn').addEventListener('click', () => {
            this.confirmDeleteCategory();
        });

        document.getElementById('cancelDeleteCategoryBtn').addEventListener('click', () => {
            UI.hideModal('deleteCategoryModal');
            this.deleteCategoryTargetId = null;
        });

        // Folder form submission
        document.getElementById('folderForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleSaveFolder();
        });

        // Cancel folder button
        document.getElementById('cancelFolderBtn').addEventListener('click', () => {
            UI.hideModal('folderModal');
            if (this.folderCreatedFromCategory) {
                this.folderCreatedFromCategory = false;
                UI.showModal('categoryModal');
            }
        });

        // Delete folder confirmation
        document.getElementById('confirmDeleteFolderBtn').addEventListener('click', () => {
            this.confirmDeleteFolder();
        });

        document.getElementById('cancelDeleteFolderBtn').addEventListener('click', () => {
            UI.hideModal('deleteFolderModal');
            this.deleteFolderTargetId = null;
        });

        // ==================== TRANSACTIONS ====================

        // Transaction actions (edit/delete/duplicate) - delegated
        document.getElementById('transactionsContainer').addEventListener('click', (e) => {
            const editBtn = e.target.closest('.edit-btn');
            const deleteBtn = e.target.closest('.delete-btn');
            const duplicateBtn = e.target.closest('.duplicate-btn');
            const notesIndicator = e.target.closest('.notes-indicator');

            if (duplicateBtn) {
                this.handleDuplicateTransaction(parseInt(duplicateBtn.dataset.id));
            } else if (editBtn) {
                this.handleEditTransaction(parseInt(editBtn.dataset.id));
            } else if (deleteBtn) {
                this.handleDeleteTransaction(parseInt(deleteBtn.dataset.id));
            } else if (notesIndicator) {
                // Toggle notes tooltip on click
                const tooltip = document.getElementById('notesTooltip');
                if (tooltip.classList.contains('visible') && tooltip.textContent === notesIndicator.dataset.notes) {
                    UI.hideNotesTooltip();
                } else {
                    UI.showNotesTooltip(notesIndicator.dataset.notes, notesIndicator);
                }
            }
        });

        // Notes tooltip on hover
        document.getElementById('transactionsContainer').addEventListener('mouseover', (e) => {
            const notesIndicator = e.target.closest('.notes-indicator');
            if (notesIndicator) {
                UI.showNotesTooltip(notesIndicator.dataset.notes, notesIndicator);
            }
        });

        document.getElementById('transactionsContainer').addEventListener('mouseout', (e) => {
            const notesIndicator = e.target.closest('.notes-indicator');
            if (notesIndicator) {
                UI.hideNotesTooltip();
            }
        });

        // Inline status dropdown change
        document.getElementById('transactionsContainer').addEventListener('change', (e) => {
            if (e.target.classList.contains('status-select')) {
                const id = parseInt(e.target.dataset.id);
                const newStatus = e.target.value;
                this.handleInlineStatusChange(id, newStatus, e.target);
            }
        });

        // Delete transaction confirmation
        document.getElementById('confirmDeleteBtn').addEventListener('click', () => {
            this.confirmDelete();
        });

        document.getElementById('cancelDeleteBtn').addEventListener('click', () => {
            UI.hideModal('deleteModal');
            this.deleteTargetId = null;
        });

        // ==================== MONTH PAID PROMPT (inline) ====================

        document.getElementById('confirmMonthPaidPromptBtn').addEventListener('click', () => {
            this.confirmMonthPaidPrompt();
        });

        document.getElementById('cancelMonthPaidPromptBtn').addEventListener('click', () => {
            this.cancelMonthPaidPrompt();
        });

        document.getElementById('paidTodayBtn').addEventListener('click', () => {
            this.confirmPaidToday();
        });

        // ==================== FILTERS & SORT ====================

        // Sort mode change
        document.getElementById('sortMode').addEventListener('change', (e) => {
            this.currentSortMode = e.target.value;
            this.clearBulkSelection();
            this.refreshTransactions();
        });

        ['filterType', 'filterStatus', 'filterMonth', 'filterCategory'].forEach(id => {
            document.getElementById(id).addEventListener('change', () => {
                this.clearBulkSelection();
                this.refreshTransactions();
            });
        });

        // Folder filter - cascades to category filter
        document.getElementById('filterFolder').addEventListener('change', () => {
            const folderId = document.getElementById('filterFolder').value;
            const categories = Database.getCategories();

            // Filter category dropdown to show only categories in selected folder
            if (folderId === 'unfiled') {
                const unfiled = categories.filter(c => !c.folder_id);
                UI.populateFilterCategories(unfiled);
            } else if (folderId) {
                const folderCats = categories.filter(c => c.folder_id === parseInt(folderId));
                UI.populateFilterCategories(folderCats);
            } else {
                UI.populateFilterCategories(categories);
            }

            document.getElementById('filterCategory').value = '';
            this.clearBulkSelection();
            this.refreshTransactions();
        });

        // ==================== SAVE / LOAD ====================

        document.getElementById('saveDbBtn').addEventListener('click', () => {
            this.handleSaveDatabase();
        });

        document.getElementById('saveAllDbBtn').addEventListener('click', () => {
            this.handleSaveAllDatabases();
        });

        document.getElementById('loadDbBtn').addEventListener('click', () => {
            document.getElementById('loadDbInput').click();
        });

        document.getElementById('loadDbInput').addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                this.pendingFileLoad = e.target.files[0];
                UI.showModal('loadConfirmModal');
            }
        });

        document.getElementById('confirmLoadBtn').addEventListener('click', () => {
            this.confirmLoadDatabase();
        });

        document.getElementById('cancelLoadBtn').addEventListener('click', () => {
            UI.hideModal('loadConfirmModal');
            this._cleanupPendingLoad();
        });

        // ==================== COMPANY SWITCHER ====================

        // Company button — toggle popover
        document.getElementById('companyBtn').addEventListener('click', (e) => {
            e.stopPropagation();
            const popover = document.getElementById('companyPopover');
            const isOpen = popover.style.display !== 'none';
            if (isOpen) {
                popover.style.display = 'none';
            } else {
                const btn = document.getElementById('companyBtn');
                const rect = btn.getBoundingClientRect();
                popover.style.top = (rect.bottom + 4) + 'px';
                popover.style.left = '8px';
                popover.style.display = 'block';
                CompanyManager.renderSwitcher();
            }
        });

        // Close company popover on outside click
        document.addEventListener('click', (e) => {
            const popover = document.getElementById('companyPopover');
            if (popover && popover.style.display !== 'none' && !e.target.closest('.company-switcher-wrapper')) {
                popover.style.display = 'none';
            }
        });

        // Switch company by clicking a list item
        document.getElementById('companyList').addEventListener('click', async (e) => {
            const item = e.target.closest('.company-list-item');
            if (!item) return;
            const id = item.dataset.id;
            if (id === CompanyManager.getRegistry().activeId) {
                document.getElementById('companyPopover').style.display = 'none';
                return;
            }
            document.getElementById('companyPopover').style.display = 'none';
            await this.switchToCompany(id);
        });

        // Add new company button
        document.getElementById('addCompanyBtn').addEventListener('click', async () => {
            document.getElementById('companyPopover').style.display = 'none';
            await this.handleCreateCompany();
        });

        // Open manage modal
        document.getElementById('manageCompaniesBtn').addEventListener('click', () => {
            document.getElementById('companyPopover').style.display = 'none';
            this.openManageCompanies();
        });

        // Manage modal — close
        document.getElementById('closeManageCompaniesBtn').addEventListener('click', () => {
            UI.hideModal('manageCompaniesModal');
        });

        // Manage modal — rename / delete (event delegation on tbody)
        document.getElementById('companiesTableBody').addEventListener('click', async (e) => {
            const renameBtn = e.target.closest('.company-rename-btn');
            const deleteBtn = e.target.closest('.company-delete-btn');
            if (renameBtn) {
                await this.handleRenameCompany(renameBtn.dataset.id);
            } else if (deleteBtn && !deleteBtn.disabled) {
                await this.handleDeleteCompany(deleteBtn.dataset.id);
                this.renderManageCompanies();
                this._updateCopySectionVisibility();
            }
        });

        // Copy section button
        document.getElementById('copySectionBtn').addEventListener('click', async () => {
            const sourceId = document.getElementById('copyFromCompanySelect').value;
            const section = document.getElementById('copySectionSelect').value;
            if (!sourceId) return;

            const status = document.getElementById('copySectionStatus');
            status.textContent = 'Copying…';
            try {
                const result = await CompanyManager.copySection(sourceId, section);
                status.textContent = `Copied ${result.copied} record(s).`;
                this.refreshAll();
            } catch (err) {
                console.error('Copy section failed:', err);
                status.textContent = 'Copy failed.';
            }
        });

        // Load choice modal buttons
        document.getElementById('loadReplaceBtn').addEventListener('click', () => {
            this.confirmLoadReplace();
        });

        document.getElementById('loadAddNewBtn').addEventListener('click', () => {
            this.confirmLoadAsNew();
        });

        document.getElementById('cancelLoadChoiceBtn').addEventListener('click', () => {
            UI.hideModal('loadChoiceModal');
            this._cleanupPendingLoad();
        });

        // Save As modal (fallback)
        document.getElementById('confirmSaveAsBtn').addEventListener('click', () => {
            this.confirmSaveAs();
        });

        document.getElementById('cancelSaveAsBtn').addEventListener('click', () => {
            UI.hideModal('saveAsModal');
        });

        // ==================== MODE TOGGLE ====================

        document.querySelectorAll('.mode-toggle-btn').forEach(btn => {
            btn.addEventListener('click', () => this.switchMode(btn.dataset.mode));
        });

        // ==================== QUICK ENTRY ====================

        document.getElementById('quickAddBtn').addEventListener('click', () => {
            const activeTab = document.querySelector('.main-tab.active');
            const tab = activeTab ? activeTab.dataset.tab : 'journal';
            this.openQuickEntry(tab === 'budget' ? 'budget' : 'journal');
        });

        document.getElementById('quickEntryCloseBtn').addEventListener('click', () => UI.hideModal('quickEntryModal'));
        document.getElementById('quickEntryCancelBtn').addEventListener('click', () => UI.hideModal('quickEntryModal'));
        document.getElementById('quickEntrySaveBtn').addEventListener('click', () => this.saveQuickEntries());
        document.getElementById('quickEntryAddRowBtn').addEventListener('click', () => { this._addQuickEntryRow(); this._updateQuickEntrySaveCount(); });

        document.getElementById('quickBudgetCloseBtn').addEventListener('click', () => UI.hideModal('quickBudgetModal'));
        document.getElementById('quickBudgetCancelBtn').addEventListener('click', () => UI.hideModal('quickBudgetModal'));
        document.getElementById('quickBudgetSaveBtn').addEventListener('click', () => this.saveQuickBudgetEntries());
        document.getElementById('quickBudgetAddRowBtn').addEventListener('click', () => { this._addQuickBudgetRow(); this._updateQuickBudgetSaveCount(); });

        // ==================== MAIN TABS ====================

        document.querySelectorAll('.main-tab[data-tab]').forEach(btn => {
            btn.addEventListener('click', () => {
                this.switchMainTab(btn.dataset.tab);
            });
        });

        // Tab context menu (right-click)
        this.setupTabContextMenu();

        // Manage tabs button
        document.getElementById('manageTabsBtn').addEventListener('click', () => this.openManageTabsModal());
        document.getElementById('manageTabsDoneBtn').addEventListener('click', () => UI.hideModal('manageTabsModal'));

        // P&L Tax Mode dropdown
        document.getElementById('plTaxMode').addEventListener('change', (e) => {
            Database.setPLTaxMode(e.target.value);
            this.refreshPnL();
        });

        // ==================== GEAR ICON / THEME CONTROLS ====================

        // Gear icon toggle popover
        document.getElementById('gearBtn').addEventListener('click', (e) => {
            e.stopPropagation();
            const popover = document.getElementById('gearPopover');
            popover.style.display = popover.style.display === 'none' ? 'flex' : 'none';
        });

        // Close popover on outside click
        document.addEventListener('click', (e) => {
            const popover = document.getElementById('gearPopover');
            if (popover.style.display !== 'none' && !e.target.closest('.gear-wrapper')) {
                popover.style.display = 'none';
            }
        });

        // Theme preset dropdown
        document.getElementById('themePreset').addEventListener('change', (e) => {
            const preset = e.target.value;
            Database.setThemePreset(preset);

            const picker = document.getElementById('customColorPicker');
            if (preset === 'custom') {
                picker.style.display = 'grid';
                let colors = Database.getThemeColors();
                if (!colors) {
                    colors = this.themePresets.default;
                    Database.setThemeColors(colors);
                }
                ['themeC1', 'themeC2', 'themeC3', 'themeC4', 'themeC5', 'themeC6'].forEach((id, i) => {
                    const input = document.getElementById(id);
                    if (input) input.value = colors[`c${i + 1}`] || '#000000';
                });
                this.applyTheme('custom', colors);
            } else {
                picker.style.display = 'none';
                this.applyTheme(preset, null);
            }
        });

        // Custom color picker inputs
        ['themeC1', 'themeC2', 'themeC3', 'themeC4', 'themeC5', 'themeC6'].forEach((id) => {
            const el = document.getElementById(id);
            if (el) el.addEventListener('input', Utils.debounce(() => {
                const colors = {
                    c1: document.getElementById('themeC1').value,
                    c2: document.getElementById('themeC2').value,
                    c3: document.getElementById('themeC3').value,
                    c4: document.getElementById('themeC4').value,
                    c5: document.getElementById('themeC5').value,
                    c6: document.getElementById('themeC6').value,
                };
                Database.setThemeColors(colors);
                this.applyTheme('custom', colors);
            }, 100));
        });

        // Dark mode toggle
        const darkToggle = document.getElementById('darkModeToggle');
        if (darkToggle) {
            darkToggle.addEventListener('change', (e) => {
                const isDark = e.target.checked;
                localStorage.setItem('darkMode', isDark);
                this.setDarkMode(isDark);
            });
        }

        // Shipping fee rate + minimum
        const shippingRateInput = document.getElementById('shippingFeeRate');
        const shippingMinInput = document.getElementById('shippingFeeMin');
        const saveShippingConfig = () => {
            const pct = parseFloat(shippingRateInput?.value) || 0;
            const minFee = parseFloat(shippingMinInput?.value) || 0;
            Database.setShippingFeeConfig({ rate: Math.round(pct * 1e4) / 1e6, minFee });
            this.syncAllB2BContractEntries();
            this.refreshAll();
        };
        if (shippingRateInput) {
            shippingRateInput.addEventListener('change', saveShippingConfig);
        }
        if (shippingMinInput) {
            shippingMinInput.addEventListener('change', saveShippingConfig);
        }

        // ==================== TIMELINE ====================
        ['timelineStartMonth', 'timelineStartYear', 'timelineEndMonth', 'timelineEndYear'].forEach(id => {
            document.getElementById(id).addEventListener('change', () => {
                this.handleTimelineChange();
            });
        });

        // ==================== ADD FOLDER ENTRIES ====================

        document.getElementById('addFolderEntriesBtn').addEventListener('click', () => {
            this.openAddFolderEntriesModal();
        });

        document.getElementById('bulkFolder').addEventListener('change', () => {
            this.updateBulkStatusOptions();
            this.updateBulkPreview();
        });

        document.getElementById('bulkMonthDue').addEventListener('change', () => {
            this.updateBulkPreview();
        });

        document.getElementById('bulkStatus').addEventListener('change', (e) => {
            this.toggleBulkProcessedFields(e.target.value);
            this.updateBulkPreview();
        });

        document.getElementById('confirmBulkBtn').addEventListener('click', () => {
            this.confirmAddFolderEntries();
        });

        document.getElementById('cancelBulkBtn').addEventListener('click', () => {
            UI.hideModal('addFolderEntriesModal');
        });

        // ==================== EXPORT ====================

        document.getElementById('exportCsvBtn').addEventListener('click', () => {
            this.handleExportCsv();
        });

        // ==================== BULK SELECT ====================

        document.getElementById('bulkSelectBtn').addEventListener('click', () => {
            this.toggleBulkSelectMode();
        });

        document.getElementById('bulkModePaid').addEventListener('click', () => {
            this.switchBulkMode('to-paid');
        });

        document.getElementById('bulkModePending').addEventListener('click', () => {
            this.switchBulkMode('to-pending');
        });

        document.getElementById('bulkMarkPaidBtn').addEventListener('click', () => {
            this.handleBulkMarkPaid();
        });

        document.getElementById('bulkResetPendingBtn').addEventListener('click', () => {
            this.handleBulkResetPending();
        });

        document.getElementById('bulkCancelBtn').addEventListener('click', () => {
            this.exitBulkSelectMode();
        });

        // Delegated checkbox handlers for bulk select
        document.getElementById('transactionsContainer').addEventListener('change', (e) => {
            if (e.target.classList.contains('bulk-checkbox')) {
                const id = parseInt(e.target.dataset.id);
                this.handleBulkCheckboxChange(id, e.target.checked);
            } else if (e.target.classList.contains('bulk-select-all')) {
                this.handleBulkSelectAll(e.target.checked);
            }
        });

        // Drag-to-select for bulk checkboxes
        const container = document.getElementById('transactionsContainer');
        this._bulkDragState = null;

        container.addEventListener('mousedown', (e) => {
            const cb = e.target.closest('.bulk-checkbox');
            if (!cb || !this.bulkSelectMode) return;
            // The checkbox toggles on click naturally; record the resulting state as the drag value
            const dragValue = !cb.checked; // will become cb.checked after the click completes
            this._bulkDragState = { active: true, value: dragValue };
        });

        container.addEventListener('mouseover', (e) => {
            if (!this._bulkDragState || !this._bulkDragState.active) return;
            const cb = e.target.closest('.bulk-checkbox');
            if (!cb) return;
            const id = parseInt(cb.dataset.id);
            cb.checked = this._bulkDragState.value;
            if (this._bulkDragState.value) {
                this.bulkSelectedIds.add(id);
            } else {
                this.bulkSelectedIds.delete(id);
            }
            this.updateBulkSelectCount();
            this.updateBulkSelectAllState();
        });

        document.addEventListener('mouseup', () => {
            if (this._bulkDragState) {
                this._bulkDragState = null;
            }
        });

        // ==================== CALCULATOR SIDEBAR ====================

        document.getElementById('calcToggleBtn').addEventListener('click', () => {
            this.toggleCalcMode();
        });

        document.getElementById('calcCloseBtn').addEventListener('click', () => {
            this.exitCalcMode();
        });

        document.getElementById('calcClearBtn').addEventListener('click', () => {
            this.calcClear();
        });

        document.getElementById('calcCopyResultBtn').addEventListener('click', () => {
            const text = document.getElementById('calcResult').textContent;
            navigator.clipboard.writeText(text).then(() => UI.showNotification('Result copied: ' + text, 'success'));
        });

        document.getElementById('calcCopyFormulaBtn').addEventListener('click', () => {
            const text = document.getElementById('calcFormulaBar').textContent;
            navigator.clipboard.writeText(text).then(() => UI.showNotification('Formula copied', 'success'));
        });

        document.getElementById('calcFormulaBar').addEventListener('input', () => {
            this.calcRecalculate();
        });

        // Paste: plain text only
        document.getElementById('calcFormulaBar').addEventListener('paste', (e) => {
            e.preventDefault();
            const text = (e.clipboardData || window.clipboardData).getData('text/plain');
            document.execCommand('insertText', false, text);
        });

        // Operator buttons for manual input
        document.querySelectorAll('.calc-op-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                this.calcInsertManual(btn.dataset.op);
            });
        });

        // Manual input: Enter key inserts as addition
        document.getElementById('calcManualValue').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                this.calcInsertManual('+');
            }
        });

        // Delegated click for calc cell selection (capture phase to intercept before edit handlers)
        document.querySelector('.app-main').addEventListener('click', (e) => {
            if (!this.calcMode) return;
            const cell = e.target.closest('.pnl-calc-cell, .cf-calc-cell, .bs-value, .txn-amount, .budget-amount, .loan-amount, .asset-amount, .ps-amount');
            if (!cell) return;
            e.preventDefault();
            e.stopPropagation();
            this.handleCalcCellClick(cell, e);
        }, true);

        // Escape to close calc mode
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && this.calcMode) {
                const formulaBar = document.getElementById('calcFormulaBar');
                if (document.activeElement === formulaBar) {
                    formulaBar.blur();
                } else {
                    this.exitCalcMode();
                }
            }
        });

        // Reference list click to highlight
        document.getElementById('calcRefList').addEventListener('click', (e) => {
            const item = e.target.closest('.calc-ref-item');
            if (!item) return;
            const refId = item.dataset.refId;
            const span = document.querySelector('#calcFormulaBar .calc-cell-ref[data-ref-id="' + refId + '"]');
            if (span) {
                span.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
                span.classList.remove('calc-ref-flash');
                void span.offsetWidth; // force reflow
                span.classList.add('calc-ref-flash');
            }
        });

        // ==================== BALANCE SHEET ====================

        // BS month/year change — persist selection
        document.getElementById('bsMonthMonth').addEventListener('change', () => {
            this._saveBsMonth();
            this.refreshBalanceSheet();
            if (this.currentMode === 'analyze') this._renderAnalyzeBSChart();
        });
        document.getElementById('bsMonthYear').addEventListener('change', () => {
            this._saveBsMonth();
            this.refreshBalanceSheet();
            if (this.currentMode === 'analyze') this._renderAnalyzeBSChart();
        });

        // Fixed asset form
        document.getElementById('fixedAssetForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleSaveFixedAsset();
        });

        document.getElementById('cancelAssetBtn').addEventListener('click', () => UI.hideModal('fixedAssetModal'));

        // Delete asset
        document.getElementById('confirmDeleteAssetBtn').addEventListener('click', () => this.confirmDeleteFixedAsset());
        document.getElementById('cancelDeleteAssetBtn').addEventListener('click', () => {
            UI.hideModal('deleteAssetModal');
            this.deleteAssetTargetId = null;
        });

        // Equity config (in Assets & Equity tab)
        document.getElementById('editEquityBtn').addEventListener('click', () => this.openEquityModal());
        document.getElementById('resyncEquityBtn').addEventListener('click', () => this.resyncEquityJournalEntries());
        document.getElementById('equityForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleSaveEquity();
        });
        document.getElementById('cancelEquityBtn').addEventListener('click', () => UI.hideModal('equityModal'));

        // ==================== FIXED ASSETS TAB ====================

        document.getElementById('addAssetTabBtn').addEventListener('click', () => this.openFixedAssetModal());

        // Assets list panel click delegation
        document.getElementById('assetsListPanel').addEventListener('click', (e) => {
            const editBtn = e.target.closest('.edit-asset-btn');
            const deleteBtn = e.target.closest('.delete-asset-btn');
            const item = e.target.closest('.asset-list-item');
            if (editBtn) {
                this.handleEditFixedAsset(parseInt(editBtn.dataset.id));
                return;
            }
            if (deleteBtn) {
                this.handleDeleteFixedAsset(parseInt(deleteBtn.dataset.id));
                return;
            }
            if (item) {
                this.selectedAssetId = parseInt(item.dataset.id);
                this.refreshFixedAssets();
            }
        });

        // ==================== LOANS TAB ====================

        document.getElementById('addLoanBtn').addEventListener('click', () => this.openLoanConfigModal());
        document.getElementById('loanConfigForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleSaveLoanConfig();
        });
        document.getElementById('cancelLoanConfigBtn').addEventListener('click', () => UI.hideModal('loanConfigModal'));

        // Delete loan modal
        document.getElementById('confirmDeleteLoanBtn').addEventListener('click', () => this.confirmDeleteLoan());
        document.getElementById('cancelDeleteLoanBtn').addEventListener('click', () => {
            UI.hideModal('deleteLoanModal');
            this.deleteLoanTargetId = null;
        });

        // Reset all data
        document.getElementById('resetAllDataBtn').addEventListener('click', (e) => {
            e.stopPropagation();
            document.getElementById('gearPopover').style.display = 'none';
            document.getElementById('resetConfirmInput').value = '';
            document.getElementById('confirmResetBtn').disabled = true;
            UI.showModal('resetAllDataModal');
        });
        document.getElementById('resetConfirmInput').addEventListener('input', (e) => {
            document.getElementById('confirmResetBtn').disabled = e.target.value.trim() !== 'RESET';
        });
        document.getElementById('confirmResetBtn').addEventListener('click', () => {
            try {
                Database.resetAllData();
                UI.hideModal('resetAllDataModal');
                this.selectedLoanId = null;
                this.selectedAssetId = null;
                this.selectedBudgetExpenseId = null;
                this.selectedB2BContractId = null;
                this._timeline = null;
                this.refreshAll();
                this.loadAndApplyTimeline();
                UI.showNotification('All data has been reset', 'success');
            } catch (error) {
                console.error('Error resetting data:', error);
                UI.showNotification('Failed to reset data: ' + error.message, 'error');
            }
        });
        document.getElementById('cancelResetBtn').addEventListener('click', () => {
            UI.hideModal('resetAllDataModal');
        });

        // Loan list panel click delegation
        document.getElementById('loanListPanel').addEventListener('click', (e) => {
            const editBtn = e.target.closest('.edit-loan-btn');
            const deleteBtn = e.target.closest('.delete-loan-btn');
            const item = e.target.closest('.loan-list-item');
            if (editBtn) {
                this.openLoanConfigModal(parseInt(editBtn.dataset.id));
                return;
            }
            if (deleteBtn) {
                this.handleDeleteLoan(parseInt(deleteBtn.dataset.id));
                return;
            }
            if (item) {
                this.selectedLoanId = parseInt(item.dataset.id);
                this.refreshLoans();
            }
        });

        // Loan payment skip/restore and cell editing delegation
        document.getElementById('loanDetailPanel').addEventListener('click', (e) => {
            const skipBtn = e.target.closest('.loan-skip-btn');
            if (skipBtn) {
                const loanId = parseInt(skipBtn.dataset.loanId);
                const paymentNum = parseInt(skipBtn.dataset.payment);
                Database.toggleSkipLoanPayment(loanId, paymentNum);
                const skippedLoan = Database.getLoanById(loanId);
                if (skippedLoan) this._syncLoanReceivable(loanId, skippedLoan);
                this.refreshAll();
                return;
            }

            // Click-to-edit payment amount
            const cell = e.target.closest('.loan-payment-cell');
            if (!cell || cell.querySelector('input') || cell.closest('.loan-payment-skipped')) return;

            const loanId = parseInt(cell.dataset.loanId);
            const paymentNum = parseInt(cell.dataset.payment);
            const currentText = cell.textContent.replace(/[^0-9.\-]/g, '');

            const input = document.createElement('input');
            input.type = 'number';
            input.step = '0.01';
            input.className = 'pnl-cell-input';
            input.value = currentText || '';
            cell.textContent = '';
            cell.appendChild(input);
            input.focus();
            input.select();

            const save = () => {
                const val = input.value.trim();
                if (val === '' || val === currentText) {
                    if (val === '') {
                        Database.setLoanPaymentOverride(loanId, paymentNum, null);
                    }
                } else {
                    Database.setLoanPaymentOverride(loanId, paymentNum, parseFloat(val));
                }
                const overrideLoan = Database.getLoanById(loanId);
                if (overrideLoan) this._syncLoanReceivable(loanId, overrideLoan);
                this.refreshAll();
            };

            input.addEventListener('blur', save);
            input.addEventListener('keydown', (ke) => {
                if (ke.key === 'Enter') { ke.preventDefault(); input.blur(); }
                else if (ke.key === 'Escape') { ke.preventDefault(); this.refreshLoans(); }
            });
        });

        // ==================== BUDGET TAB ====================

        document.getElementById('addBudgetGroupBtn').addEventListener('click', () => this.handleAddBudgetGroup());
        document.getElementById('addBudgetExpenseBtn').addEventListener('click', () => this.openBudgetExpenseModal());
        document.getElementById('budgetExpenseForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleSaveBudgetExpense();
        });
        document.getElementById('cancelBudgetExpenseBtn').addEventListener('click', () => UI.hideModal('budgetExpenseModal'));

        // Delete budget expense modal
        document.getElementById('confirmDeleteBudgetExpenseBtn').addEventListener('click', () => this.confirmDeleteBudgetExpense());
        document.getElementById('cancelDeleteBudgetExpenseBtn').addEventListener('click', () => {
            UI.hideModal('deleteBudgetExpenseModal');
            this.deleteBudgetExpenseTargetId = null;
        });

        // Budget list panel click delegation (original two-panel)
        document.getElementById('budgetListPanel').addEventListener('click', (e) => {
            const editBtn = e.target.closest('.edit-budget-btn');
            const deleteBtn = e.target.closest('.delete-budget-btn');
            const item = e.target.closest('.budget-list-item');
            if (editBtn) {
                this.handleEditBudgetExpense(parseInt(editBtn.dataset.id));
                return;
            }
            if (deleteBtn) {
                this.handleDeleteBudgetExpense(parseInt(deleteBtn.dataset.id));
                return;
            }
            if (item) {
                this.selectedBudgetExpenseId = parseInt(item.dataset.id);
                this.refreshBudget();
            }
        });

        // Budget groups container click delegation
        const budgetContainer = document.getElementById('budgetGroupsContainer');
        budgetContainer.addEventListener('click', (e) => {
            // Group actions
            const groupHeader = e.target.closest('.budget-group-section-header');
            const editGroupBtn = e.target.closest('.edit-budget-group-btn');
            const deleteGroupBtn = e.target.closest('.delete-budget-group-btn');

            if (editGroupBtn) {
                this.handleRenameBudgetGroup(parseInt(editGroupBtn.dataset.groupId));
                return;
            }
            if (deleteGroupBtn) {
                this.handleDeleteBudgetGroup(parseInt(deleteGroupBtn.dataset.groupId));
                return;
            }
            if (groupHeader && !e.target.closest('.budget-group-actions') && !e.target.closest('.budget-group-name')) {
                const groupId = groupHeader.dataset.groupId;
                if (groupId === '') return; // ungrouped header — no collapse
                const gid = parseInt(groupId);
                if (this.collapsedBudgetGroups.has(gid)) {
                    this.collapsedBudgetGroups.delete(gid);
                } else {
                    this.collapsedBudgetGroups.add(gid);
                }
                this.refreshBudget();
                return;
            }

            // Expense actions
            const editBtn = e.target.closest('.edit-budget-btn');
            const deleteBtn = e.target.closest('.delete-budget-btn');
            const item = e.target.closest('.budget-expense-row');
            if (editBtn) {
                this.handleEditBudgetExpense(parseInt(editBtn.dataset.id));
                return;
            }
            if (deleteBtn) {
                this.handleDeleteBudgetExpense(parseInt(deleteBtn.dataset.id));
                return;
            }
            if (item) {
                this.selectedBudgetExpenseId = parseInt(item.dataset.id);
                this.refreshBudget();
            }
        });

        // Budget detail panel close button
        document.getElementById('budgetDetailPanel').addEventListener('click', (e) => {
            if (e.target.closest('.budget-detail-close')) {
                this.selectedBudgetExpenseId = null;
                this.refreshBudget();
                return;
            }

            // Reset override button
            const resetBtn = e.target.closest('.budget-override-reset');
            if (resetBtn) {
                e.stopPropagation();
                const expenseId = parseInt(resetBtn.dataset.expenseId);
                const month = resetBtn.dataset.month;
                Database.setBudgetExpenseOverride(expenseId, month, null);
                this.syncAllBudgetJournalEntries();
                this.refreshAll();
                return;
            }

            // Click on amount cell to edit
            const amountCell = e.target.closest('.budget-amount-cell');
            if (amountCell && !amountCell.querySelector('input')) {
                const expenseId = parseInt(amountCell.dataset.expenseId);
                const month = amountCell.dataset.month;
                const defaultAmt = parseFloat(amountCell.dataset.default);
                const currentAmt = parseFloat(amountCell.textContent.replace(/[^0-9.\-]/g, ''));

                const input = document.createElement('input');
                input.type = 'number';
                input.className = 'budget-amount-inline-input';
                input.value = currentAmt;
                input.step = '0.01';
                input.setAttribute('data-default', defaultAmt);

                amountCell.textContent = '';
                amountCell.appendChild(input);
                input.focus();
                input.select();

                const commitEdit = () => {
                    const newAmt = parseFloat(input.value);
                    if (isNaN(newAmt) || newAmt < 0) {
                        this.refreshBudget();
                        return;
                    }
                    // If same as default, remove override; otherwise set it
                    if (Math.abs(newAmt - defaultAmt) < 0.005) {
                        Database.setBudgetExpenseOverride(expenseId, month, null);
                    } else {
                        Database.setBudgetExpenseOverride(expenseId, month, newAmt);
                    }
                    this.syncAllBudgetJournalEntries();
                    this.refreshAll();
                };

                input.addEventListener('blur', commitEdit);
                input.addEventListener('keydown', (ev) => {
                    if (ev.key === 'Enter') { ev.preventDefault(); input.blur(); }
                    if (ev.key === 'Escape') { this.refreshBudget(); }
                });
            }
        });

        // Budget groups container double-click for inline rename
        budgetContainer.addEventListener('dblclick', (e) => {
            const groupName = e.target.closest('.budget-group-name');
            if (groupName && groupName.dataset.groupId) {
                this.handleRenameBudgetGroup(parseInt(groupName.dataset.groupId));
            }
        });

        // Budget drag-and-drop
        budgetContainer.addEventListener('dragstart', (e) => {
            const item = e.target.closest('.budget-expense-row');
            if (item) {
                e.dataTransfer.setData('text/plain', item.dataset.id);
                e.dataTransfer.effectAllowed = 'move';
                item.classList.add('dragging');
            }
        });
        budgetContainer.addEventListener('dragend', (e) => {
            const item = e.target.closest('.budget-expense-row');
            if (item) item.classList.remove('dragging');
            budgetContainer.querySelectorAll('.drag-over').forEach(el => el.classList.remove('drag-over'));
        });
        budgetContainer.addEventListener('dragover', (e) => {
            const header = e.target.closest('.budget-group-section-header');
            if (header) {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
                budgetContainer.querySelectorAll('.drag-over').forEach(el => el.classList.remove('drag-over'));
                header.classList.add('drag-over');
            }
        });
        budgetContainer.addEventListener('dragleave', (e) => {
            const header = e.target.closest('.budget-group-section-header');
            if (header) header.classList.remove('drag-over');
        });
        budgetContainer.addEventListener('drop', (e) => {
            e.preventDefault();
            budgetContainer.querySelectorAll('.drag-over').forEach(el => el.classList.remove('drag-over'));
            const expenseId = parseInt(e.dataTransfer.getData('text/plain'));
            if (isNaN(expenseId)) return;
            const header = e.target.closest('.budget-group-section-header');
            if (!header) return;
            const groupId = header.dataset.groupId ? parseInt(header.dataset.groupId) : null;
            Database.moveBudgetExpenseToGroup(expenseId, groupId);
            this.refreshBudget();
            UI.showNotification('Expense moved', 'success');
        });

        // Record budget to journal
        document.getElementById('recordBudgetBtn').addEventListener('click', () => this.openRecordBudgetModal());
        document.getElementById('cancelRecordBudgetBtn').addEventListener('click', () => UI.hideModal('recordBudgetModal'));
        document.getElementById('confirmRecordBudgetBtn').addEventListener('click', () => this.confirmRecordBudget());
        document.getElementById('recordBudgetMonth').addEventListener('change', () => this.updateRecordBudgetPreview());
        document.getElementById('recordBudgetYear').addEventListener('change', () => this.updateRecordBudgetPreview());
        document.getElementById('recordBudgetStatus').addEventListener('change', () => {
            const status = document.getElementById('recordBudgetStatus').value;
            document.getElementById('recordBudgetDateProcessedGroup').style.display = status !== 'pending' ? '' : 'none';
        });

        // ==================== PRODUCTS TAB ====================
        document.getElementById('addProductBtn').addEventListener('click', () => this.openProductModal());
        document.getElementById('productForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleSaveProduct();
        });
        document.getElementById('cancelProductBtn').addEventListener('click', () => UI.hideModal('productModal'));
        document.getElementById('confirmDeleteProductBtn').addEventListener('click', () => this.confirmDeleteProduct());
        document.getElementById('cancelDeleteProductBtn').addEventListener('click', () => {
            UI.hideModal('deleteProductModal');
            this.deleteProductTargetId = null;
        });
        document.getElementById('pcShowDiscontinued').addEventListener('change', () => this.refreshProducts());

        // Product table click delegation
        document.getElementById('productTableWrapper').addEventListener('click', (e) => {
            const editBtn        = e.target.closest('.edit-product-btn');
            const deleteBtn      = e.target.closest('.delete-product-btn');
            const discontinueBtn = e.target.closest('.discontinue-product-btn');
            const linksBtn       = e.target.closest('.manage-links-btn');
            if (linksBtn) {
                this.openManageLinksModal(parseInt(linksBtn.dataset.id));
            } else if (editBtn) {
                this.handleEditProduct(parseInt(editBtn.dataset.id));
            } else if (deleteBtn) {
                this.handleDeleteProduct(parseInt(deleteBtn.dataset.id));
            } else if (discontinueBtn) {
                this.handleToggleDiscontinued(parseInt(discontinueBtn.dataset.id));
            }
        });

        // Products tab analytics date filter
        document.getElementById('pcDateFrom').addEventListener('change', () => this.refreshProducts());
        document.getElementById('pcDateTo').addEventListener('change',   () => this.refreshProducts());
        document.getElementById('pcDatePreset').addEventListener('change', () => this._pcApplyDatePreset());
        document.getElementById('pcSourceFilter').addEventListener('change', () => this.refreshProducts());
        document.getElementById('pcShowCharts').addEventListener('change', (e) => {
            const section = document.getElementById('pcAnalyticsSection');
            if (!section) return;
            const chartsRows = section.querySelectorAll('.pc-analytics-charts');
            chartsRows.forEach(el => el.style.display = e.target.checked ? '' : 'none');
        });

        // Product CSV Import
        document.getElementById('pcImportToggle').addEventListener('click', () => {
            document.getElementById('pcImportPanel').classList.toggle('open');
        });
        const pcZone = document.getElementById('pcCsvZone');
        document.getElementById('pcCsvBrowse').addEventListener('click', (e) => {
            e.stopPropagation();
            document.getElementById('pcCsvInput').click();
        });
        pcZone.addEventListener('click', () => document.getElementById('pcCsvInput').click());
        pcZone.addEventListener('dragover', (e) => { e.preventDefault(); pcZone.classList.add('dragover'); });
        pcZone.addEventListener('dragleave', () => pcZone.classList.remove('dragover'));
        pcZone.addEventListener('drop', (e) => {
            e.preventDefault();
            pcZone.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file) this._pcStageFile(file);
        });
        document.getElementById('pcCsvInput').addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) this._pcStageFile(file);
        });
        document.getElementById('pcImportSubmit').addEventListener('click', () => this._pcImportCsv());

        // Product-VE Mapping modal
        document.getElementById('pvmSaveBtn').addEventListener('click',   () => this.handleSaveProductMappings());
        document.getElementById('pvmCancelBtn').addEventListener('click', () => UI.hideModal('pvmModal'));
        document.getElementById('pvmSearchInput').addEventListener('input', (e) => {
            document.getElementById('pvmSuggestBtn').classList.remove('active');
            this.handlePvmSearch(e.target.value);
        });
        document.getElementById('pvmItemList').addEventListener('change', () => this._pvmUpdateSelectedCount());
        document.getElementById('pvmSuggestBtn').addEventListener('click', () => this.handlePvmSuggest());

        // ==================== BREAK-EVEN ====================
        document.getElementById('beConfigBtn').addEventListener('click', () => this.openBeConfigModal());
        document.getElementById('beConfigForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleSaveBeConfig();
        });
        document.getElementById('cancelBeConfigBtn').addEventListener('click', () => UI.hideModal('beConfigModal'));
        document.getElementById('beClearTimelineBtn').addEventListener('click', () => {
            document.getElementById('beTimelineStartMonth').value = '';
            document.getElementById('beTimelineStartYear').value = '';
            document.getElementById('beTimelineEndMonth').value = '';
            document.getElementById('beTimelineEndYear').value = '';
        });

        // Channel toggle visibility
        document.getElementById('beConsumerEnabled').addEventListener('change', (e) => {
            document.getElementById('beConsumerFields').style.display = e.target.checked ? 'flex' : 'none';
        });
        document.getElementById('beB2bEnabled').addEventListener('change', (e) => {
            document.getElementById('beB2bFields').style.display = e.target.checked ? 'flex' : 'none';
        });

        // CM preview updates
        ['beConsumerPrice', 'beConsumerCogs'].forEach(id => {
            document.getElementById(id).addEventListener('input', () => this._updateBeCmPreview('consumer'));
        });
        ['beB2bRate', 'beB2bCogs'].forEach(id => {
            document.getElementById(id).addEventListener('input', () => this._updateBeCmPreview('b2b'));
        });

        // Data source toggle updates cost hint and as-of visibility in real-time
        document.getElementById('beDataSource').addEventListener('change', () => {
            const isProjected = document.getElementById('beDataSource').value === 'projected';
            document.getElementById('beAsOfGroup').style.display = isProjected ? '' : 'none';
            this._updateBeFixedCostHints();
        });
        // As-of month change updates cost hint in real-time
        document.getElementById('beAsOfMonth').addEventListener('change', () => this._updateBeFixedCostHints());
        // Fixed cost override updates hint in real-time
        document.getElementById('beFixedCostOverride').addEventListener('input', () => this._updateBeFixedCostHints());

        // Progress tracker month dropdown
        document.getElementById('beProgressMonth').addEventListener('change', () => {
            if (this._beProgressState) {
                const { beResult, cfg, months, timelinePoints } = this._beProgressState;
                const asOf = document.getElementById('beProgressMonth').value;
                this._computeAndRenderProgress(beResult, cfg, months, asOf, timelinePoints);
            }
        });

        // ==================== DASHBOARD SNAPSHOT ====================
        document.getElementById('dashSnapshotMonth').addEventListener('change', (e) => {
            this._dashSnapshotMonth = e.target.value || null;
            this.refreshDashboard();
        });

        // ==================== BREAK-EVEN SNAPSHOT ====================
        document.getElementById('beSnapshotMonth').addEventListener('change', () => {
            this.refreshBreakeven();
        });

        // ==================== PROJECTED SALES ====================
        document.getElementById('psSaveBtn').addEventListener('click', () => this.handleSaveProjectedSales());
        document.getElementById('psEnabled').addEventListener('change', (e) => {
            document.getElementById('psConfigPanel').style.display = e.target.checked ? 'block' : 'none';
        });
        document.getElementById('psOnlineEnabled').addEventListener('change', (e) => {
            document.getElementById('psOnlineFields').style.display = e.target.checked ? 'flex' : 'none';
        });
        document.getElementById('psTradeshowEnabled').addEventListener('change', (e) => {
            document.getElementById('psTradeshowFields').style.display = e.target.checked ? 'flex' : 'none';
        });
        // CM preview live updates
        ['psOnlinePrice', 'psOnlineCogs'].forEach(id => {
            document.getElementById(id).addEventListener('input', () => {
                const price = parseFloat(document.getElementById('psOnlinePrice').value) || 0;
                const cogs = parseFloat(document.getElementById('psOnlineCogs').value) || 0;
                document.getElementById('psOnlineCm').textContent = Utils.formatCurrency(price - cogs);
            });
        });
        ['psTradeshowPrice', 'psTradeshowCogs'].forEach(id => {
            document.getElementById(id).addEventListener('input', () => {
                const price = parseFloat(document.getElementById('psTradeshowPrice').value) || 0;
                const cogs = parseFloat(document.getElementById('psTradeshowCogs').value) || 0;
                document.getElementById('psTradeshowCm').textContent = Utils.formatCurrency(price - cogs);
            });
        });
        // Per-tab view mode toggles (P&L and Cashflow each have their own)
        const pnlViewMode = document.getElementById('pnlViewMode');
        if (pnlViewMode) {
            pnlViewMode.addEventListener('change', () => this.refreshPnL());
        }
        const cfViewMode = document.getElementById('cfViewMode');
        if (cfViewMode) {
            cfViewMode.addEventListener('change', () => this.refreshCashFlow());
        }

        // Per-tab "As of" month pickers
        const pnlAsOfMonth = document.getElementById('pnlAsOfMonth');
        if (pnlAsOfMonth) {
            pnlAsOfMonth.addEventListener('change', () => {
                Database.setAsOfMonth('pnl', pnlAsOfMonth.value);
                this.refreshPnL();
            });
        }
        const cfAsOfMonth = document.getElementById('cfAsOfMonth');
        if (cfAsOfMonth) {
            cfAsOfMonth.addEventListener('change', () => {
                Database.setAsOfMonth('cf', cfAsOfMonth.value);
                this.refreshCashFlow();
            });
        }

        // Reset Projections buttons
        const pnlResetBtn = document.getElementById('pnlResetProjectionsBtn');
        if (pnlResetBtn) {
            pnlResetBtn.addEventListener('click', () => this.resetPnLProjections());
        }
        const cfResetBtn = document.getElementById('cfResetProjectionsBtn');
        if (cfResetBtn) {
            cfResetBtn.addEventListener('click', () => this.resetCashFlowProjections());
        }

        // ==================== MODALS & KEYBOARD ====================

        // Close modals on outside click
        document.querySelectorAll('.modal').forEach(modal => {
            modal.addEventListener('click', (e) => {
                if (e.target === modal) {
                    if (modal.id === 'entryModal') {
                        UI.resetForm();
                    } else {
                        UI.hideModal(modal.id);
                    }
                    // Cancel pending inline status change if month paid prompt is closed
                    if (modal.id === 'monthPaidPromptModal') {
                        this.cancelMonthPaidPrompt();
                    }
                }
            });
        });

        // Click outside tooltip to close
        document.addEventListener('click', (e) => {
            if (!e.target.closest('.notes-indicator') && !e.target.closest('.notes-tooltip')) {
                UI.hideNotesTooltip();
            }
        });

        // Keyboard shortcuts
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                const activeModals = document.querySelectorAll('.modal.active');
                activeModals.forEach(modal => {
                    if (modal.id === 'entryModal') {
                        UI.resetForm();
                    } else {
                        UI.hideModal(modal.id);
                    }
                    if (modal.id === 'monthPaidPromptModal') {
                        this.cancelMonthPaidPrompt();
                    }
                });
                UI.hideNotesTooltip();
            }

            // N = New Entry, Q = Quick Entry (only when no modifier keys, no modal open, no input focused)
            if (e.ctrlKey || e.altKey || e.metaKey || e.shiftKey) return;
            if (this.isViewOnly) return;
            const tag = document.activeElement.tagName;
            if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || document.activeElement.isContentEditable) return;
            if (document.querySelector('.modal.active')) return;

            if (e.key === 'n' || e.key === 'N') {
                e.preventDefault();
                document.getElementById('newEntryBtn').click();
            } else if (e.key === 'q' || e.key === 'Q') {
                e.preventDefault();
                this.openQuickEntry('journal');
            }
        });

        // ==================== SYNC / GROUP SHARING ====================

        document.getElementById('syncBtn').addEventListener('click', () => {
            this.openSyncMenu();
        });

        document.getElementById('closeSyncMenuBtn').addEventListener('click', () => {
            UI.hideModal('syncMenuModal');
        });

        // Tab switching
        document.querySelectorAll('.sync-tab').forEach(tab => {
            tab.addEventListener('click', () => {
                document.querySelectorAll('.sync-tab').forEach(t => t.classList.remove('active'));
                document.querySelectorAll('.sync-tab-content').forEach(c => c.classList.remove('active'));
                tab.classList.add('active');
                document.getElementById('sync' + tab.dataset.tab.charAt(0).toUpperCase() + tab.dataset.tab.slice(1) + 'Tab').classList.add('active');
            });
        });

        document.getElementById('createGroupForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleCreateGroup();
        });

        document.getElementById('joinGroupForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleJoinGroup();
        });

        // Join mode toggle (rejoin / new member)
        this._joinMode = 'rejoin';
        document.querySelectorAll('.sync-subtab').forEach(btn => {
            btn.addEventListener('click', () => {
                document.querySelectorAll('.sync-subtab').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                this._joinMode = btn.dataset.joinMode;
                const nameLabel = document.getElementById('syncJoinNameLabel');
                const passLabel = document.getElementById('syncJoinPasswordLabel');
                const nameInput = document.getElementById('syncJoinName');
                const passInput = document.getElementById('syncJoinPassword');
                if (this._joinMode === 'new') {
                    nameLabel.textContent = 'Choose a Username';
                    passLabel.textContent = 'Choose a Password';
                    nameInput.placeholder = 'Pick a display name...';
                    passInput.placeholder = 'Choose a password (min 4 chars)';
                } else {
                    nameLabel.textContent = 'Username';
                    passLabel.textContent = 'Password';
                    nameInput.placeholder = 'Enter your username';
                    passInput.placeholder = 'Enter your password';
                }
            });
        });

        document.getElementById('syncPullBtn').addEventListener('click', () => {
            this.handleSyncPull();
        });
        document.getElementById('syncHistoryBtn').addEventListener('click', () => {
            this.openVersionHistory();
        });
        document.getElementById('syncDisconnectBtn').addEventListener('click', () => {
            this.handleSyncDisconnect();
        });
        document.getElementById('syncCopyIdBtn').addEventListener('click', () => {
            this.copyToClipboard(this._currentInviteCode || '');
        });

        document.getElementById('closeGroupCreatedBtn').addEventListener('click', () => {
            UI.hideModal('groupCreatedModal');
        });
        document.getElementById('copyInviteCodeBtn').addEventListener('click', () => {
            const code = document.getElementById('inviteCodeDisplay').textContent;
            this.copyToClipboard(code);
        });

        document.getElementById('closeVersionHistoryBtn').addEventListener('click', () => {
            UI.hideModal('versionHistoryModal');
        });

        document.getElementById('cancelRollbackBtn').addEventListener('click', () => {
            UI.hideModal('rollbackModal');
        });
        document.getElementById('confirmRollbackBtn').addEventListener('click', () => {
            this.handleConfirmRollback();
        });

        // Members management (admin)
        document.getElementById('syncMembersBtn').addEventListener('click', () => {
            this.openMembersModal();
        });
        document.getElementById('closeMembersBtn').addEventListener('click', () => {
            UI.hideModal('membersModal');
        });
        document.getElementById('cancelRemoveMemberBtn').addEventListener('click', () => {
            UI.hideModal('removeMemberModal');
        });
        document.getElementById('confirmRemoveMemberBtn').addEventListener('click', () => {
            this.handleConfirmRemoveMember();
        });

        // Share view-only link
        document.getElementById('shareBtn').addEventListener('click', () => {
            this.openShareModal();
        });
        document.getElementById('generateShareBtn').addEventListener('click', () => {
            this.generateShare();
        });
        document.getElementById('closeShareBtn').addEventListener('click', () => {
            UI.hideModal('shareModal');
        });
        document.getElementById('closeShareResultBtn').addEventListener('click', () => {
            UI.hideModal('shareModal');
        });
        document.getElementById('copyShareUrlBtn').addEventListener('click', () => {
            const url = document.getElementById('shareUrlDisplay').value;
            if (url) this.copyToClipboard(url);
        });

        // B2B Contract events
        this.initB2BContractEvents();
    },

    // ==================== FORM HANDLERS ====================

    /**
     * Handle form submission (add/edit transaction)
     */
    handleFormSubmit() {
        if (this._guardViewOnly()) return;
        const data = UI.getFormData();
        const validation = UI.validateFormData(data);

        if (!validation.valid) {
            UI.showNotification(validation.message, 'error');
            return;
        }

        const editingId = document.getElementById('editingId').value;

        try {
            let parentId;
            if (editingId) {
                parentId = parseInt(editingId);
                const beforeTx = Database.getTransactionById(parentId);
                Database.updateTransaction(parentId, data);
                const afterData = Object.assign({}, data);
                this.pushUndo({
                    type: 'transaction-edit',
                    label: 'Edit "' + (data.item_description || 'transaction') + '"',
                    undo: () => { if (beforeTx) Database.updateTransaction(parentId, beforeTx); },
                    redo: () => { Database.updateTransaction(parentId, afterData); }
                });
                UI.showNotification('Transaction updated successfully', 'success');
            } else {
                parentId = Database.addTransaction(data);
                const addedData = Object.assign({}, data);
                this.pushUndo({
                    type: 'transaction-create',
                    label: 'Add "' + (data.item_description || 'transaction') + '"',
                    undo: () => { Database.deleteTransaction(parentId); },
                    redo: () => { Database.addTransaction(addedData); }
                });
                UI.showNotification('Transaction added successfully', 'success');
            }

            // Auto-manage linked sales tax entry
            this._manageSalesTaxEntry(parentId, data);

            // Auto-manage linked inventory cost entry
            this._manageInventoryCostEntry(parentId, data);

            UI.resetForm(); // This also closes the entry modal
            this.refreshAll();
        } catch (error) {
            console.error('Error saving transaction:', error);
            UI.showNotification('Failed to save transaction', 'error');
        }
    },

    /**
     * Handle form submission then reset form for another entry (keep modal open)
     */
    handleFormSubmitAndNew() {
        if (this._guardViewOnly()) return;
        const data = UI.getFormData();
        const validation = UI.validateFormData(data);

        if (!validation.valid) {
            UI.showNotification(validation.message, 'error');
            return;
        }

        const editingId = document.getElementById('editingId').value;

        try {
            let parentId;
            if (editingId) {
                parentId = parseInt(editingId);
                const beforeTx = Database.getTransactionById(parentId);
                Database.updateTransaction(parentId, data);
                const afterData = Object.assign({}, data);
                this.pushUndo({
                    type: 'transaction-edit',
                    label: 'Edit "' + (data.item_description || 'transaction') + '"',
                    undo: () => { if (beforeTx) Database.updateTransaction(parentId, beforeTx); },
                    redo: () => { Database.updateTransaction(parentId, afterData); }
                });
                UI.showNotification('Transaction updated — ready for next entry', 'success');
            } else {
                parentId = Database.addTransaction(data);
                const addedData = Object.assign({}, data);
                this.pushUndo({
                    type: 'transaction-create',
                    label: 'Add "' + (data.item_description || 'transaction') + '"',
                    undo: () => { Database.deleteTransaction(parentId); },
                    redo: () => { Database.addTransaction(addedData); }
                });
                UI.showNotification('Transaction added — ready for next entry', 'success');
            }

            // Auto-manage linked sales tax entry
            this._manageSalesTaxEntry(parentId, data);

            // Auto-manage linked inventory cost entry
            this._manageInventoryCostEntry(parentId, data);

            // Remember category and date for the next entry
            const keepCategoryId = data.category_id;
            const keepDate = data.entry_date;
            const keepType = data.transaction_type;

            // Advance month due by 1 month
            let nextMonthDue = '';
            if (data.month_due) {
                const [y, m] = data.month_due.split('-').map(Number);
                const nm = m === 12 ? 1 : m + 1;
                const ny = m === 12 ? y + 1 : y;
                nextMonthDue = ny + '-' + String(nm).padStart(2, '0');
            }

            // Reset form but keep modal open
            UI.resetForm(); // closes modal + clears fields
            UI.showModal('entryModal'); // reopen immediately

            // Restore kept values
            document.getElementById('entryDate').value = keepDate;
            document.getElementById('category').value = keepCategoryId;
            document.getElementById('monthDue').value = nextMonthDue;
            const typeRadio = document.querySelector('input[name="transactionType"][value="' + keepType + '"]');
            if (typeRadio) typeRadio.checked = true;
            UI.updateStatusOptions(keepType);

            // Switch to add mode
            document.getElementById('editingId').value = '';
            document.getElementById('formTitle').textContent = 'Add New Entry';
            document.getElementById('submitBtn').textContent = 'Add Entry';
            this._monthDueManuallySet = true;

            // Focus amount for fast next entry
            document.getElementById('amount').focus();

            this.refreshAll();
        } catch (error) {
            console.error('Error saving transaction:', error);
            UI.showNotification('Failed to save transaction', 'error');
        }
    },

    /**
     * Create, update, or delete a linked sales tax entry based on the parent sale
     * @param {number} parentId - Parent transaction ID
     * @param {Object} data - Parent transaction form data
     */
    _manageSalesTaxEntry(parentId, data) {
        const category = Database.getCategoryById(data.category_id);
        const existingChildId = Database.getLinkedSalesTaxTransaction(parentId);

        if (category && category.is_sales) {
            const taxAmount = data.amount - (data.pretax_amount || data.amount);
            const parentName = data.item_description ? data.item_description.trim() : '';
            const dateLabel = Utils.formatSaleDateRange(data.sale_date_start, data.sale_date_end);
            const description = parentName
                ? `Sales Tax – ${parentName}`
                : (dateLabel ? `Sales Tax ${dateLabel}` : 'Sales Tax');

            if (taxAmount > 0) {
                if (existingChildId) {
                    // Update existing linked sales tax entry
                    Database.updateSalesTaxTransaction(existingChildId, taxAmount, data.entry_date, data.month_due, description);
                } else {
                    // Create new sales tax entry
                    const salesTaxCatId = Database.getOrCreateSalesTaxCategory();
                    Database.addTransaction({
                        entry_date: data.entry_date,
                        category_id: salesTaxCatId,
                        item_description: description,
                        amount: taxAmount,
                        transaction_type: 'payable',
                        status: 'pending',
                        month_due: data.month_due,
                        source_type: 'sales_tax',
                        source_id: parentId
                    });
                }
            } else if (existingChildId) {
                // Tax is 0 or negative — remove the child
                Database.deleteTransaction(existingChildId);
            }
        } else if (existingChildId) {
            // Category changed away from sales — remove orphaned child
            Database.deleteTransaction(existingChildId);
        }
    },

    /**
     * Create, update, or delete a linked inventory cost entry based on the parent sale
     * @param {number} parentId - Parent transaction ID
     * @param {Object} data - Parent transaction form data
     */
    _manageInventoryCostEntry(parentId, data) {
        const category = Database.getCategoryById(data.category_id);
        const existingChildId = Database.getLinkedInventoryCostTransaction(parentId);

        if (category && category.is_sales) {
            const baseCost = data.inventory_cost || 0;
            const shippingConfig = Database.getShippingFeeConfig();
            const shippingRate = shippingConfig.rate || 0;
            const shippingMinFee = shippingConfig.minFee || 0;
            // Include shipping fee in COGS (integer-cent math to avoid rounding errors)
            const shippingFee = baseCost > 0 ? Math.max(Math.round(baseCost * shippingRate * 100) / 100, shippingMinFee) : 0;
            const costAmount = Math.round((baseCost + shippingFee) * 100) / 100;
            const dateLabel = Utils.formatSaleDateRange(data.sale_date_start, data.sale_date_end);
            const desc = dateLabel ? 'Inventory Cost' : 'Inventory Cost';
            const description = shippingRate > 0
                ? (dateLabel ? `${desc} ${dateLabel} (incl. shipping)` : `${desc} (incl. shipping)`)
                : (dateLabel ? `${desc} ${dateLabel}` : desc);

            if (costAmount > 0) {
                if (existingChildId) {
                    Database.updateInventoryCostTransaction(existingChildId, costAmount, data.entry_date, data.month_due, description);
                } else {
                    const inventoryCostCatId = Database.getOrCreateInventoryCostCategory();
                    Database.addTransaction({
                        entry_date: data.entry_date,
                        category_id: inventoryCostCatId,
                        item_description: description,
                        amount: costAmount,
                        transaction_type: 'payable',
                        status: 'pending',
                        month_due: data.month_due,
                        source_type: 'inventory_cost',
                        source_id: parentId
                    });
                }
            } else if (existingChildId) {
                Database.deleteTransaction(existingChildId);
            }

            // Clean up any stale separate shipping entries from prior logic
            const staleShipId = Database.getLinkedShippingTransaction(parentId);
            if (staleShipId) Database.deleteTransaction(staleShipId);
        } else if (existingChildId) {
            Database.deleteTransaction(existingChildId);
        }
    },

    // ==================== CATEGORY HANDLERS ====================

    /**
     * Open category modal for adding (resets the form)
     */
    openCategoryModal() {
        this.resetCategoryForm();
        document.getElementById('categoryModalTitle').textContent = 'Add New Category';
        document.getElementById('saveCategoryBtn').textContent = 'Add Category';
        UI.populateFolderDropdown(Database.getFolders());
        document.getElementById('categoryDefaultType').disabled = false;
        UI.showModal('categoryModal');
        document.getElementById('categoryName').focus();
    },

    /**
     * Reset the category form
     */
    resetCategoryForm() {
        document.getElementById('categoryForm').reset();
        document.getElementById('editingCategoryId').value = '';
        document.getElementById('categoryDefaultAmount').value = '';
        document.getElementById('categoryDefaultType').value = '';
        document.getElementById('categoryDefaultType').disabled = false;
        document.getElementById('categoryFolder').value = '';
        document.getElementById('categoryShowOnPl').checked = false;
        document.getElementById('categoryCogs').checked = false;
        document.getElementById('categoryDepreciation').checked = false;
        document.getElementById('categorySalesTax').checked = false;
        document.getElementById('categoryB2b').checked = false;
        document.getElementById('categoryIsSales').checked = false;
        document.getElementById('categoryDefaultStatus').value = '';
    },

    /**
     * Handle saving a category (add or edit)
     */
    handleSaveCategory() {
        if (this._guardViewOnly()) return;
        const nameInput = document.getElementById('categoryName');
        const name = nameInput.value.trim();
        const isMonthly = document.getElementById('categoryMonthly').checked;
        const editingId = document.getElementById('editingCategoryId').value;

        const defaultAmountRaw = document.getElementById('categoryDefaultAmount').value;
        const defaultAmount = defaultAmountRaw ? parseFloat(defaultAmountRaw) : null;
        const defaultType = document.getElementById('categoryDefaultType').value || null;
        const folderIdRaw = document.getElementById('categoryFolder').value;
        const folderId = folderIdRaw ? parseInt(folderIdRaw) : null;

        const showOnPl = document.getElementById('categoryShowOnPl').checked;
        const isCogs = document.getElementById('categoryCogs').checked;
        const isDepreciation = document.getElementById('categoryDepreciation').checked;
        const isSalesTax = document.getElementById('categorySalesTax').checked;
        const isB2b = document.getElementById('categoryB2b').checked;
        const isSales = document.getElementById('categoryIsSales').checked;
        const defaultStatus = document.getElementById('categoryDefaultStatus').value || null;

        if (!name) {
            UI.showNotification('Please enter a category name', 'error');
            return;
        }

        try {
            if (editingId) {
                // Update existing category
                Database.updateCategory(parseInt(editingId), name, isMonthly, defaultAmount, defaultType, folderId, showOnPl, isCogs, isDepreciation, isSalesTax, isB2b, defaultStatus, isSales);
                UI.showNotification('Category updated successfully', 'success');
            } else {
                // Add new category
                const newId = Database.addCategory(name, isMonthly, defaultAmount, defaultType, folderId, showOnPl, isCogs, isDepreciation, isSalesTax, isB2b, defaultStatus, isSales);
                this.refreshCategories();

                // Select the new category in the appropriate dropdown
                if (this._categoryModalOrigin === 'budget') {
                    // Repopulate budget expense category dropdown
                    const catSelect = document.getElementById('budgetExpenseCategory');
                    catSelect.innerHTML = '<option value="">None (won\'t record)</option>';
                    const cats = Database.getCategories();
                    cats.forEach(cat => {
                        const opt = document.createElement('option');
                        opt.value = cat.id;
                        opt.textContent = cat.name;
                        catSelect.appendChild(opt);
                    });
                    catSelect.value = newId;
                } else {
                    document.getElementById('category').value = newId;
                }

                if (isMonthly) {
                    UI.togglePaymentForMonth(true, name);
                }

                UI.showNotification('Category added successfully', 'success');
            }

            const wasEditing = !!editingId;
            UI.hideModal('categoryModal');
            this.resetCategoryForm();
            this.refreshCategories();
            this._categoryModalOrigin = null;

            // If we were editing (came from manage modal), re-open manage categories
            if (wasEditing || document.getElementById('manageCategoriesModal').classList.contains('active')) {
                this.openManageCategories();
            }
        } catch (error) {
            console.error('Error saving category:', error);
            if (error.message && error.message.includes('UNIQUE')) {
                UI.showNotification('A category with this name already exists', 'error');
            } else {
                UI.showNotification('Failed to save category', 'error');
            }
        }
    },

    /**
     * Open manage categories modal
     */
    openManageCategories() {
        const categories = Database.getCategories();
        UI.renderManageCategoriesList(categories);
        UI.showModal('manageCategoriesModal');
    },

    /**
     * Handle editing a category
     * @param {number} id - Category ID
     */
    handleEditCategory(id) {
        const category = Database.getCategoryById(id);
        if (!category) return;

        UI.hideModal('manageCategoriesModal');

        // Populate folder dropdown first
        UI.populateFolderDropdown(Database.getFolders());

        // Populate category form for editing
        document.getElementById('editingCategoryId').value = category.id;
        document.getElementById('categoryName').value = category.name;
        document.getElementById('categoryMonthly').checked = !!category.is_monthly;
        document.getElementById('categoryDefaultAmount').value = category.default_amount || '';
        document.getElementById('categoryDefaultType').value = category.default_type || '';
        document.getElementById('categoryFolder').value = category.folder_id || '';
        // P&L flags
        document.getElementById('categoryShowOnPl').checked = !!category.show_on_pl;
        document.getElementById('categoryCogs').checked = !!category.is_cogs;
        document.getElementById('categoryDepreciation').checked = !!category.is_depreciation;
        document.getElementById('categorySalesTax').checked = !!category.is_sales_tax;
        document.getElementById('categoryB2b').checked = !!category.is_b2b;
        document.getElementById('categoryIsSales').checked = !!category.is_sales;
        document.getElementById('categoryDefaultStatus').value = category.default_status || '';

        document.getElementById('categoryModalTitle').textContent = 'Edit Category';
        document.getElementById('saveCategoryBtn').textContent = 'Save Changes';

        // Enforce folder type on default type
        const defaultTypeSelect = document.getElementById('categoryDefaultType');
        if (category.folder_id) {
            const folderOption = document.querySelector(`#categoryFolder option[value="${category.folder_id}"]`);
            const folderType = folderOption ? folderOption.dataset.folderType : null;
            if (folderType && folderType !== 'none') {
                defaultTypeSelect.value = folderType;
                defaultTypeSelect.disabled = true;
            } else {
                defaultTypeSelect.disabled = false;
            }
        } else {
            defaultTypeSelect.disabled = false;
        }

        UI.showModal('categoryModal');
        document.getElementById('categoryName').focus();
    },

    /**
     * Handle deleting a category (show confirmation)
     * @param {number} id - Category ID
     */
    handleDeleteCategory(id) {
        if (this._guardViewOnly()) return;
        const category = Database.getCategoryById(id);
        if (!category) return;

        this.deleteCategoryTargetId = id;
        document.getElementById('deleteCategoryMessage').textContent =
            `Are you sure you want to delete "${category.name}"?`;
        UI.showModal('deleteCategoryModal');
    },

    /**
     * Confirm and execute category delete
     */
    confirmDeleteCategory() {
        if (this.deleteCategoryTargetId) {
            try {
                const success = Database.deleteCategory(this.deleteCategoryTargetId);
                if (success) {
                    UI.showNotification('Category deleted', 'success');
                    this.refreshCategories();
                    // Refresh manage categories list
                    if (document.getElementById('manageCategoriesModal').classList.contains('active')) {
                        this.openManageCategories();
                    }
                } else {
                    UI.showNotification('Cannot delete category that is in use', 'error');
                }
            } catch (error) {
                console.error('Error deleting category:', error);
                UI.showNotification('Failed to delete category', 'error');
            }
        }

        UI.hideModal('deleteCategoryModal');
        this.deleteCategoryTargetId = null;
    },

    // ==================== FOLDER HANDLERS ====================

    /**
     * Open folder modal for adding
     */
    openFolderModal() {
        document.getElementById('folderForm').reset();
        document.getElementById('editingFolderId').value = '';
        document.getElementById('folderModalTitle').textContent = 'Add New Folder';
        document.getElementById('saveFolderBtn').textContent = 'Add Folder';
        document.querySelector('input[name="folderType"][value="none"]').checked = true;
        UI.showModal('folderModal');
        document.getElementById('folderName').focus();
    },

    /**
     * Handle saving a folder (add or edit)
     */
    handleSaveFolder() {
        if (this._guardViewOnly()) return;
        const name = document.getElementById('folderName').value.trim();
        const editingId = document.getElementById('editingFolderId').value;
        const folderType = document.querySelector('input[name="folderType"]:checked').value;

        if (!name) {
            UI.showNotification('Please enter a folder name', 'error');
            return;
        }

        try {
            let newFolderId = null;
            if (editingId) {
                Database.updateFolder(parseInt(editingId), name, folderType);
                UI.showNotification('Folder updated', 'success');
            } else {
                newFolderId = Database.addFolder(name, folderType);
                UI.showNotification('Folder created', 'success');
            }

            UI.hideModal('folderModal');
            this.refreshCategories();

            // If folder was created from category modal, return there with folder pre-selected
            if (this.folderCreatedFromCategory && newFolderId) {
                this.folderCreatedFromCategory = false;
                UI.populateFolderDropdown(Database.getFolders());
                document.getElementById('categoryFolder').value = newFolderId;
                // Enforce the new folder's type on the category default type
                if (folderType && folderType !== 'none') {
                    document.getElementById('categoryDefaultType').value = folderType;
                    document.getElementById('categoryDefaultType').disabled = true;
                } else {
                    document.getElementById('categoryDefaultType').disabled = false;
                }
                UI.showModal('categoryModal');
            } else {
                this.folderCreatedFromCategory = false;
                this.openManageCategories();
            }
        } catch (error) {
            console.error('Error saving folder:', error);
            if (error.message && error.message.includes('UNIQUE')) {
                UI.showNotification('A folder with this name already exists', 'error');
            } else {
                UI.showNotification('Failed to save folder', 'error');
            }
        }
    },

    /**
     * Handle editing a folder
     * @param {number} id - Folder ID
     */
    handleEditFolder(id) {
        const folder = Database.getFolderById(id);
        if (!folder) return;

        document.getElementById('editingFolderId').value = folder.id;
        document.getElementById('folderName').value = folder.name;
        const typeRadio = document.querySelector(`input[name="folderType"][value="${folder.folder_type || 'payable'}"]`);
        if (typeRadio) typeRadio.checked = true;
        document.getElementById('folderModalTitle').textContent = 'Edit Folder';
        document.getElementById('saveFolderBtn').textContent = 'Save Changes';

        UI.showModal('folderModal');
        document.getElementById('folderName').focus();
    },

    /**
     * Handle deleting a folder (show confirmation)
     * @param {number} id - Folder ID
     */
    handleDeleteFolder(id) {
        if (this._guardViewOnly()) return;
        const folder = Database.getFolderById(id);
        if (!folder) return;

        this.deleteFolderTargetId = id;
        document.getElementById('deleteFolderMessage').textContent =
            `Are you sure you want to delete "${folder.name}"? Categories in this folder will become unfiled.`;
        UI.showModal('deleteFolderModal');
    },

    /**
     * Confirm and execute folder delete
     */
    confirmDeleteFolder() {
        if (this.deleteFolderTargetId) {
            try {
                Database.deleteFolder(this.deleteFolderTargetId);
                UI.showNotification('Folder deleted', 'success');
                this.refreshCategories();

                if (document.getElementById('manageCategoriesModal').classList.contains('active')) {
                    this.openManageCategories();
                }
            } catch (error) {
                console.error('Error deleting folder:', error);
                UI.showNotification('Failed to delete folder', 'error');
            }
        }

        UI.hideModal('deleteFolderModal');
        this.deleteFolderTargetId = null;
    },

    // ==================== TRANSACTION HANDLERS ====================

    /**
     * Handle inline status change from table dropdown
     * @param {number} id - Transaction ID
     * @param {string} newStatus - New status value
     * @param {HTMLElement} selectElement - The select element
     */
    handleInlineStatusChange(id, newStatus, selectElement) {
        if (newStatus === 'pending') {
            // Reverting to pending - clear processed date and month paid
            try {
                Database.updateTransactionStatus(id, newStatus);
                selectElement.className = `status-select status-${newStatus}`;
                this.refreshSummary();
                this.refreshTransactions();
                if (document.getElementById('cashflowTab').style.display !== 'none') this.refreshCashFlow();
                if (document.getElementById('pnlTab').style.display !== 'none') this.refreshPnL();
                UI.showNotification('Status updated', 'success');
            } catch (error) {
                console.error('Error updating status:', error);
                UI.showNotification('Failed to update status', 'error');
                this.refreshTransactions();
            }
        } else {
            // Changing to paid/received - prompt for month paid
            this.pendingInlineStatusChange = { id, newStatus, selectElement };

            // Pre-fill with today's date
            document.getElementById('promptDateProcessed').value = Utils.getTodayDate();

            // Update prompt title and paid today button text
            const title = newStatus === 'paid' ? 'Month Paid' : 'Month Received';
            document.getElementById('monthPaidPromptTitle').textContent = title;
            document.getElementById('paidTodayBtn').textContent =
                newStatus === 'paid' ? 'Paid Today' : 'Received Today';

            UI.showModal('monthPaidPromptModal');
        }
    },

    /**
     * Confirm the month paid prompt for inline status change
     */
    confirmMonthPaidPrompt() {
        const dateProcessed = document.getElementById('promptDateProcessed').value;

        if (!dateProcessed) {
            UI.showNotification('Please select a date', 'error');
            return;
        }

        // Bulk action path
        if (this.pendingBulkAction) {
            this.executeBulkDatePaid(dateProcessed);
            UI.hideModal('monthPaidPromptModal');
            return;
        }

        const monthPaid = dateProcessed.substring(0, 7);

        if (this.pendingInlineStatusChange) {
            const { id, newStatus, selectElement } = this.pendingInlineStatusChange;
            try {
                Database.updateTransactionStatus(id, newStatus, monthPaid);
                Database.setTransactionDateProcessed(id, dateProcessed);
                selectElement.className = `status-select status-${newStatus}`;
                this.refreshSummary();
                this.refreshTransactions();
                if (document.getElementById('cashflowTab').style.display !== 'none') this.refreshCashFlow();
                if (document.getElementById('pnlTab').style.display !== 'none') this.refreshPnL();
                UI.showNotification('Status updated', 'success');
            } catch (error) {
                console.error('Error updating status:', error);
                UI.showNotification('Failed to update status', 'error');
                this.refreshTransactions();
            }
        }

        UI.hideModal('monthPaidPromptModal');
        this.pendingInlineStatusChange = null;
    },

    /**
     * Quick "Paid Today" / "Received Today" - sets status, month_paid, and date_processed to today
     */
    confirmPaidToday() {
        // Bulk action path
        if (this.pendingBulkAction) {
            this.executeBulkDatePaid(Utils.getTodayDate());
            UI.hideModal('monthPaidPromptModal');
            return;
        }

        if (!this.pendingInlineStatusChange) return;

        const { id, newStatus, selectElement } = this.pendingInlineStatusChange;
        const today = Utils.getTodayDate();
        const monthPaid = today.substring(0, 7);

        try {
            Database.updateTransactionStatus(id, newStatus, monthPaid);
            Database.setTransactionDateProcessed(id, today);
            selectElement.className = `status-select status-${newStatus}`;
            this.refreshSummary();
            this.refreshTransactions();
            if (document.getElementById('cashflowTab').style.display !== 'none') this.refreshCashFlow();
            if (document.getElementById('pnlTab').style.display !== 'none') this.refreshPnL();
            UI.showNotification(`Marked as ${newStatus} today`, 'success');
        } catch (error) {
            console.error('Error updating status:', error);
            UI.showNotification('Failed to update status', 'error');
            this.refreshTransactions();
        }

        UI.hideModal('monthPaidPromptModal');
        this.pendingInlineStatusChange = null;
    },

    /**
     * Cancel the month paid prompt - revert the select
     */
    cancelMonthPaidPrompt() {
        if (this.pendingInlineStatusChange) {
            // Revert the select element to previous status
            this.refreshTransactions();
        }
        UI.hideModal('monthPaidPromptModal');
        this.pendingInlineStatusChange = null;
        this.pendingBulkAction = false;
    },

    // ==================== BULK SELECT ====================

    toggleBulkSelectMode() {
        if (this.bulkSelectMode) {
            this.exitBulkSelectMode();
        } else {
            this.bulkSelectMode = true;
            this.bulkSelectDirection = 'to-paid';
            this.bulkSelectedIds = new Set();
            document.getElementById('bulkSelectBtn').textContent = 'Done';
            document.getElementById('bulkActionBar').style.display = 'flex';
            document.getElementById('bulkModePaid').classList.add('active');
            document.getElementById('bulkModePending').classList.remove('active');
            document.getElementById('bulkMarkPaidBtn').style.display = '';
            document.getElementById('bulkResetPendingBtn').style.display = 'none';
            this.updateBulkSelectCount();
            this.refreshTransactions();
        }
    },

    exitBulkSelectMode() {
        this.bulkSelectMode = false;
        this.bulkSelectedIds = new Set();
        this.pendingBulkAction = false;
        document.getElementById('bulkSelectBtn').textContent = 'Select';
        document.getElementById('bulkActionBar').style.display = 'none';
        this.refreshTransactions();
    },

    switchBulkMode(direction) {
        this.bulkSelectDirection = direction;
        this.bulkSelectedIds = new Set();

        document.getElementById('bulkModePaid').classList.toggle('active', direction === 'to-paid');
        document.getElementById('bulkModePending').classList.toggle('active', direction === 'to-pending');
        document.getElementById('bulkMarkPaidBtn').style.display = direction === 'to-paid' ? '' : 'none';
        document.getElementById('bulkResetPendingBtn').style.display = direction === 'to-pending' ? '' : 'none';

        this.updateBulkSelectCount();
        this.refreshTransactions();
    },

    handleBulkCheckboxChange(id, checked) {
        if (checked) {
            this.bulkSelectedIds.add(id);
        } else {
            this.bulkSelectedIds.delete(id);
        }
        this.updateBulkSelectCount();
        this.updateBulkSelectAllState();
    },

    updateBulkSelectAllState() {
        const allCheckboxes = document.querySelectorAll('.bulk-checkbox');
        const selectAll = document.querySelector('.bulk-select-all');
        if (selectAll && allCheckboxes.length > 0) {
            selectAll.checked = Array.from(allCheckboxes).every(cb => cb.checked);
        }
    },

    handleBulkSelectAll(checked) {
        const checkboxes = document.querySelectorAll('.bulk-checkbox');
        checkboxes.forEach(cb => {
            cb.checked = checked;
            const id = parseInt(cb.dataset.id);
            if (checked) {
                this.bulkSelectedIds.add(id);
            } else {
                this.bulkSelectedIds.delete(id);
            }
        });
        this.updateBulkSelectCount();
    },

    updateBulkSelectCount() {
        const count = this.bulkSelectedIds.size;
        document.getElementById('bulkSelectCount').textContent = `${count} selected`;
    },

    clearBulkSelection() {
        if (this.bulkSelectMode) {
            this.bulkSelectedIds = new Set();
            this.updateBulkSelectCount();
        }
    },

    handleBulkMarkPaid() {
        if (this.bulkSelectedIds.size === 0) {
            UI.showNotification('No transactions selected', 'error');
            return;
        }

        this.pendingBulkAction = true;

        // Determine title based on selected transaction types
        let hasPayable = false;
        let hasReceivable = false;
        for (const id of this.bulkSelectedIds) {
            const t = Database.getTransactionById(id);
            if (t) {
                if (t.transaction_type === 'payable') hasPayable = true;
                if (t.transaction_type === 'receivable') hasReceivable = true;
            }
        }

        let title = 'Mark as Paid';
        let btnText = 'Paid Today';
        if (hasReceivable && !hasPayable) {
            title = 'Mark as Received';
            btnText = 'Received Today';
        } else if (hasReceivable && hasPayable) {
            title = 'Mark as Paid/Received';
            btnText = 'Paid/Received Today';
        }

        document.getElementById('monthPaidPromptTitle').textContent = title;
        document.getElementById('paidTodayBtn').textContent = btnText;
        document.getElementById('promptDateProcessed').value = Utils.getTodayDate();

        UI.showModal('monthPaidPromptModal');
    },

    executeBulkDatePaid(dateProcessed) {
        const monthPaid = dateProcessed.substring(0, 7);
        const updates = [];

        for (const id of this.bulkSelectedIds) {
            const t = Database.getTransactionById(id);
            if (t) {
                const status = t.transaction_type === 'receivable' ? 'received' : 'paid';
                updates.push({ id, status });
            }
        }

        try {
            // Snapshot before states for undo
            const beforeStates = updates.map(u => {
                const t = Database.getTransactionById(u.id);
                return { id: u.id, status: t.status, date_processed: t.date_processed, month_paid: t.month_paid };
            });

            Database.bulkSetDatePaid(updates, dateProcessed, monthPaid);

            this.pushUndo({
                type: 'bulk-status-change',
                label: 'Bulk mark ' + updates.length + ' as paid',
                undo: () => {
                    for (const b of beforeStates) {
                        Database.updateTransactionStatus(b.id, b.status, b.month_paid);
                        if (b.date_processed) Database.setTransactionDateProcessed(b.id, b.date_processed);
                    }
                },
                redo: () => { Database.bulkSetDatePaid(updates, dateProcessed, monthPaid); }
            });

            this.refreshSummary();
            this.refreshTransactions();
            if (document.getElementById('cashflowTab').style.display !== 'none') this.refreshCashFlow();
            if (document.getElementById('pnlTab').style.display !== 'none') this.refreshPnL();
            UI.showNotification(`${updates.length} transaction${updates.length !== 1 ? 's' : ''} marked as paid`, 'success');
        } catch (error) {
            console.error('Error bulk updating:', error);
            UI.showNotification('Failed to update transactions', 'error');
        }

        this.exitBulkSelectMode();
    },

    handleBulkResetPending() {
        if (this.bulkSelectedIds.size === 0) {
            UI.showNotification('No transactions selected', 'error');
            return;
        }

        const count = this.bulkSelectedIds.size;
        if (!confirm(`Reset ${count} transaction${count !== 1 ? 's' : ''} to pending? This will clear their paid dates.`)) {
            return;
        }

        const ids = Array.from(this.bulkSelectedIds);

        try {
            // Snapshot before states for undo
            const beforeStates = ids.map(id => {
                const t = Database.getTransactionById(id);
                return { id, status: t.status, date_processed: t.date_processed, month_paid: t.month_paid };
            });

            Database.bulkResetToPending(ids);

            this.pushUndo({
                type: 'bulk-reset-pending',
                label: 'Bulk reset ' + count + ' to pending',
                undo: () => {
                    for (const b of beforeStates) {
                        Database.updateTransactionStatus(b.id, b.status, b.month_paid);
                        if (b.date_processed) Database.setTransactionDateProcessed(b.id, b.date_processed);
                    }
                },
                redo: () => { Database.bulkResetToPending(ids); }
            });

            this.refreshSummary();
            this.refreshTransactions();
            if (document.getElementById('cashflowTab').style.display !== 'none') this.refreshCashFlow();
            if (document.getElementById('pnlTab').style.display !== 'none') this.refreshPnL();
            UI.showNotification(`${count} transaction${count !== 1 ? 's' : ''} reset to pending`, 'success');
        } catch (error) {
            console.error('Error bulk resetting:', error);
            UI.showNotification('Failed to reset transactions', 'error');
        }

        this.exitBulkSelectMode();
    },

    /**
     * Handle edit transaction
     * @param {number} id - Transaction ID
     */
    handleEditTransaction(id) {
        const transaction = Database.getTransactionById(id);
        if (transaction) {
            UI.populateFormForEdit(transaction);
        }
    },

    /**
     * Handle duplicate transaction - copies entry into a new form with incremented month
     * @param {number} id - Transaction ID
     */
    handleDuplicateTransaction(id) {
        const transaction = Database.getTransactionById(id);
        if (!transaction) return;

        // Load all fields from the original transaction
        UI.populateFormForEdit(transaction);

        // Clear editing ID so it creates a new entry
        document.getElementById('editingId').value = '';

        // Set entry date to today
        const today = Utils.getTodayDate();
        document.getElementById('entryDate').value = today;

        // Increment month_due by 1 month
        if (transaction.month_due) {
            const [y, m] = transaction.month_due.split('-').map(Number);
            const newMonth = m === 12 ? 1 : m + 1;
            const newYear = m === 12 ? y + 1 : y;
            document.getElementById('monthDue').value = `${newYear}-${String(newMonth).padStart(2, '0')}`;
        } else {
            document.getElementById('monthDue').value = today.substring(0, 7);
        }

        // Increment payment_for_month by 1 month (if set)
        if (transaction.payment_for_month) {
            const [y, m] = transaction.payment_for_month.split('-').map(Number);
            const newMonth = m === 12 ? 1 : m + 1;
            const newYear = m === 12 ? y + 1 : y;
            document.getElementById('paymentForMonth').value = `${newYear}-${String(newMonth).padStart(2, '0')}`;
        }

        // Reset to pending status and clear processed fields
        document.getElementById('status').value = 'pending';
        UI.updateFormFieldVisibility('pending');
        document.getElementById('dateProcessed').value = '';
        document.getElementById('monthPaid').value = '';

        // Clear sale dates for duplicate — user must re-enter for the new period
        document.getElementById('saleDateStart').value = '';
        document.getElementById('saleDateEnd').value = '';

        // Update form title
        document.getElementById('formTitle').textContent = 'Duplicate Entry';
        document.getElementById('submitBtn').textContent = 'Add Entry';

        this._monthDueManuallySet = true; // Don't override the incremented month due
    },

    /**
     * Handle delete transaction (show confirmation)
     * @param {number} id - Transaction ID
     */
    handleDeleteTransaction(id) {
        this.deleteTargetId = id;
        UI.showModal('deleteModal');
    },

    /**
     * Confirm and execute delete
     */
    confirmDelete() {
        if (this.deleteTargetId) {
            try {
                // Snapshot for undo before deleting
                const deletedTx = Database.getTransactionById(this.deleteTargetId);
                const childId = Database.getLinkedSalesTaxTransaction(this.deleteTargetId);
                const deletedChild = childId ? Database.getTransactionById(childId) : null;
                const invCostChildId = Database.getLinkedInventoryCostTransaction(this.deleteTargetId);
                const deletedInvCost = invCostChildId ? Database.getTransactionById(invCostChildId) : null;
                const shippingChildId = Database.getLinkedShippingTransaction(this.deleteTargetId);
                const deletedShipping = shippingChildId ? Database.getTransactionById(shippingChildId) : null;

                // Cascade-delete linked sales tax entry if present
                if (childId) {
                    Database.deleteTransaction(childId);
                }
                // Cascade-delete linked inventory cost entry if present
                if (invCostChildId) {
                    Database.deleteTransaction(invCostChildId);
                }
                // Cascade-delete linked shipping fee entry if present
                if (shippingChildId) {
                    Database.deleteTransaction(shippingChildId);
                }
                Database.deleteTransaction(this.deleteTargetId);

                // Push undo action
                if (deletedTx) {
                    const label = 'Delete "' + (deletedTx.item_description || 'transaction') + '"';
                    this.pushUndo({
                        type: 'transaction-delete', label,
                        undo: () => {
                            Database.addTransaction(deletedTx);
                            if (deletedChild) Database.addTransaction(deletedChild);
                            if (deletedInvCost) Database.addTransaction(deletedInvCost);
                        },
                        redo: () => {
                            // Re-delete by description match since IDs may change
                            const all = Database.getTransactions();
                            const match = all.find(t => t.entry_date === deletedTx.entry_date && t.amount === deletedTx.amount && t.item_description === deletedTx.item_description);
                            if (match) Database.deleteTransaction(match.id);
                            if (deletedChild) {
                                const cm = all.find(t => t.entry_date === deletedChild.entry_date && t.amount === deletedChild.amount);
                                if (cm) Database.deleteTransaction(cm.id);
                            }
                            if (deletedInvCost) {
                                const im = all.find(t => t.entry_date === deletedInvCost.entry_date && t.amount === deletedInvCost.amount);
                                if (im) Database.deleteTransaction(im.id);
                            }
                        }
                    });
                }

                UI.showNotification('Transaction deleted', 'success');
                this.refreshAll();
            } catch (error) {
                console.error('Error deleting transaction:', error);
                UI.showNotification('Failed to delete transaction', 'error');
            }
        }

        UI.hideModal('deleteModal');
        this.deleteTargetId = null;
    },

    // ==================== ADD FOLDER ENTRIES ====================

    /**
     * Open the Add Folder Entries modal
     */
    openAddFolderEntriesModal() {
        const folders = Database.getFolders();
        const select = document.getElementById('bulkFolder');
        select.innerHTML = '<option value="">Select folder...</option>';
        folders.forEach(f => {
            const opt = document.createElement('option');
            opt.value = f.id;
            opt.textContent = `${f.name} (${UI.capitalizeFirst(f.folder_type)})`;
            opt.dataset.folderType = f.folder_type;
            select.appendChild(opt);
        });

        // Pre-fill month due with current month
        document.getElementById('bulkMonthDue').value = Utils.getCurrentMonth();

        // Pre-fill entry date with today
        document.getElementById('bulkEntryDate').value = Utils.getTodayDate();

        // Reset status and hide processed fields
        document.getElementById('bulkStatus').innerHTML = '<option value="pending">Pending</option>';
        document.getElementById('bulkStatus').value = 'pending';
        this.toggleBulkProcessedFields('pending');

        // Reset preview and button
        document.getElementById('bulkPreview').innerHTML = '';
        document.getElementById('confirmBulkBtn').disabled = true;

        UI.showModal('addFolderEntriesModal');
    },

    /**
     * Update bulk status dropdown options based on selected folder type
     */
    updateBulkStatusOptions() {
        const select = document.getElementById('bulkFolder');
        const selectedOption = select.options[select.selectedIndex];
        const folderType = selectedOption ? selectedOption.dataset.folderType : null;
        const statusSelect = document.getElementById('bulkStatus');

        if (!folderType) {
            statusSelect.innerHTML = '<option value="pending">Pending</option>';
            statusSelect.value = 'pending';
            this.toggleBulkProcessedFields('pending');
            return;
        }

        if (folderType === 'receivable') {
            statusSelect.innerHTML = `
                <option value="pending">Pending</option>
                <option value="received">Received</option>
            `;
            document.getElementById('bulkMonthPaidLabel').textContent = 'Month Received';
        } else if (folderType === 'payable') {
            statusSelect.innerHTML = `
                <option value="pending">Pending</option>
                <option value="paid">Paid</option>
            `;
            document.getElementById('bulkMonthPaidLabel').textContent = 'Month Paid';
        } else {
            // 'none' type - show all options
            statusSelect.innerHTML = `
                <option value="pending">Pending</option>
                <option value="paid">Paid</option>
                <option value="received">Received</option>
            `;
            document.getElementById('bulkMonthPaidLabel').textContent = 'Month Paid/Received';
        }

        statusSelect.value = 'pending';
        this.toggleBulkProcessedFields('pending');
    },

    /**
     * Show/hide date processed and month paid fields in bulk modal
     * @param {string} status - 'pending', 'paid', or 'received'
     */
    toggleBulkProcessedFields(status) {
        const dateProcessedGroup = document.getElementById('bulkDateProcessedGroup');
        const monthPaidGroup = document.getElementById('bulkMonthPaidGroup');

        if (status === 'pending') {
            dateProcessedGroup.style.display = 'none';
            monthPaidGroup.style.display = 'none';
            document.getElementById('bulkDateProcessed').value = '';
            document.getElementById('bulkMonthPaid').value = '';
        } else {
            dateProcessedGroup.style.display = 'flex';
            monthPaidGroup.style.display = 'flex';
        }
    },

    /**
     * Update the bulk add preview when folder or month changes
     */
    updateBulkPreview() {
        const folderId = document.getElementById('bulkFolder').value;
        const monthDueValue = document.getElementById('bulkMonthDue').value;
        const status = document.getElementById('bulkStatus').value;
        const preview = document.getElementById('bulkPreview');
        const confirmBtn = document.getElementById('confirmBulkBtn');

        if (!folderId) {
            preview.innerHTML = '';
            confirmBtn.disabled = true;
            return;
        }

        const categories = Database.getCategoriesByFolder(parseInt(folderId)).filter(c => !c.is_sales);
        const selectedOption = document.getElementById('bulkFolder').options[document.getElementById('bulkFolder').selectedIndex];
        const folderType = selectedOption.dataset.folderType || 'payable';

        if (categories.length === 0) {
            preview.innerHTML = '<div class="bulk-preview-empty">No categories in this folder. Add categories to this folder first.</div>';
            confirmBtn.disabled = true;
            return;
        }

        const monthDue = monthDueValue || null;
        const monthLabel = monthDue ? Utils.formatMonthDisplay(monthDue) : 'not set';
        const isNoneType = folderType === 'none';
        const folderTypeClass = folderType === 'receivable' ? 'type-receivable' : folderType === 'payable' ? 'type-payable' : 'type-none';
        const folderTypeLabel = isNoneType ? '' : ` &mdash; <span class="type-badge ${folderTypeClass}">${UI.capitalizeFirst(folderType)}</span>`;

        let html = `<div class="bulk-preview-header">${categories.length} entr${categories.length === 1 ? 'y' : 'ies'} will be created for ${monthLabel}${folderTypeLabel} &mdash; ${UI.capitalizeFirst(status)}</div>`;

        // Check for categories missing default amounts
        const missingAmounts = categories.filter(c => !c.default_amount);
        if (missingAmounts.length > 0) {
            html += `<div class="bulk-preview-warning">` +
                `${missingAmounts.length} categor${missingAmounts.length === 1 ? 'y is' : 'ies are'} missing a typical price. ` +
                `Those will use $0.00 as default. Edit categories to set defaults.</div>`;
        }

        categories.forEach(cat => {
            const amount = cat.default_amount || 0;
            const catType = isNoneType ? (cat.default_type || 'payable') : folderType;
            const catTypeClass = catType === 'receivable' ? 'type-receivable' : 'type-payable';
            html += `
                <div class="bulk-preview-item">
                    <span class="bulk-preview-name">${Utils.escapeHtml(cat.name)}</span>
                    <div class="bulk-preview-details">
                        <span class="type-badge ${catTypeClass}">${UI.capitalizeFirst(catType)}</span>
                        <span>${Utils.formatCurrency(amount)}</span>
                    </div>
                </div>
            `;
        });

        preview.innerHTML = html;
        confirmBtn.disabled = false;
    },

    /**
     * Confirm and create all folder entries
     */
    confirmAddFolderEntries() {
        const folderSelect = document.getElementById('bulkFolder');
        const folderId = parseInt(folderSelect.value);
        const entryDate = document.getElementById('bulkEntryDate').value;
        const status = document.getElementById('bulkStatus').value;
        const dateProcessed = document.getElementById('bulkDateProcessed').value || null;

        if (!folderId) {
            UI.showNotification('Please select a folder', 'error');
            return;
        }

        if (!entryDate) {
            UI.showNotification('Please enter a date recorded', 'error');
            return;
        }

        // Get folder type from selected option
        const selectedOption = folderSelect.options[folderSelect.selectedIndex];
        const folderType = selectedOption.dataset.folderType || 'payable';

        const monthDue = document.getElementById('bulkMonthDue').value || null;
        const monthPaid = (status !== 'pending') ? (document.getElementById('bulkMonthPaid').value || null) : null;
        const categories = Database.getCategoriesByFolder(folderId).filter(c => !c.is_sales);

        if (categories.length === 0) {
            UI.showNotification('No categories in this folder', 'error');
            return;
        }

        // Validate month paid if status is paid/received
        if (status !== 'pending' && !monthPaid) {
            UI.showNotification('Please select the month paid/received', 'error');
            return;
        }

        let count = 0;

        const isNoneType = folderType === 'none';

        try {
            categories.forEach(cat => {
                const amount = cat.default_amount || 0;
                const catType = isNoneType ? (cat.default_type || 'payable') : folderType;
                const paymentForMonth = cat.is_monthly ? monthDue : null;

                Database.addTransaction({
                    entry_date: entryDate,
                    category_id: cat.id,
                    item_description: null,
                    amount: amount,
                    transaction_type: catType,
                    status: status,
                    date_processed: (status !== 'pending') ? dateProcessed : null,
                    month_due: monthDue,
                    month_paid: (status !== 'pending') ? monthPaid : null,
                    payment_for_month: paymentForMonth,
                    notes: null
                });
                count++;
            });

            UI.hideModal('addFolderEntriesModal');
            this.refreshAll();
            UI.showNotification(`${count} entr${count === 1 ? 'y' : 'ies'} added successfully`, 'success');
        } catch (error) {
            console.error('Error adding folder entries:', error);
            UI.showNotification('Failed to add entries', 'error');
        }
    },

    // ==================== BALANCE SHEET ====================

    /**
     * Refresh the Balance Sheet tab
     */
    refreshBalanceSheet() {
        const month = document.getElementById('bsMonthMonth').value;
        const year = document.getElementById('bsMonthYear').value;

        if (!month || !year) {
            document.getElementById('balanceSheetContent').innerHTML =
                '<p class="empty-state">Select a date to view the Balance Sheet.</p>';
            const ratiosEl = document.getElementById('bsRatiosContent');
            if (ratiosEl) ratiosEl.innerHTML = '';
            return;
        }

        const asOfMonth = `${year}-${month}`;
        const taxMode = Database.getPLTaxMode();

        // Gather all balance sheet data
        const cash = Database.getCashAsOf(asOfMonth);
        const ar = Database.getAccountsReceivableAsOf(asOfMonth);
        const ap = Database.getAccountsPayableAsOf(asOfMonth);
        const salesTaxPayable = Database.getSalesTaxPayableAsOf(asOfMonth);

        // Fixed assets and depreciation (using Utils.computeDepreciationSchedule)
        const fixedAssets = Database.getFixedAssets();

        const round2 = (v) => Math.round(v * 100) / 100;
        let totalFixedAssetCost = 0;
        let totalAccumDepr = 0;
        const assetDetails = fixedAssets.map(asset => {
            const deprSchedule = Utils.computeDepreciationSchedule(asset);
            let accumDepr = 0;
            Object.entries(deprSchedule).forEach(([m, amt]) => {
                if (m <= asOfMonth) accumDepr = round2(accumDepr + amt);
            });

            totalFixedAssetCost = round2(totalFixedAssetCost + asset.purchase_cost);
            totalAccumDepr = round2(totalAccumDepr + accumDepr);

            return {
                ...asset,
                accum_depreciation: accumDepr,
                net_book_value: round2(asset.purchase_cost - accumDepr)
            };
        });

        const netFixedAssets = round2(totalFixedAssetCost - totalAccumDepr);
        const totalAssets = round2(cash + ar + netFixedAssets);

        // Multi-loan balances
        const loans = Database.getLoans();
        let totalLoanBalance = 0;
        const loanDetails = loans.map(loan => {
            // Don't include loan as liability until its start date
            const loanStartMonth = loan.start_date ? loan.start_date.substring(0, 7) : '';
            if (loanStartMonth && loanStartMonth > asOfMonth) {
                return { name: loan.name, balance: 0 };
            }

            const skipped = Database.getSkippedPayments(loan.id);
            const overrides = Database.getLoanPaymentOverrides(loan.id);
            const schedule = Utils.computeAmortizationSchedule({
                principal: loan.principal,
                annual_rate: loan.annual_rate,
                term_months: loan.term_months,
                payments_per_year: loan.payments_per_year,
                start_date: loan.start_date,
                first_payment_date: loan.first_payment_date
            }, skipped, overrides);

            let balance = loan.principal;
            for (let i = schedule.length - 1; i >= 0; i--) {
                if (schedule[i].month <= asOfMonth) {
                    balance = schedule[i].ending_balance;
                    break;
                }
            }
            if (schedule.length > 0 && schedule[0].month > asOfMonth) {
                balance = loan.principal;
            }

            totalLoanBalance = round2(totalLoanBalance + balance);
            return { name: loan.name, balance };
        });

        const totalLiabilities = round2(ap + salesTaxPayable + totalLoanBalance);

        // Equity — only include amounts if as-of month >= effective date
        const equityConfig = Database.getEquityConfig();
        const seedEffective = (equityConfig.seed_received_date || equityConfig.seed_expected_date || '');
        const apicEffective = (equityConfig.apic_expected_date || equityConfig.apic_received_date || '');
        const seedMonth = seedEffective ? seedEffective.substring(0, 7) : '';
        const apicMonth = apicEffective ? apicEffective.substring(0, 7) : '';

        const commonStock = (seedMonth && seedMonth > asOfMonth) ? 0 : round2(equityConfig.common_stock_par * equityConfig.common_stock_shares);
        const apicVal = (apicMonth && apicMonth > asOfMonth) ? 0 : round2(equityConfig.apic || 0);

        const retainedEarnings = round2(Database.getRetainedEarningsAsOf(asOfMonth, taxMode));
        const totalEquity = round2(commonStock + apicVal + retainedEarnings);

        const totalLiabilitiesAndEquity = round2(totalLiabilities + totalEquity);
        const isBalanced = Math.abs(totalAssets - totalLiabilitiesAndEquity) < 0.01;

        // AR/AP category breakdowns
        const arByCategory = Database.getARByCategory(asOfMonth);
        const apByCategory = Database.getAPByCategory(asOfMonth);

        // P&L totals through as-of month (for financial ratios)
        const plTotals = Database.getPLTotalsThrough(asOfMonth, taxMode);

        const bsData = {
            asOfMonth,
            cash, ar, arByCategory,
            assetDetails, totalFixedAssetCost, totalAccumDepr, netFixedAssets,
            totalAssets,
            ap, apByCategory, salesTaxPayable,
            loanDetails, totalLoanBalance,
            totalLiabilities,
            commonStock, apic: apicVal, retainedEarnings, totalEquity,
            totalLiabilitiesAndEquity, isBalanced,
            plTotals
        };

        UI.renderBalanceSheet(bsData);
    },

    /**
     * Initialize Balance Sheet month/year dropdowns, restoring saved selection
     */
    initBalanceSheetDate() {
        const saved = Database.getAsOfMonth('bs');
        const fallback = Utils.getCurrentMonth();
        const target = (saved && saved !== 'current') ? saved : fallback;
        const [year, month] = target.split('-');
        document.getElementById('bsMonthMonth').value = month;

        // Populate year dropdown (timeline-aware)
        const yearSelect = document.getElementById('bsMonthYear');
        const timeline = this.getTimeline();
        const years = Utils.getYearsInTimeline(timeline.start, timeline.end);
        yearSelect.innerHTML = '<option value="">Year...</option>';
        years.forEach(y => {
            const opt = document.createElement('option');
            opt.value = y;
            opt.textContent = y;
            yearSelect.appendChild(opt);
        });
        yearSelect.value = year;
    },

    /**
     * Persist the current Balance Sheet month/year selection
     */
    _saveBsMonth() {
        const month = document.getElementById('bsMonthMonth').value;
        const year = document.getElementById('bsMonthYear').value;
        if (month && year) {
            Database.setAsOfMonth('bs', `${year}-${month}`);
        }
    },

    // ==================== FIXED ASSETS & LOANS ====================

    /**
     * Refresh the Fixed Assets tab
     */
    refreshFixedAssets() {
        const assets = Database.getFixedAssets();
        UI.renderFixedAssetsTab(assets, this.selectedAssetId);
        // Also render equity section
        const equityConfig = Database.getEquityConfig();
        UI.renderEquitySection(equityConfig);
    },

    /**
     * Refresh the Loans tab (multi-loan)
     */
    refreshLoans() {
        const loans = Database.getLoans();
        UI.renderLoansTab(loans, this.selectedLoanId);
    },

    // ==================== FIXED ASSET HANDLERS ====================

    openFixedAssetModal() {
        document.getElementById('fixedAssetForm').reset();
        document.getElementById('editingAssetId').value = '';
        document.getElementById('fixedAssetModalTitle').textContent = 'Add Fixed Asset';
        document.getElementById('saveAssetBtn').textContent = 'Add Asset';
        document.getElementById('assetDate').value = Utils.getTodayDate();
        document.getElementById('assetDeprMethod').value = 'straight_line';
        document.getElementById('assetAutoTransaction').checked = true;
        UI.showModal('fixedAssetModal');
        document.getElementById('assetName').focus();
    },

    handleSaveFixedAsset() {
        const name = document.getElementById('assetName').value.trim();
        const cost = parseFloat(document.getElementById('assetCost').value);
        const salvage = parseFloat(document.getElementById('assetSalvage').value) || 0;
        const life = parseInt(document.getElementById('assetLife').value) || 0;
        const date = document.getElementById('assetDate').value;
        const deprMethod = document.getElementById('assetDeprMethod').value;
        const deprStart = document.getElementById('assetDeprStart').value || null;
        const isDepreciable = deprMethod !== 'none';
        const notes = document.getElementById('assetNotes').value.trim() || null;
        const autoTransaction = document.getElementById('assetAutoTransaction').checked;
        const editingId = document.getElementById('editingAssetId').value;

        if (!name || isNaN(cost) || !date) {
            UI.showNotification('Please fill in name, cost, and date', 'error');
            return;
        }

        if (isDepreciable && life <= 0) {
            UI.showNotification('Depreciable assets need a useful life > 0', 'error');
            return;
        }

        if (isDepreciable && salvage >= cost) {
            UI.showNotification('Salvage value must be less than cost for depreciable assets', 'error');
            return;
        }

        try {
            if (editingId) {
                Database.updateFixedAsset(parseInt(editingId), name, cost, life, date, salvage, deprMethod, deprStart, isDepreciable, notes);
                UI.showNotification('Asset updated', 'success');
            } else {
                const assetId = Database.addFixedAsset(name, cost, life, date, salvage, deprMethod, deprStart, isDepreciable, notes);
                if (autoTransaction) {
                    this._autoCreateAssetTransaction(assetId, name, cost, date, deprStart);
                }
                UI.showNotification('Asset added', 'success');
            }
            UI.hideModal('fixedAssetModal');
            this.refreshFixedAssets();
            this.refreshBalanceSheet();
        } catch (error) {
            console.error('Error saving asset:', error);
            UI.showNotification('Failed to save asset', 'error');
        }
    },

    /**
     * Auto-create a payable transaction for a fixed asset purchase
     */
    _autoCreateAssetTransaction(assetId, name, cost, date, deprStart) {
        // Find or create an "Asset Purchases" category (hidden from P&L — capital expenditure, not OpEx)
        let categories = Database.getCategories();
        let cat = categories.find(c => c.name === 'Asset Purchases');
        let catId;
        if (!cat) {
            // show_on_pl=true means hidden from P&L (inverted semantics)
            catId = Database.addCategory('Asset Purchases', false, null, 'payable', null, true, false, false, false);
        } else {
            catId = cat.id;
        }

        const dueMonth = deprStart ? deprStart.substring(0, 7) : date.substring(0, 7);
        const paidMonth = date.substring(0, 7);
        Database.addTransaction({
            entry_date: date,
            category_id: catId,
            item_description: `Purchase: ${name}`,
            amount: cost,
            transaction_type: 'payable',
            status: 'paid',
            month_due: dueMonth,
            month_paid: paidMonth,
            date_processed: date,
            source_type: 'asset_purchase',
            source_id: assetId
        });

        // Get the last inserted transaction ID and link it back
        const result = Database.db.exec('SELECT last_insert_rowid() as id');
        const txId = result[0].values[0][0];
        Database.linkTransactionToAsset(assetId, txId);
    },

    handleEditFixedAsset(id) {
        const asset = Database.getFixedAssetById(id);
        if (!asset) return;

        document.getElementById('editingAssetId').value = asset.id;
        document.getElementById('assetName').value = asset.name;
        document.getElementById('assetCost').value = asset.purchase_cost;
        document.getElementById('assetSalvage').value = asset.salvage_value || 0;
        document.getElementById('assetLife').value = asset.useful_life_months;
        document.getElementById('assetDate').value = asset.purchase_date;
        document.getElementById('assetDeprMethod').value = asset.depreciation_method || 'straight_line';
        document.getElementById('assetDeprStart').value = asset.dep_start_date || '';
        document.getElementById('assetNotes').value = asset.notes || '';
        document.getElementById('assetAutoTransaction').checked = false; // Don't re-create on edit
        document.getElementById('fixedAssetModalTitle').textContent = 'Edit Fixed Asset';
        document.getElementById('saveAssetBtn').textContent = 'Save Changes';
        UI.showModal('fixedAssetModal');
    },

    handleDeleteFixedAsset(id) {
        this.deleteAssetTargetId = id;
        UI.showModal('deleteAssetModal');
    },

    confirmDeleteFixedAsset() {
        if (this.deleteAssetTargetId) {
            Database.deleteFixedAsset(this.deleteAssetTargetId);
            this.selectedAssetId = null;
            UI.showNotification('Asset deleted', 'success');
            this.refreshFixedAssets();
            this.refreshBalanceSheet();
        }
        UI.hideModal('deleteAssetModal');
        this.deleteAssetTargetId = null;
    },

    // ==================== EQUITY CONFIG HANDLERS ====================

    openEquityModal() {
        const config = Database.getEquityConfig();
        document.getElementById('equityPar').value = config.common_stock_par || '';
        document.getElementById('equityShares').value = config.common_stock_shares || '';
        document.getElementById('equityApic').value = config.apic || '';
        document.getElementById('seedExpectedDate').value = config.seed_expected_date || '';
        document.getElementById('seedReceivedDate').value = config.seed_received_date || '';
        document.getElementById('apicExpectedDate').value = config.apic_expected_date || '';
        document.getElementById('apicReceivedDate').value = config.apic_received_date || '';
        document.getElementById('equityAutoTransaction').checked = true;
        UI.showModal('equityModal');
    },

    handleSaveEquity() {
        const par = parseFloat(document.getElementById('equityPar').value) || 0;
        const shares = parseInt(document.getElementById('equityShares').value) || 0;
        const apic = parseFloat(document.getElementById('equityApic').value) || 0;
        const seedExpected = document.getElementById('seedExpectedDate').value || null;
        const seedReceived = document.getElementById('seedReceivedDate').value || null;
        const apicExpected = document.getElementById('apicExpectedDate').value || null;
        const apicReceived = document.getElementById('apicReceivedDate').value || null;
        const autoTx = document.getElementById('equityAutoTransaction').checked;

        const config = Database.getEquityConfig();
        const prevSeedReceived = config.seed_received_date || null;
        const prevApicReceived = config.apic_received_date || null;

        config.common_stock_par = par;
        config.common_stock_shares = shares;
        config.apic = apic;
        config.seed_expected_date = seedExpected;
        config.seed_received_date = seedReceived;
        config.apic_expected_date = apicExpected;
        config.apic_received_date = apicReceived;
        Database.setEquityConfig(config);

        // Auto-create journal entries for newly received amounts
        if (autoTx) {
            const seedAmount = Math.round(par * shares * 100) / 100;
            if (seedReceived && !prevSeedReceived && seedAmount > 0) {
                this._autoCreateEquityTransaction('Seed Money (Common Stock)', seedAmount, seedExpected || seedReceived, seedReceived);
            }
            if (apicReceived && !prevApicReceived && apic > 0) {
                this._autoCreateEquityTransaction('Additional Paid-In Capital', apic, apicExpected || apicReceived, apicReceived);
            }
        }

        UI.hideModal('equityModal');
        UI.showNotification('Equity saved', 'success');
        this.refreshFixedAssets();
    },

    _autoCreateEquityTransaction(description, amount, expectedDate, receivedDate) {
        // Find or create an "Equity Investment" category (hidden from P&L)
        let categories = Database.getCategories();
        let cat = categories.find(c => c.name === 'Equity Investment');
        let catId;
        if (!cat) {
            catId = Database.addCategory('Equity Investment', false, null, 'receivable', null, true, false, false, false);
        } else {
            catId = cat.id;
        }

        const monthDue = expectedDate.substring(0, 7);
        const monthPaid = receivedDate.substring(0, 7);

        Database.addTransaction({
            entry_date: expectedDate,
            category_id: catId,
            item_description: description,
            amount: amount,
            transaction_type: 'receivable',
            status: 'received',
            month_due: monthDue,
            month_paid: monthPaid,
            date_processed: receivedDate,
            source_type: 'investment',
            source_id: null
        });
    },

    resyncEquityJournalEntries() {
        const config = Database.getEquityConfig();
        const seedAmount = Math.round(config.common_stock_par * config.common_stock_shares * 100) / 100;
        const apicVal = Math.round((config.apic || 0) * 100) / 100;

        // Check for existing equity investment transactions
        const existing = Database.db.exec(
            "SELECT item_description FROM transactions WHERE source_type = 'investment'"
        );
        const existingDescs = existing.length ? existing[0].values.map(r => r[0]) : [];

        let created = 0;

        if (seedAmount > 0 && config.seed_received_date && !existingDescs.includes('Seed Money (Common Stock)')) {
            this._autoCreateEquityTransaction('Seed Money (Common Stock)', seedAmount, config.seed_expected_date || config.seed_received_date, config.seed_received_date);
            created++;
        }

        if (apicVal > 0 && config.apic_received_date && !existingDescs.includes('Additional Paid-In Capital')) {
            this._autoCreateEquityTransaction('Additional Paid-In Capital', apicVal, config.apic_expected_date || config.apic_received_date, config.apic_received_date);
            created++;
        }

        if (created > 0) {
            UI.showNotification(`Created ${created} equity journal ${created === 1 ? 'entry' : 'entries'}`, 'success');
            this.refreshAll();
        } else {
            UI.showNotification('Equity journal entries already exist', 'info');
        }
    },

    // ==================== LOAN HANDLERS ====================

    openLoanConfigModal(editId) {
        document.getElementById('loanConfigForm').reset();
        document.getElementById('editingLoanId').value = '';

        if (editId) {
            const loan = Database.getLoanById(editId);
            if (loan) {
                document.getElementById('editingLoanId').value = loan.id;
                document.getElementById('loanName').value = loan.name;
                document.getElementById('loanPrincipal').value = loan.principal;
                document.getElementById('loanRate').value = loan.annual_rate;
                document.getElementById('loanTermMonths').value = loan.term_months;
                document.getElementById('loanPayments').value = loan.payments_per_year;
                document.getElementById('loanStartDate').value = loan.start_date;
                document.getElementById('loanFirstPaymentDate').value = loan.first_payment_date || '';
                document.getElementById('loanNotes').value = loan.notes || '';
                document.getElementById('loanConfigModalTitle').textContent = 'Edit Loan';
            }
        } else {
            document.getElementById('loanConfigModalTitle').textContent = 'Add Loan';
            document.getElementById('loanStartDate').value = Utils.getTodayDate();
        }

        UI.showModal('loanConfigModal');
        document.getElementById('loanName').focus();
    },

    handleSaveLoanConfig() {
        const name = document.getElementById('loanName').value.trim();
        const principal = parseFloat(document.getElementById('loanPrincipal').value);
        const rate = parseFloat(document.getElementById('loanRate').value);
        const termMonths = parseInt(document.getElementById('loanTermMonths').value);
        const payments = parseInt(document.getElementById('loanPayments').value);
        const startDate = document.getElementById('loanStartDate').value;
        const firstPaymentDate = document.getElementById('loanFirstPaymentDate').value || null;
        const notes = document.getElementById('loanNotes').value.trim();
        const editingId = document.getElementById('editingLoanId').value;

        if (!name || isNaN(principal) || isNaN(rate) || isNaN(termMonths) || isNaN(payments) || !startDate) {
            UI.showNotification('Please fill in all required fields', 'error');
            return;
        }

        const params = { name, principal, annual_rate: rate, term_months: termMonths, payments_per_year: payments, start_date: startDate, first_payment_date: firstPaymentDate, notes };

        if (editingId) {
            Database.updateLoan(parseInt(editingId), params);
            this._syncLoanReceivable(parseInt(editingId), params);
            this._syncLoanBudgetExpense(parseInt(editingId), params);
            UI.showNotification('Loan updated', 'success');
        } else {
            const loanId = Database.addLoan(params);
            this.selectedLoanId = loanId;
            this._syncLoanReceivable(loanId, params);
            this._autoCreateLoanBudgetAndCategory(loanId, name, params);
            UI.showNotification('Loan added', 'success');
        }

        UI.hideModal('loanConfigModal');
        this.refreshAll();
    },

    handleDeleteLoan(id) {
        this.deleteLoanTargetId = id;
        const loan = Database.getLoanById(id);
        document.getElementById('deleteLoanMessage').textContent =
            `Are you sure you want to delete "${loan ? loan.name : 'this loan'}"?`;
        UI.showModal('deleteLoanModal');
    },

    confirmDeleteLoan() {
        if (this.deleteLoanTargetId) {
            // Delete budget expenses and ALL their journal entries (pending + paid)
            const budgetExpenses = Database.getBudgetExpenses();
            const autoCreatedNote = `Auto-created from loan #${this.deleteLoanTargetId}`;
            budgetExpenses.forEach(be => {
                if (be.notes === autoCreatedNote) {
                    Database.db.run(
                        "DELETE FROM transactions WHERE source_type = 'budget' AND source_id = ?",
                        [be.id]
                    );
                    Database.deleteBudgetExpense(be.id);
                }
            });

            // Delete loan receivable journal entry
            Database.db.run(
                "DELETE FROM transactions WHERE source_type = 'loan_receivable' AND source_id = ?",
                [this.deleteLoanTargetId]
            );

            // Delete any old loan_payment journal entries (legacy, before budget-driven flow)
            Database.db.run(
                "DELETE FROM transactions WHERE source_type = 'loan_payment' AND source_id = ?",
                [this.deleteLoanTargetId]
            );

            Database.deleteLoan(this.deleteLoanTargetId);
            if (this.selectedLoanId === this.deleteLoanTargetId) {
                this.selectedLoanId = null;
            }
            UI.showNotification('Loan deleted', 'success');
            this.refreshAll();
        }
        UI.hideModal('deleteLoanModal');
        this.deleteLoanTargetId = null;
    },

    /**
     * Auto-create a budget expense and journal category for a new loan
     */
    _autoCreateLoanBudgetAndCategory(loanId, loanName, params) {
        const schedule = Utils.computeAmortizationSchedule(
            { principal: params.principal, annual_rate: params.annual_rate,
              payments_per_year: params.payments_per_year, term_months: params.term_months,
              start_date: params.start_date, first_payment_date: params.first_payment_date },
            new Set(), {}
        );
        const monthlyPayment = schedule.length > 0 ? schedule[0].payment : 0;
        if (monthlyPayment <= 0) return;

        // Find or create a category named after the loan
        let categories = Database.getCategories();
        let cat = categories.find(c => c.name === loanName);
        let catId;
        if (!cat) {
            catId = Database.addCategory(loanName, true, monthlyPayment, 'payable', null, true);
        } else {
            catId = cat.id;
        }

        const firstPayMonth = schedule[0].month;
        const lastPayMonth = schedule[schedule.length - 1].month;

        Database.addBudgetExpense(
            `${loanName} Payment`,
            monthlyPayment,
            firstPayMonth,
            lastPayMonth,
            catId,
            `Auto-created from loan #${loanId}`
        );
    },

    /**
     * Sync ALL loan-linked budget expenses on every refreshAll().
     * Ensures budget expenses always reflect current loan data.
     */
    _syncAllLoanBudgetExpenses() {
        const expenses = Database.getBudgetExpenses();
        expenses.forEach(exp => {
            if (!exp.notes) return;
            const match = exp.notes.match(/^Auto-created from loan #(\d+)$/);
            if (!match) return;
            const loanId = parseInt(match[1]);
            const loan = Database.getLoanById(loanId);
            if (!loan) return;
            this._syncLoanBudgetExpense(loanId, loan);
        });
    },

    /**
     * Sync the linked budget expense when a loan is edited.
     * Updates name, amount, date range, and category to match the loan.
     */
    _syncLoanBudgetExpense(loanId, params) {
        const sentinel = `Auto-created from loan #${loanId}`;
        const expenses = Database.getBudgetExpenses();
        const linked = expenses.find(e => e.notes === sentinel);
        if (!linked) return;

        const schedule = Utils.computeAmortizationSchedule(
            { principal: params.principal, annual_rate: params.annual_rate,
              payments_per_year: params.payments_per_year, term_months: params.term_months,
              start_date: params.start_date, first_payment_date: params.first_payment_date },
            new Set(), {}
        );
        if (schedule.length === 0) return;

        const monthlyPayment = schedule[0].payment;
        const firstPayMonth = schedule[0].month;
        const lastPayMonth = schedule[schedule.length - 1].month;

        // Update or create the category to match the (possibly renamed) loan
        let categories = Database.getCategories();
        let cat = categories.find(c => c.id === linked.category_id);
        let catId = linked.category_id;
        if (cat && cat.name !== params.name) {
            Database.updateCategory(cat.id, params.name, cat.is_monthly, cat.default_amount, cat.transaction_type, cat.folder_id, cat.show_on_pl, cat.is_cogs, cat.is_depreciation, cat.is_sales_tax, cat.is_b2b, cat.default_status, cat.is_sales, cat.is_inventory_cost);
        } else if (!cat) {
            catId = Database.addCategory(params.name, true, monthlyPayment, 'payable', null, true);
        }

        Database.updateBudgetExpense(
            linked.id,
            `${params.name} Payment`,
            monthlyPayment,
            firstPayMonth,
            lastPayMonth,
            catId,
            sentinel,
            linked.group_id
        );
    },

    /**
     * Create the loan receivable journal entry (one-time, on loan add/edit).
     * Payment entries are handled entirely by the budget auto-sync.
     */
    _syncLoanReceivable(loanId, params) {
        const { name, principal, start_date, first_payment_date } = params;

        // --- Category: Loan Proceeds (receivable, hidden from P&L) ---
        let categories = Database.getCategories();
        let proceedsCat = categories.find(c => c.name === 'Loan Proceeds' && c.default_type === 'receivable');
        let proceedsCatId;
        if (!proceedsCat) {
            proceedsCatId = Database.addCategory('Loan Proceeds', false, null, 'receivable', null, true);
        } else {
            proceedsCatId = proceedsCat.id;
            if (!proceedsCat.show_on_pl) {
                Database.db.run('UPDATE categories SET show_on_pl = 1 WHERE id = ?', [proceedsCatId]);
            }
        }

        // --- Upsert receivable ---
        const existingReceivable = Database.db.exec(
            "SELECT id FROM transactions WHERE source_type = 'loan_receivable' AND source_id = ?",
            [loanId]
        );
        const fpd = first_payment_date || start_date;
        const dueMonth = fpd.substring(0, 7);
        const paidMonth = start_date.substring(0, 7);
        if (existingReceivable.length > 0 && existingReceivable[0].values.length > 0) {
            const rxId = existingReceivable[0].values[0][0];
            Database.db.run(
                'UPDATE transactions SET amount = ?, entry_date = ?, month_due = ?, month_paid = ?, category_id = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
                [principal, start_date, dueMonth, paidMonth, proceedsCatId, rxId]
            );
        } else {
            Database.addTransaction({
                entry_date: start_date,
                category_id: proceedsCatId,
                item_description: `${name} \u2013 Loan Proceeds`,
                amount: principal,
                transaction_type: 'receivable',
                status: 'received',
                date_processed: start_date,
                month_due: dueMonth,
                month_paid: paidMonth,
                source_type: 'loan_receivable',
                source_id: loanId
            });
        }

        Database.autoSave();
    },

    // ==================== BUDGET AUTO-SYNC ====================

    /**
     * Sync budget expenses to journal entries for all months up to the current month.
     * Similar to syncAllLoanJournalEntries — idempotent, runs on every refreshAll().
     */
    syncAllBudgetJournalEntries() {
        const expenses = Database.getBudgetExpenses();
        const currentMonth = Utils.getCurrentMonth();

        // Get all existing budget transactions in one query
        const existingResult = Database.db.exec(
            "SELECT id, source_id, payment_for_month, amount, status FROM transactions WHERE source_type = 'budget'"
        );
        const existingByExpenseMonth = {};
        if (existingResult.length > 0) {
            existingResult[0].values.forEach(([id, sourceId, month, amount, status]) => {
                const key = `${sourceId}:${month}`;
                existingByExpenseMonth[key] = { id, amount, status };
            });
        }

        // Build loan schedule lookup for loan-linked budget expenses
        const loanScheduleByMonth = {};
        expenses.forEach(exp => {
            if (!exp.notes) return;
            const loanMatch = exp.notes.match(/^Auto-created from loan #(\d+)$/);
            if (!loanMatch) return;
            const loanId = parseInt(loanMatch[1]);
            const loan = Database.getLoanById(loanId);
            if (!loan) return;
            const skippedPayments = Database.getSkippedPayments(loanId);
            const paymentOverrides = Database.getLoanPaymentOverrides(loanId);
            const schedule = Utils.computeAmortizationSchedule({
                principal: loan.principal, annual_rate: loan.annual_rate,
                payments_per_year: loan.payments_per_year, term_months: loan.term_months,
                start_date: loan.start_date, first_payment_date: loan.first_payment_date
            }, skippedPayments, paymentOverrides);
            const byMonth = {};
            schedule.forEach(p => { byMonth[p.month] = p; });
            loanScheduleByMonth[exp.id] = byMonth;
        });

        // Load all budget expense overrides in one query
        const allOverrides = Database.getAllBudgetExpenseOverrides();

        expenses.forEach(exp => {
            // Skip expenses without a category — can't create journal entries without one
            if (!exp.category_id) return;

            // Determine the month range: start_month to min(end_month, currentMonth)
            const endMonth = exp.end_month && exp.end_month < currentMonth ? exp.end_month : currentMonth;
            if (exp.start_month > currentMonth) return; // Not active yet

            const months = Utils.generateMonthRange(exp.start_month, endMonth);
            const activeMonths = new Set(months);
            const loanSchedule = loanScheduleByMonth[exp.id];

            // Create missing entries, update pending entries with changed amounts
            months.forEach(month => {
                const key = `${exp.id}:${month}`;
                const existing = existingByExpenseMonth[key];

                // For loan-linked expenses, use schedule amounts and respect skips
                let amount = exp.monthly_amount;
                let isSkipped = false;
                if (loanSchedule) {
                    const entry = loanSchedule[month];
                    if (entry) {
                        if (entry.skipped) {
                            isSkipped = true;
                        } else {
                            amount = entry.payment;
                        }
                    }
                }

                // Apply per-month budget override if set
                const overrideKey = `${exp.id}:${month}`;
                if (allOverrides[overrideKey] !== undefined) {
                    amount = allOverrides[overrideKey];
                }

                if (isSkipped) {
                    // Remove existing entry for skipped payment months
                    if (existing && existing.status === 'pending') {
                        Database.db.run('DELETE FROM transactions WHERE id = ?', [existing.id]);
                    }
                    return;
                }

                if (!existing) {
                    Database.addTransaction({
                        entry_date: month + '-01',
                        category_id: exp.category_id,
                        item_description: exp.name,
                        amount: amount,
                        transaction_type: 'payable',
                        status: 'pending',
                        month_due: month,
                        payment_for_month: month,
                        source_type: 'budget',
                        source_id: exp.id
                    });
                } else if (existing.amount !== amount) {
                    // Update amount to match budget (including overrides) regardless of status
                    Database.db.run(
                        'UPDATE transactions SET amount = ?, item_description = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
                        [amount, exp.name, existing.id]
                    );
                }
            });

            // Remove pending entries for months no longer in range (end_month was shortened)
            Object.entries(existingByExpenseMonth).forEach(([key, tx]) => {
                const [sourceId, month] = key.split(':');
                if (parseInt(sourceId) === exp.id && !activeMonths.has(month) && tx.status === 'pending') {
                    Database.db.run('DELETE FROM transactions WHERE id = ?', [tx.id]);
                }
            });
        });

        // Also clean up entries for deleted budget expenses
        const expenseIds = new Set(expenses.map(e => e.id));
        Object.entries(existingByExpenseMonth).forEach(([key, tx]) => {
            const sourceId = parseInt(key.split(':')[0]);
            if (!expenseIds.has(sourceId) && tx.status === 'pending') {
                Database.db.run('DELETE FROM transactions WHERE id = ?', [tx.id]);
            }
        });

        Database.autoSave();
    },

    // ==================== BUDGET HANDLERS ====================

    refreshBudget() {
        const expenses = Database.getBudgetExpenses();
        const groups = Database.getBudgetGroups();
        UI.renderBudgetTab(expenses, groups, this.selectedBudgetExpenseId, this.collapsedBudgetGroups);
    },

    openBudgetExpenseModal(editId) {
        document.getElementById('budgetExpenseForm').reset();
        document.getElementById('editingBudgetExpenseId').value = '';
        document.getElementById('budgetExpenseModalTitle').textContent = 'Add Budget Expense';
        document.getElementById('saveBudgetExpenseBtn').textContent = 'Add Expense';

        // Re-enable all fields (may have been disabled for loan-linked expense)
        ['budgetExpenseAmount', 'budgetStartMonth', 'budgetStartYear', 'budgetEndMonth', 'budgetEndYear', 'budgetExpenseNotes'].forEach(fieldId => {
            document.getElementById(fieldId).disabled = false;
        });
        const notice = document.getElementById('loanLinkNotice');
        if (notice) notice.style.display = 'none';

        // Populate category dropdown
        const catSelect = document.getElementById('budgetExpenseCategory');
        catSelect.innerHTML = '<option value="">None (won\'t record)</option>';
        const categories = Database.getCategories();
        categories.forEach(cat => {
            const opt = document.createElement('option');
            opt.value = cat.id;
            opt.textContent = cat.folder_name ? `${cat.folder_name} / ${cat.name}` : cat.name;
            catSelect.appendChild(opt);
        });

        // Populate group dropdown
        const groupSelect = document.getElementById('budgetExpenseGroup');
        groupSelect.innerHTML = '<option value="">No Group</option>';
        Database.getBudgetGroups().forEach(g => {
            const opt = document.createElement('option');
            opt.value = g.id;
            opt.textContent = g.name;
            groupSelect.appendChild(opt);
        });

        // Populate year dropdowns
        this._populateBudgetYearDropdowns();

        if (!editId) {
            // Auto-fill start/end from timeline settings
            const timeline = Database.getTimeline();
            if (timeline.start) {
                const [startYear, startMonth] = timeline.start.split('-');
                document.getElementById('budgetStartMonth').value = startMonth;
                document.getElementById('budgetStartYear').value = startYear;
            } else {
                const [year, month] = Utils.getCurrentMonth().split('-');
                document.getElementById('budgetStartMonth').value = month;
                document.getElementById('budgetStartYear').value = year;
            }
            if (timeline.end) {
                const [endYear, endMonth] = timeline.end.split('-');
                document.getElementById('budgetEndMonth').value = endMonth;
                document.getElementById('budgetEndYear').value = endYear;
            }
        }

        UI.showModal('budgetExpenseModal');
        document.getElementById('budgetExpenseName').focus();
    },

    _populateBudgetYearDropdowns() {
        const yearOptions = Utils.generateYearOptions();
        const placeholders = { budgetStartYear: 'Year...', budgetEndYear: '\u2014' };
        ['budgetStartYear', 'budgetEndYear'].forEach(id => {
            const sel = document.getElementById(id);
            const currentVal = sel.value;
            sel.innerHTML = `<option value="">${placeholders[id]}</option>`;
            yearOptions.forEach(y => {
                const opt = document.createElement('option');
                opt.value = y;
                opt.textContent = y;
                sel.appendChild(opt);
            });
            if (currentVal) sel.value = currentVal;
        });
    },

    handleSaveBudgetExpense() {
        const name = document.getElementById('budgetExpenseName').value.trim();
        const amount = parseFloat(document.getElementById('budgetExpenseAmount').value);
        const startMonth = document.getElementById('budgetStartMonth').value;
        const startYear = document.getElementById('budgetStartYear').value;
        const endMonth = document.getElementById('budgetEndMonth').value;
        const endYear = document.getElementById('budgetEndYear').value;
        const categoryId = document.getElementById('budgetExpenseCategory').value || null;
        const groupId = document.getElementById('budgetExpenseGroup').value || null;
        const notes = document.getElementById('budgetExpenseNotes').value.trim() || null;
        const editingId = document.getElementById('editingBudgetExpenseId').value;

        if (!name || isNaN(amount) || amount <= 0) {
            UI.showNotification('Please enter a name and valid amount', 'error');
            return;
        }
        if (!startMonth || !startYear) {
            UI.showNotification('Please select a start month', 'error');
            return;
        }

        const start = `${startYear}-${startMonth}`;

        if ((endMonth && !endYear) || (!endMonth && endYear)) {
            UI.showNotification('Please select both an end month and year, or leave both blank', 'error');
            return;
        }

        const end = (endMonth && endYear) ? `${endYear}-${endMonth}` : null;

        if (end && end < start) {
            UI.showNotification('End month cannot be before start month', 'error');
            return;
        }

        let didSyncLoan = false;
        try {
            if (editingId) {
                // For loan-linked expenses: preserve sentinel notes, use loan-controlled values for locked fields
                const existing = Database.getBudgetExpenseById(parseInt(editingId));
                const isLoanLinked = existing && existing.notes && /^Auto-created from loan #(\d+)$/.test(existing.notes);
                let finalNotes = notes;
                let finalAmount = amount;
                let finalStart = start;
                let finalEnd = end;
                if (isLoanLinked) {
                    finalNotes = existing.notes; // preserve sentinel
                    finalAmount = existing.monthly_amount; // controlled by loan
                    finalStart = existing.start_month;
                    finalEnd = existing.end_month;

                    // Sync name change back to the loan
                    const loanId = parseInt(existing.notes.match(/^Auto-created from loan #(\d+)$/)[1]);
                    const loan = Database.getLoanById(loanId);
                    if (loan) {
                        const newLoanName = name.replace(/ Payment$/, '');
                        if (newLoanName !== loan.name) {
                            Database.updateLoan(loanId, { ...loan, name: newLoanName });
                            // Update the loan's category name too
                            if (existing.category_id) {
                                const cat = Database.getCategories().find(c => c.id === existing.category_id);
                                if (cat) Database.updateCategory(cat.id, newLoanName, cat.is_monthly, cat.default_amount, cat.transaction_type, cat.folder_id, cat.show_on_pl, cat.is_cogs, cat.is_depreciation, cat.is_sales_tax, cat.is_b2b, cat.default_status, cat.is_sales, cat.is_inventory_cost);
                            }
                            didSyncLoan = true;
                        }
                    }
                }
                Database.updateBudgetExpense(parseInt(editingId), name, finalAmount, finalStart, finalEnd, categoryId ? parseInt(categoryId) : null, finalNotes, groupId ? parseInt(groupId) : null);
                UI.showNotification('Expense updated', 'success');
            } else {
                const id = Database.addBudgetExpense(name, amount, start, end, categoryId ? parseInt(categoryId) : null, notes, groupId ? parseInt(groupId) : null);
                this.selectedBudgetExpenseId = id;
                UI.showNotification('Expense added', 'success');
            }
            UI.hideModal('budgetExpenseModal');
            if (didSyncLoan) {
                this.refreshAll();
            } else {
                this.refreshBudget();
            }
        } catch (error) {
            console.error('Error saving budget expense:', error);
            UI.showNotification('Failed to save expense', 'error');
        }
    },

    handleEditBudgetExpense(id) {
        const expense = Database.getBudgetExpenseById(id);
        if (!expense) return;

        this.openBudgetExpenseModal(id);
        document.getElementById('editingBudgetExpenseId').value = expense.id;
        document.getElementById('budgetExpenseName').value = expense.name;
        document.getElementById('budgetExpenseAmount').value = expense.monthly_amount;
        document.getElementById('budgetExpenseCategory').value = expense.category_id || '';
        document.getElementById('budgetExpenseGroup').value = expense.group_id || '';
        document.getElementById('budgetExpenseNotes').value = expense.notes || '';
        document.getElementById('budgetExpenseModalTitle').textContent = 'Edit Budget Expense';
        document.getElementById('saveBudgetExpenseBtn').textContent = 'Save Changes';

        const [startYear, startMo] = expense.start_month.split('-');
        document.getElementById('budgetStartMonth').value = startMo;
        document.getElementById('budgetStartYear').value = startYear;

        if (expense.end_month) {
            const [endYear, endMo] = expense.end_month.split('-');
            document.getElementById('budgetEndMonth').value = endMo;
            document.getElementById('budgetEndYear').value = endYear;
        }

        // Lock loan-controlled fields if this expense is linked to a loan
        const isLoanLinked = expense.notes && /^Auto-created from loan #\d+$/.test(expense.notes);
        const loanControlledFields = ['budgetExpenseAmount', 'budgetStartMonth', 'budgetStartYear', 'budgetEndMonth', 'budgetEndYear', 'budgetExpenseNotes'];
        loanControlledFields.forEach(fieldId => {
            document.getElementById(fieldId).disabled = isLoanLinked;
        });
        // Show/hide loan link notice
        let notice = document.getElementById('loanLinkNotice');
        if (!notice) {
            notice = document.createElement('div');
            notice.id = 'loanLinkNotice';
            notice.style.cssText = 'padding:8px 12px;background:var(--accent-color,#4f46e5);color:#fff;border-radius:6px;margin-bottom:12px;font-size:0.85rem;';
            document.getElementById('budgetExpenseForm').prepend(notice);
        }
        if (isLoanLinked) {
            const loanId = expense.notes.match(/^Auto-created from loan #(\d+)$/)[1];
            const loan = Database.getLoanById(parseInt(loanId));
            notice.textContent = `Linked to loan "${loan ? loan.name : '#' + loanId}". Amount and dates are controlled by the loan. Edit the loan to change them.`;
            notice.style.display = '';
        } else {
            notice.style.display = 'none';
        }
    },

    handleDeleteBudgetExpense(id) {
        const expense = Database.getBudgetExpenseById(id);
        if (expense && expense.notes && /^Auto-created from loan #(\d+)$/.test(expense.notes)) {
            const loanId = parseInt(expense.notes.match(/^Auto-created from loan #(\d+)$/)[1]);
            const loan = Database.getLoanById(loanId);
            UI.showNotification(`This expense is linked to loan "${loan ? loan.name : '#' + loanId}". Delete the loan instead to remove both.`, 'warning');
            return;
        }
        this.deleteBudgetExpenseTargetId = id;
        UI.showModal('deleteBudgetExpenseModal');
    },

    confirmDeleteBudgetExpense() {
        if (this.deleteBudgetExpenseTargetId) {
            // Delete pending journal entries created by budget auto-sync for this expense
            Database.db.run(
                "DELETE FROM transactions WHERE source_type = 'budget' AND source_id = ? AND status = 'pending'",
                [this.deleteBudgetExpenseTargetId]
            );
            Database.deleteBudgetExpense(this.deleteBudgetExpenseTargetId);
            if (this.selectedBudgetExpenseId === this.deleteBudgetExpenseTargetId) {
                this.selectedBudgetExpenseId = null;
            }
            UI.showNotification('Expense deleted', 'success');
            this.refreshAll();
        }
        UI.hideModal('deleteBudgetExpenseModal');
        this.deleteBudgetExpenseTargetId = null;
    },

    handleAddBudgetGroup() {
        const name = prompt('Enter group name:');
        if (!name || !name.trim()) return;
        try {
            Database.addBudgetGroup(name.trim());
            this.refreshBudget();
            UI.showNotification('Group created', 'success');
        } catch (e) {
            UI.showNotification('Failed to create group', 'error');
        }
    },

    handleRenameBudgetGroup(groupId) {
        const nameEl = document.querySelector(`.budget-group-name[data-group-id="${groupId}"]`);
        if (!nameEl || nameEl.querySelector('input')) return;
        const currentName = nameEl.textContent;
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'budget-group-rename-input';
        input.value = currentName;
        nameEl.textContent = '';
        nameEl.appendChild(input);
        input.focus();
        input.select();

        const save = () => {
            const newName = input.value.trim();
            if (newName && newName !== currentName) {
                Database.updateBudgetGroup(groupId, newName);
                UI.showNotification('Group renamed', 'success');
            }
            this.refreshBudget();
        };
        input.addEventListener('blur', save);
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
            if (e.key === 'Escape') { input.removeEventListener('blur', save); this.refreshBudget(); }
        });
    },

    handleDeleteBudgetGroup(groupId) {
        if (!confirm('Delete this group? Expenses in this group will become ungrouped.')) return;
        Database.deleteBudgetGroup(groupId);
        this.refreshBudget();
        UI.showNotification('Group deleted', 'success');
    },

    openRecordBudgetModal() {
        const expenses = Database.getBudgetExpenses();
        if (expenses.length === 0) {
            UI.showNotification('No budget expenses to record', 'error');
            return;
        }

        // Populate year dropdown
        const yearSel = document.getElementById('recordBudgetYear');
        const yearOptions = Utils.generateYearOptions();
        yearSel.innerHTML = '<option value="">Year...</option>';
        yearOptions.forEach(y => {
            const opt = document.createElement('option');
            opt.value = y;
            opt.textContent = y;
            yearSel.appendChild(opt);
        });

        // Set defaults to current month
        const [curYear, curMonth] = Utils.getCurrentMonth().split('-');
        document.getElementById('recordBudgetMonth').value = curMonth;
        yearSel.value = curYear;
        document.getElementById('recordBudgetDate').value = Utils.getTodayDate();
        document.getElementById('recordBudgetStatus').value = 'pending';
        document.getElementById('recordBudgetDateProcessedGroup').style.display = 'none';
        document.getElementById('recordBudgetDateProcessed').value = '';

        this.updateRecordBudgetPreview();
        UI.showModal('recordBudgetModal');
    },

    updateRecordBudgetPreview() {
        const month = document.getElementById('recordBudgetMonth').value;
        const year = document.getElementById('recordBudgetYear').value;
        const preview = document.getElementById('recordBudgetPreview');
        const confirmBtn = document.getElementById('confirmRecordBudgetBtn');

        if (!month || !year) {
            preview.innerHTML = '<p class="empty-state">Select a month to see active expenses.</p>';
            confirmBtn.disabled = true;
            return;
        }

        const targetMonth = `${year}-${month}`;
        const activeExpenses = Database.getActiveBudgetExpensesForMonth(targetMonth);

        if (activeExpenses.length === 0) {
            preview.innerHTML = '<p class="empty-state">No active expenses for this month.</p>';
            confirmBtn.disabled = true;
            return;
        }

        const withCategory = activeExpenses.filter(e => e.category_id);
        const withoutCategory = activeExpenses.filter(e => !e.category_id);

        let html = '<div class="bulk-preview-list">';
        html += `<div class="bulk-preview-header">${activeExpenses.length} active expense${activeExpenses.length > 1 ? 's' : ''} for ${Utils.formatMonthShort(targetMonth)}</div>`;

        activeExpenses.forEach(exp => {
            const catLabel = exp.category_name
                ? `<span class="type-badge type-payable">${Utils.escapeHtml(exp.category_name)}</span>`
                : '<span class="type-badge type-none">No category</span>';
            html += `<div class="bulk-preview-item">
                <span class="bulk-preview-name">${Utils.escapeHtml(exp.name)}</span>
                ${catLabel}
                <span class="bulk-preview-amount">${Utils.formatCurrency(exp.monthly_amount)}</span>
            </div>`;
        });

        const total = activeExpenses.reduce((sum, e) => sum + e.monthly_amount, 0);
        html += `<div class="bulk-preview-total">Total: ${Utils.formatCurrency(total)}</div>`;

        if (withoutCategory.length > 0) {
            html += `<div class="bulk-preview-warning">${withoutCategory.length} expense${withoutCategory.length > 1 ? 's' : ''} without a category will be skipped.</div>`;
        }

        html += '</div>';
        preview.innerHTML = html;
        confirmBtn.disabled = withCategory.length === 0;
    },

    confirmRecordBudget() {
        const month = document.getElementById('recordBudgetMonth').value;
        const year = document.getElementById('recordBudgetYear').value;
        const entryDate = document.getElementById('recordBudgetDate').value;
        const status = document.getElementById('recordBudgetStatus').value;
        const dateProcessed = document.getElementById('recordBudgetDateProcessed').value || null;

        if (!month || !year || !entryDate) {
            UI.showNotification('Please fill in all required fields', 'error');
            return;
        }

        const targetMonth = `${year}-${month}`;
        const activeExpenses = Database.getActiveBudgetExpensesForMonth(targetMonth);
        const recordable = activeExpenses.filter(e => e.category_id);

        if (recordable.length === 0) {
            UI.showNotification('No expenses with categories to record', 'error');
            return;
        }

        let count = 0;
        try {
            recordable.forEach(exp => {
                // Check if auto-sync already created an entry for this expense+month
                const existingResult = Database.db.exec(
                    "SELECT id, status FROM transactions WHERE source_type = 'budget' AND source_id = ? AND payment_for_month = ?",
                    [exp.id, targetMonth]
                );
                if (existingResult.length > 0 && existingResult[0].values.length > 0) {
                    const [existingId, existingStatus] = existingResult[0].values[0];
                    // Update the existing entry with the user's chosen status/date
                    Database.db.run(
                        'UPDATE transactions SET entry_date = ?, status = ?, date_processed = ?, month_paid = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
                        [entryDate, status, status !== 'pending' ? dateProcessed : null, status !== 'pending' ? entryDate.substring(0, 7) : null, existingId]
                    );
                } else {
                    Database.addTransaction({
                        entry_date: entryDate,
                        category_id: exp.category_id,
                        item_description: exp.name,
                        amount: exp.monthly_amount,
                        transaction_type: 'payable',
                        status: status,
                        date_processed: status !== 'pending' ? dateProcessed : null,
                        month_due: targetMonth,
                        month_paid: status !== 'pending' ? entryDate.substring(0, 7) : null,
                        payment_for_month: targetMonth,
                        notes: null,
                        source_type: 'budget',
                        source_id: exp.id
                    });
                }
                count++;
            });

            UI.hideModal('recordBudgetModal');
            this.refreshAll();
            UI.showNotification(`${count} budget entr${count === 1 ? 'y' : 'ies'} recorded to journal`, 'success');
        } catch (error) {
            console.error('Error recording budget entries:', error);
            UI.showNotification('Failed to record entries', 'error');
        }
    },

    // ==================== CALCULATOR SIDEBAR ====================

    toggleCalcMode() {
        if (this.calcMode) {
            this.exitCalcMode();
        } else {
            this.enterCalcMode();
        }
    },

    enterCalcMode() {
        this.calcMode = true;
        this.calcRefCounter = 0;
        const btn = document.getElementById('calcToggleBtn');
        const sidebar = document.getElementById('calcSidebar');
        const appMain = document.querySelector('.app-main');
        const appContainer = document.querySelector('.app-container');

        btn.classList.add('active');
        appContainer.classList.add('calc-mode');
        appMain.classList.add('calc-sidebar-active');
        sidebar.style.display = 'flex';
        requestAnimationFrame(() => {
            sidebar.classList.add('calc-open');
        });

        // Auto-collapse left sidebar on narrow screens
        if (window.matchMedia('(max-width: 1024px)').matches) {
            appContainer.classList.remove('sidebar-open');
        }

        this.calcClear();
    },

    exitCalcMode() {
        if (!this.calcMode) return;
        this.calcMode = false;
        const btn = document.getElementById('calcToggleBtn');
        const sidebar = document.getElementById('calcSidebar');
        const appMain = document.querySelector('.app-main');
        const appContainer = document.querySelector('.app-container');

        btn.classList.remove('active');
        appContainer.classList.remove('calc-mode');
        appMain.classList.remove('calc-sidebar-active');
        sidebar.classList.remove('calc-open');

        // Remove all cell highlights
        document.querySelectorAll('.calc-cell-selected').forEach(el => el.classList.remove('calc-cell-selected'));

        setTimeout(() => {
            sidebar.style.display = '';
        }, 250);
    },

    handleCalcCellClick(cell, e) {
        // Extract numeric value
        let rawText = cell.textContent.replace(/[^0-9.\-]/g, '');
        const rawValue = parseFloat(rawText) || 0;

        // Determine label from cell context
        const label = this._calcCellLabel(cell);

        // Determine operation: Ctrl = subtract, Shift = multiply, Alt = divide
        let op = '+';
        if (e.ctrlKey || e.metaKey) op = '-';
        else if (e.shiftKey) op = '*';
        else if (e.altKey) op = '/';

        const refId = ++this.calcRefCounter;
        const formulaBar = document.getElementById('calcFormulaBar');

        // If formula bar already has content, insert operator before new value
        const existingText = formulaBar.textContent.trim();
        if (existingText.length > 0) {
            const opSymbol = { '+': ' + ', '-': ' - ', '*': ' × ', '/': ' ÷ ' }[op];
            formulaBar.appendChild(document.createTextNode(opSymbol));
        }

        // Create the cell reference span
        const span = document.createElement('span');
        span.className = 'calc-cell-ref';
        span.contentEditable = 'false';
        span.dataset.refId = refId;
        span.dataset.rawValue = rawValue;
        span.dataset.op = op;
        span.title = label;
        span.textContent = Utils.formatCurrency(rawValue);
        formulaBar.appendChild(span);

        // Highlight the cell on the spreadsheet
        cell.classList.add('calc-cell-selected');
        cell.dataset.calcRefId = refId;

        // Add to reference list
        this._calcAddRefItem(refId, rawValue, label);

        this.calcRecalculate();
    },

    _calcCellLabel(cell) {
        // P&L or Cash Flow cells
        if (cell.classList.contains('pnl-calc-cell') || cell.classList.contains('cf-calc-cell')) {
            const month = cell.dataset.month || '';
            // Check for summary labels
            if (cell.dataset.pnlLabel) return cell.dataset.pnlLabel + (month !== 'total' ? ' · ' + Utils.formatMonthShort(month) : ' · Total');
            if (cell.dataset.cfLabel) return cell.dataset.cfLabel + (month !== 'total' ? ' · ' + Utils.formatMonthShort(month) : ' · Total');
            // Category row — get name from first cell in the row
            const row = cell.closest('tr');
            const nameCell = row ? row.querySelector('td:first-child') : null;
            const catName = nameCell ? nameCell.textContent.trim() : 'Cell';
            const monthLabel = month === 'total' ? 'Total' : Utils.formatMonthShort(month);
            return catName + ' · ' + monthLabel;
        }
        // Balance Sheet
        if (cell.classList.contains('bs-value')) {
            const key = cell.dataset.bsKey || '';
            const row = cell.closest('tr');
            const nameCell = row ? row.querySelector('td:first-child') : null;
            return nameCell ? nameCell.textContent.trim() : key;
        }
        // Journal transactions
        if (cell.classList.contains('txn-amount')) {
            const row = cell.closest('tr');
            const catCell = row ? row.querySelector('td:nth-child(3)') : null;
            return catCell ? catCell.textContent.trim() : 'Transaction #' + (cell.dataset.txnId || '');
        }
        // Budget
        if (cell.classList.contains('budget-amount')) {
            const row = cell.closest('.budget-expense-row');
            const nameEl = row ? row.querySelector('.budget-expense-name') : null;
            return nameEl ? nameEl.textContent.trim() : 'Budget item';
        }
        // Loan
        if (cell.classList.contains('loan-amount')) {
            return 'Loan Payment #' + (cell.dataset.payment || '');
        }
        // Projected Sales
        if (cell.classList.contains('ps-amount')) {
            const row = cell.closest('tr');
            const nameCell = row ? row.querySelector('td:first-child') : null;
            const month = cell.dataset.psMonth || '';
            const label = nameCell ? nameCell.textContent.trim() : 'Projected';
            return label + (month !== 'total' ? ' · ' + Utils.formatMonthShort(month) : ' · Total');
        }
        return 'Cell';
    },

    _calcAddRefItem(refId, value, label) {
        const list = document.getElementById('calcRefList');
        const item = document.createElement('div');
        item.className = 'calc-ref-item';
        item.dataset.refId = refId;
        item.innerHTML = `<span class="calc-ref-value">${Utils.formatCurrency(value)}</span><span class="calc-ref-label">${Utils.escapeHtml(label)}</span>`;
        list.appendChild(item);

        // Update count
        const count = list.children.length;
        document.getElementById('calcCellCount').textContent = count + ' cell' + (count !== 1 ? 's' : '') + ' selected';
        document.getElementById('calcReferencesHeader').dataset.count = count;
    },

    calcRecalculate() {
        const formulaBar = document.getElementById('calcFormulaBar');
        const resultEl = document.getElementById('calcResult');

        // Walk DOM to build expression string
        let expr = '';
        formulaBar.childNodes.forEach(node => {
            if (node.nodeType === Node.TEXT_NODE) {
                expr += node.textContent;
            } else if (node.classList && node.classList.contains('calc-cell-ref')) {
                expr += node.dataset.rawValue || '0';
            } else {
                expr += node.textContent;
            }
        });

        if (!expr.trim()) {
            resultEl.textContent = '$0.00';
            resultEl.classList.remove('calc-error');
            return;
        }

        const result = Utils.evaluateExpression(expr);
        if (result.error) {
            resultEl.textContent = 'Error: ' + result.error;
            resultEl.classList.add('calc-error');
        } else {
            resultEl.textContent = Utils.formatCurrency(result.value);
            resultEl.classList.remove('calc-error');
        }
    },

    calcInsertManual(op) {
        const input = document.getElementById('calcManualValue');
        const formulaBar = document.getElementById('calcFormulaBar');
        const val = input.value.trim();

        if (val) {
            // If bar has content, insert operator first
            const existing = formulaBar.textContent.trim();
            if (existing.length > 0) {
                const opSymbol = { '+': ' + ', '-': ' - ', '*': ' × ', '/': ' ÷ ' }[op];
                formulaBar.appendChild(document.createTextNode(opSymbol));
            }
            formulaBar.appendChild(document.createTextNode(val));
            input.value = '';
        } else if (formulaBar.textContent.trim().length > 0) {
            // No value, just insert operator
            const opSymbol = { '+': ' + ', '-': ' - ', '*': ' × ', '/': ' ÷ ' }[op];
            formulaBar.appendChild(document.createTextNode(opSymbol));
        }

        this.calcRecalculate();
        formulaBar.focus();
    },

    calcClear() {
        document.getElementById('calcFormulaBar').innerHTML = '';
        document.getElementById('calcResult').textContent = '$0.00';
        document.getElementById('calcResult').classList.remove('calc-error');
        document.getElementById('calcRefList').innerHTML = '';
        document.getElementById('calcCellCount').textContent = '';
        document.getElementById('calcReferencesHeader').dataset.count = '0';
        document.querySelectorAll('.calc-cell-selected').forEach(el => el.classList.remove('calc-cell-selected'));
        this.calcRefCounter = 0;
    },

    // ==================== EXPORT ====================

    /**
     * Export all transactions as CSV
     */
    async handleExportCsv() {
        const transactions = Database.getTransactionsForExport();

        if (transactions.length === 0) {
            UI.showNotification('No transactions to export', 'error');
            return;
        }

        const csv = UI.generateCsv(transactions);
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });

        const owner = document.getElementById('journalOwner').value.trim();
        const prefix = owner ? Utils.sanitizeFilename(owner) : 'accounting_journal';
        const date = new Date().toISOString().split('T')[0];
        const filename = `${prefix}_export_${date}.csv`;

        await this.downloadBlob(blob, filename);
        UI.showNotification('CSV exported successfully', 'success');
    },

    // ==================== SAVE / LOAD ====================

    /**
     * Get the suggested filename for saving
     * @returns {string} Suggested filename
     */
    getSuggestedFilename() {
        const owner = document.getElementById('journalOwner').value.trim();
        if (owner) {
            return `${Utils.sanitizeFilename(owner)}_accounting_journal.db`;
        }
        return `accounting_journal_${new Date().toISOString().split('T')[0]}.db`;
    },

    /**
     * Handle save database (always prompts for location)
     */
    async handleSaveDatabase() {
        const owner = document.getElementById('journalOwner').value.trim();
        const suggestedName = owner
            ? `${Utils.sanitizeFilename(owner)}_accounting_journal.db`
            : `accounting_journal_${new Date().toISOString().split('T')[0]}.db`;

        if (window.showSaveFilePicker) {
            const blob = Database.exportToFile();
            await this.downloadBlob(blob, suggestedName);
            UI.showNotification('Database saved successfully', 'success');
        } else {
            // Fallback: show naming modal for browsers without file picker
            document.getElementById('saveAsName').value = suggestedName.replace(/\.db$/, '');
            UI.showModal('saveAsModal');
        }
    },

    /**
     * Handle save all — downloads a zip containing all company .db files
     */
    async handleSaveAllDatabases() {
        try {
            const companies = CompanyManager.getAll();
            if (!companies || companies.length === 0) {
                UI.showNotification('No companies to save', 'error');
                return;
            }

            // Save current company to IDB first so we get latest data
            const currentData = Database.db.export();
            await CompanyManager._writeIDB(CompanyManager._activeKey, new Uint8Array(currentData));

            const zip = new JSZip();

            for (const company of companies) {
                const bytes = await CompanyManager._readIDB(company.id);
                if (bytes) {
                    const filename = `${Utils.sanitizeFilename(company.name)}_accounting_journal.db`;
                    zip.file(filename, bytes);
                }
            }

            const blob = await zip.generateAsync({ type: 'blob' });
            const dateSuffix = new Date().toISOString().split('T')[0];
            const suggestedName = `all_companies_${dateSuffix}.zip`;

            this.downloadBlob(blob, suggestedName);
            UI.showNotification(`Saved ${companies.length} company database(s)`, 'success');
        } catch (e) {
            console.error('Save all failed:', e);
            UI.showNotification('Failed to save all databases', 'error');
        }
    },

    /**
     * Confirm save as from modal (fallback)
     */
    async confirmSaveAs() {
        const name = document.getElementById('saveAsName').value.trim();
        if (!name) {
            UI.showNotification('Please enter a file name', 'error');
            return;
        }

        const blob = Database.exportToFile();
        await this.downloadBlob(blob, `${Utils.sanitizeFilename(name)}.db`);
        UI.hideModal('saveAsModal');
        UI.showNotification('Database saved successfully', 'success');
    },

    // ==================== PRODUCTS HANDLERS ====================

    _pcStageFile(file) {
        this._pcStagedFile = file;
        document.getElementById('pcCsvFileName').textContent = file.name;
        document.getElementById('pcCsvZone').classList.add('has-file');
        document.getElementById('pcImportSubmit').disabled = false;
        document.getElementById('pcImportStatus').textContent = '';
    },

    _pcImportCsv() {
        const file = this._pcStagedFile;
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const text = e.target.result;
                const lines = text.split(/\r?\n/).filter(l => l.trim());
                if (lines.length < 2) throw new Error('File appears empty');

                // Parse header row
                const headers = lines[0].split(',').map(h => h.trim().toLowerCase().replace(/^"|"$/g, ''));
                const col = (name) => headers.indexOf(name);

                if (col('name') === -1) throw new Error('Missing required column: name');

                const nameIdx    = col('name');
                const skuIdx     = col('sku');
                const priceIdx   = col('price');
                const cogsIdx    = col('cogs');
                const taxIdx     = col('tax_rate');
                const notesIdx   = col('notes');

                const parseMoney = (v) => {
                    if (v === undefined || v === null || v === '') return 0;
                    return parseFloat(String(v).replace(/[$,%\s"]/g, '')) || 0;
                };

                // Get existing SKUs to avoid duplicates
                const existing = Database.getProducts();
                const existingSkus = new Set(existing.map(p => (p.sku || '').trim().toLowerCase()).filter(Boolean));

                let imported = 0, skipped = 0;
                for (let i = 1; i < lines.length; i++) {
                    const cols = lines[i].split(',').map(c => c.trim().replace(/^"|"$/g, ''));
                    const name = cols[nameIdx] || '';
                    if (!name) continue;

                    const sku = skuIdx !== -1 ? (cols[skuIdx] || '') : '';
                    if (sku && existingSkus.has(sku.toLowerCase())) { skipped++; continue; }

                    const price   = priceIdx   !== -1 ? parseMoney(cols[priceIdx])   : 0;
                    const cogs    = cogsIdx    !== -1 ? parseMoney(cols[cogsIdx])     : 0;
                    const taxRate = taxIdx     !== -1 ? parseMoney(cols[taxIdx])      : 0;
                    const notes   = notesIdx   !== -1 ? (cols[notesIdx] || '') : '';

                    Database.addProduct(name, sku || null, price, taxRate, cogs, notes || null);
                    if (sku) existingSkus.add(sku.toLowerCase());
                    imported++;
                }

                // Reset UI
                this._pcStagedFile = null;
                document.getElementById('pcCsvFileName').textContent = '';
                document.getElementById('pcCsvZone').classList.remove('has-file');
                document.getElementById('pcImportSubmit').disabled = true;
                document.getElementById('pcCsvInput').value = '';
                document.getElementById('pcImportPanel').classList.remove('open');

                const msg = skipped > 0
                    ? `Imported ${imported} product${imported !== 1 ? 's' : ''}. Skipped ${skipped} (duplicate SKU).`
                    : `Imported ${imported} product${imported !== 1 ? 's' : ''} successfully.`;
                UI.showNotification(msg, 'success');
                this.refreshProducts();
            } catch (err) {
                document.getElementById('pcImportStatus').textContent = 'Error: ' + err.message;
                UI.showNotification('Import failed: ' + err.message, 'error');
            }
        };
        reader.readAsText(file);
    },

    refreshProducts() {
        const showDiscontinued = document.getElementById('pcShowDiscontinued').checked;
        const products = Database.getProducts();
        const dateFrom = document.getElementById('pcDateFrom').value || null;
        const dateTo   = document.getElementById('pcDateTo').value   || null;
        const source   = document.getElementById('pcSourceFilter').value || 'all';
        const analytics = Database.getLinkedProductAnalytics(dateFrom, dateTo, source);
        analytics.linkedProductIds = new Set(analytics.byProduct.map(p => p.id));
        UI.renderProductsTab(products, showDiscontinued, analytics);
        this._destroyPcCharts();
        if (analytics.byProduct.length > 0) {
            this._renderPcChartsAndAnalytics(analytics);
        } else {
            document.getElementById('pcAnalyticsSection').style.display = 'none';
        }
    },

    openProductModal() {
        document.getElementById('productForm').reset();
        document.getElementById('editingProductId').value = '';
        document.getElementById('productModalTitle').textContent = 'Add Product';
        document.getElementById('saveProductBtn').textContent = 'Add Product';

        UI.showModal('productModal');
        document.getElementById('productName').focus();
    },

    handleSaveProduct() {
        const name = document.getElementById('productName').value.trim();
        const sku = document.getElementById('productSku').value.trim() || null;
        const price = parseFloat(document.getElementById('productPrice').value);
        const taxRate = 0;
        const cogs = parseFloat(document.getElementById('productCogs').value) || 0;
        const notes = document.getElementById('productNotes').value.trim() || null;
        const editingId = document.getElementById('editingProductId').value;

        if (!name || isNaN(price) || price < 0) {
            UI.showNotification('Please enter a product name and valid price', 'error');
            return;
        }

        try {
            if (editingId) {
                Database.updateProduct(parseInt(editingId), name, sku, price, taxRate, cogs, notes);
                UI.showNotification('Product updated', 'success');
            } else {
                Database.addProduct(name, sku, price, taxRate, cogs, notes);
                UI.showNotification('Product added', 'success');
            }
            UI.hideModal('productModal');
            this.refreshProducts();
        } catch (error) {
            console.error('Error saving product:', error);
            UI.showNotification('Failed to save product', 'error');
        }
    },

    handleEditProduct(id) {
        const product = Database.getProductById(id);
        if (!product) return;

        this.openProductModal(id);
        document.getElementById('editingProductId').value = product.id;
        document.getElementById('productName').value = product.name;
        document.getElementById('productSku').value = product.sku || '';
        document.getElementById('productPrice').value = product.price;
        document.getElementById('productCogs').value = product.cogs || '';
        document.getElementById('productNotes').value = product.notes || '';
        document.getElementById('productModalTitle').textContent = 'Edit Product';
        document.getElementById('saveProductBtn').textContent = 'Save Changes';
    },

    handleDeleteProduct(id) {
        this.deleteProductTargetId = id;
        UI.showModal('deleteProductModal');
    },

    confirmDeleteProduct() {
        if (this.deleteProductTargetId) {
            Database.deleteProduct(this.deleteProductTargetId);
            UI.showNotification('Product deleted', 'success');
            this.refreshProducts();
        }
        UI.hideModal('deleteProductModal');
        this.deleteProductTargetId = null;
    },

    handleToggleDiscontinued(id) {
        Database.toggleProductDiscontinued(id);
        this.refreshProducts();
    },

    // ==================== PRODUCT-VE MAPPING HANDLERS ====================

    openManageLinksModal(productId) {
        const product = Database.getProductById(productId);
        if (!product) return;
        const allItems = Database.getDistinctVeItemNamesWithPrice();
        const mappedArr = Database.getMappingsForProduct(productId);
        const mapped = new Set(mappedArr.map(m => m.name + '|||' + (m.price || 0).toFixed(2)));

        document.getElementById('pvmProductId').value = productId;
        document.getElementById('pvmProductName').textContent = product.name;
        document.getElementById('pvmProductPrice').textContent = Utils.formatCurrency(product.price);
        document.getElementById('pvmSearchInput').value = '';
        // Reset suggest button
        const suggestBtn = document.getElementById('pvmSuggestBtn');
        suggestBtn.classList.remove('active');
        suggestBtn.dataset.productPrice = product.price.toFixed(2);

        const list = document.getElementById('pvmItemList');
        if (allItems.length === 0) {
            list.innerHTML = '<p class="pvm-empty-msg">No VE item names found. Import VE Sales data first.</p>';
        } else {
            list.innerHTML = allItems.map(({ name, price }) => {
                const key = name + '|||' + (price || 0).toFixed(2);
                const checked = mapped.has(key) ? 'checked' : '';
                const esc = Utils.escapeHtml(name);
                const priceStr = (price != null) ? Utils.formatCurrency(price) : '—';
                return `<label class="pvm-item">
                    <input type="checkbox" class="pvm-check" value="${esc}|||${(price || 0).toFixed(2)}" ${checked}>
                    <span class="pvm-item-name">${esc}</span>
                    <span class="pvm-item-price">${priceStr}</span>
                </label>`;
            }).join('');
        }
        this._pvmUpdateSelectedCount();
        UI.showModal('pvmModal');
        document.getElementById('pvmSearchInput').focus();
    },

    handlePvmSearch(query) {
        const q = query.toLowerCase();
        document.querySelectorAll('#pvmItemList .pvm-item').forEach(item => {
            const name = item.querySelector('.pvm-item-name').textContent.toLowerCase();
            item.classList.toggle('pvm-item--hidden', q.length > 0 && !name.includes(q));
        });
    },

    handlePvmSuggest() {
        const btn = document.getElementById('pvmSuggestBtn');
        const isActive = btn.classList.toggle('active');
        const searchInput = document.getElementById('pvmSearchInput');

        if (isActive) {
            searchInput.value = '';
            const targetPrice = parseFloat(btn.dataset.productPrice) || 0;
            document.querySelectorAll('#pvmItemList .pvm-item').forEach(item => {
                const priceText = item.querySelector('.pvm-item-price').textContent.replace(/[^0-9.\-]/g, '');
                const itemPrice = parseFloat(priceText) || 0;
                item.classList.toggle('pvm-item--hidden', Math.abs(itemPrice - targetPrice) > 0.01);
            });
        } else {
            // Deactivate: show all
            document.querySelectorAll('#pvmItemList .pvm-item').forEach(item => {
                item.classList.remove('pvm-item--hidden');
            });
        }
    },

    _pvmUpdateSelectedCount() {
        const count = document.querySelectorAll('#pvmItemList .pvm-check:checked').length;
        document.getElementById('pvmSelectedCount').textContent = `${count} selected`;
    },

    handleSaveProductMappings() {
        const productId = parseInt(document.getElementById('pvmProductId').value);
        const checked = [...document.querySelectorAll('#pvmItemList .pvm-check:checked')]
            .map(cb => {
                const [name, price] = cb.value.split('|||');
                return { name, price: parseFloat(price) || 0 };
            });
        Database.setMappingsForProduct(productId, checked);
        UI.hideModal('pvmModal');
        UI.showNotification('Links saved', 'success');
        this.refreshProducts();
        // Refresh VE Sales if visible to update badges
        if (document.getElementById('vesalesTab').style.display !== 'none') {
            this.veRenderProducts();
        }
    },

    _destroyPcCharts() {
        if (this._pvmChartUnits)      { this._pvmChartUnits.destroy();      this._pvmChartUnits      = null; }
        if (this._pvmChartRevenue)    { this._pvmChartRevenue.destroy();    this._pvmChartRevenue    = null; }
        if (this._pvmBarUnits)        { this._pvmBarUnits.destroy();        this._pvmBarUnits        = null; }
        if (this._pvmBarRevenue)      { this._pvmBarRevenue.destroy();      this._pvmBarRevenue      = null; }
    },

    _renderPcChartsAndAnalytics(analytics) {
        document.getElementById('pcAnalyticsSection').style.display = 'block';

        // --- Cohesive color palette (blue-teal-green gradient + accents) ---
        const palette = [
            '#2d6a8a','#3a8f6e','#e6913e','#7c5cbf','#d14d4d',
            '#1fa5a5','#c49b1c','#4a7fd9','#b84f8a','#6a9e3a',
            '#e07850','#5878d8','#8eb04a','#d85090','#4a9eb0'
        ];
        const otherColor = '#b0b8c0';

        const TOP_N = 6;
        const truncate = (s, n = 30) => s.length > n ? s.slice(0, n - 1) + '\u2026' : s;

        // Full data for bar charts
        const allLabels  = analytics.byProduct.map(p => p.name);
        const allUnits   = analytics.byProduct.map(p => p.units_sold);
        const allRevenue = analytics.byProduct.map(p => p.revenue);
        const allColors  = allLabels.map((_, i) => palette[i % palette.length]);

        // --- Group into Top N + "Other" for donuts ---
        const sortedByUnits = analytics.byProduct.slice().sort((a, b) => b.units_sold - a.units_sold);
        const sortedByRev   = analytics.byProduct.slice().sort((a, b) => b.revenue - a.revenue);

        const buildGrouped = (sorted, valKey) => {
            const top   = sorted.slice(0, TOP_N);
            const rest  = sorted.slice(TOP_N);
            const labels = top.map(p => truncate(p.name));
            const data   = top.map(p => p[valKey]);
            const colors = top.map((_, i) => palette[i % palette.length]);
            const otherItems = rest.map(p => ({ name: p.name, value: p[valKey] }));
            if (rest.length > 0) {
                const otherTotal = rest.reduce((s, p) => s + p[valKey], 0);
                labels.push(`Other (${rest.length})`);
                data.push(otherTotal);
                colors.push(otherColor);
            }
            return { labels, data, colors, otherItems };
        };

        const gUnits = buildGrouped(sortedByUnits, 'units_sold');
        const gRev   = buildGrouped(sortedByRev, 'revenue');

        // --- Inline percentage label plugin ---
        const pctLabelPlugin = {
            id: 'pctLabels',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                const meta = chart.getDatasetMeta(0);
                const total = meta.total || chart.data.datasets[0].data.reduce((a, b) => a + b, 0);
                meta.data.forEach((arc, i) => {
                    const val = chart.data.datasets[0].data[i];
                    const pct = total > 0 ? (val / total * 100) : 0;
                    if (pct < 5) return; // skip tiny slices
                    const { x, y } = arc.tooltipPosition();
                    ctx.save();
                    ctx.fillStyle = '#fff';
                    ctx.font = 'bold 11px system-ui, sans-serif';
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'middle';
                    ctx.shadowColor = 'rgba(0,0,0,0.4)';
                    ctx.shadowBlur = 3;
                    ctx.fillText(pct.toFixed(0) + '%', x, y);
                    ctx.restore();
                });
            }
        };

        // --- Donut chart options builder ---
        const donutOpts = (otherItems, isRevenue) => ({
            responsive: true,
            maintainAspectRatio: true,
            cutout: '50%',
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: { boxWidth: 12, font: { size: 11 }, padding: 8 }
                },
                tooltip: {
                    callbacks: {
                        label: (ctx) => {
                            const val = ctx.raw;
                            const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                            const pct = total > 0 ? (val / total * 100).toFixed(1) : 0;
                            const formatted = isRevenue ? Utils.formatCurrency(val) : val.toLocaleString() + ' units';
                            return ` ${formatted} (${pct}%)`;
                        },
                        afterBody: (items) => {
                            const ctx = items[0];
                            if (!ctx || !ctx.label.startsWith('Other')) return '';
                            const lines = ['\n--- Breakdown ---'];
                            const sorted = otherItems.slice().sort((a, b) => b.value - a.value);
                            for (const item of sorted) {
                                const val = isRevenue ? Utils.formatCurrency(item.value) : item.value.toLocaleString() + ' units';
                                lines.push(`  ${truncate(item.name, 35)}: ${val}`);
                            }
                            return lines;
                        }
                    }
                }
            }
        });

        // --- Render donuts ---
        const uCanvas = document.getElementById('pcChartUnitsCanvas');
        const rCanvas = document.getElementById('pcChartRevenueCanvas');
        if (uCanvas && typeof Chart !== 'undefined') {
            this._pvmChartUnits = new Chart(uCanvas, {
                type: 'doughnut',
                data: { labels: gUnits.labels, datasets: [{ data: gUnits.data, backgroundColor: gUnits.colors, borderWidth: 2, borderColor: '#fff' }] },
                options: donutOpts(gUnits.otherItems, false),
                plugins: [pctLabelPlugin]
            });
        }
        if (rCanvas && typeof Chart !== 'undefined') {
            this._pvmChartRevenue = new Chart(rCanvas, {
                type: 'doughnut',
                data: { labels: gRev.labels, datasets: [{ data: gRev.data, backgroundColor: gRev.colors, borderWidth: 2, borderColor: '#fff' }] },
                options: donutOpts(gRev.otherItems, true),
                plugins: [pctLabelPlugin]
            });
        }

        // --- Horizontal bar charts (all products, sorted desc) ---
        const barSortedUnits = analytics.byProduct.slice().sort((a, b) => b.units_sold - a.units_sold);
        const barSortedRev   = analytics.byProduct.slice().sort((a, b) => b.revenue - a.revenue);

        const barOpts = (isRevenue) => ({
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: (ctx) => isRevenue ? ` ${Utils.formatCurrency(ctx.raw)}` : ` ${ctx.raw.toLocaleString()} units`
                    }
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    grid: { color: 'rgba(0,0,0,0.05)' },
                    ticks: {
                        callback: (v) => isRevenue ? '$' + (v >= 1000 ? (v / 1000).toFixed(0) + 'k' : v) : v.toLocaleString(),
                        font: { size: 10 }
                    }
                },
                y: {
                    grid: { display: false },
                    ticks: {
                        font: { size: 10 },
                        callback: function(val) {
                            const label = this.getLabelForValue(val);
                            return truncate(label, 28);
                        }
                    }
                }
            }
        });

        const barHeight = Math.max(200, analytics.byProduct.length * 28 + 40);

        const buCanvas = document.getElementById('pcBarUnitsCanvas');
        const brCanvas = document.getElementById('pcBarRevenueCanvas');
        if (buCanvas && typeof Chart !== 'undefined') {
            buCanvas.parentElement.style.height = barHeight + 'px';
            buCanvas.style.maxHeight = barHeight + 'px';
            this._pvmBarUnits = new Chart(buCanvas, {
                type: 'bar',
                data: {
                    labels: barSortedUnits.map(p => p.name),
                    datasets: [{
                        data: barSortedUnits.map(p => p.units_sold),
                        backgroundColor: barSortedUnits.map((_, i) => palette[i % palette.length]),
                        borderRadius: 3,
                        barThickness: 18
                    }]
                },
                options: barOpts(false)
            });
        }
        if (brCanvas && typeof Chart !== 'undefined') {
            brCanvas.parentElement.style.height = barHeight + 'px';
            brCanvas.style.maxHeight = barHeight + 'px';
            this._pvmBarRevenue = new Chart(brCanvas, {
                type: 'bar',
                data: {
                    labels: barSortedRev.map(p => p.name),
                    datasets: [{
                        data: barSortedRev.map(p => p.revenue),
                        backgroundColor: barSortedRev.map((_, i) => palette[i % palette.length]),
                        borderRadius: 3,
                        barThickness: 18
                    }]
                },
                options: barOpts(true)
            });
        }

        // Monthly COGS pivot table
        const rows     = analytics.monthlyCogs;
        const months   = [...new Set(rows.map(r => r.month))].sort();
        const products = analytics.byProduct;
        const fmt      = (n) => Utils.formatCurrency(n);

        let tHead = '<tr><th>Product</th>' + months.map(m => {
            const [yr, mo] = m.split('-');
            const label = new Date(parseInt(yr), parseInt(mo) - 1, 1)
                .toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
            return `<th>${label}</th>`;
        }).join('') + '<th>Total COGS</th></tr>';

        let tBody = '';
        const colTotals = new Array(months.length).fill(0);
        let grandTotal  = 0;

        for (const p of products) {
            let rowTotal = 0;
            const cells = months.map((m, mi) => {
                const hit  = rows.find(r => r.month === m && Number(r.product_id) === Number(p.id));
                const cogs = hit ? (hit.qty_sold * (p.cogs || 0)) : 0;
                colTotals[mi] += cogs;
                rowTotal += cogs;
                return `<td class="${cogs === 0 ? 'pc-cogs-zero' : ''}">${cogs > 0 ? fmt(cogs) : '—'}</td>`;
            }).join('');
            grandTotal += rowTotal;
            tBody += `<tr><td>${Utils.escapeHtml(p.name)}</td>${cells}<td>${fmt(rowTotal)}</td></tr>`;
        }
        const footCells = colTotals.map(t => `<td>${fmt(t)}</td>`).join('');
        tBody += `<tr><td>Total</td>${footCells}<td>${fmt(grandTotal)}</td></tr>`;

        document.getElementById('pcCogsTableWrapper').innerHTML =
            `<div class="pc-analytics-cogs-title">Monthly COGS Summary</div>
             <div class="pc-cogs-table-wrap">
               <table class="pc-cogs-table"><thead>${tHead}</thead><tbody>${tBody}</tbody></table>
             </div>`;
    },

    _pcApplyDatePreset() {
        const preset = document.getElementById('pcDatePreset').value;
        const fromEl = document.getElementById('pcDateFrom');
        const toEl   = document.getElementById('pcDateTo');
        const now    = new Date();
        switch (preset) {
            case 'all':
                fromEl.value = ''; toEl.value = '';
                break;
            case 'thisMonth': {
                const y = now.getFullYear(), m = String(now.getMonth() + 1).padStart(2, '0');
                fromEl.value = `${y}-${m}-01`;
                const lastDay = new Date(y, now.getMonth() + 1, 0).getDate();
                toEl.value = `${y}-${m}-${String(lastDay).padStart(2, '0')}`;
                break;
            }
            case 'lastMonth': {
                const last = new Date(now.getFullYear(), now.getMonth() - 1, 1);
                const y = last.getFullYear(), m = String(last.getMonth() + 1).padStart(2, '0');
                fromEl.value = `${y}-${m}-01`;
                const lastDay = new Date(last.getFullYear(), last.getMonth() + 1, 0).getDate();
                toEl.value = `${y}-${m}-${String(lastDay).padStart(2, '0')}`;
                break;
            }
            case 'thisQuarter': {
                const qMonth = Math.floor(now.getMonth() / 3) * 3;
                const y = now.getFullYear();
                fromEl.value = `${y}-${String(qMonth + 1).padStart(2, '0')}-01`;
                const lastQMonth = qMonth + 2;
                const lastDay = new Date(y, lastQMonth + 1, 0).getDate();
                toEl.value = `${y}-${String(lastQMonth + 1).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
                break;
            }
        }
        this.refreshProducts();
    },

    // ==================== BREAK-EVEN ANALYSIS ====================

    /**
     * Get the effective timeline for break-even (local override or global)
     * @returns {Object} {start, end}
     */
    getBreakevenTimeline(cfg) {
        cfg = cfg || Database.getBreakevenConfig();
        const local = cfg.timeline || {};
        if (local.start || local.end) return local;
        return this.getTimeline();
    },

    /**
     * Refresh the entire break-even tab
     */
    refreshBreakeven() {
        const cfg = Database.getBreakevenConfig();
        const timeline = this.getBreakevenTimeline(cfg);
        const realCurrentMonth = Utils.getCurrentMonth();
        const useProjected = cfg.dataSource === 'projected';
        // When projected mode has an as-of month, use it as the cutoff for actual vs projected
        const currentMonth = (useProjected && cfg.asOfMonth) ? cfg.asOfMonth : realCurrentMonth;

        // Populate snapshot month dropdown
        const beSnapshotSelect = document.getElementById('beSnapshotMonth');
        if (beSnapshotSelect) {
            let allTimelineMonths = [];
            if (timeline.start && timeline.end) {
                allTimelineMonths = Utils.generateMonthRange(timeline.start, timeline.end);
            } else {
                allTimelineMonths = this._getTimelineMonths();
            }
            const availableMonths = allTimelineMonths.filter(m => m <= realCurrentMonth);
            const prevVal = beSnapshotSelect.value;
            beSnapshotSelect.innerHTML = '<option value="">Current Month</option>' +
                availableMonths.map(m => `<option value="${m}">${Utils.formatMonthShort(m)}</option>`).join('');
            if (prevVal && availableMonths.includes(prevVal)) {
                beSnapshotSelect.value = prevVal;
            } else {
                beSnapshotSelect.value = '';
            }
        }
        const beSnapshotMonth = beSnapshotSelect ? beSnapshotSelect.value : '';

        // Timeline banner
        const banner = document.getElementById('beTimelineBanner');
        if (timeline.start || timeline.end) {
            const startLabel = timeline.start ? Utils.formatMonthShort(timeline.start) : 'Start';
            const endLabel = timeline.end ? Utils.formatMonthShort(timeline.end) : 'Present';
            const isLocal = (cfg.timeline && (cfg.timeline.start || cfg.timeline.end));
            const asOfLabel = (useProjected && cfg.asOfMonth) ? ` as of ${Utils.formatMonthShort(cfg.asOfMonth)}` : '';
            const snapshotLabel = beSnapshotMonth ? ` • Snapshot: ${Utils.formatMonthShort(beSnapshotMonth)}` : '';
            banner.textContent = `Timeline: ${startLabel} \u2013 ${endLabel}${isLocal ? ' (local override)' : ''} \u2022 ${useProjected ? 'Projected' : 'Actual'}${asOfLabel}${snapshotLabel}`;
            banner.style.display = 'block';
        } else {
            if (beSnapshotMonth) {
                banner.textContent = `Snapshot: ${Utils.formatMonthShort(beSnapshotMonth)}`;
                banner.style.display = 'block';
            } else {
                banner.style.display = 'none';
            }
        }

        // Compute monthly fixed costs
        let months = [];
        if (timeline.start && timeline.end) {
            months = Utils.generateMonthRange(timeline.start, timeline.end);
        } else {
            months = [currentMonth];
        }

        // Apply snapshot filter — only include months up to snapshot
        if (beSnapshotMonth) {
            months = months.filter(m => m <= beSnapshotMonth);
            if (months.length === 0) months = [beSnapshotMonth];
        }

        // Ensure P&L has been rendered so UI._pnlMonthOpex is available
        if (!UI._pnlMonthOpex) {
            this.refreshPnL();
        }
        const totalOpexByMonth = UI._pnlMonthOpex || {};

        // Determine monthly fixed costs — sum all months from P&L to match exactly
        let avgMonthlyFixed;
        let totalFixedFromPL;
        if (cfg.fixedCostOverride != null && cfg.fixedCostOverride > 0) {
            // Manual override takes priority
            avgMonthlyFixed = cfg.fixedCostOverride;
            totalFixedFromPL = avgMonthlyFixed * months.length;
        } else {
            // Sum all months (actual + projected) to match P&L total exactly
            totalFixedFromPL = months.reduce((acc, m) => acc + (totalOpexByMonth[m] || 0), 0);
            avgMonthlyFixed = months.length > 0 ? totalFixedFromPL / months.length : 0;
        }

        // Core break-even calculation
        const beResult = Utils.computeBreakEven(cfg, avgMonthlyFixed);

        // Render summary and channel breakdown
        UI.renderBreakevenSummaryCards(beResult, null, cfg);
        UI.renderBreakevenChannelBreakdown(beResult, cfg);

        // Chart data points
        this._destroyBeCharts();

        if (beResult.isValid) {
            const monthCount = months.length;
            const increment = cfg.unitIncrement || 100;
            const b2b = cfg.b2b || {};
            const b2bMonthly = (b2b.enabled && b2b.monthlyUnits > 0) ? b2b.monthlyUnits : 0;
            const b2bTotal = b2bMonthly * monthCount;

            const points = Utils.computeBreakEvenChartPoints(
                cfg, avgMonthlyFixed, beResult.consumerUnitsNeededExact || 0, increment, monthCount
            );
            // Use exact (non-ceiled) monthly value to avoid double-rounding
            const consumerBETotal = Math.ceil((beResult.consumerUnitsNeededExact || 0) * monthCount);

            // Compute exact break-even point for the table (not injected into chart)
            const consumer = cfg.consumer || {};
            const consumerPrice = consumer.enabled ? (consumer.avgPrice || 0) : 0;
            const consumerCogs = consumer.enabled ? (consumer.avgCogs || 0) : 0;
            const b2bRate = b2b.enabled ? (b2b.ratePerUnit || 0) : 0;
            const b2bCogs = b2b.enabled ? (b2b.cogsPerUnit || 0) : 0;
            const fixedTotal = Math.round(totalFixedFromPL * 100) / 100;
            const beRevenue = (b2bTotal * b2bRate) + (consumerBETotal * consumerPrice);
            const beVarCosts = (b2bTotal * b2bCogs) + (consumerBETotal * consumerCogs);
            const exactBEPoint = {
                consumerUnits: consumerBETotal,
                b2bUnits: b2bTotal,
                revenue: Math.round(beRevenue * 100) / 100,
                variableCosts: Math.round(beVarCosts * 100) / 100,
                fixedCosts: fixedTotal,
                totalCosts: Math.round((beVarCosts + fixedTotal) * 100) / 100
            };

            this._renderBeBreakevenChart(points);
            UI.renderBreakevenDataTable(points, b2bTotal, increment, consumerBETotal, exactBEPoint);

            // Timeline chart
            let timelinePoints = null;
            if (months.length > 1) {
                const hasOverride = cfg.fixedCostOverride != null && cfg.fixedCostOverride > 0;
                timelinePoints = Utils.computeBreakevenTimeline(
                    cfg, months, {}, {},
                    hasOverride
                        ? () => avgMonthlyFixed
                        : (m) => totalOpexByMonth[m] || 0
                );
                // Use P&L gross profit for the timeline chart
                const tlMonthList = timelinePoints.map(tp => tp.month);
                const actualProfitByMonth = Database.getMonthlyGrossProfit(tlMonthList);
                this._renderBeTimelineChart(timelinePoints, actualProfitByMonth);
            }

            // Progress tracker ("as of" analysis)
            this.refreshBreakevenProgress(beResult, cfg, months, timelinePoints);
        } else {
            UI.renderBreakevenDataTable([], 0, 100);
            document.getElementById('beProgressSection').style.display = 'none';
        }
    },

    /**
     * Destroy existing Chart.js instances
     */
    _destroyBeCharts() {
        if (this._beChartBreakeven) {
            this._beChartBreakeven.destroy();
            this._beChartBreakeven = null;
        }
        if (this._beChartTimeline) {
            this._beChartTimeline.destroy();
            this._beChartTimeline = null;
        }
        if (this._beChartProgress) {
            this._beChartProgress.destroy();
            this._beChartProgress = null;
        }
    },

    /**
     * Render the break-even chart (revenue vs total costs, consumer units on X axis)
     * @param {Array} points - Chart data points with consumerUnits field
     */
    _renderBeBreakevenChart(points) {
        const canvas = document.getElementById('beBreakevenChart');
        if (!canvas || typeof Chart === 'undefined') return;

        const primaryColor = getComputedStyle(document.documentElement).getPropertyValue('--color-primary').trim() || '#4a90a4';
        const dangerColor = '#dc3545';
        const successColor = '#28a745';
        const warningColor = '#f59e0b';
        const grayColor = '#6c757d';

        // Find break-even index: first point where revenue >= total costs
        const beIdx = points.findIndex(p => p.revenue >= p.totalCosts);

        // Build pointRadius and pointBackgroundColor arrays for break-even emphasis
        const defaultRadius = 2;
        const beRadius = 6;
        const revenueRadii = points.map((_, i) => i === beIdx ? beRadius : defaultRadius);
        const costRadii = points.map((_, i) => i === beIdx ? beRadius : defaultRadius);
        const revenueBgColors = points.map((_, i) => i === beIdx ? primaryColor : successColor);
        const costBgColors = points.map((_, i) => i === beIdx ? primaryColor : dangerColor);

        this._beChartBreakeven = new Chart(canvas, {
            type: 'line',
            data: {
                labels: points.map(p => p.consumerUnits.toLocaleString()),
                datasets: [
                    {
                        label: 'Revenue',
                        data: points.map(p => p.revenue),
                        borderColor: successColor,
                        backgroundColor: successColor + '20',
                        fill: false,
                        tension: 0.1,
                        pointRadius: revenueRadii,
                        pointBackgroundColor: revenueBgColors,
                        pointBorderColor: revenueBgColors
                    },
                    {
                        label: 'Total Costs',
                        data: points.map(p => p.totalCosts),
                        borderColor: dangerColor,
                        backgroundColor: dangerColor + '20',
                        fill: false,
                        tension: 0.1,
                        pointRadius: costRadii,
                        pointBackgroundColor: costBgColors,
                        pointBorderColor: costBgColors
                    },
                    {
                        label: 'Variable Costs',
                        data: points.map(p => p.variableCosts),
                        borderColor: warningColor,
                        borderDash: [4, 4],
                        fill: false,
                        tension: 0.1,
                        pointRadius: 0
                    },
                    {
                        label: 'Fixed Costs',
                        data: points.map(p => p.fixedCosts),
                        borderColor: grayColor,
                        borderDash: [5, 5],
                        fill: false,
                        tension: 0,
                        pointRadius: 0
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                scales: {
                    x: {
                        title: { display: true, text: 'Consumer Units Sold' },
                        grid: { display: false }
                    },
                    y: {
                        title: { display: true, text: 'Amount ($)' },
                        ticks: {
                            callback: (v) => '$' + v.toLocaleString()
                        }
                    }
                },
                plugins: {
                    tooltip: {
                        callbacks: {
                            label: (ctx) => `${ctx.dataset.label}: $${ctx.parsed.y.toLocaleString(undefined, { minimumFractionDigits: 2 })}`
                        }
                    }
                }
            }
        });
    },

    /**
     * Render the monthly timeline chart comparing actual gross profit to break-even target
     * @param {Array} timelinePoints - From Utils.computeBreakevenTimeline()
     * @param {Object} actualProfitByMonth - { 'YYYY-MM': grossProfit }
     */
    _renderBeTimelineChart(timelinePoints, actualProfitByMonth) {
        const canvas = document.getElementById('beTimelineChart');
        if (!canvas || typeof Chart === 'undefined') return;

        const primaryColor = getComputedStyle(document.documentElement).getPropertyValue('--color-primary').trim() || '#4a90a4';
        const successColor = '#28a745';

        this._beChartTimeline = new Chart(canvas, {
            type: 'bar',
            data: {
                labels: timelinePoints.map(p => Utils.formatMonthShort(p.month)),
                datasets: [
                    {
                        label: 'Actual Gross Profit',
                        data: timelinePoints.map(p => actualProfitByMonth[p.month] || 0),
                        backgroundColor: successColor + '60',
                        borderColor: successColor,
                        borderWidth: 1
                    },
                    {
                        label: 'Profit Needed (excl. B2B)',
                        data: timelinePoints.map(p => p.consumerBERevenue),
                        backgroundColor: primaryColor + '40',
                        borderColor: primaryColor,
                        borderWidth: 1
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                scales: {
                    x: { grid: { display: false } },
                    y: {
                        title: { display: true, text: 'Revenue ($)' },
                        ticks: {
                            callback: (v) => '$' + v.toLocaleString()
                        }
                    }
                },
                plugins: {
                    tooltip: {
                        callbacks: {
                            label: (ctx) => `${ctx.dataset.label}: $${ctx.parsed.y.toLocaleString(undefined, { minimumFractionDigits: 2 })}`
                        }
                    }
                }
            }
        });
    },

    /**
     * Refresh the break-even progress tracker ("as of" analysis)
     * @param {Object} beResult - From Utils.computeBreakEven()
     * @param {Object} cfg - Break-even config
     * @param {Array} months - Full timeline months array
     */
    refreshBreakevenProgress(beResult, cfg, months, timelinePoints) {
        const section = document.getElementById('beProgressSection');

        if (!beResult || !beResult.isValid || !months || months.length < 2) {
            section.style.display = 'none';
            return;
        }

        section.style.display = '';

        // Store state for dropdown change handler
        this._beProgressState = { beResult, cfg, months, timelinePoints };

        // Populate month dropdown
        const monthSelect = document.getElementById('beProgressMonth');
        const currentMonth = Utils.getCurrentMonth();
        const prevValue = monthSelect.value;

        // Only months up to current month (can't track future progress)
        const availableMonths = months.filter(m => m <= currentMonth);
        if (availableMonths.length === 0) {
            section.style.display = 'none';
            return;
        }

        monthSelect.innerHTML = availableMonths.map(m =>
            `<option value="${m}">${Utils.formatMonthShort(m)}</option>`
        ).join('');

        // Restore previous selection or default to latest available month
        const defaultMonth = prevValue && availableMonths.includes(prevValue) ? prevValue : availableMonths[availableMonths.length - 1];
        monthSelect.value = defaultMonth;

        this._computeAndRenderProgress(beResult, cfg, months, monthSelect.value, this._beProgressState.timelinePoints);
    },

    /**
     * Compute progress data and render cards + chart
     */
    _computeAndRenderProgress(beResult, cfg, months, asOfMonth, timelinePoints) {
        // Gross profit directly from P&L (Revenue - COGS, with overrides)
        const gpByMonth = Database.getMonthlyGrossProfit(months);
        const plRevenue = Database.getPLRevenueByMonth();
        const totalMonths = months.length;
        const timelineStart = months[0];
        const timelineEnd = months[months.length - 1];

        // Elapsed months (from start through asOfMonth)
        const elapsedMonths = months.filter(m => m <= asOfMonth);
        const elapsedCount = elapsedMonths.length;
        const remainingCount = totalMonths - elapsedCount;

        // Actual cumulative gross profit through asOfMonth (from P&L)
        let actualTotal = 0, actualB2b = 0, actualConsumer = 0;
        elapsedMonths.forEach(m => {
            actualTotal += (gpByMonth[m] || 0);
            actualB2b += (plRevenue.b2b[m] || 0);
            actualConsumer += (plRevenue.consumer[m] || 0);
        });

        // Total break-even revenue needed — use per-month targets if available
        let totalBERevenue, targetByNow;
        if (timelinePoints && timelinePoints.length > 0) {
            // Sum per-month break-even revenue from timeline (accounts for variable fixed costs)
            totalBERevenue = timelinePoints.reduce((sum, tp) => sum + tp.revenue, 0);
            // Target by now = sum of per-month targets for elapsed months
            targetByNow = 0;
            elapsedMonths.forEach(m => {
                const tp = timelinePoints.find(p => p.month === m);
                if (tp) targetByNow += tp.revenue;
            });
        } else {
            // Fallback: flat average
            totalBERevenue = beResult.breakEvenRevenue * totalMonths;
            targetByNow = (totalBERevenue / totalMonths) * elapsedCount;
        }

        // Remaining revenue needed
        const remainingRevenue = Math.max(0, totalBERevenue - actualTotal);

        // Monthly revenue needed going forward
        const monthlyNeeded = remainingCount > 0 ? remainingRevenue / remainingCount : 0;

        // On track?
        const onTrack = actualTotal >= targetByNow;

        // Build per-month chart data for the full timeline
        const chartMonths = [];
        const chartActual = [];
        const chartTarget = [];
        const chartProjected = []; // Placeholder for future projected data tab
        const chartCumulativeActual = [];
        const chartCumulativePace = [];
        let cumActual = 0, cumPace = 0;

        months.forEach(m => {
            chartMonths.push(Utils.formatMonthShort(m));
            const monthActual = gpByMonth[m] || 0;
            chartActual.push(Math.round(monthActual * 100) / 100);

            // Break-even target for this month
            let monthTarget = 0;
            if (timelinePoints) {
                const tp = timelinePoints.find(p => p.month === m);
                if (tp) monthTarget = tp.consumerBERevenue;
            } else {
                monthTarget = beResult.breakEvenRevenue;
            }
            chartTarget.push(Math.round(monthTarget * 100) / 100);

            // Projected: placeholder (null = not shown on chart)
            chartProjected.push(null);

            // Cumulative lines
            cumActual += monthActual;
            cumPace += monthTarget;
            chartCumulativeActual.push(Math.round(cumActual * 100) / 100);
            chartCumulativePace.push(Math.round(cumPace * 100) / 100);
        });

        const progressData = {
            actualB2b: Math.round(actualB2b * 100) / 100,
            actualConsumer: Math.round(actualConsumer * 100) / 100,
            actualTotal: Math.round(actualTotal * 100) / 100,
            targetByNow: Math.round(targetByNow * 100) / 100,
            totalBERevenue: Math.round(totalBERevenue * 100) / 100,
            remainingRevenue: Math.round(remainingRevenue * 100) / 100,
            monthlyNeeded: Math.round(monthlyNeeded * 100) / 100,
            onTrack,
            elapsedCount,
            remainingCount,
            totalMonths,
            asOfMonth,
            timelineStart,
            timelineEnd,
            // Per-month chart arrays
            chartMonths,
            chartActual,
            chartTarget,
            chartProjected,
            chartCumulativeActual,
            chartCumulativePace
        };

        UI.renderBreakevenProgressCards(progressData);
        this._renderBeProgressChart(progressData);
    },

    /**
     * Render the break-even progress chart: monthly bars (actual vs target) + cumulative pace line
     */
    _renderBeProgressChart(data) {
        const canvas = document.getElementById('beProgressChart');
        if (!canvas || typeof Chart === 'undefined') return;

        if (this._beChartProgress) {
            this._beChartProgress.destroy();
            this._beChartProgress = null;
        }

        const successColor = '#28a745';
        const projectedColor = '#17a2b8';
        const paceColor = '#1a3a4a';

        const datasets = [
            {
                type: 'line',
                label: 'Gross Profit',
                data: data.chartCumulativeActual,
                borderColor: successColor,
                backgroundColor: successColor + '20',
                fill: true,
                tension: 0.2,
                pointRadius: 4,
                pointBackgroundColor: successColor,
                borderWidth: 2,
                order: 2
            },
            {
                type: 'line',
                label: 'Projected Profit',
                data: data.chartProjected,
                borderColor: projectedColor,
                backgroundColor: projectedColor + '15',
                fill: true,
                tension: 0.2,
                pointRadius: 4,
                pointBackgroundColor: projectedColor,
                borderWidth: 2,
                borderDash: [5, 5],
                order: 1
            },
            {
                type: 'line',
                label: 'Monthly Pace',
                data: data.chartCumulativePace,
                borderColor: paceColor,
                backgroundColor: paceColor,
                fill: false,
                tension: 0.2,
                pointRadius: 4,
                pointBackgroundColor: paceColor,
                borderWidth: 2,
                order: 3
            }
        ];

        this._beChartProgress = new Chart(canvas, {
            type: 'line',
            data: {
                labels: data.chartMonths,
                datasets
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                scales: {
                    x: {
                        grid: { display: false }
                    },
                    y: {
                        title: { display: true, text: 'Cumulative Revenue ($)' },
                        ticks: {
                            callback: (v) => '$' + v.toLocaleString()
                        },
                        beginAtZero: true
                    }
                },
                plugins: {
                    tooltip: {
                        callbacks: {
                            label: (ctx) => {
                                if (ctx.parsed.y === null || ctx.parsed.y === 0) return null;
                                return `${ctx.dataset.label}: $${ctx.parsed.y.toLocaleString(undefined, { minimumFractionDigits: 2 })}`;
                            }
                        }
                    },
                    legend: {
                        labels: {
                            filter: (item) => {
                                // Hide 'Projected Revenue' from legend if all values are null
                                if (item.text === 'Projected Revenue') {
                                    return data.chartProjected.some(v => v !== null);
                                }
                                return true;
                            }
                        }
                    }
                }
            }
        });
    },

    /**
     * Open the break-even config modal and populate fields
     */
    openBeConfigModal() {
        const cfg = Database.getBreakevenConfig();

        // Populate year dropdowns FIRST (before setting values)
        const years = Utils.generateYearOptions();
        ['beTimelineStartYear', 'beTimelineEndYear'].forEach(id => {
            const sel = document.getElementById(id);
            sel.innerHTML = '<option value="">Year...</option>';
            years.forEach(y => {
                const opt = document.createElement('option');
                opt.value = y;
                opt.textContent = y;
                sel.appendChild(opt);
            });
        });

        // Timeline override values
        const tl = cfg.timeline || {};
        if (tl.start) {
            const [y, m] = tl.start.split('-');
            document.getElementById('beTimelineStartMonth').value = m;
            document.getElementById('beTimelineStartYear').value = y;
        } else {
            document.getElementById('beTimelineStartMonth').value = '';
            document.getElementById('beTimelineStartYear').value = '';
        }
        if (tl.end) {
            const [y, m] = tl.end.split('-');
            document.getElementById('beTimelineEndMonth').value = m;
            document.getElementById('beTimelineEndYear').value = y;
        } else {
            document.getElementById('beTimelineEndMonth').value = '';
            document.getElementById('beTimelineEndYear').value = '';
        }

        // Data source toggle
        document.getElementById('beDataSource').value = cfg.dataSource || 'actual';

        // As-of month picker: populate and show/hide based on data source
        const isProjected = (cfg.dataSource || 'actual') === 'projected';
        document.getElementById('beAsOfGroup').style.display = isProjected ? '' : 'none';
        const beAsOfEl = document.getElementById('beAsOfMonth');
        if (beAsOfEl) {
            const beTl = this.getBreakevenTimeline(cfg);
            const tlMonths = (beTl.start && beTl.end)
                ? Utils.generateMonthRange(beTl.start, beTl.end)
                : [Utils.getCurrentMonth()];
            this._populateAsOfSelect(beAsOfEl, tlMonths, cfg.asOfMonth || 'current');
        }

        // Consumer channel
        const consumer = cfg.consumer || {};
        document.getElementById('beConsumerEnabled').checked = consumer.enabled !== false;
        document.getElementById('beConsumerPrice').value = consumer.avgPrice || '';
        document.getElementById('beConsumerCogs').value = consumer.avgCogs || '';
        document.getElementById('beConsumerFields').style.display = consumer.enabled !== false ? 'flex' : 'none';

        // B2B channel
        const b2b = cfg.b2b || {};
        document.getElementById('beB2bEnabled').checked = b2b.enabled === true;
        document.getElementById('beB2bUnits').value = b2b.monthlyUnits || '';
        document.getElementById('beB2bRate').value = b2b.ratePerUnit || '';
        document.getElementById('beB2bCogs').value = b2b.cogsPerUnit || '';
        document.getElementById('beB2bFields').style.display = b2b.enabled ? 'flex' : 'none';

        // Unit increment
        document.getElementById('beUnitIncrement').value = cfg.unitIncrement || 100;

        // Fixed cost override
        document.getElementById('beFixedCostOverride').value = cfg.fixedCostOverride || '';

        // Update CM previews and cost hints
        this._updateBeCmPreview('consumer');
        this._updateBeCmPreview('b2b');
        this._updateBeFixedCostHints();

        UI.showModal('beConfigModal');
    },

    /**
     * Update contribution margin preview for a channel
     * @param {string} channel - 'consumer' or 'b2b'
     */
    _updateBeCmPreview(channel) {
        if (channel === 'consumer') {
            const price = parseFloat(document.getElementById('beConsumerPrice').value) || 0;
            const cogs = parseFloat(document.getElementById('beConsumerCogs').value) || 0;
            document.getElementById('beConsumerCmPreview').textContent = Utils.formatCurrency(price - cogs);
        } else {
            const rate = parseFloat(document.getElementById('beB2bRate').value) || 0;
            const cogs = parseFloat(document.getElementById('beB2bCogs').value) || 0;
            document.getElementById('beB2bCmPreview').textContent = Utils.formatCurrency(rate - cogs);
        }
    },

    /**
     * Update fixed cost source hints in the config modal.
     * Shows average $/mo across the break-even timeline (not just current month).
     */
    _updateBeFixedCostHints() {
        const cfg = Database.getBreakevenConfig();
        const timeline = this.getBreakevenTimeline(cfg);
        const realCurrentMonth = Utils.getCurrentMonth();
        const useProjected = (document.getElementById('beDataSource').value || cfg.dataSource) === 'projected';
        // Use as-of month from the modal picker (may not be saved yet)
        const beAsOfVal = document.getElementById('beAsOfMonth').value;
        const currentMonth = (useProjected && beAsOfVal && beAsOfVal !== 'current') ? beAsOfVal : realCurrentMonth;

        let months = [];
        if (timeline.start && timeline.end) {
            months = Utils.generateMonthRange(timeline.start, timeline.end);
        } else {
            months = [currentMonth];
        }

        // Check manual override first
        const overrideVal = parseFloat(document.getElementById('beFixedCostOverride').value);
        const hint = document.getElementById('bePLCostHint');

        if (!isNaN(overrideVal) && overrideVal > 0) {
            if (hint) hint.textContent = `Monthly fixed costs: ${Utils.formatCurrency(overrideVal)}/mo (manual override)`;
            return;
        }

        // Fall back to P&L renderer's computed per-month operating expenses
        const totalOpexByMonth = UI._pnlMonthOpex || {};

        // Sum all months (actual + projected) to match P&L total
        const totalSum = months.reduce((acc, m) => acc + (totalOpexByMonth[m] || 0), 0);
        const totalFixed = months.length > 0 ? totalSum / months.length : 0;

        if (hint) hint.textContent = totalFixed > 0
            ? `Avg. monthly fixed costs: ${Utils.formatCurrency(totalFixed)}/mo (from P&L)`
            : '';
    },

    /**
     * Save break-even config from modal form
     */
    handleSaveBeConfig() {
        // Timeline override
        const startM = document.getElementById('beTimelineStartMonth').value;
        const startY = document.getElementById('beTimelineStartYear').value;
        const endM = document.getElementById('beTimelineEndMonth').value;
        const endY = document.getElementById('beTimelineEndYear').value;

        const timeline = {
            start: (startM && startY) ? `${startY}-${startM}` : null,
            end: (endM && endY) ? `${endY}-${endM}` : null
        };

        const beAsOfVal = document.getElementById('beAsOfMonth').value;
        const fixedOverrideVal = parseFloat(document.getElementById('beFixedCostOverride').value);
        const cfg = {
            timeline,
            dataSource: document.getElementById('beDataSource').value || 'actual',
            asOfMonth: (beAsOfVal && beAsOfVal !== 'current') ? beAsOfVal : null,
            fixedCostOverride: (!isNaN(fixedOverrideVal) && fixedOverrideVal > 0) ? fixedOverrideVal : null,
            consumer: {
                enabled: document.getElementById('beConsumerEnabled').checked,
                avgPrice: parseFloat(document.getElementById('beConsumerPrice').value) || 0,
                avgCogs: parseFloat(document.getElementById('beConsumerCogs').value) || 0
            },
            b2b: {
                enabled: document.getElementById('beB2bEnabled').checked,
                monthlyUnits: parseInt(document.getElementById('beB2bUnits').value) || 0,
                ratePerUnit: parseFloat(document.getElementById('beB2bRate').value) || 0,
                cogsPerUnit: parseFloat(document.getElementById('beB2bCogs').value) || 0
            },
            unitIncrement: parseInt(document.getElementById('beUnitIncrement').value) || 100
        };

        Database.setBreakevenConfig(cfg);
        UI.hideModal('beConfigModal');
        this.refreshBreakeven();
    },

    /**
     * Download a blob as a file
     * @param {Blob} blob - The file blob
     * @param {string} filename - The filename
     */
    async downloadBlob(blob, filename) {
        if (window.showSaveFilePicker) {
            try {
                const ext = filename.split('.').pop().toLowerCase();
                const typeMap = {
                    zip: { description: 'ZIP Archive', mime: 'application/zip' },
                    csv: { description: 'CSV File', mime: 'text/csv' },
                    db:  { description: 'Database File', mime: 'application/octet-stream' }
                };
                const fileType = typeMap[ext] || typeMap.db;
                const handle = await window.showSaveFilePicker({
                    suggestedName: filename,
                    types: [{
                        description: fileType.description,
                        accept: { [fileType.mime]: [`.${ext}`] }
                    }]
                });
                const writable = await handle.createWritable();
                await writable.write(blob);
                await writable.close();
                return;
            } catch (e) {
                if (e.name === 'AbortError') return;
            }
        }
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    },

    /**
     * Stage a file load — reads the file and shows the load-choice modal.
     */
    async confirmLoadDatabase() {
        if (this._guardViewOnly()) return;
        if (!this.pendingFileLoad) return;

        try {
            this.pendingLoadBuffer = await this.pendingFileLoad.arrayBuffer();
            UI.hideModal('loadConfirmModal');

            // Update the active-company name in the replace-button label
            const active = CompanyManager.getActiveCompany();
            const targetEl = document.getElementById('loadReplaceTarget');
            if (targetEl && active) targetEl.textContent = active.name;

            // Show which option the user wants
            UI.showModal('loadChoiceModal');
        } catch (error) {
            console.error('Error reading file:', error);
            UI.hideModal('loadConfirmModal');
            this._cleanupPendingLoad();
            UI.showNotification('Failed to read file.', 'error');
        }
    },

    /**
     * Replace the active company's DB with the staged file.
     */
    async confirmLoadReplace() {
        if (!this.pendingLoadBuffer) return;
        try {
            await CompanyManager.replaceActive(this.pendingLoadBuffer);

            this._finalizeLoad('Database loaded successfully');
        } catch (error) {
            console.error('Error loading database:', error);
            UI.showNotification('Failed to load database. The file may be corrupted.', 'error');
        }
    },

    /**
     * Import the staged file as a brand-new company.
     */
    async confirmLoadAsNew() {
        if (!this.pendingLoadBuffer) return;
        const registry = CompanyManager.getRegistry();
        if (registry.companies.length >= CompanyManager.MAX_COMPANIES) {
            UI.showNotification('Maximum 5 companies reached. Delete one first.', 'error');
            UI.hideModal('loadChoiceModal');
            this._cleanupPendingLoad();
            return;
        }
        const name = prompt('Name for this imported company:');
        if (!name || !name.trim()) {
            return; // user cancelled
        }
        try {
            const newId = await CompanyManager.importAsNew(this.pendingLoadBuffer, name.trim());
            UI.hideModal('loadChoiceModal');
            this._cleanupPendingLoad();

            if (confirm(`Switch to "${name.trim()}" now?`)) {
                await this.switchToCompany(newId);
            } else {
                CompanyManager.renderSwitcher();
                UI.showNotification(`"${name.trim()}" added. Switch to it from the company menu.`, 'success');
            }
        } catch (error) {
            console.error('Error importing company:', error);
            UI.showNotification('Failed to import. The file may be corrupted.', 'error');
        }
    },

    /**
     * Shared post-load cleanup and UI reload.
     */
    _finalizeLoad(message) {
        UI.hideModal('loadChoiceModal');
        this._cleanupPendingLoad();
        this._reloadUI();
        UI.showNotification(message || 'Database loaded successfully', 'success');
    },

    _cleanupPendingLoad() {
        this.pendingFileLoad = null;
        this.pendingLoadBuffer = null;
        document.getElementById('loadDbInput').value = '';
    },

    /**
     * Reload all UI state after a company switch or DB load.
     */
    _reloadUI() {
        this.savedFileHandle = null;

        const owner = Database.getJournalOwner();
        document.getElementById('journalOwner').value = owner;
        UI.updateJournalTitle(owner);
        document.getElementById('journalOwner').dispatchEvent(new Event('input'));

        UI.populateYearDropdowns();
        UI.populatePaymentForMonthDropdown();
        this.loadAndApplyTheme();
        this.loadShippingFeeRate();
        this.restoreTabOrder();
        this.setupTabDragDrop();
        this.applyHiddenTabs();
        this.setupTabScrollFade();
        this.loadAndApplyTimeline();
        this.initBalanceSheetDate();
        this.refreshAll();
        CompanyManager.renderSwitcher();
    },

    // ==================== COMPANY MANAGEMENT ====================

    /**
     * Show the first-time company naming modal and wait for the user's input.
     */
    promptInitialCompanyName() {
        return new Promise((resolve) => {
            UI.showModal('companyNameModal');
            const input = document.getElementById('initialCompanyNameInput');
            input.value = '';
            input.focus();

            const confirmBtn = document.getElementById('confirmCompanyNameBtn');
            const skipBtn = document.getElementById('skipCompanyNameBtn');

            const onConfirm = async () => {
                const name = input.value.trim() || 'My Company';
                await CompanyManager.rename(CompanyManager.getActiveCompany().id, name);
                CompanyManager.clearNamingPrompt();
                UI.hideModal('companyNameModal');
                cleanup();
                resolve();
            };

            const onSkip = () => {
                CompanyManager.clearNamingPrompt();
                UI.hideModal('companyNameModal');
                cleanup();
                resolve();
            };

            const onKeydown = (e) => { if (e.key === 'Enter') onConfirm(); };

            confirmBtn.addEventListener('click', onConfirm);
            skipBtn.addEventListener('click', onSkip);
            input.addEventListener('keydown', onKeydown);

            function cleanup() {
                confirmBtn.removeEventListener('click', onConfirm);
                skipBtn.removeEventListener('click', onSkip);
                input.removeEventListener('keydown', onKeydown);
            }
        });
    },

    /**
     * Switch the active company and reload the full UI.
     */
    async switchToCompany(companyId) {
        document.body.style.opacity = '0.5';
        this.clearUndoHistory();
        try {
            await CompanyManager.switchTo(companyId);
            this._reloadUI();
        } catch (err) {
            console.error('Company switch failed:', err);
            UI.showNotification('Failed to switch company.', 'error');
        } finally {
            document.body.style.opacity = '1';
        }
    },

    /**
     * Create a new blank company and switch to it.
     */
    async handleCreateCompany() {
        const registry = CompanyManager.getRegistry();
        if (registry.companies.length >= CompanyManager.MAX_COMPANIES) {
            UI.showNotification('Maximum 5 companies reached. Delete one first.', 'error');
            return;
        }
        const name = prompt('New company name:');
        if (!name || !name.trim()) return;

        try {
            const newId = await CompanyManager.createNew(name.trim());
            await this.switchToCompany(newId);
        } catch (err) {
            UI.showNotification('Failed to create company.', 'error');
        }
    },

    /**
     * Delete the specified company. If it was active, switchToCompany handles the transition.
     */
    async handleDeleteCompany(companyId) {
        const registry = CompanyManager.getRegistry();
        const comp = registry.companies.find(c => c.id === companyId);
        if (!comp) return;

        if (!confirm(`Delete "${comp.name}" and all its data? This cannot be undone.`)) return;

        document.body.style.opacity = '0.5';
        try {
            const wasActive = companyId === registry.activeId;
            const updatedRegistry = await CompanyManager.delete(companyId);

            if (wasActive) {
                if (updatedRegistry.activeId) {
                    // Already switched by CompanyManager.delete(); just load new bytes
                    const bytes = await CompanyManager._readIDB(updatedRegistry.activeId);
                    Database.loadBytes(bytes);
                    this._reloadUI();
                } else {
                    // No companies left — create a blank one
                    await this.handleCreateCompany();
                }
            } else {
                // Deleted a non-active company — just re-render
                CompanyManager.renderSwitcher();
                this.renderManageCompanies();
            }
            UI.showNotification(`"${comp.name}" deleted.`, 'success');
        } catch (err) {
            console.error('Delete failed:', err);
            UI.showNotification('Failed to delete company.', 'error');
        } finally {
            document.body.style.opacity = '1';
        }
    },

    /**
     * Rename a company inline from the manage modal.
     */
    async handleRenameCompany(companyId) {
        const registry = CompanyManager.getRegistry();
        const comp = registry.companies.find(c => c.id === companyId);
        if (!comp) return;

        const name = prompt('New name:', comp.name);
        if (!name || !name.trim() || name.trim() === comp.name) return;

        await CompanyManager.rename(companyId, name.trim());
        CompanyManager.renderSwitcher();
        this.renderManageCompanies();

        // If renaming the active company, update the journal owner display too
        if (companyId === CompanyManager.getRegistry().activeId) {
            // Update company button label only (journal owner field is separate)
        }
        UI.showNotification('Company renamed.', 'success');
    },

    /**
     * Open the manage companies modal.
     */
    openManageCompanies() {
        this.renderManageCompanies();
        this._updateCopySectionVisibility();
        UI.showModal('manageCompaniesModal');
    },

    renderManageCompanies() {
        CompanyManager.renderManageTable();
    },

    _updateCopySectionVisibility() {
        const all = CompanyManager.getAll();
        const section = document.getElementById('copySectionWrapper');
        if (section) section.style.display = all.length > 1 ? '' : 'none';
    },

    // ==================== PROJECTION HELPERS ====================

    _populateAsOfSelect(selectEl, months, savedValue) {
        if (!selectEl || !months || months.length === 0) return;
        // Build continuous month range from first to last (fills gaps)
        const first = months[0];
        const last = months[months.length - 1];
        const allMonths = [];
        let [y, m] = first.split('-').map(Number);
        const [endY, endM] = last.split('-').map(Number);
        while (y < endY || (y === endY && m <= endM)) {
            allMonths.push(y + '-' + String(m).padStart(2, '0'));
            m++;
            if (m > 12) { m = 1; y++; }
        }
        const currentVal = selectEl.value;
        // Only rebuild if month count changed (avoid flicker)
        if (selectEl.options.length === allMonths.length + 1) return;
        selectEl.innerHTML = '<option value="current">Current</option>';
        const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        allMonths.forEach(mo => {
            const opt = document.createElement('option');
            opt.value = mo;
            const [yr, mn] = mo.split('-');
            opt.textContent = monthNames[parseInt(mn) - 1] + ' ' + yr;
            selectEl.appendChild(opt);
        });
        // Restore saved value from DB, or previous in-memory selection
        const restoreVal = savedValue || currentVal;
        if (restoreVal && [...selectEl.options].some(o => o.value === restoreVal)) {
            selectEl.value = restoreVal;
        }
    },

    resetPnLProjections() {
        const psConfig = Database.getProjectedSalesConfig();
        const pnlAsOfEl = document.getElementById('pnlAsOfMonth');
        const asOfVal = pnlAsOfEl ? pnlAsOfEl.value : 'current';
        const startMonth = asOfVal !== 'current' ? asOfVal : (psConfig.projectionStartMonth || Utils.getCurrentMonth());
        Database.clearPLOverridesFrom(startMonth);
        this.refreshPnL();
        UI.showNotification('P&L projections reset from ' + startMonth, 'success');
    },

    resetCashFlowProjections() {
        const psConfig = Database.getProjectedSalesConfig();
        const cfAsOfEl = document.getElementById('cfAsOfMonth');
        const asOfVal = cfAsOfEl ? cfAsOfEl.value : 'current';
        const startMonth = asOfVal !== 'current' ? asOfVal : (psConfig.projectionStartMonth || Utils.getCurrentMonth());
        Database.clearCashFlowOverridesFrom(startMonth);
        this.refreshCashFlow();
        UI.showNotification('Cash flow projections reset from ' + startMonth, 'success');
    },

    // ==================== PROJECTED SALES ====================

    refreshProjectedSales() {
        const config = Database.getProjectedSalesConfig();
        const timeline = this.getTimeline();
        const currentMonth = Utils.getCurrentMonth();

        // Sync UI controls
        document.getElementById('psEnabled').checked = config.enabled;
        document.getElementById('psConfigPanel').style.display = config.enabled ? 'block' : 'none';

        // Populate year selects
        const years = Utils.generateYearOptions();
        const yearSelect = document.getElementById('psStartYear');
        if (yearSelect && yearSelect.options.length <= 1) {
            yearSelect.innerHTML = '<option value="">Year...</option>';
            years.forEach(y => {
                const opt = document.createElement('option');
                opt.value = y; opt.textContent = y;
                yearSelect.appendChild(opt);
            });
        }

        // Sync projection start month
        if (config.projectionStartMonth) {
            const [y, m] = config.projectionStartMonth.split('-');
            document.getElementById('psStartMonth').value = m;
            document.getElementById('psStartYear').value = y;
        } else {
            document.getElementById('psStartMonth').value = '';
            document.getElementById('psStartYear').value = '';
        }

        // Sync sales tax rate
        const taxRateEl = document.getElementById('psSalesTaxRate');
        if (taxRateEl) taxRateEl.value = config.salesTaxRate || '';

        // Sync channel fields
        this._syncPsChannel('online', config.channels.online);
        this._syncPsChannel('tradeshow', config.channels.tradeshow);

        // Determine months from timeline
        let months = [];
        const startMonth = timeline.start || config.projectionStartMonth;
        const endMonth = timeline.end;
        if (startMonth && endMonth) {
            months = Utils.generateMonthRange(startMonth, endMonth);
        } else if (config.projectionStartMonth) {
            // Fallback: projection start to 12 months from now
            let end = currentMonth;
            for (let i = 0; i < 12; i++) end = Utils.nextMonth(end);
            months = Utils.generateMonthRange(config.projectionStartMonth, endMonth || end);
        }

        // Get spreadsheet data and render
        const psData = Database.getProjectedSalesSpreadsheet(config, months);
        UI.renderProjectedSalesSummaryCards(psData, months);
        UI.renderProjectedSalesGrid(psData, months, currentMonth);
        this.setupPsUnitEditing();
    },

    _syncPsChannel(key, ch) {
        const prefix = key === 'online' ? 'psOnline' : 'psTradeshow';
        document.getElementById(prefix + 'Enabled').checked = ch.enabled;
        document.getElementById(prefix + 'Fields').style.display = ch.enabled ? 'flex' : 'none';
        document.getElementById(prefix + 'Price').value = ch.avgPrice || '';
        document.getElementById(prefix + 'Cogs').value = ch.avgCogs || '';
        const cm = (ch.avgPrice || 0) - (ch.avgCogs || 0);
        document.getElementById(prefix + 'Cm').textContent = Utils.formatCurrency(cm);
    },

    handleSaveProjectedSales() {
        const startM = document.getElementById('psStartMonth').value;
        const startY = document.getElementById('psStartYear').value;
        const existing = Database.getProjectedSalesConfig();

        const config = {
            enabled: document.getElementById('psEnabled').checked,
            projectionStartMonth: (startM && startY) ? `${startY}-${startM}` : null,
            salesTaxRate: parseFloat(document.getElementById('psSalesTaxRate').value) || 0,
            channels: {
                online: {
                    enabled: document.getElementById('psOnlineEnabled').checked,
                    avgPrice: parseFloat(document.getElementById('psOnlinePrice').value) || 0,
                    avgCogs: parseFloat(document.getElementById('psOnlineCogs').value) || 0,
                    units: existing.channels.online.units || {}
                },
                tradeshow: {
                    enabled: document.getElementById('psTradeshowEnabled').checked,
                    avgPrice: parseFloat(document.getElementById('psTradeshowPrice').value) || 0,
                    avgCogs: parseFloat(document.getElementById('psTradeshowCogs').value) || 0,
                    units: existing.channels.tradeshow.units || {}
                }
            }
        };

        Database.setProjectedSalesConfig(config);
        this.refreshProjectedSales();
        UI.showNotification('Projected sales saved', 'success');
    },

    setupPsUnitEditing() {
        const container = document.getElementById('psMonthlyGrid');
        if (!container || container.dataset.psUnitSetup) return;
        container.dataset.psUnitSetup = '1';

        container.addEventListener('click', (e) => {
            const cell = e.target.closest('.ps-unit-editable');
            if (!cell || cell.querySelector('input')) return;

            const channel = cell.dataset.channel;
            const month = cell.dataset.month;
            const currentVal = parseInt(cell.textContent.trim()) || 0;

            const input = document.createElement('input');
            input.type = 'number';
            input.min = '0';
            input.step = '1';
            input.className = 'pnl-cell-input';
            input.value = currentVal || '';

            cell.textContent = '';
            cell.appendChild(input);
            input.focus();
            input.select();

            const save = () => {
                const newVal = parseInt(input.value) || 0;
                const config = Database.getProjectedSalesConfig();
                if (!config.channels[channel].units) config.channels[channel].units = {};
                if (newVal > 0) {
                    config.channels[channel].units[month] = newVal;
                } else {
                    delete config.channels[channel].units[month];
                }
                Database.setProjectedSalesConfig(config);
                // Re-render grid (reset flag so editing can re-bind)
                delete container.dataset.psUnitSetup;
                this.refreshProjectedSales();
            };

            input.addEventListener('blur', save);
            input.addEventListener('keydown', (ke) => {
                if (ke.key === 'Enter') { ke.preventDefault(); input.blur(); }
                else if (ke.key === 'Escape') {
                    ke.preventDefault();
                    delete container.dataset.psUnitSetup;
                    this.refreshProjectedSales();
                }
            });
        });
    },

    // ==================== SYNC / GROUP SHARING ====================

    async initSync() {
        const supaConfig = Database.getSupabaseConfig();
        if (!supaConfig || !supaConfig.url || !supaConfig.anonKey) return;

        // Show share button when Supabase is configured
        const shareBtn = document.getElementById('shareBtn');
        if (shareBtn) shareBtn.style.display = '';

        SupabaseAdapter.init(supaConfig.url, supaConfig.anonKey);
        SyncService.api = SupabaseAdapter;
        SyncService.onStatusChange = (info) => this.handleSyncStatusChange(info);
        SyncService.onRemoteUpdate = (info) => this.handleSyncRemoteUpdate(info);
        SyncService.onConflict = (err) => this.handleSyncConflict(err);

        const syncConfig = Database.getSyncConfig();
        if (syncConfig && syncConfig.groupId && syncConfig.userName && syncConfig.memberId) {
            try {
                // Verify member still exists (not removed by admin)
                const memberInfo = await SupabaseAdapter.verifyMember(
                    syncConfig.groupId, syncConfig.memberId
                );

                if (!memberInfo) {
                    Database.clearSyncConfig();
                    this.updateSyncUI();
                    UI.showNotification('You were removed from the group by an admin.', 'error');
                    return;
                }

                await SyncService.joinGroup(syncConfig.groupId, syncConfig.userName, memberInfo);
                if (!this._syncAutoSaveWrapped) {
                    SyncService.wrapAutoSave(Database);
                    this._syncAutoSaveWrapped = true;
                }
                SyncService.startPolling();
                this.updateSyncUI();
            } catch (err) {
                console.error('Failed to reconnect to group:', err);
            }
        }
    },

    encodeInviteCode(groupId, url, anonKey) {
        return btoa(JSON.stringify({ g: groupId, u: url, k: anonKey }));
    },

    decodeInviteCode(code) {
        try {
            const parsed = JSON.parse(atob(code));
            if (!parsed.g || !parsed.u || !parsed.k) return null;
            return { groupId: parsed.g, url: parsed.u, anonKey: parsed.k };
        } catch {
            return null;
        }
    },

    _initSupabase(url, anonKey) {
        Database.setSupabaseConfig({ url, anonKey });
        SupabaseAdapter.init(url, anonKey);
        SyncService.api = SupabaseAdapter;
        SyncService.onStatusChange = (info) => this.handleSyncStatusChange(info);
        SyncService.onRemoteUpdate = (info) => this.handleSyncRemoteUpdate(info);
        SyncService.onConflict = (err) => this.handleSyncConflict(err);
    },

    openSyncMenu() {
        this.updateSyncUI();

        // Pre-fill create form with saved config
        const supaConfig = Database.getSupabaseConfig();
        if (supaConfig) {
            document.getElementById('supabaseUrl').value = supaConfig.url || '';
            document.getElementById('supabaseAnonKey').value = supaConfig.anonKey || '';
        }
        const syncConfig = Database.getSyncConfig();
        if (syncConfig && syncConfig.userName) {
            document.getElementById('syncCreateName').value = syncConfig.userName;
            document.getElementById('syncJoinName').value = syncConfig.userName;
        }

        UI.showModal('syncMenuModal');
    },

    async handleCreateGroup() {
        const userName = document.getElementById('syncCreateName').value.trim();
        const password = document.getElementById('syncCreatePassword').value;
        const groupName = document.getElementById('newGroupName').value.trim();
        const url = document.getElementById('supabaseUrl').value.trim();
        const anonKey = document.getElementById('supabaseAnonKey').value.trim();
        if (!userName || !password || !groupName || !url || !anonKey) {
            UI.showNotification('All fields are required', 'error');
            return;
        }

        try {
            this._initSupabase(url, anonKey);

            // Create group first, then register creator as admin
            const groupResult = await SupabaseAdapter.createGroup(groupName);
            const memberInfo = await SupabaseAdapter.registerMember(
                groupResult.groupId, userName, password, 'admin'
            );

            // Set up SyncService state
            SyncService.groupId = groupResult.groupId;
            SyncService.currentUser = userName;
            SyncService.memberId = memberInfo.id;
            SyncService.memberRole = memberInfo.role;
            SyncService.localVersion = 0;
            SyncService.isConnected = true;

            const blob = new Uint8Array(Database.db.export());
            await SyncService.push(blob);

            Database.setSyncConfig({
                groupId: groupResult.groupId, groupName, userName,
                memberId: memberInfo.id, memberRole: memberInfo.role
            });

            if (!this._syncAutoSaveWrapped) {
                SyncService.wrapAutoSave(Database);
                this._syncAutoSaveWrapped = true;
            }
            SyncService.startPolling();

            const inviteCode = this.encodeInviteCode(groupResult.groupId, url, anonKey);
            this._currentInviteCode = inviteCode;

            UI.hideModal('syncMenuModal');
            document.getElementById('inviteCodeDisplay').textContent = inviteCode;
            UI.showModal('groupCreatedModal');
            this.updateSyncUI();
        } catch (err) {
            console.error('Failed to create group:', err);
            UI.showNotification('Failed to create group: ' + err.message, 'error');
        }
    },

    async handleJoinGroup() {
        const userName = document.getElementById('syncJoinName').value.trim();
        const password = document.getElementById('syncJoinPassword').value;
        const code = document.getElementById('joinGroupCode').value.trim();
        if (!userName || !password || !code) {
            UI.showNotification('All fields are required', 'error');
            return;
        }

        const decoded = this.decodeInviteCode(code);
        if (!decoded) {
            UI.showNotification('Invalid invite code. Ask the group creator for a new one.', 'error');
            return;
        }

        try {
            this._initSupabase(decoded.url, decoded.anonKey);

            let memberInfo;

            if (this._joinMode === 'rejoin') {
                // Rejoin: authenticate only — don't auto-register
                try {
                    memberInfo = await SupabaseAdapter.authenticateMember(
                        decoded.groupId, userName, password
                    );
                } catch (authErr) {
                    if (authErr.code === 'INVALID_PASSWORD') {
                        UI.showNotification('Incorrect password for this username.', 'error');
                        return;
                    }
                    throw authErr;
                }
                if (!memberInfo) {
                    UI.showNotification('No account found with this username. Switch to "New member" to register.', 'error');
                    return;
                }
            } else {
                // New member: register only — don't auto-authenticate
                try {
                    memberInfo = await SupabaseAdapter.registerMember(
                        decoded.groupId, userName, password, 'member'
                    );
                } catch (regErr) {
                    if (regErr.code === 'MEMBER_EXISTS') {
                        UI.showNotification('Username already taken. Switch to "I have an account" to sign in.', 'error');
                        return;
                    }
                    throw regErr;
                }
            }

            const result = await SyncService.joinGroup(decoded.groupId, userName, memberInfo);

            const loaded = await SyncService.loadRemoteIntoDatabase(Database);
            if (loaded) {
                this.refreshAll();
                UI.showNotification('Joined "' + result.name + '" and loaded latest version', 'success');
            } else {
                const blob = new Uint8Array(Database.db.export());
                await SyncService.push(blob);
                UI.showNotification('Joined "' + result.name + '"', 'success');
            }

            Database.setSyncConfig({
                groupId: decoded.groupId, groupName: result.name, userName,
                memberId: memberInfo.id, memberRole: memberInfo.role
            });
            this._currentInviteCode = code;

            if (!this._syncAutoSaveWrapped) {
                SyncService.wrapAutoSave(Database);
                this._syncAutoSaveWrapped = true;
            }
            SyncService.startPolling();

            UI.hideModal('syncMenuModal');
            this.updateSyncUI();
        } catch (err) {
            console.error('Failed to join group:', err);
            UI.showNotification('Failed to join group: ' + err.message, 'error');
        }
    },

    async handleSyncPull() {
        try {
            const result = await SyncService.pull();
            if (result.updated && result.data) {
                Database.db = new Database.SQL.Database(result.data);
                Database.migrateSchema();
                await Database.saveToIndexedDB();
                this.refreshAll();
                UI.showNotification('Updated to v' + result.version + ' (by ' + result.savedBy + ')', 'success');
            } else {
                UI.showNotification('Already up to date', 'info');
            }
            this.updateSyncUI();
        } catch (err) {
            console.error('Pull failed:', err);
            UI.showNotification('Failed to pull: ' + err.message, 'error');
        }
    },

    handleSyncDisconnect() {
        SyncService.disconnect();
        Database.clearSyncConfig();
        this._currentInviteCode = null;
        this.updateSyncUI();
        UI.hideModal('syncMenuModal');
        UI.showNotification('Disconnected from group', 'info');
    },

    handleSyncStatusChange(info) {
        const dot = document.getElementById('syncStatusDot');
        const text = document.getElementById('syncStatusText');

        dot.className = 'sync-status-dot';
        switch (info.status) {
            case 'connected':
            case 'saved':
            case 'updated':
                dot.classList.add('connected');
                break;
            case 'conflict':
                dot.classList.add('conflict');
                break;
            case 'error':
            case 'disconnected':
                dot.classList.add('error');
                break;
            default:
                dot.style.display = 'none';
                return;
        }
        dot.style.display = '';
        if (text) text.textContent = info.message || info.status;
    },

    handleSyncRemoteUpdate(info) {
        if (info.data) {
            Database.db = new Database.SQL.Database(info.data);
            Database.migrateSchema();
            Database.saveToIndexedDB();
            this.refreshAll();
            UI.showNotification('Updated to v' + info.version + ' (by ' + info.savedBy + ')', 'info');
        }
        this.updateSyncUI();
    },

    handleSyncConflict(err) {
        UI.showNotification('Conflict: someone else saved. Pulling their changes...', 'error');
        this.handleSyncPull();
    },

    updateSyncUI() {
        const dot = document.getElementById('syncStatusDot');
        const disconnectedPanel = document.getElementById('syncDisconnectedPanel');
        const connectedPanel = document.getElementById('syncConnectedPanel');

        if (SyncService.isConnected) {
            dot.className = 'sync-status-dot connected';
            dot.style.display = '';

            if (disconnectedPanel) disconnectedPanel.style.display = 'none';
            if (connectedPanel) connectedPanel.style.display = '';

            const syncConfig = Database.getSyncConfig();
            const supaConfig = Database.getSupabaseConfig();
            document.getElementById('syncGroupName').textContent = (syncConfig && syncConfig.groupName) || '--';
            document.getElementById('syncCurrentUser').textContent = SyncService.currentUser || '--';
            document.getElementById('syncVersion').textContent = 'v' + SyncService.localVersion;

            // Show role
            const roleEl = document.getElementById('syncCurrentRole');
            if (roleEl) {
                roleEl.textContent = SyncService.memberRole === 'admin' ? 'Admin' : 'Member';
            }

            // Show/hide Members button based on role
            const membersBtn = document.getElementById('syncMembersBtn');
            if (membersBtn) {
                membersBtn.style.display = SyncService.memberRole === 'admin' ? '' : 'none';
            }

            // Rebuild invite code for connected panel
            if (!this._currentInviteCode && SyncService.groupId && supaConfig) {
                this._currentInviteCode = this.encodeInviteCode(SyncService.groupId, supaConfig.url, supaConfig.anonKey);
            }
            const codeEl = document.getElementById('syncInviteCode');
            codeEl.textContent = this._currentInviteCode || '--';
            codeEl.title = this._currentInviteCode || '';
        } else {
            dot.style.display = 'none';
            if (disconnectedPanel) disconnectedPanel.style.display = '';
            if (connectedPanel) connectedPanel.style.display = 'none';
        }
    },

    // ==================== SHARE VIEW-ONLY LINK ====================

    openShareModal() {
        const supaConfig = Database.getSupabaseConfig();
        if (!supaConfig || !supaConfig.url || !supaConfig.anonKey) {
            UI.showNotification('Sharing requires Supabase. Set up Group Sync first.', 'error');
            return;
        }

        // Reset to phase 1
        document.getElementById('shareConfigSection').style.display = '';
        document.getElementById('shareResultSection').style.display = 'none';
        document.getElementById('shareUrlSection').style.display = 'none';
        document.getElementById('shareStatus').textContent = '';

        // Populate report month dropdown
        const select = document.getElementById('shareReportMonth');
        const realCurrent = Utils.getCurrentMonth();
        const allMonths = this._getTimelineMonths().filter(m => m <= realCurrent);
        select.innerHTML = '<option value="">Current (no filter)</option>' +
            allMonths.map(m => `<option value="${m}">${Utils.formatMonthShort(m)}</option>`).join('');
        select.value = '';

        UI.showModal('shareModal');
    },

    async generateShare() {
        const supaConfig = Database.getSupabaseConfig();
        if (!SupabaseAdapter.isInitialized()) {
            SupabaseAdapter.init(supaConfig.url, supaConfig.anonKey);
        }

        const reportMonth = document.getElementById('shareReportMonth').value || null;

        // Switch to phase 2 with loading state
        document.getElementById('shareConfigSection').style.display = 'none';
        document.getElementById('shareResultSection').style.display = '';
        document.getElementById('shareStatus').textContent = 'Uploading snapshot...';
        document.getElementById('shareUrlSection').style.display = 'none';

        try {
            // Clone DB and inject report_month if set
            let blob;
            if (reportMonth) {
                const bytes = Database.db.export();
                const clone = new Database.SQL.Database(new Uint8Array(bytes));
                clone.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('report_month', ?)", [reportMonth]);
                blob = new Uint8Array(clone.export());
                clone.close();
            } else {
                blob = new Uint8Array(Database.db.export());
            }

            const journalName = Database.getJournalOwner() || 'Accounting Journal';
            const syncConfig = Database.getSyncConfig();
            const createdBy = (syncConfig && syncConfig.userName) || journalName;

            const result = await SupabaseAdapter.createShare(blob, createdBy, journalName);

            // Build share URL
            const shareToken = btoa(JSON.stringify({
                s: result.shareId,
                u: supaConfig.url,
                k: supaConfig.anonKey
            }));
            const shareUrl = window.location.origin + window.location.pathname + '#share=' + shareToken;

            // Display results
            document.getElementById('shareUrlDisplay').value = shareUrl;
            document.getElementById('shareStatus').textContent = '';
            document.getElementById('shareUrlSection').style.display = '';

            // Show report month badge if set
            const badge = document.getElementById('shareReportBadge');
            if (reportMonth) {
                badge.textContent = 'Report: ' + Utils.formatMonthShort(reportMonth);
                badge.style.display = '';
            } else {
                badge.style.display = 'none';
            }

            document.getElementById('shareExpiry').textContent =
                'Expires: ' + new Date(result.expiresAt).toLocaleDateString();

            this.generateShareQR(shareUrl);
        } catch (err) {
            console.error('Share failed:', err);
            document.getElementById('shareStatus').textContent = 'Share failed: ' + err.message;
        }
    },

    generateShareQR(url) {
        const container = document.getElementById('shareQrCode');
        container.innerHTML = '';
        if (typeof qrcode === 'undefined') {
            container.textContent = 'QR library not loaded';
            return;
        }
        const qr = qrcode(0, 'M');
        qr.addData(url);
        qr.make();
        container.innerHTML = qr.createSvgTag({ cellSize: 4, margin: 4 });
    },

    // ==================== VIEW-ONLY MODE ====================

    async initViewOnlyMode(token) {
        try {
            const decoded = JSON.parse(atob(token));
            if (!decoded.s || !decoded.u || !decoded.k) {
                throw new Error('Invalid share link');
            }

            // Initialize a temporary Supabase client (do NOT save to localStorage)
            const tempClient = supabase.createClient(decoded.u, decoded.k);

            // Fetch share metadata
            const { data: shareMeta, error: metaErr } = await tempClient
                .from('shares')
                .select('id, created_at, expires_at, size_bytes, storage_path, created_by, journal_name')
                .eq('id', decoded.s)
                .single();

            if (metaErr || !shareMeta) {
                throw new Error('Share not found. It may have been deleted.');
            }

            if (new Date(shareMeta.expires_at) < new Date()) {
                throw new Error('This share link has expired.');
            }

            // Download the blob from storage
            const { data: blobData, error: blobErr } = await tempClient.storage
                .from('db-blobs')
                .download(shareMeta.storage_path);
            if (blobErr) throw new Error('Failed to download shared journal');

            const arrayBuffer = await blobData.arrayBuffer();
            const uint8Array = new Uint8Array(arrayBuffer);

            // Initialize sql.js and load the blob (skip IndexedDB entirely)
            const SQL = await initSqlJs({
                locateFile: file => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.8.0/${file}`
            });
            Database.SQL = SQL;
            Database.db = new SQL.Database(uint8Array);
            Database.migrateSchema();

            // Disable auto-save — never write to IndexedDB in view-only mode
            Database.autoSave = function() {};

            this.isViewOnly = true;

            // Set up UI
            document.getElementById('entryDate').value = Utils.getTodayDate();
            UI.populateYearDropdowns();
            UI.populatePaymentForMonthDropdown();

            const owner = shareMeta.journal_name || Database.getJournalOwner();
            document.getElementById('journalOwner').value = owner;
            document.getElementById('journalOwner').disabled = true;
            UI.updateJournalTitle(owner);

            this.loadAndApplyTheme();
            this.loadShippingFeeRate();
            this.restoreTabOrder();
            this.setupTabDragDrop();
            this.applyHiddenTabs();
            this.setupTabScrollFade();
            this.loadAndApplyTimeline();
            this.initBalanceSheetDate();
            this.refreshAll();
            this.setupEventListeners();
            this.setupVESalesListeners();
            this.applyViewOnlyRestrictions();
            this.showViewOnlyBanner(shareMeta);

            document.body.style.opacity = '1';
            console.log('View-only mode initialized');
        } catch (err) {
            console.error('Failed to load shared journal:', err);
            document.body.style.opacity = '1';
            document.body.innerHTML =
                '<div style="display:flex;align-items:center;justify-content:center;min-height:100vh;font-family:DM Sans,sans-serif;">' +
                '<div style="text-align:center;max-width:420px;padding:40px;">' +
                '<h2 style="margin-bottom:12px;">Unable to Load Shared Journal</h2>' +
                '<p style="color:#6c757d;">' + this._escapeHtml(err.message) + '</p>' +
                '<p style="margin-top:24px;"><a href="' + window.location.origin + window.location.pathname + '">Go to app</a></p>' +
                '</div></div>';
        }
    },

    applyViewOnlyRestrictions() {
        const hideIds = [
            'newEntryBtn', 'addFolderEntriesBtn', 'manageCategoriesBtn',
            'saveDbBtn', 'saveAllDbBtn', 'loadDbBtn', 'shareBtn',
            'veCreateJournalBtn', 'veClearBtn', 'veImportSubmit',
            'veSyncFromServer', 'veAddEventBtn', 'veAddAllToJournalBtn',
            'veRemoveAllFromJournalBtn'
        ];
        hideIds.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.style.display = 'none';
        });

        const loadInput = document.getElementById('loadDbInput');
        if (loadInput) loadInput.style.display = 'none';

        const syncWrapper = document.querySelector('.sync-wrapper');
        if (syncWrapper) syncWrapper.style.display = 'none';

        const gearWrapper = document.querySelector('.gear-wrapper');
        if (gearWrapper) gearWrapper.style.display = 'none';

        // Hide VE import panel in view-only mode
        const veImportPanel = document.getElementById('veImportPanel');
        if (veImportPanel) veImportPanel.style.display = 'none';

        document.body.classList.add('view-only-mode');
    },

    showViewOnlyBanner(shareMeta) {
        const banner = document.createElement('div');
        banner.className = 'view-only-banner';

        const createdDate = new Date(shareMeta.created_at).toLocaleDateString();
        const createdBy = shareMeta.created_by || 'Unknown';
        const journalName = shareMeta.journal_name || 'Journal';

        banner.innerHTML =
            '<div class="view-only-banner-content">' +
            '<span>&#128274;</span> ' +
            '<span><strong>View-Only Snapshot</strong> &mdash; ' +
            this._escapeHtml(journalName) + ' &middot; Shared by ' +
            this._escapeHtml(createdBy) + ' on ' + createdDate + '</span>' +
            '</div>';

        document.body.insertBefore(banner, document.body.firstChild);
    },

    _guardViewOnly() {
        if (this.isViewOnly) {
            UI.showNotification('This is a view-only snapshot. Editing is not available.', 'info');
            return true;
        }
        return false;
    },

    _escapeHtml(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    },

    async openVersionHistory() {
        UI.hideModal('syncMenuModal');
        UI.showModal('versionHistoryModal');
        document.getElementById('versionHistoryList').innerHTML = '<p class="empty-state">Loading...</p>';

        try {
            const history = await SyncService.getHistory(20);
            const container = document.getElementById('versionHistoryList');

            if (!history || history.length === 0) {
                container.innerHTML = '<p class="empty-state">No versions yet.</p>';
                return;
            }

            container.innerHTML = history.map(v => {
                const date = new Date(v.savedAt).toLocaleString();
                const size = (v.sizeBytes / 1024).toFixed(1) + ' KB';
                const isCurrent = v.version === SyncService.localVersion;
                return '<div class="version-item ' + (isCurrent ? 'version-item-current' : '') + '">' +
                    '<div class="version-item-info">' +
                    '<div class="version-item-number">v' + v.version + (isCurrent ? ' (current)' : '') + '</div>' +
                    '<div class="version-item-meta">' + v.savedBy + ' &middot; ' + date + ' &middot; ' + size + '</div>' +
                    '</div>' +
                    (!isCurrent ? '<button type="button" class="btn btn-small btn-secondary rollback-btn" data-version="' + v.version + '">Rollback</button>' : '') +
                    '</div>';
            }).join('');

            container.querySelectorAll('.rollback-btn').forEach(btn => {
                btn.addEventListener('click', () => {
                    this._rollbackTargetVersion = parseInt(btn.dataset.version);
                    document.getElementById('rollbackMessage').textContent =
                        'This will replace your current data with version ' + btn.dataset.version + '. Continue?';
                    UI.showModal('rollbackModal');
                });
            });
        } catch (err) {
            console.error('Failed to load version history:', err);
            document.getElementById('versionHistoryList').innerHTML = '<p class="empty-state">Failed to load history.</p>';
        }
    },

    async handleConfirmRollback() {
        if (!this._rollbackTargetVersion) return;

        try {
            const result = await SyncService.pullVersion(this._rollbackTargetVersion);
            Database.db = new Database.SQL.Database(result.data);
            Database.migrateSchema();
            await Database.saveToIndexedDB();
            SyncService.localVersion = result.version;

            this.refreshAll();
            UI.hideModal('rollbackModal');
            UI.hideModal('versionHistoryModal');
            UI.showNotification('Rolled back to v' + result.version, 'success');

            const blob = new Uint8Array(Database.db.export());
            await SyncService.push(blob);
            this.updateSyncUI();
        } catch (err) {
            console.error('Rollback failed:', err);
            UI.showNotification('Rollback failed: ' + err.message, 'error');
        }
        this._rollbackTargetVersion = null;
    },

    async openMembersModal() {
        UI.showModal('membersModal');
        document.getElementById('membersList').innerHTML = '<p class="empty-state">Loading...</p>';

        try {
            const members = await SupabaseAdapter.listMembers(SyncService.groupId);
            const container = document.getElementById('membersList');

            if (!members || members.length === 0) {
                container.innerHTML = '<p class="empty-state">No members found.</p>';
                return;
            }

            container.innerHTML = members.map(m => {
                const isCurrentUser = m.id === SyncService.memberId;
                const joinDate = new Date(m.joined_at).toLocaleDateString();
                const roleLabel = m.role === 'admin' ? ' (Admin)' : '';
                const removeBtnHtml = (!isCurrentUser && SyncService.memberRole === 'admin')
                    ? '<button type="button" class="btn btn-small btn-danger remove-member-btn" data-member-id="' + m.id + '" data-member-name="' + m.display_name + '">Remove</button>'
                    : '';

                return '<div class="member-item">' +
                    '<div class="member-item-info">' +
                    '<div class="member-item-name">' + m.display_name + roleLabel + (isCurrentUser ? ' (you)' : '') + '</div>' +
                    '<div class="member-item-meta">Joined ' + joinDate + '</div>' +
                    '</div>' +
                    removeBtnHtml +
                    '</div>';
            }).join('');

            container.querySelectorAll('.remove-member-btn').forEach(btn => {
                btn.addEventListener('click', () => {
                    this._removeMemberTarget = {
                        id: btn.dataset.memberId,
                        name: btn.dataset.memberName
                    };
                    document.getElementById('removeMemberMessage').textContent =
                        'Remove "' + btn.dataset.memberName + '" from this group? They will need to re-register to rejoin.';
                    UI.showModal('removeMemberModal');
                });
            });
        } catch (err) {
            console.error('Failed to load members:', err);
            document.getElementById('membersList').innerHTML = '<p class="empty-state">Failed to load members.</p>';
        }
    },

    async handleConfirmRemoveMember() {
        if (!this._removeMemberTarget) return;

        try {
            await SupabaseAdapter.removeMember(
                SyncService.groupId,
                this._removeMemberTarget.id,
                SyncService.memberId
            );
            UI.hideModal('removeMemberModal');
            UI.showNotification('Removed "' + this._removeMemberTarget.name + '" from the group', 'success');
            this._removeMemberTarget = null;
            this.openMembersModal();
        } catch (err) {
            console.error('Failed to remove member:', err);
            UI.showNotification('Failed to remove member: ' + err.message, 'error');
        }
    },

    copyToClipboard(text) {
        navigator.clipboard.writeText(text).then(() => {
            UI.showNotification('Copied to clipboard', 'success');
        }).catch(() => {
            const textarea = document.createElement('textarea');
            textarea.value = text;
            document.body.appendChild(textarea);
            textarea.select();
            document.execCommand('copy');
            document.body.removeChild(textarea);
            UI.showNotification('Copied to clipboard', 'success');
        });
    },

    // ==================== VE SALES TAB ====================

    _veSales: [],
    _veFiltered: [],
    _veItemsCache: new Map(),
    _veEventsMap: new Map(),
    _veProductSortCol: 'qty',
    _veProductSortDir: 'desc',

    refreshVESales() {
        this._veSales = Database.getVESales();
        this._veItemsCache = Database.getAllVESaleItems();
        const events = Database.getAllVEEvents();
        this._veEventsMap = new Map(events.map(e => [e.id, e.name]));
        const meta = Database.getVEImportMeta();

        // Show import info
        const infoEl = document.getElementById('veImportInfo');
        if (meta && meta.importDate) {
            const d = new Date(meta.importDate);
            infoEl.textContent = 'Last import: ' + d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' });
        } else {
            infoEl.textContent = '';
        }

        // Toggle empty state vs dashboard
        const hasData = this._veSales.length > 0;
        document.getElementById('veEmptyState').style.display = hasData ? 'none' : 'block';
        document.getElementById('veDashboard').style.display = hasData ? 'block' : 'none';
        document.getElementById('veControls').style.display = hasData ? 'flex' : 'none';

        // Auto-open import panel when no data, collapse when data exists
        const panel = document.getElementById('veImportPanel');
        if (!hasData) {
            panel.classList.add('open');
        }

        document.getElementById('veSubtabs').style.display = hasData ? 'flex' : 'none';

        if (hasData) {
            this.vePopulatePresetDropdown();
            this.vePopulateEventDropdown();
            this.veApplyFilters();
            // Re-render events panel if it's the active sub-tab
            const eventsTab = document.getElementById('veSubtabEvents');
            if (eventsTab && eventsTab.style.display !== 'none') this.veRenderEventsPanel();
        }
    },

    vePopulatePresetDropdown() {
        const sel = document.getElementById('ve-filterPreset');
        const current = sel.value;
        const timeline = this.getTimeline();
        const now = new Date();
        const currentYM = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

        let start = timeline.start;
        let end = timeline.end;
        if (!start && this._veSales.length > 0) {
            const dates = this._veSales.map(s => s.date).filter(Boolean).sort();
            start = dates[0] ? dates[0].substring(0, 7) : null;
        }
        if (!end && this._veSales.length > 0) {
            const dates = this._veSales.map(s => s.date).filter(Boolean).sort();
            end = dates[dates.length - 1] ? dates[dates.length - 1].substring(0, 7) : null;
        }

        sel.innerHTML = '<option value="">Date preset...</option><option value="all">All time</option>';
        if (!start || !end) return;
        if (end > currentYM) end = currentYM;

        const months = [];
        let [sy, sm] = start.split('-').map(Number);
        const [ey, em] = end.split('-').map(Number);
        while (sy < ey || (sy === ey && sm <= em)) {
            const val = `${sy}-${String(sm).padStart(2, '0')}`;
            const label = new Date(sy, sm - 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
            months.push({ value: val, label, year: sy });
            sm++;
            if (sm > 12) { sm = 1; sy++; }
        }
        months.reverse();

        let currentYear = null;
        let optgroup = null;
        for (const m of months) {
            if (m.year !== currentYear) {
                currentYear = m.year;
                optgroup = document.createElement('optgroup');
                optgroup.label = String(m.year);
                sel.appendChild(optgroup);
            }
            const opt = document.createElement('option');
            opt.value = m.value;
            opt.textContent = m.label;
            optgroup.appendChild(opt);
        }
        sel.value = current;
    },

    vePopulateEventDropdown() {
        const events = Database.getAllVEEvents();
        const sel = document.getElementById('ve-filterEvent');
        const current = sel.value;
        sel.innerHTML = '<option value="">All sales</option><option value="__unassigned__">Unassigned only</option>';
        for (const evt of events) {
            const opt = document.createElement('option');
            opt.value = evt.id;
            opt.textContent = evt.name;
            sel.appendChild(opt);
        }
        sel.value = current;
    },

    veApplyFilters() {
        const source = document.getElementById('ve-filterSource').value;
        const from = document.getElementById('ve-filterFrom').value;
        const to = document.getElementById('ve-filterTo').value;
        const sortBy = document.getElementById('ve-sortBy').value;
        const eventFilter = document.getElementById('ve-filterEvent').value;

        this._veFiltered = this._veSales.filter(s => {
            if (source !== 'both' && s.source !== source) return false;
            if (from && s.date < from) return false;
            if (to && s.date > to) return false;
            if (eventFilter === '__unassigned__' && s.event_id) return false;
            if (eventFilter && eventFilter !== '__unassigned__' && String(s.event_id) !== eventFilter) return false;
            return true;
        });

        this._veFiltered.sort((a, b) => {
            switch (sortBy) {
                case 'date-desc': return (b.date || '').localeCompare(a.date || '');
                case 'date-asc': return (a.date || '').localeCompare(b.date || '');
                case 'total-desc': return b.total - a.total;
                case 'total-asc': return a.total - b.total;
                case 'txno': return String(a.transaction_no).localeCompare(String(b.transaction_no));
                default: return 0;
            }
        });

        this.veRenderAll();
    },

    veRenderAll() {
        this.veRenderSummary();
        this.veRenderSourceBreakdown();
        this.veRenderProducts();
        this.veRenderTransactions();
    },

    veFmt(n) {
        return '$' + (n || 0).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    },

    veFmtDate(isoStr) {
        if (!isoStr) return '';
        const [y, m, d] = isoStr.split('-').map(Number);
        const date = new Date(y, m - 1, d);
        return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
    },

    veRenderSummary() {
        let subtotal = 0, tax = 0, total = 0, shipping = 0, discount = 0;
        for (const s of this._veFiltered) {
            subtotal = Math.round((subtotal + (s.subtotal || 0)) * 100) / 100;
            tax = Math.round((tax + (s.tax || 0)) * 100) / 100;
            total = Math.round((total + (s.total || 0)) * 100) / 100;
            shipping = Math.round((shipping + (s.shipping || 0)) * 100) / 100;
            discount = Math.round((discount + (s.discount || 0)) * 100) / 100;
        }
        document.getElementById('ve-cardPretax').textContent = this.veFmt(subtotal);
        document.getElementById('ve-cardTax').textContent = this.veFmt(tax);
        document.getElementById('ve-cardTotal').textContent = this.veFmt(total);
        document.getElementById('ve-cardCount').textContent = `${this._veFiltered.length} transaction${this._veFiltered.length !== 1 ? 's' : ''}`;

        const discountCard = document.getElementById('ve-discountCard');
        const pretaxAfterDiscCard = document.getElementById('ve-pretaxAfterDiscCard');
        if (discount > 0) {
            discountCard.style.display = '';
            document.getElementById('ve-cardDiscount').textContent = '-' + this.veFmt(discount);
            const dc = this._veFiltered.filter(s => s.discount > 0).length;
            document.getElementById('ve-cardDiscountCount').textContent = `${dc} order${dc !== 1 ? 's' : ''} with discounts`;
            const pretaxAfterDisc = subtotal - discount;
            pretaxAfterDiscCard.style.display = '';
            document.getElementById('ve-cardPretaxAfterDisc').textContent = this.veFmt(pretaxAfterDisc);
        } else {
            discountCard.style.display = 'none';
            pretaxAfterDiscCard.style.display = 'none';
        }

        const shippingCard = document.getElementById('ve-shippingCard');
        if (shipping > 0) {
            shippingCard.style.display = '';
            document.getElementById('ve-cardShipping').textContent = this.veFmt(shipping);
            const sc = this._veFiltered.filter(s => s.shipping > 0).length;
            document.getElementById('ve-cardShippingCount').textContent = `${sc} order${sc !== 1 ? 's' : ''} with shipping`;
        } else {
            shippingCard.style.display = 'none';
        }
    },

    veRenderSourceBreakdown() {
        const online = this._veFiltered.filter(s => s.source === 'online');
        const tradeshow = this._veFiltered.filter(s => s.source === 'tradeshow');
        const container = document.getElementById('ve-sourceBreakdown');

        const calcTotals = (arr) => {
            let subtotal = 0, tax = 0, total = 0;
            for (const s of arr) { subtotal += s.subtotal || 0; tax += s.tax || 0; total += s.total || 0; }
            return { subtotal, tax, total, count: arr.length };
        };

        let html = '';
        if (online.length > 0) {
            const t = calcTotals(online);
            html += `<div class="ve-source-card">
                <h4>Online Sales (Store Manager)</h4>
                <div class="ve-source-stat"><span class="ve-stat-label">Transactions</span><span>${t.count}</span></div>
                <div class="ve-source-stat"><span class="ve-stat-label">Pretax</span><span>${this.veFmt(t.subtotal)}</span></div>
                <div class="ve-source-stat"><span class="ve-stat-label">Tax</span><span>${this.veFmt(t.tax)}</span></div>
                <div class="ve-source-stat"><span class="ve-stat-label">Total</span><span>${this.veFmt(t.total)}</span></div>
            </div>`;
        }
        if (tradeshow.length > 0) {
            const t = calcTotals(tradeshow);
            html += `<div class="ve-source-card">
                <h4>Trade Show Sales (POS)</h4>
                <div class="ve-source-stat"><span class="ve-stat-label">Transactions</span><span>${t.count}</span></div>
                <div class="ve-source-stat"><span class="ve-stat-label">Pretax</span><span>${this.veFmt(t.subtotal)}</span></div>
                <div class="ve-source-stat"><span class="ve-stat-label">Tax</span><span>${this.veFmt(t.tax)}</span></div>
                <div class="ve-source-stat"><span class="ve-stat-label">Total</span><span>${this.veFmt(t.total)}</span></div>
            </div>`;
        }
        container.innerHTML = html;
    },

    veRenderProducts() {
        const tbody = document.getElementById('ve-productBody');
        const mappedNames = Database.getMappedVeItemNames();
        const map = new Map();
        let unscrapedSubtotal = 0, unscrapedTax = 0, unscrapedCount = 0;

        for (const s of this._veFiltered) {
            const items = this._veItemsCache.get(s.transaction_no);
            if (items && items.length > 0) {
                for (const item of items) {
                    const price = item.price || 0;
                    const key = (item.name || 'Unknown').toLowerCase().trim() + '|' + price.toFixed(2);
                    if (!map.has(key)) {
                        map.set(key, { name: item.name || 'Unknown', productNumber: item.productNumber || null, price, totalQty: 0, totalRevenue: 0, totalTax: 0, hasInferred: false });
                    }
                    const entry = map.get(key);
                    entry.totalQty += item.quantity || 1;
                    entry.totalRevenue = Math.round((entry.totalRevenue + (item.amount || 0)) * 100) / 100;
                    if ((s.subtotal || 0) > 0) {
                        entry.totalTax = Math.round((entry.totalTax + (s.tax || 0) * ((item.amount || 0) / s.subtotal)) * 100) / 100;
                    }
                    if (item.inferred) entry.hasInferred = true;
                    if (item.productNumber && !entry.productNumber) entry.productNumber = item.productNumber;
                }
            } else {
                unscrapedSubtotal = Math.round((unscrapedSubtotal + (s.subtotal || 0)) * 100) / 100;
                unscrapedTax = Math.round((unscrapedTax + (s.tax || 0)) * 100) / 100;
                unscrapedCount++;
            }
        }

        const products = [...map.values()];
        const dir = this._veProductSortDir === 'asc' ? 1 : -1;
        products.sort((a, b) => {
            switch (this._veProductSortCol) {
                case 'name': return dir * a.name.localeCompare(b.name);
                case 'qty': return dir * (a.totalQty - b.totalQty);
                case 'price': return dir * (a.price - b.price);
                case 'revenue': return dir * (a.totalRevenue - b.totalRevenue);
                case 'tax': return dir * (a.totalTax - b.totalTax);
                default: return dir * (a.totalQty - b.totalQty);
            }
        });

        // Highlight similar products
        const highlightOn = document.getElementById('ve-highlightDupes').checked;
        const colorMap = new Map();
        if (highlightOn) {
            const palette = ['#fce4ec','#e8f5e9','#e3f2fd','#fff3e0','#f3e5f5','#e0f7fa','#fff9c4','#fbe9e7','#e8eaf6','#f1f8e9'];
            // Group by prefix
            const prefixGroups = new Map();
            for (let i = 0; i < products.length; i++) {
                const prefix = products[i].name.toLowerCase().trim().substring(0, 12);
                if (!prefixGroups.has(prefix)) prefixGroups.set(prefix, []);
                prefixGroups.get(prefix).push(i);
            }
            // Re-order: groups (2+ items) first, then singles
            const reordered = [];
            const singles = [];
            for (const [, indices] of prefixGroups) {
                if (indices.length > 1) {
                    for (const i of indices) reordered.push(products[i]);
                } else {
                    singles.push(products[indices[0]]);
                }
            }
            products.splice(0, products.length, ...reordered, ...singles);
            // Rebuild colorMap on new order
            const newPrefixGroups = new Map();
            for (let i = 0; i < products.length; i++) {
                const prefix = products[i].name.toLowerCase().trim().substring(0, 12);
                if (!newPrefixGroups.has(prefix)) newPrefixGroups.set(prefix, []);
                newPrefixGroups.get(prefix).push(i);
            }
            let colorIdx = 0;
            for (const [, indices] of newPrefixGroups) {
                if (indices.length > 1) {
                    const color = palette[colorIdx % palette.length];
                    for (const i of indices) colorMap.set(i, color);
                    colorIdx++;
                }
            }
        }

        let totalQty = 0, totalRev = 0, totalTax = 0;
        let html = '';
        for (let i = 0; i < products.length; i++) {
            const p = products[i];
            const inferBadge  = p.hasInferred ? '<span class="ve-inferred-badge">inferred</span>' : '';
            const linkedBadge = mappedNames.has(p.name + '|' + p.price.toFixed(2)) ? '<span class="ve-linked-badge">linked</span>' : '';
            const prodNum = p.productNumber ? `<br><span class="ve-product-num">${Utils.escapeHtml(p.productNumber)}</span>` : '';
            const bgStyle = colorMap.has(i) ? ` style="background:${colorMap.get(i)}"` : '';
            html += `<tr${bgStyle}>
                <td>${Utils.escapeHtml(p.name)}${prodNum} ${inferBadge}${linkedBadge}</td>
                <td class="num">${p.totalQty}</td>
                <td class="num">${this.veFmt(p.price)}</td>
                <td class="num">${this.veFmt(p.totalRevenue)}</td>
                <td class="num">${this.veFmt(p.totalTax)}</td>
            </tr>`;
            totalQty += p.totalQty;
            totalRev += p.totalRevenue;
            totalTax += p.totalTax;
        }
        if (unscrapedCount > 0) {
            html += `<tr><td><span class="ve-muted">No Item Data (${unscrapedCount})</span></td><td class="num">${unscrapedCount}</td><td class="num"></td><td class="num">${this.veFmt(unscrapedSubtotal)}</td><td class="num">${this.veFmt(unscrapedTax)}</td></tr>`;
            totalQty += unscrapedCount; totalRev += unscrapedSubtotal; totalTax += unscrapedTax;
        }
        html += `<tr class="ve-total-row"><td>Total</td><td class="num">${totalQty}</td><td class="num"></td><td class="num">${this.veFmt(totalRev)}</td><td class="num">${this.veFmt(totalTax)}</td></tr>`;
        tbody.innerHTML = html;
        document.getElementById('ve-productCount').textContent = `(${products.length} products)`;
    },

    veRenderTransactions() {
        const tbody = document.getElementById('ve-txBody');
        document.getElementById('ve-txCount').textContent = `(${this._veFiltered.length})`;

        if (this._veFiltered.length === 0) {
            tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;color:#6c757d;padding:20px;">No transactions match filters</td></tr>';
            return;
        }

        let html = '';
        for (const s of this._veFiltered) {
            const items = this._veItemsCache.get(s.transaction_no);
            let statusHtml = '';
            if (items && items.length > 0) {
                const isInferred = items.some(i => i.inferred);
                statusHtml = isInferred
                    ? '<span class="ve-status-dot inferred"></span>Inferred'
                    : '<span class="ve-status-dot scraped"></span>' + items.length + ' item' + (items.length !== 1 ? 's' : '');
            } else {
                statusHtml = '<span class="ve-status-dot unknown"></span>Unknown';
            }

            const eventName = s.event_id && this._veEventsMap ? (this._veEventsMap.get(s.event_id) || '') : '';
            const eventBadge = eventName ? `<span class="ve-event-badge">${Utils.escapeHtml(eventName)}</span>` : '<span class="ve-no-event">&mdash;</span>';

            html += `<tr>
                <td>${Utils.escapeHtml(s.transaction_no)}</td>
                <td>${this.veFmtDate(s.date)}</td>
                <td><span class="ve-source-badge ${s.source}">${s.source === 'online' ? 'Online' : 'Trade Show'}</span></td>
                <td>${eventBadge}</td>
                <td>${Utils.escapeHtml(s.billing_name || '')}</td>
                <td class="num">${this.veFmt(s.subtotal)}</td>
                <td class="num">${this.veFmt(s.tax)}</td>
                <td class="num">${this.veFmt(s.total)}</td>
                <td>${statusHtml}</td>
            </tr>`;
        }
        tbody.innerHTML = html;
    },

    // --- VE Excel Parsing (client-side via SheetJS) ---

    veFindCol(headers, candidates) {
        for (const name of candidates) {
            if (headers[name] !== undefined) return name;
        }
        return null;
    },

    veParseNumber(val) {
        if (val === null || val === undefined) return 0;
        if (typeof val === 'number') return Math.round(val * 100) / 100;
        const cleaned = String(val).replace(/[$,\s]/g, '');
        const num = parseFloat(cleaned);
        return isNaN(num) ? 0 : Math.round(num * 100) / 100;
    },

    veParseDate(val) {
        if (!val) return null;
        const formatLocal = (dt) => {
            const y = dt.getFullYear();
            const m = String(dt.getMonth() + 1).padStart(2, '0');
            const day = String(dt.getDate()).padStart(2, '0');
            return `${y}-${m}-${day}`;
        };
        if (val instanceof Date) return isNaN(val.getTime()) ? null : formatLocal(val);
        const str = String(val).trim();
        const d = new Date(str);
        return isNaN(d.getTime()) ? null : formatLocal(d);
    },

    veResolveColumns(headers) {
        return {
            transactionNo: this.veFindCol(headers, ['transaction no', 'transaction_no', 'transactionno', 'trans no', 'order no', 'order number']),
            date: this.veFindCol(headers, ['date', 'order date', 'transaction date']),
            billingName: this.veFindCol(headers, ['billing name', 'billing_name', 'customer', 'name']),
            description: this.veFindCol(headers, ['description', 'item', 'product', 'item description', 'product name']),
            subtotal: this.veFindCol(headers, ['subtotal', 'sub total', 'sub-total', 'pretax']),
            tax: this.veFindCol(headers, ['tax', 'sales tax', 'tax amount']),
            shipping: this.veFindCol(headers, ['shipping', 'shipping cost', 'freight']),
            discount: this.veFindCol(headers, ['discount', 'discount amount']),
            total: this.veFindCol(headers, ['total', 'grand total', 'order total', 'amount']),
        };
    },

    veDetectSource(workbook) {
        // Detect if this is a store (online) or tradeshow file by sheet names and content
        const sheetNames = workbook.SheetNames.map(n => n.toLowerCase());
        if (sheetNames.some(n => n.includes('trade show') || n.includes('pos'))) return 'tradeshow';
        if (sheetNames.some(n => n.includes('store') || n.includes('checkout'))) return 'online';
        // Fallback: check first sheet header row for clues
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', range: 0 });
        if (rows.length > 0) {
            const keys = Object.keys(rows[0]).map(k => k.toLowerCase());
            if (keys.some(k => k.includes('pos') || k.includes('trade'))) return 'tradeshow';
        }
        return 'online'; // default
    },

    veParseExcel(arrayBuffer) {
        if (typeof XLSX === 'undefined') throw new Error('SheetJS library not loaded');
        const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
        const source = this.veDetectSource(workbook);

        // Parse main sales sheet (first sheet)
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        if (rows.length === 0) return { sales: [], lineItems: new Map(), source };

        // Build header map (lowercase)
        const sampleRow = rows[0];
        const headers = {};
        for (const key of Object.keys(sampleRow)) {
            headers[key.trim().toLowerCase()] = key;
        }
        const cols = this.veResolveColumns(headers);

        if (!cols.transactionNo || !cols.date) {
            throw new Error('Required columns not found: Transaction No and Date are required. Found headers: ' + Object.keys(sampleRow).join(', '));
        }

        const sales = [];
        for (const row of rows) {
            const txNo = row[headers[cols.transactionNo]];
            const dateRaw = row[headers[cols.date]];
            const total = this.veParseNumber(cols.total ? row[headers[cols.total]] : 0);

            if (!txNo && !dateRaw && !total) continue;

            const date = this.veParseDate(dateRaw);
            if (!date) continue;

            sales.push({
                transactionNo: String(txNo || ''),
                date,
                billingName: String(cols.billingName ? row[headers[cols.billingName]] : '' || ''),
                description: String(cols.description ? row[headers[cols.description]] : '' || ''),
                subtotal: this.veParseNumber(cols.subtotal ? row[headers[cols.subtotal]] : 0),
                tax: this.veParseNumber(cols.tax ? row[headers[cols.tax]] : 0),
                shipping: this.veParseNumber(cols.shipping ? row[headers[cols.shipping]] : 0),
                discount: Math.abs(this.veParseNumber(cols.discount ? row[headers[cols.discount]] : 0)),
                total,
                source,
            });
        }

        // Parse line items from items sheet
        const lineItems = new Map();
        const itemsSheetName = workbook.SheetNames.find(n => {
            const lower = n.toLowerCase();
            return lower.includes('item') || lower.includes('line');
        });

        if (itemsSheetName) {
            const itemSheet = workbook.Sheets[itemsSheetName];
            const itemRows = XLSX.utils.sheet_to_json(itemSheet, { defval: '' });
            if (itemRows.length > 0) {
                const itemHeaders = {};
                for (const key of Object.keys(itemRows[0])) {
                    itemHeaders[key.trim().toLowerCase()] = key;
                }
                const txCol = this.veFindCol(itemHeaders, ['transaction no', 'transaction_no', 'transactionno']);
                const nameCol = this.veFindCol(itemHeaders, ['item name', 'product name', 'name', 'item', 'product']);
                const numCol = this.veFindCol(itemHeaders, ['item number', 'product number', 'item no', 'product no']);
                const priceCol = this.veFindCol(itemHeaders, ['price', 'unit price']);
                const qtyCol = this.veFindCol(itemHeaders, ['quantity', 'qty']);
                const taxableCol = this.veFindCol(itemHeaders, ['taxable']);
                const amountCol = this.veFindCol(itemHeaders, ['amount', 'total', 'line total']);

                if (txCol && nameCol) {
                    for (const row of itemRows) {
                        const txNoVal = String(row[itemHeaders[txCol]] || '').trim();
                        if (!txNoVal) continue;
                        const item = {
                            name: String(row[itemHeaders[nameCol]] || 'Unknown').trim(),
                            productNumber: numCol ? String(row[itemHeaders[numCol]] || '').trim() || null : null,
                            price: this.veParseNumber(priceCol ? row[itemHeaders[priceCol]] : 0),
                            quantity: parseInt(qtyCol ? row[itemHeaders[qtyCol]] : '1', 10) || 1,
                            taxable: taxableCol ? String(row[itemHeaders[taxableCol]] || '').toLowerCase() === 'yes' : false,
                            amount: this.veParseNumber(amountCol ? row[itemHeaders[amountCol]] : 0),
                        };
                        if (!lineItems.has(txNoVal)) lineItems.set(txNoVal, []);
                        lineItems.get(txNoVal).push(item);
                    }
                }
            }
        }

        return { sales, lineItems, source };
    },

    // --- VE Import Handlers ---

    veImportJson(fileContent) {
        let data;
        try {
            data = JSON.parse(fileContent);
        } catch {
            throw new Error('Invalid JSON file');
        }
        if (!data.exportVersion || !Array.isArray(data.sales)) {
            throw new Error('Invalid VE Sales export file. Expected exportVersion and sales array.');
        }

        // Preserve event assignments before clearing
        const eventMap = Database.getVEEventAssignments();

        Database.clearVESales();

        // Normalize dates to YYYY-MM-DD
        const sales = data.sales.map(s => ({
            ...s,
            date: s.date ? s.date.split('T')[0] : s.date,
        }));
        Database.upsertVESales(sales);

        // Restore event assignments
        Database.restoreVEEventAssignments(eventMap);

        // Import line items
        if (data.lineItems && typeof data.lineItems === 'object') {
            for (const [txNo, items] of Object.entries(data.lineItems)) {
                if (Array.isArray(items) && items.length > 0) {
                    Database.upsertVESaleItems(txNo, items);
                }
            }
        }

        Database.setVEImportMeta({
            importDate: new Date().toISOString(),
            source: 'json',
            companyName: data.companyName || 'Unknown',
            salesCount: sales.length,
        });
    },

    async veHandleExcelImport(file, forceSource) {
        const arrayBuffer = await file.arrayBuffer();
        const { sales, lineItems, source } = this.veParseExcel(arrayBuffer);
        const effectiveSource = forceSource || source;

        // Override source on all parsed sales
        for (const s of sales) s.source = effectiveSource;

        // Preserve event assignments before clearing (same pattern as JSON import)
        const eventMap = Database.getVEEventAssignments(effectiveSource);

        // Only clear the specific source being imported
        Database.clearVESales(effectiveSource);
        Database.upsertVESales(sales);

        // Restore event assignments to matching transaction numbers
        Database.restoreVEEventAssignments(eventMap);
        let totalItems = 0;
        for (const [txNo, items] of lineItems) {
            Database.upsertVESaleItems(txNo, items);
            totalItems += items.length;
        }
        return { salesCount: sales.length, itemsCount: totalItems, source: effectiveSource };
    },

    // ---- VE Create Journal Entry ----

    veOpenJournalModal() {
        this._veJournalEventId = null;
        // Populate sales category dropdown (prefer is_sales categories, fall back to all)
        const catSelect = document.getElementById('veJournalCategory');
        let cats = Database.getSalesCategories();
        if (cats.length === 0) {
            cats = Database.getCategories().filter(c => !c.is_sales_tax && !c.is_cogs && !c.is_depreciation && !c.is_inventory_cost);
        }
        catSelect.innerHTML = '';
        if (cats.length === 0) {
            catSelect.innerHTML = '<option value="">No categories found</option>';
            document.getElementById('veJournalSubmitBtn').disabled = true;
        } else {
            document.getElementById('veJournalSubmitBtn').disabled = false;
            for (const cat of cats) {
                const opt = document.createElement('option');
                opt.value = cat.id;
                opt.textContent = cat.name;
                catSelect.appendChild(opt);
            }
        }

        // Default source from current filter
        const currentSource = document.getElementById('ve-filterSource').value;
        const sourceSelect = document.getElementById('veJournalSource');
        if (currentSource === 'online' || currentSource === 'tradeshow') {
            sourceSelect.value = currentSource;
        } else {
            sourceSelect.value = 'both';
        }

        // Default dates from current filter or from data range
        const fromInput = document.getElementById('ve-filterFrom').value;
        const toInput = document.getElementById('ve-filterTo').value;
        if (fromInput) {
            document.getElementById('veJournalDateFrom').value = fromInput;
        } else if (this._veSales.length > 0) {
            const dates = this._veSales.map(s => s.date).filter(Boolean).sort();
            document.getElementById('veJournalDateFrom').value = dates[0] || '';
        }
        if (toInput) {
            document.getElementById('veJournalDateTo').value = toInput;
        } else if (this._veSales.length > 0) {
            const dates = this._veSales.map(s => s.date).filter(Boolean).sort();
            document.getElementById('veJournalDateTo').value = dates[dates.length - 1] || '';
        }

        this.veUpdateJournalPreview();
        UI.showModal('veJournalModal');
    },

    veOpenJournalModalForEvent(eventId) {
        const evt = Database.getAllVEEvents().find(e => e.id === eventId);
        if (!evt) return;

        const eventSales = this._veSales.filter(s => s.event_id === eventId);
        if (eventSales.length === 0) {
            UI.showNotification('No sales in this event.', 'info');
            return;
        }

        // Open the normal journal modal first (populates categories etc.)
        this.veOpenJournalModal();

        // Override the source, dates, and preview with event-specific data
        const sources = new Set(eventSales.map(s => s.source));
        const sourceSelect = document.getElementById('veJournalSource');
        if (sources.size === 1) {
            sourceSelect.value = [...sources][0];
        } else {
            sourceSelect.value = 'both';
        }

        document.getElementById('veJournalDateFrom').value = evt.start_date;
        document.getElementById('veJournalDateTo').value = evt.end_date;

        // Store the event ID so veGetJournalFilteredSales can use it
        this._veJournalEventId = eventId;

        this.veUpdateJournalPreview();
    },

    veGetJournalFilteredSales() {
        const source = document.getElementById('veJournalSource').value;
        const from = document.getElementById('veJournalDateFrom').value;
        const to = document.getElementById('veJournalDateTo').value;

        return this._veSales.filter(s => {
            // When creating journal for a specific event, only include that event's sales
            if (this._veJournalEventId && s.event_id !== this._veJournalEventId) return false;
            if (source !== 'both' && s.source !== source) return false;
            if (from && s.date < from) return false;
            if (to && s.date > to) return false;
            return true;
        });
    },

    veUpdateJournalPreview() {
        const source = document.getElementById('veJournalSource').value;
        const from = document.getElementById('veJournalDateFrom').value;
        const to = document.getElementById('veJournalDateTo').value;
        const filtered = this.veGetJournalFilteredSales();

        let subtotal = 0, total = 0, discount = 0;
        for (const s of filtered) {
            subtotal = Math.round((subtotal + (s.subtotal || 0)) * 100) / 100;
            total = Math.round((total + (s.total || 0)) * 100) / 100;
            discount = Math.round((discount + (s.discount || 0)) * 100) / 100;
        }
        const pretaxAfterDisc = Math.round((subtotal - discount) * 100) / 100;

        // Build description preview
        const sourceLabel = source === 'online' ? 'Online' : source === 'tradeshow' ? 'Tradeshow' : 'Tradeshow + Online';
        let dateLabel = '';
        if (from && to && from === to) {
            dateLabel = this.veFmtDate(from);
        } else if (from && to) {
            dateLabel = this.veFmtDate(from) + ' - ' + this.veFmtDate(to);
        } else if (from) {
            dateLabel = this.veFmtDate(from) + '+';
        } else if (to) {
            dateLabel = 'through ' + this.veFmtDate(to);
        }

        // Use event name when available
        let nameLabel;
        if (this._veJournalEventId) {
            const evt = Database.getAllVEEvents().find(e => e.id === this._veJournalEventId);
            nameLabel = evt ? evt.name : 'Event';
        } else {
            const eventNames = [...new Set(
                filtered.map(s => s.event_id && this._veEventsMap ? this._veEventsMap.get(s.event_id) : null).filter(Boolean)
            )];
            nameLabel = eventNames.length > 0
                ? `${sourceLabel} sales (${eventNames.join(', ')})`
                : `${sourceLabel} sales`;
        }
        const desc = `${nameLabel} ${dateLabel}`.trim();
        document.getElementById('veJournalPreviewDesc').textContent = desc || '--';
        document.getElementById('veJournalPreviewPretax').textContent = this.veFmt(pretaxAfterDisc);
        document.getElementById('veJournalPreviewTotal').textContent = this.veFmt(total);
        document.getElementById('veJournalPreviewCount').textContent = filtered.length;

        // Disable submit if no matching sales
        const submitBtn = document.getElementById('veJournalSubmitBtn');
        const hasCats = document.getElementById('veJournalCategory').value;
        submitBtn.disabled = filtered.length === 0 || !hasCats;
    },

    veCreateJournalEntry() {
        if (this._guardViewOnly()) return;

        const source = document.getElementById('veJournalSource').value;
        const from = document.getElementById('veJournalDateFrom').value;
        const to = document.getElementById('veJournalDateTo').value;
        const categoryId = parseInt(document.getElementById('veJournalCategory').value);
        const filtered = this.veGetJournalFilteredSales();

        if (filtered.length === 0) {
            UI.showNotification('No sales data matches the selected filters', 'error');
            return;
        }

        let subtotal = 0, total = 0, discount = 0;
        for (const s of filtered) {
            subtotal = Math.round((subtotal + (s.subtotal || 0)) * 100) / 100;
            total = Math.round((total + (s.total || 0)) * 100) / 100;
            discount = Math.round((discount + (s.discount || 0)) * 100) / 100;
        }
        const pretaxAfterDisc = Math.round((subtotal - discount) * 100) / 100;

        // Build description
        let dateLabel = '';
        if (from && to && from === to) {
            dateLabel = this.veFmtDate(from);
        } else if (from && to) {
            dateLabel = this.veFmtDate(from) + ' - ' + this.veFmtDate(to);
        } else if (from) {
            dateLabel = this.veFmtDate(from) + '+';
        } else if (to) {
            dateLabel = 'through ' + this.veFmtDate(to);
        }
        let nameLabel;
        if (this._veJournalEventId) {
            const evt = Database.getAllVEEvents().find(e => e.id === this._veJournalEventId);
            nameLabel = evt ? evt.name : 'Event';
        } else {
            // Collect unique event names from filtered sales
            const eventNames = [...new Set(
                filtered.map(s => s.event_id && this._veEventsMap ? this._veEventsMap.get(s.event_id) : null).filter(Boolean)
            )];
            const sourceLabel = source === 'online' ? 'Online' : source === 'tradeshow' ? 'Tradeshow' : 'Tradeshow + Online';
            nameLabel = eventNames.length > 0
                ? `${sourceLabel} sales (${eventNames.join(', ')})`
                : `${sourceLabel} sales`;
        }
        const description = `${nameLabel} ${dateLabel}`.trim();

        // Use the latest date as entry date; received date = last transaction date
        const sortedDates = filtered.map(s => s.date).filter(Boolean).sort();
        const lastTxDate = sortedDates[sortedDates.length - 1] || to || from;
        const entryDate = to || from || Utils.getTodayDate();
        const monthDue = entryDate.substring(0, 7);

        try {
            const txData = {
                entry_date: entryDate,
                category_id: categoryId,
                item_description: description,
                amount: total,
                pretax_amount: pretaxAfterDisc,
                transaction_type: 'receivable',
                status: 'received',
                date_processed: lastTxDate || entryDate,
                month_due: monthDue,
                month_paid: monthDue,
                sale_date_start: from || null,
                sale_date_end: to || null,
            };
            let eventCogs = 0;
            if (this._veJournalEventId) {
                txData.source_type = 've_event';
                txData.source_id = this._veJournalEventId;
                const evt = Database.getAllVEEvents().find(e => e.id === this._veJournalEventId);
                if (evt) eventCogs = evt.cogs || 0;
            }
            const parentId = Database.addTransaction(txData);

            // Auto-manage linked sales tax entry
            const data = {
                entry_date: entryDate,
                category_id: categoryId,
                item_description: description,
                amount: total,
                pretax_amount: pretaxAfterDisc,
                sale_date_start: from || null,
                sale_date_end: to || null,
                month_due: monthDue,
                inventory_cost: eventCogs,
            };
            this._manageSalesTaxEntry(parentId, data);
            this._manageInventoryCostEntry(parentId, data);

            // Mark event as added to journal if created from an event
            if (this._veJournalEventId) {
                Database.markVEEventJournalAdded(this._veJournalEventId);
            }

            UI.hideModal('veJournalModal');
            this.refreshAll();
            UI.showNotification(`Journal entry created: ${description}`, 'success');
        } catch (error) {
            console.error('Error creating VE journal entry:', error);
            UI.showNotification('Failed to create journal entry', 'error');
        }
    },

    veRemoveJournalForEvent(eventId) {
        if (this._guardViewOnly()) return;

        const txId = Database.getVEEventTransaction(eventId);
        if (!txId) {
            // No linked transaction found — just reset the flag
            Database.markVEEventJournalAdded(eventId, 0);
            this.veRenderEventsPanel();
            UI.showNotification('Event unmarked (no linked journal entry found)', 'info');
            return;
        }

        if (!confirm('Remove the journal entry for this event? This will delete the transaction and any linked sales tax entry.')) return;

        try {
            // Cascade-delete linked children
            const childId = Database.getLinkedSalesTaxTransaction(txId);
            if (childId) Database.deleteTransaction(childId);
            const invCostId = Database.getLinkedInventoryCostTransaction(txId);
            if (invCostId) Database.deleteTransaction(invCostId);
            const shipId = Database.getLinkedShippingTransaction(txId);
            if (shipId) Database.deleteTransaction(shipId);

            Database.deleteTransaction(txId);
            Database.markVEEventJournalAdded(eventId, 0);

            this.refreshAll();
            UI.showNotification('Journal entry removed', 'success');
        } catch (error) {
            console.error('Error removing VE journal entry:', error);
            UI.showNotification('Failed to remove journal entry', 'error');
        }
    },

    veAddAllEventsToJournal() {
        if (this._guardViewOnly()) return;

        const events = Database.getAllVEEvents();
        const unadded = events.filter(evt => {
            if (evt.journal_added === 1) return false;
            const salesCount = this._veSales.filter(s => s.event_id === evt.id).length;
            return salesCount > 0;
        });

        if (unadded.length === 0) {
            UI.showNotification('No events to add — all events are already in the journal or have no sales.', 'info');
            return;
        }

        // Get default sales category
        let cats = Database.getSalesCategories();
        if (cats.length === 0) {
            cats = Database.getCategories().filter(c => !c.is_sales_tax && !c.is_cogs && !c.is_depreciation && !c.is_inventory_cost);
        }
        if (cats.length === 0) {
            UI.showNotification('No sales categories found. Please create one first.', 'error');
            return;
        }
        const categoryId = cats[0].id;

        if (!confirm(`Add ${unadded.length} event(s) to journal using category "${cats[0].name}"?`)) return;

        let added = 0;
        for (const evt of unadded) {
            const eventSales = this._veSales.filter(s => s.event_id === evt.id);
            if (eventSales.length === 0) continue;

            let subtotal = 0, total = 0, discount = 0;
            for (const s of eventSales) {
                subtotal = Math.round((subtotal + (s.subtotal || 0)) * 100) / 100;
                total = Math.round((total + (s.total || 0)) * 100) / 100;
                discount = Math.round((discount + (s.discount || 0)) * 100) / 100;
            }
            const pretaxAfterDisc = Math.round((subtotal - discount) * 100) / 100;

            const dateRange = evt.start_date === evt.end_date
                ? this.veFmtDate(evt.start_date)
                : `${this.veFmtDate(evt.start_date)} - ${this.veFmtDate(evt.end_date)}`;
            const description = `${evt.name} ${dateRange}`.trim();

            const sortedDates = eventSales.map(s => s.date).filter(Boolean).sort();
            const lastTxDate = sortedDates[sortedDates.length - 1] || evt.end_date || evt.start_date;
            const entryDate = evt.end_date || evt.start_date;
            const monthDue = entryDate.substring(0, 7);

            try {
                const txData = {
                    entry_date: entryDate,
                    category_id: categoryId,
                    item_description: description,
                    amount: total,
                    pretax_amount: pretaxAfterDisc,
                    transaction_type: 'receivable',
                    status: 'received',
                    date_processed: lastTxDate || entryDate,
                    month_due: monthDue,
                    month_paid: monthDue,
                    sale_date_start: evt.start_date || null,
                    sale_date_end: evt.end_date || null,
                    source_type: 've_event',
                    source_id: evt.id,
                };
                const parentId = Database.addTransaction(txData);

                const data = {
                    entry_date: entryDate,
                    category_id: categoryId,
                    item_description: description,
                    amount: total,
                    pretax_amount: pretaxAfterDisc,
                    sale_date_start: evt.start_date || null,
                    sale_date_end: evt.end_date || null,
                    month_due: monthDue,
                    inventory_cost: evt.cogs || 0,
                };
                this._manageSalesTaxEntry(parentId, data);
                this._manageInventoryCostEntry(parentId, data);

                Database.markVEEventJournalAdded(evt.id);
                added++;
            } catch (error) {
                console.error(`Error adding event "${evt.name}" to journal:`, error);
            }
        }

        this.refreshAll();
        UI.showNotification(`Added ${added} event(s) to journal`, 'success');
    },

    veRemoveAllEventsFromJournal() {
        if (this._guardViewOnly()) return;

        const events = Database.getAllVEEvents();
        const added = events.filter(evt => evt.journal_added === 1);

        if (added.length === 0) {
            UI.showNotification('No events are currently in the journal.', 'info');
            return;
        }

        if (!confirm(`Remove ${added.length} event(s) from journal? This will delete their journal entries and any linked sales tax entries.`)) return;

        let removed = 0;
        for (const evt of added) {
            const txId = Database.getVEEventTransaction(evt.id);
            if (txId) {
                try {
                    const childId = Database.getLinkedSalesTaxTransaction(txId);
                    if (childId) Database.deleteTransaction(childId);
                    const invCostId = Database.getLinkedInventoryCostTransaction(txId);
                    if (invCostId) Database.deleteTransaction(invCostId);
                    const shipChildId = Database.getLinkedShippingTransaction(txId);
                    if (shipChildId) Database.deleteTransaction(shipChildId);
                    Database.deleteTransaction(txId);
                } catch (error) {
                    console.error(`Error removing journal for event "${evt.name}":`, error);
                }
            }
            Database.markVEEventJournalAdded(evt.id, 0);
            removed++;
        }

        this.refreshAll();
        UI.showNotification(`Removed ${removed} event(s) from journal`, 'success');
    },

    // Staged files for the import panel
    _veStagedFiles: { online: null, tradeshow: null, json: null },

    veUpdateImportButton() {
        const btn = document.getElementById('veImportSubmit');
        const { online, tradeshow, json } = this._veStagedFiles;
        btn.disabled = !online && !tradeshow && !json;
    },

    async veSubmitImport() {
        const { online, tradeshow, json } = this._veStagedFiles;
        const btn = document.getElementById('veImportSubmit');
        btn.disabled = true;
        btn.textContent = 'Importing...';

        try {
            if (json) {
                const text = await json.text();
                this.veImportJson(text);
                UI.showNotification('VE Sales JSON imported successfully', 'success');
            } else {
                let totalSales = 0;
                if (online) {
                    const result = await this.veHandleExcelImport(online, 'online');
                    totalSales += result.salesCount;
                }
                if (tradeshow) {
                    const result = await this.veHandleExcelImport(tradeshow, 'tradeshow');
                    totalSales += result.salesCount;
                }
                const parts = [];
                if (online) parts.push('Online');
                if (tradeshow) parts.push('Trade Show');
                UI.showNotification(`Imported ${totalSales} ${parts.join(' + ')} sales successfully`, 'success');
            }

            Database.setVEImportMeta({ importDate: new Date().toISOString(), source: json ? 'json' : 'excel' });

            // Reset staged files and UI
            this._veStagedFiles = { online: null, tradeshow: null, json: null };
            document.getElementById('veOnlineFileName').textContent = '';
            document.getElementById('veTradeshowFileName').textContent = '';
            document.getElementById('veJsonFileName').textContent = '';
            document.getElementById('veOnlineZone').classList.remove('has-file');
            document.getElementById('veTradeshowZone').classList.remove('has-file');
            document.getElementById('veJsonZone').classList.remove('has-file');

            this.refreshVESales();
        } catch (err) {
            UI.showNotification('Import failed: ' + err.message, 'error');
        }

        btn.textContent = 'Import Selected Files';
        this.veUpdateImportButton();
    },

    async veSyncFromServer() {
        const btn = document.getElementById('veSyncFromServer');
        const hint = document.getElementById('veSyncHint');
        btn.disabled = true;
        btn.textContent = 'Connecting...';
        hint.textContent = '';

        try {
            const res = await fetch('http://localhost:3000/api/export', { signal: AbortSignal.timeout(5000) });
            if (!res.ok) throw new Error(`Server returned ${res.status}`);
            const text = await res.text();
            this.veImportJson(text);
            Database.setVEImportMeta({ importDate: new Date().toISOString(), source: 'json-sync' });
            this.refreshVESales();
            UI.showNotification('VE Sales synced from dashboard server', 'success');
            hint.textContent = 'Synced successfully';
            hint.style.color = 'var(--green, #28a745)';
        } catch (err) {
            if (err.name === 'TimeoutError' || err.message.includes('Failed to fetch') || err.message.includes('NetworkError')) {
                hint.textContent = 'Server not running. Start it with start-ve-dashboard.bat';
                hint.style.color = 'var(--red, #dc3545)';
            } else {
                hint.textContent = 'Sync failed: ' + err.message;
                hint.style.color = 'var(--red, #dc3545)';
            }
        }

        btn.disabled = false;
        btn.innerHTML = '<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M17 1l4 4-4 4"/><path d="M3 11V9a4 4 0 0 1 4-4h14"/><path d="M7 23l-4-4 4-4"/><path d="M21 13v2a4 4 0 0 1-4 4H3"/></svg> Sync from VE Dashboard';
    },

    veApplyPreset() {
        const preset = document.getElementById('ve-filterPreset').value;
        const fromEl = document.getElementById('ve-filterFrom');
        const toEl = document.getElementById('ve-filterTo');

        if (!preset) return;
        if (preset === 'all') {
            fromEl.value = '';
            toEl.value = '';
        } else if (/^\d{4}-\d{2}$/.test(preset)) {
            const [y, m] = preset.split('-').map(Number);
            fromEl.value = `${preset}-01`;
            const lastDay = new Date(y, m, 0).getDate();
            toEl.value = `${preset}-${String(lastDay).padStart(2, '0')}`;
        }
        this.veApplyFilters();
    },

    // ==================== VE EVENTS SUB-TAB ====================

    veSwitchSubtab(tab) {
        document.querySelectorAll('.ve-subtab').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.veSubtab === tab);
        });
        document.getElementById('veSubtabSales').style.display = tab === 'sales' ? '' : 'none';
        document.getElementById('veSubtabEvents').style.display = tab === 'events' ? '' : 'none';
        if (tab === 'events') this.veRenderEventsPanel();
    },

    veRenderEventsPanel() {
        const events = Database.getAllVEEvents();
        const list = document.getElementById('veEventsList');
        document.getElementById('veEventsCount').textContent = `(${events.length})`;

        if (events.length === 0) {
            list.innerHTML = '<div class="ve-events-empty">No events yet. Use Auto-Detect or create one manually.</div>';
            return;
        }

        let html = '';
        for (const evt of events) {
            const salesCount = this._veSales.filter(s => s.event_id === evt.id).length;
            const salesTotal = this._veSales.filter(s => s.event_id === evt.id).reduce((sum, s) => sum + (s.total || 0), 0);
            const dateRange = evt.start_date === evt.end_date
                ? this.veFmtDate(evt.start_date)
                : `${this.veFmtDate(evt.start_date)} — ${this.veFmtDate(evt.end_date)}`;

            const isAdded = evt.journal_added === 1;
            const cogs = evt.cogs || 0;
            const grossMargin = salesTotal - cogs;
            const marginPct = salesTotal > 0 ? Math.round((grossMargin / salesTotal) * 100) : 0;
            const cogsPct = salesTotal > 0 ? Math.min(100, Math.round((cogs / salesTotal) * 100)) : 0;
            const marginColor = marginPct >= 50 ? 'var(--green, #28a745)' : marginPct >= 20 ? 'var(--orange, #fd7e14)' : 'var(--red, #dc3545)';
            html += `<div class="ve-event-card${isAdded ? ' ve-event-added' : ''}" data-event-id="${evt.id}">
                <div class="ve-event-card-header">
                    <input class="ve-event-name-input" type="text" value="${Utils.escapeHtml(evt.name)}" data-event-id="${evt.id}" data-field="name">
                    ${isAdded ? '<span class="ve-event-added-badge">In Journal</span>' : ''}
                    <button class="ve-event-edit-toggle" data-event-id="${evt.id}" title="Edit event">&#9998;</button>
                    <button class="ve-event-delete" data-event-id="${evt.id}" title="Delete event">&times;</button>
                </div>
                <div class="ve-event-card-meta">
                    <select class="ve-event-type-select" data-event-id="${evt.id}" data-field="type">
                        <option value="tradeshow"${evt.type === 'tradeshow' ? ' selected' : ''}>Trade Show</option>
                        <option value="online_event"${evt.type === 'online_event' ? ' selected' : ''}>Online Event</option>
                        <option value="custom"${evt.type !== 'tradeshow' && evt.type !== 'online_event' ? ' selected' : ''}>Custom</option>
                    </select>
                    <span class="ve-event-dates">${dateRange}</span>
                </div>
                <div class="ve-event-card-stats">
                    <span>${salesCount} sale${salesCount !== 1 ? 's' : ''}</span>
                    <span>${this.veFmt(salesTotal)}</span>
                </div>
                <div class="ve-margin-visual">
                    <div class="ve-margin-bar">
                        <div class="ve-margin-bar-cogs" style="width:${cogsPct}%;" title="COGS: ${this.veFmt(cogs)}"></div>
                        <div class="ve-margin-bar-margin" style="width:${100 - cogsPct}%; background:${cogs > 0 ? marginColor : 'var(--c5, var(--border))'};" title="Gross Margin: ${this.veFmt(grossMargin)}"></div>
                    </div>
                    <div class="ve-margin-labels">
                        <span class="ve-margin-label-cogs">${cogs > 0 ? `COGS ${this.veFmt(cogs)}` : 'No COGS set'}</span>
                        <span class="ve-margin-label-pct" style="color:${cogs > 0 ? marginColor : 'var(--text-muted)'};">${cogs > 0 ? `${marginPct}% margin` : ''}</span>
                        <span class="ve-margin-label-gm">${cogs > 0 ? `GM ${this.veFmt(grossMargin)}` : ''}</span>
                    </div>
                </div>
                <div class="ve-event-edit-panel" data-event-id="${evt.id}" style="display:none;">
                    <div class="ve-edit-row">
                        <label>Start Date</label>
                        <input type="date" class="ve-edit-start" value="${Utils.escapeHtml(evt.start_date)}" data-event-id="${evt.id}">
                    </div>
                    <div class="ve-edit-row">
                        <label>End Date</label>
                        <input type="date" class="ve-edit-end" value="${Utils.escapeHtml(evt.end_date)}" data-event-id="${evt.id}">
                    </div>
                    <div class="ve-edit-row">
                        <label>COGS / Inventory Cost</label>
                        <input type="number" step="0.01" min="0" class="ve-edit-cogs" value="${Utils.escapeHtml(String(evt.cogs || ''))}" placeholder="0.00" data-event-id="${evt.id}">
                    </div>
                    <div class="ve-edit-row">
                        <label>Notes</label>
                        <textarea class="ve-edit-notes" rows="2" placeholder="Optional notes..." data-event-id="${evt.id}">${Utils.escapeHtml(evt.notes || '')}</textarea>
                    </div>
                    <div class="ve-edit-actions">
                        <button class="btn btn-primary btn-sm ve-edit-save" data-event-id="${evt.id}">Save</button>
                        <button class="btn btn-sm ve-edit-cancel" data-event-id="${evt.id}">Cancel</button>
                    </div>
                </div>
                <div class="ve-event-card-actions">
                    ${isAdded
                        ? `<div class="ve-event-journal-status">
                            <span class="ve-event-added-badge">In Journal</span>
                            <button class="btn btn-sm ve-event-remove-btn" data-event-id="${evt.id}">Remove</button>
                           </div>`
                        : `<button class="btn btn-sm ve-event-journal-btn" data-event-id="${evt.id}" ${salesCount === 0 ? 'disabled' : ''}>+ Add to Journal</button>`
                    }
                </div>
            </div>`;
        }
        list.innerHTML = html;

        // Attach inline edit listeners
        list.querySelectorAll('.ve-event-name-input').forEach(input => {
            input.addEventListener('change', (e) => {
                const id = Number(e.target.dataset.eventId);
                const evt = Database.getAllVEEvents().find(ev => ev.id === id);
                if (evt) {
                    const updated = { ...evt, name: e.target.value };
                    Database.updateVEEvent(id, updated);
                    this._veEventsMap.set(id, e.target.value);
                    this.vePopulateEventDropdown();
                    this.veRenderTransactions();
                    if (evt.journal_added === 1) {
                        this.veUpdateJournalForEvent(id, updated);
                        this.refreshAll();
                    }
                }
            });
        });
        list.querySelectorAll('.ve-event-type-select').forEach(sel => {
            sel.addEventListener('change', (e) => {
                const id = Number(e.target.dataset.eventId);
                const evt = Database.getAllVEEvents().find(ev => ev.id === id);
                if (evt) {
                    const updated = { ...evt, type: e.target.value };
                    Database.updateVEEvent(id, updated);
                }
            });
        });
        // Edit toggle
        list.querySelectorAll('.ve-event-edit-toggle').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = Number(e.target.dataset.eventId);
                const panel = list.querySelector(`.ve-event-edit-panel[data-event-id="${id}"]`);
                if (panel) panel.style.display = panel.style.display === 'none' ? '' : 'none';
            });
        });
        // Edit save
        list.querySelectorAll('.ve-edit-save').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = Number(e.target.dataset.eventId);
                this.veSaveEventEdit(id);
            });
        });
        // Edit cancel
        list.querySelectorAll('.ve-edit-cancel').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = Number(e.target.dataset.eventId);
                const panel = list.querySelector(`.ve-event-edit-panel[data-event-id="${id}"]`);
                if (panel) panel.style.display = 'none';
            });
        });
        list.querySelectorAll('.ve-event-delete').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = Number(e.target.dataset.eventId);
                if (confirm('Delete this event? Sales will be unassigned.')) {
                    Database.deleteVEEvent(id);
                    this.refreshVESales();
                }
            });
        });
        list.querySelectorAll('.ve-event-journal-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = Number(e.target.dataset.eventId);
                this.veOpenJournalModalForEvent(id);
            });
        });
        list.querySelectorAll('.ve-event-remove-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = Number(e.target.dataset.eventId);
                this.veRemoveJournalForEvent(id);
            });
        });
    },

    veSaveEventEdit(eventId) {
        const list = document.getElementById('veEventsList');
        const panel = list.querySelector(`.ve-event-edit-panel[data-event-id="${eventId}"]`);
        if (!panel) return;

        const evt = Database.getAllVEEvents().find(e => e.id === eventId);
        if (!evt) return;

        const startDate = panel.querySelector('.ve-edit-start').value;
        const endDate = panel.querySelector('.ve-edit-end').value || startDate;
        const cogs = parseFloat(panel.querySelector('.ve-edit-cogs').value) || 0;
        const notes = panel.querySelector('.ve-edit-notes').value.trim();

        if (!startDate) {
            UI.showNotification('Start date is required.', 'error');
            return;
        }

        // Also pick up the current name and type from inline inputs
        const card = list.querySelector(`.ve-event-card[data-event-id="${eventId}"]`);
        const name = card.querySelector('.ve-event-name-input').value.trim() || evt.name;
        const type = card.querySelector('.ve-event-type-select').value || evt.type;

        const updated = { name, type, start_date: startDate, end_date: endDate, notes: notes || null, cogs };
        Database.updateVEEvent(eventId, updated);

        // Update events map
        this._veEventsMap.set(eventId, name);
        this.vePopulateEventDropdown();

        // If event is in journal, auto-update the linked transaction
        if (evt.journal_added === 1) {
            this.veUpdateJournalForEvent(eventId, updated);
        }

        this.veRenderEventsPanel();
        this.veRenderTransactions();
        UI.showNotification('Event updated.', 'success');
    },

    veUpdateJournalForEvent(eventId, evt) {
        const txId = Database.getVEEventTransaction(eventId);
        if (!txId) return;

        const tx = Database.getTransactionById(txId);
        if (!tx) return;

        const eventSales = this._veSales.filter(s => s.event_id === eventId);
        let subtotal = 0, total = 0, discount = 0;
        for (const s of eventSales) {
            subtotal = Math.round((subtotal + (s.subtotal || 0)) * 100) / 100;
            total = Math.round((total + (s.total || 0)) * 100) / 100;
            discount = Math.round((discount + (s.discount || 0)) * 100) / 100;
        }
        const pretaxAfterDisc = Math.round((subtotal - discount) * 100) / 100;

        const dateRange = evt.start_date === evt.end_date
            ? this.veFmtDate(evt.start_date)
            : `${this.veFmtDate(evt.start_date)} - ${this.veFmtDate(evt.end_date)}`;
        const description = `${evt.name} ${dateRange}`.trim();

        const sortedDates = eventSales.map(s => s.date).filter(Boolean).sort();
        const lastTxDate = sortedDates[sortedDates.length - 1] || evt.end_date || evt.start_date;
        const entryDate = evt.end_date || evt.start_date;
        const monthDue = entryDate.substring(0, 7);

        const updateData = {
            entry_date: entryDate,
            category_id: tx.category_id,
            item_description: description,
            amount: total,
            pretax_amount: pretaxAfterDisc,
            transaction_type: 'receivable',
            status: 'received',
            date_processed: lastTxDate || entryDate,
            month_due: monthDue,
            month_paid: monthDue,
            sale_date_start: evt.start_date || null,
            sale_date_end: evt.end_date || null,
            source_type: 've_event',
            source_id: eventId,
            inventory_cost: evt.cogs || 0,
        };

        Database.updateTransaction(txId, updateData);

        // Update linked sales tax and inventory cost entries
        this._manageSalesTaxEntry(txId, updateData);
        this._manageInventoryCostEntry(txId, updateData);
    },

    veAutoDetectEvents(mode = 'day') {
        const unassigned = this._veSales.filter(s => !s.event_id && s.date);
        if (unassigned.length === 0) {
            UI.showNotification('No unassigned sales to group into events.', 'info');
            return;
        }

        // Split by source first — online and tradeshow sales never group together
        const bySource = { online: [], tradeshow: [] };
        for (const s of unassigned) {
            (bySource[s.source] || (bySource[s.source] = [])).push(s);
        }

        const clusters = [];
        if (mode === 'month') {
            // Group all sales in the same month together per source
            for (const source of Object.keys(bySource)) {
                const sorted = bySource[source].sort((a, b) => a.date.localeCompare(b.date));
                if (sorted.length === 0) continue;
                const byMonth = {};
                for (const s of sorted) {
                    const monthKey = s.date.substring(0, 7); // "YYYY-MM"
                    if (!byMonth[monthKey]) byMonth[monthKey] = [];
                    byMonth[monthKey].push(s);
                }
                for (const key of Object.keys(byMonth).sort()) {
                    clusters.push(byMonth[key]);
                }
            }
        } else {
            // Cluster consecutive sales within each source where date gap <= 1 day
            for (const source of Object.keys(bySource)) {
                const sorted = bySource[source].sort((a, b) => a.date.localeCompare(b.date));
                if (sorted.length === 0) continue;
                let current = [sorted[0]];
                for (let i = 1; i < sorted.length; i++) {
                    const lastDate = new Date(current[current.length - 1].date + 'T00:00:00');
                    const thisDate = new Date(sorted[i].date + 'T00:00:00');
                    const diffDays = (thisDate - lastDate) / (1000 * 60 * 60 * 24);
                    if (diffDays <= 1) {
                        current.push(sorted[i]);
                    } else {
                        clusters.push([...current]);
                        current = [sorted[i]];
                    }
                }
                clusters.push(current);
            }
        }
        // Sort clusters by start date
        clusters.sort((a, b) => a[0].date.localeCompare(b[0].date));

        // Show preview
        const list = document.getElementById('veEventsList');
        let html = '<div class="ve-autodetect-preview"><h4>Detected Events</h4>';
        html += '<p class="ve-autodetect-hint">Select events to create:</p>';
        for (let i = 0; i < clusters.length; i++) {
            const c = clusters[i];
            const dates = c.map(s => s.date).sort();
            const startDate = dates[0];
            const endDate = dates[dates.length - 1];
            const sources = new Set(c.map(s => s.source));
            const suggestedType = sources.has('tradeshow') ? 'tradeshow' : 'online_event';
            let suggestedName;
            if (mode === 'month') {
                const monthLabel = new Date(startDate + 'T00:00:00').toLocaleString('default', { month: 'long', year: 'numeric' });
                suggestedName = suggestedType === 'tradeshow' ? `Trade Show ${monthLabel}` : `Online Sales ${monthLabel}`;
            } else {
                suggestedName = suggestedType === 'tradeshow' ? `Trade Show ${startDate}` : `Online Event ${startDate}`;
            }
            const dateRange = startDate === endDate ? this.veFmtDate(startDate) : `${this.veFmtDate(startDate)} — ${this.veFmtDate(endDate)}`;
            const total = c.reduce((sum, s) => sum + (s.total || 0), 0);

            html += `<label class="ve-autodetect-item">
                <input type="checkbox" checked data-cluster-idx="${i}">
                <div class="ve-autodetect-item-info">
                    <input type="text" class="ve-autodetect-name" value="${Utils.escapeHtml(suggestedName)}" data-cluster-idx="${i}">
                    <span class="ve-autodetect-details">${dateRange} &middot; ${c.length} sale${c.length !== 1 ? 's' : ''} &middot; ${this.veFmt(total)}</span>
                </div>
            </label>`;
        }
        html += '<div class="ve-autodetect-actions"><button class="btn btn-primary btn-sm" id="veAutoDetectConfirm">Create Selected</button> <button class="btn btn-sm" id="veAutoDetectCancel">Cancel</button></div></div>';
        list.innerHTML = html;

        // Store clusters for confirm handler
        this._veAutoDetectClusters = clusters;

        document.getElementById('veAutoDetectConfirm').addEventListener('click', () => {
            const checkboxes = list.querySelectorAll('.ve-autodetect-item input[type="checkbox"]');
            let created = 0;
            checkboxes.forEach(cb => {
                if (!cb.checked) return;
                const idx = Number(cb.dataset.clusterIdx);
                const cluster = this._veAutoDetectClusters[idx];
                const nameInput = list.querySelector(`.ve-autodetect-name[data-cluster-idx="${idx}"]`);
                const name = nameInput ? nameInput.value : `Event ${idx + 1}`;
                const dates = cluster.map(s => s.date).sort();
                const sources = new Set(cluster.map(s => s.source));
                const type = sources.has('tradeshow') ? 'tradeshow' : 'online_event';

                const eventId = Database.createVEEvent({
                    name,
                    type,
                    start_date: dates[0],
                    end_date: dates[dates.length - 1]
                });
                Database.assignSalesToEvent(eventId, cluster.map(s => s.transaction_no));
                created++;
            });
            UI.showNotification(`Created ${created} event${created !== 1 ? 's' : ''}.`, 'success');
            this.refreshVESales();
        });

        document.getElementById('veAutoDetectCancel').addEventListener('click', () => {
            this.veRenderEventsPanel();
        });
    },

    veShowNewEventForm() {
        const list = document.getElementById('veEventsList');
        const existingForm = list.querySelector('.ve-new-event-form');
        if (existingForm) return;

        const form = document.createElement('div');
        form.className = 've-new-event-form ve-event-card';
        form.innerHTML = `
            <div class="ve-event-card-header">
                <input class="ve-event-name-input" type="text" placeholder="Event name..." id="veNewEventName" autofocus>
            </div>
            <div class="ve-event-card-meta">
                <select id="veNewEventType">
                    <option value="tradeshow">Trade Show</option>
                    <option value="online_event">Online Event</option>
                    <option value="custom">Custom</option>
                </select>
                <input type="date" id="veNewEventStart" style="flex:1;">
                <input type="date" id="veNewEventEnd" style="flex:1;">
            </div>
            <div class="ve-edit-row">
                <label>COGS / Inv. Cost</label>
                <input type="number" step="0.01" min="0" id="veNewEventCogs" placeholder="0.00" style="flex:1;">
            </div>
            <div class="ve-edit-row">
                <label>Notes</label>
                <textarea id="veNewEventNotes" rows="2" placeholder="Optional notes..." style="flex:1;"></textarea>
            </div>
            <div class="ve-autodetect-actions">
                <button class="btn btn-primary btn-sm" id="veNewEventSave">Save</button>
                <button class="btn btn-sm" id="veNewEventCancel">Cancel</button>
            </div>
        `;
        list.insertBefore(form, list.firstChild);

        document.getElementById('veNewEventSave').addEventListener('click', () => {
            const name = document.getElementById('veNewEventName').value.trim();
            const type = document.getElementById('veNewEventType').value;
            const startDate = document.getElementById('veNewEventStart').value;
            const endDate = document.getElementById('veNewEventEnd').value || startDate;
            const cogs = parseFloat(document.getElementById('veNewEventCogs').value) || 0;
            const notes = (document.getElementById('veNewEventNotes').value || '').trim();
            if (!name || !startDate) {
                UI.showNotification('Name and start date are required.', 'error');
                return;
            }
            const eventId = Database.createVEEvent({ name, type, start_date: startDate, end_date: endDate, cogs, notes: notes || null });
            // Auto-assign sales within date range
            const matching = this._veSales.filter(s => !s.event_id && s.date >= startDate && s.date <= endDate);
            if (matching.length > 0) {
                Database.assignSalesToEvent(eventId, matching.map(s => s.transaction_no));
            }
            UI.showNotification(`Event "${name}" created${matching.length > 0 ? ` with ${matching.length} sales` : ''}.`, 'success');
            this.refreshVESales();
        });

        document.getElementById('veNewEventCancel').addEventListener('click', () => {
            this.veRenderEventsPanel();
        });
    },

    setupVESalesListeners() {
        // Filter controls
        document.getElementById('ve-filterSource').addEventListener('change', () => this.veApplyFilters());
        document.getElementById('ve-filterFrom').addEventListener('change', () => this.veApplyFilters());
        document.getElementById('ve-filterTo').addEventListener('change', () => this.veApplyFilters());
        document.getElementById('ve-sortBy').addEventListener('change', () => this.veApplyFilters());
        document.getElementById('ve-filterPreset').addEventListener('change', () => this.veApplyPreset());
        document.getElementById('ve-filterEvent').addEventListener('change', () => this.veApplyFilters());
        document.getElementById('ve-highlightDupes').addEventListener('change', () => this.veRenderProducts());

        // Sub-tabs
        document.querySelectorAll('.ve-subtab').forEach(btn => {
            btn.addEventListener('click', () => this.veSwitchSubtab(btn.dataset.veSubtab));
        });

        // Events panel
        document.getElementById('veAutoDetectBtn').addEventListener('click', (e) => {
            e.stopPropagation();
            const menu = document.getElementById('veAutoDetectMenu');
            const isOpen = menu.style.display !== 'none';
            menu.style.display = isOpen ? 'none' : 'block';
            if (!isOpen) {
                const closeMenu = (ev) => { if (!menu.contains(ev.target)) { menu.style.display = 'none'; document.removeEventListener('click', closeMenu); } };
                setTimeout(() => document.addEventListener('click', closeMenu), 0);
            }
        });
        document.getElementById('veAutoDetectByDay').addEventListener('click', () => { document.getElementById('veAutoDetectMenu').style.display = 'none'; this.veAutoDetectEvents('day'); });
        document.getElementById('veAutoDetectByMonth').addEventListener('click', () => { document.getElementById('veAutoDetectMenu').style.display = 'none'; this.veAutoDetectEvents('month'); });
        document.getElementById('veAddEventBtn').addEventListener('click', () => this.veShowNewEventForm());
        document.getElementById('veAddAllToJournalBtn').addEventListener('click', () => this.veAddAllEventsToJournal());
        document.getElementById('veRemoveAllFromJournalBtn').addEventListener('click', () => this.veRemoveAllEventsFromJournal());

        // Import panel toggle
        document.getElementById('veImportToggle').addEventListener('click', () => {
            document.getElementById('veImportPanel').classList.toggle('open');
        });

        // File staging — Online Excel
        const stageFile = (inputId, zoneId, fileNameId, key) => {
            document.getElementById(zoneId).addEventListener('click', (e) => {
                if (e.target.closest('.ve-drop-btn') || e.target === document.getElementById(zoneId)) {
                    document.getElementById(inputId).click();
                }
            });
            document.getElementById(zoneId).querySelector('.ve-drop-btn').addEventListener('click', (e) => {
                e.stopPropagation();
                document.getElementById(inputId).click();
            });
            document.getElementById(inputId).addEventListener('change', (e) => {
                const file = e.target.files[0];
                if (!file) return;
                this._veStagedFiles[key] = file;
                document.getElementById(fileNameId).textContent = file.name;
                document.getElementById(zoneId).classList.add('has-file');
                // If JSON is staged, clear excel and vice versa
                if (key === 'json') {
                    this._veStagedFiles.online = null;
                    this._veStagedFiles.tradeshow = null;
                    document.getElementById('veOnlineFileName').textContent = '';
                    document.getElementById('veTradeshowFileName').textContent = '';
                    document.getElementById('veOnlineZone').classList.remove('has-file');
                    document.getElementById('veTradeshowZone').classList.remove('has-file');
                    document.getElementById('veOnlineExcelInput').value = '';
                    document.getElementById('veTradeshowExcelInput').value = '';
                } else {
                    this._veStagedFiles.json = null;
                    document.getElementById('veJsonFileName').textContent = '';
                    document.getElementById('veJsonZone').classList.remove('has-file');
                    document.getElementById('veJsonInput').value = '';
                }
                this.veUpdateImportButton();
            });

            // Drag and drop
            const zone = document.getElementById(zoneId);
            zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
            zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
            zone.addEventListener('drop', (e) => {
                e.preventDefault();
                zone.classList.remove('dragover');
                const file = e.dataTransfer.files[0];
                if (!file) return;
                this._veStagedFiles[key] = file;
                document.getElementById(fileNameId).textContent = file.name;
                zone.classList.add('has-file');
                if (key === 'json') {
                    this._veStagedFiles.online = null;
                    this._veStagedFiles.tradeshow = null;
                    document.getElementById('veOnlineFileName').textContent = '';
                    document.getElementById('veTradeshowFileName').textContent = '';
                    document.getElementById('veOnlineZone').classList.remove('has-file');
                    document.getElementById('veTradeshowZone').classList.remove('has-file');
                } else {
                    this._veStagedFiles.json = null;
                    document.getElementById('veJsonFileName').textContent = '';
                    document.getElementById('veJsonZone').classList.remove('has-file');
                }
                this.veUpdateImportButton();
            });
        };

        stageFile('veOnlineExcelInput', 'veOnlineZone', 'veOnlineFileName', 'online');
        stageFile('veTradeshowExcelInput', 'veTradeshowZone', 'veTradeshowFileName', 'tradeshow');
        stageFile('veJsonInput', 'veJsonZone', 'veJsonFileName', 'json');

        // Submit import
        document.getElementById('veImportSubmit').addEventListener('click', () => this.veSubmitImport());

        // Sync from VE Dashboard server
        document.getElementById('veSyncFromServer').addEventListener('click', () => this.veSyncFromServer());

        // Clear data
        document.getElementById('veClearBtn').addEventListener('click', () => {
            if (!confirm('Clear all VE Sales data? This cannot be undone.')) return;
            Database.clearVESales();
            Database.setVEImportMeta(null);
            this.refreshVESales();
            UI.showNotification('VE Sales data cleared', 'success');
        });

        // Create Journal Entry button
        document.getElementById('veCreateJournalBtn').addEventListener('click', () => {
            this.veOpenJournalModal();
        });

        // VE Journal modal events
        document.getElementById('veJournalCancelBtn').addEventListener('click', () => {
            UI.hideModal('veJournalModal');
        });
        document.getElementById('veJournalSource').addEventListener('change', () => this.veUpdateJournalPreview());
        document.getElementById('veJournalDateFrom').addEventListener('change', () => this.veUpdateJournalPreview());
        document.getElementById('veJournalDateTo').addEventListener('change', () => this.veUpdateJournalPreview());
        document.getElementById('veJournalSubmitBtn').addEventListener('click', () => {
            this.veCreateJournalEntry();
        });

        // Product table sorting
        document.querySelectorAll('#ve-productTable th[data-vesort]').forEach(th => {
            th.addEventListener('click', () => {
                const col = th.dataset.vesort;
                if (this._veProductSortCol === col) {
                    this._veProductSortDir = this._veProductSortDir === 'desc' ? 'asc' : 'desc';
                } else {
                    this._veProductSortCol = col;
                    this._veProductSortDir = col === 'name' ? 'asc' : 'desc';
                }
                document.querySelectorAll('#ve-productTable th[data-vesort] .ve-sort-arrow').forEach(el => el.textContent = '');
                th.querySelector('.ve-sort-arrow').textContent = this._veProductSortDir === 'asc' ? '\u25B2' : '\u25BC';
                this.veRenderProducts();
            });
        });
    },

    // ==================== B2B CONTRACT TAB ====================

    /**
     * Compute all B2B contract values from input parameters
     * @param {Object} c - Contract params (monthly_payroll, fiscal_months, bundle_price, cogs_per_unit, contract_start, contract_end)
     * @returns {Object} Computed values
     */
    _computeB2BContract(c) {
        // Derive COGS from gross margin if cost_mode is 'margin'
        let cogsPerUnit = c.cogs_per_unit;
        let grossMargin;
        if (c.cost_mode === 'margin' && c.gross_margin_pct != null) {
            grossMargin = c.gross_margin_pct / 100;
            cogsPerUnit = Math.round((c.bundle_price * (1 - grossMargin)) * 100) / 100;
        } else {
            grossMargin = c.bundle_price > 0 ? (c.bundle_price - cogsPerUnit) / c.bundle_price : 0;
        }

        const totalGrossPayroll = c.monthly_payroll * c.fiscal_months;
        const maxAllowedSales = Math.round(totalGrossPayroll * 0.75 * 100) / 100;
        const maxRevenue = grossMargin > 0 ? Math.round((maxAllowedSales / grossMargin) * 100) / 100 : 0;
        const maxAllowedForContract = maxRevenue;
        const monthlyContractedSales = c.fiscal_months > 0 ? Math.round((maxAllowedForContract / c.fiscal_months) * 100) / 100 : 0;
        const maxUnitsPerMonth = c.bundle_price > 0 ? Math.floor(monthlyContractedSales / c.bundle_price) : 0;

        // Use user-specified units_sold if provided, otherwise use max
        const unitsSold = (c.units_sold && c.units_sold > 0) ? c.units_sold : maxUnitsPerMonth;
        const monthlyContractedRounded = Math.round(unitsSold * c.bundle_price * 100) / 100;
        const monthlyGrossProfit = Math.round(monthlyContractedRounded * grossMargin * 100) / 100;
        const totalContractValue = Math.round(monthlyContractedRounded * c.fiscal_months * 100) / 100;
        const totalGrossProfit = Math.round(monthlyGrossProfit * c.fiscal_months * 100) / 100;
        const isWithinLimit = totalGrossProfit <= maxAllowedSales;
        const months = (c.contract_start && c.contract_end) ? Utils.generateMonthRange(c.contract_start, c.contract_end) : [];

        return {
            grossMargin,
            cogsPerUnit,
            totalGrossPayroll,
            maxAllowedSales,
            maxRevenue,
            maxAllowedForContract,
            monthlyContractedSales,
            maxUnitsPerMonth,
            unitsSold,
            monthlyContractedRounded,
            monthlyGrossProfit,
            totalContractValue,
            totalGrossProfit,
            isWithinLimit,
            months
        };
    },

    /**
     * Sync all finalized B2B contract journal entries
     */
    syncAllB2BContractEntries() {
        try {
            const contracts = Database.getB2BContracts();
            contracts.forEach(contract => {
                if (contract.is_finalized) {
                    this.syncB2BContractJournalEntries(contract);
                }
            });
        } catch (e) {
            console.warn('B2B contract sync skipped:', e.message);
        }
    },

    /**
     * Sync journal entries for a single finalized B2B contract
     * @param {Object} contract - Contract object (or pass id and it will be fetched)
     */
    syncB2BContractJournalEntries(contract) {
        if (typeof contract === 'number') {
            contract = Database.getB2BContractById(contract);
        }
        if (!contract || !contract.is_finalized) {
            console.warn('[B2B Sync] Skipped – contract:', contract?.id, 'is_finalized:', contract?.is_finalized);
            return;
        }

        const computed = this._computeB2BContract(contract);
        const catId = contract.category_id || this._getOrCreateB2BCategory();
        const cogsCatId = Database.getOrCreateInventoryCostCategory();
        const shippingConfig = Database.getShippingFeeConfig();
        const shippingRate = shippingConfig.rate || 0;
        const shippingMinFee = shippingConfig.minFee || 0;
        const currentMonth = Utils.getCurrentMonth();
        const isDirect = contract.entry_mode === 'direct';
        const monthlyRevenue = isDirect ? Math.round((contract.direct_revenue || 0) * 100) / 100 : computed.monthlyContractedRounded;
        const baseCogs = isDirect ? Math.round((contract.direct_cogs || 0) * 100) / 100 : Math.round(computed.cogsPerUnit * computed.unitsSold * 100) / 100;
        // Include shipping fee in COGS (integer-cent math to avoid rounding errors)
        const shippingFee = baseCogs > 0 ? Math.max(Math.round(baseCogs * shippingRate * 100) / 100, shippingMinFee) : 0;
        const monthlyCogs = Math.round((baseCogs + shippingFee) * 100) / 100;
        const revenueDesc = isDirect
            ? `${contract.company_name} – B2B Contract`
            : `${contract.company_name} – B2B Contract (${computed.unitsSold} units)`;
        const cogsDesc = isDirect
            ? `${contract.company_name} – B2B COGS` + (shippingRate > 0 ? ' (incl. shipping)' : '')
            : `${contract.company_name} – B2B COGS (${computed.unitsSold} × ${Utils.formatCurrency(computed.cogsPerUnit)})` + (shippingRate > 0 ? ' + shipping' : '');

        computed.months.forEach(month => {
            // Only create entries up to current month (like budget sync)
            if (month > currentMonth) return;

            // --- Receivable entry (revenue) ---
            const existing = Database.db.exec(
                "SELECT id, amount, status FROM transactions WHERE source_type = 'b2b_contract' AND source_id = ? AND payment_for_month = ? AND transaction_type = 'receivable'",
                [contract.id, month]
            );

            if (existing.length > 0 && existing[0].values.length > 0) {
                const [existingId, existingAmount, existingStatus] = existing[0].values[0];
                if (Math.abs(existingAmount - monthlyRevenue) > 0.01) {
                    // Update amount/description but preserve current status and month_paid
                    Database.db.run(
                        'UPDATE transactions SET amount = ?, item_description = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
                        [monthlyRevenue, revenueDesc, existingId]
                    );
                    Database.autoSave();
                }
            } else {
                Database.addTransaction({
                    entry_date: month + '-01',
                    category_id: catId,
                    item_description: revenueDesc,
                    amount: monthlyRevenue,
                    transaction_type: 'receivable',
                    status: 'pending',
                    month_due: month,
                    payment_for_month: month,
                    source_type: 'b2b_contract',
                    source_id: contract.id
                });
            }

            // --- Payable entry (COGS + shipping, combined) ---
            if (monthlyCogs > 0) {
                const existingCogs = Database.db.exec(
                    "SELECT id, amount, status FROM transactions WHERE source_type = 'b2b_contract_cogs' AND source_id = ? AND payment_for_month = ?",
                    [contract.id, month]
                );

                if (existingCogs.length > 0 && existingCogs[0].values.length > 0) {
                    const [cogsId, cogsAmount] = existingCogs[0].values[0];
                    if (Math.abs(cogsAmount - monthlyCogs) > 0.01) {
                        // Update amount/description but preserve current status and month_paid
                        Database.db.run(
                            'UPDATE transactions SET amount = ?, item_description = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
                            [monthlyCogs, cogsDesc, cogsId]
                        );
                        Database.autoSave();
                    }
                } else {
                    Database.addTransaction({
                        entry_date: month + '-01',
                        category_id: cogsCatId,
                        item_description: cogsDesc,
                        amount: monthlyCogs,
                        transaction_type: 'payable',
                        status: 'pending',
                        month_due: month,
                        payment_for_month: month,
                        source_type: 'b2b_contract_cogs',
                        source_id: contract.id
                    });
                }
            }

            // Clean up any stale separate shipping entries from prior logic
            const staleShip = Database.db.exec(
                "SELECT id FROM transactions WHERE source_type = 'b2b_contract_shipping' AND source_id = ? AND payment_for_month = ? AND status = 'pending'",
                [contract.id, month]
            );
            if (staleShip.length > 0) {
                staleShip[0].values.forEach(row => Database.deleteTransaction(row[0]));
            }
        });
    },

    /**
     * Get or create a B2B revenue category
     * @returns {number} Category ID
     */
    _getOrCreateB2BCategory() {
        const categories = Database.getCategories();
        const existing = categories.find(c => c.is_b2b && c.default_type === 'receivable');
        if (existing) return existing.id;
        // addCategory params: name, isMonthly, defaultAmount, defaultType, folderId, showOnPl, isCogs, isDepreciation, isSalesTax, isB2b, defaultStatus, isSales, isInventoryCost
        return Database.addCategory('B2B Contract Revenue', false, null, 'receivable', null, false, false, false, false, true, 'pending', false, false);
    },

    /**
     * Unfinalize a B2B contract - delete only pending entries
     * @param {number} contractId
     */
    unfinalizeB2BContract(contractId) {
        // Only delete pending entries - preserve received/paid ones
        Database.db.run(
            "DELETE FROM transactions WHERE source_type IN ('b2b_contract', 'b2b_contract_cogs', 'b2b_contract_shipping') AND source_id = ? AND status = 'pending'",
            [contractId]
        );
        Database.setB2BContractFinalized(contractId, false);
        Database.autoSave();
    },

    /**
     * Refresh B2B contracts tab
     */
    refreshB2BContracts() {
        try {
            const contracts = Database.getB2BContracts();
            // Get linked transactions for timeline display
            const contractTransactions = {};
            const contractCogsTransactions = {};
            contracts.forEach(c => {
                const results = Database.db.exec(
                    "SELECT payment_for_month, amount, status, month_paid FROM transactions WHERE source_type = 'b2b_contract' AND source_id = ? ORDER BY payment_for_month",
                    [c.id]
                );
                contractTransactions[c.id] = (results.length > 0) ? Database.rowsToObjects(results[0]) : [];

                const cogsResults = Database.db.exec(
                    "SELECT payment_for_month, amount, status, month_paid FROM transactions WHERE source_type = 'b2b_contract_cogs' AND source_id = ? ORDER BY payment_for_month",
                    [c.id]
                );
                contractCogsTransactions[c.id] = (cogsResults.length > 0) ? Database.rowsToObjects(cogsResults[0]) : [];
            });
            UI.renderB2BContractTab(contracts, this.selectedB2BContractId, contractTransactions, contractCogsTransactions, this);
        } catch (e) {
            console.warn('B2B contract refresh skipped:', e.message);
        }
    },

    /**
     * Open B2B contract modal for add/edit
     * @param {number|null} editId - Contract ID to edit, or null for new
     */
    openB2BContractModal(editId) {
        const form = document.getElementById('b2bContractForm');
        const title = document.getElementById('b2bContractModalTitle');
        const saveBtn = document.getElementById('saveB2BContractBtn');

        form.reset();
        document.getElementById('editingB2BContractId').value = '';

        // Populate category dropdown
        const catSelect = document.getElementById('b2bCategoryId');
        const categories = Database.getCategories();
        const b2bCats = categories.filter(c => c.is_b2b);
        catSelect.innerHTML = '<option value="">Auto-create category</option>' +
            b2bCats.map(c => `<option value="${c.id}">${Utils.escapeHtml(c.name)}</option>`).join('');

        // Clear product rows and direct fields
        document.getElementById('b2bProductsBody').innerHTML = '';
        document.getElementById('b2bDirectRevenue').value = '';
        document.getElementById('b2bDirectCogs').value = '';

        if (editId) {
            const contract = Database.getB2BContractById(editId);
            if (!contract) return;
            title.textContent = 'Edit B2B Contract';
            saveBtn.textContent = 'Update';
            document.getElementById('editingB2BContractId').value = editId;
            document.getElementById('b2bCompanyName').value = contract.company_name;
            document.getElementById('b2bBundleDescription').value = contract.bundle_description || '';
            document.getElementById('b2bContractStart').value = contract.contract_start;
            document.getElementById('b2bContractEnd').value = contract.contract_end;
            document.getElementById('b2bFiscalMonths').value = contract.fiscal_months;
            document.getElementById('b2bMonthlyPayroll').value = contract.monthly_payroll;
            catSelect.value = contract.category_id || '';
            document.getElementById('b2bNotes').value = contract.notes || '';

            // Restore entry mode
            const mode = contract.entry_mode || 'products';
            this._setB2BEntryMode(mode);

            if (mode === 'direct') {
                document.getElementById('b2bDirectRevenue').value = contract.direct_revenue || '';
                document.getElementById('b2bDirectCogs').value = contract.direct_cogs || '';
                this._updateB2BDirectComputed();
            } else {
                // Load product lines
                const products = Database.getB2BContractProducts(editId);
                if (products.length > 0) {
                    products.forEach(p => this._addB2BProductRow(p));
                } else {
                    // Legacy: create a product row from the single-field values
                    this._addB2BProductRow({
                        product_name: contract.bundle_description || 'Product',
                        b2b_price: contract.bundle_price,
                        cogs: contract.cogs_per_unit,
                        quantity: contract.units_sold || 0
                    });
                }
            }
        } else {
            title.textContent = 'Add B2B Contract';
            saveBtn.textContent = 'Save';
            // Default to products mode
            this._setB2BEntryMode('products');
            // Auto-fill from timeline settings
            const timeline = Database.getTimeline();
            if (timeline.start) document.getElementById('b2bContractStart').value = timeline.start;
            if (timeline.end) document.getElementById('b2bContractEnd').value = timeline.end;
            if (timeline.start && timeline.end) {
                const months = Utils.generateMonthRange(timeline.start, timeline.end);
                document.getElementById('b2bFiscalMonths').value = months.length;
            }
            // Start with one empty product row
            this._addB2BProductRow();
        }

        this._updateB2BCalcPreview();
        UI.showModal('b2bContractModal');
    },

    /**
     * Add a product row to the B2B modal products table
     */
    _addB2BProductRow(data) {
        const tbody = document.getElementById('b2bProductsBody');
        const tr = document.createElement('tr');
        const b2bPrice = data && data.b2b_price != null ? data.b2b_price : '';
        const cogs = data && data.cogs != null ? data.cogs : '';
        const qty = data && data.quantity ? data.quantity : '';
        const pm = (b2bPrice && cogs !== '' && b2bPrice > 0) ? (((b2bPrice - cogs) / b2bPrice) * 100).toFixed(2) + '%' : '—';
        const total = (b2bPrice && qty) ? Utils.formatCurrency(b2bPrice * qty) : '—';
        tr.innerHTML = `
            <td><input type="text" class="b2b-p-name" value="${Utils.escapeHtml((data && data.product_name) || '')}" placeholder="Product name"></td>
            <td><input type="number" class="b2b-p-price" step="0.01" min="0" value="${b2bPrice}" placeholder="0.00"></td>
            <td><input type="number" class="b2b-p-cogs" step="0.01" min="0" value="${cogs}" placeholder="0.00"></td>
            <td><span class="b2b-product-computed b2b-p-pm">${pm}</span></td>
            <td><input type="number" class="b2b-p-qty" min="0" value="${qty}" placeholder="0"></td>
            <td><span class="b2b-product-computed b2b-p-total">${total}</span></td>
            <td><button type="button" class="b2b-product-delete" title="Remove">&times;</button></td>
        `;
        tbody.appendChild(tr);

        // Wire up live calc for this row
        const inputs = tr.querySelectorAll('input');
        inputs.forEach(inp => inp.addEventListener('input', () => this._updateB2BProductRow(tr)));
        tr.querySelector('.b2b-product-delete').addEventListener('click', () => {
            tr.remove();
            this._updateB2BCalcPreview();
        });
    },

    /**
     * Recalculate a single product row's derived fields
     */
    _updateB2BProductRow(tr) {
        const b2bPrice = parseFloat(tr.querySelector('.b2b-p-price').value) || 0;
        const cogs = parseFloat(tr.querySelector('.b2b-p-cogs').value) || 0;
        const qty = parseInt(tr.querySelector('.b2b-p-qty').value) || 0;
        const pm = b2bPrice > 0 ? (((b2bPrice - cogs) / b2bPrice) * 100) : 0;
        const total = Math.round(b2bPrice * qty * 100) / 100;

        tr.querySelector('.b2b-p-pm').textContent = b2bPrice > 0 ? pm.toFixed(2) + '%' : '—';
        tr.querySelector('.b2b-p-total').textContent = (b2bPrice > 0 && qty > 0) ? Utils.formatCurrency(total) : '—';

        this._updateB2BCalcPreview();
    },

    /**
     * Toggle between products and direct entry mode in the B2B modal
     */
    _setB2BEntryMode(mode) {
        const productsBtn = document.getElementById('b2bModeProducts');
        const directBtn = document.getElementById('b2bModeDirect');
        const productsSection = document.getElementById('b2bProductsSection');
        const directSection = document.getElementById('b2bDirectSection');

        if (mode === 'direct') {
            productsBtn.classList.remove('active');
            directBtn.classList.add('active');
            productsSection.style.display = 'none';
            directSection.style.display = 'block';
        } else {
            directBtn.classList.remove('active');
            productsBtn.classList.add('active');
            productsSection.style.display = '';
            directSection.style.display = 'none';
        }
        this._updateB2BCalcPreview();
    },

    /**
     * Get the current B2B entry mode from the toggle
     */
    _getB2BEntryMode() {
        return document.getElementById('b2bModeDirect').classList.contains('active') ? 'direct' : 'products';
    },

    /**
     * Update the computed fields in the direct entry section
     */
    _updateB2BDirectComputed() {
        const revenue = parseFloat(document.getElementById('b2bDirectRevenue').value) || 0;
        const cogs = parseFloat(document.getElementById('b2bDirectCogs').value) || 0;
        const profit = Math.round((revenue - cogs) * 100) / 100;
        const pm = revenue > 0 ? ((revenue - cogs) / revenue * 100) : 0;
        document.getElementById('b2bDirectProfit').textContent = revenue > 0 ? Utils.formatCurrency(profit) : '—';
        document.getElementById('b2bDirectPM').textContent = revenue > 0 ? pm.toFixed(2) + '%' : '—';
    },

    /**
     * Get all product rows from the modal as an array of objects
     */
    _getB2BProductRows() {
        const rows = document.querySelectorAll('#b2bProductsBody tr');
        const products = [];
        rows.forEach(tr => {
            const name = tr.querySelector('.b2b-p-name').value.trim();
            const b2bPrice = parseFloat(tr.querySelector('.b2b-p-price').value) || 0;
            const cogs = parseFloat(tr.querySelector('.b2b-p-cogs').value) || 0;
            const qty = parseInt(tr.querySelector('.b2b-p-qty').value) || 0;
            if (b2bPrice > 0 || cogs > 0 || qty > 0 || name) {
                products.push({ product_name: name || 'Unnamed', b2b_price: b2bPrice, cogs, quantity: qty });
            }
        });
        return products;
    },

    /**
     * Update the live calculation preview in the modal
     */
    _updateB2BCalcPreview() {
        const preview = document.getElementById('b2bCalcPreview');
        const monthlyPayroll = parseFloat(document.getElementById('b2bMonthlyPayroll').value) || 0;
        const fiscalMonths = parseInt(document.getElementById('b2bFiscalMonths').value) || 0;
        const contractStart = document.getElementById('b2bContractStart').value;
        const contractEnd = document.getElementById('b2bContractEnd').value;
        const entryMode = this._getB2BEntryMode();

        let monthlyRevenue = 0, monthlyCogs = 0, avgPm = 0, totalQty = 0;

        if (entryMode === 'direct') {
            monthlyRevenue = Math.round((parseFloat(document.getElementById('b2bDirectRevenue').value) || 0) * 100) / 100;
            monthlyCogs = Math.round((parseFloat(document.getElementById('b2bDirectCogs').value) || 0) * 100) / 100;
            avgPm = monthlyRevenue > 0 ? (monthlyRevenue - monthlyCogs) / monthlyRevenue : 0;
        } else {
            const products = this._getB2BProductRows();
            let weightedPmSum = 0, weightedPmDenom = 0;
            products.forEach(p => {
                totalQty += p.quantity;
                monthlyRevenue += Math.round(p.b2b_price * p.quantity * 100) / 100;
                monthlyCogs += Math.round(p.cogs * p.quantity * 100) / 100;
                if (p.b2b_price > 0 && p.quantity > 0) {
                    const pm = (p.b2b_price - p.cogs) / p.b2b_price;
                    weightedPmSum += pm * (p.b2b_price * p.quantity);
                    weightedPmDenom += p.b2b_price * p.quantity;
                }
            });
            avgPm = weightedPmDenom > 0 ? (weightedPmSum / weightedPmDenom) : 0;

            // Update summary footer (products mode only)
            document.getElementById('b2bAvgPM').textContent = products.length > 0 ? (avgPm * 100).toFixed(2) + '%' : '—';
            document.getElementById('b2bTotalQty').textContent = totalQty.toLocaleString();
            document.getElementById('b2bMonthlyTotal').textContent = Utils.formatCurrency(monthlyRevenue);

            if (products.length === 0) monthlyRevenue = 0;
        }

        if (!monthlyPayroll || !fiscalMonths || monthlyRevenue <= 0) {
            const hint = entryMode === 'direct'
                ? 'Enter revenue, COGS, and fill in payroll/fiscal months to see calculations.'
                : 'Add product lines and fill in payroll/fiscal months to see calculations.';
            preview.innerHTML = `<p class="b2b-preview-empty">${hint}</p>`;
            document.getElementById('b2bMaxUnitsHint').textContent = '';
            return;
        }

        // Compute 75% payroll limit
        const monthlyGrossProfit = Math.round((monthlyRevenue - monthlyCogs) * 100) / 100;
        const totalGrossPayroll = monthlyPayroll * fiscalMonths;
        const maxAllowedSales = Math.round(totalGrossPayroll * 0.75 * 100) / 100;
        const totalGrossProfit = Math.round(monthlyGrossProfit * fiscalMonths * 100) / 100;
        const totalContractValue = Math.round(monthlyRevenue * fiscalMonths * 100) / 100;
        const isWithinLimit = totalGrossProfit <= maxAllowedSales;

        const fmt = Utils.formatCurrency;
        const limitClass = isWithinLimit ? 'b2b-limit-ok' : 'b2b-limit-exceeded';
        const limitLabel = isWithinLimit ? 'Within Limit' : 'Exceeds Limit';

        // Show max revenue hint
        const grossMargin = monthlyRevenue > 0 ? (monthlyRevenue - monthlyCogs) / monthlyRevenue : 0;
        const maxRevenueFromLimit = grossMargin > 0 ? Math.round((maxAllowedSales / grossMargin) * 100) / 100 : 0;
        const monthlyGoal = Math.round((maxRevenueFromLimit / fiscalMonths) * 100) / 100;
        document.getElementById('b2bMaxUnitsHint').textContent = maxRevenueFromLimit > 0
            ? `Monthly GOAL: ${fmt(monthlyGoal)} · Yearly GOAL: ${fmt(maxRevenueFromLimit)}`
            : '';

        preview.innerHTML = `
            <div class="b2b-preview-grid">
                <div class="b2b-preview-row">
                    <span>Avg Profit Margin (weighted)</span>
                    <span>${(avgPm * 100).toFixed(2)}%</span>
                </div>
                <div class="b2b-preview-row">
                    <span>Total Gross Payroll</span>
                    <span>${fmt(totalGrossPayroll)}</span>
                </div>
                <div class="b2b-preview-row">
                    <span>Max 75% of Payroll</span>
                    <span>${fmt(maxAllowedSales)}</span>
                </div>
                <div class="b2b-preview-row b2b-preview-divider">
                    <span>Monthly Revenue</span>
                    <span>${fmt(monthlyRevenue)}</span>
                </div>
                <div class="b2b-preview-row">
                    <span>Monthly COGS</span>
                    <span>${fmt(monthlyCogs)}</span>
                </div>
                <div class="b2b-preview-row">
                    <span>Monthly Gross Profit</span>
                    <span>${fmt(monthlyGrossProfit)}</span>
                </div>
                <div class="b2b-preview-row b2b-preview-divider">
                    <span>Total Contract Value (${fiscalMonths} mo)</span>
                    <strong>${fmt(totalContractValue)}</strong>
                </div>
                <div class="b2b-preview-row">
                    <span>Total Gross Profit</span>
                    <span class="${limitClass}">${fmt(totalGrossProfit)} <small>(${limitLabel})</small></span>
                </div>
            </div>
        `;
    },

    /**
     * Handle save B2B contract form submission
     */
    handleSaveB2BContract() {
        try {
            const editId = document.getElementById('editingB2BContractId').value;
            const companyName = document.getElementById('b2bCompanyName').value.trim();
            const bundleDescription = document.getElementById('b2bBundleDescription').value.trim();
            const contractStart = document.getElementById('b2bContractStart').value;
            const contractEnd = document.getElementById('b2bContractEnd').value;
            const fiscalMonths = parseInt(document.getElementById('b2bFiscalMonths').value) || 0;
            const monthlyPayroll = parseFloat(document.getElementById('b2bMonthlyPayroll').value) || 0;
            const categoryId = document.getElementById('b2bCategoryId').value || null;
            const notes = document.getElementById('b2bNotes').value.trim();
            const entryMode = this._getB2BEntryMode();

            // Validation
            if (!companyName) { UI.showNotification('Company name is required', 'error'); return; }
            if (!contractStart || !contractEnd) { UI.showNotification('Contract start and end months are required', 'error'); return; }
            if (contractStart > contractEnd) { UI.showNotification('Contract start must be before end', 'error'); return; }
            if (fiscalMonths <= 0) { UI.showNotification('Fiscal months must be greater than 0', 'error'); return; }
            if (monthlyPayroll <= 0) { UI.showNotification('Monthly payroll must be greater than 0', 'error'); return; }

            let params;

            if (entryMode === 'direct') {
                const directRevenue = Math.round((parseFloat(document.getElementById('b2bDirectRevenue').value) || 0) * 100) / 100;
                const directCogs = Math.round((parseFloat(document.getElementById('b2bDirectCogs').value) || 0) * 100) / 100;
                if (directRevenue <= 0) { UI.showNotification('Monthly revenue must be greater than 0', 'error'); return; }

                params = {
                    company_name: companyName,
                    bundle_description: bundleDescription,
                    contract_start: contractStart,
                    contract_end: contractEnd,
                    fiscal_months: fiscalMonths,
                    monthly_payroll: monthlyPayroll,
                    bundle_price: directRevenue,
                    cogs_per_unit: directCogs,
                    cost_mode: 'direct',
                    gross_margin_pct: null,
                    units_sold: 1,
                    category_id: categoryId,
                    notes: notes,
                    entry_mode: 'direct',
                    direct_revenue: directRevenue,
                    direct_cogs: directCogs
                };
            } else {
                const products = this._getB2BProductRows();
                if (products.length === 0) { UI.showNotification('Add at least one product line', 'error'); return; }

                // Aggregate from products for the contract-level fields
                let totalQty = 0, monthlyRevenue = 0, totalCogs = 0;
                products.forEach(p => {
                    totalQty += p.quantity;
                    monthlyRevenue += Math.round(p.b2b_price * p.quantity * 100) / 100;
                    totalCogs += Math.round(p.cogs * p.quantity * 100) / 100;
                });

                if (monthlyRevenue <= 0) { UI.showNotification('Products must have positive revenue', 'error'); return; }

                // Weighted average price and cogs for backward-compatible storage
                const avgBundlePrice = totalQty > 0 ? Math.round((monthlyRevenue / totalQty) * 100) / 100 : 0;
                const avgCogsPerUnit = totalQty > 0 ? Math.round((totalCogs / totalQty) * 100) / 100 : 0;

                params = {
                    company_name: companyName,
                    bundle_description: bundleDescription,
                    contract_start: contractStart,
                    contract_end: contractEnd,
                    fiscal_months: fiscalMonths,
                    monthly_payroll: monthlyPayroll,
                    bundle_price: avgBundlePrice,
                    cogs_per_unit: avgCogsPerUnit,
                    cost_mode: 'products',
                    gross_margin_pct: null,
                    units_sold: totalQty,
                    category_id: categoryId,
                    notes: notes,
                    entry_mode: 'products',
                    direct_revenue: null,
                    direct_cogs: null
                };
            }

            const products = entryMode === 'products' ? this._getB2BProductRows() : [];

            if (editId) {
                const id = parseInt(editId);
                Database.updateB2BContract(id, params);
                Database.setB2BContractProducts(id, products);
                // If finalized, re-sync entries
                const contract = Database.getB2BContractById(id);
                if (contract && contract.is_finalized) {
                    this.syncB2BContractJournalEntries(contract);
                }
                UI.showNotification('Contract updated', 'success');
            } else {
                const newId = Database.addB2BContract(params);
                Database.setB2BContractProducts(newId, products);
                this.selectedB2BContractId = newId;
                UI.showNotification('Contract added', 'success');
            }

            UI.hideModal('b2bContractModal');
            this.refreshAll();
        } catch (error) {
            console.error('Error saving B2B contract:', error);
            UI.showNotification('Error saving contract: ' + error.message, 'error');
        }
    },

    /**
     * Handle finalize B2B contract
     * @param {number} id
     */
    handleFinalizeB2BContract(id) {
        const contract = Database.getB2BContractById(id);
        if (!contract) return;
        console.log('[B2B Finalize] id:', id, '| contract_start:', contract.contract_start, '| contract_end:', contract.contract_end);

        // Ensure category exists
        if (!contract.category_id) {
            const catId = this._getOrCreateB2BCategory();
            Database.updateB2BContract(id, { ...contract, category_id: catId });
            contract.category_id = catId;
        }

        Database.setB2BContractFinalized(id, true);
        this.syncB2BContractJournalEntries(id);

        // Verify entries were created
        const check = Database.db.exec(
            "SELECT payment_for_month, status, amount FROM transactions WHERE source_type = 'b2b_contract' AND source_id = ? ORDER BY payment_for_month",
            [id]
        );
        console.log('[B2B Finalize] Entries after sync:', check.length > 0 ? check[0].values : 'NONE');

        UI.showNotification('Contract finalized – journal entries created', 'success');
        this.refreshAll();
    },

    /**
     * Handle unfinalize B2B contract
     * @param {number} id
     */
    handleUnfinalizeB2BContract(id) {
        // Check for received/paid entries
        const received = Database.db.exec(
            "SELECT COUNT(*) FROM transactions WHERE source_type IN ('b2b_contract', 'b2b_contract_cogs') AND source_id = ? AND status IN ('received', 'paid')",
            [id]
        );
        const receivedCount = (received.length > 0) ? received[0].values[0][0] : 0;

        // Log entries before unfinalize
        const beforeEntries = Database.db.exec(
            "SELECT payment_for_month, status, source_type FROM transactions WHERE source_type IN ('b2b_contract', 'b2b_contract_cogs') AND source_id = ? ORDER BY payment_for_month",
            [id]
        );
        console.log('[B2B Unfinalize] id:', id, '| entries before:', beforeEntries.length > 0 ? beforeEntries[0].values : 'NONE');

        this.unfinalizeB2BContract(id);

        // Log entries after unfinalize
        const afterEntries = Database.db.exec(
            "SELECT payment_for_month, status, source_type FROM transactions WHERE source_type IN ('b2b_contract', 'b2b_contract_cogs') AND source_id = ? ORDER BY payment_for_month",
            [id]
        );
        console.log('[B2B Unfinalize] entries after:', afterEntries.length > 0 ? afterEntries[0].values : 'NONE (all cleared)');

        let msg = 'Contract unfinalized – pending entries removed';
        if (receivedCount > 0) {
            msg += ` (${receivedCount} received entries preserved)`;
        }
        UI.showNotification(msg, 'success');
        this.refreshAll();
    },

    /**
     * Handle delete B2B contract
     * @param {number} id
     */
    handleDeleteB2BContract(id) {
        this.deleteB2BContractTargetId = id;
        const contract = Database.getB2BContractById(id);
        document.getElementById('deleteB2BContractMessage').textContent =
            `Delete contract "${contract.company_name}"? This will also remove all auto-created journal entries.`;
        UI.showModal('deleteB2BContractModal');
    },

    /**
     * Confirm delete B2B contract
     */
    confirmDeleteB2BContract() {
        if (!this.deleteB2BContractTargetId) return;
        Database.deleteB2BContract(this.deleteB2BContractTargetId);
        if (this.selectedB2BContractId === this.deleteB2BContractTargetId) {
            this.selectedB2BContractId = null;
        }
        this.deleteB2BContractTargetId = null;
        UI.hideModal('deleteB2BContractModal');
        UI.showNotification('Contract deleted', 'success');
        this.refreshAll();
    },

    /**
     * Initialize B2B contract event listeners
     */
    initB2BContractEvents() {
        document.getElementById('addB2BContractBtn').addEventListener('click', () => this.openB2BContractModal());
        document.getElementById('saveB2BContractBtn').addEventListener('click', () => this.handleSaveB2BContract());
        document.getElementById('cancelB2BContractBtn').addEventListener('click', () => UI.hideModal('b2bContractModal'));
        document.getElementById('confirmDeleteB2BContractBtn').addEventListener('click', () => this.confirmDeleteB2BContract());
        document.getElementById('cancelDeleteB2BContractBtn').addEventListener('click', () => {
            UI.hideModal('deleteB2BContractModal');
            this.deleteB2BContractTargetId = null;
        });

        // Entry mode toggle
        document.getElementById('b2bModeProducts').addEventListener('click', () => this._setB2BEntryMode('products'));
        document.getElementById('b2bModeDirect').addEventListener('click', () => this._setB2BEntryMode('direct'));

        // Direct entry live calc
        ['b2bDirectRevenue', 'b2bDirectCogs'].forEach(id => {
            document.getElementById(id).addEventListener('input', () => {
                this._updateB2BDirectComputed();
                this._updateB2BCalcPreview();
            });
        });

        // Live calculation preview
        ['b2bMonthlyPayroll', 'b2bFiscalMonths', 'b2bContractStart', 'b2bContractEnd'].forEach(id => {
            document.getElementById(id).addEventListener('input', () => this._updateB2BCalcPreview());
        });

        // Add product row button
        document.getElementById('addB2BProductRowBtn').addEventListener('click', () => this._addB2BProductRow());

        // Auto-calculate fiscal months when start/end change
        ['b2bContractStart', 'b2bContractEnd'].forEach(id => {
            document.getElementById(id).addEventListener('change', () => {
                const start = document.getElementById('b2bContractStart').value;
                const end = document.getElementById('b2bContractEnd').value;
                if (start && end && start <= end) {
                    const months = Utils.generateMonthRange(start, end);
                    document.getElementById('b2bFiscalMonths').value = months.length;
                    this._updateB2BCalcPreview();
                }
            });
        });

        // List panel click delegation
        document.getElementById('b2bListPanel').addEventListener('click', (e) => {
            const item = e.target.closest('.b2b-list-item');
            if (item) {
                this.selectedB2BContractId = parseInt(item.dataset.id);
                this.refreshB2BContracts();
            }
        });

        // Detail panel button delegation
        document.getElementById('b2bDetailPanel').addEventListener('click', (e) => {
            const btn = e.target.closest('button');
            if (!btn) return;
            const id = parseInt(btn.dataset.id);
            if (btn.classList.contains('b2b-edit-btn')) {
                this.openB2BContractModal(id);
            } else if (btn.classList.contains('b2b-delete-btn')) {
                this.handleDeleteB2BContract(id);
            } else if (btn.classList.contains('b2b-finalize-btn')) {
                this.handleFinalizeB2BContract(id);
            } else if (btn.classList.contains('b2b-unfinalize-btn')) {
                this.handleUnfinalizeB2BContract(id);
            }
        });
    },

    // ==================== CHANGE LOG ====================

    changelogData: [
        {
            version: '2.2',
            date: '2026-03-24',
            title: 'KPI Dashboard & Financial Metrics',
            changes: [
                { type: 'feature', text: 'Interactive KPI dashboard — clickable section headers open analysis modals with grouped financial metrics' },
                { type: 'feature', text: 'KPI drill-down detail modals — click any KPI card to see monthly breakdowns, trends, and explanatory tables' },
                { type: 'feature', text: '12 KPI metrics: Cash Position, Gross Burn, Net Burn, Revenue Trend, EBITDA, CMGR, Non-B2B CMGR, Rule of 40, Rule of 40 (Non-B2B), Working Capital, Overdue AR, DSCR' },
                { type: 'feature', text: 'Revenue Concentration pie chart on dashboard showing B2B vs consumer revenue mix' },
                { type: 'feature', text: 'AR Aging breakdown with 30/60/90/90+ day overdue buckets' },
                { type: 'feature', text: 'Shipping Fee auto-entries — configurable rate applied to inventory cost payments, with gear popover setting' },
                { type: 'feature', text: 'Cash Ratio added to Balance Sheet financial ratios (Cash / Current Liabilities)' },
                { type: 'improved', text: 'Dashboard KPIs now use accrual-basis P&L data for burn rate and expense metrics instead of cash-basis' },
                { type: 'improved', text: 'Revenue metrics use pretax amounts for consistent accrual accounting' },
                { type: 'improved', text: 'Sparklines exclude future months for accurate trailing data' },
                { type: 'removed', text: 'Removed Post-Tax After Discounts metric from VE analytics (redundant with pretax reporting)' }
            ]
        },
        {
            version: '2.1',
            date: '2026-03-15',
            title: 'B2B Contracts & Budget Overrides',
            changes: [
                { type: 'feature', text: 'B2B Contracts tab — manage wholesale contracts with 75% payroll profit limit, COGS tracking, payment timelines, and auto-generated journal entries' },
                { type: 'feature', text: 'Budget per-month overrides — click any monthly amount in budget detail to set a custom value for that month' },
                { type: 'feature', text: 'VE Events: inline edit panel for event details, COGS tracking, and gross margin bar visualization' },
                { type: 'feature', text: 'VE Sales: bulk "Add All to Journal" and "Remove All from Journal" buttons for events' },
                { type: 'improved', text: 'Dashboard and Analyze charts now use P&L data with overrides for accurate revenue/expense totals' },
                { type: 'improved', text: 'Analyze Balance Sheet chart includes fixed assets, loan balances, and sales tax payable' },
                { type: 'improved', text: 'Financial ratios use proper current assets/liabilities breakdown' },
                { type: 'fix', text: 'Filter month dropdown now retains selection when data refreshes' },
                { type: 'fix', text: 'Auto-generated budget journal entries no longer show edit/duplicate buttons' }
            ]
        },
        {
            version: '2.0',
            date: '2026-03-14',
            title: 'Design Polish & Data Integrity',
            changes: [
                { type: 'feature', text: 'Rich empty states with SVG icons, descriptions, and call-to-action buttons across all tabs' },
                { type: 'feature', text: 'Enhanced notifications with type-specific icons and animated progress bar' },
                { type: 'feature', text: 'Item description labels shown inline on journal entry rows' },
                { type: 'feature', text: 'Sales tax entries now include parent item name in description' },
                { type: 'improved', text: 'CSS transitions unified via design token variables (--transition-fast)' },
                { type: 'improved', text: 'Input fields refined with softer borders, subtle shadows, and consistent focus rings' },
                { type: 'improved', text: 'Loan payment summary shows most common (mode) payment amount instead of first payment' },
                { type: 'improved', text: 'Budget expenses from loans now respect skip and override schedules' },
                { type: 'fix', text: 'Fixed dark mode not overriding inline theme colors for background, surface, border, and text' },
                { type: 'fix', text: 'Loan deletion now properly cleans up all associated budget journal entries' }
            ]
        },
        {
            version: '1.9',
            date: '2026-03-14',
            title: 'Dashboard, Quick Entry & Undo System',
            changes: [
                { type: 'feature', text: 'Added Work / Analyze mode toggle in sidebar — separates data entry tabs from reporting tabs' },
                { type: 'feature', text: 'New Dashboard tab with KPI cards, sparklines, and break-even progress bar' },
                { type: 'feature', text: 'Quick Entry modal — batch-add multiple journal entries in a spreadsheet-style grid' },
                { type: 'feature', text: 'Undo/Redo system with toast notifications and Ctrl+Z / Ctrl+Y support' },
                { type: 'feature', text: 'Analyze mode chart panels for Cash Flow, P&L, Balance Sheet, and Break-Even tabs' },
                { type: 'feature', text: 'Interactive tutorial lessons with guided input and impact visualization' },
                { type: 'feature', text: 'VE Events: journal_added tracking for event-to-journal linking' },
                { type: 'improved', text: 'Sidebar auto-hides group labels when all tabs in a group are hidden' },
                { type: 'fix', text: 'Fixed Chart.js legend point style rendering (circle style with consistent sizing)' }
            ]
        },
        {
            version: '1.8',
            date: '2026-03-13',
            title: 'Change Log',
            changes: [
                { type: 'feature', text: 'Added Change Log page — view all versions, features, fixes, and changes in a timeline' },
                { type: 'feature', text: 'New "Info" section in sidebar with Change Log tab' },
                { type: 'feature', text: 'Tagged change items (New, Fix, Removed) with color-coded badges' },
                { type: 'feature', text: 'Dark mode support for changelog styles' }
            ]
        },
        {
            version: '1.7',
            date: '2026-03-12',
            title: 'Precision & Connectivity',
            changes: [
                { type: 'fix', text: 'Fixed floating-point rounding — all monetary values now round to 2 decimal places across parsers, exports, and UI aggregations' },
                { type: 'feature', text: 'Added "Sync from VE Dashboard" button for direct localhost data import without file upload' },
                { type: 'feature', text: 'Added CORS middleware to Express server for cross-origin requests' }
            ]
        },
        {
            version: '1.6',
            date: '2026-03-12',
            title: 'UI Redesign',
            changes: [
                { type: 'feature', text: 'Replaced top header with collapsible sidebar navigation (Zoho/Wave hybrid design)' },
                { type: 'feature', text: 'Added interactive tutorial system with guided walkthrough' },
                { type: 'feature', text: 'Extended theme presets with sidebar, border, and text color channels' },
                { type: 'feature', text: 'Complete CSS rewrite with new design tokens' }
            ]
        },
        {
            version: '1.5',
            date: '2026-03-11',
            title: 'Tab Management & VE Journal Entries',
            changes: [
                { type: 'feature', text: 'Added tab context menu — right-click to hide/reset tabs' },
                { type: 'feature', text: 'Added manage tabs modal with gear button for tab visibility control' },
                { type: 'feature', text: 'VE Sales: create journal entries from filtered sales data with preview' },
                { type: 'feature', text: 'VE Sales: after-discount summary cards (pre-tax and post-tax)' },
                { type: 'feature', text: 'Products: source filter dropdown and linked row highlighting' },
                { type: 'feature', text: 'Product-VE mapping: suggest button for price-based matching' },
                { type: 'feature', text: 'Product-VE mapping: show product price in modal title' },
                { type: 'fix', text: 'Fixed date parsing in VE date formatter to avoid timezone offset issues' },
                { type: 'removed', text: 'Removed unused product tax rate field' }
            ]
        },
        {
            version: '1.4',
            date: '2026-03-11',
            title: 'Export Tools & Product Catalog Enhancements',
            changes: [
                { type: 'feature', text: 'Added Excel and HTML export tools for UpdateAccountsFlow with filtered data export' },
                { type: 'feature', text: 'Added company name autocomplete' },
                { type: 'feature', text: 'Improved product VE mapping display' },
                { type: 'feature', text: 'Save All workflow refinements' },
                { type: 'feature', text: 'Siply products CSV import support' }
            ]
        },
        {
            version: '1.3',
            date: '2026-03-11',
            title: 'Save All & Price-Based Mappings',
            changes: [
                { type: 'feature', text: 'Added Save All button — exports all company databases as a zip file with file picker' },
                { type: 'feature', text: 'Added price-aware product-VE linking (ve_item_price in mappings)' },
                { type: 'feature', text: 'Mappings now join on both name and price for accurate sales attribution' },
                { type: 'feature', text: 'Show prices next to VE item names in mapping modal' },
                { type: 'removed', text: 'Replaced Save As with Save — now always prompts for file location' },
                { type: 'removed', text: 'Removed tax rate column from product catalog table' }
            ]
        },
        {
            version: '1.2',
            date: '2026-03-10',
            title: 'Product Catalog & Companies',
            changes: [
                { type: 'feature', text: 'Added product catalog with full CRUD, SKU tracking, cost/price fields, and notes' },
                { type: 'feature', text: 'Added company management module with company switcher in sidebar' },
                { type: 'feature', text: 'Icon-only action buttons for products (edit, discontinue, reactivate, delete)' },
                { type: 'feature', text: 'Product CSV import with drag-and-drop file upload' },
                { type: 'feature', text: 'Inline notes tooltip indicator on product names' },
                { type: 'feature', text: 'Added example database (Siply Accounting Journal)' }
            ]
        },
        {
            version: '1.1',
            date: '2026-03-10',
            title: 'VE Sales Dashboard',
            changes: [
                { type: 'feature', text: 'Added VE Sales Dashboard with smart scraping and product breakdown' },
                { type: 'feature', text: 'Extract product data from Excel line items instead of web scraping' },
                { type: 'feature', text: 'Added discount and shipping summary cards that auto-show when data exists' },
                { type: 'feature', text: 'Added pretax after-discounts summary card' },
                { type: 'fix', text: 'Fixed Excel parsers to handle negative discount values with Math.abs()' }
            ]
        },
        {
            version: '1.0',
            date: '2026-03-10',
            title: 'Initial Release',
            changes: [
                { type: 'feature', text: 'Journal entry management with debit/credit tracking' },
                { type: 'feature', text: 'Cash Flow statement generation' },
                { type: 'feature', text: 'Profit & Loss report' },
                { type: 'feature', text: 'Balance Sheet report' },
                { type: 'feature', text: 'Assets & Equity tracking' },
                { type: 'feature', text: 'Loan management' },
                { type: 'feature', text: 'Budget planning' },
                { type: 'feature', text: 'Break-Even analysis' },
                { type: 'feature', text: 'Projected Sales forecasting' },
                { type: 'feature', text: 'SQLite database with browser persistence' },
                { type: 'feature', text: 'Theme system with multiple presets and dark mode' },
                { type: 'feature', text: 'Timeline-based filtering' }
            ]
        }
    ],

    quickGuideSteps: [
        {
            step: 1, title: 'Equity', tab: 'assets', tabLabel: 'Assets & Equity',
            description: 'Set up your seed money and additional paid-in capital (APIC).',
            auto: 'Auto-creates investment receivable entries in the Journal.'
        },
        {
            step: 2, title: 'Loans', tab: 'loan', tabLabel: 'Loans',
            description: 'Add any business loans with their terms and payment schedules.',
            auto: 'Auto-creates a loan receivable entry and a budget expense for monthly payments, which then auto-creates monthly payable entries in the Journal.'
        },
        {
            step: 3, title: 'Fixed Assets', tab: 'assets', tabLabel: 'Assets & Equity',
            description: 'Add equipment, property, or other fixed assets.',
            auto: 'Optionally auto-creates an asset purchase entry in the Journal.'
        },
        {
            step: 4, title: 'Budget', tab: 'budget', tabLabel: 'Budget',
            description: 'Add recurring monthly expenses like rent, utilities, and subscriptions. Don\'t add these to the Journal manually\u2014they\'re auto-filled!',
            auto: 'Auto-creates monthly payable entries in the Journal for each budget expense.'
        },
        {
            step: 5, title: 'Products', tab: 'products', tabLabel: 'Products',
            description: 'Set up your product catalog with prices and inventory costs before recording sales.',
            auto: null
        },
        {
            step: 6, title: 'Sales', tab: 'vesales', tabLabel: 'VE Sales',
            description: 'Record sales transactions using the VE Sales tab. Sales tax and cost of goods sold are handled for you.',
            auto: 'Auto-creates sales tax payable and inventory cost (COGS) entries in the Journal.'
        },
        {
            step: 7, title: 'Journal', tab: 'journal', tabLabel: 'Journal',
            description: 'Use the Journal only for one-time or manual entries not covered above\u2014things like one-off expenses, miscellaneous income, or adjustments.',
            auto: null
        }
    ],

    quickGuideTips: [
        'Budget items are auto-synced monthly\u2014don\'t duplicate them in the Journal.',
        'Loan payments auto-appear via budget sync\u2014no manual journal entries needed.',
        'Use the Sales workflow for revenue\u2014it handles tax and COGS automatically.',
        'The Journal tab is for things that don\'t fit elsewhere.',
        'You can always check the Journal tab to see all auto-created entries.'
    ],

    renderQuickGuide() {
        const stepsContainer = document.getElementById('quickguideSteps');
        const tipsContainer = document.getElementById('quickguideTips');
        if (!stepsContainer) return;

        stepsContainer.innerHTML = this.quickGuideSteps.map(s => {
            const autoHtml = s.auto
                ? '<div class="quickguide-auto"><span class="quickguide-auto-badge">Auto</span><span class="quickguide-auto-text">' + s.auto + '</span></div>'
                : '';
            return '<div class="quickguide-step">' +
                '<div class="quickguide-step-header">' +
                    '<div class="quickguide-step-number">' + s.step + '</div>' +
                    '<div class="quickguide-step-meta">' +
                        '<h3 class="quickguide-step-title">' + s.title + '</h3>' +
                        '<button class="quickguide-tab-link" onclick="App.switchMainTab(\'' + s.tab + '\')">' + s.tabLabel + ' &rarr;</button>' +
                    '</div>' +
                '</div>' +
                '<p class="quickguide-step-desc">' + s.description + '</p>' +
                autoHtml +
            '</div>';
        }).join('');

        if (tipsContainer) {
            tipsContainer.innerHTML = '<div class="quickguide-tips-card">' +
                '<h3 class="quickguide-tips-title">Key Tips</h3>' +
                '<ul class="quickguide-tips-list">' +
                this.quickGuideTips.map(t => '<li class="quickguide-tip-item">' + t + '</li>').join('') +
                '</ul></div>';
        }
    },

    renderChangelog() {
        const container = document.getElementById('changelogTimeline');
        if (!container) return;

        const typeConfig = {
            feature: { label: 'New', cls: 'changelog-tag-feature' },
            fix: { label: 'Fix', cls: 'changelog-tag-fix' },
            removed: { label: 'Removed', cls: 'changelog-tag-removed' },
            improved: { label: 'Improved', cls: 'changelog-tag-improved' }
        };

        container.innerHTML = this.changelogData.map((release, i) => {
            const changesHtml = release.changes.map(c => {
                const cfg = typeConfig[c.type] || typeConfig.feature;
                return '<li class="changelog-item"><span class="changelog-tag ' + cfg.cls + '">' + cfg.label + '</span><span class="changelog-text">' + c.text + '</span></li>';
            }).join('');

            return '<div class="changelog-release' + (i === 0 ? ' changelog-latest' : '') + '">' +
                '<div class="changelog-release-header">' +
                    '<div class="changelog-version-badge">v' + release.version + '</div>' +
                    '<div class="changelog-release-meta">' +
                        '<h3 class="changelog-release-title">' + release.title + '</h3>' +
                        '<span class="changelog-date">' + release.date + '</span>' +
                    '</div>' +
                    (i === 0 ? '<span class="changelog-latest-badge">Latest</span>' : '') +
                '</div>' +
                '<ul class="changelog-changes">' + changesHtml + '</ul>' +
            '</div>';
        }).join('');
    }
};

// Initialize application when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    App.init();
});

// Export for use in other modules
window.App = App;
