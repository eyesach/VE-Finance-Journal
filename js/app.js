/**
 * Main application logic for the Accounting Journal Calculator
 */

const App = {
    deleteTargetId: null,
    deleteCategoryTargetId: null,
    deleteFolderTargetId: null,
    deleteAssetTargetId: null,
    deleteLoanTargetId: null,
    deleteBudgetExpenseTargetId: null,
    deleteProductTargetId: null,
    selectedAssetId: null,
    selectedLoanId: null,
    selectedBudgetExpenseId: null,
    folderCreatedFromCategory: false,
    pendingFileLoad: null,
    pendingLoadBuffer: null,
    pendingInlineStatusChange: null, // {id, newStatus, selectElement}
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

    // Theme preset palettes: { c1: primary, c2: accent, c3: background, c4: surface, style?: string }
    themePresets: {
        // Color-only themes
        default:  { c1: '#4a90a4', c2: '#e8f0f3', c3: '#f8f9fa', c4: '#ffffff' },
        ocean:    { c1: '#9EBAC2', c2: '#BAE0EB', c3: '#f0f7fa', c4: '#ffffff' },
        forest:   { c1: '#2d6a4f', c2: '#d8f3dc', c3: '#f0f7f0', c4: '#ffffff' },
        sunset:   { c1: '#e76f51', c2: '#fce4d6', c3: '#fdf8f4', c4: '#ffffff' },
        midnight: { c1: '#6c63ff', c2: '#e8e6ff', c3: '#f5f4ff', c4: '#ffffff' },
        // Extreme design styles (CSS handles font, radius, shadows via data-theme-style)
        modern:     { c1: '#0f172a', c2: '#f1f5f9', c3: '#f8fafc', c4: '#ffffff', style: 'modern' },
        futuristic: { c1: '#00d4ff', c2: '#0f1729', c3: '#080e1a', c4: '#111827', style: 'futuristic' },
        vintage:    { c1: '#7c5a3c', c2: '#f0e6d6', c3: '#faf6f0', c4: '#fffdf8', style: 'vintage' },
        accounting: { c1: '#1e3a5f', c2: '#f0f2f5', c3: '#ffffff', c4: '#ffffff', style: 'accounting' },
    },

    /**
     * Initialize the application
     */
    isViewOnly: false,

    async init() {
        try {
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

            // Restore tab order and set up tab drag-drop
            this.restoreTabOrder();
            this.setupTabDragDrop();
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
        this.syncAllLoanJournalEntries();
        this.refreshCategories();
        this.refreshTransactions();
        this.refreshSummary();
        this.refreshCashFlow();
        this.refreshPnL();
        this.refreshBalanceSheet();
        this.refreshFixedAssets();
        this.refreshLoans();
        this.refreshBudget();
        this.refreshBreakeven();
        this.refreshProjectedSales();
        this.refreshProducts();
        this.refreshVESales();
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

        nav.querySelectorAll('.main-tab').forEach(btn => {
            btn.setAttribute('draggable', 'true');
        });

        nav.addEventListener('dragstart', (e) => {
            const tab = e.target.closest('.main-tab');
            if (!tab) return;
            draggedTab = tab;
            tab.classList.add('tab-dragging');
            e.dataTransfer.effectAllowed = 'move';
            e.dataTransfer.setData('text/plain', tab.dataset.tab);
        });

        nav.addEventListener('dragover', (e) => {
            e.preventDefault();
            const tab = e.target.closest('.main-tab');
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
            const targetTab = e.target.closest('.main-tab');
            if (!targetTab || !draggedTab || targetTab === draggedTab) return;

            nav.insertBefore(draggedTab, targetTab);
            targetTab.classList.remove('tab-drag-over');

            // Save new order
            const tabs = nav.querySelectorAll('.main-tab');
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
        nav.querySelectorAll('.main-tab').forEach(tab => {
            if (!order.includes(tab.dataset.tab)) {
                nav.appendChild(tab);
            }
        });
    },

    /**
     * Set up inline cell editing for Cash Flow projected cells (only binds once)
     */
    setupCashFlowCellEditing() {
        const container = document.getElementById('cashflowSpreadsheet');
        if (!container || container.dataset.cfCellSetup) return;
        container.dataset.cfCellSetup = '1';

        container.addEventListener('click', (e) => {
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
        this.setupPnLCellEditing();
    },

    /**
     * Set up inline cell editing for P&L spreadsheet (only binds once)
     */
    setupPnLCellEditing() {
        const container = document.getElementById('pnlSpreadsheet');
        if (!container || container.dataset.pnlEditing) return;
        container.dataset.pnlEditing = '1';

        container.addEventListener('click', (e) => {
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
        document.querySelectorAll('.main-tab').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tab === tab);
        });

        const tabs = ['journalTab', 'cashflowTab', 'pnlTab', 'balancesheetTab', 'assetsTab', 'loanTab', 'budgetTab', 'breakevenTab', 'projectedsalesTab', 'productsTab', 'vesalesTab'];
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
        } else {
            document.getElementById('journalTab').style.display = 'block';
        }
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
            ['themeC1', 'themeC2', 'themeC3', 'themeC4'].forEach((id, i) => {
                const input = document.getElementById(id);
                if (input) input.value = customColors[`c${i + 1}`] || '#000000';
            });
        }

        this.applyTheme(preset, customColors);
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

        // Primary and derived
        root.style.setProperty('--color-primary', colors.c1);
        root.style.setProperty('--color-primary-hover', Utils.adjustLightness(colors.c1, -12));
        root.style.setProperty('--color-primary-light', Utils.adjustLightness(colors.c1, -6));
        root.style.setProperty('--color-primary-rgb', Utils.hexToRGBString(colors.c1));

        // Accent (section headers, hover tints)
        root.style.setProperty('--color-accent-bg', colors.c2);
        root.style.setProperty('--color-accent-bg-hover', Utils.adjustLightness(colors.c2, -5));

        // Background and surface
        root.style.setProperty('--color-bg', colors.c3);
        root.style.setProperty('--color-bg-dark', Utils.adjustLightness(colors.c3, -3));
        root.style.setProperty('--color-white', colors.c4);
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

        // Auto-size the owner input using a hidden measurement span
        const measureSpan = document.createElement('span');
        measureSpan.style.cssText = 'position:absolute;visibility:hidden;white-space:pre;font-size:1.5rem;font-weight:600;';
        document.body.appendChild(measureSpan);

        const autoSizeOwnerInput = () => {
            const text = journalOwnerInput.value || journalOwnerInput.placeholder;
            measureSpan.textContent = text;
            journalOwnerInput.style.width = (measureSpan.offsetWidth + 8) + 'px';
        };
        journalOwnerInput.addEventListener('input', autoSizeOwnerInput);
        autoSizeOwnerInput();

        // ==================== CATEGORIES ====================

        // Add category button (from entry form)
        document.getElementById('addCategoryBtn').addEventListener('click', () => {
            this._categoryModalOrigin = 'entry';
            this.openCategoryModal();
        });

        // Add category button (from budget expense form)
        document.getElementById('addCategoryFromBudgetBtn').addEventListener('click', () => {
            this._categoryModalOrigin = 'budget';
            this.openCategoryModal();
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
            this.refreshTransactions();
        });

        ['filterType', 'filterStatus', 'filterMonth', 'filterCategory'].forEach(id => {
            document.getElementById(id).addEventListener('change', () => {
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
            popover.style.display = isOpen ? 'none' : 'block';
            if (!isOpen) CompanyManager.renderSwitcher();
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

        // ==================== MAIN TABS ====================

        document.querySelectorAll('.main-tab').forEach(btn => {
            btn.addEventListener('click', () => {
                this.switchMainTab(btn.dataset.tab);
            });
        });

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
            popover.style.display = popover.style.display === 'none' ? 'block' : 'none';
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
                ['themeC1', 'themeC2', 'themeC3', 'themeC4'].forEach((id, i) => {
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
        ['themeC1', 'themeC2', 'themeC3', 'themeC4'].forEach((id) => {
            document.getElementById(id).addEventListener('input', Utils.debounce(() => {
                const colors = {
                    c1: document.getElementById('themeC1').value,
                    c2: document.getElementById('themeC2').value,
                    c3: document.getElementById('themeC3').value,
                    c4: document.getElementById('themeC4').value,
                };
                Database.setThemeColors(colors);
                this.applyTheme('custom', colors);
            }, 100));
        });

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

        // ==================== BALANCE SHEET ====================

        // BS month/year change — persist selection
        document.getElementById('bsMonthMonth').addEventListener('change', () => {
            this._saveBsMonth();
            this.refreshBalanceSheet();
        });
        document.getElementById('bsMonthYear').addEventListener('change', () => {
            this._saveBsMonth();
            this.refreshBalanceSheet();
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
                if (skippedLoan) this._syncLoanJournalEntries(loanId, skippedLoan);
                this.refreshLoans();
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
                if (overrideLoan) this._syncLoanJournalEntries(loanId, overrideLoan);
                this.refreshLoans();
            };

            input.addEventListener('blur', save);
            input.addEventListener('keydown', (ke) => {
                if (ke.key === 'Enter') { ke.preventDefault(); input.blur(); }
                else if (ke.key === 'Escape') { ke.preventDefault(); this.refreshLoans(); }
            });
        });

        // ==================== BUDGET TAB ====================

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

        // Budget list panel click delegation
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
        document.getElementById('pvmSearchInput').addEventListener('input', (e) => this.handlePvmSearch(e.target.value));
        document.getElementById('pvmItemList').addEventListener('change', () => this._pvmUpdateSelectedCount());

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
            this.shareJournal();
        });
        document.getElementById('closeShareBtn').addEventListener('click', () => {
            UI.hideModal('shareModal');
        });
        document.getElementById('copyShareUrlBtn').addEventListener('click', () => {
            const url = document.getElementById('shareUrlDisplay').value;
            if (url) this.copyToClipboard(url);
        });
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
                Database.updateTransaction(parentId, data);
                UI.showNotification('Transaction updated successfully', 'success');
            } else {
                parentId = Database.addTransaction(data);
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
     * Create, update, or delete a linked sales tax entry based on the parent sale
     * @param {number} parentId - Parent transaction ID
     * @param {Object} data - Parent transaction form data
     */
    _manageSalesTaxEntry(parentId, data) {
        const category = Database.getCategoryById(data.category_id);
        const existingChildId = Database.getLinkedSalesTaxTransaction(parentId);

        if (category && category.is_sales) {
            const taxAmount = data.amount - (data.pretax_amount || data.amount);
            const dateLabel = Utils.formatSaleDateRange(data.sale_date_start, data.sale_date_end);
            const description = dateLabel ? `Sales Tax ${dateLabel}` : 'Sales Tax';

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
            const costAmount = data.inventory_cost || 0;
            const dateLabel = Utils.formatSaleDateRange(data.sale_date_start, data.sale_date_end);
            const description = dateLabel ? `Inventory Cost ${dateLabel}` : 'Inventory Cost';

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
                // Cascade-delete linked sales tax entry if present
                const childId = Database.getLinkedSalesTaxTransaction(this.deleteTargetId);
                if (childId) {
                    Database.deleteTransaction(childId);
                }
                // Cascade-delete linked inventory cost entry if present
                const invCostChildId = Database.getLinkedInventoryCostTransaction(this.deleteTargetId);
                if (invCostChildId) {
                    Database.deleteTransaction(invCostChildId);
                }
                Database.deleteTransaction(this.deleteTargetId);
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
        const apicEffective = (equityConfig.apic_received_date || equityConfig.apic_expected_date || '');
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
                    this._autoCreateAssetTransaction(assetId, name, cost, date);
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
    _autoCreateAssetTransaction(assetId, name, cost, date) {
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

        const month = date.substring(0, 7);
        Database.addTransaction({
            entry_date: date,
            category_id: catId,
            item_description: `Purchase: ${name}`,
            amount: cost,
            transaction_type: 'payable',
            status: 'paid',
            month_due: month,
            month_paid: month,
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

        // Show auto-create checkbox only for new loans
        const autoCreateGroup = document.getElementById('loanAutoCreate').closest('.form-group');
        if (editId) {
            autoCreateGroup.style.display = 'none';
        } else {
            autoCreateGroup.style.display = '';
            document.getElementById('loanAutoCreate').checked = true;
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
            this._syncLoanJournalEntries(parseInt(editingId), params);
            UI.showNotification('Loan updated', 'success');
        } else {
            const loanId = Database.addLoan(params);
            this.selectedLoanId = loanId;
            this._syncLoanJournalEntries(loanId, params);
            if (document.getElementById('loanAutoCreate').checked) {
                this._autoCreateLoanBudgetAndCategory(loanId, name, params);
            }
            UI.showNotification('Loan added', 'success');
        }

        UI.hideModal('loanConfigModal');
        this.refreshLoans();
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
            Database.deleteLoan(this.deleteLoanTargetId);
            if (this.selectedLoanId === this.deleteLoanTargetId) {
                this.selectedLoanId = null;
            }
            UI.showNotification('Loan deleted', 'success');
            this.refreshLoans();
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

        // Compute start/end months from amortization schedule
        const firstPayMonth = schedule[0].month;
        const lastPayMonth = schedule[schedule.length - 1].month;

        // Create budget expense for the loan payment
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
     * Sync journal entries for a single loan:
     * - Upserts the loan receivable (principal amount, status=received, dated start_date)
     * - Adds/updates pending payable entries for each payment month up to currentMonth
     * - Removes pending entries for months no longer in schedule (e.g. loan shortened)
     * - Paid entries are never touched
     */
    _syncLoanJournalEntries(loanId, params) {
        const { name, principal, annual_rate, term_months, payments_per_year, start_date, first_payment_date } = params;

        // --- Category: Loan Proceeds (receivable, hidden from P&L) ---
        let categories = Database.getCategories();
        let proceedsCat = categories.find(c => c.name === 'Loan Proceeds' && c.default_type === 'receivable');
        let proceedsCatId;
        if (!proceedsCat) {
            proceedsCatId = Database.addCategory('Loan Proceeds', false, null, 'receivable', null, true);
        } else {
            proceedsCatId = proceedsCat.id;
            // Ensure existing Loan Proceeds category is hidden from P&L
            if (!proceedsCat.show_on_pl) {
                Database.db.run('UPDATE categories SET show_on_pl = 1 WHERE id = ?', [proceedsCatId]);
            }
        }

        // --- Category: loan name (payable, hidden from P&L — only interest belongs on P&L) ---
        categories = Database.getCategories();
        let paymentCat = categories.find(c => c.name === name);
        let paymentCatId;
        if (!paymentCat) {
            paymentCatId = Database.addCategory(name, true, null, 'payable', null, true);
        } else {
            paymentCatId = paymentCat.id;
            // Ensure existing loan payment category is hidden from P&L
            if (!paymentCat.show_on_pl) {
                Database.db.run('UPDATE categories SET show_on_pl = 1 WHERE id = ?', [paymentCatId]);
            }
        }

        // --- Upsert receivable ---
        const existingReceivable = Database.db.exec(
            "SELECT id FROM transactions WHERE source_type = 'loan_receivable' AND source_id = ?",
            [loanId]
        );
        const startMonth = start_date.substring(0, 7);
        if (existingReceivable.length > 0 && existingReceivable[0].values.length > 0) {
            const rxId = existingReceivable[0].values[0][0];
            Database.db.run(
                'UPDATE transactions SET amount = ?, entry_date = ?, month_due = ?, month_paid = ?, category_id = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
                [principal, start_date, startMonth, startMonth, proceedsCatId, rxId]
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
                month_due: startMonth,
                month_paid: startMonth,
                source_type: 'loan_receivable',
                source_id: loanId
            });
        }

        // --- Sync payment entries ---
        const skipped = Database.getLoanSkippedPayments(loanId);
        const overrides = Database.getLoanPaymentOverrides(loanId);
        const schedule = Utils.computeAmortizationSchedule(
            { principal, annual_rate, payments_per_year, term_months, start_date, first_payment_date },
            skipped, overrides
        );
        const currentMonth = Utils.getCurrentMonth();

        // Build set of months that should have entries (up to currentMonth, non-skipped)
        const scheduledMonths = new Set();
        schedule.forEach(p => {
            if (!p.skipped && p.payment > 0 && p.month <= currentMonth) scheduledMonths.add(p.month);
        });

        // Get existing payment transactions for this loan
        const existingPayments = Database.db.exec(
            "SELECT id, payment_for_month, amount, status FROM transactions WHERE source_type = 'loan_payment' AND source_id = ?",
            [loanId]
        );
        const existingByMonth = {};
        if (existingPayments.length > 0) {
            existingPayments[0].values.forEach(([id, month, amount, status]) => {
                existingByMonth[month] = { id, amount, status };
            });
        }

        // Add or update payment entries
        schedule.forEach(p => {
            if (p.skipped || p.payment <= 0 || p.month > currentMonth) return;
            const existing = existingByMonth[p.month];
            if (!existing) {
                Database.addTransaction({
                    entry_date: p.month + '-01',
                    category_id: paymentCatId,
                    item_description: `${name} \u2013 Payment`,
                    amount: p.payment,
                    transaction_type: 'payable',
                    status: 'pending',
                    month_due: p.month,
                    payment_for_month: p.month,
                    source_type: 'loan_payment',
                    source_id: loanId
                });
            } else if (existing.status === 'pending' && existing.amount !== p.payment) {
                Database.db.run(
                    'UPDATE transactions SET amount = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
                    [p.payment, existing.id]
                );
            }
        });

        // Remove pending entries for months no longer in schedule (loan was shortened/edited)
        Object.entries(existingByMonth).forEach(([month, tx]) => {
            if (!scheduledMonths.has(month) && tx.status === 'pending') {
                Database.db.run('DELETE FROM transactions WHERE id = ?', [tx.id]);
            }
        });

        Database.autoSave();
    },

    /**
     * Sync journal entries for all loans. Called from refreshAll() so new months auto-appear.
     */
    syncAllLoanJournalEntries() {
        const loans = Database.getLoans();
        loans.forEach(loan => {
            this._syncLoanJournalEntries(loan.id, {
                name: loan.name,
                principal: loan.principal,
                annual_rate: loan.annual_rate,
                term_months: loan.term_months,
                payments_per_year: loan.payments_per_year,
                start_date: loan.start_date,
                first_payment_date: loan.first_payment_date
            });
        });
    },

    // ==================== BUDGET HANDLERS ====================

    refreshBudget() {
        const expenses = Database.getBudgetExpenses();
        UI.renderBudgetTab(expenses, this.selectedBudgetExpenseId);
    },

    openBudgetExpenseModal(editId) {
        document.getElementById('budgetExpenseForm').reset();
        document.getElementById('editingBudgetExpenseId').value = '';
        document.getElementById('budgetExpenseModalTitle').textContent = 'Add Budget Expense';
        document.getElementById('saveBudgetExpenseBtn').textContent = 'Add Expense';

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

        // Populate year dropdowns
        this._populateBudgetYearDropdowns();

        if (!editId) {
            const [year, month] = Utils.getCurrentMonth().split('-');
            document.getElementById('budgetStartMonth').value = month;
            document.getElementById('budgetStartYear').value = year;
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

        try {
            if (editingId) {
                Database.updateBudgetExpense(parseInt(editingId), name, amount, start, end, categoryId ? parseInt(categoryId) : null, notes);
                UI.showNotification('Expense updated', 'success');
            } else {
                const id = Database.addBudgetExpense(name, amount, start, end, categoryId ? parseInt(categoryId) : null, notes);
                this.selectedBudgetExpenseId = id;
                UI.showNotification('Expense added', 'success');
            }
            UI.hideModal('budgetExpenseModal');
            this.refreshBudget();
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
    },

    handleDeleteBudgetExpense(id) {
        this.deleteBudgetExpenseTargetId = id;
        UI.showModal('deleteBudgetExpenseModal');
    },

    confirmDeleteBudgetExpense() {
        if (this.deleteBudgetExpenseTargetId) {
            Database.deleteBudgetExpense(this.deleteBudgetExpenseTargetId);
            if (this.selectedBudgetExpenseId === this.deleteBudgetExpenseTargetId) {
                this.selectedBudgetExpenseId = null;
            }
            UI.showNotification('Expense deleted', 'success');
            this.refreshBudget();
        }
        UI.hideModal('deleteBudgetExpenseModal');
        this.deleteBudgetExpenseTargetId = null;
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

    // ==================== EXPORT ====================

    /**
     * Export all transactions as CSV
     */
    handleExportCsv() {
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

        this.downloadBlob(blob, filename);
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
        const blob = Database.exportToFile();

        if (window.showSaveFilePicker) {
            try {
                const handle = await window.showSaveFilePicker({
                    suggestedName: this.getSuggestedFilename(),
                    types: [{
                        description: 'Database Files',
                        accept: { 'application/x-sqlite3': ['.db'] }
                    }]
                });

                const writable = await handle.createWritable();
                await writable.write(blob);
                await writable.close();
                UI.showNotification('Journal saved', 'success');
            } catch (e) {
                if (e.name === 'AbortError') return;
                // Fall through to download
                this.downloadBlob(blob, this.getSuggestedFilename());
                UI.showNotification('Database saved successfully', 'success');
            }
        } else {
            // Fallback: show save as modal for naming
            const owner = document.getElementById('journalOwner').value.trim();
            document.getElementById('saveAsName').value = owner
                ? `${Utils.sanitizeFilename(owner)}_accounting_journal`
                : `accounting_journal_${new Date().toISOString().split('T')[0]}`;
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

            if (window.showSaveFilePicker) {
                try {
                    const handle = await window.showSaveFilePicker({
                        suggestedName,
                        types: [{
                            description: 'ZIP Archive',
                            accept: { 'application/zip': ['.zip'] }
                        }]
                    });
                    const writable = await handle.createWritable();
                    await writable.write(blob);
                    await writable.close();
                    UI.showNotification(`Saved ${companies.length} company database(s)`, 'success');
                    return;
                } catch (e) {
                    if (e.name === 'AbortError') return;
                }
            }

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
    confirmSaveAs() {
        const name = document.getElementById('saveAsName').value.trim();
        if (!name) {
            UI.showNotification('Please enter a file name', 'error');
            return;
        }

        const blob = Database.exportToFile();
        this.downloadBlob(blob, `${Utils.sanitizeFilename(name)}.db`);
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
        const analytics = Database.getLinkedProductAnalytics(dateFrom, dateTo);
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
        const taxRate = parseFloat(document.getElementById('productTaxRate').value) || 0;
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
        document.getElementById('productTaxRate').value = product.tax_rate || '';
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
        document.getElementById('pvmSearchInput').value = '';

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

        // Timeline banner
        const banner = document.getElementById('beTimelineBanner');
        if (timeline.start || timeline.end) {
            const startLabel = timeline.start ? Utils.formatMonthShort(timeline.start) : 'Start';
            const endLabel = timeline.end ? Utils.formatMonthShort(timeline.end) : 'Present';
            const isLocal = (cfg.timeline && (cfg.timeline.start || cfg.timeline.end));
            const asOfLabel = (useProjected && cfg.asOfMonth) ? ` as of ${Utils.formatMonthShort(cfg.asOfMonth)}` : '';
            banner.textContent = `Timeline: ${startLabel} \u2013 ${endLabel}${isLocal ? ' (local override)' : ''} \u2022 ${useProjected ? 'Projected' : 'Actual'}${asOfLabel}`;
            banner.style.display = 'block';
        } else {
            banner.style.display = 'none';
        }

        // Compute monthly fixed costs
        let months = [];
        if (timeline.start && timeline.end) {
            months = Utils.generateMonthRange(timeline.start, timeline.end);
        } else {
            months = [currentMonth];
        }

        // Ensure P&L has been rendered so UI._pnlMonthOpex is available
        if (!UI._pnlMonthOpex) {
            this.refreshPnL();
        }
        const totalOpexByMonth = UI._pnlMonthOpex || {};

        // Determine monthly fixed costs
        let avgMonthlyFixed;
        if (cfg.fixedCostOverride != null && cfg.fixedCostOverride > 0) {
            // Manual override takes priority
            avgMonthlyFixed = cfg.fixedCostOverride;
        } else {
            const actualMonths = months.filter(m => m <= currentMonth);
            const na = actualMonths.length || 1;

            if (useProjected) {
                const pastValues = actualMonths.map(m => totalOpexByMonth[m] || 0).filter(v => v > 0);
                avgMonthlyFixed = pastValues.length > 0 ? pastValues.reduce((a, b) => a + b, 0) / pastValues.length : 0;
            } else {
                const sum = actualMonths.reduce((acc, m) => acc + (totalOpexByMonth[m] || 0), 0);
                avgMonthlyFixed = sum / na;
            }
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
            const fixedTotal = Math.round(avgMonthlyFixed * monthCount * 100) / 100;
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
                const projAvgFixed = avgMonthlyFixed;
                timelinePoints = Utils.computeBreakevenTimeline(
                    cfg, months, {}, {},
                    hasOverride
                        ? () => projAvgFixed
                        : (useProjected
                            ? () => projAvgFixed
                            : (m) => totalOpexByMonth[m] || 0)
                );
                const actualRevenueByMonth = Database.getActualRevenueByMonth();
                this._renderBeTimelineChart(timelinePoints, actualRevenueByMonth);
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
     * Render the monthly timeline chart comparing actual revenue to break-even revenue needed
     * @param {Array} timelinePoints - From Utils.computeBreakevenTimeline()
     * @param {Object} actualRevenueByMonth - { 'YYYY-MM': totalRevenue } from P&L
     */
    _renderBeTimelineChart(timelinePoints, actualRevenueByMonth) {
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
                        label: 'Actual Revenue',
                        data: timelinePoints.map(p => actualRevenueByMonth[p.month] || 0),
                        backgroundColor: successColor + '60',
                        borderColor: successColor,
                        borderWidth: 1
                    },
                    {
                        label: 'Revenue Needed (excl. B2B)',
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
        // Pull from P&L (respects overrides, works with all old entries)
        const plRevenue = Database.getPLRevenueByMonth();
        const totalMonths = months.length;
        const timelineStart = months[0];
        const timelineEnd = months[months.length - 1];

        // Elapsed months (from start through asOfMonth)
        const elapsedMonths = months.filter(m => m <= asOfMonth);
        const elapsedCount = elapsedMonths.length;
        const remainingCount = totalMonths - elapsedCount;

        // Actual cumulative revenue through asOfMonth (from P&L data)
        let actualTotal = 0, actualB2b = 0, actualConsumer = 0;
        elapsedMonths.forEach(m => {
            actualTotal += (plRevenue.total[m] || 0);
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
            const monthActual = plRevenue.consumer[m] || 0;
            chartActual.push(Math.round(monthActual * 100) / 100);

            // Consumer BE revenue target for this month (excludes B2B)
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
                label: 'Revenue',
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
                label: 'Projected Revenue',
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

        const actualMonths = months.filter(m => m <= currentMonth);
        let totalFixed = 0;

        if (useProjected) {
            const values = actualMonths.map(m => totalOpexByMonth[m] || 0).filter(v => v > 0);
            totalFixed = values.length > 0 ? values.reduce((a, b) => a + b, 0) / values.length : 0;
        } else {
            const sum = actualMonths.reduce((acc, m) => acc + (totalOpexByMonth[m] || 0), 0);
            totalFixed = sum / (actualMonths.length || 1);
        }

        if (hint) hint.textContent = totalFixed > 0
            ? `Avg. monthly fixed costs: ${Utils.formatCurrency(totalFixed)}/mo (${useProjected ? 'projected' : 'actual'}, from P&L)`
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
    downloadBlob(blob, filename) {
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
        this.restoreTabOrder();
        this.setupTabDragDrop();
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

    async shareJournal() {
        const supaConfig = Database.getSupabaseConfig();
        if (!supaConfig || !supaConfig.url || !supaConfig.anonKey) {
            UI.showNotification('Sharing requires Supabase. Set up Group Sync first.', 'error');
            return;
        }

        if (!SupabaseAdapter.isInitialized()) {
            SupabaseAdapter.init(supaConfig.url, supaConfig.anonKey);
        }

        // Show modal with loading state
        document.getElementById('shareStatus').textContent = 'Uploading snapshot...';
        document.getElementById('shareUrlSection').style.display = 'none';
        UI.showModal('shareModal');

        try {
            const blob = new Uint8Array(Database.db.export());
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

            // Display
            document.getElementById('shareUrlDisplay').value = shareUrl;
            document.getElementById('shareStatus').textContent = '';
            document.getElementById('shareUrlSection').style.display = '';
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
            this.restoreTabOrder();
            this.setupTabDragDrop();
            this.setupTabScrollFade();
            this.loadAndApplyTimeline();
            this.refreshAll();
            this.setupEventListeners();
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
            'saveDbBtn', 'saveAllDbBtn', 'loadDbBtn', 'shareBtn'
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
    _veProductSortCol: 'qty',
    _veProductSortDir: 'desc',

    refreshVESales() {
        this._veSales = Database.getVESales();
        this._veItemsCache = Database.getAllVESaleItems();
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

        if (hasData) {
            this.veApplyFilters();
        }
    },

    veApplyFilters() {
        const source = document.getElementById('ve-filterSource').value;
        const from = document.getElementById('ve-filterFrom').value;
        const to = document.getElementById('ve-filterTo').value;
        const sortBy = document.getElementById('ve-sortBy').value;

        this._veFiltered = this._veSales.filter(s => {
            if (source !== 'both' && s.source !== source) return false;
            if (from && s.date < from) return false;
            if (to && s.date > to) return false;
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
        const d = new Date(isoStr);
        return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
    },

    veRenderSummary() {
        let subtotal = 0, tax = 0, total = 0, shipping = 0, discount = 0;
        for (const s of this._veFiltered) {
            subtotal += s.subtotal || 0;
            tax += s.tax || 0;
            total += s.total || 0;
            shipping += s.shipping || 0;
            discount += s.discount || 0;
        }
        document.getElementById('ve-cardPretax').textContent = this.veFmt(subtotal);
        document.getElementById('ve-cardTax').textContent = this.veFmt(tax);
        document.getElementById('ve-cardTotal').textContent = this.veFmt(total);
        document.getElementById('ve-cardCount').textContent = `${this._veFiltered.length} transaction${this._veFiltered.length !== 1 ? 's' : ''}`;

        const discountCard = document.getElementById('ve-discountCard');
        if (discount > 0) {
            discountCard.style.display = '';
            document.getElementById('ve-cardDiscount').textContent = '-' + this.veFmt(discount);
            const dc = this._veFiltered.filter(s => s.discount > 0).length;
            document.getElementById('ve-cardDiscountCount').textContent = `${dc} order${dc !== 1 ? 's' : ''} with discounts`;
        } else {
            discountCard.style.display = 'none';
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
                    entry.totalRevenue += item.amount || 0;
                    if ((s.subtotal || 0) > 0) {
                        entry.totalTax += (s.tax || 0) * ((item.amount || 0) / s.subtotal);
                    }
                    if (item.inferred) entry.hasInferred = true;
                    if (item.productNumber && !entry.productNumber) entry.productNumber = item.productNumber;
                }
            } else {
                unscrapedSubtotal += s.subtotal || 0;
                unscrapedTax += s.tax || 0;
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
            tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;color:#6c757d;padding:20px;">No transactions match filters</td></tr>';
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

            html += `<tr>
                <td>${Utils.escapeHtml(s.transaction_no)}</td>
                <td>${this.veFmtDate(s.date)}</td>
                <td><span class="ve-source-badge ${s.source}">${s.source === 'online' ? 'Online' : 'Trade Show'}</span></td>
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
        if (typeof val === 'number') return val;
        const cleaned = String(val).replace(/[$,\s]/g, '');
        const num = parseFloat(cleaned);
        return isNaN(num) ? 0 : num;
    },

    veParseDate(val) {
        if (!val) return null;
        if (val instanceof Date) return isNaN(val.getTime()) ? null : val.toISOString().split('T')[0];
        const str = String(val).trim();
        const d = new Date(str);
        return isNaN(d.getTime()) ? null : d.toISOString().split('T')[0];
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

        Database.clearVESales();

        // Normalize dates to YYYY-MM-DD
        const sales = data.sales.map(s => ({
            ...s,
            date: s.date ? s.date.split('T')[0] : s.date,
        }));
        Database.upsertVESales(sales);

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

        // Only clear the specific source being imported
        Database.clearVESales(effectiveSource);
        Database.upsertVESales(sales);
        let totalItems = 0;
        for (const [txNo, items] of lineItems) {
            Database.upsertVESaleItems(txNo, items);
            totalItems += items.length;
        }
        return { salesCount: sales.length, itemsCount: totalItems, source: effectiveSource };
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

    veApplyPreset() {
        const preset = document.getElementById('ve-filterPreset').value;
        const fromEl = document.getElementById('ve-filterFrom');
        const toEl = document.getElementById('ve-filterTo');
        const now = new Date();

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
                const lastDay = new Date(y, last.getMonth() + 1, 0).getDate();
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
        this.veApplyFilters();
    },

    setupVESalesListeners() {
        // Filter controls
        document.getElementById('ve-filterSource').addEventListener('change', () => this.veApplyFilters());
        document.getElementById('ve-filterFrom').addEventListener('change', () => this.veApplyFilters());
        document.getElementById('ve-filterTo').addEventListener('change', () => this.veApplyFilters());
        document.getElementById('ve-sortBy').addEventListener('change', () => this.veApplyFilters());
        document.getElementById('ve-filterPreset').addEventListener('change', () => this.veApplyPreset());
        document.getElementById('ve-highlightDupes').addEventListener('change', () => this.veRenderProducts());

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

        // Clear data
        document.getElementById('veClearBtn').addEventListener('click', () => {
            if (!confirm('Clear all VE Sales data? This cannot be undone.')) return;
            Database.clearVESales();
            Database.setVEImportMeta(null);
            this.refreshVESales();
            UI.showNotification('VE Sales data cleared', 'success');
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
    }
};

// Initialize application when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    App.init();
});

// Export for use in other modules
window.App = App;
