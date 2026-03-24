# Accounting Journal Calculator - Code Map

Vanilla HTML/CSS/JS accounting app with SQLite (sql.js) in-browser database, multi-company support, and optional Supabase sync.

## File Overview

| File | Lines | Role |
|------|-------|------|
| `index.html` | ~2736 | All UI structure: tabs, modals, forms |
| `js/app.js` | ~13120 | Core logic, event handlers, all feature controllers |
| `js/database.js` | ~4023 | SQLite schema, CRUD, financial calculations |
| `js/ui.js` | ~3420 | DOM rendering, spreadsheets, forms, notifications |
| `js/utils.js` | ~850 | Formatting, date helpers, financial math (amortization, depreciation, break-even) |
| `js/companies.js` | ~450 | Multi-company manager (IndexedDB, up to 5 companies) |
| `js/sync.js` | ~318 | Group sync with version history, polling |
| `js/supabase-adapter.js` | ~306 | Supabase backend for sync, shares, members |
| `js/tutorial.js` | ~1924 | Guided tours, interactive lessons, spotlight overlay |
| `css/styles.css` | ~3994 | All styling, themes, dark mode, responsive breakpoints |

## Feature → Code Location Map

### Journal Entries / Transactions
- **HTML**: `index.html:341-589` (journalTab, entryModal, quickEntryModal, filters, bulk bar)
- **Logic**: `app.js:3754-3810` handleFormSubmit, `app.js:4587-4722` edit/duplicate/delete
- **DB**: `database.js:1041-1278` getTransactions, addTransaction, updateTransaction, delete, bulk ops
- **UI render**: `ui.js:317-505` renderTransactions, renderTransactionRow
- **Quick entry**: `app.js:1783-2052` openQuickEntry, addQuickEntryRow, saveQuickEntries
- **Inline status**: `app.js:4244-4369` handleInlineStatusChange, confirmMonthPaidPrompt

### Bulk Actions (mark paid, reset pending)
- **HTML**: `index.html:567-576` bulkActionBar
- **Logic**: `app.js:4370-4588` toggleBulkSelectMode, handleBulkMarkPaid, handleBulkResetPending
- **DB**: `database.js:1238-1268` bulkSetDatePaid, bulkResetToPending

### Categories
- **HTML**: `index.html:2101-2197` categoryModal, manageCategoriesModal
- **Logic**: `app.js:3896-4114` openCategoryModal, handleSaveCategory, confirmDeleteCategory
- **DB**: `database.js:832-1040` getCategories, addCategory, updateCategory, deleteCategory
- **UI render**: `ui.js:507-673` renderCategoryItem, renderManageCategoriesList

### Folders (category grouping)
- **HTML**: `index.html:2200-2256` folderModal, deleteFolderModal
- **Logic**: `app.js:4115-4243` openFolderModal, handleSaveFolder, confirmDeleteFolder
- **DB**: `database.js:775-830` getFolders, addFolder, updateFolder, deleteFolder
- **Batch add**: `app.js:4723-4947` openAddFolderEntriesModal, confirmAddFolderEntries

### Cash Flow
- **HTML**: `index.html:592-621` cashflowTab, cashflowSpreadsheet
- **Logic**: `app.js:266-393` refreshCashFlow, setupCashFlowDragDrop, setupCashFlowCellEditing (`app.js:680-724`)
- **DB**: `database.js:1447-1493` getCashFlowSummary, getCashFlowSpreadsheet
- **DB overrides**: `database.js:2946-2972` getAllCashFlowOverrides, setCashFlowOverride
- **UI render**: `ui.js:674-1053` renderCashFlowSpreadsheet

### Profit & Loss (P&L)
- **HTML**: `index.html:624-659` pnlTab, pnlSpreadsheet
- **Logic**: `app.js:861-984` refreshPnL, setupPnLCellEditing
- **DB**: `database.js:1522-1795` getPLSpreadsheet, getPLRevenueByMonth, getMonthlyTotalOpex, getMonthlyTotalRevenue
- **DB overrides**: `database.js:3409-3445` getAllPLOverrides, setPLOverride
- **DB tax mode**: `database.js:1798-1816` getPLTaxMode, setPLTaxMode
- **UI render**: `ui.js:1055-1441` renderProfitLossSpreadsheet

### Balance Sheet
- **HTML**: `index.html:662-705` balancesheetTab, balanceSheetContent, bsRatiosSection
- **Logic**: `app.js:4954-5200` refreshBalanceSheet
- **DB**: `database.js:2974-3403` getCashAsOf, getARByCategory, getAPByCategory, getRetainedEarningsAsOf, getPLTotalsThrough
- **UI render**: `ui.js:1443-1694` renderBalanceSheet, renderFinancialRatios (includes Cash Ratio)

### Fixed Assets & Depreciation
- **HTML**: `index.html:708-736` assetsTab; `index.html:1805-1876` fixedAssetModal, deleteAssetModal
- **Logic**: `app.js:~5200+` refreshFixedAssets
- **DB**: `database.js:1898-2016` getFixedAssets, addFixedAsset, updateFixedAsset, deleteFixedAsset, getAssetDepreciationByMonth
- **UI render**: `ui.js:1686-1784` renderFixedAssetsTab
- **Depreciation math**: `utils.js:~600+` computeDepreciationSchedule

### Equity (Stockholders' Equity)
- **HTML**: `index.html:727-735` equityDisplayPanel; `index.html:1879-1935` equityModal
- **DB**: `database.js:2756-2808` getEquityConfig, setEquityConfig
- **UI render**: `ui.js:1786-1853` renderEquitySection

### Loans & Amortization
- **HTML**: `index.html:739-757` loanTab; `index.html:1938-2001` loanConfigModal, deleteLoanModal
- **Logic**: `app.js:~5600+` refreshLoans
- **DB**: `database.js:2018-2191` getLoans, addLoan, updateLoan, deleteLoan, getLoanInterestByMonth, skip/override payments
- **UI render**: `ui.js:1855-1957` renderLoansTab
- **Amortization math**: `utils.js:515-580` computeAmortizationSchedule

### Budget
- **HTML**: `index.html:760-788` budgetTab; `index.html:1645-1802` budgetExpenseModal, recordBudgetModal
- **Logic**: `app.js:~5900+` refreshBudget; `app.js:1874-2052` quick budget entry
- **DB**: `database.js:2279-2416` getBudgetExpenses, addBudgetExpense, budget groups, moveBudgetExpenseToGroup
- **UI render**: `ui.js:1959-2249` renderBudgetTab, _renderBudgetPieChart

### Break-Even Analysis
- **HTML**: `index.html:791-845` breakevenTab; `index.html:1508-1642` beConfigModal
- **Logic**: `app.js:6740-7496` refreshBreakeven, openBeConfigModal, handleSaveBeConfig, chart rendering
- **DB**: `database.js:2610-2668` getBreakevenConfig, setBreakevenConfig, getBudgetFixedCostsForMonth
- **UI render**: `ui.js:2626-2884` renderBreakevenSummaryCards, renderBreakevenChannelBreakdown, renderBreakevenDataTable
- **Math**: `utils.js:~700+` computeBreakEven, computeBreakEvenChartPoints, computeBreakevenTimeline

### Projected Sales
- **HTML**: `index.html:848-950` projectedsalesTab
- **Logic**: `app.js:7840-8034` refreshProjectedSales, handleSaveProjectedSales, setupPsUnitEditing
- **DB**: `database.js:2670-2754` getProjectedSalesConfig, setProjectedSalesConfig, getProjectedSalesSpreadsheet
- **UI render**: `ui.js:2886-2999` renderProjectedSalesSummaryCards, renderProjectedSalesGrid

### Products & Inventory
- **HTML**: `index.html:953-1079` productsTab; `index.html:1378-1450` productModal, pvmModal
- **Logic**: `app.js:~6200+` refreshProducts; `app.js:6400-6739` PVM search/suggest, product charts/analytics
- **DB**: `database.js:2418-2607` addProduct, getProducts, getLinkedProductAnalytics, product-VE mappings

### VE Sales (Vendor/Event Sales)
- **HTML**: `index.html:1082-1316` vesalesTab (import panel, controls, subtabs, dashboard)
- **Logic**: `app.js:8685-10362` refreshVESales, Excel/JSON import, event management, journal creation
- **DB**: `database.js:3544-3716` VE sales CRUD, VE events CRUD, event assignments

### B2B Contracts
- **HTML**: `index.html:1327-1352` b2bcontractTab; `index.html:2004-2084` b2bContractModal
- **Logic**: `app.js:10363-11099` B2B contract CRUD, sync journal entries, finalize/unfinalize
- **DB**: `database.js:2193-2277` getB2BContracts, addB2BContract, setB2BContractFinalized

### Dashboard
- **HTML**: `index.html:1319-1324` dashboardTab; `index.html:2330-2364` kpiDetailModal, analyzeKpiModal
- **Logic**: `app.js:1088-2224` refreshDashboard, _computeKpiData, _renderDashboardSections
- **KPI system**: `app.js:1161-1208` _computeKpiData (cached metrics: cash, burn, EBITDA, CMGR, Rule of 40, DSCR, working capital)
- **KPI modals**: `app.js:1208-1280` _openSectionAnalysis (section→KPI grid modal), _openKpiDetail (drill-down detail tables)
- **KPI detail renderers**: `app.js:~1300-2100` _kpiDetail_cashposition, _kpiDetail_grossburn, _kpiDetail_netburn, _kpiDetail_revtrend, _kpiDetail_ebitda, _kpiDetail_cmgr, _kpiDetail_cmgrnonb2b, _kpiDetail_overdue, _kpiDetail_workingcapital, _kpiDetail_rule40, _kpiDetail_rule40nb, _kpiDetail_dscr
- **Dashboard charts**: `app.js:~2100+` Revenue Concentration pie chart, AR Aging breakdown
- **CSS**: `styles.css:2629-2690` KPI cards, detail modal, analyze modal, section clickable states

### Analyze Mode Charts
- **Logic**: `app.js:~2500+` _updateAnalyzeCharts, _renderAnalyzeCFChart, _renderAnalyzePnLChart, _renderAnalyzeBSChart

### Work/Analyze Mode Toggle
- **HTML**: `index.html:46-52` mode-toggle buttons
- **Logic**: `app.js:1061-1084` switchMode
- **CSS**: `styles.css:2579-2585`

### Themes & Dark Mode
- **HTML**: `index.html:166-216` themePreset select, darkModeToggle, custom color pickers
- **Logic**: `app.js:2065-2187` loadAndApplyTheme, applyTheme, loadDarkMode, setDarkMode
- **DB**: `database.js:1835-1897` getThemePreset, setThemePreset, getThemeColors, getThemeDark
- **CSS tokens**: `styles.css:6-121` design tokens; `styles.css:122-182` dark mode overrides
- **CSS theme variants**: `styles.css:2934-2963` modern/futuristic/vintage/accounting
- **Presets defined**: `app.js:109-121` themePresets object (8 themes)

### Timeline (fiscal date range)
- **HTML**: `index.html:217-265` timeline start/end month/year selects
- **Logic**: `app.js:730-856` getTimeline, loadAndApplyTimeline, applyTimelineConstraints, handleTimelineChange
- **DB**: `database.js:2899-2938` getTimeline, setTimelineStart, setTimelineEnd

### Tab Management (reorder, hide, context menu)
- **HTML**: `index.html:54-126` mainTabs; `index.html:323-338` tabContextMenu, manageTabsModal
- **Logic**: `app.js:398-675` setupTabDragDrop, restoreTabOrder, applyHiddenTabs, setupTabContextMenu, openManageTabsModal
- **DB**: `database.js:2810-2897` getTabOrder, setTabOrder, getHiddenTabs, setHiddenTabs, resetTabData

### Company Management (multi-company)
- **HTML**: `index.html:17-38` companyBtn/companyPopover; `index.html:2304-2374` manageCompaniesModal, saveAsModal
- **Logic**: `app.js:7627-7900+` promptInitialCompanyName, openManageCompanies, switchToCompany
- **Module**: `companies.js` CompanyManager class (init, switchTo, createNew, rename, delete, copySection, renderSwitcher)

### Save / Load / Export
- **HTML**: `index.html:294-297` saveDbBtn, saveAllDbBtn, loadDbBtn
- **Logic**: `app.js:6045-6115` handleExportCsv, handleSaveDatabase; `app.js:7497-7626` downloadBlob, _finalizeLoad, _reloadUI
- **DB**: `database.js:3447-3540` autoSave, saveToIndexedDB, loadFromIndexedDB, exportToFile, importFromFile
- **CSV**: `ui.js:2256-2300` generateCsv

### Sync & Collaboration
- **HTML**: `index.html:131-139` syncBtn; `index.html:2441-2601` syncMenuModal, groupCreatedModal, versionHistoryModal, membersModal
- **Logic**: `app.js:8035-8474` openSyncMenu, initSync, encodeInviteCode, handleSyncStatusChange, generateShareQR
- **Module**: `sync.js` (createGroup, joinGroup, push, pull, getHistory, polling)
- **Backend**: `supabase-adapter.js` (Supabase RPC, storage, members, shares)

### Share (view-only)
- **HTML**: `index.html:141-149` shareBtn; `index.html:2604-2623` shareModal
- **Logic**: `app.js:8475-8524` initViewOnlyMode, applyViewOnlyRestrictions, showViewOnlyBanner

### Undo/Redo
- **Logic**: `app.js:6-71` pushUndo, undo, redo, _showUndoToast, clearUndoHistory

### Tutorial / Quick Guide
- **HTML**: `index.html:1355-1364` quickguideTab; `index.html:150-156` helpBtn
- **Logic**: `app.js:11099-11168` renderQuickGuide, renderChangelog
- **Module**: `js/tutorial.js` full tour system with spotlight overlay and interactive lessons
- **CSS**: `styles.css:3721-3798`

### Notifications / Toasts
- **UI render**: `ui.js:2576-2614` showNotification

### Modals (generic)
- **UI**: `ui.js:2334-2392` showModal, hideModal, resetForm
- **CSS**: `styles.css:1642-1718`

### Forms / Entry Form
- **UI**: `ui.js:2398-2574` populateFormForEdit, getFormData, validateFormData
- **CSS**: `styles.css:1719-1849`

### Sales Tax Auto-Entries
- **Logic**: `app.js:3811-3895` _manageSalesTaxEntry, _manageInventoryCostEntry
- **DB**: `database.js:893-969` getOrCreateSalesTaxCategory, getLinkedSalesTaxTransaction, getOrCreateInventoryCostCategory

### Shipping Fee Auto-Entries
- **HTML**: `index.html:266-276` shippingFeeRate input in gear popover
- **Logic**: `app.js` _manageShippingEntry, loadShippingFeeRate
- **DB**: `database.js:1055-1126` getOrCreateShippingCategory, getLinkedShippingTransaction, updateShippingTransaction, getShippingFeeConfig, setShippingFeeConfig

## CSS Responsive Breakpoints
- `1024px`: Sidebar auto-collapse (`styles.css:3000-3005`)
- `768px`: Stack layout, smaller fonts (`styles.css:3008-3052`)
- `480px`: Mobile-optimized, touch targets (`styles.css:3054-3077`)

## Database Schema (18 tables)
Defined in `database.js:70-337`. Key tables: `transactions`, `categories`, `category_folders`, `balance_sheet_assets`, `loans`, `budget_expenses`, `budget_groups`, `products`, `b2b_contracts`, `ve_sales`, `ve_sale_items`, `ve_events`, `product_ve_mappings`, `pl_overrides`, `cashflow_overrides`, `app_meta`, `loan_skipped_payments`, `loan_payment_overrides`. Migrations: `database.js:342-773`.

## Key Utilities (utils.js)
- Formatting: `formatCurrency`, `formatDate`, `formatMonthDisplay` (lines 11-119)
- Date helpers: `getTodayDate`, `getCurrentMonth`, `isPaidLate`, `isOverdue` (lines 68-213)
- Financial math: `computeAmortizationSchedule` (515+), `computeDepreciationSchedule` (600+), `computeBreakEven` (700+)
- Color helpers: `hexToHSL`, `adjustLightness` (446-503)

## Event Listeners
All wired in `app.js:2206-3753` setupEventListeners — keyboard shortcuts, form submissions, filter changes, theme controls, sidebar toggle, all button clicks.
