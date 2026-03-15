/**
 * Interactive Tutorial / Guided Walkthrough System
 * Spotlight + tooltip walkthrough for all app features
 * Includes interactive lessons with guided input and impact visualization
 */

const Tutorial = {
    // ─── State ───────────────────────────────────────────────
    state: {
        active: false,
        currentTourId: null,
        currentStepIndex: 0,
        steps: [],
        isLesson: false,
    },

    // DOM element references (created in init)
    overlayEl: null,
    spotlightEl: null,
    tooltipEl: null,
    pickerEl: null,

    // Sample data tracking for lessons
    _sampleData: {
        transactionIds: [],
        categoryIds: [],
        budgetExpenseIds: [],
        budgetGroupIds: [],
        folderIds: [],
    },

    // Active input prompt listener (cleaned up between steps)
    _inputCleanup: null,

    // ─── Tour Definitions ────────────────────────────────────
    tours: {
        quickStart: {
            label: 'Quick Start',
            description: 'Learn the basics in 2 minutes',
            steps: [
                {
                    target: '.app-sidebar',
                    title: 'Sidebar Navigation',
                    text: 'The sidebar is your main navigation. Tabs are organized into sections: Overview, Reports, Assets, Planning, and Sales.',
                    position: 'right',
                },
                {
                    target: '#companyBtn',
                    title: 'Company Switcher',
                    text: 'Switch between multiple companies. Each company has its own journal, categories, and reports.',
                    position: 'right',
                },
                {
                    target: '#newEntryBtn',
                    title: 'Create Entries',
                    text: 'Click here to add a new journal entry. Entries can be receivables (money owed to you) or payables (money you owe).',
                    position: 'bottom',
                    tab: 'journal',
                },
                {
                    target: '.summary-section',
                    title: 'Summary Cards',
                    text: 'These cards show your Cash Balance, Accounts Receivable, and Accounts Payable in real time.',
                    position: 'bottom',
                    tab: 'journal',
                },
                {
                    target: '.main-tab[data-tab="cashflow"]',
                    title: 'Financial Reports',
                    text: 'The Reports section includes Cash Flow, Profit & Loss, and Balance Sheet tabs for a full financial picture.',
                    position: 'right',
                },
                {
                    target: '.main-tab[data-tab="budget"]',
                    title: 'Planning Tools',
                    text: 'Use Budget, Break-Even, and Projected Sales tabs to plan and forecast your business finances.',
                    position: 'right',
                },
                {
                    target: '.main-tab[data-tab="products"]',
                    title: 'Sales Management',
                    text: 'Track your product catalog and import sales data from VE (Volusion) in the Sales section.',
                    position: 'right',
                },
                {
                    target: '#gearBtn',
                    title: 'Settings',
                    text: 'Customize themes, dark mode, and your financial timeline here.',
                    position: 'top',
                },
                {
                    target: '#saveDbBtn',
                    title: 'Save & Load',
                    text: 'Save your data to a file and load it later. Use "Save All" to export all companies at once.',
                    position: 'bottom',
                    tab: 'journal',
                },
                {
                    target: '#helpBtn',
                    title: 'Help & Tutorials',
                    text: 'You can always come back here to take a tour of any specific feature area. Happy accounting!',
                    position: 'top',
                },
            ],
        },

        journal: {
            label: 'Journal Entries',
            description: 'Creating and managing transactions',
            steps: [
                {
                    target: '.main-tab[data-tab="journal"]',
                    title: 'Journal Tab',
                    text: 'This is your main workspace. All receivables and payables are recorded here as journal entries.',
                    position: 'right',
                    tab: 'journal',
                },
                {
                    target: '#newEntryBtn',
                    title: 'New Entry Button',
                    text: 'Click here to open the entry form and create a new receivable or payable transaction.',
                    position: 'bottom',
                    tab: 'journal',
                },
                {
                    target: '#manageCategoriesBtn',
                    title: 'Categories',
                    text: 'Manage your categories and folders. Categories help organize entries by type (e.g., "Rent", "Sales Revenue").',
                    position: 'bottom',
                    tab: 'journal',
                },
                {
                    target: '#addFolderEntriesBtn',
                    title: 'Add Folder Entries',
                    text: 'Quickly create entries for all categories in a folder at once \u2014 great for recurring monthly expenses.',
                    position: 'bottom',
                    tab: 'journal',
                },
                {
                    target: '.summary-section',
                    title: 'Financial Summary',
                    text: 'Cash Balance = total received minus total paid. A/R shows money owed to you. A/P shows what you owe others.',
                    position: 'bottom',
                    tab: 'journal',
                },
                {
                    target: '#exportCsvBtn',
                    title: 'Export to CSV',
                    text: 'Export your journal entries to a CSV file for use in Excel or other tools.',
                    position: 'bottom',
                    tab: 'journal',
                },
                {
                    target: '#saveDbBtn',
                    title: 'Save Database',
                    text: 'Save your current company\'s data to a .db file. Use "Save All" to export every company in one file.',
                    position: 'bottom',
                    tab: 'journal',
                },
                {
                    target: '#loadDbBtn',
                    title: 'Load Database',
                    text: 'Load a previously saved .db file. You can replace the current company or import as a new one.',
                    position: 'bottom',
                    tab: 'journal',
                },
            ],
        },

        reports: {
            label: 'Reports',
            description: 'Cash Flow, P&L, and Balance Sheet',
            steps: [
                {
                    target: '.main-tab[data-tab="cashflow"]',
                    title: 'Cash Flow Tab',
                    text: 'View your cash flow summary \u2014 money coming in and going out over time.',
                    position: 'right',
                    tab: 'cashflow',
                },
                {
                    target: '#cashflowTab',
                    title: 'Cash Flow Statement',
                    text: 'Shows monthly cash inflows and outflows. Toggle between Projected and Actual data views.',
                    position: 'left',
                    tab: 'cashflow',
                },
                {
                    target: '.main-tab[data-tab="pnl"]',
                    title: 'Profit & Loss Tab',
                    text: 'Your income statement \u2014 shows revenue, expenses, and net income.',
                    position: 'right',
                    tab: 'pnl',
                },
                {
                    target: '#pnlTab',
                    title: 'P&L Statement',
                    text: 'View income vs expenses. Choose Corporate (21%) or Pass-through tax mode. Toggle Projected vs Actual.',
                    position: 'left',
                    tab: 'pnl',
                },
                {
                    target: '.main-tab[data-tab="balancesheet"]',
                    title: 'Balance Sheet Tab',
                    text: 'View assets, liabilities, and equity at a point in time.',
                    position: 'right',
                    tab: 'balancesheet',
                },
                {
                    target: '#balancesheetTab',
                    title: 'Balance Sheet',
                    text: 'Shows your financial position: Assets = Liabilities + Equity. Includes financial ratios at the bottom.',
                    position: 'left',
                    tab: 'balancesheet',
                },
            ],
        },

        assets: {
            label: 'Assets & Loans',
            description: 'Fixed assets, depreciation, and loans',
            steps: [
                {
                    target: '.main-tab[data-tab="assets"]',
                    title: 'Assets & Equity Tab',
                    text: 'Manage fixed assets with depreciation tracking and stockholders\' equity.',
                    position: 'right',
                    tab: 'assets',
                },
                {
                    target: '#assetsTab',
                    title: 'Fixed Assets',
                    text: 'Add assets like equipment or vehicles. Set depreciation method (straight-line or double-declining), useful life, and salvage value.',
                    position: 'left',
                    tab: 'assets',
                },
                {
                    target: '.main-tab[data-tab="loan"]',
                    title: 'Loans Tab',
                    text: 'Track loans with automatic amortization schedules and monthly payment calculations.',
                    position: 'right',
                    tab: 'loan',
                },
                {
                    target: '#loanTab',
                    title: 'Loan Management',
                    text: 'Add loans with principal, interest rate, and term. The app generates a full amortization schedule and syncs payments to your journal.',
                    position: 'left',
                    tab: 'loan',
                },
            ],
        },

        planning: {
            label: 'Planning',
            description: 'Budget, break-even, and projections',
            steps: [
                {
                    target: '.main-tab[data-tab="budget"]',
                    title: 'Budget Tab',
                    text: 'Plan monthly expenses and link them to journal categories.',
                    position: 'right',
                    tab: 'budget',
                },
                {
                    target: '#budgetTab',
                    title: 'Budget Planning',
                    text: 'Add recurring monthly expenses with start/end dates. Link to categories, then record them to your journal in one click.',
                    position: 'left',
                    tab: 'budget',
                },
                {
                    target: '.main-tab[data-tab="breakeven"]',
                    title: 'Break-Even Tab',
                    text: 'Analyze when your business becomes profitable with break-even analysis.',
                    position: 'right',
                    tab: 'breakeven',
                },
                {
                    target: '#breakevenTab',
                    title: 'Break-Even Analysis',
                    text: 'Configure sales channels (Consumer & B2B) with pricing and COGS. See break-even charts, contribution margins, and monthly projections.',
                    position: 'left',
                    tab: 'breakeven',
                },
                {
                    target: '.main-tab[data-tab="projectedsales"]',
                    title: 'Projected Sales Tab',
                    text: 'Forecast future sales with multiple channels and sales tax.',
                    position: 'right',
                    tab: 'projectedsales',
                },
                {
                    target: '#projectedsalesTab',
                    title: 'Sales Projections',
                    text: 'Set projected units per month for Online and Tradeshow channels. These projections feed into Cash Flow and P&L reports.',
                    position: 'left',
                    tab: 'projectedsales',
                },
            ],
        },

        sales: {
            label: 'Products & Sales',
            description: 'Product catalog and VE sales imports',
            steps: [
                {
                    target: '.main-tab[data-tab="products"]',
                    title: 'Products Tab',
                    text: 'Manage your product catalog with pricing, COGS, and analytics.',
                    position: 'right',
                    tab: 'products',
                },
                {
                    target: '#productsTab',
                    title: 'Product Catalog',
                    text: 'Add products with SKU, selling price, and COGS. Import products in bulk via CSV. View analytics charts for units sold and revenue.',
                    position: 'left',
                    tab: 'products',
                },
                {
                    target: '.main-tab[data-tab="vesales"]',
                    title: 'VE Sales Tab',
                    text: 'Import and analyze sales data from Volusion (VE) sources.',
                    position: 'right',
                    tab: 'vesales',
                },
                {
                    target: '#vesalesTab',
                    title: 'VE Sales Dashboard',
                    text: 'Import sales from Store Manager Excel, Trade Show POS, or JSON exports. View breakdowns by source and product, then create journal entries from the data.',
                    position: 'left',
                    tab: 'vesales',
                },
            ],
        },

        settings: {
            label: 'Settings & Sync',
            description: 'Themes, sync, and data management',
            steps: [
                {
                    target: '#gearBtn',
                    title: 'Settings Panel',
                    text: 'Click the gear icon to open settings. Customize your theme, toggle dark mode, and set your financial timeline.',
                    position: 'top',
                },
                {
                    target: '#companyBtn',
                    title: 'Multi-Company Support',
                    text: 'Create and switch between multiple companies. Each has isolated data. Use "Manage" to rename or delete companies.',
                    position: 'right',
                },
                {
                    target: '#syncBtn',
                    title: 'Group Sync',
                    text: 'Collaborate with others! Create a sync group and share an invite code. Members can sync data in real time via the cloud.',
                    position: 'top',
                },
                {
                    target: '#manageTabsBtn',
                    title: 'Manage Tabs',
                    text: 'Show or hide sidebar tabs you don\'t need. Hidden tabs keep their data \u2014 you can always bring them back.',
                    position: 'right',
                },
                {
                    target: '#saveDbBtn',
                    title: 'Save & Load',
                    text: 'Save your database locally as a .db file. Load it on any device. "Save All" exports every company in one file.',
                    position: 'bottom',
                    tab: 'journal',
                },
            ],
        },
    },

    // ─── Lesson Definitions ─────────────────────────────────
    lessons: {
        firstEntry: {
            label: 'Your First Journal Entry',
            description: 'Create an entry step-by-step with guided input',
            steps: [],  // Built dynamically in _buildLessonSteps()
        },
        impactReports: {
            label: 'How Entries Impact Reports',
            description: 'See how one entry flows to Cash Flow, P&L, and more',
            steps: [],
        },
        salesCascade: {
            label: 'Sales Entry Cascade',
            description: 'Sales tax and COGS entries created automatically',
            steps: [],
        },
        budgetPipeline: {
            label: 'Budget to Journal Pipeline',
            description: 'Budget expenses auto-generate journal entries',
            steps: [],
        },
    },

    // ─── Build Lesson Steps ─────────────────────────────────
    _buildLessonSteps() {
        const now = new Date();
        const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
        const todayStr = now.toISOString().split('T')[0];
        const self = this;

        // ── Lesson 1: Your First Journal Entry ──────────────
        this.lessons.firstEntry.steps = [
            {
                target: '.app-header',
                title: 'Your First Journal Entry',
                html: '<p>Let\'s create your first journal entry together.</p><p>You\'ll fill out each field yourself \u2014 I\'ll guide you through what everything means.</p>',
                position: 'bottom',
                noSpotlight: true,
                tab: 'journal',
                btnLabel: 'Let\'s Go',
            },
            {
                target: '#newEntryBtn',
                title: 'Open the Entry Form',
                html: '<p><strong>Click the "+ New Entry" button</strong> to open the entry form.</p>',
                position: 'bottom',
                tab: 'journal',
                inputPrompt: {
                    target: '#newEntryBtn',
                    type: 'click',
                    advanceDelay: 600,  // Wait for modal open animation
                },
            },
            {
                target: '.date-input-wrapper',
                title: 'Entry Date',
                html: '<p>This is the date the transaction occurred.</p><p><strong>Click the "Today" button</strong> or type a date.</p>',
                position: 'bottom',
                inputPrompt: {
                    target: '#entryDate',
                    type: 'value',
                    tolerance: 'any',
                    placeholder: 'Pick a date',
                    alsoLift: ['#todayBtn'],
                    // Watch for Today button clicks that set the value programmatically
                    watchClicks: ['#todayBtn'],
                },
                prepare() {
                    // Make sure the entry modal is open
                    const modal = document.getElementById('entryModal');
                    if (modal && !modal.classList.contains('active')) {
                        if (typeof UI !== 'undefined') UI.showModal('entryModal');
                    }
                },
            },
            {
                target: '.category-input-wrapper',
                title: 'Category',
                html: '<p>Categories organize your entries by type (e.g., "Sales Revenue", "Rent").</p><p><strong>Select "[Tutorial] Product Sales"</strong> from the dropdown.</p>',
                position: 'bottom',
                async action() {
                    // Create tutorial category if needed
                    const catId = self._getOrCreateTutorialCategory('Product Sales', 'receivable', {});
                    // Refresh categories dropdown
                    if (typeof App !== 'undefined' && App.refreshCategories) {
                        App.refreshCategories();
                    }
                    self._lessonCategoryId = catId;
                    // Wait for dropdown to populate
                    await new Promise(r => setTimeout(r, 100));
                },
                inputPrompt: {
                    target: '#category',
                    type: 'select',
                    get expectedValue() { return String(self._lessonCategoryId || ''); },
                },
            },
            {
                target: '#amount',
                title: 'Amount',
                html: '<p>This is the total dollar amount of the transaction.</p><p><strong>Type <code>1000</code></strong> \u2014 this represents a $1,000 sale.</p>',
                position: 'bottom',
                inputPrompt: {
                    target: '#amount',
                    type: 'value',
                    expectedValue: '1000',
                    placeholder: 'Type 1000',
                },
            },
            {
                target: '.radio-group',
                title: 'Transaction Type',
                html: '<p><strong>Receivable</strong> = money owed <em>to you</em> (income, sales).</p><p><strong>Payable</strong> = money <em>you owe</em> (rent, supplies, expenses).</p><p><strong>Make sure "Receivable" is selected</strong> \u2014 this is a sale.</p>',
                position: 'bottom',
                inputPrompt: {
                    target: 'input[name="transactionType"][value="receivable"]',
                    type: 'radio',
                    expectedValue: 'receivable',
                },
            },
            {
                target: '#status',
                title: 'Status',
                html: '<p><strong>Pending</strong> = not yet paid/received. The obligation exists but no cash has moved.</p><p><strong>Received/Paid</strong> = cash is in hand (or has been sent).</p><p><strong>Select "Pending"</strong> for now \u2014 we\'ll receive it later.</p>',
                position: 'bottom',
                inputPrompt: {
                    target: '#status',
                    type: 'select',
                    expectedValue: 'pending',
                },
            },
            {
                target: '#monthDue',
                title: 'Month Due',
                html: `<p>Month Due = when this obligation exists. This drives your <strong>P&L report</strong> (accrual accounting).</p><p><strong>Set it to the current month</strong> (${currentMonth}).</p>`,
                position: 'bottom',
                inputPrompt: {
                    target: '#monthDue',
                    type: 'value',
                    tolerance: 'any',
                    placeholder: currentMonth,
                },
            },
            {
                target: '#notes',
                title: 'Notes (Optional)',
                html: '<p>Add any notes to help you remember what this entry is for.</p><p><strong>Type anything you like</strong>, or just write "My first entry".</p>',
                position: 'bottom',
                inputPrompt: {
                    target: '#notes',
                    type: 'value',
                    tolerance: 'any',
                    placeholder: 'My first entry',
                },
            },
            {
                target: '#submitBtn',
                title: 'Submit Your Entry',
                html: '<p>Your entry is ready!</p><p><strong>Click "Add Entry"</strong> to save it to your journal.</p>',
                position: 'top',
                inputPrompt: {
                    target: '#submitBtn',
                    type: 'click',
                    advanceDelay: 800,  // Wait for form submit, modal close, and refreshAll
                },
            },
            {
                target: '.summary-section',
                title: 'Summary Cards Updated!',
                html: '<p>Look at the <strong>Accounts Receivable</strong> card \u2014 it increased by $1,000.</p><p>Cash Balance didn\'t change yet because the entry is still <em>Pending</em>. No cash has moved.</p>',
                position: 'bottom',
                tab: 'journal',
                impactBanner: 'A/R increased by $1,000 \u2014 you have a pending receivable!',
                highlight: ['#receivablesCard'],
                async prepare() {
                    // Track the transaction that was just created by the form submission
                    await new Promise(r => setTimeout(r, 200));
                    if (typeof Database !== 'undefined') {
                        const result = Database.db.exec('SELECT id FROM transactions ORDER BY id DESC LIMIT 1');
                        if (result.length > 0 && result[0].values.length > 0) {
                            self._sampleData.transactionIds.push(result[0].values[0][0]);
                            self._lastEntryId = result[0].values[0][0];
                        }
                    }
                },
            },
            {
                target: '#transactionsContainer',
                title: 'Your Entry in the Journal',
                html: '<p>Here\'s your entry in the journal table. You can click on it anytime to edit.</p>',
                position: 'top',
                tab: 'journal',
            },
            {
                target: '.summary-section',
                title: 'Marking as Received',
                html: '<p>Now let\'s mark this entry as <strong>Received</strong> to see how the Cash Balance changes.</p><p>I\'ll update it for you...</p>',
                position: 'bottom',
                tab: 'journal',
                btnLabel: 'Mark as Received',
                async action() {
                    // Update the entry to received status
                    if (typeof Database !== 'undefined' && self._lastEntryId) {
                        const rows = Database.db.exec('SELECT * FROM transactions WHERE id = ?', [self._lastEntryId]);
                        if (rows.length > 0 && rows[0].values.length > 0) {
                            const cols = rows[0].columns;
                            const vals = rows[0].values[0];
                            const tx = {};
                            cols.forEach((c, i) => tx[c] = vals[i]);
                            tx.status = 'received';
                            tx.month_paid = currentMonth;
                            tx.date_processed = todayStr;
                            Database.updateTransaction(self._lastEntryId, tx);
                        }
                    }
                    if (typeof App !== 'undefined' && App.refreshAll) App.refreshAll();
                    await new Promise(r => setTimeout(r, 200));
                },
            },
            {
                target: '.summary-section',
                title: 'Cash Balance Updated!',
                html: '<p><strong>Cash Balance increased</strong> by $1,000 and <strong>A/R decreased</strong> back down.</p><p>The money moved from "owed to you" to "in hand". This is the core of double-entry accounting.</p>',
                position: 'bottom',
                tab: 'journal',
                impactBanner: 'Cash Balance went up, A/R went down \u2014 cash received!',
                highlight: ['#cashBalanceCard', '#receivablesCard'],
            },
            {
                target: '.app-header',
                title: 'Great Job!',
                html: '<p>You created and received your first journal entry!</p><p><strong>Next up:</strong> Try <em>"How Entries Impact Reports"</em> to see how this entry flows to Cash Flow, P&L, Balance Sheet, and more.</p>',
                position: 'bottom',
                noSpotlight: true,
                btnLabel: 'Finish',
            },
        ];

        // ── Lesson 2: How Entries Impact Reports ────────────
        this.lessons.impactReports.steps = [
            {
                target: '.app-header',
                title: 'How Entries Impact Reports',
                html: '<p>Let\'s trace how a single journal entry flows through <strong>all your financial reports</strong>.</p><p>I\'ll create a sample entry and then walk you through Cash Flow, P&L, and Balance Sheet to show exactly where it appears.</p>',
                position: 'bottom',
                noSpotlight: true,
                tab: 'journal',
                btnLabel: 'Let\'s Go',
            },
            {
                target: '#transactionsContainer',
                title: 'Creating a Sample Entry',
                html: '<p>I just created a <strong>$2,000 Consulting Revenue</strong> receivable, marked as <em>Received</em> this month.</p><p>This means cash has been collected \u2014 both month_due and month_paid are set to this month.</p>',
                position: 'top',
                tab: 'journal',
                async action() {
                    const catId = self._getOrCreateTutorialCategory('Consulting Revenue', 'receivable', {});
                    const txId = Database.addTransaction({
                        entry_date: todayStr,
                        category_id: catId,
                        item_description: '[Tutorial] Consulting Revenue',
                        amount: 2000,
                        transaction_type: 'receivable',
                        status: 'received',
                        month_due: currentMonth,
                        month_paid: currentMonth,
                        date_processed: todayStr,
                        payment_for_month: currentMonth,
                        source_type: 'tutorial',
                        notes: 'Tutorial sample - consulting revenue',
                    });
                    self._sampleData.transactionIds.push(txId);
                    if (typeof App !== 'undefined' && App.refreshAll) App.refreshAll();
                    await new Promise(r => setTimeout(r, 200));
                },
            },
            {
                target: '.summary-section',
                title: 'Summary Cards React',
                html: '<p><strong>Cash Balance</strong> went up by $2,000 because the entry is marked Received.</p><p>If it were still Pending, only A/R would increase \u2014 Cash Balance wouldn\'t change.</p>',
                position: 'bottom',
                tab: 'journal',
                impactBanner: 'Cash Balance +$2,000 \u2014 entry is Received, so cash moved.',
                highlight: ['#cashBalanceCard'],
            },
            {
                target: '.app-header',
                title: 'Next: Cash Flow',
                html: '<p>Now let\'s see how this entry appears on the <strong>Cash Flow</strong> report...</p>',
                position: 'bottom',
                noSpotlight: true,
                btnLabel: 'See Cash Flow',
            },
            {
                target: '#cashflowTab',
                title: 'Cash Flow Statement',
                html: '<p>Cash Flow groups entries by <strong>month_paid</strong> \u2014 the month cash actually moved.</p><p>Your $2,000 appears here under the current month because that\'s when it was received.</p><p><em>Key insight: Cash Flow = cash basis accounting.</em></p>',
                position: 'left',
                tab: 'cashflow',
                impactBanner: '$2,000 shows in Cash Flow under the month it was paid/received.',
            },
            {
                target: '.app-header',
                title: 'Next: Profit & Loss',
                html: '<p>Now let\'s check the <strong>Profit & Loss</strong> report...</p>',
                position: 'bottom',
                noSpotlight: true,
                btnLabel: 'See P&L',
            },
            {
                target: '#pnlTab',
                title: 'Profit & Loss Statement',
                html: '<p>P&L groups entries by <strong>month_due</strong> \u2014 when the revenue was earned or expense was incurred.</p><p>Your $2,000 appears as revenue for this month.</p><p><em>Key insight: P&L = accrual basis accounting.</em></p>',
                position: 'left',
                tab: 'pnl',
                impactBanner: '$2,000 shows as revenue in P&L under the month_due.',
            },
            {
                target: '.app-header',
                title: 'Now Let\'s Add an Expense',
                html: '<p>To see the full picture, let\'s add an <strong>expense</strong> (payable) and see how both income and expenses appear on reports.</p>',
                position: 'bottom',
                noSpotlight: true,
                btnLabel: 'Add Expense',
                async action() {
                    const catId = self._getOrCreateTutorialCategory('Office Rent', 'payable', {});
                    const txId = Database.addTransaction({
                        entry_date: todayStr,
                        category_id: catId,
                        item_description: '[Tutorial] Office Rent',
                        amount: 300,
                        transaction_type: 'payable',
                        status: 'paid',
                        month_due: currentMonth,
                        month_paid: currentMonth,
                        date_processed: todayStr,
                        payment_for_month: currentMonth,
                        source_type: 'tutorial',
                        notes: 'Tutorial sample - office rent',
                    });
                    self._sampleData.transactionIds.push(txId);
                    if (typeof App !== 'undefined' && App.refreshAll) App.refreshAll();
                    await new Promise(r => setTimeout(r, 200));
                },
            },
            {
                target: '#transactionsContainer',
                title: 'Two Entries in Journal',
                html: '<p>Now you have two entries:</p><ul><li><strong>$2,000 receivable</strong> (income)</li><li><strong>$300 payable</strong> (expense)</li></ul>',
                position: 'top',
                tab: 'journal',
            },
            {
                target: '#cashflowTab',
                title: 'Cash Flow: Income & Expense',
                html: '<p>Cash Flow now shows <strong>both sides</strong>:</p><ul><li>Cash in: $2,000</li><li>Cash out: $300</li><li>Net: $1,700</li></ul><p>This is the actual cash movement for the month.</p>',
                position: 'left',
                tab: 'cashflow',
                impactBanner: 'Net Cash Flow = $2,000 in - $300 out = $1,700',
            },
            {
                target: '#pnlTab',
                title: 'P&L: Revenue vs Expenses',
                html: '<p>P&L shows your profitability:</p><ul><li>Revenue: $2,000</li><li>Operating Expenses: $300</li><li>Net Income: $1,700</li></ul><p>This is your accrual-basis income statement.</p>',
                position: 'left',
                tab: 'pnl',
                impactBanner: 'Net Income = Revenue ($2,000) - Expenses ($300) = $1,700',
            },
            {
                target: '#balancesheetTab',
                title: 'Balance Sheet',
                html: '<p>The Balance Sheet shows your financial position at a point in time:</p><ul><li><strong>Assets</strong> (what you own) increased</li><li><strong>Equity</strong> reflects retained earnings</li><li>Assets = Liabilities + Equity</li></ul>',
                position: 'left',
                tab: 'balancesheet',
                impactBanner: 'Balance Sheet reflects cumulative impact of all entries.',
            },
            {
                target: '.app-header',
                title: 'Key Takeaway',
                html: '<p><strong>month_due</strong> drives the P&L (when revenue was earned).</p><p><strong>month_paid</strong> drives Cash Flow (when cash moved).</p><p>These can be different months! That\'s the difference between <em>accrual</em> and <em>cash basis</em> accounting.</p>',
                position: 'bottom',
                noSpotlight: true,
                btnLabel: 'Finish',
            },
        ];

        // ── Lesson 3: Sales Entry Cascade ───────────────────
        this.lessons.salesCascade.steps = [
            {
                target: '.app-header',
                title: 'Sales Entry Cascade',
                html: '<p>When you use a <strong>Sales category</strong>, the system automatically creates linked entries:</p><ul><li><strong>Sales Tax</strong> payable (from the difference between total and pretax amount)</li><li><strong>Inventory Cost</strong> payable (COGS \u2014 Cost of Goods Sold)</li></ul><p>Let\'s see this in action.</p>',
                position: 'bottom',
                noSpotlight: true,
                tab: 'journal',
                btnLabel: 'Let\'s Go',
            },
            {
                target: '.app-header',
                title: 'Creating a Sales Entry',
                html: '<p>I\'m creating a sales receivable:</p><ul><li>Total: <strong>$500</strong></li><li>Pretax: <strong>$450</strong> (so $50 is sales tax)</li><li>Inventory Cost: <strong>$200</strong> (COGS)</li></ul>',
                position: 'bottom',
                noSpotlight: true,
                btnLabel: 'Create Entry',
                async action() {
                    // Create sales category with is_sales flag
                    const folderId = self._getOrCreateTutorialFolder('Tutorial Sales', 'receivable');
                    const catId = self._getOrCreateTutorialCategory('Online Store Sales', 'receivable', { isSales: true, folderId });
                    self._lessonSalesCatId = catId;

                    // We need to create the entry with pretax and inventory cost
                    // Using Database.addTransaction directly, then manually trigger the linked entries
                    const txId = Database.addTransaction({
                        entry_date: todayStr,
                        category_id: catId,
                        item_description: '[Tutorial] Online Store Sales',
                        amount: 500,
                        pretax_amount: 450,
                        inventory_cost: 200,
                        transaction_type: 'receivable',
                        status: 'received',
                        month_due: currentMonth,
                        month_paid: currentMonth,
                        date_processed: todayStr,
                        payment_for_month: currentMonth,
                        source_type: 'tutorial',
                        notes: 'Tutorial sample - online store sale',
                    });
                    self._sampleData.transactionIds.push(txId);

                    // Trigger the auto-creation of Sales Tax and Inventory Cost entries
                    const formData = {
                        category_id: catId,
                        amount: 500,
                        pretax_amount: 450,
                        inventory_cost: 200,
                        entry_date: todayStr,
                        month_due: currentMonth,
                    };
                    if (typeof App !== 'undefined' && App._manageSalesTaxEntry) {
                        App._manageSalesTaxEntry(txId, formData);
                    }
                    if (typeof App !== 'undefined' && App._manageInventoryCostEntry) {
                        App._manageInventoryCostEntry(txId, formData);
                    }

                    // Track the auto-created linked entries
                    if (typeof Database !== 'undefined') {
                        const linkedTx = Database.db.exec(
                            "SELECT id FROM transactions WHERE source_id = ? AND source_type IN ('sales_tax', 'inventory_cost')",
                            [txId]
                        );
                        if (linkedTx.length > 0) {
                            linkedTx[0].values.forEach(row => self._sampleData.transactionIds.push(row[0]));
                        }
                    }

                    if (typeof App !== 'undefined' && App.refreshAll) App.refreshAll();
                    await new Promise(r => setTimeout(r, 300));
                },
            },
            {
                target: '#transactionsContainer',
                title: 'Three Entries Created!',
                html: '<p>Look \u2014 <strong>one sale created three entries</strong>:</p><ol><li><strong>$500 receivable</strong> \u2014 the sale itself</li><li><strong>$50 payable</strong> \u2014 Sales Tax (auto-created)</li><li><strong>$200 payable</strong> \u2014 Inventory Cost / COGS (auto-created)</li></ol>',
                position: 'top',
                tab: 'journal',
                impactBanner: 'One sale = 3 journal entries, automatically linked!',
            },
            {
                target: '#transactionsContainer',
                title: 'Sales Tax Entry',
                html: '<p>The <strong>Sales Tax entry</strong> is the difference between the total ($500) and pretax ($450) = <strong>$50</strong>.</p><p>It\'s created as a <em>payable</em> because you owe this tax to the government.</p>',
                position: 'top',
                tab: 'journal',
            },
            {
                target: '#transactionsContainer',
                title: 'Inventory Cost (COGS)',
                html: '<p>The <strong>Inventory Cost entry</strong> ($200) represents your Cost of Goods Sold.</p><p>It\'s also a <em>payable</em> \u2014 this is what you spent to produce/buy the product you sold.</p><p><strong>Gross Margin</strong> = Pretax Revenue ($450) - COGS ($200) = <strong>$250</strong></p>',
                position: 'top',
                tab: 'journal',
            },
            {
                target: '.app-header',
                title: 'Impact on P&L',
                html: '<p>Let\'s see how this looks on the Profit & Loss statement...</p>',
                position: 'bottom',
                noSpotlight: true,
                btnLabel: 'See P&L',
            },
            {
                target: '#pnlTab',
                title: 'P&L: Revenue, COGS, Gross Margin',
                html: '<p>On the P&L, you can see:</p><ul><li><strong>Revenue</strong>: $450 (pretax sales)</li><li><strong>COGS</strong>: $200 (inventory cost)</li><li><strong>Gross Margin</strong>: $250</li></ul><p>Sales categories automatically separate revenue from costs on your P&L.</p>',
                position: 'left',
                tab: 'pnl',
                impactBanner: 'Revenue $450 - COGS $200 = Gross Margin $250',
            },
            {
                target: '#cashflowTab',
                title: 'Cash Flow Impact',
                html: '<p>Cash Flow shows the actual cash movement:</p><ul><li>Cash in: $500 (full sale amount)</li><li>Cash out: $50 (tax) + $200 (COGS) = $250</li><li>Net: $250</li></ul>',
                position: 'left',
                tab: 'cashflow',
            },
            {
                target: '.app-header',
                title: 'Sales Automation Complete',
                html: '<p>Sales categories save you time by <strong>automatically creating linked entries</strong> for tax and inventory costs.</p><p>This keeps your P&L accurate with proper revenue, COGS, and gross margin separation.</p>',
                position: 'bottom',
                noSpotlight: true,
                btnLabel: 'Finish',
            },
        ];

        // ── Lesson 4: Budget to Journal Pipeline ────────────
        this.lessons.budgetPipeline.steps = [
            {
                target: '.app-header',
                title: 'Budget to Journal Pipeline',
                html: '<p>Budget expenses automatically generate <strong>pending journal entries</strong> for each month in their range.</p><p>This lets you plan future expenses and see them in your projections.</p>',
                position: 'bottom',
                noSpotlight: true,
                btnLabel: 'Let\'s Go',
            },
            {
                target: '#budgetTab',
                title: 'Creating a Budget Expense',
                html: '<p>I\'m creating a budget expense:</p><ul><li>Name: <strong>Marketing Budget</strong></li><li>Amount: <strong>$500/month</strong></li><li>Duration: <strong>3 months</strong></li></ul><p>This will auto-generate 3 pending journal entries.</p>',
                position: 'left',
                tab: 'budget',
                btnLabel: 'Create Budget',
                async action() {
                    // Create tutorial category for budget
                    const catId = self._getOrCreateTutorialCategory('Marketing Expenses', 'payable', {});
                    self._lessonBudgetCatId = catId;

                    // Create budget group
                    const groupId = Database.addBudgetGroup('[Tutorial] Operating Expenses');
                    self._sampleData.budgetGroupIds.push(groupId);

                    // Calculate 3-month range ending at current month
                    const endDate = new Date(now.getFullYear(), now.getMonth(), 1);
                    const startDate = new Date(endDate.getFullYear(), endDate.getMonth() - 2, 1);
                    const startMonth = `${startDate.getFullYear()}-${String(startDate.getMonth() + 1).padStart(2, '0')}`;
                    const endMonth = currentMonth;

                    // Create budget expense
                    const expId = Database.addBudgetExpense(
                        '[Tutorial] Marketing Budget',
                        500,
                        startMonth,
                        endMonth,
                        catId,
                        'Tutorial sample - marketing budget',
                        groupId
                    );
                    self._sampleData.budgetExpenseIds.push(expId);

                    // Trigger budget sync to create journal entries
                    if (typeof App !== 'undefined' && App.syncAllBudgetJournalEntries) {
                        App.syncAllBudgetJournalEntries();
                    }
                    if (typeof App !== 'undefined' && App.refreshAll) App.refreshAll();

                    // Track the auto-created budget transactions
                    await new Promise(r => setTimeout(r, 300));
                    if (typeof Database !== 'undefined') {
                        const budgetTx = Database.db.exec(
                            "SELECT id FROM transactions WHERE source_type = 'budget' AND source_id = ?",
                            [expId]
                        );
                        if (budgetTx.length > 0) {
                            budgetTx[0].values.forEach(row => self._sampleData.transactionIds.push(row[0]));
                        }
                    }
                },
            },
            {
                target: '#budgetTab',
                title: 'Budget Expense Created',
                html: '<p>Your budget expense is now visible on the Budget tab.</p><p>It shows <strong>$500/month</strong> for 3 months = <strong>$1,500 total</strong>.</p>',
                position: 'left',
                tab: 'budget',
            },
            {
                target: '.app-header',
                title: 'Auto-Generated Journal Entries',
                html: '<p>The budget sync just created <strong>pending payable entries</strong> in your journal \u2014 one for each month.</p><p>Let\'s go see them...</p>',
                position: 'bottom',
                noSpotlight: true,
                btnLabel: 'See Journal',
            },
            {
                target: '#transactionsContainer',
                title: 'Budget Entries in Journal',
                html: '<p>Three <strong>pending payable</strong> entries were auto-created:</p><ul><li>$500 \u2014 2 months ago</li><li>$500 \u2014 last month</li><li>$500 \u2014 this month</li></ul><p>Each is linked to the budget expense. They\'re pending because they haven\'t been paid yet.</p>',
                position: 'top',
                tab: 'journal',
                impactBanner: 'Budget sync created 3 pending payable entries automatically!',
            },
            {
                target: '.app-header',
                title: 'Impact on Reports',
                html: '<p>Pending entries still appear on <strong>P&L</strong> (accrual basis) but <strong>not on Cash Flow</strong> until they\'re marked as paid.</p><p>Let\'s check...</p>',
                position: 'bottom',
                noSpotlight: true,
                btnLabel: 'See P&L',
            },
            {
                target: '#pnlTab',
                title: 'P&L Shows Pending Expenses',
                html: '<p>P&L uses <strong>month_due</strong>, so these budget expenses appear as operating expenses even though they\'re unpaid.</p><p>This is accrual accounting \u2014 expenses are recorded when incurred, not when paid.</p>',
                position: 'left',
                tab: 'pnl',
                impactBanner: 'P&L shows $500/month in expenses (accrual basis).',
            },
            {
                target: '#cashflowTab',
                title: 'Cash Flow: No Impact Yet',
                html: '<p>Cash Flow uses <strong>month_paid</strong>. Since these entries are still <em>Pending</em>, they don\'t appear here yet.</p><p>When you mark them as Paid and set a month_paid, they\'ll show up in Cash Flow.</p>',
                position: 'left',
                tab: 'cashflow',
                impactBanner: 'Pending entries don\'t affect Cash Flow until paid.',
            },
            {
                target: '.app-header',
                title: 'The Budget Lifecycle',
                html: '<p>Here\'s how the budget pipeline works:</p><ol><li><strong>Create budget expense</strong> \u2192 sets the plan</li><li><strong>Auto-sync</strong> \u2192 creates pending journal entries</li><li><strong>Mark as Paid</strong> \u2192 when you actually pay, set status to Paid</li><li><strong>Cash Flow updates</strong> \u2192 reflects real cash movement</li></ol>',
                position: 'bottom',
                noSpotlight: true,
            },
            {
                target: '.app-header',
                title: 'Budget Pipeline Complete',
                html: '<p>Budget automation saves you from creating the same entries every month. Just set it up once and the system handles the rest!</p>',
                position: 'bottom',
                noSpotlight: true,
                btnLabel: 'Finish',
            },
        ];
    },

    // ─── Sample Data Helpers ────────────────────────────────
    _getOrCreateTutorialCategory(name, folderType, flags = {}) {
        const fullName = `[Tutorial] ${name}`;
        // Check if it already exists
        const existing = Database.db.exec('SELECT id FROM categories WHERE name = ?', [fullName]);
        if (existing.length > 0 && existing[0].values.length > 0) {
            return existing[0].values[0][0];
        }

        const catId = Database.addCategory(
            fullName,
            false,                     // isMonthly
            null,                      // defaultAmount
            folderType || null,        // defaultType
            flags.folderId || null,    // folderId
            false,                     // showOnPl
            false,                     // isCogs
            false,                     // isDepreciation
            false,                     // isSalesTax
            false,                     // isB2b
            null,                      // defaultStatus
            flags.isSales || false,    // isSales
            false                      // isInventoryCost
        );
        this._sampleData.categoryIds.push(catId);
        return catId;
    },

    _getOrCreateTutorialFolder(name, type) {
        const fullName = `[Tutorial] ${name}`;
        const existing = Database.db.exec('SELECT id FROM category_folders WHERE name = ?', [fullName]);
        if (existing.length > 0 && existing[0].values.length > 0) {
            return existing[0].values[0][0];
        }
        const folderId = Database.addFolder(fullName, type);
        this._sampleData.folderIds.push(folderId);
        return folderId;
    },

    _cleanupSampleData() {
        if (typeof Database === 'undefined') return;

        // Delete transactions first (categories can't be deleted if in use)
        this._sampleData.transactionIds.forEach(id => {
            try { Database.deleteTransaction(id); } catch (e) { /* ignore */ }
        });

        // Delete budget expenses
        this._sampleData.budgetExpenseIds.forEach(id => {
            try { Database.deleteBudgetExpense(id); } catch (e) { /* ignore */ }
        });

        // Delete budget groups
        this._sampleData.budgetGroupIds.forEach(id => {
            try { Database.deleteBudgetGroup(id); } catch (e) { /* ignore */ }
        });

        // Also clean up any budget-synced transactions that reference tutorial budget expenses
        try {
            const tutorialBudgetTx = Database.db.exec(
                "SELECT id FROM transactions WHERE source_type = 'tutorial'"
            );
            if (tutorialBudgetTx.length > 0) {
                tutorialBudgetTx[0].values.forEach(row => {
                    try { Database.deleteTransaction(row[0]); } catch (e) { /* ignore */ }
                });
            }
        } catch (e) { /* ignore */ }

        // Delete categories (now safe since transactions are gone)
        this._sampleData.categoryIds.forEach(id => {
            try { Database.deleteCategory(id); } catch (e) { /* ignore */ }
        });

        // Delete folders
        this._sampleData.folderIds.forEach(id => {
            try {
                Database.db.run('DELETE FROM category_folders WHERE id = ?', [id]);
                Database.autoSave();
            } catch (e) { /* ignore */ }
        });

        // Reset tracker
        this._sampleData = {
            transactionIds: [],
            categoryIds: [],
            budgetExpenseIds: [],
            budgetGroupIds: [],
            folderIds: [],
        };

        if (typeof App !== 'undefined' && App.refreshAll) App.refreshAll();
    },

    _showCleanupPrompt() {
        const hasData = this._sampleData.transactionIds.length > 0 ||
                       this._sampleData.categoryIds.length > 0 ||
                       this._sampleData.budgetExpenseIds.length > 0 ||
                       this._sampleData.budgetGroupIds.length > 0 ||
                       this._sampleData.folderIds.length > 0;
        if (!hasData) return;

        const prompt = document.createElement('div');
        prompt.className = 'tutorial-cleanup-modal';
        prompt.innerHTML = `
            <div class="tutorial-cleanup-content">
                <div class="tutorial-cleanup-title">Lesson Complete!</div>
                <div class="tutorial-cleanup-text">This lesson created sample data to demonstrate features. What would you like to do?</div>
                <div class="tutorial-cleanup-actions">
                    <button class="btn btn-secondary btn-small tutorial-cleanup-keep">Keep for Exploration</button>
                    <button class="btn btn-primary btn-small tutorial-cleanup-remove">Remove Sample Data</button>
                </div>
            </div>
        `;
        document.body.appendChild(prompt);

        prompt.querySelector('.tutorial-cleanup-remove').addEventListener('click', () => {
            this._cleanupSampleData();
            prompt.remove();
        });
        prompt.querySelector('.tutorial-cleanup-keep').addEventListener('click', () => {
            // Reset tracker but keep data
            this._sampleData = {
                transactionIds: [],
                categoryIds: [],
                budgetExpenseIds: [],
                budgetGroupIds: [],
                folderIds: [],
            };
            prompt.remove();
        });
    },

    // ─── Guided Input System ────────────────────────────────
    _setupInputPrompt(step) {
        const prompt = step.inputPrompt;
        if (!prompt) return;

        const target = document.querySelector(prompt.target);
        if (!target) return;

        // Track all elements we lift so we can clean them up
        this._liftedEls = [];

        const liftEl = (el) => {
            if (el && !el.classList.contains('tutorial-input-active')) {
                el.classList.add('tutorial-input-active');
                this._liftedEls.push(el);
            }
        };

        // Lift target above overlay so user can interact
        liftEl(target);

        // Lift parent containers up to (and including) any modal
        const formGroup = target.closest('.form-group');
        if (formGroup) liftEl(formGroup);

        const radioGroup = target.closest('.radio-group');
        if (radioGroup) liftEl(radioGroup);

        // Lift wrapper containers (date-input-wrapper, category-input-wrapper, etc.)
        const wrapper = target.closest('[class*="-wrapper"]');
        if (wrapper) liftEl(wrapper);

        // Lift all labels inside the radio group so they're clickable
        if (radioGroup) {
            radioGroup.querySelectorAll('.radio-label').forEach(label => liftEl(label));
        }

        // CRITICAL: If target is inside a modal, lift the ENTIRE modal above the overlay
        const modal = target.closest('.modal');
        if (modal) {
            modal.style.zIndex = '5003';
            this._liftedModal = modal;
        }

        // For "alsoLift" selectors — lift sibling elements the user needs (e.g., Today button)
        if (prompt.alsoLift) {
            prompt.alsoLift.forEach(sel => {
                const el = document.querySelector(sel);
                if (el) liftEl(el);
            });
        }

        // Save and set placeholder
        if (prompt.placeholder && target.placeholder !== undefined) {
            this._savedPlaceholder = target.placeholder;
            target.placeholder = prompt.placeholder;
        }

        // Disable Next button
        const nextBtn = this.tooltipEl.querySelector('.tutorial-btn-next');
        nextBtn.disabled = true;
        nextBtn.classList.add('tutorial-btn-waiting');
        nextBtn.textContent = 'Waiting...';

        // Disable overlay click advance and keyboard advance
        this._inputPromptActive = true;

        // Create hint element
        const hintEl = document.createElement('div');
        hintEl.className = 'tutorial-input-hint';
        hintEl.style.display = 'none';
        this.tooltipEl.querySelector('.tutorial-tooltip-body').appendChild(hintEl);

        const enableNext = () => {
            nextBtn.disabled = false;
            nextBtn.classList.remove('tutorial-btn-waiting');
            nextBtn.textContent = step.btnLabel || 'Continue';
            nextBtn.classList.add('tutorial-btn-ready');
            target.classList.add('tutorial-input-correct');
            setTimeout(() => target.classList.remove('tutorial-input-correct'), 600);
            this._inputPromptActive = false;
        };

        const checkValue = () => {
            if (prompt.type === 'click') return; // Handled by click listener

            let currentVal = '';
            if (prompt.type === 'select') {
                currentVal = target.value;
            } else if (prompt.type === 'radio') {
                currentVal = target.checked ? target.value : '';
            } else {
                currentVal = target.value;
            }

            if (prompt.tolerance === 'any') {
                if (currentVal && currentVal.trim() !== '') {
                    enableNext();
                    return true;
                }
            } else {
                const expected = typeof prompt.expectedValue === 'string' ? prompt.expectedValue : String(prompt.expectedValue || '');
                if (currentVal === expected) {
                    enableNext();
                    return true;
                } else if (currentVal && currentVal.trim() !== '' && expected) {
                    // Show hint
                    hintEl.innerHTML = `<small>Hint: try entering <strong>${expected}</strong></small>`;
                    hintEl.style.display = 'block';
                }
            }
            return false;
        };

        if (prompt.type === 'click') {
            // Determine auto-advance delay (longer for clicks that trigger DOM changes)
            const advanceDelay = prompt.advanceDelay || 400;

            const clickHandler = (e) => {
                enableNext();
                target.removeEventListener('click', clickHandler);
                // Auto-advance after a delay to let DOM changes settle
                setTimeout(() => {
                    if (this.state.active) this.next();
                }, advanceDelay);
            };
            target.addEventListener('click', clickHandler);

            // For radio inputs, also listen on the parent label (users click labels)
            let labelHandler = null;
            const parentLabel = target.closest('label.radio-label');
            if (parentLabel && target !== parentLabel) {
                labelHandler = (e) => {
                    if (e.target === target) return; // Avoid double-fire
                    enableNext();
                    target.removeEventListener('click', clickHandler);
                    parentLabel.removeEventListener('click', labelHandler);
                    setTimeout(() => {
                        if (this.state.active) this.next();
                    }, advanceDelay);
                };
                parentLabel.addEventListener('click', labelHandler);
            }

            this._inputCleanup = () => {
                target.removeEventListener('click', clickHandler);
                if (parentLabel && labelHandler) parentLabel.removeEventListener('click', labelHandler);
                this._cleanupLifted();
                hintEl.remove();
                this._inputPromptActive = false;
            };
        } else {
            const events = prompt.type === 'select' ? ['change'] : prompt.type === 'radio' ? ['change'] : ['input', 'change'];
            const handler = () => checkValue();
            events.forEach(evt => target.addEventListener(evt, handler));

            // Watch for external buttons that set the value programmatically (e.g., "Today" button)
            const watchClickHandlers = [];
            if (prompt.watchClicks) {
                prompt.watchClicks.forEach(sel => {
                    const btn = document.querySelector(sel);
                    if (btn) {
                        const wcHandler = () => {
                            // Delay check to let the button's own handler set the value first
                            setTimeout(() => checkValue(), 50);
                        };
                        btn.addEventListener('click', wcHandler);
                        watchClickHandlers.push({ el: btn, handler: wcHandler });
                    }
                });
            }

            // Also check initially in case value is pre-filled
            setTimeout(() => checkValue(), 100);

            this._inputCleanup = () => {
                events.forEach(evt => target.removeEventListener(evt, handler));
                watchClickHandlers.forEach(({ el, handler }) => el.removeEventListener('click', handler));
                this._cleanupLifted();
                if (prompt.placeholder && this._savedPlaceholder !== undefined) {
                    target.placeholder = this._savedPlaceholder;
                }
                hintEl.remove();
                this._inputPromptActive = false;
            };
        }
    },

    _cleanupLifted() {
        // Remove tutorial-input-active from all lifted elements
        if (this._liftedEls) {
            this._liftedEls.forEach(el => el.classList.remove('tutorial-input-active'));
            this._liftedEls = [];
        }
        // Restore modal z-index
        if (this._liftedModal) {
            this._liftedModal.style.zIndex = '';
            this._liftedModal = null;
        }
    },

    _teardownInputPrompt() {
        if (this._inputCleanup) {
            this._inputCleanup();
            this._inputCleanup = null;
        }
        // Also cleanup lifted elements in case cleanup didn't run
        this._cleanupLifted();
        // Reset Next button state
        const nextBtn = this.tooltipEl.querySelector('.tutorial-btn-next');
        if (nextBtn) {
            nextBtn.disabled = false;
            nextBtn.classList.remove('tutorial-btn-waiting', 'tutorial-btn-ready');
        }
        this._savedPlaceholder = undefined;
    },

    // ─── Initialization ──────────────────────────────────────
    init() {
        this._buildLessonSteps();
        this.createDOMElements();
        this.bindEvents();

        // First-visit welcome
        if (!localStorage.getItem('tutorialSeen')) {
            setTimeout(() => this.showWelcome(), 800);
        }
    },

    createDOMElements() {
        // Overlay (blocks clicks outside spotlight)
        this.overlayEl = document.createElement('div');
        this.overlayEl.id = 'tutorial-overlay';
        document.body.appendChild(this.overlayEl);

        // Spotlight cutout
        this.spotlightEl = document.createElement('div');
        this.spotlightEl.id = 'tutorial-spotlight';
        document.body.appendChild(this.spotlightEl);

        // Tooltip
        this.tooltipEl = document.createElement('div');
        this.tooltipEl.id = 'tutorial-tooltip';
        this.tooltipEl.className = 'tutorial-tooltip';
        this.tooltipEl.innerHTML = `
            <div class="tutorial-tooltip-arrow"></div>
            <div class="tutorial-tooltip-header">
                <span class="tutorial-tooltip-title"></span>
                <button class="tutorial-tooltip-close" title="Exit tutorial">&times;</button>
            </div>
            <div class="tutorial-tooltip-impact" style="display:none"></div>
            <div class="tutorial-tooltip-body"></div>
            <div class="tutorial-tooltip-footer">
                <div class="tutorial-tooltip-progress">
                    <span class="tutorial-tooltip-step-text"></span>
                    <div class="tutorial-tooltip-progress-bar">
                        <div class="tutorial-tooltip-progress-fill"></div>
                    </div>
                </div>
                <div class="tutorial-tooltip-nav">
                    <button class="btn btn-secondary btn-small tutorial-btn-prev">Back</button>
                    <button class="btn btn-primary btn-small tutorial-btn-next">Next</button>
                </div>
            </div>
        `;
        document.body.appendChild(this.tooltipEl);

        // Tour picker popover (with sections)
        this.pickerEl = document.createElement('div');
        this.pickerEl.id = 'tutorial-picker';
        this.pickerEl.className = 'tutorial-picker';
        this.pickerEl.style.display = 'none';

        let pickerHTML = '<div class="tutorial-picker-section">Quick Reference</div>';
        for (const [id, tour] of Object.entries(this.tours)) {
            pickerHTML += `
                <button class="tutorial-picker-item" data-tour="${id}">
                    <span class="tutorial-picker-item-label">${tour.label}</span>
                    <span class="tutorial-picker-item-desc">${tour.description}</span>
                </button>
            `;
        }
        pickerHTML += '<div class="tutorial-picker-divider"></div>';
        pickerHTML += '<div class="tutorial-picker-section">Interactive Lessons <span class="tutorial-picker-badge">Hands-on</span></div>';
        for (const [id, lesson] of Object.entries(this.lessons)) {
            pickerHTML += `
                <button class="tutorial-picker-item tutorial-picker-lesson" data-tour="${id}">
                    <span class="tutorial-picker-item-label">${lesson.label}</span>
                    <span class="tutorial-picker-item-desc">${lesson.description}</span>
                </button>
            `;
        }
        pickerHTML += `
            <div class="tutorial-picker-divider"></div>
            <button class="tutorial-picker-item tutorial-picker-full" data-tour="all">
                <span class="tutorial-picker-item-label">Full Tour</span>
                <span class="tutorial-picker-item-desc">Walk through every feature</span>
            </button>
        `;
        this.pickerEl.innerHTML = pickerHTML;
        document.body.appendChild(this.pickerEl);
    },

    bindEvents() {
        // Tooltip buttons
        this.tooltipEl.querySelector('.tutorial-tooltip-close').addEventListener('click', () => this.end());
        this.tooltipEl.querySelector('.tutorial-btn-next').addEventListener('click', () => this.next());
        this.tooltipEl.querySelector('.tutorial-btn-prev').addEventListener('click', () => this.prev());

        // Overlay click = next (but not during input prompts)
        this.overlayEl.addEventListener('click', () => {
            if (!this._inputPromptActive) this.next();
        });

        // Keyboard
        document.addEventListener('keydown', (e) => {
            if (!this.state.active) return;
            if (e.key === 'Escape') { this.end(); e.stopPropagation(); }
            else if (!this._inputPromptActive) {
                if (e.key === 'ArrowRight' || e.key === 'Enter') { this.next(); e.stopPropagation(); }
                else if (e.key === 'ArrowLeft') { this.prev(); e.stopPropagation(); }
            }
        });

        // Resize/scroll repositioning
        let resizeTimer;
        const reposition = () => {
            clearTimeout(resizeTimer);
            resizeTimer = setTimeout(() => {
                if (this.state.active) this.showStep(this.state.currentStepIndex, true);
            }, 100);
        };
        window.addEventListener('resize', reposition);
        document.querySelector('.app-main')?.addEventListener('scroll', reposition);

        // Help button
        document.getElementById('helpBtn')?.addEventListener('click', (e) => {
            e.stopPropagation();
            this.togglePicker();
        });

        // Tour picker item clicks
        this.pickerEl.addEventListener('click', (e) => {
            const item = e.target.closest('.tutorial-picker-item');
            if (item) {
                const tourId = item.dataset.tour;
                this.hidePicker();
                this.start(tourId);
            }
        });

        // Close picker when clicking outside
        document.addEventListener('click', (e) => {
            if (this.pickerEl.style.display !== 'none' && !this.pickerEl.contains(e.target) && e.target.id !== 'helpBtn') {
                this.hidePicker();
            }
        });
    },

    // ─── Tour Picker ─────────────────────────────────────────
    togglePicker() {
        if (this.pickerEl.style.display === 'none') {
            this.showPicker();
        } else {
            this.hidePicker();
        }
    },

    showPicker() {
        const btn = document.getElementById('helpBtn');
        if (!btn) return;
        const rect = btn.getBoundingClientRect();
        this.pickerEl.style.display = 'block';
        this.pickerEl.style.bottom = (window.innerHeight - rect.top + 8) + 'px';
        this.pickerEl.style.left = rect.left + 'px';
    },

    hidePicker() {
        this.pickerEl.style.display = 'none';
    },

    // ─── Welcome (First Visit) ──────────────────────────────
    showWelcome() {
        const btn = document.getElementById('helpBtn');
        if (!btn) return;

        // Create a simple welcome popover
        const welcome = document.createElement('div');
        welcome.id = 'tutorial-welcome';
        welcome.className = 'tutorial-welcome';
        welcome.innerHTML = `
            <div class="tutorial-welcome-title">Welcome to Accounting Journal!</div>
            <div class="tutorial-welcome-text">Want a quick tour of the features, or dive into a hands-on lesson?</div>
            <div class="tutorial-welcome-actions">
                <button class="btn btn-secondary btn-small tutorial-welcome-skip">Skip</button>
                <button class="btn btn-primary btn-small tutorial-welcome-start">Start Tour</button>
                <button class="btn btn-primary btn-small tutorial-welcome-lesson">Hands-on Lesson</button>
            </div>
        `;
        document.body.appendChild(welcome);

        // Position near help button
        const rect = btn.getBoundingClientRect();
        welcome.style.bottom = (window.innerHeight - rect.top + 8) + 'px';
        welcome.style.left = rect.left + 'px';

        welcome.querySelector('.tutorial-welcome-skip').addEventListener('click', () => {
            welcome.remove();
            localStorage.setItem('tutorialSeen', 'true');
        });
        welcome.querySelector('.tutorial-welcome-start').addEventListener('click', () => {
            welcome.remove();
            localStorage.setItem('tutorialSeen', 'true');
            this.start('quickStart');
        });
        welcome.querySelector('.tutorial-welcome-lesson').addEventListener('click', () => {
            welcome.remove();
            localStorage.setItem('tutorialSeen', 'true');
            this.start('firstEntry');
        });

        // Auto-dismiss on outside click
        const dismiss = (e) => {
            if (!welcome.contains(e.target) && e.target !== btn) {
                welcome.remove();
                localStorage.setItem('tutorialSeen', 'true');
                document.removeEventListener('click', dismiss);
            }
        };
        setTimeout(() => document.addEventListener('click', dismiss), 100);
    },

    // ─── Tour Control ────────────────────────────────────────
    start(tourId) {
        // Check lessons first
        if (this.lessons[tourId]) {
            this.state.steps = [...this.lessons[tourId].steps];
            this.state.isLesson = true;
        } else if (tourId === 'all') {
            // Flatten all tours into one sequence
            this.state.steps = [];
            for (const tour of Object.values(this.tours)) {
                this.state.steps.push(...tour.steps);
            }
            this.state.isLesson = false;
        } else {
            const tour = this.tours[tourId];
            if (!tour) return;
            this.state.steps = [...tour.steps];
            this.state.isLesson = false;
        }

        this.state.currentTourId = tourId;
        this.state.currentStepIndex = 0;
        this.state.active = true;
        localStorage.setItem('tutorialSeen', 'true');

        // Add lesson class for wider tooltips
        if (this.state.isLesson) {
            this.tooltipEl.classList.add('tutorial-lesson-tooltip');
        } else {
            this.tooltipEl.classList.remove('tutorial-lesson-tooltip');
        }

        this.overlayEl.classList.add('active');
        this.spotlightEl.style.display = 'block';
        this.tooltipEl.style.display = 'block';

        this.showStep(0);
    },

    async showStep(index, repositionOnly = false) {
        const step = this.state.steps[index];
        if (!step) { this.end(); return; }

        // Run cleanup on previous step
        if (!repositionOnly && this.state.currentStepIndex !== index) {
            const prev = this.state.steps[this.state.currentStepIndex];
            if (prev && prev.cleanup) prev.cleanup();
            this._teardownInputPrompt();
        }

        this.state.currentStepIndex = index;

        // Auto-switch tab if needed
        if (step.tab && typeof App !== 'undefined' && App.switchMainTab) {
            App.switchMainTab(step.tab);
        }

        // Run prepare function
        if (!repositionOnly && step.prepare) {
            await step.prepare();
        }

        // Run action (lesson-specific: creates data, etc.)
        if (!repositionOnly && step.action) {
            await step.action();
        }

        // Delay if specified
        if (!repositionOnly && step.delay) {
            await new Promise(r => setTimeout(r, step.delay));
        }

        // Let DOM settle
        await new Promise(r => setTimeout(r, 80));

        // Remove previous highlights and restore modal z-index from previous step
        document.querySelector('.tutorial-target-highlight')?.classList.remove('tutorial-target-highlight');
        document.querySelectorAll('.tutorial-impact-highlight').forEach(el => el.classList.remove('tutorial-impact-highlight'));
        if (this._stepModal) {
            this._stepModal.style.zIndex = '';
            this._stepModal = null;
        }

        // Handle noSpotlight steps (centered tooltip, no target needed)
        if (step.noSpotlight) {
            this.spotlightEl.style.display = 'none';
            this.updateTooltip(step, index);
            // Center tooltip on screen
            requestAnimationFrame(() => {
                const tipRect = this.tooltipEl.getBoundingClientRect();
                this.tooltipEl.style.top = Math.max(40, (window.innerHeight - tipRect.height) / 2) + 'px';
                this.tooltipEl.style.left = Math.max(40, (window.innerWidth - tipRect.width) / 2) + 'px';
                this.tooltipEl.removeAttribute('data-pos');
                this.tooltipEl.querySelector('.tutorial-tooltip-arrow').style.display = 'none';
            });

            // Set up input prompt if present
            if (!repositionOnly && step.inputPrompt) {
                this._setupInputPrompt(step);
            }

            // Apply impact highlights
            if (step.highlight) {
                step.highlight.forEach(sel => {
                    const el = document.querySelector(sel);
                    if (el) el.classList.add('tutorial-impact-highlight');
                });
            }
            return;
        }

        // Show spotlight
        this.spotlightEl.style.display = 'block';
        this.tooltipEl.querySelector('.tutorial-tooltip-arrow').style.display = '';

        const targetEl = document.querySelector(step.target);
        if (!targetEl || targetEl.offsetParent === null) {
            // Element not found or hidden - skip
            if (!repositionOnly) {
                if (index < this.state.steps.length - 1) {
                    this.showStep(index + 1);
                } else {
                    this.end();
                }
            }
            return;
        }

        // If target is inside a modal, lift the modal above the overlay
        const targetModal = targetEl.closest('.modal');
        if (targetModal) {
            targetModal.style.zIndex = '5003';
            this._stepModal = targetModal;
        }

        // Scroll into view
        targetEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

        // Add highlight class to lift target above overlay
        targetEl.classList.add('tutorial-target-highlight');

        // Position spotlight
        this.positionSpotlight(targetEl);

        // Update tooltip content
        this.updateTooltip(step, index);

        // Position tooltip (after a frame so dimensions are calculated)
        requestAnimationFrame(() => {
            this.positionTooltip(targetEl, step.position);
        });

        // Apply impact highlights
        if (step.highlight) {
            step.highlight.forEach(sel => {
                const el = document.querySelector(sel);
                if (el) el.classList.add('tutorial-impact-highlight');
            });
        }

        // Set up input prompt if present
        if (!repositionOnly && step.inputPrompt) {
            this._setupInputPrompt(step);
        }
    },

    next() {
        if (!this.state.active) return;
        if (this._inputPromptActive) return; // Block advance during input prompts
        if (this.state.currentStepIndex < this.state.steps.length - 1) {
            this.showStep(this.state.currentStepIndex + 1);
        } else {
            this.end();
        }
    },

    prev() {
        if (!this.state.active) return;
        if (this.state.currentStepIndex > 0) {
            this.showStep(this.state.currentStepIndex - 1);
        }
    },

    end() {
        // Cleanup current step
        const step = this.state.steps[this.state.currentStepIndex];
        if (step && step.cleanup) step.cleanup();
        this._teardownInputPrompt();

        // Restore any modal z-index we changed
        if (this._stepModal) {
            this._stepModal.style.zIndex = '';
            this._stepModal = null;
        }

        this.state.active = false;
        this.overlayEl.classList.remove('active');
        this.spotlightEl.style.display = 'none';
        this.tooltipEl.style.display = 'none';
        this.tooltipEl.classList.remove('tutorial-lesson-tooltip');

        // Remove highlights
        document.querySelector('.tutorial-target-highlight')?.classList.remove('tutorial-target-highlight');
        document.querySelectorAll('.tutorial-impact-highlight').forEach(el => el.classList.remove('tutorial-impact-highlight'));

        // Show cleanup prompt if lesson had sample data
        if (this.state.isLesson) {
            this._showCleanupPrompt();
        }
    },

    // ─── Spotlight Positioning ───────────────────────────────
    positionSpotlight(targetEl) {
        const rect = targetEl.getBoundingClientRect();
        const pad = 8;
        this.spotlightEl.style.top = (rect.top - pad) + 'px';
        this.spotlightEl.style.left = (rect.left - pad) + 'px';
        this.spotlightEl.style.width = (rect.width + pad * 2) + 'px';
        this.spotlightEl.style.height = (rect.height + pad * 2) + 'px';
        this.spotlightEl.style.borderRadius = window.getComputedStyle(targetEl).borderRadius || '8px';
    },

    // ─── Tooltip ─────────────────────────────────────────────
    updateTooltip(step, index) {
        const total = this.state.steps.length;
        this.tooltipEl.querySelector('.tutorial-tooltip-title').textContent = step.title;

        // Handle html vs text content
        const bodyEl = this.tooltipEl.querySelector('.tutorial-tooltip-body');
        if (step.html) {
            bodyEl.innerHTML = step.html;
        } else {
            bodyEl.textContent = step.text || '';
        }

        // Impact banner
        const impactEl = this.tooltipEl.querySelector('.tutorial-tooltip-impact');
        if (step.impactBanner) {
            impactEl.textContent = step.impactBanner;
            impactEl.style.display = 'block';
        } else {
            impactEl.style.display = 'none';
        }

        this.tooltipEl.querySelector('.tutorial-tooltip-step-text').textContent = `Step ${index + 1} of ${total}`;
        this.tooltipEl.querySelector('.tutorial-tooltip-progress-fill').style.width = ((index + 1) / total * 100) + '%';

        // Back button state
        const prevBtn = this.tooltipEl.querySelector('.tutorial-btn-prev');
        prevBtn.style.display = index === 0 ? 'none' : '';

        // Next button label
        const nextBtn = this.tooltipEl.querySelector('.tutorial-btn-next');
        if (step.btnLabel) {
            nextBtn.textContent = step.btnLabel;
        } else {
            nextBtn.textContent = index === total - 1 ? 'Done' : 'Next';
        }
    },

    positionTooltip(targetEl, preferred) {
        const rect = targetEl.getBoundingClientRect();
        const tip = this.tooltipEl;
        const tipRect = tip.getBoundingClientRect();
        const gap = 16; // gap between spotlight edge and tooltip
        const pad = 8;  // spotlight padding

        // Calculate available space in each direction
        const space = {
            top: rect.top - pad - gap,
            bottom: window.innerHeight - rect.bottom - pad - gap,
            left: rect.left - pad - gap,
            right: window.innerWidth - rect.right - pad - gap,
        };

        // Try positions in order: preferred, then fallbacks
        const order = [preferred, 'bottom', 'right', 'top', 'left'].filter((v, i, a) => a.indexOf(v) === i);
        let pos = preferred;

        for (const dir of order) {
            if (dir === 'bottom' && space.bottom >= tipRect.height) { pos = 'bottom'; break; }
            if (dir === 'top' && space.top >= tipRect.height) { pos = 'top'; break; }
            if (dir === 'right' && space.right >= tipRect.width) { pos = 'right'; break; }
            if (dir === 'left' && space.left >= tipRect.width) { pos = 'left'; break; }
        }

        let top, left;
        switch (pos) {
            case 'bottom':
                top = rect.bottom + pad + gap;
                left = rect.left + rect.width / 2 - tipRect.width / 2;
                break;
            case 'top':
                top = rect.top - pad - gap - tipRect.height;
                left = rect.left + rect.width / 2 - tipRect.width / 2;
                break;
            case 'right':
                top = rect.top + rect.height / 2 - tipRect.height / 2;
                left = rect.right + pad + gap;
                break;
            case 'left':
                top = rect.top + rect.height / 2 - tipRect.height / 2;
                left = rect.left - pad - gap - tipRect.width;
                break;
        }

        // Clamp to viewport
        top = Math.max(8, Math.min(top, window.innerHeight - tipRect.height - 8));
        left = Math.max(8, Math.min(left, window.innerWidth - tipRect.width - 8));

        tip.style.top = top + 'px';
        tip.style.left = left + 'px';
        tip.setAttribute('data-pos', pos);

        // Compute arrow offset so it points at the target's center
        const targetCenterX = rect.left + rect.width / 2;
        const targetCenterY = rect.top + rect.height / 2;
        let arrowOffset;

        if (pos === 'bottom' || pos === 'top') {
            // Arrow is horizontal — compute left offset within tooltip
            arrowOffset = targetCenterX - left;
            // Clamp so arrow stays within tooltip bounds (12px from edges)
            arrowOffset = Math.max(16, Math.min(arrowOffset, tipRect.width - 16));
        } else {
            // Arrow is vertical — compute top offset within tooltip
            arrowOffset = targetCenterY - top;
            arrowOffset = Math.max(16, Math.min(arrowOffset, tipRect.height - 16));
        }

        tip.style.setProperty('--arrow-offset', arrowOffset + 'px');
    },
};

// Auto-init when DOM is ready (called from App.init, but also safe as standalone)
if (typeof App === 'undefined') {
    document.addEventListener('DOMContentLoaded', () => Tutorial.init());
}
