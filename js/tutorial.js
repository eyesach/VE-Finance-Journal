/**
 * Interactive Tutorial / Guided Walkthrough System
 * Spotlight + tooltip walkthrough for all app features
 */

const Tutorial = {
    // ─── State ───────────────────────────────────────────────
    state: {
        active: false,
        currentTourId: null,
        currentStepIndex: 0,
        steps: [],
    },

    // DOM element references (created in init)
    overlayEl: null,
    spotlightEl: null,
    tooltipEl: null,
    pickerEl: null,

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

    // ─── Initialization ──────────────────────────────────────
    init() {
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

        // Tour picker popover
        this.pickerEl = document.createElement('div');
        this.pickerEl.id = 'tutorial-picker';
        this.pickerEl.className = 'tutorial-picker';
        this.pickerEl.style.display = 'none';

        let pickerHTML = '<div class="tutorial-picker-header">Tutorials</div>';
        for (const [id, tour] of Object.entries(this.tours)) {
            pickerHTML += `
                <button class="tutorial-picker-item" data-tour="${id}">
                    <span class="tutorial-picker-item-label">${tour.label}</span>
                    <span class="tutorial-picker-item-desc">${tour.description}</span>
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

        // Overlay click = next
        this.overlayEl.addEventListener('click', () => this.next());

        // Keyboard
        document.addEventListener('keydown', (e) => {
            if (!this.state.active) return;
            if (e.key === 'Escape') { this.end(); e.stopPropagation(); }
            else if (e.key === 'ArrowRight' || e.key === 'Enter') { this.next(); e.stopPropagation(); }
            else if (e.key === 'ArrowLeft') { this.prev(); e.stopPropagation(); }
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
            <div class="tutorial-welcome-text">Want a quick tour of the features?</div>
            <div class="tutorial-welcome-actions">
                <button class="btn btn-secondary btn-small tutorial-welcome-skip">Skip</button>
                <button class="btn btn-primary btn-small tutorial-welcome-start">Start Tour</button>
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
        if (tourId === 'all') {
            // Flatten all tours into one sequence
            this.state.steps = [];
            for (const tour of Object.values(this.tours)) {
                this.state.steps.push(...tour.steps);
            }
        } else {
            const tour = this.tours[tourId];
            if (!tour) return;
            this.state.steps = [...tour.steps];
        }

        this.state.currentTourId = tourId;
        this.state.currentStepIndex = 0;
        this.state.active = true;
        localStorage.setItem('tutorialSeen', 'true');

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

        // Let DOM settle
        await new Promise(r => setTimeout(r, 80));

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

        // Remove previous highlight
        document.querySelector('.tutorial-target-highlight')?.classList.remove('tutorial-target-highlight');

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
    },

    next() {
        if (!this.state.active) return;
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

        this.state.active = false;
        this.overlayEl.classList.remove('active');
        this.spotlightEl.style.display = 'none';
        this.tooltipEl.style.display = 'none';

        // Remove highlight
        document.querySelector('.tutorial-target-highlight')?.classList.remove('tutorial-target-highlight');
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
        this.tooltipEl.querySelector('.tutorial-tooltip-body').textContent = step.text;
        this.tooltipEl.querySelector('.tutorial-tooltip-step-text').textContent = `Step ${index + 1} of ${total}`;
        this.tooltipEl.querySelector('.tutorial-tooltip-progress-fill').style.width = ((index + 1) / total * 100) + '%';

        // Back button state
        const prevBtn = this.tooltipEl.querySelector('.tutorial-btn-prev');
        prevBtn.style.display = index === 0 ? 'none' : '';

        // Next button label
        const nextBtn = this.tooltipEl.querySelector('.tutorial-btn-next');
        nextBtn.textContent = index === total - 1 ? 'Done' : 'Next';
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
