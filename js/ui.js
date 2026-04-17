/**
 * UI rendering module for the Accounting Journal Calculator
 */

const UI = {
    /**
     * Update summary cards with calculated values and late payment shading
     * @param {Object} summary - Summary object with cashBalance, receivables, payables
     */
    updateSummary(summary) {
        document.getElementById('cashBalance').textContent =
            Utils.formatCurrency(summary.cashBalance);
        document.getElementById('accountsReceivable').textContent =
            Utils.formatCurrency(summary.accountsReceivable);
        document.getElementById('accountsPayable').textContent =
            Utils.formatCurrency(summary.accountsPayable);

        // Check for late payments and apply shading
        const lateInfo = Database.checkLatePayments();
        const cashCard = document.getElementById('cashBalanceCard');
        const receivablesCard = document.getElementById('receivablesCard');
        const payablesCard = document.getElementById('payablesCard');
        const lateIndicator = document.getElementById('cashBalanceLate');

        // Reset
        cashCard.classList.remove('has-late');
        receivablesCard.classList.remove('has-late');
        payablesCard.classList.remove('has-late');
        lateIndicator.style.display = 'none';

        if (lateInfo.hasLateReceivables || lateInfo.hasLatePayables) {
            cashCard.classList.add('has-late');
            lateIndicator.style.display = 'block';

            const parts = [];
            if (lateInfo.hasLateReceivables) {
                parts.push(`${Utils.formatCurrency(lateInfo.lateReceivedAmount)} received late`);
            }
            if (lateInfo.hasLatePayables) {
                parts.push(`${Utils.formatCurrency(lateInfo.latePaidAmount)} paid late`);
            }
            lateIndicator.textContent = parts.join(', ');
        }

        if (lateInfo.hasLateReceivables) {
            receivablesCard.classList.add('has-late');
        }
        if (lateInfo.hasLatePayables) {
            payablesCard.classList.add('has-late');
        }
    },

    /**
     * Populate all year dropdown selects
     * @param {Object} [timeline] - Optional {start, end} to constrain years
     */
    populateYearDropdowns(timeline) {
        const years = (timeline && (timeline.start || timeline.end))
            ? Utils.getYearsInTimeline(timeline.start, timeline.end)
            : Utils.generateYearOptions();
        // Only populate year dropdowns that still exist (balance sheet, etc.)
        const yearSelects = [
            'bsMonthYear'
        ];

        yearSelects.forEach(selectId => {
            const select = document.getElementById(selectId);
            if (!select) return;
            const currentValue = select.value;
            select.innerHTML = '<option value="">Year...</option>';
            years.forEach(year => {
                const option = document.createElement('option');
                option.value = year;
                option.textContent = year;
                select.appendChild(option);
            });
            if (currentValue) select.value = currentValue;
        });
    },

    /**
     * Update the journal title display based on owner name
     * @param {string} owner - Owner/company name
     */
    updateJournalTitle(owner) {
        const suffix = document.getElementById('journalSuffix');
        if (owner && owner.trim()) {
            suffix.textContent = "'s Accounting Journal";
            document.title = `${owner.trim()}'s Accounting Journal`;
        } else {
            suffix.textContent = "Accounting Journal";
            document.title = 'Accounting Journal';
        }
    },

    /**
     * Populate category dropdown with folder grouping (optgroups)
     * @param {Array} categories - Array of category objects (with folder_name, folder_id)
     * @param {string} selectId - ID of the select element
     */
    populateCategoryDropdown(categories, selectId = 'category') {
        const select = document.getElementById(selectId);
        const currentValue = select.value;

        select.innerHTML = '<option value="">Select category...</option>';

        // Group categories by folder
        const folders = {};
        const unfiled = [];

        categories.forEach(cat => {
            // Filter out system-managed Sales Tax categories
            if (cat.is_sales_tax || cat.is_inventory_cost) return;
            if (cat.folder_id && cat.folder_name) {
                if (!folders[cat.folder_name]) {
                    folders[cat.folder_name] = [];
                }
                folders[cat.folder_name].push(cat);
            } else {
                unfiled.push(cat);
            }
        });

        // Add folder optgroups
        const folderNames = Object.keys(folders).sort();
        folderNames.forEach(folderName => {
            const optgroup = document.createElement('optgroup');
            optgroup.label = folderName;
            folders[folderName].forEach(cat => {
                const option = document.createElement('option');
                option.value = cat.id;
                option.textContent = cat.name;
                option.dataset.isMonthly = cat.is_monthly ? '1' : '0';
                option.dataset.isSales = cat.is_sales ? '1' : '0';
                option.dataset.defaultAmount = cat.default_amount || '';
                option.dataset.defaultType = cat.default_type || '';
                option.dataset.defaultStatus = cat.default_status || '';
                optgroup.appendChild(option);
            });
            select.appendChild(optgroup);
        });

        // Add unfiled categories
        unfiled.forEach(cat => {
            const option = document.createElement('option');
            option.value = cat.id;
            option.textContent = cat.name;
            option.dataset.isMonthly = cat.is_monthly ? '1' : '0';
            option.dataset.isSales = cat.is_sales ? '1' : '0';
            option.dataset.defaultAmount = cat.default_amount || '';
            option.dataset.defaultType = cat.default_type || '';
            select.appendChild(option);
        });

        if (currentValue) {
            select.value = currentValue;
        }
    },

    /**
     * Populate payment for month dropdown
     */
    populatePaymentForMonthDropdown() {
        const select = document.getElementById('paymentForMonth');
        const months = Utils.generateMonthOptions();

        select.innerHTML = '<option value="">Select month...</option>';

        months.forEach(({ value, label }) => {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = label;
            select.appendChild(option);
        });
    },

    /**
     * Show/hide payment for month field based on category
     * @param {boolean} show - Whether to show the field
     * @param {string} categoryName - Name of the category for labeling
     */
    togglePaymentForMonth(show, categoryName = '') {
        const group = document.getElementById('paymentForGroup');
        if (show) {
            group.style.display = 'flex';
            const label = group.querySelector('label');
            if (categoryName) {
                label.textContent = `${categoryName} for`;
            } else {
                label.textContent = 'Payment For';
            }
        } else {
            group.style.display = 'none';
            document.getElementById('paymentForMonth').value = '';
        }
    },

    /**
     * Update form field visibility based on status
     * Shows/hides dateProcessed and monthPaid groups
     * @param {string} status - 'pending', 'paid', or 'received'
     */
    updateFormFieldVisibility(status) {
        const dateProcessedGroup = document.getElementById('dateProcessedGroup');
        const monthPaidGroup = document.getElementById('monthPaidGroup');

        if (status === 'pending') {
            dateProcessedGroup.style.display = 'none';
            monthPaidGroup.style.display = 'none';
            // Clear values when switching to pending
            document.getElementById('dateProcessed').value = '';
            document.getElementById('monthPaid').value = '';
        } else {
            dateProcessedGroup.style.display = 'flex';
            monthPaidGroup.style.display = 'flex';
        }
    },

    /**
     * Populate filter category dropdown (with folder optgroups)
     * @param {Array} categories - Array of category objects
     */
    populateFilterCategories(categories) {
        const select = document.getElementById('filterCategory');
        const currentValue = select.value;
        select.innerHTML = '<option value="">All Categories</option>';

        // Group categories by folder
        const folders = {};
        const unfiled = [];

        categories.forEach(cat => {
            if (cat.folder_id && cat.folder_name) {
                if (!folders[cat.folder_name]) {
                    folders[cat.folder_name] = [];
                }
                folders[cat.folder_name].push(cat);
            } else {
                unfiled.push(cat);
            }
        });

        const folderNames = Object.keys(folders).sort();
        folderNames.forEach(folderName => {
            const optgroup = document.createElement('optgroup');
            optgroup.label = folderName;
            folders[folderName].forEach(cat => {
                const option = document.createElement('option');
                option.value = cat.id;
                option.textContent = cat.name;
                optgroup.appendChild(option);
            });
            select.appendChild(optgroup);
        });

        unfiled.forEach(cat => {
            const option = document.createElement('option');
            option.value = cat.id;
            option.textContent = cat.name;
            select.appendChild(option);
        });

        if (currentValue) {
            select.value = currentValue;
        }
    },

    /**
     * Populate filter folder dropdown
     * @param {Array} folders - Array of folder objects
     */
    populateFilterFolders(folders) {
        const select = document.getElementById('filterFolder');
        const currentValue = select.value;
        select.innerHTML = '<option value="">All Folders</option>';

        folders.forEach(folder => {
            const option = document.createElement('option');
            option.value = folder.id;
            option.textContent = folder.name;
            select.appendChild(option);
        });

        // Add "Unfiled" option
        const unfiledOption = document.createElement('option');
        unfiledOption.value = 'unfiled';
        unfiledOption.textContent = 'Unfiled';
        select.appendChild(unfiledOption);

        if (currentValue) {
            select.value = currentValue;
        }
    },

    /**
     * Populate filter month dropdown
     * @param {Array} months - Array of month strings (YYYY-MM)
     */
    populateFilterMonths(months) {
        const select = document.getElementById('filterMonth');
        const savedValue = select.value;
        select.innerHTML = '<option value="">All Months</option>';

        months.forEach(month => {
            const option = document.createElement('option');
            option.value = month;
            option.textContent = Utils.formatMonthDisplay(month);
            select.appendChild(option);
        });

        // Restore previously selected month if it still exists
        if (savedValue) {
            select.value = savedValue;
        }
    },

    /**
     * Render transactions table grouped by the selected sort mode
     * @param {Array} transactions - Array of transaction objects
     * @param {string} sortMode - 'entryDate', 'monthDue', or 'category'
     */
    renderTransactions(transactions, sortMode = 'entryDate') {
        const container = document.getElementById('transactionsContainer');

        if (transactions.length === 0) {
            container.innerHTML = '<p class="empty-state">No transactions yet. Add your first entry above.</p>';
            return;
        }

        let grouped;
        let formatHeader;

        switch (sortMode) {
            case 'monthDue':
                grouped = Utils.groupByMonthDue(transactions);
                formatHeader = (key) => key === 'No Due Date' ? 'No Due Date' : Utils.formatMonthDisplay(key);
                break;
            case 'category':
                grouped = Utils.groupByCategory(transactions);
                formatHeader = (key) => key;
                break;
            case 'entryDate':
            default:
                grouped = Utils.groupByMonth(transactions);
                formatHeader = (key) => Utils.formatMonthDisplay(key);
                break;
        }

        let html = '';

        for (const [key, groupTransactions] of Object.entries(grouped)) {
            html += `
                <div class="month-group">
                    <div class="month-header">${formatHeader(key)}</div>
                    <table class="transaction-table${App.bulkSelectMode ? ' bulk-select-active' : ''}">
                        <thead>
                            <tr>
                                ${App.bulkSelectMode ? '<th class="bulk-checkbox-col"><input type="checkbox" class="bulk-select-all"></th>' : ''}
                                <th>Date</th>
                                <th>Category</th>
                                <th>Type</th>
                                <th>Amount</th>
                                <th>Due</th>
                                <th>Status</th>
                                <th>Processed</th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            ${groupTransactions.map(t => this.renderTransactionRow(t)).join('')}
                        </tbody>
                    </table>
                </div>
            `;
        }

        container.innerHTML = html;
    },

    /**
     * Render a single transaction row
     * @param {Object} t - Transaction object
     * @returns {string} HTML string
     */
    renderTransactionRow(t) {
        // Report month rollback: treat payments after report month as pending
        let effectiveStatus = t.status;
        if (App._reportMonth && t.month_paid && t.month_paid > App._reportMonth) {
            effectiveStatus = 'pending';
        }

        const isOverdue = Utils.isOverdue(t.month_due, effectiveStatus);
        const statusClass = isOverdue ? 'status-overdue' : `status-${effectiveStatus}`;
        const amountClass = t.transaction_type === 'receivable' ? 'amount-receivable' : 'amount-payable';
        const typeClass = `type-${t.transaction_type}`;

        // Status dropdown options based on transaction type
        const statusOptions = t.transaction_type === 'receivable'
            ? ['pending', 'received']
            : ['pending', 'paid'];

        const statusDropdown = `
            <select class="status-select ${statusClass}" data-id="${t.id}">
                ${statusOptions.map(s => `
                    <option value="${s}" ${effectiveStatus === s ? 'selected' : ''}>
                        ${this.capitalizeFirst(s)}
                    </option>
                `).join('')}
            </select>
        `;

        // Late payment info
        const isPaidLate = (effectiveStatus !== 'pending') && Utils.isPaidLate(t.month_due, t.month_paid);
        const lateInfo = isPaidLate
            ? `<span class="late-info">in ${Utils.formatMonthShort(t.month_paid)}</span>`
            : '';

        // Category name with item description and payment for month if applicable
        let categoryDisplay = Utils.escapeHtml(t.category_name || 'Unknown');
        if (t.item_description && t.item_description !== t.category_name) {
            categoryDisplay += ` <span class="item-description-label">${Utils.escapeHtml(t.item_description)}</span>`;
        }
        if (t.payment_for_month) {
            categoryDisplay += `<span class="payment-for-label"> for ${Utils.formatMonthShort(t.payment_for_month)}</span>`;
        }
        // Sale date range display
        if (t.sale_date_start && !t.item_description) {
            categoryDisplay += ` <span class="sale-date-range">${Utils.formatSaleDateRange(t.sale_date_start, t.sale_date_end)}</span>`;
        }

        // Notes indicator icon (shown only when notes exist)
        const notesIcon = t.notes ? `
            <span class="notes-indicator" data-notes="${Utils.escapeHtml(t.notes)}">
                <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                    <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path>
                </svg>
            </span>
        ` : '';

        // Month due display
        const monthDueDisplay = t.month_due
            ? `<span class="month-due-badge ${isOverdue ? 'overdue' : ''}">${Utils.formatMonthShort(t.month_due)}</span>`
            : '-';

        // Processed date display
        const processedDisplay = t.date_processed ? Utils.formatDateShort(t.date_processed) : '-';

        // Row class for late payment shade
        const rowClass = isPaidLate ? 'late-payment-row' : '';

        // Bulk select checkbox
        let bulkCheckboxCell = '';
        if (App.bulkSelectMode) {
            const isAutoGenerated = t.source_type === 'sales_tax' || t.source_type === 'inventory_cost';
            const isEligible = !isAutoGenerated && (
                App.bulkSelectDirection === 'to-paid'
                    ? t.status === 'pending'
                    : (t.status === 'paid' || t.status === 'received')
            );
            if (isEligible) {
                const isChecked = App.bulkSelectedIds.has(t.id) ? 'checked' : '';
                bulkCheckboxCell = `<td class="bulk-checkbox-col"><input type="checkbox" class="bulk-checkbox" data-id="${t.id}" ${isChecked}></td>`;
            } else {
                bulkCheckboxCell = '<td class="bulk-checkbox-col"></td>';
            }
        }

        return `
            <tr data-id="${t.id}" class="${rowClass}">
                ${bulkCheckboxCell}
                <td>${Utils.formatDateShort(t.entry_date)}</td>
                <td>${categoryDisplay} ${notesIcon}</td>
                <td>
                    <span class="type-badge ${typeClass}">
                        ${this.capitalizeFirst(t.transaction_type)}
                    </span>
                </td>
                <td class="${amountClass} txn-amount" data-txn-id="${t.id}">${Utils.formatCurrency(t.amount)}</td>
                <td>${monthDueDisplay}</td>
                <td>
                    ${statusDropdown}
                    ${lateInfo}
                </td>
                <td>${processedDisplay}</td>
                <td class="actions-cell">
                    ${(t.source_type === 'sales_tax' || t.source_type === 'inventory_cost' || t.source_type === 'budget') ? '' : `
                    <button class="btn-icon duplicate-btn" data-id="${t.id}" title="Duplicate">
                        <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                            <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
                            <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
                        </svg>
                    </button>
                    <button class="btn-icon edit-btn" data-id="${t.id}" title="Edit">
                        <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                            <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
                            <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
                        </svg>
                    </button>
                    `}
                    <button class="btn-icon delete-btn" data-id="${t.id}" title="Delete">
                        <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                            <polyline points="3 6 5 6 21 6"></polyline>
                            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
                            <line x1="10" y1="11" x2="10" y2="17"></line>
                            <line x1="14" y1="11" x2="14" y2="17"></line>
                        </svg>
                    </button>
                </td>
            </tr>
        `;
    },

    /**
     * Render a single category item HTML
     * @param {Object} cat - Category object
     * @returns {string} HTML string
     */
    renderCategoryItem(cat) {
        // Skip system-managed Sales Tax categories
        if (cat.is_sales_tax || cat.is_inventory_cost) return '';
        const usageCount = Database.getCategoryUsageCount(cat.id);
        const monthlyBadge = cat.is_monthly
            ? '<span class="category-badge monthly">Monthly</span>'
            : '';
        const defaultTypeBadge = cat.default_type
            ? `<span class="category-badge default-type">${this.capitalizeFirst(cat.default_type)}</span>`
            : '';
        const defaultAmountBadge = cat.default_amount
            ? `<span class="category-badge default-amount">${Utils.formatCurrency(cat.default_amount)}</span>`
            : '';
        const plBadge = cat.show_on_pl
            ? '<span class="category-badge pl">Hidden</span>'
            : '';
        const cogsBadge = cat.is_cogs
            ? '<span class="category-badge cogs">COGS</span>'
            : '';
        const deprBadge = cat.is_depreciation
            ? '<span class="category-badge depr">Depr.</span>'
            : '';
        const salesTaxBadge = cat.is_sales_tax
            ? '<span class="category-badge sales-tax">Sales Tax</span>'
            : '';
        const salesBadge = cat.is_sales
            ? '<span class="category-badge sales">Sales</span>'
            : '';
        const statusBadge = cat.default_status
            ? `<span class="category-badge default-status">${this.capitalizeFirst(cat.default_status)}</span>`
            : '';

        return `
            <div class="category-item" data-id="${cat.id}">
                <div class="category-info">
                    <span class="category-name">${Utils.escapeHtml(cat.name)} ${monthlyBadge} ${defaultTypeBadge} ${defaultAmountBadge} ${statusBadge} ${plBadge} ${cogsBadge} ${deprBadge} ${salesTaxBadge} ${salesBadge}</span>
                    <span class="category-meta">${usageCount} transaction${usageCount !== 1 ? 's' : ''}</span>
                </div>
                <div class="category-actions">
                    <button class="btn-icon always-visible edit-category-btn" data-id="${cat.id}" title="Edit">
                        <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                            <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
                            <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
                        </svg>
                    </button>
                    <button class="btn-icon always-visible delete-category-btn" data-id="${cat.id}" title="Delete"
                            ${usageCount > 0 ? 'disabled style="opacity:0.3;cursor:not-allowed;"' : ''}>
                        <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                            <polyline points="3 6 5 6 21 6"></polyline>
                            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
                            <line x1="10" y1="11" x2="10" y2="17"></line>
                            <line x1="14" y1="11" x2="14" y2="17"></line>
                        </svg>
                    </button>
                </div>
            </div>
        `;
    },

    /**
     * Render the manage categories list in the modal with folder structure
     * @param {Array} categories - Array of category objects
     */
    renderManageCategoriesList(categories) {
        const container = document.getElementById('categoriesList');
        const folders = Database.getFolders();

        if (categories.length === 0 && folders.length === 0) {
            container.innerHTML = '<p class="empty-state">No categories yet.</p>';
            return;
        }

        // Group categories by folder
        const folderMap = {};
        const unfiled = [];

        categories.forEach(cat => {
            if (cat.folder_id) {
                if (!folderMap[cat.folder_id]) folderMap[cat.folder_id] = [];
                folderMap[cat.folder_id].push(cat);
            } else {
                unfiled.push(cat);
            }
        });

        let html = '';

        // Render folders
        folders.forEach(folder => {
            const folderCats = folderMap[folder.id] || [];
            html += `
                <div class="folder-group" data-folder-id="${folder.id}">
                    <div class="folder-header" data-folder-id="${folder.id}">
                        <div class="folder-header-left">
                            <svg class="folder-toggle" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                                <path d="M6 9l6 6 6-6"></path>
                            </svg>
                            <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                                <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path>
                            </svg>
                            ${Utils.escapeHtml(folder.name)} <span class="type-badge type-${folder.folder_type || 'payable'}">${this.capitalizeFirst(folder.folder_type || 'payable')}</span> <span class="category-meta">(${folderCats.length})</span>
                        </div>
                        <div class="folder-actions">
                            <button class="btn-icon always-visible edit-folder-btn" data-id="${folder.id}" title="Edit Folder">
                                <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                                    <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
                                    <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
                                </svg>
                            </button>
                            <button class="btn-icon always-visible delete-folder-btn" data-id="${folder.id}" title="Delete Folder">
                                <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                                    <polyline points="3 6 5 6 21 6"></polyline>
                                    <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
                                </svg>
                            </button>
                        </div>
                    </div>
                    <div class="folder-children" data-folder-id="${folder.id}">
                        ${folderCats.map(cat => this.renderCategoryItem(cat)).join('')}
                        ${folderCats.length === 0 ? '<p class="empty-state" style="padding:8px 12px;font-size:0.8rem;">No categories in this folder</p>' : ''}
                    </div>
                </div>
            `;
        });

        // Render unfiled categories
        if (unfiled.length > 0) {
            if (folders.length > 0) {
                html += '<div class="unfiled-header">Unfiled</div>';
            }
            unfiled.forEach(cat => {
                html += this.renderCategoryItem(cat);
            });
        }

        container.innerHTML = html;
    },

    /**
     * Populate the folder dropdown in category modal
     * @param {Array} folders - Array of folder objects
     * @param {string} selectId - ID of the select element
     */
    populateFolderDropdown(folders, selectId = 'categoryFolder') {
        const select = document.getElementById(selectId);
        const currentValue = select.value;

        select.innerHTML = '<option value="">No Folder</option>';

        folders.forEach(folder => {
            const option = document.createElement('option');
            option.value = folder.id;
            option.textContent = `${folder.name} (${this.capitalizeFirst(folder.folder_type || 'payable')})`;
            option.dataset.folderType = folder.folder_type || 'payable';
            select.appendChild(option);
        });

        if (currentValue) {
            select.value = currentValue;
        }
    },

    /**
     * Render the cash flow spreadsheet table (categories as rows, months as columns)
     * @param {Object} spreadsheetData - { months: string[], data: Object[] } from getCashFlowSpreadsheet()
     * @param {Object} [cfOverrides] - Map of "categoryId-month" => override_amount
     * @param {string} [currentMonth] - Current month YYYY-MM (for projection styling)
     */
    renderCashFlowSpreadsheet(spreadsheetData, cfOverrides, currentMonth, projectedSales) {
        const container = document.getElementById('cashflowSpreadsheet');
        let { months, data } = spreadsheetData;
        cfOverrides = cfOverrides || {};

        // Merge projected sales months into timeline (fills gaps like missing current month)
        if (projectedSales && projectedSales.enabled && projectedSales.byMonth) {
            const allMonths = new Set(months);
            Object.keys(projectedSales.byMonth).forEach(m => allMonths.add(m));
            months = Array.from(allMonths).sort();
        }

        if (months.length === 0) {
            container.innerHTML = '<p class="empty-state">No completed transactions yet.</p>';
            return;
        }

        // Group data by category_id and type using Map (preserves insertion order from DB)
        const receivableCatMap = new Map();
        const payableCatMap = new Map();

        data.forEach(row => {
            const target = row.transaction_type === 'receivable' ? receivableCatMap : payableCatMap;
            if (!target.has(row.category_id)) {
                target.set(row.category_id, { name: row.category_name, is_b2b: row.is_b2b, is_cogs: row.is_cogs, months: {} });
            }
            const catData = target.get(row.category_id);
            catData.months[row.month] = (catData.months[row.month] || 0) + row.total;
        });

        // Entries ordered by DB sort (cashflow_sort_order ASC, name ASC) - Map preserves insertion order
        const receivableEntries = Array.from(receivableCatMap.entries());
        const payableEntries = Array.from(payableCatMap.entries());

        // Helpers (defined before totals so projections are reflected in subtotals)
        const fmtMonth = (m) => Utils.formatMonthShort(m);
        const fmtAmt = (amt) => Utils.formatCurrency(amt);
        const effectiveCurrentMonth = (projectedSales && projectedSales.asOfMonth) || currentMonth;
        const isFuture = (m) => effectiveCurrentMonth && m > effectiveCurrentMonth;

        const getCFVal = (catId, month, computed) => {
            const key = `${catId}-${month}`;
            return (key in cfOverrides) ? cfOverrides[key] : computed;
        };
        const isCFOverridden = (catId, month) => `${catId}-${month}` in cfOverrides;

        const computeCFProjectedAvg = (catMonths) => {
            if (!effectiveCurrentMonth) return 0;
            const pastValues = months.filter(m => !isFuture(m)).map(m => catMonths[m] || 0).filter(v => v > 0);
            return pastValues.length > 0 ? pastValues.reduce((a, b) => a + b, 0) / pastValues.length : 0;
        };

        // Effective value for a category in a month (with projection + override)
        const getEffectiveVal = (catId, catData, m) => {
            const raw = catData.months[m] || 0;
            const fallback = isFuture(m) && raw === 0 ? computeCFProjectedAvg(catData.months) : raw;
            return getCFVal(catId, m, fallback);
        };

        // Projected sales integration for cashflow
        const psActive = projectedSales && projectedSales.enabled && projectedSales.projectionStartMonth;
        const isProjectedSalesMonth = (m) => psActive && m >= projectedSales.projectionStartMonth;
        const useProjectedView = (m) => {
            if (!isProjectedSalesMonth(m)) return false;
            if (isFuture(m)) return true;
            return projectedSales.viewMode === 'projected';
        };

        // Calculate per-month totals (using effective projected/override values)
        const monthReceipts = {};
        const monthPayments = {};
        months.forEach(m => {
            monthReceipts[m] = 0;
            monthPayments[m] = 0;
        });

        receivableEntries.forEach(([catId, catData]) => {
            months.forEach(m => {
                // Skip non-B2B categories for projected-view months (projected sales replaces them)
                if (psActive && !catData.is_b2b && useProjectedView(m)) return;
                monthReceipts[m] += getEffectiveVal(catId, catData, m);
            });
        });

        // Add projected sales revenue + collected sales tax for projected months
        if (psActive) {
            months.forEach(m => {
                if (useProjectedView(m) && projectedSales.byMonth[m]) {
                    monthReceipts[m] += projectedSales.byMonth[m].revenue + projectedSales.byMonth[m].salesTax;
                }
            });
        }

        payableEntries.forEach(([catId, catData]) => {
            months.forEach(m => {
                monthPayments[m] += getEffectiveVal(catId, catData, m);
            });
        });

        // Add projected sales COGS + sales tax remittance for projected months
        if (psActive) {
            months.forEach(m => {
                if (useProjectedView(m) && projectedSales.byMonth[m]) {
                    monthPayments[m] += projectedSales.byMonth[m].cogs + projectedSales.byMonth[m].salesTax;
                }
            });
        }

        // Calculate running beginning balance (ending balance of previous month)
        const beginningBalance = {};
        let runningBalance = 0;
        months.forEach(m => {
            beginningBalance[m] = runningBalance;
            runningBalance += monthReceipts[m] - monthPayments[m];
        });

        let html = '<table class="cashflow-table"><thead><tr>';
        html += '<th></th>';
        months.forEach(m => {
            const futureClass = isFuture(m) ? ' cashflow-future-header' : '';
            const badge = isFuture(m) ? ' <span class="projected-badge">P</span>' : '';
            html += `<th class="${futureClass}">${fmtMonth(m)}${badge}</th>`;
        });
        html += '<th>Total</th>';
        html += '</tr></thead><tbody>';

        // Beginning Cash Balance row
        html += '<tr class="cashflow-subtotal"><td>Beginning Cash Balance</td>';
        months.forEach(m => { html += `<td>${fmtAmt(beginningBalance[m])}</td>`; });
        html += `<td></td></tr>`;

        // CASH RECEIPTS section header
        html += '<tr class="cashflow-section-header"><td colspan="' + (months.length + 2) + '">Cash Receipts</td></tr>';

        // Projected sales: single "Total Sales (Projected)" row = revenue + sales tax collected
        if (psActive) {
            let rowTotal = 0;
            html += `<tr class="ps-projected-row"><td class="cashflow-indent">Total Sales (Projected)</td>`;
            months.forEach(m => {
                if (useProjectedView(m) && projectedSales.byMonth[m]) {
                    const val = projectedSales.byMonth[m].revenue + projectedSales.byMonth[m].salesTax;
                    rowTotal += val;
                    html += `<td class="amount-receivable cashflow-projected">${val ? fmtAmt(val) : ''}</td>`;
                } else {
                    html += '<td></td>';
                }
            });
            html += `<td class="amount-receivable">${fmtAmt(rowTotal)}</td></tr>`;
        }

        // Individual receivable category rows
        receivableEntries.forEach(([catId, catData]) => {
            let rowTotal = 0;
            html += `<tr draggable="true" data-category-id="${catId}" data-section="receivable">`;
            html += '<td class="cashflow-indent cashflow-drag-handle">' + Utils.escapeHtml(catData.name) + '</td>';
            months.forEach(m => {
                if (psActive && !catData.is_b2b && useProjectedView(m)) {
                    html += '<td class="ps-replaced">-</td>';
                } else {
                    const amt = getEffectiveVal(catId, catData, m);
                    rowTotal += amt;
                    const overClass = isCFOverridden(catId, m) ? ' pnl-overridden' : '';
                    const projClass = isFuture(m) && !isCFOverridden(catId, m) ? ' cashflow-projected' : '';
                    const editClass = isFuture(m) ? ' cf-editable' : '';
                    html += `<td class="amount-receivable cf-calc-cell${overClass}${projClass}${editClass}" data-cat-id="${catId}" data-month="${m}">${amt ? fmtAmt(amt) : ''}</td>`;
                }
            });
            html += `<td class="amount-receivable cf-calc-cell" data-cat-id="${catId}" data-month="total">${fmtAmt(rowTotal)}</td></tr>`;
        });

        if (receivableEntries.length === 0 && !psActive) {
            html += '<tr><td class="cashflow-indent" style="color:var(--color-text-muted);font-style:italic;">No receipts</td>';
            months.forEach(() => { html += '<td></td>'; });
            html += '<td></td></tr>';
        }

        // Total Cash Receipts subtotal
        let totalAllReceipts = 0;
        html += '<tr class="cashflow-subtotal"><td>Total Cash Receipts</td>';
        months.forEach(m => {
            totalAllReceipts += monthReceipts[m];
            html += `<td class="amount-receivable cf-calc-cell" data-cf-label="Total Receipts" data-month="${m}">${fmtAmt(monthReceipts[m])}</td>`;
        });
        html += `<td class="amount-receivable">${fmtAmt(totalAllReceipts)}</td></tr>`;

        // Beginning Balance + Receipts subtotal
        html += '<tr class="cashflow-subtotal"><td>Beginning Balance + Receipts</td>';
        months.forEach(m => {
            html += `<td>${fmtAmt(beginningBalance[m] + monthReceipts[m])}</td>`;
        });
        html += '<td></td></tr>';

        // CASH PAYMENTS section header
        html += '<tr class="cashflow-section-header"><td colspan="' + (months.length + 2) + '">Cash Payments</td></tr>';

        // Projected sales synthetic COGS payment rows
        if (psActive) {
            ['online', 'tradeshow'].forEach(key => {
                const ch = projectedSales.channels && projectedSales.channels[key];
                if (!ch || !ch.enabled) return;
                const label = key === 'online' ? 'Online COGS (Projected)' : 'Tradeshow COGS (Projected)';
                let rowTotal = 0;
                html += `<tr class="ps-projected-row"><td class="cashflow-indent">${label}</td>`;
                months.forEach(m => {
                    if (useProjectedView(m)) {
                        const val = (projectedSales.byMonth[m] && projectedSales.byMonth[m][key + 'Cogs']) || 0;
                        rowTotal += val;
                        html += `<td class="amount-payable cashflow-projected">${val ? fmtAmt(val) : ''}</td>`;
                    } else {
                        html += '<td></td>';
                    }
                });
                html += `<td class="amount-payable">${fmtAmt(rowTotal)}</td></tr>`;
            });

            // Sales Tax (Projected) payment row
            let taxRowTotal = 0;
            html += `<tr class="ps-projected-row"><td class="cashflow-indent">Sales Tax (Projected)</td>`;
            months.forEach(m => {
                if (useProjectedView(m) && projectedSales.byMonth[m]) {
                    const val = projectedSales.byMonth[m].salesTax || 0;
                    taxRowTotal += val;
                    html += `<td class="amount-payable cashflow-projected">${val ? fmtAmt(val) : ''}</td>`;
                } else {
                    html += '<td></td>';
                }
            });
            html += `<td class="amount-payable">${fmtAmt(taxRowTotal)}</td></tr>`;
        }

        // Individual payable category rows — all remain visible and editable
        payableEntries.forEach(([catId, catData]) => {
            let rowTotal = 0;
            html += `<tr draggable="true" data-category-id="${catId}" data-section="payable">`;
            html += '<td class="cashflow-indent cashflow-drag-handle">' + Utils.escapeHtml(catData.name) + '</td>';
            months.forEach(m => {
                const amt = getEffectiveVal(catId, catData, m);
                rowTotal += amt;
                const overClass = isCFOverridden(catId, m) ? ' pnl-overridden' : '';
                const projClass = isFuture(m) && !isCFOverridden(catId, m) ? ' cashflow-projected' : '';
                const editClass = isFuture(m) ? ' cf-editable' : '';
                html += `<td class="amount-payable cf-calc-cell${overClass}${projClass}${editClass}" data-cat-id="${catId}" data-month="${m}">${amt ? fmtAmt(amt) : ''}</td>`;
            });
            html += `<td class="amount-payable cf-calc-cell" data-cat-id="${catId}" data-month="total">${fmtAmt(rowTotal)}</td></tr>`;
        });

        if (payableEntries.length === 0 && !psActive) {
            html += '<tr><td class="cashflow-indent" style="color:var(--color-text-muted);font-style:italic;">No payments</td>';
            months.forEach(() => { html += '<td></td>'; });
            html += '<td></td></tr>';
        }

        // Total Cash Payments subtotal
        let totalAllPayments = 0;
        html += '<tr class="cashflow-subtotal"><td>Total Cash Payments</td>';
        months.forEach(m => {
            totalAllPayments += monthPayments[m];
            html += `<td class="amount-payable cf-calc-cell" data-cf-label="Total Payments" data-month="${m}">${fmtAmt(monthPayments[m])}</td>`;
        });
        html += `<td class="amount-payable">${fmtAmt(totalAllPayments)}</td></tr>`;

        // Ending Cash Balance total row
        html += '<tr class="cashflow-total"><td>Ending Cash Balance</td>';
        months.forEach(m => {
            const ending = beginningBalance[m] + monthReceipts[m] - monthPayments[m];
            html += `<td>${fmtAmt(ending)}</td>`;
        });
        const netTotal = totalAllReceipts - totalAllPayments;
        html += `<td>${fmtAmt(netTotal)}</td></tr>`;

        // Net Cash Inflow (Outflow) row
        html += '<tr class="cashflow-subtotal"><td>Net Cash Inflow (Outflow)</td>';
        months.forEach(m => {
            const netCash = monthReceipts[m] - monthPayments[m];
            const colorClass = netCash >= 0 ? 'amount-receivable' : 'amount-payable';
            html += `<td class="${colorClass}">${fmtAmt(netCash)}</td>`;
        });
        const totalNetCash = totalAllReceipts - totalAllPayments;
        const totalNetClass = totalNetCash >= 0 ? 'amount-receivable' : 'amount-payable';
        html += `<td class="${totalNetClass}">${fmtAmt(totalNetCash)}</td></tr>`;

        html += '</tbody></table>';
        container.innerHTML = html;

        // ==================== CASHFLOW SUMMARY ====================
        const summaryEl = document.getElementById('cashflowSummary');
        if (summaryEl && months.length > 0) {
            // Find actuals boundary
            const actualMonths = months.filter(m => !isFuture(m));
            const projectedMonths = months.filter(m => isFuture(m));
            const lastActualMonth = actualMonths.length > 0 ? actualMonths[actualMonths.length - 1] : null;

            // Starting cash (first month's beginning balance)
            const startingCash = beginningBalance[months[0]] || 0;

            // Actuals ending cash (ending balance of last actual month)
            const actualsEndingCash = lastActualMonth
                ? (beginningBalance[lastActualMonth] + monthReceipts[lastActualMonth] - monthPayments[lastActualMonth])
                : startingCash;

            // Projected ending cash (ending balance of last month)
            const lastMonth = months[months.length - 1];
            const projectedEndingCash = beginningBalance[lastMonth] + monthReceipts[lastMonth] - monthPayments[lastMonth];

            // Actuals total change
            const actualsChange = actualsEndingCash - startingCash;

            // Projected total change (from actuals end to final end)
            const projectedChange = projectedEndingCash - actualsEndingCash;

            // Format month short name (e.g., "OCT")
            const fmtMonthName = (m) => {
                const [y, mo] = m.split('-');
                return new Date(y, parseInt(mo) - 1, 1).toLocaleDateString('en-US', { month: 'short' }).toUpperCase();
            };

            // Build monthly cards
            let summaryHtml = '<div class="cashflow-summary-months">';
            months.forEach(m => {
                const net = monthReceipts[m] - monthPayments[m];
                const ending = beginningBalance[m] + monthReceipts[m] - monthPayments[m];
                const future = isFuture(m);
                const netClass = net >= 0 ? 'cf-net-positive' : 'cf-net-negative';
                const prefix = net >= 0 ? '+' : '';

                summaryHtml += `<div class="cashflow-summary-card${future ? ' cf-card-projected' : ''}">
                    <div class="cashflow-summary-month-label">${fmtMonthName(m)}</div>
                    <div class="cashflow-summary-type">${future ? 'Projected' : 'Actual'}</div>
                    <div class="cashflow-summary-net ${netClass}">${prefix}${fmtAmt(net)}</div>
                    <div class="cashflow-summary-net-label">NET CASH FLOW</div>
                    <div class="cashflow-summary-balance">${fmtAmt(ending)}</div>
                    <div class="cashflow-summary-balance-label">ENDING BALANCE</div>
                </div>`;
            });
            summaryHtml += '</div>';

            // Timeline bar
            const actualsChangeSign = actualsChange >= 0 ? '+' : '';
            const projectedChangeSign = projectedChange >= 0 ? '+' : '';
            summaryHtml += `<div class="cashflow-summary-timeline">
                <div class="cf-timeline-point">
                    <div class="cf-timeline-label">STARTING CASH</div>
                    <div class="cf-timeline-value">${fmtAmt(startingCash)}</div>
                </div>
                <div class="cf-timeline-arrow">
                    <div class="cf-timeline-change">Actuals: ${actualsChangeSign}${fmtAmt(actualsChange)}</div>
                    <div class="cf-timeline-line cf-timeline-actual"></div>
                </div>`;

            if (lastActualMonth) {
                summaryHtml += `<div class="cf-timeline-point cf-timeline-mid">
                    <div class="cf-timeline-value">${fmtAmt(actualsEndingCash)}</div>
                    <div class="cf-timeline-label">Actuals Ending Cash</div>
                </div>`;
            }

            if (projectedMonths.length > 0) {
                summaryHtml += `<div class="cf-timeline-arrow">
                    <div class="cf-timeline-change">Projected: ${projectedChangeSign}${fmtAmt(projectedChange)}</div>
                    <div class="cf-timeline-line cf-timeline-projected"></div>
                </div>`;
            }

            summaryHtml += `<div class="cf-timeline-point">
                <div class="cf-timeline-label">PROJECTED ENDING CASH</div>
                <div class="cf-timeline-value">${fmtAmt(projectedEndingCash)}</div>
            </div>
            </div>`;

            summaryEl.innerHTML = summaryHtml;
        } else if (summaryEl) {
            summaryEl.innerHTML = '';
        }
    },

    /**
     * Render the Profit & Loss spreadsheet (VE format, accrual-based)
     * @param {Object} plData - { months, revenue, cogs, opex } from getPLSpreadsheet()
     * @param {Object} overrides - Map of "categoryId-month" => override_amount
     * @param {string} taxMode - 'corporate' (21%) or 'passthrough' ($0)
     * @param {string} [currentMonth] - Current month YYYY-MM (for projection styling)
     */
    renderProfitLossSpreadsheet(plData, overrides, taxMode, currentMonth, projectedSales) {
        const container = document.getElementById('pnlSpreadsheet');
        let { months, revenue, cogs, opex, depreciation, assetDeprByMonth, loanInterestByMonth } = plData;

        // Merge projected sales months into timeline (fills gaps like missing current month)
        if (projectedSales && projectedSales.enabled && projectedSales.byMonth) {
            const allMonths = new Set(months);
            Object.keys(projectedSales.byMonth).forEach(m => allMonths.add(m));
            months = Array.from(allMonths).sort();
        }

        if (months.length === 0) {
            container.innerHTML = '<p class="empty-state">No transactions with a month due yet.</p>';
            return;
        }

        const fmtMonth = (m) => Utils.formatMonthShort(m);
        const fmtAmt = (amt) => Utils.formatCurrency(amt);
        const colSpan = months.length + 2;

        // Helper to get value (override or computed)
        const getVal = (catId, month, computed) => {
            const key = `${catId}-${month}`;
            return (key in overrides) ? overrides[key] : computed;
        };

        const isOverridden = (catId, month) => {
            return `${catId}-${month}` in overrides;
        };

        // Color helper for negative values
        const negStyle = (val) => val >= 0 ? '' : ' style="color:var(--color-payable)"';

        // Group data by category_id using Map (preserves insertion order from DB)
        const groupByCategory = (rows) => {
            const map = new Map();
            rows.forEach(row => {
                if (!map.has(row.category_id)) {
                    map.set(row.category_id, { name: row.category_name, is_b2b: row.is_b2b, months: {} });
                }
                map.get(row.category_id).months[row.month] = row.total;
            });
            return Array.from(map.entries());
        };

        const revenueEntries = groupByCategory(revenue);
        const cogsEntries = groupByCategory(cogs);
        const opexEntries = groupByCategory(opex);

        // Helper: check if month is projected (future), respecting as-of override
        const effectiveCurrentMonth = (projectedSales && projectedSales.asOfMonth) || currentMonth;
        const isFuture = (m) => effectiveCurrentMonth && m > effectiveCurrentMonth;

        // Projected sales integration
        const psActive = projectedSales && projectedSales.enabled && projectedSales.projectionStartMonth;
        const isProjectedSalesMonth = (m) => psActive && m >= projectedSales.projectionStartMonth;
        const useProjectedView = (m) => {
            if (!isProjectedSalesMonth(m)) return false;
            if (isFuture(m)) return true;
            return projectedSales.viewMode === 'projected';
        };

        // Compute projected averages per category from past months
        const computeProjectedAvg = (catMonths) => {
            if (!effectiveCurrentMonth) return 0;
            const pastValues = months.filter(m => !isFuture(m)).map(m => catMonths[m] || 0).filter(v => v > 0);
            return pastValues.length > 0 ? pastValues.reduce((a, b) => a + b, 0) / pastValues.length : 0;
        };

        // Start building table
        let html = '<table class="pnl-table"><thead><tr>';
        html += '<th></th>';
        months.forEach(m => {
            const futureClass = isFuture(m) ? ' pnl-future-header' : '';
            const badge = isFuture(m) ? ' <span class="projected-badge">P</span>' : '';
            html += `<th class="${futureClass}">${fmtMonth(m)}${badge}</th>`;
        });
        html += '<th>Total</th>';
        html += '</tr></thead><tbody>';

        // ===== REVENUE =====
        html += `<tr class="pnl-section-header"><td colspan="${colSpan}">Revenue</td></tr>`;

        const monthRevenue = {};
        months.forEach(m => { monthRevenue[m] = 0; });

        // Separate B2B from non-B2B revenue entries
        const b2bRevenueEntries = revenueEntries.filter(([, d]) => d.is_b2b);
        const nonB2bRevenueEntries = revenueEntries.filter(([, d]) => !d.is_b2b);

        // Render projected sales synthetic rows for non-B2B (when active)
        if (psActive) {
            ['online', 'tradeshow'].forEach(key => {
                const ch = projectedSales.channels && projectedSales.channels[key];
                if (!ch || !ch.enabled) return;
                const label = key === 'online' ? 'Online Revenue (Projected)' : 'Tradeshow Revenue (Projected)';
                let rowTotal = 0;
                html += `<tr class="pnl-indent ps-projected-row"><td>${label}</td>`;
                months.forEach(m => {
                    if (useProjectedView(m)) {
                        const val = (projectedSales.byMonth[m] && projectedSales.byMonth[m][key + 'Revenue']) || 0;
                        monthRevenue[m] += val;
                        rowTotal += val;
                        html += `<td class="pnl-projected">${fmtAmt(val)}</td>`;
                    } else {
                        html += '<td></td>';
                    }
                });
                html += `<td>${fmtAmt(rowTotal)}</td></tr>`;
            });
        }

        // Non-B2B actual revenue rows (blanked out for projected-view months)
        nonB2bRevenueEntries.forEach(([catId, catData]) => {
            let rowTotal = 0;
            const projAvg = computeProjectedAvg(catData.months);
            html += `<tr class="pnl-indent" draggable="true" data-category-id="${catId}" data-section="revenue"><td class="pnl-drag-handle">${Utils.escapeHtml(catData.name)}</td>`;
            months.forEach(m => {
                if (psActive && useProjectedView(m)) {
                    // Projected sales covers this month for non-B2B
                    html += '<td class="ps-replaced">-</td>';
                } else {
                    const computed = catData.months[m] || 0;
                    const fallback = isFuture(m) && computed === 0 ? projAvg : computed;
                    const val = getVal(catId, m, fallback);
                    monthRevenue[m] += val;
                    rowTotal += val;
                    const overriddenClass = isOverridden(catId, m) ? ' pnl-overridden' : '';
                    const projClass = isFuture(m) && !isOverridden(catId, m) ? ' pnl-projected' : '';
                    html += `<td class="pnl-editable pnl-calc-cell${overriddenClass}${projClass}" data-cat-id="${catId}" data-month="${m}">${fmtAmt(val)}</td>`;
                }
            });
            html += `<td>${fmtAmt(rowTotal)}</td></tr>`;
        });

        // B2B revenue rows — always render normally (unaffected by projected sales)
        b2bRevenueEntries.forEach(([catId, catData]) => {
            let rowTotal = 0;
            const projAvg = computeProjectedAvg(catData.months);
            html += `<tr class="pnl-indent" draggable="true" data-category-id="${catId}" data-section="revenue"><td class="pnl-drag-handle">${Utils.escapeHtml(catData.name)}</td>`;
            months.forEach(m => {
                const computed = catData.months[m] || 0;
                const fallback = isFuture(m) && computed === 0 ? projAvg : computed;
                const val = getVal(catId, m, fallback);
                monthRevenue[m] += val;
                rowTotal += val;
                const overriddenClass = isOverridden(catId, m) ? ' pnl-overridden' : '';
                const projClass = isFuture(m) && !isOverridden(catId, m) ? ' pnl-projected' : '';
                html += `<td class="pnl-editable pnl-calc-cell${overriddenClass}${projClass}" data-cat-id="${catId}" data-month="${m}">${fmtAmt(val)}</td>`;
            });
            html += `<td>${fmtAmt(rowTotal)}</td></tr>`;
        });

        // Total Revenue
        let totalRevenue = 0;
        html += '<tr class="pnl-subtotal"><td>Total Revenue</td>';
        months.forEach(m => {
            totalRevenue += monthRevenue[m];
            html += `<td class="pnl-calc-cell" data-pnl-label="Total Revenue" data-month="${m}">${fmtAmt(monthRevenue[m])}</td>`;
        });
        html += `<td class="pnl-calc-cell" data-pnl-label="Total Revenue" data-month="total">${fmtAmt(totalRevenue)}</td></tr>`;

        // ===== COST OF GOODS SOLD =====
        html += `<tr class="pnl-section-header"><td colspan="${colSpan}">Cost of Goods Sold</td></tr>`;

        const monthCogs = {};
        months.forEach(m => { monthCogs[m] = 0; });

        // Projected sales synthetic COGS rows (additive — existing COGS rows remain untouched)
        if (psActive) {
            ['online', 'tradeshow'].forEach(key => {
                const ch = projectedSales.channels && projectedSales.channels[key];
                if (!ch || !ch.enabled) return;
                const label = key === 'online' ? 'Online COGS (Projected)' : 'Tradeshow COGS (Projected)';
                let rowTotal = 0;
                html += `<tr class="pnl-indent ps-projected-row"><td>${label}</td>`;
                months.forEach(m => {
                    if (useProjectedView(m)) {
                        const val = (projectedSales.byMonth[m] && projectedSales.byMonth[m][key + 'Cogs']) || 0;
                        monthCogs[m] += val;
                        rowTotal += val;
                        html += `<td class="pnl-projected">${fmtAmt(val)}</td>`;
                    } else {
                        html += '<td></td>';
                    }
                });
                html += `<td>${fmtAmt(rowTotal)}</td></tr>`;
            });
        }

        // ALL existing COGS rows — always visible and editable (not replaced by projections)
        cogsEntries.forEach(([catId, catData]) => {
            let rowTotal = 0;
            const projAvg = computeProjectedAvg(catData.months);
            html += `<tr class="pnl-indent" draggable="true" data-category-id="${catId}" data-section="cogs"><td class="pnl-drag-handle">${Utils.escapeHtml(catData.name)}</td>`;
            months.forEach(m => {
                const computed = catData.months[m] || 0;
                const fallback = isFuture(m) && computed === 0 ? projAvg : computed;
                const val = getVal(catId, m, fallback);
                monthCogs[m] += val;
                rowTotal += val;
                const overriddenClass = isOverridden(catId, m) ? ' pnl-overridden' : '';
                const projClass = isFuture(m) && !isOverridden(catId, m) ? ' pnl-projected' : '';
                html += `<td class="pnl-editable pnl-calc-cell${overriddenClass}${projClass}" data-cat-id="${catId}" data-month="${m}">${fmtAmt(val)}</td>`;
            });
            html += `<td>${fmtAmt(rowTotal)}</td></tr>`;
        });

        // Total COGS
        let totalCogs = 0;
        html += '<tr class="pnl-subtotal"><td>Total Cost of Goods Sold</td>';
        months.forEach(m => {
            totalCogs += monthCogs[m];
            html += `<td>${fmtAmt(monthCogs[m])}</td>`;
        });
        html += `<td>${fmtAmt(totalCogs)}</td></tr>`;

        // ===== GROSS PROFIT =====
        let totalGP = 0;
        html += '<tr class="pnl-total"><td>Gross Profit</td>';
        months.forEach(m => {
            const gp = monthRevenue[m] - monthCogs[m];
            totalGP += gp;
            html += `<td${negStyle(gp)}>${fmtAmt(gp)}</td>`;
        });
        html += `<td${negStyle(totalGP)}>${fmtAmt(totalGP)}</td></tr>`;

        // Gross Margin %
        html += '<tr class="pnl-percentage"><td>Gross Margin %</td>';
        months.forEach(m => {
            const gp = monthRevenue[m] - monthCogs[m];
            const pct = monthRevenue[m] ? ((gp / monthRevenue[m]) * 100).toFixed(1) + '%' : '-';
            html += `<td>${pct}</td>`;
        });
        const totalGPPct = totalRevenue ? ((totalGP / totalRevenue) * 100).toFixed(1) + '%' : '-';
        html += `<td>${totalGPPct}</td></tr>`;

        // ===== OPERATING EXPENSES =====
        html += `<tr class="pnl-section-header"><td colspan="${colSpan}">Operating Expenses</td></tr>`;

        const monthOpex = {};
        months.forEach(m => { monthOpex[m] = 0; });

        opexEntries.forEach(([catId, catData]) => {
            let rowTotal = 0;
            const projAvg = computeProjectedAvg(catData.months);
            html += `<tr class="pnl-indent" draggable="true" data-category-id="${catId}" data-section="opex"><td class="pnl-drag-handle">${Utils.escapeHtml(catData.name)}</td>`;
            months.forEach(m => {
                const computed = catData.months[m] || 0;
                const fallback = isFuture(m) && computed === 0 ? projAvg : computed;
                const val = getVal(catId, m, fallback);
                monthOpex[m] += val;
                rowTotal += val;
                const overriddenClass = isOverridden(catId, m) ? ' pnl-overridden' : '';
                const projClass = isFuture(m) && !isOverridden(catId, m) ? ' pnl-projected' : '';
                html += `<td class="pnl-editable pnl-calc-cell${overriddenClass}${projClass}" data-cat-id="${catId}" data-month="${m}">${fmtAmt(val)}</td>`;
            });
            html += `<td>${fmtAmt(rowTotal)}</td></tr>`;
        });

        // Depreciation rows (manual input only — values come from pl_overrides)
        depreciation.forEach(cat => {
            let rowTotal = 0;
            html += `<tr class="pnl-indent" draggable="true" data-category-id="${cat.category_id}" data-section="opex"><td class="pnl-drag-handle">${Utils.escapeHtml(cat.category_name)}</td>`;
            months.forEach(m => {
                const val = getVal(cat.category_id, m, 0);
                monthOpex[m] += val;
                rowTotal += val;
                const overriddenClass = isOverridden(cat.category_id, m) ? ' pnl-overridden' : '';
                html += `<td class="pnl-editable${overriddenClass}" data-cat-id="${cat.category_id}" data-month="${m}">${fmtAmt(val)}</td>`;
            });
            html += `<td>${fmtAmt(rowTotal)}</td></tr>`;
        });

        // Computed: Depreciation from Fixed Assets tab
        if (assetDeprByMonth && Object.keys(assetDeprByMonth).length > 0) {
            let rowTotal = 0;
            html += '<tr class="pnl-indent pnl-computed-row"><td>Depreciation (Fixed Assets)</td>';
            months.forEach(m => {
                const val = assetDeprByMonth[m] || 0;
                monthOpex[m] += val;
                rowTotal += val;
                html += `<td>${fmtAmt(val)}</td>`;
            });
            html += `<td>${fmtAmt(rowTotal)}</td></tr>`;
        }

        // Computed: Interest Expense from Loans
        if (loanInterestByMonth && Object.keys(loanInterestByMonth).length > 0) {
            let rowTotal = 0;
            html += '<tr class="pnl-indent pnl-computed-row"><td>Interest Expense (Loans)</td>';
            months.forEach(m => {
                const val = loanInterestByMonth[m] || 0;
                monthOpex[m] += val;
                rowTotal += val;
                html += `<td>${fmtAmt(val)}</td>`;
            });
            html += `<td>${fmtAmt(rowTotal)}</td></tr>`;
        }

        // Total Operating Expenses
        let totalOpex = 0;
        html += '<tr class="pnl-subtotal"><td>Total Operating Expenses</td>';
        months.forEach(m => {
            totalOpex += monthOpex[m];
            html += `<td>${fmtAmt(monthOpex[m])}</td>`;
        });
        html += `<td>${fmtAmt(totalOpex)}</td></tr>`;

        // ===== NET INCOME (LOSS) BEFORE TAXES =====
        let totalNIBT = 0;
        html += '<tr class="pnl-total"><td>Net Income (Loss) Before Taxes</td>';
        months.forEach(m => {
            const nibt = monthRevenue[m] - monthCogs[m] - monthOpex[m];
            totalNIBT += nibt;
            html += `<td${negStyle(nibt)}>${fmtAmt(nibt)}</td>`;
        });
        html += `<td${negStyle(totalNIBT)}>${fmtAmt(totalNIBT)}</td></tr>`;

        // ===== INCOME TAX =====
        const monthTax = {};
        if (taxMode === 'corporate') {
            // Corporate: 21% of Net Income Before Taxes (only when positive)
            months.forEach(m => {
                const nibt = monthRevenue[m] - monthCogs[m] - monthOpex[m];
                const autoTax = nibt > 0 ? nibt * 0.21 : 0;
                monthTax[m] = getVal(-1, m, autoTax);
            });
        } else {
            // Pass-through: $0
            months.forEach(m => { monthTax[m] = 0; });
        }

        let totalTax = 0;
        const taxLabel = taxMode === 'passthrough'
            ? 'Income Tax (pass-through to owners)'
            : 'Income Tax (21%)';
        html += `<tr class="pnl-indent"><td>${taxLabel}</td>`;
        months.forEach(m => {
            totalTax += monthTax[m];
            if (taxMode === 'passthrough') {
                html += '<td style="text-align:center;">&mdash;</td>';
            } else {
                const overriddenClass = isOverridden(-1, m) ? ' pnl-overridden' : '';
                html += `<td class="pnl-editable${overriddenClass}" data-cat-id="-1" data-month="${m}">${fmtAmt(monthTax[m])}</td>`;
            }
        });
        if (taxMode === 'passthrough') {
            html += '<td style="text-align:center;">&mdash;</td></tr>';
        } else {
            html += `<td>${fmtAmt(totalTax)}</td></tr>`;
        }

        // ===== NET INCOME (LOSS) AFTER TAXES =====
        let totalNIAT = 0;
        html += '<tr class="pnl-total"><td>Net Income (Loss) After Taxes</td>';
        months.forEach(m => {
            const niat = monthRevenue[m] - monthCogs[m] - monthOpex[m] - monthTax[m];
            totalNIAT += niat;
            html += `<td${negStyle(niat)}>${fmtAmt(niat)}</td>`;
        });
        html += `<td${negStyle(totalNIAT)}>${fmtAmt(totalNIAT)}</td></tr>`;

        // ===== CUMULATIVE NET INCOME =====
        let cumulative = 0;
        html += '<tr class="pnl-cumulative"><td>Cumulative Net Income (Loss)</td>';
        months.forEach(m => {
            const niat = monthRevenue[m] - monthCogs[m] - monthOpex[m] - monthTax[m];
            cumulative += niat;
            html += `<td${negStyle(cumulative)}>${fmtAmt(cumulative)}</td>`;
        });
        html += `<td>${fmtAmt(cumulative)}</td></tr>`;

        html += '</tbody></table>';
        container.innerHTML = html;

        // Expose computed per-month operating expenses for break-even to use
        this._pnlMonthOpex = monthOpex;
        this._pnlMonths = months;
    },

    /**
     * Render the Balance Sheet
     * @param {Object} data - Balance sheet data from App.refreshBalanceSheet()
     */
    renderBalanceSheet(data) {
        const container = document.getElementById('balanceSheetContent');
        const fmtAmt = (amt) => Utils.formatCurrency(amt);
        const monthLabel = Utils.formatMonthDisplay(data.asOfMonth);

        const isProjected = Utils.isFutureMonth(data.asOfMonth);
        const projLabel = isProjected ? ' <span class="bs-projected-label">Projected</span>' : '';
        let html = `<div class="bs-date-label" style="font-size:0.8rem;color:var(--color-text-muted);margin-bottom:12px;">As of ${monthLabel}${projLabel}</div>`;

        html += '<table class="bs-table"><tbody>';

        // ===== ASSETS =====
        html += '<tr class="bs-section-header"><td colspan="2">Assets</td></tr>';

        // Current Assets
        html += '<tr class="bs-subsection"><td colspan="2">Current Assets</td></tr>';
        html += `<tr class="bs-indent"><td>Cash</td><td class="bs-value" data-bs-key="cash">${fmtAmt(data.cash)}</td></tr>`;
        html += `<tr class="bs-indent"><td>Accounts Receivable</td><td class="bs-value" data-bs-key="ar">${fmtAmt(data.ar)}</td></tr>`;
        if (data.arByCategory && data.arByCategory.length > 0) {
            data.arByCategory.forEach(cat => {
                html += `<tr class="bs-detail-indent"><td>${Utils.escapeHtml(cat.category_name)}</td><td>${fmtAmt(cat.total)}</td></tr>`;
            });
        }

        const totalCurrentAssets = data.cash + data.ar;
        html += `<tr class="bs-subtotal"><td>Total Current Assets</td><td class="bs-value" data-bs-key="totalCurrentAssets">${fmtAmt(totalCurrentAssets)}</td></tr>`;

        // Fixed Assets
        html += '<tr class="bs-subsection"><td colspan="2">Fixed Assets</td></tr>';
        if (data.assetDetails.length === 0) {
            html += '<tr class="bs-indent"><td style="color:var(--color-text-muted);font-style:italic;">No fixed assets</td><td></td></tr>';
        } else {
            data.assetDetails.forEach(asset => {
                html += `<tr class="bs-indent"><td>${Utils.escapeHtml(asset.name)}</td><td>${fmtAmt(asset.purchase_cost)}</td></tr>`;
            });
        }

        html += `<tr class="bs-indent"><td>Total Fixed Assets (Cost)</td><td>${fmtAmt(data.totalFixedAssetCost)}</td></tr>`;
        html += `<tr class="bs-indent"><td>Less: Accumulated Depreciation</td><td>(${fmtAmt(data.totalAccumDepr)})</td></tr>`;
        html += `<tr class="bs-subtotal"><td>Net Fixed Assets</td><td class="bs-value" data-bs-key="netFixedAssets">${fmtAmt(data.netFixedAssets)}</td></tr>`;

        // Total Assets
        html += `<tr class="bs-total"><td>Total Assets</td><td class="bs-value" data-bs-key="totalAssets">${fmtAmt(data.totalAssets)}</td></tr>`;

        // Spacer
        html += '<tr><td colspan="2" style="padding:8px;"></td></tr>';

        // ===== LIABILITIES =====
        html += '<tr class="bs-section-header"><td colspan="2">Liabilities</td></tr>';

        // Current Liabilities
        html += '<tr class="bs-subsection"><td colspan="2">Current Liabilities</td></tr>';
        html += `<tr class="bs-indent"><td>Accounts Payable</td><td class="bs-value" data-bs-key="ap">${fmtAmt(data.ap)}</td></tr>`;
        if (data.apByCategory && data.apByCategory.length > 0) {
            data.apByCategory.forEach(cat => {
                html += `<tr class="bs-detail-indent"><td>${Utils.escapeHtml(cat.category_name)}</td><td>${fmtAmt(cat.total)}</td></tr>`;
            });
        }
        html += `<tr class="bs-indent"><td>Sales Tax Payable</td><td class="bs-value" data-bs-key="salesTaxPayable">${fmtAmt(data.salesTaxPayable)}</td></tr>`;

        const totalCurrentLiabilities = data.ap + data.salesTaxPayable;
        html += `<tr class="bs-subtotal"><td>Total Current Liabilities</td><td class="bs-value" data-bs-key="totalCurrentLiabilities">${fmtAmt(totalCurrentLiabilities)}</td></tr>`;

        // Long-Term Liabilities
        if (data.totalLoanBalance > 0 && data.loanDetails && data.loanDetails.length > 0) {
            html += '<tr class="bs-subsection"><td colspan="2">Long-Term Liabilities</td></tr>';
            data.loanDetails.forEach(loan => {
                if (loan.balance > 0) {
                    html += `<tr class="bs-indent"><td>${Utils.escapeHtml(loan.name)}</td><td>${fmtAmt(loan.balance)}</td></tr>`;
                }
            });
            html += `<tr class="bs-indent" style="font-weight:500;"><td>Total Loans Payable</td><td>${fmtAmt(data.totalLoanBalance)}</td></tr>`;
        }

        html += `<tr class="bs-subtotal"><td>Total Liabilities</td><td class="bs-value" data-bs-key="totalLiabilities">${fmtAmt(data.totalLiabilities)}</td></tr>`;

        // Spacer
        html += '<tr><td colspan="2" style="padding:4px;"></td></tr>';

        // ===== STOCKHOLDERS' EQUITY =====
        html += '<tr class="bs-section-header"><td colspan="2">Stockholders\' Equity</td></tr>';
        html += `<tr class="bs-indent"><td>Common Stock</td><td class="bs-value" data-bs-key="commonStock">${fmtAmt(data.commonStock)}</td></tr>`;
        html += `<tr class="bs-indent"><td>Additional Paid-In Capital</td><td class="bs-value" data-bs-key="apic">${fmtAmt(data.apic)}</td></tr>`;
        html += `<tr class="bs-indent"><td>Retained Earnings</td><td class="bs-value" data-bs-key="retainedEarnings">${fmtAmt(data.retainedEarnings)}</td></tr>`;
        html += `<tr class="bs-subtotal"><td>Total Stockholders' Equity</td><td class="bs-value" data-bs-key="totalEquity">${fmtAmt(data.totalEquity)}</td></tr>`;

        // Total Liabilities + Equity
        html += `<tr class="bs-total"><td>Total Liabilities + Equity</td><td class="bs-value" data-bs-key="totalLiabilitiesAndEquity">${fmtAmt(data.totalLiabilitiesAndEquity)}</td></tr>`;

        html += '</tbody></table>';

        // Validation
        if (data.isBalanced) {
            html += '<div class="bs-validation balanced">Balanced &mdash; Assets = Liabilities + Equity</div>';
        } else {
            const diff = data.totalAssets - data.totalLiabilitiesAndEquity;
            html += `<div class="bs-validation unbalanced">Unbalanced &mdash; Difference: ${fmtAmt(Math.abs(diff))}</div>`;
        }

        container.innerHTML = html;
        this.renderFinancialRatios(data, document.getElementById('bsRatiosContent'));
    },

    /**
     * Render the Financial Ratios summary section below the Balance Sheet.
     */
    renderFinancialRatios(data, container) {
        if (!container || !data.asOfMonth) return;

        const pl = data.plTotals || {};
        const totalRevenue = pl.totalRevenue || 0;
        const totalNIAT = pl.totalNIAT || 0;
        const totalNIBT = pl.totalNIBT || 0;
        const totalLoanInterest = pl.totalLoanInterest || 0;
        const totalGP = pl.totalGP || 0;

        const totalEquity = data.totalEquity || 0;
        const totalAssets = data.totalAssets || 0;
        const totalLiabilities = data.totalLiabilities || 0;
        const currentAssets = (data.cash || 0) + (data.ar || 0);
        const currentLiabilities = (data.ap || 0) + (data.salesTaxPayable || 0);

        const fmtPct = (num, den) => {
            if (!den || den === 0) return '<span class="ratio-na">N/A</span>';
            return ((num / den) * 100).toFixed(1) + '%';
        };
        const fmtX = (num, den) => {
            if (!den || den === 0) return '<span class="ratio-na">N/A</span>';
            return (num / den).toFixed(2) + 'x';
        };
        const cls = (num, den, higherIsBetter = true) => {
            if (!den || den === 0) return '';
            const val = num / den;
            if (higherIsBetter) return val >= 0 ? ' ratio-positive' : ' ratio-negative';
            return val <= 1 ? ' ratio-positive' : ' ratio-negative';
        };

        const monthLabel = Utils.formatMonthDisplay(data.asOfMonth);

        let html = `
        <div class="bs-ratios-header">
            <h3>Financial Ratios</h3>
            <span class="bs-ratios-period">P&amp;L cumulative through ${monthLabel}</span>
        </div>
        <div class="bs-ratios-body">`;

        // Profitability
        html += `
        <div class="bs-ratios-group">
            <div class="bs-ratios-group-title">Profitability</div>
            <div class="bs-ratios-grid">
                <div class="bs-ratio-card">
                    <span class="bs-ratio-label">Gross Margin</span>
                    <span class="bs-ratio-value${cls(totalGP, totalRevenue)}">${fmtPct(totalGP, totalRevenue)}</span>
                    <span class="bs-ratio-formula">Gross Profit / Revenue</span>
                </div>
                <div class="bs-ratio-card">
                    <span class="bs-ratio-label">Net Profit Margin</span>
                    <span class="bs-ratio-value${cls(totalNIAT, totalRevenue)}">${fmtPct(totalNIAT, totalRevenue)}</span>
                    <span class="bs-ratio-formula">Net Income (AT) / Revenue</span>
                </div>
                <div class="bs-ratio-card">
                    <span class="bs-ratio-label">Return on Equity</span>
                    <span class="bs-ratio-value${cls(totalNIAT, totalEquity)}">${fmtPct(totalNIAT, totalEquity)}</span>
                    <span class="bs-ratio-formula">Net Income (AT) / Total Equity</span>
                </div>
                <div class="bs-ratio-card">
                    <span class="bs-ratio-label">Return on Assets</span>
                    <span class="bs-ratio-value${cls(totalNIAT, totalAssets)}">${fmtPct(totalNIAT, totalAssets)}</span>
                    <span class="bs-ratio-formula">Net Income (AT) / Total Assets</span>
                </div>
            </div>
        </div>`;

        // Liquidity
        const liqVal = currentLiabilities > 0 ? fmtX(currentAssets, currentLiabilities) : '<span class="ratio-na">N/A</span>';
        const liqCls = currentLiabilities > 0
            ? ((currentAssets / currentLiabilities) >= 1 ? ' ratio-positive' : ' ratio-negative')
            : '';

        const cashRatioVal = currentLiabilities > 0 ? fmtX(data.cash || 0, currentLiabilities) : '<span class="ratio-na">N/A</span>';
        const cashRatioCls = currentLiabilities > 0
            ? (((data.cash || 0) / currentLiabilities) >= 0.5 ? ' ratio-positive' : ' ratio-negative')
            : '';

        html += `
        <div class="bs-ratios-group">
            <div class="bs-ratios-group-title">Liquidity</div>
            <div class="bs-ratios-grid">
                <div class="bs-ratio-card">
                    <span class="bs-ratio-label">Current Ratio</span>
                    <span class="bs-ratio-value${liqCls}">${liqVal}</span>
                    <span class="bs-ratio-formula">Current Assets / Current Liabilities</span>
                </div>
                <div class="bs-ratio-card">
                    <span class="bs-ratio-label">Quick Ratio</span>
                    <span class="bs-ratio-value${liqCls}">${liqVal}</span>
                    <span class="bs-ratio-formula">Same as Current Ratio (no inventory tracked)</span>
                </div>
                <div class="bs-ratio-card">
                    <span class="bs-ratio-label">Cash Ratio</span>
                    <span class="bs-ratio-value${cashRatioCls}">${cashRatioVal}</span>
                    <span class="bs-ratio-formula">Cash / Current Liabilities</span>
                </div>
            </div>
        </div>`;

        // Solvency & Leverage
        const deVal = totalEquity !== 0 ? fmtX(totalLiabilities, totalEquity) : '<span class="ratio-na">N/A</span>';
        const deCls = totalEquity !== 0
            ? ((totalLiabilities / totalEquity) <= 2 ? ' ratio-positive' : ' ratio-negative')
            : '';

        html += `
        <div class="bs-ratios-group">
            <div class="bs-ratios-group-title">Solvency &amp; Leverage</div>
            <div class="bs-ratios-grid">
                <div class="bs-ratio-card">
                    <span class="bs-ratio-label">Debt-to-Equity</span>
                    <span class="bs-ratio-value${deCls}">${deVal}</span>
                    <span class="bs-ratio-formula">Total Liabilities / Total Equity</span>
                </div>
                <div class="bs-ratio-card">
                    <span class="bs-ratio-label">Interest Coverage</span>
                    <span class="bs-ratio-value${totalLoanInterest > 0 ? cls(totalNIBT, totalLoanInterest) : ''}">${totalLoanInterest > 0 ? fmtX(totalNIBT, totalLoanInterest) : '<span class="ratio-na">No debt</span>'}</span>
                    <span class="bs-ratio-formula">Net Income (BT) / Loan Interest</span>
                </div>
            </div>
        </div>`;

        // Efficiency
        html += `
        <div class="bs-ratios-group">
            <div class="bs-ratios-group-title">Efficiency</div>
            <div class="bs-ratios-grid">
                <div class="bs-ratio-card">
                    <span class="bs-ratio-label">Asset Turnover</span>
                    <span class="bs-ratio-value${cls(totalRevenue, totalAssets)}">${fmtX(totalRevenue, totalAssets)}</span>
                    <span class="bs-ratio-formula">Revenue / Total Assets</span>
                </div>
            </div>
        </div>`;

        html += '</div>';
        container.innerHTML = html;
    },

    /**
     * Render the Fixed Assets tab with list/detail layout
     * @param {Array} assets - Array of asset objects
     * @param {number|null} selectedAssetId - Currently selected asset ID
     */
    renderFixedAssetsTab(assets, selectedAssetId) {
        const fmtAmt = (amt) => Utils.formatCurrency(amt);

        // Summary cards
        let totalCost = 0, totalAccumDepr = 0;
        assets.forEach(asset => {
            totalCost += asset.purchase_cost;
            const schedule = Utils.computeDepreciationSchedule(asset);
            const currentMonth = Utils.getCurrentMonth();
            let accumDepr = 0;
            Object.entries(schedule).forEach(([m, amt]) => {
                if (m <= currentMonth) accumDepr += amt;
            });
            asset._accumDepr = accumDepr;
            asset._nbv = asset.purchase_cost - accumDepr;
            totalAccumDepr += accumDepr;
        });

        const summaryContainer = document.getElementById('assetsSummaryCards');
        summaryContainer.innerHTML = `
            <div class="assets-summary-card"><span class="assets-summary-label">Total Cost</span><span class="assets-summary-value">${fmtAmt(totalCost)}</span></div>
            <div class="assets-summary-card"><span class="assets-summary-label">Accum. Depreciation</span><span class="assets-summary-value amount-payable">${fmtAmt(totalAccumDepr)}</span></div>
            <div class="assets-summary-card"><span class="assets-summary-label">Net Book Value</span><span class="assets-summary-value">${fmtAmt(totalCost - totalAccumDepr)}</span></div>
        `;

        // Left panel: asset list
        const listPanel = document.getElementById('assetsListPanel');
        if (assets.length === 0) {
            listPanel.innerHTML = '<p class="empty-state">No fixed assets yet. Click "+ Add Asset" to begin.</p>';
        } else {
            listPanel.innerHTML = assets.map(asset => {
                const selected = asset.id === selectedAssetId ? ' selected' : '';
                const methodLabel = asset.depreciation_method === 'none' ? 'Non-depreciable'
                    : asset.depreciation_method === 'double_declining' ? 'DDB' : 'SL';
                return `<div class="asset-list-item${selected}" data-id="${asset.id}">
                    <div class="asset-list-name">${Utils.escapeHtml(asset.name)}</div>
                    <div class="asset-list-meta">${fmtAmt(asset.purchase_cost)} &middot; ${methodLabel}</div>
                    <div class="asset-list-actions">
                        <button class="btn-icon edit-asset-btn" data-id="${asset.id}" title="Edit">&#9998;</button>
                        <button class="btn-icon delete-asset-btn" data-id="${asset.id}" title="Delete">&times;</button>
                    </div>
                </div>`;
            }).join('');
        }

        // Right panel: selected asset detail
        const detailPanel = document.getElementById('assetsDetailPanel');
        const selectedAsset = assets.find(a => a.id === selectedAssetId);
        if (!selectedAsset) {
            detailPanel.innerHTML = '<p class="empty-state">Select an asset to view its depreciation schedule.</p>';
            return;
        }

        const schedule = Utils.computeDepreciationSchedule(selectedAsset);
        const scheduleEntries = Object.entries(schedule).sort((a, b) => a[0].localeCompare(b[0]));

        let html = `<div class="asset-detail-header">
            <h4>${Utils.escapeHtml(selectedAsset.name)}</h4>
            <div class="asset-detail-meta">
                <span>Cost: ${fmtAmt(selectedAsset.purchase_cost)}</span>
                <span>Salvage: ${fmtAmt(selectedAsset.salvage_value || 0)}</span>
                <span>Life: ${selectedAsset.useful_life_months} mo</span>
                <span>Purchased: ${Utils.formatDate(selectedAsset.purchase_date)}</span>
            </div>
        </div>`;

        if (scheduleEntries.length === 0) {
            html += '<p class="empty-state">This asset is non-depreciable.</p>';
        } else {
            html += '<div class="asset-depr-table-wrapper"><table class="asset-depr-table"><thead><tr>';
            html += '<th>Month</th><th>Depreciation</th><th>Accumulated</th><th>Net Book Value</th>';
            html += '</tr></thead><tbody>';

            let accumDepr = 0;
            scheduleEntries.forEach(([month, depr]) => {
                accumDepr += depr;
                const nbv = selectedAsset.purchase_cost - accumDepr;
                html += `<tr>
                    <td>${Utils.formatMonthShort(month)}</td>
                    <td>${fmtAmt(depr)}</td>
                    <td>${fmtAmt(accumDepr)}</td>
                    <td>${fmtAmt(nbv)}</td>
                </tr>`;
            });

            html += '</tbody></table></div>';
        }

        if (selectedAsset.notes) {
            html += `<div class="asset-detail-notes">Notes: ${Utils.escapeHtml(selectedAsset.notes)}</div>`;
        }

        detailPanel.innerHTML = html;
    },

    /**
     * Render the equity section in the Assets & Equity tab
     * @param {Object} equityConfig - Equity configuration
     */
    renderEquitySection(equityConfig) {
        const fmtAmt = (amt) => Utils.formatCurrency(amt);
        const round2 = (v) => Math.round(v * 100) / 100;

        const commonStock = round2(equityConfig.common_stock_par * equityConfig.common_stock_shares);
        const apicVal = round2(equityConfig.apic || 0);
        const totalEquity = round2(commonStock + apicVal);

        const panel = document.getElementById('equityDisplayPanel');

        if (totalEquity === 0 && !equityConfig.common_stock_shares) {
            panel.innerHTML = '<p class="empty-state">No equity configured. Click "Edit Equity" to set up seed money and APIC.</p>';
            return;
        }

        const seedStatus = this._equityStatusBadge(equityConfig.seed_expected_date, equityConfig.seed_received_date);
        const apicStatus = this._equityStatusBadge(equityConfig.apic_expected_date, equityConfig.apic_received_date);

        let html = '<table class="equity-display-table"><thead><tr>';
        html += '<th>Item</th><th>Amount</th><th>Expected</th><th>Received</th><th>Status</th>';
        html += '</tr></thead><tbody>';

        html += `<tr>
            <td>Seed Money (Common Stock)</td>
            <td>${fmtAmt(commonStock)}</td>
            <td>${equityConfig.seed_expected_date ? Utils.formatDate(equityConfig.seed_expected_date) : '—'}</td>
            <td>${equityConfig.seed_received_date ? Utils.formatDate(equityConfig.seed_received_date) : '—'}</td>
            <td>${seedStatus}</td>
        </tr>`;

        if (apicVal > 0) {
            html += `<tr>
                <td>Additional Paid-In Capital</td>
                <td>${fmtAmt(apicVal)}</td>
                <td>${equityConfig.apic_expected_date ? Utils.formatDate(equityConfig.apic_expected_date) : '—'}</td>
                <td>${equityConfig.apic_received_date ? Utils.formatDate(equityConfig.apic_received_date) : '—'}</td>
                <td>${apicStatus}</td>
            </tr>`;
        }

        html += `<tr class="equity-total-row">
            <td><strong>Total Stockholders' Equity</strong></td>
            <td><strong>${fmtAmt(totalEquity)}</strong></td>
            <td colspan="3"></td>
        </tr>`;

        html += '</tbody></table>';

        // Detail line
        if (equityConfig.common_stock_shares) {
            html += `<div class="equity-detail-line">${equityConfig.common_stock_shares.toLocaleString()} shares at ${fmtAmt(equityConfig.common_stock_par)} par value</div>`;
        }

        panel.innerHTML = html;
    },

    _equityStatusBadge(expectedDate, receivedDate) {
        if (receivedDate) {
            return '<span class="status-received">Received</span>';
        } else if (expectedDate) {
            return '<span class="status-pending">Pending</span>';
        }
        return '<span class="status-none">—</span>';
    },

    /**
     * Render the Loans tab with list/detail layout
     * @param {Array} loans - Array of loan objects
     * @param {number|null} selectedLoanId - Currently selected loan ID
     */
    renderLoansTab(loans, selectedLoanId) {
        const fmtAmt = (amt) => Utils.formatCurrency(amt);

        // Left panel: loan list
        const listPanel = document.getElementById('loanListPanel');
        if (loans.length === 0) {
            listPanel.innerHTML = '<p class="empty-state">No loans yet. Click "+ Add Loan" to begin.</p>';
        } else {
            listPanel.innerHTML = loans.map(loan => {
                const selected = loan.id === selectedLoanId ? ' selected' : '';
                return `<div class="loan-list-item${selected}" data-id="${loan.id}">
                    <div class="loan-list-name">${Utils.escapeHtml(loan.name)}</div>
                    <div class="loan-list-meta">${fmtAmt(loan.principal)} &middot; ${loan.annual_rate}%</div>
                    <div class="loan-list-actions">
                        <button class="btn-icon edit-loan-btn" data-id="${loan.id}" title="Edit">&#9998;</button>
                        <button class="btn-icon delete-loan-btn" data-id="${loan.id}" title="Delete">&times;</button>
                    </div>
                </div>`;
            }).join('');
        }

        // Right panel: selected loan detail
        const detailPanel = document.getElementById('loanDetailPanel');
        const selectedLoan = loans.find(l => l.id === selectedLoanId);
        if (!selectedLoan) {
            detailPanel.innerHTML = '<p class="empty-state">Select a loan to view its amortization schedule.</p>';
            return;
        }

        const skippedPayments = Database.getSkippedPayments(selectedLoan.id);
        const paymentOverrides = Database.getLoanPaymentOverrides(selectedLoan.id);
        const schedule = Utils.computeAmortizationSchedule({
            principal: selectedLoan.principal,
            annual_rate: selectedLoan.annual_rate,
            term_months: selectedLoan.term_months,
            payments_per_year: selectedLoan.payments_per_year,
            start_date: selectedLoan.start_date,
            first_payment_date: selectedLoan.first_payment_date
        }, skippedPayments, paymentOverrides);

        const totalInterest = schedule.filter(p => !p.skipped).reduce((sum, p) => sum + p.interest, 0);
        const totalPaid = schedule.filter(p => !p.skipped).reduce((sum, p) => sum + p.payment, 0);
        const skippedCount = schedule.filter(p => p.skipped).length;
        const termYears = (selectedLoan.term_months / 12).toFixed(1);

        // Use the most common (mode) payment amount, ignoring skipped and overridden payments
        const standardPayments = schedule.filter(p => !p.skipped && !p.overridden).map(p => p.payment);
        let modePayment = schedule[0]?.payment || 0;
        if (standardPayments.length > 0) {
            const freq = {};
            standardPayments.forEach(amt => {
                const key = amt.toFixed(2);
                freq[key] = (freq[key] || 0) + 1;
            });
            const modeKey = Object.keys(freq).reduce((a, b) => freq[a] >= freq[b] ? a : b);
            modePayment = parseFloat(modeKey);
        }

        let html = '<div class="loan-summary">';
        html += `<div class="loan-summary-item"><div class="loan-summary-label">Principal</div><div class="loan-summary-value">${fmtAmt(selectedLoan.principal)}</div></div>`;
        html += `<div class="loan-summary-item"><div class="loan-summary-label">Rate</div><div class="loan-summary-value">${selectedLoan.annual_rate}%</div></div>`;
        html += `<div class="loan-summary-item"><div class="loan-summary-label">Term</div><div class="loan-summary-value">${termYears} yr</div></div>`;
        html += `<div class="loan-summary-item"><div class="loan-summary-label">Payment</div><div class="loan-summary-value">${fmtAmt(modePayment)}</div></div>`;
        html += `<div class="loan-summary-item"><div class="loan-summary-label">Total Interest</div><div class="loan-summary-value amount-payable">${fmtAmt(totalInterest)}</div></div>`;
        html += `<div class="loan-summary-item"><div class="loan-summary-label">Total Paid</div><div class="loan-summary-value">${fmtAmt(totalPaid)}</div></div>`;
        if (skippedCount > 0) {
            html += `<div class="loan-summary-item"><div class="loan-summary-label">Skipped</div><div class="loan-summary-value amount-payable">${skippedCount}</div></div>`;
        }
        if (selectedLoan.first_payment_date) {
            html += `<div class="loan-summary-item"><div class="loan-summary-label">1st Payment</div><div class="loan-summary-value">${selectedLoan.first_payment_date}</div></div>`;
        }
        html += '</div>';

        html += '<div class="loan-table-wrapper"><table class="loan-table"><thead><tr>';
        html += '<th>#</th><th>Date</th><th>Payment</th><th>Principal</th><th>Interest</th><th>Balance</th><th></th>';
        html += '</tr></thead><tbody>';

        schedule.forEach(p => {
            const rowClass = p.skipped ? ' loan-payment-skipped' : (p.overridden ? ' loan-payment-overridden' : '');
            html += `<tr class="${rowClass}">`;
            html += `<td>${p.number}</td>`;
            html += `<td>${Utils.formatMonthShort(p.month)}</td>`;
            html += `<td class="loan-payment-cell loan-amount" data-loan-id="${selectedLoan.id}" data-payment="${p.number}">${p.skipped ? '—' : fmtAmt(p.payment)}</td>`;
            html += `<td>${p.skipped ? '—' : fmtAmt(p.principal)}</td>`;
            html += `<td class="amount-payable">${fmtAmt(p.interest)}</td>`;
            html += `<td>${fmtAmt(p.ending_balance)}</td>`;
            html += `<td><button class="loan-skip-btn" data-loan-id="${selectedLoan.id}" data-payment="${p.number}" title="${p.skipped ? 'Restore payment' : 'Skip payment'}">${p.skipped ? '&#8634;' : '&times;'}</button></td>`;
            html += '</tr>';
        });

        html += '</tbody></table></div>';

        if (selectedLoan.notes) {
            html += `<div class="loan-detail-notes">Notes: ${Utils.escapeHtml(selectedLoan.notes)}</div>`;
        }

        detailPanel.innerHTML = html;
    },

    /**
     * Render the Budget tab with list/detail layout
     * @param {Array} expenses - Array of budget expense objects
     * @param {number|null} selectedExpenseId - Currently selected expense ID
     */
    renderBudgetTab(expenses, groups, selectedExpenseId, collapsedGroups) {
        const fmtAmt = (amt) => Utils.formatCurrency(amt);
        const currentMonth = Utils.getCurrentMonth();
        collapsedGroups = collapsedGroups || new Set();

        const isActive = (exp) => exp.start_month <= currentMonth && (!exp.end_month || exp.end_month >= currentMonth);

        // Count active expenses and compute totals
        let activeCount = 0;
        let totalMonthly = 0;
        expenses.forEach(exp => {
            if (isActive(exp)) {
                activeCount++;
                totalMonthly += exp.monthly_amount;
            }
        });

        // Summary cards
        const summaryContainer = document.getElementById('budgetSummaryCards');
        summaryContainer.innerHTML = `
            <div class="budget-summary-card"><span class="budget-summary-label">Monthly Total</span><span class="budget-summary-value">${fmtAmt(totalMonthly)}</span></div>
            <div class="budget-summary-card"><span class="budget-summary-label">Active Expenses</span><span class="budget-summary-value">${activeCount}</span></div>
            <div class="budget-summary-card"><span class="budget-summary-label">Annual Estimate</span><span class="budget-summary-value">${fmtAmt(totalMonthly * 12)}</span></div>
        `;

        // ===== OLD LIST PANEL (left side) =====
        const listPanel = document.getElementById('budgetListPanel');
        if (expenses.length === 0) {
            listPanel.innerHTML = '<p class="empty-state">No budget expenses yet. Click "+ Add Expense" to begin.</p>';
        } else {
            listPanel.innerHTML = expenses.map(exp => {
                const selected = exp.id === selectedExpenseId ? ' selected' : '';
                const active = isActive(exp);
                const statusClass = active ? 'budget-active' : 'budget-inactive';
                return `<div class="budget-list-item${selected}" data-id="${exp.id}">
                    <div class="budget-list-name">${Utils.escapeHtml(exp.name)}</div>
                    <div class="budget-list-meta">${fmtAmt(exp.monthly_amount)} &middot; <span class="${statusClass}">${active ? 'Active' : 'Inactive'}</span></div>
                    <div class="budget-list-actions">
                        <button class="btn-icon edit-budget-btn" data-id="${exp.id}" title="Edit">&#9998;</button>
                        <button class="btn-icon delete-budget-btn" data-id="${exp.id}" title="Delete">&times;</button>
                    </div>
                </div>`;
            }).join('');
        }

        // ===== RIGHT PANEL: selected expense detail =====
        const detailPanel = document.getElementById('budgetDetailPanel');
        const selectedExpense = expenses.find(e => e.id === selectedExpenseId);
        if (!selectedExpense) {
            detailPanel.innerHTML = '<p class="empty-state">Select an expense to view its payment schedule.</p>';
        } else {
            const startMonth = selectedExpense.start_month;
            const endMonth = selectedExpense.end_month;
            const months = [];
            let m = startMonth;
            // For loan-linked expenses, show all months (up to term length); otherwise cap at 24
            const isLoanLinked = selectedExpense.notes && /^Auto-created from loan #(\d+)$/.test(selectedExpense.notes);
            let loanScheduleByMonth = null;
            let loanData = null;
            if (isLoanLinked) {
                const loanId = parseInt(selectedExpense.notes.match(/^Auto-created from loan #(\d+)$/)[1]);
                loanData = Database.getLoanById(loanId);
                if (loanData) {
                    const skippedPayments = Database.getSkippedPayments(loanId);
                    const paymentOverrides = Database.getLoanPaymentOverrides(loanId);
                    const schedule = Utils.computeAmortizationSchedule({
                        principal: loanData.principal, annual_rate: loanData.annual_rate,
                        payments_per_year: loanData.payments_per_year, term_months: loanData.term_months,
                        start_date: loanData.start_date, first_payment_date: loanData.first_payment_date
                    }, skippedPayments, paymentOverrides);
                    loanScheduleByMonth = {};
                    schedule.forEach(p => { loanScheduleByMonth[p.month] = p; });
                }
            }
            const maxMonths = isLoanLinked ? 600 : 24;
            for (let i = 0; i < maxMonths; i++) {
                if (endMonth && m > endMonth) break;
                months.push(m);
                m = Utils.nextMonth(m);
            }

            let html = `<div class="budget-detail-header">
                <h4>${Utils.escapeHtml(selectedExpense.name)}</h4>
                <div class="budget-detail-meta">
                    <span>Amount: ${fmtAmt(selectedExpense.monthly_amount)}/mo</span>
                    <span>Start: ${Utils.formatMonthShort(selectedExpense.start_month)}</span>
                    <span>End: ${selectedExpense.end_month ? Utils.formatMonthShort(selectedExpense.end_month) : 'Indefinite'}</span>
                    ${selectedExpense.category_name ? `<span>Category: ${Utils.escapeHtml(selectedExpense.category_name)}</span>` : ''}
                    ${selectedExpense.group_name ? `<span>Group: ${Utils.escapeHtml(selectedExpense.group_name)}</span>` : ''}
                </div>
            </div>`;

            // Load per-month overrides for this expense
            const overrides = Database.getBudgetExpenseOverrides(selectedExpense.id);

            html += '<div class="budget-schedule-wrapper"><table class="budget-schedule-table"><thead><tr>';
            html += '<th>Month</th><th>Amount</th><th>Cumulative</th>';
            html += '</tr></thead><tbody>';

            let cumulative = 0;
            months.forEach(month => {
                const overrideAmt = overrides[month];
                // For loan-linked expenses, use the amortization schedule amount
                let baseAmount = selectedExpense.monthly_amount;
                if (loanScheduleByMonth) {
                    const entry = loanScheduleByMonth[month];
                    if (entry && !entry.skipped) {
                        baseAmount = entry.payment;
                    } else if (entry && entry.skipped) {
                        baseAmount = 0;
                    }
                }
                const amount = overrideAmt !== undefined ? overrideAmt : baseAmount;
                cumulative += amount;
                const isCurrent = month === currentMonth;
                const isOverridden = overrideAmt !== undefined;
                html += `<tr${isCurrent ? ' class="budget-current-month"' : ''}>
                    <td>${Utils.formatMonthShort(month)}</td>
                    <td class="budget-amount-cell${isOverridden ? ' budget-amount-overridden' : ''}" data-expense-id="${selectedExpense.id}" data-month="${month}" data-default="${baseAmount}" title="Click to edit">${fmtAmt(amount)}${isOverridden ? ' <button class="budget-override-reset" data-expense-id="' + selectedExpense.id + '" data-month="' + month + '" title="Reset to default">&times;</button>' : ''}</td>
                    <td>${fmtAmt(cumulative)}</td>
                </tr>`;
            });

            const hitMaxCap = months.length === maxMonths && (!endMonth || m <= endMonth);
            if (!endMonth) {
                html += `<tr class="budget-continues-row"><td colspan="3">Continues indefinitely...</td></tr>`;
            } else if (hitMaxCap) {
                html += `<tr class="budget-continues-row"><td colspan="3">Showing first ${maxMonths} months \u2014 ends ${Utils.formatMonthShort(endMonth)}</td></tr>`;
            }

            html += '</tbody></table></div>';

            if (selectedExpense.notes) {
                html += `<div class="budget-detail-notes">Notes: ${Utils.escapeHtml(selectedExpense.notes)}</div>`;
            }

            detailPanel.innerHTML = html;
        }

        // ===== NEW GROUPED SECTIONS (below the main section) =====
        const container = document.getElementById('budgetGroupsContainer');

        if (groups.length === 0 || expenses.length === 0) {
            container.innerHTML = '';
            container.style.display = 'none';
        } else {
            container.style.display = '';

            // Helper to render a single expense row (clean: name left, amount right)
            const renderExpenseRow = (exp) => {
                const active = isActive(exp);
                return `<div class="budget-expense-row${!active ? ' budget-expense-inactive' : ''}" data-id="${exp.id}" draggable="true" data-group-id="${exp.group_id || ''}">
                    <span class="budget-expense-name">${Utils.escapeHtml(exp.name)}</span>
                    <span class="budget-expense-amount budget-amount" data-budget-id="${exp.id}">${fmtAmt(exp.monthly_amount)}</span>
                    <div class="budget-expense-actions">
                        <button class="btn-icon edit-budget-btn" data-id="${exp.id}" title="Edit">&#9998;</button>
                        <button class="btn-icon delete-budget-btn" data-id="${exp.id}" title="Delete">&times;</button>
                    </div>
                </div>`;
            };

            // Bucket expenses by group
            const groupedExpenses = {};
            const ungrouped = [];
            groups.forEach(g => { groupedExpenses[g.id] = []; });
            expenses.forEach(exp => {
                if (exp.group_id && groupedExpenses[exp.group_id]) {
                    groupedExpenses[exp.group_id].push(exp);
                } else {
                    ungrouped.push(exp);
                }
            });

            let html = '';
            groups.forEach(g => {
                const items = groupedExpenses[g.id];
                const collapsed = collapsedGroups.has(g.id);
                let groupMonthly = 0;
                items.forEach(exp => { if (isActive(exp)) groupMonthly += exp.monthly_amount; });
                const pct = totalMonthly > 0 ? ((groupMonthly / totalMonthly) * 100).toFixed(1) : '0.0';

                html += `<div class="budget-group-section" data-group-id="${g.id}">
                    <div class="budget-group-section-header${collapsed ? ' collapsed' : ''}" data-group-id="${g.id}">
                        <div class="budget-group-section-title">
                            <span class="budget-group-name" data-group-id="${g.id}">${Utils.escapeHtml(g.name).toUpperCase()}</span>
                        </div>
                        <div class="budget-group-section-stats">
                            <span class="budget-group-section-total">${fmtAmt(groupMonthly)}</span>
                            <span class="budget-group-section-pct">${pct}%</span>
                            <div class="budget-group-actions">
                                <button class="btn-icon edit-budget-group-btn" data-group-id="${g.id}" title="Rename">&#9998;</button>
                                <button class="btn-icon delete-budget-group-btn" data-group-id="${g.id}" title="Delete Group">&times;</button>
                            </div>
                        </div>
                    </div>
                    <div class="budget-group-section-body${collapsed ? ' collapsed' : ''}" data-group-id="${g.id}">
                        ${items.length > 0 ? items.map(renderExpenseRow).join('') : '<div class="budget-group-empty">Drag expenses here or add a new expense to this group.</div>'}
                    </div>
                </div>`;
            });

            if (ungrouped.length > 0) {
                let ungroupedMonthly = 0;
                ungrouped.forEach(exp => { if (isActive(exp)) ungroupedMonthly += exp.monthly_amount; });
                const pct = totalMonthly > 0 ? ((ungroupedMonthly / totalMonthly) * 100).toFixed(1) : '0.0';

                html += `<div class="budget-group-section" data-group-id="">
                    <div class="budget-group-section-header budget-ungrouped-header" data-group-id="">
                        <div class="budget-group-section-title">
                            <span class="budget-group-name">UNGROUPED</span>
                        </div>
                        <div class="budget-group-section-stats">
                            <span class="budget-group-section-total">${fmtAmt(ungroupedMonthly)}</span>
                            <span class="budget-group-section-pct">${pct}%</span>
                        </div>
                    </div>
                    <div class="budget-group-section-body" data-group-id="">
                        ${ungrouped.map(renderExpenseRow).join('')}
                    </div>
                </div>`;
            }

            container.innerHTML = html;
        }

        // Pie chart — render when groups exist and there are active expenses
        const chartSection = document.getElementById('budgetChartSection');
        if (groups.length > 0 && totalMonthly > 0) {
            chartSection.style.display = '';
            this._renderBudgetPieChart(expenses, groups, totalMonthly, isActive);
        } else {
            chartSection.style.display = 'none';
        }
    },

    _budgetPieChart: null,

    _renderBudgetPieChart(expenses, groups, totalMonthly, isActive) {
        const canvas = document.getElementById('budgetPieChart');
        if (!canvas || typeof Chart === 'undefined') return;

        // Calculate group totals
        const groupTotals = {};
        groups.forEach(g => { groupTotals[g.id] = { name: g.name, monthly: 0 }; });
        let ungroupedMonthly = 0;
        expenses.forEach(exp => {
            if (!isActive(exp)) return;
            if (exp.group_id && groupTotals[exp.group_id]) {
                groupTotals[exp.group_id].monthly += exp.monthly_amount;
            } else {
                ungroupedMonthly += exp.monthly_amount;
            }
        });

        const labels = [];
        const data = [];
        const colors = [
            '#2c5f7c', '#3a7ca5', '#4a9ec9', '#6bb5d6', '#8ecae6',
            '#457b6b', '#5a9a89', '#78b5a3', '#95d0be', '#b2e0d4',
            '#8b6f47', '#a6885c', '#c1a172', '#d9bc90', '#e8d5b0',
            '#7c4c5e', '#996178', '#b37892', '#cc92ac', '#e0adc4'
        ];

        groups.forEach(g => {
            const gt = groupTotals[g.id];
            if (gt.monthly > 0) {
                labels.push(gt.name);
                data.push(Math.round(gt.monthly * 100) / 100);
            }
        });
        if (ungroupedMonthly > 0) {
            labels.push('Ungrouped');
            data.push(Math.round(ungroupedMonthly * 100) / 100);
        }

        if (this._budgetPieChart) {
            this._budgetPieChart.destroy();
        }

        const ctx = canvas.getContext('2d');
        this._budgetPieChart = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: labels,
                datasets: [{
                    data: data,
                    backgroundColor: colors.slice(0, labels.length),
                    borderWidth: 2,
                    borderColor: getComputedStyle(document.documentElement).getPropertyValue('--surface').trim() || '#fff'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                cutout: '50%',
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            color: getComputedStyle(document.documentElement).getPropertyValue('--text').trim() || '#333',
                            font: { size: 11 },
                            padding: 10,
                            usePointStyle: true,
                            pointStyle: 'circle',
                            boxWidth: 6,
                            boxHeight: 6
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const val = context.parsed;
                                const pct = ((val / totalMonthly) * 100).toFixed(1);
                                return ` ${context.label}: ${Utils.formatCurrency(val)}/mo (${pct}%)`;
                            }
                        }
                    }
                }
            }
        });
    },

    /**
     * Generate CSV string from transaction data
     * @param {Array} transactions - Array of transaction objects for export
     * @returns {string} CSV string
     */
    generateCsv(transactions) {
        const headers = [
            'Entry Date',
            'Category',
            'Type',
            'Amount',
            'Pretax Amount',
            'Status',
            'Month Due',
            'Month Paid',
            'Date Processed',
            'Payment For',
            'Notes'
        ];

        const escapeCsvField = (val) => {
            if (val === null || val === undefined) return '';
            const str = String(val);
            if (str.includes(',') || str.includes('"') || str.includes('\n')) {
                return '"' + str.replace(/"/g, '""') + '"';
            }
            return str;
        };

        const rows = transactions.map(t => [
            t.entry_date || '',
            t.category || '',
            t.type || '',
            t.amount || 0,
            t.pretax_amount || '',
            t.status || '',
            t.month_due ? Utils.formatMonthShort(t.month_due) : '',
            t.month_paid ? Utils.formatMonthShort(t.month_paid) : '',
            t.date_processed || '',
            t.payment_for_month ? Utils.formatMonthShort(t.payment_for_month) : '',
            t.notes || ''
        ]);

        const csvLines = [
            headers.map(escapeCsvField).join(','),
            ...rows.map(row => row.map(escapeCsvField).join(','))
        ];

        return csvLines.join('\n');
    },

    /**
     * Show the notes tooltip near the target element
     * @param {string} notes - The notes text
     * @param {HTMLElement} target - The element to position near
     */
    showNotesTooltip(notes, target) {
        const tooltip = document.getElementById('notesTooltip');
        tooltip.textContent = notes;
        tooltip.classList.add('visible');

        const rect = target.getBoundingClientRect();
        tooltip.style.left = rect.left + 'px';
        tooltip.style.top = (rect.bottom + 8) + 'px';

        // Keep tooltip on screen
        const tooltipRect = tooltip.getBoundingClientRect();
        if (tooltipRect.right > window.innerWidth) {
            tooltip.style.left = (window.innerWidth - tooltipRect.width - 16) + 'px';
        }
        if (tooltipRect.bottom > window.innerHeight) {
            tooltip.style.top = (rect.top - tooltipRect.height - 8) + 'px';
        }
    },

    /**
     * Hide the notes tooltip
     */
    hideNotesTooltip() {
        const tooltip = document.getElementById('notesTooltip');
        tooltip.classList.remove('visible');
    },

    /**
     * Show modal
     * @param {string} modalId - ID of the modal element
     */
    showModal(modalId) {
        document.getElementById(modalId).classList.add('active');
    },

    /**
     * Hide modal
     * @param {string} modalId - ID of the modal element
     */
    hideModal(modalId) {
        document.getElementById(modalId).classList.remove('active');
    },

    /**
     * Reset the entry form
     */
    resetForm() {
        document.getElementById('entryForm').reset();
        document.getElementById('editingId').value = '';
        document.getElementById('entryDate').value = Utils.getTodayDate();
        document.getElementById('formTitle').textContent = 'Add New Entry';
        document.getElementById('submitBtn').textContent = 'Add Entry';

        // Reset radio buttons
        document.querySelector('input[name="transactionType"][value="receivable"]').checked = true;

        // Update status dropdown for receivable type
        this.updateStatusOptions('receivable');

        // Reset status to pending and update field visibility
        document.getElementById('status').value = 'pending';
        this.updateFormFieldVisibility('pending');

        // Hide payment for month field
        this.togglePaymentForMonth(false);

        // Hide and clear pretax amount
        document.getElementById('pretaxAmountGroup').style.display = 'none';
        document.getElementById('pretaxAmount').value = '';

        // Hide and clear inventory cost
        document.getElementById('inventoryCostGroup').style.display = 'none';
        document.getElementById('inventoryCost').value = '';

        // Hide and clear sale date fields
        document.getElementById('saleDateGroup').style.display = 'none';
        document.getElementById('saleDateStart').value = '';
        document.getElementById('saleDateEnd').value = '';

        // Clear month due/paid inputs
        document.getElementById('monthDue').value = '';
        document.getElementById('monthPaid').value = '';

        // Close the entry modal
        this.hideModal('entryModal');
    },

    /**
     * Populate form for editing
     * @param {Object} transaction - Transaction object
     */
    populateFormForEdit(transaction) {
        document.getElementById('editingId').value = transaction.id;
        document.getElementById('entryDate').value = transaction.entry_date;
        document.getElementById('category').value = transaction.category_id;
        document.getElementById('amount').value = transaction.amount;
        document.getElementById('dateProcessed').value = transaction.date_processed || '';
        document.getElementById('notes').value = transaction.notes || '';

        // Set transaction type radio
        const typeRadio = document.querySelector(`input[name="transactionType"][value="${transaction.transaction_type}"]`);
        if (typeRadio) typeRadio.checked = true;

        // Update status options for the transaction type
        this.updateStatusOptions(transaction.transaction_type);
        document.getElementById('status').value = transaction.status;

        // Update field visibility based on status
        this.updateFormFieldVisibility(transaction.status);

        // Set month due
        document.getElementById('monthDue').value = transaction.month_due || '';

        // Set month paid
        document.getElementById('monthPaid').value = transaction.month_paid || '';

        // Handle pretax amount and sale date fields (only for sales categories)
        const pretaxGroup = document.getElementById('pretaxAmountGroup');
        const saleDateGroup = document.getElementById('saleDateGroup');
        const inventoryCostGroup = document.getElementById('inventoryCostGroup');
        if (transaction.category_is_sales) {
            pretaxGroup.style.display = 'flex';
            document.getElementById('pretaxAmount').value = transaction.pretax_amount || '';
            inventoryCostGroup.style.display = 'flex';
            document.getElementById('inventoryCost').value = transaction.inventory_cost || '';
            saleDateGroup.style.display = 'flex';
            document.getElementById('saleDateStart').value = transaction.sale_date_start || '';
            document.getElementById('saleDateEnd').value = transaction.sale_date_end || '';
        } else {
            pretaxGroup.style.display = 'none';
            document.getElementById('pretaxAmount').value = '';
            inventoryCostGroup.style.display = 'none';
            document.getElementById('inventoryCost').value = '';
            saleDateGroup.style.display = 'none';
            document.getElementById('saleDateStart').value = '';
            document.getElementById('saleDateEnd').value = '';
        }

        // Handle payment for month if category is monthly
        if (transaction.category_is_monthly) {
            this.togglePaymentForMonth(true, transaction.category_name);
            document.getElementById('paymentForMonth').value = transaction.payment_for_month || '';
        } else {
            this.togglePaymentForMonth(false);
        }

        // Update form title and button
        document.getElementById('formTitle').textContent = 'Edit Entry';
        document.getElementById('submitBtn').textContent = 'Update Entry';

        // Open entry modal
        this.showModal('entryModal');
    },

    /**
     * Update status dropdown options based on transaction type
     * @param {string} type - 'receivable' or 'payable'
     */
    updateStatusOptions(type) {
        const statusSelect = document.getElementById('status');
        const currentValue = statusSelect.value;

        if (type === 'receivable') {
            statusSelect.innerHTML = `
                <option value="pending">Pending</option>
                <option value="received">Received</option>
            `;
        } else {
            statusSelect.innerHTML = `
                <option value="pending">Pending</option>
                <option value="paid">Paid</option>
            `;
        }

        if (currentValue === 'pending') {
            statusSelect.value = 'pending';
        }
    },

    /**
     * Get form data
     * @returns {Object} Form data object
     */
    getFormData() {
        const transactionType = document.querySelector('input[name="transactionType"]:checked').value;
        const paymentForGroup = document.getElementById('paymentForGroup');
        const paymentForMonth = paymentForGroup.style.display !== 'none'
            ? document.getElementById('paymentForMonth').value || null
            : null;

        const month_due = document.getElementById('monthDue').value || null;
        const month_paid = document.getElementById('monthPaid').value || null;

        const status = document.getElementById('status').value;

        // Pretax amount (only when visible / sales categories)
        const pretaxGroup = document.getElementById('pretaxAmountGroup');
        const pretaxAmount = (pretaxGroup && pretaxGroup.style.display !== 'none')
            ? Utils.parseAmount(document.getElementById('pretaxAmount').value) || null
            : null;

        // Inventory cost (only when visible / sales categories)
        const inventoryCostGroup = document.getElementById('inventoryCostGroup');
        const inventoryCost = (inventoryCostGroup && inventoryCostGroup.style.display !== 'none')
            ? Utils.parseAmount(document.getElementById('inventoryCost').value) || null
            : null;

        // Sale date fields (only when visible / sales categories)
        const saleDateGroup = document.getElementById('saleDateGroup');
        let saleDateStart = null;
        let saleDateEnd = null;
        if (saleDateGroup && saleDateGroup.style.display !== 'none') {
            saleDateStart = document.getElementById('saleDateStart').value || null;
            saleDateEnd = document.getElementById('saleDateEnd').value || null;
            // Single day: if only start is set, treat end as same day
            if (saleDateStart && !saleDateEnd) {
                saleDateEnd = saleDateStart;
            }
        }

        return {
            entry_date: document.getElementById('entryDate').value,
            category_id: parseInt(document.getElementById('category').value),
            amount: Utils.parseAmount(document.getElementById('amount').value),
            pretax_amount: pretaxAmount,
            transaction_type: transactionType,
            status: status,
            date_processed: (status !== 'pending') ? (document.getElementById('dateProcessed').value || null) : null,
            month_due: month_due,
            month_paid: (status !== 'pending') ? month_paid : null,
            payment_for_month: paymentForMonth,
            notes: document.getElementById('notes').value.trim() || null,
            sale_date_start: saleDateStart,
            sale_date_end: saleDateEnd,
            inventory_cost: inventoryCost
        };
    },

    /**
     * Validate form data
     * @param {Object} data - Form data
     * @returns {Object} Validation result {valid: boolean, message: string}
     */
    validateFormData(data) {
        if (!data.entry_date) {
            return { valid: false, message: 'Entry date is required' };
        }
        if (!data.category_id) {
            return { valid: false, message: 'Please select a category' };
        }
        if (!data.amount || data.amount <= 0) {
            return { valid: false, message: 'Please enter a valid amount' };
        }
        // Require month paid when status is paid or received
        if (data.status !== 'pending' && !data.month_paid) {
            return { valid: false, message: 'Month paid/received is required when status is not pending' };
        }
        // Validate sale dates are in the same month
        if (data.sale_date_start && data.sale_date_end) {
            if (data.sale_date_start.substring(0, 7) !== data.sale_date_end.substring(0, 7)) {
                return { valid: false, message: 'Sale period must be within a single month' };
            }
            if (data.sale_date_start > data.sale_date_end) {
                return { valid: false, message: 'Sale start date must be before or equal to end date' };
            }
        }
        return { valid: true };
    },

    /**
     * Show notification message
     * @param {string} message - Message to show
     * @param {string} type - 'success', 'error', or 'info'
     */
    showNotification(message, type = 'info', duration = 3000) {
        const existing = document.querySelector('.notification');
        if (existing) existing.remove();

        const notification = document.createElement('div');
        notification.className = `notification notification-${type}`;

        // Icon map
        const icons = {
            success: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 6L9 17l-5-5"/></svg>',
            error: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>',
            warning: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',
            info: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>'
        };
        notification.innerHTML = `${icons[type] || icons.info}<span>${message}</span><div class="notification-progress" style="width: 100%"></div>`;

        document.body.appendChild(notification);

        // Trigger reflow then add visible class for CSS transition
        notification.offsetHeight;
        notification.classList.add('visible');

        // Animate progress bar
        const progress = notification.querySelector('.notification-progress');
        if (progress) {
            progress.style.transitionDuration = duration + 'ms';
            requestAnimationFrame(() => { progress.style.width = '0%'; });
        }

        setTimeout(() => {
            notification.classList.remove('visible');
            setTimeout(() => notification.remove(), 200);
        }, duration);
    },

    /**
     * Capitalize first letter
     * @param {string} str - String to capitalize
     * @returns {string} Capitalized string
     */
    capitalizeFirst(str) {
        if (!str) return '';
        return str.charAt(0).toUpperCase() + str.slice(1);
    },

    // ==================== BREAK-EVEN RENDERING ====================

    /**
     * Render the 4 break-even summary cards with fixed cost breakdown
     * @param {Object} result - From Utils.computeBreakEven()
     * @param {Object} fixedBreakdown - {budget, depreciation, loanInterest} avg monthly values
     */
    renderBreakevenSummaryCards(result, fixedBreakdown, cfg) {
        const container = document.getElementById('beSummaryCards');
        if (!result.isValid) {
            container.innerHTML = '<p class="empty-state">Configure selling price and COGS to calculate break-even.</p>';
            return;
        }

        // Build fixed cost breakdown lines (from P&L)
        let breakdownHtml = '';
        if (fixedBreakdown) {
            const lines = [];
            if (fixedBreakdown.opex > 0) lines.push(`OpEx: ${Utils.formatCurrency(fixedBreakdown.opex)}`);
            if (fixedBreakdown.depreciation > 0) lines.push(`Depreciation: ${Utils.formatCurrency(fixedBreakdown.depreciation)}`);
            if (fixedBreakdown.loanInterest > 0) lines.push(`Loan Interest: ${Utils.formatCurrency(fixedBreakdown.loanInterest)}`);
            if (lines.length > 0) {
                breakdownHtml = `<span class="be-card-breakdown">${lines.join('<br>')}</span>`;
            }
        }

        // Per-stream monthly revenue cards
        let revenueCardsHtml = '';
        if (cfg && cfg.b2b && cfg.b2b.enabled && result.b2bMonthlyRevenue > 0) {
            revenueCardsHtml += `
                <div class="be-card">
                    <span class="be-card-label">B2B Monthly Revenue</span>
                    <span class="be-card-value">${Utils.formatCurrency(result.b2bMonthlyRevenue)}</span>
                </div>
            `;
        }
        if (cfg && cfg.consumer && cfg.consumer.enabled && cfg.consumer.avgPrice > 0 && result.consumerUnitsNeeded > 0) {
            const consumerMonthlyRev = result.consumerUnitsNeeded * cfg.consumer.avgPrice;
            revenueCardsHtml += `
                <div class="be-card">
                    <span class="be-card-label">Consumer Monthly Revenue</span>
                    <span class="be-card-value">${Utils.formatCurrency(consumerMonthlyRev)}</span>
                    <span class="be-card-breakdown">${result.consumerUnitsNeeded.toLocaleString()} units at BE</span>
                </div>
            `;
        }

        container.innerHTML = `
            <div class="be-card">
                <span class="be-card-label">Monthly Fixed Costs</span>
                <span class="be-card-value">${Utils.formatCurrency(result.monthlyFixedCosts)}</span>
                ${breakdownHtml}
            </div>
            <div class="be-card be-card-highlight">
                <span class="be-card-label">Break-Even Units/Mo</span>
                <span class="be-card-value">${result.breakEvenUnits.toLocaleString()}</span>
            </div>
            ${(cfg && cfg.b2b && cfg.b2b.enabled && cfg.b2b.monthlyUnits > 0 && result.consumerUnitsNeeded > 0) ? `
            <div class="be-card">
                <span class="be-card-label">Consumer BE Units/Mo</span>
                <span class="be-card-value">${result.consumerUnitsNeeded.toLocaleString()}</span>
                <span class="be-card-breakdown">Excludes B2B</span>
            </div>
            ` : ''}
            <div class="be-card">
                <span class="be-card-label">Break-Even Revenue/Mo</span>
                <span class="be-card-value">${Utils.formatCurrency(result.breakEvenRevenue)}</span>
            </div>
            <div class="be-card">
                <span class="be-card-label">Weighted Gross Margin</span>
                <span class="be-card-value">${result.weightedCM.toFixed(1)}%</span>
            </div>
            ${revenueCardsHtml}
        `;
    },

    /**
     * Render per-channel breakdown cards
     * @param {Object} result - From Utils.computeBreakEven()
     * @param {Object} cfg - Break-even config
     */
    renderBreakevenChannelBreakdown(result, cfg) {
        const container = document.getElementById('beChannelBreakdown');
        let html = '';

        if (cfg.consumer && cfg.consumer.enabled && cfg.consumer.avgPrice > 0) {
            html += `
                <div class="be-channel-card">
                    <div class="be-channel-title">Consumer Sales</div>
                    <div class="be-channel-row">
                        <span>Avg. Price</span><span>${Utils.formatCurrency(cfg.consumer.avgPrice)}</span>
                    </div>
                    <div class="be-channel-row">
                        <span>Avg. COGS</span><span>${Utils.formatCurrency(cfg.consumer.avgCogs || 0)}</span>
                    </div>
                    <div class="be-channel-row be-channel-highlight">
                        <span>CM / Unit</span><span>${Utils.formatCurrency(result.consumerCM)}</span>
                    </div>
                    <div class="be-channel-row">
                        <span>Gross Margin</span><span>${result.consumerCMPercent.toFixed(1)}%</span>
                    </div>
                    <div class="be-channel-row be-channel-highlight">
                        <span>Units to Break-Even</span><span>${result.consumerUnitsNeeded.toLocaleString()}</span>
                    </div>
                </div>
            `;
        }

        if (cfg.b2b && cfg.b2b.enabled && cfg.b2b.monthlyUnits > 0) {
            html += `
                <div class="be-channel-card">
                    <div class="be-channel-title">B2B Contract</div>
                    <div class="be-channel-row">
                        <span>Rate / Unit</span><span>${Utils.formatCurrency(cfg.b2b.ratePerUnit || 0)}</span>
                    </div>
                    <div class="be-channel-row">
                        <span>COGS / Unit</span><span>${Utils.formatCurrency(cfg.b2b.cogsPerUnit || 0)}</span>
                    </div>
                    <div class="be-channel-row be-channel-highlight">
                        <span>CM / Unit</span><span>${Utils.formatCurrency(result.b2bCM)}</span>
                    </div>
                    <div class="be-channel-row">
                        <span>Gross Margin</span><span>${result.b2bCMPercent.toFixed(1)}%</span>
                    </div>
                    <div class="be-channel-row">
                        <span>Monthly Units</span><span>${(cfg.b2b.monthlyUnits || 0).toLocaleString()}</span>
                    </div>
                    <div class="be-channel-row be-channel-highlight">
                        <span>Monthly Contribution</span><span>${Utils.formatCurrency(result.b2bMonthlyContribution)}</span>
                    </div>
                </div>
            `;
        }

        container.innerHTML = html;
    },

    /**
     * Render break-even progress summary cards
     * @param {Object} data - Progress data from _computeAndRenderProgress()
     */
    renderBreakevenProgressCards(data) {
        const container = document.getElementById('beProgressCards');

        const statusClass = data.onTrack ? 'be-progress-on-track' : 'be-progress-behind';
        const statusText = data.onTrack ? 'On Track' : 'Behind';
        const statusIcon = data.onTrack ? '\u2713' : '\u2717';

        // B2B/Consumer breakdown for actual revenue
        let actualBreakdown = '';
        if (data.actualB2b > 0 || data.actualConsumer > 0) {
            const lines = [];
            if (data.actualConsumer > 0) lines.push(`Consumer: ${Utils.formatCurrency(data.actualConsumer)}`);
            if (data.actualB2b > 0) lines.push(`B2B: ${Utils.formatCurrency(data.actualB2b)}`);
            actualBreakdown = `<span class="be-card-breakdown">${lines.join('<br>')}</span>`;
        }

        container.innerHTML = `
            <div class="be-card">
                <span class="be-card-label">Actual Gross Profit</span>
                <span class="be-card-value">${Utils.formatCurrency(data.actualTotal)}</span>
                ${actualBreakdown}
                <span class="be-card-breakdown">${data.elapsedCount} of ${data.totalMonths} months</span>
            </div>
            <div class="be-card">
                <span class="be-card-label">Target (On Pace)</span>
                <span class="be-card-value">${Utils.formatCurrency(data.targetByNow)}</span>
                <span class="be-card-breakdown">Expected by ${Utils.formatMonthShort(data.asOfMonth)}</span>
            </div>
            <div class="be-card ${statusClass}">
                <span class="be-card-label">Status</span>
                <span class="be-card-value">${statusIcon} ${statusText}</span>
                <span class="be-card-breakdown">${data.onTrack ? 'Ahead by' : 'Behind by'} ${Utils.formatCurrency(Math.abs(data.actualTotal - data.targetByNow))}</span>
            </div>
            <div class="be-card be-card-highlight">
                <span class="be-card-label">Remaining Profit Needed</span>
                <span class="be-card-value">${Utils.formatCurrency(data.remainingRevenue)}</span>
                <span class="be-card-breakdown">${data.remainingCount} months left</span>
            </div>
            <div class="be-card">
                <span class="be-card-label">Monthly Needed (Remaining)</span>
                <span class="be-card-value">${Utils.formatCurrency(data.monthlyNeeded)}</span>
                <span class="be-card-breakdown">Per month to break even</span>
            </div>
            <div class="be-card">
                <span class="be-card-label">Total BE Profit Target</span>
                <span class="be-card-value">${Utils.formatCurrency(data.totalBERevenue)}</span>
                <span class="be-card-breakdown">Full ${data.totalMonths}-month target</span>
            </div>
        `;
    },

    /**
     * Render the break-even data points table with separate B2B and consumer columns
     * @param {Array} points - From Utils.computeBreakEvenChartPoints()
     * @param {number} consumerBETotal - Consumer break-even units (timeline total)
     * @param {number} b2bTotal - Total B2B units across timeline
     * @param {number} increment - Unit axis increment
     */
    renderBreakevenDataTable(points, b2bTotal, increment, consumerBETotal, exactBEPoint) {
        const container = document.getElementById('beDataTable');
        if (!points || points.length === 0) {
            container.innerHTML = '';
            return;
        }

        // Merge exact break-even point into table rows (if not already on an increment)
        let tableRows = points.map(p => ({ ...p, isExactBE: false }));
        if (exactBEPoint && consumerBETotal > 0) {
            const alreadyExists = tableRows.some(p => p.consumerUnits === consumerBETotal);
            if (!alreadyExists) {
                tableRows.push({ ...exactBEPoint, isExactBE: true });
                tableRows.sort((a, b) => a.consumerUnits - b.consumerUnits);
            } else {
                tableRows = tableRows.map(p =>
                    p.consumerUnits === consumerBETotal ? { ...p, isExactBE: true } : p
                );
            }
        }

        let html = `
            <div class="be-table-header-info">
                <span><strong>Commercial (B2B):</strong> ${b2bTotal.toLocaleString()} units</span>
                <span><strong>Consumer Break-Even:</strong> ${(consumerBETotal || 0).toLocaleString()} units</span>
                <span><strong>Increment:</strong> ${increment.toLocaleString()}</span>
            </div>
            <table class="be-table">
                <thead>
                    <tr>
                        <th>Commercial Sold</th>
                        <th>Consumer Units Sold</th>
                        <th>Revenue</th>
                        <th>Variable Cost</th>
                        <th>Fixed Cost</th>
                        <th>Total Cost</th>
                    </tr>
                </thead>
                <tbody>
        `;

        tableRows.forEach(p => {
            const profit = p.revenue - p.totalCosts;
            const isBreakEven = p.isExactBE || p.consumerUnits === consumerBETotal;
            const rowClass = isBreakEven ? 'be-row-breakeven' : (profit < 0 ? 'be-row-loss' : 'be-row-profit');
            html += `
                <tr class="${rowClass}">
                    <td>${p.b2bUnits.toLocaleString()}</td>
                    <td>${p.consumerUnits.toLocaleString()}${isBreakEven ? ' ★' : ''}</td>
                    <td>${Utils.formatCurrency(p.revenue)}</td>
                    <td>${Utils.formatCurrency(p.variableCosts)}</td>
                    <td>${Utils.formatCurrency(p.fixedCosts)}</td>
                    <td>${Utils.formatCurrency(p.totalCosts)}</td>
                </tr>
            `;
        });

        html += '</tbody></table>';
        container.innerHTML = html;
    },

    // ==================== PROJECTED SALES TAB ====================

    renderProjectedSalesSummaryCards(psData, months) {
        const container = document.getElementById('psSummaryCards');
        const { byMonth, config } = psData;

        if (!config.enabled) {
            container.innerHTML = '';
            return;
        }

        let totalRevenue = 0, totalCogs = 0, totalUnits = 0;
        const projMonths = months.filter(m => byMonth[m]);
        projMonths.forEach(m => {
            totalRevenue += byMonth[m].revenue;
            totalCogs += byMonth[m].cogs;
            totalUnits += (byMonth[m].onlineUnits || 0) + (byMonth[m].tradeshowUnits || 0);
        });

        const avgMonthlyRev = projMonths.length > 0 ? totalRevenue / projMonths.length : 0;
        const gp = totalRevenue - totalCogs;
        const margin = totalRevenue > 0 ? ((gp / totalRevenue) * 100).toFixed(1) + '%' : '-';

        container.innerHTML = `
            <div class="be-card"><span class="be-card-label">Projected Revenue</span><span class="be-card-value">${Utils.formatCurrency(totalRevenue)}</span></div>
            <div class="be-card"><span class="be-card-label">Projected COGS</span><span class="be-card-value">${Utils.formatCurrency(totalCogs)}</span></div>
            <div class="be-card"><span class="be-card-label">Gross Profit</span><span class="be-card-value">${Utils.formatCurrency(gp)}</span></div>
            <div class="be-card"><span class="be-card-label">Gross Margin</span><span class="be-card-value">${margin}</span></div>
            <div class="be-card"><span class="be-card-label">Total Units</span><span class="be-card-value">${totalUnits.toLocaleString()}</span></div>
            <div class="be-card"><span class="be-card-label">Avg Monthly Rev</span><span class="be-card-value">${Utils.formatCurrency(avgMonthlyRev)}</span></div>
        `;
    },

    renderProjectedSalesGrid(psData, months, currentMonth) {
        const container = document.getElementById('psMonthlyGrid');
        const { byMonth, channels, config } = psData;

        if (!config.enabled || months.length === 0) {
            container.innerHTML = '<p class="empty-state">Enable projections and set a start month to view the grid.</p>';
            return;
        }

        const fmtMonth = (m) => Utils.formatMonthShort(m);
        const fmtAmt = (amt) => Utils.formatCurrency(amt);
        const projStart = config.projectionStartMonth;
        const isProj = (m) => projStart && m >= projStart;
        const isFuture = (m) => currentMonth && m > currentMonth;
        const colSpan = months.length + 2;

        let html = '<table class="pnl-table ps-table"><thead><tr><th></th>';
        months.forEach(m => {
            const projClass = isProj(m) ? ' pnl-future-header' : '';
            const overClass = isProj(m) && !isFuture(m) ? ' ps-overwrite-header' : '';
            const badge = isProj(m) ? ' <span class="projected-badge">P</span>' : '';
            html += `<th class="${projClass}${overClass}">${fmtMonth(m)}${badge}</th>`;
        });
        html += '<th>Total</th></tr></thead><tbody>';

        // Per-channel sections
        ['online', 'tradeshow'].forEach(key => {
            const ch = channels[key];
            if (!ch || !ch.enabled) return;
            const label = key === 'online' ? 'Online' : 'Tradeshow';

            html += `<tr class="pnl-section-header"><td colspan="${colSpan}">${label} Channel (${fmtAmt(ch.avgPrice)}/unit, COGS ${fmtAmt(ch.avgCogs)}/unit)</td></tr>`;

            // Units row (editable for projected months)
            let totalUnits = 0;
            html += '<tr class="pnl-indent"><td>Units</td>';
            months.forEach(m => {
                const units = (byMonth[m] && byMonth[m][key + 'Units']) || 0;
                totalUnits += units;
                if (isProj(m)) {
                    html += `<td class="ps-unit-editable" data-channel="${key}" data-month="${m}">${units || ''}</td>`;
                } else {
                    html += `<td>${units || ''}</td>`;
                }
            });
            html += `<td>${totalUnits}</td></tr>`;

            // % Change row
            html += '<tr class="pnl-percentage"><td>% Change</td>';
            let prevUnits = 0;
            months.forEach((m, idx) => {
                const units = (byMonth[m] && byMonth[m][key + 'Units']) || 0;
                let pct = '-';
                if (idx > 0 && prevUnits > 0) {
                    const change = ((units - prevUnits) / prevUnits * 100).toFixed(1);
                    pct = (change >= 0 ? '+' : '') + change + '%';
                }
                html += `<td>${pct}</td>`;
                prevUnits = units;
            });
            html += '<td></td></tr>';

            // Revenue row
            let totalRev = 0;
            html += '<tr class="pnl-indent"><td>Revenue</td>';
            months.forEach(m => {
                const rev = (byMonth[m] && byMonth[m][key + 'Revenue']) || 0;
                totalRev += rev;
                html += `<td class="amount-receivable ps-amount" data-ps-month="${m}">${rev ? fmtAmt(rev) : ''}</td>`;
            });
            html += `<td class="amount-receivable ps-amount" data-ps-month="total">${fmtAmt(totalRev)}</td></tr>`;

            // COGS row
            let totalCogs = 0;
            html += '<tr class="pnl-indent"><td>COGS</td>';
            months.forEach(m => {
                const cg = (byMonth[m] && byMonth[m][key + 'Cogs']) || 0;
                totalCogs += cg;
                html += `<td class="amount-payable ps-amount" data-ps-month="${m}">${cg ? fmtAmt(cg) : ''}</td>`;
            });
            html += `<td class="amount-payable">${fmtAmt(totalCogs)}</td></tr>`;

            // Gross Profit row
            let totalGP = 0;
            html += '<tr class="pnl-subtotal"><td>Gross Profit</td>';
            months.forEach(m => {
                const gp = ((byMonth[m] && byMonth[m][key + 'Revenue']) || 0) - ((byMonth[m] && byMonth[m][key + 'Cogs']) || 0);
                totalGP += gp;
                html += `<td>${fmtAmt(gp)}</td>`;
            });
            html += `<td>${fmtAmt(totalGP)}</td></tr>`;
        });

        // Combined totals
        html += `<tr class="pnl-section-header"><td colspan="${colSpan}">Combined Non-B2B Projected Totals</td></tr>`;

        let grandRevenue = 0;
        html += '<tr class="pnl-subtotal"><td>Total Projected Revenue</td>';
        months.forEach(m => {
            const rev = (byMonth[m] && byMonth[m].revenue) || 0;
            grandRevenue += rev;
            html += `<td class="amount-receivable">${fmtAmt(rev)}</td>`;
        });
        html += `<td class="amount-receivable">${fmtAmt(grandRevenue)}</td></tr>`;

        let grandCogs = 0;
        html += '<tr class="pnl-subtotal"><td>Total Projected COGS</td>';
        months.forEach(m => {
            const cg = (byMonth[m] && byMonth[m].cogs) || 0;
            grandCogs += cg;
            html += `<td class="amount-payable">${fmtAmt(cg)}</td>`;
        });
        html += `<td class="amount-payable">${fmtAmt(grandCogs)}</td></tr>`;

        let grandGP = 0;
        html += '<tr class="pnl-total"><td>Total Gross Profit</td>';
        months.forEach(m => {
            const gp = ((byMonth[m] && byMonth[m].revenue) || 0) - ((byMonth[m] && byMonth[m].cogs) || 0);
            grandGP += gp;
            html += `<td>${fmtAmt(gp)}</td>`;
        });
        html += `<td>${fmtAmt(grandGP)}</td></tr>`;

        html += '</tbody></table>';
        container.innerHTML = html;
    },

    // ==================== PRODUCTS TAB ====================

    renderProductsTab(products, showDiscontinued, analytics) {
        const fmtAmt = (amt) => Utils.formatCurrency(amt);
        analytics = analytics || { totals: { pretax_total: 0, sales_tax: 0, post_tax_total: 0, discount: 0, pretax_after_discount: 0 }, byProduct: [], linkedProductIds: new Set() };
        const linkedIds = analytics.linkedProductIds || new Set();

        // Filter products based on discontinued toggle
        const visible = showDiscontinued ? products : products.filter(p => !p.is_discontinued);
        const activeProducts = products.filter(p => !p.is_discontinued);

        // Summary cards
        const summaryEl = document.getElementById('productSummaryCards');
        if (products.length === 0) {
            summaryEl.innerHTML = '';
        } else {
            const avgPrice = activeProducts.length > 0
                ? activeProducts.reduce((sum, p) => sum + p.price, 0) / activeProducts.length : 0;
            const avgMarginPct = activeProducts.length > 0
                ? activeProducts.reduce((sum, p) => sum + (p.price > 0 ? ((p.price - p.cogs) / p.price) * 100 : 0), 0) / activeProducts.length : 0;

            const hasAnalytics = analytics.byProduct.length > 0;
            const discountCards = hasAnalytics && analytics.totals.discount > 0 ? `
                <div class="budget-summary-card"><span class="budget-summary-label">VE Pretax After Discounts</span><span class="budget-summary-value">${fmtAmt(analytics.totals.pretax_after_discount)}</span></div>
            ` : '';
            const veCards = hasAnalytics ? `
                <div class="budget-summary-card"><span class="budget-summary-label">VE Pretax Revenue</span><span class="budget-summary-value">${fmtAmt(analytics.totals.pretax_total)}</span></div>
                <div class="budget-summary-card"><span class="budget-summary-label">VE Sales Tax</span><span class="budget-summary-value">${fmtAmt(analytics.totals.sales_tax)}</span></div>
                <div class="budget-summary-card"><span class="budget-summary-label">VE Post-Tax Total</span><span class="budget-summary-value">${fmtAmt(analytics.totals.post_tax_total)}</span></div>
                ${discountCards}
            ` : '';

            summaryEl.innerHTML = `
                <div class="budget-summary-card"><span class="budget-summary-label">Total Products</span><span class="budget-summary-value">${products.length}</span></div>
                <div class="budget-summary-card"><span class="budget-summary-label">Active</span><span class="budget-summary-value">${activeProducts.length}</span></div>
                <div class="budget-summary-card"><span class="budget-summary-label">Avg Price</span><span class="budget-summary-value">${fmtAmt(avgPrice)}</span></div>
                <div class="budget-summary-card"><span class="budget-summary-label">Avg Margin</span><span class="budget-summary-value ${avgMarginPct >= 0 ? 'pc-margin-positive' : 'pc-margin-negative'}">${avgMarginPct.toFixed(1)}%</span></div>
                ${veCards}
            `;
        }

        // Product table
        const wrapper = document.getElementById('productTableWrapper');
        if (visible.length === 0) {
            wrapper.innerHTML = products.length === 0
                ? '<p class="empty-state">No products yet. Click "+ Add Product" to get started.</p>'
                : '<p class="empty-state">No active products to show. Toggle "Show Discontinued" to see all.</p>';
            return;
        }

        let html = `<table class="pc-table">
            <thead><tr>
                <th>Name</th>
                <th>SKU</th>
                <th class="pc-col-num">Price</th>
                <th class="pc-col-num">COGS</th>
                <th class="pc-col-num">Margin $</th>
                <th class="pc-col-num">Margin %</th>
                <th class="pc-col-actions">Actions</th>
            </tr></thead><tbody>`;

        visible.forEach(p => {
            const margin = p.price - (p.cogs || 0);
            const marginPct = p.price > 0 ? (margin / p.price) * 100 : 0;
            const marginClass = margin >= 0 ? 'pc-margin-positive' : 'pc-margin-negative';
            const isLinked = linkedIds.has(p.id);
            const rowClass = (p.is_discontinued ? 'pc-row-discontinued' : '') + (isLinked ? ' pc-row-linked' : '');
            const toggleTitle = p.is_discontinued ? 'Reactivate' : 'Discontinue';
            const toggleIcon = p.is_discontinued
                ? `<svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M9 12l2 2 4-4"/><circle cx="12" cy="12" r="10"/></svg>`
                : `<svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="5" y1="12" x2="19" y2="12"/></svg>`;
            const toggleClass = p.is_discontinued ? 'pc-icon-btn pc-icon-btn-reactivate' : 'pc-icon-btn pc-icon-btn-discontinue';
            const linkedBadge = linkedIds.has(p.id) ? '<span class="pc-linked-badge">linked</span>' : '';
            const notesIcon = p.notes ? `<span class="notes-indicator pc-notes-indicator" data-notes="${Utils.escapeHtml(p.notes)}"><svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path></svg></span>` : '';

            html += `<tr class="${rowClass}">
                <td class="pc-col-name"><span class="pc-name-text">${Utils.escapeHtml(p.name)}</span>${linkedBadge}${notesIcon}</td>
                <td class="pc-col-sku">${p.sku ? Utils.escapeHtml(p.sku) : '<span style="color:var(--color-text-muted)">—</span>'}</td>
                <td class="pc-col-num">${fmtAmt(p.price)}</td>
                <td class="pc-col-num">${fmtAmt(p.cogs || 0)}</td>
                <td class="pc-col-num ${marginClass}">${fmtAmt(margin)}</td>
                <td class="pc-col-num ${marginClass}">${marginPct.toFixed(1)}%</td>
                <td class="pc-col-actions">
                    <div class="pc-actions">
                        <button class="btn btn-primary btn-small manage-links-btn" data-id="${p.id}">Links</button>
                        <button class="btn btn-secondary btn-small edit-product-btn" data-id="${p.id}">Edit</button>
                        <button class="${toggleClass} discontinue-product-btn" data-id="${p.id}" title="${toggleTitle}">${toggleIcon}</button>
                        <button class="pc-icon-btn pc-icon-btn-delete delete-product-btn" data-id="${p.id}" title="Delete"><svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6M14 11v6"/></svg></button>
                    </div>
                </td>
            </tr>`;
        });

        html += '</tbody></table>';
        wrapper.innerHTML = html;
    },

    // ==================== B2B CONTRACT TAB ====================

    /**
     * Render the B2B Contract tab with list/detail layout
     * @param {Array} contracts - Array of contract objects
     * @param {number|null} selectedId - Currently selected contract ID
     * @param {Object} contractTransactions - Map of contractId -> array of receivable transaction objects
     * @param {Object} contractCogsTransactions - Map of contractId -> array of COGS transaction objects
     * @param {Object} app - App reference for computing values
     */
    renderB2BContractTab(contracts, selectedId, contractTransactions, contractCogsTransactions, app) {
        const fmtAmt = (amt) => Utils.formatCurrency(amt);

        // Left panel: contract list
        const listPanel = document.getElementById('b2bListPanel');
        if (contracts.length === 0) {
            listPanel.innerHTML = '<p class="empty-state">No contracts yet. Click "+ Add Contract" to begin.</p>';
        } else {
            listPanel.innerHTML = contracts.map(contract => {
                const selected = contract.id === selectedId ? ' selected' : '';
                const statusBadge = contract.is_finalized
                    ? '<span class="b2b-status-badge b2b-finalized">Finalized</span>'
                    : '<span class="b2b-status-badge b2b-draft">Draft</span>';
                return `<div class="b2b-list-item${selected}" data-id="${contract.id}">
                    <div class="b2b-list-name">${Utils.escapeHtml(contract.company_name)}</div>
                    <div class="b2b-list-meta">${contract.contract_start} to ${contract.contract_end} ${statusBadge}</div>
                </div>`;
            }).join('');
        }

        // Right panel: selected contract detail
        const detailPanel = document.getElementById('b2bDetailPanel');
        const selectedContract = contracts.find(c => c.id === selectedId);
        if (!selectedContract) {
            detailPanel.innerHTML = `<div class="empty-state">
                <svg class="empty-state-icon" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                    <polyline points="14 2 14 8 20 8"/>
                    <line x1="16" y1="13" x2="8" y2="13"/>
                    <line x1="16" y1="17" x2="8" y2="17"/>
                </svg>
                <p class="empty-state-title">No Contract Selected</p>
                <p class="empty-state-desc">Select a contract to view its details and payment timeline, or add a new one.</p>
            </div>`;
            return;
        }

        const computed = app._computeB2BContract(selectedContract);
        const transactions = contractTransactions[selectedContract.id] || [];
        const cogsTransactions = contractCogsTransactions[selectedContract.id] || [];
        const isDirect = selectedContract.entry_mode === 'direct';
        const monthlyRevenue = isDirect ? Math.round((selectedContract.direct_revenue || 0) * 100) / 100 : computed.monthlyContractedRounded;
        const monthlyCogs = isDirect ? Math.round((selectedContract.direct_cogs || 0) * 100) / 100 : Math.round(computed.cogsPerUnit * computed.unitsSold * 100) / 100;
        const limitClass = computed.isWithinLimit ? 'b2b-limit-ok' : 'b2b-limit-exceeded';
        const limitLabel = computed.isWithinLimit ? 'Within 75% Limit' : 'Exceeds 75% Limit';

        let html = '';

        // Header with actions
        html += `<div class="b2b-detail-header">
            <div>
                <h3>${Utils.escapeHtml(selectedContract.company_name)}</h3>
                ${selectedContract.bundle_description ? `<p class="b2b-bundle-desc">${Utils.escapeHtml(selectedContract.bundle_description)}</p>` : ''}
                <span class="b2b-detail-dates">${selectedContract.contract_start} to ${selectedContract.contract_end} (${selectedContract.fiscal_months} months)</span>
            </div>
            <div class="b2b-detail-actions">
                <button class="btn btn-secondary btn-small b2b-edit-btn" data-id="${selectedContract.id}">Edit</button>
                <button class="btn btn-danger btn-small b2b-delete-btn" data-id="${selectedContract.id}">Delete</button>
            </div>
        </div>`;

        // Summary cards
        const monthlyGrossProfit = Math.round((monthlyRevenue - monthlyCogs) * 100) / 100;
        const totalContractValue = Math.round(monthlyRevenue * selectedContract.fiscal_months * 100) / 100;
        const totalGrossProfit = Math.round(monthlyGrossProfit * selectedContract.fiscal_months * 100) / 100;

        html += `<div class="b2b-summary-cards">
            <div class="b2b-summary-card">
                <div class="b2b-summary-label">Total Contract Value</div>
                <div class="b2b-summary-value">${fmtAmt(totalContractValue)}</div>
            </div>
            <div class="b2b-summary-card">
                <div class="b2b-summary-label">Monthly Revenue</div>
                <div class="b2b-summary-value">${fmtAmt(monthlyRevenue)}</div>
            </div>
            ${isDirect ? `
            <div class="b2b-summary-card">
                <div class="b2b-summary-label">Entry Mode</div>
                <div class="b2b-summary-value" style="font-size:var(--font-sm)">Direct</div>
            </div>` : `
            <div class="b2b-summary-card">
                <div class="b2b-summary-label">Units / Month</div>
                <div class="b2b-summary-value">${computed.unitsSold.toLocaleString()} <small style="font-weight:400;color:var(--text-muted)">/ ${computed.maxUnitsPerMonth.toLocaleString()} max</small></div>
            </div>`}
            <div class="b2b-summary-card">
                <div class="b2b-summary-label">Monthly COGS</div>
                <div class="b2b-summary-value">${fmtAmt(monthlyCogs)}</div>
            </div>
            <div class="b2b-summary-card">
                <div class="b2b-summary-label">Monthly Gross Profit</div>
                <div class="b2b-summary-value">${fmtAmt(monthlyGrossProfit)}</div>
            </div>
            <div class="b2b-summary-card">
                <div class="b2b-summary-label">Total Gross Profit</div>
                <div class="b2b-summary-value">${fmtAmt(totalGrossProfit)}</div>
            </div>
            <div class="b2b-summary-card">
                <div class="b2b-summary-label">Profit Limit Status</div>
                <div class="b2b-summary-value ${limitClass}">${limitLabel}</div>
            </div>
        </div>`;

        // Product lines table (only for products mode)
        const products = !isDirect ? Database.getB2BContractProducts(selectedContract.id) : [];
        if (products.length > 0) {
            let totalQty = 0, monthlyTotal = 0, weightedPmSum = 0, weightedPmDenom = 0;
            html += `<div class="b2b-calc-breakdown">
                <h4>Product Lines</h4>
                <table class="b2b-breakdown-table b2b-products-detail-table">
                    <thead><tr><th>Product</th><th>B2B Price</th><th>COGS</th><th>PM%</th><th>Qty</th><th>Total</th></tr></thead>
                    <tbody>`;
            products.forEach(p => {
                const pm = p.b2b_price > 0 ? ((p.b2b_price - p.cogs) / p.b2b_price * 100) : 0;
                const total = Math.round(p.b2b_price * p.quantity * 100) / 100;
                totalQty += p.quantity;
                monthlyTotal += total;
                if (p.b2b_price > 0 && p.quantity > 0) {
                    weightedPmSum += pm * total;
                    weightedPmDenom += total;
                }
                html += `<tr>
                    <td>${Utils.escapeHtml(p.product_name)}</td>
                    <td>${fmtAmt(p.b2b_price)}</td>
                    <td>${fmtAmt(p.cogs)}</td>
                    <td>${pm.toFixed(2)}%</td>
                    <td>${p.quantity.toLocaleString()}</td>
                    <td>${fmtAmt(total)}</td>
                </tr>`;
            });
            const avgPm = weightedPmDenom > 0 ? (weightedPmSum / weightedPmDenom) : 0;
            html += `</tbody>
                <tfoot><tr class="b2b-row-highlight">
                    <td><strong>Totals</strong></td><td></td><td></td>
                    <td>${avgPm.toFixed(2)}%</td>
                    <td>${totalQty.toLocaleString()}</td>
                    <td>${fmtAmt(monthlyTotal)}</td>
                </tr></tfoot></table></div>`;
        }

        // Calculation breakdown
        const grossMarginPct = monthlyRevenue > 0 ? ((monthlyRevenue - monthlyCogs) / monthlyRevenue * 100) : 0;
        html += `<div class="b2b-calc-breakdown">
            <h4>Calculation Breakdown</h4>
            <table class="b2b-breakdown-table">
                <tbody>
                    <tr><td>Monthly Payroll</td><td>${fmtAmt(selectedContract.monthly_payroll)}</td></tr>
                    <tr><td>Total Gross Payroll (${selectedContract.fiscal_months} months)</td><td>${fmtAmt(computed.totalGrossPayroll)}</td></tr>
                    <tr class="b2b-row-highlight"><td>Max 75% of Payroll</td><td>${fmtAmt(computed.maxAllowedSales)}</td></tr>
                    <tr><td>Gross Margin</td><td>${grossMarginPct.toFixed(2)}%</td></tr>
                    <tr><td>Monthly Revenue</td><td>${fmtAmt(monthlyRevenue)}</td></tr>
                    <tr><td>Monthly COGS</td><td>${fmtAmt(monthlyCogs)}</td></tr>
                    <tr><td>Monthly Gross Profit</td><td>${fmtAmt(monthlyGrossProfit)}</td></tr>
                    <tr class="b2b-row-highlight"><td>Total Contract Value</td><td>${fmtAmt(totalContractValue)}</td></tr>
                    <tr><td>Total Gross Profit</td><td class="${limitClass}">${fmtAmt(totalGrossProfit)}</td></tr>
                </tbody>
            </table>
        </div>`;

        // Finalize / Unfinalize button
        if (selectedContract.is_finalized) {
            html += `<div class="b2b-finalize-section">
                <span class="b2b-status-badge b2b-finalized">Finalized</span>
                <button class="btn btn-secondary btn-small b2b-unfinalize-btn" data-id="${selectedContract.id}">Unfinalize</button>
            </div>`;
        } else {
            html += `<div class="b2b-finalize-section">
                <button class="btn btn-primary b2b-finalize-btn" data-id="${selectedContract.id}">Finalize Contract</button>
                <small>Creates monthly receivable entries in the journal</small>
            </div>`;
        }

        // Monthly timeline
        if (selectedContract.is_finalized && computed.months.length > 0) {
            const currentMonth = Utils.getCurrentMonth();
            html += `<div class="b2b-timeline">
                <h4>Monthly Payment Timeline</h4>
                <table class="b2b-timeline-table">
                    <thead>
                        <tr><th>Month</th><th>Revenue</th><th>Status</th><th>COGS</th><th>Status</th></tr>
                    </thead>
                    <tbody>`;

            computed.months.forEach(month => {
                const isFuture = month > currentMonth;
                const monthLabel = Utils.formatMonthDisplay(month) || month;

                // Revenue
                const txn = transactions.find(t => t.payment_for_month === month);
                const revAmount = txn ? txn.amount : computed.monthlyContractedRounded;
                const revStatus = isFuture ? 'future' : (txn ? txn.status : 'pending');
                const revStatusClass = revStatus === 'received' ? 'b2b-status-received' : (revStatus === 'future' ? 'b2b-status-future' : 'b2b-status-pending');
                const revStatusLabel = revStatus === 'received' ? 'Received' : (revStatus === 'future' ? 'Future' : 'Pending');

                // COGS
                const cogsTxn = cogsTransactions.find(t => t.payment_for_month === month);
                const cogsAmount = cogsTxn ? cogsTxn.amount : monthlyCogs;
                const cogsStatus = isFuture ? 'future' : (cogsTxn ? cogsTxn.status : 'pending');
                const cogsStatusClass = cogsStatus === 'paid' ? 'b2b-status-received' : (cogsStatus === 'future' ? 'b2b-status-future' : 'b2b-status-pending');
                const cogsStatusLabel = cogsStatus === 'paid' ? 'Paid' : (cogsStatus === 'future' ? 'Future' : 'Pending');

                const rowClass = isFuture ? ' class="b2b-row-future"' : '';
                html += `<tr${rowClass}>
                    <td>${monthLabel}</td>
                    <td>${fmtAmt(revAmount)}</td>
                    <td><span class="b2b-payment-status ${revStatusClass}">${revStatusLabel}</span></td>
                    <td>${fmtAmt(cogsAmount)}</td>
                    <td><span class="b2b-payment-status ${cogsStatusClass}">${cogsStatusLabel}</span></td>
                </tr>`;
            });

            html += `</tbody></table></div>`;
        }

        // Notes
        if (selectedContract.notes) {
            html += `<div class="b2b-notes"><strong>Notes:</strong> ${Utils.escapeHtml(selectedContract.notes)}</div>`;
        }

        detailPanel.innerHTML = html;
    },

};

// Export for use in other modules
window.UI = UI;
