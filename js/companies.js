/**
 * CompanyManager — Multi-company support for the Accounting Journal
 *
 * Stores up to 5 company databases in IndexedDB, each under its own key.
 * A registry document tracks which companies exist and which is active.
 *
 * IDB layout (store: 'database' in DB 'AccountingJournalDB'):
 *   'companyRegistry'  → plain object (structured-clone)
 *   'co_<id>'          → Uint8Array (SQLite binary for that company)
 *
 * The legacy 'sqliteDb' key is migrated to 'co_<id>' on first run.
 */

const CompanyManager = {
    _registry: null,      // CompanyRegistry loaded at init
    _activeKey: null,     // IDB key for the active company's SQLite bytes
    _needsNamingPrompt: false,

    IDB_NAME: 'AccountingJournalDB',
    IDB_STORE: 'database',
    REGISTRY_KEY: 'companyRegistry',
    MAX_COMPANIES: 5,

    // ==================== INIT ====================

    async init() {
        const registry = await this._readIDB(this.REGISTRY_KEY);
        if (!registry) {
            await this.bootstrap();
        } else {
            this._registry = registry;
            this._activeKey = this._registry.activeId;
        }
    },

    /**
     * First-ever run: migrate legacy 'sqliteDb' data (if any) into a new company slot.
     */
    async bootstrap() {
        const legacyBytes = await this._readIDB('sqliteDb');
        const id = this._generateId();
        const now = new Date().toISOString();

        this._registry = {
            version: 1,
            activeId: id,
            companies: [{ id, name: 'My Company', createdAt: now, lastAccessedAt: now }]
        };
        this._activeKey = id;

        if (legacyBytes) {
            await this._writeIDB(id, legacyBytes);
        }
        // If no legacy bytes: new blank DB — Database.init() handles this via loadFromIndexedDB returning null

        await this._persistRegistry();
        this._needsNamingPrompt = true;
    },

    // ==================== ACCESSORS ====================

    getActiveKey() {
        return this._activeKey;
    },

    getRegistry() {
        return this._registry;
    },

    getActiveCompany() {
        if (!this._registry) return null;
        return this._registry.companies.find(c => c.id === this._registry.activeId) || null;
    },

    getAll() {
        if (!this._registry) return [];
        return [...this._registry.companies].sort(
            (a, b) => new Date(b.lastAccessedAt) - new Date(a.lastAccessedAt)
        );
    },

    needsNamingPrompt() {
        return this._needsNamingPrompt;
    },

    clearNamingPrompt() {
        this._needsNamingPrompt = false;
    },

    // ==================== OPERATIONS ====================

    /**
     * Switch to another company. Saves current company first, then swaps DB.
     */
    async switchTo(companyId) {
        // Save current company under its own key (bypass debounce for reliability)
        const currentData = Database.db.export();
        await this._writeIDB(this._activeKey, new Uint8Array(currentData));

        // Update registry
        this._registry.activeId = companyId;
        const comp = this._registry.companies.find(c => c.id === companyId);
        comp.lastAccessedAt = new Date().toISOString();
        await this._persistRegistry();

        // Swap key and load new DB
        this._activeKey = companyId;
        const bytes = await this._readIDB(companyId);
        Database.loadBytes(bytes);
    },

    /**
     * Create a new blank company. Does NOT switch to it — caller does that.
     */
    async createNew(name) {
        if (this._registry.companies.length >= this.MAX_COMPANIES) {
            throw new Error('MAX_COMPANIES');
        }
        const id = this._generateId();
        const now = new Date().toISOString();
        this._registry.companies.push({ id, name, createdAt: now, lastAccessedAt: now });
        await this._persistRegistry();
        return id;
        // No bytes written — switchTo(id) calls loadBytes(null) which creates fresh schema
    },

    /**
     * Rename any company.
     */
    async rename(companyId, name) {
        const comp = this._registry.companies.find(c => c.id === companyId);
        if (comp) {
            comp.name = name;
            await this._persistRegistry();
        }
    },

    /**
     * Delete a company. If it was active, auto-switch to the most recent other company.
     * Returns updated registry so caller knows new active state.
     */
    async delete(companyId) {
        const isActive = this._registry.activeId === companyId;
        this._registry.companies = this._registry.companies.filter(c => c.id !== companyId);

        if (isActive) {
            const remaining = [...this._registry.companies].sort(
                (a, b) => new Date(b.lastAccessedAt) - new Date(a.lastAccessedAt)
            );
            if (remaining.length > 0) {
                this._registry.activeId = remaining[0].id;
                this._activeKey = remaining[0].id;
            } else {
                this._registry.activeId = null;
                this._activeKey = null;
            }
        }

        // Persist registry first, THEN delete bytes — so on failure the registry still points to a valid key
        await this._persistRegistry();
        await this._deleteIDB(companyId);
        return this._registry;
    },

    /**
     * Import an ArrayBuffer as a new company (does not switch to it).
     */
    async importAsNew(buffer, name) {
        if (this._registry.companies.length >= this.MAX_COMPANIES) {
            throw new Error('MAX_COMPANIES');
        }
        const id = this._generateId();
        const now = new Date().toISOString();
        this._registry.companies.push({ id, name, createdAt: now, lastAccessedAt: now });
        await this._writeIDB(id, new Uint8Array(buffer));
        await this._persistRegistry();
        return id;
    },

    /**
     * Replace active company's data with an imported ArrayBuffer.
     */
    async replaceActive(buffer) {
        const bytes = new Uint8Array(buffer);
        await this._writeIDB(this._activeKey, bytes);
        Database.loadBytes(bytes);
    },

    /**
     * Copy a section of data from sourceCompanyId into the active company.
     * section: 'breakeven' | 'loans' | 'budget' | 'products' | 'assets'
     */
    async copySection(sourceCompanyId, section) {
        const sourceBytes = await this._readIDB(sourceCompanyId);
        if (!sourceBytes) return { copied: 0 };

        const tempDb = new Database.SQL.Database(sourceBytes);
        let copied = 0;

        try {
            if (section === 'breakeven') {
                const r = tempDb.exec("SELECT value FROM app_meta WHERE key = 'breakeven_config'");
                if (r.length && r[0].values.length) {
                    Database.db.run(
                        "INSERT OR REPLACE INTO app_meta (key, value) VALUES ('breakeven_config', ?)",
                        [r[0].values[0][0]]
                    );
                    copied = 1;
                }
            } else if (section === 'loans') {
                Database.db.run('DELETE FROM loan_payment_overrides');
                Database.db.run('DELETE FROM loan_skipped_payments');
                Database.db.run('DELETE FROM loans');
                const rows = tempDb.exec(
                    'SELECT name, principal, annual_rate, term_months, payments_per_year, start_date, first_payment_date, notes, is_active FROM loans'
                );
                if (rows.length) {
                    for (const row of rows[0].values) {
                        Database.db.run(
                            'INSERT INTO loans (name, principal, annual_rate, term_months, payments_per_year, start_date, first_payment_date, notes, is_active) VALUES (?,?,?,?,?,?,?,?,?)',
                            row
                        );
                        copied++;
                    }
                }
            } else if (section === 'budget') {
                Database.db.run('DELETE FROM budget_expenses');
                const rows = tempDb.exec(
                    'SELECT name, monthly_amount, start_month, end_month, notes FROM budget_expenses'
                );
                if (rows.length) {
                    for (const row of rows[0].values) {
                        Database.db.run(
                            'INSERT INTO budget_expenses (name, monthly_amount, start_month, end_month, notes) VALUES (?,?,?,?,?)',
                            row
                        );
                        copied++;
                    }
                }
            } else if (section === 'products') {
                Database.db.run('DELETE FROM products');
                const rows = tempDb.exec(
                    'SELECT name, sku, price, tax_rate, cogs, notes, is_discontinued FROM products'
                );
                if (rows.length) {
                    for (const row of rows[0].values) {
                        Database.db.run(
                            'INSERT INTO products (name, sku, price, tax_rate, cogs, notes, is_discontinued) VALUES (?,?,?,?,?,?,?)',
                            row
                        );
                        copied++;
                    }
                }
            } else if (section === 'assets') {
                Database.db.run('DELETE FROM balance_sheet_assets');
                const rows = tempDb.exec(
                    'SELECT name, purchase_cost, useful_life_months, purchase_date, salvage_value, depreciation_method, dep_start_date, is_depreciable, notes FROM balance_sheet_assets'
                );
                if (rows.length) {
                    for (const row of rows[0].values) {
                        Database.db.run(
                            'INSERT INTO balance_sheet_assets (name, purchase_cost, useful_life_months, purchase_date, salvage_value, depreciation_method, dep_start_date, is_depreciable, notes) VALUES (?,?,?,?,?,?,?,?,?)',
                            row
                        );
                        copied++;
                    }
                }
            }
        } finally {
            tempDb.close();
        }

        Database.autoSave();
        return { copied };
    },

    // ==================== UI RENDERING ====================

    /**
     * Rebuild the company switcher button label and popover list.
     */
    renderSwitcher() {
        const active = this.getActiveCompany();
        const all = this.getAll();

        // Update button label
        const label = document.getElementById('companyBtnLabel');
        if (label) label.textContent = active ? active.name : 'Companies';

        // Rebuild list
        const list = document.getElementById('companyList');
        if (!list) return;
        list.innerHTML = '';

        all.forEach(company => {
            const li = document.createElement('li');
            li.className = 'company-list-item' + (company.id === this._registry.activeId ? ' active' : '');
            li.dataset.id = company.id;

            const nameSpan = document.createElement('span');
            nameSpan.className = 'company-list-name';
            nameSpan.textContent = company.name;

            const metaSpan = document.createElement('span');
            metaSpan.className = 'company-list-meta';
            metaSpan.textContent = this._formatRelativeTime(company.lastAccessedAt);

            li.appendChild(nameSpan);
            li.appendChild(metaSpan);
            list.appendChild(li);
        });

        // Show/hide add button based on limit
        const addBtn = document.getElementById('addCompanyBtn');
        if (addBtn) {
            addBtn.style.display = all.length >= this.MAX_COMPANIES ? 'none' : '';
        }

        // Update copy-from selects in manage modal
        this._updateCopyFromSelect();
    },

    /**
     * Populate the manage companies table body.
     */
    renderManageTable() {
        const tbody = document.getElementById('companiesTableBody');
        if (!tbody) return;
        tbody.innerHTML = '';

        const all = this.getAll();
        all.forEach(company => {
            const isActive = company.id === this._registry.activeId;
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>
                    <span class="company-table-name" data-id="${company.id}">${this._escHtml(company.name)}</span>
                    ${isActive ? '<span class="company-active-badge">active</span>' : ''}
                </td>
                <td class="company-table-date">${this._formatDate(company.createdAt)}</td>
                <td class="company-table-date">${this._formatRelativeTime(company.lastAccessedAt)}</td>
                <td>
                    <button class="btn btn-sm btn-secondary company-rename-btn" data-id="${company.id}">Rename</button>
                    <button class="btn btn-sm btn-danger company-delete-btn" data-id="${company.id}" ${isActive && all.length === 1 ? 'disabled' : ''}>Delete</button>
                </td>
            `;
            tbody.appendChild(tr);
        });
    },

    // ==================== PRIVATE HELPERS ====================

    _generateId() {
        const rand = Math.random().toString(36).slice(2, 7);
        return 'co_' + Date.now() + '_' + rand;
    },

    async _openIDB() {
        return new Promise((resolve, reject) => {
            const req = indexedDB.open(this.IDB_NAME, 1);
            req.onerror = () => reject(req.error);
            req.onupgradeneeded = (e) => {
                const db = e.target.result;
                if (!db.objectStoreNames.contains(this.IDB_STORE)) {
                    db.createObjectStore(this.IDB_STORE);
                }
            };
            req.onsuccess = (e) => resolve(e.target.result);
        });
    },

    async _readIDB(key) {
        const db = await this._openIDB();
        return new Promise((resolve, reject) => {
            const tx = db.transaction([this.IDB_STORE], 'readonly');
            const store = tx.objectStore(this.IDB_STORE);
            const req = store.get(key);
            req.onsuccess = () => resolve(req.result || null);
            req.onerror = () => reject(req.error);
        });
    },

    async _writeIDB(key, data) {
        const db = await this._openIDB();
        return new Promise((resolve, reject) => {
            const tx = db.transaction([this.IDB_STORE], 'readwrite');
            const store = tx.objectStore(this.IDB_STORE);
            store.put(data, key);
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        });
    },

    async _deleteIDB(key) {
        const db = await this._openIDB();
        return new Promise((resolve, reject) => {
            const tx = db.transaction([this.IDB_STORE], 'readwrite');
            const store = tx.objectStore(this.IDB_STORE);
            store.delete(key);
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        });
    },

    async _persistRegistry() {
        await this._writeIDB(this.REGISTRY_KEY, this._registry);
    },

    _updateCopyFromSelect() {
        const sel = document.getElementById('copyFromCompanySelect');
        if (!sel) return;
        const active = this._registry.activeId;
        const others = this._registry.companies.filter(c => c.id !== active);
        sel.innerHTML = '';
        if (others.length === 0) {
            sel.innerHTML = '<option value="">No other companies</option>';
            return;
        }
        others.forEach(c => {
            const opt = document.createElement('option');
            opt.value = c.id;
            opt.textContent = c.name;
            sel.appendChild(opt);
        });
    },

    _formatDate(iso) {
        if (!iso) return '—';
        const d = new Date(iso);
        return d.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' });
    },

    _formatRelativeTime(iso) {
        if (!iso) return '—';
        const diffMs = Date.now() - new Date(iso).getTime();
        const diffMin = Math.floor(diffMs / 60000);
        if (diffMin < 2) return 'just now';
        if (diffMin < 60) return `${diffMin}m ago`;
        const diffHr = Math.floor(diffMin / 60);
        if (diffHr < 24) return `${diffHr}h ago`;
        const diffDay = Math.floor(diffHr / 24);
        if (diffDay === 1) return 'yesterday';
        if (diffDay < 30) return `${diffDay}d ago`;
        return this._formatDate(iso);
    },

    _escHtml(str) {
        return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    },
};
