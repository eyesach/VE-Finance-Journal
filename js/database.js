/**
 * Database module for SQLite operations using sql.js
 * Includes IndexedDB auto-save functionality
 */

const Database = {
    db: null,
    SQL: null,
    IDB_NAME: 'AccountingJournalDB',
    IDB_STORE: 'database',
    get IDB_KEY() {
        const key = CompanyManager.getActiveKey();
        if (!key) throw new Error('CompanyManager not initialized — no active company key');
        return key;
    },

    /**
     * Initialize the database
     * @returns {Promise<void>}
     */
    async init() {
        this.SQL = await initSqlJs({
            locateFile: file => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.8.0/${file}`
        });

        await CompanyManager.init();

        const savedData = await this.loadFromIndexedDB();

        if (savedData) {
            this.db = new this.SQL.Database(savedData);
            this.migrateSchema();
            console.log('Database loaded from IndexedDB');
        } else {
            this.db = new this.SQL.Database();
            this.createSchema();
            console.log('New database created');
        }
    },

    /**
     * Swap the active in-memory database with the given bytes.
     * Called by CompanyManager.switchTo() and replaceActive().
     * @param {Uint8Array|null} bytes - SQLite binary, or null for a fresh blank DB
     */
    loadBytes(bytes) {
        if (this.db) this.db.close();
        if (bytes) {
            this.db = new this.SQL.Database(bytes);
            this.migrateSchema();
        } else {
            this.db = new this.SQL.Database();
            this.createSchema();
        }
    },

    /**
     * Reset all data by recreating the database from scratch
     */
    resetAllData() {
        this.db.close();
        this.db = new this.SQL.Database();
        this.createSchema();
        this.autoSave();
    },

    /**
     * Create database schema
     */
    createSchema() {
        this.db.run(`
            CREATE TABLE IF NOT EXISTS category_folders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                folder_type TEXT NOT NULL DEFAULT 'payable',
                sort_order INTEGER DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS categories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                type TEXT DEFAULT 'both',
                is_monthly INTEGER DEFAULT 0,
                default_amount DECIMAL(10,2),
                default_type TEXT,
                folder_id INTEGER,
                cashflow_sort_order INTEGER DEFAULT 0,
                show_on_pl INTEGER DEFAULT 0,
                is_cogs INTEGER DEFAULT 0,
                is_depreciation INTEGER DEFAULT 0,
                is_sales_tax INTEGER DEFAULT 0,
                is_b2b INTEGER DEFAULT 0,
                is_sales INTEGER DEFAULT 0,
                is_inventory_cost INTEGER DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (folder_id) REFERENCES category_folders(id)
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entry_date DATE NOT NULL,
                category_id INTEGER NOT NULL,
                item_description TEXT,
                amount DECIMAL(10,2) NOT NULL,
                pretax_amount DECIMAL(10,2),
                transaction_type TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'pending',
                date_processed DATE,
                month_due TEXT,
                month_paid TEXT,
                payment_for_month TEXT,
                notes TEXT,
                source_type TEXT,
                source_id INTEGER,
                sale_date_start DATE,
                sale_date_end DATE,
                inventory_cost DECIMAL(10,2),
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (category_id) REFERENCES categories(id)
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS pl_overrides (
                category_id INTEGER,
                month TEXT,
                override_amount DECIMAL(10,2),
                PRIMARY KEY(category_id, month)
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS app_meta (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS cashflow_overrides (
                category_id INTEGER,
                month TEXT,
                override_amount DECIMAL(10,2),
                PRIMARY KEY(category_id, month)
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS balance_sheet_assets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                purchase_cost DECIMAL(10,2) NOT NULL,
                useful_life_months INTEGER NOT NULL,
                purchase_date DATE NOT NULL,
                salvage_value DECIMAL(10,2) DEFAULT 0,
                depreciation_method TEXT DEFAULT 'straight_line',
                dep_start_date DATE,
                is_depreciable INTEGER DEFAULT 1,
                linked_transaction_id INTEGER,
                notes TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS loans (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                principal DECIMAL(10,2) NOT NULL,
                annual_rate DECIMAL(8,4) NOT NULL,
                term_months INTEGER NOT NULL,
                payments_per_year INTEGER NOT NULL DEFAULT 12,
                start_date DATE NOT NULL,
                first_payment_date DATE,
                notes TEXT,
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS loan_skipped_payments (
                loan_id INTEGER NOT NULL,
                payment_number INTEGER NOT NULL,
                PRIMARY KEY(loan_id, payment_number)
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS loan_payment_overrides (
                loan_id INTEGER NOT NULL,
                payment_number INTEGER NOT NULL,
                override_amount DECIMAL(10,2) NOT NULL,
                PRIMARY KEY(loan_id, payment_number)
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS budget_groups (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                sort_order INTEGER DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS budget_expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                monthly_amount DECIMAL(10,2) NOT NULL,
                start_month TEXT NOT NULL,
                end_month TEXT,
                category_id INTEGER,
                group_id INTEGER,
                notes TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (category_id) REFERENCES categories(id),
                FOREIGN KEY (group_id) REFERENCES budget_groups(id)
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                sku TEXT,
                price DECIMAL(10,2) NOT NULL,
                tax_rate DECIMAL(6,4) DEFAULT 0,
                cogs DECIMAL(10,2) DEFAULT 0,
                notes TEXT,
                is_discontinued INTEGER DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS ve_sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                transaction_no TEXT NOT NULL,
                date DATE NOT NULL,
                billing_name TEXT,
                description TEXT,
                subtotal DECIMAL(10,2) DEFAULT 0,
                tax DECIMAL(10,2) DEFAULT 0,
                shipping DECIMAL(10,2) DEFAULT 0,
                discount DECIMAL(10,2) DEFAULT 0,
                total DECIMAL(10,2) DEFAULT 0,
                source TEXT NOT NULL DEFAULT 'online',
                event_id INTEGER,
                UNIQUE(transaction_no, source)
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS ve_sale_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                transaction_no TEXT NOT NULL,
                name TEXT NOT NULL,
                product_number TEXT,
                price DECIMAL(10,2) DEFAULT 0,
                quantity INTEGER DEFAULT 1,
                taxable INTEGER DEFAULT 0,
                amount DECIMAL(10,2) DEFAULT 0,
                inferred INTEGER DEFAULT 0
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS ve_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                type TEXT NOT NULL DEFAULT 'tradeshow',
                start_date DATE NOT NULL,
                end_date DATE NOT NULL,
                notes TEXT,
                journal_added INTEGER DEFAULT 0
            )
        `);

        this.db.run(`
            CREATE TABLE IF NOT EXISTS product_ve_mappings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER NOT NULL,
                ve_item_name TEXT NOT NULL,
                ve_item_price REAL NOT NULL DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(product_id, ve_item_name, ve_item_price)
            )
        `);

        // Create default "Monthly Expenses" folder
        this.db.run('INSERT OR IGNORE INTO category_folders (name, folder_type, sort_order) VALUES (?, ?, ?)', ['Monthly Expenses', 'payable', 0]);

        const defaultCategories = [
            'General Income',
            'General Expense',
            'Loan',
            'Investment',
            'Salary',
            'Utilities',
            'Supplies'
        ];

        const stmt = this.db.prepare('INSERT OR IGNORE INTO categories (name) VALUES (?)');
        defaultCategories.forEach(cat => {
            stmt.run([cat]);
        });
        stmt.free();
    },

    /**
     * Migrate existing database schema (add new columns if missing)
     */
    migrateSchema() {
        try {
            this.db.exec('SELECT is_monthly FROM categories LIMIT 1');
        } catch (e) {
            this.db.run('ALTER TABLE categories ADD COLUMN is_monthly INTEGER DEFAULT 0');
        }

        try {
            this.db.exec('SELECT payment_for_month FROM transactions LIMIT 1');
        } catch (e) {
            this.db.run('ALTER TABLE transactions ADD COLUMN payment_for_month TEXT');
        }

        // Add default_amount column to categories
        try {
            this.db.exec('SELECT default_amount FROM categories LIMIT 1');
        } catch (e) {
            this.db.run('ALTER TABLE categories ADD COLUMN default_amount DECIMAL(10,2)');
        }

        // Add default_type column to categories
        try {
            this.db.exec('SELECT default_type FROM categories LIMIT 1');
        } catch (e) {
            this.db.run('ALTER TABLE categories ADD COLUMN default_type TEXT');
        }

        // Add folder_id column to categories
        try {
            this.db.exec('SELECT folder_id FROM categories LIMIT 1');
        } catch (e) {
            this.db.run('ALTER TABLE categories ADD COLUMN folder_id INTEGER');
        }

        // Create category_folders table if not exists
        this.db.run(`
            CREATE TABLE IF NOT EXISTS category_folders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                folder_type TEXT NOT NULL DEFAULT 'payable',
                sort_order INTEGER DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        `);

        // Add folder_type column to category_folders if missing
        try {
            this.db.exec('SELECT folder_type FROM category_folders LIMIT 1');
        } catch (e) {
            this.db.run("ALTER TABLE category_folders ADD COLUMN folder_type TEXT NOT NULL DEFAULT 'payable'");
        }

        // Add pretax_amount column to transactions
        try {
            this.db.exec('SELECT pretax_amount FROM transactions LIMIT 1');
        } catch (e) {
            this.db.run('ALTER TABLE transactions ADD COLUMN pretax_amount DECIMAL(10,2)');
        }

        // Add cashflow_sort_order column to categories
        try {
            this.db.exec('SELECT cashflow_sort_order FROM categories LIMIT 1');
        } catch (e) {
            this.db.run('ALTER TABLE categories ADD COLUMN cashflow_sort_order INTEGER DEFAULT 0');
        }

        // Add P&L flags to categories
        try {
            this.db.exec('SELECT show_on_pl FROM categories LIMIT 1');
        } catch (e) {
            this.db.run('ALTER TABLE categories ADD COLUMN show_on_pl INTEGER DEFAULT 0');
        }
        try {
            this.db.exec('SELECT is_cogs FROM categories LIMIT 1');
        } catch (e) {
            this.db.run('ALTER TABLE categories ADD COLUMN is_cogs INTEGER DEFAULT 0');
        }
        try {
            this.db.exec('SELECT is_depreciation FROM categories LIMIT 1');
        } catch (e) {
            this.db.run('ALTER TABLE categories ADD COLUMN is_depreciation INTEGER DEFAULT 0');
        }

        // Create pl_overrides table for P&L manual overrides
        this.db.run(`
            CREATE TABLE IF NOT EXISTS pl_overrides (
                category_id INTEGER,
                month TEXT,
                override_amount DECIMAL(10,2),
                PRIMARY KEY(category_id, month)
            )
        `);

        // Ensure app_meta table exists
        this.db.run(`
            CREATE TABLE IF NOT EXISTS app_meta (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        `);

        // Add is_sales_tax flag to categories
        try {
            this.db.exec('SELECT is_sales_tax FROM categories LIMIT 1');
        } catch (e) {
            this.db.run('ALTER TABLE categories ADD COLUMN is_sales_tax INTEGER DEFAULT 0');
        }

        // Add is_b2b flag to categories
        try {
            this.db.exec('SELECT is_b2b FROM categories LIMIT 1');
        } catch (e) {
            this.db.run('ALTER TABLE categories ADD COLUMN is_b2b INTEGER DEFAULT 0');
        }

        try {
            this.db.exec('SELECT default_status FROM categories LIMIT 1');
        } catch (e) {
            this.db.run('ALTER TABLE categories ADD COLUMN default_status TEXT');
        }

        // Create balance_sheet_assets table
        this.db.run(`
            CREATE TABLE IF NOT EXISTS balance_sheet_assets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                purchase_cost DECIMAL(10,2) NOT NULL,
                useful_life_months INTEGER NOT NULL,
                purchase_date DATE NOT NULL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        `);

        // One-time migration: reset show_on_pl flags (semantics inverted from "show" to "hide")
        const migrated = this.db.exec("SELECT value FROM app_meta WHERE key = 'pl_hide_migration'");
        if (migrated.length === 0 || migrated[0].values.length === 0) {
            this.db.run('UPDATE categories SET show_on_pl = 0');
            this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('pl_hide_migration', '1')");
        }

        // === New columns on balance_sheet_assets ===
        try { this.db.exec('SELECT salvage_value FROM balance_sheet_assets LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE balance_sheet_assets ADD COLUMN salvage_value DECIMAL(10,2) DEFAULT 0'); }

        try { this.db.exec('SELECT depreciation_method FROM balance_sheet_assets LIMIT 1'); }
        catch (e) { this.db.run("ALTER TABLE balance_sheet_assets ADD COLUMN depreciation_method TEXT DEFAULT 'straight_line'"); }

        try { this.db.exec('SELECT dep_start_date FROM balance_sheet_assets LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE balance_sheet_assets ADD COLUMN dep_start_date DATE'); }

        try { this.db.exec('SELECT is_depreciable FROM balance_sheet_assets LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE balance_sheet_assets ADD COLUMN is_depreciable INTEGER DEFAULT 1'); }

        try { this.db.exec('SELECT linked_transaction_id FROM balance_sheet_assets LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE balance_sheet_assets ADD COLUMN linked_transaction_id INTEGER'); }

        try { this.db.exec('SELECT notes FROM balance_sheet_assets LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE balance_sheet_assets ADD COLUMN notes TEXT'); }

        // === source_type and source_id on transactions ===
        try { this.db.exec('SELECT source_type FROM transactions LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE transactions ADD COLUMN source_type TEXT'); }

        try { this.db.exec('SELECT source_id FROM transactions LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE transactions ADD COLUMN source_id INTEGER'); }

        // === Create loans table ===
        this.db.run(`
            CREATE TABLE IF NOT EXISTS loans (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                principal DECIMAL(10,2) NOT NULL,
                annual_rate DECIMAL(8,4) NOT NULL,
                term_months INTEGER NOT NULL,
                payments_per_year INTEGER NOT NULL DEFAULT 12,
                start_date DATE NOT NULL,
                notes TEXT,
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        `);

        // === Migrate loan_config JSON → loans table row ===
        const loanMigrated = this.db.exec("SELECT value FROM app_meta WHERE key = 'loans_migration_v1'");
        if (loanMigrated.length === 0 || loanMigrated[0].values.length === 0) {
            const loanCfg = this.db.exec("SELECT value FROM app_meta WHERE key = 'loan_config'");
            if (loanCfg.length > 0 && loanCfg[0].values.length > 0) {
                try {
                    const cfg = JSON.parse(loanCfg[0].values[0][0]);
                    if (cfg && cfg.principal && cfg.start_date) {
                        const termMonths = (cfg.term_months) ? cfg.term_months : (cfg.term_years ? cfg.term_years * 12 : 0);
                        this.db.run(
                            `INSERT INTO loans (name, principal, annual_rate, term_months, payments_per_year, start_date)
                             VALUES (?, ?, ?, ?, ?, ?)`,
                            ['Primary Loan', cfg.principal, cfg.annual_rate, termMonths, cfg.payments_per_year || 12, cfg.start_date]
                        );
                    }
                } catch (e) {
                    console.warn('Failed to migrate loan_config:', e);
                }
            }
            this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('loans_migration_v1', '1')");
        }

        // === Create cashflow_overrides table ===
        this.db.run(`
            CREATE TABLE IF NOT EXISTS cashflow_overrides (
                category_id INTEGER,
                month TEXT,
                override_amount DECIMAL(10,2),
                PRIMARY KEY(category_id, month)
            )
        `);

        // === Create loan_skipped_payments table ===
        this.db.run(`
            CREATE TABLE IF NOT EXISTS loan_skipped_payments (
                loan_id INTEGER NOT NULL,
                payment_number INTEGER NOT NULL,
                PRIMARY KEY(loan_id, payment_number)
            )
        `);

        // === Create loan_payment_overrides table ===
        this.db.run(`
            CREATE TABLE IF NOT EXISTS loan_payment_overrides (
                loan_id INTEGER NOT NULL,
                payment_number INTEGER NOT NULL,
                override_amount DECIMAL(10,2) NOT NULL,
                PRIMARY KEY(loan_id, payment_number)
            )
        `);

        // === Add first_payment_date column to loans if missing ===
        try {
            this.db.exec("SELECT first_payment_date FROM loans LIMIT 1");
        } catch (e) {
            this.db.run("ALTER TABLE loans ADD COLUMN first_payment_date DATE");
        }

        // === Add is_sales flag to categories ===
        try { this.db.exec('SELECT is_sales FROM categories LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE categories ADD COLUMN is_sales INTEGER DEFAULT 0'); }

        // === Add sale_date_start and sale_date_end to transactions ===
        try { this.db.exec('SELECT sale_date_start FROM transactions LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE transactions ADD COLUMN sale_date_start DATE'); }

        try { this.db.exec('SELECT sale_date_end FROM transactions LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE transactions ADD COLUMN sale_date_end DATE'); }

        // === Add is_inventory_cost flag to categories ===
        try { this.db.exec('SELECT is_inventory_cost FROM categories LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE categories ADD COLUMN is_inventory_cost INTEGER DEFAULT 0'); }

        // === Add inventory_cost to transactions ===
        try { this.db.exec('SELECT inventory_cost FROM transactions LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE transactions ADD COLUMN inventory_cost DECIMAL(10,2)'); }

        // === Add ve_item_price to product_ve_mappings ===
        try { this.db.exec('SELECT ve_item_price FROM product_ve_mappings LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE product_ve_mappings ADD COLUMN ve_item_price REAL NOT NULL DEFAULT 0'); }

        // === Create budget_groups table ===
        this.db.run(`
            CREATE TABLE IF NOT EXISTS budget_groups (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                sort_order INTEGER DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        `);

        // === Create budget_expenses table ===
        this.db.run(`
            CREATE TABLE IF NOT EXISTS budget_expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                monthly_amount DECIMAL(10,2) NOT NULL,
                start_month TEXT NOT NULL,
                end_month TEXT,
                category_id INTEGER,
                group_id INTEGER,
                notes TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (category_id) REFERENCES categories(id),
                FOREIGN KEY (group_id) REFERENCES budget_groups(id)
            )
        `);

        // Migrate: add group_id to budget_expenses if missing
        try { this.db.run('ALTER TABLE budget_expenses ADD COLUMN group_id INTEGER REFERENCES budget_groups(id)'); } catch(e) {}

        // === Create VE sales tables ===
        this.db.run(`
            CREATE TABLE IF NOT EXISTS ve_sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                transaction_no TEXT NOT NULL,
                date DATE NOT NULL,
                billing_name TEXT,
                description TEXT,
                subtotal DECIMAL(10,2) DEFAULT 0,
                tax DECIMAL(10,2) DEFAULT 0,
                shipping DECIMAL(10,2) DEFAULT 0,
                discount DECIMAL(10,2) DEFAULT 0,
                total DECIMAL(10,2) DEFAULT 0,
                source TEXT NOT NULL DEFAULT 'online',
                UNIQUE(transaction_no, source)
            )
        `);
        this.db.run(`
            CREATE TABLE IF NOT EXISTS ve_sale_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                transaction_no TEXT NOT NULL,
                name TEXT NOT NULL,
                product_number TEXT,
                price DECIMAL(10,2) DEFAULT 0,
                quantity INTEGER DEFAULT 1,
                taxable INTEGER DEFAULT 0,
                amount DECIMAL(10,2) DEFAULT 0,
                inferred INTEGER DEFAULT 0
            )
        `);

        // === Create product-VE mapping table ===
        this.db.run(`
            CREATE TABLE IF NOT EXISTS product_ve_mappings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER NOT NULL,
                ve_item_name TEXT NOT NULL,
                ve_item_price REAL NOT NULL DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(product_id, ve_item_name, ve_item_price)
            )
        `);

        // === Create products table if missing ===
        this.db.run(`
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                sku TEXT,
                price DECIMAL(10,2) NOT NULL,
                tax_rate DECIMAL(6,4) DEFAULT 0,
                cogs DECIMAL(10,2) DEFAULT 0,
                notes TEXT,
                is_discontinued INTEGER DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        `);

        // === Add is_discontinued column to products (if table existed before this migration) ===
        try { this.db.exec('SELECT is_discontinued FROM products LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE products ADD COLUMN is_discontinued INTEGER DEFAULT 0'); }

        // === Create ve_events table ===
        this.db.run(`
            CREATE TABLE IF NOT EXISTS ve_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                type TEXT NOT NULL DEFAULT 'tradeshow',
                start_date DATE NOT NULL,
                end_date DATE NOT NULL,
                notes TEXT,
                journal_added INTEGER DEFAULT 0
            )
        `);

        // === Add journal_added column to ve_events ===
        try { this.db.exec('SELECT journal_added FROM ve_events LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE ve_events ADD COLUMN journal_added INTEGER DEFAULT 0'); }

        // === Add event_id column to ve_sales ===
        try { this.db.exec('SELECT event_id FROM ve_sales LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE ve_sales ADD COLUMN event_id INTEGER'); }
    },

    // ==================== FOLDER OPERATIONS ====================

    /**
     * Get all category folders
     * @returns {Array} Array of folder objects
     */
    getFolders() {
        const results = this.db.exec('SELECT * FROM category_folders ORDER BY sort_order ASC, name ASC');
        if (results.length === 0) return [];
        return this.rowsToObjects(results[0]);
    },

    /**
     * Add a new folder
     * @param {string} name - Folder name
     * @param {string} type - Folder type ('payable' or 'receivable')
     * @returns {number} New folder ID
     */
    addFolder(name, type = 'payable') {
        this.db.run('INSERT INTO category_folders (name, folder_type) VALUES (?, ?)', [name.trim(), type]);
        const result = this.db.exec('SELECT last_insert_rowid() as id');
        this.autoSave();
        return result[0].values[0][0];
    },

    /**
     * Update a folder
     * @param {number} id - Folder ID
     * @param {string} name - New name
     * @param {string} type - Folder type ('payable' or 'receivable')
     */
    updateFolder(id, name, type = 'payable') {
        this.db.run('UPDATE category_folders SET name = ?, folder_type = ? WHERE id = ?', [name.trim(), type, id]);
        this.autoSave();
    },

    /**
     * Delete a folder (moves its categories to unfiled)
     * @param {number} id - Folder ID
     */
    deleteFolder(id) {
        this.db.run('UPDATE categories SET folder_id = NULL WHERE folder_id = ?', [id]);
        this.db.run('DELETE FROM category_folders WHERE id = ?', [id]);
        this.autoSave();
    },

    /**
     * Get folder by ID
     * @param {number} id - Folder ID
     * @returns {Object|null} Folder object
     */
    getFolderById(id) {
        const results = this.db.exec('SELECT * FROM category_folders WHERE id = ?', [id]);
        if (results.length === 0) return null;
        return this.rowsToObjects(results[0])[0];
    },

    // ==================== CATEGORY OPERATIONS ====================

    /**
     * Get all categories with folder info
     * @returns {Array} Array of category objects
     */
    getCategories() {
        const results = this.db.exec(`
            SELECT c.*, cf.name as folder_name
            FROM categories c
            LEFT JOIN category_folders cf ON c.folder_id = cf.id
            ORDER BY cf.sort_order ASC, cf.name ASC, c.name ASC
        `);
        if (results.length === 0) return [];
        return this.rowsToObjects(results[0]);
    },

    /**
     * Add a new category
     * @param {string} name - Category name
     * @param {boolean} isMonthly - Whether this is a monthly payment category
     * @param {number|null} defaultAmount - Default amount for this category
     * @param {string|null} defaultType - Default type ('receivable' or 'payable')
     * @param {number|null} folderId - Folder ID
     * @returns {number} New category ID
     */
    addCategory(name, isMonthly = false, defaultAmount = null, defaultType = null, folderId = null, showOnPl = false, isCogs = false, isDepreciation = false, isSalesTax = false, isB2b = false, defaultStatus = null, isSales = false, isInventoryCost = false) {
        this.db.run(
            'INSERT INTO categories (name, is_monthly, default_amount, default_type, folder_id, show_on_pl, is_cogs, is_depreciation, is_sales_tax, is_b2b, default_status, is_sales, is_inventory_cost) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
            [name.trim(), isMonthly ? 1 : 0, defaultAmount, defaultType, folderId, showOnPl ? 1 : 0, isCogs ? 1 : 0, isDepreciation ? 1 : 0, isSalesTax ? 1 : 0, isB2b ? 1 : 0, defaultStatus, isSales ? 1 : 0, isInventoryCost ? 1 : 0]
        );
        const result = this.db.exec('SELECT last_insert_rowid() as id');
        this.autoSave();
        return result[0].values[0][0];
    },

    /**
     * Delete a category
     * @param {number} id - Category ID
     * @returns {boolean} Success (false if category is in use)
     */
    deleteCategory(id) {
        const inUse = this.db.exec('SELECT COUNT(*) as count FROM transactions WHERE category_id = ?', [id]);
        if (inUse[0].values[0][0] > 0) {
            return false;
        }
        this.db.run('DELETE FROM categories WHERE id = ?', [id]);
        this.autoSave();
        return true;
    },

    /**
     * Get or create the system "Sales Tax" category
     * @returns {number} Category ID for the Sales Tax category
     */
    getSalesCategories() {
        const results = this.db.exec('SELECT id, name FROM categories WHERE is_sales = 1 ORDER BY name ASC');
        if (results.length === 0) return [];
        return this.rowsToObjects(results[0]);
    },

    getOrCreateSalesTaxCategory() {
        const result = this.db.exec('SELECT id FROM categories WHERE is_sales_tax = 1 LIMIT 1');
        if (result.length > 0 && result[0].values.length > 0) {
            return result[0].values[0][0];
        }
        return this.addCategory('Sales Tax', false, null, 'payable', null, false, false, false, true, false, 'pending', false);
    },

    /**
     * Get the linked sales tax transaction for a parent sale
     * @param {number} parentId - Parent transaction ID
     * @returns {number|null} Child transaction ID or null
     */
    getLinkedSalesTaxTransaction(parentId) {
        const result = this.db.exec("SELECT id FROM transactions WHERE source_type = 'sales_tax' AND source_id = ?", [parentId]);
        if (result.length > 0 && result[0].values.length > 0) {
            return result[0].values[0][0];
        }
        return null;
    },

    /**
     * Update only the mutable fields of a linked sales tax transaction
     * @param {number} id - Sales tax transaction ID
     * @param {number} amount - Tax amount
     * @param {string} entryDate - Entry date
     * @param {string} monthDue - Month due
     * @param {string} description - Item description
     */
    updateSalesTaxTransaction(id, amount, entryDate, monthDue, description) {
        this.db.run(
            'UPDATE transactions SET amount = ?, entry_date = ?, month_due = ?, item_description = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
            [amount, entryDate, monthDue, description, id]
        );
        this.autoSave();
    },

    /**
     * Get or create the system "Inventory Cost" category
     * @returns {number} Category ID for the Inventory Cost category
     */
    getOrCreateInventoryCostCategory() {
        const result = this.db.exec('SELECT id FROM categories WHERE is_inventory_cost = 1 LIMIT 1');
        if (result.length > 0 && result[0].values.length > 0) {
            return result[0].values[0][0];
        }
        return this.addCategory('Inventory Cost', false, null, 'payable', null, false, true, false, false, false, 'pending', false, true);
    },

    /**
     * Get the linked inventory cost transaction for a parent sale
     * @param {number} parentId - Parent transaction ID
     * @returns {number|null} Child transaction ID or null
     */
    getLinkedInventoryCostTransaction(parentId) {
        const result = this.db.exec("SELECT id FROM transactions WHERE source_type = 'inventory_cost' AND source_id = ?", [parentId]);
        if (result.length > 0 && result[0].values.length > 0) {
            return result[0].values[0][0];
        }
        return null;
    },

    /**
     * Update only the mutable fields of a linked inventory cost transaction
     * @param {number} id - Inventory cost transaction ID
     * @param {number} amount - Cost amount
     * @param {string} entryDate - Entry date
     * @param {string} monthDue - Month due
     * @param {string} description - Item description
     */
    updateInventoryCostTransaction(id, amount, entryDate, monthDue, description) {
        this.db.run(
            'UPDATE transactions SET amount = ?, entry_date = ?, month_due = ?, item_description = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
            [amount, entryDate, monthDue, description, id]
        );
        this.autoSave();
    },

    /**
     * Get category by ID
     * @param {number} id - Category ID
     * @returns {Object|null} Category object
     */
    getCategoryById(id) {
        const results = this.db.exec(`
            SELECT c.*, cf.name as folder_name
            FROM categories c
            LEFT JOIN category_folders cf ON c.folder_id = cf.id
            WHERE c.id = ?
        `, [id]);
        if (results.length === 0) return null;
        return this.rowsToObjects(results[0])[0];
    },

    /**
     * Update a category
     * @param {number} id - Category ID
     * @param {string} name - New name
     * @param {boolean} isMonthly - Whether this is a monthly payment category
     * @param {number|null} defaultAmount - Default amount
     * @param {string|null} defaultType - Default type
     * @param {number|null} folderId - Folder ID
     */
    updateCategory(id, name, isMonthly = false, defaultAmount = null, defaultType = null, folderId = null, showOnPl = false, isCogs = false, isDepreciation = false, isSalesTax = false, isB2b = false, defaultStatus = null, isSales = false, isInventoryCost = false) {
        this.db.run(
            'UPDATE categories SET name = ?, is_monthly = ?, default_amount = ?, default_type = ?, folder_id = ?, show_on_pl = ?, is_cogs = ?, is_depreciation = ?, is_sales_tax = ?, is_b2b = ?, default_status = ?, is_sales = ?, is_inventory_cost = ? WHERE id = ?',
            [name.trim(), isMonthly ? 1 : 0, defaultAmount, defaultType, folderId, showOnPl ? 1 : 0, isCogs ? 1 : 0, isDepreciation ? 1 : 0, isSalesTax ? 1 : 0, isB2b ? 1 : 0, defaultStatus, isSales ? 1 : 0, isInventoryCost ? 1 : 0, id]
        );
        this.autoSave();
    },

    /**
     * Get all categories in a specific folder
     * @param {number} folderId - Folder ID
     * @returns {Array} Array of category objects
     */
    getCategoriesByFolder(folderId) {
        const results = this.db.exec(
            'SELECT * FROM categories WHERE folder_id = ? ORDER BY name ASC',
            [folderId]
        );
        if (results.length === 0) return [];
        return this.rowsToObjects(results[0]);
    },

    /**
     * Update cashflow sort order for a list of category IDs
     * @param {Array<{id: number, sortOrder: number}>} orderList
     */
    updateCashflowSortOrder(orderList) {
        const stmt = this.db.prepare('UPDATE categories SET cashflow_sort_order = ? WHERE id = ?');
        orderList.forEach(({ id, sortOrder }) => {
            stmt.run([sortOrder, id]);
        });
        stmt.free();
        this.autoSave();
    },

    /**
     * Get count of transactions using a category
     * @param {number} categoryId - Category ID
     * @returns {number} Transaction count
     */
    getCategoryUsageCount(categoryId) {
        const result = this.db.exec('SELECT COUNT(*) as count FROM transactions WHERE category_id = ?', [categoryId]);
        return result[0].values[0][0];
    },

    // ==================== TRANSACTION OPERATIONS ====================

    /**
     * Get all transactions
     * @param {Object} filters - Optional filters
     * @returns {Array} Array of transaction objects
     */
    getTransactions(filters = {}) {
        let query = `
            SELECT t.*, c.name as category_name, c.is_monthly as category_is_monthly,
                   c.is_sales as category_is_sales, c.is_sales_tax as category_is_sales_tax
            FROM transactions t
            LEFT JOIN categories c ON t.category_id = c.id
            WHERE 1=1
        `;
        const params = [];

        if (filters.type) {
            query += ' AND t.transaction_type = ?';
            params.push(filters.type);
        }

        if (filters.status) {
            query += ' AND t.status = ?';
            params.push(filters.status);
        }

        if (filters.month) {
            query += ' AND substr(t.entry_date, 1, 7) = ?';
            params.push(filters.month);
        }

        if (filters.folderId) {
            if (filters.folderId === 'unfiled') {
                query += ' AND c.folder_id IS NULL';
            } else {
                query += ' AND c.folder_id = ?';
                params.push(filters.folderId);
            }
        }

        if (filters.categoryId) {
            query += ' AND t.category_id = ?';
            params.push(filters.categoryId);
        }

        query += ' ORDER BY t.entry_date DESC, t.id DESC';

        const results = this.db.exec(query, params);
        if (results.length === 0) return [];

        return this.rowsToObjects(results[0]);
    },

    /**
     * Get a single transaction by ID
     * @param {number} id - Transaction ID
     * @returns {Object|null} Transaction object
     */
    getTransactionById(id) {
        const results = this.db.exec(`
            SELECT t.*, c.name as category_name, c.is_monthly as category_is_monthly,
                   c.is_sales as category_is_sales, c.is_sales_tax as category_is_sales_tax
            FROM transactions t
            LEFT JOIN categories c ON t.category_id = c.id
            WHERE t.id = ?
        `, [id]);

        if (results.length === 0) return null;
        return this.rowsToObjects(results[0])[0];
    },

    /**
     * Add a new transaction
     * @param {Object} transaction - Transaction data
     * @returns {number} New transaction ID
     */
    addTransaction(transaction) {
        this.db.run(`
            INSERT INTO transactions
            (entry_date, category_id, item_description, amount, pretax_amount, transaction_type,
             status, date_processed, month_due, month_paid, payment_for_month, notes,
             source_type, source_id, sale_date_start, sale_date_end, inventory_cost)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `, [
            transaction.entry_date,
            transaction.category_id,
            transaction.item_description || null,
            transaction.amount,
            transaction.pretax_amount || null,
            transaction.transaction_type,
            transaction.status,
            transaction.date_processed || null,
            transaction.month_due || null,
            transaction.month_paid || null,
            transaction.payment_for_month || null,
            transaction.notes || null,
            transaction.source_type || null,
            transaction.source_id || null,
            transaction.sale_date_start || null,
            transaction.sale_date_end || null,
            transaction.inventory_cost || null
        ]);

        const result = this.db.exec('SELECT last_insert_rowid() as id');
        this.autoSave();
        return result[0].values[0][0];
    },

    /**
     * Update a transaction
     * @param {number} id - Transaction ID
     * @param {Object} transaction - Transaction data
     */
    updateTransaction(id, transaction) {
        this.db.run(`
            UPDATE transactions SET
                entry_date = ?,
                category_id = ?,
                item_description = ?,
                amount = ?,
                pretax_amount = ?,
                transaction_type = ?,
                status = ?,
                date_processed = ?,
                month_due = ?,
                month_paid = ?,
                payment_for_month = ?,
                notes = ?,
                sale_date_start = ?,
                sale_date_end = ?,
                inventory_cost = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        `, [
            transaction.entry_date,
            transaction.category_id,
            transaction.item_description || null,
            transaction.amount,
            transaction.pretax_amount || null,
            transaction.transaction_type,
            transaction.status,
            transaction.date_processed || null,
            transaction.month_due || null,
            transaction.month_paid || null,
            transaction.payment_for_month || null,
            transaction.notes || null,
            transaction.sale_date_start || null,
            transaction.sale_date_end || null,
            transaction.inventory_cost || null,
            id
        ]);
        this.autoSave();
    },

    /**
     * Update just the status of a transaction (for inline status changes)
     * @param {number} id - Transaction ID
     * @param {string} status - New status
     * @param {string} monthPaidValue - Optional month paid value (required for paid/received)
     */
    setTransactionDateProcessed(id, date) {
        this.db.run('UPDATE transactions SET date_processed = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?', [date, id]);
        this.autoSave();
    },

    updateTransactionStatus(id, status, monthPaidValue = null) {
        if (status === 'pending') {
            // Reverting to pending: clear date_processed and month_paid
            this.db.run(`
                UPDATE transactions SET
                    status = ?,
                    date_processed = NULL,
                    month_paid = NULL,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            `, [status, id]);
        } else {
            // Paid/received: set month_paid (required)
            const monthPaid = monthPaidValue || Utils.getCurrentMonth();
            this.db.run(`
                UPDATE transactions SET
                    status = ?,
                    month_paid = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            `, [status, monthPaid, id]);
        }
        this.autoSave();
    },

    /**
     * Bulk set date paid for multiple transactions
     * @param {Array} updates - Array of { id, status } objects
     * @param {string} dateProcessed - Date processed value
     * @param {string} monthPaid - Month paid value (YYYY-MM)
     */
    bulkSetDatePaid(updates, dateProcessed, monthPaid) {
        for (const { id, status } of updates) {
            this.db.run(`
                UPDATE transactions SET
                    status = ?,
                    date_processed = ?,
                    month_paid = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            `, [status, dateProcessed, monthPaid, id]);
        }
        this.autoSave();
    },

    /**
     * Bulk reset transactions to pending
     * @param {Array} ids - Array of transaction IDs
     */
    bulkResetToPending(ids) {
        for (const id of ids) {
            this.db.run(`
                UPDATE transactions SET
                    status = 'pending',
                    date_processed = NULL,
                    month_paid = NULL,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            `, [id]);
        }
        this.autoSave();
    },

    /**
     * Delete a transaction
     * @param {number} id - Transaction ID
     */
    deleteTransaction(id) {
        this.db.run('DELETE FROM transactions WHERE id = ?', [id]);
        this.autoSave();
    },

    // ==================== JOURNAL METADATA ====================

    /**
     * Get journal owner name
     * @returns {string} Journal owner name (empty string if not set)
     */
    getJournalOwner() {
        const result = this.db.exec("SELECT value FROM app_meta WHERE key = 'journal_owner'");
        if (result.length === 0 || result[0].values.length === 0) {
            return '';
        }
        return result[0].values[0][0] || '';
    },

    /**
     * Set journal owner name
     * @param {string} owner - Owner/company name
     */
    setJournalOwner(owner) {
        this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('journal_owner', ?)", [owner]);
        this.autoSave();
    },

    /**
     * Get journal name (legacy support)
     * @returns {string} Journal name
     */
    getJournalName() {
        const result = this.db.exec("SELECT value FROM app_meta WHERE key = 'journal_name'");
        if (result.length === 0 || result[0].values.length === 0) {
            return 'Accounting Journal';
        }
        return result[0].values[0][0];
    },

    /**
     * Set journal name (legacy support)
     * @param {string} name - Journal name
     */
    setJournalName(name) {
        this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('journal_name', ?)", [name]);
        this.autoSave();
    },

    // ==================== SYNC CONFIG (localStorage — survives DB replacement) ====================

    getSyncConfig() {
        try { return JSON.parse(localStorage.getItem('sync_config')); } catch { return null; }
    },

    setSyncConfig(config) {
        localStorage.setItem('sync_config', JSON.stringify(config));
    },

    clearSyncConfig() {
        localStorage.removeItem('sync_config');
    },

    getSupabaseConfig() {
        try { return JSON.parse(localStorage.getItem('supabase_config')); } catch { return null; }
    },

    setSupabaseConfig(config) {
        localStorage.setItem('supabase_config', JSON.stringify(config));
    },

    // ==================== CALCULATIONS ====================

    /**
     * Calculate summary totals
     * @returns {Object} Summary object with cashBalance, receivables, payables
     */
    calculateSummary() {
        const receivedResult = this.db.exec(`
            SELECT COALESCE(SUM(amount), 0) as total
            FROM transactions
            WHERE transaction_type = 'receivable' AND status = 'received'
        `);
        const totalReceived = receivedResult[0].values[0][0];

        const paidResult = this.db.exec(`
            SELECT COALESCE(SUM(amount), 0) as total
            FROM transactions
            WHERE transaction_type = 'payable' AND status = 'paid'
        `);
        const totalPaid = paidResult[0].values[0][0];

        const receivablesResult = this.db.exec(`
            SELECT COALESCE(SUM(amount), 0) as total
            FROM transactions
            WHERE transaction_type = 'receivable' AND status = 'pending'
        `);
        const pendingReceivables = receivablesResult[0].values[0][0];

        const payablesResult = this.db.exec(`
            SELECT COALESCE(SUM(amount), 0) as total
            FROM transactions
            WHERE transaction_type = 'payable' AND status = 'pending'
        `);
        const pendingPayables = payablesResult[0].values[0][0];

        return {
            cashBalance: totalReceived - totalPaid,
            accountsReceivable: pendingReceivables,
            accountsPayable: pendingPayables
        };
    },

    /**
     * Check if there are any late payments in the completed transactions
     * @returns {Object} Object with lateReceivedAmount, latePaidAmount
     */
    checkLatePayments() {
        const lateReceivedResult = this.db.exec(`
            SELECT COALESCE(SUM(amount), 0) as total FROM transactions
            WHERE transaction_type = 'receivable'
            AND status = 'received'
            AND month_due IS NOT NULL
            AND month_paid IS NOT NULL
            AND month_paid > month_due
        `);
        const lateReceivedAmount = lateReceivedResult[0].values[0][0];

        const latePaidResult = this.db.exec(`
            SELECT COALESCE(SUM(amount), 0) as total FROM transactions
            WHERE transaction_type = 'payable'
            AND status = 'paid'
            AND month_due IS NOT NULL
            AND month_paid IS NOT NULL
            AND month_paid > month_due
        `);
        const latePaidAmount = latePaidResult[0].values[0][0];

        return {
            lateReceivedAmount,
            latePaidAmount,
            hasLateReceivables: lateReceivedAmount > 0,
            hasLatePayables: latePaidAmount > 0
        };
    },

    /**
     * Get monthly summary data (grouped by entry date)
     * @returns {Array} Array of monthly summary objects
     */
    getMonthlySummary() {
        const result = this.db.exec(`
            SELECT
                substr(entry_date, 1, 7) as month,
                SUM(CASE WHEN transaction_type = 'receivable' AND status = 'received' THEN amount ELSE 0 END) as received,
                SUM(CASE WHEN transaction_type = 'payable' AND status = 'paid' THEN amount ELSE 0 END) as paid,
                SUM(CASE WHEN transaction_type = 'receivable' AND status = 'pending' THEN amount ELSE 0 END) as pending_receivables,
                SUM(CASE WHEN transaction_type = 'payable' AND status = 'pending' THEN amount ELSE 0 END) as pending_payables,
                COUNT(*) as total_entries
            FROM transactions
            GROUP BY substr(entry_date, 1, 7)
            ORDER BY month DESC
        `);

        if (result.length === 0) return [];
        return this.rowsToObjects(result[0]);
    },

    /**
     * Get cash flow summary grouped by month_paid (when money actually moved)
     * Only includes completed transactions (paid/received), grouped by the month they were processed
     * @returns {Array} Array of cash flow summary objects
     */
    getCashFlowSummary() {
        const result = this.db.exec(`
            SELECT
                month_paid as month,
                SUM(CASE WHEN transaction_type = 'receivable' AND status = 'received' THEN amount ELSE 0 END) as cash_in,
                SUM(CASE WHEN transaction_type = 'payable' AND status = 'paid' THEN amount ELSE 0 END) as cash_out,
                COUNT(*) as total_entries
            FROM transactions
            WHERE month_paid IS NOT NULL
            AND status != 'pending'
            GROUP BY month_paid
            ORDER BY month DESC
        `);

        if (result.length === 0) return [];
        return this.rowsToObjects(result[0]);
    },

    /**
     * Get cash flow data broken down by category and month for spreadsheet view
     * @returns {Object} { months: string[], data: Object[] }
     */
    getCashFlowSpreadsheet() {
        // Get all distinct months from month_paid (sorted ASC)
        const monthsResult = this.db.exec(`
            SELECT DISTINCT month_paid as month FROM transactions
            WHERE month_paid IS NOT NULL AND status != 'pending'
            ORDER BY month ASC
        `);
        const months = monthsResult.length > 0 ? monthsResult[0].values.map(r => r[0]) : [];

        // Get per-category, per-month totals for completed transactions
        const dataResult = this.db.exec(`
            SELECT c.name as category_name, c.id as category_id,
                   c.is_b2b, c.is_cogs,
                   t.transaction_type, t.month_paid as month,
                   SUM(t.amount) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.status != 'pending' AND t.month_paid IS NOT NULL
            GROUP BY c.id, t.month_paid, t.transaction_type
            ORDER BY c.cashflow_sort_order ASC, c.name ASC
        `);
        const data = dataResult.length > 0 ? this.rowsToObjects(dataResult[0]) : [];

        return { months, data };
    },

    /**
     * Get all transactions as flat data for CSV export
     * @returns {Array} Array of transaction objects with all fields
     */
    getTransactionsForExport() {
        const results = this.db.exec(`
            SELECT
                t.entry_date,
                c.name as category,
                t.transaction_type as type,
                t.amount,
                t.pretax_amount,
                t.status,
                t.month_due,
                t.month_paid,
                t.date_processed,
                t.payment_for_month,
                t.notes
            FROM transactions t
            LEFT JOIN categories c ON t.category_id = c.id
            ORDER BY t.entry_date DESC, t.id DESC
        `);

        if (results.length === 0) return [];
        return this.rowsToObjects(results[0]);
    },

    // ==================== PROFIT & LOSS ====================

    /**
     * Get total actual revenue per month (accrual-based, from receivable non-COGS categories)
     * @returns {Object} { 'YYYY-MM': totalRevenue, ... }
     */
    getActualRevenueByMonth() {
        const result = this.db.exec(`
            SELECT t.month_due as month,
                   SUM(COALESCE(t.pretax_amount, t.amount)) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.month_due IS NOT NULL
            AND t.transaction_type = 'receivable'
            AND c.is_cogs = 0
            AND c.show_on_pl != 1
            AND (c.is_b2b = 1 OR c.is_sales = 1)
            AND COALESCE(t.source_type, '') NOT IN ('loan_receivable', 'loan_payment')
            GROUP BY t.month_due
        `);
        const map = {};
        if (result.length > 0) {
            result[0].values.forEach(row => { map[row[0]] = row[1]; });
        }
        return map;
    },

    /**
     * Get P&L-consistent revenue by month, respecting pl_overrides.
     * Returns total per month and optional B2B/consumer split (via is_b2b flag).
     * Matches P&L revenue computation so progress tracker aligns with P&L tab.
     * @returns {Object} { total: { 'YYYY-MM': number }, b2b: { 'YYYY-MM': number }, consumer: { 'YYYY-MM': number } }
     */
    getPLRevenueByMonth() {
        const overrides = this.getAllPLOverrides();

        // Revenue per category per month (same query as P&L spreadsheet)
        const result = this.db.exec(`
            SELECT c.id as category_id, c.is_b2b, t.month_due as month,
                   SUM(COALESCE(t.pretax_amount, t.amount)) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.month_due IS NOT NULL
            AND t.transaction_type = 'receivable'
            AND c.is_cogs = 0 AND c.show_on_pl != 1
            AND (c.is_b2b = 1 OR c.is_sales = 1)
            AND COALESCE(t.source_type, '') NOT IN ('loan_receivable', 'loan_payment')
            GROUP BY c.id, t.month_due
        `);

        const total = {}, b2b = {}, consumer = {};
        if (result.length > 0) {
            result[0].values.forEach(row => {
                const catId = row[0];
                const isB2b = row[1];
                const month = row[2];
                let amount = row[3];

                // Apply P&L override if one exists for this category+month
                const overrideKey = `${catId}-${month}`;
                if (overrideKey in overrides) {
                    amount = overrides[overrideKey];
                }

                total[month] = (total[month] || 0) + amount;
                if (isB2b) {
                    b2b[month] = (b2b[month] || 0) + amount;
                } else {
                    consumer[month] = (consumer[month] || 0) + amount;
                }
            });
        }

        // Also include override-only categories (categories with overrides but no transactions)
        Object.keys(overrides).forEach(key => {
            const [catIdStr, month] = key.split('-');
            const catId = parseInt(catIdStr);
            if (catId < 0) return; // skip tax override
            // Check if this is a revenue category we haven't already counted
            const alreadyCounted = result.length > 0 && result[0].values.some(
                row => row[0] === catId && row[2] === month
            );
            if (!alreadyCounted) {
                const catResult = this.db.exec(
                    `SELECT is_b2b FROM categories WHERE id = ? AND is_cogs = 0 AND show_on_pl != 1`, [catId]
                );
                if (catResult.length > 0 && catResult[0].values.length > 0) {
                    const isB2b = catResult[0].values[0][0];
                    const amount = overrides[key];
                    total[month] = (total[month] || 0) + amount;
                    if (isB2b) {
                        b2b[month] = (b2b[month] || 0) + amount;
                    } else {
                        consumer[month] = (consumer[month] || 0) + amount;
                    }
                }
            }
        });

        return { total, b2b, consumer };
    },

    /**
     * Get P&L spreadsheet data (accrual-based: uses month_due, includes all statuses)
     * Revenue uses COALESCE(pretax_amount, amount) for receivable categories.
     * @returns {Object} { months, revenue, cogs, opex }
     */
    getPLSpreadsheet() {
        // Get all distinct months from month_due (accrual basis)
        const monthsResult = this.db.exec(`
            SELECT DISTINCT t.month_due as month FROM transactions t
            WHERE t.month_due IS NOT NULL
            ORDER BY month ASC
        `);
        const months = monthsResult.length > 0 ? monthsResult[0].values.map(r => r[0]) : [];

        // Revenue: receivable categories (not COGS, not hidden), using pretax_amount if available
        const revenueResult = this.db.exec(`
            SELECT c.id as category_id, c.name as category_name,
                   c.is_b2b,
                   t.month_due as month,
                   SUM(COALESCE(t.pretax_amount, t.amount)) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.month_due IS NOT NULL
            AND t.transaction_type = 'receivable'
            AND c.is_cogs = 0
            AND c.show_on_pl != 1
            AND (c.is_b2b = 1 OR c.is_sales = 1)
            AND COALESCE(t.source_type, '') NOT IN ('loan_receivable', 'loan_payment')
            GROUP BY c.id, t.month_due
            ORDER BY c.cashflow_sort_order ASC, c.name ASC
        `);
        const revenue = revenueResult.length > 0 ? this.rowsToObjects(revenueResult[0]) : [];

        // COGS: is_cogs=1 categories (not hidden), accrual basis
        const cogsResult = this.db.exec(`
            SELECT c.id as category_id, c.name as category_name,
                   c.is_b2b,
                   t.month_due as month,
                   SUM(t.amount) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.month_due IS NOT NULL
            AND c.is_cogs = 1
            AND c.show_on_pl != 1
            GROUP BY c.id, t.month_due
            ORDER BY c.cashflow_sort_order ASC, c.name ASC
        `);
        const cogs = cogsResult.length > 0 ? this.rowsToObjects(cogsResult[0]) : [];

        // OpEx: all payable categories that are not COGS, not depreciation, not sales tax, not hidden, not loan payments, accrual basis
        const opexResult = this.db.exec(`
            SELECT c.id as category_id, c.name as category_name,
                   t.month_due as month,
                   SUM(t.amount) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.month_due IS NOT NULL
            AND t.transaction_type = 'payable'
            AND c.is_cogs = 0
            AND c.is_depreciation = 0
            AND c.is_sales_tax = 0
            AND c.show_on_pl != 1
            AND COALESCE(t.source_type, '') NOT IN ('loan_payment', 'loan_receivable')
            GROUP BY c.id, t.month_due
            ORDER BY c.cashflow_sort_order ASC, c.name ASC
        `);
        const opex = opexResult.length > 0 ? this.rowsToObjects(opexResult[0]) : [];

        // Depreciation: categories flagged is_depreciation=1, shown regardless of show_on_pl
        // Values are manually entered via pl_overrides (no transaction aggregation)
        const depreciationResult = this.db.exec(`
            SELECT id as category_id, name as category_name
            FROM categories
            WHERE is_depreciation = 1
            ORDER BY cashflow_sort_order ASC, name ASC
        `);
        const depreciation = depreciationResult.length > 0 ? this.rowsToObjects(depreciationResult[0]) : [];

        // Computed asset depreciation and loan interest by month
        const assetDeprByMonth = this.getAssetDepreciationByMonth(null);
        const loanInterestByMonth = this.getLoanInterestByMonth(null);

        // Merge their month keys into the master months array
        const allMonths = new Set(months);
        Object.keys(assetDeprByMonth).forEach(m => allMonths.add(m));
        Object.keys(loanInterestByMonth).forEach(m => allMonths.add(m));
        const mergedMonths = Array.from(allMonths).sort();

        return { months: mergedMonths, revenue, cogs, opex, depreciation, assetDeprByMonth, loanInterestByMonth };
    },

    /**
     * Compute total operating expenses per month, matching P&L renderer exactly.
     * Sums: opex categories (with overrides), depreciation categories (from overrides),
     * asset depreciation, and loan interest.
     * @param {string[]} months - Array of month strings (YYYY-MM)
     * @returns {{ [month: string]: number }} Per-month total operating expenses
     */
    getMonthlyTotalOpex(months) {
        const plData = this.getPLSpreadsheet();
        const overrides = this.getAllPLOverrides();
        const result = {};
        months.forEach(m => { result[m] = 0; });

        // Group opex by category → { catId: { months: { m: total } } }
        const opexByCat = {};
        (plData.opex || []).forEach(row => {
            if (!opexByCat[row.category_id]) opexByCat[row.category_id] = {};
            opexByCat[row.category_id][row.month] = (opexByCat[row.category_id][row.month] || 0) + row.total;
        });

        // Opex categories (with overrides)
        Object.entries(opexByCat).forEach(([catId, catMonths]) => {
            months.forEach(m => {
                const key = `${catId}-${m}`;
                const computed = catMonths[m] || 0;
                const val = (key in overrides) ? overrides[key] : computed;
                result[m] += val;
            });
        });

        // Depreciation categories (values from pl_overrides only)
        (plData.depreciation || []).forEach(cat => {
            months.forEach(m => {
                const key = `${cat.category_id}-${m}`;
                const val = (key in overrides) ? overrides[key] : 0;
                result[m] += val;
            });
        });

        // Asset depreciation
        const assetDepr = plData.assetDeprByMonth || {};
        months.forEach(m => { result[m] += (assetDepr[m] || 0); });

        // Loan interest
        const loanInt = plData.loanInterestByMonth || {};
        months.forEach(m => { result[m] += (loanInt[m] || 0); });

        return result;
    },

    /**
     * Get P&L tax mode setting
     * @returns {string} 'corporate' or 'passthrough'
     */
    getPLTaxMode() {
        const result = this.db.exec("SELECT value FROM app_meta WHERE key = 'pl_tax_mode'");
        if (result.length === 0 || result[0].values.length === 0) {
            return 'corporate';
        }
        return result[0].values[0][0] || 'corporate';
    },

    /**
     * Set P&L tax mode setting
     * @param {string} mode - 'corporate' or 'passthrough'
     */
    setPLTaxMode(mode) {
        this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('pl_tax_mode', ?)", [mode]);
        this.autoSave();
    },

    /**
     * Get/set persisted as-of month for a given tab
     * @param {string} tab - 'pnl' or 'cf'
     */
    getAsOfMonth(tab) {
        const key = tab + '_as_of_month';
        const result = this.db.exec("SELECT value FROM app_meta WHERE key = ?", [key]);
        if (result.length === 0 || result[0].values.length === 0) return 'current';
        return result[0].values[0][0] || 'current';
    },

    setAsOfMonth(tab, value) {
        const key = tab + '_as_of_month';
        this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES (?, ?)", [key, value]);
        this.autoSave();
    },

    // ==================== THEME SETTINGS ====================

    /**
     * Get theme preset name
     * @returns {string} Preset name (default, ocean, forest, sunset, midnight, custom)
     */
    getThemePreset() {
        const result = this.db.exec("SELECT value FROM app_meta WHERE key = 'theme_preset'");
        if (result.length === 0 || result[0].values.length === 0) return 'default';
        return result[0].values[0][0] || 'default';
    },

    /**
     * Set theme preset name
     * @param {string} name - Preset name
     */
    setThemePreset(name) {
        this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('theme_preset', ?)", [name]);
        this.autoSave();
    },

    /**
     * Get custom theme colors
     * @returns {Object|null} Color object {c1, c2, c3, c4} or null
     */
    getThemeColors() {
        const result = this.db.exec("SELECT value FROM app_meta WHERE key = 'theme_colors'");
        if (result.length === 0 || result[0].values.length === 0) return null;
        try {
            return JSON.parse(result[0].values[0][0]);
        } catch (e) {
            return null;
        }
    },

    /**
     * Set custom theme colors
     * @param {Object} colors - Color object {c1, c2, c3, c4}
     */
    setThemeColors(colors) {
        this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('theme_colors', ?)", [JSON.stringify(colors)]);
        this.autoSave();
    },

    /**
     * Get dark mode setting
     * @returns {boolean} True if dark mode enabled
     */
    getThemeDark() {
        const result = this.db.exec("SELECT value FROM app_meta WHERE key = 'theme_dark'");
        if (result.length === 0 || result[0].values.length === 0) return false;
        return result[0].values[0][0] === '1';
    },

    /**
     * Set dark mode setting
     * @param {boolean} isDark - True for dark mode
     */
    setThemeDark(isDark) {
        this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('theme_dark', ?)", [isDark ? '1' : '0']);
        this.autoSave();
    },

    // ==================== FIXED ASSETS ====================

    /**
     * Get all fixed assets
     * @returns {Array} Array of fixed asset objects
     */
    getFixedAssets() {
        const results = this.db.exec('SELECT * FROM balance_sheet_assets ORDER BY purchase_date ASC, name ASC');
        if (results.length === 0) return [];
        return this.rowsToObjects(results[0]);
    },

    /**
     * Get total purchase cost of all fixed assets
     * @returns {number} Sum of purchase_cost from balance_sheet_assets
     */
    getTotalAssetPurchaseCost() {
        const result = this.db.exec('SELECT COALESCE(SUM(purchase_cost), 0) AS total FROM balance_sheet_assets');
        return (result.length > 0 && result[0].values.length > 0) ? result[0].values[0][0] : 0;
    },

    /**
     * Add a fixed asset
     * @param {string} name - Asset name
     * @param {number} purchaseCost - Purchase cost
     * @param {number} usefulLifeMonths - Useful life in months
     * @param {string} purchaseDate - Purchase date (YYYY-MM-DD)
     * @param {number} salvageValue - Salvage value
     * @param {string} depreciationMethod - 'straight_line' | 'double_declining' | 'none'
     * @param {string|null} depStartDate - Depreciation start date (YYYY-MM-DD) or null
     * @param {boolean} isDepreciable - Whether the asset depreciates
     * @param {string|null} notes - Notes
     * @returns {number} New asset ID
     */
    addFixedAsset(name, purchaseCost, usefulLifeMonths, purchaseDate, salvageValue = 0, depreciationMethod = 'straight_line', depStartDate = null, isDepreciable = true, notes = null) {
        this.db.run(
            `INSERT INTO balance_sheet_assets (name, purchase_cost, useful_life_months, purchase_date, salvage_value, depreciation_method, dep_start_date, is_depreciable, notes)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
            [name.trim(), purchaseCost, usefulLifeMonths, purchaseDate, salvageValue, depreciationMethod, depStartDate, isDepreciable ? 1 : 0, notes]
        );
        const result = this.db.exec('SELECT last_insert_rowid() as id');
        this.autoSave();
        return result[0].values[0][0];
    },

    /**
     * Update a fixed asset
     * @param {number} id - Asset ID
     * @param {string} name - Asset name
     * @param {number} purchaseCost - Purchase cost
     * @param {number} usefulLifeMonths - Useful life in months
     * @param {string} purchaseDate - Purchase date (YYYY-MM-DD)
     * @param {number} salvageValue - Salvage value
     * @param {string} depreciationMethod - Depreciation method
     * @param {string|null} depStartDate - Depreciation start date
     * @param {boolean} isDepreciable - Whether the asset depreciates
     * @param {string|null} notes - Notes
     */
    updateFixedAsset(id, name, purchaseCost, usefulLifeMonths, purchaseDate, salvageValue = 0, depreciationMethod = 'straight_line', depStartDate = null, isDepreciable = true, notes = null) {
        this.db.run(
            `UPDATE balance_sheet_assets SET name = ?, purchase_cost = ?, useful_life_months = ?, purchase_date = ?,
             salvage_value = ?, depreciation_method = ?, dep_start_date = ?, is_depreciable = ?, notes = ? WHERE id = ?`,
            [name.trim(), purchaseCost, usefulLifeMonths, purchaseDate, salvageValue, depreciationMethod, depStartDate, isDepreciable ? 1 : 0, notes, id]
        );
        this.autoSave();
    },

    /**
     * Delete a fixed asset (and its linked transaction if any)
     * @param {number} id - Asset ID
     */
    deleteFixedAsset(id) {
        // Remove linked transaction
        this.db.run("DELETE FROM transactions WHERE source_type = 'asset_purchase' AND source_id = ?", [id]);
        this.db.run('DELETE FROM balance_sheet_assets WHERE id = ?', [id]);
        this.autoSave();
    },

    /**
     * Get a fixed asset by ID
     * @param {number} id - Asset ID
     * @returns {Object|null} Asset object
     */
    getFixedAssetById(id) {
        const results = this.db.exec('SELECT * FROM balance_sheet_assets WHERE id = ?', [id]);
        if (results.length === 0) return null;
        return this.rowsToObjects(results[0])[0];
    },

    /**
     * Link a transaction ID to a fixed asset
     * @param {number} assetId - Asset ID
     * @param {number} transactionId - Transaction ID
     */
    linkTransactionToAsset(assetId, transactionId) {
        this.db.run('UPDATE balance_sheet_assets SET linked_transaction_id = ? WHERE id = ?', [transactionId, assetId]);
        this.autoSave();
    },

    /**
     * Get aggregated asset depreciation by month.
     * For each depreciable asset, computes its schedule and aggregates.
     * @param {string|null} asOfMonth - If provided, only returns months <= asOfMonth
     * @returns {Object} Map of { [YYYY-MM]: totalDepreciation }
     */
    getAssetDepreciationByMonth(asOfMonth) {
        const assets = this.getFixedAssets();
        const result = {};

        assets.forEach(asset => {
            const schedule = Utils.computeDepreciationSchedule(asset);
            Object.entries(schedule).forEach(([month, amount]) => {
                if (asOfMonth && month > asOfMonth) return;
                result[month] = Math.round(((result[month] || 0) + amount) * 100) / 100;
            });
        });

        return result;
    },

    // ==================== LOANS ====================

    /**
     * Get all active loans
     * @returns {Array} Array of loan objects
     */
    getLoans() {
        const results = this.db.exec('SELECT * FROM loans WHERE is_active = 1 ORDER BY start_date ASC');
        if (results.length === 0) return [];
        return this.rowsToObjects(results[0]);
    },

    /**
     * Get a loan by ID
     * @param {number} id - Loan ID
     * @returns {Object|null} Loan object
     */
    getLoanById(id) {
        const results = this.db.exec('SELECT * FROM loans WHERE id = ?', [id]);
        if (results.length === 0) return null;
        return this.rowsToObjects(results[0])[0];
    },

    /**
     * Add a new loan
     * @param {Object} params - { name, principal, annual_rate, term_months, payments_per_year, start_date, notes }
     * @returns {number} New loan ID
     */
    addLoan(params) {
        this.db.run(
            `INSERT INTO loans (name, principal, annual_rate, term_months, payments_per_year, start_date, first_payment_date, notes)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
            [params.name.trim(), params.principal, params.annual_rate, params.term_months, params.payments_per_year || 12, params.start_date, params.first_payment_date || null, params.notes || null]
        );
        const result = this.db.exec('SELECT last_insert_rowid() as id');
        this.autoSave();
        return result[0].values[0][0];
    },

    /**
     * Update a loan
     * @param {number} id - Loan ID
     * @param {Object} params - Fields to update
     */
    updateLoan(id, params) {
        this.db.run(
            `UPDATE loans SET name = ?, principal = ?, annual_rate = ?, term_months = ?,
             payments_per_year = ?, start_date = ?, first_payment_date = ?, notes = ? WHERE id = ?`,
            [params.name.trim(), params.principal, params.annual_rate, params.term_months, params.payments_per_year || 12, params.start_date, params.first_payment_date || null, params.notes || null, id]
        );
        this.autoSave();
    },

    /**
     * Permanently delete a loan and all associated data
     * @param {number} id - Loan ID
     */
    deleteLoan(id) {
        this.db.run('DELETE FROM loans WHERE id = ?', [id]);
        this.db.run('DELETE FROM loan_skipped_payments WHERE loan_id = ?', [id]);
        this.db.run('DELETE FROM loan_payment_overrides WHERE loan_id = ?', [id]);
        this.deleteLoanTransactions(id);
        this.autoSave();
    },

    /**
     * Delete all journal transactions linked to a loan (receivable + payments)
     * @param {number} id - Loan ID
     */
    deleteLoanTransactions(id) {
        this.db.run(
            "DELETE FROM transactions WHERE source_type IN ('loan_receivable', 'loan_payment') AND source_id = ?",
            [id]
        );
    },

    /**
     * Get skipped payment numbers for a loan
     * @param {number} loanId
     * @returns {Set<number>} Set of skipped payment numbers
     */
    getLoanSkippedPayments(loanId) {
        return this.getSkippedPayments(loanId);
    },

    getSkippedPayments(loanId) {
        const results = this.db.exec('SELECT payment_number FROM loan_skipped_payments WHERE loan_id = ?', [loanId]);
        const set = new Set();
        if (results.length > 0) {
            results[0].values.forEach(row => set.add(row[0]));
        }
        return set;
    },

    /**
     * Toggle a loan payment as skipped/unskipped
     * @param {number} loanId
     * @param {number} paymentNumber
     */
    toggleSkipLoanPayment(loanId, paymentNumber) {
        const existing = this.db.exec(
            'SELECT 1 FROM loan_skipped_payments WHERE loan_id = ? AND payment_number = ?',
            [loanId, paymentNumber]
        );
        if (existing.length > 0 && existing[0].values.length > 0) {
            this.db.run('DELETE FROM loan_skipped_payments WHERE loan_id = ? AND payment_number = ?', [loanId, paymentNumber]);
        } else {
            this.db.run('INSERT INTO loan_skipped_payments (loan_id, payment_number) VALUES (?, ?)', [loanId, paymentNumber]);
        }
        this.autoSave();
    },

    /**
     * Get all payment overrides for a loan
     * @param {number} loanId
     * @returns {Object} Map of paymentNumber => override_amount
     */
    getLoanPaymentOverrides(loanId) {
        const results = this.db.exec('SELECT payment_number, override_amount FROM loan_payment_overrides WHERE loan_id = ?', [loanId]);
        const map = {};
        if (results.length > 0) {
            results[0].values.forEach(([num, amt]) => { map[num] = amt; });
        }
        return map;
    },

    /**
     * Set or remove a payment override
     * @param {number} loanId
     * @param {number} paymentNumber
     * @param {number|null} amount - null to remove override
     */
    setLoanPaymentOverride(loanId, paymentNumber, amount) {
        if (amount === null || amount === undefined) {
            this.db.run('DELETE FROM loan_payment_overrides WHERE loan_id = ? AND payment_number = ?', [loanId, paymentNumber]);
        } else {
            this.db.run(
                'INSERT OR REPLACE INTO loan_payment_overrides (loan_id, payment_number, override_amount) VALUES (?, ?, ?)',
                [loanId, paymentNumber, amount]
            );
        }
        this.autoSave();
    },

    /**
     * Get aggregated loan interest by month across all active loans.
     * @param {string|null} asOfMonth - If provided, only returns months <= asOfMonth
     * @returns {Object} Map of { [YYYY-MM]: totalInterest }
     */
    getLoanInterestByMonth(asOfMonth) {
        const loans = this.getLoans();
        const result = {};

        loans.forEach(loan => {
            const skipped = this.getSkippedPayments(loan.id);
            const overrides = this.getLoanPaymentOverrides(loan.id);
            const schedule = Utils.computeAmortizationSchedule({
                principal: loan.principal,
                annual_rate: loan.annual_rate,
                term_months: loan.term_months,
                payments_per_year: loan.payments_per_year,
                start_date: loan.start_date,
                first_payment_date: loan.first_payment_date
            }, skipped, overrides);

            schedule.forEach(entry => {
                if (asOfMonth && entry.month > asOfMonth) return;
                if (entry.skipped) return;
                result[entry.month] = Math.round(((result[entry.month] || 0) + entry.interest) * 100) / 100;
            });
        });

        return result;
    },

    // ==================== BUDGET EXPENSES ====================

    /**
     * Get all budget expenses
     * @returns {Array} Array of budget expense objects
     */
    getBudgetExpenses() {
        const results = this.db.exec(`
            SELECT be.*, c.name as category_name, bg.name as group_name, bg.sort_order as group_sort_order
            FROM budget_expenses be
            LEFT JOIN categories c ON be.category_id = c.id
            LEFT JOIN budget_groups bg ON be.group_id = bg.id
            ORDER BY bg.sort_order ASC, be.name ASC
        `);
        if (results.length === 0) return [];
        return this.rowsToObjects(results[0]);
    },

    /**
     * Get a budget expense by ID
     * @param {number} id
     * @returns {Object|null}
     */
    getBudgetExpenseById(id) {
        const results = this.db.exec(`
            SELECT be.*, c.name as category_name, bg.name as group_name
            FROM budget_expenses be
            LEFT JOIN categories c ON be.category_id = c.id
            LEFT JOIN budget_groups bg ON be.group_id = bg.id
            WHERE be.id = ?
        `, [id]);
        if (results.length === 0) return null;
        return this.rowsToObjects(results[0])[0];
    },

    /**
     * Add a budget expense
     * @param {string} name
     * @param {number} monthlyAmount
     * @param {string} startMonth - YYYY-MM
     * @param {string|null} endMonth - YYYY-MM or null for indefinite
     * @param {number|null} categoryId - FK to categories
     * @param {string|null} notes
     * @returns {number} New ID
     */
    addBudgetExpense(name, monthlyAmount, startMonth, endMonth = null, categoryId = null, notes = null, groupId = null) {
        this.db.run(
            'INSERT INTO budget_expenses (name, monthly_amount, start_month, end_month, category_id, notes, group_id) VALUES (?, ?, ?, ?, ?, ?, ?)',
            [name.trim(), monthlyAmount, startMonth, endMonth || null, categoryId || null, notes || null, groupId || null]
        );
        const result = this.db.exec('SELECT last_insert_rowid() as id');
        this.autoSave();
        return result[0].values[0][0];
    },

    /**
     * Update a budget expense
     * @param {number} id
     * @param {string} name
     * @param {number} monthlyAmount
     * @param {string} startMonth
     * @param {string|null} endMonth
     * @param {number|null} categoryId
     * @param {string|null} notes
     */
    updateBudgetExpense(id, name, monthlyAmount, startMonth, endMonth = null, categoryId = null, notes = null, groupId = null) {
        this.db.run(
            'UPDATE budget_expenses SET name = ?, monthly_amount = ?, start_month = ?, end_month = ?, category_id = ?, notes = ?, group_id = ? WHERE id = ?',
            [name.trim(), monthlyAmount, startMonth, endMonth || null, categoryId || null, notes || null, groupId || null, id]
        );
        this.autoSave();
    },

    /**
     * Delete a budget expense
     * @param {number} id
     */
    deleteBudgetExpense(id) {
        this.db.run('DELETE FROM budget_expenses WHERE id = ?', [id]);
        this.autoSave();
    },

    /**
     * Get active budget expenses for a given month
     * @param {string} month - YYYY-MM
     * @returns {Array} Expenses where start_month <= month AND (end_month IS NULL OR end_month >= month)
     */
    getActiveBudgetExpensesForMonth(month) {
        const results = this.db.exec(`
            SELECT be.*, c.name as category_name
            FROM budget_expenses be
            LEFT JOIN categories c ON be.category_id = c.id
            WHERE be.start_month <= ? AND (be.end_month IS NULL OR be.end_month >= ?)
            ORDER BY be.name ASC
        `, [month, month]);
        if (results.length === 0) return [];
        return this.rowsToObjects(results[0]);
    },

    // ==================== BUDGET GROUPS ====================

    getBudgetGroups() {
        const results = this.db.exec('SELECT * FROM budget_groups ORDER BY sort_order ASC, name ASC');
        if (results.length === 0) return [];
        return this.rowsToObjects(results[0]);
    },

    addBudgetGroup(name) {
        const maxOrder = this.db.exec('SELECT COALESCE(MAX(sort_order), -1) + 1 as next_order FROM budget_groups');
        const nextOrder = maxOrder.length > 0 ? maxOrder[0].values[0][0] : 0;
        this.db.run('INSERT INTO budget_groups (name, sort_order) VALUES (?, ?)', [name.trim(), nextOrder]);
        const result = this.db.exec('SELECT last_insert_rowid() as id');
        this.autoSave();
        return result[0].values[0][0];
    },

    updateBudgetGroup(id, name) {
        this.db.run('UPDATE budget_groups SET name = ? WHERE id = ?', [name.trim(), id]);
        this.autoSave();
    },

    deleteBudgetGroup(id) {
        this.db.run('UPDATE budget_expenses SET group_id = NULL WHERE group_id = ?', [id]);
        this.db.run('DELETE FROM budget_groups WHERE id = ?', [id]);
        this.autoSave();
    },

    updateBudgetGroupOrder(orderedIds) {
        orderedIds.forEach((id, index) => {
            this.db.run('UPDATE budget_groups SET sort_order = ? WHERE id = ?', [index, id]);
        });
        this.autoSave();
    },

    moveBudgetExpenseToGroup(expenseId, groupId) {
        this.db.run('UPDATE budget_expenses SET group_id = ? WHERE id = ?', [groupId || null, expenseId]);
        this.autoSave();
    },

    // ==================== PRODUCTS ====================

    /**
     * Add a product to the catalog
     */
    addProduct(name, sku, price, taxRate, cogs, notes) {
        this.db.run(
            'INSERT INTO products (name, sku, price, tax_rate, cogs, notes) VALUES (?, ?, ?, ?, ?, ?)',
            [name.trim(), sku ? sku.trim() : null, price, taxRate || 0, cogs || 0, notes || null]
        );
        const result = this.db.exec('SELECT last_insert_rowid() as id');
        this.autoSave();
        return result[0].values[0][0];
    },

    /**
     * Update a product
     */
    updateProduct(id, name, sku, price, taxRate, cogs, notes) {
        this.db.run(
            'UPDATE products SET name = ?, sku = ?, price = ?, tax_rate = ?, cogs = ?, notes = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
            [name.trim(), sku ? sku.trim() : null, price, taxRate || 0, cogs || 0, notes || null, id]
        );
        this.autoSave();
    },

    /**
     * Toggle a product's discontinued status
     */
    toggleProductDiscontinued(id) {
        this.db.run('UPDATE products SET is_discontinued = CASE WHEN is_discontinued = 0 THEN 1 ELSE 0 END, updated_at = CURRENT_TIMESTAMP WHERE id = ?', [id]);
        this.autoSave();
    },

    /**
     * Delete a product
     */
    deleteProduct(id) {
        this.db.run('DELETE FROM product_ve_mappings WHERE product_id = ?', [id]);
        this.db.run('DELETE FROM products WHERE id = ?', [id]);
        this.autoSave();
    },

    /**
     * Get all products, ordered by discontinued status then name
     */
    getProducts() {
        const results = this.db.exec('SELECT * FROM products ORDER BY is_discontinued ASC, name ASC');
        if (results.length === 0) return [];
        return this.rowsToObjects(results[0]);
    },

    /**
     * Get a single product by ID
     */
    getProductById(id) {
        const results = this.db.exec('SELECT * FROM products WHERE id = ?', [id]);
        if (results.length === 0 || results[0].values.length === 0) return null;
        return this.rowsToObjects(results[0])[0];
    },

    // ==================== PRODUCT-VE MAPPINGS ====================

    /** Returns sorted unique VE item names from ve_sale_items */
    getDistinctVeItemNames() {
        const results = this.db.exec(
            'SELECT DISTINCT name FROM ve_sale_items WHERE name IS NOT NULL ORDER BY name ASC'
        );
        if (results.length === 0) return [];
        return results[0].values.map(r => r[0]);
    },

    /** Returns sorted distinct (name, price) pairs from ve_sale_items */
    getDistinctVeItemNamesWithPrice() {
        const results = this.db.exec(
            'SELECT DISTINCT name, ROUND(price, 2) as price FROM ve_sale_items WHERE name IS NOT NULL ORDER BY name ASC, price ASC'
        );
        if (results.length === 0) return [];
        return results[0].values.map(r => ({ name: r[0], price: r[1] }));
    },

    /** Returns {name, price} objects currently mapped to a product */
    getMappingsForProduct(productId) {
        const results = this.db.exec(
            'SELECT ve_item_name, ve_item_price FROM product_ve_mappings WHERE product_id = ? ORDER BY ve_item_name ASC',
            [productId]
        );
        if (results.length === 0) return [];
        return results[0].values.map(r => ({ name: r[0], price: r[1] || 0 }));
    },

    /** Returns a Set of "name|price.toFixed(2)" keys for all mapped VE items */
    getMappedVeItemNames() {
        const results = this.db.exec('SELECT DISTINCT ve_item_name, ve_item_price FROM product_ve_mappings');
        if (results.length === 0) return new Set();
        return new Set(results[0].values.map(r => r[0] + '|' + (r[1] || 0).toFixed(2)));
    },

    /**
     * Replace all mappings for a product with a new set of {name, price} objects.
     * Deletes existing then inserts new ones.
     */
    setMappingsForProduct(productId, items) {
        this.db.run('DELETE FROM product_ve_mappings WHERE product_id = ?', [productId]);
        if (items && items.length > 0) {
            const stmt = this.db.prepare(
                'INSERT OR IGNORE INTO product_ve_mappings (product_id, ve_item_name, ve_item_price) VALUES (?, ?, ?)'
            );
            for (const { name, price } of items) {
                stmt.run([productId, name, price || 0]);
            }
            stmt.free();
        }
        this.autoSave();
    },

    /**
     * Compute linked analytics for the Products tab.
     * Returns { totals, byProduct, monthlyCogs }
     */
    getLinkedProductAnalytics(dateFrom, dateTo, source) {
        const empty = { totals: { pretax_total: 0, sales_tax: 0, post_tax_total: 0, discount: 0, pretax_after_discount: 0, posttax_after_discount: 0 }, byProduct: [], monthlyCogs: [] };
        const hasMappings = this.db.exec('SELECT 1 FROM product_ve_mappings LIMIT 1');
        if (hasMappings.length === 0) return empty;

        const conditions = [];
        const params = [];
        if (dateFrom) { conditions.push('vs.date >= ?'); params.push(dateFrom); }
        if (dateTo)   { conditions.push('vs.date <= ?'); params.push(dateTo); }
        if (source && source !== 'all') { conditions.push('vs.source = ?'); params.push(source); }
        const whereFrag = conditions.length > 0 ? 'AND ' + conditions.join(' AND ') : '';

        // Totals
        const totalsResult = this.db.exec(`
            SELECT
                COALESCE(SUM(vi.amount), 0) AS pretax_total,
                COALESCE(SUM(CASE WHEN vs.subtotal > 0 THEN vs.tax * (vi.amount / vs.subtotal) ELSE 0 END), 0) AS sales_tax,
                COALESCE(SUM(CASE WHEN vs.subtotal > 0 THEN vs.discount * (vi.amount / vs.subtotal) ELSE 0 END), 0) AS discount
            FROM ve_sale_items vi
            JOIN ve_sales vs ON vs.transaction_no = vi.transaction_no
            JOIN product_ve_mappings m ON m.ve_item_name = vi.name AND ROUND(m.ve_item_price, 2) = ROUND(vi.price, 2)
            WHERE 1=1 ${whereFrag}
        `, params);

        let totals = { pretax_total: 0, sales_tax: 0, post_tax_total: 0, discount: 0, pretax_after_discount: 0, posttax_after_discount: 0 };
        if (totalsResult.length > 0 && totalsResult[0].values.length > 0) {
            const pt = totalsResult[0].values[0][0] || 0;
            const st = totalsResult[0].values[0][1] || 0;
            const disc = totalsResult[0].values[0][2] || 0;
            const pretaxAfterDisc = pt - disc;
            const taxAfterDisc = st > 0 && pt > 0 ? st * (pretaxAfterDisc / pt) : 0;
            totals = { pretax_total: pt, sales_tax: st, post_tax_total: pt + st, discount: disc, pretax_after_discount: pretaxAfterDisc, posttax_after_discount: pretaxAfterDisc + taxAfterDisc };
        }

        // By-product aggregates
        const byProductResult = this.db.exec(`
            SELECT p.id, p.name, p.cogs,
                   COALESCE(SUM(vi.quantity), 0) AS units_sold,
                   COALESCE(SUM(vi.amount), 0) AS revenue
            FROM ve_sale_items vi
            JOIN product_ve_mappings m ON m.ve_item_name = vi.name AND ROUND(m.ve_item_price, 2) = ROUND(vi.price, 2)
            JOIN products p ON p.id = m.product_id
            JOIN ve_sales vs ON vs.transaction_no = vi.transaction_no
            WHERE 1=1 ${whereFrag}
            GROUP BY p.id, p.name, p.cogs
            ORDER BY revenue DESC
        `, params);
        const byProduct = byProductResult.length > 0 ? this.rowsToObjects(byProductResult[0]) : [];

        // Monthly COGS pivot rows
        let monthlyCogs = [];
        if (byProduct.length > 0) {
            const mcResult = this.db.exec(`
                SELECT strftime('%Y-%m', vs.date) AS month,
                       p.id AS product_id, p.name AS product_name, p.cogs,
                       COALESCE(SUM(vi.quantity), 0) AS qty_sold
                FROM ve_sale_items vi
                JOIN product_ve_mappings m ON m.ve_item_name = vi.name AND ROUND(m.ve_item_price, 2) = ROUND(vi.price, 2)
                JOIN products p ON p.id = m.product_id
                JOIN ve_sales vs ON vs.transaction_no = vi.transaction_no
                WHERE 1=1 ${whereFrag}
                GROUP BY month, p.id
                ORDER BY month ASC, p.name ASC
            `, params);
            if (mcResult.length > 0) {
                monthlyCogs = this.rowsToObjects(mcResult[0]);
            }
        }

        return { totals, byProduct, monthlyCogs };
    },

    // ==================== BREAK-EVEN CONFIG ====================

    /**
     * Get default break-even configuration
     * @returns {Object} Default config with consumer, b2b, timeline, and cost-source flags
     */
    _defaultBreakevenConfig() {
        return {
            timeline: { start: null, end: null },
            dataSource: 'actual',
            asOfMonth: null,
            fixedCostOverride: null,
            consumer: { enabled: true, avgPrice: 0, avgCogs: 0 },
            b2b: { enabled: false, monthlyUnits: 0, ratePerUnit: 0, cogsPerUnit: 0 },
            unitIncrement: 100
        };
    },

    /**
     * Get break-even configuration, merged with defaults
     * @returns {Object} Break-even config
     */
    getBreakevenConfig() {
        const result = this.db.exec("SELECT value FROM app_meta WHERE key = 'breakeven_config'");
        if (result.length === 0 || result[0].values.length === 0) {
            return this._defaultBreakevenConfig();
        }
        try {
            return Object.assign(this._defaultBreakevenConfig(), JSON.parse(result[0].values[0][0]));
        } catch (e) {
            return this._defaultBreakevenConfig();
        }
    },

    /**
     * Save break-even configuration
     * @param {Object} config - Break-even config to persist
     */
    setBreakevenConfig(config) {
        this.db.run(
            "INSERT OR REPLACE INTO app_meta (key, value) VALUES ('breakeven_config', ?)",
            [JSON.stringify(config)]
        );
        this.autoSave();
    },

    /**
     * Get total monthly fixed costs from budget expenses active in a given month
     * @param {string} month - YYYY-MM
     * @returns {number}
     */
    getBudgetFixedCostsForMonth(month) {
        const result = this.db.exec(`
            SELECT COALESCE(SUM(monthly_amount), 0) as total
            FROM budget_expenses
            WHERE start_month <= ? AND (end_month IS NULL OR end_month >= ?)
        `, [month, month]);
        return result[0].values[0][0] || 0;
    },

    // ==================== PROJECTED SALES ====================

    _defaultProjectedSalesConfig() {
        return {
            enabled: false,
            projectionStartMonth: null,
            viewMode: 'projected',
            salesTaxRate: 0,
            channels: {
                online: { enabled: true, avgPrice: 0, avgCogs: 0, units: {} },
                tradeshow: { enabled: false, avgPrice: 0, avgCogs: 0, units: {} }
            }
        };
    },

    getProjectedSalesConfig() {
        const result = this.db.exec("SELECT value FROM app_meta WHERE key = 'projected_sales_config'");
        if (result.length === 0 || result[0].values.length === 0) {
            return this._defaultProjectedSalesConfig();
        }
        try {
            const saved = JSON.parse(result[0].values[0][0]);
            const defaults = this._defaultProjectedSalesConfig();
            const merged = Object.assign({}, defaults, saved);
            merged.channels = Object.assign({}, defaults.channels);
            if (saved.channels) {
                if (saved.channels.online) {
                    merged.channels.online = Object.assign({}, defaults.channels.online, saved.channels.online);
                }
                if (saved.channels.tradeshow) {
                    merged.channels.tradeshow = Object.assign({}, defaults.channels.tradeshow, saved.channels.tradeshow);
                }
            }
            return merged;
        } catch (e) {
            return this._defaultProjectedSalesConfig();
        }
    },

    setProjectedSalesConfig(config) {
        this.db.run(
            "INSERT OR REPLACE INTO app_meta (key, value) VALUES ('projected_sales_config', ?)",
            [JSON.stringify(config)]
        );
        this.autoSave();
    },

    /**
     * Compute projected sales spreadsheet data for given months
     * @param {Object} config - Projected sales config
     * @param {Array} months - Array of YYYY-MM strings
     * @returns {Object} { byMonth, channels, config }
     */
    getProjectedSalesSpreadsheet(config, months) {
        config = config || this.getProjectedSalesConfig();
        if (!config.enabled || !config.projectionStartMonth) {
            return { byMonth: {}, channels: config.channels, config };
        }
        const byMonth = {};
        (months || []).forEach(m => {
            if (m < config.projectionStartMonth) return;
            const entry = {
                revenue: 0, cogs: 0, salesTax: 0,
                onlineUnits: 0, onlineRevenue: 0, onlineCogs: 0,
                tradeshowUnits: 0, tradeshowRevenue: 0, tradeshowCogs: 0
            };
            ['online', 'tradeshow'].forEach(key => {
                const ch = config.channels[key];
                if (!ch || !ch.enabled) return;
                const units = (ch.units && ch.units[m]) || 0;
                const rev = units * (ch.avgPrice || 0);
                const cg = units * (ch.avgCogs || 0);
                entry[key + 'Units'] = units;
                entry[key + 'Revenue'] = rev;
                entry[key + 'Cogs'] = cg;
                entry.revenue += rev;
                entry.cogs += cg;
            });
            // Sales tax on total non-B2B revenue
            const taxRate = (config.salesTaxRate || 0) / 100;
            entry.salesTax = Math.round(entry.revenue * taxRate * 100) / 100;
            byMonth[m] = entry;
        });
        return { byMonth, channels: config.channels, config };
    },

    // ==================== EQUITY & LOAN CONFIG ====================

    /**
     * Get equity config (common stock par, shares, APIC)
     * @returns {Object} { common_stock_par: number, common_stock_shares: number, apic: number }
     */
    getEquityConfig() {
        const result = this.db.exec("SELECT value FROM app_meta WHERE key = 'equity_config'");
        if (result.length === 0 || result[0].values.length === 0) {
            return { common_stock_par: 0, common_stock_shares: 0, apic: 0 };
        }
        try {
            return JSON.parse(result[0].values[0][0]);
        } catch (e) {
            return { common_stock_par: 0, common_stock_shares: 0, apic: 0 };
        }
    },

    /**
     * Set equity config
     * @param {Object} config - { common_stock_par, common_stock_shares, apic }
     */
    setEquityConfig(config) {
        this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('equity_config', ?)", [JSON.stringify(config)]);
        this.autoSave();
    },

    /**
     * Get loan config (single loan)
     * @returns {Object|null} { principal, annual_rate, term_years, payments_per_year, start_date } or null
     */
    getLoanConfig() {
        const result = this.db.exec("SELECT value FROM app_meta WHERE key = 'loan_config'");
        if (result.length === 0 || result[0].values.length === 0) return null;
        try {
            return JSON.parse(result[0].values[0][0]);
        } catch (e) {
            return null;
        }
    },

    /**
     * Set loan config (kept for backward compat)
     * @param {Object|null} config - Loan config object or null to clear
     */
    setLoanConfig(config) {
        if (config === null) {
            this.db.run("DELETE FROM app_meta WHERE key = 'loan_config'");
        } else {
            this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('loan_config', ?)", [JSON.stringify(config)]);
        }
        this.autoSave();
    },

    // ==================== TAB ORDER ====================

    /**
     * Get saved tab order
     * @returns {Array|null} Array of tab names or null if not set
     */
    getTabOrder() {
        const result = this.db.exec("SELECT value FROM app_meta WHERE key = 'tab_order'");
        if (result.length === 0 || result[0].values.length === 0) return null;
        try {
            return JSON.parse(result[0].values[0][0]);
        } catch (e) {
            return null;
        }
    },

    /**
     * Save tab order
     * @param {Array} order - Array of tab name strings
     */
    setTabOrder(order) {
        this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('tab_order', ?)", [JSON.stringify(order)]);
        this.autoSave();
    },

    // ==================== HIDDEN TABS ====================

    getHiddenTabs() {
        const result = this.db.exec("SELECT value FROM app_meta WHERE key = 'hidden_tabs'");
        if (result.length === 0 || result[0].values.length === 0) return [];
        try {
            return JSON.parse(result[0].values[0][0]);
        } catch (e) {
            return [];
        }
    },

    setHiddenTabs(tabs) {
        this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('hidden_tabs', ?)", [JSON.stringify(tabs)]);
        this.autoSave();
    },

    // ==================== TAB RESET ====================

    resetTabData(tabName) {
        switch (tabName) {
            case 'journal':
                this.db.run('DELETE FROM transactions');
                this.db.run('DELETE FROM categories');
                this.db.run('DELETE FROM category_folders');
                break;
            case 'cashflow':
                this.db.run('DELETE FROM cashflow_overrides');
                break;
            case 'pnl':
                this.db.run('DELETE FROM pl_overrides');
                break;
            case 'assets':
                this.db.run("DELETE FROM transactions WHERE source_type = 'asset_purchase'");
                this.db.run('DELETE FROM balance_sheet_assets');
                break;
            case 'loan':
                this.db.run("DELETE FROM transactions WHERE source_type IN ('loan_receivable', 'loan_payment')");
                this.db.run('DELETE FROM loan_payment_overrides');
                this.db.run('DELETE FROM loan_skipped_payments');
                this.db.run('DELETE FROM loans');
                break;
            case 'budget':
                this.db.run('DELETE FROM budget_expenses');
                this.db.run('DELETE FROM budget_groups');
                break;
            case 'breakeven':
                this.db.run("DELETE FROM app_meta WHERE key = 'breakeven_config'");
                break;
            case 'projectedsales':
                this.db.run("DELETE FROM app_meta WHERE key = 'projected_sales_config'");
                break;
            case 'products':
                this.db.run('DELETE FROM product_ve_mappings');
                this.db.run('DELETE FROM products');
                break;
            case 'vesales':
                this.db.run('DELETE FROM ve_sale_items');
                this.db.run('DELETE FROM ve_sales');
                break;
        }
        this.autoSave();
    },

    // ==================== TIMELINE ====================

    /**
     * Get timeline settings
     * @returns {Object} { start: 'YYYY-MM' | null, end: 'YYYY-MM' | null }
     */
    getTimeline() {
        const startResult = this.db.exec("SELECT value FROM app_meta WHERE key = 'timeline_start'");
        const endResult = this.db.exec("SELECT value FROM app_meta WHERE key = 'timeline_end'");
        return {
            start: (startResult.length > 0 && startResult[0].values.length > 0) ? startResult[0].values[0][0] : null,
            end: (endResult.length > 0 && endResult[0].values.length > 0) ? endResult[0].values[0][0] : null
        };
    },

    /**
     * Set timeline start month
     * @param {string|null} month - 'YYYY-MM' or null to clear
     */
    setTimelineStart(month) {
        if (month) {
            this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('timeline_start', ?)", [month]);
        } else {
            this.db.run("DELETE FROM app_meta WHERE key = 'timeline_start'");
        }
        this.autoSave();
    },

    /**
     * Set timeline end month
     * @param {string|null} month - 'YYYY-MM' or null to clear
     */
    setTimelineEnd(month) {
        if (month) {
            this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('timeline_end', ?)", [month]);
        } else {
            this.db.run("DELETE FROM app_meta WHERE key = 'timeline_end'");
        }
        this.autoSave();
    },

    // ==================== CASH FLOW OVERRIDES ====================

    /**
     * Get all cash flow overrides
     * @returns {Object} Map of "categoryId-month" => override_amount
     */
    getAllCashFlowOverrides() {
        const results = this.db.exec('SELECT category_id, month, override_amount FROM cashflow_overrides');
        if (results.length === 0) return {};
        const overrides = {};
        this.rowsToObjects(results[0]).forEach(row => {
            overrides[`${row.category_id}-${row.month}`] = row.override_amount;
        });
        return overrides;
    },

    /**
     * Set a cash flow override value for a category+month
     * @param {number} categoryId - Category ID
     * @param {string} month - Month (YYYY-MM)
     * @param {number|null} amount - Override amount (null to remove)
     */
    setCashFlowOverride(categoryId, month, amount) {
        if (amount === null || amount === '') {
            this.db.run('DELETE FROM cashflow_overrides WHERE category_id = ? AND month = ?', [categoryId, month]);
        } else {
            this.db.run(
                'INSERT OR REPLACE INTO cashflow_overrides (category_id, month, override_amount) VALUES (?, ?, ?)',
                [categoryId, month, parseFloat(amount)]
            );
        }
        this.autoSave();
    },

    // ==================== BALANCE SHEET QUERIES ====================

    /**
     * Get cash balance as of a given month (sum of received - sum of paid where month_paid <= asOfMonth)
     * @param {string} asOfMonth - Month in YYYY-MM format
     * @returns {number} Cash balance
     */
    getCashAsOf(asOfMonth) {
        const result = this.db.exec(`
            SELECT
                COALESCE(SUM(CASE WHEN transaction_type = 'receivable' AND status = 'received' THEN amount ELSE 0 END), 0)
                - COALESCE(SUM(CASE WHEN transaction_type = 'payable' AND status = 'paid' THEN amount ELSE 0 END), 0) as cash
            FROM transactions
            WHERE month_paid IS NOT NULL AND month_paid <= ?
        `, [asOfMonth]);
        return result[0].values[0][0];
    },

    /**
     * Get accounts receivable as of a given month
     * (receivable transactions where month_due <= asOfMonth AND still unpaid as of that month)
     * @param {string} asOfMonth - Month in YYYY-MM format
     * @returns {number} AR balance
     */
    getAccountsReceivableAsOf(asOfMonth) {
        const result = this.db.exec(`
            SELECT COALESCE(SUM(amount), 0) as total
            FROM transactions
            WHERE transaction_type = 'receivable'
            AND month_due IS NOT NULL
            AND month_due <= ?
            AND (status = 'pending' OR (status = 'received' AND month_paid > ?))
        `, [asOfMonth, asOfMonth]);
        return result[0].values[0][0];
    },

    /**
     * Get accounts payable as of a given month (non-sales-tax categories)
     * @param {string} asOfMonth - Month in YYYY-MM format
     * @returns {number} AP balance
     */
    getAccountsPayableAsOf(asOfMonth) {
        const result = this.db.exec(`
            SELECT COALESCE(SUM(t.amount), 0) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.transaction_type = 'payable'
            AND c.is_sales_tax = 0
            AND t.month_due IS NOT NULL
            AND t.month_due <= ?
            AND (t.status = 'pending' OR (t.status = 'paid' AND t.month_paid > ?))
        `, [asOfMonth, asOfMonth]);
        return result[0].values[0][0];
    },

    /**
     * Get sales tax payable as of a given month (only sales tax categories)
     * @param {string} asOfMonth - Month in YYYY-MM format
     * @returns {number} Sales tax payable balance
     */
    getSalesTaxPayableAsOf(asOfMonth) {
        const result = this.db.exec(`
            SELECT COALESCE(SUM(t.amount), 0) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.transaction_type = 'payable'
            AND c.is_sales_tax = 1
            AND t.month_due IS NOT NULL
            AND t.month_due <= ?
            AND (t.status = 'pending' OR (t.status = 'paid' AND t.month_paid > ?))
        `, [asOfMonth, asOfMonth]);
        return result[0].values[0][0];
    },

    /**
     * Get accounts receivable broken down by category as of a given month
     * @param {string} asOfMonth - Month in YYYY-MM format
     * @returns {Array} [{category_id, category_name, total}]
     */
    getARByCategory(asOfMonth) {
        const results = this.db.exec(`
            SELECT c.id as category_id, c.name as category_name,
                   COALESCE(SUM(t.amount), 0) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.transaction_type = 'receivable'
            AND t.month_due IS NOT NULL
            AND t.month_due <= ?
            AND (t.status = 'pending' OR (t.status = 'received' AND t.month_paid > ?))
            GROUP BY c.id, c.name
            HAVING total > 0
            ORDER BY c.name ASC
        `, [asOfMonth, asOfMonth]);
        if (results.length === 0) return [];
        return this.rowsToObjects(results[0]);
    },

    /**
     * Get accounts payable broken down by category as of a given month (excludes sales tax)
     * @param {string} asOfMonth - Month in YYYY-MM format
     * @returns {Array} [{category_id, category_name, total}]
     */
    getAPByCategory(asOfMonth) {
        const results = this.db.exec(`
            SELECT c.id as category_id, c.name as category_name,
                   COALESCE(SUM(t.amount), 0) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.transaction_type = 'payable'
            AND c.is_sales_tax = 0
            AND t.month_due IS NOT NULL
            AND t.month_due <= ?
            AND (t.status = 'pending' OR (t.status = 'paid' AND t.month_paid > ?))
            GROUP BY c.id, c.name
            HAVING total > 0
            ORDER BY c.name ASC
        `, [asOfMonth, asOfMonth]);
        if (results.length === 0) return [];
        return this.rowsToObjects(results[0]);
    },

    /**
     * Get retained earnings as of a given month (cumulative P&L net income through asOfMonth).
     * This recomputes P&L with overrides to match what the P&L statement shows.
     * @param {string} asOfMonth - Month in YYYY-MM format
     * @param {string} taxMode - 'corporate' or 'passthrough'
     * @returns {number} Retained earnings (cumulative net income after tax)
     */
    getRetainedEarningsAsOf(asOfMonth, taxMode) {
        // Get all months up to and including asOfMonth (from transactions)
        const monthsResult = this.db.exec(`
            SELECT DISTINCT t.month_due as month FROM transactions t
            WHERE t.month_due IS NOT NULL AND t.month_due <= ?
            ORDER BY month ASC
        `, [asOfMonth]);
        const txMonths = monthsResult.length > 0 ? monthsResult[0].values.map(r => r[0]) : [];

        // Also include months from asset depreciation and loan interest
        const assetDeprByMonth = this.getAssetDepreciationByMonth(asOfMonth);
        const loanInterestByMonth = this.getLoanInterestByMonth(asOfMonth);

        const allMonths = new Set(txMonths);
        Object.keys(assetDeprByMonth).forEach(m => allMonths.add(m));
        Object.keys(loanInterestByMonth).forEach(m => allMonths.add(m));
        const months = Array.from(allMonths).sort();

        if (months.length === 0) return 0;

        const overrides = this.getAllPLOverrides();

        // Helper: get value with override
        const getVal = (catId, month, computed) => {
            const key = `${catId}-${month}`;
            return (key in overrides) ? overrides[key] : computed;
        };

        // Revenue per category per month (accrual, pretax)
        const revenueResult = this.db.exec(`
            SELECT c.id as category_id, t.month_due as month,
                   SUM(COALESCE(t.pretax_amount, t.amount)) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.month_due IS NOT NULL AND t.month_due <= ?
            AND t.transaction_type = 'receivable'
            AND c.is_cogs = 0 AND c.show_on_pl != 1
            AND (c.is_b2b = 1 OR c.is_sales = 1)
            AND COALESCE(t.source_type, '') NOT IN ('loan_receivable', 'loan_payment')
            GROUP BY c.id, t.month_due
        `, [asOfMonth]);
        const revenue = revenueResult.length > 0 ? this.rowsToObjects(revenueResult[0]) : [];

        // COGS
        const cogsResult = this.db.exec(`
            SELECT c.id as category_id, t.month_due as month, SUM(t.amount) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.month_due IS NOT NULL AND t.month_due <= ?
            AND c.is_cogs = 1 AND c.show_on_pl != 1
            GROUP BY c.id, t.month_due
        `, [asOfMonth]);
        const cogs = cogsResult.length > 0 ? this.rowsToObjects(cogsResult[0]) : [];

        // OpEx (non-COGS, non-depreciation, non-sales-tax, non-loan payables)
        const opexResult = this.db.exec(`
            SELECT c.id as category_id, t.month_due as month, SUM(t.amount) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.month_due IS NOT NULL AND t.month_due <= ?
            AND t.transaction_type = 'payable'
            AND c.is_cogs = 0 AND c.is_depreciation = 0 AND c.is_sales_tax = 0 AND c.show_on_pl != 1
            AND COALESCE(t.source_type, '') NOT IN ('loan_payment', 'loan_receivable')
            GROUP BY c.id, t.month_due
        `, [asOfMonth]);
        const opex = opexResult.length > 0 ? this.rowsToObjects(opexResult[0]) : [];

        // Depreciation categories (values from overrides only)
        const deprResult = this.db.exec(`SELECT id as category_id FROM categories WHERE is_depreciation = 1`);
        const deprCats = deprResult.length > 0 ? deprResult[0].values.map(r => r[0]) : [];

        // Build lookup maps
        const buildMap = (rows) => {
            const map = {};
            rows.forEach(r => {
                const key = `${r.category_id}-${r.month}`;
                map[key] = (map[key] || 0) + r.total;
            });
            return map;
        };

        const revMap = buildMap(revenue);
        const cogsMap = buildMap(cogs);
        const opexMap = buildMap(opex);

        // Compute cumulative net income
        const round2 = (v) => Math.round(v * 100) / 100;
        let cumulative = 0;

        // Get unique category IDs per section from transactions
        const revCatIds = [...new Set(revenue.map(r => r.category_id))];
        const cogsCatIds = [...new Set(cogs.map(r => r.category_id))];
        const opexCatIds = [...new Set(opex.map(r => r.category_id))];

        // Also include categories that have overrides but no transactions
        // (override-only categories would otherwise be missed)
        Object.keys(overrides).forEach(key => {
            const [catIdStr] = key.split('-');
            const catId = parseInt(catIdStr);
            if (catId < 0) return; // skip tax override key (-1)
            if (!revCatIds.includes(catId) && !cogsCatIds.includes(catId) &&
                !opexCatIds.includes(catId) && !deprCats.includes(catId)) {
                // Check if this category is an opex category (non-hidden, non-cogs, non-depr, non-sales-tax payable-eligible)
                const catResult = this.db.exec(`
                    SELECT id FROM categories
                    WHERE id = ? AND is_cogs = 0 AND is_depreciation = 0 AND is_sales_tax = 0 AND show_on_pl != 1
                `, [catId]);
                if (catResult.length > 0 && catResult[0].values.length > 0) {
                    opexCatIds.push(catId);
                }
            }
        });

        months.forEach(month => {
            let monthRev = 0;
            revCatIds.forEach(catId => {
                monthRev += getVal(catId, month, revMap[`${catId}-${month}`] || 0);
            });

            let monthCogs = 0;
            cogsCatIds.forEach(catId => {
                monthCogs += getVal(catId, month, cogsMap[`${catId}-${month}`] || 0);
            });

            let monthOpex = 0;
            opexCatIds.forEach(catId => {
                monthOpex += getVal(catId, month, opexMap[`${catId}-${month}`] || 0);
            });

            // Depreciation from overrides (manual depreciation categories)
            deprCats.forEach(catId => {
                monthOpex += getVal(catId, month, 0);
            });

            // Asset depreciation (computed from fixed assets)
            if (assetDeprByMonth[month]) {
                monthOpex += assetDeprByMonth[month];
            }

            // Loan interest (computed from active loans)
            if (loanInterestByMonth[month]) {
                monthOpex += loanInterestByMonth[month];
            }

            const nibt = round2(monthRev - monthCogs - monthOpex);

            let tax = 0;
            if (taxMode === 'corporate') {
                const autoTax = round2(nibt > 0 ? nibt * 0.21 : 0);
                tax = getVal(-1, month, autoTax);
            }

            cumulative = round2(cumulative + nibt - tax);
        });

        return cumulative;
    },

    /**
     * Get cumulative P&L totals for all months through asOfMonth.
     * Mirrors getRetainedEarningsAsOf logic but returns full breakdown for financial ratios.
     */
    getPLTotalsThrough(asOfMonth, taxMode) {
        const round2 = (v) => Math.round(v * 100) / 100;
        const monthsResult = this.db.exec(`
            SELECT DISTINCT t.month_due as month FROM transactions t
            WHERE t.month_due IS NOT NULL AND t.month_due <= ?
            ORDER BY month ASC
        `, [asOfMonth]);
        const txMonths = monthsResult.length > 0 ? monthsResult[0].values.map(r => r[0]) : [];

        const assetDeprByMonth = this.getAssetDepreciationByMonth(asOfMonth);
        const loanInterestByMonth = this.getLoanInterestByMonth(asOfMonth);

        const allMonths = new Set(txMonths);
        Object.keys(assetDeprByMonth).forEach(m => allMonths.add(m));
        Object.keys(loanInterestByMonth).forEach(m => allMonths.add(m));
        const months = Array.from(allMonths).sort();

        const zero = { totalRevenue: 0, totalCogs: 0, totalGP: 0, totalNIBT: 0, totalTax: 0, totalNIAT: 0, totalLoanInterest: 0 };
        if (months.length === 0) return zero;

        const overrides = this.getAllPLOverrides();
        const getVal = (catId, month, computed) => {
            const key = `${catId}-${month}`;
            return (key in overrides) ? overrides[key] : computed;
        };

        const revenueResult = this.db.exec(`
            SELECT c.id as category_id, t.month_due as month,
                   SUM(COALESCE(t.pretax_amount, t.amount)) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.month_due IS NOT NULL AND t.month_due <= ?
            AND t.transaction_type = 'receivable'
            AND c.is_cogs = 0 AND c.show_on_pl != 1
            AND (c.is_b2b = 1 OR c.is_sales = 1)
            AND COALESCE(t.source_type, '') NOT IN ('loan_receivable', 'loan_payment')
            GROUP BY c.id, t.month_due
        `, [asOfMonth]);
        const revenue = revenueResult.length > 0 ? this.rowsToObjects(revenueResult[0]) : [];

        const cogsResult = this.db.exec(`
            SELECT c.id as category_id, t.month_due as month, SUM(t.amount) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.month_due IS NOT NULL AND t.month_due <= ?
            AND c.is_cogs = 1 AND c.show_on_pl != 1
            GROUP BY c.id, t.month_due
        `, [asOfMonth]);
        const cogs = cogsResult.length > 0 ? this.rowsToObjects(cogsResult[0]) : [];

        const opexResult = this.db.exec(`
            SELECT c.id as category_id, t.month_due as month, SUM(t.amount) as total
            FROM transactions t
            JOIN categories c ON t.category_id = c.id
            WHERE t.month_due IS NOT NULL AND t.month_due <= ?
            AND t.transaction_type = 'payable'
            AND c.is_cogs = 0 AND c.is_depreciation = 0 AND c.is_sales_tax = 0 AND c.show_on_pl != 1
            AND COALESCE(t.source_type, '') NOT IN ('loan_payment', 'loan_receivable')
            GROUP BY c.id, t.month_due
        `, [asOfMonth]);
        const opex = opexResult.length > 0 ? this.rowsToObjects(opexResult[0]) : [];

        const deprResult = this.db.exec(`SELECT id as category_id FROM categories WHERE is_depreciation = 1`);
        const deprCats = deprResult.length > 0 ? deprResult[0].values.map(r => r[0]) : [];

        const buildMap = (rows) => {
            const map = {};
            rows.forEach(r => {
                const key = `${r.category_id}-${r.month}`;
                map[key] = (map[key] || 0) + r.total;
            });
            return map;
        };

        const revMap = buildMap(revenue);
        const cogsMap = buildMap(cogs);
        const opexMap = buildMap(opex);

        const revCatIds = [...new Set(revenue.map(r => r.category_id))];
        const cogsCatIds = [...new Set(cogs.map(r => r.category_id))];
        const opexCatIds = [...new Set(opex.map(r => r.category_id))];

        Object.keys(overrides).forEach(key => {
            const [catIdStr] = key.split('-');
            const catId = parseInt(catIdStr);
            if (catId < 0) return;
            if (!revCatIds.includes(catId) && !cogsCatIds.includes(catId) &&
                !opexCatIds.includes(catId) && !deprCats.includes(catId)) {
                const catResult = this.db.exec(`
                    SELECT id FROM categories
                    WHERE id = ? AND is_cogs = 0 AND is_depreciation = 0 AND is_sales_tax = 0 AND show_on_pl != 1
                `, [catId]);
                if (catResult.length > 0 && catResult[0].values.length > 0) {
                    opexCatIds.push(catId);
                }
            }
        });

        let totalRevenue = 0, totalCogs = 0, totalNIBT = 0, totalTax = 0, totalNIAT = 0, totalLoanInterest = 0;

        months.forEach(month => {
            let monthRev = 0;
            revCatIds.forEach(catId => {
                monthRev += getVal(catId, month, revMap[`${catId}-${month}`] || 0);
            });

            let monthCogs = 0;
            cogsCatIds.forEach(catId => {
                monthCogs += getVal(catId, month, cogsMap[`${catId}-${month}`] || 0);
            });

            let monthOpex = 0;
            opexCatIds.forEach(catId => {
                monthOpex += getVal(catId, month, opexMap[`${catId}-${month}`] || 0);
            });
            deprCats.forEach(catId => {
                monthOpex += getVal(catId, month, 0);
            });
            if (assetDeprByMonth[month]) monthOpex += assetDeprByMonth[month];
            if (loanInterestByMonth[month]) {
                monthOpex += loanInterestByMonth[month];
                totalLoanInterest = round2(totalLoanInterest + loanInterestByMonth[month]);
            }

            const nibt = round2(monthRev - monthCogs - monthOpex);
            let tax = 0;
            if (taxMode === 'corporate') {
                const autoTax = round2(nibt > 0 ? nibt * 0.21 : 0);
                tax = getVal(-1, month, autoTax);
            }

            totalRevenue = round2(totalRevenue + monthRev);
            totalCogs = round2(totalCogs + monthCogs);
            totalNIBT = round2(totalNIBT + nibt);
            totalTax = round2(totalTax + tax);
            totalNIAT = round2(totalNIAT + nibt - tax);
        });

        return { totalRevenue, totalCogs, totalGP: round2(totalRevenue - totalCogs), totalNIBT, totalTax, totalNIAT, totalLoanInterest };
    },

    /**
     * Get all P&L overrides
     * @returns {Object} Map of "categoryId-month" => override_amount
     */
    getAllPLOverrides() {
        const results = this.db.exec('SELECT category_id, month, override_amount FROM pl_overrides');
        if (results.length === 0) return {};
        const overrides = {};
        this.rowsToObjects(results[0]).forEach(row => {
            overrides[`${row.category_id}-${row.month}`] = row.override_amount;
        });
        return overrides;
    },

    /**
     * Set a P&L override value for a category+month
     * @param {number} categoryId - Category ID (use -1 for income tax)
     * @param {string} month - Month (YYYY-MM)
     * @param {number|null} amount - Override amount (null to remove override)
     */
    setPLOverride(categoryId, month, amount) {
        if (amount === null || amount === '') {
            this.db.run('DELETE FROM pl_overrides WHERE category_id = ? AND month = ?', [categoryId, month]);
        } else {
            this.db.run(
                'INSERT OR REPLACE INTO pl_overrides (category_id, month, override_amount) VALUES (?, ?, ?)',
                [categoryId, month, parseFloat(amount)]
            );
        }
        this.autoSave();
    },

    clearPLOverridesFrom(startMonth) {
        this.db.run('DELETE FROM pl_overrides WHERE month >= ?', [startMonth]);
        this.autoSave();
    },

    clearCashFlowOverridesFrom(startMonth) {
        this.db.run('DELETE FROM cashflow_overrides WHERE month >= ?', [startMonth]);
        this.autoSave();
    },

    // ==================== PERSISTENCE ====================

    /**
     * Auto-save to IndexedDB
     */
    autoSave: Utils.debounce(async function() {
        await Database.saveToIndexedDB();
    }, 500),

    /**
     * Save database to IndexedDB
     * @returns {Promise<void>}
     */
    async saveToIndexedDB() {
        const data = this.db.export();
        const uint8Array = new Uint8Array(data);

        return new Promise((resolve, reject) => {
            const request = indexedDB.open(this.IDB_NAME, 1);

            request.onerror = () => reject(request.error);

            request.onupgradeneeded = (event) => {
                const db = event.target.result;
                if (!db.objectStoreNames.contains(this.IDB_STORE)) {
                    db.createObjectStore(this.IDB_STORE);
                }
            };

            request.onsuccess = (event) => {
                const db = event.target.result;
                const transaction = db.transaction([this.IDB_STORE], 'readwrite');
                const store = transaction.objectStore(this.IDB_STORE);

                const putRequest = store.put(uint8Array, this.IDB_KEY);
                putRequest.onsuccess = () => {
                    console.log('Database auto-saved to IndexedDB');
                    resolve();
                };
                putRequest.onerror = () => reject(putRequest.error);
            };
        });
    },

    /**
     * Load database from IndexedDB
     * @returns {Promise<Uint8Array|null>}
     */
    async loadFromIndexedDB() {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(this.IDB_NAME, 1);

            request.onerror = () => reject(request.error);

            request.onupgradeneeded = (event) => {
                const db = event.target.result;
                if (!db.objectStoreNames.contains(this.IDB_STORE)) {
                    db.createObjectStore(this.IDB_STORE);
                }
            };

            request.onsuccess = (event) => {
                const db = event.target.result;
                const transaction = db.transaction([this.IDB_STORE], 'readonly');
                const store = transaction.objectStore(this.IDB_STORE);

                const getRequest = store.get(this.IDB_KEY);
                getRequest.onsuccess = () => {
                    resolve(getRequest.result || null);
                };
                getRequest.onerror = () => reject(getRequest.error);
            };
        });
    },

    /**
     * Export database to file
     * @returns {Blob} Database file blob
     */
    exportToFile() {
        const data = this.db.export();
        return new Blob([data], { type: 'application/x-sqlite3' });
    },

    /**
     * Import database from file
     * @param {ArrayBuffer} buffer - File content
     */
    async importFromFile(buffer) {
        const uint8Array = new Uint8Array(buffer);
        this.db = new this.SQL.Database(uint8Array);
        this.migrateSchema();
        await this.saveToIndexedDB();
    },

    // ==================== HELPERS ====================

    // ==================== VE SALES OPERATIONS ====================

    clearVESales(source) {
        if (source) {
            const txNos = this.db.exec('SELECT transaction_no FROM ve_sales WHERE source = ?', [source]);
            if (txNos.length > 0) {
                for (const row of txNos[0].values) {
                    this.db.run('DELETE FROM ve_sale_items WHERE transaction_no = ?', [row[0]]);
                }
            }
            this.db.run('DELETE FROM ve_sales WHERE source = ?', [source]);
        } else {
            this.db.run('DELETE FROM ve_sale_items');
            this.db.run('DELETE FROM ve_sales');
        }
        this.autoSave();
    },

    upsertVESales(salesArray) {
        const stmt = this.db.prepare('INSERT OR REPLACE INTO ve_sales (transaction_no, date, billing_name, description, subtotal, tax, shipping, discount, total, source) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)');
        for (const s of salesArray) {
            stmt.run([s.transactionNo, s.date, s.billingName || '', s.description || '', s.subtotal || 0, s.tax || 0, s.shipping || 0, s.discount || 0, s.total || 0, s.source || 'online']);
        }
        stmt.free();
        this.autoSave();
    },

    upsertVESaleItems(txNo, items) {
        this.db.run('DELETE FROM ve_sale_items WHERE transaction_no = ?', [txNo]);
        const stmt = this.db.prepare('INSERT INTO ve_sale_items (transaction_no, name, product_number, price, quantity, taxable, amount, inferred) VALUES (?, ?, ?, ?, ?, ?, ?, ?)');
        for (const item of items) {
            stmt.run([txNo, item.name || 'Unknown', item.productNumber || null, item.price || 0, item.quantity || 1, item.taxable ? 1 : 0, item.amount || 0, item.inferred ? 1 : 0]);
        }
        stmt.free();
        this.autoSave();
    },

    getVESales(filters = {}) {
        let sql = 'SELECT * FROM ve_sales WHERE 1=1';
        const params = [];
        if (filters.source && filters.source !== 'both') {
            sql += ' AND source = ?';
            params.push(filters.source);
        }
        if (filters.fromDate) {
            sql += ' AND date >= ?';
            params.push(filters.fromDate);
        }
        if (filters.toDate) {
            sql += ' AND date <= ?';
            params.push(filters.toDate);
        }
        sql += ' ORDER BY date DESC';
        const results = this.db.exec(sql, params);
        if (results.length === 0) return [];
        return this.rowsToObjects(results[0]);
    },

    getAllVESaleItems() {
        const results = this.db.exec('SELECT * FROM ve_sale_items');
        if (results.length === 0) return new Map();
        const rows = this.rowsToObjects(results[0]);
        const map = new Map();
        for (const row of rows) {
            if (!map.has(row.transaction_no)) map.set(row.transaction_no, []);
            map.get(row.transaction_no).push({
                name: row.name,
                productNumber: row.product_number,
                price: row.price,
                quantity: row.quantity,
                taxable: !!row.taxable,
                amount: row.amount,
                inferred: !!row.inferred,
            });
        }
        return map;
    },

    getVEImportMeta() {
        const result = this.db.exec("SELECT value FROM app_meta WHERE key = 've_import_meta'");
        if (result.length === 0 || result[0].values.length === 0) return null;
        try { return JSON.parse(result[0].values[0][0]); } catch { return null; }
    },

    setVEImportMeta(meta) {
        this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('ve_import_meta', ?)", [JSON.stringify(meta)]);
        this.autoSave();
    },

    // ==================== VE EVENTS ====================

    createVEEvent(event) {
        this.db.run(
            'INSERT INTO ve_events (name, type, start_date, end_date, notes) VALUES (?, ?, ?, ?, ?)',
            [event.name, event.type || 'tradeshow', event.start_date, event.end_date, event.notes || null]
        );
        const result = this.db.exec('SELECT last_insert_rowid() as id');
        const id = result[0].values[0][0];
        this.autoSave();
        return id;
    },

    updateVEEvent(id, event) {
        this.db.run(
            'UPDATE ve_events SET name = ?, type = ?, start_date = ?, end_date = ?, notes = ? WHERE id = ?',
            [event.name, event.type, event.start_date, event.end_date, event.notes || null, id]
        );
        this.autoSave();
    },

    getVEEventTransaction(eventId) {
        const result = this.db.exec(
            "SELECT id FROM transactions WHERE source_type = 've_event' AND source_id = ?",
            [eventId]
        );
        if (result.length > 0 && result[0].values.length > 0) {
            return result[0].values[0][0];
        }
        return null;
    },

    markVEEventJournalAdded(id, added = 1) {
        this.db.run('UPDATE ve_events SET journal_added = ? WHERE id = ?', [added ? 1 : 0, id]);
        this.autoSave();
    },

    deleteVEEvent(id) {
        this.db.run('UPDATE ve_sales SET event_id = NULL WHERE event_id = ?', [id]);
        this.db.run('DELETE FROM ve_events WHERE id = ?', [id]);
        this.autoSave();
    },

    getAllVEEvents() {
        const results = this.db.exec('SELECT * FROM ve_events ORDER BY start_date DESC');
        if (results.length === 0) return [];
        return this.rowsToObjects(results[0]);
    },

    assignSalesToEvent(eventId, transactionNos) {
        const stmt = this.db.prepare('UPDATE ve_sales SET event_id = ? WHERE transaction_no = ?');
        for (const txNo of transactionNos) {
            stmt.run([eventId, txNo]);
        }
        stmt.free();
        this.autoSave();
    },

    unassignSalesFromEvent(eventId) {
        this.db.run('UPDATE ve_sales SET event_id = NULL WHERE event_id = ?', [eventId]);
        this.autoSave();
    },

    /**
     * Convert sql.js result rows to objects
     * @param {Object} result - sql.js result with columns and values
     * @returns {Array} Array of objects
     */
    rowsToObjects(result) {
        const { columns, values } = result;
        return values.map(row => {
            const obj = {};
            columns.forEach((col, index) => {
                obj[col] = row[index];
            });
            return obj;
        });
    }
};

// Export for use in other modules
window.Database = Database;
