---
name: local-first-db
description: "Set up a complete client-side SQLite database (sql.js) with IndexedDB persistence and multi-workspace support. Use when asked to: add a database, set up local storage, add offline/local-first database, add sql.js/SQLite in browser, set up IndexedDB, create offline-first app, set up workspaces/multi-tenant client storage, add data persistence/save data locally, export/import database files, set up schema migrations, create browser database, build local-first architecture, add client-side data layer, database migration, multiple projects, workspace switching."
metadata:
  author: user
  version: "1.0.0"
  argument-hint: <app-name>
---

# Local-First SQLite Database with IndexedDB Persistence

This skill creates a complete client-side database layer using **sql.js** (SQLite compiled to WASM) with IndexedDB for blob persistence and full multi-workspace support. The architecture is framework-agnostic and works in any browser-based project.

---

## Step 1: Gather Requirements (Ask One at a Time)

Before generating any code, ask the user these questions **one at a time**, waiting for each answer:

1. **App name** — This becomes the IndexedDB database name prefix (e.g., `"my-app"` → IDB name `"my-app-db"`). If invoked with an argument, use that as the app name and confirm it.

2. **Schema tables** — Ask for each table: name, columns (name, type, constraints). An `app_meta` key-value table is always auto-created. Example prompt:
   > What tables do you need? For each, list columns with types and constraints.
   > Example: `tasks` table with `id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL, done INTEGER DEFAULT 0, created_at TEXT DEFAULT (datetime('now'))`
   > I will always auto-create an `app_meta` key-value table.

3. **Max workspaces** — Maximum number of workspaces allowed. Default: **5**. Say: "Max workspaces? (default: 5)"

4. **Default workspace name** — Name for the initial workspace created on first run. Default: **"My Workspace"**. Say: "Default workspace name? (default: 'My Workspace')"

5. **sql.js CDN version** — Which version of sql.js to load from CDN. Default: **1.8.0**. Say: "sql.js version? (default: 1.8.0)"

6. **Auto-save debounce ms** — Milliseconds to debounce auto-save after mutations. Default: **500**. Say: "Auto-save debounce in ms? (default: 500)"

7. **Copyable sections between workspaces** — Which tables/data can be copied from one workspace into another? This drives the `copySection()` method. Say: "Which tables should be copyable between workspaces? (list table names, or 'none')"

---

## Step 2: Add sql.js Script Tag

Add the sql.js CDN script to the project's HTML file, **before** any app scripts:

```html
<!-- sql.js WASM SQLite engine -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/sql.js/{{VERSION}}/sql-wasm.min.js"></script>
```

Replace `{{VERSION}}` with the chosen version (default `1.8.0`).

Also ensure `js/database.js` and `js/workspaces.js` are loaded (in that order) before any app code that uses them:

```html
<script src="js/database.js"></script>
<script src="js/workspaces.js"></script>
```

---

## Step 3: Generate `js/database.js`

Create the file `js/database.js` with the following structure. This is a **singleton** object exposed as `window.Database` (or whatever name fits the project's conventions).

### 3.1 Full File Template

```javascript
/**
 * Database — Client-side SQLite via sql.js with IndexedDB persistence.
 *
 * API:
 *   init(bytes)         — Initialize sql.js, create/open database
 *   run(sql, params)    — Execute INSERT/UPDATE/DELETE, auto-saves
 *   get(sql, params)    — SELECT single row → object or null
 *   getAll(sql, params) — SELECT multiple rows → array of objects
 *   exec(sql)           — Raw exec (DDL, multi-statement), auto-saves
 *   transaction(fn)     — Run fn() inside BEGIN/COMMIT, auto-saves once
 *   getMeta(key)         — Get string value from app_meta
 *   setMeta(key, value)  — Set string value in app_meta, auto-saves
 *   getMetaJSON(key)     — Get parsed JSON from app_meta
 *   setMetaJSON(key, val)— Set JSON-stringified value in app_meta, auto-saves
 *   exportToFile(name)   — Download .db file
 *   importFromFile(file) — Load .db file, returns Promise
 *   autoSave()           — Debounced save to IndexedDB
 *   saveToIndexedDB()    — Immediate save to IndexedDB
 *   resetAllData()       — Drop all tables, re-create schema
 *   loadBytes(bytes)     — Load Uint8Array (or null for fresh DB)
 *   rowsToObjects(stmt)  — Convert sql.js statement result to object array
 */
const Database = (() => {
    // ── Private state ──────────────────────────────────────────────
    let db = null;                    // sql.js Database instance
    let SQL = null;                   // sql.js library reference
    let _saveTimer = null;            // debounce timer
    let _idbName = '{{APP_NAME}}-db'; // IndexedDB database name
    let _storeName = 'databases';     // IDB object store name
    const DEBOUNCE_MS = {{DEBOUNCE}};
    const SQLJS_VERSION = '{{VERSION}}';

    // ── Schema Definition ──────────────────────────────────────────
    // Replace this with the user's actual schema.
    function createSchema() {
        // app_meta is ALWAYS created — key-value store for settings/flags
        db.run(`
            CREATE TABLE IF NOT EXISTS app_meta (
                key   TEXT PRIMARY KEY,
                value TEXT
            )
        `);

        // {{USER TABLES GO HERE}}
        // Example:
        // db.run(`
        //     CREATE TABLE IF NOT EXISTS tasks (
        //         id         INTEGER PRIMARY KEY AUTOINCREMENT,
        //         title      TEXT NOT NULL,
        //         done       INTEGER DEFAULT 0,
        //         created_at TEXT DEFAULT (datetime('now'))
        //     )
        // `);
    }

    // ── Schema Migrations ──────────────────────────────────────────
    // Each migration is wrapped in try/catch so it is independent.
    // Probe for the new column/table first; if it fails, run the ALTER.
    function runMigrations() {
        // Pattern for adding a column:
        // try {
        //     db.exec("SELECT new_column FROM some_table LIMIT 0");
        // } catch (e) {
        //     db.run("ALTER TABLE some_table ADD COLUMN new_column TEXT DEFAULT ''");
        // }

        // Pattern for one-time DATA migration using app_meta flag:
        // const flag = getMeta('migration_xyz_done');
        // if (!flag) {
        //     // ... do data transformation ...
        //     setMetaImmediate('migration_xyz_done', '1');  // write flag AFTER success
        // }
    }

    // ── Helpers: setMetaImmediate (no auto-save, used in migrations) ──
    function setMetaImmediate(key, value) {
        db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES (?, ?)", [key, value]);
    }

    // ── IDB Key ────────────────────────────────────────────────────
    // The IDB key is provided by WorkspaceManager (workspace-specific).
    let _activeKey = null;

    function setActiveKey(key) {
        _activeKey = key;
    }

    function getIDBKey() {
        return _activeKey || 'default';
    }

    // ── IndexedDB low-level ────────────────────────────────────────
    function _openIDB() {
        return new Promise((resolve, reject) => {
            const req = indexedDB.open(_idbName, 1);
            req.onupgradeneeded = () => {
                const idb = req.result;
                if (!idb.objectStoreNames.contains(_storeName)) {
                    idb.createObjectStore(_storeName);
                }
            };
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });
    }

    function _readIDB(key) {
        return _openIDB().then(idb => new Promise((resolve, reject) => {
            const tx = idb.transaction(_storeName, 'readonly');
            const store = tx.objectStore(_storeName);
            const req = store.get(key);
            req.onsuccess = () => { idb.close(); resolve(req.result || null); };
            req.onerror = () => { idb.close(); reject(req.error); };
        }));
    }

    function _writeIDB(key, value) {
        return _openIDB().then(idb => new Promise((resolve, reject) => {
            const tx = idb.transaction(_storeName, 'readwrite');
            const store = tx.objectStore(_storeName);
            const req = store.put(value, key);
            req.onsuccess = () => { idb.close(); resolve(); };
            req.onerror = () => { idb.close(); reject(req.error); };
        })).catch(err => {
            if (err.name === 'QuotaExceededError') {
                console.error('Storage quota exceeded. Cannot save database.');
                alert('Storage is full. Please export your data and free up space.');
            }
            throw err;
        });
    }

    function _deleteIDB(key) {
        return _openIDB().then(idb => new Promise((resolve, reject) => {
            const tx = idb.transaction(_storeName, 'readwrite');
            const store = tx.objectStore(_storeName);
            const req = store.delete(key);
            req.onsuccess = () => { idb.close(); resolve(); };
            req.onerror = () => { idb.close(); reject(req.error); };
        }));
    }

    // ── rowsToObjects ──────────────────────────────────────────────
    function rowsToObjects(result) {
        // result is from db.exec("SELECT ...")[0]
        // shape: { columns: [...], values: [[...], ...] }
        if (!result || !result.values || result.values.length === 0) return [];
        return result.values.map(row => {
            const obj = {};
            result.columns.forEach((col, i) => { obj[col] = row[i]; });
            return obj;
        });
    }

    // ── Public API ─────────────────────────────────────────────────
    return {
        /**
         * Initialize the database engine.
         * @param {Uint8Array|null} bytes - Existing DB bytes from IDB, or null for fresh.
         */
        async init(bytes) {
            SQL = await initSqlJs({
                locateFile: file => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/${SQLJS_VERSION}/${file}`
            });
            this.loadBytes(bytes);
        },

        /**
         * Load database from bytes or create fresh.
         * @param {Uint8Array|null} bytes
         */
        loadBytes(bytes) {
            if (db) {
                try { db.close(); } catch (e) { /* ignore */ }
            }
            if (bytes) {
                try {
                    db = new SQL.Database(bytes);
                } catch (e) {
                    console.error('Corrupt database bytes, creating fresh DB:', e);
                    // DO NOT delete the corrupt bytes — user may want to recover
                    db = new SQL.Database();
                    createSchema();
                    return;
                }
            } else {
                db = new SQL.Database();
                createSchema();
            }
            runMigrations();
        },

        /**
         * Execute a mutating statement (INSERT/UPDATE/DELETE).
         * @param {string} sql
         * @param {Array} [params]
         */
        run(sql, params) {
            db.run(sql, params);
            this.autoSave();
        },

        /**
         * Query a single row.
         * @param {string} sql
         * @param {Array} [params]
         * @returns {Object|null}
         */
        get(sql, params) {
            const stmt = db.prepare(sql);
            if (params) stmt.bind(params);
            let result = null;
            if (stmt.step()) {
                result = stmt.getAsObject();
            }
            stmt.free();
            return result;
        },

        /**
         * Query multiple rows.
         * @param {string} sql
         * @param {Array} [params]
         * @returns {Array<Object>}
         */
        getAll(sql, params) {
            const stmt = db.prepare(sql);
            if (params) stmt.bind(params);
            const rows = [];
            while (stmt.step()) {
                rows.push(stmt.getAsObject());
            }
            stmt.free();
            return rows;
        },

        /**
         * Execute raw SQL (DDL, multi-statement). Auto-saves.
         * @param {string} sql
         * @returns {Array} exec result
         */
        exec(sql) {
            const result = db.exec(sql);
            this.autoSave();
            return result;
        },

        /**
         * Run a function inside a transaction. Auto-saves once at end.
         * @param {Function} fn - Receives no args; use Database.run/get/getAll inside.
         */
        transaction(fn) {
            db.run("BEGIN TRANSACTION");
            try {
                fn();
                db.run("COMMIT");
            } catch (e) {
                db.run("ROLLBACK");
                throw e;
            }
            this.autoSave();
        },

        // ── Key-Value Store (app_meta) ─────────────────────────────
        getMeta(key) {
            const row = this.get("SELECT value FROM app_meta WHERE key = ?", [key]);
            return row ? row.value : null;
        },

        setMeta(key, value) {
            this.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES (?, ?)", [key, value]);
        },

        getMetaJSON(key) {
            const raw = this.getMeta(key);
            if (raw === null) return null;
            try { return JSON.parse(raw); } catch { return null; }
        },

        setMetaJSON(key, value) {
            this.setMeta(key, JSON.stringify(value));
        },

        // ── Persistence ────────────────────────────────────────────
        autoSave() {
            clearTimeout(_saveTimer);
            _saveTimer = setTimeout(() => this.saveToIndexedDB(), DEBOUNCE_MS);
        },

        async saveToIndexedDB() {
            clearTimeout(_saveTimer);
            if (!db) return;
            const bytes = db.export();
            const arr = new Uint8Array(bytes);
            await _writeIDB(getIDBKey(), arr);
        },

        // ── File Export / Import ───────────────────────────────────
        exportToFile(filename) {
            if (!db) return;
            const bytes = db.export();
            const blob = new Blob([bytes], { type: 'application/x-sqlite3' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename || '{{APP_NAME}}-backup.db';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        },

        importFromFile(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = () => {
                    try {
                        const bytes = new Uint8Array(reader.result);
                        this.loadBytes(bytes);
                        this.saveToIndexedDB().then(resolve).catch(reject);
                    } catch (e) {
                        reject(e);
                    }
                };
                reader.onerror = () => reject(reader.error);
                reader.readAsArrayBuffer(file);
            });
        },

        // ── Reset ──────────────────────────────────────────────────
        resetAllData() {
            if (db) {
                try { db.close(); } catch (e) { /* ignore */ }
            }
            db = new SQL.Database();
            createSchema();
            this.autoSave();
        },

        // ── Utilities ──────────────────────────────────────────────
        rowsToObjects,
        setActiveKey,

        /** Get the raw sql.js Database instance (advanced use). */
        getDB() { return db; },

        /** Check if initialized. */
        isReady() { return db !== null; }
    };
})();
```

### 3.2 Customization Points

When generating the file, replace:
- `{{APP_NAME}}` — the user's app name
- `{{VERSION}}` — sql.js CDN version
- `{{DEBOUNCE}}` — debounce milliseconds
- `{{USER TABLES GO HERE}}` — generate `db.run("CREATE TABLE IF NOT EXISTS ...")` calls for every table the user specified
- Add migration stubs in `runMigrations()` with comments explaining the pattern

---

## Step 4: Generate `js/workspaces.js`

Create `js/workspaces.js` as a **singleton** exposed as `window.WorkspaceManager`.

### 4.1 Full File Template

```javascript
/**
 * WorkspaceManager — Multi-workspace support over IndexedDB.
 *
 * Each workspace is a separate SQLite database blob stored in IDB.
 * A registry (also in IDB) tracks all workspaces and which is active.
 *
 * API:
 *   init()                      — Load registry, ensure default workspace
 *   getActiveKey()              — IDB key for the active workspace blob
 *   getActiveWorkspace()        — Active workspace metadata object
 *   getAll()                    — All workspace metadata objects
 *   switchTo(id)                — Save current, load another workspace
 *   createNew(name)             — Create a new empty workspace
 *   rename(id, newName)         — Rename a workspace
 *   delete(id)                  — Delete a workspace (cannot delete last)
 *   importAsNew(file, name)     — Import a .db file as a new workspace
 *   replaceActive(file)         — Replace active workspace with imported file
 *   copySection(sourceId, opts) — Copy table data from another workspace
 */
const WorkspaceManager = (() => {
    // ── Private state ──────────────────────────────────────────────
    const IDB_NAME = '{{APP_NAME}}-db';
    const STORE = 'databases';
    const REGISTRY_KEY = '__workspace_registry__';
    const MAX_WORKSPACES = {{MAX_WORKSPACES}};
    const DEFAULT_WS_NAME = '{{DEFAULT_WS_NAME}}';

    let _registry = null;  // { version, activeId, workspaces: [] }

    // ── IDB Helpers (private) ──────────────────────────────────────
    function _openIDB() {
        return new Promise((resolve, reject) => {
            const req = indexedDB.open(IDB_NAME, 1);
            req.onupgradeneeded = () => {
                const idb = req.result;
                if (!idb.objectStoreNames.contains(STORE)) {
                    idb.createObjectStore(STORE);
                }
            };
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });
    }

    function _readIDB(key) {
        return _openIDB().then(idb => new Promise((resolve, reject) => {
            const tx = idb.transaction(STORE, 'readonly');
            const store = tx.objectStore(STORE);
            const req = store.get(key);
            req.onsuccess = () => { idb.close(); resolve(req.result || null); };
            req.onerror = () => { idb.close(); reject(req.error); };
        }));
    }

    function _writeIDB(key, value) {
        return _openIDB().then(idb => new Promise((resolve, reject) => {
            const tx = idb.transaction(STORE, 'readwrite');
            const store = tx.objectStore(STORE);
            const req = store.put(value, key);
            req.onsuccess = () => { idb.close(); resolve(); };
            req.onerror = () => { idb.close(); reject(req.error); };
        }));
    }

    function _deleteIDB(key) {
        return _openIDB().then(idb => new Promise((resolve, reject) => {
            const tx = idb.transaction(STORE, 'readwrite');
            const store = tx.objectStore(STORE);
            const req = store.delete(key);
            req.onsuccess = () => { idb.close(); resolve(); };
            req.onerror = () => { idb.close(); reject(req.error); };
        }));
    }

    // ── Registry helpers ───────────────────────────────────────────
    function _generateId() {
        return 'ws_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
    }

    function _dbKey(wsId) {
        return 'workspace_' + wsId;
    }

    async function _loadRegistry() {
        _registry = await _readIDB(REGISTRY_KEY);
        if (!_registry) {
            // First run — create default workspace
            const id = _generateId();
            _registry = {
                version: 1,
                activeId: id,
                workspaces: [{
                    id,
                    name: DEFAULT_WS_NAME,
                    createdAt: new Date().toISOString(),
                    lastAccessedAt: new Date().toISOString()
                }]
            };
            await _saveRegistry();
        }
    }

    async function _saveRegistry() {
        await _writeIDB(REGISTRY_KEY, _registry);
    }

    function _findWs(id) {
        return _registry.workspaces.find(ws => ws.id === id);
    }

    // ── Public API ─────────────────────────────────────────────────
    return {
        /**
         * Initialize workspace manager. Call before Database.init().
         * Returns the bytes for the active workspace (or null if new).
         */
        async init() {
            await _loadRegistry();

            // Touch lastAccessedAt on active
            const active = _findWs(_registry.activeId);
            if (active) {
                active.lastAccessedAt = new Date().toISOString();
                await _saveRegistry();
            }

            // Tell Database which IDB key to use
            Database.setActiveKey(_dbKey(_registry.activeId));

            // Load active workspace bytes
            const bytes = await _readIDB(_dbKey(_registry.activeId));
            return bytes;  // Uint8Array or null
        },

        /** IDB key for the active workspace. */
        getActiveKey() {
            return _dbKey(_registry.activeId);
        },

        /** Metadata object for the active workspace. */
        getActiveWorkspace() {
            return _findWs(_registry.activeId);
        },

        /** All workspace metadata objects. */
        getAll() {
            return [..._registry.workspaces];
        },

        /**
         * Switch to a different workspace.
         * Saves current DB, loads target workspace bytes, re-inits Database.
         */
        async switchTo(id) {
            if (id === _registry.activeId) return;
            const target = _findWs(id);
            if (!target) throw new Error('Workspace not found: ' + id);

            // Save current workspace
            await Database.saveToIndexedDB();

            // Update registry
            _registry.activeId = id;
            target.lastAccessedAt = new Date().toISOString();
            await _saveRegistry();

            // Update Database key and load new bytes
            Database.setActiveKey(_dbKey(id));
            const bytes = await _readIDB(_dbKey(id));
            Database.loadBytes(bytes);
        },

        /**
         * Create a new empty workspace.
         * @param {string} [name] — Workspace name
         * @returns {Object} The new workspace metadata
         */
        async createNew(name) {
            if (_registry.workspaces.length >= MAX_WORKSPACES) {
                throw new Error(`Maximum of ${MAX_WORKSPACES} workspaces reached.`);
            }
            const id = _generateId();
            const ws = {
                id,
                name: name || ('Workspace ' + (_registry.workspaces.length + 1)),
                createdAt: new Date().toISOString(),
                lastAccessedAt: new Date().toISOString()
            };
            _registry.workspaces.push(ws);
            await _saveRegistry();
            // Don't switch — just create the entry. switchTo() will create fresh DB.
            return ws;
        },

        /**
         * Rename a workspace.
         */
        async rename(id, newName) {
            const ws = _findWs(id);
            if (!ws) throw new Error('Workspace not found: ' + id);
            ws.name = newName;
            await _saveRegistry();
        },

        /**
         * Delete a workspace. Cannot delete the last one.
         */
        async delete(id) {
            if (_registry.workspaces.length <= 1) {
                throw new Error('Cannot delete the last workspace.');
            }
            const idx = _registry.workspaces.findIndex(ws => ws.id === id);
            if (idx === -1) throw new Error('Workspace not found: ' + id);

            // Remove from registry
            _registry.workspaces.splice(idx, 1);

            // If deleting the active workspace, switch to another
            if (_registry.activeId === id) {
                _registry.activeId = _registry.workspaces[0].id;
                Database.setActiveKey(_dbKey(_registry.activeId));
                const bytes = await _readIDB(_dbKey(_registry.activeId));
                Database.loadBytes(bytes);
            }

            await _saveRegistry();

            // Delete the blob from IDB
            await _deleteIDB(_dbKey(id));
        },

        /**
         * Import a .db file as a new workspace.
         * @param {File} file — The .db file
         * @param {string} [name] — Workspace name (defaults to filename)
         */
        async importAsNew(file, name) {
            if (_registry.workspaces.length >= MAX_WORKSPACES) {
                throw new Error(`Maximum of ${MAX_WORKSPACES} workspaces reached.`);
            }
            const buffer = await file.arrayBuffer();
            const bytes = new Uint8Array(buffer);

            const id = _generateId();
            const ws = {
                id,
                name: name || file.name.replace(/\.db$/i, ''),
                createdAt: new Date().toISOString(),
                lastAccessedAt: new Date().toISOString()
            };
            _registry.workspaces.push(ws);
            await _writeIDB(_dbKey(id), bytes);
            await _saveRegistry();
            return ws;
        },

        /**
         * Replace the active workspace's database with an imported file.
         * @param {File} file — The .db file
         */
        async replaceActive(file) {
            const buffer = await file.arrayBuffer();
            const bytes = new Uint8Array(buffer);
            Database.loadBytes(bytes);
            await Database.saveToIndexedDB();
        },

        /**
         * Copy a section (table data) from a source workspace into the active one.
         * Opens the source as a temporary sql.js database, reads rows, inserts into active.
         *
         * @param {string} sourceId — Source workspace ID
         * @param {Object} opts
         * @param {string} opts.table — Table name to copy
         * @param {string} [opts.where] — Optional WHERE clause (without WHERE keyword)
         * @param {boolean} [opts.clearFirst] — DELETE target rows before copy (default false)
         */
        async copySection(sourceId, opts) {
            const { table, where, clearFirst } = opts;
            if (sourceId === _registry.activeId) {
                throw new Error('Source and target workspace are the same.');
            }

            // Load source bytes into a temp database
            const sourceBytes = await _readIDB(_dbKey(sourceId));
            if (!sourceBytes) throw new Error('Source workspace has no data.');

            // Need access to SQL from Database module — use sql.js global
            const tempDB = new SQL.Database(sourceBytes);

            try {
                // Read rows from source
                let query = `SELECT * FROM ${table}`;
                if (where) query += ` WHERE ${where}`;
                const result = tempDB.exec(query);
                if (!result || result.length === 0) return;

                const { columns, values } = result[0];
                if (values.length === 0) return;

                // Build insert SQL
                const placeholders = columns.map(() => '?').join(', ');
                const insertSQL = `INSERT OR REPLACE INTO ${table} (${columns.join(', ')}) VALUES (${placeholders})`;

                // Write into active database
                Database.transaction(() => {
                    if (clearFirst) {
                        Database.run(`DELETE FROM ${table}`);
                    }
                    for (const row of values) {
                        Database.run(insertSQL, row);
                    }
                });
            } finally {
                tempDB.close();
            }
        }
    };
})();
```

### 4.2 Customization Points

Replace:
- `{{APP_NAME}}` — the user's app name (same as Database)
- `{{MAX_WORKSPACES}}` — max workspaces number
- `{{DEFAULT_WS_NAME}}` — default workspace name
- Populate `copySection` documentation comments with the user's specified copyable tables

---

## Step 5: Create the Init Sequence

In the app's main entry point (e.g., `js/app.js`, `js/main.js`, or `<script>` in HTML), add the initialization sequence:

```javascript
// ── App Initialization ─────────────────────────────────────────────
async function initApp() {
    try {
        // 1. Initialize workspace manager (loads registry from IDB)
        const bytes = await WorkspaceManager.init();

        // 2. Initialize database engine (loads sql.js WASM, opens DB)
        await Database.init(bytes);

        // 3. App is ready — render UI, load data, etc.
        console.log('Database ready. Active workspace:',
            WorkspaceManager.getActiveWorkspace().name);

        // ... your app init code here ...

    } catch (err) {
        console.error('Failed to initialize:', err);
    }
}

// Start when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
} else {
    initApp();
}
```

**IMPORTANT**: This init sequence MUST follow this exact order:
1. `WorkspaceManager.init()` — returns bytes for active workspace
2. `Database.init(bytes)` — loads WASM engine, then opens DB from bytes
3. App-specific setup

---

## Step 6: Edge Cases — Handle in the Generated Code

### 6.1 Corrupt Data
Already handled in `loadBytes()`:
```javascript
try {
    db = new SQL.Database(bytes);
} catch (e) {
    console.error('Corrupt database bytes, creating fresh DB:', e);
    // DO NOT delete corrupt bytes from IDB — user may want manual recovery
    db = new SQL.Database();
    createSchema();
    return;
}
```

### 6.2 Storage Quota Exceeded
Already handled in `_writeIDB()`:
```javascript
.catch(err => {
    if (err.name === 'QuotaExceededError') {
        console.error('Storage quota exceeded.');
        alert('Storage is full. Please export your data and free up space.');
    }
    throw err;
});
```

### 6.3 Last Workspace Cannot Be Deleted
Already handled in `delete()`:
```javascript
if (_registry.workspaces.length <= 1) {
    throw new Error('Cannot delete the last workspace.');
}
```

### 6.4 Concurrent Tabs
Document this as a known limitation in a code comment at the top of `database.js`:
```javascript
// NOTE: Concurrent tab support is NOT implemented. If the app is open
// in multiple tabs, the last tab to save will overwrite changes from
// other tabs. For multi-tab support, consider BroadcastChannel or
// SharedWorker coordination (not included in this implementation).
```

### 6.5 Empty/New Workspace
When `loadBytes(null)` is called (new workspace with no saved data), the code creates a fresh `SQL.Database()` and calls `createSchema()`. This is already handled in the `loadBytes()` method.

---

## Step 7: Schema Migration Pattern — Detailed

Explain this pattern in comments within `runMigrations()`:

### Column Addition Migration
```javascript
// Migration: Add 'priority' column to 'tasks' table
// Date added: YYYY-MM-DD
try {
    db.exec("SELECT priority FROM tasks LIMIT 0");
} catch (e) {
    db.run("ALTER TABLE tasks ADD COLUMN priority INTEGER DEFAULT 0");
}
```

### One-Time Data Migration
```javascript
// Migration: Backfill priority from old 'importance' text field
// Date added: YYYY-MM-DD
const migrated = getMeta('migration_backfill_priority');
if (!migrated) {
    db.run("UPDATE tasks SET priority = CASE WHEN importance = 'high' THEN 3 WHEN importance = 'medium' THEN 2 ELSE 1 END");
    setMetaImmediate('migration_backfill_priority', '1');  // flag AFTER success
}
```

### New Table Migration
```javascript
// Migration: Add 'tags' table
// Date added: YYYY-MM-DD
db.run(`
    CREATE TABLE IF NOT EXISTS tags (
        id      INTEGER PRIMARY KEY AUTOINCREMENT,
        task_id INTEGER REFERENCES tasks(id),
        name    TEXT NOT NULL
    )
`);
```

**Key rules for migrations:**
- Each migration is **independent** — wrapped in its own try/catch or IF block
- Column probing uses `SELECT col FROM table LIMIT 0` — cheap, no data scanned
- Data migration flags are written to `app_meta` **AFTER** success, never before
- Use `setMetaImmediate()` (no auto-save) inside migrations to avoid save loops
- `CREATE TABLE IF NOT EXISTS` is inherently idempotent
- Never remove old migrations — they are no-ops on upgraded databases

---

## Step 8: Post-Generation Checklist

After generating both files, verify:

1. [ ] `js/database.js` has all user tables in `createSchema()`
2. [ ] `js/workspaces.js` has correct app name, max workspaces, default name
3. [ ] Script tags added to HTML in correct order: sql.js CDN → database.js → workspaces.js → app code
4. [ ] Init sequence added to app entry point
5. [ ] If the project already has a database or storage layer, integrate rather than duplicate
6. [ ] All `{{PLACEHOLDER}}` values are replaced with actual values
7. [ ] Concurrent-tab warning comment is present
8. [ ] The `copySection()` method in WorkspaceManager has been customized (or left generic) per user's answer to question 7

---

## API Quick Reference (include as comment block at top of each file)

### Database API
| Method | Returns | Auto-saves | Description |
|--------|---------|------------|-------------|
| `init(bytes)` | `Promise<void>` | No | Load sql.js WASM, open DB from bytes |
| `run(sql, params?)` | `void` | Yes | Execute INSERT/UPDATE/DELETE |
| `get(sql, params?)` | `Object\|null` | No | Query single row |
| `getAll(sql, params?)` | `Object[]` | No | Query multiple rows |
| `exec(sql)` | `Array` | Yes | Raw exec for DDL or multi-statement |
| `transaction(fn)` | `void` | Yes (once) | Run fn inside BEGIN/COMMIT |
| `getMeta(key)` | `string\|null` | No | Read from app_meta |
| `setMeta(key, value)` | `void` | Yes | Write to app_meta |
| `getMetaJSON(key)` | `any\|null` | No | Read + JSON.parse from app_meta |
| `setMetaJSON(key, value)` | `void` | Yes | JSON.stringify + write to app_meta |
| `exportToFile(name?)` | `void` | No | Download .db file |
| `importFromFile(file)` | `Promise<void>` | Yes | Load .db file |
| `autoSave()` | `void` | — | Debounced save trigger |
| `saveToIndexedDB()` | `Promise<void>` | — | Immediate save |
| `resetAllData()` | `void` | Yes | Nuke and re-create schema |
| `loadBytes(bytes)` | `void` | No | Load Uint8Array or null for fresh |
| `rowsToObjects(result)` | `Object[]` | No | Convert exec result to objects |

### WorkspaceManager API
| Method | Returns | Description |
|--------|---------|-------------|
| `init()` | `Promise<Uint8Array\|null>` | Load registry, return active workspace bytes |
| `getActiveKey()` | `string` | IDB key for active workspace |
| `getActiveWorkspace()` | `Object` | Active workspace metadata |
| `getAll()` | `Object[]` | All workspace metadata |
| `switchTo(id)` | `Promise<void>` | Save current, load target workspace |
| `createNew(name?)` | `Promise<Object>` | Create new empty workspace |
| `rename(id, name)` | `Promise<void>` | Rename a workspace |
| `delete(id)` | `Promise<void>` | Delete workspace (not last) |
| `importAsNew(file, name?)` | `Promise<Object>` | Import .db as new workspace |
| `replaceActive(file)` | `Promise<void>` | Replace active DB with import |
| `copySection(srcId, opts)` | `Promise<void>` | Copy table data between workspaces |

### Registry Shape
```json
{
    "version": 1,
    "activeId": "ws_1710000000000_abc123",
    "workspaces": [
        {
            "id": "ws_1710000000000_abc123",
            "name": "My Workspace",
            "createdAt": "2024-03-10T00:00:00.000Z",
            "lastAccessedAt": "2024-03-10T12:00:00.000Z"
        }
    ]
}
```
