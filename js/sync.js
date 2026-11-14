/**
 * SyncService — Collaborative group file sharing with version history.
 *
 * Architecture:
 *   - Each "group" is a shared document identified by a unique group ID.
 *   - The canonical database blob lives on a remote server.
 *   - On save, the client uploads a new version (snapshot) with metadata.
 *   - On open, the client downloads the latest version.
 *   - Optimistic locking: saves are rejected if someone else saved a newer version
 *     since you last pulled. The client must pull, merge, and retry.
 *   - Version history keeps every snapshot with who saved it and when.
 *
 * This module is backend-agnostic. Set SyncService.api to an object implementing:
 *   - createGroup(name)           → { groupId, name, createdAt }
 *   - joinGroup(groupId, user)    → { groupId, name, members, currentVersion }
 *   - pushVersion(groupId, data)  → { version } or throws ConflictError
 *   - pullLatest(groupId)         → { version, data, savedBy, savedAt } | null
 *   - getHistory(groupId, limit)  → [{ version, savedBy, savedAt, sizeBytes }]
 *   - pullVersion(groupId, ver)   → { version, data, savedBy, savedAt }
 *
 *   data = { blob: Uint8Array, baseVersion: number, user: string }
 */

const SyncService = {
    // ==================== STATE ====================

    /** @type {string|null} Current group ID */
    groupId: null,

    /** @type {number} Version number we last pulled/pushed */
    localVersion: 0,

    /** @type {string} Display name of the current user */
    currentUser: '',

    /** @type {string|null} Authenticated member ID */
    memberId: null,

    /** @type {string} Role of the current user ('admin' or 'member') */
    memberRole: 'member',

    /** @type {boolean} Whether sync is actively connected to a group */
    isConnected: false,

    /** @type {number|null} Interval ID for auto-pull polling */
    _pollInterval: null,

    /** @type {number} Polling frequency in ms (default 30s) */
    pollFrequencyMs: 30000,

    /** @type {Object|null} Pluggable API adapter (must be set before use) */
    api: null,

    /** @type {Function|null} Callback when a remote update is pulled */
    onRemoteUpdate: null,

    /** @type {Function|null} Callback when a conflict is detected */
    onConflict: null,

    /** @type {Function|null} Callback for status changes */
    onStatusChange: null,

    /** @type {number} Increments on every local edit (see wrapAutoSave) */
    _editSeq: 0,

    /** @type {Promise|null} The push currently in flight, if any */
    _pushInFlight: null,

    /** @type {boolean} A save arrived while a push was in flight */
    _pushQueued: false,

    /** @type {boolean} True while a pulled version is being loaded into the app */
    _applyingRemote: false,

    /**
     * Run fn() while a pulled version is being applied, so writes it triggers
     * aren't treated as local edits.
     * @param {Function} fn - May be async
     */
    async applyingRemote(fn) {
        this._applyingRemote = true;
        try {
            return await fn();
        } finally {
            this._applyingRemote = false;
            this.hasUnpushedChanges = false;
        }
    },

    /** @type {Function|null} Returns true to skip a background poll (e.g. a form is open) */
    shouldDeferPull: null,

    /**
     * True while local edits exist that haven't reached the server (a push is in
     * flight, failed, or is blocked by a conflict). A pull must never silently
     * overwrite these.
     * @type {boolean}
     */
    hasUnpushedChanges: false,

    /**
     * Set when the server has a newer version than ours AND we have unpushed edits.
     * Pushing and polling pause until resolveKeepLocal() or resolveTakeRemote().
     * @type {{remoteVersion: number|null}|null}
     */
    pendingConflict: null,

    // ==================== GROUP MANAGEMENT ====================

    /**
     * Create a new shared group and become its first member.
     * @param {string} groupName - Human-readable group name
     * @param {string} userName - Display name of the creator
     * @param {Object} memberInfo - { id, role } from registerMember
     * @returns {Promise<Object>} { groupId, name, createdAt }
     */
    async createGroup(groupName, userName, memberInfo) {
        this._requireApi();
        const result = await this.api.createGroup(groupName);
        this.groupId = result.groupId;
        this.currentUser = userName;
        this.memberId = memberInfo.id;
        this.memberRole = memberInfo.role;
        this.localVersion = 0;
        this.isConnected = true;
        this._emitStatus('connected', `Created group "${groupName}"`);
        return result;
    },

    /**
     * Join an existing group by ID.
     * @param {string} groupId - Group to join
     * @param {string} userName - Display name
     * @param {Object} memberInfo - { id, role } from authentication
     * @returns {Promise<Object>} { groupId, name, currentVersion }
     */
    async joinGroup(groupId, userName, memberInfo) {
        this._requireApi();
        const result = await this.api.joinGroup(groupId, userName);
        this.groupId = groupId;
        this.currentUser = userName;
        this.memberId = memberInfo.id;
        this.memberRole = memberInfo.role;
        this.localVersion = result.currentVersion || 0;
        this.isConnected = true;
        this._emitStatus('connected', `Joined group "${result.name}"`);
        return result;
    },

    /**
     * Disconnect from the current group (does not delete it).
     */
    disconnect() {
        this.stopPolling();
        this.groupId = null;
        this.localVersion = 0;
        this.currentUser = '';
        this.memberId = null;
        this.memberRole = 'member';
        this.isConnected = false;
        this.pendingConflict = null;
        this.hasUnpushedChanges = false;
        this._emitStatus('disconnected', 'Left group');
    },

    // ==================== PUSH / PULL ====================

    /**
     * Push the current database state to the server.
     * Uses optimistic locking: fails if someone else pushed since we last pulled.
     * @param {Uint8Array} dbBlob - Exported database binary
     * @returns {Promise<{version: number, conflict: boolean}>}
     */
    async push(dbBlob) {
        this._requireApi();
        this._requireGroup();

        const payload = {
            blob: dbBlob,
            baseVersion: this.localVersion,
            user: this.currentUser
        };

        this.hasUnpushedChanges = true;
        const seqAtStart = this._editSeq;
        try {
            const result = await this.api.pushVersion(this.groupId, payload);
            this.localVersion = result.version;
            // Edits made while this push was in flight aren't in its blob.
            this.hasUnpushedChanges = this._editSeq !== seqAtStart;
            this._emitStatus('saved', `v${result.version} saved by ${this.currentUser}`);
            return { version: result.version, conflict: false };
        } catch (err) {
            if (err.name === 'ConflictError' || err.code === 'VERSION_CONFLICT') {
                this._raiseConflict(err.remoteVersion != null ? err.remoteVersion : null);
                return { version: this.localVersion, conflict: true };
            }
            throw err;
        }
    },

    /**
     * Pull the latest version from the server.
     * @returns {Promise<{updated: boolean, version: number, data: Uint8Array|null, savedBy: string|null}>}
     */
    async pull() {
        this._requireApi();
        this._requireGroup();

        const result = await this.api.pullLatest(this.groupId);

        if (!result) {
            return { updated: false, version: this.localVersion, data: null, savedBy: null };
        }

        if (result.version > this.localVersion && this.hasUnpushedChanges) {
            // Remote moved on while we have edits the server hasn't seen —
            // applying it would silently discard them. Ask the user instead.
            this._raiseConflict(result.version);
            return { updated: false, conflict: true, version: this.localVersion, data: null, savedBy: null };
        }

        if (result.version > this.localVersion) {
            const prevVersion = this.localVersion;
            this.localVersion = result.version;
            this._emitStatus('updated', `Pulled v${result.version} (saved by ${result.savedBy})`);
            if (this.onRemoteUpdate) {
                this.onRemoteUpdate({
                    version: result.version,
                    data: result.data,
                    savedBy: result.savedBy,
                    savedAt: result.savedAt,
                    previousVersion: prevVersion
                });
            }
            return { updated: true, version: result.version, data: result.data, savedBy: result.savedBy };
        }

        return { updated: false, version: this.localVersion, data: null, savedBy: null };
    },

    /**
     * Push the current DB, one push at a time. Overlapping pushes would share a
     * baseVersion and the second would always conflict, so a save that arrives
     * mid-push is queued and pushed (with a fresh export) when the first finishes.
     */
    async _pushLatest(database) {
        if (this._pushInFlight) {
            this._pushQueued = true;
            return this._pushInFlight;
        }
        this._pushInFlight = (async () => {
            try {
                do {
                    this._pushQueued = false;
                    const blob = new Uint8Array(database.db.export());
                    const result = await this.push(blob);
                    if (result.conflict) break;
                } while (this._pushQueued && this.hasUnpushedChanges && !this.pendingConflict);
            } catch (err) {
                console.error('Remote sync failed (local save succeeded):', err);
            } finally {
                this._pushInFlight = null;
            }
        })();
        return this._pushInFlight;
    },

    // ==================== CONFLICT RESOLUTION ====================

    /**
     * Keep the local copy: push it on top of the server's latest version.
     * The overwritten remote version stays in version history.
     * @param {Uint8Array} dbBlob - Exported local database
     * @returns {Promise<{version: number, conflict: boolean}>}
     */
    async resolveKeepLocal(dbBlob) {
        this._requireApi();
        this._requireGroup();
        const history = await this.api.getHistory(this.groupId, 1);
        const latest = history && history.length ? history[0].version : this.localVersion;
        this.pendingConflict = null;
        this.localVersion = latest;
        // If someone saved again in the meantime this raises a fresh conflict.
        return this.push(dbBlob);
    },

    /**
     * Take the server copy: discard local unpushed edits.
     * @returns {Promise<{version: number, data: Uint8Array, savedBy: string}|null>}
     */
    async resolveTakeRemote() {
        this._requireApi();
        this._requireGroup();
        const result = await this.api.pullLatest(this.groupId);
        this.pendingConflict = null;
        this.hasUnpushedChanges = false;
        if (!result) return null;
        this.localVersion = result.version;
        this._emitStatus('updated', `Pulled v${result.version} (saved by ${result.savedBy})`);
        return { version: result.version, data: result.data, savedBy: result.savedBy };
    },

    _raiseConflict(remoteVersion) {
        const first = !this.pendingConflict;
        this.pendingConflict = { remoteVersion };
        this._emitStatus('conflict', 'Save conflict — someone else saved first');
        if (first && this.onConflict) {
            this.onConflict({ remoteVersion });
        }
    },

    // ==================== VERSION HISTORY ====================

    /**
     * Get version history for the current group.
     * @param {number} [limit=20] - Max entries to return
     * @returns {Promise<Array>} [{ version, savedBy, savedAt, sizeBytes }]
     */
    async getHistory(limit = 20) {
        this._requireApi();
        this._requireGroup();
        return this.api.getHistory(this.groupId, limit);
    },

    /**
     * Pull a specific historical version (for rollback/viewing).
     * @param {number} version - Version number to retrieve
     * @returns {Promise<Object>} { version, data, savedBy, savedAt }
     */
    async pullVersion(version) {
        this._requireApi();
        this._requireGroup();
        return this.api.pullVersion(this.groupId, version);
    },

    // ==================== AUTO-POLLING ====================

    /**
     * Start polling for remote updates at the configured frequency.
     */
    startPolling() {
        this.stopPolling();
        this._pollInterval = setInterval(() => {
            if (this.pendingConflict) return;
            if (this.shouldDeferPull && this.shouldDeferPull()) return;
            this.pull().catch(err => {
                console.error('Sync poll error:', err);
                this._emitStatus('error', 'Failed to check for updates');
            });
        }, this.pollFrequencyMs);
    },

    /**
     * Stop polling for remote updates.
     */
    stopPolling() {
        if (this._pollInterval) {
            clearInterval(this._pollInterval);
            this._pollInterval = null;
        }
    },

    // ==================== INTEGRATION WITH DATABASE ====================

    /**
     * Hook into Database.autoSave to also push to the server.
     * Call this once after SyncService is connected and Database is initialized.
     * @param {Object} database - The Database module
     */
    wrapAutoSave(database) {
        // Idempotent: this is called on every connect. Wrapping twice would push
        // each save twice (the second push conflicting with the first).
        if (database._syncWrapped) return;
        database._syncWrapped = true;

        const originalSave = database.saveToIndexedDB.bind(database);
        const sync = this;

        // Local-only save, for applying a pulled version. Using the wrapped save
        // there would push the copy we just received back as a new version, and
        // two polling clients would ping-pong new versions forever.
        database.saveLocalOnly = originalSave;

        // Mark edits as unpushed the moment they happen, not 500ms later when the
        // debounced save fires, so a poll in that window can't overwrite them.
        const originalAutoSave = database.autoSave;
        database.autoSave = function(...args) {
            // Writes made while applying a pulled version (derived-entry re-sync in
            // refreshAll) aren't user edits — don't push them back.
            if (sync.isConnected && !sync._applyingRemote) {
                sync.hasUnpushedChanges = true;
                sync._editSeq++;
            }
            return originalAutoSave.apply(this, args);
        };

        database.saveToIndexedDB = async function() {
            // Always save locally first
            await originalSave();

            // If connected to a group, also push remotely — but only real local
            // edits, and never while a conflict is waiting on the user.
            if (!sync.isConnected || !sync.groupId) return;
            if (sync.pendingConflict) { sync.hasUnpushedChanges = true; return; }
            if (!sync.hasUnpushedChanges) return;
            await sync._pushLatest(database);
        };
    },

    /**
     * Load the latest remote version into the Database module.
     * @param {Object} database - The Database module
     * @returns {Promise<boolean>} true if a remote version was loaded
     */
    async loadRemoteIntoDatabase(database) {
        const result = await this.pull();
        if (result.updated && result.data) {
            database.db = new database.SQL.Database(result.data);
            database.migrateSchema();
            await (database.saveLocalOnly || database.saveToIndexedDB).call(database);
            return true;
        }
        return false;
    },

    // ==================== INTERNALS ====================

    _requireApi() {
        if (!this.api) {
            throw new Error('SyncService.api is not set. Provide a backend adapter before using sync.');
        }
    },

    _requireGroup() {
        if (!this.groupId) {
            throw new Error('Not connected to a group. Call createGroup() or joinGroup() first.');
        }
    },

    _emitStatus(status, message) {
        if (this.onStatusChange) {
            this.onStatusChange({ status, message, timestamp: Date.now() });
        }
    },

    /**
     * Reset all state (for testing).
     */
    _reset() {
        this.stopPolling();
        this.groupId = null;
        this.localVersion = 0;
        this.currentUser = '';
        this.memberId = null;
        this.memberRole = 'member';
        this.isConnected = false;
        this.api = null;
        this.onRemoteUpdate = null;
        this.onConflict = null;
        this.onStatusChange = null;
        this.shouldDeferPull = null;
        this.pendingConflict = null;
        this.hasUnpushedChanges = false;
    }
};
