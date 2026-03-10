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

        try {
            const result = await this.api.pushVersion(this.groupId, payload);
            this.localVersion = result.version;
            this._emitStatus('saved', `v${result.version} saved by ${this.currentUser}`);
            return { version: result.version, conflict: false };
        } catch (err) {
            if (err.name === 'ConflictError' || err.code === 'VERSION_CONFLICT') {
                this._emitStatus('conflict', 'Save conflict — someone else saved first');
                if (this.onConflict) {
                    this.onConflict(err);
                }
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
        const originalSave = database.saveToIndexedDB.bind(database);
        const sync = this;

        database.saveToIndexedDB = async function() {
            // Always save locally first
            await originalSave();

            // If connected to a group, also push remotely
            if (sync.isConnected && sync.groupId) {
                try {
                    const blob = new Uint8Array(database.db.export());
                    const result = await sync.push(blob);
                    if (result.conflict) {
                        console.warn('Sync conflict detected — remote version is newer');
                    }
                } catch (err) {
                    console.error('Remote sync failed (local save succeeded):', err);
                }
            }
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
            await database.saveToIndexedDB();
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
    }
};
