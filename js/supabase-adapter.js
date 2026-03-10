/**
 * SupabaseAdapter — Implements SyncService.api using Supabase (Postgres + Storage).
 *
 * Requires the Supabase JS client loaded via CDN:
 *   <script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/dist/umd/supabase.min.js"></script>
 *
 * API contract (matches SyncService.api):
 *   createGroup(name)                          → { groupId, name, createdAt }
 *   joinGroup(groupId, user)                   → { groupId, name, currentVersion }
 *   pushVersion(groupId, data)                 → { version } or throws ConflictError
 *   pullLatest(groupId)                        → { version, data, savedBy, savedAt } | null
 *   getHistory(groupId, limit)                 → [{ version, savedBy, savedAt, sizeBytes }]
 *   pullVersion(groupId, ver)                  → { version, data, savedBy, savedAt }
 *   registerMember(groupId, name, pw, role)    → { id, group_id, display_name, role }
 *   authenticateMember(groupId, name, pw)      → JSONB | null
 *   verifyMember(groupId, memberId)            → JSONB | null
 *   removeMember(groupId, memberId, adminId)   → true
 *   listMembers(groupId)                       → [{ id, display_name, role, joined_at }]
 */

const SupabaseAdapter = {
    _client: null,
    BUCKET: 'db-blobs',

    init(url, anonKey) {
        this._client = supabase.createClient(url, anonKey);
    },

    isInitialized() {
        return !!this._client;
    },

    // ==================== API METHODS ====================

    async createGroup(name) {
        const { data, error } = await this._client
            .from('groups')
            .insert({ name })
            .select()
            .single();
        if (error) throw new Error('Failed to create group: ' + error.message);
        return { groupId: data.id, name: data.name, createdAt: data.created_at };
    },

    async joinGroup(groupId, user) {
        // Verify group exists
        const { data: group, error: groupErr } = await this._client
            .from('groups')
            .select('id, name')
            .eq('id', groupId)
            .single();
        if (groupErr || !group) throw new Error('Group not found');

        // Member creation is handled by registerMember/authenticateMember RPCs

        // Get current version
        const { data: verData } = await this._client
            .from('versions')
            .select('version')
            .eq('group_id', groupId)
            .order('version', { ascending: false })
            .limit(1);

        const currentVersion = verData && verData.length > 0 ? verData[0].version : 0;

        return {
            groupId,
            name: group.name,
            currentVersion
        };
    },

    async pushVersion(groupId, data) {
        const newVersion = await this._rpcPushVersion(groupId, data);

        // Upload blob to storage
        const path = this._storagePath(groupId, newVersion);
        await this._uploadBlob(path, data.blob);

        return { version: newVersion };
    },

    async pullLatest(groupId) {
        const { data: verData, error } = await this._client
            .from('versions')
            .select('version, saved_by, saved_at, storage_path')
            .eq('group_id', groupId)
            .order('version', { ascending: false })
            .limit(1);

        if (error) throw new Error('Failed to pull: ' + error.message);
        if (!verData || verData.length === 0) return null;

        const v = verData[0];
        const blob = await this._downloadBlob(v.storage_path);

        return {
            version: v.version,
            data: blob,
            savedBy: v.saved_by,
            savedAt: v.saved_at
        };
    },

    async getHistory(groupId, limit = 20) {
        const { data, error } = await this._client
            .from('versions')
            .select('version, saved_by, saved_at, size_bytes')
            .eq('group_id', groupId)
            .order('version', { ascending: false })
            .limit(limit);

        if (error) throw new Error('Failed to get history: ' + error.message);

        return (data || []).map(v => ({
            version: v.version,
            savedBy: v.saved_by,
            savedAt: v.saved_at,
            sizeBytes: v.size_bytes
        }));
    },

    async pullVersion(groupId, ver) {
        const { data: verData, error } = await this._client
            .from('versions')
            .select('version, saved_by, saved_at, storage_path')
            .eq('group_id', groupId)
            .eq('version', ver)
            .single();

        if (error || !verData) throw new Error('Version not found');

        const blob = await this._downloadBlob(verData.storage_path);

        return {
            version: verData.version,
            data: blob,
            savedBy: verData.saved_by,
            savedAt: verData.saved_at
        };
    },

    // ==================== MEMBER AUTH ====================

    async registerMember(groupId, displayName, password, role = 'member') {
        const { data, error } = await this._client.rpc('register_member', {
            p_group_id: groupId,
            p_display_name: displayName,
            p_password: password,
            p_role: role
        });
        if (error) {
            if (error.message && error.message.includes('MEMBER_EXISTS')) {
                const err = new Error('A member with this name already exists');
                err.code = 'MEMBER_EXISTS';
                throw err;
            }
            throw new Error('Registration failed: ' + error.message);
        }
        return data;
    },

    async authenticateMember(groupId, displayName, password) {
        const { data, error } = await this._client.rpc('authenticate_member', {
            p_group_id: groupId,
            p_display_name: displayName,
            p_password: password
        });
        if (error) {
            if (error.message && error.message.includes('INVALID_PASSWORD')) {
                const err = new Error('Incorrect password');
                err.code = 'INVALID_PASSWORD';
                throw err;
            }
            throw new Error('Authentication failed: ' + error.message);
        }
        return data;
    },

    async verifyMember(groupId, memberId) {
        const { data, error } = await this._client.rpc('verify_member', {
            p_group_id: groupId,
            p_member_id: memberId
        });
        if (error) throw new Error('Verify failed: ' + error.message);
        return data;
    },

    async removeMember(groupId, memberId, adminId) {
        const { data, error } = await this._client.rpc('remove_member', {
            p_group_id: groupId,
            p_member_id: memberId,
            p_admin_id: adminId
        });
        if (error) {
            if (error.message && error.message.includes('NOT_ADMIN')) {
                throw new Error('Only admins can remove members');
            }
            if (error.message && error.message.includes('CANNOT_REMOVE_SELF')) {
                throw new Error('Cannot remove yourself');
            }
            throw new Error('Remove failed: ' + error.message);
        }
        return data;
    },

    async listMembers(groupId) {
        const { data, error } = await this._client.rpc('list_members', {
            p_group_id: groupId
        });
        if (error) throw new Error('List members failed: ' + error.message);
        return data || [];
    },

    // ==================== SHARES (VIEW-ONLY SNAPSHOTS) ====================

    async createShare(blob, createdBy, journalName) {
        const { data, error } = await this._client
            .from('shares')
            .insert({
                size_bytes: blob.length,
                storage_path: '',
                created_by: createdBy || 'Unknown',
                journal_name: journalName || ''
            })
            .select()
            .single();
        if (error) throw new Error('Failed to create share: ' + error.message);

        const path = `shares/${data.id}.db`;
        await this._uploadBlob(path, blob);

        const { error: updateErr } = await this._client
            .from('shares')
            .update({ storage_path: path })
            .eq('id', data.id);
        if (updateErr) throw new Error('Failed to update share path: ' + updateErr.message);

        return { shareId: data.id, expiresAt: data.expires_at };
    },

    async getShare(shareId) {
        const { data, error } = await this._client
            .from('shares')
            .select('id, created_at, expires_at, size_bytes, storage_path, created_by, journal_name')
            .eq('id', shareId)
            .single();
        if (error || !data) return null;
        if (new Date(data.expires_at) < new Date()) return null;
        return data;
    },

    async downloadShareBlob(storagePath) {
        return this._downloadBlob(storagePath);
    },

    // ==================== INTERNALS ====================

    _storagePath(groupId, version) {
        return `${groupId}/v${version}.db`;
    },

    async _rpcPushVersion(groupId, data) {
        const { data: result, error } = await this._client.rpc('push_version', {
            p_group_id: groupId,
            p_base_version: data.baseVersion,
            p_saved_by: data.user,
            p_size_bytes: data.blob.length,
            p_storage_path: this._storagePath(groupId, data.baseVersion + 1)
        });

        if (error) {
            if (error.message && error.message.includes('VERSION_CONFLICT')) {
                const conflictErr = new Error('Version conflict');
                conflictErr.name = 'ConflictError';
                conflictErr.code = 'VERSION_CONFLICT';
                // Extract remote version from error message if possible
                const match = error.message.match(/VERSION_CONFLICT:(\d+)/);
                conflictErr.remoteVersion = match ? parseInt(match[1]) : null;
                throw conflictErr;
            }
            throw new Error('Push failed: ' + error.message);
        }

        return result;
    },

    async _uploadBlob(path, blob) {
        const { error } = await this._client.storage
            .from(this.BUCKET)
            .upload(path, blob, {
                contentType: 'application/octet-stream',
                upsert: false
            });
        if (error) throw new Error('Blob upload failed: ' + error.message);
    },

    async _downloadBlob(path) {
        const { data, error } = await this._client.storage
            .from(this.BUCKET)
            .download(path);
        if (error) throw new Error('Blob download failed: ' + error.message);
        const arrayBuffer = await data.arrayBuffer();
        return new Uint8Array(arrayBuffer);
    }
};
