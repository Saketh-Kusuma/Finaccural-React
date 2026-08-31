'use strict';

const axios       = require('axios');
const querystring = require('querystring');
const config      = require('../../core/config');
const CONSTANTS   = require('../../core/constants');
const { encodeBasicAuth } = require('../../core/helpers');
const XeroTokenRepository = require('./repository');
const XeroMapper  = require('./mapper');
const logger      = require('../../core/logger');
const XeroTokenManager = require('./oauth/XeroTokenManager');
const { ErpSessionExpiredError } = require('../../core/errors/AppError');

/** True if an axios error looks like an expired/revoked OAuth grant. */
function isAuthError(err) {
    if (!err) return false;
    if (err.response?.status === 401 || err.response?.status === 403) return true;
    const blob = JSON.stringify(err.response?.data || err.message || '').toLowerCase();
    return blob.includes('invalid_grant') || blob.includes('invalid_token') || blob.includes('unauthorized');
}

/**
 * XeroService
 * -----------------------------------------------------------------
 * Responsible for all Xero business logic:
 *   - OAuth token exchange, refresh & storage
 *   - Calling the Xero API
 *   - Delegating data transformation to XeroMapper
 * -----------------------------------------------------------------
 */
class XeroService {

    /**
     * Exchange the OAuth authorization code for tokens and persist them.
     * @param {string} code - OAuth authorization code
     * @returns {object} tenant object from Xero
     */
    static async exchangeAndSaveToken(code, sessionInfo, userId) {
        const credentials = encodeBasicAuth(config.XERO.CLIENT_ID, config.XERO.CLIENT_SECRET);

        const tokenResponse = await axios.post(
            CONSTANTS.XERO.TOKEN_URL,
            querystring.stringify({
                grant_type:   'authorization_code',
                code,
                redirect_uri: config.XERO.REDIRECT_URI
            }),
            {
                headers: {
                    Authorization:  `Basic ${credentials}`,
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            }
        );

        const tokens = tokenResponse.data;

        const tenantResponse = await axios.get(CONSTANTS.XERO.CONNECTIONS_URL, {
            headers: { Authorization: `Bearer ${tokens.access_token}` }
        });

        const tenants = tenantResponse.data || [];
        if (tenants.length === 0) throw new Error('No Xero organisation connected.');

        // Persist token records for ALL connected organisations/tenants
        await Promise.all(tenants.map(tenant => 
            XeroTokenRepository.upsertToken({
                tenant_id:    tenant.tenantId,
                access_token: tokens.access_token || '',
                refresh_token: tokens.refresh_token || '',
                expires_in:   Math.floor(Date.now() / 1000) + (tokens.expires_in || 0),
                token_type:   tokens.token_type || '',
                scope:        tokens.scope || '',
                session_info: sessionInfo,
                user_id:      userId,
                company_name: tenant.tenantName || 'Xero Organisation',
                // A freshly connected organisation hasn't had a Master Data
                // Pull yet, so it starts "Not Synced" rather than "Active" —
                // pullMasterData flips it to 'Active' once the first pull
                // succeeds.
                status:       'Not Synced'
            })
        ));

        return tenants[0];
    }

    /**
     * Exchange the OAuth code for tokens and fetch all available tenants,
     * but do NOT persist anything to the database yet.
     * Returns { tokens, tenants } for the selection step.
     *
     * @param {string} code - OAuth authorization code
     * @returns {{ tokens: object, tenants: Array }} raw Xero response
     */
    static async exchangeTokensOnly(code) {
        const credentials = encodeBasicAuth(config.XERO.CLIENT_ID, config.XERO.CLIENT_SECRET);

        const tokenResponse = await axios.post(
            CONSTANTS.XERO.TOKEN_URL,
            querystring.stringify({
                grant_type:   'authorization_code',
                code,
                redirect_uri: config.XERO.REDIRECT_URI
            }),
            {
                headers: {
                    Authorization:  `Basic ${credentials}`,
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            }
        );

        const tokens = tokenResponse.data;

        const tenantResponse = await axios.get(CONSTANTS.XERO.CONNECTIONS_URL, {
            headers: { Authorization: `Bearer ${tokens.access_token}` }
        });

        const tenants = tenantResponse.data || [];
        return { tokens, tenants };
    }

    /**
     * Persist tokens only for the tenant IDs that the user explicitly selected.
     *
     * @param {string[]} selectedTenantIds  - Tenant IDs chosen by the user
     * @param {object}   tokens             - Raw Xero token object (access_token, refresh_token, …)
     * @param {Array}    allTenants         - Full list of tenants from Xero /connections
     * @param {string}   userId             - User's FIN ID
     * @param {string}   sessionInfo        - Serialised session
     */
    static async saveSelectedTenants(selectedTenantIds, tokens, allTenants, userId, sessionInfo) {
        const selectedSet = new Set(selectedTenantIds);
        const toSave = allTenants.filter(t => selectedSet.has(t.tenantId));

        if (toSave.length === 0) throw new Error('No valid tenants selected.');

        await Promise.all(toSave.map(tenant =>
            XeroTokenRepository.upsertToken({
                tenant_id:     tenant.tenantId,
                access_token:  tokens.access_token  || '',
                refresh_token: tokens.refresh_token || '',
                expires_in:    Math.floor(Date.now() / 1000) + (tokens.expires_in || 0),
                token_type:    tokens.token_type    || '',
                scope:         tokens.scope         || '',
                session_info:  sessionInfo,
                user_id:       userId,
                company_name:  tenant.tenantName || 'Xero Organisation',
                // See exchangeAndSaveToken above — new connections start
                // 'Not Synced' until their first successful Master Data Pull.
                status:        'Not Synced'
            })
        ));

        return toSave;
    }

    /**
     * Use the stored refresh token to get a new access token.
     * @returns {object} updated token record
     */
    static async refreshAccessToken() {
        const connections = await XeroTokenRepository.getActiveTokens();
        if (!connections || connections.length === 0) throw new Error('Xero is not connected.');

        const token = connections[0];
        const tenantId = token.companyId || token.tenant_id;
        await XeroTokenManager.getValidToken(tenantId);
        return token;
    }

    /**
     * Helper to get list of active tokens for the calling user's connected
     * tenants only.
     * @param {string} [userId] - Owning user's FIN ID; scopes which tenants are returned.
     * @returns {Promise<Array>}
     */
    static async getAllTokens(userId) {
        const tokens = await XeroTokenRepository.getActiveTokens(userId);
        if (!tokens || tokens.length === 0) throw new Error('Xero account is not connected.');
        return tokens;
    }

    /**
     * Standard Xero request headers for one connected tenant. Always goes
     * through XeroTokenManager so an expiring access token is refreshed
     * (and a revoked connection surfaces as a reconnect) rather than being
     * used as-is.
     */
    static async _tenantHeaders(tenantId) {
        const accessToken = await XeroTokenManager.getValidToken(tenantId);
        return {
            Authorization:    `Bearer ${accessToken}`,
            'Xero-Tenant-Id': tenantId,
            Accept:           'application/json'
        };
    }

    /**
     * Resolves the display name for a tenant, preferring the live
     * Organisation record and falling back to the stored company name (and
     * then the tenant id). A failed lookup is non-fatal: the entity fetch
     * it decorates is still worth returning.
     */
    static async _resolveOrgName(token, tenantId, headers) {
        let orgName = token.companyName || tenantId;
        try {
            const orgRes = await axios.get(CONSTANTS.XERO.ORGANISATION_URL, { headers });
            const orgObj = XeroMapper.toOrganisation(orgRes.data);
            if (orgObj && orgObj.name) orgName = orgObj.name;
        } catch (_) {}
        return orgName;
    }

    /**
     * Shared implementation behind getContacts/getAccounts/getClasses/
     * getLocations. Queries one Xero endpoint across every connected tenant
     * for `userId`, in parallel, and tags each record with the tenant id and
     * its organisation name. A tenant whose request fails is logged and
     * contributes no records rather than failing the whole call.
     *
     * @param {string} url - Xero API URL from CONSTANTS.XERO.
     * @param {Function} mapFn - (responseData) => DTO[], e.g. XeroMapper.toContactList.
     * @param {string} logLabel - Plural label used in the per-tenant error log.
     * @param {string} [userId] - Owning user's FIN ID; scopes which tenants are queried.
     * @returns {Promise<object[]>}
     */
    static async _getEntityList(url, mapFn, logLabel, userId) {
        const tokens = await XeroService.getAllTokens(userId);
        const results = await Promise.all(tokens.map(async (token) => {
            try {
                const tenantId = token.companyId || token.tenant_id;
                const headers = await XeroService._tenantHeaders(tenantId);
                const orgName = await XeroService._resolveOrgName(token, tenantId, headers);

                const response = await axios.get(url, { headers });
                const records = mapFn(response.data);
                return records.map(r => ({
                    ...r,
                    clientId: tenantId,
                    clientName: orgName
                }));
            } catch (err) {
                const tenantId = token.companyId || token.tenant_id;
                logger.error(`Error fetching Xero ${logLabel} for tenant ${tenantId}:`, err.message);
                return [];
            }
        }));
        return results.flat();
    }

    /**
     * Fetch all organisation details from Xero across the calling user's
     * connected tenants only.
     * @param {string} [userId] - Owning user's FIN ID; scopes which tenants are queried.
     * @returns {Promise<Array>}
     */
    static async getOrganisation(userId) {
        const tokens = await XeroService.getAllTokens(userId);
        const results = await Promise.all(tokens.map(async (token) => {
            try {
                const tenantId = token.companyId || token.tenant_id;
                const headers = await XeroService._tenantHeaders(tenantId);
                const res = await axios.get(CONSTANTS.XERO.ORGANISATION_URL, { headers });
                const org = XeroMapper.toOrganisation(res.data);
                if (org) {
                    org.id = tenantId;
                    return org;
                }
            } catch (err) {
                const tenantId = token.companyId || token.tenant_id;
                logger.error(`Error fetching Xero organisation for tenant ${tenantId}:`, err.message);
            }
            return null;
        }));
        const orgs = results.filter(Boolean);
        return orgs.length === 1 ? orgs[0] : orgs;
    }

    /**
     * Fetch all contacts from Xero across the calling user's connected
     * tenants only.
     * @param {string} [userId] - Owning user's FIN ID; scopes which tenants are queried.
     * @returns {Promise<ContactDTO[]>}
     */
    static async getContacts(userId) {
        return XeroService._getEntityList(CONSTANTS.XERO.CONTACTS_URL, XeroMapper.toContactList, 'contacts', userId);
    }

    /**
     * Fetch all accounts from Xero across the calling user's connected
     * tenants only.
     * @param {string} [userId] - Owning user's FIN ID; scopes which tenants are queried.
     * @returns {Promise<AccountDTO[]>}
     */
    static async getAccounts(userId) {
        return XeroService._getEntityList(CONSTANTS.XERO.ACCOUNTS_URL, XeroMapper.toAccountList, 'accounts', userId);
    }

    /**
     * Fetch tracking categories for classes from Xero across the calling
     * user's connected tenants only.
     * @param {string} [userId] - Owning user's FIN ID; scopes which tenants are queried.
     * @returns {Promise<ClassDTO[]>}
     */
    static async getClasses(userId) {
        return XeroService._getEntityList(
            CONSTANTS.XERO.TRACKING_CATEGORIES_URL,
            data => XeroMapper.toTrackingList(data, "class"),
            'classes',
            userId
        );
    }

    /**
     * Fetch tracking categories for locations from Xero across the calling
     * user's connected tenants only.
     * @param {string} [userId] - Owning user's FIN ID; scopes which tenants are queried.
     * @returns {Promise<LocationDTO[]>}
     */
    static async getLocations(userId) {
        return XeroService._getEntityList(
            CONSTANTS.XERO.TRACKING_CATEGORIES_URL,
            data => XeroMapper.toTrackingList(data, "location"),
            'locations',
            userId
        );
    }

    // ── Self-contained Connections Management & Pulling ────────────────

    static PLAN_LIMITS = { trial: 1, basic: 1, standard: 3, pro: 10 };

    static getMaxConnections(plan) {
        return XeroService.PLAN_LIMITS[(plan || 'pro').toLowerCase()] ?? 10;
    }

    /**
     * Model + operator handles for the connection-management queries below.
     * Resolved per call rather than at module load so the model registry is
     * only touched once a query actually runs; require() caches, so this is
     * free after the first call.
     */
    static _db() {
        return {
            XeroToken: require('../../core/database').XeroToken,
            Op: require('sequelize').Op
        };
    }

    static async listConnections(userId) {
        const { XeroToken } = XeroService._db();
        const xeroWhere = userId ? { user_id: userId } : {};
        const xeroTokens = await XeroToken.findAll({ where: xeroWhere });

        return xeroTokens.map(t => ({
            platform:     'Xero',
            companyName:  t.company_name || 'Xero Organisation',
            companyId:    t.tenant_id,
            status:       t.status || 'Not Synced',
            // XeroToken doesn't rename its timestamp attributes like
            // QuickBooksToken does — the JS-side properties are the default
            // updatedAt/createdAt, not updated_at/created_at.
            lastSyncedAt: t.last_synced_at || t.updatedAt || null,
            createdAt:    t.createdAt || null
        }));
    }

    static async getConnectionStats(userId, plan) {
        const { XeroToken, Op } = XeroService._db();
        const maxAllowed = XeroService.getMaxConnections(plan);

        const whereClause = { status: { [Op.ne]: 'Disconnected' } };
        if (userId) whereClause.user_id = userId;

        const xeroCount = await XeroToken.count({ where: whereClause });

        return {
            plan: (plan || 'pro').toLowerCase(),
            maxAllowed,
            connected: xeroCount,
            remaining: Math.max(0, maxAllowed - xeroCount)
        };
    }

    /**
     * @param {string} companyId
     * @param {string} userId - Owning user's FIN ID. Required: without it this
     *   would disconnect a company regardless of who owns it.
     */
    static async disconnectConnection(companyId, userId) {
        if (!userId) return false;
        const { XeroToken } = XeroService._db();
        const [updated] = await XeroToken.update(
            { status: 'Disconnected' },
            { where: { tenant_id: companyId, user_id: userId } }
        );
        return updated > 0;
    }

    /**
     * @param {string} companyId
     * @param {string} userId - Owning user's FIN ID. Required — see
     *   disconnectConnection() above for why an ownership check matters here.
     */
    static async activateConnection(companyId, userId) {
        if (!userId) return false;
        const { XeroToken, Op } = XeroService._db();

        // Selecting/switching to a connection re-activates it if it was
        // 'Disconnected', but must not resurrect 'Not Synced' to 'Active' —
        // that transition only happens via a successful Master Data Pull
        // (see pullMasterData).
        const [updated] = await XeroToken.update(
            { status: 'Active' },
            { where: { tenant_id: companyId, user_id: userId, status: { [Op.ne]: 'Not Synced' } } }
        );
        if (updated > 0) return true;

        // If nothing matched, the row might legitimately be 'Not Synced'
        // (or simply not exist) — confirm it exists (and is owned by this
        // user) so the caller still gets a truthy result for "this company
        // is now the active one".
        const existing = await XeroToken.findOne({ where: { tenant_id: companyId, user_id: userId } });
        return !!existing;
    }

    /**
     * @param {string} companyId
     * @param {string} userId - Owning user's FIN ID. Required — see
     *   disconnectConnection() above for why an ownership check matters here.
     */
    static async renameConnection(companyId, userId, companyName) {
        if (!userId) return false;
        const { XeroToken } = XeroService._db();
        const [updated] = await XeroToken.update(
            { company_name: companyName },
            { where: { tenant_id: companyId, user_id: userId } }
        );
        return updated > 0;
    }

    /**
     * @param {string} companyId
     * @param {string} tier
     * @param {string} userId - Owning user's FIN ID. Required — without it a
     *   companyId-scoped pull would return (and let this user overwrite
     *   their Excel sheet with) another user's financial data.
     * @param {boolean} [isIncremental=false] - When true, sends `If-Modified-Since`
     *   on Contacts and Accounts requests so Xero returns only records modified
     *   after `last_synced_at`. TrackingCategories always does a full fetch since
     *   Xero does not support delta filtering on that endpoint.
     */
    static async pullMasterData(companyId, tier, userId, isIncremental = false) {
        if (!userId) return null;
        const { XeroToken, Op } = XeroService._db();
        const maxAllowed = XeroService.getMaxConnections(tier);

        // Include 'Not Synced' connections (not just 'Active') so a
        // freshly connected organisation that has never been pulled yet
        // is still eligible for this — and only this pull is what
        // transitions it to 'Active' below. Both branches are scoped to
        // `userId` so this can only ever touch the calling user's own
        // organisations.
        const rawTokens = companyId
            ? await XeroToken.findAll({ where: { tenant_id: companyId, user_id: userId } })
            : await XeroToken.findAll({ where: { user_id: userId, status: { [Op.ne]: 'Disconnected' } }, order: [['updated_at', 'DESC']] });

        const tokens = rawTokens.slice(0, maxAllowed).map(t => ({
            platform:    'xero',
            companyId:   t.tenant_id,
            companyName: t.company_name || 'Xero Organisation',
            tenant_id:   t.tenant_id,
            lastSyncedAt: t.last_synced_at
        }));

        const results = await Promise.all(tokens.map(async (token) => {
            try {
                // Route through XeroTokenManager for every call, same as
                // XeroService.getContacts()/getAccounts()/etc. This gives us:
                //  - proactive refresh (5 min before expiry) instead of only
                //    reacting to a 401 after the token has already died
                //  - the per-tenant lock, so two concurrent pulls can't both
                //    try to redeem the same (single-use, rotating) refresh
                //    token at once
                //  - XeroTokenRepository's multi-tenant handling, which
                //    updates every tenant row sharing the same refresh token
                //    set — a bespoke single-row update here would silently
                //    orphan a sibling company's refresh token on rotation
                //    and cause a spurious "Reconnect" for it later
                //  - the refresh-token-expired pre-flight check, so a dead
                //    connection fails fast as ErpSessionExpiredError instead
                //    of bouncing off a live 401 first
                //
                // Xero returns HTTP 304 Not Modified when If-Modified-Since
                // is sent and nothing has changed since that timestamp.
                // axios treats 304 as an error by default — validateStatus
                // below accepts it so we can detect "no delta" cleanly
                // instead of falling into the .catch(() => null) path, which
                // would wrongly silence real errors.
                const xeroGet = async (url, extraHeaders = {}) => {
                    const accessToken = await XeroTokenManager.getValidToken(token.companyId);
                    const headers = {
                        Authorization:    `Bearer ${accessToken}`,
                        'Xero-Tenant-Id': token.companyId,
                        Accept:           'application/json',
                        ...extraHeaders
                    };
                    return axios.get(url, {
                        headers,
                        // Accept 304 as a valid status so axios doesn't throw for it.
                        // A 304 means the If-Modified-Since header was honoured and
                        // there are no new records — the caller detects this via
                        // response.status === 304 and treats it as an empty result.
                        validateStatus: (status) => (status >= 200 && status < 300) || status === 304
                    });
                };

                // When isIncremental is true and we have a lastSyncedAt timestamp,
                // pass If-Modified-Since to fetch only contacts/accounts changed since last sync.
                // TrackingCategories does not support If-Modified-Since — always do a full fetch.
                const ifModifiedSince = (isIncremental && token.lastSyncedAt)
                    ? new Date(token.lastSyncedAt).toUTCString()
                    : null;
                const deltaHeaders = ifModifiedSince ? { 'If-Modified-Since': ifModifiedSince } : {};

                const orgSettled = await Promise.allSettled([xeroGet(CONSTANTS.XERO.ORGANISATION_URL)]);
                if (orgSettled[0].status === 'rejected') {
                    const reason = orgSettled[0].reason;

                    // XeroTokenManager already throws ErpSessionExpiredError
                    // (via OAuthTokenRevokedError) once refresh fails or the
                    // refresh token is expired — pass it straight through.
                    if (reason instanceof ErpSessionExpiredError) {
                        throw reason;
                    }

                    // Otherwise this is a raw axios failure from the actual
                    // Xero API call (e.g. the "valid" token was rejected
                    // anyway) — treat auth-shaped errors as a revoked
                    // connection too.
                    if (isAuthError(reason)) {
                        await XeroToken.update(
                            { status: 'Disconnected' },
                            { where: { tenant_id: token.companyId } }
                        );
                        throw new ErpSessionExpiredError(
                            'Xero',
                            `Xero refresh token expired/revoked for company "${token.companyName}" (${token.companyId}): ${reason?.message}`
                        );
                    }
                    throw reason;
                }
                const orgRes = orgSettled[0].value;

                const [contactRes, accRes, classRes] = await Promise.all([
                    xeroGet(CONSTANTS.XERO.CONTACTS_URL, deltaHeaders).catch(() => null),
                    xeroGet(CONSTANTS.XERO.ACCOUNTS_URL, deltaHeaders).catch(() => null),
                    // TrackingCategories: always full fetch — no If-Modified-Since support
                    xeroGet(CONSTANTS.XERO.TRACKING_CATEGORIES_URL).catch(() => null)
                ]);

                const company  = orgRes ? XeroMapper.toOrganisation(orgRes.data) : null;
                const orgName  = company?.name || token.companyName;

                const companyList = company ? [{ ...company, id: token.companyId }] : [];

                const contacts  = contactRes ? XeroMapper.toContactList(contactRes.data, token.lastSyncedAt) : [];
                const accounts  = accRes     ? XeroMapper.toAccountList(accRes.data, token.lastSyncedAt)     : [];
                const classes   = classRes   ? XeroMapper.toTrackingList(classRes.data, 'class', token.lastSyncedAt)    : [];
                const locations = classRes   ? XeroMapper.toTrackingList(classRes.data, 'location', token.lastSyncedAt) : [];

                const tag = items => items.map(i => ({ ...i, clientId: orgName, clientName: orgName }));

                // Captured before the update below overwrites it — lets the
                // frontend's Refresh Schedule flow distinguish "the very
                // first sync for this connection" (nothing to append
                // against yet, write everything) from "a later refresh
                // where nothing happens to be new" — both look identical if
                // you only look at isNew flags, since isNew is false either
                // way.
                const isFirstSync = !token.lastSyncedAt;

                // First successful pull (or any subsequent one) marks the
                // connection 'Active' — this is what takes it out of the
                // initial 'Not Synced' state.
                // Do not update last_synced_at unless every required API succeeds.
                if (contactRes && accRes && classRes) {
                    await XeroToken.update(
                        { last_synced_at: new Date(), status: 'Active' },
                        { where: { tenant_id: token.companyId } }
                    );
                }

                return {
                    company: companyList,
                    customers: tag(contacts.filter(c => c.isCustomer || !c.isSupplier)),
                    vendors: tag(contacts.filter(c => c.isSupplier)),
                    accounts: tag(accounts),
                    classes: tag(classes),
                    locations: tag(locations),
                    isFirstSync
                };
            } catch (err) {
                logger.error(`Error pulling Xero data for connection ${token.companyId}:`, err.message);
                throw err;
            }
        }));

        return results.reduce((acc, curr) => ({
            company: [...acc.company, ...curr.company],
            customers: [...acc.customers, ...curr.customers],
            vendors: [...acc.vendors, ...curr.vendors],
            accounts: [...acc.accounts, ...curr.accounts],
            classes: [...acc.classes, ...curr.classes],
            locations: [...acc.locations, ...curr.locations],
            // See QuickBooksService.pullMasterData's reduce for why this is
            // an AND-merge across connections.
            isFirstSync: acc.isFirstSync && curr.isFirstSync
        }), { company: [], customers: [], vendors: [], accounts: [], classes: [], locations: [], isFirstSync: true });
    }
}

// Register event listener for plan downgrades
const eventBus = require('../../core/events');

eventBus.on('user.downgraded', async ({ userId, email }) => {
    try {
        const { XeroToken, User } = require('../../core/database');
        let targetUserId = userId;
        if (!targetUserId && email) {
            const userObj = await User.findOne({ where: { email } });
            if (userObj) targetUserId = userObj.id;
        }
        if (targetUserId) {
            const deletedCount = await XeroToken.destroy({ where: { user_id: targetUserId } });
            logger.info(`[XeroService] Plan downgrade: cleared ${deletedCount} connections for ${targetUserId}`);
        }
    } catch (err) {
        logger.error(`[XeroService] Failed to clear connections on downgrade for ${userId || email}:`, err.message);
    }
});

module.exports = XeroService;