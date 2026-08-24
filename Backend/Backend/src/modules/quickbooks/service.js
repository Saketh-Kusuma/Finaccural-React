'use strict';

const axios        = require('axios');
const querystring  = require('querystring');
const config       = require('../../core/config');
const CONSTANTS    = require('../../core/constants');
const { encodeBasicAuth } = require('../../core/helpers');
const QuickBooksTokenRepository = require('./repository');
const QuickBooksMapper = require('./mapper');
const logger       = require('../../core/logger');
const QuickBooksTokenManager = require('./oauth/QuickBooksTokenManager');
const { AppError, ErpSessionExpiredError } = require('../../core/errors/AppError');

/**
 * QuickBooksService
 * -----------------------------------------------------------------
 * Responsible for all QuickBooks business logic:
 *   - OAuth token exchange & storage
 *   - Querying the QB API
 *   - Delegating data transformation to QuickBooksMapper
 * -----------------------------------------------------------------
 */
class QuickBooksService {

    /**
     * Exchange the OAuth authorization code for tokens, query CompanyInfo, and persist connection.
     * @param {string} code   - OAuth authorization code
     * @param {string} realmId - QB company ID
     */
    static async exchangeAndSaveToken(code, realmId, sessionInfo, mail) {
        const credentials = encodeBasicAuth(config.QB.CLIENT_ID, config.QB.CLIENT_SECRET);

        const response = await axios.post(
            CONSTANTS.QUICKBOOKS.TOKEN_URL,
            querystring.stringify({
                grant_type:   'authorization_code',
                code,
                redirect_uri: config.QB.REDIRECT_URI
            }),
            {
                headers: {
                    Accept:         'application/json',
                    'Content-Type': 'application/x-www-form-urlencoded',
                    Authorization:  `Basic ${credentials}`
                }
            }
        );

        const tokenData = response.data;

        // Fetch CompanyInfo using the newly issued access token directly,
        // via a raw request rather than executeQuery/QuickBooksTokenManager.
        // Those look up whatever token is already stored for this realmId,
        // which — on a reconnect — can be a stale/expired one, causing this
        // step to fail with a 401 immediately after a successful exchange.
        let companyName = 'QuickBooks Company';
        try {
            const url = `${CONSTANTS.QUICKBOOKS.BASE_URL}/v3/company/${realmId}/query`;
            const compRes = await axios.get(url, {
                headers: {
                    Authorization: `Bearer ${tokenData.access_token}`,
                    Accept: 'application/json',
                    'Content-Type': 'application/text'
                },
                params: { query: 'SELECT * FROM CompanyInfo' }
            });
            const compInfo = QuickBooksMapper.toCompanyInfo(compRes.data);
            companyName = compInfo ? (compInfo.name || compInfo.legalName || realmId) : 'QuickBooks Company';
        } catch (compErr) {
            logger.warn(`Could not fetch company info directly during OAuth exchange for realm ${realmId}:`, compErr.message);
        }

        await QuickBooksTokenRepository.upsertToken({
            realm_id: realmId,
            access_token: tokenData.access_token || '',
            refresh_token: tokenData.refresh_token || '',
            token_type: tokenData.token_type || '',
            expires_in: Math.floor(Date.now() / 1000) + (tokenData.expires_in || 0),
            x_refresh_token_expires_in: Math.floor(Date.now() / 1000) + (tokenData.x_refresh_token_expires_in || 0),
            session_info: sessionInfo,
            mail: mail,
            company_name: companyName,
            // A freshly connected company hasn't had a Master Data Pull yet,
            // so it starts "Not Synced" rather than "Active" — pullMasterData
            // flips it to 'Active' once the first pull succeeds.
            status: 'Not Synced'
        });
    }

    /**
     * Execute a raw QBQL query against the QuickBooks API.
     * @param {string} query - QuickBooks SQL-like query string
     * @param {object} [token] - Specific QuickBooks Token record to use
     * @returns {object} raw API response
     */
    static async executeQuery(query, token) {
        let realmId;
        let accessToken;

        // TEMPORARY concurrency-verification instrumentation: tags every
        // query with its entity/STARTPOSITION so the [QB-HTTP] timestamps
        // below prove whether concurrent calls really overlapped. Safe to
        // delete once concurrency is confirmed.
        const qbEntityMatch = /FROM\s+(\w+)/i.exec(query);
        const qbPosMatch = /STARTPOSITION\s+(\d+)/i.exec(query);
        const qbLabel = `${qbEntityMatch ? qbEntityMatch[1] : 'query'}${qbPosMatch ? ` @${qbPosMatch[1]}` : ''}`;
        const qbCallStart = Date.now();

        if (token) {
            realmId = token.companyId || token.realm_id;
            // Always go through QuickBooksTokenManager rather than falling
            // back to token.access_token/accessToken on failure. That token
            // is precisely the one getValidToken() just decided was
            // expiring/expired — silently using it anyway would mean a
            // failed/refused refresh (revoked connection) gets masked as a
            // doomed API call instead of surfacing as "Reconnect" here.
            accessToken = await QuickBooksTokenManager.getValidToken(realmId);
        } else {
            const connections = await QuickBooksTokenRepository.getActiveTokens();
            const activeToken = connections[0];
            if (!activeToken) {
                throw new ErpSessionExpiredError('QuickBooks', 'No active QuickBooks connection found.');
            }
            realmId = activeToken.companyId || activeToken.realm_id;
            accessToken = await QuickBooksTokenManager.getValidToken(realmId);
        }

        console.log(`[QB-HTTP] ${new Date().toISOString()} TOKEN READY     ${qbLabel} realm=${realmId} (+${Date.now() - qbCallStart}ms since executeQuery() was called)`);

        const url = `${CONSTANTS.QUICKBOOKS.BASE_URL}/v3/company/${realmId}/query`;

        const qbHttpStart = Date.now();
        console.log(`[QB-HTTP] ${new Date(qbHttpStart).toISOString()} REQUEST START   ${qbLabel} realm=${realmId}`);

        try {
            const response = await axios.get(url, {
                headers: {
                    Authorization:  `Bearer ${accessToken}`,
                    Accept:         'application/json',
                    'Content-Type': 'application/text'
                },
                params: { query }
            });
            console.log(`[QB-HTTP] ${new Date().toISOString()} RESPONSE OK     ${qbLabel} realm=${realmId} (+${Date.now() - qbHttpStart}ms)`);
            return response.data;
        } catch (error) {
            console.log(`[QB-HTTP] ${new Date().toISOString()} RESPONSE ERROR  ${qbLabel} realm=${realmId} (+${Date.now() - qbHttpStart}ms)`);
            logger.error(`Error executing QB query for realm ${realmId}:`, error.response?.data || error.message);
            throw error;
        }
    }

    /**
     * Fetches every record of `entityName` for a token, paging through
     * STARTPOSITION/MAXRESULTS until QuickBooks returns fewer than
     * `batchSize` records.
     *
     * QuickBooks' query API defaults to only the first 100 records when
     * MAXRESULTS is omitted, and caps MAXRESULTS at 1000 per request, so
     * any entity type with more records than that gets silently truncated
     * unless paged through like this. Returns a QueryResponse-shaped
     * object so existing QuickBooksMapper.toXList() calls work unchanged.
     *
     * @param {string} entityName - QBQL entity name, e.g. "Customer".
     * @param {object} token
     * @param {number} [batchSize=1000] - QuickBooks' MAXRESULTS hard cap.
     * @returns {Promise<{ QueryResponse: Object }>}
     */
    static async queryAll(entityName, token, batchSize = 1000) {
        const fetchBatch = async (startPosition) => {
            const query = `SELECT * FROM ${entityName} STARTPOSITION ${startPosition} MAXRESULTS ${batchSize}`;
            const raw = await QuickBooksService.executeQuery(query, token);
            const batch = raw?.QueryResponse?.[entityName] || [];

            if (batch.length < batchSize) {
                return batch;
            }
            const nextBatch = await fetchBatch(startPosition + batchSize);
            return batch.concat(nextBatch);
        };

        const allRecords = await fetchBatch(1);
        return { QueryResponse: { [entityName]: allRecords } };
    }

    /**
     * Executes an API function call with exponential backoff and jitter for rate-limiting (429 / 503).
     * Zero loops used — relies on pure async functional recursion.
     */
    static async executeWithRetryAndBackoff(fn, retries = 5, delay = 500) {
        try {
            return await fn();
        } catch (err) {
            const isRateLimited = err.response?.status === 429 || err.status === 429 || err.statusCode === 429 || err.code === 'THROTTLED';
            if (isRateLimited && retries > 0) {
                const jitter = Math.random() * 200;
                const backoffMs = delay * 2 + jitter;
                logger.warn(`[QB Rate Limit 429] Backing off for ${Math.round(backoffMs)}ms. Retries left: ${retries}`);
                await new Promise(resolve => setTimeout(resolve, backoffMs));
                return QuickBooksService.executeWithRetryAndBackoff(fn, retries - 1, delay * 2);
            }
            throw err;
        }
    }

    /**
     * Pre-flight query to count total records for all entity types of a token.
     * Uses fast light `SELECT COUNT(*)` queries across Customer, Vendor, Account, Class, Department.
     */
    static async getTotalRecordCountsForToken(token) {
        const entityNames = ['Customer', 'Vendor', 'Account', 'Class', 'Department'];
        const countResults = await Promise.all(entityNames.map(async (entityName) => {
            try {
                const raw = await QuickBooksService.executeQuery(`SELECT COUNT(*) FROM ${entityName}`, token);
                const count = raw?.QueryResponse?.totalCount || raw?.QueryResponse?.[entityName]?.length || 0;
                return { entityName, count };
            } catch (err) {
                return { entityName, count: 0 };
            }
        }));

        return countResults.reduce((acc, curr) => {
            acc[curr.entityName] = curr.count;
            acc.total += curr.count;
            return acc;
        }, { Customer: 0, Vendor: 0, Account: 0, Class: 0, Department: 0, total: 0 });
    }

    /**
     * Pure loopless async functional recursion to page through an entity with dynamic auto-tuned batch sizing.
     * Stack-safe due to V8 microtask queue yielding across await boundaries.
     */
    static async fetchEntityPagesRecursiveAutoTuned(entityName, token, startPosition = 1, currentBatchSize = 500, accumulatedRecords = [], onChunkCallback = null, updatedSince = null) {
        const startTime = Date.now();
        
        let query = `SELECT * FROM ${entityName} STARTPOSITION ${startPosition} MAXRESULTS ${currentBatchSize}`;
        if (updatedSince) {
            query = `SELECT * FROM ${entityName} WHERE MetaData.LastUpdatedTime >= '${updatedSince}' STARTPOSITION ${startPosition} MAXRESULTS ${currentBatchSize}`;
        }

        const pageResult = await QuickBooksService.executeWithRetryAndBackoff(() => 
            QuickBooksService.executeQuery(query, token)
        );

        const elapsedTime = Date.now() - startTime;
        const records = pageResult?.QueryResponse?.[entityName] || [];
        const updatedAccumulated = accumulatedRecords.concat(records);

        if (typeof onChunkCallback === 'function' && records.length > 0) {
            onChunkCallback(entityName, records.length);
        }

        // Base case: no more records or received fewer than batch size
        if (records.length < currentBatchSize) {
            return updatedAccumulated;
        }

        // Auto-tune next batch size based on execution time (target ~800ms)
        const targetMs = 800;
        const scaleFactor = elapsedTime > 0 ? targetMs / elapsedTime : 1.0;
        const nextBatchSize = Math.max(50, Math.min(1000, Math.floor(currentBatchSize * Math.min(Math.max(scaleFactor, 0.5), 2.0))));
        const nextPosition = startPosition + records.length;

        // Recursive step without loops
        return QuickBooksService.fetchEntityPagesRecursiveAutoTuned(
            entityName,
            token,
            nextPosition,
            nextBatchSize,
            updatedAccumulated,
            onChunkCallback,
            updatedSince
        );
    }

    /**
     * Single-Click Multithreaded Full Data Sync Engine.
     * Pulls 100% of data across all entities concurrently without any manual user refresh clicks.
     * Invokes onProgress callback with live count and percentage progress updates.
     */
    static async pullMasterDataMultithreaded(companyId, tier, mail, onProgress = null, isIncremental = false) {
        if (!mail) return null;
        const { QuickBooksToken, Op } = QuickBooksService._db();
        const maxAllowed = QuickBooksService.getMaxConnections(tier);

        const rawTokens = companyId
            ? await QuickBooksToken.findAll({ where: { realm_id: companyId, mail } })
            : await QuickBooksToken.findAll({ where: { mail, status: { [Op.ne]: 'Disconnected' } }, order: [['updated_at', 'DESC']] });

        const tokens = rawTokens.slice(0, maxAllowed).map(t => ({
            platform:     'quickbooks',
            companyId:    t.realm_id,
            companyName:  t.company_name || 'QuickBooks Company',
            realm_id:     t.realm_id,
            lastSyncedAt: t.last_synced_at
        }));

        if (!tokens || tokens.length === 0) return null;

        const results = await Promise.all(tokens.map(async (token) => {
            const rawComp = await QuickBooksService.executeQuery('SELECT * FROM CompanyInfo', token).catch(() => null);
            const comp = QuickBooksMapper.toCompanyInfo(rawComp);
            const companyList = comp ? [{ ...comp, id: token.companyId }] : [];
            const orgName = comp?.name || comp?.legalName || token.companyName;

            // Step 1: Pre-flight Count Query
            const countInfo = await QuickBooksService.getTotalRecordCountsForToken(token);
            let grandTotal = countInfo.total;
            let totalFetchedSoFar = 0;

            console.log(`[MASTER SYNC START] Realm=${token.companyId} Org="${orgName}" TotalRecords=${grandTotal}`);
            if (typeof onProgress === 'function') {
                onProgress({ type: 'start', totalRecords: grandTotal, companyName: orgName, companyId: token.companyId });
            }

            const onChunk = (entityName, chunkSize) => {
                totalFetchedSoFar += chunkSize;
                const percentage = grandTotal > 0 ? Math.min(100, Math.round((totalFetchedSoFar / grandTotal) * 100)) : 100;
                console.log(`[SYNC PROGRESS] ${percentage}% (${totalFetchedSoFar}/${grandTotal} records) entity=${entityName}`);
                if (typeof onProgress === 'function') {
                    onProgress({ type: 'progress', percentage, fetchedRecords: totalFetchedSoFar, totalRecords: grandTotal, currentEntity: entityName });
                }
            };

            const updatedSinceFilter = (isIncremental && token.lastSyncedAt)
                ? new Date(token.lastSyncedAt).toISOString()
                : null;

            // Step 2: Parallel Stream Worker Pool across all 5 Entities
            const entityList = ['Account', 'Class', 'Department', 'Customer', 'Vendor'];
            const [accountsRaw, classesRaw, locationsRaw, customersRaw, vendorsRaw] = await Promise.all([
                QuickBooksService.fetchEntityPagesRecursiveAutoTuned('Account', token, 1, 500, [], onChunk, updatedSinceFilter),
                QuickBooksService.fetchEntityPagesRecursiveAutoTuned('Class', token, 1, 500, [], onChunk, updatedSinceFilter),
                QuickBooksService.fetchEntityPagesRecursiveAutoTuned('Department', token, 1, 500, [], onChunk, updatedSinceFilter),
                QuickBooksService.fetchEntityPagesRecursiveAutoTuned('Customer', token, 1, 500, [], onChunk, updatedSinceFilter),
                QuickBooksService.fetchEntityPagesRecursiveAutoTuned('Vendor', token, 1, 500, [], onChunk, updatedSinceFilter)
            ]);

            const rawCust  = { QueryResponse: { Customer:   customersRaw } };
            const rawVend  = { QueryResponse: { Vendor:     vendorsRaw } };
            const rawAcc   = { QueryResponse: { Account:    accountsRaw } };
            const rawClass = { QueryResponse: { Class:      classesRaw } };
            const rawLoc   = { QueryResponse: { Department: locationsRaw } };

            const tag = (list) => list.map(i => ({ ...i, clientId: orgName, clientName: orgName }));
            const isFirstSync = !token.lastSyncedAt;

            await QuickBooksToken.update(
                { last_synced_at: new Date(), status: 'Active' },
                { where: { realm_id: token.companyId } }
            );

            return {
                company: companyList,
                customers: tag(QuickBooksMapper.toCustomerList(rawCust, token.lastSyncedAt)),
                vendors: tag(QuickBooksMapper.toVendorList(rawVend, token.lastSyncedAt)),
                accounts: tag(QuickBooksMapper.toAccountList(rawAcc, token.lastSyncedAt)),
                classes: tag(QuickBooksMapper.toClassList(rawClass, token.lastSyncedAt)),
                locations: tag(QuickBooksMapper.toLocationList(rawLoc, token.lastSyncedAt)),
                isFirstSync,
                isDone: true
            };
        }));

        const aggregated = results.reduce((acc, curr) => ({
            company: [...acc.company, ...curr.company],
            customers: [...acc.customers, ...curr.customers],
            vendors: [...acc.vendors, ...curr.vendors],
            accounts: [...acc.accounts, ...curr.accounts],
            classes: [...acc.classes, ...curr.classes],
            locations: [...acc.locations, ...curr.locations],
            isFirstSync: acc.isFirstSync && curr.isFirstSync,
            isDone: true
        }), { company: [], customers: [], vendors: [], accounts: [], classes: [], locations: [], isFirstSync: true, isDone: true });

        return aggregated;
    }

    /**
     * Fetches exactly ONE page of `entityName` for a single token via
     * STARTPOSITION/MAXRESULTS — the single-page counterpart to
     * queryAll() above, which recurses through every page internally
     * and returns everything at once. This is what lets a caller (see
     * exportMasterData's batch loop) drive pagination one 10-record
     * batch at a time instead of waiting for a full recursive fetch.
     * @param {string} entityName - QBQL entity name, e.g. "Customer".
     * @param {object} token
     * @param {number} startPosition - 1-based, QuickBooks STARTPOSITION.
     * @param {number} pageSize - QuickBooks MAXRESULTS for this page.
     * @returns {Promise<{ raw: object, records: object[], hasMore: boolean }>}
     */
    static async queryPage(entityName, token, startPosition, pageSize) {
        const query = `SELECT * FROM ${entityName} STARTPOSITION ${startPosition} MAXRESULTS ${pageSize}`;
        const raw = await QuickBooksService.executeQuery(query, token);
        const records = raw?.QueryResponse?.[entityName] || [];
        return { raw, records, hasMore: records.length === pageSize };
    }

    /**
     * Fetch company info and return clean CompanyDTO for a specific token or
     * all of the calling user's tokens.
     * @param {object} [token]
     * @param {string} [mail] - Owning user's email; scopes which companies are queried when `token` isn't given.
     * @returns {CompanyDTO|CompanyDTO[]|null}
     */
    static async getCompanyInfo(token, mail) {
        if (token) {
            try {
                const raw = await QuickBooksService.executeQuery('SELECT * FROM CompanyInfo', token);
                return QuickBooksMapper.toCompanyInfo(raw);
            } catch (err) {
                return null;
            }
        }

        const tokens = await QuickBooksTokenRepository.getActiveTokens(mail);
        if (!tokens || tokens.length === 0) return null;

        const companyResults = await Promise.all(tokens.map(async (t) => {
            try {
                const raw = await QuickBooksService.executeQuery('SELECT * FROM CompanyInfo', t);
                const info = QuickBooksMapper.toCompanyInfo(raw);
                if (info) {
                    info.id = t.companyId;
                    return info;
                }
            } catch (err) {}
            return null;
        }));

        const companies = companyResults.filter(Boolean);
        return companies.length === 1 ? companies[0] : companies;
    }

    /**
     * Helper to get company name and ID for tagging records of a specific token.
     * @param {object} token
     * @returns {Promise<{ orgId: string, orgName: string }>}
     */
    static async getCompanyMetadata(token) {
        const company = await QuickBooksService.getCompanyInfo(token).catch(() => null);
        const realmId = token.companyId || token.realm_id;
        const orgId   = company ? (company.id || company.name || realmId) : realmId;
        const orgName = company ? (company.name || company.legalName || company.id || "QuickBooks Company") : "QuickBooks Company";
        return { orgId, orgName };
    }

    /**
     * Fetches every active token's CompanyInfo ONCE and returns both the
     * "Company" worksheet data and a token→orgName lookup in a single
     * pass. Used by exportMasterData's batch loop so its per-batch,
     * per-entity page fetches (getCustomersPage/getVendorsPage/etc.,
     * below) can reuse an already-known org name instead of each one
     * separately re-fetching CompanyInfo per token per batch — the way
     * getCustomers()/getVendors()/etc. do today (once per entity type,
     * via getCompanyMetadata) is fine for a single non-paginated call,
     * but would mean 5x redundant CompanyInfo calls per token per batch
     * if repeated on every page.
     * @param {string} mail - Owning user's email; scopes which companies are queried.
     * @returns {Promise<{ tokens: object[], company: CompanyDTO|CompanyDTO[]|null, orgNameByTokenId: Map<string, string> }>}
     */
    static async getCompanyInfoAndOrgNames(mail) {
        const tokens = await QuickBooksTokenRepository.getActiveTokens(mail);
        const orgNameByTokenId = new Map();

        // Match getCompanyInfo(undefined, mail)'s own no-connections
        // behavior exactly (returns null, not []) — exportMasterData's
        // `if (company) {...}` check depends on that, and an empty
        // array is truthy in JS.
        if (!tokens || tokens.length === 0) {
            return { tokens: [], company: null, orgNameByTokenId };
        }

        const companies = (await Promise.all(tokens.map(async (token) => {
            const tokenId = token.companyId || token.realm_id;
            const info = await QuickBooksService.getCompanyInfo(token).catch(() => null);
            if (info) info.id = token.companyId;
            const orgName = info ? (info.name || info.legalName || info.id || "QuickBooks Company") : "QuickBooks Company";
            orgNameByTokenId.set(tokenId, orgName);
            return info;
        }))).filter(Boolean);

        const company = companies.length === 1 ? companies[0] : companies;
        return { tokens, company, orgNameByTokenId };
    }

    /**
     * Shared implementation behind getCustomers/getVendors/getAccounts/
     * getClasses/getLocations. Queries one QBQL entity across every active
     * token for `mail`, in parallel, and tags each record with the owning
     * company's org name. A token whose query fails is logged and
     * contributes no records rather than failing the whole call.
     *
     * @param {string} entityName - QBQL entity, e.g. "Customer".
     * @param {Function} mapperFn - QuickBooksMapper.toXList, e.g. toCustomerList.
     * @param {string} logLabel - Plural label used in the per-token error log.
     * @param {string} mail - Owning user's email; scopes which companies are queried.
     * @returns {Promise<object[]>}
     */
    static async _getEntityList(entityName, mapperFn, logLabel, mail) {
        const tokens = await QuickBooksTokenRepository.getActiveTokens(mail);
        const results = await Promise.all(tokens.map(async (token) => {
            try {
                const raw = await QuickBooksService.queryAll(entityName, token);
                const list = mapperFn(raw);
                const { orgName } = await QuickBooksService.getCompanyMetadata(token);
                return list.map(item => ({
                    ...item,
                    clientId: orgName,
                    clientName: orgName
                }));
            } catch (err) {
                const realmId = token.companyId || token.realm_id;
                logger.error(`Error getting ${logLabel} for realm ${realmId}:`, err.message);
                return [];
            }
        }));
        return results.flat();
    }

    /**
     * Fetch all customers and return clean CustomerDTOs across the calling
     * user's connected companies only.
     * @param {string} mail - Owning user's email; scopes which companies are queried.
     * @returns {CustomerDTO[]}
     */
    static async getCustomers(mail) {
        return QuickBooksService._getEntityList('Customer', QuickBooksMapper.toCustomerList, 'customers', mail);
    }

    /**
     * Fetch all vendors and return clean VendorDTOs across the calling
     * user's connected companies only.
     * @param {string} mail - Owning user's email; scopes which companies are queried.
     * @returns {VendorDTO[]}
     */
    static async getVendors(mail) {
        return QuickBooksService._getEntityList('Vendor', QuickBooksMapper.toVendorList, 'vendors', mail);
    }

    /**
     * Fetch all accounts and return clean AccountDTOs across the calling
     * user's connected companies only.
     * @param {string} mail - Owning user's email; scopes which companies are queried.
     * @returns {AccountDTO[]}
     */
    static async getAccounts(mail) {
        return QuickBooksService._getEntityList('Account', QuickBooksMapper.toAccountList, 'accounts', mail);
    }

    /**
     * Fetch all classes and return clean ClassDTOs across the calling
     * user's connected companies only.
     * @param {string} mail - Owning user's email; scopes which companies are queried.
     * @returns {ClassDTO[]}
     */
    static async getClasses(mail) {
        return QuickBooksService._getEntityList('Class', QuickBooksMapper.toClassList, 'classes', mail);
    }

    /**
     * Fetch all locations (departments) and return clean LocationDTOs
     * across the calling user's connected companies only.
     * @param {string} mail - Owning user's email; scopes which companies are queried.
     * @returns {LocationDTO[]}
     */
    static async getLocations(mail) {
        return QuickBooksService._getEntityList('Department', QuickBooksMapper.toLocationList, 'departments', mail);
    }

    // ── Paginated (batch) entity fetchers ───────────────────────────────
    //
    // Counterparts to getCustomers/getVendors/getAccounts/getClasses/
    // getLocations above, for callers that want to drive pagination one
    // page at a time (see exportMasterData's batch loop) instead of
    // getting every record back in a single call. Each one fetches at
    // most `pageSize` records per still-active token, in parallel, and
    // reports back which tokens have run out of pages so the caller can
    // stop querying them on the next batch — no per-token pagination
    // state is kept here between calls, that's the caller's job.

    /**
     * Shared implementation behind getCustomersPage/getVendorsPage/etc.
     * @param {string} entityName - QBQL entity, e.g. "Customer".
     * @param {Function} mapperFn - QuickBooksMapper.toXList, e.g. toCustomerList.
     * @param {object[]} activeTokens - Tokens still known to have more pages for this entity.
     * @param {number} startPosition - 1-based, QuickBooks STARTPOSITION.
     * @param {number} pageSize - QuickBooks MAXRESULTS for this page.
     * @param {Map<string,string>} orgNameByTokenId - From getCompanyInfoAndOrgNames(), so each
     *   page doesn't need its own CompanyInfo round trip just to tag clientId/clientName.
     * @returns {Promise<{ records: object[], exhaustedTokenIds: Set<string> }>}
     */
    static async _queryEntityPage(entityName, mapperFn, activeTokens, startPosition, pageSize, orgNameByTokenId) {
        const exhaustedTokenIds = new Set();
        const perToken = await Promise.all(activeTokens.map(async (token) => {
            const tokenId = token.companyId || token.realm_id;
            try {
                const { raw, records, hasMore } = await QuickBooksService.queryPage(entityName, token, startPosition, pageSize);
                if (!hasMore) exhaustedTokenIds.add(tokenId);
                const list = mapperFn(raw);
                const orgName = (orgNameByTokenId && orgNameByTokenId.get(tokenId)) || "QuickBooks Company";
                return list.map(item => ({ ...item, clientId: orgName, clientName: orgName }));
            } catch (err) {
                logger.error(`Error getting ${entityName} page (start ${startPosition}) for realm ${tokenId}:`, err.message);
                // A token whose page request failed isn't retried on the
                // next batch — same "log it, return empty, move on"
                // contract as the non-paginated getters' per-token catch.
                exhaustedTokenIds.add(tokenId);
                return [];
            }
        }));
        return { records: perToken.flat(), exhaustedTokenIds };
    }

    /** Paginated counterpart to getCustomers() — see _queryEntityPage. */
    static async getCustomersPage(activeTokens, startPosition, pageSize, orgNameByTokenId) {
        return QuickBooksService._queryEntityPage('Customer', QuickBooksMapper.toCustomerList, activeTokens, startPosition, pageSize, orgNameByTokenId);
    }

    /** Paginated counterpart to getVendors() — see _queryEntityPage. */
    static async getVendorsPage(activeTokens, startPosition, pageSize, orgNameByTokenId) {
        return QuickBooksService._queryEntityPage('Vendor', QuickBooksMapper.toVendorList, activeTokens, startPosition, pageSize, orgNameByTokenId);
    }

    /** Paginated counterpart to getAccounts() — see _queryEntityPage. */
    static async getAccountsPage(activeTokens, startPosition, pageSize, orgNameByTokenId) {
        return QuickBooksService._queryEntityPage('Account', QuickBooksMapper.toAccountList, activeTokens, startPosition, pageSize, orgNameByTokenId);
    }

    /** Paginated counterpart to getClasses() — see _queryEntityPage. */
    static async getClassesPage(activeTokens, startPosition, pageSize, orgNameByTokenId) {
        return QuickBooksService._queryEntityPage('Class', QuickBooksMapper.toClassList, activeTokens, startPosition, pageSize, orgNameByTokenId);
    }

    /** Paginated counterpart to getLocations() — see _queryEntityPage. */
    static async getLocationsPage(activeTokens, startPosition, pageSize, orgNameByTokenId) {
        return QuickBooksService._queryEntityPage('Department', QuickBooksMapper.toLocationList, activeTokens, startPosition, pageSize, orgNameByTokenId);
    }

    /**
     * Fetches every page of the 5 paginated entities (Customer, Vendor,
     * Account, Class, Department) for a SINGLE token, `pageSize` records at
     * a time: all still-active entities of a batch fire together via
     * Promise.all and the next batch only starts once that one resolves.
     * queryAll's own internal paging (MAXRESULTS up to 1000) cannot produce
     * that batch-by-batch, 10-records-at-a-time shape, which is why the
     * bulk export path uses this instead.
     *
     * Each entity independently drops out of later batches once it returns
     * a short page — Vendor can stop after 15 records while Customer, with
     * 47, keeps paging. Returns QueryResponse-shaped objects so the
     * QuickBooksMapper.toXList(raw, lastSyncedAt) calls in pullMasterData
     * work unchanged (isNew/isUpdated flagging included).
     *
     * @param {object} token
     * @param {number} [pageSize=10]
     * @returns {Promise<{ Customer: object[], Vendor: object[], Account: object[], Class: object[], Department: object[] }>}
     */
    static async _fetchAllPaginatedEntitiesForToken(token, pageSize = 10) {
        const entities = ['Customer', 'Vendor', 'Account', 'Class', 'Department'];
        const entityLabel = { Customer: 'Customers', Vendor: 'Vendors', Account: 'Accounts', Class: 'Classes', Department: 'Locations' };
        const recordsByEntity = { Customer: [], Vendor: [], Account: [], Class: [], Department: [] };
        let active = entities.slice();
        let startPosition = 1;
        let batchNumber = 0;
        const realmId = token.companyId || token.realm_id;

        while (active.length > 0) {
            batchNumber += 1;
            const batchStart = Date.now();

            // TEMPORARY diagnostic logging: every still-active entity's
            // START line prints synchronously BEFORE this batch's
            // queryPage() promises are created, so interleaved START and
            // RESPONSE lines would expose a batch being awaited serially.
            active.forEach(entityName => {
                console.log(`[BATCH ${batchNumber}][${entityLabel[entityName]}] START position=${startPosition} limit=${pageSize} realm=${realmId}`);
            });

            // All still-active entities for this token start together —
            // nothing here waits for another to finish first. The array
            // passed to Promise.all is built via .map(), which invokes
            // queryPage() for every active entity synchronously (each
            // call starts its HTTP request immediately); Promise.all only
            // waits for all of them together, it does not serialize them.
            const pages = await Promise.all(active.map(entityName =>
                QuickBooksService.queryPage(entityName, token, startPosition, pageSize)
            ));

            const stillActive = [];
            active.forEach((entityName, i) => {
                const { records, hasMore } = pages[i];
                console.log(`[BATCH ${batchNumber}][${entityLabel[entityName]}] RESPONSE count=${records.length} realm=${realmId} (+${Date.now() - batchStart}ms since this batch's requests started)`);
                recordsByEntity[entityName].push(...records);
                if (hasMore) stillActive.push(entityName);
            });
            active = stillActive;

            startPosition += pageSize;
        }

        return recordsByEntity;
    }

    /**
     * Fixed order in which the paginated master-data APIs are pulled.
     * ONE API is fully drained (PULL_PAGE_SIZE records at a time, click by click)
     * before the next one is even touched, and each API always starts
     * back at its own first record — QuickBooks `STARTPOSITION 1` — no
     * matter how far the previous API had paged.
     *
     * Accounts -> Classes -> Locations (QuickBooks `Department`) ->
     * Customers -> Vendors.
     */
    static SEQUENTIAL_ENTITY_ORDER = ['Account', 'Class', 'Department', 'Customer', 'Vendor'];

    /**
     * Records fetched per Pull Master Data / Refresh Schedule click.
     *
     * One click still means exactly one QuickBooks request for exactly one
     * entity — this only changes how much that single request asks for
     * (MAXRESULTS), so a full dataset is drained in ~10x fewer clicks than
     * the previous size of 10. QuickBooks caps MAXRESULTS at 1000.
     *
     * Safe to change mid-cycle: a cursor stored by the add-in under the old
     * size holds an absolute STARTPOSITION, so the next click simply
     * continues from that position with the new page size. No records are
     * skipped or repeated.
     */
    static PULL_PAGE_SIZE = 100;

    /**
     * Fetches exactly ONE page (up to pageSize records) for exactly ONE
     * entity per call — the click-scoped counterpart to
     * _fetchAllPaginatedEntitiesForToken above, which loops internally
     * until every entity is exhausted in one call. This is what lets
     * pullMasterData make exactly one real MAXRESULTS=PULL_PAGE_SIZE QuickBooks
     * request per Pull Master Data / Refresh Schedule click, instead of
     * silently fetching the whole dataset in one HTTP call and only
     * showing part of it — the click IS the pagination trigger, driven
     * by a caller-supplied per-entity cursor rather than an internal
     * while loop.
     *
     * Entities are processed STRICTLY SEQUENTIALLY, in
     * SEQUENTIAL_ENTITY_ORDER, each on a "start at position 1, drain
     * fully, hand over" basis. At most one entity's request is ever in
     * flight, so two APIs can never overlap. (Contrast
     * _fetchAllPaginatedEntitiesForToken, which runs all five
     * concurrently for the bulk export path.)
     *
     * A click never returns an empty batch just because the entity it
     * landed on was exhausted: a zero-record page closes that entity's
     * cursor and the SAME click moves on to the next entity in the order.
     *
     * @param {object} token
     * @param {{[entity: string]: {position: number, done: boolean}}|null} priorCursor
     *   Per-entity cursor returned by this function on a previous click
     *   for this same token, or null/undefined to start a fresh cycle
     *   (Accounts at position 1, nothing done yet).
     * @param {number} [pageSize=PULL_PAGE_SIZE]
     * @returns {Promise<{
     *   recordsByEntity: { Customer: object[], Vendor: object[], Account: object[], Class: object[], Department: object[] },
     *   cursor: {[entity: string]: {position: number, done: boolean}},
     *   isDone: boolean
     * }>}
     */
    static async _fetchOnePageForToken(token, priorCursor, pageSize = QuickBooksService.PULL_PAGE_SIZE) {
        const entities = QuickBooksService.SEQUENTIAL_ENTITY_ORDER;
        const entityLabel = { Account: 'Accounts', Class: 'Classes', Department: 'Locations', Customer: 'Customers', Vendor: 'Vendors' };
        const realmId = token.companyId || token.realm_id;
        const safePriorCursor = priorCursor && typeof priorCursor === 'object' ? priorCursor : {};

        const recordsByEntity = { Customer: [], Vendor: [], Account: [], Class: [], Department: [] };
        const nextCursor = { ...safePriorCursor };

        const isEntityDone = (entityName) => !!(nextCursor[entityName] && nextCursor[entityName].done);

        // Walk the fixed order and stop at the first entity that still
        // has records left — that single entity owns this whole click.
        // The loop only ever iterates more than once when the entity it
        // lands on turns out to be exhausted (a zero-record page), in
        // which case that entity is closed out and the next one in the
        // order takes over the click rather than wasting it.
        for (const entityName of entities) {
            if (isEntityDone(entityName)) continue;

            // Each API starts from its own first record. A cursor entry
            // only exists once THIS entity has been paged before, so a
            // freshly reached entity always begins at STARTPOSITION 1 —
            // the previous entity's position never carries over.
            const startPosition = (safePriorCursor[entityName] && safePriorCursor[entityName].position) || 1;

            console.log(`[PAGE][${entityLabel[entityName]}] START position=${startPosition} limit=${pageSize} realm=${realmId}`);
            const { records, hasMore } = await QuickBooksService.queryPage(entityName, token, startPosition, pageSize);
            console.log(`[PAGE][${entityLabel[entityName]}] RESPONSE count=${records.length} realm=${realmId}`);

            nextCursor[entityName] = {
                position: startPosition + pageSize,
                done: !hasMore
            };

            if (records.length === 0) {
                // This entity had nothing left (its record count was an
                // exact multiple of pageSize, so the previous click
                // couldn't tell it was finished). It's now closed out;
                // hand this click to the next API instead of returning
                // an empty "Batch written".
                console.log(`[PAGE][${entityLabel[entityName]}] COMPLETED realm=${realmId} — moving to next API`);
                continue;
            }

            recordsByEntity[entityName] = records;
            if (!hasMore) {
                console.log(`[PAGE][${entityLabel[entityName]}] COMPLETED realm=${realmId}`);
            }

            // Exactly one API's batch per click — nothing else is
            // fetched, even though later entities are still pending.
            return {
                recordsByEntity,
                cursor: nextCursor,
                isDone: entities.every(isEntityDone)
            };
        }

        // Every entity in the order finished — either on earlier clicks
        // or just now, in the loop above.
        return { recordsByEntity, cursor: nextCursor, isDone: true };
    }

    // ── Self-contained Connections Management & Pulling ────────────────

    static PLAN_LIMITS = { trial: 1, basic: 1, standard: 3, pro: 10 };

    static getMaxConnections(plan) {
        return QuickBooksService.PLAN_LIMITS[(plan || 'pro').toLowerCase()] ?? 10;
    }

    /**
     * Model + operator handles for the connection-management queries below.
     * Resolved per call rather than at module load so the model registry is
     * only touched once a query actually runs; require() caches, so this is
     * free after the first call.
     */
    static _db() {
        return {
            QuickBooksToken: require('../../core/database').QuickBooksToken,
            Op: require('sequelize').Op
        };
    }

    static async listConnections(mail) {
        const { QuickBooksToken } = QuickBooksService._db();
        const qbWhere = mail ? { mail } : {};
        const qbTokens = await QuickBooksToken.findAll({ where: qbWhere });

        return qbTokens.map(t => ({
            platform:     'QuickBooks',
            companyName:  t.company_name || 'QuickBooks Company',
            companyId:    t.realm_id,
            status:       t.status || 'Not Synced',
            lastSyncedAt: t.last_synced_at || t.updated_at || null,
            createdAt:    t.created_at || null
        }));
    }

    static async getConnectionStats(mail, plan) {
        const { QuickBooksToken, Op } = QuickBooksService._db();
        const maxAllowed = QuickBooksService.getMaxConnections(plan);

        const whereClause = { status: { [Op.ne]: 'Disconnected' } };
        if (mail) whereClause.mail = mail;

        const qbCount = await QuickBooksToken.count({ where: whereClause });

        return {
            plan: (plan || 'pro').toLowerCase(),
            maxAllowed,
            connected: qbCount,
            remaining: Math.max(0, maxAllowed - qbCount)
        };
    }

    /**
     * @param {string} companyId
     * @param {string} mail - Owning user's email. Required: without it this
     *   would disconnect a company regardless of who owns it, letting any
     *   authenticated user tear down another user's connection just by
     *   knowing/guessing its companyId.
     */
    static async disconnectConnection(companyId, mail) {
        if (!mail) return false;
        const { QuickBooksToken } = QuickBooksService._db();
        const [updated] = await QuickBooksToken.update(
            { status: 'Disconnected' },
            { where: { realm_id: companyId, mail } }
        );
        return updated > 0;
    }

    /**
     * @param {string} companyId
     * @param {string} mail - Owning user's email. Required — see
     *   disconnectConnection() above for why an ownership check matters here.
     */
    static async activateConnection(companyId, mail) {
        if (!mail) return false;
        const { QuickBooksToken, Op } = QuickBooksService._db();

        // Selecting/switching to a connection re-activates it if it was
        // 'Disconnected', but must not resurrect 'Not Synced' to 'Active' —
        // that transition only happens via a successful Master Data Pull
        // (see pullMasterData).
        const [updated] = await QuickBooksToken.update(
            { status: 'Active' },
            { where: { realm_id: companyId, mail, status: { [Op.ne]: 'Not Synced' } } }
        );
        if (updated > 0) return true;

        // If nothing matched, the row might legitimately be 'Not Synced'
        // (or simply not exist) — confirm it exists (and is owned by this
        // user) so the caller still gets a truthy result for "this company
        // is now the active one".
        const existing = await QuickBooksToken.findOne({ where: { realm_id: companyId, mail } });
        return !!existing;
    }

    /**
     * @param {string} companyId
     * @param {string} mail - Owning user's email. Required — see
     *   disconnectConnection() above for why an ownership check matters here.
     */
    static async renameConnection(companyId, mail, companyName) {
        if (!mail) return false;
        const { QuickBooksToken } = QuickBooksService._db();
        const [updated] = await QuickBooksToken.update(
            { company_name: companyName },
            { where: { realm_id: companyId, mail } }
        );
        return updated > 0;
    }

    /**
     * @param {string} companyId
     * @param {string} tier
     * @param {string} mail - Owning user's email. Required — without it a
     *   companyId-scoped pull would return (and let this user overwrite
     *   their Excel sheet with) another user's financial data, and a
     *   bulk (no companyId) pull would aggregate every user's connections
     *   in the system into one response.
     * @param {{[companyId: string]: {[entity: string]: {position: number, done: boolean}}}} [cursorByCompany]
     *   Per-company, per-entity pagination cursor returned by this same
     *   function on a previous click, or {}/undefined to start a fresh
     *   cycle. Each call fetches exactly ONE page (up to PULL_PAGE_SIZE records) for
     *   exactly ONE entity per token, walking the APIs strictly in order
     *   — Accounts, then Classes, then Locations, then Customers, then
     *   Vendors, each drained completely and each starting from its own
     *   first record — see _fetchOnePageForToken. So the caller (the
     *   Pull Master Data / Refresh Schedule click handler) controls
     *   pagination one click at a time instead of this function eagerly
     *   fetching everything and the frontend slicing it.
     */
    static async pullMasterData(companyId, tier, mail, cursorByCompany) {
        if (!mail) return null;
        const { QuickBooksToken, Op } = QuickBooksService._db();
        const maxAllowed = QuickBooksService.getMaxConnections(tier);

        // Exclude 'Disconnected' connections from the bulk (no companyId)
        // pull — a revoked/expired connection stops being retried
        // automatically the moment it's marked disconnected; it only comes
        // back once the user reconnects. 'Not Synced' connections are still
        // included since they've never had a chance to sync yet. Both
        // branches are scoped to `mail` so this can only ever touch the
        // calling user's own companies.
        const rawTokens = companyId
            ? await QuickBooksToken.findAll({ where: { realm_id: companyId, mail } })
            : await QuickBooksToken.findAll({ where: { mail, status: { [Op.ne]: 'Disconnected' } }, order: [['updated_at', 'DESC']] });

        const tokens = rawTokens.slice(0, maxAllowed).map(t => ({
            platform:     'quickbooks',
            companyId:    t.realm_id,
            companyName:  t.company_name || 'QuickBooks Company',
            realm_id:     t.realm_id,
            lastSyncedAt: t.last_synced_at
        }));

        const safeCursorByCompany = cursorByCompany && typeof cursorByCompany === 'object' ? cursorByCompany : {};

        const results = await Promise.all(tokens.map(async (token) => {
            try {
                const rawComp = await QuickBooksService.executeQuery('SELECT * FROM CompanyInfo', token);
                const comp = QuickBooksMapper.toCompanyInfo(rawComp);
                const companyList = comp ? [{ ...comp, id: token.companyId }] : [];

                // Exactly ONE page (up to PULL_PAGE_SIZE records) of exactly ONE
                // entity for THIS click — the APIs are drained one at a
                // time in a fixed order, so four of the five lists below
                // are always empty on any given click. See
                // _fetchOnePageForToken. CompanyInfo above is
                // deliberately outside pagination: it's a single record
                // per company, refetched every click regardless of where
                // the entity pagination cursor is.
                const priorCursor = safeCursorByCompany[token.companyId] || null;
                const pageResult = await QuickBooksService._fetchOnePageForToken(token, priorCursor, QuickBooksService.PULL_PAGE_SIZE);
                const pagedEntities = pageResult.recordsByEntity;
                const rawCust  = { QueryResponse: { Customer:   pagedEntities.Customer } };
                const rawVend  = { QueryResponse: { Vendor:     pagedEntities.Vendor } };
                const rawAcc   = { QueryResponse: { Account:    pagedEntities.Account } };
                const rawClass = { QueryResponse: { Class:      pagedEntities.Class } };
                const rawLoc   = { QueryResponse: { Department: pagedEntities.Department } };

                const orgName = comp?.name || comp?.legalName || token.companyName;
                const tag = (list) => list.map(i => ({ ...i, clientId: orgName, clientName: orgName }));

                // Captured before the update below overwrites it — this is
                // what lets the frontend's Refresh Schedule flow tell "the
                // very first sync for this connection" (nothing to append
                // against yet, write everything) apart from "a later
                // refresh where nothing happens to be new" (both look
                // identical if you only look at isNew flags, since every
                // record's isNew is false in both cases).
                const isFirstSync = !token.lastSyncedAt;

                // last_synced_at only advances once this token's cycle has
                // actually finished (every entity's cursor is done) — not
                // on every intermediate single-page click. That keeps
                // isNew/isUpdated flagging (toXList compares each click's
                // page against token.lastSyncedAt) anchored to the last
                // COMPLETE sync across every click of an in-progress
                // cycle, and keeps isFirstSync accurate for every click of
                // a first cycle instead of flipping to false after click 1.
                await QuickBooksToken.update(
                    pageResult.isDone
                        ? { last_synced_at: new Date(), status: 'Active' }
                        : { status: 'Active' },
                    { where: { realm_id: token.companyId } }
                );

                return {
                    company: companyList,
                    customers: tag(QuickBooksMapper.toCustomerList(rawCust, token.lastSyncedAt)),
                    vendors: tag(QuickBooksMapper.toVendorList(rawVend, token.lastSyncedAt)),
                    accounts: tag(QuickBooksMapper.toAccountList(rawAcc, token.lastSyncedAt)),
                    classes: tag(QuickBooksMapper.toClassList(rawClass, token.lastSyncedAt)),
                    locations: tag(QuickBooksMapper.toLocationList(rawLoc, token.lastSyncedAt)),
                    isFirstSync,
                    companyIdForCursor: token.companyId,
                    cursor: pageResult.cursor,
                    isDone: pageResult.isDone
                };
            } catch (err) {
                logger.error(`Error pulling QB data for connection ${token.companyId}:`, err.message);

                // Check for QuickBooks Subscription Expired / Suspended (Code 8020)
                const faultError = err.response?.data?.Fault?.Error?.[0];
                if (faultError?.code === '8020' || (faultError?.Message && faultError.Message.includes('Subscription is not active'))) {
                    await QuickBooksToken.update(
                        { status: 'Disconnected' },
                        { where: { realm_id: token.companyId } }
                    );
                    throw new AppError(
                        'Your QuickBooks subscription has expired or been suspended. Please log into QuickBooks to update your billing.',
                        403,
                        'ERR_QB_SUBSCRIPTION_EXPIRED'
                    );
                }

                const isTokenError = err.response?.status === 401
                    || err.statusCode === 401
                    || (err.message && (err.message.includes('Token expired') || err.message.includes('401') || err.message.includes('grant')));

                if (isTokenError || err.message?.includes('OAuth')) {
                    await QuickBooksToken.update(
                        { status: 'Disconnected' },
                        { where: { realm_id: token.companyId } }
                    );
                    throw new ErpSessionExpiredError(
                        'QuickBooks',
                        `QuickBooks refresh token expired/revoked for company "${token.companyName}" (${token.companyId}): ${err.message}`
                    );
                }

                throw err;
            }
        }));

        const aggregated = results.reduce((acc, curr) => ({
            company: [...acc.company, ...curr.company],
            customers: [...acc.customers, ...curr.customers],
            vendors: [...acc.vendors, ...curr.vendors],
            accounts: [...acc.accounts, ...curr.accounts],
            classes: [...acc.classes, ...curr.classes],
            locations: [...acc.locations, ...curr.locations],
            // A companyId-scoped pull (the normal case) is always exactly
            // one connection, so this just reflects that one token. For the
            // rare bulk (no companyId) pull spanning several connections,
            // AND-merging means "first sync" only holds if every one of
            // them is — mixing "never synced" and "already synced" here
            // would otherwise leave it ambiguous which single answer to
            // give the frontend for a mixed batch.
            isFirstSync: acc.isFirstSync && curr.isFirstSync,
            // Same AND-merge logic applied to pagination completeness:
            // a bulk pull only counts as "fully done" for this click's
            // cycle once every one of its tokens' entities are exhausted.
            isDone: acc.isDone && !!curr.isDone
        }), { company: [], customers: [], vendors: [], accounts: [], classes: [], locations: [], isFirstSync: true, isDone: true });

        // Per-company cursor map the caller must send back on its next
        // click to continue this cycle where this click left off (or to
        // detect that every token is already done and no next click is
        // needed until the cycle is manually reset).
        aggregated.cursor = {};
        results.forEach(r => {
            if (r && r.companyIdForCursor) {
                aggregated.cursor[r.companyIdForCursor] = r.cursor;
            }
        });

        return aggregated;
    }
}

// Register event listener for plan downgrades
const eventBus = require('../../core/events');
const { QuickBooksToken } = require('../../core/database');

eventBus.on('user.downgraded', async ({ email }) => {
    try {
        const deletedCount = await QuickBooksToken.destroy({ where: { mail: email } });
        logger.info(`[QuickBooksService] Plan downgrade: cleared ${deletedCount} connections for ${email}`);
    } catch (err) {
        logger.error(`[QuickBooksService] Failed to clear connections on downgrade for ${email}:`, err.message);
    }
});

module.exports = QuickBooksService;
