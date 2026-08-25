'use strict';

const querystring  = require('querystring');
const exceljs      = require('exceljs');
const config       = require('../../core/config');
const CONSTANTS    = require('../../core/constants');
const { generateOAuthState, renderOAuthBlockedPage } = require('../../core/helpers');
const QuickBooksService    = require('./service');
const QuickBooksTokenRepository = require('./repository');
const { ValidationError } = require('../../core/errors/AppError');
const asyncHandler = require('../../core/errors/asyncHandler');

/**
 * QuickbooksController
 * -----------------------------------------------------------------
 * Handles all incoming HTTP requests for the QuickBooks module.
 * Delegates all business logic to QuickBooksService.
 * Handlers are wrapped in asyncHandler, which forwards any rejection to
 * next(err) — that is why they carry no try/catch of their own; the one
 * exception is quickbooksCallback, which needs to translate the failure
 * into a ValidationError before forwarding it.
 * Does NOT contain any data-transformation or mapping logic —
 * that responsibility lives in mapper.js (used by the Service).
 * -----------------------------------------------------------------
 */
class QuickbooksController {

    /**
     * GET /api/quickbooks/connect
     * Generates the QuickBooks OAuth authorization URL and redirects.
     *
     * Requires authentication (the frontend passes the JWT as ?token=...
     * since this is a browser navigation, not a fetch). The owning email
     * is taken exclusively from the verified token (req.user.email), never
     * from a client-suppliable query param — otherwise anyone could
     * initiate a connect flow tagged with someone else's email and inject
     * a connection into that other user's account.
     */
    connectQuickbooks = asyncHandler(async (req, res, next) => {
        const { QuickBooksToken } = require('../../core/database');
        const mail = req.user.email;
        const { Op } = require('sequelize');
        const tier = (req.query.tier || 'pro').toLowerCase();

        let maxAllowed = 10;
        if (tier === 'trial') maxAllowed = 1;
        else if (tier === 'basic') maxAllowed = 1;
        else if (tier === 'standard') maxAllowed = 3;

        // The task pane appends ?reconnectId=<realmId> when the user
        // clicked "Reconnect" on one specific disconnected company. That id
        // is the user's declared intent for this entire round trip, so it is
        // parked in the session here and compared in the callback against
        // whatever realmId Intuit actually returns — the callback must never
        // read it back off its own querystring, which the provider controls.
        const reconnectId = String(req.query.reconnectId || '').trim() || null;

        if (reconnectId) {
            // An id that isn't one of this user's own companies means a
            // stale task pane or a hand-edited URL. Refuse rather than
            // quietly degrading into a normal "add company" flow, since
            // that degradation is precisely the bypass this guards.
            const reconnectTarget = await QuickBooksToken.findOne({ where: { realm_id: reconnectId, mail } });
            if (!reconnectTarget) {
                return res.send(renderOAuthBlockedPage({
                    title: 'Company Not Found',
                    lines: [
                        'The company you tried to reconnect is no longer part of your account.',
                        'Please reload the add-in and try again.'
                    ]
                }));
            }
        } else {
            // Only a genuinely new connection is measured against the plan
            // limit here. A reconnect re-authorizes a company that already
            // occupies one of the plan's slots, so counting it would block
            // the user from restoring a company they are entitled to.
            const whereClause = { status: { [Op.ne]: 'Disconnected' }, mail };
            const qbCount = await QuickBooksToken.count({ where: whereClause });

            if (qbCount >= maxAllowed) {
                return res.send(renderOAuthBlockedPage({
                    title: 'Connection Limit Reached',
                    lines: [
                        `Your subscription tier (${tier.toUpperCase()}) allows a maximum of ${maxAllowed} connected company.`,
                        'Please disconnect an existing company or upgrade your plan to connect more.'
                    ]
                }));
            }
        }

        const state = generateOAuthState();
        req.session.oauth_state = state;
        req.session.user_mail = mail;
        // Carried into the callback so the limit is re-checked against the
        // tier the flow actually started with, rather than a ?tier= the
        // client could swap mid-flight.
        req.session.qb_tier = tier;
        req.session.qb_max_allowed = maxAllowed;
        req.session.qb_reconnect_id = reconnectId;

        const params = {
            client_id:     config.QB.CLIENT_ID,
            response_type: 'code',
            scope:         CONSTANTS.QUICKBOOKS.SCOPES,
            redirect_uri:  config.QB.REDIRECT_URI,
            state
        };

        const authUrl = `${CONSTANTS.QUICKBOOKS.AUTH_URL}?${querystring.stringify(params)}`;
        res.redirect(authUrl);
    });

    /**
     * GET /api/quickbooks/callback
     * Handles the OAuth callback, exchanges code for tokens.
     */
    quickbooksCallback = async (req, res, next) => {
        try {
            const { code, realmId } = req.query;
            const mail = req.session?.user_mail || req.session?.admin?.email || req.session?.googleUser?.email || null;

            // Both values come from the session, written by /connect — the
            // tier so a client cannot widen its own limit halfway through
            // the flow, the reconnect id so Intuit's callback querystring
            // cannot claim an intent the user never expressed.
            const tier       = (req.session?.qb_tier || 'pro').toLowerCase();
            const maxAllowed = req.session?.qb_max_allowed || 10;
            const reconnectId = req.session?.qb_reconnect_id || null;

            // Scenario 2 — strict reconnect validation.
            // The user asked to restore one specific company; Intuit's
            // account picker lets them authorize any company they own. If
            // those disagree, saving the token would quietly add a NEW
            // company under the guise of a reconnect, sidestepping the
            // limit check that a normal "Add Company" flow would have run.
            if (reconnectId && String(realmId) !== String(reconnectId)) {
                delete req.session.qb_reconnect_id;
                return res.send(renderOAuthBlockedPage({
                    title: 'Invalid Company Selected',
                    icon: '🚫',
                    lines: [
                        'You started a reconnect for one specific company, but a different company was authorized in QuickBooks.',
                        'Please click Reconnect again and choose the same company in the Intuit window.'
                    ]
                }));
            }

            // Scenario 1 — lifetime company limit.
            // Disconnecting sets status = 'Disconnected' but keeps the row,
            // so counting every row this user has ever owned (minus the one
            // being authorized right now, which may be an existing row) is
            // what makes the limit a lifetime allowance rather than a
            // concurrent one. Without it, disconnect -> connect a different
            // company is an unlimited carousel on a 1-company plan.
            if (!reconnectId && mail) {
                const { QuickBooksToken } = require('../../core/database');
                const { Op } = require('sequelize');
                const otherCount = await QuickBooksToken.count({
                    where: { mail, realm_id: { [Op.ne]: realmId } }
                });

                if (otherCount + 1 > maxAllowed) {
                    return res.send(renderOAuthBlockedPage({
                        title: 'Connection Limit Reached',
                        lines: [
                            `Your subscription tier (${tier.toUpperCase()}) allows a maximum of ${maxAllowed} company in total.`,
                            `You have already connected ${otherCount} ${otherCount === 1 ? 'company' : 'companies'} on this account.`,
                            'Reconnect one of your existing companies, or upgrade your plan to add a new one.'
                        ]
                    }));
                }
            }

            const sessionInfo = JSON.stringify(req.session || {});
            await QuickBooksService.exchangeAndSaveToken(code, realmId, sessionInfo, mail);
            delete req.session.qb_reconnect_id;
            return res.send(CONSTANTS.QUICKBOOKS.SUCCESS_HTML);
        } catch (error) {
            const details = JSON.stringify(error.response?.data || error.message);
            next(new ValidationError('Failed to connect QuickBooks. Please try again.', details));
        }
    };

    /**
     * GET /api/quickbooks/tokens
     * Returns the authenticated user's own stored QuickBooks OAuth tokens.
     */
    listQuickbooksTokens = asyncHandler(async (req, res, next) => {
        const tokens = await QuickBooksTokenRepository.getAllTokens(req.user.email);
        res.json({ tokens });
    });

    /**
     * GET /api/quickbooks/customers
     * Returns a list of mapped CustomerDTOs for the authenticated user's companies.
     */
    getCustomers = asyncHandler(async (req, res, next) => {
        const customers = await QuickBooksService.getCustomers(req.user.email);
        res.json({ customers });
    });

    /**
     * GET /api/quickbooks/vendors
     * Returns a list of mapped VendorDTOs for the authenticated user's companies.
     */
    getVendors = asyncHandler(async (req, res, next) => {
        const vendors = await QuickBooksService.getVendors(req.user.email);
        res.json({ vendors });
    });

    /**
     * GET /api/quickbooks/accounts
     * Returns a list of mapped AccountDTOs for the authenticated user's companies.
     */
    getAccounts = asyncHandler(async (req, res, next) => {
        const accounts = await QuickBooksService.getAccounts(req.user.email);
        res.json({ accounts });
    });

    /**
     * GET /api/quickbooks/classes
     * Returns a list of mapped ClassDTOs for the authenticated user's companies.
     */
    getClasses = asyncHandler(async (req, res, next) => {
        const classes = await QuickBooksService.getClasses(req.user.email);
        res.json({ classes });
    });

    /**
     * GET /api/quickbooks/locations
     * Returns a list of mapped LocationDTOs for the authenticated user's companies.
     */
    getLocations = asyncHandler(async (req, res, next) => {
        const locations = await QuickBooksService.getLocations(req.user.email);
        res.json({ locations });
    });

    /**
     * GET /api/quickbooks/company
     * Returns company information DTO for the authenticated user's companies.
     */
    getCompanyInfo = asyncHandler(async (req, res, next) => {
        const company = await QuickBooksService.getCompanyInfo(undefined, req.user.email);
        res.json({ company });
    });

    /**
     * GET /api/quickbooks/export
     * Exports company, customers, vendors, accounts, classes, and locations
     * as an Excel file, scoped to the authenticated user's companies.
     *
     * Fetches in 1000-record batches per entity instead of one big
     * unpaginated call per entity. Each batch still fires every entity's
     * API concurrently via Promise.all — same shape as the original
     * single-shot Promise.all below, just repeated batch-by-batch — so
     * customers/vendors/accounts/classes/locations all pull records
     * 1–10 together, then 11–20 together, and so on, until every one of
     * them has exhausted every connected company's data. Company info
     * isn't paginated (one record per company) and is fetched once,
     * up front, alongside a token→orgName lookup so the per-batch entity
     * fetches don't each need their own CompanyInfo round trip.
     */
    exportMasterData = asyncHandler(async (req, res, next) => {
        const mail = req.user.email;
        const BATCH_SIZE = 1000;

        const { tokens: allTokens, company, orgNameByTokenId } =
            await QuickBooksService.getCompanyInfoAndOrgNames(mail)
                .catch(() => ({ tokens: [], company: null, orgNameByTokenId: new Map() }));

        // Per-entity list of tokens still known to have more pages —
        // shrinks independently as each token/company reports its
        // last (short) page, so a company with fewer records simply
        // stops being queried for that entity while others with more
        // data keep going. No two entities share a list, since one
        // entity finishing early for a company must not affect the
        // others' pagination.
        let customersTokens = allTokens.slice();
        let vendorsTokens   = allTokens.slice();
        let accountsTokens  = allTokens.slice();
        let classesTokens   = allTokens.slice();
        let locationsTokens = allTokens.slice();

        const customers = [];
        const vendors   = [];
        const accounts  = [];
        const classes   = [];
        const locations = [];

        const dropExhausted = (list, exhaustedIds) =>
            list.filter(t => !exhaustedIds.has(t.companyId || t.realm_id));
        const allTokenIds = (list) => new Set(list.map(t => t.companyId || t.realm_id));
        const emptyPage = () => ({ records: [], exhaustedTokenIds: new Set() });

        let startPosition = 1;
        let batchCount = 0;
        // Safety valve only — each entity's token list can only
        // shrink every iteration, so the loop is guaranteed to end
        // long before this; it just guards against an infinite loop
        // if that invariant is ever broken by a future change.
        const MAX_BATCHES = 100000;

        while (
            (customersTokens.length || vendorsTokens.length || accountsTokens.length ||
             classesTokens.length || locationsTokens.length) &&
            batchCount < MAX_BATCHES
        ) {
            batchCount += 1;
            const batchStartedAt = Date.now();

            // ── TEMPORARY diagnostic logging ─────────────────────
            // Logged synchronously for every still-active entity
            // BEFORE the Promise.all below is even constructed, so
            // all of a batch's START lines print together as one
            // group regardless of how long each entity's HTTP
            // response actually takes.
            if (customersTokens.length) console.log(`[BATCH ${batchCount}][Customers] START position=${startPosition} limit=${BATCH_SIZE}`);
            if (vendorsTokens.length) console.log(`[BATCH ${batchCount}][Vendors] START position=${startPosition} limit=${BATCH_SIZE}`);
            if (accountsTokens.length) console.log(`[BATCH ${batchCount}][Accounts] START position=${startPosition} limit=${BATCH_SIZE}`);
            if (classesTokens.length) console.log(`[BATCH ${batchCount}][Classes] START position=${startPosition} limit=${BATCH_SIZE}`);
            if (locationsTokens.length) console.log(`[BATCH ${batchCount}][Locations] START position=${startPosition} limit=${BATCH_SIZE}`);

            // All 5 entity APIs for this batch run together — nothing
            // here waits for another to finish first.
            const [customersResult, vendorsResult, accountsResult, classesResult, locationsResult] = await Promise.all([
                customersTokens.length
                    ? QuickBooksService.getCustomersPage(customersTokens, startPosition, BATCH_SIZE, orgNameByTokenId)
                        .catch(() => ({ records: [], exhaustedTokenIds: allTokenIds(customersTokens) }))
                    : Promise.resolve(emptyPage()),
                vendorsTokens.length
                    ? QuickBooksService.getVendorsPage(vendorsTokens, startPosition, BATCH_SIZE, orgNameByTokenId)
                        .catch(() => ({ records: [], exhaustedTokenIds: allTokenIds(vendorsTokens) }))
                    : Promise.resolve(emptyPage()),
                accountsTokens.length
                    ? QuickBooksService.getAccountsPage(accountsTokens, startPosition, BATCH_SIZE, orgNameByTokenId)
                        .catch(() => ({ records: [], exhaustedTokenIds: allTokenIds(accountsTokens) }))
                    : Promise.resolve(emptyPage()),
                classesTokens.length
                    ? QuickBooksService.getClassesPage(classesTokens, startPosition, BATCH_SIZE, orgNameByTokenId)
                        .catch(() => ({ records: [], exhaustedTokenIds: allTokenIds(classesTokens) }))
                    : Promise.resolve(emptyPage()),
                locationsTokens.length
                    ? QuickBooksService.getLocationsPage(locationsTokens, startPosition, BATCH_SIZE, orgNameByTokenId)
                        .catch(() => ({ records: [], exhaustedTokenIds: allTokenIds(locationsTokens) }))
                    : Promise.resolve(emptyPage())
            ]);

            console.log(`[BATCH ${batchCount}][Customers] RESPONSE count=${customersResult.records.length} (+${Date.now() - batchStartedAt}ms since this batch's requests started)`);
            console.log(`[BATCH ${batchCount}][Vendors] RESPONSE count=${vendorsResult.records.length} (+${Date.now() - batchStartedAt}ms since this batch's requests started)`);
            console.log(`[BATCH ${batchCount}][Accounts] RESPONSE count=${accountsResult.records.length} (+${Date.now() - batchStartedAt}ms since this batch's requests started)`);
            console.log(`[BATCH ${batchCount}][Classes] RESPONSE count=${classesResult.records.length} (+${Date.now() - batchStartedAt}ms since this batch's requests started)`);
            console.log(`[BATCH ${batchCount}][Locations] RESPONSE count=${locationsResult.records.length} (+${Date.now() - batchStartedAt}ms since this batch's requests started)`);

            customers.push(...customersResult.records);
            vendors.push(...vendorsResult.records);
            accounts.push(...accountsResult.records);
            classes.push(...classesResult.records);
            locations.push(...locationsResult.records);

            customersTokens = dropExhausted(customersTokens, customersResult.exhaustedTokenIds);
            vendorsTokens   = dropExhausted(vendorsTokens, vendorsResult.exhaustedTokenIds);
            accountsTokens  = dropExhausted(accountsTokens, accountsResult.exhaustedTokenIds);
            classesTokens   = dropExhausted(classesTokens, classesResult.exhaustedTokenIds);
            locationsTokens = dropExhausted(locationsTokens, locationsResult.exhaustedTokenIds);

            startPosition += BATCH_SIZE;
        }

        const wb = new exceljs.Workbook();

        if (company) {
            const wsCompany = wb.addWorksheet('Company');
            wsCompany.addRow(['ID', 'Company Name', 'Legal Name']);
            wsCompany.addRow([company.id, company.name, company.legalName]);
        }

        const wsCustomers = wb.addWorksheet('Customers');
        wsCustomers.addRow(['ID', 'Name', 'Company Name', 'Email', 'Balance']);
        wsCustomers.addRows(customers.map(c => [c.id, c.name, c.companyName, c.email, c.balance]));

        const wsVendors = wb.addWorksheet('Vendors');
        wsVendors.addRow(['ID', 'Name', 'Company Name', 'Email', 'Balance']);
        wsVendors.addRows(vendors.map(v => [v.id, v.name, v.companyName, v.email, v.balance]));

        const wsAccounts = wb.addWorksheet('Accounts');
        wsAccounts.addRow(['ID', 'Acct #', 'Name', 'Account Type', 'Sub Type', 'Balance']);
        wsAccounts.addRows(accounts.map(a => [a.id, a.acctNum, a.name, a.accountType, a.accountSubType, a.currentBalance]));

        const wsClasses = wb.addWorksheet('Classes');
        wsClasses.addRow(['ID', 'Name', 'Status']);
        wsClasses.addRows(classes.map(c => [c.id, c.name, c.active]));

        const wsLocations = wb.addWorksheet('Locations');
        wsLocations.addRow(['ID', 'Name', 'Status']);
        wsLocations.addRows(locations.map(l => [l.id, l.name, l.active]));

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="quickbooks_master_data.xlsx"');

        await wb.xlsx.write(res);
        res.end();
    });

    /**
     * POST /api/quickbooks/disconnect
     * Clears the authenticated user's own stored QuickBooks tokens only.
     */
    disconnectQuickbooks = asyncHandler(async (req, res, next) => {
        await QuickBooksTokenRepository.clearTokens(req.user.email);
        res.json({ success: true, message: 'QuickBooks tokens cleared successfully.' });
    });

    /**
     * GET /api/quickbooks/connections
     */
    listConnections = asyncHandler(async (req, res, next) => {
        const mail = req.user.email;
        const list = await QuickBooksService.listConnections(mail);
        return res.json(list);
    });

    /**
     * GET /api/quickbooks/connections/stats
     */
    getConnectionStats = asyncHandler(async (req, res, next) => {
        const mail = req.user.email;
        const plan = req.query.plan || 'pro';

        const stats = {
            plan: plan.toLowerCase(),
            maxPerPlatform: 10,
            quickbooks: { connected: 0, remaining: 10 }
        };

        if (plan === 'trial')    stats.maxPerPlatform = 1;
        else if (plan === 'basic')    stats.maxPerPlatform = 1;
        else if (plan === 'standard') stats.maxPerPlatform = 3;

        const qbStats = await QuickBooksService.getConnectionStats(mail, plan);
        stats.quickbooks = {
            connected: qbStats.connected,
            remaining: qbStats.remaining
        };

        return res.json(stats);
    });

    /**
     * DELETE /api/quickbooks/connections/:id
     */
    disconnectConnection = asyncHandler(async (req, res, next) => {
        const companyId = req.params.id;
        const success = await QuickBooksService.disconnectConnection(companyId, req.user.email);
        return res.json({ success: !!success });
    });

    /**
     * POST /api/quickbooks/connections/:id/activate
     */
    activateConnection = asyncHandler(async (req, res, next) => {
        const companyId = req.params.id;
        const success = await QuickBooksService.activateConnection(companyId, req.user.email);

        let totalRecords = 0;
        if (success) {
            try {
                const token = { companyId, realm_id: companyId };
                const countInfo = await QuickBooksService.getTotalRecordCountsForToken(token);
                totalRecords = countInfo.total;
            } catch (err) {}
        }

        return res.json({ success: !!success, totalRecords });
    });

    /**
     * PATCH /api/quickbooks/connections/:id/rename
     */
    renameConnection = asyncHandler(async (req, res, next) => {
        const companyId = req.params.id;
        const { companyName } = req.body;
        if (!companyName) {
            throw new ValidationError('companyName is required.');
        }

        const success = await QuickBooksService.renameConnection(companyId, req.user.email, companyName);
        return res.json({ success: !!success });
    });

    /**
     * GET /api/quickbooks/pull-master-data?companyId=...&tier=...&cursor=...
     *
     * `cursor` (optional) is the JSON-encoded per-company, per-entity
     * pagination cursor returned as `cursor` in a PREVIOUS call's
     * response body. Omitting it (or sending {}) starts a fresh cycle at
     * Accounts, position 1.
     *
     * Each call fetches exactly ONE page (up to 100 records) for exactly
     * ONE entity — the entities are processed sequentially, in the order
     * Accounts -> Classes -> Locations -> Customers -> Vendors, and an
     * entity is drained completely before the next one starts (from its
     * own first record). So on any given call four of the five record
     * arrays in the response are empty, and no two APIs are ever fetched
     * at the same time. See QuickBooksService.pullMasterData /
     * _fetchOnePageForToken.
     *
     * A single button click therefore never silently pulls more than one
     * batch; the caller must send the returned `cursor` back on its next
     * click to continue, and `isDone: true` means every entity is
     * exhausted and there is nothing left to fetch until the cycle is
     * reset.
     */
    /**
     * GET /api/quickbooks/pull-master-data?companyId=...&tier=...&stream=...
     *
     * Single-Click Auto-Pull Master Data Endpoint.
     * Fetches 100% of data across all entities (Accounts, Classes, Locations, Customers, Vendors)
     * using multithreaded async worker streams and auto-tuned pagination in a single request lifecycle.
     * Includes HTTP Keep-Alive stream pings to prevent proxy timeouts.
     */
    pullMasterData = asyncHandler(async (req, res, next) => {
        const { companyId, tier, mode, stream } = req.query;
        const isIncremental = mode === 'incremental';

        if (stream === 'true' || req.headers.accept?.includes('text/event-stream')) {
            res.setHeader('Content-Type', 'text/event-stream');
            res.setHeader('Cache-Control', 'no-cache');
            res.setHeader('Connection', 'keep-alive');
            res.setHeader('X-Content-Type-Options', 'nosniff');
            res.status(200);

            const heartbeatInterval = setInterval(() => {
                res.write(': heartbeat ping\n\n');
            }, 15000);

            try {
                const onProgress = (event) => {
                    res.write(`data: ${JSON.stringify(event)}\n\n`);
                };

                const aggregated = await QuickBooksService.pullMasterDataMultithreaded(companyId, tier, req.user.email, onProgress, isIncremental);
                clearInterval(heartbeatInterval);

                res.write(`data: ${JSON.stringify({ type: 'complete', data: aggregated })}\n\n`);
                return res.end();
            } catch (err) {
                clearInterval(heartbeatInterval);
                res.write(`data: ${JSON.stringify({ type: 'error', error: err.message })}\n\n`);
                return res.end();
            }
        }

        const aggregated = await QuickBooksService.pullMasterDataMultithreaded(companyId, tier, req.user.email, null, isIncremental);

        if (!aggregated) {
            const { AppError } = require('../../core/errors/AppError');
            throw new AppError('The requested resource was not found.', 404, 'ERR_NOT_FOUND', `No active connections found for quickbooks.`);
        }

        return res.json({
            company:   aggregated.company.length === 1 ? aggregated.company[0] : aggregated.company,
            customers: aggregated.customers,
            vendors:   aggregated.vendors,
            accounts:  aggregated.accounts,
            classes:   aggregated.classes,
            locations: aggregated.locations,
            isFirstSync: aggregated.isFirstSync,
            isDone: true
        });
    });

    /**
     * GET /api/quickbooks/refresh-incremental?companyId=...&tier=...
     *
     * Incremental Refresh Endpoint.
     * Fetches ONLY modified or newly added records since the last sync timestamp (`MetaData.LastUpdatedTime`).
     */
    refreshIncremental = asyncHandler(async (req, res, next) => {
        const { companyId, tier } = req.query;
        const aggregated = await QuickBooksService.pullMasterDataMultithreaded(companyId, tier, req.user.email, null, true);

        if (!aggregated) {
            const { AppError } = require('../../core/errors/AppError');
            throw new AppError('No active connection found for refresh.', 404, 'ERR_NOT_FOUND', 'QuickBooks connection not found.');
        }

        return res.json({
            company:   aggregated.company.length === 1 ? aggregated.company[0] : aggregated.company,
            customers: aggregated.customers,
            vendors:   aggregated.vendors,
            accounts:  aggregated.accounts,
            classes:   aggregated.classes,
            locations: aggregated.locations,
            isFirstSync: false,
            isIncremental: true,
            isDone: true
        });
    });
}

module.exports = new QuickbooksController();
