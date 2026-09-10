'use strict';

const axios = require('axios');
const CONSTANTS = require('../../../core/constants');
const XeroMapper = require('../mapper');
const XeroApiClient = require('./xeroApiClient');
const XeroTokenManager = require('../oauth/XeroTokenManager');
const logger = require('../../../core/logger');
const { ErpSessionExpiredError } = require('../../../core/errors/AppError');

class XeroConnectionService {
    static PLAN_LIMITS = { trial: 1, basic: 1, standard: 3, pro: 10 };

    static getMaxConnections(plan) {
        return XeroConnectionService.PLAN_LIMITS[(plan || 'pro').toLowerCase()] ?? 10;
    }

    static _db() {
        return {
            XeroToken: require('../../../core/database').XeroToken,
            Op: require('sequelize').Op
        };
    }

    static async listConnections(userId) {
        const { XeroToken } = XeroConnectionService._db();
        const xeroWhere = userId ? { user_id: userId } : {};
        const xeroTokens = await XeroToken.findAll({ where: xeroWhere });

        return xeroTokens.map(t => ({
            platform:     'Xero',
            companyName:  t.company_name || 'Xero Organisation',
            companyId:    t.tenant_id,
            status:       t.status || 'Not Synced',
            lastSyncedAt: t.last_synced_at || t.updatedAt || null,
            createdAt:    t.createdAt || null
        }));
    }

    static async getConnectionStats(userId, plan) {
        const { XeroToken, Op } = XeroConnectionService._db();
        const maxAllowed = XeroConnectionService.getMaxConnections(plan);

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

    static async disconnectConnection(companyId, userId) {
        if (!userId) return false;
        const { XeroToken } = XeroConnectionService._db();
        const [updated] = await XeroToken.update(
            { status: 'Disconnected' },
            { where: { tenant_id: companyId, user_id: userId } }
        );
        return updated > 0;
    }

    static async activateConnection(companyId, userId) {
        if (!userId) return false;
        const { XeroToken, Op } = XeroConnectionService._db();

        const [updated] = await XeroToken.update(
            { status: 'Active' },
            { where: { tenant_id: companyId, user_id: userId, status: { [Op.ne]: 'Not Synced' } } }
        );
        if (updated > 0) return true;

        const existing = await XeroToken.findOne({ where: { tenant_id: companyId, user_id: userId } });
        return !!existing;
    }

    static async renameConnection(companyId, userId, companyName) {
        if (!userId) return false;
        const { XeroToken } = XeroConnectionService._db();
        const [updated] = await XeroToken.update(
            { company_name: companyName },
            { where: { tenant_id: companyId, user_id: userId } }
        );
        return updated > 0;
    }

    static async pullMasterData(companyId, tier, userId, isIncremental = false) {
        if (!userId) return null;
        const { XeroToken, Op } = XeroConnectionService._db();
        const maxAllowed = XeroConnectionService.getMaxConnections(tier);

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
                        validateStatus: (status) => (status >= 200 && status < 300) || status === 304
                    });
                };

                const ifModifiedSince = (isIncremental && token.lastSyncedAt)
                    ? new Date(token.lastSyncedAt).toUTCString()
                    : null;
                const deltaHeaders = ifModifiedSince ? { 'If-Modified-Since': ifModifiedSince } : {};

                const orgSettled = await Promise.allSettled([xeroGet(CONSTANTS.XERO.ORGANISATION_URL)]);
                if (orgSettled[0].status === 'rejected') {
                    const reason = orgSettled[0].reason;

                    if (reason instanceof ErpSessionExpiredError) {
                        throw reason;
                    }

                    if (XeroApiClient.isAuthError(reason)) {
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
                const isFirstSync = !token.lastSyncedAt;

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
            isFirstSync: acc.isFirstSync && curr.isFirstSync
        }), { company: [], customers: [], vendors: [], accounts: [], classes: [], locations: [], isFirstSync: true });
    }
}

module.exports = XeroConnectionService;
