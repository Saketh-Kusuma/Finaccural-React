'use strict';

const axios = require('axios');
const CONSTANTS = require('../../../core/constants');
const XeroMapper = require('../mapper');
const XeroApiClient = require('./xeroApiClient');
const XeroAuthService = require('./xeroAuthService');
const logger = require('../../../core/logger');

class XeroEntityService {
    /**
     * Shared implementation behind getContacts/getAccounts/getClasses/getLocations.
     */
    static async _getEntityList(url, mapFn, logLabel, userId) {
        const tokens = await XeroAuthService.getAllTokens(userId);
        const results = await Promise.all(tokens.map(async (token) => {
            try {
                const tenantId = token.companyId || token.tenant_id;
                const headers = await XeroApiClient._tenantHeaders(tenantId);
                const orgName = await XeroApiClient._resolveOrgName(token, tenantId, headers);

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
     * Fetch all organisation details from Xero across the calling user's connected tenants.
     */
    static async getOrganisation(userId) {
        const tokens = await XeroAuthService.getAllTokens(userId);
        const results = await Promise.all(tokens.map(async (token) => {
            try {
                const tenantId = token.companyId || token.tenant_id;
                const headers = await XeroApiClient._tenantHeaders(tenantId);
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
     * Fetch all contacts from Xero across the calling user's connected tenants.
     */
    static async getContacts(userId) {
        return XeroEntityService._getEntityList(CONSTANTS.XERO.CONTACTS_URL, XeroMapper.toContactList, 'contacts', userId);
    }

    /**
     * Fetch all accounts from Xero across the calling user's connected tenants.
     */
    static async getAccounts(userId) {
        return XeroEntityService._getEntityList(CONSTANTS.XERO.ACCOUNTS_URL, XeroMapper.toAccountList, 'accounts', userId);
    }

    /**
     * Fetch tracking categories for classes from Xero.
     */
    static async getClasses(userId) {
        return XeroEntityService._getEntityList(
            CONSTANTS.XERO.TRACKING_CATEGORIES_URL,
            data => XeroMapper.toTrackingList(data, "class"),
            'classes',
            userId
        );
    }

    /**
     * Fetch tracking categories for locations from Xero.
     */
    static async getLocations(userId) {
        return XeroEntityService._getEntityList(
            CONSTANTS.XERO.TRACKING_CATEGORIES_URL,
            data => XeroMapper.toTrackingList(data, "location"),
            'locations',
            userId
        );
    }
}

module.exports = XeroEntityService;
