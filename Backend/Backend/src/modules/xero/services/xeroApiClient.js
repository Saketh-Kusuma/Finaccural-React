'use strict';

const axios = require('axios');
const CONSTANTS = require('../../../core/constants');
const XeroMapper = require('../mapper');
const XeroTokenManager = require('../oauth/XeroTokenManager');

/** True if an axios error looks like an expired/revoked OAuth grant. */
function isAuthError(err) {
    if (!err) return false;
    if (err.response?.status === 401 || err.response?.status === 403) return true;
    const blob = JSON.stringify(err.response?.data || err.message || '').toLowerCase();
    return blob.includes('invalid_grant') || blob.includes('invalid_token') || blob.includes('unauthorized');
}

class XeroApiClient {
    static isAuthError = isAuthError;

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
     * then the tenant id). A failed lookup is non-fatal.
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
}

module.exports = XeroApiClient;
