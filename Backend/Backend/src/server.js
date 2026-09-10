const dns = require('dns');

if (typeof dns.setDefaultResultOrder === 'function') {
    dns.setDefaultResultOrder('ipv4first');
}

const fs = require('fs');
const http = require('http');
const https = require('https');
const os = require('os');
const path = require('path');

const app = require('./app');
const { sequelize } = require('./core/database');
const config = require('./core/config');
const logger = require('./core/logger');

const shouldAlter = process.env.DB_SYNC_ALTER === 'true';

if (shouldAlter) {
    logger.info(
        'DB_SYNC_ALTER=true — syncing with { alter: true }. Turn this off once your schema is stable.'
    );
}

// ---------------------------------------------------------------------------
// Dev-HTTPS helper
// ---------------------------------------------------------------------------
// On macOS, Safari/WKWebView enforces strict mixed-content rules: an HTTPS
// taskpane (webpack dev server runs on https://localhost:3001) cannot fetch
// from an HTTP backend.  In development we reuse the TLS certs that
// office-addin-dev-certs already generated and installed for the webpack
// dev server.  The certs live at ~/.office-addin-dev-certs/ on both macOS
// and Windows.  If the files are absent we fall through to plain HTTP so
// production / CI environments are completely unaffected.
// ---------------------------------------------------------------------------
function tryLoadDevCerts() {
    try {
        const certDir = path.join(os.homedir(), '.office-addin-dev-certs');
        const ca   = fs.readFileSync(path.join(certDir, 'ca.crt'));
        const cert = fs.readFileSync(path.join(certDir, 'localhost.crt'));
        const key  = fs.readFileSync(path.join(certDir, 'localhost.key'));
        return { ca, cert, key };
    } catch (_) {
        return null;
    }
}

function startServer() {
    const isProduction = process.env.NODE_ENV === 'production';
    const port = config.PORT;

    if (!isProduction) {
        const certs = tryLoadDevCerts();
        if (certs) {
            https.createServer(certs, app).listen(port, '0.0.0.0', () => {
                logger.info(`Node.js backend running on https://localhost:${port} (dev HTTPS via office-addin-dev-certs)`);
            });
            return;
        }
        logger.warn(
            'office-addin-dev-certs not found at ~/.office-addin-dev-certs — ' +
            'starting HTTP server. Run `npx office-addin-dev-certs install` in the ' +
            'Frontend-React directory to enable HTTPS on the backend (required on macOS).'
        );
    }

    http.createServer(app).listen(port, '0.0.0.0', () => {
        logger.info(`Node.js backend running on http://localhost:${port}`);
    });
}

sequelize
    .sync({ alter: shouldAlter })
    .then(() => {
        logger.info('Database synchronized.');
        startServer();
    })
    .catch((err) => {
        logger.error('Unable to connect to the database:', err);
    });