#!/usr/bin/env node
/**
 * Outlook MCP Server entry point.
 *
 * Supports two transports:
 *   - stdio (default) for local Claude Desktop / Claude Code use
 *   - Streamable HTTP + SSE on Express when started with `--http`,
 *     gated by an authMiddleware (static MCP_API_KEY or Authelia OIDC
 *     introspection) so it can be exposed safely behind Traefik.
 *
 * In HTTP mode the same Express app also serves the Microsoft OAuth
 * bootstrap (`/auth` + `/auth/callback`), so you don't need the legacy
 * standalone `outlook-auth-server.js` on port 3333. The redirect URI
 * is whatever `OAUTH_PUBLIC_BASE_URL` resolves to (e.g.
 * `https://outlook-mcp.example.com/auth/callback`).
 */
const crypto = require('crypto');
const querystring = require('querystring');
const { parseArgs } = require('node:util');

const express = require('express');

const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const { StreamableHTTPServerTransport } = require('@modelcontextprotocol/sdk/server/streamableHttp.js');
const { SSEServerTransport } = require('@modelcontextprotocol/sdk/server/sse.js');

const config = require('./config');

// Tool modules (unchanged from upstream)
const { authTools } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const { folderTools } = require('./folder');
const { rulesTools } = require('./rules');
const { onedriveTools } = require('./onedrive');
const { powerAutomateTools } = require('./power-automate');

const TokenStorage = require('./auth/token-storage');

// ---------------------------------------------------------------------------
// CLI args
// ---------------------------------------------------------------------------
const {
  values: { http: useHttp, port }
} = parseArgs({
  options: {
    http: { type: 'boolean', default: false },
    port: { type: 'string' }
  },
  allowPositionals: true
});

const resolvedPort = port ? parseInt(port, 10) : 3000;

// ---------------------------------------------------------------------------
// Auth (MCP gateway): static key OR Authelia OIDC introspection
// ---------------------------------------------------------------------------
const mcpApiKey = process.env.MCP_API_KEY || '';
const oidcIntrospectionUrl = process.env.OIDC_INTROSPECTION_URL || '';
const oidcClientId = process.env.OIDC_CLIENT_ID || '';
const oidcClientSecret = process.env.OIDC_CLIENT_SECRET || '';

async function isAuthorized(req) {
  if (!mcpApiKey && !oidcIntrospectionUrl) return true; // no gating configured

  const auth = req.headers.authorization || '';

  if (mcpApiKey && auth === `Bearer ${mcpApiKey}`) {
    return true;
  }

  if (auth.startsWith('Bearer ') && oidcIntrospectionUrl && oidcClientId && oidcClientSecret) {
    const jwtToken = auth.slice(7);
    try {
      const credentials = Buffer.from(`${oidcClientId}:${oidcClientSecret}`).toString('base64');
      const resp = await fetch(oidcIntrospectionUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          Authorization: `Basic ${credentials}`
        },
        body: new URLSearchParams({ token: jwtToken }),
        signal: AbortSignal.timeout(5000)
      });
      const data = await resp.json();
      return data && data.active === true;
    } catch (e) {
      console.error('Introspection failed:', e.message);
    }
  }

  return false;
}

function authMiddleware(req, res, next) {
  isAuthorized(req).then((ok) => {
    if (ok) next();
    else res.status(401).send('Unauthorized');
  });
}

// ---------------------------------------------------------------------------
// MCP server (low-level Server, matches upstream tool registration)
// ---------------------------------------------------------------------------
const TOOLS = [
  ...authTools,
  ...calendarTools,
  ...emailTools,
  ...folderTools,
  ...rulesTools,
  ...onedriveTools,
  ...powerAutomateTools
];

function createServer() {
  const server = new Server(
    { name: config.SERVER_NAME, version: config.SERVER_VERSION },
    { capabilities: { tools: {} } }
  );

  server.fallbackRequestHandler = async (request) => {
    const { method, params, id } = request;

    if (method === 'initialize') {
      return {
        protocolVersion: '2025-11-25',
        capabilities: { tools: {} },
        serverInfo: { name: config.SERVER_NAME, version: config.SERVER_VERSION }
      };
    }

    if (method === 'tools/list') {
      return {
        tools: TOOLS.map(t => ({
          name: t.name,
          description: t.description,
          inputSchema: t.inputSchema
        }))
      };
    }

    if (method === 'resources/list') return { resources: [] };
    if (method === 'prompts/list') return { prompts: [] };

    if (method === 'tools/call') {
      try {
        const { name, arguments: args = {} } = params || {};
        const tool = TOOLS.find(t => t.name === name);
        if (tool && tool.handler) return await tool.handler(args);
        return { error: { code: -32601, message: `Tool not found: ${name}` } };
      } catch (error) {
        console.error('Error in tools/call:', error);
        return { error: { code: -32603, message: `Error processing tool call: ${error.message}` } };
      }
    }

    return { error: { code: -32601, message: `Method not found: ${method}` } };
  };

  return server;
}

// ---------------------------------------------------------------------------
// OAuth bootstrap routes (Microsoft Graph)
// ---------------------------------------------------------------------------
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function mountOAuthRoutes(app) {
  const tokenStorage = new TokenStorage();
  const pendingStates = new Map();
  const TEN_MINUTES = 10 * 60 * 1000;

  setInterval(() => {
    const now = Date.now();
    for (const [key, ts] of pendingStates.entries()) {
      if (now - ts > TEN_MINUTES) pendingStates.delete(key);
    }
  }, 5 * 60 * 1000).unref();

  app.get('/auth', (req, res) => {
    const { clientId, authorityHost, tenantId, redirectUri, scopes } = config.AUTH_CONFIG;
    if (!clientId) {
      res.status(500).type('html').send(`
        <h1>Configuration error</h1>
        <p>MS_CLIENT_ID is not set on the server.</p>
      `);
      return;
    }
    const state = crypto.randomBytes(32).toString('hex');
    pendingStates.set(state, Date.now());

    const authUrl = `${authorityHost}/${tenantId}/oauth2/v2.0/authorize?` + querystring.stringify({
      client_id: clientId,
      response_type: 'code',
      redirect_uri: redirectUri,
      scope: scopes.join(' '),
      response_mode: 'query',
      state
    });
    res.redirect(302, authUrl);
  });

  app.get('/auth/callback', async (req, res) => {
    const q = req.query;

    if (!q.state || !pendingStates.has(q.state)) {
      res.status(403).type('html').send(`
        <h1>Authentication error</h1>
        <p>Invalid or expired OAuth state. Please try again.</p>
      `);
      return;
    }
    pendingStates.delete(q.state);

    if (q.error) {
      res.status(400).type('html').send(`
        <h1>Authentication error</h1>
        <p><strong>${escapeHtml(q.error)}</strong>: ${escapeHtml(q.error_description || '')}</p>
      `);
      return;
    }

    if (!q.code) {
      res.status(400).type('html').send(`
        <h1>Missing authorization code</h1>
      `);
      return;
    }

    try {
      await tokenStorage.exchangeCodeForTokens(q.code);
      res.status(200).type('html').send(`
        <html><head><title>Authentication successful</title>
        <style>body{font-family:system-ui;max-width:600px;margin:40px auto;padding:0 20px}h1{color:#5cb85c}</style>
        </head><body>
        <h1>Authentication successful</h1>
        <p>Tokens stored. You can close this tab and return to Claude.</p>
        </body></html>
      `);
    } catch (error) {
      console.error('Token exchange failed:', error);
      res.status(500).type('html').send(`
        <h1>Token exchange error</h1>
        <p>${escapeHtml(error.message)}</p>
      `);
    }
  });

  app.get('/', (_req, res) => {
    res.type('html').send(`
      <html><head><title>Outlook MCP Server</title>
      <style>body{font-family:system-ui;max-width:700px;margin:40px auto;padding:0 20px}code{background:#f4f4f4;padding:2px 6px;border-radius:3px}</style>
      </head><body>
      <h1>Outlook MCP Server</h1>
      <p>This is the OAuth bootstrap landing page. To authenticate with Microsoft Graph,
      open <a href="/auth">/auth</a> in a browser and complete the sign-in flow.</p>
      <p>The MCP endpoint is <code>POST /mcp</code> (Streamable HTTP) or
      <code>GET /sse</code> + <code>POST /messages</code> (SSE), both gated by the
      configured authentication.</p>
      </body></html>
    `);
  });
}

// ---------------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------------
async function main() {
  console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} MCP SERVER`);
  console.error(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);

  const server = createServer();

  if (useHttp) {
    const app = express();
    app.use(express.json());

    // Public OAuth bootstrap (no MCP auth gate — Microsoft owns the flow,
    // CSRF state protects the callback).
    mountOAuthRoutes(app);

    // MCP routes — gated by API key / OIDC introspection.
    const mcpRouter = express.Router();
    mcpRouter.use(authMiddleware);

    const sseTransports = {};

    mcpRouter.post('/mcp', async (req, res) => {
      try {
        const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
        res.on('close', () => transport.close());
        await server.connect(transport);
        await transport.handleRequest(req, res, req.body);
      } catch (error) {
        console.error('Error handling MCP request:', error);
        if (!res.headersSent) {
          res.status(500).json({
            jsonrpc: '2.0',
            error: { code: -32603, message: 'Internal server error' },
            id: null
          });
        }
      }
    });

    mcpRouter.get('/mcp', (_req, res) => {
      res.writeHead(405).end(JSON.stringify({
        jsonrpc: '2.0',
        error: { code: -32000, message: 'Method not allowed.' },
        id: null
      }));
    });

    mcpRouter.delete('/mcp', (_req, res) => {
      res.writeHead(405).end(JSON.stringify({
        jsonrpc: '2.0',
        error: { code: -32000, message: 'Method not allowed.' },
        id: null
      }));
    });

    mcpRouter.get('/sse', async (_req, res) => {
      try {
        const transport = new SSEServerTransport('/messages', res);
        sseTransports[transport.sessionId] = transport;
        res.on('close', () => {
          delete sseTransports[transport.sessionId];
          transport.close();
        });
        await server.connect(transport);
      } catch (error) {
        console.error('Error handling SSE request:', error);
        if (!res.headersSent) {
          res.status(500).json({
            jsonrpc: '2.0',
            error: { code: -32603, message: 'Internal server error' },
            id: null
          });
        }
      }
    });

    mcpRouter.post('/messages', async (req, res) => {
      const sessionId = req.query.sessionId;
      const transport = sseTransports[sessionId];
      if (transport) {
        await transport.handlePostMessage(req, res, req.body);
      } else {
        res.status(400).send('No transport found for sessionId');
      }
    });

    app.use(mcpRouter);

    app.listen(resolvedPort, () => {
      console.error(`Outlook MCP Server listening on port ${resolvedPort}`);
      console.error(`OAuth bootstrap: ${config.AUTH_CONFIG.authServerUrl}/auth`);
      console.error(`Token store:    ${config.AUTH_CONFIG.tokenStorePath}`);
    });
  } else {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error(`${config.SERVER_NAME} connected and listening (stdio)`);
  }
}

process.on('SIGTERM', () => {
  console.error('SIGTERM received, exiting');
  process.exit(0);
});

main().catch((error) => {
  console.error(`Connection error: ${error.message}`);
  process.exit(1);
});
