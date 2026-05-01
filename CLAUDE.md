# CLAUDE.md

Briefing for Claude Code sessions working on this fork.

## What this repo is

Fork of [`ryaker/outlook-mcp`](https://github.com/ryaker/outlook-mcp) — an MCP
server for Microsoft 365 (Outlook, OneDrive, Power Automate) via the Microsoft
Graph + Flow API. The fork adds:

- Streamable HTTP + SSE transport (next to upstream's stdio) for remote MCP use
- MCP auth gateway: static `MCP_API_KEY` Bearer **or** Authelia OIDC introspection
- OAuth bootstrap (`/auth`, `/auth/callback`) and discovery endpoints
  (`/.well-known/oauth-authorization-server`, `/.well-known/oauth-protected-resource`)
  mounted on the same Express app — no separate auth server on `localhost:3333`
- Configurable token store (`TOKEN_STORE_PATH`) for Docker volume persistence
- Dockerfile + Traefik-fronted docker-compose matching the smtp-mcp-server /
  paperless-mcp-server stack pattern

## Key files

- `index.js` — entry point (stdio + HTTP transport, auth gate, OAuth routes,
  RFC 8414 + RFC 9728 discovery)
- `config.js` — env-driven Microsoft Graph / OAuth config
- `auth/token-storage.js` — token persistence + refresh (env-driven path)
- `auth/`, `calendar/`, `email/`, `folder/`, `rules/`, `onedrive/`,
  `power-automate/`, `utils/` — tool modules (untouched from upstream)
- `Dockerfile`, `docker-compose.yml`, `.env.example` — deployment
- `.github/workflows/build-push.yml` — multi-arch GHCR image build on `main`

## Commands

```bash
npm install
node index.js                          # stdio
node index.js --http --port 3000       # HTTP + OAuth (dev/local)
USE_TEST_MODE=true npm start           # mock data, no real Graph calls
npm test                               # jest
npm run inspect                        # MCP Inspector against stdio
```

## Pulling upstream updates

```bash
git remote add upstream https://github.com/ryaker/outlook-mcp.git
git fetch upstream
git merge upstream/main
```

Conflicts concentrate in `index.js` (HTTP/auth layer), `config.js` (env-driven
paths) and `auth/token-storage.js` (`TOKEN_STORE_PATH`). Tool modules under
`calendar/`, `email/` etc. should merge cleanly.

## Conventions

- TypeScript-style JSDoc on new helpers, but the server itself stays plain JS
  (CommonJS) to keep the upstream merge surface small.
- Anything secret-looking (real domains, API keys, tenant IDs, client secrets)
  goes through `.env` only — never into committed files. `.env.example` uses
  `example.com` style placeholders.
- The MCP gateway auth is layered: `MCP_API_KEY` for CLI/Claude-Desktop access,
  Authelia OIDC introspection for Claude.ai cloud connectors. Both can coexist.
