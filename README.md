# Outlook MCP Server

Ein MCP-Server für **Microsoft 365** (Outlook, OneDrive, Power Automate) über die Microsoft Graph + Flow API. Fork von [`ryaker/outlook-mcp`](https://github.com/ryaker/outlook-mcp), erweitert um:

- **Streamable HTTP + SSE Transport** (zusätzlich zu stdio) — für Remote-Use über Traefik / Claude.ai
- **MCP-Auth-Gateway** — `MCP_API_KEY` (statisches Bearer) ODER Authelia-OIDC-Introspection
- **OAuth-Callback im selben Express-App** — kein separater Server auf `localhost:3333` mehr nötig, Redirect läuft auf `https://<MCP_DOMAIN>/auth/callback`
- **Konfigurierbarer Token-Store** (`TOKEN_STORE_PATH`) — Persistenz über Docker-Volume
- **Dockerfile + docker-compose mit Traefik-Labels** — passt in das bestehende `smtp-mcp-server` / `paperless-mcp-server` Setup

## Toolset

Geerbt vom Upstream — komplettes M365 Toolset:

- **Outlook** — Mail (list/search/read/send/move/delete), Kalender (list/create/accept/decline/delete events), Folder-Operationen, Inbox-Rules
- **OneDrive** — Datei-/Ordner-Listing, Suche, Download, Upload (klein + chunked), Teilen, Löschen
- **Power Automate** — Environments, Flows (list/run/toggle), Run-History
- **Auth** — `authenticate`, `check-auth-status`, `about`

Vollständige Tool-Liste via `tools/list` oder beim Upstream-Repo nachschlagen.

## Architektur

```
┌─ Browser ────────────────┐                ┌─ Claude.ai / Claude Desktop ─┐
│ /auth, /auth/callback    │                │ POST /mcp (Streamable HTTP)  │
│ (OAuth Bootstrap, public)│                │ GET /sse + POST /messages    │
└───────────┬──────────────┘                └────────────┬─────────────────┘
            │                                            │
            │       Traefik (TLS)                        │
            ▼                                            ▼
┌──────────────────────────────────────────────────────────────────────────┐
│ Express (port 3000)                                                      │
│  ├─ public:        /, /auth, /auth/callback                              │
│  └─ authMiddleware: /mcp, /sse, /messages                                │
│      (MCP_API_KEY === Bearer)  OR  (Authelia OIDC introspection)         │
│                                                                          │
│ MCP Server  ─►  Tools (auth, calendar, email, folder, rules,             │
│                        onedrive, power-automate)                         │
│                  └─► Microsoft Graph API + Flow API                      │
│                                                                          │
│ TokenStorage  ◄─►  /data/outlook-mcp-tokens.json (Volume)                │
└──────────────────────────────────────────────────────────────────────────┘
```

## Setup

### 1. Azure App Registration

1. [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Redirect URI: **Web** → `https://<dein-host>/auth/callback`
3. **Certificates & secrets** → neuen Client-Secret-VALUE (nicht ID!) erzeugen
4. **API permissions** → Microsoft Graph → Delegated:
   - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
   - `Calendars.Read`, `Calendars.ReadWrite`
   - `Files.Read`, `Files.ReadWrite`
   - `Contacts.Read`
   - `User.Read`, `offline_access`
5. (Optional Power Automate) → Flow service: `https://service.flow.microsoft.com//.default`

### 2. Konfiguration

`cp .env.example .env` und befüllen:

```env
MS_CLIENT_ID=...
MS_CLIENT_SECRET=...
MS_TENANT_ID=...                 # tenant GUID oder "common"
OAUTH_PUBLIC_BASE_URL=https://outlook-mcp.example.com
MCP_DOMAIN=outlook-mcp.example.com
MCP_API_KEY=...                  # für CLI/Tooling-Zugriff
OIDC_CLIENT_ID=outlook-mcp       # für Claude.ai OAuth via Authelia
OIDC_CLIENT_SECRET=...
```

### 3. Deployment

```bash
docker compose up -d
```

Image kommt aus `ghcr.io/marco2901/outlook-mcp-server:latest` (build via GitHub Action). Lokal bauen: `docker build -t outlook-mcp-server .`.

### 4. OAuth-Bootstrap (einmalig)

Browser auf `https://<MCP_DOMAIN>/auth` öffnen → Microsoft-Login durchlaufen → Tokens werden ins Volume nach `/data/outlook-mcp-tokens.json` geschrieben (mode 0600). Refresh läuft automatisch.

Anschließend in Claude:
- **Claude Desktop / Code (Remote MCP)**: URL `https://<MCP_DOMAIN>/mcp`, Bearer = `MCP_API_KEY`
- **Claude.ai (Cloud)**: über Authelia-OIDC-Login authentisieren

## Lokal (stdio)

Für Claude Desktop direkt ohne Docker:

```bash
npm install
node index.js                       # stdio
node index.js --http --port 3000    # HTTP mit OAuth-Callback auf :3000
```

`claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/pfad/zu/outlook-mcp-server/index.js"],
      "env": {
        "MS_CLIENT_ID": "...",
        "MS_CLIENT_SECRET": "...",
        "MS_TENANT_ID": "..."
      }
    }
  }
}
```

Bei lokalem stdio-Setup ohne HTTP-Listener musst du den OAuth-Bootstrap einmalig per `npm run auth-server` (legacy `outlook-auth-server.js` auf Port 3333) durchlaufen — oder genauso `node index.js --http --port 3333` einmal kurz starten und im Browser `/auth` aufrufen.

## Upstream-Updates ziehen

```bash
git remote add upstream https://github.com/ryaker/outlook-mcp.git
git fetch upstream
git merge upstream/main
```

Konflikte primär in `index.js` (HTTP-Layer + Auth-Middleware), `config.js` (env-Pfade), `auth/token-storage.js` (TOKEN_STORE_PATH). Tool-Module bleiben unangetastet.

## Lizenz

MIT — wie Upstream.
