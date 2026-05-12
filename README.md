# Outlook MCP Server

Ein MCP‑Server für **Microsoft 365** (Outlook, OneDrive, Power Automate) über die Microsoft Graph + Flow API. Fork von [`ryaker/outlook-mcp`](https://github.com/ryaker/outlook-mcp), erweitert um:

- **Streamable HTTP + SSE Transport** zusätzlich zu stdio — für Remote‑Use über Traefik / Claude.ai
- **MCP‑Auth‑Gateway** — `MCP_API_KEY` (statisches Bearer) für CLI/Claude‑Desktop, **Authelia‑OIDC‑Introspection** für Claude.ai cloud
- **OAuth‑Discovery** nach RFC 8414 + RFC 9728 — `/.well-known/oauth-authorization-server` & `/.well-known/oauth-protected-resource` öffentlich, damit Claude.ai cloud den Konnektor automatisch ausrollt
- **Microsoft‑OAuth‑Bootstrap** im selben Express‑App (`/auth`, `/auth/callback`), kein zweiter Server auf `localhost:3333` mehr nötig
- **Konfigurierbarer Token‑Store** (`TOKEN_STORE_PATH`) — Persistenz über Docker‑Volume
- **Dockerfile + Traefik‑Compose** — passt 1:1 in das gleiche Setup wie `smtp-mcp-server` / `paperless-mcp-server`

## Toolset

Komplett vom Upstream geerbt:

| Bereich | Tools |
|---|---|
| **Mail** | list/search/read/send/move/delete, draft, mark‑as‑read, list/get attachments |
| **Kalender** | list/create/cancel/accept/decline/delete events |
| **Aufgaben (To Do)** | list‑lists, list/get/create/update/delete tasks |
| **Folder** | list, create, move |
| **Rules** | list, create, edit‑sequence |
| **OneDrive** | list/search/download/upload (klein + chunked), share, delete, create‑folder |
| **Power Automate** | environments, flows (list/run/toggle), run‑history |
| **Auth** | `authenticate`, `check-auth-status`, `about` |

## Architektur

```
┌─ Browser ────────────────┐                ┌─ Claude.ai / Claude Desktop ─┐
│ /auth, /auth/callback    │                │ POST /mcp (Streamable HTTP)  │
│ (MS OAuth Bootstrap)     │                │ GET /sse + POST /messages    │
│ /.well-known/oauth-*     │                │                              │
└───────────┬──────────────┘                └────────────┬─────────────────┘
            │                                            │
            │       Traefik (TLS, DNS-01)                │
            ▼                                            ▼
┌──────────────────────────────────────────────────────────────────────────┐
│ Express (port 3000)                                                      │
│  ├─ public:         /, /auth, /auth/callback, /.well-known/oauth-*       │
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

### 1. Microsoft Entra App‑Registrierung

1. [Entra Admin Center](https://entra.microsoft.com) → **Identity → Applications → App registrations → New registration**
2. **Redirect URI** (Platform = **Web**): `https://<MCP_DOMAIN>/auth/callback`
3. **Certificates & secrets** → neuen Client‑Secret‑**VALUE** (nicht ID!) erzeugen, 24 Monate
4. **API permissions → Microsoft Graph → Delegated** — anhaken:
   ```
   offline_access  User.Read  Mail.Read  Mail.ReadWrite  Mail.Send
   Calendars.Read  Calendars.ReadWrite  Files.Read  Files.ReadWrite  Contacts.Read
   Tasks.Read  Tasks.ReadWrite
   ```
   → **Grant admin consent for <Tenant>**

### 2. Authelia‑Client (Claude.ai cloud OAuth)

In `configuration.yml` unter `identity_providers.oidc.clients`:

```yaml
- client_id: outlook-mcp
  client_name: Outlook MCP Server
  client_secret: '$pbkdf2-sha512$310000$<pbkdf2-hash deines OIDC_CLIENT_SECRET>'
  public: false
  authorization_policy: one_factor
  redirect_uris:
    - https://claude.ai/api/mcp/auth_callback
  scopes: [openid, profile, email, offline_access, address, phone, groups]
  grant_types: [authorization_code, refresh_token]
  response_types: [code]
  token_endpoint_auth_method: client_secret_post
  introspection_endpoint_auth_method: client_secret_basic
```

> **Authelia 4.39+:** Client‑Secrets müssen pbkdf2 oder argon2id sein —
> bcrypt (`$2b$…`/`$2y$…`) führt zu *„client secret did not match"*-Fehlern
> am Token‑Endpoint. Hash erzeugen:
> ```bash
> docker run --rm authelia/authelia:latest \
>   authelia crypto hash generate pbkdf2 --variant sha512 \
>   --password '<plain-OIDC_CLIENT_SECRET>'
> ```
> Den `Digest:`-Wert (z.B. `$pbkdf2-sha512$310000$…`) als `client_secret`
> in `configuration.yml` eintragen, der Klartext kommt in
> `OIDC_CLIENT_SECRET` im Stack. Anschließend Authelia neu starten.

### 3. Stack in Portainer

**Stacks → Add stack → Web editor:**

```yaml
services:
  outlook-mcp-server:
    image: ghcr.io/marco2901/outlook-mcp-server:latest
    container_name: outlook-mcp-server
    restart: unless-stopped
    environment:
      - MS_CLIENT_ID=${MS_CLIENT_ID}
      - MS_CLIENT_SECRET=${MS_CLIENT_SECRET}
      - MS_TENANT_ID=${MS_TENANT_ID:-common}
      - MS_AUTHORITY_HOST=${MS_AUTHORITY_HOST:-https://login.microsoftonline.com}
      - MS_SCOPES=${MS_SCOPES:-offline_access User.Read Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Calendars.ReadWrite Files.Read Files.ReadWrite Contacts.Read Tasks.Read Tasks.ReadWrite}
      - OAUTH_PUBLIC_BASE_URL=https://${MCP_DOMAIN}
      - TOKEN_STORE_PATH=/data/outlook-mcp-tokens.json
      - MCP_API_KEY=${MCP_API_KEY}
      # Authelia 4.39+: externe HTTPS-URL nutzen, sonst X-Forwarded-Proto-Reject.
      - OIDC_INTROSPECTION_URL=https://authelia.your-domain.com/api/oidc/introspection
      - OIDC_ISSUER_URL=${OIDC_ISSUER_URL}
      - OIDC_CLIENT_ID=${OIDC_CLIENT_ID}
      - OIDC_CLIENT_SECRET=${OIDC_CLIENT_SECRET}
    volumes:
      - outlook-mcp-tokens:/data
    expose:
      - "3000"
    labels:
      - traefik.enable=true
      - traefik.docker.network=traefik
      - traefik.http.routers.outlook-mcp.rule=Host(`${MCP_DOMAIN}`)
      - traefik.http.routers.outlook-mcp.entrypoints=websecure
      - traefik.http.services.outlook-mcp.loadbalancer.server.port=3000
      - traefik.http.routers.outlook-mcp.tls.certresolver=mydnschallenge
      - traefik.http.routers.outlook-mcp.tls=true
      - traefik.http.routers.outlook-mcp.middlewares=middlewares-rate-limit@file,middlewares-secure-headers@file
    networks:
      - traefik

volumes:
  outlook-mcp-tokens:

networks:
  traefik:
    external: true
```

**Environment variables** (im Portainer‑Stack):

| Name | Beispiel‑Wert |
|---|---|
| `MS_CLIENT_ID` | aus Entra → Übersicht → Anwendungs‑(Client‑)ID |
| `MS_CLIENT_SECRET` | Secret‑Value aus Entra |
| `MS_TENANT_ID` | aus Entra → Übersicht → Verzeichnis‑(Mandanten‑)ID |
| `MCP_DOMAIN` | `outlook-mcp.example.com` |
| `MCP_API_KEY` | `openssl rand -hex 32` |
| `OIDC_ISSUER_URL` | `https://authelia.example.com` |
| `OIDC_CLIENT_ID` | `outlook-mcp` |
| `OIDC_CLIENT_SECRET` | Plain‑Wert (ohne Hash) |

→ **Deploy the stack**.

### 4. Microsoft OAuth‑Bootstrap (einmalig)

Browser auf `https://<MCP_DOMAIN>/auth` öffnen → Microsoft‑Login mit deinem M365‑Account → "Authentication successful". Tokens liegen im Volume `outlook-mcp-tokens`, Refresh läuft automatisch via `offline_access`.

### 5. Claude.ai cloud Connector

**Settings → Connectors → Add custom connector:**

| Feld | Wert |
|---|---|
| URL | `https://<MCP_DOMAIN>/mcp` |
| Client ID | `outlook-mcp` |
| Client Secret | dein Plain‑`OIDC_CLIENT_SECRET` (nicht der pbkdf2‑Hash) |

Beim ersten Connect läuft der Authelia‑Login durch, danach sind die Tools dauerhaft verfügbar.

## Lokale Nutzung (stdio, ohne Docker)

```bash
git clone https://github.com/marco2901/outlook-mcp-server.git
cd outlook-mcp-server
npm install
cp .env.example .env  # füllen
node index.js --http --port 3333  # einmalig für OAuth-Bootstrap
# Browser: http://localhost:3333/auth → Microsoft-Login

# danach für Claude Desktop:
node index.js
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

## Discovery‑Endpoints (ungeschützt, public)

| Pfad | Inhalt |
|---|---|
| `/` | Landing‑Page mit Link auf `/auth` |
| `/.well-known/oauth-authorization-server` | RFC 8414 — Authelia‑Endpoints |
| `/.well-known/oauth-protected-resource` | RFC 9728 — verweist auf Authelia |
| `/.well-known/oauth-protected-resource/mcp` | wie oben, resource = `/mcp` |
| `/auth` | Redirect zu Microsoft‑Login |
| `/auth/callback` | OAuth Callback, persistiert Tokens |

`/mcp`, `/sse`, `/messages` sind durch die `authMiddleware` geschützt.

## Upstream‑Updates ziehen

```bash
git remote add upstream https://github.com/ryaker/outlook-mcp.git
git fetch upstream
git merge upstream/main
```

Konflikte konzentrieren sich auf `index.js`, `config.js` und `auth/token-storage.js`. Die Tool‑Module (calendar/, email/, …) bleiben unangetastet.

## Lizenz

MIT — siehe [LICENSE](LICENSE). Doppel‑Copyright Upstream (Richard Yaker) + Fork (Marco Biegel).
