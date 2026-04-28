# Builder stage
FROM node:24-slim AS builder
WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm install --omit=dev
COPY . .

# Production stage
FROM node:24-slim AS production
WORKDIR /app
ENV NODE_ENV=production
ENV TOKEN_STORE_PATH=/data/outlook-mcp-tokens.json

COPY --from=builder /app/node_modules ./node_modules
COPY --from=builder /app/package.json ./package.json
COPY --from=builder /app/index.js ./index.js
COPY --from=builder /app/config.js ./config.js
COPY --from=builder /app/outlook-auth-server.js ./outlook-auth-server.js
COPY --from=builder /app/auth ./auth
COPY --from=builder /app/calendar ./calendar
COPY --from=builder /app/email ./email
COPY --from=builder /app/folder ./folder
COPY --from=builder /app/onedrive ./onedrive
COPY --from=builder /app/power-automate ./power-automate
COPY --from=builder /app/rules ./rules
COPY --from=builder /app/utils ./utils

RUN mkdir -p /data && chown node:node /data
VOLUME ["/data"]
USER node

EXPOSE 3000
ENTRYPOINT ["node", "index.js", "--http", "--port", "3000"]
