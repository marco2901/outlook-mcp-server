/**
 * Upload a Buffer to OneDrive at a given path.
 *
 * Picks the right strategy based on size:
 *   - <= 4 MB: simple PUT to /content endpoint (one request, raw bytes)
 *   - >  4 MB: upload session + chunked PUTs to the returned upload URL
 *
 * Both paths bypass the project's `callGraphAPI()` helper for the actual
 * data transfer, because `callGraphAPI()` JSON.stringify's the body — fine
 * for OData, fatal for binary content.
 */
const https = require('https');
const config = require('../config');
const { callGraphAPI } = require('./graph-api');

const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024; // 4 MB — Microsoft's threshold
const CHUNK_SIZE = 320 * 1024 * 10;          // 3.2 MB — must be multiple of 320 KB

/**
 * Public entry point.
 * @param {string} accessToken
 * @param {string} dstPath           - OneDrive path, e.g. "/Documents/foo.pdf"
 * @param {Buffer} buffer            - raw file bytes
 * @param {object} [opts]
 * @param {string} [opts.conflictBehavior='rename']  - rename|replace|fail
 * @returns {Promise<object>}        - DriveItem JSON from Graph (id, name, size, webUrl, …)
 */
async function uploadBufferToOneDrive(accessToken, dstPath, buffer, opts = {}) {
  if (!Buffer.isBuffer(buffer)) {
    throw new Error('uploadBufferToOneDrive: buffer must be a Buffer');
  }
  const conflictBehavior = opts.conflictBehavior || 'rename';
  const normalizedPath = String(dstPath).replace(/^\/+|\/+$/g, '');

  if (buffer.length <= SIMPLE_UPLOAD_LIMIT) {
    return simpleUpload(accessToken, normalizedPath, buffer, conflictBehavior);
  }
  return sessionUpload(accessToken, normalizedPath, buffer, conflictBehavior);
}

/**
 * Single-shot PUT to /me/drive/root:/{path}:/content with raw binary body.
 */
function simpleUpload(accessToken, normalizedPath, buffer, conflictBehavior) {
  return new Promise((resolve, reject) => {
    const url =
      `${config.GRAPH_API_ENDPOINT}me/drive/root:/${encodeURIPath(normalizedPath)}:/content` +
      `?@microsoft.graph.conflictBehavior=${encodeURIComponent(conflictBehavior)}`;

    const req = https.request(
      url,
      {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/octet-stream',
          'Content-Length': buffer.length
        }
      },
      (res) => collectJson(res, resolve, reject)
    );
    req.on('error', reject);
    req.write(buffer);
    req.end();
  });
}

/**
 * Chunked upload via upload session. The session is created with
 * callGraphAPI() (JSON body OK), then each chunk is a raw PUT to the
 * returned uploadUrl.
 */
async function sessionUpload(accessToken, normalizedPath, buffer, conflictBehavior) {
  const sessionEndpoint = `me/drive/root:/${normalizedPath}:/createUploadSession`;
  const session = await callGraphAPI(accessToken, 'POST', sessionEndpoint, {
    item: { '@microsoft.graph.conflictBehavior': conflictBehavior }
  });
  if (!session || !session.uploadUrl) {
    throw new Error('Failed to create OneDrive upload session.');
  }

  const totalSize = buffer.length;
  let offset = 0;
  let lastResponse = null;

  while (offset < totalSize) {
    const chunkEnd = Math.min(offset + CHUNK_SIZE, totalSize);
    const chunk = buffer.slice(offset, chunkEnd);
    lastResponse = await putChunk(session.uploadUrl, chunk, offset, chunkEnd - 1, totalSize);
    offset = chunkEnd;
  }

  if (!lastResponse || !lastResponse.id) {
    throw new Error('Chunked upload finished but server returned no DriveItem.');
  }
  return lastResponse;
}

function putChunk(uploadUrl, chunk, start, end, totalSize) {
  return new Promise((resolve, reject) => {
    const req = https.request(
      uploadUrl,
      {
        method: 'PUT',
        headers: {
          'Content-Length': chunk.length,
          'Content-Range': `bytes ${start}-${end}/${totalSize}`
        }
      },
      (res) => collectJson(res, resolve, reject)
    );
    req.on('error', reject);
    req.write(chunk);
    req.end();
  });
}

/**
 * Collect a JSON response body or reject on non-2xx.
 */
function collectJson(res, resolve, reject) {
  const chunks = [];
  res.on('data', (c) => chunks.push(c));
  res.on('end', () => {
    const body = Buffer.concat(chunks).toString('utf8');
    if (res.statusCode >= 200 && res.statusCode < 300) {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch (e) {
        reject(new Error(`Bad JSON from OneDrive (status ${res.statusCode}): ${body.slice(0, 200)}`));
      }
    } else {
      reject(new Error(`OneDrive upload failed (status ${res.statusCode}): ${body.slice(0, 500)}`));
    }
  });
  res.on('error', reject);
}

/**
 * Encode each path segment but keep the slash separators.
 */
function encodeURIPath(path) {
  return path.split('/').map(encodeURIComponent).join('/');
}

module.exports = {
  uploadBufferToOneDrive,
  SIMPLE_UPLOAD_LIMIT,
  CHUNK_SIZE
};
