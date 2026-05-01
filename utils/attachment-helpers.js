/**
 * Helpers for working with email attachments via Microsoft Graph.
 *
 * Microsoft Graph accepts inline file attachments up to ~3 MB per item when
 * sending via /me/sendMail or attaching to a draft via the create-message
 * payload. Anything larger needs a chunked upload session, which we don't
 * implement yet — those attachments are rejected with a clear error pointing
 * at the OneDrive share-link workaround.
 */
const https = require('https');
const { callGraphAPI } = require('./graph-api');

// Microsoft Graph hard limit for inline file attachments.
const MAX_INLINE_ATTACHMENT_BYTES = 3 * 1024 * 1024;

/**
 * Fetches a file from a Graph download URL and returns its raw bytes.
 * Download URLs are pre-authenticated, so no Bearer header is needed.
 * @param {string} url
 * @returns {Promise<Buffer>}
 */
function fetchBinary(url) {
  return new Promise((resolve, reject) => {
    https.get(url, (res) => {
      if (res.statusCode === 302 || res.statusCode === 301) {
        return fetchBinary(res.headers.location).then(resolve, reject);
      }
      if (res.statusCode < 200 || res.statusCode >= 300) {
        return reject(new Error(`Download failed with status ${res.statusCode}`));
      }
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
  });
}

/**
 * Loads a OneDrive item by path or itemId and returns its metadata together
 * with the binary content as a base64 string.
 * @param {string} accessToken
 * @param {object} ref - { path?: string, itemId?: string }
 * @returns {Promise<{name: string, contentType: string, contentBytes: string, size: number}>}
 */
async function loadOneDriveAsAttachment(accessToken, { path, itemId }) {
  if (!path && !itemId) {
    throw new Error('Either path or itemId is required for OneDrive attachment.');
  }

  const endpoint = itemId
    ? `me/drive/items/${itemId}`
    : `me/drive/root:/${path.replace(/^\/+|\/+$/g, '')}`;

  const meta = await callGraphAPI(accessToken, 'GET', endpoint, null, {
    $select: 'id,name,size,file,@microsoft.graph.downloadUrl'
  });

  if (!meta || !meta['@microsoft.graph.downloadUrl']) {
    throw new Error(`Could not get download URL for OneDrive item${path ? ` "${path}"` : ` ${itemId}`}.`);
  }
  if (meta.folder) {
    throw new Error(`"${meta.name}" is a folder, not a file.`);
  }
  if (meta.size > MAX_INLINE_ATTACHMENT_BYTES) {
    throw new Error(
      `OneDrive item "${meta.name}" is ${formatSize(meta.size)} which exceeds the ` +
      `${formatSize(MAX_INLINE_ATTACHMENT_BYTES)} inline-attachment limit. ` +
      `Use onedrive-share to send a link instead, or split the file.`
    );
  }

  const buffer = await fetchBinary(meta['@microsoft.graph.downloadUrl']);
  return {
    name: meta.name,
    contentType: (meta.file && meta.file.mimeType) || 'application/octet-stream',
    contentBytes: buffer.toString('base64'),
    size: meta.size
  };
}

/**
 * Normalises a single user-supplied attachment spec to the FileAttachment
 * shape Microsoft Graph wants on /sendMail or POST /me/messages.
 *
 * Accepted input shapes (one per array element):
 *   { name, contentType?, contentBytes }      // inline base64
 *   { onedrivePath, name? }                   // load from OneDrive path
 *   { onedriveItemId, name? }                 // load by OneDrive item ID
 *
 * @param {string} accessToken
 * @param {object} att
 * @returns {Promise<object>}
 */
async function normalizeAttachment(accessToken, att) {
  if (!att || typeof att !== 'object') {
    throw new Error('Invalid attachment entry — expected object.');
  }

  let resolved;

  if (att.contentBytes) {
    // base64-inlined attachment, validate size
    const sizeBytes = Buffer.byteLength(att.contentBytes, 'base64');
    if (sizeBytes > MAX_INLINE_ATTACHMENT_BYTES) {
      throw new Error(
        `Attachment "${att.name || '(unnamed)'}" decodes to ${formatSize(sizeBytes)} ` +
        `which exceeds the ${formatSize(MAX_INLINE_ATTACHMENT_BYTES)} limit. ` +
        `Use a OneDrive share link instead.`
      );
    }
    resolved = {
      name: att.name || 'attachment',
      contentType: att.contentType || 'application/octet-stream',
      contentBytes: att.contentBytes
    };
  } else if (att.onedrivePath || att.onedriveItemId) {
    const loaded = await loadOneDriveAsAttachment(accessToken, {
      path: att.onedrivePath,
      itemId: att.onedriveItemId
    });
    resolved = {
      name: att.name || loaded.name,
      contentType: att.contentType || loaded.contentType,
      contentBytes: loaded.contentBytes
    };
  } else {
    throw new Error(
      'Each attachment must provide either contentBytes (base64) ' +
      'or onedrivePath / onedriveItemId.'
    );
  }

  return {
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: resolved.name,
    contentType: resolved.contentType,
    contentBytes: resolved.contentBytes
  };
}

/**
 * Resolves an array of attachment specs into Graph-ready FileAttachments.
 * Resolves OneDrive references in parallel.
 */
async function normalizeAttachments(accessToken, attachments) {
  if (!Array.isArray(attachments) || attachments.length === 0) return [];
  return Promise.all(attachments.map((a) => normalizeAttachment(accessToken, a)));
}

function formatSize(bytes) {
  if (!bytes) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

module.exports = {
  MAX_INLINE_ATTACHMENT_BYTES,
  fetchBinary,
  loadOneDriveAsAttachment,
  normalizeAttachment,
  normalizeAttachments,
  formatSize
};
