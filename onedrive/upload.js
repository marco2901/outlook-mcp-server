/**
 * OneDrive upload tool — accepts UTF-8 text (`content`) or base64 binary
 * (`contentBase64`). Automatically picks simple PUT vs upload-session
 * chunked upload based on size, so a separate large-file tool is no
 * longer required (kept as an alias for backward compat).
 */
const { ensureAuthenticated } = require('../auth');
const { uploadBufferToOneDrive } = require('../utils/onedrive-binary-upload');

function formatSize(bytes) {
  if (!bytes) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Common binary magic byte signatures. If a supposedly-text `content` argument
// decodes as base64 into one of these, the caller almost certainly meant to
// upload binary and mistakenly passed base64 via `content` instead of
// `contentBase64`. Rescue silently so the upload isn't corrupt.
const BINARY_MAGICS = [
  Buffer.from([0x25, 0x50, 0x44, 0x46]),       // %PDF
  Buffer.from([0x89, 0x50, 0x4e, 0x47]),       // .PNG
  Buffer.from([0xff, 0xd8, 0xff]),             // JPEG
  Buffer.from([0x47, 0x49, 0x46, 0x38]),       // GIF8
  Buffer.from([0x50, 0x4b, 0x03, 0x04]),       // PK.. — ZIP / DOCX / XLSX / PPTX
  Buffer.from([0x50, 0x4b, 0x05, 0x06]),       // PK.. — empty ZIP
  Buffer.from([0xd0, 0xcf, 0x11, 0xe0]),       // MS-CFB — old .doc/.xls/.ppt
  Buffer.from([0x52, 0x61, 0x72, 0x21]),       // Rar!
  Buffer.from([0x1f, 0x8b, 0x08]),             // gzip
  Buffer.from([0x42, 0x4d]),                   // BM — bitmap
  Buffer.from([0x00, 0x00, 0x01, 0x00])        // ICO
];

function looksLikeBase64(str) {
  return typeof str === 'string' &&
    str.length >= 24 &&
    str.length % 4 === 0 &&
    /^[A-Za-z0-9+/]+={0,2}$/.test(str);
}

function tryDecodeBase64IfBinary(content) {
  if (!looksLikeBase64(content)) return null;
  let decoded;
  try {
    decoded = Buffer.from(content, 'base64');
  } catch { return null; }
  // Round-trip: real base64 re-encodes to the original string.
  if (decoded.toString('base64') !== content) return null;
  // Only rescue if the decoded bytes start with a known binary magic —
  // otherwise leave the ambiguity alone (avoids turning plain-text that
  // happens to be all-base64 into binary garbage).
  if (BINARY_MAGICS.some((m) => decoded.length >= m.length && decoded.subarray(0, m.length).equals(m))) {
    return decoded;
  }
  return null;
}

async function handleUpload(args) {
  const a = args || {};
  const path = a.path;
  const conflictBehavior = a.conflictBehavior || 'rename';

  if (!path) {
    return { content: [{ type: 'text', text: "Path is required (e.g. '/Documents/myfile.txt' — start with '/'). Include the filename." }] };
  }

  let buffer;
  let rescued = false;
  if (a.contentBase64) {
    try {
      buffer = Buffer.from(a.contentBase64, 'base64');
    } catch (e) {
      return { content: [{ type: 'text', text: `contentBase64 is not valid base64: ${e.message}` }] };
    }
  } else if (typeof a.content === 'string') {
    const rescueBinary = tryDecodeBase64IfBinary(a.content);
    if (rescueBinary) {
      buffer = rescueBinary;
      rescued = true;
    } else {
      buffer = Buffer.from(a.content, 'utf8');
    }
  } else {
    return { content: [{ type: 'text', text: 'Either content (UTF-8 text) or contentBase64 (base64 bytes) is required.' }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const uploaded = await uploadBufferToOneDrive(accessToken, path, buffer, { conflictBehavior });
    const note = rescued
      ? '\n\nNote: `content` looked like base64-encoded binary and started with a known magic byte, so it was decoded before upload. Next time pass it via `contentBase64` instead.'
      : '';
    return {
      content: [{
        type: 'text',
        text: `Successfully uploaded "${uploaded.name}" (${formatSize(uploaded.size)})\n\nID: ${uploaded.id}\nWeb URL: ${uploaded.webUrl}${note}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: 'text', text: "Authentication required. Use 'authenticate' first." }] };
    }
    return { content: [{ type: 'text', text: `Error uploading file: ${error.message}` }] };
  }
}

module.exports = handleUpload;
