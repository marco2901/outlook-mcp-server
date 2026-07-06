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

async function handleUpload(args) {
  const a = args || {};
  const path = a.path;
  const conflictBehavior = a.conflictBehavior || 'rename';

  if (!path) {
    return { content: [{ type: 'text', text: "Path is required (e.g. '/Documents/myfile.txt' — start with '/'). Include the filename." }] };
  }

  let buffer;
  if (a.contentBase64) {
    try {
      buffer = Buffer.from(a.contentBase64, 'base64');
    } catch (e) {
      return { content: [{ type: 'text', text: `contentBase64 is not valid base64: ${e.message}` }] };
    }
  } else if (typeof a.content === 'string') {
    buffer = Buffer.from(a.content, 'utf8');
  } else {
    return { content: [{ type: 'text', text: 'Either content (UTF-8 text) or contentBase64 (base64 bytes) is required.' }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const uploaded = await uploadBufferToOneDrive(accessToken, path, buffer, { conflictBehavior });
    return {
      content: [{
        type: 'text',
        text: `Successfully uploaded "${uploaded.name}" (${formatSize(uploaded.size)})\n\nID: ${uploaded.id}\nWeb URL: ${uploaded.webUrl}`
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
