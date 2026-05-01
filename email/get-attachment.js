/**
 * Download a single email attachment.
 *
 * Two output modes:
 *   - default: returns the file metadata + base64 contentBytes inline in the
 *     MCP response. Caps inline base64 at 256 KB to keep the response readable;
 *     larger files require saveToOneDrive.
 *   - saveToOneDrive: decodes the bytes and uploads to the given OneDrive path,
 *     returning the upload result instead of the bytes.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { formatSize } = require('../utils/attachment-helpers');

const INLINE_BASE64_CAP = 256 * 1024; // bytes (decoded) — keep MCP responses sane

async function handleGetAttachment(args) {
  const emailId = args && args.id;
  const attachmentId = args && args.attachmentId;
  const saveToOneDrive = args && args.saveToOneDrive;

  if (!emailId || !attachmentId) {
    return { content: [{ type: 'text', text: 'Both id (email) and attachmentId are required.' }] };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const att = await callGraphAPI(
      accessToken,
      'GET',
      `me/messages/${encodeURIComponent(emailId)}/attachments/${encodeURIComponent(attachmentId)}`
    );

    if (!att) {
      return { content: [{ type: 'text', text: 'Attachment not found.' }] };
    }

    if (att['@odata.type'] !== '#microsoft.graph.fileAttachment') {
      return {
        content: [{
          type: 'text',
          text: `Attachment "${att.name}" is of type ${att['@odata.type']} ` +
                `(item or reference attachment), which this tool does not yet support. ` +
                `For reference attachments (e.g. OneDrive links) the URL is in the message body.`
        }]
      };
    }

    if (saveToOneDrive) {
      // Upload bytes to OneDrive at the given path
      const buffer = Buffer.from(att.contentBytes, 'base64');
      const normalizedPath = String(saveToOneDrive).replace(/^\/+|\/+$/g, '');
      const endpoint = `me/drive/root:/${normalizedPath}:/content`;
      const uploaded = await callGraphAPI(
        accessToken,
        'PUT',
        endpoint,
        buffer.toString('binary'),
        { '@microsoft.graph.conflictBehavior': 'rename' }
      );
      return {
        content: [{
          type: 'text',
          text: `Saved attachment "${att.name}" (${formatSize(att.size)}) to OneDrive:\n\n` +
                `Name:    ${uploaded.name}\nSize:    ${formatSize(uploaded.size)}\n` +
                `Web URL: ${uploaded.webUrl}\nID:      ${uploaded.id}`
        }]
      };
    }

    // Inline mode: only return base64 if the file is small enough
    if (att.size > INLINE_BASE64_CAP) {
      return {
        content: [{
          type: 'text',
          text: `Attachment "${att.name}" (${formatSize(att.size)}) is too large to return ` +
                `inline (cap: ${formatSize(INLINE_BASE64_CAP)}). ` +
                `Re-call with saveToOneDrive: "/path/to/file.ext" to persist it instead.`
        }]
      };
    }

    return {
      content: [{
        type: 'text',
        text: `Name:        ${att.name}\nContentType: ${att.contentType}\n` +
              `Size:        ${formatSize(att.size)}\nIsInline:    ${att.isInline ? 'yes' : 'no'}\n\n` +
              `--- base64 contentBytes ---\n${att.contentBytes}\n--- end ---`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: 'text', text: "Authentication required. Please use the 'authenticate' tool first." }] };
    }
    return { content: [{ type: 'text', text: `Error fetching attachment: ${error.message}` }] };
  }
}

module.exports = handleGetAttachment;
