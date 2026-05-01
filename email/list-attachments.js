/**
 * List attachments on a single email.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { formatSize } = require('../utils/attachment-helpers');

async function handleListAttachments(args) {
  const emailId = args && args.id;
  if (!emailId) {
    return { content: [{ type: 'text', text: 'Email ID is required.' }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const response = await callGraphAPI(
      accessToken,
      'GET',
      `me/messages/${encodeURIComponent(emailId)}/attachments`,
      null,
      { $select: 'id,name,contentType,size,isInline' }
    );

    const items = (response && response.value) || [];
    if (items.length === 0) {
      return { content: [{ type: 'text', text: 'No attachments on this email.' }] };
    }

    const lines = items.map((a, i) =>
      `${i + 1}. ${a.name}\n` +
      `   id:           ${a.id}\n` +
      `   contentType:  ${a.contentType || 'unknown'}\n` +
      `   size:         ${formatSize(a.size)}\n` +
      `   inline:       ${a.isInline ? 'yes' : 'no'}`
    );

    return {
      content: [{
        type: 'text',
        text: `Found ${items.length} attachment(s):\n\n${lines.join('\n\n')}\n\n` +
              `Use get-attachment with id + attachmentId to download a single one.`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: 'text', text: "Authentication required. Please use the 'authenticate' tool first." }] };
    }
    return { content: [{ type: 'text', text: `Error listing attachments: ${error.message}` }] };
  }
}

module.exports = handleListAttachments;
