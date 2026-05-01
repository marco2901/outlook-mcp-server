/**
 * Email module for Outlook MCP server
 */
const handleListEmails = require('./list');
const handleSearchEmails = require('./search');
const handleReadEmail = require('./read');
const handleSendEmail = require('./send');
const handleDraftEmail = require('./draft');
const handleMarkAsRead = require('./mark-as-read');
const handleDeleteEmail = require('./delete');
const handleListAttachments = require('./list-attachments');
const handleGetAttachment = require('./get-attachment');

// Reusable schema fragment for the `attachments` array on send-email / draft-email.
const ATTACHMENTS_SCHEMA = {
  type: "array",
  description:
    "Optional file attachments. Each entry must provide either `contentBytes` (base64 of the file, " +
    "<= 3 MB) OR `onedrivePath` / `onedriveItemId` to attach a file from OneDrive. " +
    "For files larger than 3 MB use onedrive-share to send a link instead.",
  items: {
    type: "object",
    properties: {
      name: { type: "string", description: "Filename shown to the recipient" },
      contentType: { type: "string", description: "MIME type (e.g. application/pdf). Optional, sniffed from OneDrive if omitted." },
      contentBytes: { type: "string", description: "Base64-encoded file content. Mutually exclusive with onedrivePath/onedriveItemId." },
      onedrivePath: { type: "string", description: "OneDrive path, e.g. /Documents/foo.pdf. Loaded into the attachment server-side." },
      onedriveItemId: { type: "string", description: "OneDrive item ID. Alternative to onedrivePath." }
    }
  }
};

// Email tool definitions
const emailTools = [
  {
    name: "list-emails",
    description: "Lists recent emails from your inbox",
    inputSchema: {
      type: "object",
      properties: {
        folder: {
          type: "string",
          description: "Email folder to list (e.g., 'inbox', 'sent', 'drafts', default: 'inbox')"
        },
        count: {
          type: "number",
          description: "Number of emails to retrieve (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleListEmails
  },
  {
    name: "search-emails",
    description: "Search for emails using various criteria",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query text to find in emails"
        },
        folder: {
          type: "string",
          description: "Email folder to search in (default: 'inbox')"
        },
        from: {
          type: "string",
          description: "Filter by sender email address or name"
        },
        to: {
          type: "string",
          description: "Filter by recipient email address or name"
        },
        subject: {
          type: "string",
          description: "Filter by email subject"
        },
        hasAttachments: {
          type: "boolean",
          description: "Filter to only emails with attachments"
        },
        unreadOnly: {
          type: "boolean",
          description: "Filter to only unread emails"
        },
        count: {
          type: "number",
          description: "Number of results to return (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleSearchEmails
  },
  {
    name: "read-email",
    description: "Reads the content of a specific email. HTML emails are securely sanitized to extract only visible text, preventing prompt injection attacks via hidden content.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to read"
        },
        includeRawHtml: {
          type: "boolean",
          description: "Include raw HTML content (UNSAFE - for debugging only, may contain hidden prompt injection content)"
        }
      },
      required: ["id"]
    },
    handler: handleReadEmail
  },
  {
    name: "send-email",
    description: "Composes and sends a new email. Supports plain text/HTML content and file attachments (inline base64 or loaded from OneDrive).",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Email subject"
        },
        body: {
          type: "string",
          description: "Email body content (plain text or HTML)"
        },
        isHtml: {
          type: "boolean",
          description: "Set to true to send as HTML, false for plain text. If not specified, auto-detects based on <html> tag presence."
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        },
        saveToSentItems: {
          type: "boolean",
          description: "Whether to save the email to sent items"
        },
        attachments: ATTACHMENTS_SCHEMA
      },
      required: ["to", "subject", "body"]
    },
    handler: handleSendEmail
  },
  {
    name: "draft-email",
    description: "Creates and saves an email draft in Outlook. Supports file attachments (inline base64 or loaded from OneDrive).",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Draft email subject"
        },
        body: {
          type: "string",
          description: "Draft email body content (can be plain text or HTML)"
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        },
        attachments: ATTACHMENTS_SCHEMA
      },
      required: []
    },
    handler: handleDraftEmail
  },
  {
    name: "mark-as-read",
    description: "Marks an email as read or unread",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to mark as read/unread"
        },
        isRead: {
          type: "boolean",
          description: "Whether to mark as read (true) or unread (false). Default: true"
        }
      },
      required: ["id"]
    },
    handler: handleMarkAsRead
  },
  {
    name: "delete-email",
    description: "Deletes an email by moving it to Deleted Items (trash). Use permanent=true to hard delete.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to delete"
        },
        permanent: {
          type: "boolean",
          description: "If true, permanently delete the email instead of moving to Deleted Items. Default: false"
        }
      },
      required: ["id"]
    },
    handler: handleDeleteEmail
  },
  {
    name: "list-attachments",
    description: "Lists the attachments on a single email (id, name, contentType, size, isInline). Use get-attachment to download one.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "Email ID" }
      },
      required: ["id"]
    },
    handler: handleListAttachments
  },
  {
    name: "get-attachment",
    description: "Downloads a single email attachment. By default returns base64 inline (for files < 256 KB). Pass saveToOneDrive to persist the file to a given OneDrive path instead.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "Email ID containing the attachment" },
        attachmentId: { type: "string", description: "Attachment ID from list-attachments" },
        saveToOneDrive: {
          type: "string",
          description: "Optional. If set, decode the file and upload to this OneDrive path (e.g. \"/Documents/from-mail/foo.pdf\") instead of returning base64."
        }
      },
      required: ["id", "attachmentId"]
    },
    handler: handleGetAttachment
  }
];

module.exports = {
  emailTools,
  handleListEmails,
  handleSearchEmails,
  handleReadEmail,
  handleSendEmail,
  handleDraftEmail,
  handleMarkAsRead,
  handleDeleteEmail,
  handleListAttachments,
  handleGetAttachment
};
