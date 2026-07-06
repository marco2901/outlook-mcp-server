/**
 * OneDrive module for Outlook MCP server
 */
const handleListFiles = require('./list');
const handleSearchFiles = require('./search');
const handleDownload = require('./download');
const handleUpload = require('./upload');
const handleUploadLarge = require('./upload-large');
const handleShare = require('./share');
const handleMove = require('./move');
const { handleCreateFolder, handleDeleteItem } = require('./folder');

// Standard hint used across path-taking tools so Claude stops guessing.
const PATH_HINT =
  "OneDrive path starting with '/' (forward slashes), e.g. '/Documents/foo.pdf' " +
  "or '/mcp-test/Ordner mit Leerzeichen/Datei.txt'. Leading/trailing slashes are tolerated. " +
  "For files include the filename.";

const onedriveTools = [
  {
    name: "onedrive-list",
    description: "List files and folders inside a OneDrive folder. Argument: 'path' (folder). Default: root.",
    inputSchema: {
      type: "object",
      properties: {
        path: { type: "string", description: `Folder path. ${PATH_HINT} Default: root.` },
        count: { type: "number", description: "Max items to return (default 25, max 50)." }
      },
      required: []
    },
    handler: handleListFiles
  },
  {
    name: "onedrive-search",
    description: "Search OneDrive by filename or content substring. Argument: 'query' (required).",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Text to search for in filenames / contents." },
        count: { type: "number", description: "Max results (default 25, max 50)." }
      },
      required: ["query"]
    },
    handler: handleSearchFiles
  },
  {
    name: "onedrive-download",
    description: "Get a pre-authenticated download URL for a OneDrive file. Argument: 'path' OR 'itemId'.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "OneDrive item ID (from onedrive-list / -search)." },
        path: { type: "string", description: `File path. ${PATH_HINT}` }
      },
      required: []
    },
    handler: handleDownload
  },
  {
    name: "onedrive-upload",
    description:
      "Upload a file to OneDrive. Automatically switches to chunked upload session for files over 4 MB. " +
      "For TEXT files pass 'content' (UTF-8 string). For BINARY files (PDF, JPG, PNG, DOCX, XLSX, ZIP, …) " +
      "pass 'contentBase64' (base64-encoded bytes) — passing binary via 'content' produces a corrupt file.",
    inputSchema: {
      type: "object",
      properties: {
        path: { type: "string", description: `Destination file path including filename. ${PATH_HINT}` },
        content: { type: "string", description: "UTF-8 text content. Use ONLY for plain-text files." },
        contentBase64: { type: "string", description: "Base64-encoded file bytes. Use for anything binary (PDF/JPG/DOCX/…)." },
        conflictBehavior: {
          type: "string",
          description: "What to do if a file with the same name exists at path.",
          enum: ["rename", "replace", "fail"]
        }
      },
      required: ["path"]
    },
    handler: handleUpload
  },
  {
    name: "onedrive-upload-large",
    description:
      "Deprecated alias for onedrive-upload — the regular upload tool now handles files of any size. " +
      "Kept for backward compatibility. Same arguments as onedrive-upload (content/contentBase64).",
    inputSchema: {
      type: "object",
      properties: {
        path: { type: "string", description: `Destination file path including filename. ${PATH_HINT}` },
        content: { type: "string", description: "UTF-8 text content." },
        contentBase64: { type: "string", description: "Base64-encoded bytes for binary files." },
        conflictBehavior: { type: "string", enum: ["rename", "replace", "fail"] }
      },
      required: ["path"]
    },
    handler: handleUploadLarge
  },
  {
    name: "onedrive-share",
    description: "Create a sharing link for a OneDrive file or folder. Argument: 'path' OR 'itemId'.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "OneDrive item ID." },
        path: { type: "string", description: `Item path. ${PATH_HINT}` },
        type: { type: "string", enum: ["view", "edit", "embed"], description: "Link permission. Default 'view'." },
        scope: { type: "string", enum: ["anonymous", "organization"], description: "Who can use the link. Default 'anonymous'." }
      },
      required: []
    },
    handler: handleShare
  },
  {
    name: "onedrive-create-folder",
    description: "Create a new folder in OneDrive. Arguments: 'path' (parent folder), 'name' (new folder's name).",
    inputSchema: {
      type: "object",
      properties: {
        path: { type: "string", description: `Parent folder path. ${PATH_HINT} Default: root.` },
        name: { type: "string", description: "Name of the new folder (not the full path — just the folder name)." }
      },
      required: ["name"]
    },
    handler: handleCreateFolder
  },
  {
    name: "onedrive-delete",
    description:
      "Delete a OneDrive file or folder (goes to Recycle Bin, recoverable). Argument: 'path' OR 'itemId'.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: { type: "string", description: "OneDrive item ID." },
        path: { type: "string", description: `Item path. ${PATH_HINT}` }
      },
      required: []
    },
    handler: handleDeleteItem
  },
  {
    name: "onedrive-move",
    description:
      "Move and/or rename a OneDrive file or folder. Provide the source via 'sourcePath' or 'sourceItemId'. " +
      "Pass 'destPath' to move into a target folder, 'newName' to rename, or both together for move+rename.",
    inputSchema: {
      type: "object",
      properties: {
        sourcePath: { type: "string", description: `Path of the item to move. ${PATH_HINT}` },
        sourceItemId: { type: "string", description: "OneDrive item ID of the source (alternative to sourcePath)." },
        destPath: { type: "string", description: `Target FOLDER to move the item into. ${PATH_HINT} (Folder must exist.)` },
        newName: { type: "string", description: "New name for the item after moving (rename operation)." }
      },
      required: []
    },
    handler: handleMove
  }
];

module.exports = {
  onedriveTools,
  handleListFiles,
  handleSearchFiles,
  handleDownload,
  handleUpload,
  handleUploadLarge,
  handleShare,
  handleCreateFolder,
  handleDeleteItem,
  handleMove
};
