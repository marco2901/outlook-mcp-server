/**
 * Move / rename a OneDrive item.
 *
 * Uses PATCH /me/drive/items/{id} with a new parentReference (for move) and
 * optional new name (for rename). Source can be given as itemId or a
 * source path. Destination is either destPath (folder to move into) or
 * newName (rename in place) — both can be combined for "move + rename".
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

function normPath(p) {
  return String(p).replace(/^\/+|\/+$/g, '');
}

function splitParentAndName(fullPath) {
  const clean = normPath(fullPath);
  const idx = clean.lastIndexOf('/');
  if (idx === -1) return { parentPath: '', name: clean };
  return { parentPath: clean.slice(0, idx), name: clean.slice(idx + 1) };
}

async function resolveItemId(accessToken, { itemId, path }) {
  if (itemId) return itemId;
  if (!path) throw new Error("Provide either 'sourcePath' or 'itemId'.");
  const item = await callGraphAPI(accessToken, 'GET', `me/drive/root:/${normPath(path)}`);
  if (!item || !item.id) throw new Error(`Source not found: ${path}`);
  return item.id;
}

async function resolveFolderId(accessToken, folderPath) {
  const p = normPath(folderPath);
  if (p === '') {
    const root = await callGraphAPI(accessToken, 'GET', 'me/drive/root');
    return root.id;
  }
  const item = await callGraphAPI(accessToken, 'GET', `me/drive/root:/${p}`);
  if (!item || !item.id) throw new Error(`Destination folder not found: ${folderPath}`);
  return item.id;
}

async function handleMove(args) {
  const a = args || {};
  const sourceItemId = a.sourceItemId;
  const sourcePath = a.sourcePath;
  const destPath = a.destPath;    // parent folder to move into (optional)
  const newName = a.newName;      // rename (optional)

  if (!sourceItemId && !sourcePath) {
    return { content: [{ type: 'text', text: "Either 'sourcePath' or 'sourceItemId' is required." }] };
  }
  if (!destPath && !newName) {
    return { content: [{ type: 'text', text: "Provide 'destPath' (folder to move into), 'newName' (rename in place), or both." }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const itemId = await resolveItemId(accessToken, { itemId: sourceItemId, path: sourcePath });

    const patch = {};
    if (destPath) {
      const parentId = await resolveFolderId(accessToken, destPath);
      patch.parentReference = { id: parentId };
    }
    if (newName) {
      patch.name = newName;
    }

    const moved = await callGraphAPI(
      accessToken,
      'PATCH',
      `me/drive/items/${encodeURIComponent(itemId)}`,
      patch
    );

    return {
      content: [{
        type: 'text',
        text:
          `Moved "${moved.name}" successfully.\n\n` +
          `ID:      ${moved.id}\n` +
          `Web URL: ${moved.webUrl}` +
          (moved.parentReference && moved.parentReference.path
            ? `\nParent:  ${moved.parentReference.path}`
            : '')
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: 'text', text: "Authentication required. Use 'authenticate' first." }] };
    }
    return { content: [{ type: 'text', text: `Error moving item: ${error.message}` }] };
  }
}

module.exports = handleMove;
