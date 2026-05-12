/**
 * List all Microsoft To Do task lists (incl. the default Outlook "Aufgaben" list).
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

async function handleListTaskLists() {
  try {
    const accessToken = await ensureAuthenticated();
    const response = await callGraphAPI(accessToken, 'GET', 'me/todo/lists', null, {
      $select: 'id,displayName,wellknownListName,isOwner,isShared'
    });

    const lists = (response && response.value) || [];
    if (lists.length === 0) {
      return { content: [{ type: 'text', text: 'No task lists found.' }] };
    }

    const lines = lists.map((l, i) =>
      `${i + 1}. ${l.displayName}${l.wellknownListName && l.wellknownListName !== 'none' ? ` (${l.wellknownListName})` : ''}\n` +
      `   id:     ${l.id}\n` +
      `   owner:  ${l.isOwner ? 'yes' : 'no'}, shared: ${l.isShared ? 'yes' : 'no'}`
    );

    return {
      content: [{
        type: 'text',
        text: `Found ${lists.length} task list(s):\n\n${lines.join('\n\n')}\n\n` +
              `Use list-tasks with listId to see tasks inside a list.`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: 'text', text: "Authentication required. Use 'authenticate' first." }] };
    }
    return { content: [{ type: 'text', text: `Error listing task lists: ${error.message}` }] };
  }
}

module.exports = handleListTaskLists;
