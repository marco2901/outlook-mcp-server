/**
 * Delete a Microsoft To Do task.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveListId } = require('./list');

async function handleDeleteTask(args) {
  const a = args || {};
  if (!a.id) {
    return { content: [{ type: 'text', text: 'Task id is required.' }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const listId = await resolveListId(accessToken, a.listId);

    await callGraphAPI(
      accessToken,
      'DELETE',
      `me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(a.id)}`
    );

    return {
      content: [{
        type: 'text',
        text: `Task ${a.id} deleted from list ${listId}.`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: 'text', text: "Authentication required. Use 'authenticate' first." }] };
    }
    return { content: [{ type: 'text', text: `Error deleting task: ${error.message}` }] };
  }
}

module.exports = handleDeleteTask;
