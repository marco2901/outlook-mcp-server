/**
 * Fetch a single Microsoft To Do task with full body.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveListId, formatDateTimeTZ } = require('./list');

async function handleGetTask(args) {
  const taskId = args && args.id;
  const listRef = args && args.listId;
  if (!taskId) {
    return { content: [{ type: 'text', text: 'Task id is required.' }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const listId = await resolveListId(accessToken, listRef);

    const task = await callGraphAPI(
      accessToken,
      'GET',
      `me/todo/lists/${listId}/tasks/${encodeURIComponent(taskId)}`
    );

    if (!task) return { content: [{ type: 'text', text: 'Task not found.' }] };

    const body = task.body && task.body.content ? task.body.content : '(no notes)';
    const text =
      `Title:        ${task.title}\n` +
      `Status:       ${task.status}\n` +
      `Importance:   ${task.importance}\n` +
      (task.dueDateTime ? `Due:          ${formatDateTimeTZ(task.dueDateTime)}\n` : '') +
      (task.reminderDateTime ? `Reminder:     ${formatDateTimeTZ(task.reminderDateTime)}\n` : '') +
      (task.completedDateTime ? `Completed:    ${formatDateTimeTZ(task.completedDateTime)}\n` : '') +
      (task.categories && task.categories.length > 0 ? `Categories:   ${task.categories.join(', ')}\n` : '') +
      `Created:      ${task.createdDateTime}\n` +
      `LastModified: ${task.lastModifiedDateTime}\n` +
      `ID:           ${task.id}\n` +
      `ListID:       ${listId}\n` +
      `\n--- Notes ---\n${body}`;

    return { content: [{ type: 'text', text }] };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: 'text', text: "Authentication required. Use 'authenticate' first." }] };
    }
    return { content: [{ type: 'text', text: `Error fetching task: ${error.message}` }] };
  }
}

module.exports = handleGetTask;
