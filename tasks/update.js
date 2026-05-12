/**
 * Update an existing Microsoft To Do task. Common usage:
 *   - mark complete:        { id, status: "completed" }
 *   - change due date:      { id, dueDateTime: "2026-05-15T10:00:00", timeZone: "Europe/Berlin" }
 *   - change importance:    { id, importance: "high" }
 *   - clear reminder:       { id, clearReminder: true }
 *   - clear due date:       { id, clearDueDate: true }
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveListId } = require('./list');

function buildDateTimeTZ(value, timeZone) {
  if (!value) return undefined;
  return { dateTime: value, timeZone: timeZone || config.DEFAULT_TIMEZONE || 'UTC' };
}

async function handleUpdateTask(args) {
  const a = args || {};
  if (!a.id) {
    return { content: [{ type: 'text', text: 'Task id is required.' }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const listId = await resolveListId(accessToken, a.listId);

    const payload = {};
    if (a.title !== undefined) payload.title = a.title;
    if (a.status !== undefined) payload.status = a.status;
    if (a.importance !== undefined) payload.importance = a.importance;
    if (a.body !== undefined) {
      payload.body = { contentType: a.bodyIsHtml ? 'html' : 'text', content: a.body };
    }
    if (a.categories !== undefined) payload.categories = a.categories;

    if (a.clearDueDate) {
      payload.dueDateTime = null;
    } else if (a.dueDateTime !== undefined) {
      payload.dueDateTime = buildDateTimeTZ(a.dueDateTime, a.timeZone);
    }

    if (a.clearReminder) {
      payload.isReminderOn = false;
      payload.reminderDateTime = null;
    } else if (a.reminderDateTime !== undefined) {
      payload.isReminderOn = true;
      payload.reminderDateTime = buildDateTimeTZ(a.reminderDateTime, a.timeZone);
    }

    if (Object.keys(payload).length === 0) {
      return { content: [{ type: 'text', text: 'No fields to update.' }] };
    }

    const updated = await callGraphAPI(
      accessToken,
      'PATCH',
      `me/todo/lists/${listId}/tasks/${encodeURIComponent(a.id)}`,
      payload
    );

    return {
      content: [{
        type: 'text',
        text:
          `Task updated.\n\n` +
          `Title:   ${updated.title}\n` +
          `Status:  ${updated.status}\n` +
          `ID:      ${updated.id}\n` +
          `ListID:  ${listId}` +
          (updated.dueDateTime ? `\nDue:     ${updated.dueDateTime.dateTime} (${updated.dueDateTime.timeZone})` : '') +
          (updated.completedDateTime ? `\nCompleted: ${updated.completedDateTime.dateTime}` : '')
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: 'text', text: "Authentication required. Use 'authenticate' first." }] };
    }
    return { content: [{ type: 'text', text: `Error updating task: ${error.message}` }] };
  }
}

module.exports = handleUpdateTask;
