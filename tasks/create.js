/**
 * Create a new Microsoft To Do task.
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveListId } = require('./list');

function buildDateTimeTZ(value, timeZone) {
  if (!value) return undefined;
  // Accept ISO strings; pass through to Graph along with timezone.
  return { dateTime: value, timeZone: timeZone || config.DEFAULT_TIMEZONE || 'UTC' };
}

async function handleCreateTask(args) {
  const a = args || {};
  if (!a.title) {
    return { content: [{ type: 'text', text: 'Title is required.' }] };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const listId = await resolveListId(accessToken, a.listId);

    const payload = {
      title: a.title,
      importance: a.importance || 'normal',
      status: a.status || 'notStarted',
      ...(a.body
        ? { body: { contentType: a.bodyIsHtml ? 'html' : 'text', content: a.body } }
        : {}),
      ...(a.dueDateTime
        ? { dueDateTime: buildDateTimeTZ(a.dueDateTime, a.timeZone) }
        : {}),
      ...(a.reminderDateTime
        ? {
            isReminderOn: true,
            reminderDateTime: buildDateTimeTZ(a.reminderDateTime, a.timeZone)
          }
        : {}),
      ...(Array.isArray(a.categories) && a.categories.length > 0
        ? { categories: a.categories }
        : {})
    };

    const created = await callGraphAPI(
      accessToken,
      'POST',
      `me/todo/lists/${listId}/tasks`,
      payload
    );

    return {
      content: [{
        type: 'text',
        text:
          `Task created.\n\n` +
          `Title:   ${created.title}\n` +
          `Status:  ${created.status}\n` +
          `ID:      ${created.id}\n` +
          `ListID:  ${listId}` +
          (created.dueDateTime ? `\nDue:     ${created.dueDateTime.dateTime} (${created.dueDateTime.timeZone})` : '')
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: 'text', text: "Authentication required. Use 'authenticate' first." }] };
    }
    return { content: [{ type: 'text', text: `Error creating task: ${error.message}` }] };
  }
}

module.exports = handleCreateTask;
