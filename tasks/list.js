/**
 * List tasks inside a Microsoft To Do list.
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

const TASK_SELECT_FIELDS =
  'id,title,status,importance,dueDateTime,reminderDateTime,completedDateTime,' +
  'createdDateTime,lastModifiedDateTime,categories,hasAttachments,bodyLastModifiedDateTime';

async function resolveListId(accessToken, listIdOrName) {
  // Accept either a real list ID, the wellknownListName ("defaultList", "flaggedEmails"),
  // or the displayName of an existing list. Returns the resolved id.
  if (!listIdOrName) {
    // Default to the well-known "defaultList" (Outlook's "Aufgaben" / "Tasks" list)
    const response = await callGraphAPI(accessToken, 'GET', 'me/todo/lists', null, {
      $filter: "wellknownListName eq 'defaultList'",
      $select: 'id,displayName'
    });
    if (response.value && response.value.length > 0) return response.value[0].id;
    throw new Error('No default task list found.');
  }

  // Treat anything starting with "AQMk" or similar Graph ID prefixes as a real id.
  // Otherwise: try to resolve as displayName or wellknownListName.
  if (/^[A-Za-z0-9_\-]{30,}$/.test(listIdOrName)) return listIdOrName;

  const response = await callGraphAPI(accessToken, 'GET', 'me/todo/lists', null, {
    $select: 'id,displayName,wellknownListName'
  });
  const match = (response.value || []).find(
    (l) => l.displayName === listIdOrName || l.wellknownListName === listIdOrName
  );
  if (!match) throw new Error(`Task list "${listIdOrName}" not found.`);
  return match.id;
}

function formatDateTimeTZ(dt) {
  if (!dt || !dt.dateTime) return '';
  return `${dt.dateTime}${dt.timeZone ? ` (${dt.timeZone})` : ''}`;
}

async function handleListTasks(args) {
  const count = Math.min((args && args.count) || 25, config.MAX_RESULT_COUNT);
  const listRef = args && args.listId;
  const includeCompleted = args && args.includeCompleted === true;

  try {
    const accessToken = await ensureAuthenticated();
    const listId = await resolveListId(accessToken, listRef);

    const queryParams = {
      $top: count,
      $orderby: 'createdDateTime desc',
      $select: TASK_SELECT_FIELDS
    };
    if (!includeCompleted) {
      queryParams.$filter = "status ne 'completed'";
    }

    const response = await callGraphAPI(
      accessToken,
      'GET',
      `me/todo/lists/${listId}/tasks`,
      null,
      queryParams
    );

    const tasks = (response && response.value) || [];
    if (tasks.length === 0) {
      return {
        content: [{
          type: 'text',
          text: `No ${includeCompleted ? '' : 'open '}tasks in this list.`
        }]
      };
    }

    const lines = tasks.map((t, i) =>
      `${i + 1}. [${t.status}] ${t.title}` +
      `${t.importance && t.importance !== 'normal' ? ` (${t.importance})` : ''}\n` +
      `   id:        ${t.id}\n` +
      (t.dueDateTime ? `   due:       ${formatDateTimeTZ(t.dueDateTime)}\n` : '') +
      (t.reminderDateTime ? `   reminder:  ${formatDateTimeTZ(t.reminderDateTime)}\n` : '') +
      (t.categories && t.categories.length > 0 ? `   tags:      ${t.categories.join(', ')}\n` : '') +
      (t.hasAttachments ? '   attachments: yes\n' : '') +
      `   created:   ${t.createdDateTime}`
    );

    return {
      content: [{
        type: 'text',
        text: `Found ${tasks.length} task(s) (listId: ${listId}):\n\n${lines.join('\n\n')}\n\n` +
              `Use get-task / update-task / delete-task with id + listId.`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: 'text', text: "Authentication required. Use 'authenticate' first." }] };
    }
    return { content: [{ type: 'text', text: `Error listing tasks: ${error.message}` }] };
  }
}

module.exports = { handleListTasks, resolveListId, formatDateTimeTZ };
