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
  // Accept a real list id, a wellknownListName ("defaultList", "flaggedEmails"),
  // or the displayName of an existing list. Returns the resolved id.
  // Strategy: always pull the lists once and match client-side — the
  // /me/todo/lists $filter API needs the fully-qualified enum cast
  // (microsoft.toDo.wellknownListName'defaultList') and is brittle.
  const response = await callGraphAPI(accessToken, 'GET', 'me/todo/lists', null, {
    $select: 'id,displayName,wellknownListName'
  });
  const lists = (response && response.value) || [];

  if (!listIdOrName) {
    const def = lists.find((l) => l.wellknownListName === 'defaultList');
    if (def) return def.id;
    throw new Error('No default task list found.');
  }

  // Exact id match first
  const byId = lists.find((l) => l.id === listIdOrName);
  if (byId) return byId.id;

  // Then displayName / wellknownListName
  const byName = lists.find(
    (l) => l.displayName === listIdOrName || l.wellknownListName === listIdOrName
  );
  if (byName) return byName.id;

  // Fall back: treat as opaque id (lets callers pass an id the lists call
  // didn't return, e.g. a shared list)
  if (/^[A-Za-z0-9_\-=+/]{20,}$/.test(listIdOrName)) return listIdOrName;

  throw new Error(`Task list "${listIdOrName}" not found.`);
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
      `me/todo/lists/${encodeURIComponent(listId)}/tasks`,
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
