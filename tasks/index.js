/**
 * Microsoft To Do tasks module (the modern replacement for the deprecated
 * /me/outlook/tasks endpoint). Surfaces the same data as the Outlook web
 * "Aufgaben" view and the Microsoft To Do mobile/desktop apps.
 */
const handleListTaskLists = require('./list-lists');
const { handleListTasks } = require('./list');
const handleGetTask = require('./get');
const handleCreateTask = require('./create');
const handleUpdateTask = require('./update');
const handleDeleteTask = require('./delete');

const LIST_ID_DESC =
  'Task list reference. Accepts the list id, a wellknownListName ' +
  "(e.g. \"defaultList\" for Outlook's standard \"Aufgaben\" / Tasks list) " +
  'or a list displayName. If omitted, the default list is used.';

const TIMEZONE_DESC =
  'IANA timezone string (e.g. "Europe/Berlin"). Defaults to the server\'s ' +
  'DEFAULT_TIMEZONE config when omitted.';

const tasksTools = [
  {
    name: 'list-task-lists',
    description: 'Lists all Microsoft To Do task lists (incl. the default Outlook Aufgaben list and any custom lists).',
    inputSchema: { type: 'object', properties: {}, required: [] },
    handler: handleListTaskLists
  },
  {
    name: 'list-tasks',
    description: 'Lists tasks inside a Microsoft To Do list. Defaults to the standard Outlook tasks list and excludes completed items unless includeCompleted=true.',
    inputSchema: {
      type: 'object',
      properties: {
        listId: { type: 'string', description: LIST_ID_DESC },
        count: { type: 'number', description: 'Maximum number of tasks to return (default 25, max 50).' },
        includeCompleted: { type: 'boolean', description: 'Include completed tasks (default false).' }
      },
      required: []
    },
    handler: handleListTasks
  },
  {
    name: 'get-task',
    description: 'Fetches a single Microsoft To Do task with full body / notes.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'Task ID from list-tasks.' },
        listId: { type: 'string', description: LIST_ID_DESC }
      },
      required: ['id']
    },
    handler: handleGetTask
  },
  {
    name: 'create-task',
    description: 'Creates a new Microsoft To Do task (synced to Outlook Aufgaben + the To Do apps).',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'Task title.' },
        body: { type: 'string', description: 'Notes / description.' },
        bodyIsHtml: { type: 'boolean', description: 'Treat body as HTML rather than plain text.' },
        importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Default normal.' },
        status: {
          type: 'string',
          enum: ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'],
          description: 'Initial status (default notStarted).'
        },
        dueDateTime: { type: 'string', description: 'ISO 8601 due date/time.' },
        reminderDateTime: { type: 'string', description: 'ISO 8601 reminder date/time. Enables the reminder automatically.' },
        timeZone: { type: 'string', description: TIMEZONE_DESC },
        categories: {
          type: 'array',
          items: { type: 'string' },
          description: 'Outlook category labels.'
        },
        listId: { type: 'string', description: LIST_ID_DESC }
      },
      required: ['title']
    },
    handler: handleCreateTask
  },
  {
    name: 'update-task',
    description: 'Updates a Microsoft To Do task. Common cases: mark complete (status="completed"), change due date, change importance, clear reminder/due (clearReminder/clearDueDate=true).',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'Task ID.' },
        listId: { type: 'string', description: LIST_ID_DESC },
        title: { type: 'string' },
        body: { type: 'string' },
        bodyIsHtml: { type: 'boolean' },
        status: {
          type: 'string',
          enum: ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred']
        },
        importance: { type: 'string', enum: ['low', 'normal', 'high'] },
        dueDateTime: { type: 'string', description: 'ISO 8601. Set null/omit and pass clearDueDate=true to remove.' },
        clearDueDate: { type: 'boolean' },
        reminderDateTime: { type: 'string', description: 'ISO 8601. Enables the reminder.' },
        clearReminder: { type: 'boolean' },
        timeZone: { type: 'string', description: TIMEZONE_DESC },
        categories: { type: 'array', items: { type: 'string' } }
      },
      required: ['id']
    },
    handler: handleUpdateTask
  },
  {
    name: 'delete-task',
    description: 'Permanently deletes a Microsoft To Do task.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'Task ID.' },
        listId: { type: 'string', description: LIST_ID_DESC }
      },
      required: ['id']
    },
    handler: handleDeleteTask
  }
];

module.exports = {
  tasksTools,
  handleListTaskLists,
  handleListTasks,
  handleGetTask,
  handleCreateTask,
  handleUpdateTask,
  handleDeleteTask
};
