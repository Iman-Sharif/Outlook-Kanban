'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const board = require('../js/board/logic');

function mkPrivacyFilter() {
  return {
    all: { value: '0', text: 'All' },
    private: { value: '2', text: 'Private' },
    public: { value: '1', text: 'Not Private' }
  };
}

test('buildLanes() sorts by laneOrder when saveOrder enabled', () => {
  const config = {
    BOARD: { saveOrder: true },
    LANES: [{ id: 'backlog', title: 'Backlog', color: '#94a3b8', wipLimit: 0, enabled: true, outlookStatus: 0 }]
  };
  const tasks = [
    { entryID: '1', laneId: 'backlog', laneOrder: 2, dueDateMs: 1000, priority: 1, subject: 'B' },
    { entryID: '2', laneId: 'backlog', laneOrder: 1, dueDateMs: 500, priority: 2, subject: 'A' },
    { entryID: '3', laneId: 'backlog', laneOrder: null, dueDateMs: 0, priority: 0, subject: 'C' }
  ];

  const lanes = board.buildLanes(tasks, config);
  assert.equal(lanes.length, 1);
  assert.deepEqual(lanes[0].tasks.map(t => t.entryID), ['2', '1', '3']);
});

test('buildLanes() sorts by due date, priority, then subject when saveOrder disabled', () => {
  const config = {
    BOARD: { saveOrder: false },
    LANES: [{ id: 'backlog', title: 'Backlog', color: '#94a3b8', wipLimit: 0, enabled: true, outlookStatus: 0 }]
  };
  const tasks = [
    { entryID: 'a', laneId: 'backlog', dueDateMs: 2000, priority: 1, subject: 'Z' },
    { entryID: 'b', laneId: 'backlog', dueDateMs: 1000, priority: 1, subject: 'Z' },
    { entryID: 'c', laneId: 'backlog', dueDateMs: 1000, priority: 2, subject: 'A' },
    { entryID: 'd', laneId: 'backlog', dueDateMs: null, priority: 2, subject: 'B' }
  ];

  const lanes = board.buildLanes(tasks, config);
  // dueDate 1000 first; within same due date, priority desc; then subject
  assert.deepEqual(lanes[0].tasks.map(t => t.entryID), ['c', 'b', 'a', 'd']);
});

test('buildLanes() assigns unassigned/unknown tasks to first enabled lane', () => {
  const config = {
    BOARD: { saveOrder: false },
    LANES: [
      { id: 'backlog', title: 'Backlog', color: '#94a3b8', wipLimit: 0, enabled: false, outlookStatus: 0 },
      { id: 'doing', title: 'Doing', color: '#60a5fa', wipLimit: 0, enabled: true, outlookStatus: 1 }
    ]
  };

  const tasks = [
    { entryID: 'a', laneId: '', subject: 'Unassigned', dueDateMs: 1, priority: 0 },
    { entryID: 'b', laneId: 'old-lane', subject: 'Unknown lane', dueDateMs: 2, priority: 0 }
  ];

  const lanes = board.buildLanes(tasks, config);
  assert.equal(lanes.length, 2);
  assert.equal(lanes[0].id, 'backlog');
  assert.equal(lanes[1].id, 'doing');

  assert.deepEqual(lanes[0].tasks.map(t => t.entryID), []);
  assert.deepEqual(lanes[1].tasks.map(t => t.entryID), ['a', 'b']);
});

test('applyFilters() filters by privacy, search, and category', () => {
  const config = {
    BOARD: { saveOrder: false },
    LANES: [{ id: 'backlog', title: 'Backlog', color: '#94a3b8', wipLimit: 0, enabled: true, outlookStatus: 0 }]
  };
  const tasks = [
    {
      entryID: '1',
      laneId: 'backlog',
      subject: 'Pay invoice',
      notes: 'Finance',
      sensitivity: 0,
      categoriesCsv: 'Work',
      categories: [{ label: 'Work' }]
    },
    {
      entryID: '2',
      laneId: 'backlog',
      subject: 'Buy milk',
      notes: '',
      sensitivity: 2,
      categoriesCsv: '',
      categories: []
    }
  ];

  const lanes = board.buildLanes(tasks, config);
  const privacyFilter = mkPrivacyFilter();

  // privacy: private only
  let filter = { private: privacyFilter.private.value, search: '', category: '<All Categories>' };
  board.applyFilters(lanes, filter, privacyFilter);
  assert.deepEqual(lanes[0].filteredTasks.map(t => t.entryID), ['2']);

  // search
  filter = { private: privacyFilter.all.value, search: 'invoice', category: '<All Categories>' };
  board.applyFilters(lanes, filter, privacyFilter);
  assert.deepEqual(lanes[0].filteredTasks.map(t => t.entryID), ['1']);

  // category
  filter = { private: privacyFilter.all.value, search: '', category: 'Work' };
  board.applyFilters(lanes, filter, privacyFilter);
  assert.deepEqual(lanes[0].filteredTasks.map(t => t.entryID), ['1']);

  // no category
  filter = { private: privacyFilter.all.value, search: '', category: '<No Category>' };
  board.applyFilters(lanes, filter, privacyFilter);
  assert.deepEqual(lanes[0].filteredTasks.map(t => t.entryID), ['2']);
});

test('applyFilters() filters by due date buckets', () => {
  const config = {
    BOARD: { saveOrder: false },
    LANES: [{ id: 'backlog', title: 'Backlog', color: '#94a3b8', wipLimit: 0, enabled: true, outlookStatus: 0 }]
  };

  const tasks = [
    // Include dueDateMs to keep buildLanes() ordering stable.
    { entryID: 'o', laneId: 'backlog', subject: 'Overdue', sensitivity: 0, categoriesCsv: '', categories: [], dueDaysFromToday: -2, dueDateMs: 100 },
    { entryID: 't', laneId: 'backlog', subject: 'Today', sensitivity: 0, categoriesCsv: '', categories: [], dueDaysFromToday: 0, dueDateMs: 200 },
    { entryID: 'n', laneId: 'backlog', subject: 'Next7', sensitivity: 0, categoriesCsv: '', categories: [], dueDaysFromToday: 6, dueDateMs: 300 },
    { entryID: 'f', laneId: 'backlog', subject: 'Future', sensitivity: 0, categoriesCsv: '', categories: [], dueDaysFromToday: 12, dueDateMs: 400 },
    { entryID: 'x', laneId: 'backlog', subject: 'No due', sensitivity: 0, categoriesCsv: '', categories: [], dueDaysFromToday: null, dueDateMs: null }
  ];

  const lanes = board.buildLanes(tasks, config);
  const privacyFilter = mkPrivacyFilter();

  let filter = { private: privacyFilter.all.value, search: '', category: '<All Categories>', due: 'overdue' };
  board.applyFilters(lanes, filter, privacyFilter);
  assert.deepEqual(lanes[0].filteredTasks.map(t => t.entryID), ['o']);

  filter = { private: privacyFilter.all.value, search: '', category: '<All Categories>', due: 'today' };
  board.applyFilters(lanes, filter, privacyFilter);
  assert.deepEqual(lanes[0].filteredTasks.map(t => t.entryID), ['t']);

  filter = { private: privacyFilter.all.value, search: '', category: '<All Categories>', due: 'next7' };
  board.applyFilters(lanes, filter, privacyFilter);
  assert.deepEqual(lanes[0].filteredTasks.map(t => t.entryID), ['t', 'n']);

  filter = { private: privacyFilter.all.value, search: '', category: '<All Categories>', due: 'nodue' };
  board.applyFilters(lanes, filter, privacyFilter);
  assert.deepEqual(lanes[0].filteredTasks.map(t => t.entryID), ['x']);
});

test('applyFilters() filters by status (active/all/completed)', () => {
  const config = {
    BOARD: { saveOrder: false, showDoneCompletedInActiveView: false },
    LANES: [
      { id: 'doing', title: 'Doing', color: '#60a5fa', wipLimit: 0, enabled: true, outlookStatus: 1 },
      { id: 'done', title: 'Done', color: '#34d399', wipLimit: 0, enabled: true, outlookStatus: 2 }
    ]
  };

  const tasks = [
    { entryID: 'a', laneId: 'doing', subject: 'Active', sensitivity: 0, categoriesCsv: '', categories: [], statusValue: 1, dueDateMs: 100 },
    { entryID: 'b', laneId: 'doing', subject: 'Completed in Doing', sensitivity: 0, categoriesCsv: '', categories: [], statusValue: 2, dueDateMs: 200 },
    { entryID: 'c', laneId: 'done', subject: 'Completed in Done', sensitivity: 0, categoriesCsv: '', categories: [], statusValue: 2, dueDateMs: 100 },
    { entryID: 'd', laneId: 'done', subject: 'Active in Done', sensitivity: 0, categoriesCsv: '', categories: [], statusValue: 0, dueDateMs: 200 }
  ];

  const lanes = board.buildLanes(tasks, config);
  const privacyFilter = mkPrivacyFilter();

  // Active view: hide all completed tasks by default
  let filter = { private: privacyFilter.all.value, search: '', category: '<All Categories>', due: 'any', status: 'active', stale: 'any' };
  board.applyFilters(lanes, filter, privacyFilter, config);
  assert.deepEqual(lanes[0].filteredTasks.map(t => t.entryID), ['a']);
  assert.deepEqual(lanes[1].filteredTasks.map(t => t.entryID), ['d']);

  // Active view: optionally keep completed tasks visible in Done lane
  const config2 = { ...config, BOARD: { ...config.BOARD, showDoneCompletedInActiveView: true } };
  const lanes2 = board.buildLanes(tasks, config2);
  board.applyFilters(lanes2, filter, privacyFilter, config2);
  assert.deepEqual(lanes2[0].filteredTasks.map(t => t.entryID), ['a']);
  assert.deepEqual(lanes2[1].filteredTasks.map(t => t.entryID), ['c', 'd']);

  // Completed view
  filter = { private: privacyFilter.all.value, search: '', category: '<All Categories>', due: 'any', status: 'completed', stale: 'any' };
  board.applyFilters(lanes2, filter, privacyFilter, config2);
  assert.deepEqual(lanes2[0].filteredTasks.map(t => t.entryID), ['b']);
  assert.deepEqual(lanes2[1].filteredTasks.map(t => t.entryID), ['c']);
});

test('applyFilters() filters by staleness (time in lane)', () => {
  const config = {
    BOARD: { saveOrder: false },
    LANES: [{ id: 'backlog', title: 'Backlog', color: '#94a3b8', wipLimit: 0, enabled: true, outlookStatus: 0 }]
  };

  const tasks = [
    { entryID: 'n', laneId: 'backlog', subject: 'New', sensitivity: 0, categoriesCsv: '', categories: [], laneAgeDays: 2, dueDateMs: 50 },
    { entryID: 's7', laneId: 'backlog', subject: 'Stale7', sensitivity: 0, categoriesCsv: '', categories: [], laneAgeDays: 7, dueDateMs: 100 },
    { entryID: 's20', laneId: 'backlog', subject: 'Stale20', sensitivity: 0, categoriesCsv: '', categories: [], laneAgeDays: 20, dueDateMs: 200 },
    { entryID: 'u', laneId: 'backlog', subject: 'Unknown age', sensitivity: 0, categoriesCsv: '', categories: [], laneAgeDays: null, dueDateMs: null }
  ];

  const lanes = board.buildLanes(tasks, config);
  const privacyFilter = mkPrivacyFilter();

  let filter = { private: privacyFilter.all.value, search: '', category: '<All Categories>', due: 'any', status: 'all', stale: 'stale7' };
  board.applyFilters(lanes, filter, privacyFilter, config);
  assert.deepEqual(lanes[0].filteredTasks.map(t => t.entryID), ['s7', 's20']);

  filter = { private: privacyFilter.all.value, search: '', category: '<All Categories>', due: 'any', status: 'all', stale: 'stale14' };
  board.applyFilters(lanes, filter, privacyFilter, config);
  assert.deepEqual(lanes[0].filteredTasks.map(t => t.entryID), ['s20']);
});
