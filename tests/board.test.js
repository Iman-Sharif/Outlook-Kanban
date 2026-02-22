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
