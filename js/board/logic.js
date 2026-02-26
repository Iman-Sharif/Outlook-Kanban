'use strict';

(function (root, factory) {
    if (typeof module === 'object' && module && module.exports) {
        module.exports = factory(require('../core/util'));
    } else {
        root.kfoBoard = factory(root.kfoUtil);
    }
})(typeof window !== 'undefined' ? window : (typeof global !== 'undefined' ? global : this), function (util) {
    function normaliseLanes(config) {
        var lanes = [];
        (config && config.LANES ? config.LANES : []).forEach(function (l) {
            var id = util.sanitizeId(l.id);
            if (!id) return;
            lanes.push({
                id: id,
                title: l.title || id,
                color: util.isValidHexColor(l.color) ? l.color : '#94a3b8',
                wipLimit: Number(l.wipLimit || 0),
                enabled: (l.enabled !== false),
                outlookStatus: (l.outlookStatus === undefined ? null : l.outlookStatus),
                tasks: [],
                filteredTasks: []
            });
        });

        // Default to at least one lane
        if (lanes.length === 0) {
            lanes.push({ id: 'backlog', title: 'Backlog', color: '#94a3b8', wipLimit: 0, enabled: true, outlookStatus: 0, tasks: [], filteredTasks: [] });
        }

        return lanes;
    }

    function sortLaneTasks(tasks, config) {
        var saveOrder = !!(config && config.BOARD && config.BOARD.saveOrder);
        tasks.sort(function (a, b) {
            if (saveOrder) {
                var ao = (a.laneOrder === undefined || a.laneOrder === null) ? 999999 : a.laneOrder;
                var bo = (b.laneOrder === undefined || b.laneOrder === null) ? 999999 : b.laneOrder;
                if (ao !== bo) return ao - bo;
            }

            // due date asc (missing due dates last)
            var ad = a.dueDateMs || 9999999999999;
            var bd = b.dueDateMs || 9999999999999;
            if (ad !== bd) return ad - bd;

            // priority desc
            if (a.priority !== b.priority) return (b.priority || 0) - (a.priority || 0);

            // subject asc
            var as = (a.subject || '').toLowerCase();
            var bs = (b.subject || '').toLowerCase();
            if (as < bs) return -1;
            if (as > bs) return 1;
            return 0;
        });
    }

    function buildLanes(tasks, config) {
        var enabledLanes = normaliseLanes(config);
        // Default unassigned tasks to the first *enabled* lane.
        // If all lanes are disabled (or missing), fall back to the first lane.
        var defaultLaneId = enabledLanes[0].id;
        for (var di = 0; di < enabledLanes.length; di++) {
            if (enabledLanes[di] && enabledLanes[di].enabled !== false) {
                defaultLaneId = enabledLanes[di].id;
                break;
            }
        }

        var defaultLane = enabledLanes[0];
        for (var dj = 0; dj < enabledLanes.length; dj++) {
            if (enabledLanes[dj] && enabledLanes[dj].id === defaultLaneId) {
                defaultLane = enabledLanes[dj];
                break;
            }
        }

        (tasks || []).forEach(function (t) {
            var laneId = util.sanitizeId(t.laneId) || defaultLaneId;
            var lane = null;
            for (var i = 0; i < enabledLanes.length; i++) {
                if (enabledLanes[i].id === laneId) {
                    lane = enabledLanes[i];
                    break;
                }
            }
            if (!lane) {
                lane = defaultLane;
            }
            lane.tasks.push(t);
        });

        enabledLanes.forEach(function (lane) {
            sortLaneTasks(lane.tasks, config);
            lane.filteredTasks = lane.tasks.slice(0);
        });

        return enabledLanes;
    }

    function isFiltersActive(filter, privacyFilter) {
        try {
            if ((String(filter.search || '').trim()) !== '') return true;
            if ((filter.category || '<All Categories>') !== '<All Categories>') return true;
            if (String(filter.private) !== String(privacyFilter.all.value)) return true;
            if (String(filter.due || 'any') !== 'any') return true;
            if (String(filter.status || 'all') !== 'all') return true;
            if (String(filter.stale || 'any') !== 'any') return true;
        } catch (e) {
            // ignore
        }
        return false;
    }

    function applyFilters(lanes, filter, privacyFilter, config) {
        // Treat whitespace-only searches as empty.
        var search = String(filter.search || '').toLowerCase();
        search = search.replace(/^\s+|\s+$/g, '');
        var category = filter.category || '<All Categories>';
        var privacy = String(filter.private);
        var due = String(filter.due || 'any');
        var status = String(filter.status || 'all');
        var stale = String(filter.stale || 'any');

        var showDoneCompletedInActiveView = !!(config && config.BOARD && config.BOARD.showDoneCompletedInActiveView);

        (lanes || []).forEach(function (lane) {
            lane.filteredTasks = (lane.tasks || []).filter(function (t) {
                // completed/active
                if (status === 'completed') {
                    if (t.statusValue !== 2) return false;
                } else if (status === 'active') {
                    if (t.statusValue === 2) {
                        if (!(showDoneCompletedInActiveView && lane && lane.outlookStatus === 2)) {
                            return false;
                        }
                    }
                }

                // privacy
                if (privacy === String(privacyFilter.private.value)) {
                    if (t.sensitivity !== 2) return false;
                }
                if (privacy === String(privacyFilter.public.value)) {
                    if (t.sensitivity === 2) return false;
                }

                // search
                if (search) {
                    var hay = ((t.subject || '') + ' ' + (t.notes || '')).toLowerCase();
                    if (hay.indexOf(search) === -1) return false;
                }

                // category
                if (category && category !== '<All Categories>') {
                    if (category === '<No Category>') {
                        if ((t.categoriesCsv || '').trim() !== '') return false;
                    } else {
                        var found = false;
                        (t.categories || []).forEach(function (c) {
                            if (c.label === category) found = true;
                        });
                        if (!found) return false;
                    }
                }

                // due date (relative to today; computed by runtime as dueDaysFromToday)
                if (due && due !== 'any') {
                    var days = (t.dueDaysFromToday === undefined) ? null : t.dueDaysFromToday;
                    if (days === '' || days === null || days === undefined) {
                        days = null;
                    } else {
                        var dn = parseInt(days, 10);
                        days = isNaN(dn) ? null : dn;
                    }

                    if (due === 'nodue') {
                        if (days !== null) return false;
                    } else if (due === 'overdue') {
                        if (days === null || days >= 0) return false;
                    } else if (due === 'today') {
                        if (days === null || days !== 0) return false;
                    } else if (due === 'next7') {
                        if (days === null || days < 0 || days > 7) return false;
                    }
                }

                // staleness (time in lane; computed by runtime as laneAgeDays)
                if (stale && stale !== 'any') {
                    var threshold = 0;
                    if (stale === 'stale7') threshold = 7;
                    else if (stale === 'stale14') threshold = 14;
                    else if (stale === 'stale30') threshold = 30;
                    else threshold = 0;

                    if (threshold > 0) {
                        var sd = (t.laneAgeDays === undefined) ? null : t.laneAgeDays;
                        if (sd === '' || sd === null || sd === undefined) {
                            sd = null;
                        } else {
                            var sn = parseInt(sd, 10);
                            sd = isNaN(sn) ? null : sn;
                        }
                        if (sd === null || sd < threshold) return false;
                    }
                }

                return true;
            });
        });

        return isFiltersActive(filter, privacyFilter);
    }

    return {
        buildLanes: buildLanes,
        applyFilters: applyFilters,
        sortLaneTasks: sortLaneTasks,
        isFiltersActive: isFiltersActive
    };
});
