'use strict';

(function () {
    var CONFIG_ID = 'KanbanConfig';
    var STATE_ID = 'KanbanState';
    var LOG_ID = 'KanbanErrorLog';

    var SCHEMA_VERSION = 3;

    // Outlook task user properties (stored locally in Outlook)
    var PROP_LANE_ID = 'KFO_LaneId';
    var PROP_LANE_ORDER = 'KFO_LaneOrder';

    var DEFAULT_ROOT_FOLDER_NAME = 'Kanban Projects';

    var BUILTIN_THEMES = [
        { id: 'kfo-light', name: 'Professional Light', cssHref: 'themes/kfo-light/theme.css', kind: 'builtin' },
        { id: 'kfo-dark', name: 'Professional Dark', cssHref: 'themes/kfo-dark/theme.css', kind: 'builtin' }
    ];

    function nowStamp() {
        var d = new Date();
        function pad(n) { return (n < 10 ? '0' : '') + n; }
        return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '-' + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
    }

    function sanitizeId(raw) {
        if (!raw) return '';
        return String(raw)
            .toLowerCase()
            .replace(/\s+/g, '-')
            .replace(/[^a-z0-9\-]/g, '')
            .replace(/\-\-+/g, '-')
            .replace(/^\-+|\-+$/g, '');
    }

    function isValidHexColor(s) {
        return /^#[0-9a-fA-F]{6}$/.test(s || '');
    }

    function isCssLocalOnly(cssText) {
        // Best-effort guardrail: prevent accidental external loads in custom themes.
        // Users can still theme fully locally.
        try {
            var s = String(cssText || '').toLowerCase();
            if (s.indexOf('http://') !== -1) return false;
            if (s.indexOf('https://') !== -1) return false;
            if (s.indexOf('@import') !== -1) return false;
            // IE-specific scriptable CSS features
            if (s.indexOf('expression(') !== -1) return false;
            if (s.indexOf('behavior:') !== -1) return false;
            return true;
        } catch (e) {
            return false;
        }
    }

    function isSafeLocalCssPath(href) {
        // Restrict to relative paths within the install folder.
        // Example: themes/my-theme/theme.css
        try {
            var s = String(href || '').trim();
            if (!s) return false;
            if (s.indexOf('..') !== -1) return false;
            if (s.indexOf('\\') !== -1) return false;
            if (s.indexOf(':') !== -1) return false;
            if (s[0] === '/' || s[0] === '.') return false;
            if (!/\.css$/i.test(s)) return false;
            if (!/^[a-zA-Z0-9_\-\/\.]+$/.test(s)) return false;
            return true;
        } catch (e) {
            return false;
        }
    }

    function isRealDate(d) {
        try {
            if (!d) return false;
            if (isNaN(d.getTime())) return false;
            if (d.getFullYear && d.getFullYear() === 4501) return false;
            return true;
        } catch (e) {
            return false;
        }
    }

            function DEFAULT_CONFIG_V3() {
                return {
            SCHEMA_VERSION: SCHEMA_VERSION,
            SETUP: {
                completed: false
            },
            PROJECTS: {
                rootFolderName: DEFAULT_ROOT_FOLDER_NAME,
                defaultProjectEntryID: '',
                linkedProjects: [],
                hiddenProjectEntryIDs: []
            },
            UI: {
                density: 'comfortable',
                motion: 'full',
                laneWidthPx: 320,
                showDueDate: true,
                showNotes: true,
                showCategories: true,
                showOnlyFirstCategory: false,
                showPriorityPill: true,
                showPrivacyIcon: true,
                showLaneCounts: true
            },
            AUTOMATION: {
                setOutlookStatusOnLaneMove: true
            },
            LANES: [
                { id: 'backlog', title: 'Backlog', color: '#94a3b8', wipLimit: 0, enabled: true, outlookStatus: 0 },
                { id: 'doing', title: 'In Progress', color: '#60a5fa', wipLimit: 5, enabled: true, outlookStatus: 1 },
                { id: 'waiting', title: 'Waiting', color: '#fbbf24', wipLimit: 0, enabled: true, outlookStatus: 3 },
                { id: 'done', title: 'Done', color: '#34d399', wipLimit: 0, enabled: true, outlookStatus: 2 }
            ],
            THEME: {
                activeThemeId: 'kfo-light',
                folderThemes: [],
                customThemes: []
            },
            BOARD: {
                taskNoteMaxLen: 140,
                saveState: true,
                saveOrder: true
            },
            USE_CATEGORY_COLORS: true,
            USE_CATEGORY_COLOR_FOOTERS: false,
            DATE_FORMAT: 'DD-MMM',
            MULTI_MAILBOX: false,
            ACTIVE_MAILBOXES: [],
                    LOG_ERRORS: false
                };
            }

    function laneTemplate(templateId) {
        if (templateId === 'gtd') {
            return [
                { id: 'inbox', title: 'Inbox', color: '#93c5fd', wipLimit: 0, enabled: true, outlookStatus: 0 },
                { id: 'next', title: 'Next', color: '#60a5fa', wipLimit: 20, enabled: true, outlookStatus: 1 },
                { id: 'waiting', title: 'Waiting', color: '#fbbf24', wipLimit: 0, enabled: true, outlookStatus: 3 },
                { id: 'someday', title: 'Someday', color: '#a78bfa', wipLimit: 0, enabled: true, outlookStatus: 0 },
                { id: 'done', title: 'Done', color: '#34d399', wipLimit: 0, enabled: true, outlookStatus: 2 }
            ];
        }
        if (templateId === 'scrum') {
            return [
                { id: 'backlog', title: 'Backlog', color: '#94a3b8', wipLimit: 0, enabled: true, outlookStatus: 0 },
                { id: 'sprint', title: 'Sprint', color: '#60a5fa', wipLimit: 0, enabled: true, outlookStatus: 0 },
                { id: 'doing', title: 'Doing', color: '#38bdf8', wipLimit: 5, enabled: true, outlookStatus: 1 },
                { id: 'review', title: 'Review', color: '#f59e0b', wipLimit: 0, enabled: true, outlookStatus: 1 },
                { id: 'done', title: 'Done', color: '#34d399', wipLimit: 0, enabled: true, outlookStatus: 2 }
            ];
        }
        // personal (default)
        return [
            { id: 'backlog', title: 'Backlog', color: '#94a3b8', wipLimit: 0, enabled: true, outlookStatus: 0 },
            { id: 'next', title: 'Next', color: '#60a5fa', wipLimit: 20, enabled: true, outlookStatus: 0 },
            { id: 'doing', title: 'In Progress', color: '#38bdf8', wipLimit: 5, enabled: true, outlookStatus: 1 },
            { id: 'waiting', title: 'Waiting', color: '#fbbf24', wipLimit: 0, enabled: true, outlookStatus: 3 },
            { id: 'done', title: 'Done', color: '#34d399', wipLimit: 0, enabled: true, outlookStatus: 2 }
        ];
    }

    angular
        .module('taskboardApp', ['ui.sortable'])
        .controller('taskboardController', ['$scope', '$filter', '$interval', '$timeout', function ($scope, $filter, $interval, $timeout) {
            var hasReadConfig = false;
            var hasReadState = false;
            var refreshTimer;

            $scope.isBrowserSupported = false;

            $scope.version = (typeof VERSION !== 'undefined') ? VERSION : '0.0.0';

            $scope.rootClasses = {};

            $scope.ui = {
                mode: 'board',
                projectEntryID: '',
                isRefreshing: false,
                toast: { show: false, type: 'info', title: '', message: '' },
                showSetupWizard: false,
                setupStep: 1,
                setupDefaultProjectName: 'General',
                setupProjectMode: 'create',
                setupExistingProjectEntryID: '',
                setupLaneTemplate: 'personal',
                showCreateProject: false,
                showRenameProject: false,
                renameProjectEntryID: '',
                renameProjectStoreID: '',
                renameProjectName: '',
                showMoveTasks: false,
                move: {
                    fromProjectEntryID: '',
                    toProjectEntryID: '',
                    mode: 'all',
                    laneId: '',
                    running: false,
                    progress: { total: 0, done: 0, percent: 0 }
                },
                showMigration: false,
                migration: {
                    onlyUnassigned: true,
                    treatUnknownAsUnassigned: true,
                    mappingRows: [],
                    running: false,
                    progress: { total: 0, done: 0, percent: 0, updated: 0, skipped: 0, errors: 0 }
                },
                showDiagnostics: false,
                createProjectMode: 'create',
                linkProjectEntryID: '',
                newProjectName: '',
                newLaneTitle: '',
                newLaneId: '',
                newLaneColor: '#60a5fa',
                importThemeName: '',
                importThemeId: '',
                folderThemeName: '',
                folderThemeId: '',
                folderThemeHref: ''
            };

            $scope.privacyFilter = {
                all: { value: '0', text: 'All' },
                private: { value: '2', text: 'Private' },
                public: { value: '1', text: 'Not Private' }
            };

            $scope.filter = {
                private: $scope.privacyFilter.all.value,
                search: '',
                category: '<All Categories>',
                mailbox: ''
            };

            $scope.categories = ['<All Categories>', '<No Category>'];
            $scope.mailboxes = [];
            $scope.projects = [];
            $scope.projectsAll = [];
            $scope.availableProjectFolders = [];
            $scope.lanes = [];
            $scope.laneOptions = [];
            $scope.migrationLaneOptions = [];
            $scope.allThemes = [];
            $scope.diagnosticsText = '';

            var outlookCategories;

            var toastTimer;

            function showToast(type, title, message, ms) {
                try {
                    if (!$scope.ui) return;
                    $scope.ui.toast = {
                        show: true,
                        type: type || 'info',
                        title: title || '',
                        message: message || ''
                    };
                    if (toastTimer) {
                        $timeout.cancel(toastTimer);
                        toastTimer = null;
                    }
                    toastTimer = $timeout(function () {
                        if ($scope.ui && $scope.ui.toast) {
                            $scope.ui.toast.show = false;
                        }
                    }, ms || 3200);
                } catch (e) {
                    // ignore
                }
            }

            function writeLog(message) {
                try {
                    if (!$scope.config || !$scope.config.LOG_ERRORS) {
                        return;
                    }
                    var now = new Date();
                    var datetimeString = now.getFullYear() + '-' + (now.getMonth() + 1) + '-' + now.getDate() + ' ' + now.getHours() + ':' + now.getMinutes();
                    var line = datetimeString + '  ' + message;
                    var logRaw = getJournalItem(LOG_ID);
                    var log = [];
                    if (logRaw !== null) {
                        try { log = JSON.parse(logRaw); } catch (e) { log = []; }
                    }
                    log.unshift(line);
                    if (log.length > 800) {
                        log.pop();
                    }
                    saveJournalItem(LOG_ID, JSON.stringify(log, null, 2));
                } catch (e) {
                    // keep silent
                }
            }

            function backupLegacyConfig(raw) {
                try {
                    var subject = CONFIG_ID + '.legacy.' + nowStamp();
                    saveJournalItem(subject, String(raw || ''));
                } catch (e) {
                    // ignore
                }
            }

            function rebuildThemeList() {
                var list = [];
                BUILTIN_THEMES.forEach(function (t) {
                    list.push({ id: t.id, name: t.name, cssHref: t.cssHref, kind: t.kind });
                });

                if ($scope.config && $scope.config.THEME) {
                    ($scope.config.THEME.folderThemes || []).forEach(function (t) {
                        list.push({ id: t.id, name: t.name, cssHref: t.cssHref, kind: 'folder' });
                    });
                    ($scope.config.THEME.customThemes || []).forEach(function (t) {
                        list.push({ id: t.id, name: t.name, cssText: t.cssText, kind: 'imported' });
                    });
                }

                // unique by id
                var seen = {};
                var uniq = [];
                list.forEach(function (t) {
                    if (!t.id) return;
                    if (seen[t.id]) return;
                    seen[t.id] = true;
                    uniq.push(t);
                });
                $scope.allThemes = uniq;
            }

            function findThemeById(id) {
                for (var i = 0; i < $scope.allThemes.length; i++) {
                    if ($scope.allThemes[i].id === id) {
                        return $scope.allThemes[i];
                    }
                }
                return null;
            }

            function ensureThemeStyleElement() {
                var el = document.getElementById('kfo-theme-style');
                if (!el) {
                    el = document.createElement('style');
                    el.type = 'text/css';
                    el.id = 'kfo-theme-style';
                    document.getElementsByTagName('head')[0].appendChild(el);
                }
                return el;
            }

            $scope.applyRootClasses = function () {
                try {
                    var themeId = ($scope.config && $scope.config.THEME) ? $scope.config.THEME.activeThemeId : 'kfo-light';
                    var density = ($scope.config && $scope.config.UI) ? ($scope.config.UI.density || 'comfortable') : 'comfortable';
                    var motion = ($scope.config && $scope.config.UI) ? ($scope.config.UI.motion || 'full') : 'full';

                    var classes = {};
                    classes['theme-' + themeId] = true;
                    classes['density-' + density] = true;
                    classes['motion-' + motion] = true;
                    $scope.rootClasses = classes;
                } catch (e) {
                    $scope.rootClasses = {};
                }
            };

            $scope.applyTheme = function () {
                try {
                    rebuildThemeList();

                    var themeId = ($scope.config && $scope.config.THEME) ? $scope.config.THEME.activeThemeId : 'kfo-light';
                    var theme = findThemeById(themeId);
                    if (!theme) {
                        themeId = 'kfo-light';
                        if ($scope.config && $scope.config.THEME) {
                            $scope.config.THEME.activeThemeId = themeId;
                        }
                        theme = findThemeById(themeId);
                    }

                    // Apply root classes (theme + UI)
                    $scope.applyRootClasses();

                    // Apply theme CSS link (fallback to builtin light)
                    var themeLink = document.getElementById('kfo-theme-link');
                    if (themeLink) {
                        if (theme && theme.cssHref) {
                            themeLink.href = theme.cssHref;
                        } else {
                            themeLink.href = 'themes/kfo-light/theme.css';
                        }
                    }

                    // Apply imported theme css (optional)
                    var styleEl = ensureThemeStyleElement();
                    if (theme && theme.cssText) {
                        styleEl.styleSheet ? (styleEl.styleSheet.cssText = theme.cssText) : (styleEl.innerHTML = theme.cssText);
                    } else {
                        styleEl.styleSheet ? (styleEl.styleSheet.cssText = '') : (styleEl.innerHTML = '');
                    }

                    // Persist theme selection
                    if ($scope.config && $scope.config.THEME) {
                        saveConfig();
                    }
                } catch (error) {
                    writeLog('applyTheme: ' + error);
                }
            };

            function readConfig() {
                if (hasReadConfig) return;
                try {
                    var raw = getJournalItem(CONFIG_ID);
                    if (raw === null) {
                        $scope.config = DEFAULT_CONFIG_V3();
                        saveConfig();
                        hasReadConfig = true;
                        return;
                    }
                    try {
                        $scope.config = JSON.parse(JSON.minify(raw));
                    } catch (e) {
                        alert('Configuration JSON is invalid. A new configuration will be created.');
                        backupLegacyConfig(raw);
                        $scope.config = DEFAULT_CONFIG_V3();
                        saveConfig();
                        hasReadConfig = true;
                        return;
                    }

                    if (!$scope.config.SCHEMA_VERSION || $scope.config.SCHEMA_VERSION < SCHEMA_VERSION) {
                        backupLegacyConfig(raw);
                        $scope.config = DEFAULT_CONFIG_V3();
                        saveConfig();
                    }

                    // Ensure required keys exist
                    if (!$scope.config.PROJECTS) $scope.config.PROJECTS = DEFAULT_CONFIG_V3().PROJECTS;
                    if (!$scope.config.PROJECTS.linkedProjects) $scope.config.PROJECTS.linkedProjects = [];
                    if (!$scope.config.PROJECTS.hiddenProjectEntryIDs) $scope.config.PROJECTS.hiddenProjectEntryIDs = [];
                    if (!$scope.config.UI) $scope.config.UI = DEFAULT_CONFIG_V3().UI;
                    if (!$scope.config.AUTOMATION) $scope.config.AUTOMATION = DEFAULT_CONFIG_V3().AUTOMATION;
                    if ($scope.config.UI.density === undefined) $scope.config.UI.density = DEFAULT_CONFIG_V3().UI.density;
                    if ($scope.config.UI.motion === undefined) $scope.config.UI.motion = DEFAULT_CONFIG_V3().UI.motion;
                    if ($scope.config.UI.laneWidthPx === undefined) $scope.config.UI.laneWidthPx = DEFAULT_CONFIG_V3().UI.laneWidthPx;
                    if ($scope.config.UI.showDueDate === undefined) $scope.config.UI.showDueDate = DEFAULT_CONFIG_V3().UI.showDueDate;
                    if ($scope.config.UI.showNotes === undefined) $scope.config.UI.showNotes = DEFAULT_CONFIG_V3().UI.showNotes;
                    if ($scope.config.UI.showCategories === undefined) $scope.config.UI.showCategories = DEFAULT_CONFIG_V3().UI.showCategories;
                    if ($scope.config.UI.showOnlyFirstCategory === undefined) $scope.config.UI.showOnlyFirstCategory = DEFAULT_CONFIG_V3().UI.showOnlyFirstCategory;
                    if ($scope.config.UI.showPriorityPill === undefined) $scope.config.UI.showPriorityPill = DEFAULT_CONFIG_V3().UI.showPriorityPill;
                    if ($scope.config.UI.showPrivacyIcon === undefined) $scope.config.UI.showPrivacyIcon = DEFAULT_CONFIG_V3().UI.showPrivacyIcon;
                    if ($scope.config.UI.showLaneCounts === undefined) $scope.config.UI.showLaneCounts = DEFAULT_CONFIG_V3().UI.showLaneCounts;
                    if ($scope.config.AUTOMATION.setOutlookStatusOnLaneMove === undefined) {
                        $scope.config.AUTOMATION.setOutlookStatusOnLaneMove = DEFAULT_CONFIG_V3().AUTOMATION.setOutlookStatusOnLaneMove;
                    }
                    if (!$scope.config.LANES) $scope.config.LANES = DEFAULT_CONFIG_V3().LANES;
                    if (!$scope.config.THEME) $scope.config.THEME = DEFAULT_CONFIG_V3().THEME;
                    if (!$scope.config.BOARD) $scope.config.BOARD = DEFAULT_CONFIG_V3().BOARD;
                    if ($scope.config.USE_CATEGORY_COLORS === undefined) $scope.config.USE_CATEGORY_COLORS = true;
                    if ($scope.config.USE_CATEGORY_COLOR_FOOTERS === undefined) $scope.config.USE_CATEGORY_COLOR_FOOTERS = false;
                    if (!$scope.config.DATE_FORMAT) $scope.config.DATE_FORMAT = 'DD-MMM';
                    if ($scope.config.LOG_ERRORS === undefined) $scope.config.LOG_ERRORS = false;

                    // Clamp lane width
                    try {
                        var w = parseInt($scope.config.UI.laneWidthPx, 10);
                        if (isNaN(w)) w = DEFAULT_CONFIG_V3().UI.laneWidthPx;
                        if (w < 240) w = 240;
                        if (w > 520) w = 520;
                        $scope.config.UI.laneWidthPx = w;
                    } catch (e) {
                        $scope.config.UI.laneWidthPx = DEFAULT_CONFIG_V3().UI.laneWidthPx;
                    }

                } catch (error) {
                    alert('Failed to read configuration: ' + error);
                    $scope.config = DEFAULT_CONFIG_V3();
                    saveConfig();
                }
                hasReadConfig = true;
            }

            function saveConfig() {
                try {
                    // Clamp UI settings (defensive)
                    if ($scope.config && $scope.config.UI) {
                        var w = parseInt($scope.config.UI.laneWidthPx, 10);
                        if (isNaN(w)) w = DEFAULT_CONFIG_V3().UI.laneWidthPx;
                        if (w < 240) w = 240;
                        if (w > 520) w = 520;
                        $scope.config.UI.laneWidthPx = w;

                        var density = String($scope.config.UI.density || 'comfortable');
                        if (density !== 'compact' && density !== 'comfortable') density = 'comfortable';
                        $scope.config.UI.density = density;

                        var motion = String($scope.config.UI.motion || 'full');
                        if (motion !== 'full' && motion !== 'subtle' && motion !== 'off') motion = 'full';
                        $scope.config.UI.motion = motion;
                    }
                    saveJournalItem(CONFIG_ID, JSON.stringify($scope.config, null, 2));
                } catch (e) {
                    alert('Failed to save configuration: ' + e);
                }
            }

            function readState() {
                if (hasReadState) return;
                try {
                    if (!$scope.config || !$scope.config.BOARD || !$scope.config.BOARD.saveState) {
                        hasReadState = true;
                        return;
                    }
                    var state = {
                        private: $scope.privacyFilter.all.value,
                        search: '',
                        category: '<All Categories>',
                        mailbox: '',
                        projectEntryID: ''
                    };
                    var raw = getJournalItem(STATE_ID);
                    if (raw !== null) {
                        try { state = JSON.parse(raw); } catch (e) { /* ignore */ }
                    } else {
                        saveJournalItem(STATE_ID, JSON.stringify(state, null, 2));
                    }

                    $scope.filter.private = state.private || $scope.privacyFilter.all.value;
                    $scope.filter.search = state.search || '';
                    $scope.filter.category = state.category || '<All Categories>';
                    $scope.filter.mailbox = state.mailbox || '';
                    $scope.ui.projectEntryID = state.projectEntryID || '';
                } catch (e) {
                    writeLog('readState: ' + e);
                }
                hasReadState = true;
            }

            function saveState() {
                try {
                    if (!$scope.config || !$scope.config.BOARD || !$scope.config.BOARD.saveState) {
                        return;
                    }
                    var state = {
                        private: $scope.filter.private,
                        search: $scope.filter.search,
                        category: $scope.filter.category,
                        mailbox: $scope.filter.mailbox,
                        projectEntryID: $scope.ui.projectEntryID
                    };
                    saveJournalItem(STATE_ID, JSON.stringify(state, null, 2));
                } catch (e) {
                    writeLog('saveState: ' + e);
                }
            }

            function initCategories() {
                $scope.categories = ['<All Categories>', '<No Category>'];
                try {
                    outlookCategories = getOutlookCategories();
                    outlookCategories.names.forEach(function (name) {
                        $scope.categories.push(name);
                    });
                    $scope.categories = $scope.categories.sort();
                } catch (e) {
                    writeLog('initCategories: ' + e);
                }
            }

            function initMailboxes() {
                $scope.mailboxes = [];
                try {
                    var mb = getOutlookMailboxes(!!($scope.config && $scope.config.MULTI_MAILBOX));
                    mb.forEach(function (m) {
                        $scope.mailboxes.push(m);
                    });
                    if (!$scope.filter.mailbox) {
                        $scope.filter.mailbox = $scope.mailboxes[0];
                    }
                } catch (e) {
                    writeLog('initMailboxes: ' + e);
                }
            }

            function loadAvailableProjectFolders() {
                try {
                    var list = [];
                    // Include default Tasks folder
                    try {
                        var tasksFolder = getTaskFolderExisting($scope.filter.mailbox, '');
                        if (tasksFolder) {
                            list.push({
                                name: tasksFolder.Name + ' (Tasks)',
                                entryID: tasksFolder.EntryID,
                                storeID: tasksFolder.StoreID
                            });
                        }
                    } catch (e1) {
                        // ignore
                    }

                    // Include subfolders under Tasks
                    try {
                        var subs = listTaskSubFolders($scope.filter.mailbox, '');
                        subs.forEach(function (f) {
                            list.push(f);
                        });
                    } catch (e2) {
                        // ignore
                    }

                    // Unique by entryID
                    var seen = {};
                    var uniq = [];
                    list.forEach(function (p) {
                        if (!p || !p.entryID) return;
                        if (seen[p.entryID]) return;
                        seen[p.entryID] = true;
                        uniq.push(p);
                    });
                    uniq.sort(function (a, b) {
                        var an = (a.name || '').toLowerCase();
                        var bn = (b.name || '').toLowerCase();
                        if (an < bn) return -1;
                        if (an > bn) return 1;
                        return 0;
                    });
                    $scope.availableProjectFolders = uniq;
                } catch (e) {
                    writeLog('loadAvailableProjectFolders: ' + e);
                    $scope.availableProjectFolders = [];
                }
            }

            function getProjectsRootFolderExisting() {
                try {
                    return getTaskFolderExisting($scope.filter.mailbox, $scope.config.PROJECTS.rootFolderName);
                } catch (e) {
                    return null;
                }
            }

            function loadProjects() {
                try {
                    var projects = [];
                    var hidden = ($scope.config.PROJECTS.hiddenProjectEntryIDs || []);

                    var defaultTasksEntryID = '';
                    try {
                        var tf = getTaskFolderExisting($scope.filter.mailbox, '');
                        if (tf) {
                            defaultTasksEntryID = tf.EntryID;
                        }
                    } catch (e0) {
                        defaultTasksEntryID = '';
                    }

                    // Root subfolders
                    var root = getProjectsRootFolderExisting();
                    if (root) {
                        var subs = listTaskSubFolders($scope.filter.mailbox, $scope.config.PROJECTS.rootFolderName);
                        subs.forEach(function (p) {
                            projects.push({
                                name: p.name,
                                entryID: p.entryID,
                                storeID: p.storeID,
                                isLinked: false
                            });
                        });
                    }

                    // Linked projects (resolve current name when possible)
                    var linked = ($scope.config.PROJECTS.linkedProjects || []);
                    linked.forEach(function (p) {
                        if (!p || !p.entryID) return;
                        var name = p.name || 'Linked project';
                        var storeID = p.storeID;
                        try {
                            var f = getFolderFromIDs(p.entryID, p.storeID);
                            if (f) {
                                name = f.Name;
                                try { storeID = f.StoreID; } catch (e1) { /* ignore */ }
                            }
                        } catch (e2) {
                            // ignore
                        }
                        projects.push({
                            name: name,
                            entryID: p.entryID,
                            storeID: storeID,
                            isLinked: true
                        });
                    });

                    // Unique by entryID
                    var seen = {};
                    var uniq = [];
                    projects.forEach(function (p) {
                        if (!p.entryID) return;
                        if (seen[p.entryID]) return;
                        seen[p.entryID] = true;
                        uniq.push(p);
                    });

                    // Mark hidden/default
                    uniq.forEach(function (p) {
                        p.isHidden = (hidden.indexOf(p.entryID) !== -1);
                        p.isDefaultTasks = (defaultTasksEntryID && p.entryID === defaultTasksEntryID);
                    });

                    // Sort by name
                    uniq.sort(function (a, b) {
                        var an = (a.name || '').toLowerCase();
                        var bn = (b.name || '').toLowerCase();
                        if (an < bn) return -1;
                        if (an > bn) return 1;
                        return 0;
                    });

                    $scope.projectsAll = uniq;
                    $scope.projects = uniq.filter(function (p) { return !p.isHidden; });

                    // Set default project if missing/invalid/hidden
                    var defaultId = $scope.config.PROJECTS.defaultProjectEntryID;
                    var defaultOk = false;
                    if (defaultId) {
                        for (var i = 0; i < $scope.projects.length; i++) {
                            if ($scope.projects[i].entryID === defaultId) defaultOk = true;
                        }
                    }
                    if (!defaultOk) {
                        if ($scope.projects.length > 0) {
                            $scope.config.PROJECTS.defaultProjectEntryID = $scope.projects[0].entryID;
                            saveConfig();
                        }
                    }
                } catch (e) {
                    writeLog('loadProjects: ' + e);
                }
            }

            function ensureSelectedProject() {
                if ($scope.projects.length === 0) {
                    $scope.ui.projectEntryID = '';
                    return;
                }

                function exists(entryID) {
                    for (var i = 0; i < $scope.projects.length; i++) {
                        if ($scope.projects[i].entryID === entryID) return true;
                    }
                    return false;
                }

                if ($scope.ui.projectEntryID && exists($scope.ui.projectEntryID)) {
                    return;
                }

                if ($scope.config.PROJECTS.defaultProjectEntryID && exists($scope.config.PROJECTS.defaultProjectEntryID)) {
                    $scope.ui.projectEntryID = $scope.config.PROJECTS.defaultProjectEntryID;
                    return;
                }

                $scope.ui.projectEntryID = $scope.projects[0].entryID;
            }

            function getSelectedProject() {
                for (var i = 0; i < $scope.projects.length; i++) {
                    if ($scope.projects[i].entryID === $scope.ui.projectEntryID) {
                        return $scope.projects[i];
                    }
                }
                return null;
            }

            function getSelectedProjectFolder() {
                try {
                    var p = getSelectedProject();
                    if (!p) return null;
                    if (p.entryID) {
                        var f = getFolderFromIDs(p.entryID, p.storeID);
                        if (f) return f;
                    }
                } catch (e) {
                    writeLog('getSelectedProjectFolder: ' + e);
                }
                return null;
            }

            function taskBodyNotes(str, limit) {
                try {
                    if (!str) return '';
                    var s = String(str);
                    s = s.replace(/^(?=\n)$|^\s*|\s*$|\n\n+/gm, '');
                    s = s.replace(/\r\n/g, '\n');
                    if (limit && s.length > limit) {
                        s = s.substring(0, limit);
                        // trim to last whitespace
                        var i = s.lastIndexOf(' ');
                        if (i > 40) s = s.substring(0, i);
                        s = s + '...';
                    }
                    return s;
                } catch (e) {
                    return '';
                }
            }

            function taskStatusText(status) {
                // Built-in Outlook task statuses
                if (status === 0) return 'Not Started';
                if (status === 1) return 'In Progress';
                if (status === 2) return 'Completed';
                if (status === 3) return 'Waiting For Someone Else';
                if (status === 4) return 'Deferred';
                return '';
            }

            function getContrastYIQ(hexcolor) {
                try {
                    if (!hexcolor) return 'black';
                    var r = parseInt(hexcolor.substr(1, 2), 16);
                    var g = parseInt(hexcolor.substr(3, 2), 16);
                    var b = parseInt(hexcolor.substr(5, 2), 16);
                    var yiq = ((r * 299) + (g * 587) + (b * 114)) / 1000;
                    return (yiq >= 128) ? 'black' : 'white';
                } catch (e) {
                    return 'black';
                }
            }

            function getCategoryStyles(csvCategories) {
                var colorArray = [
                    '#E7A1A2', '#F9BA89', '#F7DD8F', '#FCFA90', '#78D168', '#9FDCC9', '#C6D2B0', '#9DB7E8', '#B5A1E2',
                    '#daaec2', '#dad9dc', '#6b7994', '#bfbfbf', '#6f6f6f', '#4f4f4f', '#c11a25', '#e2620d', '#c79930',
                    '#b9b300', '#368f2b', '#329b7a', '#778b45', '#2858a5', '#5c3fa3', '#93446b'
                ];

                function getColor(category) {
                    try {
                        if (!outlookCategories || !outlookCategories.names) return '#4f4f4f';
                        var c = outlookCategories.names.indexOf(category);
                        var i = outlookCategories.colors[c];
                        if (i === -1 || i === undefined) {
                            return '#4f4f4f';
                        }
                        return colorArray[i - 1];
                    } catch (e) {
                        return '#4f4f4f';
                    }
                }

                try {
                    var catStyles = [];
                    var categories = String(csvCategories || '').split(/[;,]+/);
                    for (var i = 0; i < categories.length; i++) {
                        var c = categories[i].trim();
                        if (c.length === 0) continue;
                        if ($scope.config.USE_CATEGORY_COLORS) {
                            var bg = getColor(c);
                            catStyles.push({
                                label: c,
                                style: {
                                    'background-color': bg,
                                    color: getContrastYIQ(bg)
                                }
                            });
                        } else {
                            catStyles.push({ label: c, style: { color: 'inherit' } });
                        }
                    }
                    return catStyles;
                } catch (e) {
                    return [];
                }
            }

            $scope.getFooterStyle = function (categories) {
                try {
                    if ($scope.config.USE_CATEGORY_COLOR_FOOTERS && $scope.config.USE_CATEGORY_COLORS) {
                        if (categories && categories.length === 1 && categories[0] && categories[0].style) {
                            return categories[0].style;
                        }
                        if (categories && categories.length > 1) {
                            var lightGray = '#dfdfdf';
                            return { 'background-color': lightGray, color: getContrastYIQ(lightGray) };
                        }
                    }
                } catch (e) {
                    // ignore
                }
                return;
            };

            function buildLanes(tasks) {
                var lanes = [];
                var enabledLanes = [];
                ($scope.config.LANES || []).forEach(function (l) {
                    var id = sanitizeId(l.id);
                    if (!id) return;
                    enabledLanes.push({
                        id: id,
                        title: l.title || id,
                        color: isValidHexColor(l.color) ? l.color : '#94a3b8',
                        wipLimit: Number(l.wipLimit || 0),
                        enabled: (l.enabled !== false),
                        outlookStatus: (l.outlookStatus === undefined ? null : l.outlookStatus),
                        tasks: [],
                        filteredTasks: []
                    });
                });

                // Default to at least one lane
                if (enabledLanes.length === 0) {
                    enabledLanes.push({ id: 'backlog', title: 'Backlog', color: '#94a3b8', wipLimit: 0, enabled: true, outlookStatus: 0, tasks: [], filteredTasks: [] });
                }

                var defaultLaneId = enabledLanes[0].id;

                tasks.forEach(function (t) {
                    var laneId = sanitizeId(t.laneId) || defaultLaneId;
                    var lane = null;
                    for (var i = 0; i < enabledLanes.length; i++) {
                        if (enabledLanes[i].id === laneId) {
                            lane = enabledLanes[i];
                            break;
                        }
                    }
                    if (!lane) {
                        lane = enabledLanes[0];
                    }
                    lane.tasks.push(t);
                });

                // Sorting
                enabledLanes.forEach(function (lane) {
                    lane.tasks.sort(function (a, b) {
                        if ($scope.config.BOARD.saveOrder) {
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
                    lane.filteredTasks = lane.tasks.slice(0);
                });

                lanes = enabledLanes;
                return lanes;
            }

            function readTasksFromOutlookFolder(folder) {
                var tasks = [];
                try {
                    var folderStoreID = '';
                    try { folderStoreID = folder.StoreID; } catch (e0) { folderStoreID = ''; }

                    var today0 = new Date();
                    today0.setHours(0, 0, 0, 0);

                    var items = folder.Items;
                    var count = items.Count;
                    for (var i = 1; i <= count; i++) {
                        var task = items(i);
                        try {
                            var due = new Date(task.DueDate);
                            var dueText = '';
                            var dueMs = null;
                            var dueClass = '';
                            if (isRealDate(due)) {
                                var due0 = new Date(due);
                                due0.setHours(0, 0, 0, 0);
                                dueText = moment(due).format($scope.config.DATE_FORMAT || 'DD-MMM');
                                dueMs = due.getTime();

                                // Due-state color (ignore completed tasks)
                                if (task.Status !== 2) {
                                    if (due0.getTime() < today0.getTime()) {
                                        dueClass = 'kfo-due--overdue';
                                    } else if (due0.getTime() === today0.getTime()) {
                                        dueClass = 'kfo-due--today';
                                    } else {
                                        var days = Math.round((due0.getTime() - today0.getTime()) / (24 * 60 * 60 * 1000));
                                        if (days <= 2) {
                                            dueClass = 'kfo-due--soon';
                                        }
                                    }
                                }
                            }

                            var laneId = '';
                            try { laneId = getUserProperty(task, PROP_LANE_ID); } catch (e1) { laneId = ''; }
                            var laneOrderRaw = '';
                            try { laneOrderRaw = getUserProperty(task, PROP_LANE_ORDER); } catch (e2) { laneOrderRaw = ''; }
                            var laneOrder = null;
                            if (laneOrderRaw !== '' && laneOrderRaw !== null && laneOrderRaw !== undefined) {
                                var n = parseInt(laneOrderRaw, 10);
                                if (!isNaN(n)) laneOrder = n;
                            }

                            tasks.push({
                                entryID: task.EntryID,
                                storeID: folderStoreID,
                                subject: task.Subject,
                                priority: task.Importance,
                                sensitivity: task.Sensitivity,
                                statusValue: task.Status,
                                statusText: taskStatusText(task.Status),
                                dueText: dueText,
                                dueDateMs: dueMs,
                                dueClass: dueClass,
                                categoriesCsv: task.Categories,
                                categories: getCategoryStyles(task.Categories),
                                notes: taskBodyNotes(task.Body, $scope.config.BOARD.taskNoteMaxLen),
                                oneNoteURL: (function () { try { return getUserProperty(task, 'OneNoteURL'); } catch (e3) { return ''; } })(),
                                laneId: laneId,
                                laneOrder: laneOrder
                            });
                        } catch (inner) {
                            writeLog('read task: ' + inner);
                        }
                    }
                } catch (e) {
                    writeLog('readTasksFromOutlookFolder: ' + e);
                }
                return tasks;
            }

            $scope.laneHeaderStyle = function (lane) {
                return {
                    'border-top-color': lane.color || '#94a3b8'
                };
            };

            $scope.taskCardClasses = function (task) {
                return {
                    'kfo-task--private': task && task.sensitivity === 2,
                    'kfo-task--high': task && task.priority === 2
                };
            };

            $scope.visibleCategories = function (categories) {
                try {
                    if (!$scope.config || !$scope.config.UI || !$scope.config.UI.showCategories) return [];
                    var arr = categories || [];
                    if (!$scope.config.UI.showOnlyFirstCategory) return arr;
                    if (arr.length > 0) return [arr[0]];
                    return [];
                } catch (e) {
                    return categories || [];
                }
            };

            $scope.laneContainerStyle = function (lane, laneIndex) {
                var style = {};
                try {
                    if ($scope.config && $scope.config.UI) {
                        style.width = String($scope.config.UI.laneWidthPx || 320) + 'px';
                    }

                    var motion = ($scope.config && $scope.config.UI) ? ($scope.config.UI.motion || 'full') : 'full';
                    if (motion !== 'off') {
                        var base = (motion === 'subtle') ? 40 : 60;
                        var idx = laneIndex || 0;
                        if (idx > 8) idx = 8;
                        style.animationDelay = String(idx * base) + 'ms';
                    }
                } catch (e) {
                    // ignore
                }
                return style;
            };

            $scope.taskItemStyle = function (laneIndex, taskIndex) {
                var style = {};
                try {
                    var motion = ($scope.config && $scope.config.UI) ? ($scope.config.UI.motion || 'full') : 'full';
                    if (motion === 'off') return style;

                    var laneBase = (motion === 'subtle') ? 30 : 45;
                    var taskBase = (motion === 'subtle') ? 10 : 16;
                    var li = laneIndex || 0;
                    var ti = taskIndex || 0;
                    if (li > 8) li = 8;
                    if (ti > 12) ti = 12;
                    style.animationDelay = String((li * laneBase) + (ti * taskBase)) + 'ms';
                } catch (e) {
                    // ignore
                }
                return style;
            };

            $scope.applyFilters = function () {
                try {
                    var search = ($scope.filter.search || '').toLowerCase();
                    var category = $scope.filter.category || '<All Categories>';
                    var privacy = $scope.filter.private;

                    $scope.lanes.forEach(function (lane) {
                        lane.filteredTasks = lane.tasks.filter(function (t) {
                            // privacy
                            if (privacy === $scope.privacyFilter.private.value) {
                                if (t.sensitivity !== 2) return false;
                            }
                            if (privacy === $scope.privacyFilter.public.value) {
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

                            return true;
                        });
                    });

                    // To avoid persisting partial ordering, disable drag/drop while filters are active.
                    if ($scope.sortableOptions) {
                        var filtersActive = false;
                        if (($scope.filter.search || '').trim() !== '') filtersActive = true;
                        if (($scope.filter.category || '<All Categories>') !== '<All Categories>') filtersActive = true;
                        if (String($scope.filter.private) !== String($scope.privacyFilter.all.value)) filtersActive = true;
                        $scope.sortableOptions.disabled = filtersActive;
                    }

                    saveState();
                } catch (e) {
                    writeLog('applyFilters: ' + e);
                }
            };

            function doRefreshTasks() {
                try {
                    if (!$scope.ui.projectEntryID) {
                        loadProjects();
                        ensureSelectedProject();
                    }
                    var folder = getSelectedProjectFolder();
                    if (!folder) {
                        $scope.lanes = buildLanes([]);
                        $scope.applyFilters();
                        return;
                    }
                    var tasks = readTasksFromOutlookFolder(folder);
                    $scope.lanes = buildLanes(tasks);
                    $scope.applyFilters();
                } catch (e) {
                    writeLog('refreshTasks: ' + e);
                }
            }

            $scope.refreshTasks = function () {
                try {
                    if ($scope.ui.isRefreshing) {
                        return;
                    }
                    $scope.ui.isRefreshing = true;
                    $timeout(function () {
                        try {
                            doRefreshTasks();
                        } finally {
                            $scope.ui.isRefreshing = false;
                        }
                    }, 0);
                } catch (e) {
                    $scope.ui.isRefreshing = false;
                    writeLog('refreshTasks: ' + e);
                }
            };

            $scope.onMailboxChanged = function () {
                try {
                    loadAvailableProjectFolders();
                    loadProjects();
                    ensureSelectedProject();
                    saveState();
                    $scope.refreshTasks();
                } catch (e) {
                    writeLog('onMailboxChanged: ' + e);
                }
            };

            $scope.onProjectChanged = function () {
                saveState();
                $scope.refreshTasks();
            };

            $scope.switchMode = function (mode) {
                $scope.ui.mode = mode;
                if (mode === 'board') {
                    $scope.applyFilters();
                }
            };

            function getLaneById(laneId) {
                for (var i = 0; i < $scope.lanes.length; i++) {
                    if ($scope.lanes[i].id === laneId) return $scope.lanes[i];
                }
                return null;
            }

            function getTaskItemSafe(entryID, storeID) {
                try {
                    if (typeof getTaskItemFromIDs === 'function') {
                        return getTaskItemFromIDs(entryID, storeID);
                    }
                } catch (e) {
                    // ignore
                }
                return getTaskItem(entryID);
            }

            function setTaskLane(taskEntryID, storeID, laneId) {
                var taskitem = getTaskItemSafe(taskEntryID, storeID);
                setUserProperty(taskitem, PROP_LANE_ID, laneId);
                taskitem.Save();
            }

            function maybeSetTaskOutlookStatus(taskEntryID, storeID, statusValue) {
                try {
                    if (statusValue === null || statusValue === undefined) return;
                    var taskitem = getTaskItemSafe(taskEntryID, storeID);
                    if (taskitem.Status != statusValue) {
                        taskitem.Status = statusValue;
                        taskitem.Save();
                    }
                } catch (e) {
                    writeLog('maybeSetTaskOutlookStatus: ' + e);
                }
            }

            function fixLaneOrder(lane) {
                try {
                    if (!$scope.config.BOARD.saveOrder) return;
                    for (var i = 0; i < lane.filteredTasks.length; i++) {
                        var t = lane.filteredTasks[i];
                        var taskitem = getTaskItemSafe(t.entryID, t.storeID);
                        setUserProperty(taskitem, PROP_LANE_ORDER, i, OlUserPropertyType.olNumber);
                        taskitem.Save();
                    }
                } catch (e) {
                    writeLog('fixLaneOrder: ' + e);
                }
            }

            // Drag-and-drop support
            $scope.sortableOptions = {
                connectWith: '.kfo-tasklist',
                items: 'li',
                opacity: 0.6,
                cursor: 'move',
                containment: 'document',
                distance: 6,
                placeholder: 'kfo-sort-placeholder',
                forcePlaceholderSize: true,
                stop: function (e, ui) {
                    try {
                        if (!ui.item.sortable || !ui.item.sortable.droptarget) {
                            return;
                        }
                        var fromLaneId = ui.item.sortable.source.attr('data-lane-id');
                        var toLaneId = ui.item.sortable.droptarget.attr('data-lane-id');
                        if (!fromLaneId || !toLaneId) {
                            return;
                        }

                        var toLane = getLaneById(toLaneId);
                        if (!toLane) {
                            return;
                        }

                        // WIP limit guard
                        if (toLane.wipLimit && toLane.wipLimit > 0 && toLane.filteredTasks.length > toLane.wipLimit) {
                            alert('WIP limit reached for this lane.');
                            ui.item.sortable.cancel();
                            return;
                        }

                        var model = ui.item.sortable.model;
                        if (!model || !model.entryID) {
                            return;
                        }

                        // Update lane assignment
                        if (fromLaneId !== toLaneId) {
                            setTaskLane(model.entryID, model.storeID, toLaneId);
                            if ($scope.config && $scope.config.AUTOMATION && $scope.config.AUTOMATION.setOutlookStatusOnLaneMove) {
                                maybeSetTaskOutlookStatus(model.entryID, model.storeID, toLane.outlookStatus);
                            }
                        }

                        // Persist order
                        fixLaneOrder(toLane);
                        var fromLane = getLaneById(fromLaneId);
                        if (fromLane && fromLaneId !== toLaneId) {
                            fixLaneOrder(fromLane);
                        }

                        // Resync from Outlook for correctness
                        $scope.refreshTasks();
                    } catch (error) {
                        writeLog('drag/drop: ' + error);
                    }
                }
            };

            $scope.addTask = function (lane) {
                try {
                    var folder = getSelectedProjectFolder();
                    if (!folder) {
                        alert('Please create/select a project first.');
                        return;
                    }
                    var taskitem = folder.Items.Add();

                    // Default sensitivity based on current filter
                    if ($scope.filter.private == $scope.privacyFilter.private.value) {
                        taskitem.Sensitivity = SENSITIVITY.olPrivate;
                    }

                    if (lane && lane.id) {
                        setUserProperty(taskitem, PROP_LANE_ID, lane.id, OlUserPropertyType.olText);
                        setUserProperty(taskitem, PROP_LANE_ORDER, 0, OlUserPropertyType.olNumber);
                        if ($scope.config && $scope.config.AUTOMATION && $scope.config.AUTOMATION.setOutlookStatusOnLaneMove) {
                            if (lane.outlookStatus !== null && lane.outlookStatus !== undefined) {
                                taskitem.Status = lane.outlookStatus;
                            }
                        }
                    }
                    taskitem.Save();
                    taskitem.Display();

                    // Refresh after save
                    try {
                        eval('function taskitem::Write (bStat) {window.location.reload(); return true;}');
                    } catch (e) {
                        // ignore
                    }
                } catch (e) {
                    writeLog('addTask: ' + e);
                }
            };

            $scope.editTask = function (task) {
                try {
                    var taskitem = getTaskItemSafe(task.entryID, task.storeID);
                    taskitem.Display();
                    try {
                        eval('function taskitem::Write (bStat) {window.location.reload(); return true;}');
                        eval('function taskitem::BeforeDelete (bStat) {window.location.reload(); return true;}');
                    } catch (e) {
                        // ignore
                    }
                } catch (e) {
                    writeLog('editTask: ' + e);
                }
            };

            $scope.deleteTask = function (task, askConfirmation) {
                try {
                    var ok = true;
                    if (askConfirmation) {
                        ok = window.confirm('Delete this task?');
                    }
                    if (!ok) return;
                    var taskitem = getTaskItemSafe(task.entryID, task.storeID);
                    taskitem.Delete();
                    $scope.refreshTasks();
                } catch (e) {
                    writeLog('deleteTask: ' + e);
                }
            };

            $scope.openOneNoteURL = function (url) {
                try {
                    window.event.returnValue = false;
                    if (navigator.msLaunchUri) {
                        navigator.msLaunchUri(url);
                    } else {
                        window.open(url, '_blank').close();
                    }
                    return false;
                } catch (e) {
                    writeLog('openOneNoteURL: ' + e);
                }
            };

            function rebuildLaneOptions() {
                try {
                    var opts = [];
                    ($scope.config.LANES || []).forEach(function (l) {
                        var id = sanitizeId(l.id);
                        if (!id) return;
                        opts.push({ id: id, title: (l.title || id) });
                    });
                    if (opts.length === 0) {
                        opts.push({ id: 'backlog', title: 'Backlog' });
                    }
                    $scope.laneOptions = opts;
                    $scope.migrationLaneOptions = opts;
                } catch (e) {
                    $scope.laneOptions = [];
                    $scope.migrationLaneOptions = [];
                }
            }

            // Settings: lanes
            $scope.addLane = function () {
                var title = ($scope.ui.newLaneTitle || '').trim();
                var id = sanitizeId($scope.ui.newLaneId || title);
                var color = ($scope.ui.newLaneColor || '').trim();
                if (!title || !id) {
                    alert('Lane title and id are required.');
                    return;
                }
                if (color && !isValidHexColor(color)) {
                    alert('Lane color must be in #RRGGBB format.');
                    return;
                }
                for (var i = 0; i < $scope.config.LANES.length; i++) {
                    if (sanitizeId($scope.config.LANES[i].id) === id) {
                        alert('Lane id already exists.');
                        return;
                    }
                }
                $scope.config.LANES.push({ id: id, title: title, color: color || '#94a3b8', wipLimit: 0, enabled: true, outlookStatus: null });
                $scope.ui.newLaneTitle = '';
                $scope.ui.newLaneId = '';
                saveConfig();
                rebuildLaneOptions();
                $scope.refreshTasks();
            };

            $scope.removeLane = function (index) {
                if (index < 0 || index >= $scope.config.LANES.length) return;
                if (!window.confirm('Remove this lane from the board? (Tasks will not be deleted.)')) return;
                $scope.config.LANES.splice(index, 1);
                saveConfig();
                rebuildLaneOptions();
                $scope.refreshTasks();
            };

            $scope.moveLaneUp = function (index) {
                if (index <= 0 || index >= $scope.config.LANES.length) return;
                var tmp = $scope.config.LANES[index - 1];
                $scope.config.LANES[index - 1] = $scope.config.LANES[index];
                $scope.config.LANES[index] = tmp;
                saveConfig();
                rebuildLaneOptions();
                $scope.refreshTasks();
            };

            $scope.moveLaneDown = function (index) {
                if (index < 0 || index >= $scope.config.LANES.length - 1) return;
                var tmp = $scope.config.LANES[index + 1];
                $scope.config.LANES[index + 1] = $scope.config.LANES[index];
                $scope.config.LANES[index] = tmp;
                saveConfig();
                rebuildLaneOptions();
                $scope.refreshTasks();
            };

            $scope.applyLaneTemplate = function (templateId) {
                try {
                    $scope.config.LANES = laneTemplate(templateId);
                    saveConfig();
                    rebuildLaneOptions();
                } catch (e) {
                    writeLog('applyLaneTemplate: ' + e);
                }
            };

            // Settings: themes
            $scope.importThemeFromFile = function () {
                try {
                    var fileInput = document.getElementById('themeCssFile');
                    if (!fileInput || !fileInput.files || fileInput.files.length === 0) {
                        alert('Choose a .css file first.');
                        return;
                    }
                    var name = ($scope.ui.importThemeName || '').trim();
                    var id = sanitizeId($scope.ui.importThemeId || name);
                    if (!name || !id) {
                        alert('Theme name and id are required.');
                        return;
                    }
                    var file = fileInput.files[0];
                    var reader = new FileReader();
                    reader.onload = function (evt) {
                        var cssText = String(evt.target.result || '');
                        if (!isCssLocalOnly(cssText)) {
                            alert('Theme import rejected. Themes must be local-only (no http/https/@import), and must not use IE-specific scriptable CSS (expression/behavior).');
                            return;
                        }
                        $scope.$apply(function () {
                            $scope.config.THEME.customThemes.push({ id: id, name: name, cssText: cssText });
                            $scope.config.THEME.activeThemeId = id;
                            saveConfig();
                            rebuildThemeList();
                            $scope.applyTheme();
                            showToast('success', 'Theme imported', name);
                            $scope.ui.importThemeName = '';
                            $scope.ui.importThemeId = '';
                            fileInput.value = '';
                        });
                    };
                    reader.onerror = function () {
                        alert('Failed to read theme file.');
                    };
                    reader.readAsText(file);
                } catch (e) {
                    writeLog('importThemeFromFile: ' + e);
                }
            };

            $scope.addFolderTheme = function () {
                var name = ($scope.ui.folderThemeName || '').trim();
                var id = sanitizeId($scope.ui.folderThemeId || name);
                var href = ($scope.ui.folderThemeHref || '').trim();
                if (!name || !id || !href) {
                    alert('Theme name, id and CSS path are required.');
                    return;
                }
                if (!isSafeLocalCssPath(href)) {
                    alert('Folder theme path must be a relative local path (for example: themes/my-theme/theme.css).');
                    return;
                }
                $scope.config.THEME.folderThemes.push({ id: id, name: name, cssHref: href });
                $scope.config.THEME.activeThemeId = id;
                saveConfig();
                rebuildThemeList();
                $scope.applyTheme();
                showToast('success', 'Theme added', name);
                $scope.ui.folderThemeName = '';
                $scope.ui.folderThemeId = '';
                $scope.ui.folderThemeHref = '';
            };

            // Projects
            $scope.openCreateProject = function () {
                $scope.ui.createProjectMode = 'create';
                $scope.ui.linkProjectEntryID = '';
                $scope.ui.newProjectName = '';
                $scope.ui.showCreateProject = true;
            };

            $scope.openLinkProject = function () {
                $scope.ui.createProjectMode = 'link';
                $scope.ui.linkProjectEntryID = '';
                $scope.ui.newProjectName = '';
                $scope.ui.showCreateProject = true;
            };

            function getProjectAll(entryID) {
                for (var i = 0; i < $scope.projectsAll.length; i++) {
                    if ($scope.projectsAll[i].entryID === entryID) {
                        return $scope.projectsAll[i];
                    }
                }
                return null;
            }

            function isProjectHidden(entryID) {
                return ($scope.config.PROJECTS.hiddenProjectEntryIDs || []).indexOf(entryID) !== -1;
            }

            function setProjectHidden(entryID, hidden) {
                if (!$scope.config.PROJECTS.hiddenProjectEntryIDs) {
                    $scope.config.PROJECTS.hiddenProjectEntryIDs = [];
                }
                var arr = $scope.config.PROJECTS.hiddenProjectEntryIDs;
                var idx = arr.indexOf(entryID);
                if (hidden) {
                    if (idx === -1) arr.push(entryID);
                } else {
                    if (idx !== -1) arr.splice(idx, 1);
                }
            }

            $scope.selectProject = function (entryID) {
                try {
                    if (!entryID) return;
                    // Selecting a hidden project makes it visible.
                    if (isProjectHidden(entryID)) {
                        setProjectHidden(entryID, false);
                        saveConfig();
                    }
                    loadProjects();
                    $scope.ui.projectEntryID = entryID;
                    ensureSelectedProject();
                    saveState();
                    $scope.refreshTasks();
                } catch (e) {
                    writeLog('selectProject: ' + e);
                }
            };

            $scope.toggleProjectHidden = function (p) {
                try {
                    if (!p || !p.entryID) return;
                    var hiding = !p.isHidden;

                    // Prevent hiding the last visible project
                    if (hiding && $scope.projects.length <= 1) {
                        alert('At least one project must remain visible.');
                        return;
                    }

                    setProjectHidden(p.entryID, hiding);

                    // If the default project is hidden, pick a new default
                    if (hiding && $scope.config.PROJECTS.defaultProjectEntryID === p.entryID) {
                        // choose first visible after reload
                        saveConfig();
                        loadProjects();
                        if ($scope.projects.length > 0) {
                            $scope.config.PROJECTS.defaultProjectEntryID = $scope.projects[0].entryID;
                        }
                    }

                    saveConfig();
                    loadProjects();
                    ensureSelectedProject();
                    saveState();
                    showToast('info', hiding ? 'Project hidden' : 'Project shown', p.name);
                    $scope.refreshTasks();
                } catch (e) {
                    writeLog('toggleProjectHidden: ' + e);
                }
            };

            $scope.unlinkProject = function (entryID) {
                try {
                    if (!entryID) return;
                    var p = getProjectAll(entryID);
                    if (!p || !p.isLinked) {
                        return;
                    }
                    if (!window.confirm('Unlink this project from the board? (The Outlook folder will not be deleted.)')) {
                        return;
                    }

                    var next = [];
                    ($scope.config.PROJECTS.linkedProjects || []).forEach(function (lp) {
                        if (lp && lp.entryID && lp.entryID !== entryID) {
                            next.push(lp);
                        }
                    });
                    $scope.config.PROJECTS.linkedProjects = next;

                    // Also unhide (so it does not stay hidden if re-linked later)
                    setProjectHidden(entryID, false);

                    saveConfig();
                    loadProjects();

                    if ($scope.config.PROJECTS.defaultProjectEntryID === entryID) {
                        if ($scope.projects.length > 0) {
                            $scope.config.PROJECTS.defaultProjectEntryID = $scope.projects[0].entryID;
                            saveConfig();
                        }
                    }
                    ensureSelectedProject();
                    saveState();
                    showToast('info', 'Project unlinked', p.name);
                    $scope.refreshTasks();
                } catch (e) {
                    writeLog('unlinkProject: ' + e);
                }
            };

            $scope.openRenameProject = function (p) {
                if (!p || !p.entryID) return;
                if (p.isDefaultTasks) {
                    alert('The default Tasks folder cannot be renamed from here.');
                    return;
                }
                $scope.ui.renameProjectEntryID = p.entryID;
                $scope.ui.renameProjectStoreID = p.storeID;
                $scope.ui.renameProjectName = p.name;
                $scope.ui.showRenameProject = true;
            };

            $scope.submitRenameProject = function () {
                try {
                    var entryID = $scope.ui.renameProjectEntryID;
                    var storeID = $scope.ui.renameProjectStoreID;
                    var newName = String($scope.ui.renameProjectName || '').trim();
                    if (!entryID) return;
                    if (!newName) {
                        alert('Project name is required.');
                        return;
                    }

                    var folder = getFolderFromIDs(entryID, storeID);
                    if (!folder) {
                        alert('Could not locate the project folder in Outlook.');
                        return;
                    }
                    folder.Name = newName;

                    // Update linked project display name (best-effort)
                    ($scope.config.PROJECTS.linkedProjects || []).forEach(function (lp) {
                        if (lp && lp.entryID === entryID) {
                            lp.name = newName;
                        }
                    });
                    saveConfig();
                    loadProjects();
                    $scope.ui.showRenameProject = false;
                    showToast('success', 'Project renamed', newName);
                } catch (e) {
                    writeLog('submitRenameProject: ' + e);
                    alert('Rename failed: ' + e);
                }
            };

            function linkExistingProject(entryID) {
                var id = String(entryID || '').trim();
                if (!id) {
                    alert('Please select a folder.');
                    return null;
                }
                var folder = null;
                for (var i = 0; i < $scope.availableProjectFolders.length; i++) {
                    if ($scope.availableProjectFolders[i].entryID === id) {
                        folder = $scope.availableProjectFolders[i];
                        break;
                    }
                }
                if (!folder) {
                    alert('Selected folder is not available.');
                    return null;
                }
                if (!$scope.config.PROJECTS.linkedProjects) {
                    $scope.config.PROJECTS.linkedProjects = [];
                }
                var exists = false;
                $scope.config.PROJECTS.linkedProjects.forEach(function (p) {
                    if (p && p.entryID === folder.entryID) exists = true;
                });
                if (!exists) {
                    $scope.config.PROJECTS.linkedProjects.push({
                        name: folder.name,
                        entryID: folder.entryID,
                        storeID: folder.storeID
                    });
                }
                saveConfig();
                loadProjects();
                return folder;
            }

            $scope.submitCreateProject = function () {
                if ($scope.ui.createProjectMode === 'link') {
                    var f = linkExistingProject($scope.ui.linkProjectEntryID);
                    if (!f) return;
                    $scope.ui.projectEntryID = f.entryID;
                    if (!$scope.config.PROJECTS.defaultProjectEntryID) {
                        $scope.config.PROJECTS.defaultProjectEntryID = f.entryID;
                        saveConfig();
                    }
                    saveState();
                    $scope.ui.showCreateProject = false;
                    showToast('success', 'Project linked', f.name || 'Project');
                    $scope.refreshTasks();
                    return;
                }
                $scope.createProject($scope.ui.newProjectName);
            };

            $scope.createProject = function (name) {
                try {
                    var projectName = String(name || '').trim();
                    if (!projectName) {
                        alert('Project name is required.');
                        return;
                    }

                    // Create (or reuse) root folder
                    var root = getTaskFolder($scope.filter.mailbox, $scope.config.PROJECTS.rootFolderName);
                    // Create project folder under root
                    var pf = getOrCreateFolder($scope.filter.mailbox, projectName, root.Folders, OlDefaultFolders.olFolderTasks);

                    // Refresh projects and select
                    loadProjects();
                    $scope.ui.projectEntryID = pf.EntryID;
                    if (!$scope.config.PROJECTS.defaultProjectEntryID) {
                        $scope.config.PROJECTS.defaultProjectEntryID = pf.EntryID;
                    }
                    saveConfig();
                    saveState();

                    $scope.ui.showCreateProject = false;
                    $scope.ui.newProjectName = '';
                    showToast('success', 'Project created', projectName);
                    $scope.refreshTasks();
                } catch (e) {
                    writeLog('createProject: ' + e);
                    alert('Failed to create project: ' + e);
                }
            };

            // Tools: Move tasks between projects
            $scope.openMoveTasks = function () {
                try {
                    rebuildLaneOptions();
                    $scope.ui.move.fromProjectEntryID = $scope.ui.projectEntryID;
                    $scope.ui.move.toProjectEntryID = '';
                    $scope.ui.move.mode = 'all';
                    $scope.ui.move.laneId = ($scope.laneOptions.length > 0) ? $scope.laneOptions[0].id : '';
                    $scope.ui.move.running = false;
                    $scope.ui.move.progress = { total: 0, done: 0, percent: 0 };

                    // Pick a default destination if possible
                    for (var i = 0; i < $scope.projectsAll.length; i++) {
                        if ($scope.projectsAll[i].entryID && $scope.projectsAll[i].entryID !== $scope.ui.move.fromProjectEntryID) {
                            $scope.ui.move.toProjectEntryID = $scope.projectsAll[i].entryID;
                            break;
                        }
                    }

                    $scope.ui.showMoveTasks = true;
                } catch (e) {
                    writeLog('openMoveTasks: ' + e);
                }
            };

            $scope.closeMoveTasks = function () {
                if ($scope.ui.move && $scope.ui.move.running) {
                    alert('Move is in progress.');
                    return;
                }
                $scope.ui.showMoveTasks = false;
            };

            function getProjectFolderByEntryID(entryID) {
                var p = getProjectAll(entryID);
                if (!p) return null;
                return getFolderFromIDs(p.entryID, p.storeID);
            }

            $scope.runMoveTasks = function () {
                try {
                    if ($scope.ui.move.running) return;

                    var fromId = $scope.ui.move.fromProjectEntryID;
                    var toId = $scope.ui.move.toProjectEntryID;
                    if (!fromId || !toId) {
                        alert('Please select both source and destination projects.');
                        return;
                    }
                    if (fromId === toId) {
                        alert('Source and destination must be different.');
                        return;
                    }

                    var fromFolder = getProjectFolderByEntryID(fromId);
                    var toFolder = getProjectFolderByEntryID(toId);
                    if (!fromFolder || !toFolder) {
                        alert('Could not locate one of the project folders in Outlook.');
                        return;
                    }

                    var fromStoreID = '';
                    try { fromStoreID = fromFolder.StoreID; } catch (e0) { fromStoreID = ''; }

                    var mode = $scope.ui.move.mode;
                    var laneFilter = sanitizeId($scope.ui.move.laneId);
                    if (mode === 'lane' && !laneFilter) {
                        alert('Please select a lane.');
                        return;
                    }

                    // Scan source folder
                    var moveList = [];
                    var items = fromFolder.Items;
                    var count = items.Count;
                    for (var i = 1; i <= count; i++) {
                        try {
                            var it = items(i);
                            var currentLane = '';
                            try { currentLane = sanitizeId(getUserProperty(it, PROP_LANE_ID)); } catch (e1) { currentLane = ''; }

                            var laneOrder = null;
                            try {
                                var laneOrderRaw = getUserProperty(it, PROP_LANE_ORDER);
                                if (laneOrderRaw !== '' && laneOrderRaw !== null && laneOrderRaw !== undefined) {
                                    var n = parseInt(laneOrderRaw, 10);
                                    if (!isNaN(n)) laneOrder = n;
                                }
                            } catch (e1b) {
                                laneOrder = null;
                            }

                            var match = false;
                            if (mode === 'all') {
                                match = true;
                            } else if (mode === 'unassigned') {
                                match = (currentLane === '');
                            } else if (mode === 'lane') {
                                match = (currentLane === laneFilter);
                            }

                            if (match) {
                                moveList.push({ entryID: it.EntryID, laneId: currentLane, laneOrder: laneOrder });
                            }
                        } catch (e2) {
                            // ignore
                        }
                    }

                    if (moveList.length === 0) {
                        alert('No tasks matched your selection.');
                        return;
                    }

                    var fromName = (function () { var p = getProjectAll(fromId); return p ? p.name : 'source'; })();
                    var toName = (function () { var p = getProjectAll(toId); return p ? p.name : 'destination'; })();
                    if (!window.confirm('Move ' + moveList.length + ' tasks from "' + fromName + '" to "' + toName + '"?')) {
                        return;
                    }

                    $scope.ui.move.running = true;
                    $scope.ui.move.progress.total = moveList.length;
                    $scope.ui.move.progress.done = 0;
                    $scope.ui.move.progress.percent = 0;

                    var idx = 0;
                    var batchSize = 10;

                    function step() {
                        var end = Math.min(idx + batchSize, moveList.length);
                        for (; idx < end; idx++) {
                            try {
                                var w = moveList[idx];
                                var taskitem = getTaskItemFromIDs(w.entryID, fromStoreID);
                                var moved = taskitem.Move(toFolder);
                                // Ensure lane metadata remains on the moved task
                                try {
                                    if (moved) {
                                        if (w.laneId) {
                                            setUserProperty(moved, PROP_LANE_ID, w.laneId, OlUserPropertyType.olText);
                                        }
                                        if (w.laneOrder !== null && w.laneOrder !== undefined) {
                                            setUserProperty(moved, PROP_LANE_ORDER, w.laneOrder, OlUserPropertyType.olNumber);
                                        }
                                        moved.Save();
                                    }
                                } catch (e3b) {
                                    // ignore
                                }
                            } catch (e3) {
                                writeLog('move task: ' + e3);
                            }
                            $scope.ui.move.progress.done = idx + 1;
                        }
                        $scope.ui.move.progress.percent = Math.round(($scope.ui.move.progress.done * 100) / $scope.ui.move.progress.total);

                        if (idx < moveList.length) {
                            $timeout(step, 0);
                        } else {
                            $scope.ui.move.running = false;
                            $scope.ui.showMoveTasks = false;
                            showToast('success', 'Move completed', String(moveList.length) + ' tasks moved');
                            $scope.refreshTasks();
                        }
                    }

                    $timeout(step, 0);
                } catch (e) {
                    writeLog('runMoveTasks: ' + e);
                    alert('Move failed: ' + e);
                    $scope.ui.move.running = false;
                }
            };

            // Tools: Migration (assign lane ids based on Outlook Task Status)
            function outlookStatusLabel(value) {
                return taskStatusText(value) || ('Status ' + value);
            }

            function buildKnownLaneSet() {
                var known = {};
                ($scope.config.LANES || []).forEach(function (l) {
                    var id = sanitizeId(l.id);
                    if (id) known[id] = true;
                });
                return known;
            }

            function defaultLaneForStatus(statusValue) {
                var lanes = ($scope.config.LANES || []);
                for (var i = 0; i < lanes.length; i++) {
                    if (lanes[i].outlookStatus === statusValue) {
                        var id = sanitizeId(lanes[i].id);
                        if (id) return id;
                    }
                }
                return '';
            }

            function updateMigrationCounts() {
                try {
                    if (!$scope.ui.migration || !$scope.ui.migration.scanTasks) return;
                    var known = buildKnownLaneSet();
                    var onlyUnassigned = !!$scope.ui.migration.onlyUnassigned;
                    var treatUnknown = !!$scope.ui.migration.treatUnknownAsUnassigned;

                    var rows = $scope.ui.migration.mappingRows || [];
                    var byStatus = {};
                    rows.forEach(function (r) { byStatus[r.statusValue] = r; r.count = 0; });

                    $scope.ui.migration.scanTasks.forEach(function (t) {
                        var lane = sanitizeId(t.laneId);
                        var assigned = lane && known[lane];
                        if (lane && !known[lane] && treatUnknown) {
                            assigned = false;
                            lane = '';
                        }
                        if (onlyUnassigned && assigned) {
                            return;
                        }
                        var r = byStatus[t.statusValue];
                        if (r) {
                            r.count++;
                        }
                    });
                } catch (e) {
                    writeLog('updateMigrationCounts: ' + e);
                }
            }

            $scope.$watch('ui.migration.onlyUnassigned', function () {
                if ($scope.ui.showMigration) {
                    updateMigrationCounts();
                }
            });

            $scope.$watch('ui.migration.treatUnknownAsUnassigned', function () {
                if ($scope.ui.showMigration) {
                    updateMigrationCounts();
                }
            });

            $scope.openMigration = function () {
                try {
                    rebuildLaneOptions();
                    $scope.ui.migration.running = false;
                    $scope.ui.migration.progress = { total: 0, done: 0, percent: 0, updated: 0, skipped: 0, errors: 0 };

                    var statusValues = [0, 1, 3, 2, 4];
                    var rows = [];
                    for (var i = 0; i < statusValues.length; i++) {
                        var sv = statusValues[i];
                        rows.push({
                            statusValue: sv,
                            statusText: outlookStatusLabel(sv),
                            laneId: defaultLaneForStatus(sv),
                            count: 0
                        });
                    }
                    $scope.ui.migration.mappingRows = rows;

                    // Scan tasks in the current project
                    var folder = getSelectedProjectFolder();
                    if (!folder) {
                        $scope.ui.migration.scanTasks = [];
                        updateMigrationCounts();
                        $scope.ui.showMigration = true;
                        return;
                    }

                    var folderStoreID = '';
                    try { folderStoreID = folder.StoreID; } catch (e0) { folderStoreID = ''; }

                    var scan = [];
                    var items = folder.Items;
                    var count = items.Count;
                    for (var j = 1; j <= count; j++) {
                        try {
                            var it = items(j);
                            var laneId = '';
                            try { laneId = getUserProperty(it, PROP_LANE_ID); } catch (e1) { laneId = ''; }
                            scan.push({
                                entryID: it.EntryID,
                                storeID: folderStoreID,
                                statusValue: it.Status,
                                laneId: laneId
                            });
                        } catch (e2) {
                            // ignore
                        }
                    }
                    $scope.ui.migration.scanTasks = scan;
                    updateMigrationCounts();
                    $scope.ui.showMigration = true;
                } catch (e) {
                    writeLog('openMigration: ' + e);
                }
            };

            $scope.closeMigration = function () {
                if ($scope.ui.migration && $scope.ui.migration.running) {
                    alert('Migration is in progress.');
                    return;
                }
                $scope.ui.showMigration = false;
            };

            $scope.runMigration = function () {
                try {
                    if ($scope.ui.migration.running) return;
                    var scan = $scope.ui.migration.scanTasks || [];
                    if (scan.length === 0) {
                        alert('No tasks found in the current project.');
                        return;
                    }

                    var known = buildKnownLaneSet();
                    var onlyUnassigned = !!$scope.ui.migration.onlyUnassigned;
                    var treatUnknown = !!$scope.ui.migration.treatUnknownAsUnassigned;

                    var mapping = {};
                    ($scope.ui.migration.mappingRows || []).forEach(function (r) {
                        mapping[r.statusValue] = sanitizeId(r.laneId);
                    });

                    var work = [];
                    scan.forEach(function (t) {
                        var current = sanitizeId(t.laneId);
                        var assigned = current && known[current];
                        if (current && !known[current] && treatUnknown) {
                            assigned = false;
                            current = '';
                        }
                        if (onlyUnassigned && assigned) {
                            return;
                        }
                        var target = mapping[t.statusValue] || '';
                        if (!target) {
                            return;
                        }
                        if (target === current) {
                            return;
                        }
                        work.push({ entryID: t.entryID, storeID: t.storeID, laneId: target });
                    });

                    if (work.length === 0) {
                        alert('No tasks matched your migration scope.');
                        return;
                    }

                    if (!window.confirm('Assign lanes for ' + work.length + ' tasks in this project?')) {
                        return;
                    }

                    $scope.ui.migration.running = true;
                    $scope.ui.migration.progress.total = work.length;
                    $scope.ui.migration.progress.done = 0;
                    $scope.ui.migration.progress.updated = 0;
                    $scope.ui.migration.progress.skipped = scan.length - work.length;
                    $scope.ui.migration.progress.errors = 0;
                    $scope.ui.migration.progress.percent = 0;

                    var idx = 0;
                    var batchSize = 12;
                    function step() {
                        var end = Math.min(idx + batchSize, work.length);
                        for (; idx < end; idx++) {
                            try {
                                var w = work[idx];
                                var taskitem = getTaskItemFromIDs(w.entryID, w.storeID);
                                setUserProperty(taskitem, PROP_LANE_ID, w.laneId, OlUserPropertyType.olText);
                                taskitem.Save();
                                $scope.ui.migration.progress.updated++;
                            } catch (e1) {
                                $scope.ui.migration.progress.errors++;
                                writeLog('migrate task: ' + e1);
                            }
                            $scope.ui.migration.progress.done = idx + 1;
                        }
                        $scope.ui.migration.progress.percent = Math.round(($scope.ui.migration.progress.done * 100) / $scope.ui.migration.progress.total);
                        if (idx < work.length) {
                            $timeout(step, 0);
                        } else {
                            $scope.ui.migration.running = false;
                            $scope.ui.showMigration = false;
                            showToast('success', 'Migration completed', String($scope.ui.migration.progress.updated) + ' tasks updated');
                            $scope.refreshTasks();
                        }
                    }
                    $timeout(step, 0);
                } catch (e) {
                    writeLog('runMigration: ' + e);
                    alert('Migration failed: ' + e);
                    $scope.ui.migration.running = false;
                }
            };

            // Setup wizard
            $scope.closeSetupWizard = function () {
                $scope.ui.showSetupWizard = false;
            };

            $scope.prevSetupStep = function () {
                if ($scope.ui.setupStep > 1) {
                    $scope.ui.setupStep--;
                }
            };

            $scope.nextSetupStep = function () {
                try {
                    if ($scope.ui.setupStep === 2) {
                        // Ensure root + default project exist
                        var rootName = String($scope.config.PROJECTS.rootFolderName || DEFAULT_ROOT_FOLDER_NAME).trim();
                        if (!rootName) rootName = DEFAULT_ROOT_FOLDER_NAME;
                        $scope.config.PROJECTS.rootFolderName = rootName;

                        // Always create root (recommended; used for new projects)
                        var root = getTaskFolder($scope.filter.mailbox, rootName);

                        if ($scope.ui.setupProjectMode === 'link') {
                            var lf = linkExistingProject($scope.ui.setupExistingProjectEntryID);
                            if (!lf) {
                                alert('Please select an existing folder to link.');
                                return;
                            }
                            $scope.config.PROJECTS.defaultProjectEntryID = lf.entryID;
                            $scope.ui.projectEntryID = lf.entryID;
                            saveConfig();
                        } else {
                            var projName = String($scope.ui.setupDefaultProjectName || 'General').trim();
                            if (!projName) projName = 'General';
                            var pf = getOrCreateFolder($scope.filter.mailbox, projName, root.Folders, OlDefaultFolders.olFolderTasks);
                            loadProjects();
                            $scope.config.PROJECTS.defaultProjectEntryID = pf.EntryID;
                            $scope.ui.projectEntryID = pf.EntryID;
                            saveConfig();
                        }
                    }
                    if ($scope.ui.setupStep < 4) {
                        $scope.ui.setupStep++;
                    }
                } catch (e) {
                    writeLog('nextSetupStep: ' + e);
                    alert('Setup failed: ' + e);
                }
            };

            $scope.finishSetup = function () {
                try {
                    $scope.config.SETUP.completed = true;
                    saveConfig();
                    $scope.ui.showSetupWizard = false;
                    $scope.ui.mode = 'board';
                    loadProjects();
                    ensureSelectedProject();
                    $scope.applyTheme();
                    showToast('success', 'Setup complete', '');
                    $scope.refreshTasks();
                } catch (e) {
                    writeLog('finishSetup: ' + e);
                }
            };

            $scope.saveAndReturn = function () {
                saveConfig();
                $scope.applyTheme();
                loadProjects();
                ensureSelectedProject();
                $scope.switchMode('board');
                showToast('success', 'Settings saved', '');
                $scope.refreshTasks();
            };

            // Diagnostics
            $scope.openDiagnostics = function () {
                try {
                    var logRaw = getJournalItem(LOG_ID);
                    var log = [];
                    if (logRaw !== null) {
                        try { log = JSON.parse(logRaw); } catch (e) { log = []; }
                    }
                    var payload = {
                        app: 'Kanban for Outlook',
                        version: $scope.version,
                        outlookVersion: (function () { try { return getOutlookVersion(); } catch (e) { return 'unknown'; } })(),
                        mailbox: $scope.filter.mailbox,
                        projectEntryID: $scope.ui.projectEntryID,
                        filter: $scope.filter,
                        config: $scope.config,
                        recentLog: log.slice(0, 50)
                    };
                    $scope.diagnosticsText = JSON.stringify(payload, null, 2);
                    $scope.ui.showDiagnostics = true;
                } catch (e) {
                    writeLog('openDiagnostics: ' + e);
                }
            };

            $scope.copyDiagnostics = function () {
                try {
                    var text = $scope.diagnosticsText || '';
                    if (window.clipboardData && window.clipboardData.setData) {
                        window.clipboardData.setData('Text', text);
                        return;
                    }
                    // best-effort fallback
                    var ta = document.createElement('textarea');
                    ta.value = text;
                    document.body.appendChild(ta);
                    ta.select();
                    document.execCommand('copy');
                    document.body.removeChild(ta);
                } catch (e) {
                    alert('Copy failed. You can still select and copy from the text box.');
                }
            };

            $scope.dismissToast = function () {
                try {
                    if ($scope.ui && $scope.ui.toast) {
                        $scope.ui.toast.show = false;
                    }
                } catch (e) {
                    // ignore
                }
            };

            // Init
            $scope.init = function () {
                $scope.isBrowserSupported = checkBrowser();
                if (!$scope.isBrowserSupported) {
                    return;
                }

                readConfig();
                rebuildLaneOptions();
                rebuildThemeList();
                $scope.applyTheme();

                initCategories();
                initMailboxes();
                loadAvailableProjectFolders();

                readState();

                // If mailbox was not restored, use default
                if (!$scope.filter.mailbox) {
                    $scope.filter.mailbox = $scope.mailboxes[0];
                }

                loadProjects();
                ensureSelectedProject();

                // First run: if not set up or no projects exist
                if (!$scope.config.SETUP.completed || $scope.projects.length === 0) {
                    $scope.ui.showSetupWizard = true;
                    $scope.ui.setupStep = 1;
                    $scope.ui.setupProjectMode = 'create';
                    $scope.applyLaneTemplate($scope.ui.setupLaneTemplate);
                    saveConfig();
                }

                // Auto refresh guard (lightweight)
                if (refreshTimer) {
                    $interval.cancel(refreshTimer);
                    refreshTimer = undefined;
                }
                refreshTimer = $interval(function () {
                    if ($scope.ui.mode === 'board' && !$scope.ui.showSetupWizard) {
                        // Do not auto-refresh too aggressively; Outlook can be slow.
                    }
                }, 60000);

                $scope.refreshTasks();
            };
        }]);
})();
