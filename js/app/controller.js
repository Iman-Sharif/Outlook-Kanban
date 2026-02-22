'use strict';

(function () {
    // ES5 modules loaded via script tags (see kanban.html)
    var core = (typeof kfoAppCore !== 'undefined') ? kfoAppCore : null;
    var board = core ? core.board : null;
    var outlook = core ? core.outlook : null;

    var CONFIG_ID = core ? core.CONFIG_ID : 'KanbanConfig';
    var STATE_ID = core ? core.STATE_ID : 'KanbanState';
    var LOG_ID = core ? core.LOG_ID : 'KanbanErrorLog';
    var SCHEMA_VERSION = core ? core.SCHEMA_VERSION : 3;
    var PROP_LANE_ID = core ? core.PROP_LANE_ID : 'KFO_LaneId';
    var PROP_LANE_ORDER = core ? core.PROP_LANE_ORDER : 'KFO_LaneOrder';
    var BUILTIN_THEMES = core ? core.BUILTIN_THEMES : [];

    function nowStamp() {
        return core && core.nowStamp ? core.nowStamp() : '';
    }

    function sanitizeId(raw) {
        return core && core.sanitizeId ? core.sanitizeId(raw) : '';
    }

    function isValidHexColor(s) {
        return core && core.isValidHexColor ? core.isValidHexColor(s) : false;
    }

    function isCssLocalOnly(cssText) {
        return core && core.isCssLocalOnly ? core.isCssLocalOnly(cssText) : false;
    }

    function isSafeLocalCssPath(href) {
        return core && core.isSafeLocalCssPath ? core.isSafeLocalCssPath(href) : false;
    }

    function isRealDate(d) {
        return core && core.isRealDate ? core.isRealDate(d) : false;
    }

    function DEFAULT_CONFIG_V3() {
        return core && core.DEFAULT_CONFIG_V3 ? core.DEFAULT_CONFIG_V3() : { SCHEMA_VERSION: SCHEMA_VERSION };
    }

    function laneTemplate(templateId) {
        return core && core.laneTemplate ? core.laneTemplate(templateId) : [];
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
                lastError: null,
                showErrorDetails: false,
                storage: {
                    journal: { ok: null, lastOkAt: '', lastFailAt: '', lastError: '' },
                    config: { readOk: null, writeOk: null, lastReadAt: '', lastWriteAt: '', lastError: '' },
                    state: { readOk: null, writeOk: null, lastReadAt: '', lastWriteAt: '', lastError: '' },
                    log: { readOk: null, writeOk: null, lastReadAt: '', lastWriteAt: '', lastError: '' }
                },
                perf: { lastRefresh: null, history: [] },
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
            $scope.errorDetailsText = '';

            var outlookCategories;

            var toastTimer;
            var sessionLog = [];
            var lastErrorToastSig = '';
            var lastErrorToastAt = 0;
            var storageFailureNotified = false;

            function showToast(type, title, message, ms) {
                try {
                    if (!$scope.ui) return;
                    // Ensure updates are applied even if called outside a digest.
                    $timeout(function () {
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
                    }, 0);
                } catch (e) {
                    // ignore
                }
            }

            function nowIso() {
                return core && core.nowIso ? core.nowIso() : String(new Date());
            }

            function safeErrorString(e) {
                return core && core.safeErrorString ? core.safeErrorString(e) : String(e || '');
            }

            function pushSessionLog(line) {
                try {
                    sessionLog.unshift(String(line || ''));
                    if (sessionLog.length > 200) {
                        sessionLog.pop();
                    }
                } catch (e) {
                    // ignore
                }
            }

            function reportError(context, err, userTitle, userMessage) {
                try {
                    var at = nowIso();
                    var ctx = String(context || '');
                    var msg = safeErrorString(err);
                    var stack = '';
                    try { stack = err && err.stack ? String(err.stack) : ''; } catch (e1) { stack = ''; }

                    var last = {
                        at: at,
                        context: ctx,
                        message: msg,
                        stack: stack
                    };
                    $scope.ui.lastError = last;

                    // Always keep a lightweight session log, even if persistent logging is disabled.
                    pushSessionLog(at + '  ERROR  ' + ctx + (msg ? (': ' + msg) : ''));

                    // De-dupe noisy error toasts
                    var sig = (userTitle || '') + '|' + (userMessage || '') + '|' + ctx + '|' + msg;
                    var nowMs = (new Date()).getTime();
                    if (sig !== lastErrorToastSig || (nowMs - lastErrorToastAt) > 2500) {
                        lastErrorToastSig = sig;
                        lastErrorToastAt = nowMs;
                        showToast('error', userTitle || 'Error', userMessage || 'Something went wrong. Click the ! icon for details.', 5200);
                    }

                    // Ensure UI picks up lastError updates even if no toast is shown.
                    try { $timeout(function () { }, 0); } catch (e2) { /* ignore */ }
                } catch (e) {
                    // ignore
                }
            }

            function showUserError(title, message) {
                showToast('error', title || 'Error', message || '', 4200);
            }

            function markStorage(kind, op, ok, err) {
                try {
                    if (!$scope.ui || !$scope.ui.storage) return;
                    var k = $scope.ui.storage[kind];
                    if (!k) return;
                    var at = nowIso();
                    if (op === 'read') {
                        k.readOk = !!ok;
                        k.lastReadAt = at;
                    }
                    if (op === 'write') {
                        k.writeOk = !!ok;
                        k.lastWriteAt = at;
                    }
                    if (!ok) {
                        k.lastError = safeErrorString(err);
                        $scope.ui.storage.journal.ok = false;
                        $scope.ui.storage.journal.lastFailAt = at;
                        $scope.ui.storage.journal.lastError = safeErrorString(err);
                    } else {
                        // best-effort: treat any successful storage operation as journal being available
                        $scope.ui.storage.journal.ok = true;
                        $scope.ui.storage.journal.lastOkAt = at;
                        $scope.ui.storage.journal.lastError = '';
                    }
                } catch (e) {
                    // ignore
                }
            }

            function storageRead(subject, kind, notifyOnFail) {
                var r;
                if (outlook && outlook.tryGetJournalItem) {
                    r = outlook.tryGetJournalItem(subject);
                } else {
                    r = { ok: false, value: null, error: 'Outlook adapter not available' };
                }

                if (r && r.ok) {
                    markStorage(kind, 'read', true);
                    return r.value;
                }

                markStorage(kind, 'read', false, (r && r.error) ? r.error : 'read failed');
                if (notifyOnFail && !storageFailureNotified) {
                    storageFailureNotified = true;
                    reportError('storage.' + kind + '.read', (r && r.error) ? r.error : 'read failed', 'Local storage unavailable', 'Settings and diagnostics may not be saved. Click the ! icon for details.');
                }
                return null;
            }

            function storageWrite(subject, body, kind, notifyOnFail) {
                var r;
                if (outlook && outlook.trySaveJournalItem) {
                    r = outlook.trySaveJournalItem(subject, body);
                } else {
                    r = { ok: false, value: false, error: 'Outlook adapter not available' };
                }

                if (r && r.ok) {
                    markStorage(kind, 'write', true);
                    return true;
                }

                markStorage(kind, 'write', false, (r && r.error) ? r.error : 'write failed');
                if (notifyOnFail && !storageFailureNotified) {
                    storageFailureNotified = true;
                    reportError('storage.' + kind + '.write', (r && r.error) ? r.error : 'write failed', 'Local storage unavailable', 'Settings and diagnostics may not be saved. Click the ! icon for details.');
                }
                return false;
            }

            function runStorageHealthCheck() {
                // Best-effort read/write check for local Outlook storage.
                // Uses existing subjects and writes the current in-memory values.
                try {
                    // Config
                    try {
                        storageRead(CONFIG_ID, 'config', false);
                        storageWrite(CONFIG_ID, JSON.stringify($scope.config || DEFAULT_CONFIG_V3(), null, 2), 'config', false);
                    } catch (e1) {
                        // ignore
                    }

                    // State (only if enabled)
                    try {
                        if ($scope.config && $scope.config.BOARD && $scope.config.BOARD.saveState) {
                            storageRead(STATE_ID, 'state', false);
                            var state = {
                                private: $scope.filter.private,
                                search: $scope.filter.search,
                                category: $scope.filter.category,
                                mailbox: $scope.filter.mailbox,
                                projectEntryID: $scope.ui.projectEntryID
                            };
                            storageWrite(STATE_ID, JSON.stringify(state, null, 2), 'state', false);
                        }
                    } catch (e2) {
                        // ignore
                    }
                } catch (e) {
                    // ignore
                }
            }

            function writeLog(message) {
                try {
                    var now = new Date();
                    var datetimeString = now.getFullYear() + '-' + (now.getMonth() + 1) + '-' + now.getDate() + ' ' + now.getHours() + ':' + now.getMinutes();
                    var line = datetimeString + '  ' + message;
                    pushSessionLog(line);

                    if (!$scope.config || !$scope.config.LOG_ERRORS) {
                        return;
                    }

                    var logRaw = storageRead(LOG_ID, 'log', false);
                    var log = [];
                    if (logRaw !== null) {
                        try { log = JSON.parse(logRaw); } catch (e) { log = []; }
                    }
                    log.unshift(line);
                    if (log.length > 800) {
                        log.pop();
                    }
                    storageWrite(LOG_ID, JSON.stringify(log, null, 2), 'log', false);
                } catch (e) {
                    // keep silent
                }
            }

            // Allow the Outlook bridge (js/exchange.js) to report errors without using alert().
            try {
                window.kfoReportError = function (context, error) {
                    reportError('outlook.' + String(context || ''), error, 'Outlook error', 'An Outlook operation failed. Click the ! icon for details.');
                };
            } catch (e) {
                // ignore
            }

            // Best-effort capture of unexpected script errors into session diagnostics.
            try {
                window.onerror = function (msg, url, line, col, error) {
                    try {
                        reportError('window.onerror', error || msg, 'Unexpected error', 'A script error occurred. Click the ! icon for details.');
                    } catch (e1) {
                        // ignore
                    }
                    return false;
                };
            } catch (e) {
                // ignore
            }

            function backupLegacyConfig(raw) {
                try {
                    var subject = CONFIG_ID + '.legacy.' + nowStamp();
                    storageWrite(subject, String(raw || ''), 'config', false);
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
                    var raw = storageRead(CONFIG_ID, 'config', true);
                    if (raw === null) {
                        $scope.config = DEFAULT_CONFIG_V3();
                        saveConfig();
                        hasReadConfig = true;
                        return;
                    }
                    try {
                        $scope.config = JSON.parse(JSON.minify(raw));
                    } catch (e) {
                        reportError('readConfig.parse', e, 'Configuration reset', 'Your configuration could not be read and has been reset. Click the ! icon for details.');
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
                    reportError('readConfig', error, 'Configuration error', 'Failed to read configuration. Defaults will be used. Click the ! icon for details.');
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
                    var ok = storageWrite(CONFIG_ID, JSON.stringify($scope.config, null, 2), 'config', false);
                    if (!ok) {
                        var why = '';
                        try { why = ($scope.ui && $scope.ui.storage && $scope.ui.storage.config) ? $scope.ui.storage.config.lastError : ''; } catch (e0) { why = ''; }
                        reportError('saveConfig', why || 'write failed', 'Settings not saved', 'Could not save settings to Outlook storage. Click the ! icon for details.');
                    }
                    return !!ok;
                } catch (e) {
                    reportError('saveConfig', e, 'Settings not saved', 'Could not save settings to Outlook storage. Click the ! icon for details.');
                    return false;
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
                    var raw = storageRead(STATE_ID, 'state', false);
                    if (raw !== null) {
                        try { state = JSON.parse(raw); } catch (e) { /* ignore */ }
                    } else {
                        storageWrite(STATE_ID, JSON.stringify(state, null, 2), 'state', false);
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
                    storageWrite(STATE_ID, JSON.stringify(state, null, 2), 'state', false);
                } catch (e) {
                    writeLog('saveState: ' + e);
                }
            }

            function initCategories() {
                $scope.categories = ['<All Categories>', '<No Category>'];
                try {
                    outlookCategories = outlook && outlook.getOutlookCategories ? outlook.getOutlookCategories() : { names: [], colors: [] };
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
                    var mb = outlook && outlook.getOutlookMailboxes ? outlook.getOutlookMailboxes(!!($scope.config && $scope.config.MULTI_MAILBOX)) : [];
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
                        var tasksFolder = outlook && outlook.getTaskFolderExisting ? outlook.getTaskFolderExisting($scope.filter.mailbox, '') : null;
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
                        var subs = outlook && outlook.listTaskSubFolders ? outlook.listTaskSubFolders($scope.filter.mailbox, '') : [];
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
                    return outlook && outlook.getTaskFolderExisting ? outlook.getTaskFolderExisting($scope.filter.mailbox, $scope.config.PROJECTS.rootFolderName) : null;
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
                        var tf = outlook && outlook.getTaskFolderExisting ? outlook.getTaskFolderExisting($scope.filter.mailbox, '') : null;
                        if (tf) {
                            defaultTasksEntryID = tf.EntryID;
                        }
                    } catch (e0) {
                        defaultTasksEntryID = '';
                    }

                    // Root subfolders
                    var root = getProjectsRootFolderExisting();
                    if (root) {
                        var subs = outlook && outlook.listTaskSubFolders ? outlook.listTaskSubFolders($scope.filter.mailbox, $scope.config.PROJECTS.rootFolderName) : [];
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
                            var f = outlook && outlook.getFolderFromIDs ? outlook.getFolderFromIDs(p.entryID, p.storeID) : null;
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
                        var f = outlook && outlook.getFolderFromIDs ? outlook.getFolderFromIDs(p.entryID, p.storeID) : null;
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
                try {
                    if (board && board.buildLanes) {
                        return board.buildLanes(tasks, $scope.config);
                    }
                } catch (e) {
                    writeLog('buildLanes: ' + e);
                }
                return [];
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

                            var laneId = outlook && outlook.getUserProperty ? outlook.getUserProperty(task, PROP_LANE_ID) : '';
                            var laneOrderRaw = outlook && outlook.getUserProperty ? outlook.getUserProperty(task, PROP_LANE_ORDER) : '';
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
                                oneNoteURL: (outlook && outlook.getUserProperty) ? outlook.getUserProperty(task, 'OneNoteURL') : '',
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
                    var filtersActive = false;
                    if (board && board.applyFilters) {
                        filtersActive = board.applyFilters($scope.lanes, $scope.filter, $scope.privacyFilter);
                    }

                    // To avoid persisting partial ordering, disable drag/drop while filters are active.
                    if ($scope.sortableOptions) {
                        $scope.sortableOptions.disabled = filtersActive;
                    }

                    saveState();
                } catch (e) {
                    writeLog('applyFilters: ' + e);
                }
            };

            function doRefreshTasks() {
                var perf = {
                    at: nowIso(),
                    ok: true,
                    mailbox: $scope.filter.mailbox,
                    projectEntryID: $scope.ui.projectEntryID,
                    counts: { tasks: 0, lanes: 0, filtered: 0 },
                    stepsMs: {},
                    totalMs: 0,
                    error: ''
                };
                var t0 = (new Date()).getTime();
                try {
                    if (!$scope.ui.projectEntryID) {
                        loadProjects();
                        ensureSelectedProject();
                    }
                    var t1 = (new Date()).getTime();

                    var folder = getSelectedProjectFolder();
                    var t2 = (new Date()).getTime();
                    perf.stepsMs.ensureProject = t1 - t0;
                    perf.stepsMs.selectFolder = t2 - t1;

                    if (!folder) {
                        $scope.lanes = buildLanes([]);
                        $scope.applyFilters();
                        perf.counts.lanes = ($scope.lanes || []).length;
                        perf.totalMs = (new Date()).getTime() - t0;
                        return;
                    }

                    var tasks = readTasksFromOutlookFolder(folder);
                    var t3 = (new Date()).getTime();
                    perf.stepsMs.readTasks = t3 - t2;
                    perf.counts.tasks = (tasks || []).length;

                    $scope.lanes = buildLanes(tasks);
                    var t4 = (new Date()).getTime();
                    perf.stepsMs.buildLanes = t4 - t3;
                    perf.counts.lanes = ($scope.lanes || []).length;

                    $scope.applyFilters();
                    var t5 = (new Date()).getTime();
                    perf.stepsMs.applyFilters = t5 - t4;

                    // filtered count
                    try {
                        var n = 0;
                        ($scope.lanes || []).forEach(function (lane) {
                            n += (lane.filteredTasks || []).length;
                        });
                        perf.counts.filtered = n;
                    } catch (eCount) {
                        // ignore
                    }

                    perf.totalMs = t5 - t0;
                } catch (e) {
                    perf.ok = false;
                    perf.error = safeErrorString(e);
                    perf.totalMs = (new Date()).getTime() - t0;
                    reportError('refreshTasks', e, 'Refresh failed', 'Could not read tasks from Outlook. Click the ! icon for details.');
                } finally {
                    try {
                        $scope.ui.perf.lastRefresh = perf;
                        $scope.ui.perf.history.unshift(perf);
                        if ($scope.ui.perf.history.length > 12) {
                            $scope.ui.perf.history.pop();
                        }
                    } catch (e2) {
                        // ignore
                    }
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
                    if (outlook && outlook.getTaskItemFromIDs && storeID) {
                        var it = outlook.getTaskItemFromIDs(entryID, storeID);
                        if (it) return it;
                    }
                } catch (e) {
                    // ignore
                }
                return outlook && outlook.getTaskItem ? outlook.getTaskItem(entryID) : null;
            }

            function setTaskLane(taskEntryID, storeID, laneId) {
                var taskitem = getTaskItemSafe(taskEntryID, storeID);
                if (!taskitem) {
                    reportError('setTaskLane', 'task not available', 'Move failed', 'Could not update the task in Outlook. Click the ! icon for details.');
                    return;
                }
                if (!(outlook && outlook.setUserProperty)) {
                    reportError('setTaskLane', 'Outlook adapter not available', 'Move failed', 'Outlook integration is not available. Click the ! icon for details.');
                    return;
                }
                outlook.setUserProperty(taskitem, PROP_LANE_ID, laneId);
                taskitem.Save();
            }

            function maybeSetTaskOutlookStatus(taskEntryID, storeID, statusValue) {
                try {
                    if (statusValue === null || statusValue === undefined) return;
                    var taskitem = getTaskItemSafe(taskEntryID, storeID);
                    if (!taskitem) return;
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
                    if (!(outlook && outlook.setUserProperty)) {
                        reportError('fixLaneOrder', 'Outlook adapter not available', 'Ordering failed', 'Outlook integration is not available. Click the ! icon for details.');
                        return;
                    }
                    for (var i = 0; i < lane.filteredTasks.length; i++) {
                        var t = lane.filteredTasks[i];
                        var taskitem = getTaskItemSafe(t.entryID, t.storeID);
                        if (!taskitem) {
                            continue;
                        }
                        outlook.setUserProperty(taskitem, PROP_LANE_ORDER, i, OlUserPropertyType.olNumber);
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
                            showUserError('WIP limit reached', 'This lane is over its WIP limit. Move or complete something first, or raise the WIP limit in Settings.');
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
                        showUserError('No project selected', 'Create or select a project first (Projects are Outlook Tasks folders).');
                        return;
                    }
                    var taskitem = folder.Items.Add();

                    // Default sensitivity based on current filter
                    if ($scope.filter.private == $scope.privacyFilter.private.value) {
                        taskitem.Sensitivity = SENSITIVITY.olPrivate;
                    }

                    if (lane && lane.id) {
                        if (outlook && outlook.setUserProperty) {
                            outlook.setUserProperty(taskitem, PROP_LANE_ID, lane.id, OlUserPropertyType.olText);
                            outlook.setUserProperty(taskitem, PROP_LANE_ORDER, 0, OlUserPropertyType.olNumber);
                        } else {
                            // Allow task creation to proceed even if lane metadata cannot be stored.
                            reportError('addTask', 'Outlook adapter not available', 'Lane not set', 'The task was created but could not be placed on a lane. Click the ! icon for details.');
                        }
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
                    reportError('addTask', e, 'Add task failed', 'Could not create a new task in Outlook. Click the ! icon for details.');
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
                    reportError('editTask', e, 'Open task failed', 'Could not open this task in Outlook. Click the ! icon for details.');
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
                    reportError('deleteTask', e, 'Delete failed', 'Could not delete this task. Click the ! icon for details.');
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
                    reportError('openOneNoteURL', e, 'Open link failed', 'Could not open the link. Click the ! icon for details.');
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
                    showUserError('Lane details required', 'Lane title and id are required.');
                    return;
                }
                if (color && !isValidHexColor(color)) {
                    showUserError('Invalid colour', 'Lane colour must be in #RRGGBB format.');
                    return;
                }
                for (var i = 0; i < $scope.config.LANES.length; i++) {
                    if (sanitizeId($scope.config.LANES[i].id) === id) {
                        showUserError('Lane id already exists', 'Choose a different id (letters, numbers, and dashes).');
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
                        showUserError('Theme import', 'Choose a .css file first.');
                        return;
                    }
                    var name = ($scope.ui.importThemeName || '').trim();
                    var id = sanitizeId($scope.ui.importThemeId || name);
                    if (!name || !id) {
                        showUserError('Theme import', 'Theme name and id are required.');
                        return;
                    }
                    var file = fileInput.files[0];
                    var reader = new FileReader();
                    reader.onload = function (evt) {
                        var cssText = String(evt.target.result || '');
                        if (!isCssLocalOnly(cssText)) {
                            showUserError('Theme import rejected', 'Themes must be local-only (no http/https or @import) and must not use IE scriptable CSS (expression/behaviour).');
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
                        showUserError('Theme import failed', 'Failed to read theme file.');
                    };
                    reader.readAsText(file);
                } catch (e) {
                    writeLog('importThemeFromFile: ' + e);
                    reportError('importThemeFromFile', e, 'Theme import failed', 'Could not import the theme. Click the ! icon for details.');
                }
            };

            $scope.addFolderTheme = function () {
                var name = ($scope.ui.folderThemeName || '').trim();
                var id = sanitizeId($scope.ui.folderThemeId || name);
                var href = ($scope.ui.folderThemeHref || '').trim();
                if (!name || !id || !href) {
                    showUserError('Folder theme', 'Theme name, id and CSS path are required.');
                    return;
                }
                if (!isSafeLocalCssPath(href)) {
                    showUserError('Folder theme', 'Folder theme path must be a relative local path (for example: themes/my-theme/theme.css).');
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
                        showUserError('Cannot hide project', 'At least one project must remain visible.');
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
                    showUserError('Cannot rename folder', 'The default Tasks folder cannot be renamed from here.');
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
                        showUserError('Project name required', 'Project name is required.');
                        return;
                    }

                    var folder = outlook && outlook.getFolderFromIDs ? outlook.getFolderFromIDs(entryID, storeID) : null;
                    if (!folder) {
                        showUserError('Folder not found', 'Could not locate the project folder in Outlook.');
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
                    reportError('submitRenameProject', e, 'Rename failed', 'Could not rename the project. Click the ! icon for details.');
                }
            };

            function linkExistingProject(entryID) {
                var id = String(entryID || '').trim();
                if (!id) {
                    showUserError('Select a folder', 'Please select a folder.');
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
                    showUserError('Folder not available', 'Selected folder is not available.');
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
                        showUserError('Project name required', 'Project name is required.');
                        return;
                    }

                    // Create (or reuse) root folder
                    var root = outlook && outlook.getTaskFolder ? outlook.getTaskFolder($scope.filter.mailbox, $scope.config.PROJECTS.rootFolderName) : null;
                    if (!root) {
                        showUserError('Project not created', 'Could not access or create the root Tasks folder in Outlook.');
                        return;
                    }
                    // Create project folder under root
                    var pf = outlook && outlook.getOrCreateFolder ? outlook.getOrCreateFolder($scope.filter.mailbox, projectName, root.Folders, OlDefaultFolders.olFolderTasks) : null;
                    if (!pf) {
                        showUserError('Project not created', 'Could not create the project folder in Outlook.');
                        return;
                    }

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
                    reportError('createProject', e, 'Project not created', 'Failed to create the project folder in Outlook. Click the ! icon for details.');
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
                    showUserError('Move in progress', 'Please wait for the move to complete.');
                    return;
                }
                $scope.ui.showMoveTasks = false;
            };

            function getProjectFolderByEntryID(entryID) {
                var p = getProjectAll(entryID);
                if (!p) return null;
                return outlook && outlook.getFolderFromIDs ? outlook.getFolderFromIDs(p.entryID, p.storeID) : null;
            }

            $scope.runMoveTasks = function () {
                try {
                    if ($scope.ui.move.running) return;

                    var fromId = $scope.ui.move.fromProjectEntryID;
                    var toId = $scope.ui.move.toProjectEntryID;
                    if (!fromId || !toId) {
                        showUserError('Move tasks', 'Please select both source and destination projects.');
                        return;
                    }
                    if (fromId === toId) {
                        showUserError('Move tasks', 'Source and destination must be different.');
                        return;
                    }

                    var fromFolder = getProjectFolderByEntryID(fromId);
                    var toFolder = getProjectFolderByEntryID(toId);
                    if (!fromFolder || !toFolder) {
                        showUserError('Move tasks', 'Could not locate one of the project folders in Outlook.');
                        return;
                    }

                    var fromStoreID = '';
                    try { fromStoreID = fromFolder.StoreID; } catch (e0) { fromStoreID = ''; }

                    var mode = $scope.ui.move.mode;
                    var laneFilter = sanitizeId($scope.ui.move.laneId);
                    if (mode === 'lane' && !laneFilter) {
                        showUserError('Move tasks', 'Please select a lane.');
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
                            currentLane = sanitizeId((outlook && outlook.getUserProperty) ? outlook.getUserProperty(it, PROP_LANE_ID) : '');

                            var laneOrder = null;
                            try {
                                var laneOrderRaw = (outlook && outlook.getUserProperty) ? outlook.getUserProperty(it, PROP_LANE_ORDER) : '';
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
                        showUserError('Move tasks', 'No tasks matched your selection.');
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
                                var taskitem = outlook && outlook.getTaskItemFromIDs ? outlook.getTaskItemFromIDs(w.entryID, fromStoreID) : null;
                                if (!taskitem) {
                                    writeLog('move task: task not found');
                                    continue;
                                }
                                var moved = taskitem.Move(toFolder);
                                // Ensure lane metadata remains on the moved task
                                try {
                                    if (moved) {
                                        if (outlook && outlook.setUserProperty) {
                                            if (w.laneId) {
                                                outlook.setUserProperty(moved, PROP_LANE_ID, w.laneId, OlUserPropertyType.olText);
                                            }
                                            if (w.laneOrder !== null && w.laneOrder !== undefined) {
                                                outlook.setUserProperty(moved, PROP_LANE_ORDER, w.laneOrder, OlUserPropertyType.olNumber);
                                            }
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
                    reportError('runMoveTasks', e, 'Move failed', 'Failed to move tasks between projects. Click the ! icon for details.');
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
                            var laneId = (outlook && outlook.getUserProperty) ? outlook.getUserProperty(it, PROP_LANE_ID) : '';
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
                    showUserError('Migration in progress', 'Please wait for migration to complete.');
                    return;
                }
                $scope.ui.showMigration = false;
            };

            $scope.runMigration = function () {
                try {
                    if ($scope.ui.migration.running) return;
                    var scan = $scope.ui.migration.scanTasks || [];
                    if (scan.length === 0) {
                        showUserError('Migration', 'No tasks found in the current project.');
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
                        showUserError('Migration', 'No tasks matched your migration scope.');
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
                                var taskitem = outlook && outlook.getTaskItemFromIDs ? outlook.getTaskItemFromIDs(w.entryID, w.storeID) : null;
                                if (!taskitem) {
                                    throw 'task not found';
                                }
                                if (!(outlook && outlook.setUserProperty)) {
                                    throw 'Outlook adapter not available';
                                }
                                outlook.setUserProperty(taskitem, PROP_LANE_ID, w.laneId, OlUserPropertyType.olText);
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
                    reportError('runMigration', e, 'Migration failed', 'Failed to assign lanes. Click the ! icon for details.');
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
                        var root = outlook && outlook.getTaskFolder ? outlook.getTaskFolder($scope.filter.mailbox, rootName) : null;
                        if (!root) {
                            showUserError('Setup', 'Could not access or create the root Tasks folder in Outlook.');
                            return;
                        }

                        if ($scope.ui.setupProjectMode === 'link') {
                            var lf = linkExistingProject($scope.ui.setupExistingProjectEntryID);
                            if (!lf) {
                                showUserError('Setup', 'Please select an existing folder to link.');
                                return;
                            }
                            $scope.config.PROJECTS.defaultProjectEntryID = lf.entryID;
                            $scope.ui.projectEntryID = lf.entryID;
                            saveConfig();
                        } else {
                            var projName = String($scope.ui.setupDefaultProjectName || 'General').trim();
                            if (!projName) projName = 'General';
                            var pf = outlook && outlook.getOrCreateFolder ? outlook.getOrCreateFolder($scope.filter.mailbox, projName, root.Folders, OlDefaultFolders.olFolderTasks) : null;
                            if (!pf) {
                                showUserError('Setup', 'Could not create the default project folder in Outlook.');
                                return;
                            }
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
                    reportError('nextSetupStep', e, 'Setup failed', 'Setup could not be completed. Click the ! icon for details.');
                }
            };

            $scope.finishSetup = function () {
                try {
                    $scope.config.SETUP.completed = true;
                    var ok = saveConfig();
                    $scope.ui.showSetupWizard = false;
                    $scope.ui.mode = 'board';
                    loadProjects();
                    ensureSelectedProject();
                    $scope.applyTheme();
                    if (ok) {
                        showToast('success', 'Setup complete', '');
                    } else {
                        showUserError('Setup not saved', 'Your setup could not be saved to Outlook storage. The app will still run, but settings may reset next time.');
                    }
                    $scope.refreshTasks();
                } catch (e) {
                    writeLog('finishSetup: ' + e);
                }
            };

            $scope.saveAndReturn = function () {
                var ok = saveConfig();
                $scope.applyTheme();
                loadProjects();
                ensureSelectedProject();
                $scope.switchMode('board');
                if (ok) {
                    showToast('success', 'Settings saved', '');
                } else {
                    showUserError('Settings not saved', 'Your settings could not be saved to Outlook storage.');
                }
                $scope.refreshTasks();
            };

            // Diagnostics
            $scope.openDiagnostics = function () {
                try {
                    runStorageHealthCheck();

                    var persistedLogRaw = storageRead(LOG_ID, 'log', false);
                    var persistedLog = [];
                    if (persistedLogRaw !== null) {
                        try { persistedLog = JSON.parse(persistedLogRaw); } catch (e) { persistedLog = []; }
                    }

                    var support = (function () {
                        try {
                            if (outlook && outlook.getBrowserSupportDetails) {
                                return outlook.getBrowserSupportDetails();
                            }
                        } catch (e) {
                            // ignore
                        }
                        return { supported: !!$scope.isBrowserSupported, method: 'unknown', error: '' };
                    })();

                    var outlookVersion = (outlook && outlook.getOutlookVersion) ? outlook.getOutlookVersion() : 'unknown';
                    var outlookTodayHome = (outlook && outlook.getOutlookTodayHomePageFolder) ? outlook.getOutlookTodayHomePageFolder() : 'unknown';

                    var selectedProject = null;
                    try { selectedProject = getProjectAll($scope.ui.projectEntryID); } catch (e) { selectedProject = null; }

                    var payload = {
                        app: 'Kanban for Outlook',
                        version: $scope.version,
                        generatedAt: nowIso(),
                        host: {
                            href: (function () { try { return String(window.location.href || ''); } catch (e) { return ''; } })(),
                            userAgent: (function () { try { return String(navigator.userAgent || ''); } catch (e) { return ''; } })(),
                            browserSupport: support
                        },
                        outlook: {
                            version: outlookVersion,
                            todayHomePageFolder: outlookTodayHome
                        },
                        selection: {
                            mailbox: $scope.filter.mailbox,
                            projectEntryID: $scope.ui.projectEntryID,
                            projectName: selectedProject ? selectedProject.name : ''
                        },
                        filter: $scope.filter,
                        perf: $scope.ui.perf,
                        storage: $scope.ui.storage,
                        lastError: $scope.ui.lastError,
                        sessionLog: sessionLog.slice(0, 200),
                        persistedLog: persistedLog.slice(0, 200),
                        config: $scope.config
                    };
                    $scope.diagnosticsText = JSON.stringify(payload, null, 2);
                    $scope.ui.showDiagnostics = true;
                } catch (e) {
                    reportError('openDiagnostics', e, 'Diagnostics failed', 'Could not build diagnostics output. Click the ! icon for details.');
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
                    showUserError('Copy failed', 'You can still select and copy from the text box.');
                }
            };

            $scope.openErrorDetails = function () {
                try {
                    var support = (function () {
                        try {
                            if (outlook && outlook.getBrowserSupportDetails) {
                                return outlook.getBrowserSupportDetails();
                            }
                        } catch (e) {
                            // ignore
                        }
                        return { supported: !!$scope.isBrowserSupported, method: 'unknown', error: '' };
                    })();

                    var payload = {
                        app: 'Kanban for Outlook',
                        version: $scope.version,
                        generatedAt: nowIso(),
                        lastError: $scope.ui.lastError,
                        perf: $scope.ui.perf,
                        storage: $scope.ui.storage,
                        selection: {
                            mailbox: $scope.filter.mailbox,
                            projectEntryID: $scope.ui.projectEntryID
                        },
                        host: {
                            href: (function () { try { return String(window.location.href || ''); } catch (e) { return ''; } })(),
                            userAgent: (function () { try { return String(navigator.userAgent || ''); } catch (e) { return ''; } })(),
                            browserSupport: support
                        },
                        outlookVersion: (outlook && outlook.getOutlookVersion) ? outlook.getOutlookVersion() : 'unknown'
                    };
                    $scope.errorDetailsText = JSON.stringify(payload, null, 2);
                    $scope.ui.showErrorDetails = true;
                } catch (e) {
                    reportError('openErrorDetails', e, 'Error details failed', 'Could not build error details output.');
                }
            };

            $scope.copyErrorDetails = function () {
                try {
                    var text = $scope.errorDetailsText || '';
                    if (window.clipboardData && window.clipboardData.setData) {
                        window.clipboardData.setData('Text', text);
                        return;
                    }
                    var ta = document.createElement('textarea');
                    ta.value = text;
                    document.body.appendChild(ta);
                    ta.select();
                    document.execCommand('copy');
                    document.body.removeChild(ta);
                } catch (e) {
                    showUserError('Copy failed', 'You can still select and copy from the text box.');
                }
            };

            $scope.clearLastError = function () {
                try {
                    $scope.ui.lastError = null;
                    $scope.ui.showErrorDetails = false;
                } catch (e) {
                    // ignore
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
                // Capture basic host info even when Outlook integration is not available.
                try {
                    $scope.env = {
                        href: (function () { try { return String(window.location.href || ''); } catch (e) { return ''; } })(),
                        userAgent: (function () { try { return String(navigator.userAgent || ''); } catch (e) { return ''; } })()
                    };
                } catch (e) {
                    $scope.env = { href: '', userAgent: '' };
                }

                $scope.isBrowserSupported = outlook && outlook.checkBrowser ? outlook.checkBrowser() : false;
                try {
                    $scope.browserSupport = (outlook && outlook.getBrowserSupportDetails) ? outlook.getBrowserSupportDetails() : { supported: !!$scope.isBrowserSupported, method: 'unknown', error: '' };
                } catch (e) {
                    $scope.browserSupport = { supported: !!$scope.isBrowserSupported, method: 'unknown', error: '' };
                }
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
