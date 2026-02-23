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
    var DEFAULT_ROOT_FOLDER_NAME = core ? core.DEFAULT_ROOT_FOLDER_NAME : 'Kanban Projects';
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
        .controller('taskboardController', ['$scope', '$timeout', function ($scope, $timeout) {
            var hasReadConfig = false;
            var hasReadState = false;

            $scope.isBrowserSupported = false;

            $scope.version = (typeof VERSION !== 'undefined') ? VERSION : '0.0.0';

            $scope.rootClasses = {};

            $scope.ui = {
                mode: 'board',
                projectEntryID: '',
                settingsDirty: false,
                settingsBaselineRaw: '',
                isRefreshing: false,
                lastRefreshedAtMs: 0,
                lastRefreshedAtText: '',
                filtersActive: false,
                banner: { show: false, type: 'info', title: '', message: '' },
                toast: { show: false, type: 'info', title: '', message: '', actionLabel: '', onAction: null },
                lastMove: null,
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
                setupProjectMode: 'default',
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
                showShortcuts: false,
                createProjectMode: 'create',
                linkProjectEntryID: '',
                newProjectName: '',
                newLaneTitle: '',
                newLaneId: '',
                newLaneColor: '#60a5fa',
                laneErrors: {},
                newLaneErrors: {},
                importThemeName: '',
                importThemeId: '',
                folderThemeName: '',
                folderThemeId: '',
                folderThemeHref: '',

                // Phase 3: settings transfer
                showSettingsTransfer: false,
                settingsExportIncludeState: true,
                settingsExportText: '',
                settingsImportText: '',
                settingsImportApplyConfig: true,
                settingsImportApplyState: true,

                // Lane id migration tool
                showLaneIdTool: false,
                laneIdTool: {
                    oldId: '',
                    laneTitle: '',
                    newId: '',
                    scope: 'all',
                    running: false,
                    cancelRequested: false,
                    progress: {
                        foldersTotal: 0,
                        foldersDone: 0,
                        tasksTotal: 0,
                        tasksDone: 0,
                        updated: 0,
                        errors: 0
                    }
                },

                // Phase 3: quick add
                quickAddLaneId: '',
                quickAddText: '',
                quickAddSaving: false
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
                due: 'any',
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
            var stateSaveTimer;
            var sessionLog = [];
            var lastErrorToastSig = '';
            var lastErrorToastAt = 0;
            var storageFailureNotified = false;
            var projectFallbackNotified = false;

            var lastSavedThemeId = '';

            var keyboardShortcutsBound = false;

            function focusById(id, selectAll) {
                try {
                    $timeout(function () {
                        try {
                            var el = document.getElementById(String(id || ''));
                            if (el && el.focus) {
                                el.focus();
                                if (selectAll && el.select) {
                                    try { el.select(); } catch (e1) { /* ignore */ }
                                }
                            }
                        } catch (e) {
                            // ignore
                        }
                    }, 0);
                } catch (e0) {
                    // ignore
                }
            }

            function showToast(type, title, message, ms, actionLabel, onAction) {
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
                                message: message || '',
                                actionLabel: actionLabel || '',
                                onAction: onAction || null
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

            $scope.undoLastMove = function () {
                try {
                    if (!$scope.ui || !$scope.ui.lastMove) {
                        showToast('info', 'Nothing to undo', '', 1800);
                        return;
                    }
                    var mv = $scope.ui.lastMove;
                    var ageMs = 0;
                    try { ageMs = (new Date()).getTime() - (mv.atMs || 0); } catch (e0) { ageMs = 0; }
                    if (ageMs > 20000) {
                        $scope.ui.lastMove = null;
                        showToast('info', 'Undo expired', 'Undo is only available briefly after a move.', 2800);
                        return;
                    }

                    var setStatus = !!mv.restoreStatus;
                    var statusValue = setStatus ? mv.beforeStatusValue : null;
                    if (setStatus && (statusValue === '' || statusValue === null || statusValue === undefined)) {
                        setStatus = false;
                        statusValue = null;
                    }

                    var res = updateTaskLaneAndStatus(mv.entryID, mv.storeID, mv.fromLaneId, statusValue, setStatus);
                    if (!res || !res.ok) {
                        showToast('error', 'Undo failed', 'Could not update the task in Outlook', 4200);
                        return;
                    }

                    $scope.ui.lastMove = null;
                    $scope.dismissToast();
                    showToast('success', 'Move undone', '', 1800);
                    $scope.refreshTasks();
                } catch (e) {
                    writeLog('undoLastMove: ' + e);
                    reportError('undoLastMove', e, 'Undo failed', 'Could not undo the last move. Click the ! icon for details.');
                }
            };

            function showBanner(type, title, message) {
                try {
                    if (!$scope.ui || !$scope.ui.banner) return;
                    $scope.ui.banner = {
                        show: true,
                        type: type || 'info',
                        title: title || '',
                        message: message || ''
                    };
                } catch (e) {
                    // ignore
                }
            }

            $scope.dismissBanner = function () {
                try {
                    if ($scope.ui && $scope.ui.banner) {
                        $scope.ui.banner.show = false;
                    }
                } catch (e) {
                    // ignore
                }
            };

            $scope.bannerSwitchToTasks = function () {
                try {
                    var tf = getDefaultTasksFolderExisting();
                    if (tf && tf.EntryID) {
                        $scope.ui.projectEntryID = tf.EntryID;
                        scheduleSaveState();
                        $scope.dismissBanner();
                        $scope.refreshTasks();
                        return;
                    }
                    showUserError('Tasks folder not available', 'Could not access your default Outlook Tasks folder.');
                } catch (e) {
                    writeLog('bannerSwitchToTasks: ' + e);
                }
            };

            $scope.bannerRetry = function () {
                try {
                    $scope.refreshTasks();
                } catch (e) {
                    // ignore
                }
            };

            $scope.copyText = function (text) {
                try {
                    var t = String(text || '');
                    if (window.clipboardData && window.clipboardData.setData) {
                        window.clipboardData.setData('Text', t);
                        showToast('success', 'Copied', '', 1400);
                        return;
                    }
                    var ta = document.createElement('textarea');
                    ta.value = t;
                    document.body.appendChild(ta);
                    ta.select();
                    document.execCommand('copy');
                    document.body.removeChild(ta);
                    showToast('success', 'Copied', '', 1400);
                } catch (e) {
                    showUserError('Copy failed', 'Select and copy the text manually.');
                }
            };

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
                    var dragHandleOnly = ($scope.config && $scope.config.BOARD) ? !!$scope.config.BOARD.dragHandleOnly : false;

                    var classes = {};
                    classes['theme-' + themeId] = true;
                    classes['density-' + density] = true;
                    classes['motion-' + motion] = true;
                    if (dragHandleOnly) {
                        classes['drag-handle-only'] = true;
                    }
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

                    // Phase 4: harden theme safety at apply-time as well (config can be edited outside the UI).
                    var blocked = false;
                    var blockedWhy = '';
                    if (theme && theme.cssHref) {
                        if (!isSafeLocalCssPath(theme.cssHref)) {
                            blocked = true;
                            blockedWhy = 'Theme CSS path must be a relative local path.';
                        }
                    }
                    if (!blocked && theme && theme.cssText) {
                        if (!isCssLocalOnly(theme.cssText)) {
                            blocked = true;
                            blockedWhy = 'Theme CSS must be local-only (no http/https, no protocol-relative //, no @import) and must not use scriptable CSS (expression/behaviour or javascript: URLs).';
                        }
                    }

                    if (blocked) {
                        try {
                            reportError('theme.blocked', blockedWhy, 'Theme blocked', blockedWhy);
                        } catch (e0) {
                            // ignore
                        }
                        themeId = 'kfo-light';
                        try {
                            if ($scope.config && $scope.config.THEME) {
                                $scope.config.THEME.activeThemeId = themeId;
                                saveConfig();
                            }
                        } catch (e1) {
                            // ignore
                        }
                        theme = findThemeById(themeId);
                        showUserError('Theme blocked', blockedWhy);
                    }

                    // Apply root classes (theme + UI)
                    $scope.applyRootClasses();

                    // Apply theme CSS link (fallback to builtin light)
                    var themeLink = document.getElementById('kfo-theme-link');
                    if (themeLink) {
                        if (theme && theme.cssHref) {
                            // Defensive: re-check before applying
                            themeLink.href = isSafeLocalCssPath(theme.cssHref) ? theme.cssHref : 'themes/kfo-light/theme.css';
                        } else {
                            themeLink.href = 'themes/kfo-light/theme.css';
                        }
                    }

                    // Apply imported theme css (optional)
                    var styleEl = ensureThemeStyleElement();
                    if (theme && theme.cssText) {
                        if (isCssLocalOnly(theme.cssText)) {
                            styleEl.styleSheet ? (styleEl.styleSheet.cssText = theme.cssText) : (styleEl.innerHTML = theme.cssText);
                        } else {
                            styleEl.styleSheet ? (styleEl.styleSheet.cssText = '') : (styleEl.innerHTML = '');
                        }
                    } else {
                        styleEl.styleSheet ? (styleEl.styleSheet.cssText = '') : (styleEl.innerHTML = '');
                    }

                    // Persist theme selection only when it actually changed.
                    try {
                        if ($scope.config && $scope.config.THEME) {
                            var activeId = String($scope.config.THEME.activeThemeId || '');
                            if (activeId && activeId !== lastSavedThemeId) {
                                var ok = saveConfig();
                                if (ok) {
                                    lastSavedThemeId = activeId;
                                }
                            }
                        }
                    } catch (e2) {
                        // ignore
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
                    if ($scope.config.UI.keyboardShortcuts === undefined) $scope.config.UI.keyboardShortcuts = DEFAULT_CONFIG_V3().UI.keyboardShortcuts;
                    if ($scope.config.AUTOMATION.setOutlookStatusOnLaneMove === undefined) {
                        $scope.config.AUTOMATION.setOutlookStatusOnLaneMove = DEFAULT_CONFIG_V3().AUTOMATION.setOutlookStatusOnLaneMove;
                    }
                    if (!$scope.config.LANES) $scope.config.LANES = DEFAULT_CONFIG_V3().LANES;
                    ensureAtLeastOneLaneEnabled();
                    if (!$scope.config.THEME) $scope.config.THEME = DEFAULT_CONFIG_V3().THEME;
                    if (!$scope.config.BOARD) $scope.config.BOARD = DEFAULT_CONFIG_V3().BOARD;
                    if ($scope.config.BOARD.quickAddEnabled === undefined) $scope.config.BOARD.quickAddEnabled = DEFAULT_CONFIG_V3().BOARD.quickAddEnabled;
                    if ($scope.config.BOARD.dragHandleOnly === undefined) $scope.config.BOARD.dragHandleOnly = DEFAULT_CONFIG_V3().BOARD.dragHandleOnly;
                    if ($scope.config.USE_CATEGORY_COLORS === undefined) $scope.config.USE_CATEGORY_COLORS = true;
                    if ($scope.config.USE_CATEGORY_COLOR_FOOTERS === undefined) $scope.config.USE_CATEGORY_COLOR_FOOTERS = false;
                    if (!$scope.config.DATE_FORMAT) $scope.config.DATE_FORMAT = 'DD-MMM';
                    if ($scope.config.LOG_ERRORS === undefined) $scope.config.LOG_ERRORS = false;

                    // Track persisted theme id to avoid unnecessary storage writes.
                    try {
                        if ($scope.config && $scope.config.THEME && $scope.config.THEME.activeThemeId) {
                            lastSavedThemeId = String($scope.config.THEME.activeThemeId || '');
                        }
                    } catch (e0t) {
                        // ignore
                    }

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

                    if (ok) {
                        try {
                            if ($scope.config && $scope.config.THEME) {
                                lastSavedThemeId = String($scope.config.THEME.activeThemeId || '');
                            }
                        } catch (e1) {
                            // ignore
                        }

                        updateSettingsBaseline();
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
                    $scope.filter.due = state.due || 'any';
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
                        due: $scope.filter.due,
                        mailbox: $scope.filter.mailbox,
                        projectEntryID: $scope.ui.projectEntryID
                    };
                    storageWrite(STATE_ID, JSON.stringify(state, null, 2), 'state', false);
                } catch (e) {
                    writeLog('saveState: ' + e);
                }
            }

            function scheduleSaveState() {
                try {
                    if (!$scope.config || !$scope.config.BOARD || !$scope.config.BOARD.saveState) {
                        return;
                    }
                    if (stateSaveTimer) {
                        $timeout.cancel(stateSaveTimer);
                        stateSaveTimer = null;
                    }
                    stateSaveTimer = $timeout(function () {
                        stateSaveTimer = null;
                        saveState();
                    }, 400);
                } catch (e) {
                    writeLog('scheduleSaveState: ' + e);
                }
            }

            function updateSettingsBaseline() {
                try {
                    if (!$scope.ui) return;
                    if ($scope.ui.mode !== 'settings') return;
                    $scope.ui.settingsBaselineRaw = JSON.stringify($scope.config || {});
                    $scope.ui.settingsDirty = false;
                } catch (e) {
                    // ignore
                }
            }

            $scope.markSettingsDirty = function () {
                try {
                    if ($scope.ui && $scope.ui.mode === 'settings') {
                        $scope.ui.settingsDirty = true;
                    }
                } catch (e) {
                    // ignore
                }
            };

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

            function getDefaultTasksFolderExisting() {
                try {
                    return outlook && outlook.getTaskFolderExisting ? outlook.getTaskFolderExisting($scope.filter.mailbox, '') : null;
                } catch (e) {
                    return null;
                }
            }

            function loadProjects() {
                try {
                    var projects = [];
                    var hidden = ($scope.config.PROJECTS.hiddenProjectEntryIDs || []);

                    var defaultTasksEntryID = '';
                    var defaultTasksStoreID = '';
                    try {
                        var tf = getDefaultTasksFolderExisting();
                        if (tf) {
                            defaultTasksEntryID = tf.EntryID;
                            try { defaultTasksStoreID = tf.StoreID; } catch (e0b) { defaultTasksStoreID = ''; }

                            // Always include the mailbox default Tasks folder so projects are optional.
                            var tasksName = (function () { try { return String(tf.Name || 'Tasks'); } catch (e0c) { return 'Tasks'; } })();
                            try {
                                if ($scope.config && $scope.config.MULTI_MAILBOX && $scope.filter.mailbox) {
                                    tasksName = tasksName + ' (' + String($scope.filter.mailbox) + ')';
                                }
                            } catch (e0d) {
                                // ignore
                            }
                            projects.push({
                                name: tasksName,
                                entryID: defaultTasksEntryID,
                                storeID: defaultTasksStoreID,
                                isLinked: false,
                                group: 'Default Tasks'
                            });
                        }
                    } catch (e0) {
                        defaultTasksEntryID = '';
                        defaultTasksStoreID = '';
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
                                isLinked: false,
                                group: 'Projects'
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
                            isLinked: true,
                            group: 'Linked'
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
                        p.isDefaultTasks = (defaultTasksEntryID && p.entryID === defaultTasksEntryID);
                        p.isHidden = (!p.isDefaultTasks && (hidden.indexOf(p.entryID) !== -1));

                        // Grouping for dropdown
                        if (p.isDefaultTasks) {
                            p.group = 'Default Tasks';
                        } else if (p.isLinked) {
                            p.group = 'Linked';
                        } else {
                            p.group = 'Projects';
                        }
                    });

                    // Sort by name
                    uniq.sort(function (a, b) {
                        if (a.isDefaultTasks && !b.isDefaultTasks) return -1;
                        if (!a.isDefaultTasks && b.isDefaultTasks) return 1;
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
                        var fallback = '';
                        for (var j = 0; j < $scope.projects.length; j++) {
                            if ($scope.projects[j].isDefaultTasks) {
                                fallback = $scope.projects[j].entryID;
                                break;
                            }
                        }
                        if (!fallback && $scope.projects.length > 0) {
                            fallback = $scope.projects[0].entryID;
                        }
                        if (fallback) {
                            $scope.config.PROJECTS.defaultProjectEntryID = fallback;
                            saveConfig();
                        }
                    }
                } catch (e) {
                    writeLog('loadProjects: ' + e);
                }
            }

            function ensureSelectedProject() {
                var before = '';
                try { before = String($scope.ui.projectEntryID || ''); } catch (e0) { before = ''; }
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
                    if (before !== String($scope.ui.projectEntryID || '')) {
                        scheduleSaveState();
                    }
                    return;
                }

                $scope.ui.projectEntryID = $scope.projects[0].entryID;
                if (before !== String($scope.ui.projectEntryID || '')) {
                    scheduleSaveState();
                }
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
                    if (p.isDefaultTasks) {
                        var tf = getDefaultTasksFolderExisting();
                        if (tf) return tf;
                    }
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
                            var dueDaysFromToday = null;
                            if (isRealDate(due)) {
                                var due0 = new Date(due);
                                due0.setHours(0, 0, 0, 0);
                                dueText = moment(due).format($scope.config.DATE_FORMAT || 'DD-MMM');
                                dueMs = due.getTime();

                                try {
                                    dueDaysFromToday = moment(due0).diff(moment(today0), 'days');
                                } catch (e0d) {
                                    dueDaysFromToday = Math.round((due0.getTime() - today0.getTime()) / (24 * 60 * 60 * 1000));
                                }

                                // Due-state color (ignore completed tasks)
                                if (task.Status !== 2) {
                                    if (dueDaysFromToday !== null && dueDaysFromToday < 0) {
                                        dueClass = 'kfo-due--overdue';
                                    } else if (dueDaysFromToday !== null && dueDaysFromToday === 0) {
                                        dueClass = 'kfo-due--today';
                                    } else {
                                        if (dueDaysFromToday !== null && dueDaysFromToday <= 2) {
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
                                dueDaysFromToday: dueDaysFromToday,
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
                    // Normalize whitespace-only searches (avoid invisible "active" search values).
                    try {
                        if ($scope.filter) {
                            var raw = String($scope.filter.search || '');
                            if (raw !== '') {
                                var trimmed = raw.replace(/^\s+|\s+$/g, '');
                                if (trimmed === '') {
                                    $scope.filter.search = '';
                                }
                            }
                        }
                    } catch (e0) {
                        // ignore
                    }

                    var filtersActive = false;
                    if (board && board.applyFilters) {
                        filtersActive = board.applyFilters($scope.lanes, $scope.filter, $scope.privacyFilter);
                    }

                    $scope.ui.filtersActive = !!filtersActive;

                    // To avoid persisting partial ordering, disable drag/drop while filters are active.
                    updateSortableDisabled(!!filtersActive);

                    scheduleSaveState();
                } catch (e) {
                    writeLog('applyFilters: ' + e);
                }
            };

            function updateSortableDisabled(disabled) {
                try {
                    if ($scope.sortableOptions) {
                        $scope.sortableOptions.disabled = !!disabled;
                    }

                    // Best-effort: update existing jQuery UI sortable instances (IE11/Outlook can miss option updates).
                    try {
                        if (window.jQuery && window.jQuery.fn && window.jQuery.fn.sortable) {
                            window.jQuery('.kfo-tasklist').sortable('option', 'disabled', !!disabled);
                        }
                    } catch (e1) {
                        // ignore
                    }
                } catch (e) {
                    // ignore
                }
            }

            function applySortableInteractionConfig() {
                try {
                    if (!$scope.sortableOptions) return;

                    // Never start a drag from interactive elements.
                    var cancel = 'button, input, textarea, select, a, .kfo-iconBtn, .kfo-tag';
                    $scope.sortableOptions.cancel = cancel;

                    if ($scope.config && $scope.config.BOARD && $scope.config.BOARD.dragHandleOnly) {
                        $scope.sortableOptions.handle = '.kfo-dragHandle';
                    } else {
                        $scope.sortableOptions.handle = false;
                    }

                    // Best-effort: update existing jQuery UI sortable instances.
                    try {
                        if (window.jQuery && window.jQuery.fn && window.jQuery.fn.sortable) {
                            window.jQuery('.kfo-tasklist').sortable('option', 'cancel', cancel);
                            window.jQuery('.kfo-tasklist').sortable('option', 'handle', ($scope.sortableOptions.handle || false));
                        }
                    } catch (e1) {
                        // ignore
                    }
                } catch (e) {
                    // ignore
                }
            }

            // Expose for Settings toggles (ng-change)
            $scope.applySortableInteractionConfig = function () {
                applySortableInteractionConfig();
            };

            function boardCounts() {
                var totalAll = 0;
                var totalEnabled = 0;
                var filteredEnabled = 0;
                try {
                    ($scope.lanes || []).forEach(function (lane) {
                        if (!lane) return;
                        totalAll += (lane.tasks || []).length;
                        if (lane.enabled === false) return;
                        totalEnabled += (lane.tasks || []).length;
                        filteredEnabled += (lane.filteredTasks || []).length;
                    });
                } catch (e) {
                    // ignore
                }
                return { totalAll: totalAll, totalEnabled: totalEnabled, filteredEnabled: filteredEnabled };
            }

            $scope.isBoardEmpty = function () {
                return boardCounts().totalAll === 0;
            };

            $scope.isBoardNoMatch = function () {
                var c = boardCounts();
                return !!($scope.ui && $scope.ui.filtersActive) && c.totalEnabled > 0 && c.filteredEnabled === 0;
            };

            $scope.isBoardHiddenByLanes = function () {
                var c = boardCounts();
                return c.totalAll > 0 && c.totalEnabled === 0;
            };

            $scope.enableAllLanes = function () {
                try {
                    if (!$scope.config || !$scope.config.LANES) return;
                    for (var i = 0; i < $scope.config.LANES.length; i++) {
                        $scope.config.LANES[i].enabled = true;
                    }
                    saveConfig();
                    rebuildLaneOptions();
                    showToast('success', 'Lanes enabled', 'All lanes are now visible', 2000);
                    $scope.refreshTasks();
                } catch (e) {
                    writeLog('enableAllLanes: ' + e);
                }
            };

            $scope.clearFilters = function () {
                try {
                    $scope.filter.search = '';
                    $scope.filter.category = '<All Categories>';
                    $scope.filter.private = $scope.privacyFilter.all.value;
                    $scope.filter.due = 'any';
                    $scope.applyFilters();
                } catch (e) {
                    writeLog('clearFilters: ' + e);
                }
            };

            $scope.clearSearchFilter = function () {
                try {
                    $scope.filter.search = '';
                    $scope.applyFilters();
                } catch (e) {
                    // ignore
                }
            };

            $scope.clearCategoryFilter = function () {
                try {
                    $scope.filter.category = '<All Categories>';
                    $scope.applyFilters();
                } catch (e) {
                    // ignore
                }
            };

            $scope.clearDueFilter = function () {
                try {
                    $scope.filter.due = 'any';
                    $scope.applyFilters();
                } catch (e) {
                    // ignore
                }
            };

            $scope.clearPrivacyFilter = function () {
                try {
                    $scope.filter.private = $scope.privacyFilter.all.value;
                    $scope.applyFilters();
                } catch (e) {
                    // ignore
                }
            };

            $scope.privacyFilterLabel = function (value) {
                try {
                    var v = String(value);
                    if (v === String($scope.privacyFilter.private.value)) return $scope.privacyFilter.private.text;
                    if (v === String($scope.privacyFilter.public.value)) return $scope.privacyFilter.public.text;
                    return $scope.privacyFilter.all.text;
                } catch (e) {
                    return '';
                }
            };

            $scope.dueFilterLabel = function (value) {
                try {
                    var v2 = String(value || 'any');
                    if (v2 === 'overdue') return 'Overdue';
                    if (v2 === 'today') return 'Today';
                    if (v2 === 'next7') return 'Next 7 days';
                    if (v2 === 'nodue') return 'No due date';
                    return 'Any';
                } catch (e) {
                    return '';
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
                        // If the selected project folder is missing/unavailable, fall back to the default Tasks folder.
                        try {
                            var tf = getDefaultTasksFolderExisting();
                            if (tf) {
                                if ($scope.ui && $scope.ui.projectEntryID !== tf.EntryID) {
                                    $scope.ui.projectEntryID = tf.EntryID;
                                    scheduleSaveState();
                                    if (!projectFallbackNotified) {
                                        projectFallbackNotified = true;
                                        showToast('info', 'Project unavailable', 'Switched to Outlook Tasks', 2800);
                                    }
                                    showBanner('warning', 'Project unavailable', 'The selected folder could not be opened. Switched to Outlook Tasks. You can pick another folder or relink in Settings -> Projects.');
                                }
                                folder = tf;
                                perf.projectEntryID = tf.EntryID;
                            }
                        } catch (e0) {
                            // ignore
                        }
                    }

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
                            try {
                                var d = new Date();
                                $scope.ui.lastRefreshedAtMs = d.getTime();
                                $scope.ui.lastRefreshedAtText = moment(d).format('HH:mm');
                            } catch (e1) {
                                $scope.ui.lastRefreshedAtMs = 0;
                                $scope.ui.lastRefreshedAtText = '';
                            }
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
                try {
                    // Leaving Settings: confirm and discard unsaved changes.
                    if ($scope.ui && $scope.ui.mode === 'settings' && mode !== 'settings' && $scope.ui.settingsDirty) {
                        var discard = window.confirm('You have unsaved changes in Settings. Discard changes and leave Settings?');
                        if (!discard) {
                            return;
                        }
                        try {
                            if ($scope.ui.settingsBaselineRaw) {
                                $scope.config = JSON.parse($scope.ui.settingsBaselineRaw);
                                rebuildLaneOptions();
                                rebuildThemeList();
                                applySortableInteractionConfig();
                                $scope.applyTheme();
                            }
                        } catch (e0) {
                            // ignore
                        }
                        $scope.ui.settingsDirty = false;
                    }
                } catch (e1) {
                    // ignore
                }

                $scope.ui.mode = mode;

                if (mode === 'board') {
                    $scope.applyFilters();
                }
                if (mode === 'settings') {
                    try {
                        $scope.ui.settingsBaselineRaw = JSON.stringify($scope.config || {});
                        $scope.ui.settingsDirty = false;
                        $scope.validateLanes();
                        $scope.validateNewLaneDraft();
                    } catch (e2) {
                        // ignore
                    }
                }
            };

            $scope.brandKeyDown = function (ev) {
                try {
                    var e = ev || window.event;
                    var code = e ? (e.keyCode || e.which) : 0;

                    // Enter / Space
                    if (code === 13 || code === 32) {
                        try { if (e.preventDefault) e.preventDefault(); } catch (e1) { /* ignore */ }
                        $scope.switchMode('board');
                        return false;
                    }
                } catch (e) {
                    // ignore
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

            function updateTaskLaneAndStatus(taskEntryID, storeID, laneId, statusValue, setStatus) {
                var result = { ok: false, beforeStatus: null, statusChanged: false };
                try {
                    var taskitem = getTaskItemSafe(taskEntryID, storeID);
                    if (!taskitem) {
                        reportError('updateTaskLaneAndStatus', 'task not available', 'Update failed', 'Could not update the task in Outlook. Click the ! icon for details.');
                        return result;
                    }
                    if (!(outlook && outlook.setUserProperty)) {
                        reportError('updateTaskLaneAndStatus', 'Outlook adapter not available', 'Update failed', 'Outlook integration is not available. Click the ! icon for details.');
                        return result;
                    }

                    try { result.beforeStatus = taskitem.Status; } catch (e0) { result.beforeStatus = null; }

                    outlook.setUserProperty(taskitem, PROP_LANE_ID, laneId, OlUserPropertyType.olText);

                    if (setStatus && statusValue !== null && statusValue !== undefined) {
                        try {
                            if (taskitem.Status != statusValue) {
                                taskitem.Status = statusValue;
                                result.statusChanged = true;
                            }
                        } catch (e1) {
                            // ignore
                        }
                    }

                    taskitem.Save();
                    result.ok = true;
                } catch (e) {
                    reportError('updateTaskLaneAndStatus', e, 'Update failed', 'Could not update the task in Outlook. Click the ! icon for details.');
                }
                return result;
            }

            function fixLaneOrder(lane) {
                try {
                    if (!$scope.config.BOARD.saveOrder) return;
                    if ($scope.ui && $scope.ui.filtersActive) {
                        return;
                    }
                    if (!(outlook && outlook.setUserProperty)) {
                        reportError('fixLaneOrder', 'Outlook adapter not available', 'Ordering failed', 'Outlook integration is not available. Click the ! icon for details.');
                        return;
                    }
                    // Do not persist partial ordering.
                    try {
                        var allCount = (lane && lane.tasks) ? lane.tasks.length : 0;
                        var visCount = (lane && lane.filteredTasks) ? lane.filteredTasks.length : 0;
                        if (allCount !== visCount) {
                            writeLog('fixLaneOrder skipped (filtered list)');
                            return;
                        }
                    } catch (e0) {
                        // ignore
                    }

                    // COM write reduction: only update tasks whose stored order differs.
                    var updates = [];
                    for (var i = 0; i < lane.filteredTasks.length; i++) {
                        var t = lane.filteredTasks[i];
                        if (!t || !t.entryID) continue;
                        var current = (t.laneOrder === undefined || t.laneOrder === null) ? null : t.laneOrder;
                        var cn = null;
                        try {
                            cn = (current === null) ? null : parseInt(current, 10);
                            if (isNaN(cn)) cn = null;
                        } catch (eParse) {
                            cn = null;
                        }
                        if (cn !== i) {
                            updates.push({ idx: i, task: t });
                        }
                    }

                    if (updates.length === 0) {
                        return;
                    }

                    for (var j = 0; j < updates.length; j++) {
                        var u = updates[j];
                        var t2 = u.task;
                        var taskitem = getTaskItemSafe(t2.entryID, t2.storeID);
                        if (!taskitem) {
                            continue;
                        }
                        outlook.setUserProperty(taskitem, PROP_LANE_ORDER, u.idx, OlUserPropertyType.olNumber);
                        taskitem.Save();
                        try { t2.laneOrder = u.idx; } catch (eSet) { /* ignore */ }
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

                        // Defensive: prevent reordering while filters are active (partial ordering would be persisted).
                        if ($scope.ui && $scope.ui.filtersActive) {
                            try {
                                showToast('info', 'Reordering disabled', 'Clear filters to reorder tasks', 2600);
                            } catch (e0) {
                                // ignore
                            }
                            try { ui.item.sortable.cancel(); } catch (e0b) { /* ignore */ }
                            return;
                        }
                        var fromLaneId = ui.item.sortable.source.attr('data-lane-id');
                        var toLaneId = ui.item.sortable.droptarget.attr('data-lane-id');
                        if (!fromLaneId || !toLaneId) {
                            return;
                        }

                        // If manual ordering is off, do not allow reordering within a lane.
                        if (fromLaneId === toLaneId && !($scope.config && $scope.config.BOARD && $scope.config.BOARD.saveOrder)) {
                            showToast('info', 'Manual ordering is off', 'Enable it in Settings -> Board to reorder tasks within a lane', 3200);
                            try { ui.item.sortable.cancel(); } catch (e0c) { /* ignore */ }
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

                        // Update lane assignment first (and optional status sync), then offer Undo.
                        if (fromLaneId !== toLaneId) {
                            var syncStatus = !!($scope.config && $scope.config.AUTOMATION && $scope.config.AUTOMATION.setOutlookStatusOnLaneMove);
                            var resMove = updateTaskLaneAndStatus(model.entryID, model.storeID, toLaneId, toLane.outlookStatus, syncStatus);
                            if (!resMove || !resMove.ok) {
                                try { ui.item.sortable.cancel(); } catch (e0c) { /* ignore */ }
                                $scope.refreshTasks();
                                return;
                            }

                            // Capture for one-step undo (lane moves only)
                            try {
                                $scope.ui.lastMove = {
                                    entryID: model.entryID,
                                    storeID: model.storeID,
                                    fromLaneId: fromLaneId,
                                    toLaneId: toLaneId,
                                    restoreStatus: !!resMove.statusChanged,
                                    beforeStatusValue: resMove.beforeStatus,
                                    atMs: (new Date()).getTime()
                                };
                                showToast('info', 'Moved', 'To: ' + String(toLane.title || toLaneId), 6500, 'Undo', function () { $scope.undoLastMove(); });
                            } catch (e0m) {
                                // ignore
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

            // Quick add (no inspector)
            $scope.toggleQuickAdd = function (lane) {
                try {
                    if (!$scope.ui) return;
                    if (!lane || !lane.id) return;

                    if ($scope.ui.quickAddLaneId === lane.id) {
                        $scope.closeQuickAdd();
                        return;
                    }

                    openQuickAddForLane(lane);
                } catch (e) {
                    writeLog('toggleQuickAdd: ' + e);
                }
            };

            $scope.closeQuickAdd = function () {
                try {
                    if (!$scope.ui) return;
                    $scope.ui.quickAddLaneId = '';
                    $scope.ui.quickAddText = '';
                    $scope.ui.quickAddSaving = false;
                } catch (e) {
                    // ignore
                }
            };

            $scope.quickAddKeyDown = function (ev, lane) {
                try {
                    var e = ev || window.event;
                    var code = e ? (e.keyCode || e.which) : 0;

                    // Enter
                    if (code === 13) {
                        try { if (e.preventDefault) e.preventDefault(); } catch (e1) { /* ignore */ }
                        $scope.submitQuickAdd(lane);
                        return false;
                    }

                    // Esc
                    if (code === 27) {
                        try { if (e.preventDefault) e.preventDefault(); } catch (e2) { /* ignore */ }
                        $scope.closeQuickAdd();
                        return false;
                    }
                } catch (e) {
                    // ignore
                }
            };

            $scope.submitQuickAdd = function (lane) {
                try {
                    if (!$scope.ui || $scope.ui.quickAddSaving) return;

                    var subject = String($scope.ui.quickAddText || '');
                    subject = subject.replace(/^\s+|\s+$/g, '');
                    if (!subject) {
                        showUserError('Quick add', 'Enter a task subject first.');
                        return;
                    }

                    var folder = getSelectedProjectFolder();
                    if (!folder) {
                        // Fallback: allow using the default Tasks folder even when no project is selected.
                        folder = getDefaultTasksFolderExisting();
                        try {
                            if (folder && folder.EntryID) {
                                $scope.ui.projectEntryID = folder.EntryID;
                            }
                        } catch (e0) {
                            // ignore
                        }
                    }

                    if (!folder) {
                        showUserError('Tasks folder not available', 'Could not access your Outlook Tasks folder.');
                        return;
                    }

                    $scope.ui.quickAddSaving = true;

                    var taskitem = folder.Items.Add();
                    taskitem.Subject = subject;

                    // Default sensitivity based on current filter
                    if ($scope.filter.private == $scope.privacyFilter.private.value) {
                        taskitem.Sensitivity = SENSITIVITY.olPrivate;
                    }

                    if (lane && lane.id) {
                        if (outlook && outlook.setUserProperty) {
                            outlook.setUserProperty(taskitem, PROP_LANE_ID, lane.id, OlUserPropertyType.olText);

                            // Lane order: place new tasks at the top without rewriting all tasks.
                            var order = suggestedNewLaneOrder(lane);
                            if (order !== null && order !== undefined) {
                                outlook.setUserProperty(taskitem, PROP_LANE_ORDER, order, OlUserPropertyType.olNumber);
                            }
                        } else {
                            // Allow task creation to proceed even if lane metadata cannot be stored.
                            reportError('quickAdd', 'Outlook adapter not available', 'Lane not set', 'The task was created but could not be placed on a lane. Click the ! icon for details.');
                        }
                        if ($scope.config && $scope.config.AUTOMATION && $scope.config.AUTOMATION.setOutlookStatusOnLaneMove) {
                            if (lane.outlookStatus !== null && lane.outlookStatus !== undefined) {
                                taskitem.Status = lane.outlookStatus;
                            }
                        }
                    }

                    taskitem.Save();

                    $scope.ui.quickAddText = '';
                    $scope.ui.quickAddSaving = false;
                    showToast('success', 'Task added', subject, 1400);

                    // Refresh to pick up EntryID + Outlook-calculated fields
                    $scope.refreshTasks();

                    // Re-focus input for rapid entry
                    $timeout(function () {
                        try {
                            if ($scope.ui && $scope.ui.quickAddLaneId && lane && lane.id && $scope.ui.quickAddLaneId === lane.id) {
                                var el = document.getElementById('kfo-quickadd-' + lane.id);
                                if (el && el.focus) el.focus();
                            }
                        } catch (e3) {
                            // ignore
                        }
                    }, 0);
                } catch (e) {
                    try { if ($scope.ui) $scope.ui.quickAddSaving = false; } catch (e1) { /* ignore */ }
                    writeLog('quickAdd: ' + e);
                    reportError('quickAdd', e, 'Add task failed', 'Could not create a new task in Outlook. Click the ! icon for details.');
                }
            };

            function preferredNewTaskLaneId() {
                // Prefer Backlog when enabled, else first enabled lane.
                var firstEnabled = '';
                var backlogEnabled = '';
                try {
                    ($scope.config.LANES || []).forEach(function (l) {
                        var id = sanitizeId(l.id);
                        if (!id) return;
                        if (l.enabled === false) return;
                        if (!firstEnabled) firstEnabled = id;
                        if (id === 'backlog') backlogEnabled = id;
                    });
                } catch (e) {
                    // ignore
                }
                return backlogEnabled || firstEnabled || '';
            }

            function getLaneConfigById(laneId) {
                try {
                    var id = sanitizeId(laneId);
                    if (!id) return null;
                    for (var i = 0; i < ($scope.config.LANES || []).length; i++) {
                        var l = $scope.config.LANES[i];
                        if (sanitizeId(l.id) === id) return l;
                    }
                } catch (e) {
                    // ignore
                }
                return null;
            }

            function suggestedNewLaneOrder(lane) {
                try {
                    if (!$scope.config || !$scope.config.BOARD || !$scope.config.BOARD.saveOrder) {
                        return null;
                    }
                    var min = null;
                    var list = [];
                    try {
                        list = (lane && lane.tasks) ? lane.tasks : ((lane && lane.filteredTasks) ? lane.filteredTasks : []);
                    } catch (e0) {
                        list = [];
                    }
                    for (var i = 0; i < list.length; i++) {
                        var t = list[i];
                        var raw = (t && t.laneOrder !== undefined && t.laneOrder !== null) ? t.laneOrder : null;
                        if (raw === null) continue;
                        var n = parseInt(raw, 10);
                        if (isNaN(n)) continue;
                        if (min === null || n < min) {
                            min = n;
                        }
                    }
                    if (min === null) return 0;
                    return min - 1;
                } catch (e) {
                    return 0;
                }
            }

            function scrollLaneIntoView(laneId) {
                try {
                    var id = sanitizeId(laneId);
                    if (!id) return;
                    var el = document.getElementById('kfo-lane-' + id);
                    if (el && el.scrollIntoView) {
                        // IE11-friendly: bring lane into view in the scroll container.
                        el.scrollIntoView(true);
                    }
                } catch (e) {
                    // ignore
                }
            }

            function openQuickAddForLane(lane) {
                try {
                    if (!$scope.ui) return;
                    if (!lane || !lane.id) return;

                    scrollLaneIntoView(lane.id);

                    $scope.ui.quickAddLaneId = lane.id;
                    $scope.ui.quickAddText = '';
                    $scope.ui.quickAddSaving = false;

                    $timeout(function () {
                        try {
                            var el = document.getElementById('kfo-quickadd-' + lane.id);
                            if (el && el.focus) {
                                el.focus();
                                try { el.select(); } catch (e1) { /* ignore */ }
                            }
                        } catch (e2) {
                            // ignore
                        }
                    }, 0);
                } catch (e) {
                    // ignore
                }
            }

            $scope.openQuickAddBacklog = function () {
                try {
                    var laneId = preferredNewTaskLaneId();
                    var lane = getLaneById(laneId);

                    if (!lane || lane.enabled === false) {
                        // Fall back to first enabled lane in the current board.
                        for (var i = 0; i < ($scope.lanes || []).length; i++) {
                            if ($scope.lanes[i] && $scope.lanes[i].enabled !== false) {
                                lane = $scope.lanes[i];
                                break;
                            }
                        }
                    }

                    if (lane) {
                        openQuickAddForLane(lane);
                    }
                } catch (e) {
                    writeLog('openQuickAddBacklog: ' + e);
                }
            };

            $scope.openNewTask = function () {
                try {
                    $scope.switchMode('board');
                    if ($scope.ui && $scope.ui.filtersActive) {
                        showToast('info', 'Filters active', 'A new task may be hidden by filters', 2600);
                    }

                    // If quick add is disabled, fall back to the Outlook inspector.
                    if ($scope.config && $scope.config.BOARD && $scope.config.BOARD.quickAddEnabled === false) {
                        showToast('info', 'Quick add disabled', 'Opening Outlook editor', 2400);
                        $timeout(function () {
                            try {
                                var laneId = preferredNewTaskLaneId();
                                var lane = getLaneById(laneId);
                                if (!lane && laneId) {
                                    var cfg = getLaneConfigById(laneId);
                                    lane = { id: laneId, outlookStatus: cfg ? cfg.outlookStatus : null };
                                }
                                $scope.addTask(lane);
                            } catch (e0) {
                                // ignore
                            }
                        }, 0);
                        return;
                    }

                    $timeout(function () {
                        try {
                            $scope.openQuickAddBacklog();
                        } catch (e1) {
                            // ignore
                        }
                    }, 0);
                } catch (e) {
                    writeLog('openNewTask: ' + e);
                }
            };

            $scope.addTask = function (lane) {
                try {
                    var folder = getSelectedProjectFolder();
                    if (!folder) {
                        folder = getDefaultTasksFolderExisting();
                        try {
                            if (folder && folder.EntryID) {
                                $scope.ui.projectEntryID = folder.EntryID;
                            }
                        } catch (e0) {
                            // ignore
                        }
                    }

                    if (!folder) {
                        showUserError('Tasks folder not available', 'Could not access your Outlook Tasks folder.');
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
                            var order = suggestedNewLaneOrder(lane);
                            if (order !== null && order !== undefined) {
                                outlook.setUserProperty(taskitem, PROP_LANE_ORDER, order, OlUserPropertyType.olNumber);
                            }
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
                        window.open(url, '_blank');
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
            $scope.validateLanes = function () {
                try {
                    if (!$scope.ui) return {};
                    var errs = {};

                    var ids = {};
                    var dup = {};
                    for (var i = 0; i < ($scope.config.LANES || []).length; i++) {
                        var l = $scope.config.LANES[i];
                        var id = sanitizeId(l ? l.id : '');
                        if (!id) {
                            errs[i] = errs[i] || {};
                            errs[i].id = 'Id is required.';
                            continue;
                        }
                        if (ids[id]) {
                            dup[id] = true;
                        }
                        ids[id] = true;
                    }

                    for (var j = 0; j < ($scope.config.LANES || []).length; j++) {
                        var lane = $scope.config.LANES[j];
                        if (!lane) continue;
                        var sid = sanitizeId(lane.id);
                        if (sid && dup[sid]) {
                            errs[j] = errs[j] || {};
                            errs[j].id = 'Duplicate id.';
                        }

                        var title = String(lane.title || '').trim();
                        if (!title) {
                            errs[j] = errs[j] || {};
                            errs[j].title = 'Title is required.';
                        }

                        var color = String(lane.color || '').trim();
                        if (color && !isValidHexColor(color)) {
                            errs[j] = errs[j] || {};
                            errs[j].color = 'Use #RRGGBB.';
                        }

                        var w = 0;
                        try { w = parseInt(lane.wipLimit, 10); } catch (e0) { w = 0; }
                        if (isNaN(w) || w < 0) {
                            errs[j] = errs[j] || {};
                            errs[j].wipLimit = 'Must be 0 or greater.';
                        }
                    }

                    $scope.ui.laneErrors = errs;
                    return errs;
                } catch (e) {
                    try { $scope.ui.laneErrors = {}; } catch (e1) { /* ignore */ }
                    return {};
                }
            };

            $scope.hasLaneErrors = function () {
                try {
                    var e = ($scope.ui && $scope.ui.laneErrors) ? $scope.ui.laneErrors : {};
                    for (var k in e) {
                        if (e.hasOwnProperty(k)) return true;
                    }
                } catch (e2) {
                    // ignore
                }
                return false;
            };

            $scope.validateNewLaneDraft = function () {
                try {
                    if (!$scope.ui) return {};
                    var title = String($scope.ui.newLaneTitle || '').trim();
                    var id = sanitizeId($scope.ui.newLaneId || title);
                    var color = String($scope.ui.newLaneColor || '').trim();
                    var out = {};
                    if (!title) out.title = 'Title is required.';
                    if (!id) out.id = 'Id is required.';
                    if (color && !isValidHexColor(color)) out.color = 'Use #RRGGBB.';
                    if (id) {
                        for (var i = 0; i < ($scope.config.LANES || []).length; i++) {
                            if (sanitizeId($scope.config.LANES[i].id) === id) {
                                out.id = 'Id already exists.';
                            }
                        }
                    }
                    $scope.ui.newLaneErrors = out;
                    return out;
                } catch (e) {
                    try { $scope.ui.newLaneErrors = {}; } catch (e1) { /* ignore */ }
                    return {};
                }
            };

            function enabledLaneCount() {
                var n = 0;
                try {
                    ($scope.config.LANES || []).forEach(function (l) {
                        if (!l) return;
                        if (l.enabled === false) return;
                        n++;
                    });
                } catch (e) {
                    // ignore
                }
                return n;
            }

            function ensureAtLeastOneLaneEnabled() {
                try {
                    if (!$scope.config || !$scope.config.LANES) return;
                    if ($scope.config.LANES.length === 0) return;
                    if (enabledLaneCount() > 0) return;
                    // Re-enable the first lane if all are disabled.
                    $scope.config.LANES[0].enabled = true;
                } catch (e) {
                    // ignore
                }
            }

            $scope.onLaneEnabledChanged = function (lane) {
                try {
                    if (enabledLaneCount() === 0) {
                        // Revert the change so the board is never left without a visible lane.
                        if (lane) {
                            lane.enabled = true;
                        } else {
                            ensureAtLeastOneLaneEnabled();
                        }
                        showUserError('Lane required', 'At least one lane must remain enabled.');
                    }
                } catch (e) {
                    writeLog('onLaneEnabledChanged: ' + e);
                }
            };

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
                if ($scope.config.LANES.length <= 1) {
                    showUserError('Cannot remove lane', 'At least one lane must remain on the board.');
                    return;
                }
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

            // Lane id migration tool
            function resetLaneIdToolProgress() {
                try {
                    $scope.ui.laneIdTool.progress = {
                        foldersTotal: 0,
                        foldersDone: 0,
                        tasksTotal: 0,
                        tasksDone: 0,
                        updated: 0,
                        errors: 0
                    };
                } catch (e) {
                    // ignore
                }
            }

            $scope.openLaneIdTool = function (lane) {
                try {
                    if (!$scope.ui || !$scope.ui.laneIdTool) return;
                    if (!lane || !lane.id) return;

                    $scope.ui.laneIdTool.oldId = sanitizeId(lane.id) || String(lane.id || '');
                    $scope.ui.laneIdTool.laneTitle = String(lane.title || lane.id || 'Lane');
                    $scope.ui.laneIdTool.newId = $scope.ui.laneIdTool.oldId;
                    $scope.ui.laneIdTool.scope = 'all';
                    $scope.ui.laneIdTool.running = false;
                    $scope.ui.laneIdTool.cancelRequested = false;
                    resetLaneIdToolProgress();
                    $scope.ui.showLaneIdTool = true;
                    focusById('kfo-modal-laneId-newId', true);
                } catch (e) {
                    writeLog('openLaneIdTool: ' + e);
                }
            };

            $scope.laneIdToolKeyDown = function (ev) {
                try {
                    var e = ev || window.event;
                    var code = e ? (e.keyCode || e.which) : 0;
                    if (code === 13) {
                        try { if (e.preventDefault) e.preventDefault(); } catch (e1) { /* ignore */ }
                        $scope.runLaneIdTool();
                        return false;
                    }
                } catch (e) {
                    // ignore
                }
            };

            $scope.closeLaneIdTool = function () {
                try {
                    if ($scope.ui && $scope.ui.laneIdTool && $scope.ui.laneIdTool.running) {
                        showUserError('In progress', 'Please cancel the migration first.');
                        return;
                    }
                    if ($scope.ui) {
                        $scope.ui.showLaneIdTool = false;
                    }
                } catch (e) {
                    // ignore
                }
            };

            $scope.cancelLaneIdTool = function () {
                try {
                    if ($scope.ui && $scope.ui.laneIdTool) {
                        $scope.ui.laneIdTool.cancelRequested = true;
                    }
                } catch (e) {
                    // ignore
                }
            };

            $scope.runLaneIdTool = function () {
                try {
                    if (!$scope.ui || !$scope.ui.laneIdTool) return;
                    if ($scope.ui.laneIdTool.running) return;
                    if (!(outlook && outlook.getUserProperty && outlook.setUserProperty)) {
                        showUserError('Outlook integration not available', 'Cannot migrate lane ids in this host.');
                        return;
                    }

                    var oldId = sanitizeId($scope.ui.laneIdTool.oldId);
                    var newId = sanitizeId($scope.ui.laneIdTool.newId);
                    if (!oldId || !newId) {
                        showUserError('Lane id required', 'Enter a valid new id (letters, numbers, dashes).');
                        return;
                    }
                    if (oldId === newId) {
                        showUserError('No change', 'The new id is the same as the current id.');
                        return;
                    }

                    // Uniqueness check
                    for (var i = 0; i < ($scope.config.LANES || []).length; i++) {
                        var existing = sanitizeId($scope.config.LANES[i].id);
                        if (existing === newId) {
                            showUserError('Lane id already exists', 'Choose a different id.');
                            return;
                        }
                    }

                    if (!window.confirm('Change lane id from "' + oldId + '" to "' + newId + '"? This will update tasks and then update your lane configuration.')) {
                        return;
                    }

                    // Backup current settings before touching tasks/config.
                    try { backupCurrentSettingsToOutlook(); } catch (e0) { /* ignore */ }

                    $scope.ui.laneIdTool.running = true;
                    $scope.ui.laneIdTool.cancelRequested = false;
                    resetLaneIdToolProgress();

                    // Build folder list
                    var folders = [];
                    if ($scope.ui.laneIdTool.scope === 'current') {
                        var current = getSelectedProjectFolder();
                        if (current) {
                            folders.push({ name: 'Current folder', folder: current });
                        }
                    } else {
                        // All known projects in this mailbox (includes default Tasks, root projects, and linked).
                        ($scope.projectsAll || []).forEach(function (p) {
                            try {
                                var f = null;
                                if (p && p.isDefaultTasks) {
                                    f = getDefaultTasksFolderExisting();
                                } else if (p && p.entryID) {
                                    f = outlook && outlook.getFolderFromIDs ? outlook.getFolderFromIDs(p.entryID, p.storeID) : null;
                                }
                                if (f) {
                                    folders.push({ name: p.name || 'Folder', folder: f });
                                }
                            } catch (e1) {
                                // ignore
                            }
                        });
                    }

                    if (folders.length === 0) {
                        $scope.ui.laneIdTool.running = false;
                        showUserError('No folders', 'No Tasks folders were found to migrate.');
                        return;
                    }

                    // Pre-calc task totals (best-effort)
                    var totalTasks = 0;
                    folders.forEach(function (x) {
                        try {
                            if (x && x.folder && x.folder.Items) {
                                totalTasks += x.folder.Items.Count;
                            }
                        } catch (e2) {
                            // ignore
                        }
                    });
                    $scope.ui.laneIdTool.progress.foldersTotal = folders.length;
                    $scope.ui.laneIdTool.progress.tasksTotal = totalTasks;

                    var folderIdx = 0;
                    var itemIdx = 1;
                    var currentFolder = null;
                    var currentItems = null;
                    var currentCount = 0;

                    function nextFolder() {
                        try {
                            if (folderIdx >= folders.length) {
                                return false;
                            }
                            currentFolder = folders[folderIdx].folder;
                            currentItems = currentFolder.Items;
                            currentCount = currentItems.Count;
                            itemIdx = 1;
                            return true;
                        } catch (e) {
                            $scope.ui.laneIdTool.progress.errors++;
                            folderIdx++;
                            return nextFolder();
                        }
                    }

                    nextFolder();

                    function step() {
                        try {
                            if (!$scope.ui || !$scope.ui.laneIdTool || !$scope.ui.laneIdTool.running) return;
                            if ($scope.ui.laneIdTool.cancelRequested) {
                                $scope.ui.laneIdTool.running = false;
                                showToast('info', 'Cancelled', 'Lane id migration cancelled', 2200);
                                return;
                            }

                            var chunk = 18;
                            var processed = 0;
                            while (processed < chunk) {
                                if (!currentFolder || !currentItems) {
                                    break;
                                }
                                if (itemIdx > currentCount) {
                                    // folder done
                                    folderIdx++;
                                    $scope.ui.laneIdTool.progress.foldersDone = folderIdx;
                                    if (!nextFolder()) {
                                        break;
                                    }
                                    continue;
                                }

                                var task = null;
                                try {
                                    task = currentItems(itemIdx);
                                    var lid = outlook.getUserProperty(task, PROP_LANE_ID);
                                    lid = sanitizeId(lid);
                                    if (lid && lid === oldId) {
                                        outlook.setUserProperty(task, PROP_LANE_ID, newId, OlUserPropertyType.olText);
                                        task.Save();
                                        $scope.ui.laneIdTool.progress.updated++;
                                    }
                                } catch (e3) {
                                    $scope.ui.laneIdTool.progress.errors++;
                                } finally {
                                    try { task = null; } catch (e4) { /* ignore */ }
                                }

                                itemIdx++;
                                $scope.ui.laneIdTool.progress.tasksDone++;
                                processed++;
                            }

                            // done?
                            if (folderIdx >= folders.length) {
                                // Update config lane id
                                try {
                                    for (var li = 0; li < ($scope.config.LANES || []).length; li++) {
                                        if (sanitizeId($scope.config.LANES[li].id) === oldId) {
                                            $scope.config.LANES[li].id = newId;
                                        }
                                    }
                                    saveConfig();
                                    rebuildLaneOptions();
                                } catch (e5) {
                                    // ignore
                                }

                                $scope.ui.laneIdTool.running = false;
                                showToast('success', 'Lane id updated', String($scope.ui.laneIdTool.progress.updated) + ' tasks migrated', 2600);
                                $scope.ui.showLaneIdTool = false;
                                $scope.refreshTasks();
                                return;
                            }

                            $timeout(step, 0);
                        } catch (e) {
                            $scope.ui.laneIdTool.running = false;
                            reportError('laneIdTool', e, 'Migration failed', 'Could not migrate lane ids. Click the ! icon for details.');
                        }
                    }

                    $timeout(step, 0);
                } catch (e) {
                    writeLog('runLaneIdTool: ' + e);
                    reportError('runLaneIdTool', e, 'Migration failed', 'Could not migrate lane ids. Click the ! icon for details.');
                    try { if ($scope.ui && $scope.ui.laneIdTool) $scope.ui.laneIdTool.running = false; } catch (e2) { /* ignore */ }
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
                            showUserError('Theme import rejected', 'Themes must be local-only (no http/https, no protocol-relative //, no @import) and must not use scriptable CSS (expression/behaviour or javascript: URLs).');
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
                focusById('kfo-modal-createProject-name', true);
            };

            $scope.openLinkProject = function () {
                $scope.ui.createProjectMode = 'link';
                $scope.ui.linkProjectEntryID = '';
                $scope.ui.newProjectName = '';
                $scope.ui.showCreateProject = true;
                focusById('kfo-modal-createProject-existing', false);
            };

            $scope.createProjectNameKeyDown = function (ev) {
                try {
                    var e = ev || window.event;
                    var code = e ? (e.keyCode || e.which) : 0;
                    if (code === 13) {
                        try { if (e.preventDefault) e.preventDefault(); } catch (e1) { /* ignore */ }
                        $scope.submitCreateProject();
                        return false;
                    }
                } catch (e) {
                    // ignore
                }
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

                    if (p.isDefaultTasks) {
                        showUserError('Cannot hide', 'The default Outlook Tasks folder is always available.');
                        return;
                    }
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
                focusById('kfo-modal-renameProject-name', true);
            };

            $scope.renameProjectNameKeyDown = function (ev) {
                try {
                    var e = ev || window.event;
                    var code = e ? (e.keyCode || e.which) : 0;
                    if (code === 13) {
                        try { if (e.preventDefault) e.preventDefault(); } catch (e1) { /* ignore */ }
                        $scope.submitRenameProject();
                        return false;
                    }
                } catch (e) {
                    // ignore
                }
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
                    focusById('kfo-modal-moveTasks-from', false);
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
                                        var wrote = false;
                                        if (outlook && outlook.setUserProperty) {
                                            if (w.laneId) {
                                                outlook.setUserProperty(moved, PROP_LANE_ID, w.laneId, OlUserPropertyType.olText);
                                                wrote = true;
                                            }
                                            if (w.laneOrder !== null && w.laneOrder !== undefined) {
                                                outlook.setUserProperty(moved, PROP_LANE_ORDER, w.laneOrder, OlUserPropertyType.olNumber);
                                                wrote = true;
                                            }
                                        }
                                        if (wrote) {
                                            moved.Save();
                                        }
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
                        // Setup default project (projects are optional)
                        var rootName = String($scope.config.PROJECTS.rootFolderName || DEFAULT_ROOT_FOLDER_NAME).trim();
                        if (!rootName) rootName = DEFAULT_ROOT_FOLDER_NAME;
                        $scope.config.PROJECTS.rootFolderName = rootName;

                        if ($scope.ui.setupProjectMode === 'default') {
                            var tf = getDefaultTasksFolderExisting();
                            if (!tf) {
                                showUserError('Setup', 'Could not access your default Outlook Tasks folder.');
                                return;
                            }
                            loadProjects();
                            $scope.config.PROJECTS.defaultProjectEntryID = tf.EntryID;
                            $scope.ui.projectEntryID = tf.EntryID;
                            saveConfig();
                        } else if ($scope.ui.setupProjectMode === 'link') {
                            var lf = linkExistingProject($scope.ui.setupExistingProjectEntryID);
                            if (!lf) {
                                showUserError('Setup', 'Please select an existing folder to link.');
                                return;
                            }
                            $scope.config.PROJECTS.defaultProjectEntryID = lf.entryID;
                            $scope.ui.projectEntryID = lf.entryID;
                            saveConfig();
                        } else {
                            // Create a new project folder under the projects root
                            var root = outlook && outlook.getTaskFolder ? outlook.getTaskFolder($scope.filter.mailbox, rootName) : null;
                            if (!root) {
                                showUserError('Setup', 'Could not access or create the projects root folder in Outlook.');
                                return;
                            }
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
                try {
                    $scope.validateLanes();
                    if ($scope.hasLaneErrors()) {
                        showUserError('Fix lane errors', 'Please fix lane validation errors before saving.');
                        return;
                    }
                } catch (e0) {
                    // ignore
                }
                var ok = saveConfig();
                if (!ok) {
                    showUserError('Settings not saved', 'Your settings could not be saved to Outlook storage.');
                    return;
                }

                applySortableInteractionConfig();
                $scope.applyTheme();
                loadProjects();
                ensureSelectedProject();
                $scope.switchMode('board');
                showToast('success', 'Settings saved', '');
                $scope.refreshTasks();
            };

            // Settings transfer (export/import)
            function buildViewStateObject() {
                return {
                    private: $scope.filter.private,
                    search: $scope.filter.search,
                    category: $scope.filter.category,
                    due: $scope.filter.due,
                    mailbox: $scope.filter.mailbox,
                    projectEntryID: $scope.ui.projectEntryID
                };
            }

            function buildSettingsExport(includeState) {
                var payload = {
                    kind: 'kfo-settings',
                    app: 'Kanban for Outlook',
                    version: $scope.version,
                    exportedAt: nowIso(),
                    schemaVersion: SCHEMA_VERSION,
                    config: $scope.config || DEFAULT_CONFIG_V3()
                };
                if (includeState) {
                    payload.state = buildViewStateObject();
                }
                return payload;
            }

            function tryDownloadTextFile(filename, text, contentType) {
                try {
                    var type = contentType || 'application/octet-stream';
                    var blob = null;
                    try {
                        blob = new Blob([String(text || '')], { type: type });
                    } catch (e1) {
                        blob = null;
                    }

                    if (blob && window.navigator && window.navigator.msSaveBlob) {
                        window.navigator.msSaveBlob(blob, filename);
                        return true;
                    }
                } catch (e) {
                    // ignore
                }
                return false;
            }

            function backupCurrentSettingsToOutlook() {
                try {
                    var ts = nowStamp();
                    try {
                        var cfgRaw = storageRead(CONFIG_ID, 'config', false);
                        if (cfgRaw !== null) {
                            storageWrite(CONFIG_ID + '.backup.' + ts, String(cfgRaw), 'config', false);
                        }
                    } catch (e1) {
                        // ignore
                    }
                    try {
                        var stateRaw = storageRead(STATE_ID, 'state', false);
                        if (stateRaw !== null) {
                            storageWrite(STATE_ID + '.backup.' + ts, String(stateRaw), 'state', false);
                        }
                    } catch (e2) {
                        // ignore
                    }
                } catch (e) {
                    // ignore
                }
            }

            function normaliseConfigObject(cfg) {
                // Applies defaults defensively to avoid runtime errors after import.
                var c = (cfg && typeof cfg === 'object') ? cfg : {};
                var d = DEFAULT_CONFIG_V3();

                if (!c.PROJECTS) c.PROJECTS = d.PROJECTS;
                if (!c.PROJECTS.linkedProjects) c.PROJECTS.linkedProjects = [];
                if (!c.PROJECTS.hiddenProjectEntryIDs) c.PROJECTS.hiddenProjectEntryIDs = [];

                if (!c.UI) c.UI = d.UI;
                if (!c.AUTOMATION) c.AUTOMATION = d.AUTOMATION;
                if (!c.LANES) c.LANES = d.LANES;
                if (!c.THEME) c.THEME = d.THEME;
                if (!c.BOARD) c.BOARD = d.BOARD;

                // UI defaults
                if (c.UI.density === undefined) c.UI.density = d.UI.density;
                if (c.UI.motion === undefined) c.UI.motion = d.UI.motion;
                if (c.UI.laneWidthPx === undefined) c.UI.laneWidthPx = d.UI.laneWidthPx;
                if (c.UI.showDueDate === undefined) c.UI.showDueDate = d.UI.showDueDate;
                if (c.UI.showNotes === undefined) c.UI.showNotes = d.UI.showNotes;
                if (c.UI.showCategories === undefined) c.UI.showCategories = d.UI.showCategories;
                if (c.UI.showOnlyFirstCategory === undefined) c.UI.showOnlyFirstCategory = d.UI.showOnlyFirstCategory;
                if (c.UI.showPriorityPill === undefined) c.UI.showPriorityPill = d.UI.showPriorityPill;
                if (c.UI.showPrivacyIcon === undefined) c.UI.showPrivacyIcon = d.UI.showPrivacyIcon;
                if (c.UI.showLaneCounts === undefined) c.UI.showLaneCounts = d.UI.showLaneCounts;
                if (c.UI.keyboardShortcuts === undefined) c.UI.keyboardShortcuts = d.UI.keyboardShortcuts;

                if (c.AUTOMATION.setOutlookStatusOnLaneMove === undefined) {
                    c.AUTOMATION.setOutlookStatusOnLaneMove = d.AUTOMATION.setOutlookStatusOnLaneMove;
                }

                if (c.USE_CATEGORY_COLORS === undefined) c.USE_CATEGORY_COLORS = true;
                if (c.USE_CATEGORY_COLOR_FOOTERS === undefined) c.USE_CATEGORY_COLOR_FOOTERS = false;
                if (!c.DATE_FORMAT) c.DATE_FORMAT = d.DATE_FORMAT || 'DD-MMM';
                if (c.LOG_ERRORS === undefined) c.LOG_ERRORS = false;
                if (c.MULTI_MAILBOX === undefined) c.MULTI_MAILBOX = d.MULTI_MAILBOX;
                if (!c.ACTIVE_MAILBOXES) c.ACTIVE_MAILBOXES = d.ACTIVE_MAILBOXES || [];

                if (c.BOARD.quickAddEnabled === undefined) c.BOARD.quickAddEnabled = d.BOARD.quickAddEnabled;
                if (c.BOARD.dragHandleOnly === undefined) c.BOARD.dragHandleOnly = d.BOARD.dragHandleOnly;

                // Clamp lane width
                try {
                    var w = parseInt(c.UI.laneWidthPx, 10);
                    if (isNaN(w)) w = d.UI.laneWidthPx;
                    if (w < 240) w = 240;
                    if (w > 520) w = 520;
                    c.UI.laneWidthPx = w;
                } catch (e) {
                    c.UI.laneWidthPx = d.UI.laneWidthPx;
                }

                // Clamp density + motion
                try {
                    var density = String(c.UI.density || 'comfortable');
                    if (density !== 'compact' && density !== 'comfortable') density = 'comfortable';
                    c.UI.density = density;
                } catch (e2) {
                    c.UI.density = 'comfortable';
                }
                try {
                    var motion = String(c.UI.motion || 'full');
                    if (motion !== 'full' && motion !== 'subtle' && motion !== 'off') motion = 'full';
                    c.UI.motion = motion;
                } catch (e3) {
                    c.UI.motion = 'full';
                }

                return c;
            }

            function parseSettingsImportText(text) {
                var raw = String(text || '');
                if (raw.replace(/\s+/g, '') === '') {
                    return { ok: false, error: 'No JSON provided' };
                }

                var obj = null;
                try {
                    var cleaned = raw;
                    try {
                        if (typeof JSON !== 'undefined' && JSON.minify) {
                            cleaned = JSON.minify(raw);
                        }
                    } catch (e1) {
                        cleaned = raw;
                    }
                    obj = JSON.parse(cleaned);
                } catch (e) {
                    return { ok: false, error: 'Invalid JSON: ' + safeErrorString(e) };
                }

                // Supported shapes:
                // 1) { kind: 'kfo-settings', config: {...}, state: {...} }
                // 2) { config: {...}, state: {...} }
                // 3) { SCHEMA_VERSION: 3, ... } (config only)
                var cfg = null;
                var state = null;

                if (obj && typeof obj === 'object' && obj.config) {
                    cfg = obj.config;
                    state = obj.state || null;
                } else if (obj && typeof obj === 'object' && obj.SCHEMA_VERSION) {
                    cfg = obj;
                } else {
                    return { ok: false, error: 'Unrecognised settings format' };
                }

                if (!cfg || typeof cfg !== 'object') {
                    return { ok: false, error: 'Missing config' };
                }

                if (!cfg.SCHEMA_VERSION || cfg.SCHEMA_VERSION < SCHEMA_VERSION) {
                    return { ok: false, error: 'This settings file is for an older version and cannot be imported safely.' };
                }

                if (cfg.SCHEMA_VERSION > SCHEMA_VERSION) {
                    return { ok: false, error: 'This settings file is from a newer version. Please upgrade this app first.' };
                }

                if (state && typeof state !== 'object') {
                    state = null;
                }

                return { ok: true, config: cfg, state: state };
            }

            function applyImportedSettings(parsed) {
                if (!parsed || !parsed.ok) {
                    showUserError('Import failed', parsed ? String(parsed.error || 'Invalid settings') : 'Invalid settings');
                    return;
                }

                // backup current values before overwriting
                backupCurrentSettingsToOutlook();

                if ($scope.ui && $scope.ui.settingsImportApplyConfig) {
                    $scope.config = normaliseConfigObject(parsed.config);
                    var ok1 = saveConfig();
                    if (!ok1) {
                        showUserError('Import failed', 'Could not save imported config to Outlook storage.');
                        return;
                    }
                }

                if ($scope.ui && $scope.ui.settingsImportApplyState && parsed.state) {
                    try {
                        var st = parsed.state;
                        $scope.filter.private = st.private || $scope.privacyFilter.all.value;
                        $scope.filter.search = st.search || '';
                        $scope.filter.category = st.category || '<All Categories>';
                        $scope.filter.due = st.due || 'any';
                        $scope.filter.mailbox = st.mailbox || '';
                        $scope.ui.projectEntryID = st.projectEntryID || '';
                        storageWrite(STATE_ID, JSON.stringify(buildViewStateObject(), null, 2), 'state', false);
                    } catch (e1) {
                        writeLog('import state: ' + e1);
                    }
                }

                try {
                    showToast('success', 'Import complete', 'Reloading...', 1500);
                    $timeout(function () {
                        try { window.location.reload(); } catch (e2) { /* ignore */ }
                    }, 600);
                } catch (e3) {
                    // ignore
                }
            }

            $scope.openSettingsTransfer = function () {
                try {
                    if (!$scope.ui) return;
                    if ($scope.ui.settingsExportIncludeState === undefined) {
                        $scope.ui.settingsExportIncludeState = true;
                    }
                    if ($scope.ui.settingsImportApplyConfig === undefined) {
                        $scope.ui.settingsImportApplyConfig = true;
                    }
                    if ($scope.ui.settingsImportApplyState === undefined) {
                        $scope.ui.settingsImportApplyState = true;
                    }

                    $scope.ui.settingsImportText = '';
                    $scope.refreshSettingsExportText();
                    $scope.ui.showSettingsTransfer = true;
                } catch (e) {
                    reportError('openSettingsTransfer', e, 'Export/import failed', 'Could not open settings transfer. Click the ! icon for details.');
                }
            };

            $scope.refreshSettingsExportText = function () {
                try {
                    if (!$scope.ui) return;
                    $scope.ui.settingsExportText = JSON.stringify(buildSettingsExport(!!$scope.ui.settingsExportIncludeState), null, 2);
                } catch (e) {
                    writeLog('refreshSettingsExportText: ' + e);
                }
            };

            $scope.closeSettingsTransfer = function () {
                try {
                    if ($scope.ui) {
                        $scope.ui.showSettingsTransfer = false;
                    }
                } catch (e) {
                    // ignore
                }
            };

            $scope.copySettingsExport = function () {
                try {
                    $scope.refreshSettingsExportText();
                    var text = ($scope.ui && $scope.ui.settingsExportText) ? $scope.ui.settingsExportText : '';
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

            $scope.downloadSettingsExport = function () {
                try {
                    if (!$scope.ui) return;
                    $scope.refreshSettingsExportText();
                    var name = 'kanban-for-outlook-settings-' + nowStamp() + '.json';
                    var ok = tryDownloadTextFile(name, $scope.ui.settingsExportText, 'application/json;charset=utf-8');
                    if (!ok) {
                        showUserError('Download not supported', 'Use Copy JSON, then save it into a .json file locally.');
                    }
                } catch (e) {
                    reportError('downloadSettingsExport', e, 'Export failed', 'Could not export settings. Click the ! icon for details.');
                }
            };

            $scope.importSettingsFromText = function () {
                try {
                    var parsed = parseSettingsImportText($scope.ui ? $scope.ui.settingsImportText : '');
                    applyImportedSettings(parsed);
                } catch (e) {
                    reportError('importSettingsFromText', e, 'Import failed', 'Could not import settings. Click the ! icon for details.');
                }
            };

            $scope.importSettingsFromFile = function () {
                try {
                    var input = document.getElementById('kfoSettingsFile');
                    if (!input || !input.files || input.files.length === 0) {
                        showUserError('Import', 'Choose a .json file first.');
                        return;
                    }
                    var file = input.files[0];
                    if (!window.FileReader) {
                        showUserError('Import', 'File import is not supported in this host. Paste JSON instead.');
                        return;
                    }

                    var reader = new FileReader();
                    reader.onload = function (e) {
                        $timeout(function () {
                            try {
                                var text = '';
                                try { text = String(e && e.target ? e.target.result : ''); } catch (e1) { text = ''; }
                                var parsed = parseSettingsImportText(text);
                                applyImportedSettings(parsed);
                            } catch (err) {
                                reportError('importSettingsFromFile', err, 'Import failed', 'Could not import settings from file.');
                            }
                        }, 0);
                    };
                    reader.onerror = function () {
                        $timeout(function () {
                            showUserError('Import failed', 'Could not read the selected file.');
                        }, 0);
                    };
                    reader.readAsText(file);
                } catch (e) {
                    reportError('importSettingsFromFile', e, 'Import failed', 'Could not import settings. Click the ! icon for details.');
                }
            };

            // Keyboard shortcuts (opt-in)
            function isKeyboardShortcutsEnabled() {
                try {
                    return !!($scope.config && $scope.config.UI && $scope.config.UI.keyboardShortcuts);
                } catch (e) {
                    return false;
                }
            }

            function isTypingTarget(target) {
                try {
                    if (!target) return false;
                    var tag = '';
                    try { tag = String(target.tagName || '').toLowerCase(); } catch (e1) { tag = ''; }
                    if (tag === 'input' || tag === 'textarea' || tag === 'select') return true;
                    if (target.isContentEditable) return true;
                } catch (e) {
                    // ignore
                }
                return false;
            }

            function focusSearchInput() {
                try {
                    var el = document.getElementById('kfo-search-input');
                    if (el && el.focus) {
                        el.focus();
                        try { el.select(); } catch (e1) { /* ignore */ }
                        return true;
                    }
                } catch (e) {
                    // ignore
                }
                return false;
            }

            function closeOverlaysOrReturnToBoard() {
                var closed = false;
                try {
                    if (!$scope.ui) return false;

                    if ($scope.ui.quickAddLaneId) {
                        $scope.closeQuickAdd();
                        closed = true;
                    }

                    if ($scope.ui.showShortcuts) {
                        $scope.ui.showShortcuts = false;
                        closed = true;
                    }
                    if ($scope.ui.showSettingsTransfer) {
                        $scope.ui.showSettingsTransfer = false;
                        closed = true;
                    }
                    if ($scope.ui.showErrorDetails) {
                        $scope.ui.showErrorDetails = false;
                        closed = true;
                    }
                    if ($scope.ui.showDiagnostics) {
                        $scope.ui.showDiagnostics = false;
                        closed = true;
                    }
                    if ($scope.ui.showRenameProject) {
                        $scope.ui.showRenameProject = false;
                        closed = true;
                    }
                    if ($scope.ui.showCreateProject) {
                        $scope.ui.showCreateProject = false;
                        closed = true;
                    }
                    if ($scope.ui.showMoveTasks) {
                        $scope.closeMoveTasks();
                        closed = true;
                    }
                    if ($scope.ui.showMigration) {
                        $scope.closeMigration();
                        closed = true;
                    }
                    if ($scope.ui.showLaneIdTool) {
                        try { $scope.closeLaneIdTool(); } catch (eLane) { /* ignore */ }
                        closed = true;
                    }
                    if ($scope.ui.showSetupWizard) {
                        $scope.closeSetupWizard();
                        closed = true;
                    }

                    if (!closed && $scope.ui.mode !== 'board') {
                        $scope.ui.mode = 'board';
                        closed = true;
                    }
                } catch (e) {
                    // ignore
                }
                return closed;
            }

            function bindKeyboardShortcuts() {
                if (keyboardShortcutsBound) return;
                keyboardShortcutsBound = true;

                function onKeyDown(e) {
                    var ev = e || window.event;
                    if (!ev) return;

                    var code = ev.keyCode || ev.which;
                    var shift = !!ev.shiftKey;

                    // Always allow Esc to close open dialogs (even when shortcuts are disabled).
                    if (code === 27) {
                        try {
                            var anyOverlay = false;
                            if ($scope.ui) {
                                anyOverlay = !!(
                                    $scope.ui.quickAddLaneId ||
                                    $scope.ui.showShortcuts ||
                                    $scope.ui.showSettingsTransfer ||
                                    $scope.ui.showErrorDetails ||
                                    $scope.ui.showDiagnostics ||
                                    $scope.ui.showRenameProject ||
                                    $scope.ui.showCreateProject ||
                                    $scope.ui.showMoveTasks ||
                                    $scope.ui.showMigration ||
                                    $scope.ui.showLaneIdTool ||
                                    $scope.ui.showSetupWizard
                                );
                            }
                            if (anyOverlay) {
                                $timeout(function () { closeOverlaysOrReturnToBoard(); }, 0);
                                try {
                                    if (ev.preventDefault) ev.preventDefault();
                                    ev.returnValue = false;
                                } catch (eEsc) {
                                    // ignore
                                }
                                return false;
                            }
                        } catch (e0) {
                            // ignore
                        }
                    }

                    if (!isKeyboardShortcutsEnabled()) return;
                    if (ev.altKey || ev.ctrlKey || ev.metaKey) return;

                    // Ignore while typing
                    var tgt = ev.target || ev.srcElement;
                    if (isTypingTarget(tgt)) return;

                    var handled = false;

                    // ? (Shift + /)
                    if (code === 191 && shift) {
                        handled = true;
                        $timeout(function () { $scope.openShortcuts(); }, 0);
                    }

                    // / focuses search
                    if (!handled && code === 191 && !shift) {
                        handled = focusSearchInput();
                    }

                    // r refreshes
                    if (!handled && code === 82) {
                        handled = true;
                        $timeout(function () { $scope.refreshTasks(); }, 0);
                    }

                    // Esc closes dialogs / returns to board
                    if (!handled && code === 27) {
                        handled = true;
                        $timeout(function () { closeOverlaysOrReturnToBoard(); }, 0);
                    }

                    if (handled) {
                        try {
                            if (ev.preventDefault) ev.preventDefault();
                            ev.returnValue = false;
                        } catch (e1) {
                            // ignore
                        }
                        return false;
                    }
                }

                try {
                    if (document.addEventListener) {
                        document.addEventListener('keydown', onKeyDown, false);
                    } else if (document.attachEvent) {
                        document.attachEvent('onkeydown', onKeyDown);
                    } else {
                        document.onkeydown = onKeyDown;
                    }
                } catch (e) {
                    writeLog('bindKeyboardShortcuts: ' + e);
                }
            }

            $scope.openShortcuts = function () {
                try {
                    if ($scope.ui) {
                        $scope.ui.showShortcuts = true;
                    }
                } catch (e) {
                    // ignore
                }
            };

            $scope.closeShortcuts = function () {
                try {
                    if ($scope.ui) {
                        $scope.ui.showShortcuts = false;
                    }
                } catch (e) {
                    // ignore
                }
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
                applySortableInteractionConfig();
                bindKeyboardShortcuts();
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

                // First run: show the setup wizard
                if (!$scope.config.SETUP.completed) {
                    $scope.ui.showSetupWizard = true;
                    $scope.ui.setupStep = 1;
                    $scope.ui.setupProjectMode = 'default';
                    $scope.applyLaneTemplate($scope.ui.setupLaneTemplate);
                    saveConfig();
                }

                $scope.refreshTasks();
            };
        }]);
})();
