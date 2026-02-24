'use strict';

(function (root, factory) {
    root.kfoAppCore = factory(root);
})(typeof window !== 'undefined' ? window : (typeof global !== 'undefined' ? global : this), function (root) {
    var util = root.kfoUtil || null;
    var themeSafety = root.kfoThemeSafety || null;
    var board = root.kfoBoard || null;
    var outlook = root.kfoOutlook || null;

    var CONFIG_ID = 'KanbanConfig';
    var STATE_ID = 'KanbanState';
    var LOG_ID = 'KanbanErrorLog';

    var SCHEMA_VERSION = 3;

    // Outlook task user properties (stored locally in Outlook)
    // Note: some Outlook installs reject custom field names containing underscores.
    // Use underscore-free names, but keep backward-compatible reads in the controller.
    var PROP_LANE_ID = 'KFOLaneId';
    var PROP_LANE_ORDER = 'KFOLaneOrder';
    var PROP_LANE_CHANGED_AT = 'KFOLaneChangedAt';

    var DEFAULT_ROOT_FOLDER_NAME = 'Kanban Projects';

    var BUILTIN_THEMES = [
        { id: 'kfo-light', name: 'Light mode', cssHref: 'themes/kfo-light/theme.css', kind: 'builtin' },
        { id: 'kfo-dark', name: 'Dark mode', cssHref: 'themes/kfo-dark/theme.css', kind: 'builtin' }
    ];

    function nowStamp() {
        return util && util.nowStamp ? util.nowStamp() : '';
    }

    function nowIso() {
        return util && util.nowIso ? util.nowIso() : String(new Date());
    }

    function safeErrorString(e) {
        return util && util.safeErrorString ? util.safeErrorString(e) : String(e || '');
    }

    function sanitizeId(raw) {
        return util && util.sanitizeId ? util.sanitizeId(raw) : '';
    }

    function isValidHexColor(s) {
        return util && util.isValidHexColor ? util.isValidHexColor(s) : false;
    }

    function isRealDate(d) {
        return util && util.isRealDate ? util.isRealDate(d) : false;
    }

    function isCssLocalOnly(cssText) {
        return themeSafety && themeSafety.isCssLocalOnly ? themeSafety.isCssLocalOnly(cssText) : false;
    }

    function isSafeLocalCssPath(href) {
        return themeSafety && themeSafety.isSafeLocalCssPath ? themeSafety.isSafeLocalCssPath(href) : false;
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
            VIEWS: [],
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
                showLaneCounts: true,

                // Optional: highlight search matches in cards
                highlightSearch: false,

                // Optional: custom dropdowns with typeahead (avoids native <select> rendering quirks)
                customDropdowns: false,

                // Opt-in (avoid Outlook conflicts)
                keyboardShortcuts: false
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
                saveOrder: true,

                // Quick add in lanes (no inspector)
                quickAddEnabled: true,

                // Optional: reduce accidental drags
                dragHandleOnly: false,

                // Optional: completion flow
                completeMovesToDone: true,

                // When viewing only active tasks, optionally keep completed tasks visible in the Done lane.
                showDoneCompletedInActiveView: false,

                // Staleness cues (time in lane)
                staleDaysThreshold: 7,
                showStalePill: true
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

    return {
        // dependencies
        util: util,
        themeSafety: themeSafety,
        board: board,
        outlook: outlook,

        // constants
        CONFIG_ID: CONFIG_ID,
        STATE_ID: STATE_ID,
        LOG_ID: LOG_ID,
        SCHEMA_VERSION: SCHEMA_VERSION,
        PROP_LANE_ID: PROP_LANE_ID,
        PROP_LANE_ORDER: PROP_LANE_ORDER,
        PROP_LANE_CHANGED_AT: PROP_LANE_CHANGED_AT,
        DEFAULT_ROOT_FOLDER_NAME: DEFAULT_ROOT_FOLDER_NAME,
        BUILTIN_THEMES: BUILTIN_THEMES,

        // helpers
        nowStamp: nowStamp,
        nowIso: nowIso,
        safeErrorString: safeErrorString,
        sanitizeId: sanitizeId,
        isValidHexColor: isValidHexColor,
        isRealDate: isRealDate,
        isCssLocalOnly: isCssLocalOnly,
        isSafeLocalCssPath: isSafeLocalCssPath,
        DEFAULT_CONFIG_V3: DEFAULT_CONFIG_V3,
        laneTemplate: laneTemplate
    };
});
