'use strict';

(function (root, factory) {
    if (typeof module === 'object' && module && module.exports) {
        module.exports = factory();
    } else {
        root.kfoOutlook = factory();
    }
})(typeof window !== 'undefined' ? window : (typeof global !== 'undefined' ? global : this), function () {
    function report(context, error) {
        try {
            if (typeof window !== 'undefined' && window.kfoReportError) {
                window.kfoReportError('adapter.' + String(context || ''), error);
            }
        } catch (e) {
            // ignore
        }
    }

    function tryCall(context, fn, fallbackValue) {
        try {
            return { ok: true, value: fn() };
        } catch (e) {
            report(context, e);
            return { ok: false, value: fallbackValue, error: String(e) };
        }
    }

    function callOr(context, fn, fallbackValue) {
        var r = tryCall(context, fn, fallbackValue);
        return r.value;
    }

    // Public API
    var api = {
        // environment
        checkBrowser: function () {
            return callOr('checkBrowser', function () { return checkBrowser(); }, false);
        },
        getBrowserSupportDetails: function () {
            return callOr('getBrowserSupportDetails', function () { return getBrowserSupportDetails(); }, { supported: false, method: 'unknown', error: '' });
        },
        getOutlookVersion: function () {
            return callOr('getOutlookVersion', function () { return getOutlookVersion(); }, 'unknown');
        },
        getOutlookTodayHomePageFolder: function () {
            return callOr('getOutlookTodayHomePageFolder', function () { return getOutlookTodayHomePageFolder(); }, 'unknown');
        },

        // core data
        getOutlookCategories: function () {
            return callOr('getOutlookCategories', function () { return getOutlookCategories(); }, { names: [], colors: [] });
        },
        getOutlookMailboxes: function (multiMailbox) {
            return callOr('getOutlookMailboxes', function () { return getOutlookMailboxes(!!multiMailbox); }, []);
        },

        // folders
        getTaskFolder: function (mailbox, folderName) {
            return callOr('getTaskFolder', function () { return getTaskFolder(mailbox, folderName); }, null);
        },
        getTaskFolderExisting: function (mailbox, folderName) {
            return callOr('getTaskFolderExisting', function () { return getTaskFolderExisting(mailbox, folderName); }, null);
        },
        listTaskSubFolders: function (mailbox, parentFolderName) {
            return callOr('listTaskSubFolders', function () { return listTaskSubFolders(mailbox, parentFolderName); }, []);
        },
        getFolderFromIDs: function (entryID, storeID) {
            return callOr('getFolderFromIDs', function () { return getFolderFromIDs(entryID, storeID); }, null);
        },
        getOrCreateFolder: function (mailbox, folderName, inFolders, folderType) {
            return callOr('getOrCreateFolder', function () { return getOrCreateFolder(mailbox, folderName, inFolders, folderType); }, null);
        },

        // items
        getTaskItem: function (entryID) {
            return callOr('getTaskItem', function () { return getTaskItem(entryID); }, null);
        },
        getTaskItemFromIDs: function (entryID, storeID) {
            return callOr('getTaskItemFromIDs', function () { return getTaskItemFromIDs(entryID, storeID); }, null);
        },

        // journal storage
        tryGetJournalItem: function (subject) {
            return tryCall('getJournalItem', function () { return getJournalItem(subject); }, null);
        },
        trySaveJournalItem: function (subject, body) {
            return tryCall('saveJournalItem', function () { saveJournalItem(subject, body); return true; }, false);
        },

        // user properties
        getUserProperty: function (item, prop) {
            return callOr('getUserProperty', function () { return getUserProperty(item, prop); }, '');
        },
        setUserProperty: function (item, prop, value, type) {
            return callOr('setUserProperty', function () { setUserProperty(item, prop, value, type); return true; }, false);
        }
    };

    api._tryCall = tryCall;

    return api;
});
