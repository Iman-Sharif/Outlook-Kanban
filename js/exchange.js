'use strict';

var outlookApp;
var outlookNS;

var SENSITIVITY = { olNormal: 0, olPrivate: 2 };
var OlDefaultFolders = { olFolderTasks: 13 };
var OlItemType = { olTaskItem: 3 };
var OlUserPropertyType = {
    olText: 1,
    olNumber: 3,
    olYesNo: 6
};

function checkBrowser() {
    var isBrowserSupported
    if (window.external !== undefined && window.external.OutlookApplication !== undefined) {
        isBrowserSupported = true;
        outlookApp = window.external.OutlookApplication;
        outlookNS = outlookApp.GetNameSpace("MAPI");
    } else {
        try {
            isBrowserSupported = true;
            outlookApp = new ActiveXObject("Outlook.Application");
            outlookNS = outlookApp.GetNameSpace("MAPI");
        }
        catch (e) {
            isBrowserSupported = false;
        }
    }
    return isBrowserSupported;
}

function getOutlookCategories() {
    var i;
    var catNames = [];
    var catColors = [];
    var categories = outlookNS.Categories;
    var count = outlookNS.Categories.Count;
    catNames.length = count;
    catColors.length = count;
    for (i = 1; i <= count; i++) {
        catNames[i - 1] = categories(i).Name;
        catColors[i - 1] = categories(i).Color;
    };
    return { names: catNames, colors: catColors };
}

function getOutlookMailboxes(multiMailbox) {
    var i;
    var mi = 0;
    var mailboxNames = [];
    var folders = outlookNS.Folders;
    var count = folders.count;
    mailboxNames.length = count;
    mailboxNames[0] = fixMailboxName(getDefaultMailbox().Name);
    if (!multiMailbox) 
    { 
        mailboxNames.length = 1;
        return mailboxNames; 
    }
    for (i = 1; i <= count; i++) {
        try {
            var acc = folders.Item(i).Name;
            if (acc.indexOf("Internet Calendar") == -1) {
                if (acc != mailboxNames[0]) {
                    if (hasTasksFolder(folders.Item(i))) {
                        mi++;
                        mailboxNames[mi] = fixMailboxName(acc);
                    }
                }
            }
        }
        catch (e)
        {
            // ignore this error, because this mailbox will not be useful for the kanban board
        }
    };
    mailboxNames.length = mi + 1;
    return mailboxNames;
}

function hasTasksFolder(mailbox) {
    var i;
    for (i = 1; i <= mailbox.Folders.count; i++) {
        if (mailbox.Folders(i).DefaultItemType == OlItemType.olTaskItem ) {
            return true;
        }
    }
    return false;
}

function findTasksFolder(mailboxName) {
    var i;
    var j = getFolderIndex(outlookNS.Folders, mailboxName);
    var mailbox = outlookNS.Folders(j);
    for (i = 1; i <= mailbox.Folders.count; i++) {
        if (mailbox.Folders(i).DefaultItemType == OlItemType.olTaskItem ) {
            return mailbox.Folders(i);
        }
    }
    return false;
}

function fixMailboxName(name) {
    var i = name.indexOf(' <');
    if (i > -1) {
        name = name.substring(0,i);
    }
    return name;
}

function getDefaultMailbox() {
    return outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderTasks).Parent;
}

function getOutlookTodayHomePageFolder() {
    return outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderTasks).Parent.WebViewUrl;
}

function getOutlookVersion() {
    return outlookApp.version;
}

function getFolderIndex(folders, folder) {
    try {
        var i;
        for (i = 1; i <= folders.count; i++) {
            if (folders(i).Name == folder) {
                return i;
            }
        }
        return -1;
    } catch (error) {
        alert('getFolderIndex error:' + error)
    }
}

function getTaskFolder(mailbox, folderName) {
    try {
        var folder = findTasksFolder(mailbox);
        if (folderName == '') {
            return folder;
        }
        var returnFolder = getOrCreateFolder(mailbox, folderName, folder.Folders, OlDefaultFolders.olFolderTasks);
        return returnFolder;
    } catch (error) {
        alert('getTaskFolder error:' + error)
    }
}

function getTaskFolderExisting(mailbox, folderName) {
    try {
        var folder = findTasksFolder(mailbox);
        if (folderName === '') {
            return folder;
        }
        var i = getFolderIndex(folder.Folders, folderName);
        if (i === -1) {
            return null;
        }
        return folder.Folders(folderName);
    } catch (error) {
        alert('getTaskFolderExisting error:' + error);
        return null;
    }
}

function listTaskSubFolders(mailbox, parentFolderName) {
    // Returns plain JS objects (safe for Angular bindings)
    try {
        var parent = getTaskFolderExisting(mailbox, parentFolderName);
        if (!parent) {
            return [];
        }
        var result = [];
        var folders = parent.Folders;
        var count = folders.Count;
        for (var i = 1; i <= count; i++) {
            try {
                var f = folders(i);
                if (f.DefaultItemType == OlItemType.olTaskItem) {
                    result.push({
                        name: f.Name,
                        entryID: f.EntryID,
                        storeID: f.StoreID
                    });
                }
            } catch (e) {
                // ignore unreadable folders
            }
        }
        // Sort by name for stable UI
        result.sort(function (a, b) {
            var an = (a.name || '').toLowerCase();
            var bn = (b.name || '').toLowerCase();
            if (an < bn) return -1;
            if (an > bn) return 1;
            return 0;
        });
        return result;
    } catch (error) {
        alert('listTaskSubFolders error:' + error);
        return [];
    }
}

function getFolderFromIDs(entryID, storeID) {
    try {
        // Namespace.GetFolderFromID
        if (outlookNS && outlookNS.GetFolderFromID) {
            if (storeID) {
                return outlookNS.GetFolderFromID(entryID, storeID);
            }
            return outlookNS.GetFolderFromID(entryID);
        }
    } catch (e) {
        // fall through
    }
    return null;
}

function getOrCreateFolder(mailbox, folderName, inFolders, folderType) {
    try {
        var i = getFolderIndex(inFolders, folderName);
        if (i == -1) {
            var f = inFolders.Add(folderName, folderType);
            if (f.Name != folderName) {
                inFolders.Add(folderName, folderType);
                f.Delete();
            }
        }
        return inFolders(folderName);
    } catch (error) {
        alert('getOrCreateFolder error creating folder ' + folderName + ' in ' + mailbox + '  error: ' + error)
    }
}

function getDefaultTasksFolderName() {
    return outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderTasks).Name;
}

function getJournalFolder() {
    return outlookNS.GetDefaultFolder(11);
}

function getTaskItems(mailbox, folderName) {
    return getTaskFolder(mailbox, folderName).Items;
}

function getTaskItem(id) {
    return outlookNS.GetItemFromID(id);
}

function getTaskItemFromIDs(entryID, storeID) {
    try {
        if (storeID) {
            return outlookNS.GetItemFromID(entryID, storeID);
        }
    } catch (e) {
        // fall back
    }
    return outlookNS.GetItemFromID(entryID);
}

function newMailItem() {
    return outlookApp.CreateItem(0);
}

function newJournalItem() {
    return outlookApp.CreateItem(4);
}

function newNoteItem() {
    return outlookApp.CreateItem(5);
}

function getJournalItem(subject) {
    var folder = getJournalFolder();
    var configItems = folder.Items.Restrict('[Subject] = "' + subject + '"');
    if (configItems.Count > 0) {
        var configItem = configItems(1);
        if (configItem.Body) {
            return configItem.Body;
        }
    }
    return null;
}

function getPureJournalItem(subject) {
    var folder = getJournalFolder();
    var configItems = folder.Items.Restrict('[Subject] = "' + subject + '"');
    if (configItems.Count > 0) {
        var configItem = configItems(1);
        return configItem;
    }
    return null;
}

function saveJournalItem(subject, body) {
    var folder = getJournalFolder();
    var configItems = folder.Items.Restrict('[Subject] = "' + subject + '"');
    if (configItems.Count == 0) {
        var configItem = newJournalItem();
        configItem.Subject = subject;
    }
    else {
        configItem = configItems(1);
    }
    configItem.Body = body;
    configItem.Save();
}

function getUserEmailAddress() {
    try {
        return outlookNS.folders.Item(1).SmtpAddress;
    } catch (error) {
        return 'address-unknown';
    }
}

function getUserName() {
    try {
        return outlookApp.Session.CurrentUser.Name;
    } catch (error) {
        return 'name-unknown';
    }
}

function getUserProperty(item, prop) {
    var userprop = item.UserProperties(prop);
    var value = '';
    if (userprop != null) {
        value = userprop.Value;
    }
    return value;
};

function setUserProperty(item, prop, value, type) {
    try {
        var userprop = null;
        try {
            userprop = item.UserProperties(prop);
        } catch (e1) {
            userprop = null;
        }

        if (userprop == null) {
            var t = type;
            if (t === undefined || t === null) {
                t = OlUserPropertyType.olText;
            }
            userprop = item.UserProperties.Add(prop, t);
        }
        userprop.Value = value;
    } catch (error) {
        alert('setUserProperty error: ' + error);
    }
}
