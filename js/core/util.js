'use strict';

(function (root, factory) {
    if (typeof module === 'object' && module && module.exports) {
        module.exports = factory();
    } else {
        root.kfoUtil = factory();
    }
})(typeof window !== 'undefined' ? window : (typeof global !== 'undefined' ? global : this), function () {
    function nowIso() {
        try {
            return (new Date()).toISOString();
        } catch (e) {
            return String(new Date());
        }
    }

    function nowStamp() {
        var d = new Date();
        function pad(n) { return (n < 10 ? '0' : '') + n; }
        return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '-' + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
    }

    function safeErrorString(e) {
        try {
            if (e === null || e === undefined) return '';
            if (typeof e === 'string') return e;
            if (e.message) return String(e.message);
            return String(e);
        } catch (err) {
            return 'unknown error';
        }
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

    function isRealDate(d) {
        try {
            if (!d) return false;
            if (isNaN(d.getTime())) return false;
            // Outlook sometimes returns a sentinel far-future date
            if (d.getFullYear && d.getFullYear() === 4501) return false;
            return true;
        } catch (e) {
            return false;
        }
    }

    function detectEol(text) {
        try {
            return (String(text || '').indexOf('\r\n') !== -1) ? '\r\n' : '\n';
        } catch (e) {
            return '\n';
        }
    }

    function splitLines(text) {
        return String(text || '').split(/\r?\n/);
    }

    function matchChecklistLine(line) {
        // Markdown checkbox list:
        // - [ ] item
        // - [x] item
        // Also accepts: * [ ] item
        var s = String(line || '');
        var m = s.match(/^(\s*[-*]\s+\[)( |x|X)(\]\s*)(.*)$/);
        if (!m) return null;
        return { prefix: m[1], mark: m[2], suffix: m[3], text: m[4] };
    }

    function parseChecklist(bodyText) {
        var body = String(bodyText || '');
        var lines = splitLines(body);
        var items = [];
        var other = [];
        for (var i = 0; i < lines.length; i++) {
            var m = matchChecklistLine(lines[i]);
            if (m) {
                items.push({
                    lineIndex: i,
                    checked: (String(m.mark || '').toLowerCase() === 'x'),
                    text: String(m.text || '').replace(/^\s+|\s+$/g, ''),
                    raw: String(lines[i] || '')
                });
            } else {
                other.push(lines[i]);
            }
        }
        return {
            items: items,
            otherText: other.join('\n')
        };
    }

    function toggleChecklistItem(bodyText, lineIndex, checked) {
        try {
            var body = String(bodyText || '');
            var eol = detectEol(body);
            var lines = splitLines(body);
            var idx = parseInt(lineIndex, 10);
            if (isNaN(idx) || idx < 0 || idx >= lines.length) return body;
            var m = matchChecklistLine(lines[idx]);
            if (!m) return body;
            var mark = checked ? 'x' : ' ';
            lines[idx] = String(m.prefix || '') + mark + String(m.suffix || '') + String(m.text || '');
            return lines.join(eol);
        } catch (e) {
            return String(bodyText || '');
        }
    }

    function addChecklistItem(bodyText, itemText) {
        try {
            var text = String(itemText || '').replace(/^\s+|\s+$/g, '');
            if (!text) return String(bodyText || '');

            var body = String(bodyText || '');
            var eol = detectEol(body);
            var lines = splitLines(body);

            var lastIdx = -1;
            for (var i = 0; i < lines.length; i++) {
                if (matchChecklistLine(lines[i])) lastIdx = i;
            }

            var newLine = '- [ ] ' + text;

            if (lastIdx >= 0) {
                lines.splice(lastIdx + 1, 0, newLine);
                return lines.join(eol);
            }

            // No checklist yet.
            var trimmed = body.replace(/^\s+|\s+$/g, '');
            if (!trimmed) {
                return newLine;
            }

            // Ensure a blank line separation.
            if (lines.length && String(lines[lines.length - 1] || '') !== '') {
                lines.push('');
            }
            lines.push(newLine);
            return lines.join(eol);
        } catch (e) {
            return String(bodyText || '');
        }
    }

    return {
        nowIso: nowIso,
        nowStamp: nowStamp,
        safeErrorString: safeErrorString,
        sanitizeId: sanitizeId,
        isValidHexColor: isValidHexColor,
        isRealDate: isRealDate,

        // Task body helpers (pure)
        parseChecklist: parseChecklist,
        toggleChecklistItem: toggleChecklistItem,
        addChecklistItem: addChecklistItem
    };
});
