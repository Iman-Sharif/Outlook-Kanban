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

    return {
        nowIso: nowIso,
        nowStamp: nowStamp,
        safeErrorString: safeErrorString,
        sanitizeId: sanitizeId,
        isValidHexColor: isValidHexColor,
        isRealDate: isRealDate
    };
});
