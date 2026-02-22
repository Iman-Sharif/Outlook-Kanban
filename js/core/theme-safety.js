'use strict';

(function (root, factory) {
    if (typeof module === 'object' && module && module.exports) {
        module.exports = factory();
    } else {
        root.kfoThemeSafety = factory();
    }
})(typeof window !== 'undefined' ? window : (typeof global !== 'undefined' ? global : this), function () {
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

    return {
        isCssLocalOnly: isCssLocalOnly,
        isSafeLocalCssPath: isSafeLocalCssPath
    };
});
