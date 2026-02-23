'use strict';

(function (root, factory) {
    if (typeof module === 'object' && module && module.exports) {
        module.exports = factory();
    } else {
        root.kfoThemeSafety = factory();
    }
})(typeof window !== 'undefined' ? window : (typeof global !== 'undefined' ? global : this), function () {
    function stripCssComments(s) {
        return String(s || '').replace(/\/\*[\s\S]*?\*\//g, '');
    }

    function isCssLocalOnly(cssText) {
        // Best-effort guardrail: prevent accidental external loads in custom themes.
        // Users can still theme fully locally.
        try {
            var s = stripCssComments(cssText).toLowerCase();

            // External loads
            if (/(^|[^a-z0-9])https?:\/\//.test(s)) return false;
            if (/@import\b/.test(s)) return false;

            // Protocol-relative remote references (e.g. url(//example.com/x.png))
            if (/url\s*\(\s*["']?\s*\/\//.test(s)) return false;
            // IE filter/legacy patterns that use src='//...'
            if (/\bsrc\s*=\s*["']\s*\/\//.test(s)) return false;

            // Scriptable URL schemes
            if (/url\s*\(\s*["']?\s*javascript:/.test(s)) return false;
            if (/url\s*\(\s*["']?\s*vbscript:/.test(s)) return false;
            if (/\bsrc\s*=\s*["']\s*javascript:/.test(s)) return false;
            if (/\bsrc\s*=\s*["']\s*vbscript:/.test(s)) return false;

            // IE-specific scriptable CSS features
            if (/\bexpression\s*\(/.test(s)) return false;
            if (/\bbehavior\s*:/.test(s)) return false;
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
