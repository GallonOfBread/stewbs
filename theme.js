// Light/dark mode for every page on stewbs.com.
//
// Load this with a plain <script> in <head> (not defer/async) so the theme is set
// before the page paints, which avoids a white flash in dark mode.
//
// Until a visitor uses the switch, the page follows their device setting. Their choice is
// then kept in localStorage. It stays in their browser and is never sent anywhere, so
// no cookie banner is needed.
(function () {
    var KEY = "theme";
    var root = document.documentElement;
    var media = window.matchMedia("(prefers-color-scheme: dark)");

    function saved() {
        try {
            var t = localStorage.getItem(KEY);
            return t === "dark" || t === "light" ? t : null;
        } catch (e) {
            return null; // storage blocked, e.g. some private modes
        }
    }

    function current() {
        return saved() || (media.matches ? "dark" : "light");
    }

    function apply(theme) {
        root.dataset.theme = theme;
        var meta = document.querySelector('meta[name="theme-color"]');
        if (meta) meta.content = theme === "dark" ? "#0e0f12" : "#ffffff";
        document.querySelectorAll("[data-theme-toggle]").forEach(function (b) {
            b.setAttribute("aria-checked", String(theme === "dark"));
        });
    }

    apply(current());

    // Follow device changes until the visitor picks a theme themselves
    media.addEventListener("change", function () {
        if (!saved()) apply(current());
    });
    // Keep other open tabs in sync
    window.addEventListener("storage", function (e) {
        if (e.key === KEY) apply(current());
    });

    // Fill in and wire up every <button data-theme-toggle> on the page
    var SUN = '<svg class="lucide icon-sun" viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="4"/><path d="M12 2v2"/><path d="M12 20v2"/><path d="m4.93 4.93 1.41 1.41"/><path d="m17.66 17.66 1.41 1.41"/><path d="M2 12h2"/><path d="M20 12h2"/><path d="m6.34 17.66-1.41 1.41"/><path d="m19.07 4.93-1.41 1.41"/></svg>';
    var MOON = '<svg class="lucide icon-moon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/></svg>';

    document.addEventListener("DOMContentLoaded", function () {
        document.querySelectorAll("[data-theme-toggle]").forEach(function (b) {
            b.innerHTML = '<span class="thumb">' + SUN + MOON + "</span>";
            b.addEventListener("click", function () {
                var next = root.dataset.theme === "dark" ? "light" : "dark";
                try { localStorage.setItem(KEY, next); } catch (e) {}
                apply(next);
            });
        });
        apply(current());
    });
})();
