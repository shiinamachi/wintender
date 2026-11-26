// Theme switching shared by every wintender widget.
//
// Each page ships one complete <style data-wintender-theme="..."> block per
// theme. This classic (non-module) script runs synchronously from <head>,
// after those blocks, so the stored theme is applied before the first paint.
// It also follows theme changes made from the settings window.
(() => {
    const STORAGE_KEY = "wintender:settings";
    const CHANNEL_NAME = "wintender-settings";
    const THEMES = Object.freeze(["sakura-night", "legacy"]);
    const DEFAULT_THEME = "sakura-night";

    function normalizeTheme(theme) {
        return THEMES.includes(theme) ? theme : DEFAULT_THEME;
    }

    function readStoredTheme(storage = globalThis.localStorage) {
        try {
            const raw = storage?.getItem?.(STORAGE_KEY);
            return normalizeTheme(raw ? JSON.parse(raw)?.theme : null);
        } catch {
            return DEFAULT_THEME;
        }
    }

    function applyTheme(theme) {
        const nextTheme = normalizeTheme(theme);

        document.documentElement.dataset.theme = nextTheme;
        document
            .querySelectorAll("style[data-wintender-theme]")
            .forEach((style) => {
                style.media =
                    style.dataset.wintenderTheme === nextTheme
                        ? "all"
                        : "not all";
            });

        return nextTheme;
    }

    function syncFromStorage() {
        return applyTheme(readStoredTheme());
    }

    globalThis.wintenderTheme = Object.freeze({
        THEMES,
        DEFAULT_THEME,
        normalizeTheme,
        readStoredTheme,
        applyTheme,
    });

    syncFromStorage();

    window.addEventListener("storage", (event) => {
        if (!event?.key || event.key === STORAGE_KEY) {
            syncFromStorage();
        }
    });

    if (typeof BroadcastChannel === "function") {
        try {
            new BroadcastChannel(CHANNEL_NAME).addEventListener(
                "message",
                syncFromStorage,
            );
        } catch (error) {
            console.error("Failed to follow wintender theme changes", error);
        }
    }
})();
