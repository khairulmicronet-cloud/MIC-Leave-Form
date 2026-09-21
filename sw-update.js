// Shared PWA update helper (loaded by every page).
//
// sw.js already activates a new version immediately (skipWaiting +
// clients.claim) instead of waiting for every open tab/app instance to be
// fully closed. The piece that was missing: once a new worker takes
// control, this page needs to reload itself so the fresh HTML/JS actually
// gets used -- otherwise whatever's already loaded in memory keeps running
// the old code until the app is fully closed and reopened.
//
// With this in place, simply opening the app (or bringing it back to the
// foreground) is enough to pick up the latest release automatically. No
// re-"Add to Home Screen", no manual cache clearing, no telling staff to
// reinstall anything.
(function () {
  if (!("serviceWorker" in navigator)) return;

  window.addEventListener("load", () => {
    navigator.serviceWorker.register("sw.js").catch((err) => console.warn("SW registration failed", err));
  });

  let reloaded = false;
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (reloaded) return;
    reloaded = true;
    window.location.reload();
  });
})();
