/* PWA helpers: service worker registration, install button, update and offline notices. */
(() => {
  const installBtn = document.getElementById('install-btn');
  const toast = document.getElementById('pwa-toast');
  const showToast = (html, ms = 0) => {
    toast.innerHTML = html; toast.hidden = false;
    if (ms) setTimeout(() => { toast.hidden = true; }, ms);
  };

  // Service workers only run over http(s). Opening index.html directly from disk still works, just without install/offline.
  if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
    window.addEventListener('load', async () => {
      try {
        const reg = await navigator.serviceWorker.register('./sw.js', { scope: './' });
        const offerUpdate = worker => {
          showToast('A new version is available. <button class="btn" id="pwa-reload">Update</button>');
          document.getElementById('pwa-reload').onclick = () => worker.postMessage('skipWaiting');
        };
        if (reg.waiting && navigator.serviceWorker.controller) offerUpdate(reg.waiting);
        reg.addEventListener('updatefound', () => {
          const w = reg.installing;
          w && w.addEventListener('statechange', () => {
            if (w.state === 'installed' && navigator.serviceWorker.controller) offerUpdate(w);
          });
        });
        reg.update();
        document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') reg.update(); });
        let reloaded = false;
        navigator.serviceWorker.addEventListener('controllerchange', () => { if (!reloaded) { reloaded = true; location.reload(); } });
      } catch (e) { console.warn('Service worker registration failed', e); }
    });
  }

  // Install button (Chrome, Edge, Android). Safari uses Share → Add to Home Screen / Add to Dock.
  let deferred = null;
  const standalone = () => matchMedia('(display-mode: standalone)').matches || navigator.standalone === true;
  window.addEventListener('beforeinstallprompt', e => {
    e.preventDefault(); deferred = e;
    if (!standalone()) installBtn.hidden = false;
  });
  installBtn.addEventListener('click', async () => {
    if (deferred) {
      deferred.prompt();
      await deferred.userChoice; deferred = null; installBtn.hidden = true;
    } else {
      showToast('To install: in Safari use <b>Share → Add to Home Screen</b> (iPhone/iPad) or <b>File → Add to Dock</b> (Mac).', 7000);
    }
  });
  window.addEventListener('appinstalled', () => { installBtn.hidden = true; showToast('Bulk Energy installed.', 3000); });
  const isSafari = /^((?!chrome|android|crios|fxios|edg).)*safari/i.test(navigator.userAgent);
  if (isSafari && !standalone() && /^https?:$/.test(location.protocol)) installBtn.hidden = false;

  // Offline indicator.
  const net = () => document.body.classList.toggle('offline', !navigator.onLine);
  window.addEventListener('online', () => { net(); showToast('Back online.', 2500); });
  window.addEventListener('offline', () => { net(); showToast('You are offline. The dashboard keeps working from the saved copy.', 4000); });
  net();
})();
