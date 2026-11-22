// Service worker mínimo de Novato Gestión. Sólo sirve para que la app sea instalable y abra rápido.
// NO guarda datos del negocio: las llamadas al backend (Apps Script) y a cualquier otro dominio
// pasan siempre directo a la red. Sin conexión, la app no muestra datos viejos.
const CACHE = 'novato-shell-v1';

self.addEventListener('install', (e) => { self.skipWaiting(); });

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return; // backend y APIs: directo a la red

  // Páginas: primero la red (así una versión nueva se ve enseguida); si no hay red, la última guardada.
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req)
        .then((res) => { const copia = res.clone(); caches.open(CACHE).then((c) => c.put('/', copia)); return res; })
        .catch(() => caches.match('/').then((r) => r || Response.error()))
    );
    return;
  }
  // Archivos con hash en el nombre (/assets/...): nunca cambian, se guardan.
  if (url.pathname.startsWith('/assets/')) {
    e.respondWith(
      caches.match(req).then((r) => r || fetch(req).then((res) => {
        if (res.ok) { const copia = res.clone(); caches.open(CACHE).then((c) => c.put(req, copia)); }
        return res;
      }))
    );
  }
});
