/* AccountStore: guarda la app en el celular para que abra al instante.
   Los datos de la planilla no pasan por acá (van por la API, con PIN). */
const VERSION = 'accountstore-v1';
const APP = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(APP)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

const guardar = (req, res) => { if (res && res.ok) caches.open(VERSION).then(c => c.put(req, res)); };
const guardada = req => caches.match(req, { ignoreSearch: true });

// La página: primero internet (así llegan las actualizaciones); si tarda o no hay señal, la guardada
function redPrimero(req) {
  return new Promise(resolver => {
    let listo = false;
    const usar = r => { if (!listo && r) { listo = true; resolver(r); } };
    const espera = setTimeout(() => guardada(req).then(usar), 2500);
    fetch(req).then(res => { clearTimeout(espera); guardar(req, res.clone()); usar(res); })
      .catch(() => guardada(req).then(r => usar(r || guardada('./index.html'))));
  });
}
// Íconos y letras: la copia guardada, y si no está, internet
const guardadaPrimero = req => guardada(req).then(r => r || fetch(req).then(res => { guardar(req, res.clone()); return res; }));

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                 // la API va por POST: nunca se guarda
  const url = new URL(req.url);
  if (url.origin === location.origin) {
    e.respondWith(req.mode === 'navigate' || url.pathname.endsWith('.html') || url.pathname.endsWith('/')
      ? redPrimero(req) : guardadaPrimero(req));
  } else if (/fonts\.(googleapis|gstatic)\.com$/.test(url.hostname)) {
    e.respondWith(guardadaPrimero(req));
  }
});
