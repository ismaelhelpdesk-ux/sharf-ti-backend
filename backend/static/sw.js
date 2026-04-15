/**
 * SHARF Devoluciones TI — Service Worker
 * Permite: instalación como app, cache offline, sincronización background
 */

const CACHE_NAME   = 'sharf-v1';
const API_BASE     = self.location.origin;

// Archivos a cachear para funcionamiento offline
const STATIC_FILES = [
  '/',
  '/static/app.js',
  '/static/icon-192.png',
  '/static/icon-512.png',
  '/manifest.json',
];

// ── Instalación ────────────────────────────────────────────────────────
self.addEventListener('install', evt => {
  self.skipWaiting();
  evt.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      console.log('[SW] Cacheando archivos estáticos');
      return cache.addAll(STATIC_FILES.filter(f => f !== '/manifest.json'));
    }).catch(e => console.warn('[SW] Cache parcial:', e))
  );
});

// ── Activación ─────────────────────────────────────────────────────────
self.addEventListener('activate', evt => {
  evt.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys
        .filter(k => k !== CACHE_NAME)
        .map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
  console.log('[SW] Activado — SHARF v1');
});

// ── Fetch strategy ─────────────────────────────────────────────────────
self.addEventListener('fetch', evt => {
  const url = new URL(evt.request.url);

  // API calls → Network first, sin cache (datos siempre frescos)
  if (url.pathname.startsWith('/api/')) {
    evt.respondWith(
      fetch(evt.request).catch(() =>
        new Response(
          JSON.stringify({ ok: false, error: 'Sin conexión — verifica tu red', offline: true }),
          { status: 503, headers: { 'Content-Type': 'application/json' } }
        )
      )
    );
    return;
  }

  // Recursos estáticos → Cache first, network fallback
  evt.respondWith(
    caches.match(evt.request).then(cached => {
      if (cached) return cached;
      return fetch(evt.request).then(resp => {
        if (resp && resp.status === 200 && evt.request.method === 'GET') {
          const clone = resp.clone();
          caches.open(CACHE_NAME).then(c => c.put(evt.request, clone));
        }
        return resp;
      }).catch(() => {
        // Offline fallback para navegación
        if (evt.request.mode === 'navigate') {
          return caches.match('/');
        }
      });
    })
  );
});

// ── Background Sync — reintento de devoluciones pendientes ────────────
self.addEventListener('sync', evt => {
  if (evt.tag === 'sharf-devolucion-pendiente') {
    evt.waitUntil(sincronizarPendientes());
  }
});

async function sincronizarPendientes() {
  try {
    const db    = await abrirDB();
    const items = await obtenerPendientes(db);
    for (const item of items) {
      try {
        const r = await fetch('/api/devolucion', {
          method:  'POST',
          headers: { 'Content-Type': 'application/json' },
          body:    JSON.stringify(item.payload)
        });
        if (r.ok) {
          await eliminarPendiente(db, item.id);
          self.registration.showNotification('SHARF TI', {
            body:  '✅ Devolución sincronizada exitosamente',
            icon:  '/static/icon-192.png',
            badge: '/static/icon-192.png',
            tag:   'sync-ok'
          });
        }
      } catch (e) {
        console.warn('[SW] No se pudo sincronizar:', e);
      }
    }
  } catch (e) {
    console.error('[SW] sincronizarPendientes:', e);
  }
}

// ── IndexedDB helpers ──────────────────────────────────────────────────
function abrirDB() {
  return new Promise((res, rej) => {
    const req = indexedDB.open('sharf-offline', 1);
    req.onupgradeneeded = e => {
      e.target.result.createObjectStore('pendientes', { keyPath: 'id', autoIncrement: true });
    };
    req.onsuccess = e => res(e.target.result);
    req.onerror   = e => rej(e.target.error);
  });
}
function obtenerPendientes(db) {
  return new Promise((res, rej) => {
    const tx  = db.transaction('pendientes', 'readonly');
    const req = tx.objectStore('pendientes').getAll();
    req.onsuccess = e => res(e.target.result);
    req.onerror   = e => rej(e.target.error);
  });
}
function eliminarPendiente(db, id) {
  return new Promise((res, rej) => {
    const tx = db.transaction('pendientes', 'readwrite');
    tx.objectStore('pendientes').delete(id);
    tx.oncomplete = res;
    tx.onerror    = rej;
  });
}

// ── Push Notifications ─────────────────────────────────────────────────
self.addEventListener('push', evt => {
  const data = evt.data ? evt.data.json() : { title: 'SHARF TI', body: 'Nueva notificación' };
  evt.waitUntil(
    self.registration.showNotification(data.title || 'SHARF TI', {
      body:  data.body  || '',
      icon:  '/static/icon-192.png',
      badge: '/static/icon-192.png',
      data:  data.url ? { url: data.url } : {},
    })
  );
});

self.addEventListener('notificationclick', evt => {
  evt.notification.close();
  if (evt.notification.data?.url) {
    evt.waitUntil(clients.openWindow(evt.notification.data.url));
  }
});
