const CACHE_NAME = "smel-caxias-v11";
const STATIC_ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./logo-caxias-mais-esporte.png",
  "./logo-prefeitura-smel.png"
];

self.addEventListener("install", event => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(STATIC_ASSETS)).catch(() => {})
  );
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", event => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);

  // Não intercepta chamadas da API do Google Apps Script nem do Supabase
  if (url.hostname.includes("script.google.com") || url.hostname.includes("script.googleusercontent.com") || url.hostname.includes("supabase.co")) {
    return;
  }

  // Navegação HTML: Network-First com fallback instantâneo para o Cache em caso de queda de 4G na quadra
  if (req.mode === "navigate" || url.pathname.endsWith(".html") || url.pathname.endsWith("/")) {
    event.respondWith(
      fetch(req)
        .then(res => {
          const copy = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(req, copy)).catch(() => {});
          return res;
        })
        .catch(() => caches.match(req).then(r => r || caches.match("./index.html")))
    );
    return;
  }

  // Imagens, logos e fontes: Stale-While-Revalidate (0ms de resposta do cache + atualização em background)
  event.respondWith(
    caches.match(req).then(cached => {
      const networkFetch = fetch(req)
        .then(res => {
          if (res && res.status === 200) {
            const copy = res.clone();
            caches.open(CACHE_NAME).then(c => c.put(req, copy)).catch(() => {});
          }
          return res;
        })
        .catch(() => cached);
      return cached || networkFetch;
    })
  );
});
