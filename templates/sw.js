const CACHE_NAME = 'kiosk-offline-v1';

const ASSETS_TO_CACHE = [
    '/',
    '/static/images/police_logo.png',
    '/static/images/logo.png',
    'https://cdn.tailwindcss.com',
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css'
];

self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            return cache.addAll(ASSETS_TO_CACHE);
        })
    );
});

self.addEventListener('fetch', (event) => {
    const url = new URL(event.request.url);

    // Check if the request is for an API or dynamic route
    if (url.pathname.startsWith('/api/') || url.pathname.startsWith('/services')) {
        event.respondWith(
            fetch(event.request)
                .then((networkResponse) => {
                    // Clone and store fresh response in cache
                    return caches.open(CACHE_NAME).then((cache) => {
                        cache.put(event.request, networkResponse.clone());
                        return networkResponse;
                    });
                })
                .catch(() => {
                    // Fallback to cached API response if offline
                    return caches.match(event.request);
                })
        );
    } else {
        // Standard cache-first strategy for static assets/pages
        event.respondWith(
            caches.match(event.request).then((cachedResponse) => {
                if (cachedResponse) {
                    return cachedResponse;
                }
                return fetch(event.request).catch(() => {
                    return caches.match('/');
                });
            })
        );
    }
});