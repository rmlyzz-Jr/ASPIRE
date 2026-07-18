const CACHE_NAME = 'aspire-v2.0';
const STATIC_ASSETS = [
    './',
    './index.html',
    './manifest.json',
    'https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap',
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css',
    'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js'
];

// Install Service Worker
self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => {
                console.log('✅ Caching static assets...');
                return cache.addAll(STATIC_ASSETS);
            })
            .then(() => self.skipWaiting())
    );
});

// Activate Service Worker
self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys().then(cacheNames => {
            return Promise.all(
                cacheNames.map(cacheName => {
                    if (cacheName !== CACHE_NAME) {
                        console.log('🗑️ Removing old cache:', cacheName);
                        return caches.delete(cacheName);
                    }
                })
            );
        }).then(() => self.clients.claim())
    );
});

// Fetch Strategy: Network First, Fallback to Cache
self.addEventListener('fetch', event => {
    const request = event.request;
    
    // Skip non-GET requests
    if (request.method !== 'GET') {
        event.respondWith(fetch(request));
        return;
    }

    // Skip external resources like maps.googleapis.com
    if (request.url.includes('maps.googleapis.com') || 
        request.url.includes('wasenderapi.com') ||
        request.url.includes('i.ibb.co.com')) {
        event.respondWith(fetch(request));
        return;
    }

    event.respondWith(
        fetch(request)
            .then(response => {
                // Clone response for caching
                const responseClone = response.clone();
                caches.open(CACHE_NAME)
                    .then(cache => {
                        if (request.url.startsWith('https://script.google.com')) {
                            return;
                        }
                        cache.put(request, responseClone);
                    });
                return response;
            })
            .catch(() => {
                return caches.match(request)
                    .then(cachedResponse => {
                        if (cachedResponse) {
                            return cachedResponse;
                        }
                        // Fallback for offline pages
                        if (request.headers.get('accept').includes('text/html')) {
                            return caches.match('./');
                        }
                        return new Response('Offline - Content not available', {
                            status: 503,
                            statusText: 'Service Unavailable'
                        });
                    });
            })
    );
});

// Handle push notifications
self.addEventListener('push', event => {
    let data = {
        title: 'ASPIRE Notification',
        body: 'Ada update baru di sistem aduan',
        icon: 'https://i.ibb.co.com/qMvmKCkH/aspire.png',
        badge: 'https://i.ibb.co.com/qMvmKCkH/aspire.png'
    };

    if (event.data) {
        try {
            data = event.data.json();
        } catch (e) {
            data.body = event.data.text();
        }
    }

    const options = {
        body: data.body,
        icon: data.icon || 'https://i.ibb.co.com/qMvmKCkH/aspire.png',
        badge: data.badge || 'https://i.ibb.co.com/qMvmKCkH/aspire.png',
        vibrate: [200, 100, 200],
        data: {
            url: data.url || './'
        },
        actions: [
            { action: 'open', title: '📋 Buka Aplikasi' },
            { action: 'close', title: '✖ Tutup' }
        ]
    };

    event.waitUntil(
        self.registration.showNotification(data.title, options)
    );
});

// Handle notification click
self.addEventListener('notificationclick', event => {
    event.notification.close();

    if (event.action === 'close') {
        return;
    }

    event.waitUntil(
        clients.openWindow('./')
    );
});