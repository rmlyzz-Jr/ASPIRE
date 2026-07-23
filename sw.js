// ==================== SERVICE WORKER - ASPIRE v2.0 ====================
const CACHE_NAME = 'aspire-v2.0';
const GAS_URL = 'https://script.google.com/macros/s/AKfycbxpaF_LQmPuB67yqWNbo_N0A8TFrmvJ9etJFvKoYcB4ERcmNrSRITz6H-BnIf0GT2Pu4A/exec';

// 🔥 STATIC ASSETS (HANYA UNTUK HALAMAN REDIRECT)
const STATIC_ASSETS = [
    './',
    './index.html',
    './manifest.json',
    'https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap',
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css',
    'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js',
    // 🔥 TAMBAHKAN FAVICON/LOGO
    'https://i.ibb.co.com/qMvmKCkH/aspire.png'
];

// 🔥 OFFLINE FALLBACK PAGE (HTML)
const OFFLINE_PAGE = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ASPIRE - Offline</title>
    <style>
        body {
            font-family: 'Inter', sans-serif;
            background: linear-gradient(135deg, #0F172A 0%, #1E293B 50%, #0F172A 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            margin: 0;
            padding: 20px;
            color: white;
        }
        .card {
            background: rgba(255,255,255,0.05);
            backdrop-filter: blur(20px);
            border-radius: 32px;
            padding: 40px;
            max-width: 400px;
            text-align: center;
            border: 1px solid rgba(255,255,255,0.1);
        }
        .card img { width: 80px; height: 80px; margin-bottom: 16px; }
        .card h1 { font-size: 28px; margin-bottom: 8px; }
        .card p { color: rgba(255,255,255,0.6); font-size: 14px; line-height: 1.6; }
        .card .icon { font-size: 48px; margin-bottom: 16px; }
        .btn {
            display: inline-block;
            margin-top: 20px;
            padding: 12px 32px;
            background: linear-gradient(135deg, #4F46E5, #7C3AED);
            color: white;
            border: none;
            border-radius: 60px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            text-decoration: none;
            transition: all 0.3s;
        }
        .btn:hover { transform: scale(1.05); box-shadow: 0 8px 25px rgba(79,70,229,0.4); }
        .status { font-size: 12px; color: rgba(255,255,255,0.3); margin-top: 16px; }
        .status .dot {
            display: inline-block;
            width: 8px;
            height: 8px;
            background: #EF4444;
            border-radius: 50%;
            margin-right: 8px;
            animation: pulse 1.5s ease infinite;
        }
        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.3; }
        }
    </style>
</head>
<body>
    <div class="card">
        <div class="icon">📡</div>
        <img src="https://i.ibb.co.com/qMvmKCkH/aspire.png" alt="ASPIRE">
        <h1>ASPIRE</h1>
        <p>Anda sedang offline. Silakan periksa koneksi internet Anda untuk mengakses aplikasi.</p>
        <button class="btn" onclick="location.reload()">
            <i class="fas fa-sync-alt"></i> Coba Lagi
        </button>
        <div class="status">
            <span class="dot"></span> Offline Mode
        </div>
    </div>
    <script>
        // Coba refresh otomatis saat koneksi kembali
        window.addEventListener('online', function() {
            location.reload();
        });
    </script>
</body>
</html>
`;

// ==================== INSTALL ====================
self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => {
                console.log('✅ Caching static assets...');
                // Cache static assets
                return cache.addAll(STATIC_ASSETS);
            })
            .then(() => {
                // 🔥 CACHE OFFLINE PAGE
                return caches.open(CACHE_NAME)
                    .then(cache => {
                        const response = new Response(OFFLINE_PAGE, {
                            headers: { 'Content-Type': 'text/html' }
                        });
                        return cache.put('/offline', response);
                    });
            })
            .then(() => {
                console.log('✅ Service Worker installed successfully');
                return self.skipWaiting();
            })
    );
});

// ==================== ACTIVATE ====================
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
        }).then(() => {
            console.log('✅ Service Worker activated');
            return self.clients.claim();
        })
    );
});

// ==================== FETCH ====================
self.addEventListener('fetch', event => {
    const request = event.request;
    const url = new URL(request.url);
    
    // 🔥 SKIP: Non-GET requests
    if (request.method !== 'GET') {
        event.respondWith(fetch(request));
        return;
    }

    // 🔥 SKIP: External APIs (maps, wasenderapi, image hosting)
    if (url.hostname.includes('maps.googleapis.com') || 
        url.hostname.includes('wasenderapi.com') ||
        url.hostname.includes('i.ibb.co.com')) {
        event.respondWith(fetch(request));
        return;
    }

    // 🔥 SPECIAL: Google Apps Script (CACHE FIRST, THEN NETWORK)
    if (url.href.includes('script.google.com')) {
        event.respondWith(
            caches.match(request)
                .then(cachedResponse => {
                    if (cachedResponse) {
                        // 🔥 UPDATE CACHE IN BACKGROUND
                        fetch(request)
                            .then(networkResponse => {
                                if (networkResponse && networkResponse.status === 200) {
                                    caches.open(CACHE_NAME)
                                        .then(cache => cache.put(request, networkResponse));
                                }
                            })
                            .catch(() => {});
                        return cachedResponse;
                    }
                    // 🔥 FALLBACK: FETCH FROM NETWORK
                    return fetch(request)
                        .then(networkResponse => {
                            if (networkResponse && networkResponse.status === 200) {
                                caches.open(CACHE_NAME)
                                    .then(cache => cache.put(request, networkResponse.clone()));
                            }
                            return networkResponse;
                        })
                        .catch(() => {
                            // 🔥 TAMPILKAN OFFLINE PAGE
                            return caches.match('/offline');
                        });
                })
        );
        return;
    }

    // 🔥 NORMAL: Network First, Fallback to Cache
    event.respondWith(
        fetch(request)
            .then(response => {
                // Cache successful responses
                if (response && response.status === 200) {
                    const responseClone = response.clone();
                    caches.open(CACHE_NAME)
                        .then(cache => {
                            cache.put(request, responseClone);
                        });
                }
                return response;
            })
            .catch(() => {
                return caches.match(request)
                    .then(cachedResponse => {
                        if (cachedResponse) {
                            return cachedResponse;
                        }
                        // 🔥 HTML FALLBACK
                        if (request.headers.get('accept') && 
                            request.headers.get('accept').includes('text/html')) {
                            return caches.match('/offline');
                        }
                        return new Response('Offline - Content not available', {
                            status: 503,
                            statusText: 'Service Unavailable'
                        });
                    });
            })
    );
});

// ==================== PUSH NOTIFICATION ====================
self.addEventListener('push', event => {
    let data = {
        title: '🔔 ASPIRE Notification',
        body: '📋 Ada update baru di sistem aduan',
        icon: 'https://i.ibb.co.com/qMvmKCkH/aspire.png',
        badge: 'https://i.ibb.co.com/qMvmKCkH/aspire.png',
        url: './'
    };

    if (event.data) {
        try {
            const parsed = event.data.json();
            data = { ...data, ...parsed };
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
        ],
        requireInteraction: true
    };

    event.waitUntil(
        self.registration.showNotification(data.title, options)
    );
});

// ==================== NOTIFICATION CLICK ====================
self.addEventListener('notificationclick', event => {
    event.notification.close();

    if (event.action === 'close') {
        return;
    }

    // 🔥 BUKA APLIKASI
    const urlToOpen = event.notification.data?.url || './';
    
    event.waitUntil(
        clients.matchAll({ type: 'window', includeUncontrolled: true })
            .then(windowClients => {
                // Cari tab yang sudah terbuka
                for (let i = 0; i < windowClients.length; i++) {
                    const client = windowClients[i];
                    if (client.url === urlToOpen && 'focus' in client) {
                        return client.focus();
                    }
                }
                // Buka tab baru
                if (clients.openWindow) {
                    return clients.openWindow(urlToOpen);
                }
            })
    );
});

// ==================== MESSAGE HANDLER ====================
self.addEventListener('message', event => {
    if (event.data && event.data.type === 'SKIP_WAITING') {
        self.skipWaiting();
    }
});

console.log('🚀 ASPIRE Service Worker v2.0 loaded');
