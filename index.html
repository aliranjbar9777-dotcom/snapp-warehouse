// سرویس‌ورکر ساده فقط برای «نصب‌پذیر» شدن داشبورد (Add to Home Screen).
// استراتژی: Network First — یعنی همیشه اول تلاش می‌کند نسخه‌ی تازه را از
// اینترنت بگیرد، و فقط اگر واقعاً آفلاین بود از نسخه‌ی کش‌شده استفاده می‌کند.
// این‌طوری هیچ‌وقت یک نسخه‌ی قدیمی و باگ‌دار از اپ گیر کاربر نمی‌ماند.

const CACHE_NAME = 'snapp-warehouse-shell-v2';
const SHELL_FILES = [
  './index.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES)).catch(()=>{})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // هرگز درخواست‌های Google Apps Script / Google Sheets را دست نزن؛
  // داده‌ی انبار همیشه باید مستقیم و زنده از شبکه خوانده شود.
  if (url.hostname.includes('script.google.com') || url.hostname.includes('googleusercontent.com')) {
    return;
  }

  if (event.request.method !== 'GET' || url.origin !== self.location.origin) return;

  event.respondWith(
    fetch(event.request)
      .then((networkRes) => {
        const copy = networkRes.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy)).catch(()=>{});
        return networkRes;
      })
      .catch(() => caches.match(event.request)) // فقط وقتی واقعاً آفلاین است
  );
});
