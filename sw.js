// سرویس‌ورکر ساده فقط برای «نصب‌پذیر» شدن داشبورد (Add to Home Screen) و
// بارگذاری سریع‌تر پوسته‌ی اپ. داده‌های Google Sheets همیشه به‌صورت زنده
// (fetch مستقیم) خوانده می‌شوند و در این کش ذخیره نمی‌شوند.

const CACHE_NAME = 'snapp-warehouse-shell-v1';
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

  // هرگز درخواست‌های مربوط به Google Apps Script / Google Sheets را کش نکن؛
  // داده‌ی انبار همیشه باید زنده و به‌روز باشد.
  if (url.hostname.includes('script.google.com') || url.hostname.includes('googleusercontent.com')) {
    return; // بگذار مرورگر عادی fetch کند
  }

  // فقط فایل‌های خودِ پوسته‌ی اپ را از کش سرو کن (Cache First)
  if (event.request.method === 'GET' && url.origin === self.location.origin) {
    event.respondWith(
      caches.match(event.request).then((cached) => cached || fetch(event.request))
    );
  }
});
