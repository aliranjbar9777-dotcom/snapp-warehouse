// این سرویس‌ورکر دیگر استفاده نمی‌شود و فقط برای «خودکشی» نگه داشته شده:
// اگر مرورگری قبلاً نسخه‌ی قدیمی این فایل را نصب کرده باشد، این نسخه به محض
// اجرا خودش را حذف و تمام کش‌های مرتبط را پاک می‌کند.
self.addEventListener('install', () => { self.skipWaiting(); });
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.map((k) => caches.delete(k))))
      .then(() => self.registration.unregister())
      .then(() => self.clients.matchAll())
      .then((clients) => clients.forEach((c) => c.navigate(c.url)))
  );
});
self.addEventListener('fetch', () => {}); // بدون هیچ کشی، فقط عبور می‌دهد
