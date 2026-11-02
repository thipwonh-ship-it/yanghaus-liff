// Service worker ขั้นต่ำ — ให้ Android เสนอ "ติดตั้งแอป" ได้ ไม่แคชหน้า (หน้าโหลดจาก GitHub Pages เสมอ จะได้เห็นเวอร์ชันใหม่ทันที)
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', () => {});
