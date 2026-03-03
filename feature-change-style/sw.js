const CACHE_NAME='physiocontrol-v1';
const APP_SHELL=['./','./index.html','./style.css','./manifest.json','./icons/icon.svg','./icons/icon-192.png','./icons/icon-512.png','./js/db.js','./js/utils.js','./js/auth.js','./js/log.js','./js/dashboard.js','./js/history.js','./js/admin.js','./js/profile.js','./js/export.js','./js/ui.js'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE_NAME).then(c=>c.addAll(APP_SHELL)));self.skipWaiting();});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE_NAME).map(k=>caches.delete(k)))));self.clients.claim();});
self.addEventListener('fetch',e=>{if(e.request.url.includes('supabase.co')){e.respondWith(fetch(e.request));return;}e.respondWith(caches.match(e.request).then(c=>c||fetch(e.request)));});
