// sw.js — offline app shell (cache-first)
var CACHE='pqa-sr-v1';
var SHELL=['./','./index.html','./manifest.json'];
self.addEventListener('install',function(e){
  e.waitUntil(caches.open(CACHE).then(function(c){ return c.addAll(SHELL); }).then(function(){ return self.skipWaiting(); }));
});
self.addEventListener('activate',function(e){
  e.waitUntil(caches.keys().then(function(ks){ return Promise.all(ks.map(function(k){ if(k!==CACHE) return caches.delete(k); })); }).then(function(){ return self.clients.claim(); }));
});
self.addEventListener('fetch',function(e){
  var req=e.request;
  if(req.method!=='GET') return;                 // GASへのPOSTは絶対に触らない
  var url=new URL(req.url);
  if(url.origin!==location.origin) return;        // GAS等の別ドメインは素通し
  e.respondWith(
    caches.match(req).then(function(hit){
      if(hit) return hit;
      return fetch(req).then(function(res){ var copy=res.clone(); caches.open(CACHE).then(function(c){ c.put(req,copy); }); return res; })
                       .catch(function(){ return caches.match('./index.html'); });
    })
  );
});
