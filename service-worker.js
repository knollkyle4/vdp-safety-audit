const CACHE='vdp-safety-v3';
const APP=['./','./index.html','./manifest.webmanifest','./icon.svg','./drive-integration.js'];

self.addEventListener('install',e=>{
  e.waitUntil(caches.open(CACHE).then(c=>c.addAll(APP)));
  self.skipWaiting();
});

self.addEventListener('activate',e=>{
  e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))));
  self.clients.claim();
});

self.addEventListener('fetch',e=>{
  if(e.request.method!=='GET') return;

  if(e.request.mode==='navigate'){
    e.respondWith((async()=>{
      try{
        const resp=await fetch('./index.html',{cache:'no-store'});
        let html=await resp.text();
        if(!html.includes('drive-integration.js')){
          html=html.replace('</body>','<script src="./drive-integration.js"></script></body>');
        }
        const out=new Response(html,{headers:{'Content-Type':'text/html; charset=utf-8'}});
        const cache=await caches.open(CACHE);
        cache.put('./index.html',out.clone());
        return out;
      }catch(err){
        const cached=await caches.match('./index.html');
        if(cached){
          let html=await cached.text();
          if(!html.includes('drive-integration.js')) html=html.replace('</body>','<script src="./drive-integration.js"></script></body>');
          return new Response(html,{headers:{'Content-Type':'text/html; charset=utf-8'}});
        }
        throw err;
      }
    })());
    return;
  }

  e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request).then(resp=>{
    const copy=resp.clone();
    caches.open(CACHE).then(c=>c.put(e.request,copy));
    return resp;
  })));
});
