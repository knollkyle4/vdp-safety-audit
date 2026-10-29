const CACHE='vdp-safety-v4';
const APP=['./','./index.html','./manifest.webmanifest','./icon.svg','./drive-integration.js'];

self.addEventListener('install',e=>{
  e.waitUntil(caches.open(CACHE).then(c=>c.addAll(APP)));
  self.skipWaiting();
});

self.addEventListener('activate',e=>{
  e.waitUntil(e.waitUntil ? Promise.resolve() : Promise.resolve());
  e.waitUntil(
    caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k))))
  );
  self.clients.claim();
});

async function pageResponse(){
  try{
    const resp=await fetch('./index.html',{cache:'no-store'});
    let html=await resp.text();
    if(!html.includes('drive-integration.js')){
      html=html.replace('</body>','<script src="./drive-integration.js?v=4"></script></body>');
    }
    return new Response(html,{status:200,headers:{'Content-Type':'text/html; charset=utf-8','Cache-Control':'no-store'}});
  }catch(err){
    const cached=await caches.match('./index.html');
    if(!cached) throw err;
    let html=await cached.text();
    if(!html.includes('drive-integration.js')){
      html=html.replace('</body>','<script src="./drive-integration.js?v=4"></script></body>');
    }
    return new Response(html,{status:200,headers:{'Content-Type':'text/html; charset=utf-8','Cache-Control':'no-store'}});
  }
}

self.addEventListener('fetch',e=>{
  if(e.request.method!=='GET') return;
  if(e.request.mode==='navigate'){
    e.respondWith(pageResponse());
    return;
  }
  e.respondWith(
    fetch(e.request,{cache:'no-store'}).then(async resp=>{
      const cache=await caches.open(CACHE);
      cache.put(e.request,resp.clone());
      return resp;
    }).catch(()=>caches.match(e.request))
  );
});
