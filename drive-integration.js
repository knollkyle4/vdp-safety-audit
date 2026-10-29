(()=>{
  const URL_KEY='vdpGoogleScriptUrlV1';

  function currentAudit(){
    try { return JSON.parse(localStorage.getItem('vdpSafetyStateV2')||'{}'); }
    catch(e){ return {}; }
  }

  function makeCard(){
    const summary=document.getElementById('summary');
    if(summary && !document.getElementById('driveSendCard')){
      const card=document.createElement('div');
      card.className='card';
      card.id='driveSendCard';
      card.innerHTML=`
        <h3>Google Drive Reports</h3>
        <p class="help">Creates the weekly summary as a Google Doc and the audit presentation as Google Slides in your VDP Safety Audits work Drive folder. Nothing is downloaded to this iPhone.</p>
        <button class="btn full" id="sendToDriveBtn">Send Reports to Google Drive</button>
        <div id="driveSendStatus" class="help" style="margin-top:9px"></div>`;
      summary.appendChild(card);
      document.getElementById('sendToDriveBtn').addEventListener('click',sendReports);
    }

    const backup=document.getElementById('backup');
    if(backup && !document.getElementById('driveSetupCard')){
      const card=document.createElement('div');
      card.className='card';
      card.id='driveSetupCard';
      card.innerHTML=`
        <h3>Google Drive Connection</h3>
        <p class="help">One-time setup. Paste the Google Apps Script Web App URL here. It stays only on this device.</p>
        <label>Apps Script Web App URL</label>
        <input id="driveScriptUrl" type="url" placeholder="https://script.google.com/macros/s/.../exec">
        <button class="btn" id="saveDriveUrl" style="margin-top:10px">Save Drive Connection</button>
        <div id="driveSetupStatus" class="help" style="margin-top:8px"></div>`;
      backup.insertBefore(card, backup.firstChild);
      const input=document.getElementById('driveScriptUrl');
      input.value=localStorage.getItem(URL_KEY)||'';
      document.getElementById('saveDriveUrl').addEventListener('click',()=>{
        const v=input.value.trim();
        if(!/^https:\/\/script\.google\.com\//.test(v)){
          document.getElementById('driveSetupStatus').textContent='Enter the Apps Script Web App URL ending in /exec.';
          return;
        }
        localStorage.setItem(URL_KEY,v);
        document.getElementById('driveSetupStatus').textContent='Google Drive connection saved on this device.';
      });
    }
  }

  async function sendReports(){
    const status=document.getElementById('driveSendStatus');
    const btn=document.getElementById('sendToDriveBtn');
    const url=localStorage.getItem(URL_KEY)||'';
    if(!url){
      status.textContent='Setup required: go to Backup → Google Drive Connection and save the Apps Script URL.';
      return;
    }
    const audit=currentAudit();
    if(!audit.weekEnding){ status.textContent='Choose a Week Ending date first.'; return; }
    if(!Array.isArray(audit.findings)){ status.textContent='No audit data found.'; return; }

    btn.disabled=true;
    btn.textContent='Sending…';
    status.textContent='Sending audit and photos to your work Google Drive…';

    try{
      // no-cors is intentional: Apps Script ContentService does not reliably expose CORS headers.
      // The POST still executes and creates the files in Drive.
      await fetch(url,{
        method:'POST',
        mode:'no-cors',
        headers:{'Content-Type':'text/plain;charset=utf-8'},
        body:JSON.stringify({weekEnding:audit.weekEnding,findings:audit.findings,sentAt:new Date().toISOString()})
      });
      status.textContent='Reports sent. Check the VDP Safety Audits folder in your work Google Drive for the new Google Doc and Google Slides files.';
    }catch(err){
      status.textContent='Could not send reports. Recheck the Apps Script URL and your internet connection.';
    }finally{
      btn.disabled=false;
      btn.textContent='Send Reports to Google Drive';
    }
  }

  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',makeCard);
  else makeCard();
})();
