const DESTINATION_FOLDER_ID = '1uB_EJHbK7ytV5bTLqeFfBZeK90dNiSEv';

const AREA_SLIDE_GROUPS = [
  {title:'Front Office / Hallways / Restrooms / Processing / Tanker Bay', areas:['Lobby / Offices / Hallways / Restrooms','Processing','Tanker Bay']},
  {title:'Electrical Room / Boiler / CIP', areas:['Electrical Room','Boiler & Utilities','CIP Room']},
  {title:'Dry Goods Storage', areas:['Dry Goods Storage']},
  {title:'Bottle / Crate Washer Area', areas:['Bottle / Crate Washer Area']},
  {title:'Cooler / New Cooler', areas:['Cooler / New Cooler']},
  {title:'Maintenance / Bottling / Jug Room / BIB Room', areas:['Maintenance','Bottling','Jug Room','BIB Room']},
  {title:'Ice Cream / Freezer', areas:['Ice Cream','Freezer']},
  {title:'Exterior / Loading Docks', areas:['Exterior / Loading Docks']},
  {title:'Breakroom / Warehouse', areas:['Breakroom','Warehouse']}
];

function doGet() {
  return ContentService.createTextOutput('VDP Safety Audit receiver is active.');
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const expectedToken = PropertiesService.getScriptProperties().getProperty('VDP_UPLOAD_TOKEN');

    console.log(
      'VDP upload check | scriptToken=' + !!expectedToken +
      ' | phoneToken=' + !!payload.token +
      ' | tokenMatch=' + (payload.token === expectedToken) +
      ' | weekEnding=' + (payload.weekEnding || 'MISSING') +
      ' | findings=' + (Array.isArray(payload.findings) ? payload.findings.length : 'NOT ARRAY')
    );

    if (!expectedToken) throw new Error('VDP_UPLOAD_TOKEN is missing from Script Properties.');
    if (!payload.token || payload.token !== expectedToken) throw new Error('Unauthorized: phone upload token does not match Script Property.');
    if (!payload.weekEnding) throw new Error('Phone request is missing the Week Ending date.');
    if (!Array.isArray(payload.findings)) throw new Error('Phone request is missing the findings array.');

    const folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const score = score_(payload.findings);
    const baseName = 'VDP Safety Audit - Week Ending ' + payload.weekEnding;
    const doc = buildSummaryDoc_(payload, score, baseName, folder);
    const slides = buildSlides_(payload, score, baseName, folder, doc.getUrl());

    console.log('SUCCESS | SPI=' + score.spi.toFixed(2) + ' | findings=' + payload.findings.length);
    return json_({ok:true,docUrl:doc.getUrl(),slidesUrl:slides.getUrl(),spi:score.spi,findings:payload.findings.length});
  } catch (err) {
    console.error(err);
    throw err;
  }
}

function score_(findings) {
  const ni = findings.filter(f => f.level === 'NI').length;
  const nc = findings.filter(f => f.level === 'NC').length;
  const deduction = ni * 0.25 + nc;
  return {ni,nc,deduction,spi:Math.max(0,100-deduction)};
}

function buildSummaryDoc_(payload, score, baseName, folder) {
  const doc = DocumentApp.create(baseName + ' - Summary');
  DriveApp.getFileById(doc.getId()).moveTo(folder);
  const body = doc.getBody();
  const dateLabel = formatDate_(payload.weekEnding);

  body.appendParagraph('VOLLEMAN DAIRY PROCESSING LLC')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('WEEKLY SAFETY AUDIT - FINDINGS & SCORE')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Week Ending ' + dateLabel).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const status = score.spi >= 95 ? 'PASS' : 'BELOW TARGET';
  const metrics = body.appendTable([
    ['SPI','Target','Findings','Status'],
    [score.spi.toFixed(2)+'%','95.00%',String(payload.findings.length),status]
  ]);
  styleHeader_(metrics,0);

  body.appendParagraph('PRIMARY FOCUS').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(primaryFocus_(payload.findings));

  body.appendParagraph('REPEAT / RECURRING CONCERNS').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const repeatLines = repeatConcernLines_(payload.findings);
  if (repeatLines.length) repeatLines.forEach(line => body.appendListItem(line));
  else body.appendParagraph('No repeat findings marked this week.');

  body.appendParagraph('Scoring Standard: NI = correctable unsafe condition with limited exposure while access/function remains usable. NC = required safety control failed/unavailable, access materially compromised, or significant injury exposure exists.');

  body.appendParagraph('FINDINGS').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const rows = [['Area','Finding','Category','Level']];
  payload.findings.forEach(f => rows.push([displayArea_(f.area),f.desc||'',f.category||'',f.level||'']));
  const findingsTable = body.appendTable(rows);
  styleHeader_(findingsTable,0);

  body.appendParagraph('SCORING CALCULATION').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const calc = body.appendTable([
    ['Needs Improvement (NI)', score.ni+' x 0.25', (score.ni*0.25).toFixed(2)],
    ['Non-Compliant (NC)', score.nc+' x 1.00', score.nc.toFixed(2)],
    ['Total Deduction','',score.deduction.toFixed(2)],
    ['Safety Performance Index','100 - '+score.deduction.toFixed(2),score.spi.toFixed(2)+'%']
  ]);
  calc.getRow(3).editAsText().setBold(true);

  doc.saveAndClose();
  return doc;
}

function buildSlides_(payload, score, baseName, folder, summaryUrl) {
  const pres = SlidesApp.create(baseName + ' - Weekly Safety Audit');
  DriveApp.getFileById(pres.getId()).moveTo(folder);
  const dateLabel = formatDate_(payload.weekEnding);
  const cover = pres.getSlides()[0];
  cover.getBackground().setSolidFill('#0B3A66');
  text_(cover,'WEEKLY SAFETY AUDIT\n'+dateLabel,42,52,620,90,28,true,'#FFFFFF');
  text_(cover,'Volleman Dairy Processing LLC',42,150,620,35,16,false,'#FFFFFF');
  box_(cover,390,240,270,150,'#FFFFFF');
  text_(cover,'SAFETY PERFORMANCE INDEX',412,260,230,25,12,true,'#0B3A66');
  text_(cover,score.spi.toFixed(2)+'%',412,292,230,52,30,true,score.spi>=95?'#137333':'#B42318');
  text_(cover,'Leadership Target: 95%',412,352,230,25,12,false,'#0B3A66');

  const purpose = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  titleBar_(purpose,'What This Audit Is For');
  bullets_(purpose,[
    'Identify safety hazards before they result in injury, damage, or compliance exposure.',
    'Give each team a clear list of conditions that need correction.',
    'Track repeat findings and verify that corrective actions are actually closed.',
    'Build a consistent weekly safety standard across the plant.',
    'Improve the score by removing actual hazards — not by changing how findings are counted.'
  ],55,118,610,250);

  const cat = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  titleBar_(cat,'Categories & Examples');
  const cats=[
    ['SW','Safe Work Practices','PPE, safe machine operation, LOTO when observed, ladders, lifting, chemical handling'],
    ['HH','Housekeeping & Hazard Control','Slips/trips, trash/debris, cords/hoses, clutter, unsafe general storage'],
    ['EP','Emergency Preparedness','Exits/access, signs/lights, extinguishers, eyewash/showers, first aid'],
    ['EF','Equipment & Facility Safety','Guards, damaged equipment, electrical, unsafe leaks, floors/stairs/doors, lighting'],
    ['TM','Traffic & Material Handling','Forklift/pedestrian interaction, aisles, staging, pallet stability, docks, vehicles'],
    ['FS','Food Safety Support','Safety + product/sanitary interface, chemicals, glass/brittle hazards, waste/sanitation']
  ];
  cats.forEach((c,i)=>{
    const y=88+i*55;
    text_(cat,c[0],50,y,45,28,14,true,'#FFFFFF','#0B3A66');
    text_(cat,c[1],110,y-2,210,24,13,true,'#0B3A66');
    text_(cat,c[2],325,y-2,350,38,11,false,'#333333');
  });
  text_(cat,'COMPLIANT = no action    NI = -0.25    NC = -1.00',50,425,620,26,11,true,'#0B3A66');

  AREA_SLIDE_GROUPS.forEach(group => {
    const slide=pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    titleBar_(slide,group.title);
    const findings=payload.findings.filter(f=>group.areas.indexOf(f.area)>=0);
    if(findings.some(f=>f.level==='NC')) text_(slide,'PRIORITY AREA',545,24,125,26,11,true,'#FFFFFF','#B42318');

    if (!findings.length) {
      text_(slide,'No findings noted this week.',48,112,380,42,14,false,'#64748B');
    } else {
      let y=96;
      findings.forEach((f,idx)=>{
        const label = displayArea_(f.area) + ' — ' + (f.desc||'');
        text_(slide,(idx+1)+'. '+label,45,y,390,52,12.5,false,'#1F2937');
        text_(slide,(f.category||'')+' / '+(f.level||''),45,y+47,120,22,10.5,true,f.level==='NC'?'#B42318':'#8A5A00');
        if(f.repeat) text_(slide,'Repeat',173,y+47,65,22,10,true,'#64748B');
        y+=78;
      });
    }

    const photos=[];
    findings.forEach(f=>(f.photos||[]).forEach(p=>{ if(photos.length<2) photos.push(p); }));
    if (photos.length===1) {
      insertPhoto_(slide,photos[0],462,115,205,245,'audit-photo-1.jpg');
    } else if (photos.length>=2) {
      insertPhoto_(slide,photos[0],462,105,205,145,'audit-photo-1.jpg');
      insertPhoto_(slide,photos[1],462,270,205,145,'audit-photo-2.jpg');
    }
  });

  const expectations=pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  titleBar_(expectations,'Plant-Wide Safety Expectations');
  const ex=[
    ['EXIT ACCESS','Keep emergency exits and access routes clear at all times.'],
    ['FIRE EQUIPMENT','Keep extinguishers visible and immediately accessible.'],
    ['TRAFFIC LANES','Keep pedestrian, pallet-jack, and forklift travel lanes open.'],
    ['STORAGE','Return tools, parts, equipment, crates, and materials to safe designated locations.'],
    ['WALKING SURFACES','Remove preventable trip hazards, trash, debris, leaks, and ice.']
  ];
  ex.forEach((r,i)=>{
    const y=92+i*64;
    text_(expectations,r[0],48,y,160,30,11,true,'#FFFFFF','#0B3A66');
    text_(expectations,r[1],225,y,430,38,12,false,'#1F2937');
  });

  const scoring=pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  titleBar_(scoring,'Scoring Notes');
  bullets_(scoring,[
    'Leadership target: 95% or greater.',
    'Needs Improvement (NI) = -0.25.',
    'Non-Compliant (NC) = -1.00.',
    'Repeat findings receive the normal weekly deduction and are tracked separately.',
    'Regardless of score, identified hazards still require correction.'
  ],55,105,600,230);
  text_(scoring,'Final SPI: '+score.spi.toFixed(2)+'%',55,350,260,40,22,true,score.spi>=95?'#137333':'#B42318');
  if (summaryUrl) {
    text_(scoring,'Weekly Summary of Findings & Score',55,405,280,25,12,true,'#0B3A66');
    text_(scoring,summaryUrl,55,432,600,34,9,false,'#2563EB');
  }

  const slides = pres.getSlides();
  slides.forEach((s,i)=>footer_(s,i+1,slides.length));
  return pres;
}

function testGenerateReports() {
  const folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  const payload = {
    weekEnding: '2026-10-23-TEST',
    findings: [
      {area:'Processing',desc:'Test safety finding',category:'SW',level:'NC',repeat:false,notes:'',photos:[]},
      {area:'Warehouse',desc:'Test needs improvement finding',category:'HH',level:'NI',repeat:false,notes:'',photos:[]}
    ]
  };
  const score = score_(payload.findings);
  const baseName = 'VDP Safety Audit - TEST';
  const doc = buildSummaryDoc_(payload, score, baseName, folder);
  buildSlides_(payload, score, baseName, folder, doc.getUrl());
}

function primaryFocus_(findings){
  if (!findings.length) return 'No findings this week. Maintain current controls and continue routine verification.';
  const all=(findings.map(f=>(f.desc||'')+' '+(f.area||'')+' '+(f.category||'')).join(' ')).toLowerCase();
  const parts=[];
  if (findings.some(f=>f.category==='EP')) parts.push('restore clear emergency equipment and access');
  if (findings.some(f=>f.category==='TM' || /traffic|aisle|path|forklift|dock/.test((f.desc||'').toLowerCase()))) parts.push('keep traffic and material-movement paths clear');
  if (findings.some(f=>f.category==='SW' && /ppe|eyewear|glasses|goggles/.test((f.desc||'').toLowerCase()))) parts.push('reinforce task-based PPE and eyewear compliance');
  if (/cord|hose/.test(all)) parts.push('improve control of cords and hoses in walking areas');
  if (/pallet|ladder|storage|staging/.test(all)) parts.push('improve storage and staging discipline');
  if (/glass|debris/.test(all)) parts.push('remove glass and debris from dock/walking areas');
  if (!parts.length) return 'Correct the identified conditions, verify closure, and watch for repeat findings during the next weekly audit.';
  return capitalize_(parts.join('; ') + '.');
}

function repeatConcernLines_(findings){
  const reps=findings.filter(f=>f.repeat);
  if(!reps.length) return [];
  const text=reps.map(f=>(f.desc||'')+' '+(f.area||'')).join(' ').toLowerCase();
  const out=[];
  if(/extinguisher|fire equipment/.test(text)) out.push('Blocked fire extinguisher or emergency-equipment access continues to recur.');
  if(/ppe|eyewear|goggle|glasses/.test(text)) out.push('PPE / eyewear compliance continues to recur; severity should reflect the task exposure observed.');
  if(/cord|hose/.test(text)) out.push('Cords or hoses in walking paths remain a recurring housekeeping concern.');
  if(/glass/.test(text) && /dock|loading/.test(text)) out.push('Loading-dock glass remains a recurring condition and should be corrected before exposure increases.');
  reps.forEach(f=>{
    const desc=(f.desc||'').toLowerCase();
    const covered=(/extinguisher|fire equipment|ppe|eyewear|goggle|glasses|cord|hose/.test(desc) || (/glass/.test(desc) && /dock|loading/.test((f.area||'').toLowerCase())));
    if(!covered) out.push(displayArea_(f.area)+' — '+(f.desc||''));
  });
  return [...new Set(out)];
}

function displayArea_(area){
  if(area==='Cooler / New Cooler') return 'Cooler / New Cooler';
  if(area==='BIB Room') return 'BIB Room';
  return area||'';
}

function formatDate_(value){
  const m=String(value||'').match(/^(\d{4})-(\d{2})-(\d{2})(.*)$/);
  if(!m) return value||'';
  return Number(m[2])+'/'+Number(m[3])+'/'+m[1]+(m[4]||'');
}

function capitalize_(s){return s ? s.charAt(0).toUpperCase()+s.slice(1) : s;}
function styleHeader_(table,rowIndex){const row=table.getRow(rowIndex);for(let i=0;i<row.getNumCells();i++){row.getCell(i).setBackgroundColor('#0B3A66');row.getCell(i).editAsText().setForegroundColor('#FFFFFF').setBold(true);}}
function titleBar_(slide,title){const s=slide.insertShape(SlidesApp.ShapeType.RECTANGLE,0,0,720,68);s.getFill().setSolidFill('#0B3A66');s.getBorder().setTransparent();text_(slide,title,38,17,560,36,22,true,'#FFFFFF');}
function bullets_(slide,items,x,y,w,h){const b=slide.insertTextBox(items.map(s=>'• '+s).join('\n'),x,y,w,h);b.getText().getTextStyle().setFontSize(15).setForegroundColor('#1F2937');}
function box_(slide,x,y,w,h,fill){const s=slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE,x,y,w,h);s.getFill().setSolidFill(fill);s.getBorder().setTransparent();return s;}
function text_(slide,text,x,y,w,h,size,bold,color,fill){let b;if(fill){b=slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE,x,y,w,h);b.getFill().setSolidFill(fill);b.getBorder().setTransparent();b.getText().setText(text);}else{b=slide.insertTextBox(text,x,y,w,h);}b.getText().getTextStyle().setFontSize(size).setBold(!!bold).setForegroundColor(color||'#1F2937');return b;}
function footer_(slide,index,total){text_(slide,'VDP Safety Performance System | Weekly Safety Audit | Slide '+index+' of '+total,38,514,640,18,8,false,'#64748B');}
function insertPhoto_(slide,dataUrl,x,y,w,h,name){try{slide.insertImage(dataUrlToBlob_(dataUrl,name),x,y,w,h);}catch(e){text_(slide,'Photo could not be inserted',x,y+40,w,35,10,false,'#64748B');}}
function dataUrlToBlob_(dataUrl,name){const parts=dataUrl.split(',');const mime=(parts[0].match(/data:(.*?);base64/)||[])[1]||'image/jpeg';return Utilities.newBlob(Utilities.base64Decode(parts[1]),mime,name);}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
