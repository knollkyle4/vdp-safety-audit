const DESTINATION_FOLDER_ID = '1ZZ-zfubeyNITNoHS_w7GfB4a2zxVT2yX';

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const expectedToken = PropertiesService.getScriptProperties().getProperty('VDP_UPLOAD_TOKEN');
    if (!expectedToken || payload.token !== expectedToken) {
      return json_({ok:false,error:'Unauthorized'});
    }
    if (!payload.weekEnding || !Array.isArray(payload.findings)) {
      return json_({ok:false,error:'Missing weekEnding or findings'});
    }

    const folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const score = score_(payload.findings);
    const baseName = 'VDP Safety Audit - Week Ending ' + payload.weekEnding;
    const doc = buildSummaryDoc_(payload, score, baseName, folder);
    const slides = buildSlides_(payload, score, baseName, folder);

    return json_({ok:true,docUrl:doc.getUrl(),slidesUrl:slides.getUrl(),spi:score.spi,findings:payload.findings.length});
  } catch (err) {
    return json_({ok:false,error:String(err && err.message ? err.message : err)});
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

  body.appendParagraph('VOLLEMAN DAIRY PROCESSING LLC').setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('WEEKLY SAFETY AUDIT - FINDINGS & SCORE').setHeading(DocumentApp.ParagraphHeading.HEADING2).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Week Ending ' + payload.weekEnding).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const status = score.spi >= 95 ? 'PASS' : 'BELOW TARGET';
  const metrics = body.appendTable([
    ['SPI','Target','Findings','Status'],
    [score.spi.toFixed(2)+'%','95.00%',String(payload.findings.length),status]
  ]);
  styleHeader_(metrics,0);

  body.appendParagraph('Findings').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const rows = [['Area','Finding','Category','Level']];
  payload.findings.forEach(f => rows.push([f.area||'',f.desc||'',f.category||'',f.level||'']));
  const findingsTable = body.appendTable(rows);
  styleHeader_(findingsTable,0);

  body.appendParagraph('Scoring').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Needs Improvement (NI): '+score.ni+' × 0.25 = '+(score.ni*0.25).toFixed(2));
  body.appendParagraph('Non-Compliant (NC): '+score.nc+' × 1.00 = '+score.nc.toFixed(2));
  body.appendParagraph('Total Deduction: '+score.deduction.toFixed(2));
  body.appendParagraph('Safety Performance Index: 100 - '+score.deduction.toFixed(2)+' = '+score.spi.toFixed(2)+'%');

  body.appendParagraph('Primary Focus').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(primaryFocus_(payload.findings));

  body.appendParagraph('Repeat / Recurring Concerns').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const repeats = payload.findings.filter(f=>f.repeat);
  if (repeats.length) repeats.forEach(f=>body.appendListItem((f.area||'')+' — '+(f.desc||'')));
  else body.appendParagraph('No repeat findings marked.');

  body.appendParagraph('Scoring Standard').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Compliant = no deduction. Needs Improvement (NI) = -0.25. Non-Compliant (NC) = -1.00. Leadership target = 95%.');
  doc.saveAndClose();
  return doc;
}

function buildSlides_(payload, score, baseName, folder) {
  const pres = SlidesApp.create(baseName + ' - Weekly Safety Audit');
  DriveApp.getFileById(pres.getId()).moveTo(folder);
  const cover = pres.getSlides()[0];
  cover.getBackground().setSolidFill('#0B3A66');
  text_(cover,'WEEKLY SAFETY AUDIT\n'+payload.weekEnding,42,52,620,90,28,true,'#FFFFFF');
  text_(cover,'Volleman Dairy Processing LLC',42,150,620,35,16,false,'#FFFFFF');
  box_(cover,390,240,270,150,'#FFFFFF');
  text_(cover,'SAFETY PERFORMANCE INDEX',412,260,230,25,12,true,'#0B3A66');
  text_(cover,score.spi.toFixed(2)+'%',412,292,230,52,30,true,score.spi>=95?'#137333':'#B42318');
  text_(cover,'Leadership Target: 95%',412,352,230,25,12,false,'#0B3A66');

  const purpose = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  titleBar_(purpose,'What This Audit Is For');
  bullets_(purpose,[
    'Identify safety exceptions and conditions that require action.',
    'Keep the weekly field process simple: observe more, write less.',
    'Classify each distinct condition once by its primary safety consequence.',
    'Track repeat findings separately without multiplying deductions.',
    'Use the weekly score as a trend indicator, not a substitute for correcting serious hazards.'
  ],55,120,600,250);

  const cat = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  titleBar_(cat,'Categories & Examples');
  const cats=[
    ['SW','Safe Work Practices','PPE, safe machine operation, LOTO when observed, ladder use, lifting, chemical handling'],
    ['HH','Housekeeping & Hazard Control','Slips/trips, trash/debris, cords/hoses, clutter, unsafe general storage'],
    ['EP','Emergency Preparedness','Exits, extinguishers, eyewash/showers, first aid, emergency access'],
    ['EF','Equipment & Facility Safety','Guards, damaged equipment, electrical, floors/stairs/doors, lighting'],
    ['TM','Traffic & Material Handling','Forklifts/pedestrians, aisles, staging, pallets, docks, vehicle movement'],
    ['FS','Food Safety Support','Safety conditions that also intersect product protection or sanitary operations']
  ];
  cats.forEach((c,i)=>{const y=95+i*58;text_(cat,c[0],50,y,45,28,14,true,'#FFFFFF','#0B3A66');text_(cat,c[1],110,y-2,210,24,13,true,'#0B3A66');text_(cat,c[2],325,y-2,350,36,11,false,'#333333');});

  const grouped={};
  payload.findings.forEach(f=>{(grouped[f.area]||(grouped[f.area]=[])).push(f);});
  Object.keys(grouped).forEach(area=>{
    const slide=pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    titleBar_(slide,area);
    const findings=grouped[area];
    if(findings.some(f=>f.level==='NC')) text_(slide,'PRIORITY AREA',545,24,125,26,11,true,'#FFFFFF','#B42318');
    let y=95;
    findings.forEach((f,idx)=>{text_(slide,(idx+1)+'. '+(f.desc||''),45,y,390,48,13,false,'#1F2937');text_(slide,(f.category||'')+' / '+(f.level||''),45,y+45,160,24,11,true,f.level==='NC'?'#B42318':'#8A5A00');if(f.repeat)text_(slide,'Repeat',212,y+45,70,24,10,true,'#64748B');y+=88;});
    const photoFinding=findings.find(f=>f.photos&&f.photos.length);
    if(photoFinding){try{slide.insertImage(dataUrlToBlob_(photoFinding.photos[0],'audit-photo.jpg'),455,108,215,215);}catch(e){text_(slide,'Photo could not be inserted',470,185,180,40,11,false,'#64748B');}}
  });

  const expectations=pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  titleBar_(expectations,'Plant-Wide Safety Expectations');
  const ex=[
    ['EXIT ACCESS','Emergency exits and exit paths must remain clear and immediately usable.'],
    ['FIRE EQUIPMENT','Fire extinguishers and emergency equipment must remain readily accessible.'],
    ['TRAFFIC LANES','Keep pedestrian, forklift, and material movement paths clear and predictable.'],
    ['STORAGE','Store pallets, tools, parts, and supplies in stable, designated locations.'],
    ['WALKING SURFACES','Control cords, hoses, liquids, debris, and other preventable trip/slip hazards.']
  ];
  ex.forEach((r,i)=>{const y=92+i*64;text_(expectations,r[0],48,y,160,30,11,true,'#FFFFFF','#0B3A66');text_(expectations,r[1],225,y,430,38,12,false,'#1F2937');});

  const scoring=pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  titleBar_(scoring,'Scoring Notes');
  bullets_(scoring,[
    'Start at 100%. Only findings deduct from the score.',
    'Needs Improvement (NI) = -0.25.',
    'Non-Compliant (NC) = -1.00.',
    'Leadership passing target = 95%.',
    'Repeat findings receive the normal weekly deduction and are tracked separately.',
    'A passing score does not override the need to correct serious individual findings.'
  ],55,110,600,260);
  text_(scoring,'Final SPI: '+score.spi.toFixed(2)+'%',55,400,260,40,22,true,score.spi>=95?'#137333':'#B42318');
  return pres;
}

function primaryFocus_(findings){const nc=[...new Set(findings.filter(f=>f.level==='NC').map(f=>f.area))];return nc.length?'Priority attention: '+nc.join(', ')+'. Restore failed controls, safe access, and significant injury exposures first.':'No NC findings. Focus on closing NI items before they repeat.';}
function styleHeader_(table,rowIndex){const row=table.getRow(rowIndex);for(let i=0;i<row.getNumCells();i++){row.getCell(i).setBackgroundColor('#0B3A66');row.getCell(i).editAsText().setForegroundColor('#FFFFFF').setBold(true);}}
function titleBar_(slide,title){const s=slide.insertShape(SlidesApp.ShapeType.RECTANGLE,0,0,720,68);s.getFill().setSolidFill('#0B3A66');s.getLine().setTransparent();text_(slide,title,38,17,560,36,23,true,'#FFFFFF');}
function bullets_(slide,items,x,y,w,h){const b=slide.insertTextBox(items.map(s=>'• '+s).join('\n'),x,y,w,h);b.getText().getTextStyle().setFontSize(15).setForegroundColor('#1F2937');}
function box_(slide,x,y,w,h,fill){const s=slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE,x,y,w,h);s.getFill().setSolidFill(fill);s.getLine().setTransparent();return s;}
function text_(slide,text,x,y,w,h,size,bold,color,fill){let b;if(fill){b=slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE,x,y,w,h);b.getFill().setSolidFill(fill);b.getLine().setTransparent();b.getText().setText(text);}else b=slide.insertTextBox(text,x,y,w,h);b.getText().getTextStyle().setFontSize(size).setBold(!!bold).setForegroundColor(color||'#1F2937');return b;}
function dataUrlToBlob_(dataUrl,name){const parts=dataUrl.split(',');const mime=(parts[0].match(/data:(.*?);base64/)||[])[1]||'image/jpeg';return Utilities.newBlob(Utilities.base64Decode(parts[1]),mime,name);}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
