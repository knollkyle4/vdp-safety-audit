const DESTINATION_FOLDER_ID = '1ZZ-zfubeyNITNoHS_w7GfB4a2zxVT2yX';

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    if (!payload.weekEnding || !Array.isArray(payload.findings)) {
      return json_({ok:false,error:'Missing weekEnding or findings'});
    }

    const folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const score = score_(payload.findings);
    const baseName = 'VDP Safety Audit - Week Ending ' + payload.weekEnding;

    const doc = buildSummaryDoc_(payload, score, baseName, folder);
    const slides = buildSlides_(payload, score, baseName, folder);

    return json_({
      ok:true,
      docUrl:doc.getUrl(),
      slidesUrl:slides.getUrl(),
      spi:score.spi,
      findings:payload.findings.length
    });
  } catch (err) {
    return json_({ok:false,error:String(err && err.message ? err.message : err)});
  }
}

function score_(findings) {
  const ni = findings.filter(f => f.level === 'NI').length;
  const nc = findings.filter(f => f.level === 'NC').length;
  const deduction = ni * 0.25 + nc * 1.0;
  return {ni, nc, deduction, spi: Math.max(0, 100 - deduction)};
}

function buildSummaryDoc_(payload, score, baseName, folder) {
  const doc = DocumentApp.create(baseName + ' - Summary');
  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(folder);
  const body = doc.getBody();

  body.appendParagraph('VOLLEMAN DAIRY PROCESSING LLC')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('WEEKLY SAFETY AUDIT - FINDINGS & SCORE')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Week Ending ' + payload.weekEnding)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const status = score.spi >= 95 ? 'PASS' : 'BELOW TARGET';
  const metrics = body.appendTable([
    ['SPI', 'Target', 'Findings', 'Status'],
    [score.spi.toFixed(2) + '%', '95.00%', String(payload.findings.length), status]
  ]);
  styleTableHeader_(metrics, 0);

  body.appendParagraph('Findings').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const rows = [['Area','Finding','Category','Level']];
  payload.findings.forEach(f => rows.push([f.area || '', f.desc || '', f.category || '', f.level || '']));
  const table = body.appendTable(rows);
  styleTableHeader_(table, 0);

  body.appendParagraph('Scoring').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Needs Improvement (NI): ' + score.ni + ' × 0.25 = ' + (score.ni*0.25).toFixed(2));
  body.appendParagraph('Non-Compliant (NC): ' + score.nc + ' × 1.00 = ' + score.nc.toFixed(2));
  body.appendParagraph('Total Deduction: ' + score.deduction.toFixed(2));
  body.appendParagraph('Safety Performance Index: 100 - ' + score.deduction.toFixed(2) + ' = ' + score.spi.toFixed(2) + '%');

  body.appendParagraph('Primary Focus').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(primaryFocus_(payload.findings));

  body.appendParagraph('Repeat / Recurring Concerns').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const repeats = payload.findings.filter(f => f.repeat);
  if (repeats.length) {
    repeats.forEach(f => body.appendListItem((f.area || '') + ' — ' + (f.desc || '')));
  } else {
    body.appendParagraph('No repeat findings marked.');
  }

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
  addText_(cover, 'WEEKLY SAFETY AUDIT\n' + payload.weekEnding, 42, 52, 620, 90, 28, true, '#FFFFFF');
  addText_(cover, 'Volleman Dairy Processing LLC', 42, 150, 620, 35, 16, false, '#FFFFFF');
  addBox_(cover, 390, 240, 270, 150, '#FFFFFF');
  addText_(cover, 'SAFETY PERFORMANCE INDEX', 412, 260, 230, 25, 12, true, '#0B3A66');
  addText_(cover, score.spi.toFixed(2) + '%', 412, 292, 230, 52, 30, true, score.spi >= 95 ? '#137333' : '#B42318');
  addText_(cover, 'Leadership Target: 95%', 412, 352, 230, 25, 12, false, '#0B3A66');

  const purpose = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addTitleBar_(purpose, 'What This Audit Is For');
  addBulletBlock_(purpose, [
    'Identify safety exceptions and conditions that require action.',
    'Keep the weekly field process simple: observe more, write less.',
    'Classify each distinct condition once by its primary safety consequence.',
    'Track repeat findings separately without multiplying deductions.',
    'Use the weekly score as a trend indicator, not a substitute for correcting serious hazards.'
  ], 55, 120, 600, 250);

  const cat = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addTitleBar_(cat, 'Categories & Examples');
  const cats = [
    ['SW','Safe Work Practices','PPE, safe machine operation, LOTO when observed, ladder use, lifting, chemical handling'],
    ['HH','Housekeeping & Hazard Control','Slips/trips, trash/debris, cords/hoses, clutter, unsafe general storage'],
    ['EP','Emergency Preparedness','Exits, extinguishers, eyewash/showers, first aid, emergency access'],
    ['EF','Equipment & Facility Safety','Guards, damaged equipment, electrical, floors/stairs/doors, lighting'],
    ['TM','Traffic & Material Handling','Forklifts/pedestrians, aisles, staging, pallets, docks, vehicle movement'],
    ['FS','Food Safety Support','Safety conditions that also intersect product protection or sanitary operations']
  ];
  cats.forEach((c,i)=>{
    const y=95+i*58;
    addText_(cat, c[0], 50, y, 45, 28, 14, true, '#FFFFFF', '#0B3A66');
    addText_(cat, c[1], 110, y-2, 210, 24, 13, true, '#0B3A66');
    addText_(cat, c[2], 325, y-2, 350, 36, 11, false, '#333333');
  });

  const grouped = {};
  payload.findings.forEach(f => {
    if (!grouped[f.area]) grouped[f.area] = [];
    grouped[f.area].push(f);
  });

  Object.keys(grouped).forEach(area => {
    const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    addTitleBar_(slide, area);
    const findings = grouped[area];
    if (findings.some(f => f.level === 'NC')) {
      addText_(slide, 'PRIORITY AREA', 545, 24, 125, 26, 11, true, '#FFFFFF', '#B42318');
    }
    let y = 95;
    findings.forEach((f, idx) => {
      addText_(slide, (idx+1) + '. ' + (f.desc || ''), 45, y, 390, 48, 13, false, '#1F2937');
      addText_(slide, (f.category || '') + ' / ' + (f.level || ''), 45, y+45, 160, 24, 11, true, f.level === 'NC' ? '#B42318' : '#8A5A00');
      if (f.repeat) addText_(slide, 'Repeat', 212, y+45, 70, 24, 10, true, '#64748B');
      y += 88;
    });
    const photoFinding = findings.find(f => f.photos && f.photos.length);
    if (photoFinding) {
      try {
        const blob = dataUrlToBlob_(photoFinding.photos[0], 'audit-photo.jpg');
        slide.insertImage(blob, 455, 108, 215, 215);
      } catch (e) {
        addText_(slide, 'Photo could not be inserted', 470, 185, 180, 40, 11, false, '#64748B');
      }
    }
  });

  const expectations = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addTitleBar_(expectations, 'Plant-Wide Safety Expectations');
  const ex = [
    ['EXIT ACCESS','Emergency exits and exit paths must remain clear and immediately usable.'],
    ['FIRE EQUIPMENT','Fire extinguishers and emergency equipment must remain readily accessible.'],
    ['TRAFFIC LANES','Keep pedestrian, forklift, and material movement paths clear and predictable.'],
    ['STORAGE','Store pallets, tools, parts, and supplies in stable, designated locations.'],
    ['WALKING SURFACES','Control cords, hoses, liquids, debris, and other preventable trip/slip hazards.']
  ];
  ex.forEach((r,i)=>{
    const y=92+i*64;
    addText_(expectations, r[0], 48, y, 160, 30, 11, true, '#FFFFFF', '#0B3A66');
    addText_(expectations, r[1], 225, y, 430, 38, 12, false, '#1F2937');
  });

  const scoring = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addTitleBar_(scoring, 'Scoring Notes');
  addBulletBlock_(scoring, [
    'Start at 100%. Only findings deduct from the score.',
    'Needs Improvement (NI) = -0.25.',
    'Non-Compliant (NC) = -1.00.',
    'Leadership passing target = 95%.',
    'Repeat findings receive the normal weekly deduction and are tracked separately.',
    'A passing score does not override the need to correct serious individual findings.'
  ], 55, 110, 600, 260);
  addText_(scoring, 'Final SPI: ' + score.spi.toFixed(2) + '%', 55, 400, 260, 40, 22, true, score.spi >= 95 ? '#137333' : '#B42318');

  return pres;
}

function primaryFocus_(findings) {
  const nc = [...new Set(findings.filter(f => f.level === 'NC').map(f => f.area))];
  if (!nc.length) return 'No NC findings. Focus on closing NI items before they repeat.';
  return 'Priority attention: ' + nc.join(', ') + '. Restore failed controls, safe access, and significant injury exposures first.';
}

function styleTableHeader_(table, rowIndex) {
  const row = table.getRow(rowIndex);
  for (let i=0;i<row.getNumCells();i++) {
    row.getCell(i).setBackgroundColor('#0B3A66');
    row.getCell(i).editAsText().setForegroundColor('#FFFFFF').setBold(true);
  }
}

function addTitleBar_(slide, title) {
  const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, 720, 68);
  shape.getFill().setSolidFill('#0B3A66');
  shape.getLine().setTransparent();
  addText_(slide, title, 38, 17, 560, 36, 23, true, '#FFFFFF');
}

function addBulletBlock_(slide, items, x, y, w, h) {
  const box = slide.insertTextBox(items.map(s=>'• '+s).join('\n'), x, y, w, h);
  const style = box.getText().getTextStyle();
  style.setFontSize(15).setForegroundColor('#1F2937');
}

function addBox_(slide, x, y, w, h, fill) {
  const shape = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, w, h);
  shape.getFill().setSolidFill(fill);
  shape.getLine().setTransparent();
  return shape;
}

function addText_(slide, text, x, y, w, h, size, bold, color, fill) {
  let box;
  if (fill) {
    box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, w, h);
    box.getFill().setSolidFill(fill);
    box.getLine().setTransparent();
    box.getText().setText(text);
  } else {
    box = slide.insertTextBox(text, x, y, w, h);
  }
  box.getText().getTextStyle().setFontSize(size).setBold(!!bold).setForegroundColor(color || '#1F2937');
  return box;
}

function dataUrlToBlob_(dataUrl, name) {
  const parts = dataUrl.split(',');
  const mime = (parts[0].match(/data:(.*?);base64/) || [])[1] || 'image/jpeg';
  const bytes = Utilities.base64Decode(parts[1]);
  return Utilities.newBlob(bytes, mime, name);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
