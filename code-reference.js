const SPREADSHEET_ID = '1AqKb6AyaIrHC4Qrc_m7QT98L1EOetHlw';
// Папка Google Drive для фото профилей (создать вручную, вставить ID из URL папки)
const DRIVE_FOLDER_ID = '1zkAIOO3EembU5lEil0gIfJdx5TG3iN-j';

function doGet(e){
  try {
    const action = (e.parameter.action || 'artists').toLowerCase();
    let data = [];
    if(action === 'artists')            data = getSheetData('Artists', 'artist');
    else if(action === 'venues')        data = getSheetDataAny(['Venues','Площадки'], 'venue');
    else if(action === 'events')        data = getSheetDataAny(['Мероприятия','Events'], 'event');
    else if(action === 'persons')       data = getSheetData('Persons', 'person');
    else if(action === 'pricing')       data = getSheetData('ArtistPricing', 'pricing');
    else if(action === 'reviews')       data = getSheetData('Reviews', 'review');
    else if(action === 'applications')  data = getSheetData('Applications', 'application');
    else if(action === 'projects')      data = getSheetData('Projects', 'project');
    else if(action === 'projectResponses') data = getSheetData('ProjectResponses', 'projectResponse');
    else if(action === 'venuePricing')  data = getSheetData('VenuePricing', 'venuePricing');
    else if(action === 'venueCategoryPricing') data = getSheetData('VenueCategoryPricing', 'venueCategoryPricing');
    else if(action === 'artistSchedule') data = getRawSchedule();
    else if(action === 'media')         data = getSheetData('ArtistMedia', 'media');
    else if(action === 'experience')    data = getSheetData('ArtistExperience', 'experience');
    else if(action === 'education')     data = getSheetData('ArtistEducation', 'education');
    else if(action === 'inventory')     data = getSheetData('ArtistInventory', 'inventory');
    else if(action === 'costumes')      data = getSheetData('ArtistCostumes', 'costume');
    else return jsonResponse({error:'Unknown action: '+action});
    return jsonResponse(data);
  } catch(err){ return jsonResponse({error:err.toString()}); }
}

function doPost(e){
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    // Загрузка фото в Drive. Обрабатывается до openById: на этом шаге лист не нужен
    if(action === 'uploadPhoto'){
      if(!DRIVE_FOLDER_ID || DRIVE_FOLDER_ID === '1zkAIOO3EembU5lEil0gIfJdx5TG3iN-j'){
        return jsonResponse({error:'DRIVE_FOLDER_ID is not configured'});
      }
      if(!body.base64) return jsonResponse({error:'base64 is empty'});
      try {
        const url = uploadPhoto(body.base64, body.filename || 'photo.jpg');
        return jsonResponse({success:true, url:url});
      } catch(err){
        console.error('uploadPhoto failed: ' + err);
        return jsonResponse({error:err.toString()});
      }
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    if(action === 'updateArtist'){
      updateRow(ss, 'Artists', body.artistId, body.fields);
      return jsonResponse({success:true});
    }

    if(action === 'createArtist'){
      const sheet = ss.getSheetByName('Artists');
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const idCol = headers.indexOf('id');
      const tgCol = headers.indexOf('telegramId');
      const tgId = String(body.telegramId);

      if(tgCol > -1){
        for(let i=1; i<data.length; i++){
          if(String(data[i][tgCol]) === tgId){
            updateRow(ss, 'Artists', data[i][idCol], {
              name: body.name, role: body.role, spec: body.spec,
              city: body.city, price: body.price, contacts: body.contacts,
              img: body.img || ''
            });
            if(body.institution) addEducationRow(ss, data[i][idCol], body);
            return jsonResponse({success:true, artistId: data[i][idCol], existing:true});
          }
        }
      }
      let maxId = 0;
      for(let i=1; i<data.length; i++){
        const v = Number(data[i][idCol]) || 0;
        if(v > maxId) maxId = v;
      }
      const newId = maxId + 1;
      const row = new Array(headers.length).fill('');
      const mapping = {
        'id': newId, 'Name': body.name, 'name': body.name,
        'telegramId': tgId, 'role': body.role,
        'specialty': body.specialty || 'dance', 'spec': body.spec || '',
        'price': body.price || 0, 'img': body.img || '', 'city': body.city || '',
        'rating': 5.0, 'verified': false, 'contacts': body.contacts || '',
        'readyToTravel': body.readyToTravel || false, 'countries': body.countries || '',
        'travelScope': body.travelScope || 'russia',
        'visaType': body.visaType || 'none',
        'visaValidUntil': body.visaValidUntil || '',
        'needsVisaSupport': body.needsVisaSupport || false,
        'bio': body.bio || '',
        'vacationCity': '', 'vacationUntil': '', 'isPro': false
      };
      headers.forEach((h,i)=>{ if(mapping[h] !== undefined) row[i] = mapping[h]; });
      sheet.appendRow(row);
      if(body.institution) addEducationRow(ss, newId, body);
      return jsonResponse({success:true, artistId:newId});
    }

    if(action === 'saveArtistSchedule'){
  try {
    let sheet = ss.getSheetByName('ArtistSchedule');
    if(!sheet){
      sheet = ss.insertSheet('ArtistSchedule');
      sheet.appendRow(['artistId','dayOfWeek','isWorking','startHour','endHour','busyHours']);
    }
    const artistId = String(body.artistId).trim();
    const data = sheet.getDataRange().getValues();
    const aCol = data[0].indexOf('artistId');

    // Удаляем все строки этого артиста (с конца, чтобы не сбить индексы)
    for(let i=data.length-1; i>=1; i--){
      if(String(data[i][aCol]).trim() === artistId){
        sheet.deleteRow(i+1);
      }
    }

    // Готовим массив значений — одним setValues для скорости
    const rowsToAdd = body.rows.map(r => {
      const dow = Math.floor(Number(r.dayOfWeek));
      const sh = Math.floor(Number(r.startHour));
      const eh = Math.floor(Number(r.endHour));
      return [
        artistId,
        isFinite(dow) ? dow : 0,
        r.isWorking ? 'true' : 'false',
        isFinite(sh) ? sh : 9,
        isFinite(eh) ? eh : 21,
        ''
      ];
    });

    if(rowsToAdd.length){
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rowsToAdd.length, 6).setValues(rowsToAdd);
    }

    return jsonResponse({success:true, saved: rowsToAdd.length, artistId: artistId});
  } catch(err){
    return jsonResponse({error:'saveArtistSchedule: ' + err.toString()});
  }
}

    if(action === 'setArtistDayStatus'){
      let sheet = ss.getSheetByName('ArtistDayOverride');
      if(!sheet){
        sheet = ss.insertSheet('ArtistDayOverride');
        sheet.appendRow(['artistId','date','status','startHour','endHour','busyHours']);
      }
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const aCol = headers.indexOf('artistId');
      const dCol = headers.indexOf('date');
      const artistId = String(body.artistId);
      let found = false;
      for(let i=1; i<data.length; i++){
        const rowDate = data[i][dCol] instanceof Date
          ? Utilities.formatDate(data[i][dCol], Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(data[i][dCol]);
        if(String(data[i][aCol]) === artistId && rowDate === body.date){
          const stCol = headers.indexOf('status');
          sheet.getRange(i+1, stCol+1).setValue(body.status);
          found = true;
          break;
        }
      }
      if(!found) sheet.appendRow([artistId, body.date, body.status, '', '', '']);
      return jsonResponse({success:true});
    }

    if(action === 'updatePersonCity'){
      const sheet = ss.getSheetByName('Persons');
      let data = sheet.getDataRange().getValues();
      let headers = data[0];
      let cCol = headers.indexOf('city');
      if(cCol < 0){
        sheet.getRange(1, headers.length+1).setValue('city');
        data = sheet.getDataRange().getValues();
        headers = data[0];
        cCol = headers.indexOf('city');
      }
      const tgCol = headers.indexOf('telegramId');
      for(let i=1; i<data.length; i++){
        if(String(data[i][tgCol]) === String(body.telegramId)){
          sheet.getRange(i+1, cCol+1).setValue(body.city);
          return jsonResponse({success:true});
        }
      }
      return jsonResponse({error:'Person not found'});
    }

    if(action === 'createVenue'){
      let sheet = ss.getSheetByName('Venues') || ss.getSheetByName('Площадки');
      if(!sheet){
        sheet = ss.insertSheet('Venues');
        sheet.appendRow(['id','name','city','type','capacity','basePrice','seasonMultiplier','weekendMultiplier','img','contacts','ownerTelegramId','noBargain']);
      }
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const idCol = headers.indexOf('id');
      const tgCol = headers.indexOf('ownerTelegramId');
      const tgId = String(body.telegramId);
      if(tgCol > -1){
        for(let i=1; i<data.length; i++){
          if(String(data[i][tgCol]) === tgId){
            return jsonResponse({success:true, venueId: data[i][idCol], existing:true});
          }
        }
      }
      let maxId = 0;
      for(let i=1; i<data.length; i++){
        const v = Number(data[i][idCol]) || 0;
        if(v > maxId) maxId = v;
      }
      const newId = maxId + 1;
      const row = new Array(headers.length).fill('');
      const mapping = {
        'id': newId, 'name': body.name, 'Название': body.name,
        'city': body.city, 'Город': body.city,
        'type': body.type, 'Тип': body.type,
        'capacity': body.capacity, 'Вместимость': body.capacity,
        'basePrice': body.basePrice, 'price': body.basePrice, 'Аренда': body.basePrice,
        'seasonMultiplier': 1, 'weekendMultiplier': 1,
        'img': body.img, 'Фото': body.img,
        'contacts': body.contacts, 'Контакты': body.contacts,
        'ownerTelegramId': tgId, 'noBargain': false
      };
      headers.forEach((h,i)=>{ if(mapping[h] !== undefined) row[i] = mapping[h]; });
      sheet.appendRow(row);
      return jsonResponse({success:true, venueId:newId});
    }

    if(action === 'updateVenue'){
      let sheet = ss.getSheetByName('Venues') || ss.getSheetByName('Площадки');
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const idCol = headers.indexOf('id');
      for(let i=1; i<data.length; i++){
        if(String(data[i][idCol]) === String(body.venueId)){
          Object.keys(body.fields).forEach(key=>{
            const col = headers.indexOf(key);
            if(col > -1) sheet.getRange(i+1, col+1).setValue(body.fields[key]);
          });
          return jsonResponse({success:true});
        }
      }
      return jsonResponse({error:'Venue not found'});
    }

    if(action === 'createPerson'){
      let sheet = ss.getSheetByName('Persons');
      if(!sheet){
        sheet = ss.insertSheet('Persons');
        sheet.appendRow(['telegramId','name','photoUrl','activeRole','roles','city']);
      }
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const tgCol = headers.indexOf('telegramId');
      const tgId = String(body.telegramId);
      for(let i=1; i<data.length; i++){
        if(String(data[i][tgCol]) === tgId){
          return jsonResponse({success:true, existing:true, person: rowToObj(headers, data[i])});
        }
      }
      const row = new Array(headers.length).fill('');
      const mapping = {
        'telegramId': tgId, 'name': body.name || '',
        'photoUrl': body.photoUrl || '', 'activeRole': body.activeRole || 'viewer',
        'roles': body.roles || 'organizer,artist,venue,viewer', 'city': ''
      };
      headers.forEach((h,i)=>{ if(mapping[h] !== undefined) row[i] = mapping[h]; });
      sheet.appendRow(row);
      return jsonResponse({success:true, person: rowToObj(headers, row)});
    }

    if(action === 'setActiveRole'){
      const sheet = ss.getSheetByName('Persons');
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const tgCol = headers.indexOf('telegramId');
      const arCol = headers.indexOf('activeRole');
      const tgId = String(body.telegramId);
      for(let i=1; i<data.length; i++){
        if(String(data[i][tgCol]) === tgId){
          sheet.getRange(i+1, arCol+1).setValue(body.activeRole);
          return jsonResponse({success:true});
        }
      }
      return jsonResponse({error:'Person not found'});
    }

    if(action === 'setPricing'){
      let sheet = ss.getSheetByName('ArtistPricing');
      if(!sheet){
        sheet = ss.insertSheet('ArtistPricing');
        sheet.appendRow(['artistId','date','priceModifier','note']);
      }
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const aCol = headers.indexOf('artistId');
      const dCol = headers.indexOf('date');
      let found = false;
      for(let i=1; i<data.length; i++){
        if(String(data[i][aCol]) === String(body.artistId)){
          const rowDate = data[i][dCol] instanceof Date
            ? Utilities.formatDate(data[i][dCol], Session.getScriptTimeZone(), 'yyyy-MM-dd')
            : String(data[i][dCol]);
          if(rowDate === body.date){
            sheet.getRange(i+1, headers.indexOf('priceModifier')+1).setValue(body.priceModifier);
            sheet.getRange(i+1, headers.indexOf('note')+1).setValue(body.note || '');
            found = true; break;
          }
        }
      }
      if(!found) sheet.appendRow([body.artistId, body.date, body.priceModifier, body.note || '']);
      return jsonResponse({success:true});
    }

    if(action === 'addVenuePeriodPrice'){
      let sheet = ss.getSheetByName('VenuePricing');
      if(!sheet){
        sheet = ss.insertSheet('VenuePricing');
        sheet.appendRow(['id','venueId','dateFrom','dateTo','price','note']);
      }
      const data = sheet.getDataRange().getValues();
      let maxId = 0;
      for(let i=1; i<data.length; i++){
        const v = Number(data[i][0]) || 0;
        if(v > maxId) maxId = v;
      }
      sheet.appendRow([maxId+1, body.venueId, body.dateFrom, body.dateTo, body.price, body.note || '']);
      return jsonResponse({success:true});
    }

    if(action === 'deleteVenuePeriodPrice'){
      const sheet = ss.getSheetByName('VenuePricing');
      const data = sheet.getDataRange().getValues();
      for(let i=data.length-1; i>=1; i--){
        if(String(data[i][0]) === String(body.id)){
          sheet.deleteRow(i+1);
          return jsonResponse({success:true});
        }
      }
      return jsonResponse({error:'Not found'});
    }

    if(action === 'addVenueCategoryPrice'){
      let sheet = ss.getSheetByName('VenueCategoryPricing');
      if(!sheet){
        sheet = ss.insertSheet('VenueCategoryPricing');
        sheet.appendRow(['id','venueId','eventType','multiplier','note']);
      }
      const data = sheet.getDataRange().getValues();
      let maxId = 0;
      for(let i=1; i<data.length; i++){
        const v = Number(data[i][0]) || 0;
        if(v > maxId) maxId = v;
      }
      sheet.appendRow([maxId+1, body.venueId, body.eventType, body.multiplier, body.note || '']);
      return jsonResponse({success:true});
    }

    if(action === 'deleteVenueCategoryPrice'){
      const sheet = ss.getSheetByName('VenueCategoryPricing');
      const data = sheet.getDataRange().getValues();
      for(let i=data.length-1; i>=1; i--){
        if(String(data[i][0]) === String(body.id)){
          sheet.deleteRow(i+1);
          return jsonResponse({success:true});
        }
      }
      return jsonResponse({error:'Not found'});
    }

    if(action === 'submitApplication'){
      let sheet = ss.getSheetByName('Applications');
      if(!sheet){
        sheet = ss.insertSheet('Applications');
        sheet.appendRow(['id','createdAt','organizerId','organizerName','artistId','artistName','venueId','eventType','date','city','guests','equipment','budget','status','venueResponse','venuePrice','venueNote','type']);
      }
      const data = sheet.getDataRange().getValues();
      let maxId = 0;
      for(let i=1; i<data.length; i++){
        const v = Number(data[i][0]) || 0;
        if(v > maxId) maxId = v;
      }
      const newId = maxId + 1;
      sheet.appendRow([
        newId, new Date(),
        String(body.organizerId || ''), body.organizerName || '',
        String(body.artistId || ''), body.artistName || '',
        String(body.venueId || ''),
        body.eventType || '', body.date || '', body.city || '',
        body.guests || 0, body.equipment || '', body.budget || 0,
        'new', '', '', '', body.type || 'artist'
      ]);
      return jsonResponse({success:true, applicationId:newId});
    }

    if(action === 'updateApplicationStatus'){
      const sheet = ss.getSheetByName('Applications');
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const idCol = headers.indexOf('id');
      const stCol = headers.indexOf('status');
      for(let i=1; i<data.length; i++){
        if(String(data[i][idCol]) === String(body.applicationId)){
          sheet.getRange(i+1, stCol+1).setValue(body.status);
          return jsonResponse({success:true});
        }
      }
      return jsonResponse({error:'Not found'});
    }

    if(action === 'venueRespondApplication'){
      const sheet = ss.getSheetByName('Applications');
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const idCol = headers.indexOf('id');
      const respCol = headers.indexOf('venueResponse');
      const priceCol = headers.indexOf('venuePrice');
      const noteCol = headers.indexOf('venueNote');
      for(let i=1; i<data.length; i++){
        if(String(data[i][idCol]) === String(body.applicationId)){
          sheet.getRange(i+1, respCol+1).setValue(body.response);
          if(priceCol > -1) sheet.getRange(i+1, priceCol+1).setValue(body.price || '');
          if(noteCol > -1) sheet.getRange(i+1, noteCol+1).setValue(body.note || '');
          return jsonResponse({success:true});
        }
      }
      return jsonResponse({error:'Not found'});
    }

    if(action === 'createProject'){
      let sheet = ss.getSheetByName('Projects');
      if(!sheet){
        sheet = ss.insertSheet('Projects');
        sheet.appendRow(['id','createdAt','organizerId','organizerName','title','type','startDate','endDate','cities','budget','description','isOpen','artistIds','status']);
      }
      const data = sheet.getDataRange().getValues();
      let maxId = 0;
      for(let i=1; i<data.length; i++){
        const v = Number(data[i][0]) || 0;
        if(v > maxId) maxId = v;
      }
      const newId = maxId + 1;
      sheet.appendRow([
        newId, new Date(), String(body.organizerId || ''),
        body.organizerName || '', body.title || '', body.type || 'single',
        body.startDate || '', body.endDate || '', body.cities || '',
        body.budget || 0, body.description || '',
        body.isOpen ? 'true' : 'false', body.artistIds || '', 'open'
      ]);
      return jsonResponse({success:true, projectId:newId});
    }

    if(action === 'updateProjectStatus'){
      const sheet = ss.getSheetByName('Projects');
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const idCol = headers.indexOf('id');
      const stCol = headers.indexOf('status');
      for(let i=1; i<data.length; i++){
        if(String(data[i][idCol]) === String(body.projectId)){
          sheet.getRange(i+1, stCol+1).setValue(body.status);
          return jsonResponse({success:true});
        }
      }
      return jsonResponse({error:'Not found'});
    }

    if(action === 'respondToProject'){
      let sheet = ss.getSheetByName('ProjectResponses');
      if(!sheet){
        sheet = ss.insertSheet('ProjectResponses');
        sheet.appendRow(['id','createdAt','projectId','artistId','artistName','price','comment','status']);
      }
      const data = sheet.getDataRange().getValues();
      for(let i=1; i<data.length; i++){
        if(String(data[i][2]) === String(body.projectId) && String(data[i][3]) === String(body.artistId)){
          return jsonResponse({success:true, existing:true});
        }
      }
      let maxId = 0;
      for(let i=1; i<data.length; i++){
        const v = Number(data[i][0]) || 0;
        if(v > maxId) maxId = v;
      }
      const newId = maxId + 1;
      sheet.appendRow([
        newId, new Date(), String(body.projectId || ''), String(body.artistId || ''),
        body.artistName || '', body.price || 0, body.comment || '', 'pending'
      ]);
      return jsonResponse({success:true, responseId:newId});
    }

    if(action === 'updateProjectResponse'){
      const sheet = ss.getSheetByName('ProjectResponses');
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const idCol = headers.indexOf('id');
      const stCol = headers.indexOf('status');
      for(let i=1; i<data.length; i++){
        if(String(data[i][idCol]) === String(body.responseId)){
          sheet.getRange(i+1, stCol+1).setValue(body.status);
          return jsonResponse({success:true});
        }
      }
      return jsonResponse({error:'Not found'});
    }

    const appendMap = {
      addMedia:['ArtistMedia', b=>[b.artistId, b.type||'photo', b.url||b.name, b.caption||b.description||'']],
      addExperience:['ArtistExperience', b=>[b.artistId, b.year||'', b.project||b.name, b.role||'', b.venue||'', b.city||'']],
      addEducation:['ArtistEducation', b=>[b.artistId, b.year||'', b.institution||b.name, b.specialty||b.description||'']],
      addInventory:['ArtistInventory', b=>[b.artistId, b.name, b.description||'', b.photoUrl||b.url||'']],
      addCostume:['ArtistCostumes', b=>[b.artistId, b.name, b.size||'', b.description||'', b.photoUrl||b.url||'']],
      addReview:['Reviews', b=>[b.artistId, b.author||'', b.rating||5, b.text||'', b.date||new Date()]]
    };
    if(appendMap[action]){
      const [sheetName, rowFn] = appendMap[action];
      ss.getSheetByName(sheetName).appendRow(rowFn(body));
      return jsonResponse({success:true});
    }

    return jsonResponse({error:'Unknown action: '+action});
  } catch(err){ return jsonResponse({error:err.toString()}); }
}

/* === Прямое чтение ArtistSchedule === */
function getRawSchedule(){
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ArtistSchedule');
  if(!sheet) return [];
  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return [];
  const values = sheet.getRange(1, 1, lastRow, 6).getValues();
  const headers = values[0];
  const aCol = headers.indexOf('artistId');
  const dCol = headers.indexOf('dayOfWeek');
  const wCol = headers.indexOf('isWorking');
  const sCol = headers.indexOf('startHour');
  const eCol = headers.indexOf('endHour');
  const bCol = headers.indexOf('busyHours');

  const result = [];
  for(let i=1; i<values.length; i++){
    const r = values[i];
    if(!r[aCol]) continue;
    const toNum = (v, fallback) => {
      // Date здесь означает, что колонка отформатирована как дата/время в Sheets —
      // getHours() вернул бы час, а не день недели или час начала
      if(v instanceof Date){
        return fallback;
      }
      const n = Number(v);
      return isFinite(n) ? Math.floor(n) : fallback;
    };
    result.push({
      artistId: String(r[aCol]).trim(),
      dayOfWeek: toNum(r[dCol], 0),
      isWorking: String(r[wCol]).toLowerCase() === 'true',
      startHour: toNum(r[sCol], 9),
      endHour: toNum(r[eCol], 21),
      busyHours: String(r[bCol] || '')
    });
  }
  return result;
}

function addEducationRow(ss, artistId, body){
  let eduSheet = ss.getSheetByName('ArtistEducation');
  if(!eduSheet){
    eduSheet = ss.insertSheet('ArtistEducation');
    eduSheet.appendRow(['artistId','year','institution','specialty']);
  }
  eduSheet.appendRow([artistId, body.educationYear || '', body.institution || '', body.educationSpecialty || '']);
}

function rowToObj(headers, row){
  const o = {};
  headers.forEach((h,i)=> o[h] = row[i]);
  return o;
}

function normalizeHeader(h){
  return String(h||'').trim().toLowerCase().replace(/\s+/g,'').replace(/[-_]/g,'');
}

function getSheetDataAny(names, type){
  for(const n of names){
    const s = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(n);
    if(s) return getSheetData(n, type);
  }
  return [];
}

function getSheetData(sheetName, type){
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if(!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if(values.length < 2) return [];
  const headers = values[0];
  const rows = values.slice(1);

  const maps = {
    artist: {
      'id':'id','name':'name','имя':'name','specialty':'genre','специальность':'genre','жанр':'genre',
      'spec':'spec','тип':'spec','специализация':'spec','колвочеловек':'peopleCount',
      'price':'price','гонорар':'price','цена':'price','img':'img','фото':'img','фотоurl':'img',
      'city':'city','город':'city','rating':'rating','рейтинг':'rating',
      'verified':'verified','верифицирован':'verified','freedates':'freeDates','свободныедаты':'freeDates',
      'contacts':'contacts','контакты':'contacts','readytotravel':'readyToTravel','готовквыезду':'readyToTravel',
      'countries':'countries','страны':'countries','telegramid':'telegramId','role':'role','био':'bio',
      'travelscope':'travelScope','visatype':'visaType','visavaliduntil':'visaValidUntil',
      'needsvisasupport':'needsVisaSupport',
      'vacationcity':'vacationCity','vacationuntil':'vacationUntil','ispro':'isPro',
      'height':'height','weight':'weight','clothingsize':'clothingSize','shoesize':'shoeSize',
      'haircolor':'hairColor','hairlength':'hairLength','eyecolor':'eyeColor','specialmarks':'specialMarks'
    },
    venue: {
      'id':'id','название':'name','name':'name','город':'city','city':'city','тип':'type','type':'type',
      'вместимость':'capacity','capacity':'capacity','аренда':'price','ставка':'price','price':'price',
      'baseprice':'basePrice','seasonmultiplier':'seasonMultiplier','weekendmultiplier':'weekendMultiplier',
      'personalprices':'personalPrices','фото':'img','img':'img','контакты':'contacts','contacts':'contacts',
      'ownertelegramid':'ownerTelegramId','nobargain':'noBargain','безторгов':'noBargain'
    },
    event: {
      'id':'id','дата':'date','date':'date','время':'time','time':'time','тип':'type','типсобытия':'type',
      'артистid':'artistId','artistid':'artistId','площадкаid':'venueId','venueid':'venueId',
      'организаторid':'organizerId','organizerid':'organizerId','статус':'status','status':'status',
      'место':'venue','venue':'venue','карта':'mapUrl','title':'title','название':'title'
    },
    person: {
      'telegramid':'telegramId','name':'name','имя':'name','photourl':'photoUrl',
      'фото':'photoUrl','activerole':'activeRole','roles':'roles','city':'city','город':'city'
    },
    pricing: {
      'artistid':'artistId','date':'date','дата':'date',
      'pricemodifier':'priceModifier','коэффициент':'priceModifier','note':'note','комментарий':'note'
    },
    review: {
      'artistid':'artistId','author':'author','автор':'author','rating':'rating','рейтинг':'rating',
      'text':'text','текст':'text','date':'date','дата':'date'
    },
    application: {
      'id':'id','createdat':'createdAt','organizerid':'organizerId','organizername':'organizerName',
      'artistid':'artistId','artistname':'artistName','venueid':'venueId',
      'eventtype':'eventType','date':'date','дата':'date','city':'city','город':'city',
      'guests':'guests','гостей':'guests','equipment':'equipment','оборудование':'equipment',
      'budget':'budget','бюджет':'budget','status':'status','статус':'status',
      'venueresponse':'venueResponse','venueprice':'venuePrice','venuenote':'venueNote','type':'type'
    },
    project: {
      'id':'id','createdat':'createdAt','organizerid':'organizerId','organizername':'organizerName',
      'title':'title','название':'title','type':'type','тип':'type',
      'startdate':'startDate','enddate':'endDate','cities':'cities','города':'cities',
      'budget':'budget','бюджет':'budget','description':'description','описание':'description',
      'isopen':'isOpen','artistids':'artistIds','status':'status','статус':'status'
    },
    projectResponse: {
      'id':'id','createdat':'createdAt','projectid':'projectId','artistid':'artistId',
      'artistname':'artistName','price':'price','цена':'price',
      'comment':'comment','комментарий':'comment','status':'status','статус':'status'
    },
    venuePricing: {
      'id':'id','venueid':'venueId','datefrom':'dateFrom','dateto':'dateTo',
      'price':'price','цена':'price','note':'note','комментарий':'note'
    },
    venueCategoryPricing: {
      'id':'id','venueid':'venueId','eventtype':'eventType','multiplier':'multiplier',
      'коэффициент':'multiplier','note':'note','комментарий':'note'
    },
    artistSchedule: {
      'artistid':'artistId','dayofweek':'dayOfWeek','isworking':'isWorking',
      'starthour':'startHour','endhour':'endHour','busyhours':'busyHours'
    },
    media: { 'artistid':'artistId','type':'type','тип':'type','url':'url','ссылка':'url','caption':'caption' },
    experience: { 'artistid':'artistId','year':'year','год':'year','project':'project','проект':'project','role':'role','роль':'role','venue':'venue','площадка':'venue','city':'city','город':'city' },
    education: { 'artistid':'artistId','year':'year','год':'year','institution':'institution','учреждение':'institution','specialty':'specialty','специальность':'specialty' },
    inventory: { 'artistid':'artistId','name':'name','название':'name','description':'description','описание':'description','photourl':'photoUrl','фото':'photoUrl' },
    costume: { 'artistid':'artistId','name':'name','название':'name','size':'size','размер':'size','description':'description','описание':'description','photourl':'photoUrl','фото':'photoUrl' }
  };
  const map = maps[type] || {};

  return rows.map((row,i)=>{
    const obj = {};
    headers.forEach((h,j)=>{
      const k = map[normalizeHeader(h)];
      if(!k) return;
      let v = row[j];

      if(['id','price','rating','capacity','artistId','venueId','peopleCount','year','basePrice','guests','budget','projectId','responseId','multiplier','height','weight'].includes(k)){
        if(v instanceof Date) v = 0;
        else v = Number(v)||0;
      }
      if(k === 'dayOfWeek'){
        if(v instanceof Date) v = Number(v.getDay());
        else v = Math.floor(Number(v))||0;
        if(v < 0 || v > 6) v = 0;
      }
      if(k === 'startHour' || k === 'endHour'){
        const fallback = (k === 'startHour') ? 9 : 21;
        if(v instanceof Date){
          const h = v.getHours();
          v = (h >= 0 && h <= 23) ? h : fallback;
        } else {
          v = Math.floor(Number(v));
          if(!isFinite(v) || v < 0 || v > 24) v = fallback;
        }
      }
      if(['seasonMultiplier','weekendMultiplier','priceModifier'].includes(k)){
        if(v instanceof Date) v = 1;
        else v = Number(v);
        if(!isFinite(v) || v <= 0 || v > 10) v = 1;
      }
      if(['verified','readyToTravel','needsVisaSupport','noBargain','isOpen','isWorking','isPro'].includes(k)){
        const s = String(v).toLowerCase();
        v = (v===true || s==='true' || s==='да' || s==='1');
      }
      if(k==='freeDates'&&typeof v==='string') v = v.split(',').map(d=>d.trim()).filter(Boolean);
      if(k==='roles'&&typeof v==='string') v = v.split(',').map(r=>r.trim()).filter(Boolean);
      if((k==='date'||k==='createdAt'||k==='startDate'||k==='endDate'||k==='dateFrom'||k==='dateTo'||k==='vacationUntil')&&v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(),'yyyy-MM-dd');
      if(k==='telegramId'||k==='organizerId'||k==='artistId'||k==='ownerTelegramId'||k==='venueId'||k==='projectId') v = String(v||'');
      obj[k] = v;
    });
    if(!obj.id && type==='artist') obj.id = i+1;
    return obj;
  }).filter(o=>{
    if(type==='artist') return o.name && String(o.name).trim() !== '';
    if(type==='person') return o.telegramId;
    return true;
  });
}

function updateRow(ss, sheetName, id, fields){
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  if(idCol < 0) return;
  for(let i=1; i<data.length; i++){
    if(String(data[i][idCol])===String(id)){
      Object.keys(fields).forEach(key=>{
        const col = headers.indexOf(key);
        if(col > -1) sheet.getRange(i+1, col+1).setValue(fields[key]);
      });
      return;
    }
  }
}

function jsonResponse(data){
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// Чистим префикс data:image/...;base64, если он пришёл с фронта
function extractBase64(base64){
  const s = String(base64 || '');
  const comma = s.indexOf(',');
  return (s.indexOf('base64,') > -1 && comma > -1) ? s.slice(comma + 1) : s;
}

// Сохраняет фото в папку Drive и возвращает публичную ссылку для Artists.img
function uploadPhoto(base64, filename){
  const raw = extractBase64(base64);
  const safeName = String(filename || 'photo.jpg').replace(/[^\w.\-]/g, '_');
  const bytes = Utilities.base64Decode(raw);
  const blob = Utilities.newBlob(bytes, 'image/jpeg', safeName);
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(err){
    console.error('setSharing failed: ' + err);
    throw new Error('Sharing disabled on account');
  }
  console.log('uploadPhoto: ' + safeName + ' -> ' + file.getId());
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
}