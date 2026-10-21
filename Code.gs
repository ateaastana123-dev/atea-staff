/** ATEA TEAM — Google Apps Script backend. Deploy as Web App, execute as owner.
 * Required Script Properties (Project Settings):
 * ADMIN_PIN, TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID,
 * SHEET_ID, DRIVE_FOLDER_ID, ALLOWED_ORIGIN
 * Never publish secrets in GitHub.
 */
const HEADERS = ['ID','Дата и время (Астана)','Филиал','Сотрудник','Этап','Выполнено','Всего','Комментарий','Папка Drive','Фото JSON','Telegram','Серверное время ISO','Должность'];
const BRANCHES = ['MEGA Silk Way','Достык 13','Мангилик Ел 36','Туран 22/1'];
const MAX_PHOTOS=20, MAX_BYTES=500000, MAX_REQUEST_CHARS=6500000;

function doGet(e){
 return HtmlService.createHtmlOutput('<h2>ATEA TEAM API</h2><p>Сервис доступен. Отправляйте запросы через приложение ATEA TEAM.</p>')
   .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function doPost(e){
 let nonce='',result;
 try {
  if(!e||!e.parameter||!e.parameter.payload)throw Error('Пустой запрос');
  if(e.parameter.payload.length>MAX_REQUEST_CHARS)throw Error('Запрос слишком большой. Уменьшите фотографии.');
  const req=JSON.parse(e.parameter.payload);nonce=String(req.nonce||'').slice(0,100);
  const action=String(req.action||'');
  if(action==='auth'){assertPin_(req.pin,'ADMIN_PIN');result={ok:true};}
  else if(action==='submit'){result=submit_(req);}
  else if(action==='list'){assertPin_(req.pin,'ADMIN_PIN');result=list_(req);}
  else if(action==='financeSubmit'){result=financeSubmit_(req);}
  else if(action==='financeList'){assertPin_(req.pin,'ADMIN_PIN');result=financeList_(req);}
  else throw Error('Неизвестная операция');
 }catch(err){result={ok:false,error:String(err.message||err).slice(0,300)};}
 // Fixed trusted GitHub Pages origin. Missing/misconfigured ALLOWED_ORIGIN previously
 // fell back to example.invalid and broke all iframe replies.
 const target='https://ateaastana123-dev.github.io';
 const message=JSON.stringify({channel:'atea-team-v1',nonce,result}).replace(/</g,'\\u003c');
 const html='<!doctype html><html><body><script>parent.postMessage('+message+','+JSON.stringify(target)+');<\/script></body></html>';
 return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function prop_(key){return PropertiesService.getScriptProperties().getProperty(key)||'';}
function assertPin_(candidate,key){
 const actual=prop_(key);if(!actual||actual.length<8)throw Error('Секрет '+key+' не настроен (минимум 8 символов).');
 if(typeof candidate!=='string'||candidate.length>100||candidate!==actual)throw Error('Неверный код доступа.');
}
function submit_(r){
 if(!BRANCHES.includes(r.branch))throw Error('Неизвестный филиал');
 const employee=String(r.employee||'').trim().slice(0,80),note=String(r.note||'').slice(0,1000);
 if(!employee)throw Error('Не указано имя');
 if(!['opening','closing'].includes(r.shift))throw Error('Неверный этап смены');
 if(!['cashier','barista'].includes(r.role))throw Error('Выберите должность кассир или бариста');
 if(!Array.isArray(r.items)||!Array.isArray(r.checks)||!Array.isArray(r.photos))throw Error('Неполный отчёт');
 if(r.items.length<1||r.items.length>MAX_PHOTOS||r.items.length!==r.checks.length||r.items.length!==r.photos.length)throw Error('Для каждого пункта требуется фото');
 if(r.checks.some(x=>x!==true))throw Error('Не все пункты отмечены');
 for(let i=0;i<r.photos.length;i++){
  const p=r.photos[i];if(p.item!==i+1||p.mime!=='image/jpeg'||typeof p.base64!=='string'||p.base64.length>MAX_BYTES*1.5)throw Error('Неверный формат фотографии '+(i+1));
  if(typeof r.items[i]!=='string'||r.items[i].length>300)throw Error('Некорректный пункт');
 }
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const sheet=sheet_(),folder=DriveApp.getFolderById(prop_('DRIVE_FOLDER_ID'));
  const date=new Date(),stamp=Utilities.formatDate(date,'Asia/Almaty','yyyy-MM-dd HH:mm:ss');
  const id=Utilities.formatDate(date,'Asia/Almaty','yyyyMMdd-HHmmss')+'-'+Utilities.getUuid().slice(0,8);
  const reportFolder=folder.createFolder(id+'_'+safeFile_(r.branch)+'_'+r.shift);
  const urls=[];
  try{
   for(let i=0;i<r.photos.length;i++){
    const p=r.photos[i],bytes=Utilities.base64Decode(p.base64);
    if(bytes.length>MAX_BYTES)throw Error('Фото '+(i+1)+' слишком большое.');
    // Confirm JPEG magic bytes.
    if(bytes[0]!==-1&&bytes[0]!==255)throw Error('Фото не является JPEG');
    const blob=Utilities.newBlob(bytes,'image/jpeg',String(i+1).padStart(2,'0')+'_'+safeFile_(r.items[i]).slice(0,55)+'.jpg');
    const file=reportFolder.createFile(blob);urls.push(file.getUrl());
   }
   const row=[id,stamp,safeCell_(r.branch),safeCell_(employee),r.shift,r.checks.filter(Boolean).length,r.items.length,safeCell_(note),reportFolder.getUrl(),JSON.stringify(urls),'',date.toISOString(),r.role];
   sheet.appendRow(row);
  }catch(err){reportFolder.setTrashed(true);throw err;}
  let telegramSent=false;
  try{telegramSent=telegram_(id,stamp,r,urls,reportFolder.getUrl());}catch(err){console.error('Telegram failed',err);}
  sheet.getRange(sheet.getLastRow(),11).setValue(telegramSent?'Отправлено':'Ошибка / не настроено');
  return {ok:true,id,telegramSent};
 }finally{lock.releaseLock()}
}
function telegram_(id,stamp,r,urls,folderUrl){
 const token=prop_('TELEGRAM_BOT_TOKEN'),chat=prop_('TELEGRAM_CHAT_ID');if(!token||!chat)return false;
 const stage=r.shift==='opening'?'Открытие':'Закрытие';
 const msg=['ATEA TEAM — новый чек-лист','№ '+id,stamp+' (Астана)','Филиал: '+r.branch,'Сотрудник: '+r.employee,'Должность: '+(r.role==='barista'?'Бариста':'Кассир'),'Этап: '+stage,'Выполнено: '+r.items.length+'/'+r.items.length,'Фото: '+urls.length,'Папка Drive: '+folderUrl,'Комментарий: '+(r.note||'нет')].join('\n').slice(0,3900);
 const response=UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/sendMessage',{
  method:'post',contentType:'application/json',payload:JSON.stringify({chat_id:chat,text:msg,disable_web_page_preview:true}),muteHttpExceptions:true
 });
 const body=JSON.parse(response.getContentText());if(response.getResponseCode()!==200||!body.ok)throw Error('Telegram API: '+response.getContentText().slice(0,200));return true;
}
function list_(r){
 const branch=String(r.branch||'all');if(branch!=='all'&&!BRANCHES.includes(branch))throw Error('Неверный фильтр');
 const sheet=sheet_(),last=sheet.getLastRow();if(last<2)return {ok:true,rows:[]};
 const first=Math.max(2,last-299),values=sheet.getRange(first,1,last-first+1,13).getDisplayValues();
 const rows=values.reverse().filter(x=>branch==='all'||x[2]===branch).slice(0,50).map(x=>{
  let photos=[];try{photos=JSON.parse(x[9]||'[]')}catch(err){}
  return {id:x[0],date:x[1],branch:x[2],employee:x[3],shift:x[4],role:x[12]||'',completed:x[5],total:x[6],note:x[7],photos};
 });
 return {ok:true,rows};
}
function sheet_(){
 const id=prop_('SHEET_ID');if(!id)throw Error('SHEET_ID не настроен');
 const book=SpreadsheetApp.openById(id);let sheet=book.getSheetByName('ATEA_TEAM_Reports');
 if(!sheet){sheet=book.insertSheet('ATEA_TEAM_Reports');sheet.appendRow(HEADERS);sheet.setFrozenRows(1)}
 return sheet;
}
function safeFile_(s){return String(s||'').replace(/[^\p{L}\p{N} _-]/gu,'_').slice(0,75)}
function safeCell_(s){s=String(s||'');return /^[=+\-@]/.test(s)?"'"+s:s}


// Financial reports: separate tab, server-calculated metrics, Telegram notification.
const FINANCE_HEADERS=['ID','Дата отправки (Астана)','Дата отчёта','Филиал','Отправитель','План','Факт','Отклонение','Отклонение %','Выполнение %','Наличные','Безнал','Бонусы','Доставки','Чеки','Средний чек','Списания сумма','Стоп-лист','Возвраты','Жалобы','Опоздания','Сотрудники смены','Списания подробно','Технические проблемы','Проблема дня','Telegram'];
function financeSheet_(){const id=prop_('SHEET_ID');if(!id)throw Error('SHEET_ID не настроен');const book=SpreadsheetApp.openById(id);let sh=book.getSheetByName('ATEA_TEAM_Finance');if(!sh){sh=book.insertSheet('ATEA_TEAM_Finance');sh.appendRow(FINANCE_HEADERS);sh.setFrozenRows(1)}return sh}
function financeSubmit_(req){
 if(!BRANCHES.includes(req.branch))throw Error('Неверный филиал');const employee=String(req.employee||'').trim().slice(0,80);if(!employee)throw Error('Укажите имя отправителя');const r=req.report;if(!r||typeof r!=='object')throw Error('Нет данных отчёта');
 const date=String(r.rDate||'');if(!/^20\d\d-\d\d-\d\d$/.test(date)||isNaN(new Date(date+'T12:00:00Z').getTime()))throw Error('Неверная дата отчёта');
 const num=k=>{const v=r[k];if(v===undefined||v===null||v==='')return 0;const x=Number(v);if(!Number.isFinite(x)||x<0||x>1e10)throw Error('Некорректное число: '+k);return x};
 const plan=num('rPlan'),fact=num('rFact'),checks=num('rChecks');if(!checks||!String(r.rPlan||'')||!String(r.rFact||''))throw Error('Укажите план, факт и количество чеков');
 const diff=fact-plan,pct=plan>0?diff/plan*100:0,done=plan>0?fact/plan*100:0,avg=fact/checks;
 const str=k=>safeCell_(String(r[k]||'').trim().slice(0,1000));if(!str('rEmployees'))throw Error('Укажите сотрудников смены');
 const now=new Date(),stamp=Utilities.formatDate(now,'Asia/Almaty','yyyy-MM-dd HH:mm:ss'),id=Utilities.getUuid().slice(0,12);
 const row=[id,stamp,date,safeCell_(req.branch),safeCell_(employee),plan,fact,diff,pct,done,num('rCash'),num('rCard'),num('rBonus'),num('rDelivery'),checks,avg,num('rWrite'),str('rStop'),str('rReturns'),str('rComplaints'),str('rLate'),str('rEmployees'),str('rWriteDetails'),str('rTechnical'),str('rProblem'),''];
 const lock=LockService.getScriptLock();lock.waitLock(30000);let sh,rowNum;try{sh=financeSheet_();sh.appendRow(row);rowNum=sh.getLastRow()}finally{lock.releaseLock()}
 let telegramSent=false;try{telegramSent=financeTelegram_(id,date,req.branch,employee,row)}catch(e){console.error('Finance Telegram failed',e)}
 sh.getRange(rowNum,26).setValue(telegramSent?'Отправлено':'Ошибка / не настроено');return {ok:true,id,telegramSent};
}
function financeTelegram_(id,date,branch,employee,row){const token=prop_('TELEGRAM_BOT_TOKEN'),chat=prop_('TELEGRAM_CHAT_ID');if(!token||!chat)return false;
 const money=n=>Math.round(n).toLocaleString('ru-RU')+' ₸';
 const msg=['ATEA TEAM — ЕЖЕДНЕВНЫЙ ОТЧЁТ','Дата: '+date,'Филиал: '+branch,'Отправитель: '+employee,'Сотрудники смены: '+row[21],'Выручка: '+money(row[6]),'План: '+money(row[5]),'Отклонение: '+money(row[7])+' ('+row[8].toFixed(2)+'%)','Выполнение плана: '+row[9].toFixed(2)+'%','Наличные: '+money(row[10]),'Безнал: '+money(row[11]),'Бонусы: '+money(row[12]),'Доставки: '+money(row[13]),'Чеков: '+row[14],'Средний чек: '+money(row[15]),'Списания: '+money(row[16]),'Стоп-лист: '+(row[17]||'нет'),'Возвраты: '+(row[18]||'нет'),'Жалобы: '+(row[19]||'нет'),'Опоздания: '+(row[20]||'нет'),'Списания подробно: '+(row[22]||'нет'),'Технические проблемы: '+(row[23]||'нет'),'Проблема дня: '+(row[24]||'нет'),'ID: '+id].join('\n').slice(0,4000);
 const response=UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/sendMessage',{method:'post',contentType:'application/json',payload:JSON.stringify({chat_id:chat,text:msg,disable_web_page_preview:true}),muteHttpExceptions:true});const body=JSON.parse(response.getContentText());if(response.getResponseCode()!==200||!body.ok)throw Error('Telegram API: '+response.getContentText().slice(0,200));return true}
function financeList_(req){const branch=String(req.branch||'all');if(branch!=='all'&&!BRANCHES.includes(branch))throw Error('Неверный филиал');const sh=financeSheet_(),last=sh.getLastRow();if(last<2)return {ok:true,rows:[]};const first=Math.max(2,last-299),rows=sh.getRange(first,1,last-first+1,26).getDisplayValues().reverse().filter(x=>branch==='all'||x[3]===branch).slice(0,50).map(x=>({id:x[0],date:x[2],branch:x[3],employee:x[4],plan:x[5],fact:x[6],deviation:x[7],deviationPercent:x[8],completion:x[9],cash:x[10],card:x[11],bonus:x[12],delivery:x[13],checks:x[14],average:x[15],writeoffs:x[16],stop:x[17],returns:x[18],complaints:x[19],late:x[20],employees:x[21],writeoffDetails:x[22],technical:x[23],problem:x[24],telegram:x[25]}));return {ok:true,rows}}
