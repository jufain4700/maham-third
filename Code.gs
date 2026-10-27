/** اقرأ البيانات فقط؛ جميع عمليات الإضافة والتعديل والحذف داخل Google Sheet بصلاحيات Google. */
const SHEET_TASKS = 'المهام';
const SHEET_SUBJECTS = 'المواد';
const SHEET_TYPES = 'أنواع المهام';

function doGet(e) {
  const callback = String((e && e.parameter && e.parameter.callback) || '');
  // callback مقيد بالأحرف المسموحة لمنع إدخال كود ضار
  if (!/^classTasksCallback_[0-9]{10,20}$/.test(callback)) {
    return ContentService.createTextOutput('Invalid callback').setMimeType(ContentService.MimeType.TEXT);
  }
  let data;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tasks = getRows_(ss.getSheetByName(SHEET_TASKS));
    const subjects = getList_(ss.getSheetByName(SHEET_SUBJECTS));
    const types = getList_(ss.getSheetByName(SHEET_TYPES));
    data = {ok:true,subjects:subjects,types:types,tasks:tasks.filter(x => x.title && x.active !== 'لا').map(x=>({id:x.id,title:x.title,subject:x.subject,type:x.type,notes:x.notes,link:x.link,start:x.start,end:x.end,created:x.created}))};
  } catch(err) { data = {ok:false,error:'تعذر قراءة الجدول'}; }
  return ContentService.createTextOutput(callback+'('+JSON.stringify(data)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
}
function getRows_(sheet){
  if (!sheet) throw new Error('Missing tasks tab');
  const rows = sheet.getDataRange().getDisplayValues();
  return rows.slice(1).filter(r=>r.some(x=>x)).map(r=>({
    id:r[0],title:r[1],subject:r[2],type:r[3],notes:r[4],link:r[5],
    start:normalizeDate_(sheet,r[6]),end:normalizeDate_(sheet,r[7]),created:normalizeDate_(sheet,r[8]),active:r[9]
  }));
}
function normalizeDate_(sheet,s){return String(s||'').trim();}
function getList_(sheet){if(!sheet)throw new Error('Missing list tab');return sheet.getRange(2,1,Math.max(1,sheet.getLastRow()-1),1).getDisplayValues().flat().map(x=>x.trim()).filter(Boolean);}
