/**
 * 우리 반 게시물 아카이브 — Google Apps Script backend
 * Deploy as the school Workspace owner, Execute as: Me, Access: school domain.
 * Session.getActiveUser().getEmail() MUST be nonempty; otherwise FAIL CLOSED.
 * Never publish the Spreadsheet or Drive folder or expose their file URLs.
 */
const USER_COLUMNS = ['email', 'role', 'active', 'addedAt'];
const POST_COLUMNS = ['id','title','category','displayDate','description','visibility','uploadedAt','uploadedBy','originalId','enhancedId','thumbId','status','updatedAt'];
const REPORT_COLUMNS = ['postId','reason','reportedBy','reportedAt'];
const CATEGORIES = ['행사','대회','공지','동아리','학습','기타'];
const MAX_IMAGE_BYTES = 3500000; // each image after browser preprocessing
const DAILY_UPLOAD_LIMIT = 8;

function setup() {
  const adminEmail = (Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if (!adminEmail || !adminEmail.includes('@') || adminEmail.endsWith('@gmail.com')) {
    throw new Error('학교 Google Workspace 계정으로 로그인하여 setup을 실행하세요.');
  }
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('READY') === 'yes') throw new Error('이미 설정되었습니다. 중복 설정을 하지 마세요.');
  const root = DriveApp.createFolder('우리 반 게시물 아카이브 - 비공개 원본');
  const orig = root.createFolder('originals_private');
  const enhanced = root.createFolder('enhanced_private');
  const thumbs = root.createFolder('thumbnails_private');
  const ss = SpreadsheetApp.create('우리 반 게시물 아카이브 - 비공개 목록');
  const s0 = ss.getSheets()[0];
  s0.setName('Users'); s0.appendRow(USER_COLUMNS); s0.appendRow([adminEmail, 'admin', true, new Date().toISOString()]);
  const posts = ss.insertSheet('Posts'); posts.appendRow(POST_COLUMNS);
  const reports = ss.insertSheet('Reports'); reports.appendRow(REPORT_COLUMNS);
  p.setProperties({
    READY: 'yes', DOMAIN: adminEmail.split('@')[1], ADMIN_EMAIL: adminEmail,
    SHEET_ID: ss.getId(), ROOT_FOLDER_ID: root.getId(), ORIGINAL_FOLDER_ID: orig.getId(),
    ENHANCED_FOLDER_ID: enhanced.getId(), THUMB_FOLDER_ID: thumbs.getId()
  });
  console.log('초기화 완료. 허용 관리자: ' + adminEmail);
  console.log('Google Sheet: ' + ss.getUrl());
  console.log('Private Drive Folder: ' + root.getUrl());
}

function doGet() {
  try { requireMember_(); }
  catch (e) {
    return HtmlService.createHtmlOutput(
      '<!doctype html><html lang="ko"><meta name="viewport" content="width=device-width,initial-scale=1"><meta charset="utf-8">' +
      '<body style="font-family:system-ui;max-width:600px;margin:16vh auto;padding:24px;line-height:1.8"><h2>🔒 우리 반 전용 아카이브</h2><p>학교에서 허용된 Google 계정으로 로그인해야 합니다.</p><p>접속할 수 없다면 학교 계정으로 바꿔 로그인하거나 관리자에게 허용 계정 등록을 요청하세요.</p><p>이 서비스는 학생의 Google 비밀번호를 요구하지 않습니다.</p></body></html>'
    ).setTitle('접근 제한 · 학교 게시물 아카이브');
  }
  const tpl = HtmlService.createTemplateFromFile('Index');
  return tpl.evaluate().setTitle('학교 게시물 아카이브').addMetaTag('viewport','width=device-width, initial-scale=1');
}
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

function env_() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('READY') !== 'yes') throw new Error('관리자가 setup()을 먼저 실행해야 합니다.');
  return p;
}
function tab_(name) { return SpreadsheetApp.openById(env_().getProperty('SHEET_ID')).getSheetByName(name); }
function rows_(name) {
  const s = tab_(name); const n = s.getLastRow();
  if (n < 2) return [];
  const values = s.getRange(2,1,n-1,s.getLastColumn()).getValues();
  const heads = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
  return values.map((r, i) => ({ rowNumber:i + 2, ...Object.fromEntries(heads.map((h,j) => [h,r[j]])) }));
}
function requireMember_() {
  const p = env_();
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email || email.split('@')[1] !== p.getProperty('DOMAIN')) throw new Error('학교 계정을 확인할 수 없거나 허용된 학교 계정이 아닙니다.');
  const user = rows_('Users').find(r => String(r.email).trim().toLowerCase() === email && (r.active === true || r.active === 'TRUE'));
  if (!user) throw new Error('허용된 우리 반 학생 계정이 아닙니다.');
  return {email,role:user.role === 'admin' ? 'admin' : 'student'};
}
function requireAdmin_() {
  const u = requireMember_();
  if (u.role !== 'admin' || u.email !== env_().getProperty('ADMIN_EMAIL')) throw new Error('관리자만 사용할 수 있습니다.');
  return u;
}
function cleanText_(v, maxLen) {
  const s = String(v == null ? '' : v).trim().slice(0,maxLen);
  return /^[\s]*[=+@-]/.test(s) ? "'" + s : s;
}
function dateOk_(d) { return /^20\d{2}-\d{2}-\d{2}$/.test(d) && !Number.isNaN(new Date(d+'T12:00:00Z').valueOf()) && new Date(d+'T12:00:00Z').toISOString().slice(0,10) === d; }
function sheetDate_(value, dateOnly) {
  if (value instanceof Date) return Utilities.formatDate(value, 'Asia/Seoul', dateOnly ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm:ssXXX");
  return String(value == null ? '' : value);
}
function basicPost_(r) {
  return {id:r.id,title:r.title,category:r.category,displayDate:sheetDate_(r.displayDate,true),description:r.description,
    visibility:r.visibility,uploadedAt:sheetDate_(r.uploadedAt,false),updatedAt:sheetDate_(r.updatedAt,false)};
}
function getBootstrap() {
  const user=requireMember_();
  return {user,categories:CATEGORIES,limitPerDay:DAILY_UPLOAD_LIMIT};
}
function listPosts() {
  const user=requireMember_();
  return rows_('Posts').filter(r => r.status === 'active' && (r.visibility === 'class' || user.role === 'admin'))
    .map(basicPost_).sort((a,b) => b.displayDate.localeCompare(a.displayDate) || b.uploadedAt.localeCompare(a.uploadedAt));
}
function findPost_(id) {
  return rows_('Posts').find(r => r.id === String(id) && r.status === 'active');
}
function requirePostAccess_(id,kind) {
  const user = requireMember_();
  const post = findPost_(id);
  if (!post) throw new Error('게시물을 찾을 수 없습니다.');
  if (post.visibility !== 'class' && user.role !== 'admin') throw new Error('접근 권한이 없습니다.');
  if (kind === 'original' && user.role !== 'admin') throw new Error('원본 이미지는 관리자만 열람할 수 있습니다.');
  return post;
}
function getImage(postId,kind) {
  const imageKind = kind === 'original' ? 'original' : (kind === 'thumbnail' ? 'thumbnail' : 'enhanced');
  const post = requirePostAccess_(postId,imageKind);
  const fileId = imageKind === 'original' ? post.originalId : (imageKind === 'thumbnail' ? post.thumbId : post.enhancedId);
  const blob = DriveApp.getFileById(fileId).getBlob();
  if (blob.getSize() > MAX_IMAGE_BYTES) throw new Error('파일 크기 제한을 초과했습니다.');
  return {mime: blob.getContentType(), base64: Utilities.base64Encode(blob.getBytes())};
}
function imageBlob_(data,name) {
  if (!data || typeof data !== 'object') throw new Error('이미지 형식이 올바르지 않습니다.');
  const mt = data.mime;
  if (!['image/jpeg','image/png','image/webp'].includes(mt)) throw new Error('JPG, PNG 또는 WebP만 허용합니다.');
  const base = String(data.base64 || '');
  if (!base || base.length > 4900000 || !/^[A-Za-z0-9+/]+={0,2}$/.test(base)) throw new Error('이미지 데이터가 너무 크거나 손상되었습니다.');
  const bytes = Utilities.base64Decode(base);
  if (bytes.length < 100 || bytes.length > MAX_IMAGE_BYTES) throw new Error('이미지 크기를 확인해주세요.');
  // Verify binary signatures, not merely claimed MIME.
  const b = bytes.map ? bytes : Array.from(bytes);
  const unsigned = n => (b[n] + 256) % 256;
  const jpeg = unsigned(0) === 0xff && unsigned(1) === 0xd8 && unsigned(2) === 0xff;
  const png = unsigned(0) === 0x89 && unsigned(1) === 0x50 && unsigned(2) === 0x4e && unsigned(3) === 0x47;
  const webp = String.fromCharCode(...b.slice(0,4)) === 'RIFF' && String.fromCharCode(...b.slice(8,12)) === 'WEBP';
  if (!((mt === 'image/jpeg' && jpeg) || (mt === 'image/png' && png) || (mt === 'image/webp' && webp))) throw new Error('파일 형식과 실제 이미지 형식이 일치하지 않습니다.');
  return Utilities.newBlob(bytes,mt,name);
}
function savePost(payload) {
  const who = requireMember_();
  const title = cleanText_(payload && payload.title,90);
  const description = cleanText_(payload && payload.description,500);
  const category = String(payload && payload.category || '');
  const date = String(payload && payload.displayDate || '');
  const visibility = payload && payload.visibility === 'admin' ? 'admin' : 'class';
  if (!title || !CATEGORIES.includes(category) || !dateOk_(date)) throw new Error('제목, 분류, 날짜를 확인해주세요.');
  if (!payload || payload.privacyConfirmed !== true) throw new Error('개인정보 확인에 동의해야 등록할 수 있습니다.');
  const original = imageBlob_(payload.original, 'original.jpg');
  const enhanced = imageBlob_(payload.enhanced, 'enhanced.jpg');
  const thumb = imageBlob_(payload.thumbnail, 'thumb.jpg');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('동시 업로드가 많습니다. 다시 시도해주세요.');
  let oFile=null, eFile=null, tFile=null;
  try {
    const today = new Date().toISOString().slice(0,10);
    const todayCount = rows_('Posts').filter(r => r.uploadedBy===who.email && sheetDate_(r.uploadedAt,false).startsWith(today)).length;
    if (todayCount >= DAILY_UPLOAD_LIMIT) throw new Error('하루 등록 제한('+DAILY_UPLOAD_LIMIT+'개)을 초과했습니다.');
    const id=Utilities.getUuid(); const p=env_();
    oFile = DriveApp.getFolderById(p.getProperty('ORIGINAL_FOLDER_ID')).createFile(original.setName(id+'-before.jpg'));
    eFile = DriveApp.getFolderById(p.getProperty('ENHANCED_FOLDER_ID')).createFile(enhanced.setName(id+'-after.jpg'));
    tFile = DriveApp.getFolderById(p.getProperty('THUMB_FOLDER_ID')).createFile(thumb.setName(id+'-thumb.jpg'));
    const now = new Date().toISOString();
    tab_('Posts').appendRow([id,title,category,date,description,visibility,now,who.email,oFile.getId(),eFile.getId(),tFile.getId(),'active',now]);
    return {ok:true,id};
  } catch(e) {
    if(oFile) try{oFile.setTrashed(true);}catch(_){}
    if(eFile) try{eFile.setTrashed(true);}catch(_){}
    if(tFile) try{tFile.setTrashed(true);}catch(_){}
    throw e;
  } finally { lock.releaseLock(); }
}
function updatePost(payload) {
  requireAdmin_();
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const post=findPost_(payload && payload.id);
    if(!post) throw new Error('게시물을 찾을 수 없습니다.');
    if (sheetDate_(post.updatedAt,false) !== payload.version) throw new Error('다른 곳에서 수정됐습니다. 새로고침 후 다시 시도해주세요.');
    const title=cleanText_(payload.title,90); const category=String(payload.category);
    const date=String(payload.displayDate); const desc=cleanText_(payload.description,500);
    if(!title || !CATEGORIES.includes(category) || !dateOk_(date)) throw new Error('입력값을 확인해주세요.');
    const values=[post.id,title,category,date,desc,payload.visibility==='admin'?'admin':'class',post.uploadedAt,post.uploadedBy,post.originalId,post.enhancedId,post.thumbId,'active',new Date().toISOString()];
    tab_('Posts').getRange(post.rowNumber,1,1,POST_COLUMNS.length).setValues([values]);
    return {ok:true};
  } finally {lock.releaseLock();}
}
function removePost(id) {
  requireAdmin_();
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const p=findPost_(id);
    if(!p) throw new Error('삭제할 게시물이 없습니다.');
    const s=tab_('Posts');
    s.getRange(p.rowNumber,POST_COLUMNS.indexOf('status')+1).setValue('deleted');
    s.getRange(p.rowNumber,POST_COLUMNS.indexOf('updatedAt')+1).setValue(new Date().toISOString());
    // Removed from all views immediately; Drive files moved to trash for retention control.
    DriveApp.getFileById(p.originalId).setTrashed(true);
    DriveApp.getFileById(p.enhancedId).setTrashed(true);
    DriveApp.getFileById(p.thumbId).setTrashed(true);
    return {ok:true};
  } finally {lock.releaseLock();}
}
function listUsers() {
  requireAdmin_();
  return rows_('Users').map(r => ({email:r.email,role:r.role,active:r.active === true || r.active === 'TRUE'}));
}
function setUser(email,active) {
  const admin=requireAdmin_();
  const e=String(email||'').trim().toLowerCase();
  const dom=env_().getProperty('DOMAIN');
  if(!/^[a-z0-9][a-z0-9._%+-]*@[a-z0-9.-]+\.[a-z]{2,}$/.test(e) || e.split('@')[1] !== dom) throw new Error('같은 학교 Google Workspace 이메일만 등록할 수 있습니다.');
  if(e===admin.email) throw new Error('최초 관리자 계정은 변경할 수 없습니다.');
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const list=rows_('Users'); const found=list.find(r => r.email===e);
    if(found) tab_('Users').getRange(found.rowNumber,3).setValue(active===true);
    else if(active === true) tab_('Users').appendRow([e,'student',true,new Date().toISOString()]);
    return {ok:true};
  } finally {lock.releaseLock();}
}
function reportPost(id,reason) {
  const who=requireMember_(); const p=requirePostAccess_(id,'enhanced');
  const cleaned=cleanText_(reason,250); if(cleaned.length < 4) throw new Error('신고 내용을 4자 이상 입력해주세요.');
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const today=new Date().toISOString().slice(0,10);
    if(rows_('Reports').filter(r=>r.reportedBy===who.email && sheetDate_(r.reportedAt,false).startsWith(today)).length>=3) throw new Error('하루 신고 횟수 제한을 초과했습니다.');
    tab_('Reports').appendRow([p.id,cleaned,who.email,new Date().toISOString()]);
    return {ok:true};
  } finally {lock.releaseLock();}
}
function listReports() {
  requireAdmin_();
  return rows_('Reports').map(r=>({postId:r.postId,reason:r.reason,reportedBy:r.reportedBy,reportedAt:sheetDate_(r.reportedAt,false)})).reverse().slice(0,100);
}
