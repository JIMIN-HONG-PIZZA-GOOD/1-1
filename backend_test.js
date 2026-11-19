const vm = require('node:vm');
const fs = require('node:fs');
const assert = require('node:assert/strict');
const code = fs.readFileSync('/mnt/data/class-bulletin-archive/apps-script/Code.gs','utf8');
let email = 'student@school.example';
const props={READY:'yes',DOMAIN:'school.example',ADMIN_EMAIL:'admin@school.example',SHEET_ID:'sheet001',ORIGINAL_FOLDER_ID:'orig',ENHANCED_FOLDER_ID:'enh',THUMB_FOLDER_ID:'thumb'};
const tables={
 Users:[['email','role','active','addedAt'],['admin@school.example','admin',true,'2026-11-19'],['student@school.example','student',true,'2026-11-19']],
 Posts:[['id','title','category','displayDate','description','visibility','uploadedAt','uploadedBy','originalId','enhancedId','thumbId','status','updatedAt'],['p1','축제','행사','2026-11-19','공개용','class','2026-11-19T08:00:00Z','admin@school.example','o1','e1','t1','active','v1'],['p2','제한','공지','2026-11-19','관리자만','admin','2026-11-19T08:00:00Z','admin@school.example','o2','e2','t2','active','v1']],
 Reports:[['postId','reason','reportedBy','reportedAt']]
};
function range(sheet,row,col,nr=1,nc=1){return {getValues:()=>tables[sheet].slice(row-1,row-1+nr).map(r=>r.slice(col-1,col-1+nc)),setValue:(x)=>tables[sheet][row-1][col-1]=x,setValues:(vs)=>vs.forEach((r,i)=>r.forEach((x,j)=>tables[sheet][row-1+i][col-1+j]=x))};}
const sheetStub=name=>({getLastRow:()=>tables[name].length,getLastColumn:()=>tables[name][0].length,getRange:(r,c,nr,nc)=>range(name,r,c,nr,nc),appendRow:r=>tables[name].push(r)});
let openedDriveFiles=[];
const fakeDrive={getFileById:id=>{openedDriveFiles.push(id);return{getBlob:()=>({getSize:()=>130,getContentType:()=> 'image/jpeg',getBytes:()=>Array(130).fill(1)}),setTrashed:x=>{}};},getFolderById:id=>({createFile:blob=>({getId:()=>id+'-saved',setTrashed:()=>{}})})};
const context={
 Session:{getActiveUser:()=>({getEmail:()=>email})},
 PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k]})},
 SpreadsheetApp:{openById:()=>({getSheetByName:sheetStub})},
 DriveApp:fakeDrive,
 Utilities:{formatDate:(d,z,f)=>d.toISOString(),base64Encode:()=>'',base64Decode: b=>Buffer.from(b,'base64'),newBlob:(b,m,n)=>({setName:()=>({})}),getUuid:()=> 'p3'},
 LockService:{getScriptLock:()=>({tryLock:()=>true,waitLock:()=>true,releaseLock:()=>{}})},
 console,Date,Buffer
};
vm.createContext(context);
vm.runInContext(code,context);
function run(expr){return vm.runInContext(expr,context);}
assert.equal(run('getBootstrap().user.role'),'student');
assert.equal(run('listPosts().length'),1);
assert.throws(()=>run("getImage('p1','original')"),/원본/);
assert.throws(()=>run("getImage('p2','enhanced')"),/접근/);
assert.deepEqual(openedDriveFiles,[]);
assert.throws(()=>run("listUsers()"),/관리자/);
assert.throws(()=>run("setUser('new@school.example',true)"),/관리자/);
email='outside@gmail.com';
assert.throws(()=>run('getBootstrap()'),/학교 계정/);
email='unlisted@school.example';
assert.throws(()=>run('listPosts()'),/허용된/);
email='';
assert.throws(()=>run('listPosts()'),/학교 계정/);
email='admin@school.example';
assert.equal(run('getBootstrap().user.role'),'admin');
assert.equal(run('listPosts().length'),2);
assert.equal(run("getImage('p1','original').mime"),'image/jpeg');
assert.equal(openedDriveFiles.at(-1),'o1');
assert.throws(()=>run("setUser('=danger@school.example',true)"),/이메일/);
assert.equal(run("setUser('new@school.example',true).ok"),true);
email='new@school.example';
assert.equal(run('getBootstrap().user.role'),'student');
assert.equal(run('listPosts().length'),1);
console.log('backend access-control tests passed');
