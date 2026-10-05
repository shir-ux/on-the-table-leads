/* QA מקצה לקצה: הדף האמיתי מול הסקריפט האמיתי, על גיליון מדומה בזיכרון */
(function(){
const errors=[]; window.addEventListener('error',e=>errors.push(String(e.message)));
window.addEventListener('unhandledrejection',e=>errors.push('promise: '+String(e.reason)));
document.addEventListener('click',e=>{ if(e.target.closest('a[href^="tel:"],a[target="_blank"]')) e.preventDefault(); },true);
const p2=n=>String(n).padStart(2,'0');
function fmt(d,f){ return f.replace('dd',p2(d.getDate())).replace('MM',p2(d.getMonth()+1)).replace('yyyy',d.getFullYear()).replace('HH',p2(d.getHours())).replace('mm',p2(d.getMinutes())); }
function mkSheet(name, rows, validation){
  const sh={name, rows, fmt:{},
    getLastRow(){ let n=this.rows.length; while(n>0 && this.rows[n-1].every(v=>v===''||v==null)) n--; return n; },
    getLastColumn(){ return Math.max(0,...this.rows.map(r=>r.length)); }, getMaxRows(){ return 1000; },
    getRange(r,c,nr,nc){ nr=nr||1; nc=nc||1; const self=this;
      const get=()=>{const out=[]; for(let i=0;i<nr;i++){const row=[]; for(let j=0;j<nc;j++){ const v=(self.rows[r-1+i]||[])[c-1+j]; row.push(v==null?'':v);} out.push(row);} return out;};
      const disp=v=> v instanceof Date ? fmt(v,'dd/MM/yyyy HH:mm') : (v===true?'TRUE':v===false?'FALSE':String(v));
      const rng={ getValues:get, getDisplayValues:()=>get().map(row=>row.map(disp)), getDisplayValue:()=>disp(get()[0][0]),
        setValue(v){ while(self.rows.length<r) self.rows.push([]); self.rows[r-1][c-1]=v; return rng; },
        setValues(vs){ vs.forEach((row,i)=>{ while(self.rows.length<r+i) self.rows.push([]); row.forEach((v,j)=>{ self.rows[r-1+i][c-1+j]=v; }); }); return rng; },
        setNumberFormat(f){ self.fmt[r+','+c]=f; return rng; }, getNumberFormat(){ return self.fmt[r+','+c]||''; },
        setFontWeight(){ return rng; }, insertCheckboxes(){ return rng; },
        getDataValidation(){ return validation && validation[c] ? {getCriteriaType:()=>'LIST', getCriteriaValues:()=>[validation[c]]} : null; },
        getColumn:()=>c, getLastColumn:()=>c+nc-1 };
      return rng; },
    getConditionalFormatRules(){ return []; }, setRightToLeft(){}, setFrozenRows(){}, hideColumns(){}, setColumnWidth(){} };
  return sh; }
const now=new Date(), at=(days,h,m)=>new Date(now.getFullYear(),now.getMonth(),now.getDate()+days,h,m);
const ST=['נוצר קשר ראשוני','אין מענה 1','אין מענה 2','אין מענה 3','מעוניינת - לחזור בהמשך','נשלחה הצעה','הזמנה נסגרה','לא רלוונטי'];
const H=['תאריך פנייה','שם מלא','טלפון','קמפיין','מודעה','שם הקהל','סטטוס','סיבת אי-סגירה','דיווח והערות','תאריך פולואפ','נסגרה עסקה','שווי הזמנה ₪','עדכון אחרון'];
const C={created:0,name:1,phone:2,campaign:3,ad:4,audience:5,status:6,reason:7,notes:8,followup:9,closed:10,value:11,updated:12};
const mk=(n,name,phone,status,extra)=>{ const r=['0'+n+'-10-2026 10:0'+n,name,phone,'לידים | מגשי אירוח | 699₪','קרוסלה','קהל רגיל | נשים | 22-50',status,'','','',false,'','']; Object.assign(r,extra||{}); return r; };
const rows=[H,
  mk(1,'דנה','972500000001',''),
  mk(2,'מיכל','972500000002','אין מענה 1'),
  mk(3,'רותי','972500000003','נשלחה הצעה',{9:at(0,16,0)}),
  mk(4,'שרה','972500000004','הזמנה נסגרה',{10:true,11:700}),
  mk(5,'לאה','972500000005','לא רלוונטי',{7:'אחר',8:'לא מתאים'}),
  mk(6,'נועה','972500000006',''),
  mk(7,'טל','972500000007','מעוניינת - לחזור בהמשך',{9:at(5,0,0)})];
const sheets={ 'גיליון1': mkSheet('גיליון1', rows, {7:ST, 8:['יקר לה','תאריך תפוס','אחר']}) };
const ss={ getSheetByName:n=>sheets[n]||null, insertSheet:n=>(sheets[n]=mkSheet(n,[])) };
const env={ SpreadsheetApp:{ getActive:()=>ss, DataValidationCriteria:{VALUE_IN_LIST:'LIST',VALUE_IN_RANGE:'RANGE'}, BooleanCriteria:{} },
  Utilities:{ formatDate:(d,tz,f)=>fmt(d,f), getUuid:()=>'id'+Math.random().toString(36).slice(2) },
  Session:{ getScriptTimeZone:()=>'Asia/Jerusalem' }, LockService:{ getScriptLock:()=>({waitLock(){},releaseLock(){}}) },
  ContentService:{ createTextOutput:t=>({t, setMimeType(){return this;}}), MimeType:{JSON:'json'} } };
const src=JSON.parse(document.getElementById('gsSrc').textContent).replace("var PIN = 'CHANGE_ME'","var PIN = 'T'");
const doPost=new Function(...Object.keys(env), src+'; return doPost;')(...Object.values(env));
let badPin=false; const calls=[]; const failNext={};   // failNext[action]=n: n הבקשות הבאות מהסוג הזה נופלות ברשת
window.fetch=async(u,o)=>{ const b=JSON.parse(o.body); if(badPin) b.pin='zzz'; calls.push(b);
  if(failNext[b.action]>0){ failNext[b.action]--; throw new TypeError('Load failed'); }
  const out=JSON.parse(doPost({postData:{contents:JSON.stringify(b)}}).t); return {ok:true, json:async()=>out}; };

const wait=ms=>new Promise(x=>setTimeout(x,ms)), settle=()=>wait(120);
const L=sheets['גיליון1'], cell=(row,k)=>L.rows[row-1][C[k]];
const card=row=>document.querySelector('#leadsList .lead[data-row="'+row+'"]');
const chip=k=>document.querySelector('.chip[data-f="'+k+'"]');
const pick=v=>[...document.querySelectorAll('#pStatus .pick')].find(b=>b.dataset.v===v);
const shown=()=>[...document.querySelectorAll('#leadsList .lead')].map(c=>+c.dataset.row);
const type=(el,v)=>{ el.value=v; el.dispatchEvent(new Event('input',{bubbles:true})); };
async function pickDate(day,time){ [...document.querySelectorAll('#dtGrid .day')].find(d=>d.textContent===String(day)).click(); const s=$('dtTime'); s.value=time; s.dispatchEvent(new Event('change')); $('dtOk').click(); await settle(); }
const results=[]; async function T(id,name,fn){ try{ const r=await fn(); results.push({id,name,ok:r===true||r&&r.ok===true,detail:r===true?'':JSON.stringify(r)}); }catch(e){ results.push({id,name,ok:false,detail:'EXC '+e.message}); } }
const ok=(cond,detail)=>cond?true:{ok:false,detail};

(async()=>{
await wait(200);
await T(1,'כניסה עם קוד שגוי נחסמת',async()=>{ type($('pinInput'),'9999'); $('loginBtn').click(); await settle(); return ok($('appView').hidden && !!$('loginErr').textContent,{app:$('appView').hidden,err:$('loginErr').textContent}); });
await T(2,'כניסה עם קוד נכון טוענת לידים, ברירת מחדל "חדשים"',async()=>{ type($('pinInput'),'T'); $('loginBtn').click(); await settle(); return ok(!$('appView').hidden && state.leads.length===7 && String(shown())==='7,2',{n:state.leads.length,shown:shown()}); });
await T(3,'סרגל הקבוצות: ספירות וסינון נכונים',async()=>{ const cnt=k=>+chip(k).querySelector('b').textContent; chip('noanswer').click(); const a=String(shown()); chip('progress').click(); const b=String(shown()); chip('all').click(); return ok(cnt('new')===2&&cnt('noanswer')===1&&cnt('progress')===2&&cnt('all')===7&&a==='3'&&b==='8,4'&&shown().length===7,{a,b,c:[cnt('new'),cnt('noanswer'),cnt('progress'),cnt('all')]}); });
await T(4,'חיפוש לפי טלפון מוצא גם ליד סגור',async()=>{ chip('new').click(); type($('searchInput'),'0500000004'); const s=String(shown()); type($('searchInput'),''); chip('all').click(); return ok(s==='5',s); });
await T(5,'הערה בכרטיס נשמרת לבד; קהל לא נדרס; עדכון אחרון נחתם; נרשם בהיסטוריה',async()=>{ const n=card(3).querySelector('[data-q=notes]'); n.focus(); type(n,'הערת בדיקה'); const before=calls.filter(c=>c.action==='update').length; n.blur(); n.dispatchEvent(new Event('change',{bubbles:true})); await settle(); const h=sheets['היסטוריה']; return ok(calls.filter(c=>c.action==='update').length===before+1 && cell(3,'notes')==='הערת בדיקה' && cell(3,'audience')==='קהל רגיל | נשים | 22-50' && cell(3,'updated') instanceof Date && h && h.rows.some(r=>r[3]==='הערה'&&r[4]==='הערת בדיקה'),{notes:cell(3,'notes'),aud:cell(3,'audience'),upd:String(cell(3,'updated')),hist:!!h}); });
await T(6,'תאריך ושעת מעקב מהכרטיס: נשמר, ונקרא חזרה נכון אחרי רענון',async()=>{ card(3).querySelector('[data-dt]').click(); await pickDate(20,'10:30'); const v=cell(3,'followup'); await loadLeads(false); const label=card(3).querySelector('[data-dt]').textContent; return ok(v instanceof Date && v.getDate()===20 && v.getHours()===10 && v.getMinutes()===30 && label.includes('20.') && label.includes('10:30'),{v:String(v),label}); });
await T(7,'חיוג: ליד חדש הופך ל"נוצר קשר ראשוני", ליד מתקדם לא משתנה',async()=>{ card(2).querySelector('[data-call]').click(); await settle(); card(4).querySelector('[data-call]').click(); await settle(); return ok(cell(2,'status')==='נוצר קשר ראשוני' && cell(4,'status')==='נשלחה הצעה',{a:cell(2,'status'),b:cell(4,'status')}); });
await T(8,'תיעוד ליד: לחיצה על סטטוס שומרת מיד וסוגרת',async()=>{ card(3).querySelector('.doc-btn').click(); const hid=$('saveBtn').hidden; pick('אין מענה 2').click(); await settle(); return ok(hid && cell(3,'status')==='אין מענה 2' && $('overlay').hidden,{hid,st:cell(3,'status'),closed:$('overlay').hidden}); });
await T(9,'"לא רלוונטי": אי אפשר לשמור בלי סיבה ודיווח; אחרי מילוי הכל נכתב',async()=>{ card(7).querySelector('.doc-btn').click(); pick('לא רלוונטי').click(); const d0=$('outSave').disabled, stBefore=cell(7,'status'); document.querySelector('#pReason .pick').click(); const d1=$('outSave').disabled; type($('outNote'),'רק ביררה מחיר'); const d2=$('outSave').disabled; $('outSave').click(); await settle(); return ok(d0&&d1&&!d2&&stBefore===''&&cell(7,'status')==='לא רלוונטי'&&cell(7,'reason')==='יקר לה'&&cell(7,'notes')==='רק ביררה מחיר'&&$('outPop').hidden&&$('overlay').hidden,{d0,d1,d2,st:cell(7,'status'),rs:cell(7,'reason'),nt:cell(7,'notes')}); });
await T(10,'יציאה מ"לא רלוונטי" מנקה את הסיבה',async()=>{ card(7).querySelector('.doc-btn').click(); pick('אין מענה 1').click(); await settle(); return ok(cell(7,'status')==='אין מענה 1'&&cell(7,'reason')==='',{st:cell(7,'status'),rs:cell(7,'reason')}); });
await T(11,'"הזמנה נסגרה": חובה שווי; נכתבים סטטוס, סכום ווי',async()=>{ card(2).querySelector('.doc-btn').click(); pick('הזמנה נסגרה').click(); const open=!$('winPop').hidden; type($('winValue'),'0'); const d0=$('winSave').disabled; type($('winValue'),'850'); const d1=$('winSave').disabled; $('winSave').click(); await settle(); return ok(open&&d0&&!d1&&cell(2,'status')==='הזמנה נסגרה'&&cell(2,'value')===850&&cell(2,'closed')===true,{open,d0,d1,st:cell(2,'status'),v:cell(2,'value'),c:cell(2,'closed')}); });
await T(12,'יציאה מ"הזמנה נסגרה" מורידה את הווי, הסכום נשאר',async()=>{ card(2).querySelector('.doc-btn').click(); pick('נשלחה הצעה').click(); await settle(); return ok(cell(2,'status')==='נשלחה הצעה'&&cell(2,'closed')===false&&cell(2,'value')===850,{st:cell(2,'status'),c:cell(2,'closed'),v:cell(2,'value')}); });
await T(13,'הוספת ליד חדש ידנית',async()=>{ $('addBtn').click(); const vis=!$('saveBtn').hidden; type($('fName'),'ליד ידני'); type($('fPhone'),'0541112222'); $('saveBtn').click(); await settle(); const r=L.rows[8]||[]; return ok(vis&&r[C.name]==='ליד ידני'&&String(r[C.phone])==='0541112222'&&state.leads.some(l=>l.name==='ליד ידני'),{vis,row:r.slice(0,7)}); });
await T(14,'הוספת משימה לליד: נשמרת בלשונית משימות, מונה בכרטיס, מופיעה בדף המשימות',async()=>{ card(3).querySelector('[data-task]').click(); type($('taskText'),'לשלוח תפריט'); $('taskDue').click(); await pickDate(25,'09:00'); const noLeadSave=cell(3,'followup').getDate()===20; $('taskSave').click(); await settle(); const ts=sheets['משימות'], tr=ts&&ts.rows[1]; const cnt=(card(3).querySelector('.tcount')||{}).textContent; const row=[...document.querySelectorAll('#taskList .task')].find(x=>x.querySelector('.ttext').textContent.includes('לשלוח תפריט')); return ok(noLeadSave&&tr&&tr[4]==='לשלוח תפריט'&&tr[5] instanceof Date&&tr[5].getDate()===25&&tr[6]===false&&cnt==='1'&&row&&row.querySelector('.who').textContent==='מיכל'&&row.querySelector('.tdue').textContent.includes('09:00'),{noLeadSave,tr:tr&&tr.slice(2,7).map(String),cnt,row:!!row}); });
await T(14.5,'מונה המשימות בכרטיס סופר גם שיחת מעקב שהגיע זמנה',async()=>{ const c=(card(4).querySelector('.tcount')||{}).textContent, none=!card(8).querySelector('.tcount'); return ok(c==='1'&&none,{c,none}); });
await T(15,'פולואפ להיום מופיע כ"שיחת מעקב"; פולואפ עתידי לא',async()=>{ const rs=[...document.querySelectorAll('#taskList .task[data-kind=follow]')].map(x=>x.querySelector('.who').textContent); return ok(rs.includes('רותי')&&!rs.includes('טל'),rs); });
await T(16,'סימון משימה כבוצעה, הצגה ב"בוצעו" בכרטיס הלקוח, ופתיחה מחדש',async()=>{ const row=[...document.querySelectorAll('#taskList .task[data-kind=task]')][0]; row.querySelector('[data-done]').click(); await settle(); const ts=sheets['משימות'].rows[1]; const a=ts[6]===true&&!!ts[7]&&state.tasks.length===0; openLeadPage(state.leads.find(l=>l.row===3)); document.querySelector('.ltab[data-lptab=tasks]').click(); await settle(); const done=document.querySelectorAll('#lpDone .task').length; document.querySelector('#lpDone [data-undone]').click(); await settle(); const b=sheets['משימות'].rows[1][6]===false&&state.tasks.length===1&&document.querySelectorAll('#lpTasks .task').length>=1; return ok(a&&done===1&&b,{a,done,b}); });
await T(17,'כרטיס לקוח: פרטים, תגיות מקור, תגובה שנכנסת להיסטוריה (החדשה ראשונה)',async()=>{ document.querySelector('.ltab[data-lptab=details]').click(); const tags=[...document.querySelectorAll('#lpDetails .tag')].map(t=>t.textContent); const upd=$('lpUpdated').textContent; type($('lpComment'),'תגובה מהכרטיס'); $('lpSend').click(); await settle(); await wait(150); const first=(document.querySelector('#lpHistory .ev-text')||{}).textContent; return ok(tags.length===3&&tags[1].includes('קהל רגיל')&&/\d\d\.\d\d\.\d{4}/.test(upd)&&cell(3,'notes')==='תגובה מהכרטיס'&&first==='תגובה מהכרטיס'&&document.querySelectorAll('#lpHistory .ev').length>=5,{tags,upd,first,n:document.querySelectorAll('#lpHistory .ev').length}); });
await T(18,'תאריך מעקב מתוך כרטיס הלקוח נשמר',async()=>{ $('lpFollow').click(); await pickDate(22,'14:00'); const v=cell(3,'followup'); return ok(v instanceof Date&&v.getDate()===22&&v.getHours()===14&&!$('leadPage').hidden,String(v)); });
await T(19,'קישור אישי לליד (#lead=שורה) פותח את הליד הנכון',async()=>{ $('lpBack').click(); history.replaceState(null,'','#lead=5'); window.dispatchEvent(new HashChangeEvent('hashchange')); await settle(); const t=$('sheetTitle').textContent, o=!$('overlay').hidden; closeEdit(); history.replaceState(null,'','#'); return ok(o&&t==='שרה',{o,t}); });
await T(19.1,'בקשת היסטוריה שנופלת ברשת מנסה שוב לבד, בלי הודעת שגיאה',async()=>{ failNext.history=1; openLeadPage(state.leads.find(l=>l.row===3)); await wait(1500); const txt=$('lpHistory').textContent, n=document.querySelectorAll('#lpHistory .ev').length; $('lpBack').click(); return ok(n>=5&&!txt.includes('לא נטענה'),{n,txt:txt.slice(0,60)}); });
await T(19.2,'אם גם הניסיונות החוזרים נכשלים: הודעה וכפתור "נסי שוב" שעובד',async()=>{ failNext.history=3; openLeadPage(state.leads.find(l=>l.row===3)); await wait(4200); const hasBtn=!!$('lpRetry'); if(hasBtn) $('lpRetry').click(); await settle(); const n=document.querySelectorAll('#lpHistory .ev').length; $('lpBack').click(); return ok(hasBtn&&n>=5,{hasBtn,n}); });
await T(19.3,'בקשת כתיבה שנופלת לא נשלחת פעמיים',async()=>{ const before=calls.filter(c=>c.action==='update').length; failNext.update=1; card(4).querySelector('.doc-btn').click(); pick('אין מענה 3').click(); await wait(1500); const sent=calls.filter(c=>c.action==='update').length-before; const st=cell(4,'status'); closeEdit(); await loadLeads(false); return ok(sent===1&&st==='נשלחה הצעה',{sent,st}); });
await T(20,'הגיליון מוין מאחורי הגב: הדף לא כותב לשורה הלא נכונה',async()=>{ const a=L.rows[5].slice(), b=L.rows[6].slice(); L.rows[5]=b; L.rows[6]=a; card(6).querySelector('.doc-btn').click(); pick('אין מענה 3').click(); await settle(); await wait(150); const wrote=L.rows.some(r=>r[C.status]==='אין מענה 3'); closeEdit(); return ok(!wrote,{wrote}); });
await T(21,'קוד כניסה שהוחלף באמצע עבודה מחזיר למסך הכניסה',async()=>{ badPin=true; await loadLeads(false); await settle(); const r=$('appView').hidden&&!$('loginView').hidden; badPin=false; return ok(r,{app:$('appView').hidden,login:$('loginView').hidden}); });
await T(22,'אין שגיאות קוד בדף לאורך כל הריצה',async()=>ok(errors.length===0,errors));
window.__qa={done:true, vw:innerWidth, pass:results.filter(r=>r.ok).length, total:results.length, failed:results.filter(r=>!r.ok)};
const out=document.createElement('pre'); out.id='qaOut'; out.textContent=JSON.stringify({pass:window.__qa.pass,total:window.__qa.total,vw:innerWidth,results}); document.body.appendChild(out);
})();
})();
