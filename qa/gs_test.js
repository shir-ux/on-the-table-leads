// הרצה יבשה של הסקריפט מול גיליון מדומה בזיכרון
const fs=require('fs');
function mkSheet(name, rows, opts){ opts=opts||{};
  const sh={name, rows, fmt:{},
    getLastRow(){ let n=this.rows.length; while(n>0 && this.rows[n-1].every(v=>v===''||v==null)) n--; return n; },
    getLastColumn(){ return Math.max(0,...this.rows.map(r=>r.length)); },
    getMaxRows(){ return 1000; },
    getRange(r,c,nr,nc){ nr=nr||1; nc=nc||1; const self=this;
      const get=()=>{const out=[]; for(let i=0;i<nr;i++){const row=[]; for(let j=0;j<nc;j++){ const v=(self.rows[r-1+i]||[])[c-1+j]; row.push(v==null?'':v);} out.push(row);} return out;};
      const disp=v=> v instanceof Date ? 'DATE' : (v===true?'TRUE':v===false?'FALSE':String(v));
      const rng={ getValues:get, getDisplayValues:()=>get().map(row=>row.map(disp)), getDisplayValue:()=>disp(get()[0][0]),
        setValue(v){ while(self.rows.length<r) self.rows.push([]); self.rows[r-1][c-1]=v; return rng; },
        setValues(vs){ vs.forEach((row,i)=>{ while(self.rows.length<r+i) self.rows.push([]); row.forEach((v,j)=>{ self.rows[r-1+i][c-1+j]=v; }); }); return rng; },
        setNumberFormat(f){ self.fmt[r+','+c]=f; return rng; }, getNumberFormat(){ return self.fmt[r+','+c]||''; },
        setFontWeight(){ return rng; }, insertCheckboxes(){ return rng; },
        getDataValidation(){ return opts.validation && opts.validation[c] ? {getCriteriaType:()=>'LIST', getCriteriaValues:()=>[opts.validation[c]]} : null; },
        getColumn:()=>c, getLastColumn:()=>c+nc-1 };
      return rng; },
    getConditionalFormatRules(){ return []; }, setRightToLeft(){}, setFrozenRows(){}, hideColumns(){}, setColumnWidth(){} };
  return sh; }
function run(leadRows, validation){
  const sheets={ 'גיליון1': mkSheet('גיליון1', leadRows, {validation}) };
  const ss={ getSheetByName:n=>sheets[n]||null, insertSheet:n=>(sheets[n]=mkSheet(n,[])) };
  const env={ SpreadsheetApp:{ getActive:()=>ss, DataValidationCriteria:{VALUE_IN_LIST:'LIST',VALUE_IN_RANGE:'RANGE'}, BooleanCriteria:{} },
    Utilities:{ formatDate:(d,tz,f)=>{const p=n=>String(n).padStart(2,'0'); return f.replace('dd',p(d.getDate())).replace('MM',p(d.getMonth()+1)).replace('yyyy',d.getFullYear()).replace('HH',p(d.getHours())).replace('mm',p(d.getMinutes()));}, getUuid:()=>'id'+Math.random() },
    Session:{ getScriptTimeZone:()=>'Asia/Jerusalem' }, LockService:{ getScriptLock:()=>({waitLock(){},releaseLock(){}}) },
    ContentService:{ createTextOutput:t=>({t, setMimeType(){return this;}}), MimeType:{JSON:'json'} } };
  const src=fs.readFileSync(require('path').join(__dirname,'..','apps-script.gs'),'utf8').replace("var PIN = 'CHANGE_ME'","var PIN = 'T'");
  const api=new Function(...Object.keys(env), src+'; return (body)=>JSON.parse(doPost({postData:{contents:JSON.stringify(Object.assign({pin:"T"},body))}}).t);')(...Object.values(env));
  return {api, sheets};
}
const H=['תאריך פנייה','שם מלא','טלפון','קמפיין','מודעה','שם הקהל','סטטוס','סיבת אי-סגירה','דיווח והערות','תאריך פולואפ','נסגרה עסקה','שווי הזמנה ₪'];
const row=['15-09-2026 00:01','דנה','972500000001','קמפיין א','קרוסלה','קהל רגיל','אין מענה 1','','הערה ישנה','',false,''];
const V={7:['אין מענה 1','נשלחה הצעה','לא רלוונטי']};
let t=run([H,row.slice()], V), L, u, h;
L=t.api({action:'list'});
console.log('1 list:', L.ok, '| audience:', L.leads[0].audience, '| updated key present, empty:', L.leads[0].updated==='');
u=t.api({action:'update',row:2,expect:{name:'דנה',phone:'972500000001'},fields:{status:'נשלחה הצעה',notes:'שלחתי הצעה',followup:'2026-10-08T10:00'}});
h=t.api({action:'history',phone:'0500000001'});
console.log('2 update:', u.ok, '| audience untouched:', t.sheets['גיליון1'].rows[1][5], '| history events:', h.events.length, '| no extra column written:', t.sheets['גיליון1'].rows[1].length===12);
// ---- עם עמודת "עדכון אחרון" ----
t=run([H.concat(['עדכון אחרון']),row.slice()], V);
t.api({action:'update',row:2,expect:{name:'דנה',phone:'972500000001'},fields:{status:'נשלחה הצעה'}});
console.log('3 stamped is Date:', t.sheets['גיליון1'].rows[1][12] instanceof Date, '| format:', t.sheets['גיליון1'].fmt['2,13']);
L=t.api({action:'list'}); console.log('4 list returns it:', L.leads[0].updated);
t.sheets['גיליון1'].rows[1][12]='';
t.api({action:'taskAdd',fields:{leadName:'דנה',leadPhone:'0500000001',text:'x',due:''}});
console.log('5 adding a task stamps the lead (matched by phone):', t.sheets['גיליון1'].rows[1][12] instanceof Date);
t.sheets['גיליון1'].rows[1][12]='';
t.api({action:'taskDone',id:t.sheets['משימות'].rows[1][0]});
console.log('6 completing a task stamps the lead:', t.sheets['גיליון1'].rows[1][12] instanceof Date);
// ---- העמודה במקום אחר ----
t=run([['עדכון אחרון'].concat(H),[''].concat(row)], {8:V[7]});
u=t.api({action:'update',row:2,expect:{name:'דנה',phone:'972500000001'},fields:{notes:'בדיקה'}});
console.log('7 column placed first:', u.ok, t.sheets['גיליון1'].rows[1][0] instanceof Date, '| notes col:', t.sheets['גיליון1'].rows[1][9]);
// ---- משימות שבוצעו ----
t=run([H.concat(['עדכון אחרון']),row.slice()], V);
t.api({action:'taskAdd',fields:{leadName:'דנה',leadPhone:'0500000001',text:'ראשונה',due:''}});
t.api({action:'taskAdd',fields:{leadName:'דנה',leadPhone:'972500000001',text:'שנייה',due:'2026-10-08T10:00'}});
t.api({action:'taskAdd',fields:{leadName:'אחרת',leadPhone:'0529999999',text:'לא שלה',due:''}});
const ids=t.sheets['משימות'].rows.slice(1).map(r=>r[0]);
t.api({action:'taskDone',id:ids[0]}); t.api({action:'taskDone',id:ids[2]});
let dn=t.api({action:'tasksDone',phone:'972500000001'});
console.log('8 done for this lead only:', dn.ok, dn.tasks.map(x=>x.text).join(','), '| has doneAt:', !!dn.tasks[0].doneAt);
console.log('9 open list excludes done:', t.api({action:'list'}).tasks.map(x=>x.text).join(','));
t.api({action:'taskDone',id:ids[0],done:false});
console.log('10 reopened:', t.api({action:'tasksDone',phone:'0500000001'}).tasks.length, '| open now:', t.api({action:'list'}).tasks.length, '| doneAt cleared:', t.sheets['משימות'].rows[1][7]==='');
console.log('11 reopen logged:', t.api({action:'history',phone:'0500000001'}).events[0].type);
