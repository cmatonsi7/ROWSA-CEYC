const FIELDS=['reached','bornAgain','spirit','materials'];
const sheet=()=>SpreadsheetApp.getActive().getSheetByName('Cells');
function read(){const v=sheet().getDataRange().getValues(),h=v.shift();
  return v.filter(r=>r[0]).map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i]);FIELDS.forEach(f=>o[f]=Number(o[f])||0);o.updated=o.updated?new Date(o.updated).getTime():null;return o;});}
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function doGet(){const c=CacheService.getScriptCache();let j=c.get('d');
  if(!j){j=JSON.stringify(read());c.put('d',j,5);}
  return ContentService.createTextOutput(j).setMimeType(ContentService.MimeType.JSON);}
function doPost(e){const L=LockService.getScriptLock();L.waitLock(10000);
  try{const p=JSON.parse(e.postData.contents),s=sheet(),v=s.getDataRange().getValues(),h=v[0];
    const col=h.indexOf(p.field),row=v.findIndex((r,i)=>i>0&&r[0]===p.id);
    if(col<0||!FIELDS.includes(p.field)||row<0||(p.delta!==1&&p.delta!==-1))return out({ok:false});
    const n=Math.max(0,(Number(v[row][col])||0)+p.delta);
    s.getRange(row+1,col+1).setValue(n);s.getRange(row+1,h.indexOf('updated')+1).setValue(new Date());
    CacheService.getScriptCache().remove('d');return out({ok:true,value:n});
  }finally{L.releaseLock();}}
