
const DB_NAME='xingyu_v1';
const DB_VERSION=2;
let db;
let route='home';
let exchangeView='all';
let exchangeEventFilter='all';
let squatFilter='all';
let squatEventFilter='all';
let homeEventId=null;

function openDB(){
  return new Promise((resolve,reject)=>{
    const req=indexedDB.open(DB_NAME,DB_VERSION);
    req.onupgradeneeded=()=>{
      const d=req.result;
      ['events','materials','contacts','partnerMaterials','exchanges','squats'].forEach(s=>{
        if(!d.objectStoreNames.contains(s)) d.createObjectStore(s,{keyPath:'id'});
      });
    };
    req.onsuccess=()=>resolve(req.result);
    req.onerror=()=>reject(req.error);
  });
}
function tx(store,mode='readonly'){return db.transaction(store,mode).objectStore(store)}
function getAll(store){return new Promise((res,rej)=>{const r=tx(store).getAll();r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error)})}
function getOne(store,id){return new Promise((res,rej)=>{const r=tx(store).get(id);r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error)})}
function put(store,obj){return new Promise((res,rej)=>{const r=tx(store,'readwrite').put(obj);r.onsuccess=()=>res(obj);r.onerror=()=>rej(r.error)})}
function del(store,id){return new Promise((res,rej)=>{const r=tx(store,'readwrite').delete(id);r.onsuccess=()=>res();r.onerror=()=>rej(r.error)})}
function clearStore(store){return new Promise((res,rej)=>{const r=tx(store,'readwrite').clear();r.onsuccess=()=>res();r.onerror=()=>rej(r.error)})}
function uid(prefix){return prefix+'_'+Math.random().toString(36).slice(2,9)}
function el(html){const t=document.createElement('template');t.innerHTML=html.trim();return t.content.firstElementChild}
function cloneTemplate(id){return document.getElementById(id).content.cloneNode(true)}
function setTitle(t){document.getElementById('page-title').textContent=t}
function setActiveTab(){document.querySelectorAll('.tabbar button').forEach(b=>b.classList.toggle('active',b.dataset.route===route))}
function statusLabel(s){return ({booked:'已约定',onsite:'待交换',done:'已完成',cancelled:'已取消'})[s]||s}
function squatStatusLabel(s){return ({want:'想蹲',got:'已拿到',missed:'没拿到'})[s]||s}
function fileToDataURL(file){
  return new Promise((resolve,reject)=>{
    if(!file){resolve('');return;}
    const r=new FileReader(); r.onload=()=>resolve(r.result); r.onerror=()=>reject(r.error); r.readAsDataURL(file);
  });
}
async function filesToDataURLs(files){
  const arr=[...(files||[])];
  return Promise.all(arr.map(fileToDataURL));
}
function escapeHtml(s=''){return String(s).replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]))}
function materialChoiceLabel(material, subItemId){
  if(!material) return '物料';
  if(subItemId){
    const sub=(material.subItems||[]).find(s=>s.id===subItemId);
    return sub ? `${material.name} · ${sub.name}` : material.name;
  }
  return material.name;
}
function materialChoices(materials){
  const out=[];
  materials.forEach(m=>{
    const subs=m.subItems||[];
    if(subs.length){
      subs.forEach(s=>out.push({value:`${m.id}::${s.id}`,label:`${m.name} · ${s.name}`}));
    }else{
      out.push({value:`${m.id}::`,label:m.name});
    }
  });
  return out;
}

function ensureImageLightbox(){
  let box=document.getElementById('image-lightbox');
  if(box) return box;
  box=document.createElement('div');
  box.id='image-lightbox';
  box.className='image-lightbox';
  box.innerHTML=`<div class="image-lightbox-inner">
    <button type="button" class="image-lightbox-close" aria-label="关闭">×</button>
    <img id="image-lightbox-img" alt="大图预览">
  </div>`;
  document.body.appendChild(box);
  box.addEventListener('click',e=>{
    if(e.target===box || e.target.closest('.image-lightbox-close')) box.classList.remove('open');
  });
  return box;
}
function openImageLightbox(src){
  const box=ensureImageLightbox();
  const img=box.querySelector('#image-lightbox-img');
  img.src=src;
  box.classList.add('open');
}
function countdownInfo(dateStr){
  if(!dateStr) return {days:null,msg:'还没有设置场次日期'};
  const p=dateStr.split('-').map(Number);
  const target=Date.UTC(p[0],p[1]-1,p[2]);
  const now=new Date();
  const today=Date.UTC(now.getFullYear(),now.getMonth(),now.getDate());
  const days=Math.ceil((target-today)/86400000);
  let msg='';
  if(days>30) msg='可以慢慢准备 ✦';
  else if(days>14) msg='无料可以开始动工了';
  else if(days>7) msg='无料需要滑铲了 🛼';
  else if(days>3) msg='打印 / 打包 / 交换确认！';
  else if(days>1) msg='救命，进入最后冲刺 🚨';
  else if(days===1) msg='明天！检查行李和交换清单';
  else if(days===0) msg='就是今天！去现场！！';
  else msg='本场已结束';
  return {days,msg};
}
function migrateExchange(x, contacts, partnerMaterials){
  const c=contacts.find(c=>c.id===x.contactId);
  if(!x.partnerName) x.partnerName=c?.name||c?.wechat||'未命名';
  if(!x.partnerWechat) x.partnerWechat=c?.wechat||'';
  if(!x.receiveItems){
    x.receiveItems=(x.receive||[]).map(r=>{
      const p=partnerMaterials.find(p=>p.id===r.partnerMaterialId);
      return {name:p?.name||'对方物料',images:p?.image?[p.image]:[],qty:Number(r.qty||1)};
    });
  }
  x.receiveItems=x.receiveItems.map(r=>({name:r.name||'对方物料',images:r.images||(r.image?[r.image]:[]),qty:Number(r.qty||1)}));
  return x;
}


/* ---------- V2.4 本地 XLSX / 备份工具 ---------- */
function xmlEscape(v=''){
  return String(v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
}
function colName(n){
  let s='';
  while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}
  return s;
}
function crc32(bytes){
  let table=crc32._table;
  if(!table){
    table=[];
    for(let n=0;n<256;n++){
      let c=n;
      for(let k=0;k<8;k++) c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1);
      table[n]=c>>>0;
    }
    crc32._table=table;
  }
  let c=0xFFFFFFFF;
  for(const b of bytes)c=table[(c^b)&0xFF]^(c>>>8);
  return (c^0xFFFFFFFF)>>>0;
}
function u16(v){return [v&255,(v>>>8)&255]}
function u32(v){return [v&255,(v>>>8)&255,(v>>>16)&255,(v>>>24)&255]}
function zipStore(entries){
  const enc=new TextEncoder();
  const chunks=[], central=[];
  let offset=0;
  for(const entry of entries){
    const nameBytes=enc.encode(entry.name);
    const data=entry.data instanceof Uint8Array?entry.data:enc.encode(entry.data);
    const crc=crc32(data);
    const local=new Uint8Array([
      ...u32(0x04034b50),...u16(20),...u16(0),...u16(0),...u16(0),...u16(0),
      ...u32(crc),...u32(data.length),...u32(data.length),...u16(nameBytes.length),...u16(0),
      ...nameBytes
    ]);
    chunks.push(local,data);
    const cent=new Uint8Array([
      ...u32(0x02014b50),...u16(20),...u16(20),...u16(0),...u16(0),...u16(0),...u16(0),
      ...u32(crc),...u32(data.length),...u32(data.length),...u16(nameBytes.length),...u16(0),...u16(0),
      ...u16(0),...u16(0),...u32(0),...u32(offset),...nameBytes
    ]);
    central.push(cent);
    offset+=local.length+data.length;
  }
  const centralSize=central.reduce((a,b)=>a+b.length,0);
  const end=new Uint8Array([
    ...u32(0x06054b50),...u16(0),...u16(0),...u16(entries.length),...u16(entries.length),
    ...u32(centralSize),...u32(offset),...u16(0)
  ]);
  const total=chunks.reduce((a,b)=>a+b.length,0)+centralSize+end.length;
  const out=new Uint8Array(total);let pos=0;
  for(const c of chunks){out.set(c,pos);pos+=c.length}
  for(const c of central){out.set(c,pos);pos+=c.length}
  out.set(end,pos);
  return out;
}
function makeSheetXml(rows){
  const body=rows.map((row,ri)=>{
    const cells=row.map((v,ci)=>{
      if(v===null||v===undefined||v==='') return '';
      const ref=colName(ci+1)+(ri+1);
      if(typeof v==='number' && Number.isFinite(v)) return `<c r="${ref}"><v>${v}</v></c>`;
      return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(v)}</t></is></c>`;
    }).join('');
    return `<row r="${ri+1}">${cells}</row>`;
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`;
}
function buildXlsx(sheetDefs){
  const contentTypes=[
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`,
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`,
    `<Default Extension="xml" ContentType="application/xml"/>`,
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`,
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`,
    ...sheetDefs.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`),
    `</Types>`
  ].join('');
  const rootRels=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  const workbook=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>${sheetDefs.map((s,i)=>`<sheet name="${xmlEscape(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets>
</workbook>`;
  const wbRels=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheetDefs.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}
<Relationship Id="rId${sheetDefs.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  const styles=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`;
  const entries=[
    {name:'[Content_Types].xml',data:contentTypes},
    {name:'_rels/.rels',data:rootRels},
    {name:'xl/workbook.xml',data:workbook},
    {name:'xl/_rels/workbook.xml.rels',data:wbRels},
    {name:'xl/styles.xml',data:styles},
    ...sheetDefs.map((s,i)=>({name:`xl/worksheets/sheet${i+1}.xml`,data:makeSheetXml(s.rows)}))
  ];
  return zipStore(entries);
}
function downloadBlob(blob,name){
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),1000);
}
function dataMessage(text,type=''){
  const node=document.getElementById('data-message');
  if(!node)return;
  node.textContent=text;node.className='data-message '+type;
}
async function exportExcel(){
  const [events,materials,exchanges,squats]=await Promise.all([getAll('events'),getAll('materials'),getAll('exchanges'),getAll('squats')]);
  const methodLabels={exchange:'互换',handout:'伸手',both:'互换+伸手',keep:'仅自留'};
  const inventoryLabels={shared:'共用总量',allocated:'按场次分配'};
  const exchangeStatus={booked:'已约定',onsite:'待交换',done:'已完成',cancelled:'已取消'};
  const squatStatus={want:'想蹲',got:'已拿到',missed:'没拿到'};
  const sheets=[
    {name:'场次',rows:[
      ['内部ID','场次名称','日期','城市','场馆','倒计时提示'],
      ...events.map(e=>[e.id,e.name,e.date,e.city,e.venue,e.countdownMessage||''])
    ]},
    {name:'我的物料',rows:[
      ['内部ID','名称','总量','预留','发放方式','伸手计划量','伸手已发','库存模式','制作状态','绑定场次ID','备注','图片数量','子项数量','子项明细'],
      ...materials.map(m=>[m.id,m.name,Number(m.qty||0),Number(m.reserved||0),methodLabels[m.distributionMethod||'exchange']||'',Number(m.handoutPlan||0),Number(m.handoutDone||0),inventoryLabels[m.inventoryMode]||'',m.status||'',(m.eventIds||[]).join(';'),m.note||'',m.image?1:0,(m.subItems||[]).length,(m.subItems||[]).map(s=>`${s.name}:${s.qty}`).join('；')])
    ]},
    {name:'互换',rows:[
      ['内部ID','对方昵称','微信号','场次ID','我给物料ID','我给数量','对方无料名称','对方数量','地点','时间','状态','备注','对方图片数量'],
      ...exchanges.map(x=>[x.id,x.partnerName||'',x.partnerWechat||'',x.eventId||'',x.give?.[0]?.materialId||'',Number(x.give?.[0]?.qty||0),x.receiveItems?.[0]?.name||'',Number(x.receiveItems?.[0]?.qty||0),x.place||'',x.time||'',exchangeStatus[x.status]||x.status||'',x.note||'',x.receiveItems?.[0]?.images?.length||0])
    ]},
    {name:'蹲蹲',rows:[
      ['内部ID','老师昵称','小红书','场次ID','地点','时间','领取条件','状态','备注','OOTD是否有图','无料图片数量'],
      ...squats.map(s=>[s.id,s.teacherName||'',s.xhs||'',s.eventId||'',s.place||'',s.time||'',s.condition||'',squatStatus[s.status]||s.status||'',s.note||'',s.ootdImage?1:0,s.freebieImages?.length||0])
    ]}
  ];
  const bytes=buildXlsx(sheets);
  const d=new Date().toISOString().slice(0,10);
  downloadBlob(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),`星屿物料_${d}.xlsx`);
  dataMessage('Excel 已导出。图片不会写入 Excel，请用“完整备份”保存图片。','ok');
}
function findEOCD(bytes){
  for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--){
    if(bytes[i]===0x50&&bytes[i+1]===0x4b&&bytes[i+2]===0x05&&bytes[i+3]===0x06)return i;
  }
  return -1;
}
function readU16(b,o){return b[o]|(b[o+1]<<8)}
function readU32(b,o){return (b[o]|(b[o+1]<<8)|(b[o+2]<<16)|(b[o+3]<<24))>>>0}
async function unzipEntries(arrayBuffer){
  const bytes=new Uint8Array(arrayBuffer),dec=new TextDecoder();
  const eocd=findEOCD(bytes);if(eocd<0)throw new Error('不是有效的 XLSX/ZIP 文件');
  const count=readU16(bytes,eocd+10),cdOffset=readU32(bytes,eocd+16);
  const files={};let p=cdOffset;
  for(let i=0;i<count;i++){
    if(readU32(bytes,p)!==0x02014b50)throw new Error('XLSX 目录结构异常');
    const method=readU16(bytes,p+10),compSize=readU32(bytes,p+20),uncompSize=readU32(bytes,p+24);
    const nameLen=readU16(bytes,p+28),extraLen=readU16(bytes,p+30),commentLen=readU16(bytes,p+32),localOffset=readU32(bytes,p+42);
    const name=dec.decode(bytes.slice(p+46,p+46+nameLen));
    const localNameLen=readU16(bytes,localOffset+26),localExtraLen=readU16(bytes,localOffset+28);
    const dataStart=localOffset+30+localNameLen+localExtraLen;
    const compressed=bytes.slice(dataStart,dataStart+compSize);
    let out;
    if(method===0)out=compressed;
    else if(method===8){
      if(typeof DecompressionStream==='undefined')throw new Error('当前浏览器无法解压这个 Excel，请使用 Chrome/Safari 新版');
      let ds;
      try{ds=new DecompressionStream('deflate-raw')}catch(e){throw new Error('当前浏览器暂不支持 Excel 压缩格式')}
      out=new Uint8Array(await new Response(new Blob([compressed]).stream().pipeThrough(ds)).arrayBuffer());
    }else throw new Error('暂不支持此 Excel 的 ZIP 压缩方式');
    files[name]=out;
    p+=46+nameLen+extraLen+commentLen;
  }
  return files;
}
function nodeText(n){return n?n.textContent:''}
function parseSheetRows(xmlText,sharedStrings=[]){
  const doc=new DOMParser().parseFromString(xmlText,'application/xml');
  const rows=[];
  doc.querySelectorAll('sheetData > row').forEach(row=>{
    const arr=[];
    row.querySelectorAll('c').forEach(c=>{
      const ref=c.getAttribute('r')||'A1';
      const letters=(ref.match(/[A-Z]+/)||['A'])[0];
      let ci=0;for(const ch of letters)ci=ci*26+(ch.charCodeAt(0)-64);ci--;
      const type=c.getAttribute('t');
      let val='';
      if(type==='inlineStr')val=nodeText(c.querySelector('is t'));
      else if(type==='s')val=sharedStrings[Number(nodeText(c.querySelector('v')))]||'';
      else val=nodeText(c.querySelector('v'));
      arr[ci]=val;
    });
    rows.push(arr.map(v=>v??''));
  });
  return rows;
}
async function parseXlsx(file){
  const files=await unzipEntries(await file.arrayBuffer());
  const dec=new TextDecoder();
  const workbookXml=dec.decode(files['xl/workbook.xml']||new Uint8Array());
  const relXml=dec.decode(files['xl/_rels/workbook.xml.rels']||new Uint8Array());
  if(!workbookXml)throw new Error('Excel 缺少 workbook.xml');
  let shared=[];
  if(files['xl/sharedStrings.xml']){
    const sdoc=new DOMParser().parseFromString(dec.decode(files['xl/sharedStrings.xml']),'application/xml');
    shared=[...sdoc.querySelectorAll('si')].map(si=>[...si.querySelectorAll('t')].map(nodeText).join(''));
  }
  const wdoc=new DOMParser().parseFromString(workbookXml,'application/xml');
  const rdoc=new DOMParser().parseFromString(relXml,'application/xml');
  const rels={};
  rdoc.querySelectorAll('Relationship').forEach(r=>rels[r.getAttribute('Id')]=r.getAttribute('Target'));
  const result={};
  wdoc.querySelectorAll('sheet').forEach(s=>{
    const name=s.getAttribute('name'),rid=s.getAttribute('r:id')||s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
    let target=rels[rid]||'';
    if(target.startsWith('/'))target=target.slice(1);
    else if(!target.startsWith('xl/'))target='xl/'+target;
    const bytes=files[target];
    if(bytes)result[name]=parseSheetRows(dec.decode(bytes),shared);
  });
  return result;
}
function rowsToObjects(rows){
  if(!rows?.length)return[];
  const headers=rows[0].map(x=>String(x||'').trim());
  return rows.slice(1).filter(r=>r.some(v=>String(v||'').trim()!=='')).map(r=>{
    const o={};headers.forEach((h,i)=>{if(h)o[h]=r[i]??''});return o;
  });
}
function revMap(obj,label){return Object.entries(obj).find(([,v])=>v===label)?.[0]||label}
async function importExcelFile(file,mode){
  const book=await parseXlsx(file);
  const eventRows=rowsToObjects(book['场次']||[]);
  const materialRows=rowsToObjects(book['我的物料']||[]);
  const exchangeRows=rowsToObjects(book['互换']||[]);
  const squatRows=rowsToObjects(book['蹲蹲']||[]);
  if(!eventRows.length&&!materialRows.length&&!exchangeRows.length&&!squatRows.length)throw new Error('没有找到本 App 的工作表结构');

  const methodLabels={exchange:'互换',handout:'伸手',both:'互换+伸手',keep:'仅自留'};
  const inventoryLabels={shared:'共用总量',allocated:'按场次分配'};
  const exchangeStatus={booked:'已约定',onsite:'待交换',done:'已完成',cancelled:'已取消'};
  const squatStatus={want:'想蹲',got:'已拿到',missed:'没拿到'};

  if(mode==='replace'){
    for(const s of ['events','materials','exchanges','squats'])await clearStore(s);
  }
  const maps={events:{},materials:{}};
  async function newId(store,prefix,oldId){
    if(mode==='merge')return oldId||uid(prefix);
    if(mode==='replace')return oldId||uid(prefix);
    const exists=oldId?await getOne(store,oldId):null;
    return exists?uid(prefix):(oldId||uid(prefix));
  }
  for(const r of eventRows){
    const old=r['内部ID']||'';const id=await newId('events','ev',old);maps.events[old]=id;
    await put('events',{id,name:r['场次名称']||'未命名场次',date:r['日期']||'',city:r['城市']||'',venue:r['场馆']||'',countdownMessage:r['倒计时提示']||'',pinned:false});
  }
  for(const r of materialRows){
    const old=r['内部ID']||'';const id=await newId('materials','mat',old);maps.materials[old]=id;
    const eventIds=String(r['绑定场次ID']||'').split(';').filter(Boolean).map(x=>maps.events[x]||x);
    await put('materials',{
      id,name:r['名称']||'未命名物料',image:'',qty:Number(r['总量']||0),reserved:Number(r['预留']||0),
      distributionMethod:revMap(methodLabels,r['发放方式'])||'exchange',handoutPlan:Number(r['伸手计划量']||0),handoutDone:Number(r['伸手已发']||0),
      inventoryMode:revMap(inventoryLabels,r['库存模式'])||'shared',status:r['制作状态']||'有成品',eventIds,allocations:{},note:r['备注']||'',exchangeable:true
    });
  }
  for(const r of exchangeRows){
    const old=r['内部ID']||'';const id=await newId('exchanges','x',old);
    await put('exchanges',{
      id,partnerName:r['对方昵称']||'',partnerWechat:r['微信号']||'',eventId:maps.events[r['场次ID']]||r['场次ID']||'',
      give:[{materialId:maps.materials[r['我给物料ID']]||r['我给物料ID']||'',qty:Number(r['我给数量']||1)}],
      receiveItems:[{name:r['对方无料名称']||'',images:[],qty:Number(r['对方数量']||1)}],
      place:r['地点']||'',time:r['时间']||'',status:revMap(exchangeStatus,r['状态'])||'onsite',note:r['备注']||'',createdAt:new Date().toISOString(),completedAt:null
    });
  }
  for(const r of squatRows){
    const old=r['内部ID']||'';const id=await newId('squats','sq',old);
    await put('squats',{
      id,teacherName:r['老师昵称']||'',xhs:r['小红书']||'',eventId:maps.events[r['场次ID']]||r['场次ID']||'',
      place:r['地点']||'',time:r['时间']||'',condition:r['领取条件']||'',status:revMap(squatStatus,r['状态'])||'want',
      note:r['备注']||'',ootdImage:'',freebieImages:[],createdAt:new Date().toISOString()
    });
  }
  return {events:eventRows.length,materials:materialRows.length,exchanges:exchangeRows.length,squats:squatRows.length};
}
async function exportBackup(){
  const payload={
    format:'xingyu-material-backup',version:'2.4',exportedAt:new Date().toISOString(),
    events:await getAll('events'),materials:await getAll('materials'),exchanges:await getAll('exchanges'),squats:await getAll('squats')
  };
  const d=new Date().toISOString().slice(0,10);
  downloadBlob(new Blob([JSON.stringify(payload,null,2)],{type:'application/json'}),`星屿物料_完整备份_${d}.json`);
  dataMessage('完整备份已导出，包含本地图片。','ok');
}
async function importBackupFile(file){
  const data=JSON.parse(await file.text());
  if(data.format!=='xingyu-material-backup')throw new Error('不是星屿物料完整备份文件');
  for(const s of ['events','materials','exchanges','squats'])await clearStore(s);
  for(const item of (data.events||[]))await put('events',item);
  for(const item of (data.materials||[]))await put('materials',item);
  for(const item of (data.exchanges||[]))await put('exchanges',item);
  for(const item of (data.squats||[]))await put('squats',item);
  homeEventId=null;exchangeEventFilter='all';squatEventFilter='all';
}

async function render(){
  setActiveTab();
  const view=document.getElementById('view'); view.innerHTML='';
  if(route==='home') return renderHome(view);
  if(route==='materials') return renderMaterials(view);
  if(route==='exchanges') return renderExchanges(view);
  if(route==='squats') return renderSquats(view);
}

async function renderHome(view){
  setTitle('首页'); view.append(cloneTemplate('home-template'));
  const [events,materials,exchanges,squats,contacts,partnerMaterials]=await Promise.all([
    getAll('events'),getAll('materials'),getAll('exchanges'),getAll('squats'),getAll('contacts'),getAll('partnerMaterials')
  ]);
  const sorted=events.slice().sort((a,b)=>(a.date||'9999').localeCompare(b.date||'9999'));
  if(!homeEventId || !events.some(e=>e.id===homeEventId)){
    const today=new Date().toISOString().slice(0,10);
    homeEventId=(sorted.find(e=>(e.date||'')>=today)||sorted[0])?.id||null;
  }
  const select=document.getElementById('home-event-select');
  select.innerHTML=sorted.map(e=>`<option value="${e.id}">${escapeHtml(e.name)}</option>`).join('');
  if(homeEventId) select.value=homeEventId;
  const current=events.find(e=>e.id===homeEventId);
  if(current){
    const c=countdownInfo(current.date);
    document.getElementById('countdown-event-name').textContent=current.name;
    document.getElementById('countdown-days').textContent=c.days===null?'--':Math.max(c.days,0);
    document.getElementById('countdown-message').textContent=(current.countdownMessage||'').trim() || c.msg;
    document.getElementById('countdown-meta').textContent=[current.date,current.city,current.venue].filter(Boolean).join(' · ');
  }else{
    document.getElementById('countdown-message').textContent='先新增一个场次吧';
  }
  select.addEventListener('change',()=>{homeEventId=select.value;render()});

  const all=exchanges.map(x=>migrateExchange(x,contacts,partnerMaterials));
  document.getElementById('m-materials').textContent=materials.length;
  document.getElementById('m-pending').textContent=all.filter(x=>(x.status==='booked'||x.status==='onsite') && (!homeEventId||x.eventId===homeEventId)).length;
  document.getElementById('m-squats').textContent=squats.filter(s=>s.status==='want' && (!homeEventId||s.eventId===homeEventId)).length;
  document.getElementById('m-events').textContent=events.length;

  const list=document.getElementById('home-events-list');
  sorted.forEach(ev=>{
    const node=el(`<div class="item clickable" data-edit-event="${ev.id}">
      <div class="thumb">场次</div>
      <div class="item-main">
        <div class="item-title">${escapeHtml(ev.name)}</div>
        <div class="sub">${[ev.date,ev.city,ev.venue].filter(Boolean).map(escapeHtml).join(' · ')}</div>
      </div>
      ${ev.id===homeEventId?'<div class="badge">当前</div>':''}
    </div>`);
    list.append(node);
  });

  const excelInput=document.getElementById('excel-file-input');
  if(excelInput)excelInput.addEventListener('change',async()=>{
    const file=excelInput.files?.[0];if(!file)return;
    try{
      dataMessage('正在导入 Excel…');
      const mode=document.getElementById('excel-import-mode')?.value||'append';
      const counts=await importExcelFile(file,mode);
      dataMessage(`Excel 导入完成：场次 ${counts.events}、物料 ${counts.materials}、互换 ${counts.exchanges}、蹲蹲 ${counts.squats}`,'ok');
      excelInput.value='';
      setTimeout(()=>render(),500);
    }catch(err){dataMessage(err.message||String(err),'error');excelInput.value='';}
  });
  const backupInput=document.getElementById('backup-file-input');
  if(backupInput)backupInput.addEventListener('change',async()=>{
    const file=backupInput.files?.[0];if(!file)return;
    try{
      dataMessage('正在恢复完整备份…');
      await importBackupFile(file);
      dataMessage('完整备份恢复成功。','ok');
      backupInput.value='';
      setTimeout(()=>render(),500);
    }catch(err){dataMessage(err.message||String(err),'error');backupInput.value='';}
  });
}

async function renderMaterials(view){
  setTitle('我的物料'); view.append(cloneTemplate('materials-template'));
  const [materials,events,exchanges]=await Promise.all([getAll('materials'),getAll('events'),getAll('exchanges')]);
  const list=document.getElementById('materials-list');
  materials.forEach(m=>{
    const activeBooked=exchanges.filter(x=>x.status!=='cancelled'&&x.give?.some(g=>g.materialId===m.id))
      .reduce((sum,x)=>sum+x.give.filter(g=>g.materialId===m.id).reduce((a,g)=>a+Number(g.qty||0),0),0);
    const eventNames=(m.eventIds||[]).map(id=>events.find(e=>e.id===id)?.name).filter(Boolean).join(' / ');
    const available=Math.max(0,Number(m.qty||0)-Number(m.reserved||0)-activeBooked);
    const handoutPlan=Number(m.handoutPlan||0);
    const handoutDone=Number(m.handoutDone||0);
    const method=m.distributionMethod||'exchange';
    const methodLabels={exchange:'互换',handout:'伸手',both:'互换+伸手',keep:'仅自留'};
    const trueAvailable=Math.max(0,Number(m.qty||0)-Number(m.reserved||0)-activeBooked-handoutPlan);
    const subItems=m.subItems||[];
    const subBooked={};
    exchanges.filter(x=>x.status!=='cancelled').forEach(x=>{
      (x.give||[]).filter(g=>g.materialId===m.id && g.subItemId).forEach(g=>{
        subBooked[g.subItemId]=(subBooked[g.subItemId]||0)+Number(g.qty||0);
      });
    });
    const subHtml=subItems.length?`
      <div class="subitem-mini-list">
        ${subItems.map(s=>`<span class="subitem-chip">${escapeHtml(s.name)} ${Number(s.qty||0)}${subBooked[s.id]?` / 已约${subBooked[s.id]}`:''}</span>`).join('')}
      </div>`:'';
    list.append(el(`<div class="item clickable" data-edit-material="${m.id}">
      <div class="thumb">${m.image?`<img src="${m.image}">`:subItems.find(s=>s.image)?.image?`<img src="${subItems.find(s=>s.image).image}">`:'物料图'}</div>
      <div class="item-main">
        <div class="item-title">${escapeHtml(m.name)}</div>
        <div class="sub">${escapeHtml(eventNames||'未绑定场次')}</div>
        <div class="method-tags"><span class="method-tag">${methodLabels[method]||'互换'}</span>${subItems.length?`<span class="method-tag">${subItems.length} 个子项</span>`:''}</div>
        ${subHtml}
        ${subItems.length
          ? `<div class="sub">子项数量分别管理 · 伸手计划 ${handoutPlan}</div>`
          : `<div class="sub">总量 ${m.qty||0} · 预留 ${m.reserved||0} · 已约互换 ${activeBooked} · 伸手计划 ${handoutPlan} · 可用 ${trueAvailable}</div>`}
        ${handoutPlan?`<div class="sub">伸手已发 ${handoutDone} / ${handoutPlan}</div>`:''}
      </div>
      <div class="badge">${m.inventoryMode==='shared'?'共用':'分场'}</div>
    </div>`));
  });
}

async function renderExchanges(view){
  setTitle('互换'); view.append(cloneTemplate('exchanges-template'));
  const [exchanges,events,materials,contacts,partnerMaterials]=await Promise.all([
    getAll('exchanges'),getAll('events'),getAll('materials'),getAll('contacts'),getAll('partnerMaterials')
  ]);
  const all=exchanges.map(x=>migrateExchange(x,contacts,partnerMaterials));
  const eventSelect=document.getElementById('exchange-event-filter');
  eventSelect.innerHTML='<option value="all">全部场次</option>'+events.map(e=>`<option value="${e.id}">${escapeHtml(e.name)}</option>`).join('');
  if(exchangeView==='onsite' && exchangeEventFilter==='all'){
    exchangeEventFilter=homeEventId||events[0]?.id||'all';
  }
  eventSelect.value=exchangeEventFilter;
  const list=document.getElementById('exchanges-list');
  const search=document.getElementById('exchange-search');

  function matchesView(x){
    if(exchangeView==='all') return true;
    if(exchangeView==='pending') return x.status==='booked'||x.status==='onsite';
    if(exchangeView==='done') return x.status==='done';
    if(exchangeView==='onsite') return x.status==='booked'||x.status==='onsite';
    return true;
  }
  function draw(){
    list.innerHTML='';
    const q=(search.value||'').toLowerCase();
    let rows=all.filter(matchesView).filter(x=>exchangeEventFilter==='all'||x.eventId===exchangeEventFilter);
    rows=rows.filter(x=>{
      const ev=events.find(e=>e.id===x.eventId);
      const give=(x.give||[]).map(g=>materialChoiceLabel(materials.find(m=>m.id===g.materialId),g.subItemId)).join(' ');
      const recv=(x.receiveItems||[]).map(r=>r.name||'').join(' ');
      return [x.partnerName,x.partnerWechat,ev?.name,x.place,x.note,give,recv].join(' ').toLowerCase().includes(q);
    });
    if(exchangeView==='onsite'){
      const currentEvent=events.find(e=>e.id===exchangeEventFilter);
      if(exchangeEventFilter==='all'){
        list.append(el('<div class="small-note">现场模式请选择一个具体场次。</div>'));
        return;
      }
      if(currentEvent){
        list.append(el(`<div class="small-note"><strong>当前现场：</strong>${escapeHtml(currentEvent.name)} · 卡片右上角优先显示集合地点。</div>`));
      }
      if(!rows.length){list.append(el('<div class="small-note">当前没有待交换记录。</div>'));return;}
      rows.forEach(x=>{
        const give=(x.give||[]).map(g=>`${escapeHtml(materialChoiceLabel(materials.find(m=>m.id===g.materialId),g.subItemId))} ×${g.qty}`).join('、');
        const recv=(x.receiveItems||[]).map(r=>`${escapeHtml(r.name||'对方物料')} ×${r.qty}`).join('、');
        list.append(el(`<div class="card onsite-card clickable" data-edit-exchange="${x.id}">
          <div class="section-head">
            <div class="people">${escapeHtml(x.partnerName||'未命名')}</div>
            <div class="location-pill">${escapeHtml(x.place||'未填地点')}</div>
          </div>
          ${x.partnerWechat?`<div class="onsite-meta">微信：${escapeHtml(x.partnerWechat)}</div>`:''}
          ${x.time?`<div class="onsite-meta">🕒 ${escapeHtml(x.time)}</div>`:''}
          ${x.note?`<div class="onsite-note">备注：${escapeHtml(x.note)}</div>`:''}
          <div class="exchange-visual">
            <div class="exchange-box"><strong>我给</strong>${give||'-'}</div>
            <div>⇄</div>
            <div class="exchange-box"><strong>她给</strong>${recv||'-'}</div>
          </div>
          <button class="primary" data-complete="${x.id}" style="margin-top:10px">✓ 完成交换</button>
        </div>`));
      });
    }else{
      rows.forEach(x=>{
        const ev=events.find(e=>e.id===x.eventId);
        list.append(el(`<div class="item clickable" data-edit-exchange="${x.id}">
          <div class="thumb">${x.receiveItems?.[0]?.images?.[0]?`<img src="${x.receiveItems[0].images[0]}">`:'互换'}</div>
          <div class="item-main">
            <div class="item-title">${escapeHtml(x.partnerName||'未命名')}</div>
            <div class="sub">${escapeHtml(ev?.name||'未绑定场次')} · ${escapeHtml(x.place||'未填地点')}</div>
            <div class="sub">${(x.receiveItems||[]).map(r=>escapeHtml(r.name)).join('、')||'未填写对方物料'}</div>
          </div>
          <div class="badge">${statusLabel(x.status)}</div>
        </div>`));
      });
    }
    list.querySelectorAll('[data-complete]').forEach(b=>b.addEventListener('click',async(e)=>{
      e.stopPropagation();
      const x=all.find(i=>i.id===b.dataset.complete); x.status='done'; x.completedAt=new Date().toISOString(); await put('exchanges',x); render();
    }));
  }
  document.querySelectorAll('[data-view]').forEach(b=>{
    b.classList.toggle('active',b.dataset.view===exchangeView);
    b.addEventListener('click',()=>{exchangeView=b.dataset.view;render()});
  });
  eventSelect.addEventListener('change',()=>{exchangeEventFilter=eventSelect.value;draw()});
  search.addEventListener('input',draw);
  draw();
}

async function renderSquats(view){
  setTitle('蹲蹲'); view.append(cloneTemplate('squats-template'));
  const [squats,events]=await Promise.all([getAll('squats'),getAll('events')]);
  const eventSelect=document.getElementById('squat-event-filter');
  eventSelect.innerHTML='<option value="all">全部场次</option>'+events.map(e=>`<option value="${e.id}">${escapeHtml(e.name)}</option>`).join('');
  if(squatEventFilter==='all' && homeEventId) squatEventFilter=homeEventId;
  eventSelect.value=squatEventFilter;
  const list=document.getElementById('squats-list');

  function draw(){
    list.innerHTML='';
    let rows=squats.filter(s=>squatFilter==='all'||s.status===squatFilter)
      .filter(s=>squatEventFilter==='all'||s.eventId===squatEventFilter)
      .sort((a,b)=>(a.time||'99:99').localeCompare(b.time||'99:99'));
    if(!rows.length){list.append(el('<div class="small-note">还没有蹲蹲记录。</div>'));return;}
    rows.forEach(s=>{
      const ev=events.find(e=>e.id===s.eventId);
      const freebies=(s.freebieImages||[]).slice(0,6).map(img=>`<img src="${img}" alt="无料图片" class="image-lightbox-trigger" data-preview-src="${img}">`).join('');
      list.append(el(`<div class="squat-card clickable" data-edit-squat="${s.id}">
        <div class="squat-top">
          <div>
            <div class="squat-name">${escapeHtml(s.teacherName||'未命名老师')}</div>
            <div class="sub">${s.xhs?`小红书：${escapeHtml(s.xhs)}`:'未填小红书'}${ev?` · ${escapeHtml(ev.name)}`:''}</div>
            <div class="squat-status-inline">状态：${squatStatusLabel(s.status)}${s.time?` · 🕒 ${escapeHtml(s.time)}`:''}</div>
          </div>
          <div class="squat-location">${escapeHtml(s.place||'未填地点')}</div>
        </div>
        <div class="squat-body-v231">
          <div class="freebie-main">
            <div class="label" style="margin-bottom:6px">无料</div>
            <div class="freebie-grid">${freebies||'<div class="sub">暂无图片</div>'}</div>
            <div class="small-note" style="margin-top:8px">${escapeHtml(s.note||'无备注')}</div>
          </div>
          ${s.ootdImage?`
          <div class="ootd-side">
            <div class="label" style="margin-bottom:6px">OOTD</div>
            <div class="ootd-box"><img src="${s.ootdImage}" alt="老师 OOTD" class="image-lightbox-trigger" data-preview-src="${s.ootdImage}"></div>
          </div>`:''}
        </div>
      </div>`));
    });
  }
  document.querySelectorAll('[data-squat-filter]').forEach(b=>{
    b.classList.toggle('active',b.dataset.squatFilter===squatFilter);
    b.addEventListener('click',()=>{squatFilter=b.dataset.squatFilter;draw()});
  });
  eventSelect.addEventListener('change',()=>{squatEventFilter=eventSelect.value;draw()});
  draw();
}

document.querySelectorAll('.tabbar button').forEach(b=>b.addEventListener('click',()=>{route=b.dataset.route;render()}));
document.addEventListener('click',e=>{
  const preview=e.target.closest('.image-lightbox-trigger');
  if(preview){
    e.preventDefault();
    e.stopPropagation();
    openImageLightbox(preview.dataset.previewSrc);
    return;
  }
  const dataAction=e.target.closest('[data-data-action]')?.dataset.dataAction;
  if(dataAction){
    if(dataAction==='export-excel'){exportExcel().catch(err=>dataMessage(err.message,'error'));return;}
    if(dataAction==='import-excel'){document.getElementById('excel-file-input')?.click();return;}
    if(dataAction==='export-backup'){exportBackup().catch(err=>dataMessage(err.message,'error'));return;}
    if(dataAction==='import-backup'){document.getElementById('backup-file-input')?.click();return;}
  }
  const shortcut=e.target.closest('[data-home-shortcut]')?.dataset.homeShortcut;
  if(shortcut){
    if(shortcut==='materials'){route='materials';render();return;}
    if(shortcut==='pending'){route='exchanges';exchangeView='pending';exchangeEventFilter=homeEventId||'all';render();return;}
    if(shortcut==='squats'){route='squats';squatFilter='want';squatEventFilter=homeEventId||'all';render();return;}
    if(shortcut==='events'){
      const target=document.querySelector('#home-events-list');
      if(target) target.scrollIntoView({behavior:'smooth',block:'start'});
      return;
    }
  }
  const action=e.target.closest('[data-action]')?.dataset.action;
  const jump=e.target.closest('[data-route-jump]')?.dataset.routeJump;
  const onsite=e.target.closest('[data-onsite]')?.dataset.onsite;
  const editMat=e.target.closest('[data-edit-material]')?.dataset.editMaterial;
  const editEx=e.target.closest('[data-edit-exchange]')?.dataset.editExchange;
  const editEvent=e.target.closest('[data-edit-event]')?.dataset.editEvent;
  const editSquat=e.target.closest('[data-edit-squat]')?.dataset.editSquat;
  if(jump){route=jump;if(onsite) exchangeView='onsite';render();return;}
  if(editMat){openModal('edit-material',editMat);return;}
  if(editEx){openModal('edit-exchange',editEx);return;}
  if(editEvent){openModal('edit-event',editEvent);return;}
  if(editSquat){openModal('edit-squat',editSquat);return;}
  if(action) openModal(action);
});

async function openModal(action,id=null){
  const modal=document.getElementById('modal'), body=document.getElementById('modal-body'), title=document.getElementById('modal-title');
  const [events,materials,exchanges,squats]=await Promise.all([getAll('events'),getAll('materials'),getAll('exchanges'),getAll('squats')]);
  body.innerHTML=''; let saveHandler=async()=>{};

  if(action==='new-event'||action==='edit-event'){
    const current=action==='edit-event'?await getOne('events',id):null;
    title.textContent=current?'编辑场次':'新建场次';
    body.innerHTML=`
      <div class="field"><label>场次名称</label><input class="input" name="name" value="${escapeHtml(current?.name||'')}" placeholder="例如：PLAVE 仁川 D1"></div>
      <div class="row"><div class="field"><label>日期</label><input class="input" name="date" type="date" value="${current?.date||''}"></div><div class="field"><label>城市</label><input class="input" name="city" value="${escapeHtml(current?.city||'')}"></div></div>
      <div class="field"><label>场馆</label><input class="input" name="venue" value="${escapeHtml(current?.venue||'')}"></div>
      <div class="field"><label>首页倒计时提示（可选）</label>
        <input class="input" name="countdownMessage" value="${escapeHtml(current?.countdownMessage||'')}" placeholder="例如：无料真的要开始滑铲了啊啊啊">
        <div class="small-note" style="margin-top:6px">留空时会继续使用系统根据剩余天数自动生成的提示。</div>
      </div>`;
    saveHandler=async()=>{
      const f=new FormData(document.getElementById('modal-form'));
      const obj=current||{id:uid('ev'),pinned:false};
      obj.name=f.get('name');obj.date=f.get('date');obj.city=f.get('city');obj.venue=f.get('venue');
      obj.countdownMessage=f.get('countdownMessage')||'';
      await put('events',obj);if(!homeEventId) homeEventId=obj.id;
    };
  }

  if(action==='new-material'||action==='edit-material'){
    const current=action==='edit-material'?await getOne('materials',id):null;
    title.textContent=current?'编辑物料':'新建物料';

    const subState=(current?.subItems||[]).map(s=>({
      id:s.id||uid('sub'),
      name:s.name||'',
      qty:Number(s.qty||0),
      image:s.image||'',
      newImage:''
    }));

    body.innerHTML=`
      <div class="field"><label>物料主图（可选）</label>
        <div class="image-preview" id="material-image-preview">${current?.image?`<img src="${current.image}">`:'<span>暂无主图</span>'}</div>
        <input class="input" name="imageFile" id="material-image-file" type="file" accept="image/*" style="margin-top:8px">
      </div>
      <div class="field"><label>物料名称</label><input class="input" name="name" value="${escapeHtml(current?.name||'')}" placeholder="例如：镭射票 / XXXX生日Set"></div>

      <div class="field">
        <label>子项（可选）</label>
        <div class="small-note">例如：镭射票 → 成员A / 成员B / 成员C；生日Set → 镭射票 / 吧唧。没有子项就保持为空。</div>
        <div id="subitems-editor" style="margin-top:8px"></div>
        <button type="button" class="editor-add" id="add-subitem">＋ 添加子项</button>
      </div>

      <div class="row">
        <div class="field"><label>总量（无子项时使用）</label><input class="input" name="qty" type="number" value="${current?.qty??1}" min="0"></div>
        <div class="field"><label>预留</label><input class="input" name="reserved" type="number" value="${current?.reserved??0}" min="0"></div>
      </div>
      <div class="field"><label>发放方式</label>
        <select name="distributionMethod">
          <option value="exchange" ${(current?.distributionMethod||'exchange')==='exchange'?'selected':''}>互换</option>
          <option value="handout" ${current?.distributionMethod==='handout'?'selected':''}>伸手</option>
          <option value="both" ${current?.distributionMethod==='both'?'selected':''}>互换 + 伸手</option>
          <option value="keep" ${current?.distributionMethod==='keep'?'selected':''}>仅自留 / 不发放</option>
        </select>
      </div>
      <div class="row">
        <div class="field"><label>伸手计划量</label><input class="input" name="handoutPlan" type="number" min="0" value="${current?.handoutPlan??0}"></div>
        <div class="field"><label>伸手已发</label><input class="input" name="handoutDone" type="number" min="0" value="${current?.handoutDone??0}"></div>
      </div>
      <div class="field"><label>库存模式</label><select name="inventoryMode"><option value="shared" ${current?.inventoryMode==='shared'?'selected':''}>共用总量</option><option value="allocated" ${current?.inventoryMode==='allocated'?'selected':''}>按场次分配</option></select></div>
      <div class="field"><label>绑定场次（可多选）</label>${events.map(ev=>`<label class="checkline"><input type="checkbox" name="eventIds" value="${ev.id}" ${(current?.eventIds||[]).includes(ev.id)?'checked':''}>${escapeHtml(ev.name)}</label>`).join('')}</div>
      <div class="field"><label>制作状态</label><select name="status">${['构思中','设计中','打样中','有成品'].map(s=>`<option ${current?.status===s?'selected':''}>${s}</option>`).join('')}</select></div>
      <div class="field"><label>备注</label><textarea name="note">${escapeHtml(current?.note||'')}</textarea></div>
      ${current?'<button type="button" class="danger" id="delete-material">删除物料</button>':''}`;

    const drawSubItems=()=>{
      const host=document.getElementById('subitems-editor');
      if(!host)return;
      host.innerHTML=subState.map((s,i)=>`
        <div class="subitem-editor">
          <div class="editor-head">
            <div class="editor-title">子项 ${i+1}</div>
            <button type="button" class="editor-remove" data-remove-sub="${i}">删除</button>
          </div>
          <div class="field"><label>名称</label><input class="input sub-name" data-index="${i}" value="${escapeHtml(s.name)}" placeholder="例如：成员A / 镭射票 / 吧唧"></div>
          <div class="field"><label>数量</label><input class="input sub-qty" data-index="${i}" type="number" min="0" value="${s.qty}"></div>
          <div class="field"><label>图片</label>
            <input class="input sub-image-file" data-index="${i}" type="file" accept="image/*">
            <div class="subitem-preview" data-preview-index="${i}">${(s.newImage||s.image)?`<img src="${s.newImage||s.image}">`:''}</div>
          </div>
        </div>
      `).join('');

      host.querySelectorAll('.sub-name').forEach(inp=>inp.addEventListener('input',()=>subState[Number(inp.dataset.index)].name=inp.value));
      host.querySelectorAll('.sub-qty').forEach(inp=>inp.addEventListener('input',()=>subState[Number(inp.dataset.index)].qty=Math.max(0,Number(inp.value||0))));
      host.querySelectorAll('.sub-image-file').forEach(inp=>inp.addEventListener('change',async()=>{
        const idx=Number(inp.dataset.index), file=inp.files?.[0];
        if(!file)return;
        subState[idx].newImage=await fileToDataURL(file);
        const prev=host.querySelector(`[data-preview-index="${idx}"]`);
        if(prev)prev.innerHTML=`<img src="${subState[idx].newImage}">`;
      }));
      host.querySelectorAll('[data-remove-sub]').forEach(btn=>btn.addEventListener('click',()=>{
        subState.splice(Number(btn.dataset.removeSub),1);
        drawSubItems();
      }));
    };

    saveHandler=async()=>{
      const f=new FormData(document.getElementById('modal-form'));
      const obj=current||{id:uid('mat'),image:'',allocations:{},exchangeable:true};
      const file=f.get('imageFile'); if(file&&file.size) obj.image=await fileToDataURL(file);
      obj.name=f.get('name');
      obj.qty=Number(f.get('qty')||0);
      obj.reserved=Number(f.get('reserved')||0);
      obj.subItems=subState
        .filter(s=>s.name.trim())
        .map(s=>({id:s.id||uid('sub'),name:s.name.trim(),qty:Math.max(0,Number(s.qty||0)),image:s.newImage||s.image||''}));
      obj.status=f.get('status');obj.inventoryMode=f.get('inventoryMode');obj.eventIds=f.getAll('eventIds');obj.note=f.get('note');
      obj.distributionMethod=f.get('distributionMethod')||'exchange';
      obj.handoutPlan=Math.max(0,Number(f.get('handoutPlan')||0));
      obj.handoutDone=Math.max(0,Number(f.get('handoutDone')||0));
      if(obj.handoutDone>obj.handoutPlan && obj.handoutPlan>0) obj.handoutDone=obj.handoutPlan;
      await put('materials',obj);
    };

    setTimeout(()=>{
      drawSubItems();
      const addBtn=document.getElementById('add-subitem');
      if(addBtn)addBtn.onclick=()=>{subState.push({id:uid('sub'),name:'',qty:0,image:'',newImage:''});drawSubItems();};

      const inp=document.getElementById('material-image-file');
      if(inp) inp.addEventListener('change',async()=>{
        const file=inp.files?.[0];if(!file)return;
        document.getElementById('material-image-preview').innerHTML=`<img src="${await fileToDataURL(file)}">`;
      });
      const delBtn=document.getElementById('delete-material');
      if(delBtn) delBtn.onclick=async()=>{
        const used=exchanges.some(x=>x.give?.some(g=>g.materialId===current.id));
        if(used){alert('这个物料已经关联互换记录，暂时不能直接删除。');return;}
        await del('materials',current.id);modal.close();render();
      };
    },0);
  }

  if(action==='new-exchange'||action==='edit-exchange'){
    const current=action==='edit-exchange'?await getOne('exchanges',id):null;
    title.textContent=current?'编辑互换':'新建互换';
    const receive=current?.receiveItems?.[0]||{name:'',images:[],qty:1};

    const choiceList=materialChoices(materials);
    const giveState=(current?.give?.length?current.give:[{materialId:materials[0]?.id||'',subItemId:null,qty:1}]).map(g=>({
      materialId:g.materialId||'',
      subItemId:g.subItemId||null,
      qty:Number(g.qty||1)
    }));

    body.innerHTML=`
      <div class="field"><label>对方微信昵称 / 备注</label><input class="input" name="partnerName" value="${escapeHtml(current?.partnerName||'')}" placeholder="例如：77"></div>
      <div class="field"><label>微信号（可选）</label><input class="input" name="partnerWechat" value="${escapeHtml(current?.partnerWechat||'')}"></div>
      <div class="field"><label>场次</label><select name="eventId">${events.map(ev=>`<option value="${ev.id}" ${current?.eventId===ev.id?'selected':''}>${escapeHtml(ev.name)}</option>`).join('')}</select></div>

      <div class="field">
        <label>我给她</label>
        <div class="small-note">可一次添加多个自己的无料；如果物料有子项，会显示为“父物料 · 子项”。</div>
        <div id="give-items-editor" style="margin-top:8px"></div>
        <button type="button" class="editor-add" id="add-give-item">＋ 添加我的物料</button>
      </div>

      <div class="field"><label>她给我：物料名称</label><input class="input" name="receiveName" value="${escapeHtml(receive.name||'')}" placeholder="例如：虎小卡"></div>
      <div class="field"><label>对方无料图片（可一次多选）</label>
        <input class="input" name="receiveImages" id="receive-images" type="file" accept="image/*" multiple>
        <div class="multi-preview" id="receive-preview" style="margin-top:8px">${(receive.images||[]).map(img=>`<img src="${img}">`).join('')}</div>
      </div>
      <div class="field"><label>对方给我数量</label><input class="input" name="receiveQty" type="number" min="1" value="${receive.qty||1}"></div>
      <div class="row"><div class="field"><label>交换地点</label><input class="input" name="place" value="${escapeHtml(current?.place||'')}" placeholder="例如：F6场内"></div><div class="field"><label>时间</label><input class="input" name="time" type="time" value="${current?.time||''}"></div></div>
      <div class="field"><label>现场备注</label><textarea name="note">${escapeHtml(current?.note||'')}</textarea></div>
      <div class="field"><label>状态</label><select name="status">
        ${[['booked','已约定'],['onsite','待交换'],['done','已完成'],['cancelled','已取消']].map(([v,l])=>`<option value="${v}" ${(current?.status||'onsite')===v?'selected':''}>${l}</option>`).join('')}
      </select></div>
      ${current?'<button type="button" class="danger" id="delete-exchange">删除互换</button>':''}`;

    const drawGiveItems=()=>{
      const host=document.getElementById('give-items-editor');
      if(!host)return;
      host.innerHTML=giveState.map((g,i)=>{
        const selectedValue=`${g.materialId||''}::${g.subItemId||''}`;
        return `
          <div class="give-item-editor">
            <div class="editor-head">
              <div class="editor-title">我的物料 ${i+1}</div>
              ${giveState.length>1?`<button type="button" class="editor-remove" data-remove-give="${i}">删除</button>`:''}
            </div>
            <div class="field"><label>选择物料 / 子项</label>
              <select class="give-choice" data-index="${i}">
                ${choiceList.map(c=>`<option value="${c.value}" ${c.value===selectedValue?'selected':''}>${escapeHtml(c.label)}</option>`).join('')}
              </select>
            </div>
            <div class="field"><label>数量</label><input class="input give-qty" data-index="${i}" type="number" min="1" value="${g.qty||1}"></div>
          </div>`;
      }).join('');

      host.querySelectorAll('.give-choice').forEach(sel=>sel.addEventListener('change',()=>{
        const idx=Number(sel.dataset.index);
        const [materialId,subItemId]=sel.value.split('::');
        giveState[idx].materialId=materialId;
        giveState[idx].subItemId=subItemId||null;
      }));
      host.querySelectorAll('.give-qty').forEach(inp=>inp.addEventListener('input',()=>{
        giveState[Number(inp.dataset.index)].qty=Math.max(1,Number(inp.value||1));
      }));
      host.querySelectorAll('[data-remove-give]').forEach(btn=>btn.addEventListener('click',()=>{
        giveState.splice(Number(btn.dataset.removeGive),1);
        drawGiveItems();
      }));
    };

    saveHandler=async()=>{
      const f=new FormData(document.getElementById('modal-form'));
      const obj=current||{id:uid('x'),createdAt:new Date().toISOString(),completedAt:null};
      const files=document.getElementById('receive-images')?.files;
      const newImgs=files?.length?await filesToDataURLs(files):[];
      const oldImgs=current?.receiveItems?.[0]?.images||[];

      obj.partnerName=f.get('partnerName');
      obj.partnerWechat=f.get('partnerWechat');
      obj.eventId=f.get('eventId');
      obj.give=giveState
        .filter(g=>g.materialId)
        .map(g=>({materialId:g.materialId,subItemId:g.subItemId||null,qty:Math.max(1,Number(g.qty||1))}));
      obj.receiveItems=[{name:f.get('receiveName'),images:newImgs.length?newImgs:oldImgs,qty:Number(f.get('receiveQty')||1)}];
      obj.place=f.get('place');obj.time=f.get('time');obj.note=f.get('note');obj.status=f.get('status');
      if(obj.status==='done'&&!obj.completedAt)obj.completedAt=new Date().toISOString();
      if(obj.status!=='done')obj.completedAt=null;
      await put('exchanges',obj);
    };

    setTimeout(()=>{
      drawGiveItems();
      const addBtn=document.getElementById('add-give-item');
      if(addBtn)addBtn.onclick=()=>{
        const first=choiceList[0]?.value||'::';
        const [materialId,subItemId]=first.split('::');
        giveState.push({materialId,subItemId:subItemId||null,qty:1});
        drawGiveItems();
      };

      const inp=document.getElementById('receive-images');
      if(inp) inp.addEventListener('change',async()=>{
        const urls=await filesToDataURLs(inp.files);
        document.getElementById('receive-preview').innerHTML=urls.map(u=>`<img src="${u}">`).join('');
      });
      const delBtn=document.getElementById('delete-exchange');
      if(delBtn) delBtn.onclick=async()=>{await del('exchanges',current.id);modal.close();render();};
    },0);
  }

  if(action==='new-squat'||action==='edit-squat'){
    const current=action==='edit-squat'?await getOne('squats',id):null;
    title.textContent=current?'编辑蹲蹲':'新建蹲蹲';
    body.innerHTML=`
      <div class="field"><label>老师昵称 / 称呼</label><input class="input" name="teacherName" value="${escapeHtml(current?.teacherName||'')}" placeholder="例如：77老师"></div>
      <div class="field"><label>小红书账号</label><input class="input" name="xhs" value="${escapeHtml(current?.xhs||'')}" placeholder="例如：@xxxxxx"></div>
      <div class="field"><label>场次</label><select name="eventId">${events.map(ev=>`<option value="${ev.id}" ${current?.eventId===ev.id?'selected':''}>${escapeHtml(ev.name)}</option>`).join('')}</select></div>
      <div class="field"><label>老师的无料图片（可多选）</label>
        <input class="input" name="freebieFiles" id="freebie-files" type="file" accept="image/*" multiple>
        <div class="multi-preview" id="freebie-preview" style="margin-top:8px">${(current?.freebieImages||[]).map(img=>`<img src="${img}">`).join('')}</div>
      </div>
      <div class="field"><label>老师 OOTD（可选）</label>
        <div class="image-preview" id="ootd-preview">${current?.ootdImage?`<img src="${current.ootdImage}">`:'<span>未上传则卡片中不显示</span>'}</div>
        <input class="input" name="ootdFile" id="ootd-file" type="file" accept="image/*" style="margin-top:8px">
      </div>
      <div class="row"><div class="field"><label>发放地点</label><input class="input" name="place" value="${escapeHtml(current?.place||'')}" placeholder="例如：F6场内"></div><div class="field"><label>发放时间</label><input class="input" name="time" type="time" value="${current?.time||''}"></div></div>
      <div class="field"><label>数量 / 领取条件</label><input class="input" name="condition" value="${escapeHtml(current?.condition||'')}" placeholder="例如：限100份 / 关注小红书"></div>
      <div class="field"><label>备注</label><textarea name="note">${escapeHtml(current?.note||'')}</textarea></div>
      <div class="field"><label>状态</label><select name="status">
        ${[['want','想蹲'],['got','已拿到'],['missed','没拿到']].map(([v,l])=>`<option value="${v}" ${(current?.status||'want')===v?'selected':''}>${l}</option>`).join('')}
      </select></div>
      ${current?'<button type="button" class="danger" id="delete-squat">删除蹲蹲</button>':''}`;
    saveHandler=async()=>{
      const f=new FormData(document.getElementById('modal-form'));
      const obj=current||{id:uid('sq'),createdAt:new Date().toISOString(),ootdImage:'',freebieImages:[]};
      const ootd=f.get('ootdFile');if(ootd&&ootd.size)obj.ootdImage=await fileToDataURL(ootd);
      const freebieFiles=document.getElementById('freebie-files')?.files;
      if(freebieFiles?.length)obj.freebieImages=await filesToDataURLs(freebieFiles);
      obj.teacherName=f.get('teacherName');obj.xhs=f.get('xhs');obj.eventId=f.get('eventId');
      obj.place=f.get('place');obj.time=f.get('time');obj.condition=f.get('condition');obj.note=f.get('note');obj.status=f.get('status');
      await put('squats',obj);
    };
    setTimeout(()=>{
      const ootd=document.getElementById('ootd-file');
      if(ootd)ootd.addEventListener('change',async()=>{const f=ootd.files?.[0];if(f)document.getElementById('ootd-preview').innerHTML=`<img src="${await fileToDataURL(f)}">`;});
      const freebies=document.getElementById('freebie-files');
      if(freebies)freebies.addEventListener('change',async()=>{const urls=await filesToDataURLs(freebies.files);document.getElementById('freebie-preview').innerHTML=urls.map(u=>`<img src="${u}">`).join('');});
      const delBtn=document.getElementById('delete-squat');
      if(delBtn)delBtn.onclick=async()=>{await del('squats',current.id);modal.close();render();};
    },0);
  }

  modal.showModal();
  const form=document.getElementById('modal-form');
  form.onsubmit=async(e)=>{
    if(e.submitter?.value==='cancel')return;
    e.preventDefault();await saveHandler();modal.close();await render();
  };
}

window.addEventListener('online',()=>document.getElementById('offline-status').textContent='在线');
window.addEventListener('offline',()=>document.getElementById('offline-status').textContent='离线可用');
window.addEventListener('keydown',e=>{
  if(e.key==='Escape'){
    const box=document.getElementById('image-lightbox');
    if(box) box.classList.remove('open');
  }
});

(async()=>{
  if('caches' in window){
    const keys=await caches.keys();
    await Promise.all(keys.filter(k=>k!=='xingyu-v2-5-0').map(k=>caches.delete(k)));
  }
  db=await openDB();await render();
  if('serviceWorker' in navigator)navigator.serviceWorker.register('./sw.js?v=2.5.0').catch(()=>{});
})();
