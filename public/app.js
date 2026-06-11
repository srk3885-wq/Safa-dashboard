/* ================= data model ================= */
const LS_KEY='safa_dashboard_v1';
const STATUSES={
  production:{label:'In Production / Pending',cls:'st-production'},
  sea:{label:'In Transit — Sea',cls:'st-sea'},
  air:{label:'In Transit — Air',cls:'st-air'},
  port:{label:'Arrived at Port',cls:'st-port'},
  truck:{label:'In Truck / Delivery',cls:'st-truck'},
  delivered:{label:'Delivered',cls:'st-delivered'},
  delayed:{label:'Delayed / Hold',cls:'st-delayed'},
};
const INBOUND=['production','sea','air','port','truck','delayed']; // counted as inbound
let state={shipments:[],inventory:[],lowStockThreshold:50,asOf:null,fileName:null};
let ui={view:'overview',q:'',mode:'',status:'',brand:'',expanded:{},invSort:'low'};

function load(){try{const s=JSON.parse(localStorage.getItem(LS_KEY));if(s&&s.shipments){state=s;if(!state.inventory)state.inventory=[];if(state.lowStockThreshold==null)state.lowStockThreshold=50;}}catch(e){}}
function normKey(s){return String(s??'').toUpperCase().replace(/[^A-Z0-9]/g,'');}
function save(){localStorage.setItem(LS_KEY,JSON.stringify(state));}
function uid(){return Date.now().toString(36)+Math.random().toString(36).slice(2,7);}
function toast(msg){const t=document.getElementById('toast');t.textContent=msg;t.classList.add('show');clearTimeout(t._h);t._h=setTimeout(()=>t.classList.remove('show'),2600);}
function fmt(n){return (n==null||isNaN(n))?'—':Number(n).toLocaleString();}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}

/* ================= excel parsing ================= */
const HEADER_ALIASES={
  mode:['shipmenttype','mode','shipmode','transport','shippingtype','seaorair'],
  date:['pickupdate','date','pickup','etd','shipdate','departuredate'],
  brand:['brand','brandname'],
  product:['producttype','product','spu','item','model','productname'],
  flavors:['flavors','flavor','flavours','flavour','skus','skudetails','flavorsctns'],
  count:['boxcount','cases','cartons','qty','quantity','totalcases','count','boxes','totalboxcount','ctns'],
  cargo:['cargostatus','status','remarks','remark','statusupdate','eta','shipmentstatus'],
};
function normHeader(h){return String(h??'').toLowerCase().replace(/[^a-z]/g,'');}
function mapHeaders(row){
  const map={};
  row.forEach((cell,i)=>{
    const n=normHeader(cell); if(!n)return;
    for(const[field,aliases]of Object.entries(HEADER_ALIASES)){
      if(map[field]===undefined && aliases.some(a=>n===a||n.startsWith(a)||a.startsWith(n)&&n.length>=4)){map[field]=i;return;}
    }
  });
  return map;
}
function parseFlavorLines(text){
  if(!text) return [];
  return String(text).split(/\r?\n|;/).map(l=>l.trim()).filter(Boolean).map(line=>{
    const m=line.match(/^(.*?)[\s\-–:]*([\d.,]+)\s*(?:ctns?|cartons?|cases?|boxes?|bx|pcs)?\.?\s*$/i);
    if(m && m[2] && /\d/.test(m[2]) && m[1].trim()){
      const qty=parseFloat(m[2].replace(/,/g,''));
      if(!isNaN(qty)) return {name:m[1].trim(),cartons:qty};
    }
    return {name:line,cartons:null};
  });
}
function deriveStatus(raw,mode){
  const t=String(raw??'').toUpperCase();
  if(/DELIVERED|RECEIVED AT WAREHOUSE|COMPLETED/.test(t))return 'delivered';
  if(/IN TRUCK|TRUCKING|OUT FOR DELIVERY|ON TRUCK|LAST MILE/.test(t))return 'truck';
  if(/DELAY|ON HOLD|HOLD AT|EXAM|INSPECTION/.test(t))return 'delayed';
  if(/ARRIV|AT PORT|PORT ARRIVAL|LANDED|DISCHARGED|CUSTOMS|CLEARANCE/.test(t))return 'port';
  if(/FLY|FLIE|FLIGHT|AIRPORT|AWB/.test(t))return 'air';
  if(/VESSEL|SAIL|ON BOARD|OCEAN/.test(t))return 'sea';
  if(/DEPART|ETD|ETA|TRANSIT|SHIPPED|ON THE WAY|LOADED/.test(t))return mode==='AIR'?'air':'sea';
  return 'production';
}
function normMode(v,cargo){
  const t=String(v??'').toUpperCase();
  const hasSea=/SEA|OCEAN|VESSEL/.test(t), hasAir=/AIR/.test(t);
  if(hasSea&&hasAir){ // ambiguous placeholder like "SEA or AIR" — infer from cargo text
    const c=String(cargo??'').toUpperCase();
    if(/FLY|FLIE|FLIGHT|AWB|AIRPORT/.test(c))return 'AIR';
    return 'SEA';
  }
  if(hasAir)return 'AIR';
  if(hasSea)return 'SEA';
  return t.trim()?'OTHER':'SEA';
}
function parseDate(v){
  if(v==null||v==='')return null;
  if(v instanceof Date && !isNaN(v))return v.toISOString().slice(0,10);
  if(typeof v==='number'){const d=XLSX.SSF.parse_date_code(v);if(d)return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;}
  const s=String(v).trim();
  let m=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if(m){let y=+m[3];if(y<100)y+=2000;return `${y}-${String(+m[1]).padStart(2,'0')}-${String(+m[2]).padStart(2,'0')}`;}
  const d=new Date(s);return isNaN(d)?null:d.toISOString().slice(0,10);
}
function fingerprint(s){return [s.brand,s.productType,s.pickupDate,s.mode,s.boxCount].map(x=>String(x??'').toUpperCase().trim()).join('|');}

/* --- inventory file: "Brand - Product [pack] | Flavor" + Current Stock --- */
const INV_ALIASES={
  name:['productname','product','itemname','item','sku','description','productdescription','name'],
  stock:['currentstock','stock','onhand','qtyonhand','stockqty','inventory','available','balance','quantity','qty','units'],
};
function mapInvHeaders(row){
  const map={};
  row.forEach((cell,i)=>{
    const n=normHeader(cell); if(!n)return;
    if(map.name===undefined && INV_ALIASES.name.some(a=>n===a||n.startsWith(a)))map.name=i;
    else if(map.stock===undefined && INV_ALIASES.stock.some(a=>n===a||n.startsWith(a)))map.stock=i;
  });
  return map;
}
function parseProductName(s){
  s=String(s??'').trim();
  let brand='—',product='—',flavor=s;
  const pi=s.indexOf('|');
  if(pi>=0){
    flavor=s.slice(pi+1).trim()||'—';
    const left=s.slice(0,pi).trim();
    const di=left.indexOf(' - ');
    if(di>=0){brand=left.slice(0,di).trim();product=left.slice(di+3).trim();}
    else product=left||'—';
  }else{
    const di=s.indexOf(' - ');
    if(di>=0){brand=s.slice(0,di).trim();flavor=s.slice(di+3).trim();}
  }
  return {brand:brand||'—',product:product||'—',flavor:flavor||'—'};
}

function importWorkbook(wb,fileName){
  const shipRows=[],invRows=[];
  wb.SheetNames.forEach(name=>{
    const aoa=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,raw:false,defval:''});
    if(!aoa.length)return;
    // classify sheet: shipment header (>=3 mapped fields) wins, else inventory header (name+stock)
    let sIdx=-1,sMap=null,vIdx=-1,vMap=null;
    for(let i=0;i<Math.min(aoa.length,10);i++){
      if(sIdx<0){const m=mapHeaders(aoa[i]);if(Object.keys(m).length>=3){sIdx=i;sMap=m;}}
      if(vIdx<0){const m=mapInvHeaders(aoa[i]);if(m.name!==undefined&&m.stock!==undefined){vIdx=i;vMap=m;}}
    }
    if(sIdx>=0){
      const aoaRaw=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,raw:true,defval:''});
      for(let i=sIdx+1;i<aoaRaw.length;i++){
        const r=aoaRaw[i];
        const get=f=>sMap[f]===undefined?'':r[sMap[f]];
        const brand=String(get('brand')).trim(), product=String(get('product')).trim();
        const flavors=parseFlavorLines(get('flavors'));
        let count=parseFloat(String(get('count')).replace(/[^\d.]/g,''));
        const flavSum=flavors.reduce((a,f)=>a+(f.cartons||0),0);
        if(isNaN(count)||count===0)count=flavSum||null;
        if(!brand&&!product&&!flavors.length&&!count)continue;
        if(/total/i.test(brand)||/total/i.test(product))continue;
        const cargo=String(get('cargo')).trim();
        const mode=normMode(get('mode'),cargo);
        shipRows.push({
          id:uid(),mode,pickupDate:parseDate(get('date')),brand:brand||'—',productType:product||'—',
          flavors,boxCount:count,cargoStatus:cargo,status:deriveStatus(cargo,mode),
          notes:'',manual:false,statusManual:false,
        });
      }
    }else if(vIdx>=0){
      for(let i=vIdx+1;i<aoa.length;i++){
        const r=aoa[i];
        const nameV=String(r[vMap.name]??'').trim();
        if(!nameV||/^(grand\s*)?total/i.test(nameV))continue;
        const stock=parseFloat(String(r[vMap.stock]??'').replace(/[^\d.-]/g,''));
        const p=parseProductName(nameV);
        invRows.push({id:uid(),...p,stock:isNaN(stock)?0:stock,manual:false});
      }
    }
  });
  if(!shipRows.length&&!invRows.length){toast('No shipment or inventory rows found — check the file headers');return;}
  const parts=[];
  if(shipRows.length){
    const oldByFp={};
    state.shipments.filter(s=>!s.manual).forEach(s=>oldByFp[fingerprint(s)]=s);
    shipRows.forEach(r=>{
      const old=oldByFp[fingerprint(r)];
      if(old){r.notes=old.notes;if(old.statusManual){r.status=old.status;r.statusManual=true;}}
    });
    state.shipments=[...shipRows,...state.shipments.filter(s=>s.manual)];
    parts.push(`${shipRows.length} shipment${shipRows.length>1?'s':''}`);
  }
  if(invRows.length){
    const keys=new Set(invRows.map(i=>normKey(i.brand+'|'+i.product+'|'+i.flavor)));
    const manualKept=state.inventory.filter(i=>i.manual&&!keys.has(normKey(i.brand+'|'+i.product+'|'+i.flavor)));
    state.inventory=[...invRows,...manualKept];
    parts.push(`${invRows.length} inventory SKU${invRows.length>1?'s':''}`);
    if(!shipRows.length)ui.view='inventory',syncTabs();
  }
  state.asOf=new Date().toISOString();state.fileName=fileName;
  save();render();
  toast(`Imported ${parts.join(' + ')} from ${fileName}`);
}
function handleFile(file){
  const reader=new FileReader();
  reader.onload=e=>{
    try{importWorkbook(XLSX.read(new Uint8Array(e.target.result),{type:'array',cellDates:true}),file.name);}
    catch(err){toast('Could not read file: '+err.message);}
  };
  reader.readAsArrayBuffer(file);
}

/* ================= filtering & aggregation ================= */
function filtered(){
  return state.shipments.filter(s=>{
    if(ui.mode&&s.mode!==ui.mode)return false;
    if(ui.status&&s.status!==ui.status)return false;
    if(ui.brand&&s.brand!==ui.brand)return false;
    if(ui.q){
      const hay=[s.brand,s.productType,s.cargoStatus,s.notes,STATUSES[s.status]?.label,...s.flavors.map(f=>f.name)].join(' ').toLowerCase();
      if(!hay.includes(ui.q.toLowerCase()))return false;
    }
    return true;
  });
}
function shipCases(s){return (s.boxCount ?? s.flavors.reduce((a,f)=>a+(f.cartons||0),0)) || 0;}

/* ================= rendering ================= */
function render(){
  renderHeader();renderStats();renderFilters();
  document.getElementById('cntShip').textContent=state.shipments.length;
  const skuCount=new Set(state.shipments.flatMap(s=>s.flavors.map(f=>`${s.brand}|${s.productType}|${f.name}`))).size;
  document.getElementById('cntSku').textContent=skuCount;
  document.getElementById('cntInv').textContent=state.inventory.length;
  const v=document.getElementById('view');
  if(ui.view==='inventory'){v.innerHTML=renderInventory();return;}
  if(ui.view==='breakdown'){v.innerHTML=renderBreakdown();return;}
  const list=filtered();
  if(!state.shipments.length){v.innerHTML=`<div class="tablecard"><div class="empty"><b>No shipments yet</b>Upload the Excel file or add a manual entry to get started.</div></div>`;return;}
  if(ui.view==='overview')v.innerHTML=renderOverview(list);
  else if(ui.view==='shipments')v.innerHTML=renderShipments(list);
  else v.innerHTML=renderSku(list);
}
function syncTabs(){document.querySelectorAll('.tab').forEach(t=>t.classList.toggle('active',t.dataset.view===ui.view));}
function renderHeader(){
  const a=document.getElementById('asof');
  a.textContent=state.asOf
    ?`As of ${new Date(state.asOf).toLocaleString(undefined,{month:'short',day:'numeric',year:'numeric',hour:'numeric',minute:'2-digit'})}${state.fileName?' · '+state.fileName:''}`
    :'No data loaded yet';
}
function renderStats(){
  const inb=state.shipments.filter(s=>INBOUND.includes(s.status));
  const sum=arr=>arr.reduce((a,s)=>a+shipCases(s),0);
  const seaT=inb.filter(s=>s.status==='sea'),airT=inb.filter(s=>s.status==='air');
  const ground=inb.filter(s=>['port','truck'].includes(s.status));
  const pend=inb.filter(s=>['production','delayed'].includes(s.status));
  document.getElementById('stats').innerHTML=`
    <div class="stat dark"><div class="lbl">Total cases inbound</div><div class="num">${fmt(sum(inb))}</div><div class="sub">${inb.length} active shipment${inb.length!==1?'s':''}</div></div>
    <div class="stat"><div class="lbl"><span class="dot" style="background:var(--sea)"></span>Sea freight · in transit</div><div class="num">${fmt(sum(seaT))}</div><div class="sub">${seaT.length} shipment${seaT.length!==1?'s':''} on the water</div></div>
    <div class="stat"><div class="lbl"><span class="dot" style="background:var(--air)"></span>Air freight · in transit</div><div class="num">${fmt(sum(airT))}</div><div class="sub">${airT.length} shipment${airT.length!==1?'s':''} flying</div></div>
    <div class="stat"><div class="lbl"><span class="dot" style="background:var(--truck)"></span>Arrived / in truck</div><div class="num">${fmt(sum(ground))}</div><div class="sub">${ground.length} at port or on delivery</div></div>
    <div class="stat"><div class="lbl"><span class="dot" style="background:var(--gold)"></span>Pending / pre-departure</div><div class="num">${fmt(sum(pend))}</div><div class="sub">${pend.length} not yet moving</div></div>
    ${state.inventory.length?(()=>{
      const thr=state.lowStockThreshold;
      const total=state.inventory.reduce((a,i)=>a+(i.stock||0),0);
      const low=state.inventory.filter(i=>i.stock>0&&i.stock<=thr).length;
      const out=state.inventory.filter(i=>!i.stock).length;
      return `<div class="stat"><div class="lbl"><span class="dot" style="background:#6941c6"></span>Warehouse on-hand</div><div class="num">${fmt(total)}</div><div class="sub">${state.inventory.length} SKUs · ${low} low · ${out} out</div></div>`;
    })():''}`;
}
function renderFilters(){
  const fs=document.getElementById('fStatus');
  if(fs.options.length<=1)Object.entries(STATUSES).forEach(([k,v])=>fs.add(new Option(v.label,k)));
  const brands=[...new Set([...state.shipments.map(s=>s.brand),...state.inventory.map(i=>i.brand)])].sort();
  document.getElementById('brandChips').innerHTML=brands.map(b=>
    `<button class="chip ${ui.brand===b?'active':''}" onclick="setBrand('${esc(b)}')">${esc(b)}</button>`).join('');
}
function setBrand(b){ui.brand=ui.brand===b?'':b;render();}

function renderOverview(list){
  // group brand -> product -> flavor
  const groups={};
  list.filter(s=>INBOUND.includes(s.status)).forEach(s=>{
    const k=s.brand+'|'+s.productType;
    if(!groups[k])groups[k]={brand:s.brand,product:s.productType,flavors:{},total:0,modes:new Set()};
    const g=groups[k];g.total+=shipCases(s);g.modes.add(s.mode);
    s.flavors.forEach(f=>{g.flavors[f.name]=(g.flavors[f.name]||0)+(f.cartons||0);});
  });
  const cards=Object.values(groups).sort((a,b)=>b.total-a.total);
  if(!cards.length)return `<div class="tablecard"><div class="empty"><b>Nothing matches the current filters</b>Try clearing the search or filters.</div></div>`;
  return `<div class="sec">Product breakdown — inbound cases by brand / SPU / flavor</div><div class="grid">`+cards.map(g=>{
    const fl=Object.entries(g.flavors).sort((a,b)=>b[1]-a[1]);
    const max=Math.max(...fl.map(f=>f[1]),1);
    return `<div class="pcard">
      <div class="pbrand">${esc(g.brand)} · ${[...g.modes].join(' + ')}</div>
      <div class="pname">${esc(g.product)}</div>
      ${fl.map(([n,q])=>`<div class="frow"><div class="ftop"><span class="fname" title="${esc(n)}">${esc(n)}</span><span class="fqty">${fmt(q)}</span></div><div class="bar"><i style="width:${Math.max(q/max*100,2)}%"></i></div></div>`).join('')||'<div class="frow" style="color:var(--mut);font-size:12.5px">No flavor detail provided</div>'}
      <div class="ptotal"><b>${fmt(g.total)}</b> <span>cases total ${esc(g.product)}</span></div>
    </div>`;
  }).join('')+`</div>`;
}

/* ================= breakdown (graphs) ================= */
function donutSVG(segs,centerNum,centerLbl){
  const total=segs.reduce((a,s)=>a+s.value,0)||1;
  const R=42,C=2*Math.PI*R;
  let off=0;
  const circles=segs.filter(s=>s.value>0).map(s=>{
    const len=s.value/total*C;
    const el=`<circle cx="55" cy="55" r="${R}" fill="none" stroke="${s.color}" stroke-width="13" stroke-dasharray="${len} ${C-len}" stroke-dashoffset="${-off}" transform="rotate(-90 55 55)"/>`;
    off+=len;return el;
  }).join('');
  return `<svg width="110" height="110" viewBox="0 0 110 110">
    <circle cx="55" cy="55" r="${R}" fill="none" stroke="var(--line2)" stroke-width="13"/>${circles}
    <text x="55" y="52" text-anchor="middle" class="donut-center" fill="var(--ink)">${centerNum}</text>
    <text x="55" y="68" text-anchor="middle" style="font-size:9.5px;font-weight:600;fill:var(--mut);letter-spacing:.05em">${centerLbl}</text>
  </svg>`;
}
function donutCard(title,segs,centerNum,centerLbl){
  const total=segs.reduce((a,s)=>a+s.value,0);
  return `<div class="bd-card"><h4>${title}</h4><div class="donut-flex">
    ${donutSVG(segs,centerNum,centerLbl)}
    <div class="leg">${segs.map(s=>`<div class="li"><span class="dot2" style="background:${s.color}"></span>${esc(s.label)} <span class="lv">${fmt(s.value)}${total?' · '+Math.round(s.value/total*100)+'%':''}</span></div>`).join('')}</div>
  </div></div>`;
}
function barCard(title,obj,unit,limit){
  const entries=Object.entries(obj).sort((a,b)=>b[1]-a[1]);
  const total=entries.reduce((a,e)=>a+e[1],0)||1;
  const max=Math.max(...entries.map(e=>e[1]),1);
  const shown=entries.slice(0,limit);
  return `<div class="bd-card"><h4>${title}</h4>
    ${shown.map(([k,v])=>`<div class="brow">
      <div class="btop"><span class="bname" title="${esc(k)}">${esc(k)}</span>
      <span class="bval">${fmt(v)}<span class="bpct">${Math.round(v/total*100)}%</span></span></div>
      <div class="bar"><i style="width:${Math.max(v/max*100,1.5)}%"></i></div>
    </div>`).join('')||'<div class="bd-more">No data</div>'}
    ${entries.length>limit?`<div class="bd-more">+ ${entries.length-limit} more — use search or brand filter to narrow down</div>`:''}
  </div>`;
}
function renderBreakdown(){
  if(!state.shipments.length&&!state.inventory.length)
    return `<div class="tablecard"><div class="empty"><b>No data yet</b>Upload a shipment or inventory file to see breakdowns.</div></div>`;
  const src=ui.bdSrc||(state.shipments.length?'inbound':'stock');
  const add=(o,k,v)=>{o[k]=(o[k]||0)+v;};
  const byBrand={},bySpu={},bySku={};
  let donuts='',unitLbl='cases';
  if(src==='stock'){
    unitLbl='units';
    const thr=state.lowStockThreshold;
    let ok=0,low=0,out=0,total=0;
    state.inventory.filter(i=>{
      if(ui.brand&&i.brand!==ui.brand)return false;
      if(ui.q){const hay=[i.brand,i.product,i.flavor].join(' ').toLowerCase();if(!hay.includes(ui.q.toLowerCase()))return false;}
      return true;
    }).forEach(i=>{
      const v=i.stock||0;total+=v;
      add(byBrand,i.brand,v);
      add(bySpu,i.brand+' · '+i.product,v);
      add(bySku,i.brand+' · '+i.flavor,v);
      if(!i.stock)out++;else if(i.stock<=thr)low++;else ok++;
    });
    donuts=donutCard('SKU health — stock alerts',[
      {label:'OK',value:ok,color:'var(--truck)'},
      {label:'Low stock',value:low,color:'var(--gold)'},
      {label:'Out of stock',value:out,color:'var(--warn)'},
    ],fmt(ok+low+out),'SKUS');
  }else{
    const list=filtered().filter(s=>INBOUND.includes(s.status));
    const byMode={},byStatus={};
    list.forEach(s=>{
      const v=shipCases(s);
      add(byBrand,s.brand,v);
      add(bySpu,s.brand+' · '+s.productType,v);
      add(byMode,s.mode,v);
      add(byStatus,s.status,v);
      s.flavors.forEach(f=>add(bySku,s.brand+' · '+f.name,f.cartons||0));
    });
    const totalCases=Object.values(byBrand).reduce((a,v)=>a+v,0);
    const MODE_COLORS={SEA:'var(--sea)',AIR:'var(--air)',OTHER:'var(--mut)'};
    const ST_COLORS={production:'var(--gold)',sea:'var(--sea)',air:'var(--air)',port:'#175cd3',truck:'var(--truck)',delayed:'var(--warn)'};
    donuts=
      donutCard('By shipment mode',Object.entries(byMode).map(([k,v])=>({label:k,value:v,color:MODE_COLORS[k]||'var(--mut)'})),fmt(totalCases),'CASES')+
      donutCard('By status',Object.entries(byStatus).map(([k,v])=>({label:STATUSES[k].label,value:v,color:ST_COLORS[k]||'var(--mut)'})),fmt(totalCases),'CASES');
  }
  return `
  <div class="bd-toolbar">
    <span class="lbl2">Data</span>
    <button class="chip ${src==='inbound'?'active':''}" onclick="ui.bdSrc='inbound';render()">Inbound shipments</button>
    <button class="chip ${src==='stock'?'active':''}" onclick="ui.bdSrc='stock';render()" ${state.inventory.length?'':'disabled style="opacity:.4;cursor:not-allowed"'}>Warehouse stock</button>
  </div>
  <div class="bd-grid">${donuts}</div>
  <div class="bd-grid">
    ${barCard('By brand — '+unitLbl,byBrand,unitLbl,12)}
    ${barCard('By product (SPU) — '+unitLbl,bySpu,unitLbl,12)}
  </div>
  <div class="bd-grid" style="grid-template-columns:1fr">
    ${barCard('By flavor (SKU) — top 20 by '+unitLbl,bySku,unitLbl,20)}
  </div>`;
}

function renderShipments(list){
  if(!list.length)return `<div class="tablecard"><div class="empty"><b>Nothing matches the current filters</b>Try clearing the search or filters.</div></div>`;
  const rows=list.map(s=>{
    const flavSum=s.flavors.reduce((a,f)=>a+(f.cartons||0),0);
    const mismatch=s.boxCount!=null&&flavSum>0&&flavSum!==s.boxCount;
    const exp=ui.expanded[s.id];
    return `
    <tr class="mainrow">
      <td><span class="modechip ${s.mode}">${s.mode}</span></td>
      <td style="white-space:nowrap">${s.pickupDate?new Date(s.pickupDate+'T00:00').toLocaleDateString(undefined,{month:'short',day:'numeric'}):'—'}</td>
      <td><b>${esc(s.brand)}</b>${s.manual?'<span class="manualtag">MANUAL</span>':''}</td>
      <td>${esc(s.productType)}</td>
      <td><span class="exp" onclick="toggleExp('${s.id}')">${exp?'▾':'▸'} ${s.flavors.length} flavor${s.flavors.length!==1?'s':''}</span>${mismatch?'<div class="warnflag">⚠ flavors sum to '+fmt(flavSum)+'</div>':''}</td>
      <td style="text-align:right;font-weight:700;font-variant-numeric:tabular-nums">${fmt(shipCases(s))}</td>
      <td><select class="statussel ${STATUSES[s.status].cls}" onchange="setStatus('${s.id}',this.value)">
        ${Object.entries(STATUSES).map(([k,v])=>`<option value="${k}" ${k===s.status?'selected':''}>${v.label}</option>`).join('')}
      </select></td>
      <td class="cargotxt">${esc(s.cargoStatus)||'—'}${s.notes?`<div class="notebadge" onclick="toggleExp('${s.id}')">📝 note</div>`:''}</td>
      <td><div class="rowbtns">
        <button class="iconbtn" title="Edit" onclick="openEdit('${s.id}')">✎</button>
        <button class="iconbtn" title="Note" onclick="toggleExp('${s.id}')">📝</button>
        <button class="iconbtn" title="Delete" onclick="delShip('${s.id}')">🗑</button>
      </div></td>
    </tr>
    ${exp?`<tr class="subrow"><td colspan="9">
      <table class="flavtable">${s.flavors.map(f=>`<tr><td>${esc(f.name)}</td><td>${f.cartons!=null?fmt(f.cartons)+' ctns':'—'}</td></tr>`).join('')||'<tr><td style="color:var(--mut)">No flavor lines</td></tr>'}</table>
      <div style="margin-top:10px;max-width:480px"><textarea class="notearea" placeholder="Add a note (customs hold, partial delivery…)" onblur="setNote('${s.id}',this.value)">${esc(s.notes)}</textarea></div>
    </td></tr>`:''}`;
  }).join('');
  return `<div class="tablecard"><table>
    <thead><tr><th>Mode</th><th>Pick up</th><th>Brand</th><th>Product (SPU)</th><th>Flavors</th><th style="text-align:right">Cases</th><th>Status</th><th>Cargo status</th><th></th></tr></thead>
    <tbody>${rows}</tbody></table></div>`;
}

function renderSku(list){
  const agg={};
  list.filter(s=>INBOUND.includes(s.status)).forEach(s=>{
    s.flavors.forEach(f=>{
      const k=`${s.brand}|${s.productType}|${f.name}`;
      if(!agg[k])agg[k]={brand:s.brand,product:s.productType,flavor:f.name,cartons:0,ships:0};
      agg[k].cartons+=f.cartons||0;agg[k].ships++;
    });
  });
  const rows=Object.values(agg).sort((a,b)=>b.cartons-a.cartons);
  if(!rows.length)return `<div class="tablecard"><div class="empty"><b>No SKU data</b>Flavors appear here once shipments are loaded.</div></div>`;
  const max=Math.max(...rows.map(r=>r.cartons),1);
  return `<div class="sec">Every SKU inbound — searchable, sorted by volume</div><div class="tablecard"><table>
    <thead><tr><th>Brand</th><th>Product (SPU)</th><th>Flavor (SKU)</th><th class="skuqty">Cartons</th><th>Shipments</th><th class="skubarcell"></th></tr></thead>
    <tbody>${rows.map(r=>`<tr>
      <td><b>${esc(r.brand)}</b></td><td>${esc(r.product)}</td><td>${esc(r.flavor)}</td>
      <td class="skuqty">${fmt(r.cartons)}</td><td style="color:var(--mut)">${r.ships}</td>
      <td class="skubarcell"><div class="bar"><i style="width:${Math.max(r.cartons/max*100,2)}%"></i></div></td>
    </tr>`).join('')}</tbody></table></div>`;
}

function renderInventory(){
  if(!state.inventory.length)return `<div class="tablecard"><div class="empty"><b>No inventory loaded</b>Upload the stock Excel (Product Name / Current Stock) or add items manually.</div></div>`;
  const thr=state.lowStockThreshold;
  const inbMap={};
  state.shipments.filter(s=>INBOUND.includes(s.status)).forEach(s=>s.flavors.forEach(f=>{
    const k=normKey(s.brand)+'|'+normKey(f.name);
    inbMap[k]=(inbMap[k]||0)+(f.cartons||0);
  }));
  let items=state.inventory.filter(i=>{
    if(ui.brand&&i.brand!==ui.brand)return false;
    if(ui.q){const hay=[i.brand,i.product,i.flavor].join(' ').toLowerCase();if(!hay.includes(ui.q.toLowerCase()))return false;}
    return true;
  });
  const sort=ui.invSort;
  if(sort==='low')items.sort((a,b)=>(a.stock||0)-(b.stock||0));
  else if(sort==='high')items.sort((a,b)=>(b.stock||0)-(a.stock||0));
  else items.sort((a,b)=>(a.brand+a.product+a.flavor).localeCompare(b.brand+b.product+b.flavor));
  const total=state.inventory.reduce((a,i)=>a+(i.stock||0),0);
  const low=state.inventory.filter(i=>i.stock>0&&i.stock<=thr).length;
  const out=state.inventory.filter(i=>!i.stock).length;
  const max=Math.max(...state.inventory.map(i=>i.stock||0),1);
  const rows=items.map(i=>{
    const inb=inbMap[normKey(i.brand)+'|'+normKey(i.flavor)];
    const flag=!i.stock?'<span class="flag out">OUT</span>':i.stock<=thr?'<span class="flag low">LOW</span>':'<span class="flag ok">OK</span>';
    return `<tr>
      <td><b>${esc(i.brand)}</b>${i.manual?'<span class="manualtag">MANUAL</span>':''}</td>
      <td>${esc(i.product)}</td><td>${esc(i.flavor)}</td>
      <td style="text-align:right"><input type="number" class="stockedit" value="${i.stock??0}" min="0" onchange="setStock('${i.id}',this.value)"></td>
      <td class="skuqty" style="color:var(--mut);font-weight:400" title="Inbound cartons matched by brand + flavor">${inb?'+'+fmt(inb):'—'}</td>
      <td>${flag}</td>
      <td class="skubarcell"><div class="bar"><i style="width:${Math.max((i.stock||0)/max*100,1)}%"></i></div></td>
      <td><button class="iconbtn" title="Delete" onclick="delInv('${i.id}')">🗑</button></td>
    </tr>`;
  }).join('');
  return `
  <div class="invbar">
    <span><b style="color:var(--ink);font-size:14px">${fmt(total)}</b> units on hand · ${state.inventory.length} SKUs</span>
    <span class="flag low">${low} LOW</span><span class="flag out">${out} OUT</span>
    <span style="margin-left:auto">Low-stock alert ≤ <input type="number" min="0" value="${thr}" onchange="setThreshold(this.value)"></span>
    <select class="flt" onchange="ui.invSort=this.value;render()">
      <option value="low" ${sort==='low'?'selected':''}>Stock low → high</option>
      <option value="high" ${sort==='high'?'selected':''}>Stock high → low</option>
      <option value="name" ${sort==='name'?'selected':''}>Name A–Z</option>
    </select>
    <button class="btn small" onclick="invModal.showModal()">+ Add item</button>
  </div>
  ${items.length?`<div class="tablecard"><table>
    <thead><tr><th>Brand</th><th>Product (SPU)</th><th>Flavor (SKU)</th><th class="skuqty">In stock</th><th class="skuqty">Inbound</th><th></th><th class="skubarcell"></th><th></th></tr></thead>
    <tbody>${rows}</tbody></table></div>`
  :`<div class="tablecard"><div class="empty"><b>Nothing matches the current filters</b>Try clearing the search or brand filter.</div></div>`}`;
}

/* ================= row actions ================= */
function setStock(id,val){const i=state.inventory.find(x=>x.id===id);if(i){i.stock=Math.max(0,parseFloat(val)||0);save();render();}}
function setThreshold(val){state.lowStockThreshold=Math.max(0,parseInt(val)||0);save();render();}
function delInv(id){const i=state.inventory.find(x=>x.id===id);if(i&&confirm(`Delete ${i.brand} ${i.product} | ${i.flavor}?`)){state.inventory=state.inventory.filter(x=>x.id!==id);save();render();}}
function toggleExp(id){ui.expanded[id]=!ui.expanded[id];render();}
function setStatus(id,val){const s=state.shipments.find(x=>x.id===id);if(s){s.status=val;s.statusManual=true;save();render();}}
function setNote(id,val){const s=state.shipments.find(x=>x.id===id);if(s&&s.notes!==val.trim()){s.notes=val.trim();save();render();}}
function delShip(id){const s=state.shipments.find(x=>x.id===id);if(s&&confirm(`Delete ${s.brand} ${s.productType} (${fmt(shipCases(s))} cases)?`)){state.shipments=state.shipments.filter(x=>x.id!==id);save();render();}}

/* ================= modal ================= */
const modal=document.getElementById('modal');
let editId=null;
function fillSelStatus(){const sel=document.getElementById('inStatus');sel.innerHTML='';Object.entries(STATUSES).forEach(([k,v])=>sel.add(new Option(v.label,k)));}
function openAdd(){
  editId=null;fillSelStatus();
  document.getElementById('mtitle').textContent='Manual entry — new shipment';
  shipForm.reset();document.getElementById('inStatus').value='production';
  fillDatalists();updateFlavCalc();modal.showModal();
}
function openEdit(id){
  const s=state.shipments.find(x=>x.id===id);if(!s)return;
  editId=id;fillSelStatus();
  document.getElementById('mtitle').textContent='Edit shipment';
  inMode.value=s.mode;inDate.value=s.pickupDate||'';inBrand.value=s.brand;inProduct.value=s.productType;
  inFlavors.value=s.flavors.map(f=>f.cartons!=null?`${f.name} ${f.cartons}ctns`:f.name).join('\n');
  inCount.value=s.boxCount??'';inStatus.value=s.status;inCargo.value=s.cargoStatus;inNotes.value=s.notes;
  fillDatalists();updateFlavCalc();modal.showModal();
}
function fillDatalists(){
  document.getElementById('brandList').innerHTML=[...new Set(state.shipments.map(s=>s.brand))].map(b=>`<option value="${esc(b)}">`).join('');
  document.getElementById('productList').innerHTML=[...new Set(state.shipments.map(s=>s.productType))].map(p=>`<option value="${esc(p)}">`).join('');
}
function updateFlavCalc(){
  const fl=parseFlavorLines(inFlavors.value);
  const sum=fl.reduce((a,f)=>a+(f.cartons||0),0);
  document.getElementById('flavCalc').textContent=fl.length?`${fl.length} flavor line${fl.length>1?'s':''} · ${fmt(sum)} cartons detected`:'';
}
document.getElementById('inFlavors').addEventListener('input',updateFlavCalc);
document.getElementById('shipForm').addEventListener('submit',e=>{
  const flavors=parseFlavorLines(inFlavors.value);
  let count=inCount.value!==''?parseFloat(inCount.value):NaN;
  if(isNaN(count))count=flavors.reduce((a,f)=>a+(f.cartons||0),0)||null;
  const data={
    mode:inMode.value,pickupDate:inDate.value||null,brand:inBrand.value.trim()||'—',
    productType:inProduct.value.trim()||'—',flavors,boxCount:count,
    cargoStatus:inCargo.value.trim(),status:inStatus.value,notes:inNotes.value.trim(),
  };
  if(editId){
    const s=state.shipments.find(x=>x.id===editId);
    Object.assign(s,data,{statusManual:true});
    toast('Shipment updated');
  }else{
    state.shipments.unshift({id:uid(),manual:true,statusManual:true,...data});
    if(!state.asOf)state.asOf=new Date().toISOString();
    toast('Manual shipment added');
  }
  save();render();
});

document.getElementById('invForm').addEventListener('submit',()=>{
  state.inventory.unshift({id:uid(),brand:invBrand.value.trim()||'—',product:invProduct.value.trim()||'—',
    flavor:invFlavor.value.trim()||'—',stock:Math.max(0,parseFloat(invStock.value)||0),manual:true});
  if(!state.asOf)state.asOf=new Date().toISOString();
  document.getElementById('invForm').reset();
  save();render();toast('Inventory item added');
});

/* ================= export / backup ================= */
function exportExcel(){
  if(!state.shipments.length&&!state.inventory.length){toast('Nothing to export yet');return;}
  const header=['Shipment Type','Pick Up Date','Brand','Product Type','Flavors','Box Count','Cargo Status','Dashboard Status','Notes'];
  const rows=state.shipments.map(s=>[
    s.mode,s.pickupDate||'',s.brand,s.productType,
    s.flavors.map(f=>f.cartons!=null?`${f.name} ${f.cartons}ctns`:f.name).join('\n'),
    shipCases(s),s.cargoStatus,STATUSES[s.status].label,s.notes,
  ]);
  const skuHeader=['Brand','Product (SPU)','Flavor (SKU)','Cartons','Shipments'];
  const agg={};
  state.shipments.filter(s=>INBOUND.includes(s.status)).forEach(s=>s.flavors.forEach(f=>{
    const k=`${s.brand}|${s.productType}|${f.name}`;
    if(!agg[k])agg[k]=[s.brand,s.productType,f.name,0,0];
    agg[k][3]+=f.cartons||0;agg[k][4]++;
  }));
  const wb=XLSX.utils.book_new();
  const ws1=XLSX.utils.aoa_to_sheet([header,...rows]);
  ws1['!cols']=[{wch:10},{wch:11},{wch:10},{wch:16},{wch:42},{wch:10},{wch:34},{wch:22},{wch:24}];
  XLSX.utils.book_append_sheet(wb,ws1,'Shipments');
  const ws2=XLSX.utils.aoa_to_sheet([skuHeader,...Object.values(agg).sort((a,b)=>b[3]-a[3])]);
  ws2['!cols']=[{wch:10},{wch:16},{wch:34},{wch:10},{wch:10}];
  XLSX.utils.book_append_sheet(wb,ws2,'SKU Summary');
  if(state.inventory.length){
    const thr=state.lowStockThreshold;
    const invRows=[...state.inventory].sort((a,b)=>(a.stock||0)-(b.stock||0))
      .map(i=>[i.brand,i.product,i.flavor,i.stock||0,!i.stock?'OUT':i.stock<=thr?'LOW':'OK']);
    const ws3=XLSX.utils.aoa_to_sheet([['Brand','Product (SPU)','Flavor (SKU)','Current Stock','Alert'],...invRows]);
    ws3['!cols']=[{wch:10},{wch:28},{wch:26},{wch:13},{wch:8}];
    XLSX.utils.book_append_sheet(wb,ws3,'Inventory');
  }
  XLSX.writeFile(wb,`SAFA_Dashboard_${new Date().toISOString().slice(0,10)}.xlsx`);
}
function backup(){
  const blob=new Blob([JSON.stringify(state,null,2)],{type:'application/json'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download=`safa_dashboard_backup_${new Date().toISOString().slice(0,10)}.json`;a.click();
  toast('Backup downloaded');
}
function restore(file){
  const r=new FileReader();
  r.onload=e=>{try{const s=JSON.parse(e.target.result);if(!s.shipments)throw 0;state=s;save();render();toast('Backup restored');}catch{toast('Not a valid backup file');}};
  r.readAsText(file);
}

/* ================= wiring ================= */
document.getElementById('btnUpload').onclick=()=>fileInput.click();
document.getElementById('dropzone').onclick=()=>fileInput.click();
fileInput.onchange=e=>{if(e.target.files[0])handleFile(e.target.files[0]);e.target.value='';};
document.getElementById('btnAdd').onclick=openAdd;
document.getElementById('btnExport').onclick=exportExcel;
document.getElementById('btnBackup').onclick=backup;
document.getElementById('btnRestore').onclick=()=>restoreInput.click();
restoreInput.onchange=e=>{if(e.target.files[0])restore(e.target.files[0]);e.target.value='';};

['dragenter','dragover'].forEach(ev=>document.addEventListener(ev,e=>{e.preventDefault();document.body.classList.add('dragging');}));
['dragleave','drop'].forEach(ev=>document.addEventListener(ev,e=>{e.preventDefault();if(ev==='drop'||e.target===document.documentElement||!e.relatedTarget)document.body.classList.remove('dragging');}));
document.addEventListener('drop',e=>{const f=e.dataTransfer?.files?.[0];if(f)handleFile(f);});

document.querySelectorAll('.tab').forEach(t=>t.onclick=()=>{
  document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));
  t.classList.add('active');ui.view=t.dataset.view;render();
});
document.getElementById('q').addEventListener('input',e=>{ui.q=e.target.value;render();});
document.getElementById('fMode').onchange=e=>{ui.mode=e.target.value;render();};
document.getElementById('fStatus').onchange=e=>{ui.status=e.target.value;render();};

load();render();
