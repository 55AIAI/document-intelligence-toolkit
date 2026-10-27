"""
Интерактивный редактор схемы (HTML+JS) — для веб-платформы.

Строит схему из МОДЕЛИ ДАННЫХ (дерево заказчик → исполнители → субподряд) и
даёт универсальный редактор диаграммы:
  • блоки: организация, договор, фото (ОТДЕЛЬНЫЙ блок, место зарезервировано),
    текст, произвольная фигура;
  • организация — это СПИСОК строк: любую строку можно править, удалить (✕) или
    добавить (＋ строка); заголовок «Руководитель» тоже редактируемый;
  • перетаскивание за шапку; изменение размера с ЛЮБОЙ стороны (8 маркеров);
  • форматирование выделенного текста: жирный, крупнее/мельче;
  • добавление блоков, рисование связей вручную, удаление блока/связи, Delete;
  • «Очистить всё»; вставка фото; экспорт PNG / JPEG / SVG; JSON-состояние.

Связи хранятся явно (edges) — стрелка идёт куда указана и пересчитывается при
перемещении. build_html(data) → самодостаточный HTML (DATA встроена).
Компонент переносится в веб-платформу: данные через DATA, состояние в DATA._editor.
"""

import json

_TEMPLATE = r"""<!doctype html>
<html lang="ru"><head><meta charset="utf-8">
<title>Редактор схемы связей</title>
<style>
  :root{--cust:#2f5597;--con:#548235;--sub:#bf8f00;--con2:#c00000;}
  *{box-sizing:border-box;font-family:Arial,sans-serif;}
  body{margin:0;background:#f4f5f7;}
  #bar{position:sticky;top:0;z-index:50;background:#fff;border-bottom:1px solid #ddd;
       padding:8px 10px;display:flex;gap:6px;align-items:center;flex-wrap:wrap;}
  #bar b{margin-right:6px;} #bar .sep{width:1px;height:22px;background:#ddd;margin:0 4px;}
  button{padding:6px 10px;border:1px solid #cbd3dc;background:#fff;border-radius:6px;
         cursor:pointer;font-size:13px;} button:hover{background:#eef2f7;}
  button.on{background:#2f5597;color:#fff;border-color:#2f5597;}
  #hint{color:#777;font-size:12px;margin-left:4px;}
  #canvas{position:relative;width:100%;height:calc(100vh - 48px);overflow:hidden;background:#fafbfc;cursor:grab;}
  #canvas.linking{cursor:crosshair;}
  #canvas.panning{cursor:grabbing;}
  #stage{position:absolute;left:0;top:0;width:1px;height:1px;transform-origin:0 0;}
  #wire{position:absolute;left:0;top:0;}
  .blk{position:absolute;background:#fff;border:2px solid #999;border-radius:7px;
       box-shadow:0 1px 5px rgba(0,0,0,.13);display:flex;flex-direction:column;min-width:150px;overflow:hidden;}
  .blk.sel{outline:3px solid #4c9aff;outline-offset:1px;}
  .blk.src{outline:3px dashed #ff8800;outline-offset:1px;}
  .blk.cust{border-color:var(--cust);} .blk.con{border-color:var(--con);}
  .blk.sub{border-color:var(--sub);} .blk.contract{border-color:var(--con2);}
  .blk.photo{border-color:#666;} .blk.text{border-color:#888;} .blk.shape{border-color:#777;}
  .hdr{color:#fff;font-size:13px;padding:5px 9px;cursor:move;display:flex;
       justify-content:space-between;align-items:center;user-select:none;flex:none;}
  .cust .hdr{background:var(--cust);} .con .hdr{background:var(--con);}
  .sub .hdr{background:var(--sub);} .contract .hdr{background:var(--con2);}
  .photo .hdr{background:#666;} .text .hdr{background:#888;} .shape .hdr{background:#777;}
  .hdr .ttl{outline:none;flex:1;} .hdr .x{cursor:pointer;font-weight:bold;padding:0 4px;opacity:.85;}
  .hdr .clr{width:20px;height:18px;border:none;background:none;cursor:pointer;padding:0;flex:none;}
  .body{padding:6px 9px;}
  .line{display:flex;align-items:flex-start;gap:4px;margin-top:2px;}
  .line .t{flex:1;outline:none;font-size:14px;line-height:1.3;overflow-wrap:anywhere;}
  .line.ln-name .t{font-weight:bold;font-size:16px;}
  .line.ln-sub .t{font-weight:bold;font-size:14px;border-top:1px dashed #ddd;padding-top:5px;margin-top:3px;}
  .line.ln-viol{background:#fff3f3;border-left:3px solid #c00;padding:2px 5px;border-radius:2px;}
  .line.ln-viol .t{color:#c00;}
  .line .lx{cursor:pointer;color:#c33;font-size:12px;opacity:0;padding:0 2px;flex:none;}
  .line:hover .lx{opacity:.7;}
  .addln{font-size:12px;color:#2f5597;cursor:pointer;margin-top:5px;user-select:none;}
  .ctext{padding:10px;font-size:14px;text-align:center;outline:none;flex:1;
         display:flex;flex-direction:column;justify-content:center;overflow-wrap:anywhere;}
  .ctext .amt{font-weight:bold;margin-top:4px;}
  .ttext{padding:9px;font-size:14px;outline:none;flex:1;overflow-wrap:anywhere;}
  .photobox{flex:1;background:#f2f2f2;display:flex;align-items:center;justify-content:center;
       color:#999;font-size:12px;cursor:pointer;background-size:cover;background-position:center;
       text-align:center;min-height:120px;}
  /* маркеры изменения размера со всех сторон */
  .h{position:absolute;background:transparent;} .blk:hover .h{background:rgba(76,154,255,.25);}
  .h.n{top:-3px;left:8px;right:8px;height:7px;cursor:ns-resize;}
  .h.s{bottom:-3px;left:8px;right:8px;height:7px;cursor:ns-resize;}
  .h.e{right:-3px;top:8px;bottom:8px;width:7px;cursor:ew-resize;}
  .h.w{left:-3px;top:8px;bottom:8px;width:7px;cursor:ew-resize;}
  .h.ne{top:-4px;right:-4px;width:12px;height:12px;cursor:nesw-resize;}
  .h.nw{top:-4px;left:-4px;width:12px;height:12px;cursor:nwse-resize;}
  .h.se{bottom:-4px;right:-4px;width:14px;height:14px;cursor:nwse-resize;
        background:linear-gradient(135deg,transparent 50%,#aaa 50%);}
  .h.sw{bottom:-4px;left:-4px;width:12px;height:12px;cursor:nesw-resize;}
  #search{padding:5px 8px;border:1px solid #cbd3dc;border-radius:6px;font-size:13px;width:160px;}
  .blk.hit{outline:3px solid #ffd54a !important;}
  .blk.aff{box-shadow:0 0 0 3px rgba(192,0,0,.45),0 1px 5px rgba(0,0,0,.13);}
  #panel{position:fixed;right:0;top:48px;width:330px;max-height:calc(100vh - 48px);overflow:auto;
    background:#fff;border-left:1px solid #ddd;padding:12px 14px;font-size:13px;z-index:60;
    box-shadow:-2px 0 10px rgba(0,0,0,.12);display:none;}
  #panel h3{margin:12px 0 4px;font-size:13px;text-transform:uppercase;color:#555;letter-spacing:.3px;}
  #panel .f{margin:3px 0;padding:5px 7px;border-radius:4px;background:#f6f7f9;line-height:1.35;}
  #panel .risk{font-size:26px;font-weight:bold;} .bad{color:#c00;} .ok{color:#2e7d32;} .warn{color:#b26a00;}
  #panel .close{float:right;cursor:pointer;color:#888;font-weight:bold;}
</style></head>
<body>
<div id="bar">
  <b>Редактор</b>
  <button onclick="fmtBold()"><b>Ж</b></button>
  <button onclick="fmtSize(1)">A+</button>
  <button onclick="fmtSize(-1)">A−</button>
  <span class="sep"></span>
  <button onclick="addOrg('sub')">➕ Организация</button>
  <button onclick="addContract()">➕ Договор</button>
  <button onclick="addPhoto()">➕ Фото</button>
  <button onclick="addText()">➕ Текст</button>
  <button onclick="addShape()">➕ Фигура</button>
  <button id="linkBtn" onclick="toggleLink()">🔗 Связь</button>
  <button onclick="delSel()">🗑 Удалить</button>
  <button onclick="clearAll()">🧹 Очистить всё</button>
  <span class="sep"></span>
  <button onclick="exportRaster('image/png')">🖼 PNG</button>
  <button onclick="exportRaster('image/jpeg')">🖼 JPEG</button>
  <button onclick="exportSVG()">⬇ SVG</button>
  <button onclick="save()">💾 JSON</button>
  <button onclick="location.reload()">↺ Сброс</button>
  <span class="sep"></span>
  <button onclick="setZoom(0.83)" title="уменьшить">🔍−</button>
  <button onclick="setZoom(1.2)" title="увеличить">🔍+</button>
  <button onclick="spread(1.25)" title="раздвинуть блоки">↔ Разложить</button>
  <button onclick="spread(0.8)" title="сжать блоки">↝ Сжать</button>
  <span class="sep"></span>
  <input id="search" placeholder="поиск: БИН/ИИН/ФИО" oninput="doSearch(this.value)">
  <button onclick="analyze()">🔎 Проверка схемы</button>
  <span id="hint"></span>
</div>
<div id="canvas"><div id="stage"><svg id="wire"></svg></div></div>
<div id="panel"><span class="close" onclick="closePanel()">✕</span><div id="pbody"></div></div>

<script>
const DATA = /*DATA*/;
const ROLE_RU={cust:'Заказчик',con:'Исполнитель',sub:'Субподрядчик',contract:'Договор',photo:'Фото',text:'Заметка',shape:'Фигура',person:'Лицо',violation:'Нарушение'};
const COL={cust:'#2f5597',con:'#548235',sub:'#bf8f00',contract:'#c00000',photo:'#666',text:'#888',shape:'#777',person:'#7b5ea7',violation:'#c00000'};
const canvas=document.getElementById('canvas'),stage=document.getElementById('stage'),wire=document.getElementById('wire'),hint=document.getElementById('hint');
let blocks=[],edges=[],seq=1,selected=null,linkMode=false,linkFrom=null,curEdit=null,aff=[];
const uid=()=>'b'+(seq++);
const esc=s=>(s==null?'':String(s)).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const setHint=t=>hint.textContent=t;
setHint('Тяни за шапку · тяни за любой край — размер · строки: ✕ удалить, ＋ добавить · выдели текст → Ж/A± · «Связь» → два блока');
document.addEventListener('focusin',e=>{if(e.target.isContentEditable)curEdit=e.target;});
document.addEventListener('keydown',e=>{if((e.key==='Delete')&&selected&&!(document.activeElement&&document.activeElement.isContentEditable))removeBlock(selected);});

// ── начальное состояние ──
function init(){
  if(DATA._editor){blocks=DATA._editor.blocks;edges=DATA._editor.edges;seq=blocks.length+edges.length+9;renderAll();return;}
  blocks=[];edges=[];
  const COLW=600, ROWH=500;                 // больше пространства по горизонтали и вертикали
  const cust=DATA.customer||{name:'Заказчик'};
  const leaf={n:0};
  function walk(node,role,depth){
    const kids=(node===cust?(DATA.contractors||[]):(node.subcontractors||[]));
    const childBlks=kids.map(k=>walk(k,'sub',depth+1));
    let x; if(!childBlks.length){x=40+leaf.n*COLW;leaf.n++;}
    else x=(childBlks[0]._cx+childBlks[childBlks.length-1]._cx)/2;
    const y=40+depth*ROWH;
    const b=orgBlock(node,role,x,y);b._cx=x;
    photoBlock(x+360,y+34);
    childBlks.forEach(cb=>{const cd=cb._data,ct=cd.contract||{amount:cd.amount,description:cd.description,year:cd.year};
      if(ct.amount||ct.description){const con=contractBlock(ct,cb.x+10,cb.y-190);   // договор выше субподрядчика, с зазором
        edges.push({id:uid(),from:b.id,to:con.id});edges.push({id:uid(),from:con.id,to:cb.id});}
      else edges.push({id:uid(),from:b.id,to:cb.id});});
    return b;
  }
  walk(cust,'cust',0);renderAll();
}
function orgLines(n){const h=n.header||n.director||{};const L=[];
  L.push({k:'name',t:n.name||'—'});
  L.push({k:'fld',t:'БИН '+(n.bin||'')});
  L.push({k:'fld',t:'Виды деятельности: '+(n.activity||'')});
  L.push({k:'fld',t:'Юр.адрес: '+(n.legal_address||n.address||'')});
  (n.violations||[]).forEach(v=>L.push({k:'viol',t:'⚠ '+v}));
  L.push({k:'sub',t:'Руководитель'});
  L.push({k:'fld',t:'ФИО: '+(h.fio||'')});
  L.push({k:'fld',t:'ИИН: '+(h.iin||'')});
  L.push({k:'fld',t:'Тел.номер: '+(h.phone||'')});
  L.push({k:'fld',t:'Адрес прописки: '+(h.address||'')});
  return L;}
function orgBlock(n,role,x,y){const b={id:uid(),type:'org',role,x,y,w:340,_data:n,lines:orgLines(n)};blocks.push(b);return b;}
function contractBlock(ct,x,y){const b={id:uid(),type:'contract',role:'contract',x,y,w:300,fields:{desc:ct.description||'Предмет договора',amt:ct.amount||''}};blocks.push(b);return b;}
function photoBlock(x,y){const b={id:uid(),type:'photo',role:'photo',x,y,w:160,h:210,photo:null};blocks.push(b);return b;}

// ── добавление вручную ──
function spawnXY(){const r=canvas.getBoundingClientRect();return [Math.round((r.width/2-panX)/zoom),Math.round((r.height/2-panY)/zoom)];}
function addOrg(role){const [x,y]=spawnXY();blocks.push({id:uid(),type:'org',role,x,y,w:340,
  lines:[{k:'name',t:'Новая организация'},{k:'fld',t:'БИН '},{k:'sub',t:'Руководитель'},{k:'fld',t:'ФИО: '}]});renderAll();}
function addContract(){const [x,y]=spawnXY();blocks.push({id:uid(),type:'contract',role:'contract',x,y,w:300,fields:{desc:'Предмет договора',amt:'сумма'}});renderAll();}
function addPhoto(){const [x,y]=spawnXY();blocks.push({id:uid(),type:'photo',role:'photo',x,y,w:160,h:210,photo:null});renderAll();}
function addText(){const [x,y]=spawnXY();blocks.push({id:uid(),type:'text',role:'text',x,y,w:230,minh:100,fields:{text:'Текст…'}});renderAll();}
function addShape(){const [x,y]=spawnXY();blocks.push({id:uid(),type:'shape',role:'shape',x,y,w:200,minh:120,title:'Фигура',fields:{text:''}});renderAll();}

// ── рендер ──
const HANDLES=['n','s','e','w','ne','nw','se','sw'];
function renderAll(){[...canvas.querySelectorAll('.blk')].forEach(e=>e.remove());blocks.forEach(renderBlock);redraw();}
function renderBlock(b){
  const color=/^#[0-9a-f]{6}$/i.test(b.color||'')?b.color:(COL[b.role]||'#777777');
  const title=(b.title!=null?b.title:(ROLE_RU[b.role]||''));
  if(b.photo && !/^data:image\/(png|jpeg|gif|webp);base64,[a-z0-9+/=]+$/i.test(b.photo)) b.photo=null;
  const el=document.createElement('div');el.className='blk '+b.role;el.dataset.id=b.id;
  el.style.left=b.x+'px';el.style.top=b.y+'px';if(b.w)el.style.width=b.w+'px';
  if(b.type==='photo'&&b.h)el.style.height=b.h+'px';
  if(b.minh)el.style.minHeight=b.minh+'px';
  el.style.borderColor=color;
  let inner=`<div class="hdr" style="background:${color}"><span class="ttl" contenteditable>${esc(title)}</span>`+
            `<input type="color" class="clr" value="${color}" title="цвет блока"><span class="x" title="удалить блок">✕</span></div>`;
  if(b.type==='org'){
    inner+=`<div class="body">`+ (b.lines||[]).map((l,i)=>
      `<div class="line ln-${esc(l.k)}" data-k="${esc(l.k)}"><span class="t" contenteditable>${esc(l.t)}</span><span class="lx" data-i="${i}" title="удалить строку">✕</span></div>`).join('')+
      `<div class="addln">＋ строка</div></div>`;
  }else if(b.type==='contract'){
    inner+=`<div class="ctext"><div contenteditable data-k="desc">${esc(b.fields.desc)}</div><div class="amt" contenteditable data-k="amt">${esc(b.fields.amt)}</div></div>`;
  }else if(b.type==='photo'){
    inner+=`<div class="photobox" style="${b.photo?`background-image:url(${b.photo})`:''}">${b.photo?'':'ФОТО<br>(клик — выбрать)'}</div>`;
  }else if(b.type==='text'||b.type==='shape'){
    inner+=`<div class="ttext" contenteditable data-k="text">${esc(b.fields.text)}</div>`;
  }
  inner+=HANDLES.map(d=>`<div class="h ${d}" data-dir="${d}"></div>`).join('');
  el.innerHTML=inner;stage.appendChild(el);

  el.addEventListener('mousedown',e=>onBlockDown(e,b));
  el.addEventListener('click',()=>{if(linkMode)clickLink(b.id);});
  const hdr=el.querySelector('.hdr');
  el.querySelector('.ttl').addEventListener('input',ev=>{b.title=ev.target.textContent;});
  el.querySelector('.clr').addEventListener('input',ev=>{b.color=ev.target.value;el.style.borderColor=b.color;hdr.style.background=b.color;});
  el.querySelector('.x').addEventListener('click',e=>{e.stopPropagation();removeBlock(b.id);});
  el.querySelectorAll('.h').forEach(h=>h.addEventListener('mousedown',e=>{e.stopPropagation();startDrag(e,b,'size',h.dataset.dir);}));
  const pb=el.querySelector('.photobox');if(pb)pb.addEventListener('click',e=>{if(!linkMode){e.stopPropagation();pickPhoto(b);}});
  el.querySelectorAll('[contenteditable]').forEach(n=>n.addEventListener('input',()=>syncText(b,el)));
  el.querySelectorAll('.lx').forEach(x=>x.addEventListener('click',e=>{e.stopPropagation();
    b.lines.splice(+x.dataset.i,1);renderAll();select(b.id);}));
  const add=el.querySelector('.addln');if(add)add.addEventListener('click',e=>{e.stopPropagation();
    b.lines.push({k:'fld',t:''});renderAll();select(b.id);});
}
function syncText(b,el){
  if(b.type==='org'){b.lines=[...el.querySelectorAll('.line')].map(d=>({k:d.dataset.k,t:d.querySelector('.t').textContent}));}
  else if(b.type==='contract'){b.fields.desc=txt(el,'[data-k=desc]');b.fields.amt=txt(el,'[data-k=amt]');}
  else if(b.type==='text'||b.type==='shape'){b.fields.text=txt(el,'[data-k=text]');}
}
function txt(el,q){const n=el.querySelector(q);return n?n.textContent.trim():'';}

// ── выделение / удаление ──
function select(id){selected=id;[...canvas.querySelectorAll('.blk')].forEach(e=>e.classList.toggle('sel',e.dataset.id===id));}
function removeBlock(id){blocks=blocks.filter(b=>b.id!==id);edges=edges.filter(e=>e.from!==id&&e.to!==id);selected=null;renderAll();}
function delSel(){if(selected)removeBlock(selected);else setHint('Сначала кликни блок, затем «Блок»');}
function clearAll(){if(confirm('Удалить ВСЕ блоки и связи?')){blocks=[];edges=[];selected=null;renderAll();}}

// ── связи ──
function toggleLink(){linkMode=!linkMode;linkFrom=null;document.getElementById('linkBtn').classList.toggle('on',linkMode);
  canvas.classList.toggle('linking',linkMode);[...canvas.querySelectorAll('.blk')].forEach(e=>e.classList.remove('src'));
  setHint(linkMode?'Режим связи: кликни блок-ИСТОЧНИК, затем блок-ПОЛУЧАТЕЛЬ':'');}
function clickLink(id){if(!linkFrom){linkFrom=id;const el=canvas.querySelector(`.blk[data-id="${id}"]`);if(el)el.classList.add('src');setHint('Теперь кликни блок-ПОЛУЧАТЕЛЬ');return;}
  if(id!==linkFrom)edges.push({id:uid(),from:linkFrom,to:id});linkFrom=null;linkMode=false;
  document.getElementById('linkBtn').classList.remove('on');canvas.classList.remove('linking');renderAll();setHint('Связь добавлена');}

// ── drag / resize (8 направлений) ──
let drag=null, pend=null, pan=null;
function startDrag(e,b,mode,dir){e.preventDefault();const el=canvas.querySelector(`.blk[data-id="${b.id}"]`);
  if(document.activeElement&&document.activeElement.blur)document.activeElement.blur();
  drag={b,el,mode,dir:dir||'',sx:e.clientX,sy:e.clientY,ox:el.offsetLeft,oy:el.offsetTop,ow:el.offsetWidth,oh:el.offsetHeight};}
// нажатие на блок → потенциальное перетаскивание за ЛЮБОЕ место; клик без сдвига = редактирование
function onBlockDown(e,b){
  if(linkMode)return; select(b.id);
  const t=e.target;
  if(['lx','addln','clr','x','h','photobox'].some(c=>t.classList.contains(c))) return;
  const el=canvas.querySelector(`.blk[data-id="${b.id}"]`);
  pend={b,el,sx:e.clientX,sy:e.clientY,ox:el.offsetLeft,oy:el.offsetTop,editable:t.isContentEditable};
}
// панорама холста — тянем пустое место
canvas.addEventListener('mousedown',e=>{if(e.target===canvas||e.target===stage||e.target===wire){
  selected=null;[...canvas.querySelectorAll('.blk')].forEach(x=>x.classList.remove('sel'));
  pan={sx:e.clientX,sy:e.clientY,px:panX,py:panY};canvas.classList.add('panning');}});
window.addEventListener('mousemove',e=>{
  if(pend&&!drag){const dd=Math.hypot(e.clientX-pend.sx,e.clientY-pend.sy);
    if(dd>(pend.editable?7:3)){
      drag={b:pend.b,el:pend.el,mode:'move',dir:'',sx:pend.sx,sy:pend.sy,ox:pend.ox,oy:pend.oy,
            ow:pend.el.offsetWidth,oh:pend.el.offsetHeight};
      if(document.activeElement&&document.activeElement.blur)document.activeElement.blur();}}
  if(drag){const dx=(e.clientX-drag.sx)/zoom,dy=(e.clientY-drag.sy)/zoom,el=drag.el,b=drag.b,d=drag.dir;
    if(drag.mode==='move'){b.x=Math.round(drag.ox+dx);b.y=Math.round(drag.oy+dy);el.style.left=b.x+'px';el.style.top=b.y+'px';}
    else{let x=drag.ox,y=drag.oy,w=drag.ow,h=drag.oh;
      if(d.includes('e'))w=Math.max(150,drag.ow+dx);
      if(d.includes('s'))h=Math.max(70,drag.oh+dy);
      if(d.includes('w')){w=Math.max(150,drag.ow-dx);x=drag.ox+(drag.ow-w);}
      if(d.includes('n')){h=Math.max(70,drag.oh-dy);y=drag.oy+(drag.oh-h);}
      b.x=Math.round(x);b.y=Math.round(y);b.w=Math.round(w);
      el.style.left=b.x+'px';el.style.top=b.y+'px';el.style.width=b.w+'px';
      if(b.type==='photo'){b.h=Math.round(h);el.style.height=b.h+'px';}
      else{b.minh=Math.round(h);el.style.minHeight=b.minh+'px';el.style.height='';}}
    redraw();
  } else if(pan){panX=pan.px+(e.clientX-pan.sx);panY=pan.py+(e.clientY-pan.sy);applyView();}
});
window.addEventListener('mouseup',()=>{drag=null;pend=null;pan=null;canvas.classList.remove('panning');});
canvas.addEventListener('wheel',e=>{e.preventDefault();const r=canvas.getBoundingClientRect();
  const mx=e.clientX-r.left,my=e.clientY-r.top,wx=(mx-panX)/zoom,wy=(my-panY)/zoom;
  zoom=Math.max(0.3,Math.min(2.5,zoom*(e.deltaY<0?1.1:0.9)));panX=mx-wx*zoom;panY=my-wy*zoom;applyView();},{passive:false});

// ── форматирование ──
function fmtBold(){document.execCommand('bold');}
function fmtSize(dir){if(!curEdit)return setHint('Поставь курсор в текст, потом A+/A−');
  const cur=parseFloat(getComputedStyle(curEdit).fontSize)||14;curEdit.style.fontSize=Math.max(9,Math.min(48,cur+dir*2))+'px';redraw();}

// ── стрелки ──
function anchors(el){const x=el.offsetLeft,y=el.offsetTop,w=el.offsetWidth,h=el.offsetHeight;
  return [[x+w/2,y],[x+w/2,y+h],[x,y+h/2],[x+w,y+h/2]];}
function elOf(id){return canvas.querySelector(`.blk[data-id="${id}"]`);}
function redraw(){let W=0,H=0;blocks.forEach(b=>{const el=elOf(b.id);if(!el)return;W=Math.max(W,el.offsetLeft+el.offsetWidth);H=Math.max(H,el.offsetTop+el.offsetHeight);});
  wire.setAttribute('width',W+120);wire.setAttribute('height',H+120);
  let s='<defs><marker id="a" markerWidth="11" markerHeight="11" refX="8" refY="3" orient="auto"><path d="M0,0 L8,3 L0,6 Z" fill="#555"/></marker></defs>';
  edges.forEach(ed=>{const A=elOf(ed.from),B=elOf(ed.to);if(!A||!B)return;const aa=anchors(A),ba=anchors(B);let best=null;
    aa.forEach(p=>ba.forEach(q=>{const dd=(p[0]-q[0])**2+(p[1]-q[1])**2;if(!best||dd<best.d)best={d:dd,p,q};}));
    s+=`<line data-eid="${esc(ed.id)}" x1="${best.p[0]}" y1="${best.p[1]}" x2="${best.q[0]}" y2="${best.q[1]}" stroke="#555" stroke-width="2" marker-end="url(#a)" style="pointer-events:stroke;cursor:pointer"/>`;});
  aff.forEach(p=>{const A=elOf(p.a),B=elOf(p.b);if(!A||!B)return;const aa=anchors(A),ba=anchors(B);let best=null;
    aa.forEach(u=>ba.forEach(v=>{const dd=(u[0]-v[0])**2+(u[1]-v[1])**2;if(!best||dd<best.d)best={d:dd,u,v};}));
    s+=`<line x1="${best.u[0]}" y1="${best.u[1]}" x2="${best.v[0]}" y2="${best.v[1]}" stroke="#c00" stroke-width="2" stroke-dasharray="6 4"/>`;});
  wire.innerHTML=s;
  wire.querySelectorAll('line').forEach(ln=>ln.addEventListener('click',()=>{if(confirm('Удалить связь?')){edges=edges.filter(e=>e.id!==ln.dataset.eid);redraw();}}));}

// ── фото ──
function pickPhoto(b){const inp=document.createElement('input');inp.type='file';inp.accept='image/*';
  inp.onchange=()=>{const f=inp.files[0];if(!f)return;const fr=new FileReader();fr.onload=()=>{b.photo=fr.result;renderAll();};fr.readAsDataURL(f);};inp.click();}

// ── сохранение ──
function save(){download('schema_state.json',JSON.stringify({project_name:DATA.project_name||'',_editor:{blocks,edges}},null,2),'application/json');}

// ── экспорт ──
function buildSVG(){let W=0,H=0;blocks.forEach(b=>{const el=elOf(b.id);W=Math.max(W,el.offsetLeft+el.offsetWidth);H=Math.max(H,el.offsetTop+el.offsetHeight);});W+=40;H+=40;
  let s=`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" font-family="Arial">`;
  s+='<defs><marker id="a" markerWidth="11" markerHeight="11" refX="8" refY="3" orient="auto"><path d="M0,0 L8,3 L0,6 Z" fill="#555"/></marker></defs>';
  s+=`<rect width="${W}" height="${H}" fill="#fff"/>`;
  edges.forEach(ed=>{const A=elOf(ed.from),B=elOf(ed.to);if(!A||!B)return;const aa=anchors(A),ba=anchors(B);let best=null;
    aa.forEach(p=>ba.forEach(q=>{const dd=(p[0]-q[0])**2+(p[1]-q[1])**2;if(!best||dd<best.d)best={d:dd,p,q};}));
    s+=`<line x1="${best.p[0]}" y1="${best.p[1]}" x2="${best.q[0]}" y2="${best.q[1]}" stroke="#555" stroke-width="2" marker-end="url(#a)"/>`;});
  blocks.forEach(b=>{const el=elOf(b.id),x=el.offsetLeft,y=el.offsetTop,w=el.offsetWidth,h=el.offsetHeight;
    const clr=b.color||COL[b.role]||'#777', ttl=(b.title!=null?b.title:(ROLE_RU[b.role]||''));
    if(b.type==='photo'){s+=`<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="#f2f2f2" stroke="${clr}"/>`;
      if(b.photo)s+=`<image x="${x+2}" y="${y+22}" width="${w-4}" height="${h-24}" preserveAspectRatio="xMidYMid slice" href="${b.photo}"/>`;else s+=svgWrap('ФОТО',x+8,y+h/2,w-16,12,false,'#999');return;}
    s+=`<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="6" fill="#fff" stroke="${clr}" stroke-width="2"/>`;
    s+=`<rect x="${x}" y="${y}" width="${w}" height="20" fill="${clr}"/>`;
    s+=`<text x="${x+7}" y="${y+14}" font-size="11" fill="#fff">${xmlesc(ttl)}</text>`;
    let ty=y+38;const ln=(t,bo,c)=>{s+=svgWrap(t,x+8,ty,w-16,bo?15:14,bo,c||'#222');ty+=(bo?20:17)*wrapCount(t,w-16,bo?15:14);};
    if(b.type==='org'){(b.lines||[]).forEach(l=>ln(l.t,l.k==='name'||l.k==='sub',l.k==='viol'?'#c00':'#222'));}
    else if(b.type==='contract'){ln(b.fields.desc);ln(b.fields.amt,true,'#c00');}
    else if(b.type==='text'||b.type==='shape'){ln(b.fields.text);}});
  s+='</svg>';return {svg:s,w:W,h:H};}
function wrapCount(t,w,fs){return Math.max(1,Math.ceil((esc(t).length*fs*0.55)/w));}
function svgWrap(t,x,y,w,fs,bold,col){t=esc(t);const cpl=Math.max(6,Math.floor(w/(fs*0.55)));let out='',line='',yy=y;
  t.split(' ').forEach(word=>{if((line+' '+word).trim().length>cpl){out+=`<text x="${x}" y="${yy}" font-size="${fs}" ${bold?'font-weight="bold"':''} fill="${col}">${xmlesc(line)}</text>`;line=word;yy+=fs+3;}else line=(line+' '+word).trim();});
  out+=`<text x="${x}" y="${yy}" font-size="${fs}" ${bold?'font-weight="bold"':''} fill="${col}">${xmlesc(line)}</text>`;return out;}
function xmlesc(s){return esc(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function exportSVG(){download('schema.svg',buildSVG().svg,'image/svg+xml');}
function exportRaster(type){const {svg,w,h}=buildSVG();const img=new Image();
  img.onload=()=>{const sc=2,c=document.createElement('canvas');c.width=w*sc;c.height=h*sc;const g=c.getContext('2d');
    g.fillStyle='#fff';g.fillRect(0,0,c.width,c.height);g.scale(sc,sc);g.drawImage(img,0,0);
    c.toBlob(bl=>{const a=document.createElement('a');a.href=URL.createObjectURL(bl);a.download='schema.'+(type==='image/png'?'png':'jpg');a.click();},type,0.95);};
  img.src='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(svg);}
function download(name,text,type){const b=new Blob([text],{type});const a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=name;a.click();}

// ── вид: панорама + зум (transform), раздвижка блоков ──
let zoom=1,panX=0,panY=0;
function applyView(){stage.style.transform=`translate(${panX}px,${panY}px) scale(${zoom})`;}
function setZoom(f){zoom=Math.max(0.3,Math.min(2.5,zoom*f));applyView();}
function spread(f){                       // раздвинуть/сжать всю раскладку вокруг центра
  let cx=0,cy=0,n=0;blocks.forEach(b=>{cx+=b.x;cy+=b.y;n++;});cx/=n||1;cy/=n||1;
  blocks.forEach(b=>{b.x=Math.round(cx+(b.x-cx)*f);b.y=Math.round(cy+(b.y-cy)*f);});
  renderAll();}

// ══════════════ ПРОВЕРКА СХЕМЫ ══════════════
function blockText(b){
  if(b.type==='org') return (b.lines||[]).map(l=>l.t).join(' \n');
  if(b.type==='contract') return (b.fields.desc||'')+' '+(b.fields.amt||'');
  if(b.fields&&b.fields.text) return b.fields.text; return '';
}
function label(b){return (b.type==='org'&&b.lines&&b.lines[0])?b.lines[0].t:(ROLE_RU[b.role]||b.type);}

// поиск/подсветка
function doSearch(q){q=(q||'').trim().toLowerCase();
  blocks.forEach(b=>{const el=elOf(b.id);if(!el)return;
    el.classList.toggle('hit', q.length>1 && blockText(b).toLowerCase().includes(q));});}

// контрольный разряд ИИН/БИН РК (12 цифр)
function idnValid(s){if(!/^\d{12}$/.test(s))return false;
  const d=s.split('').map(Number);
  const w1=[1,2,3,4,5,6,7,8,9,10,11],w2=[3,4,5,6,7,8,9,10,11,1,2];
  let c=d.slice(0,11).reduce((a,x,i)=>a+x*w1[i],0)%11;
  if(c===10){c=d.slice(0,11).reduce((a,x,i)=>a+x*w2[i],0)%11; if(c===10)return false;}
  return c===d[11];}

function parseAmount(s){ // «267,5 млн» → тенге
  if(!s)return 0;const m=String(s).match(/(\d[\d\s]*(?:[.,]\d+)?)\s*(млрд|млн|тыс)/i);
  if(!m)return 0;let n=parseFloat(m[1].replace(/\s/g,'').replace(',','.'));
  const k={'млрд':1e9,'млн':1e6,'тыс':1e3}[m[2].toLowerCase()];return n*k;}
function fmtTenge(n){return n>=1e9?(n/1e9).toFixed(2)+' млрд':n>=1e6?(n/1e6).toFixed(1)+' млн':Math.round(n).toLocaleString('ru');}

function norm(s){return (s||'').toLowerCase().replace(/\s+/g,' ').trim();}
function graphDepth(adj,roots){
  const memo=new Map(),active=new Set();let cycle=false;
  function visit(id){
    if(active.has(id)){cycle=true;return 0;}
    if(memo.has(id))return memo.get(id);
    active.add(id);let depth=1;
    for(const child of adj[id]||[])depth=Math.max(depth,1+visit(child));
    active.delete(id);memo.set(id,depth);return depth;
  }
  let depth=0;for(const root of roots)depth=Math.max(depth,visit(root));
  for(const id of Object.keys(adj))if(!memo.has(id))visit(id);
  return {depth,cycle};
}
function analyze(){
  aff=[]; const F=[];
  // ── аффилированность: общий 12-значный код / телефон / адрес между блоками ──
  const idMap={},telMap={},adrMap={};
  const push=(m,v,id)=>{if(!v)return;(m[v]=m[v]||new Set()).add(id);};
  blocks.forEach(b=>{const t=blockText(b);
    (t.match(/\b\d{12}\b/g)||[]).forEach(v=>push(idMap,v,b.id));
    (t.match(/\+?[78][\d\s()\-]{9,13}/g)||[]).forEach(v=>push(telMap,v.replace(/\D/g,''),b.id));
    const am=t.match(/адрес[^:]*:\s*([^\n]+)/i); if(am&&norm(am[1]).length>6)push(adrMap,norm(am[1]),b.id);
  });
  const affReport=[];
  const scan=(m,kind)=>{Object.entries(m).forEach(([v,set])=>{const ids=[...set];
    if(ids.length>=2){for(let i=0;i<ids.length-1;i++)aff.push({a:ids[i],b:ids[i+1]});
      affReport.push(`${kind}: совпадает у «${label(byId(ids[0]))}» и «${label(byId(ids[1]))}»${ids.length>2?' (+'+(ids.length-2)+')':''}`);}});};
  scan(idMap,'ИИН/БИН'); scan(telMap,'Телефон'); scan(adrMap,'Адрес');

  // ── деньги ──
  let total=0,cnt=0,amts=[];
  blocks.forEach(b=>{if(b.type==='contract'){const a=parseAmount(b.fields.amt);if(a){total+=a;cnt++;amts.push(a);}}});

  // ── глубина цепочки (по связям от заказчика) ──
  const adj={};edges.forEach(e=>{(adj[e.from]=adj[e.from]||[]).push(e.to);});
  const roots=blocks.filter(b=>b.role==='cust').map(b=>b.id);
  const depthInfo=graphDepth(adj,roots.length?roots:blocks.map(b=>b.id));
  const maxDepth=depthInfo.depth;

  // ── счётчики ──
  const by=r=>blocks.filter(b=>b.role===r).length;
  const orgs=by('cust')+by('con')+by('sub'), viol=by('violation'), people=by('person');

  // ── валидация ──
  const val=[];if(depthInfo.cycle)val.push("Обнаружен цикл в связях: проверьте схему");
  blocks.forEach(b=>{const t=blockText(b);
    (t.match(/(?:иин|бин)\D{0,3}(\d{6,13})/gi)||[]).forEach(m=>{const num=m.match(/\d{6,13}/)[0];
      if(num.length!==12) val.push(`${label(b)}: код «${num}» не 12 цифр`);
      else if(!idnValid(num)) val.push(`${label(b)}: ИИН/БИН «${num}» — неверный контрольный разряд`);});
  });
  const linked=new Set();edges.forEach(e=>{linked.add(e.from);linked.add(e.to);});
  const orphan=blocks.filter(b=>['cust','con','sub','person','contract'].includes(b.role)&&!linked.has(b.id));
  if(orphan.length) val.push(`Без связей: ${orphan.length} блок(ов) — «${esc(orphan.slice(0,3).map(label).join('», «'))}»`);

  // ── риск-скор ──
  let risk=affReport.length*3 + viol*2 + (maxDepth>=4?3:maxDepth>=3?1:0);
  const riskTxt=risk>=8?'ВЫСОКИЙ':risk>=4?'СРЕДНИЙ':'низкий';
  const riskCls=risk>=8?'bad':risk>=4?'warn':'ok';

  // ── рендер панели ──
  let h=`<h3>Риск</h3><div class="risk ${riskCls}">${riskTxt} (${risk})</div>`;
  h+=`<h3>Сводка</h3>
    <div class="f">Организаций: <b>${orgs}</b> · лиц: <b>${people}</b> · договоров: <b>${cnt}</b> · нарушений: <b class="${viol?'bad':''}">${viol}</b></div>
    <div class="f">Общая сумма договоров: <b>${fmtTenge(total)} тг</b></div>
    <div class="f">Глубина субподряда: <b class="${maxDepth>=4?'bad':maxDepth>=3?'warn':''}">${maxDepth}</b></div>`;
  h+=`<h3>Аффилированность ${affReport.length?'<span class="bad">('+affReport.length+')</span>':''}</h3>`;
  h+= affReport.length? affReport.map(x=>`<div class="f bad">🔴 ${esc(x)}</div>`).join('') : '<div class="f ok">Явных совпадений не найдено</div>';
  h+=`<h3>Валидация ${val.length?'<span class="warn">('+val.length+')</span>':''}</h3>`;
  h+= val.length? val.map(x=>`<div class="f warn">⚠ ${esc(x)}</div>`).join('') : '<div class="f ok">Замечаний нет</div>';
  h+=`<div class="f" style="margin-top:10px;color:#888">Красные пунктирные линии на схеме — предполагаемая аффилированность (проверить документально).</div>`;
  document.getElementById('pbody').innerHTML=h;
  document.getElementById('panel').style.display='block';
  blocks.forEach(b=>{const el=elOf(b.id);if(el)el.classList.toggle('aff',aff.some(p=>p.a===b.id||p.b===b.id));});
  redraw();
}
function byId(id){return blocks.find(b=>b.id===id)||{};}
function closePanel(){document.getElementById('panel').style.display='none';clearAff();}
function clearAff(){aff=[];blocks.forEach(b=>{const el=elOf(b.id);if(el)el.classList.remove('aff');});redraw();}

init(); applyView();
</script>
</body></html>
"""


def build_html(data):
    from validation import validate_editor
    if "_editor" in data:
        validate_editor(data["_editor"])
    payload = json.dumps(data, ensure_ascii=False)
    payload = payload.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    return _TEMPLATE.replace("/*DATA*/", payload)
