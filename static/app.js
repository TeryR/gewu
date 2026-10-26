let S=null, edits={}, sel=null;
const COLORS={table:"#2f6fdb",note:"#d9a406",mixed:"#d43b3b",ignore:"#888888"};
const LABELS={table:"🟦数据表",note:"🟨文本/批注",mixed:"🟥混合待复核",ignore:"⬜忽略"};
function lbl(k){return t("lbl_"+k)||LABELS[k]}   // 类型标签双语

async function loadSheets(){
  const sheets=await (await fetch('/api/sheets')).json();
  SHEETS=sheets;
  renderScope();loadSessions();refreshLLMCfg();
  const sel=document.getElementById('sheetSel');
  sel.innerHTML=sheets.map(s=>`<option value="${esc(s.name)}">${s.visible?'':'['+t('m_hidden_sheet')+'] '}${esc(s.name)}</option>`).join('');
  sel.onchange=()=>loadState(sel.value);
  loadState(sheets[0].name);
  refreshStatus();
  applyLang();
  loadTabs();setTimeout(syncTabs,600);
}
async function refreshStatus(){
  const st=await (await fetch('/api/status')).json();
  const el=document.getElementById('status');
  el.textContent=st.status==='frozen'?t('status_frozen_pre')+st.frozen_at:t('status_parsed');
  el.className=st.status==='frozen'?'frozen':'';
  window.COPY=!!st.copy;
  document.getElementById('wbName').textContent=(S?S.workbook.path:'')+(st.copy?t('m_copy_badge'):'');
  document.getElementById('exportBtn').style.display=st.copy?'':'none';
  const bar=document.getElementById('changeBar');
  const ch=st.change;
  if(ch&&ch.type==='structure'){
    bar.style.display='';
    bar.innerHTML=`${t('chg_struct_pre')}${ch.added.length}${t('chg_struct_mid')}${ch.removed.length}${t('chg_struct_mid2')}${ch.changed.length}${t('chg_struct_mid3')}`+
      (ch.added.length?`${t('chg_added')}${ch.added.join('、')}。`:'')+
      (ch.removed.length?`${t('chg_removed')}${ch.removed.join('、')}。`:'')+
      `${t('chg_tail')}${ch.marks_transferred}${t('chg_tail2')}`;
  }else if(ch&&ch.type==='value-only'){
    bar.style.display='';
    bar.innerHTML=t('chg_value');
  }else{bar.style.display='none'}
}
async function loadState(name){
  S=await (await fetch('/api/state?sheet='+encodeURIComponent(name))).json();
  document.getElementById('wbName').textContent=S.workbook.path;
  render();
}
function render(){
  const [r1,c1,r2,c2]=S.range, g=document.getElementById('grid');
  const covered=new Set(), topleft={};
  S.merges.forEach(m=>{topleft[`${m[0]},${m[1]}`]=m;
    for(let r=m[0];r<=m[2];r++)for(let c=m[1];c<=m[3];c++)if(!(r==m[0]&&c==m[1]))covered.add(`${r},${c}`)});
  const hr=new Set(S.hidden_rows), hc=new Set(S.hidden_cols);
  let html='<tr><td class="hcol hrow"></td>';
  for(let c=c1;c<=c2;c++)html+=`<td class="hcol">${hc.has(c)?t('m_hidden_cell'):''}${colName(c)}</td>`;
  html+='</tr>';
  for(let r=r1;r<=r2;r++){
    html+=`<tr><td class="hrow">${hr.has(r)?t('m_hidden_cell'):''}${r}</td>`;
    for(let c=c1;c<=c2;c++){
      if(covered.has(`${r},${c}`))continue;
      const cell=S.cells[`${r},${c}`]; const m=topleft[`${r},${c}`];
      const span=m?` rowspan="${m[2]-m[0]+1}" colspan="${m[3]-m[1]+1}"`:'';
      let style='';
      if(cell&&cell.fm&&!String(cell.fm).startsWith('theme'))style=` style="background:#${cell.fm.slice(2)}55"`;
      let txt='';
      if(cell){txt=cell.v;if(cell.pct&&typeof txt==='number')txt=(txt*100).toFixed(1)+'%';
        if(cell.f)txt='ƒ'+txt}
      html+=`<td data-r="${r}" data-c="${c}"${span}${style} title="${esc(String(txt).slice(0,120))}">${esc(String(txt))}</td>`;
    }
    html+='</tr>';
  }
  g.innerHTML=html;
  autoFitCols();
  drawOverlays();
}
function esc(s){return s.replace(/&/g,'&amp;').replace(/</g,'&lt;')}
function colName(c){let s='';while(c>0){s=String.fromCharCode(65+(c-1)%26)+s;c=Math.floor((c-1)/26)}return s}
function hexA(hex,a){a=Math.min(Math.max(a,0),1);
  return hex+Math.round(a*255).toString(16).padStart(2,'0')}
function paintAll(){const base=(+document.getElementById('opSlider').value)/100;
  document.querySelectorAll('.blk').forEach(d=>{
    const quiet=d.dataset.quiet==='1';
    const a=base*(quiet?0.25:1)+(d.dataset.boost==='hover'?0.18:0)
      +(d.classList.contains('sel')?0.25:0);
    d.style.background=hexA(d.dataset.color,a)})}
function drawOverlays(){
  const ov=document.getElementById('overlays');ov.innerHTML='';
  ov.style.display=document.getElementById('showMask').checked?'':'none';
  // 表内批注：黄框几何上落在某个蓝表范围内 → 标注出来，避免"蓝框套黄框"看不懂
  const insideOf={};
  S.blocks.forEach(a=>{
    if(typeOf(a)!=='table')return;
    S.blocks.forEach(b=>{
      if(a===b||typeOf(b)!=='note')return;
      const [ar1,ac1,ar2,ac2]=bboxOf(a),[br1,bc1,br2,bc2]=bboxOf(b);
      if(br1>=ar1&&br2<=ar2&&bc1>=ac1&&bc2<=ac2)insideOf[b.id]=a.id;
    });
  });
  const wrap=document.getElementById('gridWrap'), wr=wrap.getBoundingClientRect();
  const scTop=document.getElementById('stage').getBoundingClientRect().top;  // 裁切边界=滚动视口顶
  S.blocks.forEach(b=>{
    const t=typeOf(b), a=anchor(bboxOf(b)), [r1,c1,r2,c2]=bboxOf(b);
    if(!a)return;
    const d=document.createElement('div');d.className='blk'+(sel===b.id?' sel':'');
    d.dataset.color=COLORS[t];
    if(t==='note'||t==='ignore'){d.dataset.quiet='1';d.style.borderStyle='dashed'}
    const e1=anchorCell(r1,c1), e2=anchorCell(r2,c2)||a;
    if(!e1)return;
    const p=e1.getBoundingClientRect(), q=e2.getBoundingClientRect();
    d.style.cssText+=`left:${p.left-wr.left}px;top:${p.top-wr.top}px;width:${q.right-p.left}px;
      height:${q.bottom-p.top}px;border-color:${COLORS[t]}`;
    const nested=insideOf[b.id]!==undefined;
    d.innerHTML=`<span class="tag" style="color:${COLORS[t]}">${lbl(t)} #${b.id} ${b.score??''}${nested?' · '+t('nested_note'):''}</span>`;
    d.onmouseenter=()=>{d.dataset.boost='hover';paintAll()};
    d.onmouseleave=()=>{d.dataset.boost='';paintAll()};
    d.onclick=()=>{if(adj)return;sel=b.id;showDetail(b);drawOverlays()};
    if(adj)d.style.pointerEvents='none';  // 调整模式：彩框不吃点击，点两角/拖拽直达格子
    if(sel===b.id&&!adj)_adjHandles(d,b.id);   // 选中即出把手：直接拖边/角调整
    ov.appendChild(d);
    if(d.getBoundingClientRect().top-scTop<12)
      d.querySelector('.tag').style.top='2px';   // 贴顶块：标签翻进框内，避免被裁一半
  });
  if(adj){                       // 调整范围模式的橙色虚线预览（带把手可直接拖）
    const e1=anchorCell(adj.vals[0],adj.vals[1]), e2=anchorCell(adj.vals[2],adj.vals[3]);
    if(e1&&e2){
      const p=e1.getBoundingClientRect(), q=e2.getBoundingClientRect();
      const d=document.createElement('div');
      d.id='adjPrev';
      d.style.cssText=`position:absolute;left:${p.left-wr.left}px;top:${p.top-wr.top}px;
        width:${q.right-p.left}px;height:${q.bottom-p.top}px;border:2px dashed #ea580c;
        border-radius:4px;pointer-events:none;z-index:6`;
      d.innerHTML=`<span style="position:absolute;left:4px;top:2px;font-size:11px;
        color:#ea580c;font-weight:700;white-space:nowrap;background:rgba(255,255,255,.9);
        padding:0 4px;border-radius:6px">${t('adj_prev')}</span>`;
      _adjHandles(d,null);
      ov.appendChild(d);
    }
  }
  paintAll();
}
function typeOf(b){return (edits[b.id]&&edits[b.id].user_type)||b.type}
function anchor(bbox){return anchorCell(bbox[0],bbox[1])}
function anchorCell(r,c){
  // 精确格 → 若被合并区盖住（这些位置不渲染 td），用合并区左上格（其矩形覆盖该位置）
  const exact=document.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
  if(exact)return exact;
  const m=(S.merges||[]).find(m=>m[0]<=r&&r<=m[2]&&m[1]<=c&&c<=m[3]);
  if(m)return document.querySelector(`td[data-r="${m[0]}"][data-c="${m[1]}"]`);
  const rowTd=document.querySelector(`td[data-r="${r}"][data-c="${c}"]`)
    ||document.querySelector(`td[data-r="${r}"]`);
  if(rowTd)return rowTd;
  // 行超渲染范围（serve.max_rows 限制）→ 钳制到最后渲染行，取该列（或最后格）
  const all=document.querySelectorAll('#grid td[data-r]');
  if(!all.length)return null;
  const lastRow=all[all.length-1].dataset.r;
  return document.querySelector(`td[data-r="${lastRow}"][data-c="${c}"]`)||all[all.length-1];
}
function showHelp(){document.getElementById('helpBox').style.display=''}
const EMAIL="";


/* ===== 框微调：点两角选范围（Esc 取消）===== */
let adj=null;
document.addEventListener('keydown',e=>{if(adj&&e.key==='Escape')closeAdjust()});
document.getElementById('gridWrap').addEventListener('click',e=>{
  if(!adj)return;
  const td=e.target.closest('td[data-r]');
  if(td)adjCellClick(td);
});
/* 滚动时贴顶块标签动态翻进框内（只动标签样式，不整页重绘） */
let _tagT;
document.getElementById('stage').addEventListener('scroll',()=>{
  clearTimeout(_tagT);
  _tagT=setTimeout(()=>{
    const stTop=document.getElementById('stage').getBoundingClientRect().top;   // 视口顶才是裁切边界
    document.querySelectorAll('#overlays .blk').forEach(b=>{
      const tag=b.querySelector('.tag');if(!tag)return;
      tag.style.top=(b.getBoundingClientRect().top-stTop<12)?'2px':'';
    });
  },80);
});
function bboxOf(b){return (edits[b.id]&&edits[b.id].user_bbox)||b.bbox}
function startAdjust(id){
  const b=S.blocks.find(x=>x.id===id);if(!b)return;
  if(adj&&adj.id===id)return;              // 同块已在调整中：不重置当前值
  const [r1,c1,r2,c2]=bboxOf(b);
  adj={id,vals:[r1,c1,r2,c2],clicks:0};
  showAdjustBox();drawOverlays();
}
function showAdjustBox(){
  document.getElementById('adjustBox').style.display='';
  ['ar1','ac1','ar2','ac2'].forEach((k,i)=>document.getElementById(k).value=adj.vals[i]);
  document.getElementById('adjTitle').textContent=t('adj_title')+' #'+adj.id;
}
function adjInput(){
  adj.vals=['ar1','ac1','ar2','ac2'].map(k=>parseInt(document.getElementById(k).value)||0);
  drawOverlays();
}
function adjCellClick(td){        // 第1点=一角，第2点=对角（自动归位大小）
  const r=+td.dataset.r,c=+td.dataset.c;
  if(adj.clicks===0){adj.vals=[r,c,r,c];adj.clicks=1}
  else{adj.vals=[Math.min(adj.vals[0],r),Math.min(adj.vals[1],c),
                 Math.max(adj.vals[2],r),Math.max(adj.vals[3],c)];adj.clicks=0}
  showAdjustBox();drawOverlays();
}
/* 拖拽调整：选中块的蓝框带 8 把手；调整模式的橙虚线带把手+移动钮 */
function _adjHandles(host, blockId){
  const P={nw:'left:-5px;top:-5px',n:'left:50%;top:-5px;transform:translateX(-50%)',
    ne:'right:-5px;top:-5px',e:'right:-5px;top:50%;transform:translateY(-50%)',
    se:'right:-5px;bottom:-5px',s:'left:50%;bottom:-5px;transform:translateX(-50%)',
    sw:'left:-5px;bottom:-5px',w:'left:-5px;top:50%;transform:translateY(-50%)'};
  [['nw','nwse-resize'],['n','ns-resize'],['ne','nesw-resize'],['e','ew-resize'],
   ['se','nwse-resize'],['s','ns-resize'],['sw','nesw-resize'],['w','ew-resize']]
  .forEach(([h,cur])=>{
    const hd=document.createElement('div');
    hd.style.cssText=`position:absolute;${P[h]};width:11px;height:11px;background:#fff;
      border:2px solid ${blockId==null?'#ea580c':'#2563eb'};border-radius:3px;cursor:${cur};z-index:7;
      pointer-events:auto`;
    hd.onmousedown=e=>adjHandleDown(e,blockId,h);
    host.appendChild(hd);
  });
  if(blockId==null){                  // 仅调整模式预览：整框平移把手（块 id 可为 0，不能当假值）
    const mv=document.createElement('div');
    mv.title=t('adj_move_tt');
    mv.style.cssText='position:absolute;right:4px;top:2px;padding:1px 8px;'+
      'background:rgba(234,88,12,.85);color:#fff;border-radius:8px;font-size:10px;'+
      'cursor:move;z-index:8;white-space:nowrap;pointer-events:auto';
    mv.textContent=t('adj_move');
    mv.onmousedown=e=>adjHandleDown(e,null,'move');
    host.appendChild(mv);
  }
}
function adjHandleDown(e,blockId,h){
  e.stopPropagation();e.preventDefault();
  if(!adj){
    if(blockId==null)return;          // null=预览把手以外的误触；0 是合法块 id
    startAdjust(blockId);
  }
  if(!adj)return;
  adj.drag={h,vals:[...adj.vals],startCell:cellAt(e)};
  document.addEventListener('mousemove',adjDragMove);
  document.addEventListener('mouseup',adjDragUp);
}
function cellAt(e){
  // 命中穿透：把手/彩框盖在格子上（缩边时光标在框内），取该点元素栈里第一个 td
  const stack=document.elementsFromPoint(e.clientX,e.clientY)||[];
  for(const el of stack){
    const td=el.closest?el.closest('td[data-r]'):null;
    if(td)return{r:+td.dataset.r,c:+td.dataset.c};
  }
  return null;
}
function adjDragMove(e){
  if(!adj||!adj.drag)return;
  const cell=cellAt(e);if(!cell)return;
  const {h,vals:d,startCell}=adj.drag;
  let v;
  if(h==='move'){
    const dr=cell.r-startCell.r,dc=cell.c-startCell.c;
    v=[d[0]+dr,d[1]+dc,d[2]+dr,d[3]+dc];
  }else{
    v=[...d];
    if(h.indexOf('n')>=0)v[0]=cell.r;
    if(h.indexOf('s')>=0)v[2]=cell.r;
    if(h.indexOf('w')>=0)v[1]=cell.c;
    if(h.indexOf('e')>=0)v[3]=cell.c;
  }
  adj.vals=v.map(x=>Math.max(1,Math.round(x)));
  syncAdjustUI();
}
function adjDragUp(){
  document.removeEventListener('mousemove',adjDragMove);
  document.removeEventListener('mouseup',adjDragUp);
  if(adj){
    if(adj.vals[0]>adj.vals[2])[adj.vals[0],adj.vals[2]]=[adj.vals[2],adj.vals[0]];
    if(adj.vals[1]>adj.vals[3])[adj.vals[1],adj.vals[3]]=[adj.vals[3],adj.vals[1]];
    adj.drag=null;drawOverlays();
  }
}
function syncAdjustUI(){          // 拖动中轻量同步：输入框 + 预览框几何（不整页重绘）
  ['ar1','ac1','ar2','ac2'].forEach((k,i)=>document.getElementById(k).value=adj.vals[i]);
  const pv=document.getElementById('adjPrev');
  if(pv){
    const e1=anchorCell(adj.vals[0],adj.vals[1]), e2=anchorCell(adj.vals[2],adj.vals[3]);
    if(e1&&e2){
      const wr=document.getElementById('gridWrap').getBoundingClientRect();
      const p=e1.getBoundingClientRect(), q=e2.getBoundingClientRect();
      pv.style.left=(p.left-wr.left)+'px';pv.style.top=(p.top-wr.top)+'px';
      pv.style.width=(q.right-p.left)+'px';pv.style.height=(q.bottom-p.top)+'px';
    }
  }
}
function confirmAdjust(){
  const [r1,c1,r2,c2]=adj.vals;
  if(!(r1>=1&&c1>=1&&r2>=r1&&c2>=c1)){alert(t('adj_bad'));return}
  const id=adj.id;closeAdjust();
  setEdit(id,'user_bbox',[r1,c1,r2,c2]);
}
function closeAdjust(){
  adj=null;
  document.getElementById('adjustBox').style.display='none';
  drawOverlays();
}

/* 网页内拖入 xlsx → 上传副本打开（桌面壳里由 Rust 拦截为"打开原文件"，此 handler 不触发） */
window.addEventListener('dragover',e=>{e.preventDefault()});
window.addEventListener('drop',async e=>{
  e.preventDefault();
  const f=[...e.dataTransfer.files].find(f=>f.name.toLowerCase().endsWith('.xlsx'));
  if(!f)return;
  const bar=document.getElementById('dropBar')||document.body.appendChild(Object.assign(
    document.createElement('div'),{id:'dropBar'}));
  Object.assign(bar.style,{position:'fixed',left:'50%',transform:'translateX(-50%)',bottom:'14px',
    zIndex:99,background:'#fff7ed',border:'1px solid #fdba74',borderRadius:'8px',
    padding:'8px 14px',fontSize:'12.5px',color:'#92400e'});
  bar.textContent=t('drop_pre')+f.name+t('drop_suf');
  const buf=await f.arrayBuffer();
  const r=await (await fetch('/api/upload?name='+encodeURIComponent(f.name),{method:'POST',body:buf})).json();
  if(!r.ok){bar.textContent=t('upload_fail')+' '+(r.reason||'');return}
  bar.remove();await reloadForNewFile();syncTabs();
});

const LLM_PRESETS={
  siliconflow:{proto:'openai',base:'https://api.siliconflow.cn/v1',model:'',key:''},
  openai:{proto:'openai',base:'https://api.openai.com/v1',model:'gpt-4o-mini',key:''},
  deepseek:{proto:'openai',base:'https://api.deepseek.com/v1',model:'deepseek-chat',key:''},
  glm:{proto:'openai',base:'https://open.bigmodel.cn/api/paas/v4',model:'glm-4.6',key:''},
  anthropic:{proto:'anthropic',base:'https://api.anthropic.com',model:'claude-sonnet-4-5',key:''},
  ollama:{proto:'openai',base:'http://localhost:11434/v1',model:'qwen3:8b',key:''},
  lmstudio:{proto:'openai',base:'http://localhost:1234/v1',model:'',key:''},
  custom:{proto:'openai',base:'',model:'',key:''}};
function showLLM(){
  loadTraySetting();
  document.getElementById('llmBox').style.display='';
  const st=document.getElementById('llmState');
  st.textContent=(window.__llmCfg&&window.__llmCfg.model)?(window.__llmCfg.model+' @ '+(window.__llmCfg.base_url||'')):t('llm_none');
  if(window.__llmCfg&&window.__llmCfg.model){
    document.getElementById('llmProto').value=window.__llmCfg.protocol||'openai';
    document.getElementById('llmBase').value=window.__llmCfg.base_url||'';
    document.getElementById('llmModel').value=window.__llmCfg.model||'';
    document.getElementById('llmKey').value='';   // 不回显密钥
  }
}
function applyPreset(){
  const p=LLM_PRESETS[document.getElementById('llmPreset').value];
  if(!p)return;
  document.getElementById('llmProto').value=p.proto;
  document.getElementById('llmBase').value=p.base;
  document.getElementById('llmModel').value=p.model;
}

async function loadTraySetting(){
  try{
    const r=await (await fetch('/api/app_settings')).json();
    document.getElementById('trayToggle').checked=!!r.close_to_tray;
  }catch(e){}
}
async function saveTraySetting(v){
  await fetch('/api/app_settings_save',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({close_to_tray:v})});
}
async function refreshLLMCfg(){
  try{window.__llmCfg=await (await fetch('/api/llm_config')).json()}catch(e){window.__llmCfg=null}
}
async function llmSave(){
  const r=await (await fetch('/api/llm_save',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({protocol:document.getElementById('llmProto').value,
      base_url:document.getElementById('llmBase').value.trim(),
      model:document.getElementById('llmModel').value.trim(),
      api_key:document.getElementById('llmKey').value.trim()})})).json();
  document.getElementById('llmState').textContent=r.ok?t('llm_saved'):(r.reason||'');
  refreshLLMCfg();
}
async function llmTest(){
  const st=document.getElementById('llmState');st.textContent=t('llm_testing');
  const r=await (await fetch('/api/llm_test',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({protocol:document.getElementById('llmProto').value,
      base_url:document.getElementById('llmBase').value.trim(),
      model:document.getElementById('llmModel').value.trim(),
      api_key:document.getElementById('llmKey').value.trim()})})).json();
  st.innerHTML=r.ok?`<span style="color:#15803d">✓ ${r.ms}ms · ${esc(String(r.reply).slice(0,30))}</span>`
    :`<span style="color:#dc2626">✗ ${esc(String(r.reason).slice(0,90))}</span>`;
}
function showOpen(){document.getElementById('openBox').style.display='';
  document.getElementById('openErr').textContent='';
  document.getElementById('openPath').value=S?S.workbook.path:''}
async function doOpenPath(){
  const p=document.getElementById('openPath').value.trim().replace(/^"+|"+$/g,'');
  const err=document.getElementById('openErr');
  if(!p){err.textContent=t('open_err_empty');return}
  err.textContent=t('open_parsing');
  const r=await (await fetch('/api/open',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({path:p})})).json();
  if(!r.ok){err.textContent=r.reason||t('open_fail');return}
  document.getElementById('openBox').style.display='none';
  await reloadForNewFile();syncTabs();
}
async function doOpenUpload(){
  const f=document.getElementById('openFile').files[0];if(!f)return;
  const st=document.getElementById('openUpState');st.textContent=t('uploading');
  const buf=await f.arrayBuffer();
  const r=await (await fetch('/api/upload?name='+encodeURIComponent(f.name),{method:'POST',
    body:buf})).json();
  if(!r.ok){st.textContent=r.reason||t('upload_fail');return}
  st.textContent=t('parsing');
  await reloadForNewFile();syncTabs();
}
function showAbout(){
  const zh=LANG==='zh';
  document.getElementById('aboutBody').innerHTML=zh?`
   <div style="display:flex;align-items:baseline;gap:10px;border-bottom:2px solid #e2e8f0;padding-bottom:10px;margin-bottom:12px">
     <b style="font-size:22px;letter-spacing:2px">格物 <span style="color:#2563eb">Gewu</span></b>
     <span class="hint" style="font-size:12px">复杂表格理解工作台</span>
     <button onclick="document.getElementById('aboutBox').style.display='none'" style="margin-left:auto">关闭 ✕</button></div>
   <p style="margin:6px 0">格物致知——穷究每一格，理解整张表。</p>
   <p style="margin:6px 0;color:#475569">格物是<b>本地优先</b>的复杂表格理解工作台：机器解析结构，人工一次确认，随后即可问答、解读与安全改数。除显式改数外<b>永不修改原文件</b>；语义镜像、口径词典与对话会话全部保存在你自己的电脑上。</p>
   <div style="margin:14px 0 4px;padding:10px 12px;background:#f8fafc;border-left:3px solid #2563eb;border-radius:6px">
     <b>识玥 shyue</b><br>
     <span style="color:#64748b">📧 ${EMAIL}</span><br>
     <span class="hint" style="font-size:11px">Gewu/1.0</span>
   </div>`:`
   <div style="display:flex;align-items:baseline;gap:10px;border-bottom:2px solid #e2e8f0;padding-bottom:10px;margin-bottom:12px">
     <b style="font-size:22px;letter-spacing:2px">格物 <span style="color:#2563eb">Gewu</span></b>
     <span class="hint" style="font-size:12px">Intelligence workbench for complex spreadsheets</span>
     <button onclick="document.getElementById('aboutBox').style.display='none'" style="margin-left:auto">Close ✕</button></div>
   <p style="margin:6px 0">Investigate every cell; understand the whole sheet.</p>
   <p style="margin:6px 0;color:#475569">Gewu is a <b>local-first</b> workbench for understanding complex spreadsheets: the machine parses structure, you confirm once, then ask, interpret and safely edit. It <b>never touches your file</b> except explicit edit commands; the semantic mirror, glossary and chat sessions all stay on your own machine.</p>
   <div style="margin:14px 0 4px;padding:10px 12px;background:#f8fafc;border-left:3px solid #2563eb;border-radius:6px">
     <b>识玥 shyue</b><br>
     <span style="color:#64748b">📧 ${EMAIL}</span><br>
     <span class="hint" style="font-size:11px">Gewu/1.0</span>
   </div>`;
  document.getElementById('aboutBox').style.display='';
}
function showDetail(b){
  const ty=typeOf(b), d=document.getElementById('detail');
  const h=b.header||{};
  let rows=`
   <b>${t('d_type')}</b><span>${lbl(ty)}</span>
   <b>${t('d_pos')}</b><span>r${b.bbox[0]}-r${b.bbox[2]} × c${b.bbox[1]}-c${b.bbox[3]}</span>
   ${b.engine_bbox?`<b>${t('d_orig')}</b><span class="hint">r${b.engine_bbox[0]}-r${b.engine_bbox[2]} × c${b.engine_bbox[1]}-c${b.engine_bbox[3]}</span>`:''}
   <b>${t('d_score')}</b><span>${b.score}</span>`;
  if(h.header_rows)rows+=`<b>${t('d_hdr')}</b><span>r${h.header_rows.join(',r')}</span>`;
  if(h.title_text)rows+=`<b>${t('d_title')}</b><span>${esc(h.title_text)}</span>`;
  if(h.floating_rows&&h.floating_rows.length)rows+=`<b>${t('d_float')}</b><span>r${h.floating_rows.join(',r')}</span>`;
  if(h.summary_rows&&h.summary_rows.length)rows+=`<b>${t('d_summary')}</b><span>r${h.summary_rows.join(',r')}${t('d_summary_prot')}</span>`;
  if(h.dup_header_names&&h.dup_header_names.length)rows+=`<b>${t('d_dups')}</b><span>${h.dup_header_names.join('、')}</span>`;
  if(h.cross_tab)rows+=`<b>${t('d_cross')}</b><span>${t('d_yes')}</span>`;
  if(h.col_tree)rows+=`<b>${t('d_fields')}</b><span>${h.col_tree.filter(n=>n.path&&n.path.length).length} ${t('d_leaves')}</span>`;
  if(b.signals&&!b.annotation)rows+=`<b>${t('d_signal')}</b><span>${t('d_sig_text')}${b.signals.text_r} ${t('d_sig_num')}${b.signals.num_r} ${t('d_sig_border')}${b.signals.perimeter_border}</span>`;
  if(b.annotation)rows+=`<b>${t('d_anno')}</b><span>${t('d_anno_v')}</span>`;
  if(b.attach){
    rows+=`<b>${t('attach_h')}</b><span>${t('attach_full').replace('{id}',b.attach.table_block_id).replace('{r}',b.attach.row_no)}<br>
      <span class="hint">${esc((b.attach.text||'').slice(0,80))}</span><br>
      <button onclick="doDetach(${b.id})">${t('detach_btn')}</button></span>`;
  }
  const switchTitle=(ty==='note')?t('switch_tt_note'):t('switch_tt');
  rows+=`<b>${t('ops_h')}</b><span>
    <button onclick="setEdit(${b.id},'user_type','${nextType(ty)}')" title="${switchTitle}">${t('btn_switch')}${lbl(nextType(ty))}</button>
    <button onclick="setEdit(${b.id},'readonly',1)" title="${t('btn_ro_tt')}">${t('btn_ro')}</button>
    <button onclick="setEdit(${b.id},'readonly',0)" title="${t('btn_rw_tt')}">${t('btn_rw')}</button>
    <button onclick="startAdjust(${b.id})" title="${t('adj_btn_t')}">${t('adj_btn')}</button></span>`;
  if(edits[b.id]){
    const e=edits[b.id];const parts=[];
    if(e.user_type)parts.push(t('pend_type')+lbl(e.user_type));
    if(e.user_bbox)parts.push(t('adj_pending')+` r${e.user_bbox[0]}-r${e.user_bbox[2]} × c${e.user_bbox[1]}-c${e.user_bbox[3]}`);
    if(e.readonly!==undefined)parts.push(e.readonly?t('pend_ro'):t('pend_rw'));
    rows+=`<b>${t('pend_h')}</b><span>${parts.join('，')} <button onclick="revertEdit(${b.id})" title="${t('revert_tt')}">${t('revert_btn')}</button></span>`;
  }
  d.innerHTML=rows;
  if(ty==='table'){renderGlossary(b);renderRules()}
}

async function renderRules(){
  const box=document.getElementById('rulesBox');
  const cols=[...new Set((sel!==null&&S.blocks.find(b=>b.id===sel&&b.type!=='note')?.header?.col_tree||[])
    .filter(n=>n.path&&n.path.length).map(n=>n.path[n.path.length-1]))].slice(0,30);
  if(!cols.length){box.style.display='none';return}
  box.style.display='';
  document.getElementById('ruleCol').innerHTML=cols.map(c=>`<option>${esc(c)}</option>`).join('');
  const rs=await (await fetch('/api/rules?sheet='+encodeURIComponent(S.sheet))).json();
  document.getElementById('rulesList').innerHTML=rs.length?rs.map(r=>
    `· <b>${esc(r.col_path.split('/').pop())}</b> [${r.rmin??'-∞'},${r.rmax??'+∞'}] ${r.severity}`).join('<br>')
    :t('rules_none');
}
async function saveRule(){
  await fetch('/api/rules_save',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({sheet:S.sheet,column:document.getElementById('ruleCol').value,
      min:document.getElementById('ruleMin').value,max:document.getElementById('ruleMax').value,
      pattern:document.getElementById('rulePat').value.trim(),
      severity:document.getElementById('ruleSev').value})});
  renderRules();
}
async function delRule(){
  await fetch('/api/rules_delete',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({sheet:S.sheet,column:document.getElementById('ruleCol').value})});
  renderRules();
}
async function runRulesCheck(){
  const out=await (await fetch('/api/rules_check')).json();
  document.getElementById('rulesCheckOut').innerHTML=out.length?
    `<b style="color:#b91c1c">✗ ${out.length} ${t('rules_violations')}</b><br>`+out.slice(0,20).map(x=>
      `· [${esc(x.sheet)}] r${x.row} 「${esc(x.column)}」=${x.value} ${esc(x.问题)}(${x.severity})`).join('<br>')
    :`<span style="color:#15803d">${t('rules_all_pass')}</span>`;
}
async function renderGlossary(b){
  const h=b.header||{};
  const cols=[...new Set((h.col_tree||[]).filter(n=>n.path&&n.path.length)
    .map(n=>n.path[n.path.length-1]))].slice(0,25);
  let gl=[];
  try{gl=await (await fetch('/api/glossary?sheet='+encodeURIComponent(S.sheet))).json();}catch(e){}
  const box=document.getElementById('glossaryBox');
  if(!cols.length){box.style.display='none';return}
  box.style.display='';
  const byCol={};gl.forEach(g=>byCol[g.term]=g.meaning);
  box.querySelector('#glossaryList').innerHTML=cols.map(c=>`
    <div style="display:flex;gap:3px;align-items:center;margin:2px 0">
      <span style="flex:0 0 38%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"
            title="${esc(c)}">${esc(c)}</span>
      <input id="gl_${b.id}_${cols.indexOf(c)}" style="flex:1;font:11px system-ui;padding:2px 4px"
             placeholder="${t('gl_ph')}" value="${esc(byCol[c]||'')}">
      <button onclick="saveGlossary('${esc(c)}',${b.id},${cols.indexOf(c)})">${t('gl_save')}</button>
    </div>`).join('')+
    `<div class="hint" style="margin-top:4px">${t('gl_hint')}</div>`;
}
async function saveGlossary(col,bid,idx){
  const v=document.getElementById(`gl_${bid}_${idx}`).value.trim();
  if(!v)return;
  const r=await (await fetch('/api/glossary_save',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({sheet:S.sheet,column:col,meaning:v})})).json();
  if(r.ok)document.getElementById(`gl_${bid}_${idx}`).style.background='#dcfce7';
}
function nextType(t){return {table:'note',note:'mixed',mixed:'ignore',ignore:'table'}[t]}
function setEdit(id,k,v){
  edits[id]=edits[id]||{};
  if(!edits[id]._orig){const b0=S.blocks.find(x=>x.id===id);edits[id]._orig={type:b0?b0.type:'table',bbox:b0?b0.bbox:null}}
  edits[id][k]=v;
  if(k==='user_type'){const b=S.blocks.find(x=>x.id===id);if(b){b.type=v}}
  drawOverlays();updatePending();
  const b=S.blocks.find(x=>x.id===id);if(b)showDetail(b);
}
function revertEdit(id){          // 误触保护：撤销本块全部未冻结修改
  const e=edits[id];if(!e)return;
  const b=S.blocks.find(x=>x.id===id);
  if(b&&e._orig){b.type=e._orig.type;if(e._orig.bbox)b.bbox=e._orig.bbox}
  delete edits[id];
  drawOverlays();updatePending();
  if(b)showDetail(b);
}
function updatePending(){
  const n=Object.keys(edits).length;
  const chip=document.getElementById('pendingChip');
  chip.style.display=n?'':'none';
  chip.textContent=`${t('pending_pre')}${n}${t('pending_suf')}`;
  document.getElementById('freezeBtn').textContent=n?`${t('freeze_n_pre')}${n}${t('freeze_n_suf')}`:t('freeze');
}
async function doFreeze(){
  const list=[];
  S.blocks.forEach(b=>{if(edits[b.id]&&b.db_id){
    const {_orig,...rest}=edits[b.id];
    list.push({block_db_id:b.db_id,...rest})}});
  const r=await (await fetch('/api/freeze',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({edits:list,sheet:S.sheet})})).json();
  if(r.ok){
    edits={};updatePending();
    if(r.rejected&&r.rejected.length)
      alert(t('frozen_part')+'\n\n'+r.rejected.map(x=>'· '+x.reason).join('\n'));
    else alert(t('frozen_ok'));
    await loadState(S.sheet);refreshStatus();   // 几何/标记变了：重载确认页视图
  }else alert(t('frozen_fail')+(r.warning?'\n'+r.warning:''));
}
async function doDetach(blockId){
  const r=await (await fetch('/api/note_detach',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({sheet:S.sheet,block_id:blockId})})).json();
  if(r.ok){S=await (await fetch('/api/state?sheet='+encodeURIComponent(S.sheet))).json();
    sel=blockId;const b=S.blocks.find(x=>x.id===blockId);if(b)showDetail(b);drawOverlays();}
  else alert(t('detach_fail')+(r.reason||''));
}
function esc(s){return String(s).replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}

/* ===== 列宽：初次内容自适应 + 表头右缘拖拽 + 双击复位 ===== */
const _measure=(()=>{const c=document.createElement('canvas').getContext('2d');
  c.font='12px system-ui';return t=>Math.ceil(c.measureText(String(t)).width)})();
function colWKey(c){return (S&&S.workbook?S.workbook.path.slice(-40):'wb')+'|'+(S?S.sheet:'')+'|c'+c}
function loadColW(){try{return JSON.parse(localStorage.getItem('gewu-colw')||'{}')}catch(e){return{}}}
function saveColW(o){localStorage.setItem('gewu-colw',JSON.stringify(o))}
function autoFitCols(){
  const g=document.getElementById('grid');
  if(!g.rows.length)return;
  const [r1,c1,r2,c2]=S.range, hr=new Set(S.hidden_cols);
  const ov=loadColW();
  const widths={};let total=0;
  for(let c=c1;c<=c2;c++){
    if(hr.has(c)){widths[c]=40;total+=40;continue}
    let w=54;                                  // 列头宽底
    for(const tr of g.rows){for(const td of tr.cells){
      if(td.cellIndex-1===c-c1 && !td.classList.contains('hcol') && !td.classList.contains('hrow')){
        const t=td.textContent;if(!t)continue;
        w=Math.max(w,_measure(t)+18);          // 内容自适应（含 padding）
      }}}
    w=Math.min(w,320);
    const k=colWKey(c);
    if(ov[k])w=ov[k];                          // 手动调过的列尊重用户
    widths[c]=w;total+=w;
  }
  const tmp=document.createElement('table');    // colgroup 需在 table 上下文里解析
  tmp.innerHTML=`<colgroup><col style="width:44px">`+
    Array.from({length:c2-c1+1},(_,i)=>`<col data-c="${c1+i}" style="width:${widths[c1+i]}px">`).join('')+`</colgroup>`;
  g.insertBefore(tmp.firstChild,g.firstChild);
  g.querySelectorAll('td.hrow').forEach(td=>{td.style.width='44px'});
  g.style.width=(44+total)+'px';
  // 表头挂拖拽/双击复位
  g.querySelectorAll('td.hcol:not(.hrow)').forEach(td=>{
    td.style.cursor='col-resize';
    td.onmousedown=e=>{ if(e.offsetX>=td.offsetWidth-6) dragCol(td,e) };
    td.ondblclick=()=>{ const c=c1+td.cellIndex-1,k=colWKey(c),ov=loadColW();
      delete ov[k];saveColW(ov);autoFitCols();drawOverlays(); };
  });
}
function dragCol(td,e0){
  const c=S.range[1]+td.cellIndex-1,k=colWKey(c);
  const x0=e0.clientX,w0=td.offsetWidth;
  const mv=e=>{const w=Math.max(40,Math.min(600,w0+e.clientX-x0));
    document.querySelector(`#grid colgroup col[data-c="${c}"]`).style.width=w+'px';
    document.getElementById('grid').style.width=(parseInt(document.getElementById('grid').style.width)-w0+w)+'px';_w=w};
  let _w=w0;
  const up=()=>{window.removeEventListener('mousemove',mv);window.removeEventListener('mouseup',up);
    const ov=loadColW();ov[k]=_w;saveColW(ov);drawOverlays()};
  window.addEventListener('mousemove',mv);window.addEventListener('mouseup',up);
  e0.preventDefault();
}

function mdRich(s){
  /* 富文本渲染：先转义再逐级替换（安全：输入先 esc）。
     支持：**加粗** · `行内代码` · ```代码块``` · =公式= · > 引用 · - 列表 · 数字列表 · 表格行 | */
  let t=esc(s);
  // 代码块（三反引号）
  t=t.replace(/```(\w*)\n?([\s\S]*?)```/g,(m,lang,code)=>{
    return `<pre class="mdPre"><code>${code.trim()}</code></pre>`});
  // Excel 公式：=SUM(...) / =(A1+B2)/C3 / ==SUM(A1:A10)——等宽蓝底
  t=t.replace(/={1,2}\(?[A-Z][A-Z0-9]*(?:\([^)]*\))?[^<\s=]{0,40}\)?/g,
    m=>{ if(m.length<3||m.length>100)return m;
         return `<span class="mdFormula">${m}</span>`});
  // 行内代码
  t=t.replace(/`([^`\n]+)`/g,'<code class="mdCode">$1</code>');
  // 加粗
  t=t.replace(/\*\*([^*\n]+)\*\*/g,'<b class="mdBold">$1</b>');
  // 列表行（- 开头）→ 圆点
  t=t.replace(/^[\s]*[-·]\s+(.+)$/gm,'<span class="mdLi">· $1</span>');
  // 数字列表行（1. 2. 开头）
  t=t.replace(/^[\s]*(\d+)[.、]\s+(.+)$/gm,'<span class="mdOl"><b>$1.</b> $2</span>');
  // 引用行（> 开头；esc 后 &gt;）
  t=t.replace(/^[\s]*(?:>|&gt;)\s*(.+)$/gm,'<span class="mdQuote">$1</span>');
  // 表格行（| 分隔）→ 简单表格样式
  t=t.replace(/^[\s]*(\|[^|\n]+\|[^|\n]*)$/gm,'<span class="mdTr">$1</span>');
  // 换行保留
  t=t.replace(/\n/g,'<br>');
  return t;
}
// 兼容旧调用
function mdLite(s){return mdRich(s)}

/* ===== 双语字典（zh 为默认原文；en 为英文版）===== */
const I18N={
 "zh":{
  brand_tag:"复杂表格理解工作台", mask:"蒙版", opacity:"浓度",
  freeze:"✔ 检查无误，冻结", help_t:"注意事项", about_t:"关于格物", open_t:"打开文件",
  export_btn:"📥 导出改后文件", export_t:"下载改后的文件（副本模式的改动只有导出才是成品）",
  open_title:"📂 打开文件", open_path_h:"① 文件路径",
  open_path_hint:"在资源管理器按住 Shift 右键文件 →「复制文件地址」粘贴到这。改数直接写回这个原文件。",
  open_ph:"C:\\Users\\...\\表.xlsx", open_do:"打开",
  open_copy_h:"② 上传副本",
  open_copy_note:"（分析体验用；改数只写副本——成品须点页眉「📥 导出改后文件」下载，别把原上传文件当成品发人）",
  open_up:"上传并打开", open_err_empty:"请输入文件路径", open_parsing:"解析中…（大文件可能十几秒）",
  drop_pre:"📥 收到 ", drop_suf:"：网页拖入=上传副本（改动只进副本，成品用页眉「导出改后文件」）；要写回原文件请用 📂 路径打开。上传解析中…", collapse:"收起面板（表格全宽）",
  tab_review:"📋 审阅", tab_chat:"💬 问答", tab_write:"✏️ 改数", tab_history:"🕘 历史",
  review_h:"区块审阅", review_hint:"点色块查看详情；不确定怎么操作就点右上 ❓——大多数文件直接冻结即可。",
  glossary:"📖 字段口径词典", scope:"范围", whole_file:"整个文件",
  sess_hist:"历史会话列表", sess_new:"新对话", cur_title_empty:"（提问后自动建会话）",
  chat_hint:"问数值（合计/最高的）→ 引擎计算；问理解（这表是做什么的/为什么这么算）→ LLM 解读；问评价（分析一下）→ 数据解读。追问自动带最近 4 轮上下文。",
  ask_ph:"问：销售额合计 / 这表是做什么的 / 分析评价一下 （回车发送）", ask_btn:"问",
  write_h:"改数", write_note:"：用日常中文描述要改什么，系统会定位目标位置并计算新值，确认后才写入原文件（自动备份，支持撤销）。写入范围不受对话所选 sheet 限制。",
  write_ph:"自然语言：把销售部1月差旅费的金额改成20000\n或手写计划 JSON：{\"find\":{\"keys\":{\"列\":\"值\"}},\"set\":{\"column\":\"列\",\"value\":7}}",
  btn_nl:"💬 自然语言", btn_plan:"📋 计划 JSON", btn_undo:"↩ 撤销一批",
  diff_hint:"语义解析 → 硬校验（冻结/表头/汇总行等全不许碰）→ diff 预览 → 确认执行",
  write_log:"最近写入", write_log_empty:"（切到本页签时加载）", write_log_none:"（本文件还没有写入记录）",
  hist_h:"本文件的会话", hist_clear:"清空全部", hist_clear_confirm:"确定清空本文件全部会话？此操作不可撤销。", hist_search_h:"🔎 搜全部文件的历史",
  hist_ph:"关键词（跨工作簿检索，只读参考）", hist_none:"（本文件尚无会话）", hist_empty:"（尚无会话，提问后自动创建）",
  status_parsed:"未冻结（解析态）", status_frozen_pre:"已冻结 ", undone:"（已撤销）",
  pending_pre:"✏️ 待提交修改 ", pending_suf:" 处（点冻结才生效；刷新页面=全部撤销）",
  freeze_n_pre:"✔ 检查无误，冻结（含 ", freeze_n_suf:" 处修改）",
  frozen_ok:"已冻结：语义镜像已写入 mirror.db，下游问答/写回将只认此版本", frozen_fail:"失败",
  new_chat:"已开新对话。", sess_deleted:"会话已删除。", sess_empty:"（空会话）",
  chg_struct_pre:"⚠ <b>文件内容有变化（结构级）</b>：新增 ", chg_struct_mid:" 块、消失 ", chg_struct_mid2:" 块、结构变动 ", chg_struct_mid3:" 块。",
  chg_added:" 新增：", chg_removed:" 消失：", chg_tail:" 未变化块的人工标记已自动保留（", chg_tail2:" 处），<b>请复核变化后重新冻结</b>；冻结前改数仍被拦截。",
  chg_value:"✅ 检测到文件数值更新（结构无变化）：冻结的表结构判定仍然有效，数值已自动刷新。",
  hit_row:"命中行：", dep_pre:"本次写入将联动重算 ", dep_suf:" 个公式格（Excel 自动重算）",
  dep_done_pre:"联动公式格已重算：", dep_done_suf:"（切回表格可见新值）",
  rules_h:"📏 业务规则（值域）", rules_note:"给数值列设上下限：block=越界拒写，warn=仅警告。规则由引擎硬执行，LLM 只知情。",
  rules_save:"保存", rules_del:"删除", rules_check:"全表体检",
  rules_check_note:"按现有数据跑一遍规则，列出违规行",
  rules_min_ph:"≥", rules_max_ph:"≤",
  d_type:"类型", d_pos:"位置", d_score:"代码置信", d_hdr:"表头行",
  d_title:"标题", d_float:"浮动草稿区", d_summary:"⚠汇总行", d_dups:"⚠同名列",
  d_cross:"交叉表", d_fields:"字段", d_signal:"信号", d_anno:"批注",
  user_q:"问：", v_pass:"✓ 硬校验全部通过", v_reject:"✗ 拒绝：",
  v_confirm:"确认写入（COM，自动备份）", v_cancel:"取消", v_cancelled:"已取消",
  hist_no_hit:"（没有命中）", load_on_switch:"（切到本页签时加载）",
  rules_none:"（本 sheet 尚无规则）", rules_all_pass:"✓ 现有数据全部符合规则",
  rules_violations:"处违规：", no_hits:"没有命中记录", ask_empty:"请输入问题",
  parsing:"解析中…", upload_fail:"上传失败：", switch_fail:"切换失败:",
  btn_nl_t:"需在 ⚙️ 配置 LLM（本地模型亦可）", dep_unchanged:"⚠ 预测的联动格缓存值未变化——若预期应重算，请打开 Excel 确认", rules_pat_ph:"正则(选填)",
  llm_t:"LLM 设置", llm_h:"⚙️ LLM 设置", llm_close:"关闭 ✕", llm_preset_h:"预设",
  llm_proto:"协议", llm_base:"接口地址", llm_model:"模型名", llm_key:"API Key",
  llm_privacy:"密钥只保存在本机 llm.json，请求直连你所选的服务商；不配置则问答走内置纯代码引擎。",
  llm_save:"保存", llm_test:"测试连接", llm_testing:"测试中…",
  llm_saved:"已保存（新对话生效）", llm_none:"（未配置：问答走内置规则引擎）",
  tray_label:"关闭窗口时缩到系统托盘（引擎保持运行，重开秒进；取消勾选=直接退出）",
  tray_note:"仅桌面应用生效，改后重启应用即可。", copy_badge:"副本",
  guide:'<b>怎么用：</b><b>蓝框=表格</b>（问答和改数都发生在蓝框里，只需扫一眼框得对不对）；<b>黄虚线=标题/备注</b>（不当表格处理，<u>不用管它</u>；标了「表内批注」的是表格中间夹的一行备注，正常现象）；<b>红框=机器拿不准</b>（最好点开看一眼）。没问题 → 点右上「检查无误，冻结」→ 右侧面板开始问答与改数。框不对 → 点色块改类型；范围不对 → 选中后直接拖蓝框边角调整。<button onclick="this.parentNode.remove()" style="margin-left:auto">知道了 ✕</button>',
  adj_title:'调整范围（区块',adj_hint:'直接拖橙框的边/角，拖「✥ 移动整框」钮可平移位置，或在表格上点两下选范围，或改数字。确定后随「冻结」生效，表头和列名会在新范围里自动重新识别。',
  adj_move:'✥ 移动整框',adj_move_tt:'按住拖动可平移整个框的位置（不改变大小）',
  adj_r1:'行起',adj_c1:'列起',adj_r2:'行止',adj_c2:'列止',
  adj_ok2:'确定调整',adj_cancel:'取消',adj_prev:'新范围预览',
  adj_bad:'范围不合法：行止≥行起、列止≥列起，且都≥1',
  adj_btn:'📐 调整范围',adj_btn_t:'框盖小了/劈开了/圈多了？自己重画这块的范围；表头树和记录会在新范围里由机器重新识别，点冻结生效。与其他未忽略区块重叠会拦下',
  adj_pending:'范围调整为',d_orig:'引擎原范围',frozen_part:'部分修改未生效（其余已冻结）：',
  lbl_table:'🟦数据表',lbl_note:'🟨文本/批注',lbl_mixed:'🟥混合待复核',lbl_ignore:'⬜忽略',
  nested_note:'表内批注',d_summary_prot:'（写保护）',d_leaves:'个叶子列',d_anno_v:'是（已从表格剥离）',
  attach_h:'行内备注',attach_full:'挂到 表#{id} 第{r}行',detach_btn:'取消挂接（不算这行的信息）',
  switch_tt:'把这块按另一种类型对待；点冻结才生效，可反复切换',
  switch_tt_note:'该块引擎判定为文本（无列结构）：切为数据表仅作视觉标记，不会进入问答',
  btn_switch:'切为',btn_ro:'🔒只读',btn_rw:'🔓可写',
  btn_ro_tt:'应用内只读：本工具禁止改这块的数（问答不受影响）；文件本身不变，Excel/WPS 里仍可编辑',
  btn_rw_tt:'解除只读',
  ops_h:'操作',pend_h:'已改(未冻结)',pend_type:'切为',pend_ro:'设为只读🔒',pend_rw:'设为可写🔓',
  revert_btn:'↩ 撤销本块修改',revert_tt:'撤销这块的全部未冻结修改',
  gl_ph:'口径注释…',gl_save:'存',gl_hint:'口径注释会自动注入问答与写指令的模型上下文（同名不同口径的列靠它消歧）',
  glossary_init:'选中数据表区块后可给字段加口径注释',rules_init:'（选中数据表后可配置）',
  m_hidden_sheet:'隐', m_copy_badge:'〔副本〕', m_hidden_cell:'⟨隐⟩',
  open_fail:'打开失败', uploading:'上传中…', upload_fail:'上传失败',
  d_yes:'是', d_sig_text:'文本', d_sig_num:'数值', d_sig_border:'边框',
  detach_fail:'取消失败：', s_untitled:'(未命名)', m_sheets:'表', s_del_tt:'删除本会话',
  m_q:'问：', tag_fb:'⚙️(LLM挂,规则兜底)',
  wp_none:'未生成写计划', wp_cand:'候选：', wp_copy_note:'⚠ <b>副本模式</b>：本次写入只进副本。要发给别人的成品，写完点页眉<b>「📥 导出改后文件」</b>下载，别发原上传的文件。',
  wp_title:'✏️ 写计划', wp_locate:'📍 定位到表格',
  v_allpass:'✓ 硬校验全部通过', v_reject:'✗ 拒绝：',
  wm_set:'改为',wm_add:'加',wm_sub:'减',wm_mul:'乘',wm_div:'除',wm_add_pct:'涨%',wm_sub_pct:'降%',
  wr_ok_pre:'✓ 已写入（',wr_ok_mid:'，备份 ',wr_ok_suf:'）写后校验通过。',
  wr_copy_hint:'改动在副本上——点此拿到成品：', wr_undo_hint:'出错了可点「撤销最近写入」。',
  wr_fail:'✗ 执行失败：', wr_undone_pre:'↩ 已撤销 ',wr_undone_suf:' 格（',wr_undone_end:'）', wr_none:'无可撤销',
  boot_fail:'⚠ 解析失败：', boot_unknown:'未知原因', boot_hint:'请确认文件是有效的 .xlsx（.xls/.csv 暂不支持），或换一个文件重试。',
  llm_preset_ph:'— 选择预设自动填充 —', llm_key_ph:'sk-...（本地模型可留空）',
  panel_expand_tt:'展开面板', panel_drag_tt:'拖动调宽窄', scope_hint:'点击选择范围（可多选 sheet）',
  sess_none:'（尚无会话）', ask_btn:'问', search_btn:'搜',
  doc_title:'格物 Gewu · 复杂表格理解工作台', open_path_note:'（写回原文件，推荐）'
 },
 "en":{
  brand_tag:"Intelligence workbench for complex spreadsheets", mask:"Mask", opacity:"Opacity",
  freeze:"✔ All good — Freeze", help_t:"Notes", about_t:"About Gewu", open_t:"Open file",
  export_btn:"📥 Export edited file", export_t:"Download the edited file (in copy mode, only the export is the finished product)",
  open_title:"📂 Open file", open_path_h:"1. File path",
  open_path_hint:"Shift + right-click the file in Explorer → “Copy as path”, then paste here. Edits write back to this original file.",
  open_ph:"C:\\Users\\...\\file.xlsx", open_do:"Open",
  open_copy_h:"2. Upload a copy",
  open_copy_note:"(for analysis; edits go to the copy only — the finished product must be downloaded via “📥 Export edited file”; never send the originally uploaded file as your result)",
  open_up:"Upload & open", open_err_empty:"Please enter a file path", open_parsing:"Parsing… (large files may take a while)",
  drop_pre:"📥 Received ", drop_suf:": dropping onto the web page uploads a COPY (edits stay in the copy; use “Export edited file” for the result). To edit the original, open it by path via 📂. Uploading…", collapse:"Collapse panel (full-width grid)",
  tab_review:"📋 Review", tab_chat:"💬 Chat", tab_write:"✏️ Edit", tab_history:"🕘 History",
  review_h:"Block review", review_hint:"Click a colored block for details; most files can be frozen directly — see ❓ if unsure.",
  glossary:"📖 Field glossary", scope:"Scope", whole_file:"Whole file",
  sess_hist:"Session history", sess_new:"New chat", cur_title_empty:"(session starts on first question)",
  chat_hint:"Numbers (sum / top) → engine computes · Understanding (what is this table / why this formula) → LLM explains · Evaluation (analyze this) → data interpretation. Follow-ups carry the last 4 turns.",
  ask_ph:"Ask: total of sales / what is this table / analyze this  (Enter to send)", ask_btn:"Ask",
  write_h:"Edit", write_note:"Say which row & column → engine computes the new value → you review the diff → confirm to write (auto-backup, undoable). Edits ignore chat scope and locate semantically across the workbook.",
  write_ph:"Natural language: set Sales of Project X in Jan to 7\nOr plan JSON: {\"find\":{\"keys\":{\"col\":\"value\"}},\"set\":{\"column\":\"col\",\"value\":7}}",
  btn_nl:"💬 Natural language", btn_plan:"📋 Plan JSON", btn_undo:"↩ Undo batch",
  diff_hint:"Semantic parse → hard validation (frozen/header/summary rows are protected) → diff preview → confirm to execute",
  write_log:"Recent writes", write_log_empty:"(loads when this tab opens)", write_log_none:"(no writes yet for this file)",
  hist_h:"Sessions of this file", hist_clear:"Clear all", hist_clear_confirm:"Clear ALL sessions of this file? This cannot be undone.", hist_search_h:"🔎 Search history across all files",
  hist_ph:"Keyword (cross-workbook, read-only)", hist_none:"(no sessions yet for this file)", hist_empty:"(no sessions yet — created on first question)",
  status_parsed:"Not frozen (parsed)", status_frozen_pre:"Frozen ", undone:" (undone)",
  pending_pre:"✏️ Pending edits: ", pending_suf:" (applied on freeze; refresh discards all)",
  freeze_n_pre:"✔ All good — Freeze (", freeze_n_suf:" edits)",
  frozen_ok:"Frozen: the semantic mirror is now the single source of truth for Q&A and write-back", frozen_fail:"Failed",
  new_chat:"New conversation started.", sess_deleted:"Session deleted.", sess_empty:"(empty session)",
  chg_struct_pre:"⚠ <b>File changed (structural)</b>: ", chg_struct_mid:" added, ", chg_struct_mid2:" removed, ", chg_struct_mid3:" changed blocks.",
  chg_added:" Added: ", chg_removed:" Removed: ", chg_tail:" Manual marks on unchanged blocks were kept (", chg_tail2:"). <b>Please review and re-freeze</b>; edits stay blocked until then.",
  chg_value:"✅ Values updated (no structural change): frozen structure remains valid; numbers refreshed automatically.",
  hit_row:"Matched row:", dep_pre:"This edit will recalculate ", dep_suf:" dependent formula cells (Excel recalcs automatically)",
  dep_done_pre:"Dependent formulas recalculated: ", dep_done_suf:" (switch to the grid to see new values)",
  rules_h:"📏 Business rules (value ranges)", rules_note:"Set min/max for numeric columns: block = reject out-of-range writes, warn = warn only. Rules are engine-enforced; the LLM is only informed.",
  rules_save:"Save", rules_del:"Delete", rules_check:"Check all data",
  rules_check_note:"Run rules against existing data and list violations",
  rules_min_ph:"min", rules_max_ph:"max",
  d_type:"Type", d_pos:"Position", d_score:"Confidence", d_hdr:"Header rows",
  d_title:"Title", d_float:"Floating zone", d_summary:"⚠Summary rows", d_dups:"⚠Dup columns",
  d_cross:"Cross-tab", d_fields:"Fields", d_signal:"Signals", d_anno:"Annotation",
  user_q:"Q:", v_pass:"✓ All validations passed", v_reject:"✗ Rejected: ",
  v_confirm:"Confirm write (COM, auto-backup)", v_cancel:"Cancel", v_cancelled:"Cancelled",
  hist_no_hit:"(no results)", load_on_switch:"(loads when tab opens)",
  rules_none:"(no rules for this sheet)", rules_all_pass:"✓ All data passes rules",
  rules_violations:" violations: ", no_hits:"No matching records", ask_empty:"Please enter a question",
  parsing:"Parsing…", upload_fail:"Upload failed: ", switch_fail:"Switch failed: ",
  btn_nl_t:"Requires LLM in ⚙️ (local models also work)", dep_unchanged:"⚠ Predicted dependent cells' cached values unchanged — open Excel to verify if recalc was expected", rules_pat_ph:"regex (optional)",
  llm_t:"LLM settings", llm_h:"⚙️ LLM Settings", llm_close:"Close ✕", llm_preset_h:"Preset",
  llm_proto:"Protocol", llm_base:"Base URL", llm_model:"Model", llm_key:"API Key",
  llm_privacy:"The key is stored only in the local llm.json; requests go directly to your chosen provider. Without a key, Q&A runs on the built-in rules engine.",
  llm_save:"Save", llm_test:"Test connection", llm_testing:"Testing…",
  llm_saved:"Saved (takes effect on next question)", llm_none:"(not configured: Q&A uses the built-in rules engine)",
  tray_label:"Minimize to system tray on close (engine keeps running, instant reopen; uncheck = quit directly)",
  tray_note:"Desktop app only. Restart the app to apply.", copy_badge:"copy",
  help_full:`<div style="display:flex;align-items:center;border-bottom:2px solid #e2e8f0;padding-bottom:8px;margin-bottom:4px"><b style="font-size:16.5px">Notes</b><button onclick="document.getElementById('helpBox').style.display='none'" style="margin-left:auto">Close ✕</button></div>
  <div style="margin:12px 0 4px;font-weight:700;color:#1d4ed8">1. What freezing means</div>
  <div style="padding-left:10px">Freeze = you confirm "the machine's frames are right", commit your manual fixes, and open the master switch for <b>cell edits</b>.<br><b>Not frozen = read-only mode</b>: Q&amp;A / table interpretation / tie-checking all work; every edit is blocked and the file is absolutely safe.</div>
  <div style="margin:14px 0 4px;font-weight:700;color:#1d4ed8">2. How to ask</div>
  <div style="padding-left:10px">"This table / current table" refers to the sheet you are viewing; "the whole file / how do these tables relate" asks about the whole workbook.<br>Follow-ups (e.g. "then why is it computed that way") automatically carry the last 4 turns; click <b>🗑 New chat</b> to start over.</div>
  <div style="margin:14px 0 4px;font-weight:700;color:#1d4ed8">3. How to edit cells</div>
  <div style="padding-left:10px">Edits <b>do not follow the chat scope</b> — they locate semantically across the whole workbook by sheet keyword + row keys + column name, so include the sheet name ("add 3000 to the January travel budget of <b>Sales</b>"). Multiple hits are blocked until clarified.<br>Flow: semantic parse → hard validation → <b>diff preview (old → new)</b> → your confirmation → write (auto-backup, undoable). In uploaded-copy mode edits go to the copy only — download the finished file via “📥 Export edited file” in the header.</div>
  <div style="margin:14px 0 4px;font-weight:700;color:#1d4ed8">4. Block buttons (effective after freeze)</div>
  <table style="width:100%;border-collapse:collapse;font-size:13px;margin:4px 0"><tr style="background:#f1f5f9"><th style="text-align:left;padding:4px 8px;border:1px solid #e2e8f0">Button</th><th style="text-align:left;padding:4px 8px;border:1px solid #e2e8f0">Effect</th></tr><tr><td style="padding:4px 8px;border:1px solid #e2e8f0">→ 🟨 Text/note</td><td style="padding:4px 8px;border:1px solid #e2e8f0">The block leaves Q&amp;A and editing (data stays in the file; just not treated as a table)</td></tr><tr><td style="padding:4px 8px;border:1px solid #e2e8f0">→ ⬜ Ignore</td><td style="padding:4px 8px;border:1px solid #e2e8f0">Same as above: out of Q&amp;A + edits rejected</td></tr><tr><td style="padding:4px 8px;border:1px solid #e2e8f0">→ 🟥 Mixed review</td><td style="padding:4px 8px;border:1px solid #e2e8f0">Isolated: no edits, no Q&amp;A, until re-typed as table or text</td></tr><tr><td style="padding:4px 8px;border:1px solid #e2e8f0">🔒 Read-only</td><td style="padding:4px 8px;border:1px solid #e2e8f0"><b>In-app only</b>: this tool refuses to edit the block (Q&amp;A unaffected). <b>The file itself is untouched</b> — still editable in Excel/WPS, file attributes unchanged</td></tr><tr><td style="padding:4px 8px;border:1px solid #e2e8f0">🔓 Writable</td><td style="padding:4px 8px;border:1px solid #e2e8f0">Clears the in-app read-only mark</td></tr><tr><td style="padding:4px 8px;border:1px solid #e2e8f0">📐 Adjust range</td><td style="padding:4px 8px;border:1px solid #e2e8f0">Select a block and drag its frame edges/corners directly (panel offers orange-frame drag / numbers / two clicks); headers and records are <b>re-detected automatically</b> inside the new range, applied on freeze. Overlaps with a non-ignored block are rejected (set that block to Ignore first)</td></tr></table>
  <div style="padding-left:10px;color:#475569;font-size:12.5px">Accidental clicks: edits apply only on freeze (a "pending N" chip shows at the top); refreshing before freeze discards everything; each block has "↩ revert"; if you froze something wrong, switch it again and re-freeze.<br>Note: switching an engine-detected text block to "table" is a visual mark only (no column structure, never enters Q&amp;A).</div>
  <div style="margin:14px 0 4px;font-weight:700;color:#1d4ed8">5. After the file changes</div>
  <div style="padding-left:10px">Edits made by this tool → reopening continues seamlessly (freeze kept).<br>Someone changed values → re-drop the file: green bar "values refreshed", freeze still valid, no re-confirmation.<br>Structure changed (tables added/removed, headers edited) → re-drop: orange bar lists the changes; marks on unchanged blocks are kept automatically; review then re-freeze.</div>
  <div style="margin-top:14px;padding:8px 10px;background:#f8fafc;border-left:3px solid #94a3b8;color:#64748b;font-size:12.5px">Every edit is re-validated before writing; headers, summary rows, formula cells and ignored blocks are never written.   <div style="margin:14px 0 4px;font-weight:700;color:#1d4ed8">6. Opening files & copies</div>
  <div style="padding-left:10px">
    Three ways to open: <b>📂 path</b> (Shift+right-click → Copy as path — writes back to the original, recommended);
    <b>upload a copy</b> (for analysis — edits go to the copy only, a〔copy〕badge appears in the header);
    drop a file onto the desktop app icon / into the desktop window (= original); drop onto the web page (= uploads a copy).<br>
    <b>In copy mode the finished file = the one downloaded via “📥 Export edited file”</b> — the originally uploaded file never contains your edits; don't send it to anyone.
  </div>
  <div style="margin-top:14px;padding:8px 10px;background:#f8fafc;border-left:3px solid #94a3b8;color:#64748b;font-size:12.5px">
    Every edit is re-validated before writing; headers, summary rows, formula cells and ignored blocks are never written. All manual marks (types/read-only/glossary/row notes) live only in this tool's mirror and <b>never modify your file</b> (except explicit edit commands).</div>`,
   guide:'<b>How to use:</b> <b>Blue = table</b> (Q&A and edits happen inside blue frames — just glance they are right); <b>dashed yellow = title/notes</b> (not treated as a table; blocks tagged “in-table note” are remark rows inside a table — normal); <b>red = machine unsure</b> (worth a look). All good → click “Freeze” top-right → start in the right panel. Wrong frame → click the block and switch its type; wrong range → select it and drag the frame\u2019s edges/corners. <button onclick="this.parentNode.remove()" style="margin-left:auto">Got it ✕</button>',
   adj_title:'Adjust range (block',adj_hint:'Drag the orange frame by its edges or corners, use the "✥ Move" pill to slide it, click two corners on the grid, or edit the numbers. Applies on Freeze, with headers and columns re-detected inside the new range.',
   adj_move:'✥ Move',adj_move_tt:'Hold and drag to slide the whole frame (size unchanged)',
   adj_r1:'Row from',adj_c1:'Col from',adj_r2:'Row to',adj_c2:'Col to',
   adj_ok2:'Apply',adj_cancel:'Cancel',adj_prev:'New range preview',
   adj_bad:'Invalid range: row-to ≥ row-from, col-to ≥ col-from, all ≥ 1',
   adj_btn:'📐 Adjust range',adj_btn_t:'Frame too small, split, or too wide? Redraw it; the header tree and records are re-detected inside the new range. Applies on Freeze. Overlaps with another non-ignored block are rejected.',
   adj_pending:'Range set to',d_orig:'Engine range',frozen_part:'Some edits were rejected (the rest is frozen):',
   lbl_table:'🟦 Table',lbl_note:'🟨 Text/note',lbl_mixed:'🟥 Mixed review',lbl_ignore:'⬜ Ignore',
   nested_note:'in-table note',d_summary_prot:' (write-protected)',d_leaves:'leaf columns',d_anno_v:'Yes (carved out of the table)',
   attach_h:'In-row note',attach_full:'attached to block #{id}, row {r}',detach_btn:'Detach (not counted as this row\u2019s info)',
   switch_tt:'Treat this block as the other type; applies on freeze, switchable anytime',
   switch_tt_note:'Engine classified this as text (no column structure): switching to table is a visual mark only, never enters Q&A',
   btn_switch:'Switch to ',btn_ro:'🔒 Read-only',btn_rw:'🔓 Writable',
   btn_ro_tt:'In-app read-only: this tool refuses to edit the block (Q&A unaffected). The file itself is untouched — still editable in Excel/WPS',
   btn_rw_tt:'Clear the read-only mark',
   ops_h:'Actions',pend_h:'Changed (not frozen)',pend_type:'switch to ',pend_ro:'set read-only 🔒',pend_rw:'set writable 🔓',
   revert_btn:'↩ Revert this block',revert_tt:'Revert all unfrozen changes of this block',
   gl_ph:'definition…',gl_save:'Save',gl_hint:'Definitions are injected into the LLM context of Q&A and edit commands (disambiguates same-named columns)',
   glossary_init:'Select a table block to add field definitions',rules_init:'(select a table block to configure)',
   m_hidden_sheet:'hidden', m_copy_badge:' [COPY]', m_hidden_cell:'⟨H⟩',
   open_fail:'Failed to open', uploading:'Uploading…', upload_fail:'Upload failed',
   d_yes:'Yes', d_sig_text:'text ', d_sig_num:'nums ', d_sig_border:'border ',
   detach_fail:'Detach failed: ', s_untitled:'(untitled)', m_sheets:' sheets', s_del_tt:'Delete this session',
   m_q:'Q: ', tag_fb:'⚙️(rules fallback)',
   wp_none:'No write plan generated', wp_cand:'Candidates:', wp_copy_note:'⚠ <b>Copy mode</b>: this write goes to the copy only. For the finished file to share, click <b>“📥 Export edited file”</b> in the header when done — do not send the originally uploaded file.',
   wp_title:'✏️ Write plan', wp_locate:'📍 Locate in table',
   v_allpass:'✓ All hard validations passed', v_reject:'✗ Rejected: ',
   wm_set:'set to',wm_add:'add',wm_sub:'subtract',wm_mul:'multiply',wm_div:'divide',wm_add_pct:'increase %',wm_sub_pct:'decrease %',
   wr_ok_pre:'✓ Written (',wr_ok_mid:', backup ',wr_ok_suf:') — post-write validation passed.',
   wr_copy_hint:'Changes are on the copy — get the finished file here: ', wr_undo_hint:'If something went wrong, use “Undo recent write”.',
   wr_fail:'✗ Execution failed: ', wr_undone_pre:'↩ Undone ',wr_undone_suf:' cells (',wr_undone_end:')', wr_none:'Nothing to undo',
   boot_fail:'⚠ Parse failed: ', boot_unknown:'unknown reason', boot_hint:'Make sure the file is a valid .xlsx (.xls/.csv not supported yet), or try another file.',
   llm_preset_ph:'— pick a preset to autofill —', llm_key_ph:'sk-... (optional for local models)',
   panel_expand_tt:'Expand panel', panel_drag_tt:'Drag to resize', scope_hint:'Click to pick scope (multi-sheet allowed)',
   sess_none:'(no sessions yet)', ask_btn:'Ask', search_btn:'Search',
   doc_title:'Gewu · Spreadsheet understanding workbench', open_path_note:'(writes back to the original file — recommended)'
 }
};
let LANG=localStorage.getItem('gewu-lang')||'zh';
function t(k){const d=I18N[LANG]&&I18N[LANG][k];return d!==undefined?d:(I18N.zh[k]??k)}
function applyLang(){
  document.documentElement.lang=LANG==='en'?'en':'zh';
  document.title=t('doc_title');
  document.querySelectorAll('[data-i18n]').forEach(el=>{el.textContent=t(el.dataset.i18n)});
  document.querySelectorAll('[data-i18n-html]').forEach(el=>{const k=el.dataset.i18nHtml;
    if(!el.dataset.orig)el.dataset.orig=el.innerHTML;      // zh=DOM 原文，避免长文重复维护
    el.innerHTML=(LANG==='zh')?el.dataset.orig:(I18N.en[k]??el.dataset.orig);});
  document.querySelectorAll('[data-i18n-t]').forEach(el=>{el.title=t(el.dataset.i18nT)});
  document.querySelectorAll('[data-i18n-ph]').forEach(el=>{el.placeholder=t(el.dataset.i18nPh)});
  document.getElementById('brandTag').textContent=t('brand_tag');
  refreshStatus();
}
function toggleLang(){LANG=LANG==='zh'?'en':'zh';localStorage.setItem('gewu-lang',LANG);applyLang();loadSessions();if(MODE==='write')loadWriteLog();if(MODE==='history')loadHistList()}

let MODE='review', SHEETS=[], SCOPE=null, CUR_SESSION=null;
function showPane(name){
  MODE=name;
  document.querySelectorAll('.ptab').forEach(b=>b.classList.toggle('on',b.dataset.p===name));
  document.querySelectorAll('.pane').forEach(m=>m.classList.toggle('on',m.id==='p-'+name));
  if(name==='write')loadWriteLog();
  if(name==='history')loadHistList();
}
document.querySelectorAll('.ptab').forEach(btn=>{btn.onclick=()=>showPane(btn.dataset.p)});
document.getElementById('panelCollapse').onclick=()=>{
  document.getElementById('sidePanel').style.display='none';
  document.getElementById('panelHandle').style.display='none';
  document.getElementById('panelExpand').style.display='';};
document.getElementById('panelExpand').onclick=()=>{
  document.getElementById('sidePanel').style.display='';
  document.getElementById('panelHandle').style.display='';
  document.getElementById('panelExpand').style.display='none';};
(function(){  // 拖动调面板宽窄
  const h=document.getElementById('panelHandle'),sp=document.getElementById('sidePanel');
  let drag=false;
  h.onmousedown=e=>{drag=true;e.preventDefault()};
  window.addEventListener('mousemove',e=>{if(!drag)return;
    const w=window.innerWidth-e.clientX;
    sp.style.width=Math.min(Math.max(w,300),window.innerWidth*0.7)+'px';});
  window.addEventListener('mouseup',()=>drag=false);
})();
function toggleSessPop(){
  const p=document.getElementById('sessPop');
  // CSS 初始 display:none + JS 空串切换会永远隐藏——用显式 block/none
  p.style.display=p.style.display==='block'?'none':'block';
}
function renderScope(){
  const box=document.getElementById('scopeChips');
  let html=`<span class="scopeChip${SCOPE===null?' on':''}" onclick="setScope(null)">${t('whole_file')}</span>`;
  html+=SHEETS.filter(s=>s.visible).map(s=>
    `<span class="scopeChip${SCOPE&&SCOPE.includes(s.name)?' on':''}" onclick="toggleScope('${esc(s.name)}')">${esc(s.name)}</span>`).join('');
  box.innerHTML=html;
  // 折叠按钮上的标签
  const lbl=document.getElementById('scopeLabel');
  if(!lbl)return;
  if(SCOPE===null)lbl.textContent=t('whole_file');
  else if(SCOPE.length===1)lbl.textContent=SCOPE[0].slice(0,18);
  else lbl.textContent=SCOPE.length+' sheets';
}
function toggleScopePop(){
  const p=document.getElementById('scopePop');
  p.style.display=p.style.display==='block'?'none':'block';
}
function setScope(v){SCOPE=v;renderScope()}
function toggleScope(name){
  if(SCOPE===null)SCOPE=[name];
  else if(SCOPE.includes(name))SCOPE=SCOPE.filter(x=>x!==name);
  else SCOPE=SCOPE.concat([name]);
  if(!SCOPE.length)SCOPE=null;
  renderScope();
}
async function newChat(){
  const r=await (await fetch('/api/session_new',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({scope:SCOPE})})).json();
  CUR_SESSION=r.session_id;
  document.getElementById('chatLog').innerHTML=`<div class="a hint">${t('new_chat')}</div>`;
  loadSessions();
}
async function loadSessions(){
  const list=await (await fetch('/api/sessions')).json();
  const el=document.getElementById('sessionList');
  const cur=list.find(s=>s.id===CUR_SESSION);
  document.getElementById('curTitle').textContent=cur?(cur.title||t('s_untitled')):t('cur_title_empty');
  if(!list.length){el.innerHTML=`<div class="hint">${t('hist_empty')}</div>`;return}
  el.innerHTML=list.map(s=>`<div class="sItem${s.id===CUR_SESSION?' on':''}" onclick="switchSession(${s.id})">
    <div style="font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(s.title||t('s_untitled'))}</div>
    <div class="hint" style="font-size:11px">${s.updated_at||''}${s.scope?' · '+s.scope.length+t('m_sheets'):''}</div>
    <button class="sDel" title="${t('s_del_tt')}" onclick="event.stopPropagation();delSession(${s.id})">✕</button></div>`).join('');
}
async function switchSession(id){
  const r=await (await fetch('/api/session_switch',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({id})})).json();
  CUR_SESSION=id;
  SCOPE=r.session.scope||null;renderScope();
  const log=document.getElementById('chatLog');
  log.innerHTML=(r.messages.length?'':`<div class="a hint">${t('sess_empty')}</div>`)+
    r.messages.map(m=>m.role==='user'?`<div class="u">${t('m_q')}${esc(m.content)}</div>`
      :`<div class="a">${mdRich(m.content)}</div>`).join('');
  log.scrollTop=log.scrollHeight;
  loadSessions();
}
async function delSession(id){
  await fetch('/api/session_delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({id})});
  if(CUR_SESSION===id){CUR_SESSION=null;document.getElementById('chatLog').innerHTML=`<div class="a hint">${t('sess_deleted')}</div>`}
  loadSessions();
}
async function clearSessions(){
  if(!confirm(t('hist_clear_confirm')))return;
  await fetch('/api/sessions_clear',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  CUR_SESSION=null;loadSessions();loadHistList();
}
async function loadHistList(){
  const list=await (await fetch('/api/sessions')).json();
  document.getElementById('histList').innerHTML=list.length?list.map(s=>
    `<div style="display:flex;gap:6px;padding:4px 2px;border-bottom:1px solid #f1f5f9">
      <div style="flex:1;cursor:pointer" onclick="showPane('chat');switchSession(${s.id})">
      <b>${esc(s.title||t('s_untitled'))}</b> <span class="hint">${s.updated_at||''}</span></div>
      <button style="border:0;background:transparent;color:#dc2626;cursor:pointer" onclick="delSession(${s.id});loadHistList()">✕</button></div>`).join('')
    :t('hist_none');
}
async function doHistSearch(){
  const q=document.getElementById('histSearch').value.trim();if(!q)return;
  const r=await (await fetch('/api/search_history?q='+encodeURIComponent(q))).json();
  document.getElementById('histResults').innerHTML=r.length?r.map(h=>
    `<div style="padding:5px 2px;border-bottom:1px solid #f1f5f9"><b>[${esc(h.file)}]</b> ${esc(h.title)}<br>
     <span class="hint">…${esc(h.snippet)}…</span></div>`).join(''):t('hist_no_hit');
}
async function loadWriteLog(){
  const r=await (await fetch('/api/write_log')).json();
  document.getElementById('writeLog').innerHTML=r.length?r.map(w=>
    `<div style="padding:3px 2px;border-bottom:1px solid #f1f5f9">${w.undone?'<s>':''}
     <span class="hint">${w.ts}</span> [${esc(w.sheet)}] r${w.row}c${w.col}：${esc(w.old)} → <b>${esc(w.new)}</b>${w.undone?t('undone'):''}${w.undone?'</s>':''}</div>`).join('')
    :t('write_log_none');
}
async function doAsk(){
  const box=document.getElementById('askBox');
  const q=box.value.trim();if(!q)return;
  box.value='';box.focus();  // 发送瞬间清空并聚焦，不等回答
  const log=document.getElementById('chatLog');
  log.insertAdjacentHTML('beforeend',`<div class="u">${t('m_q')}${esc(q)}</div>`);
  log.scrollTop=log.scrollHeight;
  const r=await (await fetch('/api/ask',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({question:q,sheet:S?S.sheet:null,scope:SCOPE,session_id:CUR_SESSION,lang:LANG})})).json();
  if(r.session_id)CUR_SESSION=r.session_id;
  const tag=r.source==='llm'?'🤖':(r.source==='rules-fallback'?t('tag_fb'):'⚙️');
  log.insertAdjacentHTML('beforeend',`<div class="a">${tag} ${mdRich(String(r.answer??r).slice(0,1200))}</div>`);
  log.scrollTop=log.scrollHeight;
  loadSessions();
}
let pendingOp=null,pendingOpSaved=null;
async function doPlanWrite(mode){
  const text=document.getElementById('writeBox').value.trim();if(!text)return;
  const card=document.getElementById('diffCard');
  card.className='';card.textContent=t('parsing');
  const r=await (await fetch('/api/plan_write',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({mode,text})})).json();
  if(!r.ok){
    pendingOp=null;
    card.className='bad';
    card.innerHTML=`<b>${t('wp_none')}</b><br>${esc(r.clarify||JSON.stringify(r).slice(0,200))}`+
      (r.candidates?`<br>${t('wp_cand')}<br>${r.candidates.map(c=>'· '+esc(c.via)+' @r'+c.row).join('<br>')}`:'');
    return;
  }
  pendingOp=r.op;pendingOpSaved=r.op;window.__lastDeps=r.deps||[];
  const v=r.validation;
  card.className=v.ok?'ok':'bad';
  const modeZh={set:t('wm_set'),add:t('wm_add'),sub:t('wm_sub'),mul:t('wm_mul'),div:t('wm_div'),add_pct:t('wm_add_pct'),sub_pct:t('wm_sub_pct')};
  card.innerHTML=
   (window.COPY?`<div style="background:#fef3c7;border:1px solid #f59e0b;border-radius:6px;
      padding:4px 8px;margin-bottom:6px;font-size:12px;color:#92400e">
      ${t('wp_copy_note')}</div>`:'')+
   (r.hit&&Object.keys(r.hit).length?`<div style="display:flex;align-items:center;gap:4px;flex-wrap:wrap;margin:4px 0">
     <span class="hint" style="font-size:11px">${t('hit_row')}</span>`+
     Object.entries(r.hit).map(([k,vv])=>`<span style="background:#eef2ff;border:1px solid #c7d2fe;border-radius:10px;padding:0 8px;font-size:11px">${esc(k)}=${esc(String(vv).slice(0,14))}</span>`).join('')+`</div>`:'')+
   `<div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap">
     <b>${t('wp_title')}</b><span class="hint">${esc(r.via||'')}</span>
     ${r.op.mode&&modeZh[r.op.mode]?`<span style="background:#e0e7ff;border-radius:8px;padding:0 8px;font-size:11.5px">${modeZh[r.op.mode]}</span>`:''}
     <button style="margin-left:auto;font-size:11px" onclick="locateCell('${esc(r.op.sheet)}',${r.op.row},${r.op.col})">${t('wp_locate')}</button></div>
   <div style="display:flex;align-items:center;gap:8px;margin:6px 0;flex-wrap:wrap">
     <span class="hint">${esc(r.op.sheet)}!r${r.op.row}c${r.op.col}</span>
     <span style="background:#fee2e2;color:#b91c1c;border-radius:6px;padding:2px 10px;font-weight:600">${esc(String(r.old??'∅'))}</span>
     <span style="color:#94a3b8;font-size:16px">→</span>
     <span style="background:#dcfce7;color:#15803d;border-radius:6px;padding:2px 10px;font-weight:700">${esc(String(r.op.value))}</span>
   </div>`+
   (r.deps&&r.deps.length?`<div style="color:#7c3aed;font-size:12px">🔗 ${t('dep_pre')}${r.deps.length}${t('dep_suf')}（r${r.deps[0].row}c${r.deps[0].col} ${esc(r.deps[0].formula)}）</div>`:'')+
   (v.ok?`<div style="color:#15803d">${t('v_allpass')}</div>`:
        `<div style="color:#b91c1c">${t('v_reject')}${v.rejections.map(esc).join(LANG==='en'?'; ':'；')}</div>`)+
   (v.warnings.length?`<div style="color:#b45309">⚠ ${v.warnings.map(esc).join(LANG==='en'?'; ':'；')}</div>`:'')+
   (v.ok?`<div style="display:flex;gap:6px;margin-top:8px">
      <button class="primary" onclick="doExec()">${t('v_confirm')}</button>
      <button onclick="pendingOp=null;document.getElementById('diffCard').innerText=t('v_cancelled')">${t('v_cancel')}</button></div>`:'');
  locateCell(r.op.sheet,r.op.row,r.op.col);
}
async function locateCell(sheet,row,col){
  if(S.sheet!==sheet){
    document.getElementById('sheetSel').value=sheet;
    await loadState(sheet);
  }
  const td=anchorCell(row,col);
  if(!td)return;
  td.scrollIntoView({block:'center',inline:'center'});
  td.classList.remove('flashCell');void td.offsetWidth;td.classList.add('flashCell');
}
async function doExec(){
  if(!pendingOp)return;
  const r=await (await fetch('/api/execute_write',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({op:pendingOp})})).json();
  const card=document.getElementById('diffCard');
  card.className=r.ok?'ok':'bad';
  card.innerHTML=r.ok
    ?`${t('wr_ok_pre')}${esc(r.engine)}${t('wr_ok_mid')}${esc(r.backup)}${t('wr_ok_suf')}`+
      (window.__lastDeps&&window.__lastDeps.length?
        (r.deps_changed&&r.deps_changed.length?
          `<br><span style="color:#7c3aed;font-size:12px">🔗 ${t('dep_done_pre')}r`+
          r.deps_changed.slice(0,4).map(d=>`${d.row}c${d.col}`).join(LANG==='en'?', ':'、')+
          `${t('dep_done_suf')}</span>`:
          `<br><span style="color:#b45309;font-size:12px">${t('dep_unchanged')}</span>`):'')+
      (window.COPY?`<br><b style="color:#92400e">${t('wr_copy_hint')}</b>
        <a href="/api/export" download style="color:#d97706;font-weight:700">${t('export_btn')}</a>`:'')+
      `<br>${t('wr_undo_hint')}`
    :`${t('wr_fail')}${esc(r.reason||JSON.stringify(r).slice(0,200))}`;
  pendingOp=null;loadState(document.getElementById('sheetSel').value);
  if(r.ok){const op=r.op||pendingOpSaved; if(op){locateCell(op.sheet,op.row,op.col).then(()=>{
    const td=anchorCell(op.row,op.col);if(td){td.classList.remove('flashCell');
      void td.offsetWidth;td.classList.add('flashDone')}})}}
}
async function doUndo(){
  const r=await (await fetch('/api/undo',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'})).json();
  const log=document.getElementById('chatLog');
  log.insertAdjacentHTML('beforeend',
    `<div class="a">${r.ok?t('wr_undone_pre')+r.restored+t('wr_undone_suf')+esc(r.engine)+t('wr_undone_end'):esc(r.reason||t('wr_none'))}</div>`);
  loadState(document.getElementById('sheetSel').value);
}
window.addEventListener('resize',()=>S&&drawOverlays());
bootWait();
/* ===== 文件标签页 ===== */
let TABS=[], ACTIVE_TAB=-1;
function loadTabs(){try{TABS=JSON.parse(localStorage.getItem('gewu-tabs')||'[]')}catch(e){TABS=[]}
  renderTabs()}
function saveTabs(){localStorage.setItem('gewu-tabs',JSON.stringify(TABS))}
function renderTabs(){
  const bar=document.getElementById('tabBar');
  if(TABS.length<=1){bar.style.display='none';return}
  bar.style.display='flex';
  bar.innerHTML=TABS.map((tab,i)=>{
    const nm=tab.name||'?';
    return `<div class="ftab${i===ACTIVE_TAB?' on':''}" onclick="switchToTab(${i})" title="${esc(tab.path)}">
      <span class="tName">${esc(nm)}</span>
      ${tab.copy?`<span class="tBadge">〔${t('copy_badge')}〕</span>`:''}
      <button class="tClose" onclick="event.stopPropagation();closeTab(${i})">✕</button>
    </div>`}).join('')+
    `<button style="border:0;background:transparent;cursor:pointer;font-size:14px;padding:3px 8px;color:#64748b"
      onclick="showOpen()">＋</button>`;
}
async function syncTabs(){
  if(!S||!S.workbook)return;
  const cur=S.workbook.path;
  const i=TABS.findIndex(t=>t.path===cur);
  if(i>=0){ACTIVE_TAB=i;TABS[i].sheet=S.sheet;TABS[i].scope=SCOPE;
    TABS[i].sessionId=CUR_SESSION;TABS[i].copy=window.COPY}
  else{TABS.push({path:cur,name:cur.split(/[\\/]/).pop(),sheet:S.sheet,
    scope:SCOPE,sessionId:CUR_SESSION,copy:window.COPY});
    ACTIVE_TAB=TABS.length-1}
  saveTabs();renderTabs();
}
async function switchToTab(i){
  if(i===ACTIVE_TAB||!TABS[i])return;
  if(TABS[ACTIVE_TAB]){TABS[ACTIVE_TAB].sheet=S.sheet;TABS[ACTIVE_TAB].scope=SCOPE;
    TABS[ACTIVE_TAB].sessionId=CUR_SESSION}
  const tgt=TABS[i];ACTIVE_TAB=i;
  const r=await (await fetch('/api/open',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({path:tgt.path})})).json();
  if(!r.ok){alert(t('switch_fail')+' '+(r.reason||''));return}
  await reloadForNewFile();
  if(tgt.sheet){const sel=document.getElementById('sheetSel');
    const opt=[...sel.options].find(o=>o.textContent.includes(tgt.sheet));
    if(opt){sel.value=opt.value;await loadState(tgt.sheet)}}
  if(tgt.scope!==undefined&&tgt.scope!==null){SCOPE=tgt.scope;renderScope()}
  else{SCOPE=null;renderScope()}
  saveTabs();renderTabs();
}
async function closeTab(i){
  if(TABS.length<=1)return;
  TABS.splice(i,1);
  if(i===ACTIVE_TAB){ACTIVE_TAB=Math.min(i,TABS.length-1);
    if(TABS[ACTIVE_TAB])await switchToTab(ACTIVE_TAB)}
  else if(i<ACTIVE_TAB)ACTIVE_TAB--;
  saveTabs();renderTabs();
}
async function reloadForNewFile(){
  edits={};sel=null;CUR_SESSION=null;SCOPE=null;
  await loadSheets();loadSessions();refreshStatus();
}
async function bootWait(){
  const ov=document.createElement('div');
  ov.id='bootOverlay';
  ov.style.cssText='position:fixed;inset:0;z-index:200;background:linear-gradient(135deg,#0f172a,#1e3a5f);color:#e2e8f0;display:flex;flex-direction:column;align-items:center;justify-content:center;font:14px/1.8 system-ui';
  ov.innerHTML='<div style="font-size:26px;font-weight:700;letter-spacing:4px">格物 <span style="background:linear-gradient(90deg,#60a5fa,#93c5fd);-webkit-background-clip:text;background-clip:text;color:transparent">Gewu</span></div><div id="bootMsg" style="color:#94a3b8;margin-top:10px">…</div><div style="width:30px;height:30px;border:3px solid #334155;border-top-color:#60a5fa;border-radius:50%;animation:gwspin .9s linear infinite;margin-top:16px"></div><style>@keyframes gwspin{to{transform:rotate(360deg)}}</style>';
  document.body.appendChild(ov);
  for(let i=0;i<300;i++){
    let b=null;
    try{b=await (await fetch('/api/boot')).json()}catch(e){}
    if(b&&b.ready){ov.remove();loadSheets();return}
    if(b&&b.phase==='failed'){
      ov.querySelector('#bootMsg').innerHTML=t('boot_fail')+(b.error||t('boot_unknown'))+`<br><span style="font-size:12px;color:#64748b">${t('boot_hint')}</span>`;
      ov.lastElementChild?.remove?.(); ov.querySelector('div[style*="gwspin"],div[style*="rotate"]')?.remove?.();
      return;
    }
    if(b&&b.phase)document.getElementById('bootMsg').textContent=b.phase;
    await new Promise(r=>setTimeout(r,800));
  }
}
//
