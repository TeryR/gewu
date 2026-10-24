// 真浏览器 CDP 实测：框微调修复回归（独立无头 Chrome，端口 9333，不动用户浏览器）
// 验证：①表格块缩小拖拽 ②黄框批注(id=0)拖拽 ③贴顶标签翻进框内（自然贴顶或滚动贴顶）
// 用法：node tests/cdp_drag_check.js   （需 serve.py 已在 8765 运行）
const { spawn } = require('child_process');
const os = require('os'), path = require('path'), fs = require('fs');

const CHROME = 'C:/Program Files/Google/Chrome/Application/chrome.exe';
const PORT = 9333, APP = 'http://127.0.0.1:8765';
const PROFILE = path.join(os.tmpdir(), 'gewu-cdp-9333');
let fails = 0;
const check = (name, ok, extra) => {
  console.log((ok ? '  ok ' : '  FAIL ') + name + (extra ? '  [' + extra + ']' : ''));
  if (!ok) fails++;
};

const chrome = spawn(CHROME, [
  '--headless=new', `--remote-debugging-port=${PORT}`, `--user-data-dir=${PROFILE}`,
  '--no-first-run', '--no-default-browser-check', '--window-size=1440,960', APP,
], { stdio: 'ignore' });

const sleep = ms => new Promise(r => setTimeout(r, ms));

async function pageWs() {
  for (let i = 0; i < 40; i++) {
    try {
      const list = await (await fetch(`http://127.0.0.1:${PORT}/json`)).json();
      const pg = list.find(t => t.type === 'page' && t.url.startsWith('http'));
      if (pg) return pg.webSocketDebuggerUrl;
    } catch (e) { /* chrome 未起 */ }
    await sleep(500);
  }
  throw new Error('chrome CDP 未就绪');
}

function cdp(ws) {
  let seq = 0; const pending = {};
  ws.addEventListener('message', m => {
    const d = JSON.parse(m.data);
    if (d.id && pending[d.id]) { pending[d.id](d); delete pending[d.id]; }
  });
  const send = (method, params = {}) => new Promise(res => {
    const id = ++seq; pending[id] = res; ws.send(JSON.stringify({ id, method, params }));
  });
  return {
    eval: async (expr, awaitPromise = true) => {
      const r = await send('Runtime.evaluate', { expression: expr, returnByValue: true, awaitPromise });
      if (r.exceptionDetails) throw new Error('页面异常: ' +
        (r.exceptionDetails.exception?.description || r.exceptionDetails.text));
      return r.result.result.value;
    },
    close: () => send('Browser.close'),
    send,
  };
}

const PAGE = `(async()=>{
  const ev=(el,t,x,y)=>el.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,clientX:x,clientY:y,button:0}));
  const rd=()=>['ar1','ac1','ar2','ac2'].map(k=>document.getElementById(k).value);
  const esc=()=>document.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}));
  const out={visited:[],table:null,note:null,top:null,scroll:null};
  const wt=()=>document.getElementById('stage').getBoundingClientRect().top;   // 裁切边界=视口顶
  for(const s of (typeof SHEETS!=='undefined'?SHEETS:[])){
    await loadState(s.name);
    await new Promise(r=>setTimeout(r,700));
    const blocks=[...document.querySelectorAll('#overlays .blk')];
    out.visited.push(s.name+'('+blocks.length+')');
    // ① 表格缩小拖拽：面积最大的蓝表，东把手拖到块内 30% 处
    if(!out.table){
      let tb=null,area=0;
      for(const b of blocks){ if(b.dataset.color!=='#2f6fdb')continue;
        const g=b.getBoundingClientRect(); if(g.width*g.height>area){area=g.width*g.height;tb=b} }
      if(tb&&area>3000){
        tb.click();
        const selB=document.querySelector('.blk.sel');
        const hs=[...selB.children].filter(x=>x.tagName==='DIV');
        const eH=hs.find(h=>h.style.cursor==='ew-resize'&&h.style.right==='-5px');
        const hg=eH.getBoundingClientRect(), br=selB.getBoundingClientRect();
        ev(eH,'mousedown',hg.left+5,hg.top+5);
        const before=rd();
        const cx=br.left+br.width*0.3, cy=br.top+br.height*0.5;
        ev(document.body,'mousemove',cx,cy);
        const mid=rd();
        ev(document.body,'mouseup',cx,cy); esc();
        out.table={sheet:s.name,handles:hs.length,before:before.join(','),mid:mid.join(','),
                   shrunk:+mid[3]<+before[3]&&mid[0]===before[0]&&mid[2]===before[2]};
      }
    }
    // ② 黄批注拖拽：先向右扩，扩不动（贴边）再向下扩
    if(!out.note){
      const nt=blocks.find(b=>b.dataset.quiet==='1'&&b.dataset.color==='#d9a406');
      if(nt){
        nt.click();
        const selB=document.querySelector('.blk.sel');
        const hs=[...selB.children].filter(x=>x.tagName==='DIV');
        const eH=hs.find(h=>h.style.cursor==='ew-resize'&&h.style.right==='-5px');
        const sH=hs.find(h=>h.style.cursor==='ns-resize'&&h.style.bottom==='-5px');
        const hg=eH.getBoundingClientRect(), br=selB.getBoundingClientRect();
        const gr=document.getElementById('grid').getBoundingClientRect();
        ev(eH,'mousedown',hg.left+5,hg.top+5);
        const before=rd();
        let cx=Math.min(br.right+80,gr.right-8), cy=br.top+br.height/2;
        ev(document.body,'mousemove',cx,cy);
        let after=rd();
        if(+after[3]<=+before[3]){                    // 右边到头 → 换南把手向下
          ev(document.body,'mouseup',cx,cy); esc();
          const sg=sH.getBoundingClientRect();
          nt.click();
          const sb2=document.querySelector('.blk.sel');
          const hs2=[...sb2.children].filter(x=>x.tagName==='DIV');
          const s2=hs2.find(h=>h.style.cursor==='ns-resize'&&h.style.bottom==='-5px');
          ev(s2,'mousedown',sg.left+5,sg.top+5);
          ev(document.body,'mousemove',br.left+br.width/2,br.bottom+40);
          after=rd();
          ev(document.body,'mouseup',br.left+br.width/2,br.bottom+40); esc();
        }else{
          ev(document.body,'mouseup',cx,cy); esc();
        }
        out.note={sheet:s.name,handles:hs.length,before:before.join(','),after:after.join(','),
                  resized:+after[3]>+before[3]||+after[2]>+before[2]};
      }
    }
    // ③ 自然贴顶块
    if(!out.top){
      const near=blocks.filter(b=>{const d=b.getBoundingClientRect().top-wt();return d<12&&d>=-4});
      if(near.length) out.top={sheet:s.name,
        flips:near.map(b=>b.querySelector('.tag').style.top||'(外)'),
        ok:near.every(b=>b.querySelector('.tag').style.top==='2px')};
    }
    if(out.table&&out.note&&out.top)break;
  }
  return JSON.stringify(out);
})()`;

// 贴顶翻转兜底（小视口强制滚动）：由 driver 先 Emulation 压低窗口再执行
const PAGE_SCROLL = `(async()=>{
  const st=document.getElementById('stage');
  const wt=()=>document.getElementById('stage').getBoundingClientRect().top;   // 裁切边界=视口顶
  const scrollable=st.scrollHeight>st.clientHeight+4;
  const b=document.querySelector('#overlays .blk');
  let d=null,tagTop=null,ok=false;
  if(b&&scrollable){
    st.scrollTop=Math.max(0,b.getBoundingClientRect().top-wt()-2);
    await new Promise(r=>setTimeout(r,300));   // 滚动防抖 80ms 后翻标签
    const nb=document.querySelector('#overlays .blk');
    d=Math.round(nb.getBoundingClientRect().top-wt());
    tagTop=nb.querySelector('.tag').style.top||'(外)';
    ok=d<12&&tagTop==='2px';
  }
  return JSON.stringify({scrollable,scrollTop:st.scrollTop,topDelta:d,tagTop,ok});
})()`;

(async () => {
  try {
    const ws = new globalThis.WebSocket(await pageWs());   // Node 22 内置
    await new Promise((res, rej) => {
      ws.addEventListener('open', res);
      ws.addEventListener('error', e => rej(new Error('ws error: ' + (e.message || e.type))));
    });
    const c = cdp(ws);
    for (let i = 0; i < 30; i++) {          // 等应用 boot + 块渲染
      const n = await c.eval('document.querySelectorAll("#overlays .blk").length', false);
      if (n > 0) break;
      await sleep(600);
    }

    const R = JSON.parse(await c.eval(PAGE));
    console.log('遍历工作表:', R.visited.join(' | '));
    console.log('── 表格块·缩小拖拽（真实 Chrome 事件）──');
    if (R.table) {
      check('选中蓝表出 8 把手（无多余药丸）', R.table.handles === 8, 'handles=' + R.table.handles);
      check('缩小：东把手左拖 ' + R.table.before + ' → ' + R.table.mid, R.table.shrunk === true, R.table.sheet);
    } else check('表格缩小拖拽', false, '没找到面积足够的蓝表块');

    console.log('── 黄框批注（id=0 假值回归）·拖拽 ──');
    if (R.note) {
      check('黄批注选中出 8 把手', R.note.handles === 8, R.note.sheet);
      check('黄批注拖拽真实变形 ' + R.note.before + ' → ' + R.note.after, R.note.resized === true, R.note.sheet);
    } else check('黄框批注可拖', false, '未找到 note 块');

    console.log('── 贴顶标签翻转 ──');
    if (R.top) {
      check('自然贴顶块标签翻进框内', R.top.ok === true, R.top.sheet + ' ' + JSON.stringify(R.top.flips));
    } else {
      await c.send('Emulation.setDeviceMetricsOverride',
        { width: 1440, height: 300, deviceScaleFactor: 1, mobile: false });
      await sleep(400);
      const S = JSON.parse(await c.eval(PAGE_SCROLL));
      check('小视口滚动：stage 可滚', S.scrollable === true, 'scrollTop=' + S.scrollTop);
      check('滚动贴顶后标签翻进框内', S.ok === true,
        'topΔ=' + S.topDelta + ' tag=' + S.tagTop);
    }

    await c.close();
  } catch (e) {
    console.error('DRIVER ERROR:', e.message);
    fails++;
  } finally {
    chrome.kill();
    setTimeout(() => { try { fs.rmSync(PROFILE, { recursive: true, force: true }); } catch (e) {} }, 1500);
  }
  console.log(fails ? `\nFAILED: ${fails}` : '\nALL PASS（真浏览器）');
  process.exit(fails ? 1 : 0);
})();
