// 无头 DOM 测试：拖拽命中穿透 + 缩小/平移数学（node tests/test_adj_drag.js）
// 模拟浏览器层：elementsFromPoint 返回"该点元素栈"——缩边时光标正下方是
// 把手(11px)与 .blk 彩框（pointer-events:auto 的老实现会在此丢失 td）
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC = fs.readFileSync(path.join(__dirname, '..', 'static', 'app.js'), 'utf8');

function fn(name) {           // 从源码里抠出单个函数体（不依赖整个 app 上下文）
  const i = SRC.indexOf('function ' + name + '(');
  if (i < 0) throw new Error('fn not found: ' + name);
  let d = 0, j = SRC.indexOf('{', i);
  for (let k = j; k < SRC.length; k++) {
    if (SRC[k] === '{') d++;
    else if (SRC[k] === '}') { d--; if (!d) return SRC.slice(i, k + 1); }
  }
  throw new Error('unbalanced');
}

let fails = 0;
const check = (name, ok) => { console.log((ok ? '  ok ' : '  FAIL ') + name); if (!ok) fails++; };

// ---- 最小 DOM 桩 ----
function el(extra) {
  return Object.assign({
    dataset: {}, style: {},
    closest(sel) { return sel === 'td[data-r]' && this._isTd ? this : null; },
    appendChild() {}, addEventListener() {}, removeEventListener() {},
  }, extra || {});
}
const td5x4 = el({ _isTd: true }); td5x4.dataset = { r: '5', c: '4' };
let stack;                    // elementsFromPoint 的可编程返回
let syncCalls = 0;
const sandbox = {
  document: {
    elementsFromPoint: () => stack,
    elementFromPoint: () => stack[0],   // 老实现语义：只取最顶
    addEventListener() {}, removeEventListener() {},
  },
  adj: null,
  syncAdjustUI: () => syncCalls++,
};
vm.createContext(sandbox);
vm.runInContext(fn('cellAt') + '; ' + fn('adjDragMove'), sandbox);
const { cellAt, adjDragMove: move } = sandbox;

// ---- 1. cellAt 命中穿透 ----
// 场景 A：缩小拖拽中——光标下方依次是把手、彩框、格子（真实层叠结构）
stack = [el(), el(), td5x4];
check('穿透：栈顶是把手+彩框也能拿到底层格',
  JSON.stringify(cellAt({ clientX: 1, clientY: 1 })) === '{"r":5,"c":4}');
// 场景 B：光标下只有把手（拖出格子区域边缘）→ 无 td，返回 null（不误报）
stack = [el(), el()];
check('穿透：栈里无 td 时返回 null', cellAt({ clientX: 1, clientY: 1 }) === null);

// ---- 2. adjDragMove 数学：缩小、平移 ----
// 东边把手向左拖（缩小宽度）：原始 c2=8，光标停在 c4
sandbox.adj = { drag: { h: 'e', vals: [2, 2, 8, 8], startCell: { r: 2, c: 2 } }, vals: [2, 2, 8, 8] };
stack = [el(), el(), td5x4];
move({ clientX: 1, clientY: 1 });
check('缩小：东把手左拖 c2 8→4', JSON.stringify(sandbox.adj.vals) === '[2,2,8,4]');
check('缩小：其余边不动',
  sandbox.adj.vals[0] === 2 && sandbox.adj.vals[1] === 2 && sandbox.adj.vals[2] === 8);

// 平移：move 模式整体位移（cell 与 startCell 差 dr=3, dc=2）
sandbox.adj = { drag: { h: 'move', vals: [2, 2, 8, 8], startCell: { r: 2, c: 2 } }, vals: [2, 2, 8, 8] };
move({ clientX: 1, clientY: 1 });
check('平移：整体 +3行 +2列', JSON.stringify(sandbox.adj.vals) === '[5,4,11,10]');

// 北把手下拖（缩小高度）：r1 2→5（跨沿自动翻转留给 mouseup）
sandbox.adj = { drag: { h: 'n', vals: [2, 2, 8, 8], startCell: { r: 2, c: 2 } }, vals: [2, 2, 8, 8] };
move({ clientX: 1, clientY: 1 });
check('缩小：北把手下拖 r1 2→5', sandbox.adj.vals[0] === 5 && sandbox.adj.vals[2] === 8);

check('拖中轻量同步被调用', syncCalls === 3);

// ---- 3. 源码级断言：彩框不吃点击（调整模式） ----
check('源码：调整模式 .blk pointerEvents=none', SRC.includes("if(adj)d.style.pointerEvents='none'"));
check('源码：贴顶标签翻进框内', SRC.includes('标签翻进框内'));
check('源码：平移钮悬停说明 adj_move_tt×3（中/英/引用）', SRC.split('adj_move_tt').length - 1 >= 3);
check('源码：块 id=0 假值陷阱已修（把手判定）', SRC.includes('if(blockId==null){'));
check('源码：块 id=0 假值陷阱已修（入口判定）', SRC.includes('if(blockId==null)return'));
check('源码：块 id=0 假值陷阱已修（把手配色）', SRC.includes("${blockId==null?'#ea580c':'#2563eb'}"));

// ---- 4. 块 id=0 功能回归：蓝框 8 把手、无平移药丸 ----
{
  const mk = () => Object.assign(el(), { children: [], appendChild(c) { this.children = this.children || []; this.children.push(c); } });
  const host0 = mk(), hostP = mk();
  const sb = {
    document: { createElement: mk, addEventListener() {}, removeEventListener() {} },
    adjHandleDown() {}, t: () => 'x', host0, hostP,
  };
  vm.createContext(sb);
  vm.runInContext(fn('_adjHandles') + ';_adjHandles(host0,0);_adjHandles(hostP,null);', sb);
  const css = c => c.style.cssText || '';
  const cnt = h => h.children.filter(c => css(c).includes('resize')).length;
  const pill = h => h.children.filter(c => css(c).includes('cursor:move')).length;
  check('id=0 蓝框：8 个缩放把手', cnt(host0) === 8, 'got ' + cnt(host0));
  check('id=0 蓝框：无平移药丸（药丸只属于预览）', pill(host0) === 0);
  check('预览(null)：8 把手 + 1 药丸', cnt(hostP) === 8 && pill(hostP) === 1);
}

console.log(fails ? '\nFAILED: ' + fails : '\nALL PASS');
process.exit(fails ? 1 : 0);
