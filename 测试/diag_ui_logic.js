/* 界面逻辑无浏览器自检（回归测试）
 * ==================================
 * 给 gui/app.js 套一层最小 DOM 桩，直接调用渲染函数，检查产出的 HTML。
 * 目的：不开浏览器 / 不弹窗口 / 不启动本地服务，就能验证
 *   · 台账表格的列数（表头 vs 每行单元格数）有没有错位
 *   · 「凭证类型」「行程/明细」有没有渲染出来
 *   · 查询条件（含凭证类型/批次/报销人/金额区间）有没有被 currentFilter 收集齐
 *   · 删除选中 / 清空台账 / 删除生成文件 的二次确认弹层
 *   · 凭证明细弹层的分区结构、空字段、脏数据
 *   · 入库结果卡片（含按凭证类型拆分）
 *
 * 跑法（node 用托管版）：
 *   C:\Users\JM\.workbuddy\binaries\node\versions\22.22.2-3\node.exe 测试\diag_ui_logic.js
 * 退出码 0 = 全通过。
 */
const fs = require('fs');
const path = require('path');

const GUI = path.join(__dirname, '..', 'gui');
const src = fs.readFileSync(path.join(GUI, 'app.js'), 'utf8');

const mkEl = (id) => ({
  id, innerHTML: '', textContent: '', value: '', checked: false, indeterminate: false,
  style: {}, dataset: {}, scrollTop: 0, disabled: false, title: '', hidden: false,
  // classList 以前是空的假桩（contains 永远 false）—— 权限压灰要断言 perm-off，
  // 假桩测不出来。换成真 Set，语义跟浏览器一致，对既有用例等价（没人加过 active）。
  classList: {
    _s: new Set(),
    add(c) { this._s.add(c); },
    remove(c) { this._s.delete(c); },
    toggle(c) { this._s.has(c) ? this._s.delete(c) : this._s.add(c); },
    contains(c) { return this._s.has(c); },
  },
  setAttribute(k, v) { if (k === 'title') this.title = v; },
  removeAttribute(k) { if (k === 'title') this.title = ''; },
  addEventListener() {}, insertAdjacentElement() {}, appendChild() {}, remove() {},
  closest() { return null; }, querySelector() { return mkEl('inner'); }, querySelectorAll() { return []; },
});
const els = {};
global.window = { addEventListener() {}, pywebview: null, location: { search: '', hash: '' } };
global.document = {
  body: mkEl('body'),
  getElementById: (id) => els[id] || (els[id] = mkEl(id)),
  querySelector: (s) => global.document.getElementById('__' + s),
  querySelectorAll: () => [],
  createElement: () => mkEl('new'),
  addEventListener() {},
};
global.localStorage = { getItem: () => null, setItem() {}, removeItem() {} };
global.location = global.window.location;

/* 真的可写 localStorage 桩：折叠状态得靠「存了再读回来」验证，
   原来那个 setItem 是空转、读回来永远是 null，测了等于没测。
   （把 getItem 从「永远 null」换成「空表里读」对其他用例是等价的：
     测试期间没有任何东西写过别的键。） */
const _store = {};
global.localStorage = {
  getItem: (k) => (k in _store ? _store[k] : null),
  setItem: (k, v) => { _store[k] = String(v); },
  removeItem: (k) => { delete _store[k]; },
};

/* 带**真实** classList 的元素桩。折叠就是往元素上挂/摘 .collapsed，
   用上面那个 mkEl（add/remove 空转、contains 永远 false）当桩，
   断言会全部假通过 —— 这种「测了但没测到」最坑，所以单独造一个。 */
const mkFoldEl = (id) => {
  const e = mkEl(id);
  const set = new Set();
  e.classList = {
    add: (c) => set.add(c),
    remove: (c) => set.delete(c),
    toggle: (c, on) => {
      const want = (on === undefined) ? !set.has(c) : !!on;
      if (want) set.add(c); else set.delete(c);
      return want;
    },
    contains: (c) => set.has(c),
  };
  e.setAttribute = (k, v) => { e['attr:' + k] = String(v); };
  e.getAttribute = (k) => e['attr:' + k];
  e.querySelector = () => null;      // 卡头：setFold 里对 null 有兜底分支
  return e;
};

const box = new Function(
  src + '\n; return {showInvoiceDetail, renderLedgerBody, syncSelectionUI, onImportResult,'
      + ' currentFilter, filterSummary, deleteSelected, clearLedger, deleteOneRow,'
      + ' renderOutputs, deleteOutput, clearOutputs, syncFmtPills, loadOutputs,'
      + ' missListHtml, state, renderReportRows, renderAlw, alwRowAmount, alwSum,'
      + ' alwPayload, reportMeta, syncKindPills, initAlw, stampTag, refreshSelected,'
      + ' setTraveler, loadPhoneMap, savePhoneMap,'
      + ' renderJy, jyPayload, jySum, jyReal, jyNames, JY_KIND, refreshJyPeople, fillJyByDays,'
      + ' clearAlw, clearJy, resetAllowances, pruneJyAlw, purgeLegacyAllowanceKeys,'
      + ' renderJobs, jobRowHtml, fmtSecs, JOB_CLS,'
      + ' tvrBadge, renderTvr, syncTvBadge, saveTravelerReview,'
      // 卡片折叠（2026-09-17）：存取、开关、以及「整组展开」那一下
      + ' readFolds, saveFold, setFold, toggleFold, setFoldGroup, openWorkPanel, applyFolds, LS,'
      // 使用说明（2026-09-17）：内容表、拼 HTML、入口
      + ' GUIDE, guideHtml, showGuide, initGuide, markGuideSeen,'
      // 设置页「账号与权限」（2026-09-17）：卡片为什么显示/隐藏、怎么解释
      + ' renderUsers, loadSettingsExtras, usersBlockedHtml, call,'
      // 权限压灰（2026-09-17）：点不了的按钮不许亮着
      + ' applyPermissions, canDo, PERM_BTNS,'
      // 「按批次分开报」+ 合计口径 + 票面姓名保护（2026-09-17 第三轮）
      + ' renderBatchBar, onBatchBarClick, isAttach, askFaceOverwrite, doSaveTraveler,'
      + ' selectedSum, renderLedgerStats, syncSelectionUI,'
      // meInfo 是模块级 let，这里开个读写口子只为测试用
      + ' setMeTest: (v) => { meInfo = v; }, getMeTest: () => meInfo,'
      // tvr 是 app.js 里的模块级 let，这里用闭包拿一个读写口子（只为测试用）
      + ' setTvrTest: (v) => { tvr = v; }, getTvrTest: () => tvr};'
)();

let fails = 0;
const chk = (cond, label, extra) => {
  console.log((cond ? '  ok   ' : '  FAIL ') + label + (cond ? '' : '  <<< ' + extra));
  if (!cond) fails++;
};
const setVal = (id, v) => { els[id] = els[id] || mkEl(id); els[id].value = v; };
const getVal = (id) => (els[id] || {}).value || '';

/* ---------- 造两条真实形状的台账行 ---------- */
const train = {
  '序号': 5, '状态': '未报销', '凭证类型': '火车票', '发票号码': '', '开票日期': '2025-03-02',
  '票种': '电子发票（铁路电子客票）', '出行人': '张三',
  '行程/明细': 'G1234 武汉-宜昌东 2025-03-02 08:15 二等座 05车12A',
  '购方名称': '国网湖北省电力有限公司', '销方名称': '中国铁路武汉局集团有限公司',
  '不含税金额': '72.48', '税额': '6.52', '价税合计': '79.00',
  '项目/事由': '', '项目编号': '', '费用类别': '交通费', '报销人': '张三',
  '报销批次': '', '入库时间': '2026-09-15 10:00',
  '文件路径': 'E:\\桌\\湖北报销发票整理\\子目录\\铁路电子客票_79元_张三.pdf', '提示': '',
};
const einv = {
  '序号': 6, '状态': '已报销', '凭证类型': '发票', '发票号码': '26349119423005870096',
  '开票日期': '2025-03-05', '票种': '电子发票（普通发票）', '出行人': '',
  '行程/明细': '', '购方名称': '国网湖北省电力有限公司', '销方名称': '某某商贸有限公司',
  '不含税金额': '943.40', '税额': '56.60', '价税合计': '1000.00',
  '项目/事由': '湖北恩施建始龙坪网格10千伏线路工程 1815P825000P', '项目编号': '1815P825000P',
  '费用类别': '材料费', '报销人': '李四', '报销批次': '2025-03-10 报销单',
  '入库时间': '2026-09-15 10:00',
  '文件路径': 'E:\\桌\\湖北报销发票整理\\x\\dzfp_26349119423005870096.pdf',
  '提示': '与第 3 行可能重复',
};

const html = fs.readFileSync(path.join(GUI, 'index.html'), 'utf8');

/* =====================================================================
 * 主题默认值 + 卡片折叠（2026-09-17）
 * ===================================================================*/
console.log('— 主题默认值 —');
{
  const css = fs.readFileSync(path.join(GUI, 'style.css'), 'utf8');
  const extra = fs.readFileSync(path.join(GUI, 'extra.css'), 'utf8');

  // 默认外观＝亮色清新。三处必须对上，漏一处就会出现「html 写着 light、
  // JS 又套回 dark」的首屏闪色（先深一帧再变浅，很难看）。
  chk(/<html[^>]*data-theme="light"/.test(html),
    'index.html 静态写着 data-theme="light"（不然首屏会先闪一帧深色）', html.slice(0, 260));
  chk(!/data-theme="dark"/.test(html.split('<body>')[0]), 'html 上没有 data-theme="dark" 的残留');
  chk(/\{\s*id:\s*'light'/.test(src.slice(src.indexOf('const THEMES'), src.indexOf('const THEMES') + 400)),
    'THEMES 里「亮色清新」排第一位（主题面板第一个就是默认那套）');
  chk(/getItem\(LS\.theme\)\s*\|\|\s*'light'/.test(src), "boot() 的兜底是 'light'");
  chk(!/t\.id === id\)\s*id = 'dark'/.test(src), "applyTheme() 找不到 id 时兜底 light，不是 dark");
  for (const id of ['light', 'warm-gold', 'tiffany']) {
    chk(css.includes(`[data-theme="${id}"]`), `style.css 里有 [data-theme="${id}"] 变量块`);
  }

  // ⭐ 黑行回归：:root 装的是深色那套，浅色主题漏一条颜色变量就会漏出深色值。
  //    「凭证明细表头压着一条黑带」就是这么来的（--table-header-bg 漏写 → #0a1521）。
  // ⚠️ 必须用正则定位那个变量块：文件开头的注释里也出现过 `[data-theme="light"]`
  //    这几个字，用 indexOf 会捞到注释，然后切出一段垃圾（踩过）。
  const lightBody = (css.match(/\[data-theme="light"\]\s*\{([\s\S]*?)\n\}/) || ['', ''])[1];
  chk(/--table-header-bg\s*:\s*#eef1f5/.test(lightBody),
    'light 主题有 --table-header-bg（漏了 ⇒ 表头/分组头变近黑）', lightBody.slice(-360));
  for (const v of ['--table-row-hover', '--scrollbar-thumb', '--scrollbar-thumb-hover', '--border-strong',
                   '--text-primary', '--bg-card', '--green', '--red']) {
    chk(new RegExp(v + '\\s*:').test(lightBody), `light 主题定义了 ${v}`);
  }
}

console.log('— 卡片折叠 —');
{
  const extra = fs.readFileSync(path.join(GUI, 'extra.css'), 'utf8');

  chk(/id="work-panel"[^>]*data-fold="closed"/.test(html),
    '出单区整组默认收起（data-fold="closed"）');
  chk(/id="work-panel-body"/.test(html), '出单区正文包了 fold-body，收起时整块藏');
  chk(/id="work-panel-head"[^>]*aria-expanded="false"/.test(html) ||
      /aria-expanded="false"[^>]*id="work-panel-head"/.test(html),
    '把手上带 aria-expanded="false"（状态可读、无障碍）');
  chk(/class="card card-wide" data-fold="closed"/.test(html),
    '设置页「出行人手机号对照」默认收起');
  chk(/data-fold="closed"[\s\S]{0,400}台账数据库/.test(html), '设置页「台账数据库」默认收起');
  chk(/data-fold="closed"[\s\S]{0,200}目录信息/.test(html), '设置页「目录信息」默认收起');
  const nOpen = (html.match(/<div/g) || []).length, nClose = (html.match(/<\/div>/g) || []).length;
  chk(nOpen === nClose, `<div> 开合配平（${nOpen} 开 / ${nClose} 闭）`);

  // 折叠开关必须有样式落地，否则就是「写了 class 没人管」
  chk(/\.card\.collapsed\s*>\s*\*:not\(\.card-head\)\s*\{[^}]*display:\s*none/.test(extra),
    '收起状态真的把正文藏起来（.card.collapsed > *:not(.card-head)）');
  chk(/#work-panel\.collapsed\s+\.fold-body/.test(extra), '出单区收起时藏 fold-body');
  chk(/\.card\s*>\s*\.card-head::after/.test(extra), '卡头右侧有折叠三角');
  chk(/justify-content:\s*flex-start/.test(extra.slice(extra.indexOf('.card-head {'), extra.indexOf('.card-head {') + 400)),
    '卡头改成 flex-start（space-between 会把说明挤到正中间）');

  // 状态存取：存了要能原样读回来
  box.saveFold('t.k1', true);
  chk(box.readFolds()['t.k1'] === 1, '存「收起」能读回来（1）');
  box.saveFold('t.k1', false);
  chk(box.readFolds()['t.k1'] === 0, '再存「展开」能读回来（0）');
  chk(typeof box.LS.folds === 'string' && box.LS.folds.length > 0, 'LS.folds 有键名');

  // 出单区整组开关（用带真 classList 的桩，不然断言会假通过）
  const g = els['work-panel'] = mkFoldEl('work-panel');
  els['work-panel-head'] = mkFoldEl('work-panel-head');
  els['work-panel-hint'] = mkFoldEl('work-panel-hint');
  box.setFoldGroup(g, true);
  chk(g.classList.contains('collapsed'), 'setFoldGroup(true) 把整组收起来');
  chk(els['work-panel-head'].getAttribute('aria-expanded') === 'false', '收起时把手 aria-expanded=false');
  chk(/展开/.test(els['work-panel-hint'].textContent), '收起时说明改写成「点这里展开…」',
    els['work-panel-hint'].textContent);
  box.openWorkPanel();
  chk(!g.classList.contains('collapsed'), '「生成报销单 ↓」能把整组重新展开');
  chk(els['work-panel-head'].getAttribute('aria-expanded') === 'true', '展开时把手 aria-expanded=true');
  box.applyFolds();                 // 一张卡都没有（DOM 桩）也不能抛
  chk(true, 'applyFolds() 在空 DOM 上不抛错');

  // 路径框悬停出完整路径（长路径被省略号截了，得能看全）
  chk(/src-path'\)\.addEventListener\('mouseenter'/.test(src),
    '路径框在 mouseenter 时把当前值挂到 title（悬停显示完整路径）');
  chk(/flex: 0 1 380px/.test(extra), '路径框限宽 380，不再撑满一整行');
}

console.log('— 使用说明（说明书） —');
{
  const extra = fs.readFileSync(path.join(GUI, 'extra.css'), 'utf8');

  // 用户要的三件事：① 刚进来就让人读 ② 可以关 ③ 别处留个地方，忘了能再去点
  chk(typeof box.LS.guideSeen === 'string' && box.LS.guideSeen.length > 0,
    'LS 里有 guideSeen（记「看过了」，不然每次进来都弹）');
  chk(/id="btn-guide"/.test(html), '侧栏底部有常驻入口（忘了可以去那儿再打开）');
  chk(/data-guide-open/.test(html), '入库页副标题里也留了一个入口');
  chk(/getItem\(LS\.guideSeen\)/.test(src) && /setItem\(LS\.guideSeen/.test(src),
    'initGuide：没看过才自动弹；看完 / 点开就记上标记');
  chk(/setTimeout\([\s\S]{0,240}showGuide\(\)/.test(src.slice(src.indexOf('function initGuide'))),
    '首访自动弹（延后一点，别跟台账加载抢同一帧）');
  chk(/Escape/.test(src.slice(src.indexOf('function initGuide'), src.indexOf('function boot'))),
    'ESC 也能关弹层');
  // 弹层里如果已经在显示别的东西（比如刚弹了确认框），就别再压一层
  chk(/classList\.contains\('show'\)[\s\S]{0,60}showGuide\(\)/.test(src),
    '已经有弹层在显示时不抢着弹（先看 overlay 有没有 show）');

  // 说明书是「一栏长文」，要有自己的宽度和「只要一个知道了」
  chk(/noCancel/.test(src) && /modal-cancel'\)\.style\.display = o\.noCancel/.test(src),
    'showModal 支持 noCancel（说明书只留一个「知道了」）');
  chk(/modal-doc/.test(src) && /\.modal\.modal-doc\s*\{[^}]*width:\s*780px/.test(extra),
    '说明书走 .modal-doc 宽版单栏（780px）');
  chk(/\.link-btn\s*\{/.test(extra), '副标题里那个文字链按钮有样式（否则是个灰底方块）');

  // 内容结构：每一节都要有标题和正文；目录胶囊数 = 节数
  const g = box.guideHtml();
  chk(box.GUIDE.length >= 8, `说明书至少 8 节（实际 ${box.GUIDE.length} 节）`);
  chk((g.match(/class="g-sec"/g) || []).length === box.GUIDE.length, '每一节都渲染出来了');
  chk((g.match(/data-g-jump="/g) || []).length === box.GUIDE.length,
    '目录胶囊数 = 节数（点一下跳到那一节）');
  chk(/class="g-lead"/.test(g) && /class="g-toc"/.test(g), '开头有引导语 + 目录');
  chk(box.GUIDE.every((x) => x.t && x.h && x.h.length > 40), '每节都是实打实的正文，不是空壳');
  chk(/\.g-sec\s+h3/.test(extra) && /\.g-toc button/.test(extra) && /\.g-tip\s*\{/.test(extra),
    '说明书的排版样式（节标题 / 目录 / 提示块）在 extra.css 里都有定义');
  chk(/class="g-k"/.test(g) && /\.g-k\s*\{/.test(extra),
    '正文里的按钮名用 .g-k 标出来，跟普通文字区分');

  // ⭐ 说明书里提到的界面控件，界面上必须真的有 ——
  //   防的是「说明书写着点某某按钮，可界面上根本没这颗按钮」（说明书最容易烂在这儿）
  const names = [...g.matchAll(/<span class="g-k">([^<]+)<\/span>/g)].map((m) => m[1]);
  const missNames = [...new Set(names)].filter((n) => !html.includes(n));
  chk(names.length > 15 && !missNames.length,
    `说明书提到的 ${names.length} 处控件名在 index.html 里都找得到`, JSON.stringify(missNames));

  let threw = '';
  try { box.showGuide(); } catch (e) { threw = String(e && e.message); }
  chk(!threw, 'showGuide() 跑得通不抛错', threw);
  box.markGuideSeen();
  chk(global.localStorage.getItem(box.LS.guideSeen) === '1',
    'markGuideSeen() 真把「看过了」写进去了（不然下次进来还弹）');

  // 入库页副标题：原来又长又写着「登记到 Excel 台账」，而台账真身早就换成 SQLite 了
  const sub = (html.match(/id="page-in"[\s\S]*?class="page-sub">([\s\S]*?)<\/div>/) || ['', ''])[1];
  const subText = sub.replace(/<[^>]+>/g, '').replace(/\s+/g, '');
  chk(subText.length > 0 && subText.length <= 70,
    `入库页副标题精简到一行（${subText.length} 字）`, sub);
  chk(!/Excel/.test(sub), '副标题不再写「登记到 Excel 台账」（台账真身是 SQLite，这句是过时的）');
  chk(/#page-in\s*\{\s*max-width:\s*880px/.test(extra),
    '入库页栏宽收到 880（原来 1120 → 卡片拉满、右边空一大片，用户说「不能居中吗」）');
}

console.log('— 台账表格 —');
box.state.rows = [train, einv];
box.renderLedgerBody();
const bodyHtml = els['ledger-body'].innerHTML;
// 分组头（.grp-row）只占一格，不算「一行几个单元格」
const rowHtmls = bodyHtml.split('<tr ').slice(1).filter((h) => !/^class="grp-row/.test(h));
const perRow = rowHtmls.map((h) => (h.match(/<td[ >]/g) || []).length);
chk(perRow.every((n) => n === 7), `每行 7 个 td（实际 ${JSON.stringify(perRow)}）`, bodyHtml.slice(0, 500));
chk(bodyHtml.includes('grp-row') && bodyHtml.includes('grp-todo') && bodyHtml.includes('grp-done'),
  '列表按状态分了组');
chk(bodyHtml.indexOf('待报销</span>') < bodyHtml.indexOf('grp-done'),
  '「待报销」那组排在「已报销」前面（先干活的）');
// 表头列数必须和行内单元格数一致，否则表格一定错位
const ledgerHead = html.split('id="ledger-body"')[0].split('<table').pop();
const headCols = (ledgerHead.match(/<th[ >]/g) || []).length;
chk(headCols === 7, `台账表头 7 列（实际 ${headCols}）`);
const repHead = html.split('id="rep-body"')[0].split('<table').pop();
const repCols = (repHead.match(/<th[ >]/g) || []).length;
const repRowCells = ((html.split('id="rep-body"')[1] || '').match(/colspan="(\d+)"/) || [])[1];
chk(repCols === Number(repRowCells), `报销单表头 ${repCols} 列 = 空态 colspan ${repRowCells}`);
chk(bodyHtml.includes('ct-train') && bodyHtml.includes('ct-inv'), '凭证类型色标有渲染');
chk(bodyHtml.includes('G1234 武汉-宜昌东'), '行程明细并进「事项 / 行程」格里了');
chk(bodyHtml.includes('湖北恩施建始龙坪网格'), '项目/事由列还在');
chk(bodyHtml.includes('td-sub'), '事项格里还有一行副行（行程/销方），信息没丢');
chk(bodyHtml.includes('td-dash'), '空字段显示占位符而不是空白');
chk(bodyHtml.includes('tag-ok') && bodyHtml.includes('tag-mute'), '已/未报销标签都在');
chk(bodyHtml.includes('tag-dup'), '提示含「重复」时打重复标签');
// 界面上不再平铺的字段（发票号/费用类别/批次/税号）不该出现在行里 —— 它们进了详情弹层
chk(!bodyHtml.includes('26349119423005870096'), '发票号不再挤在列表里（点行看详情）');

console.log('— 查询条件收集（这次修的重点） —');
const FILTER_IDS = ['f-status', 'f-ctype', 'f-category', 'f-batch', 'f-person',
                    'f-month', 'f-min', 'f-max', 'f-kw'];
const missing = FILTER_IDS.filter((id) => !html.includes(`id="${id}"`));
chk(!missing.length, '每个筛选控件在 index.html 里都存在', missing);
setVal('f-ctype', '火车票');
setVal('f-batch', '9月现场审计');
setVal('f-person', '张三');
setVal('f-min', '100');
setVal('f-max', '2000');
setVal('f-kw', '宜昌');
setVal('f-month', '2025-03');
setVal('f-status', '未报销');
setVal('f-category', '交通费');
const cf = box.currentFilter();
chk(cf.ctype === '火车票',
  'currentFilter 带上了「凭证类型」（原先漏了 → 下拉选了没反应）', JSON.stringify(cf));
chk(cf.batch === '9月现场审计' && cf.person === '张三', '带上了「报销批次 / 报销人」');
chk(cf.min === '100' && cf.max === '2000', '带上了「金额区间」');
chk(cf.status === '未报销' && cf.category === '交通费' && cf.month === '2025-03' && cf.kw === '宜昌',
  '状态 / 类别 / 月份 / 关键词都在', JSON.stringify(cf));
const fs2 = box.filterSummary(cf);
chk(fs2.includes('凭证类型：火车票') && fs2.includes('金额：100 ~ 2000') && fs2.includes('报销批次'),
  '条件摘要翻译成人话', fs2);
FILTER_IDS.forEach((id) => { if (!getVal(id)) setVal(id, 'x'); });
const cf2 = box.currentFilter();
chk(FILTER_IDS.every((id) => {
  const key = { 'f-status': 'status', 'f-ctype': 'ctype', 'f-category': 'category',
                'f-batch': 'batch', 'f-person': 'person', 'f-month': 'month',
                'f-min': 'min', 'f-max': 'max', 'f-kw': 'kw' }[id];
  return cf2[key] !== undefined && cf2[key] !== null;
}), '每个筛选控件都映射到了 filter 的一个键（没有控件是「摆设」）');

console.log('— 选中提示条 —');
box.state.rows = [train, einv];
box.state.sel = new Set(['5']);
box.syncSelectionUI();
chk(els['ledger-hint'].textContent.includes('火车票 1 / 发票 1'),
  '顶部提示带按类型分布', els['ledger-hint'].textContent);
chk(els['sel-count'].innerHTML.includes('1'), '选中数正确');

console.log('— 删除选中 / 清空台账（二次确认） —');
box.state.sel = new Set(['5']);
box.deleteSelected();
chk(els['modal-title'].textContent.includes('删除选中的 1 条记录'), '删除选中要先确认',
  els['modal-title'].textContent);
chk(els['modal-text'].innerHTML.includes('不会动原始发票文件'), '明确说明只删记录、不动原始文件');
chk(els['modal-text'].innerHTML.includes('79.00'), '确认框里带出这批合计金额',
  els['modal-text'].innerHTML);
chk(els['modal-note'].innerHTML.includes('台账备份'), '提示会自动备份',
  els['modal-note'].innerHTML);
box.state.sel = new Set();
box.deleteSelected();
chk(els['modal-title'].textContent.includes('删除选中的 1 条记录'),
  '没勾选时不弹删除框（标题保持不变）', els['modal-title'].textContent);
box.state.allCount = 88;
box.clearLedger();
chk(els['modal-title'].textContent.includes('清空整张台账'), '清空台账有二次确认',
  els['modal-title'].textContent);
chk(els['modal-text'].innerHTML.includes('88'), '清空确认框里报出当前行数');

console.log('— 明细弹层：火车票（有行程、无发票号） —');
box.showInvoiceDetail(train);
let m = els['modal-text'].innerHTML;
chk(m.includes('class="id-top"') && m.includes('id-amt-n'), '顶部号码+金额大数字块在');
chk(m.includes('无发票号码'), '无号凭证有友好占位');
chk(m.includes('id-trip') && m.includes('G1234 武汉-宜昌东'), '行程明细整段显示');
chk(m.includes('行程 / 明细'), '有「行程 / 明细」分区');
chk(m.includes('data-act="open_file"') && m.includes('data-act="reveal"'), '文件区带两个按钮');
chk(m.includes('data-act="del_row"') && m.includes('从台账删除这一条'),
  '明细弹层里有「从台账删除这一条」');
chk(m.includes('E:\\桌\\湖北报销发票整理'), '完整路径单独一行显示');
chk(!/id-v">\s*<\/div>/.test(m), '空字段没有渲染成空格子');
chk((m.match(/id-sec-t/g) || []).length === 6, '分区数=6（基本信息/行程/金额/归属/文件/操作）实际 ' +
  (m.match(/id-sec-t/g) || []).length, m.match(/id-sec-t[^<]*<\/div>/g));
chk(!m.includes('undefined') && !m.includes('null'), '没有 undefined/null 漏出来');
chk(els['modal-title'].textContent.includes('第 5 行'), '标题带行号');

console.log('— 明细弹层：发票（无行程、有项目、已报销、重复提示） —');
box.showInvoiceDetail(einv);
m = els['modal-text'].innerHTML;
chk(!m.includes('id-trip'), '没行程就不显示行程明细');
chk(m.includes('tag-ok') && m.includes('tag-dup'), '已报销 + 重复 两个标签都在');
chk(m.includes('湖北恩施建始龙坪网格'), '项目/事由进明细');
chk(m.includes('1815P825000P'), '项目编号进明细');
chk(m.includes('id-money'), '金额区突出显示');
chk((m.match(/id-sec-t/g) || []).length === 6, '分区数=6（基本信息/金额/归属/提示/文件/操作）实际 ' +
  (m.match(/id-sec-t/g) || []).length, m.match(/id-sec-t[^<]*<\/div>/g));

console.log('— 明细弹层：脏数据（字段全是 undefined） —');
let dirtyOk = true;
try { box.showInvoiceDetail({ '序号': 9 }); } catch (e) { dirtyOk = false; console.log('   抛错：' + e); }
chk(dirtyOk, '缺字段不抛异常');
chk(!els['modal-text'].innerHTML.includes('undefined'), '缺字段不会渲染出 undefined');

console.log('— 单独删一条（明细弹层里点的） —');
box.state.rows = [train, einv];
box.deleteOneRow('6');
chk(els['modal-title'].textContent.includes('删除这一条记录'), '单条删除也要确认',
  els['modal-title'].textContent);
chk(els['modal-text'].innerHTML.includes('26349119423005870096'), '确认框里带出这条的关键信息',
  els['modal-text'].innerHTML);

console.log('— 报销单输出格式 —');
chk(['pdf', 'xlsx', 'both'].every((f) => html.includes(`data-fmt="${f}"`)),
  'HTML 里有 PDF / Excel / 两种都出 三档');
setVal('f-min', ''); setVal('f-max', '');
box.state.fmt = 'xlsx';
box.syncFmtPills();
chk(els['btn-make-report'].textContent.includes('Excel'),
  '按钮文案跟着格式变', els['btn-make-report'].textContent);
box.state.fmt = 'both';
box.syncFmtPills();
chk(els['btn-make-report'].textContent.includes('PDF + Excel'), '两种都出文案对');
box.state.fmt = 'pdf';

console.log('— 生成文件列表（带删除） —');
box.state.outputs = [
  { path: 'E:\\out\\费用报销单_张三.pdf', name: '费用报销单_张三.pdf', kind: '费用报销单',
    time: '2026-09-15 13:00:00', size: 153600, sum: 1252.58, count: 4, exists: true },
  { path: 'E:\\out\\发票合并.pdf', name: '发票合并.pdf', kind: '发票合并',
    time: '2026-09-15 13:05:00', size: 0, pages: 6, exists: false },
];
box.renderOutputs();
const oh = els['out-list'].innerHTML;
chk((oh.match(/data-act="del"/g) || []).length === 2, '每个生成文件都有「删除」按钮');
chk((oh.match(/data-act="open"/g) || []).length === 2, '「打开」按钮还在');
chk(oh.includes('out-gone') && oh.includes('已不在'), '文件已不在的会置灰并标出来');
chk(oh.includes('4 条') && oh.includes('150 KB') && oh.includes('6 页'),
  '文件行带条数/大小/页数摘要', oh.slice(0, 500));
chk(els['out-hint'].textContent.includes('共 2 个') && els['out-hint'].textContent.includes('1 个还在'),
  '提示条统计对', els['out-hint'].textContent);
box.deleteOutput(box.state.outputs[0]);
chk(els['modal-title'].textContent.includes('删除这个生成的文件'), '删除文件有二次确认',
  els['modal-title'].textContent);
chk(els['modal-note'].innerHTML.includes('回收站'), '说明是送回收站而不是永久删除',
  els['modal-note'].innerHTML);
box.state.outputs = [];
box.clearOutputs();
chk(!els['modal-title'].textContent.includes('删除全部'), '没有可删文件时不弹「全部删除」框',
  els['modal-title'].textContent);
box.state.outputs = [{ path: 'E:\\out\\a.pdf', name: 'a.pdf', kind: '费用报销单', exists: true }];
box.clearOutputs();
chk(els['modal-title'].textContent.includes('删除全部 1 个生成文件'), '「全部删除」也走确认',
  els['modal-title'].textContent);

console.log('— 合并少了几张的提示 —');
// 这次修的点：PDF 文件名不含发票号时按票面内容认票；认不出来的必须点名+说原因，
// 不能只闪一条 toast（用户就是被「选 4 张只合进去 1 张」坑了）
const missHtml = box.missListHtml([
  { name: '26419165773005057163.ofd', no: '26419165773005057163',
    why: '只有 OFD 版式文件，附近也没找到同一张票的 PDF' },
]);
chk(missHtml.includes('miss-list'), '少的那几张用小列表列出来');
chk(missHtml.includes('26419165773005057163.ofd'), '列表里点出了文件名');
chk(missHtml.includes('只有 OFD 版式文件'), '列表里写了原因');
chk(box.missListHtml([]) === '' && box.missListHtml(null) === '', '没有缺失时不产生空列表');
chk(src.includes('合并好了，但有'), '少几张时弹层会说明（不是一闪而过的 toast）');
chk(!src.includes('张找不到 PDF 被跳过'), '旧的「只提示不说明」写法已清掉');
chk(src.includes('missListHtml(r.missing_list)'), '弹层内容用的是同一个列表函数');
chk(src.includes('missing_list'), '结果里读了后端的 missing_list');
chk(/o\.missing \? `少 \$\{o\.missing\} 张`/.test(src), '生成文件列表里标出「少 N 张」');

console.log('— 入库结果卡片 —');
box.onImportResult({
  added: 5, in_ledger: 2, dup: 1, manual: 2, total: 88, files: 132, pending_amount: 1234.5,
  patched: 3, by_type: { '火车票': 3, '发票': 2 },
  manual_list: [{ name: 'a.jpg', why: '没找到 20 位发票号码' }],
  dup_list: [{ no: '火车票', tip: '与第 3 行重复' }],
}, true);
chk(els['in-stats'].innerHTML.includes('新增入库'), '统计方块渲染');
chk(els['in-hint'].innerHTML.includes('本次新增构成') && els['in-hint'].innerHTML.includes('火车票 3'),
  '新增构成按凭证类型拆分', els['in-hint'].innerHTML);
chk(els['in-hint'].innerHTML.includes('补全老记录 3 条'), '老记录补全条数有显示');
chk(els['dup-body'].innerHTML.includes('第 3 行'), '重复清单渲染');

console.log('— 入库必须用界面上那个路径 —');
// 用户报过的 bug：界面上把路径改成别的目录，点入库却还是扫老目录（后端只认配置文件）。
// 这里直接盯住那句调用：第 5 个实参必须是 state.src。
const importCall = src.match(/call\(\s*'start_import'[^)]*\)/);
chk(!!importCall, 'app.js 里能找到 start_import 调用');
chk(!!importCall && importCall[0].includes('state.src'),
  'start_import 把界面路径（state.src）传给了后端', importCall && importCall[0]);
// 源码里也不该再出现"不传路径"的旧写法
chk(!/'start_import',\s*state\.recursive,\s*''\s*,\s*person\s*,\s*batch\s*\)/.test(src),
  '没有残留的 4 参调用（会退回配置里的旧目录）');

console.log('— 行程单跟在发票后面（附件标记 + 不重复计钱） —');
// 用户要求：「打车行程单要跟在那个发票之后……他是证明这个发票的附件要在一块」
const invRow = {
  '序号': 11, '状态': '未报销', '凭证类型': '发票', '发票号码': '26372000004097056996',
  '开票日期': '2026-08-20', '票种': '', '行程/明细': '', '购方名称': '某某公司',
  '销方名称': '某某汽车服务有限公司', '不含税金额': '', '税额': '', '价税合计': '76.10',
  '项目/事由': '某某科技公司-烟台站打车票', '项目编号': '', '费用类别': '交通费',
  '报销人': '赵大勇', '报销批次': '', '入库时间': '2026-09-15 15:00',
  '文件路径': 'E:\\桌\\报销\\赵大勇\\某某科技-烟台站打车票.pdf', '提示': '',
};
const tripRow = {
  '序号': 12, '状态': '未报销', '凭证类型': '打车行程单', '发票号码': '',
  '开票日期': '2026-08-20', '票种': '', '行程/明细': '悦行优选  某某科技→烟台站  76.10元',
  '购方名称': '', '销方名称': '', '不含税金额': '', '税额': '', '价税合计': '76.10',
  '项目/事由': '', '项目编号': '', '费用类别': '交通费', '报销人': '赵大勇',
  '报销批次': '', '入库时间': '2026-09-15 15:00',
  '文件路径': 'E:\\桌\\报销\\赵大勇\\某某科技-烟台站行程报销单.pdf', '提示': '',
};
box.state.rows = [invRow, tripRow];
box.state.attachLinks = { '12': { main_seq: '11', main_amount: 76.1, main_desc: '打车票' } };
box.state.sel = new Set(['11', '12']);
box.renderLedgerBody();
const lh = els['ledger-body'].innerHTML;
chk(lh.includes('att-mark') && lh.includes('↳'), '台账里附件行标了 ↳', lh.slice(0, 300));
chk(lh.includes('tag-att'), '凭证类型后面跟了「附件」标签');
chk(lh.includes('row-att'), '附件行有单独的底色 class');
chk(lh.includes('第 11 号发票'), '鼠标停上去能看出是谁的附件', lh.slice(0, 400));

box.state.selRows = [invRow, tripRow];
box.renderReportRows();
const rh = els['rep-body'].innerHTML;
chk(rh.includes('↳ 附件'), '报销单预览里也标了「↳ 附件」');
chk(rh.includes('row-att'), '预览里附件行有底色');
chk(els['rep-total'].innerHTML.includes('76.10') && !els['rep-total'].innerHTML.includes('152.20'),
  '金额没重复计（附件那 76.10 不算）', els['rep-total'].innerHTML);
chk(els['rep-total'].innerHTML.includes('不重复计'), '并且写明了原因', els['rep-total'].innerHTML);
box.state.selRows = [tripRow];            // 只选行程单（对应发票没选）
box.renderReportRows();
chk(!els['rep-body'].innerHTML.includes('tag-att'), '发票不在选择里时，不标附件');
chk(els['rep-total'].innerHTML.includes('76.10'), '这时行程单自己算一张凭证（76.10）',
  els['rep-total'].innerHTML);

console.log('— 差旅费补助（人数 × 天数 × 标准，自己填） —');
chk(html.includes('id="r-alw-on"') && html.includes('id="alw-body"') && html.includes('id="btn-alw-add"'),
  'index.html 里有补助控件（勾选框 / 表格 / 加行按钮）');
chk(/initAlw\(\)/.test(src), '初始化里调了 initAlw（不然表是死的）');
chk(/addEventListener\('input'/.test(src) && /data-act="alw-del"/.test(src),
  '补助表的输入 / 删行事件有绑定');

/* ---------- ⭐ 补助不许「自己填好」（用户 2026-09-17 报的 bug） ----------
 * 用户原话：「出差补助，我之前填了怎么会保存上了，我这次啥也没填的，就都填好了，
 *           出差人员怎么也直接自己输入完了，就是会有残留信息」。
 * 根因：这三样被存进 localStorage（invoice-rp.alw / .alwOn / .jyAlw）并在下次进来读回来。
 * 补助会跟着进单据 → 那是**算错钱**，不是省事。下面这组断言把它钉死。 */
console.log('\n— 补助不做「记忆」（不残留上一次的值）—');
chk(box.state.alw.length === 0, '打开页面时补助表是空的（不读上次的值）', box.state.alw.length);
chk(box.state.alwOn === false, '「计入差旅费补助」也不会被上次的勾选带回来');
chk(Object.keys(box.state.jyAlw).length === 0, '按人补助的金额同样是空的');
// 只认「读 / 写」这种用法；启动时**清掉**老键是另一回事（下面单独测）
chk(!/getItem\(\s*'invoice-rp\.(alw|alwOn|jyAlw)'/.test(src)
  && !/setItem\(\s*'invoice-rp\.(alw|alwOn|jyAlw)'/.test(src),
  'localStorage 里不再读也不再写补助的键（alw / alwOn / jyAlw）');
// 老版本已经写进去的值要主动清掉，不能只是「不读」
const _ls = global.localStorage;
_ls.setItem('invoice-rp.alw', '[{"rate":"100"}]');
_ls.setItem('invoice-rp.alwOn', '1');
_ls.setItem('invoice-rp.jyAlw', '{"张三":500}');
box.purgeLegacyAllowanceKeys();
chk(_ls.getItem('invoice-rp.alw') === null && _ls.getItem('invoice-rp.alwOn') === null
  && _ls.getItem('invoice-rp.jyAlw') === null,
  '启动时把老版本残留的三个键从浏览器里清掉（不是只不读）');
chk(!/loadAlw|saveAlw|loadJyAlw|saveJyAlw/.test(src),
  '装载 / 保存补助的那几个函数整个删掉，不留半死不活的空壳');
// 「出差人员」名单只认本次勾选的凭证
box.state.jyPeople = [];
box.state.jyAlw = { 张三: 500 };
chk(box.jyNames().length === 0,
  '没勾凭证时「出差人员」一个都不冒出来（以前会把上次填过的人补回来）', JSON.stringify(box.jyNames()));
chk(box.jyPayload().length === 0, '所以也不会把上次的人发给后端');
box.state.jyPeople = [{ name: '李强', bills: 100 }];
box.state.jyAlw = { 李强: 200, 王五: 999 };
box.pruneJyAlw();
chk(box.state.jyAlw['王五'] === undefined && box.state.jyAlw['李强'] === 200,
  '取消勾选的人，他的补助金额会被丢掉（不挂在表里）', JSON.stringify(box.state.jyAlw));
// 清空 / 出单后归零
box.state.alwOn = true;
box.state.alw = [{ name: '伙食补助费', people: '3', days: '4', rate: '100', amount: '', manual: false }];
box.clearAlw();
chk(box.state.alw.length === 0 && box.state.alwOn === false && box.alwSum() === 0,
  '「清空」按钮：行清掉、勾选取消、合计归零');
box.state.alwOn = true;
box.state.alw = [{ name: '伙食补助费', people: '3', days: '4', rate: '100', amount: '', manual: false }];
box.state.jyPeople = [{ name: '李强', bills: 100 }];
box.state.jyAlw = { 李强: 800 };
box.resetAllowances();
chk(box.alwSum() === 0 && box.alwRowAmount(box.state.alw[0]) === 0,
  '生成成功后补助自动归零（金额不会流到下一张单）', JSON.stringify(box.state.alw));
chk(box.state.alwOn === true, '归零时保留「计入补助」的勾选（不用重新勾一次）');
chk(box.jySum() === 0 && Object.keys(box.state.jyAlw).length === 0, '按人补助也一起清零');
chk(html.includes('id="btn-alw-clear"') && html.includes('id="btn-jy-clear"'),
  '两块补助各有「清空」入口');
chk(/btn-alw-clear/.test(src) && /btn-jy-clear/.test(src), '两个清空按钮都绑了事件');
chk(/resetAllowances\(\)/.test(src), 'makeReport 生成成功后调用了归零');
// 还原，别影响后面的用例
box.state.alw = []; box.state.alwOn = false; box.state.jyPeople = []; box.state.jyAlw = {};

box.state.alwOn = true;
box.state.alw = [{ name: '伙食补助费', people: '3', days: '4', rate: '100', amount: '', manual: false }];
chk(box.alwRowAmount(box.state.alw[0]) === 1200, '3 人 × 4 天 × 100 = 1200',
  box.alwRowAmount(box.state.alw[0]));
box.state.alw[0].manual = true;
box.state.alw[0].amount = '156.5';
chk(box.alwRowAmount(box.state.alw[0]) === 156.5, '手改过的金额优先（156.50）',
  box.alwRowAmount(box.state.alw[0]));
box.state.alw[0].manual = false;
box.state.alw.push({ name: '市内交通费', people: '2', days: '3', rate: '80' });
chk(box.alwSum() === 1680, '补助合计 = 1200 + 480 = 1680', box.alwSum());
const pl = box.alwPayload();
chk(pl.length === 2 && pl[0].name === '伙食补助费' && pl[0].amount === 1200 && pl[1].amount === 480,
  '提交给后端的内容对', JSON.stringify(pl));
box.state.alw.push({ name: '', people: '', days: '', rate: '' });
chk(box.alwPayload().length === 2, '空行不发给后端', JSON.stringify(box.alwPayload()));
box.state.alw = box.state.alw.slice(0, 2);
box.state.selRows = [tripRow];
box.renderReportRows();
chk(els['rep-total'].innerHTML.includes('1,680.00'), '预览里看得见补助合计',
  els['rep-total'].innerHTML);
chk(els['rep-total'].innerHTML.includes('报销合计'), '预览里给了「凭证 + 补助 = 报销合计」');
const meta2 = box.reportMeta();
chk(meta2.allowance && meta2.allowance.length === 2 && meta2.allowance[0].amount === 1200,
  'reportMeta 把 allowance 带上了', JSON.stringify(meta2.allowance));
box.state.alwOn = false;
chk(!box.reportMeta().allowance, '没勾选时就不带（不写进单据）');
box.state.alwOn = false;
box.state.alwOff = false;
box.state.kind = '差旅费报销单';
box.syncKindPills();
chk(box.state.alwOn === true, '切到「差旅费报销单」自动把补助打开（少点一下）');
chk(/r\.linked/.test(src), '合并结果里读了后端的 linked（几组发票+行程单挨在一起）');
chk(/att\.0?attach_links|attach_links/.test(src), '台账读了后端的 attach_links（附件标记）');

/* ---------- 票面标记（退票 / 改签 / 红字）与「重新识别」 ---------- */
console.log('\n--- 票面标记 / 重新识别 ---');
chk(typeof box.stampTag === 'function', '导出了 stampTag');
const tRef = box.stampTag('退票');
chk(tRef.includes('tag-stamp') && tRef.includes('stamp-red') && tRef.includes('退票'),
  '退票渲染成红色标签', tRef);
chk(box.stampTag('差额退票').includes('stamp-red'), '差额退票也是红色');
chk(box.stampTag('改签').includes('stamp-amber'), '改签渲染成橙色标签', box.stampTag('改签'));
chk(box.stampTag('') === '' && box.stampTag(null) === '', '没有标记时不渲染空标签');

box.state.rows = [Object.assign({}, train, { '票面标记': '差额退票' }), einv];
box.state.attachLinks = {};
box.state.sel = new Set();
box.renderLedgerBody();
chk(els['ledger-body'].innerHTML.includes('tag-stamp'), '台账行里画出了票面标记标签',
  els['ledger-body'].innerHTML.slice(0, 220));
chk(els['ledger-body'].innerHTML.includes('差额退票'), '标签上写的就是票面标记本身');
const markRows = els['ledger-body'].innerHTML
  .split('<tr ').slice(1).filter((h) => !/^class="grp-row/.test(h));
chk(!/tag-stamp/.test(markRows[1] || ''),
  '没有标记的那一行不会多出标签');

box.state.selRows = [Object.assign({}, train, { '票面标记': '改签' })];
box.state.attachLinks = {};
box.renderReportRows();
chk(els['rep-body'].innerHTML.includes('tag-stamp'), '报销单预览里也带上票面标记',
  els['rep-body'].innerHTML.slice(0, 200));

let noThrow = true;
try { box.showInvoiceDetail(Object.assign({}, train, { '票面标记': '退票' })); }
catch (e) { noThrow = false; console.log('   ' + e.message); }
chk(noThrow, '带票面标记的行打开明细弹层不报错');

chk(typeof box.refreshSelected === 'function', '导出了 refreshSelected');
chk(/'btn-reparse'/.test(src), '界面绑定了「重新识别」按钮');
chk(/call\('refresh_rows', seqs\)/.test(src), '重新识别把选中序号传给后端 refresh_rows');
chk(/r\.miss/.test(src) && /r\.fail/.test(src), '结果里读了 miss / fail（原件找不到、读不了的条数）');
chk(/重新识别/.test(src), '按钮与确认弹层上写的是人话「重新识别」');

/* =====================================================================
 * 出行人（2026-09-15 新增）
 * 用户：打车行程单票面没有人名 → 三层认人 + 认不出就人工填
 * ===================================================================== */
console.log('— 出行人 —');
box.state.rows = [train, einv];
box.state.attachLinks = {};
box.state.sel = new Set();
box.renderLedgerBody();
const lHtml = els['ledger-body'].innerHTML;
chk(lHtml.includes('出行人'), '台账表头/行里有「出行人」这一列');
chk(lHtml.includes('张三'), '票面的出行人渲染进表格');
chk(/td-nowrap td-c/.test(lHtml), '出行人这一列有自己的单元格类');
chk((lHtml.match(/td-dash/g) || []).length >= 1,
  '出行人没认出来的那张显示占位符而不是空白');
// 事项 / 行程 / 销方 / 出行人全空的一行，也不该留白
box.state.rows = [Object.assign({}, einv, {
  '项目/事由': '', '行程/明细': '', '销方名称': '', '出行人': '' })];
box.renderLedgerBody();
chk((els['ledger-body'].innerHTML.match(/td-dash/g) || []).length >= 2,
  '整行都没内容时，事项格和出行人列都给占位符');
box.state.rows = [train, einv];
box.renderLedgerBody();
chk(/设置出行人/.test(lHtml) || /设置出行人/.test(html), '占位符/条目里指到了「设置出行人」这个动作');

// 明细弹层：有人名 → 显示人名；票面没印人名 → 4 处都提「未识别」并指向手填
let throwA = true, detHtml = '';
const origShow = els['modal-text'];
try {
  box.showInvoiceDetail(Object.assign({}, train, { '出行人': '李贵清' }));
  detHtml = origShow.innerHTML;
} catch (e) { throwA = false; console.log('   ' + e.message); }
chk(throwA, '带出行人的明细弹层不报错');
chk(detHtml.includes('李贵清'), '明细弹层里显示出行人姓名');

try {
  box.showInvoiceDetail(Object.assign({}, train, { '出行人': '', '提示': '出行人「张三」按文件夹名推断，请核对' }));
  detHtml = origShow.innerHTML;
} catch (e) { console.log('   ' + e.message); }
chk(detHtml.includes('未识别'), '行程单没认出行人时明细里写「未识别」而不是空着');
chk(/按文件夹名推断/.test(src) && detHtml.includes('推断'), '按文件夹推出来的人名会打「推断」标并提示核对');

// 手填：弹层带输入框、快捷按钮、留空 = 清空；保存走 update_rows
box.state.sel = new Set(['5', '6']);
box.setTraveler();
const modalHtml = els['modal-text'].innerHTML;
chk(/设置出行人（2 张）/.test(els['modal-title'].textContent), '点「设置出行人」会弹出对应张数的弹层',
  els['modal-title'].textContent);
chk(/id="tv-input"/.test(modalHtml), '弹层里有姓名输入框');
chk(/data-tv="张三"/.test(modalHtml), '弹层里给出台账里已知的人名做快捷选择');
chk(/留空/.test(modalHtml), '弹层里写明留空 = 清空');
chk(/'btn-set-tv'/.test(src) && /setTraveler/.test(src), '台账工具条绑定了「设置出行人」按钮');
chk(/call\('save_traveler_review'/.test(src),
  '保存出行人走的是 save_traveler_review —— 手工指定也算「人工核对过」，不会被核对面板再提醒一遍');

// 手机号对照表：设置页可读写
chk(/id="phone-map"/.test(html), '设置页有手机号对照表的输入框');
chk(/'btn-save-pmap'/.test(src) && /'btn-reload-pmap'/.test(src), '对照表有保存/重载按钮');
chk(/call\('get_phone_map'\)/.test(src), '读对照表调 get_phone_map');
chk(/call\('set_phone_name'/.test(src), '写对照表调 set_phone_name');
chk(/pageId === 'page-set'\) \{[^}]*loadPhoneMap\(\)/.test(src),
  '切到设置页会自动载入对照表');
chk(/pageId === 'page-set'\) \{[^}]*loadSettingsExtras\(\)/.test(src),
  '切到设置页也会载入「账号与权限 / 台账数据库」那块');
// 解析「手机号 姓名」文本：空格 / 等号 / 逗号都要能认
const parsed = {};
'13800001111 王小明\n13900000000=赵大勇\n13800000000，张三\n\n乱七八糟'.split('\n').forEach((ln) => {
  const t = ln.replace(/[=,，\t]+/g, ' ').trim();
  const m = t.match(/^(\d{6,11})\s*(.*)$/);
  if (m) parsed[m[1]] = m[2].trim();
});
chk(Object.keys(parsed).length === 3, '对照表三种写法都解析出来了', parsed);
chk(parsed['13800001111'] === '王小明' && parsed['13900000000'] === '赵大勇'
  && parsed['13800000000'] === '张三', '解析结果对得上', parsed);
chk(!parsed['乱七八糟'], '认不出来的行不会被当成手机号吞掉');

/* ---------- 出行人核对：三层兜底之后那一步「人工核对」 ---------- */
console.log('\n— 出行人核对面板 —');
chk(/id="btn-review-tv"/.test(html), '工具条上有「出行人核对」按钮');
chk(/id="tv-todo-badge"/.test(html), '按钮上挂着待核对角标');
chk(/id="tv-people"/.test(html), '页面里有姓名 datalist（输入框自动补全）');
chk(/modal-xl/.test(src) && /modal-xl/.test(fs.readFileSync(path.join(GUI, 'extra.css'), 'utf8')),
  '核对弹层用了更宽的 modal-xl（showModal 支持 xwide）');

box.syncTvBadge(7);
chk(els['tv-todo-badge'].textContent === '7' && els['tv-todo-badge'].hidden === false,
  '角标显示待核对张数', els['tv-todo-badge']);
box.syncTvBadge(0);
chk(els['tv-todo-badge'].hidden === true, '核对完归零就把角标藏起来');
chk(/syncTvBadge\(r\.tv_todo\)/.test(src), '角标数来自后端 get_ledger 的 tv_todo');

// 「凭什么是这个人」的各种来源
chk(box.tvrBadge({ done: true, cur: '王小明' })[1] === '已核对', '已核对 → 绿标签');
chk(box.tvrBadge({ state: 'blank' })[1] === '没认出来', '三层都没认出来 → 「没认出来」');
chk(box.tvrBadge({ state: 'blank' })[0] === 'tag-dup', '没认出来的用红标签（最急）');
chk(box.tvrBadge({ state: 'diff', cur: '张伟', auto: '李强', src: '文件夹' })[1] === '跟自动值不符',
  '台账值跟自动值打架 → 「跟自动值不符」');
chk(box.tvrBadge({ state: 'todo', auto: '李强' })[1] === '还没写进台账', '能认出来但没写 → 提示写入');
chk(box.tvrBadge({ src: '手机号' })[1] === '手机号译出', '手机号译的 → 标「手机号译出」');
chk(box.tvrBadge({ src: '文件夹' })[1] === '文件夹推断', '文件夹推的 → 标「文件夹推断」');
chk(box.tvrBadge({ src: '手机号' })[0] === 'tag-pending', '猜出来的用橙色标签（要看一眼）');
chk(box.tvrBadge({ src: '票面' })[0] === 'tag-ok', '票面白纸黑字写的 → 绿标签，默认不用核');

// 面板渲染
const tvRows = [
  { seq: '4', ctype: '打车行程单', date: '2026-08-21', total: 18, detail: '济南西-大明湖',
    person: '张三', no: '', cur: '', auto: '', src: '', state: 'blank', need: true,
    done: false, phone: '', file: '' },
  { seq: '2', ctype: '打车行程单', date: '2026-08-19', total: 13.58, detail: '合肥南-财智中心',
    person: '张三', no: '', cur: '王小明', auto: '王小明', src: '手机号', state: 'ok',
    need: true, done: false, phone: '13800000000', file: '' },
  { seq: '1', ctype: '火车票', date: '2026-08-19', total: 204, detail: 'G123 济南东-烟台',
    person: '张三', no: '', cur: '王小明', auto: '王小明', src: '票面', state: 'ok',
    need: false, done: false, phone: '', file: '' },
];
const mkTvr = (only) => ({
  rows: tvRows,
  stat: { total: 3, need: 2, blank: 1, done: 0, '票面': 1, '文件名': 0, '手机号': 1, '文件夹': 0 },
  people: ['王小明', '李强'], onlyTodo: only, syncPhone: true, remember: true, edits: {},
});
box.setTvrTest(mkTvr(true));
box.renderTvr();
const tvrHtml = els['tvr-box'].innerHTML;
const nInp = (h) => (h.match(/class="fi tvr-inp"/g) || []).length;
chk(nInp(tvrHtml) === 2, `勾着「只看要核对的」→ 只渲染 2 行（实际 ${nInp(tvrHtml)}）`, tvrHtml.slice(0, 400));
chk(tvrHtml.includes('13800000000'), '手机号摆出来了（好判断是哪部手机开的）');
chk(tvrHtml.includes('list="tv-people"'), '出行人输入框挂上了姓名补全');
chk(tvrHtml.includes('data-tvr="only"') && tvrHtml.includes('data-tvr="sync"')
  && tvrHtml.includes('data-tvr="remember"'),
  '三个开关都在（只看要核对 / 同手机号一起改 / 记住手机号）');
chk(tvrHtml.includes('2</b> 待核对'), '统计条上写着 2 张待核对', tvrHtml.slice(0, 300));
chk(tvrHtml.includes('id="tvr-bulk-apply"'), '有「批量填个名字，应用到勾选的」入口');
chk(/is-need/.test(tvrHtml), '要核对的行加了高亮类');

box.setTvrTest(mkTvr(false));
box.renderTvr();
chk(nInp(els['tvr-box'].innerHTML) === 3, '取消勾选 → 3 行全出来');

// 编辑过的值要带得走（保存时以输入框 / edits 为准）
const st = box.getTvrTest();
st.edits = { 4: '王小明' };
box.renderTvr();
chk(/value="王小明"/.test(els['tvr-box'].innerHTML), '改过的名字在重渲染后还在（edits 记着）');

// 老台账里出行人整列空着、但系统已经能认出来 → 先预填，用户扫一眼改错的就行
box.setTvrTest({
  rows: [{ seq: '9', ctype: '火车票', date: '2026-09-02', total: 198, detail: 'G1 济南-烟台',
           person: '', no: '', cur: '', auto: '王常凯', src: '票面', state: 'todo',
           need: true, done: false, phone: '', file: '' }],
  stat: { total: 1, need: 1, blank: 0, done: 0, '票面': 1, '文件名': 0, '手机号': 0, '文件夹': 0 },
  people: ['王常凯'], onlyTodo: true, syncPhone: true, remember: true, edits: {},
});
box.renderTvr();
const pre = els['tvr-box'].innerHTML;
chk(/value="王常凯"/.test(pre), '台账空着但能认出来 → 输入框预填系统认出来的名字（不用一张张手打）');
chk(/is-auto/.test(pre), '预填的格子用斜体虚线标出来（一看就知道是系统猜的）');
chk(/vals\[el\.dataset\.seq\]/.test(src), '保存以输入框里的值为准 —— 预填的、没动过的也照样落库');
chk(/is-auto/.test(fs.readFileSync(path.join(GUI, 'extra.css'), 'utf8')), '斜体预填样式在 extra.css 里有定义');

/* ---------- 模板二版式（出差报销明细）：补助按人 ---------- */
console.log('\n— 模板二版式（出差报销明细）—');
chk(html.includes('data-kind="模板二"'),
  '单据类型里多了「模板二」这一档');
chk(html.includes('id="jy-body"') && html.includes('id="btn-jy-fill"')
  && html.includes('id="jy-days"') && html.includes('id="jy-rate"'),
  '有「按人」的补助表和一键预填控件');
chk(/report_persons/.test(src),
  '界面会向后端要「按出行人的票据合计」（跟打印出来的单据同一个口径）');
chk(/call\('report_persons', seqs\)/.test(src), '要数时把选中的序号一起传过去');
chk(box.JY_KIND === '模板二', '版本名跟后端 report_pdf.JY_KIND 一致');

// 名单 / 金额：票据合计来自 state.jyPeople（后端算的），补助是用户填的
box.state.jyPeople = [{ name: '李强', bills: 1560 }, { name: '李贵清', bills: 0 }];
box.state.jyTotal = 1560;
box.state.jyAlw = { 李强: 1200 };
box.renderJy();
const jyHtml = els['jy-body'].innerHTML;
chk(jyHtml.includes('李强') && jyHtml.includes('李贵清'),
  '按出行人列出名单（后端给的人一个不少，票据合计为 0 的也在）', jyHtml.slice(0, 300));
chk(jyHtml.includes('1,560.00'), '票据合计显示后端给的数（不是界面自己估的）');
chk(/value="1200"/.test(jyHtml), '补助金额回填进输入框（本次会话内还在；页面不留存）');
chk(jyHtml.includes('2,760.00'), '实际＝票据合计＋补助（1560＋1200）', jyHtml.slice(-260));
chk(box.jyReal('李强') === 2760 && box.jyReal('李贵清') === 0, '没填补助的人，实际＝他自己的票据合计');
chk(box.jySum() === 1200, `补助合计只算补助（实际 ${box.jySum()}）`);
chk(box.jyPayload().length === 2, '提交时两个人都带上（0 补助的那位由后端过滤）');
chk(box.jyNames().length === 2, '名单去重、不带空名字');

// reportMeta：把按人补助带给后端
setVal('r-person', '');
setVal('r-dept', '');
setVal('r-reason', '国网湖北决算审核差旅费');
setVal('r-note', '');
setVal('r-date', '2026年09月15日');
box.state.kind = '模板二';
const jm = box.reportMeta();
chk(jm.kind === '模板二', 'kind 传的就是这一档名（等号两边要一模一样）');
chk(Array.isArray(jm.allowance_by) && jm.allowance_by[0].name === '李强'
  && jm.allowance_by[0].amount === 1200,
  'reportMeta 把按人补助（allowance_by）带上了', JSON.stringify(jm));
chk(!jm.allowance, '不会同时带老的 allowance —— 两套补助不混着发');
// 一分补助都没填 → 不带 allowance_by，后端也就不出那块
box.state.jyAlw = {};
chk(!box.reportMeta().allowance_by, '一个人都没填补助 → 不带，单据上不挂一排 0');

// 两种补助块是二选一的
box.state.kind = '模板二';
box.syncKindPills();
chk(els['jy-block'].hidden === false, '选了这个类型 → 「按人」那块显示出来');
chk(els['__#work-panel .alw-block:not(#jy-block)'].hidden === true,
  '同时把「人数×天数×标准」那块收起来（免得两块都摆着）');
box.state.kind = '差旅费报销单';
box.syncKindPills();
chk(els['jy-block'].hidden === true, '切回差旅费报销单 → 「按人」那块收起来');
chk(els['__#work-panel .alw-block:not(#jy-block)'].hidden === false,
  '「人数×天数×标准」那块放出来');

/* ==================================================================
 * 任务中心（P0-4）：谁在跑、跑到哪、失败能不能点「重试」
 * ================================================================== */
console.log('\n— 任务中心 —');
chk(box.fmtSecs(45) === '45 秒', `秒级显示（实际 ${box.fmtSecs(45)}）`);
chk(box.fmtSecs(125) === '2 分 5 秒', `分钟级显示（实际 ${box.fmtSecs(125)}）`);
chk(box.fmtSecs(7325) === '2 时 2 分', `小时级显示（实际 ${box.fmtSecs(7325)}）`);
chk(box.fmtSecs(null) === '0 秒', '没有耗时不显示 NaN');

// 渲染单行：状态色、按钮、转义
const rowRun = box.jobRowHtml({ id: 'a1', kind: 'start_import', kind_cn: '入库建账',
  state: '运行中', username: '小张', created_at: '2026-09-17 10:00:00',
  started_at: '2026-09-17 10:00:01', error: '', label: '', total: 10, done: 4,
  secs: 12, retryable: false });
const rowBad = box.jobRowHtml({ id: 'a2', kind: 'make_report', kind_cn: '生成报销单',
  state: '失败', username: 'admin', created_at: '2026-09-17 09:00:00',
  started_at: '2026-09-17 09:00:01', error: 'RuntimeError: 打印失败<script>x</script>',
  label: '', total: 0, done: 0, secs: 19, retryable: true });
const rowOk = box.jobRowHtml({ id: 'a3', kind: 'merge_invoices', kind_cn: '合并 PDF',
  state: '成功', username: '李四', created_at: '2026-09-17 08:00:00',
  started_at: '2026-09-17 08:00:00', error: '', label: '', total: 0, done: 0,
  secs: 5, retryable: false });
const rowQ = box.jobRowHtml({ id: 'a4', kind: 'export_summary', kind_cn: '导出统计',
  state: '排队', username: '张三', created_at: '2026-09-17 08:00:00',
  started_at: '', error: '', label: '', total: 0, done: 0, secs: 0, retryable: false });

chk(rowRun.includes('job-state run') && rowRun.includes('入库建账'),
  '运行中的任务带 run 状态色', rowRun);
chk(!rowRun.includes('data-act="retry"'), '没失败的任务不摆「重试」按钮');
chk(rowBad.includes('data-act="retry"') && rowBad.includes('data-jid="a2"'),
  '失败的任务有「重试」按钮，还带着任务号', rowBad);
chk(rowBad.includes('job-state err'), '失败带 err 状态色');
chk(rowBad.includes('&lt;script&gt;') && !rowBad.includes('<script>'),
  '失败原因里的尖括号被转义（任务名/报错是外部来的）', rowBad);
chk(!rowOk.includes('data-act='), '成功的任务没有操作按钮');
chk(rowQ.includes('data-act="cancel"'), '排队中的任务可以取消', rowQ);
chk(rowQ.includes('—'), '没开始的任务耗时显示「—」，不是 0 秒');

// 渲染整页：走一遍 renderJobs（api 打桩，不连后端）
const FAKE_JOBS = {
  rows: [
    { id: 'a1', kind: 'start_import', kind_cn: '入库建账', state: '运行中', username: '小张',
      created_at: '2026-09-17 10:00:00', started_at: '2026-09-17 10:00:01', ended_at: '',
      error: '', label: '入库建账 · /票据', total: 10, done: 4, secs: 12, retryable: false },
    { id: 'a2', kind: 'make_report', kind_cn: '生成报销单', state: '失败', username: 'admin',
      created_at: '2026-09-17 09:00:00', started_at: '2026-09-17 09:00:01',
      ended_at: '2026-09-17 09:00:20', error: 'RuntimeError: 打印失败',
      label: '', total: 0, done: 0, secs: 19, retryable: true },
  ],
  counts: { running: 1, queued: 0, done_today: 3, failed: 1 },
  busy: true, busy_label: '入库建账', busy_user: '小张', kinds: {},
};
let fakeJobs = FAKE_JOBS;
global.window.pywebview = { api: { list_jobs: async () => fakeJobs } };

(function runJobTests() {
  return box.renderJobs().then(() => {
    const stat = els['job-stat'].innerHTML;
    chk((stat.match(/job-chip/g) || []).length === 4, '汇总区 4 个数字块', stat);
    chk(stat.includes('>1<') && stat.includes('今日成功'), '把「运行中 1 / 今日成功 3」显示出来', stat);
    chk(stat.includes('job-chip err'), '有失败时汇总块标红');

    const rb = els['job-running-body'].innerHTML;
    chk(rb.includes('入库建账') && rb.includes('小张'),
      '正在执行区写着是谁在跑什么', rb);
    chk(rb.includes('width:40%'), `进度条按 done/total 走（4/10 → 40%，实际 ${
      (rb.match(/width:\d+%/) || ['无'])[0]}）`, rb);

    const list = els['job-list'].innerHTML;
    chk(list.includes('job-tbl') && list.includes('<thead>'), '历史列表渲染成表格', list.slice(0, 200));
    chk(list.includes('data-act="retry"'), '列表里能直接点重试');
    chk((els['job-running-hint'] || {}).textContent === '1 个在进行',
      '标题上写着有几个在跑');

    // 空列表分支
    fakeJobs = { rows: [], counts: { running: 0, queued: 0, done_today: 0, failed: 0 },
                 busy: false, kinds: {} };
    return box.renderJobs();
  }).then(() => {
    chk(els['job-list'].innerHTML.includes('还没有任务记录'), '一条任务都没有时给句人话',
      els['job-list'].innerHTML);
    chk(els['job-running-body'].innerHTML.includes('空闲'), '没任务在跑时说明空闲');

    // 后端报错时不能白屏
    fakeJobs = { error: 'db.py: database is locked' };
    return box.renderJobs();
  }).then(async () => {
    chk(els['job-list'].innerHTML.includes('读不到任务列表'), '后端出错时如实显示，不白屏',
      els['job-list'].innerHTML);

  /* =====================================================================
   * 设置页「账号与权限」为什么不见了（2026-09-17 用户报的 bug）
   * ---------------------------------------------------------------------
   * 原来 renderUsers() 在「没权限 / 读不到身份」时直接把整块卡片
   * display:none 藏掉 —— 用户只看到「这块没了」，既不知道为什么、
   * 也不知道要去哪拿权限，只能来问「怎么没有账号权限了」。
   * 现在必须把原因和办法写在页面上。
   * ===================================================================*/
  console.log('— 设置页账号与权限（不许静默消失）—');
  {
    const setMe = (o) => box.setMeTest(Object.assign(
      { user: '', role: '', role_cn: '', can: {}, auth: false, err: '' }, o));
    const card = els['card-users'] || (els['card-users'] = mkEl('card-users'));
    const list = els['user-list'] || (els['user-list'] = mkEl('user-list'));

    // ① 当前身份是业务人员：卡片不能消失，要写明「谁 / 为什么 / 怎么办」
    setMe({ user: '王涛', role: 'biz', role_cn: '业务人员', auth: true,
            can: { list_users: false } });
    await box.renderUsers();
    chk(card.style.display !== 'none',
      '非管理员：卡片不再整块藏掉（藏了就等于「功能凭空没了」）',
      'display=' + card.style.display);
    chk(list.innerHTML.includes('业务人员') && list.innerHTML.includes('没有账号管理权限'),
      '非管理员：写明当前身份和原因', list.innerHTML);
    chk(list.innerHTML.includes('访问口令') && list.innerHTML.includes('留空'),
      '非管理员：告诉用户怎么才能拿到权限（访问口令 / 用户名留空）', list.innerHTML);

    // ② 身份读不到（会话失效）：说「重新登录」，别诬赖成「你没权限」
    setMe({ err: 'HTTP 401' });
    await box.renderUsers();
    chk(card.style.display !== 'none', '读不到身份：卡片照留');
    chk(list.innerHTML.includes('读不到当前身份') && list.innerHTML.includes('重新登录'),
      '读不到身份：提示重新登录', list.innerHTML);
    chk(!list.innerHTML.includes('没有账号管理权限'),
      '会话过期不能说成「你没权限」（两回事，用户照着做会白折腾）', list.innerHTML);

    // ③ 管理员：给 api() 塞个假的 list_users，列表要真的渲染出来
    global.window.pywebview = { api: { list_users: () => Promise.resolve({
      roles: { admin: '管理员', biz: '业务人员' },
      rows: [{ username: 'admin', role: 'admin', role_cn: '管理员',
               display_name: '管理员', dept: '', disabled: false,
               last_login: '2026-09-17 09:58' }] }) } };
    setMe({ user: 'admin', role: 'admin', role_cn: '管理员', auth: true,
            can: { list_users: true } });
    await box.renderUsers();
    chk(card.style.display !== 'none', '管理员：卡片显示');
    chk(list.innerHTML.includes('admin') && list.innerHTML.includes('管理员'),
      '管理员：账号列表真的渲染出来了', list.innerHTML.slice(0, 200));
    chk(els['nu-role'].innerHTML.includes('value="admin"'), '角色下拉被填上选项');
    chk(!list.innerHTML.includes('没有账号管理权限'), '管理员不该看到「没权限」那套话');
    delete global.window.pywebview;

    // ④ 源码级护栏：别再犯回去
    chk(!/if \(!meInfo\.can \|\| !meInfo\.can\.list_users\) \{ card\.style\.display = 'none'/
      .test(src), 'renderUsers 里不再有「查不到就 display=none 直接藏」的老写法');
    const lse = src.slice(src.indexOf('async function loadSettingsExtras'),
                          src.indexOf('function initSettingsEvents'));
    chk(/for \(const fn of steps\)/.test(lse) && /catch/.test(lse),
      'loadSettingsExtras 逐步兜底：一步失败不再带停后面两步', lse.slice(0, 160));
    // ⚠️ bridge.js 以前只 return j.result，把后端的 error 整个丢掉 ——
    //    于是「权限不足：当前身份是…」永远到不了界面。这条盯着别再丢。
    const bridge = fs.readFileSync(path.join(GUI, 'bridge.js'), 'utf8');
    chk(/if \(j && j\.error\) throw/.test(bridge),
      'bridge.js 把后端的 error 抛出去（不再静默丢掉具体原因）');
  }

  /* =====================================================================
   * 点不了的按钮不许亮着（2026-09-17 用户报「删除选中报错」）
   * ---------------------------------------------------------------------
   * db.ROLE_ALLOW 里 delete_rows / clear_ledger 只有 admin 能调，可这两颗红按钮
   * 对**所有角色**都显示 —— 业务员一点必然被后端挡住，用户看到的就是「报错」。
   * 现在按 whoami().allowed 压灰 + 说明；真正的卡口仍然在后端 do_POST。
   * ===================================================================*/
  console.log('— 权限：点不了的按钮压灰 —');
  {
    const setMe = (o) => box.setMeTest(Object.assign(
      { user: '', role: '', role_cn: '', can: {}, auth: false, err: '', allowed: null }, o));
    const hint = els['perm-hint'] || (els['perm-hint'] = mkEl('perm-hint'));

    chk(box.PERM_BTNS['btn-del-sel'] === 'delete_rows',
      '「删除选中」登记为需要 delete_rows 权限');
    chk(box.PERM_BTNS['btn-clear-ledger'] === 'clear_ledger',
      '「清空台账」登记为需要 clear_ledger 权限');

    // ① 业务人员：两颗红按钮必须灰掉，并且说明为什么
    setMe({ user: '王涛', role: 'biz', role_cn: '业务人员', auth: true,
            allowed: ['get_ledger', 'make_report'] });
    box.applyPermissions();
    chk(els['btn-del-sel'].disabled === true,
      '业务人员：「删除选中」被压灰（不再是点了才报错）');
    chk(/没有这个权限/.test(els['btn-del-sel'].title || ''),
      '压灰的同时说清原因', els['btn-del-sel'].title);
    chk(els['btn-clear-ledger'].disabled === true, '业务人员：「清空台账」同样压灰');
    chk(els['btn-to-report'].disabled === false, '但该给的按钮不受影响（业务员能出单）');
    chk(hint.hidden === false && /业务人员/.test(hint.textContent || ''),
      '页面上有一句「当前登录是…灰掉的按钮不能用」', hint.textContent);

    // ② 管理员：全部恢复可点，提示收起
    setMe({ user: '口令用户', role: 'admin', role_cn: '管理员', auth: true,
            allowed: Object.keys(box.PERM_BTNS).map((k) => box.PERM_BTNS[k])
              .concat(['get_ledger']) });
    box.applyPermissions();
    chk(els['btn-del-sel'].disabled === false, '管理员：「删除选中」恢复可点');
    chk(els['btn-clear-ledger'].disabled === false, '管理员：「清空台账」恢复可点');
    chk(hint.hidden === true, '管理员：那句提示收起来');

    // ③ 本机模式 / 老后端没下发 allowed → 不许乱灰（否则界面被锁死）
    setMe({ auth: false });
    box.applyPermissions();
    chk(els['btn-del-sel'].disabled === false, '本机模式（没有登录这回事）：全部放行');
    setMe({ user: '王涛', role: 'biz', role_cn: '业务人员', auth: true, allowed: null });
    box.applyPermissions();
    chk(els['btn-del-sel'].disabled === false,
      '后端没下发 allowed（版本对不上）：宁可不灰，也别把界面锁死');

    // ④ 源码守则：canDo 认的是后端下发的 allowed，不许凭角色名字硬编码
    const permSrc = src.slice(src.indexOf('function canDo'),
      src.indexOf('function applyPermissions'));
    chk(/meInfo\.allowed/.test(permSrc),
      'canDo 看的是后端下发的 allowed', permSrc.slice(0, 120));
    chk(permSrc.indexOf("'admin'") < 0 && permSrc.indexOf("'biz'") < 0,
      'canDo 不靠角色名字硬编码（白名单以后改了这里也不会飘）');
  }
/* =====================================================================
 * 「按批次分开报」+ 合计口径 + 票面姓名保护（2026-09-17 第三轮）
 * ---------------------------------------------------------------------
 * 用户报：「首先，识别的 4361，加了两边打车的；而且报销的里边有高铁票，
 *          高铁票不是有名字吗，怎么识别不到，全都得是我自己输入的人名。」
 *   · 4361.00 = 太仓批 3795.80 + 烟台批 565.20（两个文件夹的票被硬加成一个数），
 *     且台账页把附件（那张 76.10 的行程单）也加了一遍 → 合计口径改成跟单据一致；
 *   · 高铁姓名其实**读到了**，是被「设置出行人」批量填名字盖掉的 → 加票面姓名保护。
 * ===================================================================*/
console.log('\n— 按批次分开报：批次条 —');
{
  const rows = [
    { '序号': 1, '状态': '未报销', '凭证类型': '发票', '价税合计': '942.54', '报销批次': '太仓批', '开票日期': '2026-07-23', '出行人': '孙振强' },
    { '序号': 2, '状态': '未报销', '凭证类型': '火车票', '价税合计': '456.00', '报销批次': '太仓批', '开票日期': '2026-08-05', '出行人': '孙振强' },
    { '序号': 3, '状态': '未报销', '凭证类型': '发票', '价税合计': '204.00', '报销批次': '', '开票日期': '2026-06-01', '出行人': '' },
  ];
  box.state.rows = rows;
  box.state.attachSeqs = new Set();
  box.state.sel = new Set();
  box.state.filter = { status: '', ctype: '', category: '', batch: '', person: '',
                       month: '', min: '', max: '', kw: '' };
  box.state.byBatch = [
    { batch: '太仓批', n: 2, n_attach: 0, sum: 1398.54, pending: 2, done: 0,
      first: '2026-07-23', last: '2026-08-05', persons: ['孙振强'] },
    { batch: '', n: 1, n_attach: 0, sum: 204, pending: 1, done: 0,
      first: '2026-06-01', last: '2026-06-01', persons: [] },
  ];
  box.renderBatchBar();
  const bar = els['batch-bar'];
  chk(bar.hidden === false, '两个批次 → 批次条显示出来');
  chk(/太仓批/.test(bar.innerHTML), '批次条里有「太仓批」');
  chk(/未填批次/.test(bar.innerHTML), '空批次名显示成「未填批次」（不然一片空白看不懂）');
  chk(/1,398\.54/.test(bar.innerHTML), '批次条里给的是这一批的计费金额 1,398.54');
  chk(/data-act="sel"/.test(bar.innerHTML), '每一批都有「全选这批」按钮（可以单独拿它出单）');
  chk(/按批次分开报/.test(bar.innerHTML), '有条标题说明这块是干什么的');
  chk(/数据来源/.test('数据来源') && /共 2 批/.test(bar.innerHTML), '总览写着「共 N 批」');

  // 只有一个批次、又没按批次筛 → 藏起来，不白占一块版面
  box.state.byBatch = [box.state.byBatch[0]];
  box.renderBatchBar();
  chk(els['batch-bar'].hidden === true, '只有一个批次时不显示这条（省版面）');
}

console.log('\n— 「全选这批」只勾这一批 —');
{
  box.state.rows = [
    { '序号': 1, '状态': '未报销', '凭证类型': '发票', '价税合计': '942.54', '报销批次': '太仓批', '出行人': '孙振强' },
    { '序号': 2, '状态': '未报销', '凭证类型': '火车票', '价税合计': '456.00', '报销批次': '太仓批', '出行人': '孙振强' },
    { '序号': 3, '状态': '未报销', '凭证类型': '发票', '价税合计': '204.00', '报销批次': '', '出行人': '' },
  ];
  box.state.sel = new Set();
  box.state.attachSeqs = new Set();
  box.onBatchBarClick({ target: { closest: () => ({ dataset: { act: 'sel', batch: '太仓批' } }) } });
  chk(box.state.sel.has('1') && box.state.sel.has('2'), '「全选这批」把太仓批两张都勾上');
  chk(!box.state.sel.has('3'), '没勾到另一批（两批分开报，勾错就混账了）');
}

console.log('\n— 合计口径：附件不重复计钱 —');
{
  box.state.rows = [
    { '序号': 11, '状态': '未报销', '凭证类型': '发票', '价税合计': '76.10', '报销批次': '烟台批', '开票日期': '2026-06-01', '出行人': '郑玉杰' },
    { '序号': 12, '状态': '未报销', '凭证类型': '打车行程单', '价税合计': '76.10', '报销批次': '烟台批', '开票日期': '2026-06-01', '出行人': '郑玉杰' },
  ];
  box.state.attachLinks = { '12': { main_seq: '11', main_amount: 76.1, main_desc: '打车票' } };
  box.state.attachSeqs = new Set(['12']);
  box.state.sel = new Set(['11', '12']);
  chk(box.isAttach('12') === true, 'isAttach 认得出附件行');
  chk(box.isAttach('11') === false, '主票不是附件');
  chk(Math.abs(box.selectedSum() - 76.10) < 0.001,
    '勾「发票 + 它的行程单」只算 76.10，不是 152.20（这就是「加了两遍打车」的根因）',
    box.selectedSum());
  box.renderLedgerBody();
  const tbody = els['ledger-body'].innerHTML;
  chk(/76\.10/.test(tbody), '台账表里两行都照常显示金额');
  chk(!/152\.20/.test(tbody), '但「待报销」组的小计没有把附件算两遍');
}

console.log('\n— 票面姓名保护：批量填名字不许盖掉票面 —');
{
  box.askFaceOverwrite(
    [{ seq: '25', name: '孙振强' }],
    [{ seq: '25', face: '邵自杰', cur: '孙振强', new: '孙振强',
       no: '26379116295002204819', detail: 'G2699 济南-太仓' }],
    '孙振强');
  const tt = els['modal-title'].textContent || '';
  const h2 = els['modal-text'].innerHTML || '';
  chk(/票面/.test(tt) && /保留/.test(tt), '弹层标题说明「票面姓名已保留」', tt);
  chk(/邵自杰/.test(h2), '弹层里列出票面真正写的姓名（邵自杰）');
  chk(/保持票面姓名/.test(els['modal-cancel'].textContent || ''),
    '「取消」写成「保持票面姓名」—— 默认就是不覆盖');
  chk(/覆盖这 1 张/.test(els['modal-ok'].textContent || ''),
    '「确定」要明说「覆盖这 1 张」（不能只写「确定」，点下去会改错人）',
    els['modal-ok'].textContent);
  chk(/票面是可信来源/.test(h2), '写清楚了为什么拦（票面是可信来源）');
}
  });                            // 收掉上面那个 await …then 回调
})().then(() => {
  console.log(fails ? `\n❌ ${fails} 项不通过` : '\n✅ 全部通过');
  process.exit(fails ? 1 : 0);
}).catch((e) => {
  console.log('\n❌ 任务中心测试自己抛错了：' + e);
  process.exit(1);
});
