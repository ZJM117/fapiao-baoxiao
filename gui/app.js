/* =====================================================================
 * 发票报销工具 · 界面交互
 * ---------------------------------------------------------------------
 * 后端在 app_web.py 里，通过 bridge.js 暴露成 window.pywebview.api.xxx()。
 * 六个页面：凭证入库 / 凭证台账 / 报销单 / 统计汇总 / 运行记录 / 设置。
 *
 * 后端的约定（改后端时请同步这里）：
 *   poll(since)          → {logs:[{i,t,m}], seq, progress:{on,cur,total,label},
 *                           status:{text,kind}, busy, result}
 *   get_paths()          → {source, ledger, output, app, ledger_exists}
 *   get_config()/set_config(patch) → 配置字典（含 categories 数组）
 *   pick_folder(initial) → 路径字符串（服务器模式回 {server:true,hint}，容器里弹不出选择框）
 *   check_source(path)   → {ok, files, kinds, msg}
 *   upload_info()        → {server, root, base, ok, msg, max_size, max_text}
 *                          服务器模式下才显示「上传本机文件夹」；文件走
 *                          POST /api/upload_raw（body=原始字节，X-Upload-Path=base64 相对路径）
 *   start_import(recursive, category, person, batch, source) → true（结果走 poll 的 result，只推一次）
 *                          source = 界面上那个路径框的值，后端以它为准
 *   get_ledger(filt)     → {rows, total, sum, pending_count, pending_sum, all_count,
 *                           categories, ctypes, batches, persons, months, aggregates}
 *                          filt: {status, ctype, category, batch, person, month, min, max, kw}
 *   mark_rows(seqs, status, batch, person)          → {updated}
 *   update_rows(seqs, fields)                       → {updated}
 *   refresh_rows(seqs)                              → {updated, miss, fail, miss_list, fail_list}
 *                          按最新规则重读原件、刷新票面字段（状态/类别/报销人不动）
 *   delete_rows(seqs)                               → {deleted}   （删台账记录，删前自动备份）
 *   clear_ledger()                                  → {cleared}   （清空台账，只留表头）
 *   make_report(meta, seqs, fmt)                    → {outs:[...],count,sum,upper}|{error}
 *                          fmt = 'pdf' | 'xlsx' | 'both'
 *   merge_invoices(seqs, mode='2up'|'plain')
 *        → {out,name,pages,files,missing,missing_list:[{name,no,why,row}],asked,skipped}|{error}
 *   export_summary(filt)                            → {out,name,rows}|{error}
 *   list_outputs()                                  → {items:[{path,name,kind,time,size,exists}]}
 *   delete_output(path)                             → {ok,name}   （送回收站，不是永久删除）
 *   open_folder(which) / open_file(path) / reveal(path)
 *
 * 台账列（ledger.COLUMNS，顺序即 Excel 里的列序；弹层/表格都按这些中文键取值）：
 *   序号 状态 凭证类型 发票号码 开票日期 票种 行程/明细 购方名称 销方名称 不含税金额
 *   税额 价税合计 项目/事由 项目编号 费用类别 报销人 报销批次 入库时间 文件路径 提示
 *   其中「凭证类型」取值：发票 / 火车票 / 飞机行程单 / 打车行程单 / 登机牌 / 图片凭证 / 其他凭证
 * ===================================================================== */
'use strict';

const THEMES = [
  { id: 'dark',      name: '深色经典', bg: '#0b0f19', card: '#111820', accent: '#6cb6ff' },
  { id: 'light',     name: '亮色清新', bg: '#f0f3f6', card: '#ffffff', accent: '#0969da' },
  { id: 'warm-gold', name: '暖阳金沙', bg: 'hsl(35,60%,96%)',  card: '#ffffff', accent: 'hsl(35,80%,55%)' },
  { id: 'tiffany',   name: '蒂芙尼蓝', bg: 'hsl(173,50%,96%)', card: '#ffffff', accent: 'hsl(173,70%,52%)' },
  { id: 'twilight',  name: '霞光紫',   bg: 'hsl(270,50%,97%)', card: '#ffffff', accent: 'hsl(260,80%,68%)' },
  { id: 'rosegold',  name: '金粉世家', bg: 'hsl(330,90%,98%)', card: '#ffffff', accent: 'hsl(330,80%,70%)' },
  { id: 'forest',    name: '森林绿',   bg: 'hsl(150,40%,96%)', card: '#ffffff', accent: 'hsl(150,60%,42%)' },
  { id: 'sakura',    name: '樱花粉',   bg: 'hsl(340,22%,96%)', card: '#ffffff', accent: 'hsl(340,65%,58%)' },
  { id: 'arctic',    name: '北极蓝',   bg: 'hsl(210,25%,95%)', card: '#ffffff', accent: 'hsl(210,80%,55%)' },
  { id: 'lime',      name: '青柠活力', bg: 'hsl(80,18%,95%)',  card: '#ffffff', accent: 'hsl(80,65%,40%)' },
  { id: 'peach',     name: '蜜桃橙',   bg: 'hsl(25,55%,96%)',  card: '#ffffff', accent: 'hsl(25,80%,58%)' },
  { id: 'coral',     name: '珊瑚红',   bg: 'hsl(355,45%,96%)', card: '#ffffff', accent: 'hsl(355,72%,55%)' },
];

const LS = {
  theme: 'invoice-rp.theme',
  sidebar: 'invoice-rp.sidebar',
  kind: 'invoice-rp.kind',
  fmt: 'invoice-rp.fmt',
  lastSrc: 'invoice-rp.lastSrc',
  alw: 'invoice-rp.alw',          // 差旅费补助那几行（自己填的，下次进来还在）
  alwOn: 'invoice-rp.alwOn',
  jyAlw: 'invoice-rp.jyAlw',      // 模板二版式：按人的补助金额 {姓名: 金额}
};

const $ = (id) => document.getElementById(id);
const api = () => (window.pywebview && window.pywebview.api) || null;

/* ---------------- 小工具 ---------------- */
function esc(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
/** 金额字符串/数字 → 数字（认不出给 0） */
function num(v) {
  const f = parseFloat(String(v === null || v === undefined ? '' : v).replace(/[,¥\s]/g, ''));
  return isNaN(f) ? 0 : f;
}
/** 金额显示：1314.5 → ¥1,314.50 */
function money(v) {
  return '¥' + num(v).toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
/** 只要千分位，不带符号 */
function money2(v) {
  return num(v).toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
const pad2 = (n) => String(n).padStart(2, '0');
function todayCn() {
  const d = new Date();
  return `${d.getFullYear()}年${pad2(d.getMonth() + 1)}月${pad2(d.getDate())}日`;
}
function baseName(p) {
  return String(p || '').replace(/[\\/]+$/, '').split(/[\\/]/).pop() || '';
}

/* ---------------- 全局状态 ---------------- */
const state = {
  busy: false,
  logs: [], unread: 0, logOpen: false,
  src: '',                       // 当前发票文件夹
  cfgSrc: '',                    // 配置里记着的那个（用来判断要不要回写配置）
  defaultSrc: '',                // 后端 DEFAULTS 里的默认路径
  recursive: true,
  cats: [],                      // 费用类别候选（设置页可改）
  rows: [],                      // 当前台账查询结果
  sel: new Set(),                // 选中的台账序号
  selRows: [],                   // 要出报销单的发票（从台账带过来）
  kind: localStorage.getItem(LS.kind) || '费用报销单',
  fmt: localStorage.getItem(LS.fmt) || 'pdf',      // 报销单输出格式 pdf / xlsx / both
  // 差旅费补助（用户要求：「出差有差旅费啊，要计算几个人几天，多少钱，可以我自己填」）
  alw: loadAlw(),                // [{name, people, days, rate, amount, manual}]
  alwOn: localStorage.getItem(LS.alwOn) === '1',
  alwOff: localStorage.getItem(LS.alwOn) === '0',  // 手动关掉过 → 换单据类型别自动开回来
  // 模板二版式：出差补助按人填（{姓名: 金额}）；票据合计由后端按出行人算（附件不重复计）
  jyAlw: loadJyAlw(),
  jyPeople: [],                  // [{name, bills}] 后端 report_persons 给的
  jyTotal: 0,
  attachLinks: {},               // {附件序号: {main_seq, main_amount}} —— 台账里标「↳ 附件」
  outputs: [],                   // 生成过的文件（来自后端 输出/生成记录.json）
  filter: { status: '', ctype: '', category: '', batch: '', person: '',
            month: '', min: '', max: '', kw: '' },
  aggregates: null,
  aggScope: '全部',
  cfg: null,
  serverMode: false,             // 后端是不是服务器模式（NAS）—— 决定「选择文件夹」弹本机还是填路径
  uploadRoot: '',                // 服务器模式：上传落点的根目录
};

/* =====================================================================
 * 主题
 * ===================================================================== */
function applyTheme(id) {
  if (!THEMES.some((t) => t.id === id)) id = 'dark';
  document.documentElement.setAttribute('data-theme', id);
  localStorage.setItem(LS.theme, id);
  document.querySelectorAll('.theme-card').forEach((el) =>
    el.classList.toggle('active', el.dataset.theme === id));
  const btn = $('btn-theme');
  if (btn) {
    const cur = THEMES.find((t) => t.id === id);
    btn.title = `界面主题（当前：${cur ? cur.name : id}）`;
  }
}

function renderThemes() {
  $('theme-grid').innerHTML = THEMES.map((t) => `
    <button class="theme-card" data-theme="${t.id}">
      <div class="theme-preview" style="background:linear-gradient(135deg, ${t.bg} 0%, ${t.bg} 55%, ${t.accent} 100%)">
        <span class="theme-swatch" style="background:${t.bg}"></span>
        <span class="theme-swatch" style="background:${t.card}"></span>
        <span class="theme-swatch" style="background:${t.accent}"></span>
      </div>
      <div class="theme-meta"><span class="theme-name">${t.name}</span><span class="theme-check">&#10003;</span></div>
    </button>`).join('');
  $('theme-grid').querySelectorAll('.theme-card').forEach((el) =>
    el.addEventListener('click', () => applyTheme(el.dataset.theme)));
}

function initThemePop() {
  const pop = $('theme-pop'), btn = $('btn-theme');
  const close = () => pop.classList.remove('show');
  btn.addEventListener('click', (e) => { e.stopPropagation(); pop.classList.toggle('show'); });
  document.addEventListener('click', (e) => {
    if (!pop.classList.contains('show')) return;
    if (pop.contains(e.target) || btn.contains(e.target)) return;
    close();
  });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') { close(); hideModal(); } });
}

/* =====================================================================
 * 侧边栏 / 页面切换
 * ===================================================================== */
function switchPage(pageId) {
  document.querySelectorAll('.nav-item').forEach((b) =>
    b.classList.toggle('active', b.dataset.page === pageId));
  document.querySelectorAll('.page').forEach((p) =>
    p.classList.toggle('active', p.id === pageId));
  try {
    if (location.hash !== '#' + pageId) history.replaceState(null, '', '#' + pageId);
  } catch (e) { /* file:// 下可能不允许改 URL，忽略 */ }
  onPageEnter(pageId);
}

function onPageEnter(pageId) {
  if (pageId === 'page-ledger') loadLedger();
  else if (pageId === 'page-stats') loadStats();
  else if (pageId === 'page-set') { loadConfigIntoForm(); loadPhoneMap(); loadSettingsExtras(); }
  else if (pageId === 'page-report') loadOutputs();
}

function setSidebarCollapsed(collapsed) {
  $('sidebar').classList.toggle('collapsed', collapsed);
  document.documentElement.dataset.sb = collapsed ? 'collapsed' : 'open';
  localStorage.setItem(LS.sidebar, collapsed ? '1' : '0');
}

function initNav() {
  document.querySelectorAll('.nav-item').forEach((btn) => {
    btn.addEventListener('click', () => switchPage(btn.dataset.page));
  });
  $('btn-collapse').addEventListener('click', () =>
    setSidebarCollapsed(!$('sidebar').classList.contains('collapsed')));
  setSidebarCollapsed(localStorage.getItem(LS.sidebar) === '1');
  // 报销单页里点了「生成报销单 →」要能跳过去
  $('btn-to-report').addEventListener('click', goToReport);
}

/* =====================================================================
 * 状态 / 进度
 * ===================================================================== */
function setStatus(text, kind) {
  $('status').className = 'status-badge ' + (kind || 'ready');
  $('status-text').textContent = text;
}

function setProgress(on, cur, total, label) {
  $('progress-wrap').style.display = on ? '' : 'none';
  if (!on) return;
  $('progress-label').textContent = label || '处理中…';
  $('progress-num').textContent = total ? `${cur} / ${total}` : '';
  $('progress-bar').style.width = total ? `${Math.round((cur / total) * 100)}%` : '0%';
}

function setImportBusy(busy) {
  state.busy = busy;
  const btn = $('btn-import');
  btn.disabled = busy;
  btn.textContent = busy ? '入库中…' : '开始入库';
  $('btn-pick').disabled = busy;
  $('src-path').disabled = busy;
  if ($('btn-upload')) $('btn-upload').disabled = busy;
}

/* =====================================================================
 * 日志
 * ===================================================================== */
function logClass(line) {
  if (line.includes('❌') || line.includes('失败')) return 'err';
  if (line.includes('⚠') || line.includes('跳过') || line.includes('需人工')) return 'warn';
  if (line.includes('✅') || line.includes('完成')) return 'ok';
  if (line.includes('开始') || line.includes('→')) return 'info';
  return '';
}

function pushLogs(items) {
  (items || []).forEach((it) => {
    const line = typeof it === 'string' ? it : `[${it.t}] ${it.m}`;
    state.logs.push(line);
    const cls = logClass(line);
    [$('log-body'), $('log-full')].forEach((box) => {
      if (!box) return;
      const atBottom = box.scrollHeight - box.scrollTop - box.clientHeight < 40;
      const el = document.createElement('div');
      el.className = 'log-line ' + cls;
      el.textContent = line;
      box.appendChild(el);
      if (atBottom) box.scrollTop = box.scrollHeight;
    });
  });
  $('log-count').textContent = state.logs.length + ' 条';
  if (!state.logOpen && (items || []).length) {
    state.unread += items.length;
    $('log-toggle-hint').textContent = `点击展开 · ${state.unread} 条新日志`;
  }
}

function initLogPanel() {
  $('log-head').addEventListener('click', () => {
    state.logOpen = !state.logOpen;
    $('log-panel').classList.toggle('open', state.logOpen);
    $('log-toggle-hint').textContent = state.logOpen ? '点击收起' : '点击展开';
    document.documentElement.dataset.log = state.logOpen ? 'open' : 'closed';
    if (state.logOpen) {
      state.unread = 0;
      const b = $('log-body');
      b.scrollTop = b.scrollHeight;
    }
  });
  document.documentElement.dataset.log = 'closed';
  $('btn-clear-log').addEventListener('click', () => {
    $('log-full').innerHTML = '';
    $('log-body').innerHTML = '';
    state.logs = []; state.unread = 0;
    $('log-count').textContent = '0 条';
    pushLogs(['已清空显示（日志文件 launch.log 未删除）']);
  });
  $('btn-open-app').addEventListener('click', () => call('open_folder', 'app'));
}

/* =====================================================================
 * 提示 / 弹窗
 * ===================================================================== */
function toast(text, kind) {
  const el = document.createElement('div');
  el.className = 'toast ' + (kind || '');
  el.textContent = text;
  $('toasts').appendChild(el);
  setTimeout(() => { el.style.opacity = '0'; el.style.transition = 'opacity .3s'; }, 3600);
  setTimeout(() => el.remove(), 4000);
}

let modalOk = null;
function showModal(o) {
  $('modal-title').textContent = o.title || '确认';
  $('modal-text').innerHTML = o.html || '';
  $('modal-note').innerHTML = o.note || '';
  $('modal-ok').textContent = o.okText || '确定';
  $('modal-ok').style.display = o.noOk ? 'none' : '';
  $('modal-cancel').textContent = o.cancelText || '取消';
  // 明细这类内容多的弹层要更宽一点（.modal-lg / .modal-xl 在 extra.css 里）
  const m = document.querySelector('#overlay .modal');
  if (m) {
    m.classList.toggle('modal-lg', !!o.wide);
    m.classList.toggle('modal-xl', !!o.xwide);
  }
  modalOk = o.onOk || null;
  $('overlay').classList.add('show');
  const box = document.querySelector('#overlay .modal');
  if (box) box.scrollTop = 0;
}
function hideModal() {
  $('overlay').classList.remove('show');
  modalOk = null;
}

/** 调后端接口的薄包装：出错统一转成 toast，不让界面炸掉 */
async function call(name, ...args) {
  const a = api();
  if (!a) { toast('界面还没初始化好，请稍候…', 'warn'); return null; }
  try {
    return await a[name](...args);
  } catch (e) {
    toast(`${name} 调用失败：${e}`, 'err');
    return null;
  }
}

/* =====================================================================
 * 一、凭证入库
 * ===================================================================== */
function initImportPage() {
  document.querySelectorAll('#scope-pills .pill').forEach((p) => {
    p.addEventListener('click', () => {
      state.recursive = p.dataset.scope === 'deep';
      syncScopePills();
      localStorage.setItem('invoice-rp.recursive', state.recursive ? '1' : '0');
      call('set_config', { recursive: state.recursive });
      checkSource();
    });
  });

  $('src-path').addEventListener('input', () => {
    state.src = $('src-path').value.trim();
    localStorage.setItem(LS.lastSrc, state.src);
    hideSrcOrigin();          // 手动改了路径，那条「本机文件夹 → NAS」的来源就不再成立了
    checkSource();
  });

  $('btn-pick').addEventListener('click', async () => {
    // 服务器模式（NAS）：直接弹「本机文件夹」选择框，选完立刻上传 ——
    // 不再让人「先去上面传一遍、再回来把路径改成 NAS 的」（见 initUpload）。
    if (state.serverMode) { $('up-input').click(); return; }
    const note = $('src-note');
    note.className = 'src-note';
    note.textContent = '正在等待你选择文件夹…';
    const path = await call('pick_folder', state.src || state.defaultSrc);
    // 服务器模式（Docker/NAS）没有目录选择框，后端会回 {server:true, hint:...}
    if (path && typeof path === 'object' && path.server) {
      note.className = 'src-note warn';
      note.textContent = path.hint || '服务器模式下请直接填写路径';
      return;
    }
    if (!path) { checkSource(); return; }
    state.src = path;
    $('src-path').value = path;
    localStorage.setItem(LS.lastSrc, path);
    call('set_config', { source: path });
    checkSource();
  });

  $('btn-src-reset').addEventListener('click', () => {
    state.src = state.defaultSrc;
    $('src-path').value = state.src;
    localStorage.setItem(LS.lastSrc, state.src);
    hideSrcOrigin();
    call('set_config', { source: state.src });
    checkSource();
  });

  $('btn-import').addEventListener('click', doImport);
  $('btn-open-ledger').addEventListener('click', openLedgerFile);
  $('btn-open-out').addEventListener('click', () => call('open_folder', 'output'));
  initUpload();
}

/* ---------------------------------------------------------------------
 * 上传本机文件夹（只在服务器模式 / NAS 下出现）
 * ---------------------------------------------------------------------
 * 程序跑在 NAS 的容器里，看不见用户电脑的磁盘 —— 「处理我自己电脑上那个文件夹」
 * 就只能先把文件传上去。传的时候原样保留子目录（webkitRelativePath），发票和它的
 * 行程单还是待在同一个文件夹里，配对规则（attachment.py）照常成立。
 * ------------------------------------------------------------------- */
function b64utf8(s) {
  const bytes = new TextEncoder().encode(s);
  let bin = '';
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return btoa(bin);
}

function relOf(f) {
  return f.relPath || f.webkitRelativePath || f.name;
}

function setUploadProgress(on, cur, total, label) {
  $('up-progress').style.display = on ? '' : 'none';
  if (!on) return;
  $('up-label').textContent = label || '上传中…';
  $('up-num').textContent = total ? `${cur} / ${total}` : '';
  $('up-bar').style.width = total ? `${Math.round((cur / total) * 100)}%` : '0%';
}

async function initUpload() {
  const row = $('up-row');
  if (!row) return;
  const info = await call('upload_info');
  // 本机模式不摆这套：能直接指路径，再给一套「传上去」的做法（还多一份副本）
  // 只会让人犯嘀咕。界面看见的必须等于程序在做的，所以不该出现的一律藏掉。
  if (!info || !info.server) { row.hidden = true; return; }
  row.hidden = false;
  state.serverMode = true;
  state.uploadRoot = info.root || '';

  // 主入口就是「选择文件夹…」那颗：服务器模式下它弹的是**你这台电脑**的文件夹，
  // 选完自动上传，不用再自己把路径改成 NAS 的。
  const pick = $('btn-pick');
  if (pick) {
    pick.textContent = '选择本机文件夹…';
    pick.title = '挑一个你这台电脑上的文件夹，选完自动传到 NAS（子目录原样保留）';
  }
  // 上面那行「上传本机文件夹…」按钮就多余了，撤掉；拖拽入口和提示留着。
  const upBtn = $('btn-upload');
  if (upBtn) upBtn.style.display = 'none';

  const hint = $('up-hint');
  if (!info.ok) {
    hint.textContent = `⚠️ 传不进去：${info.msg || '目录不可写'}`;
  } else {
    hint.innerHTML = `也可以直接把文件夹拖到这里　传到 NAS 上的 <span class="mono">${esc(info.root)}</span>`
      + `　单文件上限 ${esc(info.max_text || '')}`;
  }

  $('up-input').addEventListener('change', () => {
    const files = Array.from($('up-input').files || []);
    $('up-input').value = '';        // 清掉，同一个文件夹还能再选一次
    if (files.length) startUpload(files, baseOf(files[0]));
  });
  initDropUpload();
}

/* 来源条：本机选的是哪个文件夹 → 它落到 NAS 哪儿（一行看完，不用自己比对路径）
 * ⚠️ 浏览器不会把本机绝对路径交给网页（webkitdirectory 只给「文件夹名/子路径」），
 *    所以这里只能显示文件夹名 —— 但"我选的是哪个"一眼就认得出来。 */
function showSrcOrigin(localFolder, nasPath) {
  const el = $('src-origin');
  if (!el) return;
  if (!localFolder || !nasPath) { el.hidden = true; return; }
  el.hidden = false;
  el.innerHTML = `<span class="so-k">本机文件夹</span>`
    + `<b class="so-v">${esc(localFolder)}</b>`
    + `<span class="so-arrow">→</span>`
    + `<span class="so-k">NAS</span><code class="so-path">${esc(nasPath)}</code>`
    + `<span class="so-tip">已传好，直接点「开始入库」</span>`;
}

function hideSrcOrigin() {
  const el = $('src-origin');
  if (el) el.hidden = true;
}

/** 挑一个能代表这批文件的文件夹名（用来提示「正在传哪个文件夹」） */
function baseOf(f) {
  const rel = relOf(f).replace(/\\/g, '/');
  const i = rel.indexOf('/');
  return i > 0 ? rel.slice(0, i) : '';
}

/** 把文件夹直接拖进页面也能传（资源管理器里拖过来最顺手） */
function initDropUpload() {
  const page = $('page-in');
  if (!page) return;
  let depth = 0;
  const hasFiles = (e) => !!(e.dataTransfer
    && Array.from(e.dataTransfer.types || []).indexOf('Files') >= 0);

  page.addEventListener('dragenter', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault(); depth++; page.classList.add('drop-on');
  });
  page.addEventListener('dragover', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault(); e.dataTransfer.dropEffect = 'copy';
  });
  page.addEventListener('dragleave', () => {
    depth--; if (depth <= 0) { depth = 0; page.classList.remove('drop-on'); }
  });
  page.addEventListener('drop', async (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault(); depth = 0; page.classList.remove('drop-on');
    if (state.busy) { toast('正在忙，等这次跑完再拖进来', 'warn'); return; }
    // webkitGetAsEntry 必须在这个同步阶段取：await 一过，items 就失效了
    const entries = [];
    const plain = [];
    Array.from(e.dataTransfer.items || []).forEach((it) => {
      const en = it.webkitGetAsEntry ? it.webkitGetAsEntry() : null;
      if (en) { entries.push(en); return; }
      const f = it.getAsFile();
      if (f) plain.push(f);
    });
    const files = [];
    for (const en of entries) await walkEntry(en, '', files);
    files.push(...plain);
    if (!files.length) { toast('没读到文件 —— 拖的是文件夹吗？', 'warn'); return; }
    startUpload(files, (entries[0] && entries[0].name) || '');
  });
}

/** 把一个拖进来的条目递归读成文件列表，顺带把相对路径补到 f.relPath 上 */
function walkEntry(entry, prefix, out) {
  return new Promise((resolve) => {
    if (!entry) { resolve(); return; }
    if (entry.isFile) {
      entry.file((f) => {
        try { Object.defineProperty(f, 'relPath', { value: prefix + f.name }); }
        catch (err) { f.relPath = prefix + f.name; }
        out.push(f); resolve();
      }, () => resolve());
      return;
    }
    if (!entry.isDirectory) { resolve(); return; }
    const reader = entry.createReader();
    const all = [];
    const readMore = () => reader.readEntries((ents) => {
      if (ents.length) { all.push(...ents); readMore(); return; }
      (async () => {
        for (const en of all) await walkEntry(en, prefix + entry.name + '/', out);
        resolve();
      })();
    }, () => resolve());
    readMore();
  });
}

/** 一个文件一条请求（串行）。串行是有意的：NAS 磁盘 + 千兆网，并发反而更慢更容易断 */
function uploadOne(f) {
  return fetch('/api/upload_raw', {
    method: 'POST',
    cache: 'no-store',
    headers: { 'X-Upload-Path': b64utf8(relOf(f)), 'X-File-Size': String(f.size || 0) },
    body: f,
  }).then((r) => r.json()).then((j) => {
    if (j && j.error) throw new Error(j.error);
    return (j && j.result) || null;
  });
}

async function startUpload(files, label) {
  if (state.busy) return;
  const total = files.length;
  const totalBytes = files.reduce((s, f) => s + (f.size || 0), 0);
  setImportBusy(true);
  setUploadProgress(true, 0, total,
    `准备上传 ${label || ''}（${total} 个文件 / ${fmtSize(totalBytes)}）`);
  setStatus('上传中…', 'running');
  pushLogs([`开始上传本机文件夹：${label || '(直接拖入的文件)'} —— ${total} 个文件 ${fmtSize(totalBytes)}`]);

  let added = 0, skipped = 0, failed = 0, bytes = 0;
  const errs = [];
  let root = state.uploadRoot || '';
  const t0 = Date.now();
  for (let i = 0; i < total; i++) {
    const f = files[i];
    setUploadProgress(true, i, total, `正在上传（${i + 1}/${total}）${f.name}`);
    try {
      const r = await uploadOne(f);
      if (r && r.root) root = r.root;
      if (r && r.mode === 'skip') {
        skipped++;
      } else {
        added++;
        bytes += (r && r.size) || f.size || 0;
      }
    } catch (err) {
      failed++;
      if (errs.length < 6) errs.push(`${relOf(f)}：${err.message || err}`);
    }
    setUploadProgress(true, i + 1, total, `已上传 ${i + 1}/${total}`);
  }

  setUploadProgress(false);
  setImportBusy(false);
  setStatus('就绪', 'ready');

  const secs = Math.max(0.1, (Date.now() - t0) / 1000);
  const bits = [`新增 ${added} 个`];
  if (skipped) bits.push(`跳过 ${skipped} 个（NAS 上已经有）`);
  if (failed) bits.push(`失败 ${failed} 个`);
  const msg = `上传完成：${bits.join('，')}，共 ${fmtSize(bytes)}，用时 ${secs.toFixed(1)} 秒`;
  pushLogs([`${failed ? '⚠️' : '✅'} ${msg}`]);
  errs.forEach((e) => pushLogs([`⚠️ 上传失败：${e}`]));
  toast(msg, failed ? 'warn' : 'ok');

  if (added && root) {
    // 把路径框切到上传目录：范围一目了然（就是刚传上来的这些），
    // 并且必须带上「含子文件夹」—— 每个文件夹一层，只看当前层会漏掉里面的票。
    state.src = root;
    $('src-path').value = root;
    localStorage.setItem(LS.lastSrc, root);
    state.cfgSrc = root;
    state.recursive = true;
    syncScopePills();
    localStorage.setItem('invoice-rp.recursive', '1');
    call('set_config', { source: root, recursive: true });
    checkSource();
    showSrcOrigin(label, root);
    $('in-hint').textContent = `已上传 ${added} 个文件到 NAS，可以点「开始入库」建账了`;
  }
}

function syncScopePills() {
  document.querySelectorAll('#scope-pills .pill').forEach((p) =>
    p.classList.toggle('active', (p.dataset.scope === 'deep') === state.recursive));
}

let checkTimer = null;
function checkSource() {
  const note = $('src-note');
  const path = state.src;
  if (!path) {
    note.className = 'src-note warn';
    note.textContent = '请先指定放发票的文件夹';
    return;
  }
  clearTimeout(checkTimer);
  checkTimer = setTimeout(async () => {
    const a = api();
    if (!a || !a.check_source) return;
    let r;
    try { r = await a.check_source(path); } catch (e) { return; }
    if (!r || path !== state.src) return;            // 输入框已经改了，丢弃这次结果
    if (!r.ok) {
      note.className = 'src-note err';
      note.textContent = '❌ ' + (r.msg || '路径不可用');
      return;
    }
    const kinds = r.kinds || {};
    const detail = Object.keys(kinds).sort()
      .map((k) => `${k}×${kinds[k]}`).join('　');
    note.className = r.files ? 'src-note ok' : 'src-note warn';
    note.textContent = (r.files ? '✅ ' + r.msg : '⚠️ ' + r.msg)
      + (detail ? `（${detail}）` : '')
      + `　—　${state.recursive ? '含子文件夹' : '只看当前层'}`;
    // 路径确认能用了就同步进配置：否则「打开凭证文件夹」等入口还指着上一个目录
    if (r.ok && path !== state.cfgSrc) {
      state.cfgSrc = path;
      call('set_config', { source: path });
    }
  }, 380);
}

function doImport() {
  if (state.busy) return;
  if (!state.src) { toast('请先指定放发票的文件夹', 'warn'); return; }
  const person = $('in-person').value.trim();
  const batch = $('in-batch').value.trim();
  showModal({
    title: '开始入库建账',
    html: `将遍历 <b>${esc(state.src)}</b>${state.recursive ? '（含子文件夹）' : '（只当前层）'}，
           解析里面的 PDF / OFD / XML / zip，把发票要素登记到台账。<br>
           同一张票（同发票号码）不会重复登记；台账里已有且<b>已报销</b>的会标红提示。`,
    note: (person ? `报销人：${esc(person)}　` : '')
      + (batch ? `批次：${esc(batch)}　` : '')
      + '解析过程后台跑，不弹窗、不抢焦点。',
    okText: '开始入库',
    onOk: async () => {
      hideModal();
      setImportBusy(true);
      setStatus('入库中…', 'running');
      setProgress(true, 0, 0, '正在读取文件…');
      $('in-hint').textContent = '入库进行中…';
      pushLogs([`开始入库：${state.src}`]);
      // ⚠️ 第 5 个参数必须是界面上这个路径。不传的话后端会去读配置里的旧目录，
      //    于是"扫的目录 ≠ 你看到的目录"——曾经就因此把老目录 80 多条全灌回台账。
      const ok = await call('start_import', state.recursive, '', person, batch, state.src);
      if (!ok) {
        setImportBusy(false);
        setProgress(false);
        toast('入库没能启动（可能已经有一个任务在跑）', 'warn');
      }
    },
  });
}

/** 入库结果（后端带 result_id，靠它去重；刷新页面时会把上次结果显示出来但不再弹提示） */
function onImportResult(res, silent) {
  setImportBusy(false);
  setProgress(false);
  $('manual-wrap').style.display = 'none';
  $('dup-wrap').style.display = 'none';

  if (!res) return;
  if (res.error) {
    setStatus('入库失败', 'err');
    $('in-hint').textContent = '入库失败';
    $('in-stats').innerHTML = '';
    toast('入库失败：' + res.error, 'err');
    return;
  }
  if (res.found === 0) {
    setStatus('没找到凭证', 'ready');
    $('in-hint').textContent = '这个文件夹里没有 PDF / OFD / XML / zip / 图片';
    $('in-stats').innerHTML = '';
    if (!silent) toast('这个文件夹里没找到可识别的凭证', 'warn');
    return;
  }

  const boxes = [
    ['新增入库', res.added, 'accent', '条'],
    ['台账已有', res.in_ledger, 'mute', '条（未重复登记）'],
    ['重复报销风险', res.dup, 'err', '条'],
    ['需人工处理', res.manual, 'warn', '个文件'],
    ['台账合计', res.total, '', '条'],
    ['待报销金额', money(res.pending_amount), 'ok', ''],
  ];
  $('in-stats').innerHTML = boxes.map(([lab, val, cls, unit]) =>
    `<div class="stat-box ${cls}"><div class="sb-num">${esc(val)}</div>
     <div class="sb-lab">${esc(lab)}${unit ? ' · ' + esc(unit) : ''}</div></div>`).join('');

  // 本次新增都识别成了什么（发票 / 火车票 / 行程单…）
  const byType = res.by_type || {};
  const typeStr = Object.entries(byType).map(([t, c]) => `${t} ${c}`).join(' / ');
  $('in-hint').innerHTML =
    (res.source ? `扫描目录：<span class="mono">${esc(res.source)}</span><br>` : '')
    + `共扫描 ${esc(res.files)} 个文件　·　新增 ${esc(res.added)} 条　·　台账现有 ${esc(res.total)} 行`
    + (res.patched ? `　·　补全老记录 ${esc(res.patched)} 条` : '')
    + (typeStr ? `<br><span style="color:var(--text-tertiary)">本次新增构成：${esc(typeStr)}</span>` : '');

  // 需人工
  const manual = res.manual_list || [];
  $('manual-wrap').style.display = manual.length ? '' : 'none';
  $('manual-body').innerHTML = manual.map((m, i) =>
    `<tr><td class="td-idx">${pad2(i + 1)}</td>
     <td style="font-family:var(--font);color:var(--text-primary)">${esc(m.name)}</td>
     <td class="td-time">${esc(m.why)}</td></tr>`).join('');

  // 重复报销风险
  const dups = res.dup_list || [];
  $('dup-wrap').style.display = dups.length ? '' : 'none';
  $('dup-body').innerHTML = dups.map((d) =>
    `<tr><td class="td-no">${esc(d.no)}</td><td class="td-time">${esc(d.tip)}</td></tr>`).join('');

  const parts = [`新增 ${res.added} 条`];
  if (res.in_ledger) parts.push(`台账已有 ${res.in_ledger} 条`);
  if (res.dup) parts.push(`重复风险 ${res.dup} 条`);
  if (res.manual) parts.push(`需人工 ${res.manual} 个`);
  if (!silent) toast('入库完成：' + parts.join('，'), res.dup || res.manual ? 'warn' : 'ok');

  // 台账页若正在显示，顺手刷一下
  if ($('page-ledger').classList.contains('active')) loadLedger();
}

/* =====================================================================
 * 二、凭证台账
 * ===================================================================== */
/** 台账筛选条件。⚠️ 每个筛选控件都必须在这里出现一次 ——
 *  之前漏了「凭证类型」（下拉选了没反应），用户反馈就是「查询不能用」。 */
function currentFilter() {
  const val = (id) => (($(id) && $(id).value) || '').trim();
  return {
    status: val('f-status'),
    ctype: val('f-ctype'),
    category: val('f-category'),
    batch: val('f-batch'),
    person: val('f-person'),
    month: val('f-month'),
    min: val('f-min'),
    max: val('f-max'),
    kw: val('f-kw'),
  };
}

/** 把当前筛选条件翻译成人话，显示在筛选条下面，让用户知道「查询到底生效了没」 */
function filterSummary(f) {
  const labels = {
    status: '状态', ctype: '凭证类型', category: '费用类别', batch: '报销批次',
    person: '报销人', month: '月份', kw: '关键词',
  };
  const parts = [];
  Object.keys(labels).forEach((k) => { if (f[k]) parts.push(`${labels[k]}：${f[k]}`); });
  if (f.min || f.max) parts.push(`金额：${f.min || '不限'} ~ ${f.max || '不限'}`);
  return parts.length ? '当前条件 — ' + parts.join('　·　') : '当前条件 — 全部（未加任何筛选）';
}

function fillSelect(id, values, current, allLabel) {
  const sel = $(id);
  const prev = current !== undefined ? current : sel.value;
  sel.innerHTML = `<option value="">${allLabel}</option>`
    + (values || []).map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
  if (prev && (values || []).includes(prev)) sel.value = prev;
  else sel.value = '';
  return sel.value;
}

function initLedgerPage() {
  $('btn-refresh').addEventListener('click', () => loadLedger());
  $('btn-reset-filter').addEventListener('click', () => {
    ['f-status', 'f-ctype', 'f-category', 'f-batch', 'f-person', 'f-month',
     'f-min', 'f-max', 'f-kw'].forEach((id) => { if ($(id)) $(id).value = ''; });
    state.sel.clear();
    loadLedger();
  });
  // 下拉改动即查；文本类（关键词 / 金额）回车或点「查询」再查
  ['f-status', 'f-ctype', 'f-category', 'f-batch', 'f-person', 'f-month'].forEach((id) =>
    $(id).addEventListener('change', () => loadLedger()));
  ['f-kw', 'f-min', 'f-max'].forEach((id) =>
    $(id).addEventListener('keydown', (e) => { if (e.key === 'Enter') loadLedger(); }));

  $('btn-export').addEventListener('click', () => doExport(currentFilter()));
  $('btn-open-ledger2').addEventListener('click', openLedgerFile);

  $('ledger-body').addEventListener('change', (e) => {
    const cb = e.target.closest('input[type="checkbox"]');
    if (!cb) return;
    if (cb.checked) state.sel.add(cb.dataset.seq);
    else state.sel.delete(cb.dataset.seq);
    syncSelectionUI();
  });
  $('ledger-body').addEventListener('click', (e) => {
    if (e.target.closest('input[type="checkbox"]')) return;
    const tr = e.target.closest('tr[data-seq]');
    if (!tr) return;
    const row = state.rows.find((r) => String(r['序号']) === tr.dataset.seq);
    if (row) showInvoiceDetail(row);
  });

  $('ck-all').addEventListener('change', () => {
    state.rows.forEach((r) => {
      const s = String(r['序号']);
      if ($('ck-all').checked) state.sel.add(s); else state.sel.delete(s);
    });
    renderLedgerBody();
    syncSelectionUI();
  });

  $('btn-sel-all').addEventListener('click', () => {
    const all = state.rows.length && state.rows.every((r) => state.sel.has(String(r['序号'])));
    state.rows.forEach((r) => {
      const s = String(r['序号']);
      if (all) state.sel.delete(s); else state.sel.add(s);
    });
    renderLedgerBody();
    syncSelectionUI();
  });

  $('btn-mark-done').addEventListener('click', () => {
    if (state.sel.size) markSelected('已报销');
  });
  $('btn-mark-undone').addEventListener('click', () => {
    if (state.sel.size) markSelected('未报销');
  });
  $('btn-apply-cat').addEventListener('click', applyCategory);
  $('btn-set-tv').addEventListener('click', setTraveler);
  $('btn-review-tv').addEventListener('click', openTravelerReview);
  $('btn-reparse').addEventListener('click', refreshSelected);
  $('btn-del-sel').addEventListener('click', deleteSelected);
  $('btn-clear-ledger').addEventListener('click', clearLedger);
}

/**
 * 「重新识别」：按最新的识别规则把选中凭证的原件重读一遍。
 *
 * 解析规则是会改的（识别错了、票面版式变了），可台账里已经躺着的旧值就是错的。
 * 直接删了重新入库当然也行，但那样连「已报销状态、报销批次」一起丢。
 * 这条只刷新**票面上读来的**字段：状态 / 费用类别 / 报销人 / 报销批次 / 入库时间都不动。
 */
function refreshSelected() {
  const seqs = selectedSeqs();
  if (!seqs.length) { toast('先在表格里勾选要重新识别的记录', 'warn'); return; }
  showModal({
    title: `重新识别选中的 ${seqs.length} 条？`,
    html: `会按当前的识别规则<b>重新读一遍原件</b>，刷新发票号码、开票日期、购销方名称、
           金额、行程明细、票面标记（退票 / 改签）这些<b>票面上读来的字段</b>。<br>
           <span style="color:var(--text-secondary)">你手工设的状态、费用类别、报销人、报销批次
           一律不动；手改过的「项目 / 事由」也会保留。</span>`,
    note: '改前会自动备份一份台账到「台账备份」；原始发票文件不会被修改。',
    okText: '重新识别',
    cancelText: '取消',
    onOk: async () => {
      hideModal();
      const r = await call('refresh_rows', seqs);
      if (!r || r.error) { toast('重新识别失败：' + ((r && r.error) || '未知错误'), 'err'); return; }
      let msg = `已重新识别 ${r.updated} 条`;
      if (r.miss) msg += `，${r.miss} 条原件在磁盘上找不到`;
      if (r.fail) msg += `，${r.fail} 条读不了`;
      toast(msg, (r.miss || r.fail) ? 'warn' : 'ok');
      await loadLedger();
    },
  });
}

/** 删除台账里选中的记录（后端会先备份一张台账，删错了能去「台账备份」捞） */
function deleteSelected() {
  const seqs = selectedSeqs();
  if (!seqs.length) { toast('先在表格里勾选要删的记录', 'warn'); return; }
  showModal({
    title: `删除选中的 ${seqs.length} 条记录？`,
    html: `将从台账里删掉这 <b>${seqs.length}</b> 条记录，合计 <b>${money(selectedSum())}</b>。<br>
           <span style="color:var(--text-secondary)">只删台账里的这一行记录，<b>不会动原始发票文件</b>。</span>`,
    note: '删除前会自动备份一份台账到「台账备份」文件夹，反悔了可以从那儿恢复。',
    okText: `删除 ${seqs.length} 条`,
    cancelText: '再想想',
    onOk: async () => {
      hideModal();
      const r = await call('delete_rows', seqs);
      if (!r || r.error) { toast('删除失败：' + ((r && r.error) || '未知错误'), 'err'); return; }
      state.sel.clear();
      toast(`已删除 ${r.deleted} 条记录`, 'ok');
      await loadLedger();
    },
  });
}

/** 清空整张台账（只留表头） */
function clearLedger() {
  const n = state.allCount || state.rows.length;
  showModal({
    title: '清空整张台账？',
    html: `会把台账里现有的 <b>${n}</b> 行记录<b>全部删掉</b>，只留下表头，
           之后重新入库就是一份干净的台账。<br>
           <span style="color:var(--text-secondary)">原始发票文件不会被删除，重新扫一遍还能再建账。</span>`,
    note: '清空前会自动备份一份台账到「台账备份」文件夹；如果你只想删掉某几条，请用「删除选中」。',
    okText: '确认清空',
    cancelText: '取消',
    onOk: async () => {
      hideModal();
      const r = await call('clear_ledger');
      if (!r || r.error) { toast('清空失败：' + ((r && r.error) || '未知错误'), 'err'); return; }
      state.sel.clear();
      state.selRows = [];
      toast(`台账已清空（删掉 ${r.cleared} 条，备份在「台账备份」）`, 'ok');
      await loadLedger();
    },
  });
}

function selectedSeqs() {
  return Array.from(state.sel);
}

function selectedSum() {
  let s = 0;
  state.rows.forEach((r) => {
    if (state.sel.has(String(r['序号']))) s += num(r['价税合计']);
  });
  return s;
}

function syncSelectionUI() {
  const n = state.sel.size;
  $('sel-count').innerHTML = n ? `已选中 <b>${n}</b> 张` : '未选中';
  // 本页按凭证类型给个小分布，一眼看出这页里几张发票、几张行程单
  const byType = {};
  state.rows.forEach((r) => {
    const t = String(r['凭证类型'] || '其他凭证').trim() || '其他凭证';
    byType[t] = (byType[t] || 0) + 1;
  });
  const typeStr = Object.keys(byType).length > 1
    ? '　·　' + Object.entries(byType).sort((a, b) => b[1] - a[1])
      .map(([t, c]) => `${t} ${c}`).join(' / ')
    : '';
  $('ledger-hint').textContent =
    `本页 ${state.rows.length} 行，台账合计 ${state.allCount || 0} 行${typeStr}` +
    (n ? `　·　已选 ${n} 张（${money(selectedSum())}）` : '');
  const all = state.rows.length > 0 && state.rows.every((r) => state.sel.has(String(r['序号'])));
  const some = state.rows.some((r) => state.sel.has(String(r['序号'])));
  $('ck-all').checked = all;
  $('ck-all').indeterminate = !all && some;
}

async function loadLedger() {
  const hint = $('ledger-hint');
  const btn = $('btn-refresh');
  const f = currentFilter();
  hint.textContent = '查询中…';
  if (btn) { btn.disabled = true; btn.textContent = '查询中…'; }
  const r = await call('get_ledger', f);
  if (btn) { btn.disabled = false; btn.textContent = '查询'; }
  if (!r) { hint.textContent = '查询失败'; return; }
  if (r.error) {
    state.rows = [];
    state.attachLinks = {};
    $('ledger-body').innerHTML =
      `<tr><td colspan="13" class="empty">读取台账失败：${esc(r.error)}<br>
       （如果 Excel 正开着这个文件，请先关掉）</td></tr>`;
    hint.textContent = '读取失败';
    return;
  }
  state.rows = r.rows || [];
  state.allCount = r.all_count || state.rows.length;
  state.aggregates = r.aggregates || null;
  // 后端已经把「打车行程单」挪到它对应发票的紧后面，并告诉我们哪几行是附件
  state.attachLinks = r.attach_links || {};

  // 保留还存在的勾选
  const alive = new Set(state.rows.map((x) => String(x['序号'])));
  state.sel = new Set(Array.from(state.sel).filter((s) => alive.has(s)));

  fillSelect('f-ctype', r.ctypes, undefined, '全部类型');
  fillSelect('f-category', r.categories, undefined, '全部类别');
  fillSelect('f-batch', r.batches, undefined, '全部批次');
  fillSelect('f-person', r.persons, undefined, '全部人员');
  fillSelect('f-month', r.months, undefined, '全部月份');
  // 下拉被重置成空值时，说明筛掉了一部分，这里同步回 state.filter
  state.filter = currentFilter();

  const sum = $('f-summary');
  if (sum) {
    const hit = r.total === r.all_count
      ? `命中 <b>${r.total}</b> 条（全部）`
      : `命中 <b>${r.total}</b> 条 / 全部 ${r.all_count} 条`;
    sum.innerHTML = `${esc(filterSummary(state.filter))}　—　${hit}，合计 ${money(r.sum)}`;
  }

  renderLedgerBody();
  renderLedgerStats(r);
  syncSelectionUI();
  syncTvBadge(r.tv_todo);          // 「出行人核对」按钮上那个角标
}

/** 「出行人核对」按钮上挂个数字：还有几张的出行人没认出来 / 没人工核对过 */
function syncTvBadge(n) {
  const b = $('tv-todo-badge');
  if (!b) return;
  const k = Number(n) || 0;
  b.hidden = !k;
  b.textContent = k > 99 ? '99+' : String(k);
}

function renderLedgerStats(r) {
  const boxes = [
    ['本页张数', r.total, 'accent'],
    ['本页金额', money(r.sum), ''],
    ['待报销', `${r.pending_count} 张`, 'warn'],
    ['待报销金额', money(r.pending_sum), 'ok'],
    ['台账合计', `${r.all_count} 行`, 'mute'],
  ];
  const html = boxes.map(([lab, val, cls]) =>
    `<div class="stat-box ${cls}" style="flex:1 1 130px;padding:9px 12px">
      <div class="sb-num" style="font-size:17px">${esc(val)}</div>
      <div class="sb-lab">${esc(lab)}</div></div>`).join('');
  let holder = $('ledger-stats');
  if (!holder) {
    holder = document.createElement('div');
    holder.id = 'ledger-stats';
    holder.className = 'stat-row';
    holder.style.margin = '0 0 14px';
    // 放在卡片标题下面、批量操作条上面
    const head = $('ledger-body').closest('.card').querySelector('.card-head');
    head.insertAdjacentElement('afterend', holder);
  }
  holder.innerHTML = html;
}

/* 凭证类型 → 配色。后端可能给出各种叫法，认不出的一律走 ct-other */
const CTYPE_CLS = {
  '发票': 'ct-inv', '数电票': 'ct-inv', '增值税发票': 'ct-inv',
  '火车票': 'ct-train', '铁路电子客票': 'ct-train',
  '飞机行程单': 'ct-air', '航空行程单': 'ct-air', '飞机票': 'ct-air',
  '打车行程单': 'ct-taxi', '出租车票': 'ct-taxi', '网约车': 'ct-taxi',
  '登机牌': 'ct-board', '图片凭证': 'ct-img', '其他凭证': 'ct-other',
};
const ctypeClass = (t) => CTYPE_CLS[String(t || '').trim()] || 'ct-other';
/** 一行里给凭证类型画个小色标 */
function ctypeBadge(t) {
  if (!t) return '<span class="td-dash">—</span>';
  return `<span class="tag ct ${ctypeClass(t)}">${esc(t)}</span>`;
}

/**
 * 票面标记（退票 / 差额退票 / 改签 / 作废 / 红字）单独给个醒目标签。
 * 票面上这些字是**红字**印的，意思是「这张不能当普通票价看」——
 * 退票费 10 元、改签差价 1 元都不是车票钱，台上必须一眼能看见。
 */
function stampTag(s) {
  const t = String(s || '').trim();
  if (!t) return '';
  const amber = /改签/.test(t);            // 改签用橙，退票/作废/红字用红
  return `<span class="tag tag-stamp ${amber ? 'stamp-amber' : 'stamp-red'}"` +
         ` title="票面标记：${esc(t)}">${esc(t)}</span>`;
}

function renderLedgerBody() {
  const body = $('ledger-body');
  if (!state.rows.length) {
    body.innerHTML = `<tr><td colspan="13" class="empty">
      没有符合条件的凭证。<br>先去「凭证入库」建账，或把上面的筛选条件清空。</td></tr>`;
    return;
  }
  body.innerHTML = state.rows.map((r) => {
    const seq = String(r['序号']);
    const done = r['状态'] === '已报销';
    const tip = String(r['提示'] || '');
    const tags = [];
    if (done) tags.push('<span class="tag tag-ok">已报销</span>');
    else tags.push('<span class="tag tag-mute">未报销</span>');
    if (tip.includes('重复')) tags.push('<span class="tag tag-dup">重复</span>');
    if (tip.includes('红字')) tags.push('<span class="tag tag-pending">红字</span>');
    const seller = r['销方名称'] || '';
    const proj = r['项目/事由'] || '';
    const trip = String(r['行程/明细'] || '').trim();
    const no = String(r['发票号码'] || '').trim();
    const tvName = String(r['出行人'] || '').trim();
    // 是「跟在发票后面的附件」吗（后端算出的一对一配对）
    const lk = state.attachLinks[seq];
    const rowTip = lk
      ? `是第 ${lk.main_seq} 号发票（${money(lk.main_amount)}）的证明附件，排在它后面；点这一行看全部信息`
      : '点这一行看凭证全部信息';
    return `
      <tr data-seq="${seq}" class="${done ? 'row-done' : ''}${lk ? ' row-att' : ''}" title="${esc(rowTip)}">
        <td class="td-c"><input type="checkbox" data-seq="${seq}"${state.sel.has(seq) ? ' checked' : ''}></td>
        <td class="td-idx">${lk ? '<span class="att-mark" title="发票的附件">↳</span>' : ''}${pad2(seq)}</td>
        <td>${tags.join('')}</td>
        <td class="td-nowrap">${ctypeBadge(r['凭证类型'])}${stampTag(r['票面标记'])}${lk ? '<span class="tag tag-att">附件</span>' : ''}</td>
        <td class="td-nowrap td-c" title="${tvName ? esc('出行人：' + tvName) : '票面上没有人名（打车行程单常见），选中这行点「设置出行人」手填'}">${tvName
          ? esc(tvName)
          : '<span class="td-dash">—</span>'}</td>
        <td class="td-time">${esc(r['开票日期'])}</td>
        <td class="td-no">${no ? esc(no) : '<span class="td-dash">—</span>'}</td>
        <td class="td-trip" title="${esc(trip)}">${trip ? esc(trip) : '<span class="td-dash">—</span>'}</td>
        <td class="td-proj" title="${esc(seller)}">${esc(seller)}</td>
        <td class="td-proj" title="${esc(proj)}">${esc(proj)}</td>
        <td class="td-c td-nowrap">${esc(r['费用类别'])}</td>
        <td class="td-r">${money2(r['价税合计'])}</td>
        <td class="td-time">${esc(r['报销批次'])}</td>
      </tr>`;
  }).join('');
}

/* =====================================================================
 * 凭证明细弹层
 * ---------------------------------------------------------------------
 * 用户反馈「点开的预览很乱」——原来是 17 个字段竖着堆一列、长路径撑破版面。
 * 现在改成：顶部一行大号号码 + 金额，下面按「基本信息 / 行程明细 / 金额 /
 * 归属报销 / 文件」分区卡片；路径单独折行 + 两个按钮，不塞进字段列表里。
 * ===================================================================== */
function showInvoiceDetail(r) {
  const v = (k) => String(r[k] === null || r[k] === undefined ? '' : r[k]).trim();
  const ctype = v('凭证类型') || (v('发票号码') ? '发票' : '其他凭证');
  const trip = v('行程/明细');
  const path = v('文件路径');
  const tip = v('提示');
  const done = v('状态') === '已报销';
  const no = v('发票号码');
  const pay = v('价税合计');

  const badges = [`<span class="tag ct ${ctypeClass(ctype)}">${esc(ctype)}</span>`];
  badges.push(done ? '<span class="tag tag-ok">已报销</span>' : '<span class="tag tag-mute">未报销</span>');
  badges.push(stampTag(v('票面标记')));
  // 「票面标记」已经写着红字了就不再叠一个 —— 同一条信息不显示两遍
  if (tip.includes('红字') && !v('票面标记')) badges.push('<span class="tag tag-pending">红字</span>');
  if (tip.includes('重复')) badges.push('<span class="tag tag-dup">重复</span>');
  // 打车行程单这类「发票的证明附件」：说清楚它是谁的附件、金额算在谁身上
  const lk = state.attachLinks[v('序号')];
  if (lk) {
    badges.push(`<span class="tag tag-att">↳ 第 ${esc(lk.main_seq)} 号发票的附件（${money(lk.main_amount)}，不重复计）</span>`);
  }

  const head = `
    <div class="id-top">
      <div class="id-main">
        <div class="id-no">${no ? esc(no) : '<span class="id-none">无发票号码</span>'}</div>
        <div class="id-sub">${esc(v('票种')) || esc(ctype)}</div>
      </div>
      <div class="id-amt">
        <div class="id-amt-n">${pay ? money(pay) : '—'}</div>
        <div class="id-amt-l">价税合计</div>
      </div>
    </div>
    <div class="id-badges">${badges.join('')}</div>`;

  // 一个分区：rows = [标签, 值(已转义 html), 是否占整行]
  const sec = (title, rows) => {
    const cells = rows.filter(([, val]) => val !== undefined && val !== null && val !== '')
      .map(([k, val, wide]) => `<div class="id-cell${wide ? ' wide' : ''}">
          <div class="id-k">${esc(k)}</div><div class="id-v">${val}</div></div>`).join('');
    return cells ? `<div class="id-sec"><div class="id-sec-t">${esc(title)}</div>
      <div class="id-grid">${cells}</div></div>` : '';
  };
  const mono = (s) => (s ? `<span class="mono">${esc(s)}</span>` : '');

  let body = head;
  body += sec('基本信息', [
    ['开票日期', esc(v('开票日期'))],
    ['凭证类型', esc(ctype)],
    ['票种', esc(v('票种'))],
    ['发票号码', mono(no)],
    ['购方名称', esc(v('购方名称')) || '<span class="id-na">—</span>', true],
    ['销方名称', esc(v('销方名称')) || '<span class="id-na">—</span>', true],
  ]);
  const tvName = v('出行人');
  // 出行人这格：打车行程单票面只有手机号、没有姓名，认不出时就给个能点的提示
  const tvCell = tvName
    ? esc(tvName) + (tip.includes('按文件夹名推断')
      ? ' <span class="tag tag-pending" title="票面没有印人名，是按你分文件夹的名字推出来的">推断</span>' : '')
    : '<span class="id-na" title="票面没写人名">未识别 · 选中这行点「设置出行人」手填</span>';
  if (trip || tvName) {
    body += sec('行程 / 明细', [
      ['行程明细', trip ? `<span class="id-trip">${esc(trip)}</span>` : '', true],
      ['出行人', tvCell],
    ]);
  }
  body += sec('金额', [
    ['不含税金额', v('不含税金额') ? money(v('不含税金额')) : (pay ? '—' : '')],
    ['税额', v('税额') ? money(v('税额')) : (pay ? '—' : '')],
    ['价税合计', `<span class="id-money">${pay ? money(pay) : '—'}</span>`],
  ]);
  body += sec('归属 / 报销', [
    ['费用类别', esc(v('费用类别')) || '<span class="id-na">—</span>'],
    ['报销人', esc(v('报销人'))],
    ['状态', esc(v('状态'))],
    ['报销批次', esc(v('报销批次')) || '<span class="id-na">—</span>'],
    ['项目 / 事由', esc(v('项目/事由')) || '<span class="id-na">—</span>', true],
    ['项目编号', mono(v('项目编号'))],
    ['入库时间', esc(v('入库时间'))],
  ]);
  if (tip) {
    body += sec('提示', [[tip.includes('重复') ? '重复风险' : (tip.includes('红字') ? '红字发票' : '说明'),
      `<span class="id-tip">${esc(tip)}</span>`, true]]);
  }
  if (path) {
    body += `<div class="id-sec"><div class="id-sec-t">文件</div>
      <div class="id-file">
        <div class="id-file-name">${esc(baseName(path))}</div>
        <div class="id-file-dir" title="${esc(path)}">${esc(path)}</div>
        <div class="id-acts">
          <button class="btn btn-sm btn-primary" data-act="open_file" data-path="${esc(path)}">打开文件</button>
          <button class="btn btn-sm" data-act="reveal" data-path="${esc(path)}">在文件夹中显示</button>
        </div>
      </div></div>`;
  }

  body += `<div class="id-sec"><div class="id-sec-t">操作</div>
    <div class="id-acts">
      <button class="btn btn-sm btn-danger" data-act="del_row" data-seq="${esc(r['序号'])}">
        从台账删除这一条</button>
    </div></div>`;

  showModal({
    title: `凭证明细 · 第 ${esc(r['序号'])} 行`,
    html: body,
    note: path ? '' : '这条记录没有关联的原始文件（可能是在别的机器上入库的，或原文已被移走）。',
    noOk: true,
    cancelText: '关闭',
    wide: true,
  });
}

/** 从明细弹层里删掉单独一条 */
function deleteOneRow(seq) {
  const row = state.rows.find((r) => String(r['序号']) === String(seq));
  const label = row
    ? (row['发票号码'] || row['行程/明细'] || `${row['凭证类型'] || '凭证'} ${money(row['价税合计'])}`)
    : `第 ${seq} 行`;
  showModal({
    title: '删除这一条记录？',
    html: `将删除台账第 <b>${esc(seq)}</b> 行：<br><b>${esc(String(label).slice(0, 70))}</b>`,
    note: '删除前会自动备份台账；原始发票文件不受影响。',
    okText: '删除',
    cancelText: '取消',
    onOk: async () => {
      hideModal();
      const r = await call('delete_rows', [String(seq)]);
      if (!r || r.error) { toast('删除失败：' + ((r && r.error) || '未知错误'), 'err'); return; }
      toast('已删除 1 条记录', 'ok');
      await loadLedger();
    },
  });
}

async function markSelected(status) {
  const seqs = selectedSeqs();
  if (!seqs.length) { toast('先在表格里勾选发票', 'warn'); return; }
  const r = await call('mark_rows', seqs, status, '', $('in-person').value.trim());
  if (!r) return;
  toast(`已把 ${r.updated} 张标记为「${status}」`, 'ok');
  await loadLedger();
}

async function applyCategory() {
  const seqs = selectedSeqs();
  if (!seqs.length) { toast('先在表格里勾选发票', 'warn'); return; }
  const cat = $('set-category').value;
  if (!cat) { toast('先选一个费用类别', 'warn'); return; }
  const r = await call('update_rows', seqs, { 费用类别: cat });
  if (!r) return;
  toast(`已把 ${r.updated} 张的费用类别改成「${cat}」`, 'ok');
  await loadLedger();
}

/* ---------------------------------------------------------------------
 * 出行人：票面读不到（打车行程单最常见）就人工填。
 * 后端认人分三层：票面 / 手机号对照表 / 文件夹名；三层都没认出来才轮到这儿。
 * ------------------------------------------------------------------- */
async function setTraveler() {
  const seqs = selectedSeqs();
  if (!seqs.length) { toast('先在表格里勾选凭证', 'warn'); return; }
  const pickedRows = state.rows.filter((r) => seqs.includes(String(r['序号'])));
  const picked = pickedRows.map((r) => String(r['出行人'] || '').trim());
  const uniq = [...new Set(picked.filter(Boolean))];
  const known = [...new Set(state.rows.map((r) => String(r['出行人'] || '').trim()).filter(Boolean))];
  const nowTxt = uniq.length
    ? `现在这 ${seqs.length} 张里填着：${uniq.map(esc).join('、')}`
    : `现在这 ${seqs.length} 张都没填出行人`;
  showModal({
    title: `设置出行人（${seqs.length} 张）`,
    html: `
      <div class="card-hint" style="display:block;margin-bottom:12px">${nowTxt}。</div>
      <label class="fi-label" style="display:inline-flex;align-items:center;gap:8px">出行人
        <input class="fi" id="tv-input" type="text" placeholder="比如 王小明" style="width:180px">
      </label>
      ${known.length ? `<div class="card-hint" style="display:block;margin:14px 0 6px">台账里已经认出来的（点一下填进输入框）：</div>
        <div class="id-acts">${known.slice(0, 14).map((n) =>
    `<button class="btn btn-sm" data-tv="${esc(n)}">${esc(n)}</button>`).join('')}</div>` : ''}
      <div class="card-hint" style="display:block;margin-top:14px">保存时留空 = 把这 ${seqs.length} 张的出行人清掉。</div>`,
    okText: '保存',
    cancelText: '取消',
    onOk: async () => {
      const name = ($('tv-input').value || '').trim();
      hideModal();
      // 走核对接口、而不是通用的 update_rows：手工指定也算「人工核对过」，
      // 免得在核对面板里又被提醒一遍。留空 = 清掉 + 撤掉核对标记。
      const r = await call('save_traveler_review',
        seqs.map((s) => ({ seq: String(s), name })), false);
      if (!r || r.error) { toast('设置失败：' + ((r && r.error) || '未知错误'), 'err'); return; }
      toast(name ? `已把 ${r.updated} 张的出行人设为「${name}」` : `已清空 ${r.updated} 张的出行人`, 'ok');
      await loadLedger();
    },
  });
  document.querySelectorAll('#modal-text [data-tv]').forEach((b) =>
    b.addEventListener('click', () => { $('tv-input').value = b.dataset.tv; }));
}

/* ---------------------------------------------------------------------
 * 出行人核对 —— 「三层兜底 + 人工核对」里的后半截。
 *
 * 自动认人是三层兜底：票面姓名 → 手机号对照表 → 文件夹名。后两层终究是**猜**，
 * 票面读不到的干脆是空的。所以要有一个人工过一遍的环节：这里把整本台账摆出来，
 * 每张票「是怎么认出来的、认成了谁」一眼可见，当场改掉。确认过的记一笔，
 * 以后不再提醒；名字再被改掉，提醒会自动回来。
 * ------------------------------------------------------------------- */
let tvr = null;      // {rows, stat, people, onlyTodo, syncPhone, remember, edits}

async function openTravelerReview() {
  const r = await call('traveler_review');
  if (!r || r.error) { toast('读取核对清单失败：' + ((r && r.error) || '未知错误'), 'err'); return; }
  if (!r.rows || !r.rows.length) { toast('台账还是空的，先做一次「凭证入库」', 'warn'); return; }
  tvr = {
    rows: r.rows, stat: r.stat || {}, people: r.people || [],
    onlyTodo: true, syncPhone: true, remember: true, edits: {},
  };
  const dl = $('tv-people');     // 姓名补全：把台账里出现过的名字喂给 datalist
  if (dl) dl.innerHTML = tvr.people.map((n) => `<option value="${esc(n)}"></option>`).join('');
  showModal({
    title: '出行人核对',
    xwide: true,
    html: '<div id="tvr-box"></div>',
    note: '保存 = 这些行记成「已人工核对」，名字以后被改掉才会再提醒；勾了「记住手机号」的行，会顺手把手机号 → 姓名写进设置里的对照表，下次同一部手机开的行程单就能自动认人。',
    okText: '保存核对结果',
    cancelText: '关闭',
    onOk: saveTravelerReview,
  });
  renderTvr();
}

/** 这一行的出行人是**怎么认出来的** —— 直接决定要不要人工看一眼 */
function tvrBadge(x) {
  if (x.done) return ['tag-ok', '已核对', '这张人工确认过了：' + (x.cur || '（空）')];
  if (x.state === 'blank') return ['tag-dup', '没认出来', '票面上没姓名、手机号也不在对照表里 —— 请手填'];
  if (x.state === 'diff') {
    return ['tag-pending', '跟自动值不符',
            `台账里是「${x.cur}」，按${x.src || '票面'}重算应该是「${x.auto}」`];
  }
  if (x.state === 'todo') return ['tag-pending', '还没写进台账', `能认成「${x.auto}」，台账这格还空着`];
  if (x.src === '手机号') return ['tag-pending', '手机号译出', '按手机号对照表翻出来的，请核对'];
  if (x.src === '文件夹') return ['tag-pending', '文件夹推断', '按存放文件夹的名字推断的，请核对'];
  if (x.src === '票面' || x.src === '文件名') return ['tag-ok', x.src + '读到', '票面上白纸黑字写着姓名'];
  return ['tag-mute', x.src || '—', ''];
}

function tvrRowHtml(x) {
  const edited = Object.prototype.hasOwnProperty.call(tvr.edits, x.seq);
  let name = edited ? tvr.edits[x.seq] : String(x.cur || '');
  // 台账这格空着、但系统能认出来 → 先把名字填上（斜体虚线），用户扫一眼、改掉错的就行，
  // 不用一张张手打；不改就等于确认。票面上写着姓名 / 手机号译出来的，多半是对的。
  const autoFill = !edited && !name && !!x.auto;
  if (autoFill) name = x.auto;
  const b = tvrBadge(x);
  const det = String(x.detail || '').trim();
  const meta = [x.date, x.person ? '报销人 ' + x.person : ''].filter(Boolean).join(' · ');
  const tip = autoFill ? `系统自动认出来的「${x.auto}」——不改等于确认，改掉就是你的` : '';
  return `<tr class="tvr-row ${x.done ? 'is-done' : ''}${x.need ? ' is-need' : ''}" data-seq="${esc(x.seq)}">
    <td class="tvr-ck"><input type="checkbox" class="tvr-ckb" data-seq="${esc(x.seq)}"></td>
    <td class="tvr-idx">${esc(x.seq)}</td>
    <td class="tvr-inv">
      <div class="tvr-t1">${ctypeBadge(x.ctype)}${det ? esc(det.slice(0, 30)) : '<span class="td-dash">—</span>'}</div>
      <div class="tvr-t2">${esc(meta)}${x.no ? ' · ' + esc(x.no) : ''}</div>
    </td>
    <td class="td-r td-nowrap">${money(x.total)}</td>
    <td class="td-nowrap tvr-ph">${x.phone ? esc(x.phone) : '<span class="td-dash">—</span>'}</td>
    <td><input class="fi tvr-inp${autoFill ? ' is-auto' : ''}" data-seq="${esc(x.seq)}"
        data-phone="${esc(x.phone || '')}" list="tv-people" value="${esc(name)}"
        placeholder="填姓名" autocomplete="off" title="${esc(tip)}"></td>
    <td class="tvr-src"><span class="tag ${b[0]}" title="${esc(b[2])}">${esc(b[1])}</span></td>
  </tr>`;
}

function renderTvr() {
  const box = $('tvr-box');
  if (!box || !tvr) return;
  const st = tvr.stat;
  const n = (k) => Number(st[k] || 0);
  const list = tvr.onlyTodo ? tvr.rows.filter((x) => x.need) : tvr.rows;
  const rowsHtml = list.length
    ? list.map(tvrRowHtml).join('')
    : `<tr><td colspan="7" class="empty">${
        tvr.onlyTodo ? '要核对的都过完了 ✅ 取消勾选「只看要核对的」可以再全部过一遍'
                     : '台账里没有凭证'}</td></tr>`;
  box.innerHTML = `
    <div class="tvr-stats">
      <span class="tvr-st"><b>${n('total')}</b> 张凭证</span>
      <span class="tvr-st${n('need') ? ' is-need' : ''}"><b>${n('need')}</b> 待核对</span>
      ${n('blank') ? `<span class="tvr-st is-err">没认出来 <b>${n('blank')}</b></span>` : ''}
      <span class="tvr-st">票面 / 文件名 <b>${n('票面') + n('文件名')}</b></span>
      <span class="tvr-st">手机号译出 <b>${n('手机号')}</b></span>
      <span class="tvr-st">文件夹推断 <b>${n('文件夹')}</b></span>
      <span class="tvr-st is-ok">已核对 <b>${n('done')}</b></span>
    </div>
    <div class="tvr-bar">
      <label class="ck"><input type="checkbox" data-tvr="only" ${tvr.onlyTodo ? 'checked' : ''}> 只看要核对的</label>
      <label class="ck"><input type="checkbox" data-tvr="sync" ${tvr.syncPhone ? 'checked' : ''}> 同手机号一起改</label>
      <label class="ck"><input type="checkbox" data-tvr="remember" ${tvr.remember ? 'checked' : ''}> 记住手机号对应的人</label>
      <span class="tvr-sp"></span>
      <input class="fi tvr-bulk" id="tvr-bulk" type="text" list="tv-people"
             placeholder="批量填个名字" autocomplete="off">
      <button class="btn btn-sm" id="tvr-bulk-apply">应用到勾选的</button>
    </div>
    <div class="tvr-wrap">
      <table class="tvr-tbl">
        <thead><tr>
          <th class="tvr-ck"><input type="checkbox" id="tvr-all" title="全选 / 全不选"></th>
          <th class="tvr-idx">#</th>
          <th>凭证 / 明细</th>
          <th class="td-r">金额</th>
          <th>手机号</th>
          <th>出行人</th>
          <th>凭什么是这个人</th>
        </tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    </div>`;
  bindTvr(box);
}

/** 事件委托只绑一次：内容重渲染后照样管用 */
function bindTvr(box) {
  if (box.dataset.bound) return;
  box.dataset.bound = '1';
  box.addEventListener('change', onTvrChange);
  box.addEventListener('input', onTvrInput);
  box.addEventListener('click', onTvrClick);
}

function onTvrInput(e) {
  const t = e.target;
  if (t.classList && t.classList.contains('tvr-inp')) tvr.edits[t.dataset.seq] = t.value.trim();
}

function onTvrChange(e) {
  const t = e.target;
  if (t.dataset && t.dataset.tvr) {
    const k = t.dataset.tvr;
    if (k === 'only') tvr.onlyTodo = t.checked;
    else if (k === 'sync') tvr.syncPhone = t.checked;
    else if (k === 'remember') { tvr.remember = t.checked; return; }   // 不影响列表，不用重画
    renderTvr();
    return;
  }
  if (!t.classList || !t.classList.contains('tvr-inp')) return;
  const seq = t.dataset.seq;
  const name = t.value.trim();
  tvr.edits[seq] = name;
  const ph = t.dataset.phone;
  if (!tvr.syncPhone || !ph) return;
  // 同一部手机开的票，往往就是同一个人 —— 改一个，把这批一起改掉
  tvr.rows.forEach((x) => { if (x.phone === ph) tvr.edits[x.seq] = name; });
  let n = 0;
  document.querySelectorAll('.tvr-inp').forEach((el) => {
    if (el !== t && el.dataset.phone === ph && el.value.trim() !== name) { el.value = name; n++; }
  });
  if (n) toast(`同手机号的另外 ${n} 张也改成了「${name || '空'}」`, 'ok');
}

function onTvrClick(e) {
  const t = e.target;
  if (t.id === 'tvr-all') {
    document.querySelectorAll('.tvr-ckb').forEach((c) => { c.checked = t.checked; });
    return;
  }
  if (t.id !== 'tvr-bulk-apply') return;
  const cks = Array.from(document.querySelectorAll('.tvr-ckb')).filter((c) => c.checked);
  if (!cks.length) { toast('先在左边勾几行', 'warn'); return; }
  const name = ($('tvr-bulk').value || '').trim();
  cks.forEach((c) => {
    const seq = c.dataset.seq;
    tvr.edits[seq] = name;
    const inp = document.querySelector(`.tvr-inp[data-seq="${CSS.escape(seq)}"]`);
    if (inp) inp.value = name;
  });
  toast(`已把勾选的 ${cks.length} 张填成「${name || '空'}」，记得点右下角保存`, 'ok');
}

async function saveTravelerReview() {
  if (!tvr) { hideModal(); return; }
  // 以输入框里的值为准：没动过的行里也可能带着「系统预填的名字」，那也要落进台账
  const vals = {};
  document.querySelectorAll('.tvr-inp').forEach((el) => { vals[el.dataset.seq] = el.value.trim(); });
  const items = [];
  tvr.rows.forEach((x) => {
    const onScreen = Object.prototype.hasOwnProperty.call(vals, x.seq);
    const edited = Object.prototype.hasOwnProperty.call(tvr.edits, x.seq);
    if (!onScreen && !edited) return;              // 没显示、也没动过的，别去碰
    items.push({
      seq: x.seq,
      name: onScreen ? vals[x.seq] : String(tvr.edits[x.seq] || ''),
      phone: x.phone || '',
    });
  });
  if (!items.length) { hideModal(); toast('没有要保存的内容', 'warn'); return; }
  const empty = items.filter((it) => !it.name).length;
  const remember = !!tvr.remember;
  hideModal();
  const r = await call('save_traveler_review', items, remember);
  if (!r || r.error) { toast('保存失败：' + ((r && r.error) || '未知错误'), 'err'); return; }
  let msg = `核对完成：确认 ${r.updated} 张`;
  if (r.remembered) msg += `，记住 ${r.remembered} 个手机号`;
  if (empty) msg += `；还有 ${empty} 张没填姓名`;
  toast(msg, empty ? 'warn' : 'ok');
  tvr = null;
  await loadLedger();
}

function openLedgerFile() {
  const a = api();
  if (!a) return;
  a.get_paths().then((p) => {
    if (!p) return;
    if (!p.ledger_exists) {
      toast('台账文件还没生成，先做一次「凭证入库」', 'warn');
      call('open_folder', 'app');
      return;
    }
    call('open_file', p.ledger);
  }).catch(() => {});
}

/* =====================================================================
 * 三、报销单 / 合并
 * ===================================================================== */
function initReportPage() {
  document.querySelectorAll('#kind-pills .pill').forEach((p) => {
    p.addEventListener('click', () => {
      state.kind = p.dataset.kind;
      localStorage.setItem(LS.kind, state.kind);
      syncKindPills();
    });
  });
  document.querySelectorAll('#fmt-pills .pill').forEach((p) => {
    p.addEventListener('click', () => {
      state.fmt = p.dataset.fmt;
      localStorage.setItem(LS.fmt, state.fmt);
      syncFmtPills();
    });
  });
  $('r-date').value = todayCn();
  initAlw();
  renderAlw();
  initJy();
  renderJy();
  $('btn-make-report').addEventListener('click', makeReport);
  $('btn-merge-2up').addEventListener('click', () => doMerge('2up'));
  $('btn-merge-plain').addEventListener('click', () => doMerge('plain'));

  $('btn-out-refresh').addEventListener('click', loadOutputs);
  $('btn-out-folder').addEventListener('click', () => call('open_folder', 'output'));
  $('btn-out-clear').addEventListener('click', clearOutputs);

  $('out-list').addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-i]');
    if (!btn) return;
    const item = state.outputs[Number(btn.dataset.i)];
    if (!item) return;
    if (btn.dataset.act === 'open') call('open_file', item.path);
    else if (btn.dataset.act === 'reveal') call('reveal', item.path);
    else if (btn.dataset.act === 'del') deleteOutput(item);
  });
}

function syncKindPills() {
  document.querySelectorAll('#kind-pills .pill').forEach((p) =>
    p.classList.toggle('active', p.dataset.kind === state.kind));
  const isTrip = state.kind === '差旅费报销单';
  const isJy = state.kind === JY_KIND;
  $('trip-row').style.display = isTrip ? '' : 'none';
  // 两块补助是二选一的：模板二用「按人」，另外两种用「人数×天数×标准」
  const alwBlock = document.querySelector('#page-report .alw-block:not(#jy-block)');
  if (alwBlock) alwBlock.hidden = isJy;
  const jyBlock = $('jy-block');
  if (jyBlock) jyBlock.hidden = !isJy;
  // 选「差旅费报销单」时自动把补助那块打开（用户没手动关过的话），省一步点击
  if (isTrip && !state.alwOn && !state.alwOff) {
    state.alwOn = true;
    if (!state.alw.length) state.alw = [newAlwRow()];
    saveAlw();
  }
  renderAlw();
  if (isJy) {
    renderJy();
    refreshJyPeople();                 // 拉一份「按出行人的票据合计」（跟单据同一口径）
  }
}

function syncFmtPills() {
  document.querySelectorAll('#fmt-pills .pill').forEach((p) =>
    p.classList.toggle('active', p.dataset.fmt === state.fmt));
  const label = { pdf: 'PDF', xlsx: 'Excel', both: 'PDF + Excel' }[state.fmt] || 'PDF';
  const btn = $('btn-make-report');
  if (btn) btn.textContent = `生成报销单（${label}）`;
}

/** 删除一个生成的文件（后端送回收站，不是永久删除） */
function deleteOutput(item) {
  const kb = item.size ? `（${(item.size / 1024).toFixed(0)} KB）` : '';
  showModal({
    title: '删除这个生成的文件？',
    html: `<b>${esc(item.name)}</b>${kb}<br>
           <span style="color:var(--text-secondary)">${esc(item.kind || '文件')}　${esc(item.time || '')}</span>`,
    note: '文件会被送进 Windows 回收站（不是永久删除），需要的话可以去回收站还原。台账数据不受影响。',
    okText: '删除',
    cancelText: '取消',
    onOk: async () => {
      hideModal();
      const r = await call('delete_output', item.path);
      if (!r || r.error) { toast('删除失败：' + ((r && r.error) || '未知错误'), 'err'); return; }
      toast(`已删除 ${r.name || item.name}（已送回收站）`, 'ok');
      await loadOutputs();
    },
  });
}

/** 一次把列出来的生成文件全删掉 */
function clearOutputs() {
  const items = state.outputs.filter((o) => o.exists);
  if (!items.length) { toast('列表里没有可删除的文件', 'warn'); return; }
  showModal({
    title: `删除全部 ${items.length} 个生成文件？`,
    html: `会把列表里这 <b>${items.length}</b> 个文件（报销单 / 合并 PDF / 统计 Excel）<b>全部送进回收站</b>。<br>
           <span style="color:var(--text-secondary)">标签页刷新后列表就空了；台账数据不受影响。</span>`,
    note: '送回收站不是永久删除，反悔了可以去回收站还原。',
    okText: `全部删除（${items.length} 个）`,
    cancelText: '取消',
    onOk: async () => {
      hideModal();
      let n = 0;
      for (const o of items) {
        const r = await call('delete_output', o.path);
        if (r && !r.error) n += 1;
      }
      toast(`已删除 ${n} 个文件（已送回收站）`, 'ok');
      await loadOutputs();
    },
  });
}

/** 后端记着所有生成过的文件（输出/生成记录.json），进页面就能看到历史 */
async function loadOutputs() {
  const r = await call('list_outputs');
  state.outputs = (r && r.items) || [];
  renderOutputs();
}

function goToReport() {
  const rows = state.rows.filter((r) => state.sel.has(String(r['序号'])));
  if (!rows.length) { toast('先在台账里勾选要报销的发票', 'warn'); return; }
  state.selRows = rows;
  renderReportRows();
  switchPage('page-report');
  refreshJyPeople();                 // 模板二版式要用：按出行人的票据合计（其他类型里它是空转）
}

function renderReportRows() {
  const rows = state.selRows;
  const body = $('rep-body');
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="7" class="empty">还没有选凭证 —— 去「凭证台账」勾选后点「生成报销单 →」</td></tr>';
    $('rep-total').innerHTML = '';
    $('rep-hint').textContent = '未选中';
    return;
  }
  // 附件判定要跟后端一致：行程单和它的发票**都在这次选择里**才算附件（不重复计金额）
  const seqs = new Set(rows.map((r) => String(r['序号'])));
  const linkOf = (r) => {
    const lk = state.attachLinks[String(r['序号'])];
    return lk && seqs.has(String(lk.main_seq)) ? lk : null;
  };
  body.innerHTML = rows.map((r, i) => {
    // 行程单没有「项目/事由」，就用行程明细顶上，别让这一列空着
    const desc = String(r['项目/事由'] || '').trim() || String(r['行程/明细'] || '').trim();
    const lk = linkOf(r);
    const badges = ctypeBadge(r['凭证类型']) + stampTag(r['票面标记'])
      + (lk ? `<span class="tag tag-att" title="是第 ${esc(lk.main_seq)} 号发票（${money(lk.main_amount)}）的证明附件，金额不重复计">↳ 附件</span>` : '');
    return `
    <tr class="${lk ? 'row-att' : ''}">
      <td class="td-idx">${lk ? '<span class="att-mark" title="跟在发票后面">↳</span>' : ''}${pad2(i + 1)}</td>
      <td class="td-nowrap">${badges}</td>
      <td class="td-time">${esc(r['开票日期'])}</td>
      <td class="td-no">${esc(r['发票号码']) || '<span class="td-dash">—</span>'}</td>
      <td class="td-c">${esc(r['费用类别'])}</td>
      <td class="td-proj" title="${esc(desc)}">${esc(desc)}</td>
      <td class="td-r">${money2(r['价税合计'])}</td>
    </tr>`;
  }).join('');
  const bills = rows.filter((r) => !linkOf(r));
  const nAtt = rows.length - bills.length;
  const sum = bills.reduce((s, r) => s + num(r['价税合计']), 0);
  const alw = state.alwOn ? alwSum() : 0;
  $('rep-hint').textContent = `共 ${rows.length} 条`;
  $('rep-total').innerHTML =
    `<span>凭证条数：<b>${rows.length}</b> 条</span>
     <span>凭证金额：<b>${money(sum)}</b></span>`
    + (nAtt ? `<span class="rep-att-note">其中 ${nAtt} 张行程单是发票的附件，不重复计</span>` : '')
    + (alw ? `<span>差旅费补助：<b>${money(alw)}</b></span>
       <span>报销合计：<b>${money(sum + alw)}</b></span>` : '');
  updateAlwSum();
}

/* =====================================================================
 * 差旅费补助（几个人 × 几天 × 标准）
 * ---------------------------------------------------------------------
 * 用户要求：「在生成报销单的时候，要能够增加差旅费，因为出差有差旅费啊，
 * 要计算几个人几天，多少钱，可以我自己设置这个，我自己填写」。
 * 所以这里给一张小表，一行一个补助项目，金额 = 人数 × 天数 × 标准；
 * 界面上实时算给你看，也可以直接改金额（改过就按你填的算，不再自动覆盖）。
 * 这几行存在 localStorage 里，下次进来不用重填。
 * ===================================================================== */
const ALW_DEFAULT_NAME = '伙食补助费';

function loadAlw() {
  try {
    const raw = JSON.parse(localStorage.getItem(LS.alw) || '[]');
    return Array.isArray(raw) ? raw : [];
  } catch (e) {
    return [];
  }
}
function saveAlw() {
  try {
    localStorage.setItem(LS.alw, JSON.stringify(state.alw));
    localStorage.setItem(LS.alwOn, state.alwOn ? '1' : '0');
  } catch (e) { /* 隐私模式下存不了，不影响使用 */ }
}
/** 输入框里的数字（空 / 乱填都当 0） */
function alwNumber(v) {
  const f = parseFloat(String(v === null || v === undefined ? '' : v).replace(/[^\d.\-]/g, ''));
  return isNaN(f) ? 0 : f;
}
/** 一行的金额：手改过就用他填的，否则 人数 × 天数 × 标准 */
function alwRowAmount(r) {
  if (r && r.manual) return Math.round(alwNumber(r.amount) * 100) / 100;
  return Math.round(alwNumber(r.people) * alwNumber(r.days) * alwNumber(r.rate) * 100) / 100;
}
function alwSum() {
  return Math.round((state.alw || []).reduce((s, r) => s + alwRowAmount(r), 0) * 100) / 100;
}
/** 出差起止 → 天数（填了才算），用来预填补助天数 */
function tripDays() {
  const a = $('r-start') && $('r-start').value.trim();
  const b = $('r-end') && $('r-end').value.trim();
  if (!a || !b) return '';
  const d1 = new Date(a.replace(/\//g, '-') + 'T00:00:00');
  const d2 = new Date(b.replace(/\//g, '-') + 'T00:00:00');
  if (isNaN(d1) || isNaN(d2)) return '';
  const n = Math.round((d2 - d1) / 86400000) + 1;
  return n > 0 ? String(n) : '';
}
function newAlwRow(seed) {
  const d = tripDays();
  return {
    name: (seed && seed.name) || ALW_DEFAULT_NAME,
    people: '1',
    days: d || '1',
    rate: '',
    amount: '',
    manual: false,
  };
}
function addAlwRow(seed) {
  if (!state.alw.length || !state.alwOn) state.alw = [newAlwRow(seed)];
  else state.alw.push(newAlwRow({ name: '' }));
  saveAlw();
  renderAlw();
}
function alwPayload() {
  if (!state.alwOn) return [];
  return (state.alw || [])
    .map((r) => ({ name: String(r.name || '').trim(), people: r.people, days: r.days,
                   rate: r.rate, amount: alwRowAmount(r) }))
    .filter((r) => r.name || r.amount > 0);
}
function renderAlw() {
  const body = $('alw-body');
  if (!body) return;
  const box = $('alw-box');
  if (box) box.hidden = !state.alwOn;
  const on = $('r-alw-on');
  if (on) on.checked = !!state.alwOn;
  const rows = state.alw || [];
  body.innerHTML = rows.length ? rows.map((r, i) => {
    const auto = alwRowAmount(r);
    const shown = r.manual ? (r.amount === '' || r.amount === undefined ? '' : r.amount)
                           : money2(auto);
    return `<tr>
      <td><input class="fi alw-in" data-k="name" data-i="${i}" value="${esc(r.name || '')}"
            placeholder="如 伙食补助费 / 市内交通费 / 杂费"></td>
      <td><input class="fi alw-in alw-num" data-k="people" data-i="${i}" value="${esc(r.people)}" placeholder="人数"></td>
      <td><input class="fi alw-in alw-num" data-k="days" data-i="${i}" value="${esc(r.days)}" placeholder="天数"></td>
      <td><input class="fi alw-in alw-num" data-k="rate" data-i="${i}" value="${esc(r.rate)}" placeholder="元/人·天"></td>
      <td><input class="fi alw-in alw-num" data-k="amount" data-i="${i}" value="${esc(shown)}"
            title="${r.manual ? '你手填的金额（改人数/天数/标准会自动恢复成自动计算）' : '自动＝人数×天数×标准；直接改这里也能覆盖'}"></td>
      <td class="alw-del"><button class="btn btn-sm btn-danger" data-act="alw-del" data-i="${i}"
            type="button" title="删掉这一行">×</button></td>
    </tr>`;
  }).join('') : '<tr><td colspan="6" class="empty" style="padding:8px 4px">还没有补助项，点下面「+ 加一行」</td></tr>';
  updateAlwSum();
}
function updateAlwSum() {
  const el = $('alw-sum');
  if (el) el.textContent = money(alwSum());
  const grand = $('alw-grand');
  if (grand) {
    const bill = (state.selRows || []).reduce((s, r) => s + num(r['价税合计']), 0);
    grand.innerHTML = state.alwOn && alwSum()
      ? `<b>${money(bill)}</b>（凭证）＋ <b>${money(alwSum())}</b>（补助）＝ 报销合计 <b>${money(bill + alwSum())}</b>`
      : '';
  }
}
function initAlw() {
  const on = $('r-alw-on');
  if (on) {
    on.addEventListener('change', () => {
      state.alwOn = on.checked;
      state.alwOff = !on.checked;
      if (state.alwOn && !state.alw.length) state.alw = [newAlwRow()];
      saveAlw();
      renderAlw();
    });
  }
  const add = $('btn-alw-add');
  if (add) add.addEventListener('click', () => addAlwRow());
  const body = $('alw-body');
  if (body) {
    body.addEventListener('input', (e) => {
      const inp = e.target.closest('input[data-k]');
      if (!inp) return;
      const r = state.alw[Number(inp.dataset.i)];
      if (!r) return;
      const k = inp.dataset.k;
      if (k === 'amount') {
        r.amount = inp.value;
        r.manual = true;
      } else {
        r[k] = inp.value;
        if (k !== 'name') r.manual = false;          // 改人数/天数/标准 → 回到自动算
      }
      saveAlw();
      // 只重算显示，不重建整张表（否则输入框会失焦）
      const idx = Number(inp.dataset.i);
      const amtInp = body.querySelector(`input[data-k="amount"][data-i="${idx}"]`);
      if (amtInp && k !== 'amount') amtInp.value = money2(alwRowAmount(r));
      if (amtInp) amtInp.title = r.manual ? '你手填的金额（改人数/天数/标准会自动恢复成自动计算）'
                                          : '自动＝人数×天数×标准；直接改这里也能覆盖';
      updateAlwSum();
    });
    body.addEventListener('click', (e) => {
      const btn = e.target.closest('button[data-act="alw-del"]');
      if (!btn) return;
      state.alw.splice(Number(btn.dataset.i), 1);
      saveAlw();
      renderAlw();
    });
  }
}

/* =====================================================================
 * 模板二版式（出差报销明细）：出差补助按人填
 * 票据合计是后端按「出行人」算好的（复用 report_data，附件不重复计钱），
 * 所以界面上看到的数和打印出来的一定一样。
 * ===================================================================== */
const JY_KIND = '模板二';

function loadJyAlw() {
  try {
    const raw = JSON.parse(localStorage.getItem(LS.jyAlw) || '{}');
    return raw && typeof raw === 'object' && !Array.isArray(raw) ? raw : {};
  } catch (e) {
    return {};
  }
}
function saveJyAlw() {
  try {
    localStorage.setItem(LS.jyAlw, JSON.stringify(state.jyAlw || {}));
  } catch (e) { /* 隐私模式下存不了，不影响使用 */ }
}
/** 名单 = 后端算出来的人 + 已经填过补助但名单里还没有的人（顺序：先后端给的） */
function jyNames() {
  const names = (state.jyPeople || []).map((p) => p.name);
  Object.keys(state.jyAlw || {}).forEach((n) => {
    if (n && names.indexOf(n) < 0 && num(state.jyAlw[n])) names.push(n);
  });
  return names;
}
function jyBills(name) {
  const p = (state.jyPeople || []).find((x) => x.name === name);
  return p ? num(p.bills) : 0;
}
function jyAmount(name) { return num((state.jyAlw || {})[name]); }
function jyReal(name) { return Math.round((jyBills(name) + jyAmount(name)) * 100) / 100; }
function jySum() {
  return Math.round(jyNames().reduce((s, n) => s + jyAmount(n), 0) * 100) / 100;
}
function jyPayload() {
  return jyNames().map((n) => ({ name: n, amount: jyAmount(n) }));
}
function updateJySum() {
  const sum = $('jy-sum');
  if (sum) sum.textContent = money(jySum());
  const grand = $('jy-grand');
  if (grand) {
    const bill = num(state.jyTotal);
    grand.innerHTML = jySum()
      ? `票据合计 <b>${money2(bill)}</b> ＋ 补助 <b>${money2(jySum())}</b> ＝ 合计 <b>${money2(bill + jySum())}</b>`
      : `票据合计 <b>${money2(bill)}</b>`;
  }
}
function renderJy() {
  const body = $('jy-body');
  if (!body) return;
  const names = jyNames();
  const box = $('jy-box');
  if (box) box.hidden = !names.length;
  body.innerHTML = names.length ? names.map((n) => {
    const amt = jyAmount(n);
    return `<tr>
      <td>${esc(n)}</td>
      <td class="td-r td-nowrap">${money2(jyBills(n))}</td>
      <td><input class="fi alw-num jy-in" data-name="${esc(n)}" value="${amt ? esc(amt) : ''}"
            placeholder="0.00" autocomplete="off"></td>
      <td class="td-r td-nowrap jy-real">${money2(jyReal(n))}</td>
    </tr>`;
  }).join('') : '<tr><td colspan="4" class="empty" style="padding:8px 4px">先在台账勾选凭证，这里会按出行人自动列出名单</td></tr>';
  updateJySum();
}
/** 向后端要「按出行人的票据合计」（跟生成的单据同一个口径） */
async function refreshJyPeople() {
  if (state.kind !== JY_KIND) return;
  const seqs = (state.selRows || []).map((r) => r['序号']);
  if (!seqs.length) { state.jyPeople = []; state.jyTotal = 0; renderJy(); return; }
  const r = await call('report_persons', seqs);
  if (!r || r.error) { renderJy(); return; }
  state.jyPeople = r.people || [];
  state.jyTotal = num(r.total);
  renderJy();
}
/** 一键按「天数 × 元/人·天」预填到每个人（省得一张一张打） */
function fillJyByDays() {
  const days = num($('jy-days').value);
  const rate = num($('jy-rate').value);
  if (!days || !rate) { toast('天数和标准都填上才能算', 'warn'); return; }
  const amt = Math.round(days * rate * 100) / 100;
  const names = jyNames();
  if (!names.length) { toast('先去台账勾选凭证', 'warn'); return; }
  names.forEach((n) => { state.jyAlw[n] = amt; });
  saveJyAlw();
  renderJy();
  toast(`已按 ${days} 天 × ${rate} 元/人·天 ＝ ${money(amt)} 预填到 ${names.length} 个人`, 'ok');
}
function initJy() {
  const body = $('jy-body');
  if (body) {
    body.addEventListener('input', (e) => {
      const inp = e.target.closest('input[data-name]');
      if (!inp) return;
      const n = inp.dataset.name;
      state.jyAlw[n] = inp.value;
      saveJyAlw();
      // 只刷这一行的「实际」和两个合计数，不重建整张表（否则输入框会失焦）
      const tr = inp.closest('tr');
      const real = tr && tr.querySelector('.jy-real');
      if (real) real.textContent = money2(jyReal(n));
      updateJySum();
    });
  }
  const btn = $('btn-jy-fill');
  if (btn) btn.addEventListener('click', fillJyByDays);
}

function reportMeta() {
  const meta = {
    kind: state.kind,
    person: $('r-person').value.trim(),
    dept: $('r-dept').value.trim(),
    reason: $('r-reason').value.trim(),
    note: $('r-note').value.trim(),
    date: $('r-date').value.trim() || todayCn(),
  };
  if (state.kind === '差旅费报销单') {
    meta.start = $('r-start').value.trim();
    meta.end = $('r-end').value.trim();
    meta.place = $('r-place').value.trim();
  }
  if (state.kind === JY_KIND) {
    // 模板二版式：补助是「按人」的。只带填了钱的人 —— 没填的不必在单据上挂一排 0
    // （那些人照旧会出现在补助块里，只是补助为 0，实际＝他自己的票据合计）
    const by = jyPayload().filter((x) => x.name && x.amount);
    if (by.length) meta.allowance_by = by;
  } else {
    const alw = alwPayload();        // 差旅费补助（勾了才有）
    if (alw.length) meta.allowance = alw;
  }
  return meta;
}

async function makeReport() {
  if (!state.selRows.length) { toast('先去「凭证台账」勾选凭证', 'warn'); return; }
  const seqs = state.selRows.map((r) => r['序号']);
  const meta = reportMeta();
  const fmt = state.fmt || 'pdf';
  const fmtLabel = { pdf: 'PDF', xlsx: 'Excel', both: 'PDF + Excel' }[fmt] || fmt;
  const btn = $('btn-make-report');
  btn.disabled = true;
  setStatus('生成报销单…', 'running');
  pushLogs([`开始生成《${meta.kind}》（${fmtLabel}）：${seqs.length} 张凭证`]);
  const r = await call('make_report', meta, seqs, fmt);
  btn.disabled = false;
  if (!r) { setStatus('生成失败', 'err'); return; }
  if (r.error) {
    setStatus('生成失败', 'err');
    toast('生成失败：' + r.error, 'err');
    return;
  }
  await loadOutputs();
  setStatus('报销单已生成', 'ok');
  toast(`已生成报销单（${fmtLabel}）：${r.count} 张，合计 ${money(r.sum)}`, 'ok');
  const extra = [];
  if (r.bill !== undefined && num(r.alw)) {
    const alwName = meta.kind === JY_KIND ? '出差补助' : '差旅费补助';
    extra.push(`凭证金额 <b>${money(r.bill)}</b> ＋ ${alwName} <b>${money(r.alw)}</b>`);
  }
  if (r.attach) {
    extra.push(`${r.attach} 张行程单作为发票的证明附件排在对应发票后面（金额已含在发票里，未重复计）`);
  }
  // 顺手问一句要不要标已报销（用户要求：不做自动改状态，改之前先问）
  showModal({
    title: '报销单已生成',
    html: `《${esc(meta.kind)}》已生成（${esc(fmtLabel)}），共 <b>${r.count}</b> 张，
           报销合计 <b>${money(r.sum)}</b>。<br>大写：${esc(r.upper)}<br>
           ${extra.length ? `<span style="color:var(--text-secondary);font-size:12px">${extra.join('；')}</span><br>` : ''}
           <span style="color:var(--text-secondary);font-size:12px">文件：${(r.outs || [r.out]).map((p) => esc(baseName(p))).join('　')}</span>`,
    note: '要把这几张凭证在台账里标成【已报销】吗？标了以后再入库同一张票会提示「重复报销风险」。',
    okText: `标为已报销（${r.count} 张）`,
    cancelText: '先不标',
    onOk: async () => {
      hideModal();
      const rr = await call('mark_rows', seqs, '已报销', '', meta.person);
      if (!rr) return;
      toast(`已把 ${rr.updated} 张标为已报销`, 'ok');
      state.selRows = [];
      renderReportRows();
      loadLedger();
    },
  });
}

async function doMerge(mode) {
  if (!state.selRows.length) { toast('先去「凭证台账」勾选凭证', 'warn'); return; }
  const seqs = state.selRows.map((r) => r['序号']);
  const tag = mode === '2up' ? 'A4 上下两张（中间裁开）' : '每页一张（原尺寸）';
  const btn = mode === '2up' ? $('btn-merge-2up') : $('btn-merge-plain');
  btn.disabled = true;
  setStatus('合并中…', 'running');
  pushLogs([`开始合并 ${seqs.length} 张发票（${tag}）`]);
  const r = await call('merge_invoices', seqs, mode);
  btn.disabled = false;
  if (!r) { setStatus('合并失败', 'err'); return; }
  if (r.error) {
    setStatus('合并失败', 'err');
    if (r.missing_list && r.missing_list.length) {
      showModal({
        title: '这些凭证没有 PDF，合并不了',
        html: `选中的 <b>${r.missing_list.length}</b> 张凭证都没有可合并的 PDF 版式文件：`
          + missListHtml(r.missing_list),
        note: '把 PDF 版式文件放到它们各自文件夹的同一层（名字随便起都行，程序会读票面里的发票号认票），再点一次合并。',
        noOk: true, cancelText: '知道了', wide: true,
      });
    } else {
      toast('合并失败：' + r.error, 'err');
    }
    return;
  }
  await loadOutputs();
  if (r.missing) {
    setStatus(`合并完成（少 ${r.missing} 张）`, 'warn');
    showModal({
      title: `合并好了，但有 ${r.missing} 张没进去`,
      html: `你选了 <b>${r.asked}</b> 张，其中 <b>${r.files}</b> 张已合并进
        <span class="name">${esc(r.name)}</span>（共 ${r.pages} 页 A4）。
        <br><br>下面这 <b>${r.missing}</b> 张没有 PDF 版式文件，所以没进去：`
        + missListHtml(r.missing_list),
      note: '想让它们也进合并件：把这几张的 PDF 版式文件放到它们各自文件夹的同一层'
        + '（文件名随便起，「张三-汉口-郑州东.pdf」这种也行，程序会读票面里的发票号认票），再点一次合并。',
      noOk: true, cancelText: '知道了', wide: true,
    });
    return;
  }
  setStatus('合并完成', 'ok');
  const where = mode === '2up' ? '同一张 A4 的上下两半' : '相邻两页';
  toast(`已合并 ${r.files} 张 → ${r.pages} 页 A4`
    + (r.linked ? `（${r.linked} 组发票+行程单在${where}）` : ''), 'ok');
  if (r.linked) {
    pushLogs([`🚕 ${r.linked} 组「发票 + 打车行程单」已排在一起（${where}），按顺序打印装订即可`]);
  }
}

function missListHtml(list) {
  if (!list || !list.length) return '';
  return '<ul class="miss-list">' + list.map((m) =>
    `<li><b>${esc(m.name)}</b>${m.no && m.no !== '（无号码）' ? `　<span class="id-na">${esc(m.no)}</span>` : ''}
      <br><span class="id-na">${esc(m.why)}</span></li>`).join('') + '</ul>';
}

function fmtSize(n) {
  if (!n) return '';
  if (n < 1024) return n + ' B';
  if (n < 1024 * 1024) return (n / 1024).toFixed(0) + ' KB';
  return (n / 1024 / 1024).toFixed(1) + ' MB';
}

function renderOutputs() {
  const box = $('out-list');
  const items = state.outputs || [];
  const hint = $('out-hint');
  if (hint) {
    const alive = items.filter((o) => o.exists).length;
    hint.textContent = items.length
      ? `共 ${items.length} 个（${alive} 个还在，「删除」送回收站还能还原）`
      : '「打开」直接看，「删除」送进回收站（还能还原）';
  }
  if (!items.length) {
    box.innerHTML = '<div class="empty" style="padding:14px 4px">还没有生成文件</div>';
    return;
  }
  box.innerHTML = items.map((o, i) => {
    const meta = [o.kind || '文件',
      o.count ? `${o.count} 条` : '',
      o.missing ? `少 ${o.missing} 张` : '',
      o.pages ? `${o.pages} 页` : '',
      o.sum ? money(o.sum) : '',
      fmtSize(o.size), o.time || ''].filter(Boolean).join('　·　');
    return `
    <div class="out-item${o.exists ? '' : ' out-gone'}">
      <span class="oi-kind">${esc(o.kind || '文件')}</span>
      <span class="oi-name" title="${esc(o.path)}">${esc(o.name)}
        ${o.exists ? '' : '<span class="tag tag-mute">已不在</span>'}</span>
      <span class="oi-time" title="${esc(meta)}">${esc(o.time || '')}</span>
      <button class="btn btn-sm" data-act="open" data-i="${i}"${o.exists ? '' : ' disabled'}>打开</button>
      <button class="btn btn-sm btn-ghost" data-act="reveal" data-i="${i}"${o.exists ? '' : ' disabled'}>定位</button>
      <button class="btn btn-sm btn-danger" data-act="del" data-i="${i}">删除</button>
    </div>`;
  }).join('');
}

/* =====================================================================
 * 四、统计汇总
 * ===================================================================== */
function initStatsPage() {
  $('btn-export2').addEventListener('click', () => doExport(currentFilter()));
}

async function loadStats() {
  const f = currentFilter();
  const r = await call('get_ledger', f);
  if (!r || r.error) {
    $('st-stats').innerHTML = '';
    $('st-tables').innerHTML =
      `<div class="card"><div class="empty">读取台账失败${r && r.error ? '：' + esc(r.error) : ''}</div></div>`;
    return;
  }
  state.aggregates = r.aggregates || null;
  const filtered = ['status', 'ctype', 'category', 'batch', 'person', 'month',
                    'min', 'max', 'kw'].some((k) => f[k]);
  state.aggScope = filtered ? '按台账页当前筛选条件' : '全部台账';
  const boxes = [
    ['台账条数', r.all_count, 'accent'],
    ['统计条数', r.total, ''],
    ['价税合计', money(r.sum), ''],
    ['待报销', `${r.pending_count} 条`, 'warn'],
    ['待报销金额', money(r.pending_sum), 'ok'],
    ['已报销金额', money(num(r.sum) - num(r.pending_sum)), 'mute'],
  ];
  $('st-stats').innerHTML = boxes.map(([lab, val, cls]) =>
    `<div class="stat-box ${cls}"><div class="sb-num">${esc(val)}</div>
     <div class="sb-lab">${esc(lab)}</div></div>`).join('')
    + `<div class="stat-box mute"><div class="sb-num" style="font-size:13px">${esc(state.aggScope)}</div>
       <div class="sb-lab">统计口径</div></div>`;
  renderAggTables(r.aggregates || {});
}

function renderAggTables(agg) {
  const defs = [['按月', '月份'], ['按费用类别', '费用类别'], ['按销方', '销方'],
                ['按项目', '项目 / 事由'], ['按状态', '状态']];
  $('st-tables').innerHTML = defs.map(([key, title]) => {
    const data = agg[key] || {};
    const keys = Object.keys(data).sort((a, b) => data[b]['价税合计'] - data[a]['价税合计']);
    if (!keys.length) return '';
    let sumN = 0, sumT = 0;
    const rows = keys.map((k) => {
      const v = data[k];
      sumN += v['张数'];
      sumT += num(v['价税合计']);
      return `<tr><td class="name" title="${esc(k)}">${esc(k)}</td>
              <td class="num">${v['张数']}</td>
              <td class="num">${money2(v['价税合计'])}</td></tr>`;
    }).join('');
    return `<div class="card">
      <div class="card-head"><div class="card-title">${esc(title)}</div>
        <div class="card-hint">${keys.length} 类</div></div>
      <div class="table-wrap" style="max-height:288px;overflow:auto">
        <table class="agg-table">
          <thead><tr><th>${esc(title)}</th><th class="num" style="width:56px">条数</th>
          <th class="num" style="width:104px">价税合计</th></tr></thead>
          <tbody>${rows}
            <tr class="total-row"><td>合计</td><td class="num">${sumN}</td>
            <td class="num">${money2(sumT)}</td></tr>
          </tbody>
        </table>
      </div></div>`;
  }).join('') || `<div class="card"><div class="empty">台账里还没有数据，先去「凭证入库」建账</div></div>`;
}

async function doExport(filt) {
  const r = await call('export_summary', filt || {});
  if (!r) return;
  if (r.error) { toast('导出失败：' + r.error, 'err'); return; }
  toast(`已导出统计 Excel（${r.rows} 行）`, 'ok');
  await loadOutputs();
}

/* =====================================================================
 * 五、设置
 * ===================================================================== */
function initSetPage() {
  $('btn-save-cfg').addEventListener('click', async () => {
    const cats = $('cfg-cats').value.split(/[,，;；]/).map((s) => s.trim()).filter(Boolean);
    const cfg = await call('set_config', {
      people: $('cfg-person').value.trim(),
      dept: $('cfg-dept').value.trim(),
      categories: cats,
    });
    if (!cfg) return;
    state.cfg = cfg;
    state.cats = cfg.categories || [];
    fillCategorySelects();
    toast('设置已保存', 'ok');
  });
  // 出行人手机号对照：每行「手机号 姓名」
  $('btn-reload-pmap').addEventListener('click', loadPhoneMap);
  $('btn-save-pmap').addEventListener('click', savePhoneMap);
}

/** 把「手机号 → 姓名」对照读进设置页的输入框 */
async function loadPhoneMap() {
  const r = await call('get_phone_map');
  if (!r || r.error) return;
  const lines = (r.map || []).map((x) => `${x.phone} ${x.name || ''}`.trim());
  $('phone-map').value = lines.join('\n');
}

/** 存对照表：先在本地解析成 手机号/姓名，再把改动一条条同步到后端 */
async function savePhoneMap() {
  const want = {};
  String($('phone-map').value || '').split('\n').forEach((ln) => {
    const t = ln.replace(/[=,，\t]+/g, ' ').trim();
    if (!t) return;
    const m = t.match(/^(\d{6,11})\s*(.*)$/);
    if (m) want[m[1]] = m[2].trim();
  });
  const now = await call('get_phone_map');
  if (!now || now.error) { toast('读取对照表失败', 'err'); return; }
  const have = {};
  (now.map || []).forEach((x) => { if (x.name) have[x.phone] = x.name; });
  let n = 0;
  // 新增 / 改名
  for (const [ph, nm] of Object.entries(want)) {
    if (nm && have[ph] !== nm) { await call('set_phone_name', ph, nm); n += 1; }
  }
  // 删掉的
  for (const ph of Object.keys(have)) {
    if (!(ph in want)) { await call('set_phone_name', ph, ''); n += 1; }
  }
  toast(n ? `对照表已更新 ${n} 条` : '对照表没有变化', n ? 'ok' : 'warn');
  await loadPhoneMap();
}

function fillCategorySelects() {
  fillSelect('set-category', state.cats, $('set-category').value, '选类别');
}

function loadConfigIntoForm() {
  const a = api();
  if (!a) return;
  a.get_config().then((cfg) => {
    if (!cfg) return;
    state.cfg = cfg;
    state.cats = cfg.categories || [];
    state.defaultSrc = cfg.source || state.defaultSrc;
    state.cfgSrc = cfg.source || '';
    state.recursive = cfg.recursive !== false;
    syncScopePills();
    $('cfg-person').value = cfg.people || '';
    $('cfg-dept').value = cfg.dept || '';
    $('cfg-cats').value = state.cats.join('，');
    if (!$('r-person').value) $('r-person').value = cfg.people || '';
    if (!$('r-dept').value) $('r-dept').value = cfg.dept || '';
    // 入库表单的「本批报销人」也预填上配置里的人 —— 出行人拆成单独一列后，
    // 报销人不再自动顶替，空着就是空着，所以这里给个默认更省事
    if (!$('in-person').value) $('in-person').value = cfg.people || '';
    fillCategorySelects();
  }).catch(() => {});
}

function renderPaths(p) {
  if (!p) return;
  const rows = [
    ['发票文件夹', p.source || '（未设置）'],
    ['台账文件', p.ledger + (p.ledger_exists ? '' : '（还没生成）')],
    ['输出目录', p.output],
    ['程序目录', p.app],
  ];
  $('path-info').innerHTML = rows.map(([k, v]) =>
    `<div class="info-row"><span class="info-label">${esc(k)}</span>
     <span class="info-value" title="${esc(v)}">${esc(v)}</span></div>`).join('');
}

/* =====================================================================
 * 六、轮询后端
 * ===================================================================== */
let lastSeq = 0;
let lastResultId = 0;
let firstPoll = true;
let quitMode = false;

async function poll() {
  if (quitMode) return;
  const a = api();
  if (!a) return;
  let r;
  try { r = await a.poll(lastSeq); } catch (e) { return; }
  if (!r) return;

  if (r.logs && r.logs.length) {
    lastSeq = r.seq;
    pushLogs(r.logs);
  }
  if (r.status) setStatus(r.status.text, r.status.kind);
  if (r.progress) {
    const p = r.progress;
    setProgress(p.on, p.cur, p.total, p.label);
  }
  if (r.busy !== undefined && r.busy !== state.busy) setImportBusy(r.busy);
  // 结果按 result_id 去重；本次会话第一次拿到的（多半是刷新页面时带出来的旧结果）不弹提示
  if (r.result && r.result_id !== lastResultId) {
    lastResultId = r.result_id;
    onImportResult(r.result, firstPoll);
  }
  firstPoll = false;
}
setInterval(poll, 400);

/* =====================================================================
 * 窗口 / 退出
 * ===================================================================== */
function initWindow() {
  // 界面在系统默认浏览器里，窗口的最小化/最大化/关闭交给浏览器自己；
  // 这里只有「退出程序」需要管（= 关掉本地服务）。
  $('btn-close').addEventListener('click', doQuit);
  $('btn-quit').addEventListener('click', doQuit);
}

/* 退出 = 关掉本地服务。浏览器不让网页关掉用户自己开的标签页，
   所以这里只结束后端，剩下那个标签页手动关掉即可（服务停了它就是张静态页面）。 */
async function doQuit() {
  const a = api();
  if (!a) return;
  if (!confirm('结束程序？\n\n本地服务会关闭，界面随即停止工作；然后直接关掉这个标签页即可。')) return;
  try { await a.shutdown(); } catch (e) { /* 服务关闭时请求会断，属正常 */ }
  quitMode = true;
  document.body.innerHTML =
    '<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;'
    + 'height:100vh;gap:10px">'
    + '<div style="font-size:15px;font-weight:500;color:var(--text-primary)">程序已退出</div>'
    + '<div style="font-size:13px;color:var(--text-secondary)">本地服务已关闭，现在可以关掉这个标签页了</div>'
    + '</div>';
}

/* =====================================================================
 * 启动
 * ===================================================================== */
/* =====================================================================
 * 账号与权限 / 台账数据库（设置页）
 * ---------------------------------------------------------------------
 * 界面藏起按钮只是「不显眼」，真正的卡口在后端（每个接口都会再查一次权限）。
 * 这里拿到 whoami 的 can 表来决定显示哪些入口。
 * ===================================================================== */
let meInfo = { user: '', role: '', role_cn: '', can: {}, auth: false };

/* ---------------------------------------------------------------------
 * 侧栏底部：当前登录是谁 + 退出登录
 * ---------------------------------------------------------------------
 * 后端 /logout 早就在了（会真的把服务端会话删掉，不只是让浏览器丢 Cookie），
 * 只是界面上一直没挂出来 —— 结果「多人」在界面上完全看不见。这里补上。
 * 本机模式（没设访问口令）没有登录这回事，整块藏掉，不摆假身份。
 * ------------------------------------------------------------------- */
function initMe() {
  const btn = $('btn-logout');
  if (btn) {
    btn.addEventListener('click', () => { window.location.href = '/logout'; });
  }
  renderMe();
}

async function renderMe() {
  const box = $('sb-user');
  if (!box) return;
  const r = await call('whoami');
  if (r && !r.error) meInfo = r;
  if (!meInfo.auth) { box.hidden = true; return; }
  box.hidden = false;
  $('sb-user-name').textContent = meInfo.user || '未登录';
  $('sb-user-role').textContent = meInfo.role_cn || meInfo.role || '';
  box.title = `当前登录：${meInfo.user || ''}（${meInfo.role_cn || ''}）`;
}

async function loadSettingsExtras() {
  await renderMe();
  initSettingsEvents();
  await renderDbInfo();
  await renderUsers();
}

function initSettingsEvents() {
  const card = $('card-users');
  if (card.dataset.bound) return;
  card.dataset.bound = '1';
  $('btn-add-user').addEventListener('click', addUser);
  $('user-list').addEventListener('click', onUserOp);
  $('user-list').addEventListener('change', onUserOp);
  $('btn-db-refresh').addEventListener('click', async () => {
    await renderDbInfo();
    await renderUsers();
    toast('已刷新', 'ok');
  });
  $('btn-import-excel').addEventListener('click', importExcel);
  $('btn-audit').addEventListener('click', showAudit);
}

async function renderUsers() {
  const card = $('card-users');
  if (!meInfo.can || !meInfo.can.list_users) { card.style.display = 'none'; return; }
  const r = await call('list_users');
  if (!r || r.error) { card.style.display = 'none'; return; }
  card.style.display = '';
  const roles = r.roles || {};
  const roleOpts = (cur) => Object.entries(roles).map(([k, v]) =>
    `<option value="${k}"${k === cur ? ' selected' : ''}>${esc(v)}</option>`).join('');
  const sel = $('nu-role');
  if (!sel.dataset.done) { sel.innerHTML = roleOpts('biz'); sel.dataset.done = '1'; }

  const rows = (r.rows || []).map((u) => `
    <div class="usr-row${u.disabled ? ' is-off' : ''}">
      <div class="usr-line">
        <b>${esc(u.username)}</b>
        <span class="usr-tag">${esc(u.role_cn || u.role)}</span>
        ${u.disabled ? '<span class="usr-tag warn">已停用</span>' : ''}
        ${u.username === meInfo.user ? '<span class="usr-tag me">当前登录</span>' : ''}
        <span class="usr-meta">${esc(u.display_name || '')}${u.dept ? ' · ' + esc(u.dept) : ''}${
          u.last_login ? ' · 上次登录 ' + esc(u.last_login) : ' · 还没登录过'}</span>
      </div>
      <div class="usr-ops">
        <select class="fi" data-u="${esc(u.username)}" data-act="role">${roleOpts(u.role)}</select>
        <button class="btn btn-sm" data-u="${esc(u.username)}" data-act="pw">改密码</button>
        <button class="btn btn-sm" data-u="${esc(u.username)}" data-act="off">${
          u.disabled ? '启用' : '停用'}</button>
        <button class="btn btn-sm" data-u="${esc(u.username)}" data-act="kick">强制退出</button>
      </div>
    </div>`).join('');
  $('user-list').innerHTML = rows
    || '<div class="card-hint" style="display:block">还没有别的账号，用下面那行添一个</div>';
}

async function onUserOp(ev) {
  const el = ev.target.closest('[data-u]');
  if (!el) return;
  // ⚠️ 下拉框只认 change：点开下拉框本身也会冒泡一个 click 上来，
  //    要是把那个 click 当成「改角色」，就会拿旧值去写库、紧接着 renderUsers() 把
  //    整个列表 innerHTML 换掉 —— 表现就是下拉框一闪就没了、根本选不中（用户报的"一直弹跳"）。
  if (el.tagName === 'SELECT') {
    if (ev.type !== 'change') return;
  } else if (ev.type !== 'click') {
    return;
  }
  const u = el.dataset.u;
  const act = el.dataset.act;
  let r = null;
  if (act === 'role') {
    r = await call('set_user_role', u, el.value);
    if (r && r.error) return toast(r.error, 'err');
    toast(`${u} 的角色已改成「${el.options[el.selectedIndex].text}」`, 'ok');
    return renderUsers();
  }
  if (act === 'pw') {
    const pw = prompt(`给「${u}」设置新密码（至少 6 位）。\n改完他的旧登录会立刻失效，需要重新登录。`);
    if (!pw) return;
    r = await call('set_user_password', u, pw);
    if (r && r.error) return toast(r.error, 'err');
    toast(`已重置 ${u} 的密码`, 'ok');
    return renderUsers();
  }
  if (act === 'off') {
    const off = el.textContent.trim() === '停用';
    r = await call('set_user_disabled', u, off);
    if (r && r.error) return toast(r.error, 'err');
    toast(`${u} 已${off ? '停用' : '启用'}`, 'ok');
    return renderUsers();
  }
  if (act === 'kick') {
    r = await call('force_logout', u);
    if (r && r.error) return toast(r.error, 'err');
    toast(`已让 ${u} 退出（${r.kicked} 个会话）`, 'ok');
    return renderUsers();
  }
}

async function addUser() {
  const name = $('nu-name').value.trim();
  const disp = $('nu-display').value.trim();
  const dept = $('nu-dept') ? $('nu-dept').value.trim() : '';
  const pw = $('nu-pw').value;
  if (!name || !pw) return toast('用户名和初始密码都要填', 'warn');
  const r = await call('create_user', name, pw, $('nu-role').value, disp, dept);
  if (r && r.error) return toast(r.error, 'err');
  $('nu-name').value = ''; $('nu-display').value = ''; $('nu-pw').value = '';
  if ($('nu-dept')) $('nu-dept').value = '';
  toast(`已新增账号 ${name}`, 'ok');
  renderUsers();
}

async function renderDbInfo() {
  const box = $('db-info');
  // 「从 Excel 导入」和「审计日志」是管理动作：后端只放行管理员，
  // 界面上就也别露出来（露出来点一下弹「权限不足」最让人困惑）。
  if (meInfo.can) {
    $('btn-import-excel').style.display = meInfo.can.import_from_excel ? '' : 'none';
    $('btn-audit').style.display = meInfo.can.list_users ? '' : 'none';
  }
  const r = await call('ledger_db_info');
  if (!r || r.error) {
    box.innerHTML = `<div class="card-hint" style="display:block">读不到数据库信息：${
      esc((r && r.error) || '未知')}</div>`;
    return;
  }
  const kb = ((r.size || 0) / 1024).toFixed(1);
  const kv = (k, v, cls) => `<div class="kv"><span>${k}</span><b class="${cls || ''}">${v}</b></div>`;
  box.innerHTML = `<div class="kv-grid">
    ${kv('数据库文件', esc(r.db || ''))}
    ${kv('台账行数', r.invoices)}
    ${kv('已删除（留痕可追溯）', r.deleted)}
    ${kv('审计记录', r.audit)}
    ${kv('账号数', r.users)}
    ${kv('占用空间', kb + ' KB')}
    ${kv('唯一索引', r.unique_index_ok ? '正常' : '已降级（台账里有重复号码）',
         r.unique_index_ok ? '' : 'warn')}
    ${kv('Excel 快照', r.excel_exists ? '有' : '还没有')}
    ${kv('迁移自 Excel', r.migrated_at ? esc(r.migrated_at) : '——')}
  </div>`;
}

function importExcel() {
  showModal({
    title: '从 Excel 导入台账',
    html: '把 <b>发票台账.xlsx</b> 的内容读进数据库。<b>当前台账会被整表替换</b>，'
        + '导入前程序会自动把原文件备份到「台账备份」里。',
    note: '只有你在 Excel 里改过东西、想把改动并回系统时才需要这么做。平时的改动请在界面上做。',
    okText: '确认导入',
    onOk: async () => {
      hideModal();
      const r = await call('import_from_excel', '');
      if (!r || r.error) return toast('导入失败：' + ((r && r.error) || '未知'), 'err');
      toast(`已导入 ${r.migrated} 行（导入前 ${r.before} 行）`, 'ok');
      renderDbInfo();
    },
  });
}

async function showAudit() {
  const r = await call('audit_list', 150);
  if (!r || r.error) return toast('读审计日志失败：' + ((r && r.error) || '未知'), 'err');
  const rows = (r.rows || []).map((x) => `<tr class="${x.ok ? '' : 'bad'}">
      <td class="nowrap">${esc(x.ts)}</td>
      <td>${esc(x.username || '—')}</td>
      <td class="nowrap">${esc(x.ip || '—')}</td>
      <td>${esc(x.action)}</td>
      <td>${esc(x.object_id || '')}</td>
      <td>${esc((x.before_json || '').slice(0, 80))}</td>
      <td>${esc((x.after_json || '').slice(0, 80))}</td>
    </tr>`).join('');
  showModal({
    title: `审计日志（显示 ${(r.rows || []).length} 条，共 ${r.total} 条）`,
    xwide: true, noOk: true, cancelText: '关闭',
    html: `<div class="audit-wrap"><table class="audit-tbl"><thead><tr>
        <th>时间</th><th>谁</th><th>来源 IP</th><th>动作</th><th>对象</th><th>改前</th><th>改后</th>
      </tr></thead><tbody>${rows || '<tr><td colspan="7">还没有记录</td></tr>'}</tbody></table></div>`,
  });
}

function boot() {
  // 界面跑在系统默认浏览器里，自绘标题栏没意义（浏览器自己有），隐藏掉
  document.body.classList.add('native-frame');

  applyTheme(localStorage.getItem(LS.theme) || 'dark');
  renderThemes();
  initThemePop();
  initNav();
  initImportPage();
  initLedgerPage();
  initReportPage();
  initStatsPage();
  initSetPage();
  initLogPanel();
  initWindow();
  initMe();
  syncKindPills();
  syncFmtPills();
  syncScopePills();
  renderReportRows();
  loadOutputs();

  $('modal-cancel').addEventListener('click', hideModal);
  $('modal-ok').addEventListener('click', () => {
    const fn = modalOk;
    if (fn) fn();
    else hideModal();
  });
  $('overlay').addEventListener('click', (e) => { if (e.target === $('overlay')) hideModal(); });
  // 凭证明细弹层里的「打开文件 / 在文件夹中显示 / 删除这一条」（弹层内容是动态塞的，用委托）
  $('modal-text').addEventListener('click', (e) => {
    const btn = e.target.closest('[data-act]');
    if (!btn) return;
    const act = btn.dataset.act;
    if (act === 'del_row') { hideModal(); deleteOneRow(btn.dataset.seq); return; }
    const p = btn.dataset.path || '';
    if (p) call(act, p);
  });

  pushLogs(['界面已就绪。放发票的文件夹 + 点「开始入库」就能建账。']);

  // 深链：?pick=pending 直接把台账里所有未报销发票带进「报销单」页
  // （可以把它存成书签，一键给这批票出单据）
  if (new URLSearchParams(location.search).get('pick') === 'pending') {
    const a0 = api();
    if (a0) {
      a0.get_ledger({ status: '未报销' }).then((r) => {
        if (!r || !r.rows || !r.rows.length) return;
        state.selRows = r.rows;
        renderReportRows();
        $('rep-hint').textContent = `共 ${r.rows.length} 张（已带入全部未报销）`;
        toast(`已带入全部未报销发票 ${r.rows.length} 张`, 'ok');
      }).catch(() => {});
    }
  }

  // 支持 #page-ledger 这种锚点直达（刷新后停在原页，也方便做书签）
  const hash = (location.hash || '').replace('#', '');
  if (hash && $(hash) && $(hash).classList.contains('page')) {
    switchPage(hash);
  }
  window.addEventListener('hashchange', () => {
    const h = (location.hash || '').replace('#', '');
    if (h && $(h) && $(h).classList.contains('page')) switchPage(h);
  });

  loadConfigIntoForm();
  const a = api();
  if (a && a.get_paths) {
    a.get_paths().then((p) => {
      renderPaths(p);
      if (!p) return;
      state.defaultSrc = p.source || '';
      state.cfgSrc = p.source || '';
      state.src = localStorage.getItem(LS.lastSrc) || p.source || '';
      $('src-path').value = state.src;
      const savedRec = localStorage.getItem('invoice-rp.recursive');
      if (savedRec !== null) state.recursive = savedRec === '1';
      syncScopePills();
      checkSource();
      // 台账已有数据就直接显示出来，省得每次都要点一下
      if (p.ledger_exists) loadLedger();
    }).catch(() => {});
  }
}

window.addEventListener('pywebviewready', boot);
if (window.pywebview) boot();
