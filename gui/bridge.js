/* ------------------------------------------------------------------
 * bridge.js —— 用本地 HTTP 后端模拟 pywebview 的 JS 桥
 *
 * 背景：pywebview 依赖的 WebView2 宿主控件在本机不稳定（窗口黑屏），
 *      所以本项目直接起一个只监听 127.0.0.1 的小 HTTP 服务，
 *      再用系统默认浏览器打开界面（见 app_web.py）。
 *      本文件让 gui/app.js 无需任何改动即可照常工作。
 *
 * 约定：app.js 通过 window.pywebview.api.xxx(...) 调用后端，
 *      这里把它映射为 POST /api/xxx  {args:[...]}  →  {result:...}
 *
 * 注意：METHODS 只是「显式列出」方便阅读，真正的转发靠下面那个 Proxy 兜底，
 *      所以后端新增接口时这里不写也能用（但建议同步一下，方便对照）。
 * ------------------------------------------------------------------ */
(function () {
  'use strict';

  var METHODS = ['poll', 'get_paths', 'get_config', 'set_config', 'pick_folder',
                 'check_source', 'upload_info', 'start_import', 'get_ledger', 'mark_rows',
                 'update_rows', 'make_report', 'merge_invoices', 'export_summary',
                 'ledger_db_info', 'import_from_excel', 'audit_list', 'list_jobs', 'gen_files',
                 'list_users', 'create_user', 'set_user_password', 'set_user_role',
                 'set_user_disabled', 'force_logout',
                 'open_folder', 'open_file', 'reveal', 'shutdown'];
  // 注意：上传文件不走这里（raw body 塞不进 {args:[...]} 这个壳），
  //      app.js 直接 POST /api/upload_raw。除它以外的接口都在上面。

  function call(name) {
    var args = Array.prototype.slice.call(arguments, 1);
    return fetch('/api/' + encodeURIComponent(name), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ args: args }),
      cache: 'no-store'
    }).then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.json();
    }).then(function (j) { return j.result; });
  }

  var apiObj = {};
  METHODS.forEach(function (m) {
    apiObj[m] = function () { return call.apply(null, [m].concat(Array.prototype.slice.call(arguments))); };
  });

  // 服务器模式（Docker/NAS）下后端开不了资源管理器 —— 它会回一个 {url}，
  // 这里替它开个新标签页。本机模式后端不返回 url，所以完全不受影响，
  // app.js 那边一个字都不用改。
  ['open_file', 'open_folder', 'reveal'].forEach(function (m) {
    var base = apiObj[m];
    apiObj[m] = function () {
      return base.apply(null, arguments).then(function (r) {
        if (r && r.url) { try { window.open(r.url, '_blank'); } catch (e) { /* 忽略 */ } }
        return r;
      });
    };
  });

  // 兜底：任何未列出的方法也照样转发（方便以后加功能不用改这里）
  apiObj = new Proxy(apiObj, {
    get: function (t, k) {
      if (typeof k !== 'string') return t[k];
      if (k in t) return t[k];
      return function () { return call.apply(null, [k].concat(Array.prototype.slice.call(arguments))); };
    }
  });

  function define() {
    window.pywebview = { api: apiObj };
    try { window.dispatchEvent(new Event('pywebviewready')); } catch (e) { /* 忽略 */ }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', define);
  } else {
    define();
  }
})();
