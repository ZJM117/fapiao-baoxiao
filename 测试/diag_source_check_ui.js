/* 文件夹扫描提示回归测试：验证无服务、扫描异常和空目录都有可见状态。 */
const fs = require('fs');
const path = require('path');

const source = fs.readFileSync(path.join(__dirname, '..', 'gui', 'app.js'), 'utf8');
const elements = {};
const element = (id) => elements[id] || (elements[id] = {
  id, value: '', textContent: '', innerHTML: '', className: '',
  style: {}, dataset: {}, classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
  addEventListener() {}, setAttribute() {}, removeAttribute() {},
});

global.window = { addEventListener() {}, pywebview: null, location: { search: '', hash: '' } };
global.document = {
  getElementById: element,
  querySelector() { return element('query'); },
  querySelectorAll() { return []; },
  addEventListener() {},
};
global.localStorage = { getItem() { return null; }, setItem() {}, removeItem() {} };

const app = new Function(source + '\n; return {checkSource, state};')();
const note = element('src-note');
const wait = (ms = 420) => new Promise((resolve) => setTimeout(resolve, ms));
const check = (ok, label) => {
  console.log(`${ok ? '  ok   ' : '  FAIL '}${label}`);
  if (!ok) process.exitCode = 1;
};

(async () => {
  app.state.src = 'C:\\票据';
  app.checkSource();
  await wait();
  check(/本地服务未连接/.test(note.textContent), '后端未连接时给出明确提示');

  window.pywebview = { api: {
    check_source: async () => { throw new Error('PermissionError: access denied'); },
    set_config: async () => true,
  } };
  app.checkSource();
  await wait();
  check(/扫描失败/.test(note.textContent) && /PermissionError/.test(note.textContent),
    '接口异常显示具体错误');

  window.pywebview.api.check_source = async () => ({
    ok: true, files: 0, kinds: {}, msg: '没找到可识别的凭证文件',
  });
  app.checkSource();
  await wait();
  check(note.className === 'src-note warn' && /没找到可识别的凭证文件/.test(note.textContent),
    '没有可识别文件时保留后端说明并显示警告');

  check(/选择失败：\$\{call\.lastError\}/.test(source)
    && /已取消选择，当前路径未更改/.test(source),
  '文件夹选择异常与用户取消显示不同提示');
  check(/pickButton\.disabled = true/.test(source)
    && /pickButton\.disabled = false/.test(source),
  '等待系统选择器时禁用按钮，避免重复排队');

  const status = process.exitCode || 0;
  console.log(status ? '\n❌ 有测试失败' : '\n✅ 全部通过');
  process.exit(status);
})();
