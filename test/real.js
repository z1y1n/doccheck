/* 拿真实文件渲染一遍，数一数用户到底会看到几行。
   文件不在就跳过，不报错——这是给真人看的体检报告，不是回归测试。 */
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

const ROOT = path.join(__dirname, '..');
const files = process.argv.slice(2);
if (!files.length) {
  console.error('用法: node test/real.js <文件1.docx> [文件2.docx ...]');
  process.exit(1);
}

const dom = new JSDOM(fs.readFileSync(path.join(ROOT, 'src', 'index.html'), 'utf-8'), {
  url: 'http://localhost/', runScripts: 'outside-only', pretendToBeVisual: true
});
const { window } = dom;
const bundle = ['engine.js', 'rules.js', 'ui.js']
  .map(f => fs.readFileSync(path.join(ROOT, 'src', f), 'utf-8')).join('\n;\n');
window.eval(bundle + '\n;window.DocEngine = DocEngine; window.RULES = RULES;');
// jszip 在 eval 之后再挂：写进 eval 串里的话那一刻它还不存在
window.JSZip = require('jszip');
window.URL.createObjectURL = () => 'blob:fake';
window.URL.revokeObjectURL = () => {};
window.HTMLAnchorElement.prototype.click = function () {};

const DC = window.DocCheck;
const doc = window.document;

(async () => {
  const out = [];
  out.push('文件'.padEnd(34) + '问题  报告里看得见的行  收起来的行');
  out.push('-'.repeat(78));

  for (const f of files) {
    const p = path.isAbsolute(f) ? f : path.join(ROOT, f);
    if (!fs.existsSync(p)) { out.push(path.basename(f).padEnd(34) + '（文件不存在，跳过）'); continue; }
    const buf = fs.readFileSync(p);
    await DC.openFile({
      name: path.basename(p),
      arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
      text: async () => buf.toString('utf-8')
    });

    const groups = [...doc.querySelectorAll('#report .group')];
    let visible = 0, hidden = 0;
    for (const g of groups) {
      const n = g.querySelectorAll('.body .item').length;
      if (g.classList.contains('collapsed')) hidden += n;
      else visible += n;
    }
    const total = DC.state.issues.length;
    const name = path.basename(p);
    out.push(
      (name.length > 32 ? name.slice(0, 31) + '…' : name).padEnd(34) +
      String(total).padStart(4) + '  ' +
      String(visible).padStart(17) + '  ' +
      String(hidden).padStart(13)
    );

    // 说清楚收起来的是哪一类
    for (const g of groups) {
      if (!g.classList.contains('collapsed')) continue;
      const n = g.querySelectorAll('.body .item').length;
      if (n) out.push('      └ 收起来的是「' + g.querySelector('.title').textContent + '」' + n + ' 处');
    }
  }

  // 全部展开做对比
  out.push('');
  out.push('把收起来的都点开之后：');
  doc.querySelectorAll('#report .group.collapsed .head').forEach(h =>
    h.dispatchEvent(new window.Event('click', { bubbles: true })));
  const vis = [...doc.querySelectorAll('#report .body .item')]
    .filter(el => el.offsetParent !== null || !el.closest('.group').classList.contains('collapsed')).length;
  out.push('  当前这份文件展开状态下共 ' + vis + ' 行');

  fs.writeFileSync(path.join(ROOT, '_real.txt'), out.join('\n'), 'utf-8');
  console.log(out.join('\n'));
})();
