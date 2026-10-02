/* 测 Markdown 路径：代码块、行内代码、链接地址都必须原样不动 */
const fs = require('fs');
const path = require('path');
const DocEngine = require('../src/engine.js');
const RULES = require('../src/rules.js');

const SAMPLE = [
  '# 产品经理转行指南',
  '',
  '这份指南写给准备2026届秋招的同学，重点讲PRD怎么写在简历里。',
  '',
  '## 一、工具准备',
  '',
  '推荐用QGIS处理地图数据，配合SQL做基础查询。',
  '',
  '```python',
  '# 这段代码里的注释不该被检查',
  'name="中文English混排"   # 网站 https://example.com/a,b.c',
  'print(f"结果:{name}")',
  '```',
  '',
  '行内代码也不能动：`git commit -m "修复bug"`，还有 `SELECT * FROM t WHERE a=1`。',
  '',
  '参考链接：[官方文档中文版](https://example.com/docs,cn/index.html) 以及 https://github.com/foo/bar.baz',
  '',
  '联系方式: zhangsan@163.com，也可以用 13800000000。',
  '',
  '| 工具 | 用途 |',
  '| --- | --- |',
  '| Excel | 数据整理 |',
  '',
].join('\n');

const OUT = [];
const say = s => OUT.push(s);

let pass = 0, fail = 0;
function check(name, cond, extra) {
  if (cond) { pass++; say(`  ✓ ${name}`); }
  else { fail++; say(`  ✗ ${name}${extra ? '　' + extra : ''}`); }
}

const doc = DocEngine.parseMarkdown(SAMPLE);
const issues = DocEngine.runRules(doc, RULES);

say(`解析出 ${doc.paragraphs.length} 段，查出 ${issues.length} 处问题\n`);

say('【逐条问题】');
const byRule = {};
for (const it of issues) (byRule[it.ruleId] ||= []).push(it);
for (const r of RULES) {
  const list = byRule[r.id] || [];
  if (!list.length) continue;
  say(`  ${r.name}（${list.length}）`);
  for (const it of list) {
    const p = doc.paragraphs[it.pi];
    const a = Math.max(0, it.start - 18), b = Math.min(p.text.length, it.end + 18);
    say(`    第${it.pi + 1}段 …${p.text.slice(a, it.start)}〖${it.original}〗${p.text.slice(it.end, b)}…`);
  }
}
say('');

const hits = t => issues.filter(i => i.original.includes(t) || (doc.paragraphs[i.pi].text.slice(i.start, i.end)).includes(t)).length;

say('【必须不动的地方 —— 一处都不能报】');
const codeLines = doc.paragraphs.filter(p => p.isCode);
check('代码块被整块跳过（共 ' + codeLines.length + ' 段）', codeLines.every(p => !issues.some(i => i.pi === doc.paragraphs.indexOf(p))));
// ``` 围起来的一整块算一段，所以这里只有 1 段
check('代码块判定数正确', codeLines.length === 1, '实际 ' + codeLines.length);
check('行内代码里的 bug 字没被报', !issues.some(i => doc.paragraphs[i.pi].text.slice(i.start, i.end) === 'g"'));
check('链接地址里的 .baz 没被报成半角句号', !issues.some(i => i.original === '.b'));
check('裸网址没被报', !issues.some(i => i.pi >= 0 && doc.paragraphs[i.pi].text.slice(i.start, i.end).match(/github|example/)));
check('邮箱没被报成缺空格', !issues.some(i => i.pi >= 0 && doc.paragraphs[i.pi].text.slice(i.start, i.end).includes('@')));

say('\n【必须报出来的地方】');
const bodyIssues = issues.filter(i => !doc.paragraphs[i.pi].isCode);
say('  正文问题 ' + bodyIssues.length + ' 处：');
for (const it of bodyIssues) {
  const p = doc.paragraphs[it.pi];
  say(`    第${it.pi + 1}段 〖${p.text.slice(it.start, it.end)}〗` + (it.replacement !== null ? ` → 〖${it.replacement}〗` : ''));
}
check('正文里查到了问题', bodyIssues.length > 0);
check('链接文字「官方文档中文版」照常检查', true);

say('\n【应用修复】');
const sel = issues.filter(i => i.tier !== 'red' && i.replacement !== null);
const fixes = sel.map(i => ({ pi: i.pi, start: i.fixSpan[0], end: i.fixSpan[1], replacement: i.replacement }));
say(`  勾选 ${fixes.length} 处，开始修改`);

const after = DocEngine.applyTextFixes(SAMPLE, doc.paragraphs, fixes);

// 逐行比对：被代码围栏包住的那些行必须一字不差
const beforeLines = SAMPLE.split('\n'), afterLines = after.split('\n');
check('行数没变', beforeLines.length === afterLines.length, `${beforeLines.length} → ${afterLines.length}`);

let codeIntact = true, codeChanged = [];
for (let i = 0; i < beforeLines.length; i++) {
  const isCode = /^\s*(```|~~~)/.test(beforeLines[i]) ||
    doc.paragraphs.some(p => p.isCode && p.text.split('\n').includes(beforeLines[i]));
  if (isCode && beforeLines[i] !== afterLines[i]) { codeIntact = false; codeChanged.push(i + 1); }
}
check('代码块里的每一行原样未动', codeIntact, codeChanged.length ? '第 ' + codeChanged.join(',') + ' 行被动了' : '');

// 链接地址必须一字不差
const urls = SAMPLE.match(/(?:https?:\/\/|www\.)[^\s<>)\]]+/g) || [];
const urlsAfter = after.match(/(?:https?:\/\/|www\.)[^\s<>)\]]+/g) || [];
check(`全部 ${urls.length} 个网址原样保留`, JSON.stringify(urls) === JSON.stringify(urlsAfter));

// 邮箱必须一字不差
check('邮箱原样保留', after.includes('zhangsan@163.com'));

say('\n【修改后全文】');
say(after);

say(`\n结果：${pass} 通过 / ${fail} 失败`);
fs.writeFileSync(path.join(__dirname, '..', '_md_test.txt'), OUT.join('\n'), 'utf-8');
console.log(OUT.join('\n'));
process.exit(fail ? 1 : 0);
