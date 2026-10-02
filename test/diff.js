/* 对比「原件」和「导出件」的压缩包内部，看看究竟改了什么、有没有丢东西。 */
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const ROOT = path.join(__dirname, '..');

async function peek(file) {
  const p = path.isAbsolute(file) ? file : path.join(ROOT, file);
  const zip = await JSZip.loadAsync(fs.readFileSync(p));
  const names = Object.keys(zip.files).filter(n => !zip.files[n].dir).sort();
  const media = names.filter(n => /^word\/media\//.test(n));
  const xml = await zip.file('word/document.xml').async('string');
  return { names, media, xml, zip };
}

(async () => {
  const fileA = process.argv[2];
  const fileB = process.argv[3];
  if (!fileA || !fileB) {
    console.error('用法: node test/diff.js <原始.docx> <导出.docx>');
    process.exit(1);
  }

  const a = await peek(fileA);
  const b = await peek(fileB);

  const out = [];
  out.push(`【${path.basename(fileA)} 内部结构】`);
  out.push('  部件数：' + a.names.length);
  out.push('  图片：' + (a.media.length ? a.media.join(', ') : '无'));
  out.push('  部件清单：' + a.names.join('\n              '));

  out.push('');
  out.push(`【和 ${path.basename(fileB)} 比】`);
  const onlyA = a.names.filter(n => !b.names.includes(n));
  const onlyB = b.names.filter(n => !a.names.includes(n));
  out.push('  只在原件里的部件：' + (onlyA.length ? onlyA.join(', ') : '无'));
  out.push('  只在导出件里的部件：' + (onlyB.length ? onlyB.join(', ') : '无'));

  let same = 0, diff = [];
  for (const n of a.names) {
    if (!b.names.includes(n)) continue;
    const x = await a.zip.file(n).async('string');
    const y = await b.zip.file(n).async('string');
    if (x === y) same++; else diff.push(n);
  }
  out.push('  内容完全相同的部件：' + same + ' 个');
  out.push('  内容不同的部件：' + (diff.length ? diff.join(', ') : '无'));

  // 逐段比正文
  const DocEngine = require('../src/engine.js');
  const pa = DocEngine.parseDocx(a.xml).paragraphs.map(p => p.text);
  const pb = DocEngine.parseDocx(b.xml).paragraphs.map(p => p.text);
  out.push('');
  out.push('【正文逐段对比】');
  out.push('  段落数：' + pa.length + ' vs ' + pb.length);
  let changed = 0;
  for (let i = 0; i < Math.min(pa.length, pb.length); i++) {
    if (pa[i] !== pb[i]) {
      changed++;
      out.push(`  第 ${i + 1} 段不同：`);
      out.push('    原：' + pa[i]);
      out.push('    新：' + pb[i]);
    }
  }
  out.push('  不同的段落数：' + changed);

  // 统计原始文档查出的问题，看看这次导出勾了哪些
  const RULES = require('../src/rules.js');
  const d = DocEngine.parseDocx(a.xml);
  const issues = DocEngine.runRules(d, RULES);
  const byRule = {};
  issues.forEach(i => { byRule[i.ruleId] = (byRule[i.ruleId] || 0) + 1; });
  out.push('');
  out.push(`【${path.basename(fileA)} 查出的问题分布】`);
  out.push('  合计 ' + issues.length + ' 处');
  Object.entries(byRule).forEach(([k, v]) => {
    const r = RULES.find(x => x.id === k);
    out.push(`    ${r.name}（${r.tier}）：${v} 处`);
  });

  fs.writeFileSync(path.join(ROOT, '_diff.txt'), out.join('\n'), 'utf-8');
  console.log(out.join('\n'));
})();
