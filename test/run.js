const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');
const DocEngine = require('../src/engine.js');
const RULES = require('../src/rules.js');

const OUT = [];
const say = s => OUT.push(s);

async function repack(inFile, outFile, newXml) {
  const zip = await JSZip.loadAsync(fs.readFileSync(inFile));
  zip.file('word/document.xml', newXml);
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  fs.writeFileSync(outFile, buf);
  return buf.length;
}

(async () => {
  const inFile = process.argv[2];
  if (!inFile) {
    console.error('用法: node test/run.js <文档.docx> [标签]');
    console.error('  会给这份文档跑一遍全部规则，导出 out_<标签>_fixed.docx 和 report_<标签>.md');
    process.exit(1);
  }
  const tag = process.argv[3] || 'r1';

  const zip = await JSZip.loadAsync(fs.readFileSync(inFile));
  const xml = await zip.file('word/document.xml').async('string');

  const t0 = Date.now();
  const doc = DocEngine.parseDocx(xml);
  const issues = DocEngine.runRules(doc, RULES);
  const ms = Date.now() - t0;

  const tierLabel = { green: '🟢可自动修', yellow: '🟡需确认', red: '🔴只报不改' };

  say(`文件：${path.basename(inFile)}`);
  say(`解析 ${doc.paragraphs.length} 段，耗时 ${ms}ms`);
  say(`共发现 ${issues.length} 处问题\n`);

  const byRule = {};
  for (const it of issues) (byRule[it.ruleId] ||= []).push(it);
  for (const r of RULES) {
    const list = byRule[r.id] || [];
    say(`${tierLabel[r.tier].padEnd(12)} ${r.name}  —— ${list.length} 处`);
  }

  say('\n' + '='.repeat(60));
  for (const r of RULES) {
    const list = byRule[r.id] || [];
    if (!list.length) continue;
    say(`\n【${r.name}】${list.length} 处`);
    list.slice(0, 6).forEach(it => {
      const p = doc.paragraphs[it.pi].text;
      const a = Math.max(0, it.start - 16), b = Math.min(p.length, it.end + 16);
      say(`   第${it.pi + 1}段  …${p.slice(a, it.start)}〖${it.original}〗${p.slice(it.end, b)}…`);
      if (it.replacement !== null) say(`          → 改成 〖${it.replacement}〗`);
    });
    if (list.length > 6) say(`    …还有 ${list.length - 6} 处`);
  }

  // ---- 应用修复：绿 + 黄（模拟用户全勾） ----
  const selected = issues.filter(i => i.tier !== 'red' && i.replacement !== null);
  const fixes = selected.map(i => ({ pi: i.pi, start: i.fixSpan[0], end: i.fixSpan[1], replacement: i.replacement }));
  const newXml = DocEngine.applyFixes(xml, doc.paragraphs, fixes);

  say('\n' + '='.repeat(60));
  say(`\n准备应用 ${fixes.length} 处修复（绿 + 黄，全部）`);

  const outDocx = path.join(__dirname, '..', `out_${tag}_fixed.docx`);
  const size = await repack(inFile, outDocx, newXml);
  say(`已生成 ${path.basename(outDocx)}（${size} 字节）`);

  // 回读校验：修复后的文字是否真的变了
  const zip2 = await JSZip.loadAsync(fs.readFileSync(outDocx));
  const xml2 = await zip2.file('word/document.xml').async('string');
  const doc2 = DocEngine.parseDocx(xml2);
  const issues2 = DocEngine.runRules(doc2, RULES);
  say(`回读校验：重跑规则后剩余 ${issues2.length} 处（原 ${issues.length} 处）`);
  const left = {};
  for (const it of issues2) left[it.ruleId] = (left[it.ruleId] || 0) + 1;
  say(`剩余分布：${JSON.stringify(left)}`);

  // 报告文件
  const md = DocEngine.buildMarkdownReport(path.basename(inFile), issues, RULES, doc.paragraphs);
  fs.writeFileSync(path.join(__dirname, '..', `report_${tag}.md`), md, 'utf-8');
  say(`已生成 report_${tag}.md`);

  fs.writeFileSync(path.join(__dirname, '..', `_test_${tag}.txt`), OUT.join('\n'), 'utf-8');
  console.log(OUT.join('\n'));
})().catch(e => {
  console.error('ERROR', e);
  process.exit(1);
});
