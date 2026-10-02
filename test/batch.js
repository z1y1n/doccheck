/* 批量压测：拿一批真实 docx 全部过一遍引擎，看有没有会崩、会卡、会改坏的 */
const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');
const DocEngine = require('../src/engine.js');
const RULES = require('../src/rules.js');

const DIRS = process.argv.slice(2);
const OUT = [];
const say = s => OUT.push(s);

function walk(dir, acc, depth) {
  if (depth > 3) return acc;
  let ents = [];
  try { ents = fs.readdirSync(dir, { withFileTypes: true }); } catch { return acc; }
  for (const e of ents) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) walk(p, acc, depth + 1);
    else if (/\.docx$/i.test(e.name) && !e.name.startsWith('~$')) acc.push(p);
  }
  return acc;
}

(async () => {
  const files = [];
  for (const d of DIRS) walk(d, files, 0);
  say(`扫描到 ${files.length} 个 docx\n`);
  say('文件'.padEnd(34) + '段数  问题  耗时   修复  回读  状态');
  say('-'.repeat(88));

  let bad = 0;
  for (const f of files) {
    const name = path.basename(f);
    const short = name.length > 32 ? name.slice(0, 31) + '…' : name;
    try {
      const zip = await JSZip.loadAsync(fs.readFileSync(f));
      const xf = zip.file('word/document.xml');
      if (!xf) { say(short.padEnd(34) + '   ——   （不是 Word 文档，跳过）'); continue; }
      const xml = await xf.async('string');

      const t0 = Date.now();
      const doc = DocEngine.parseDocx(xml);
      const issues = DocEngine.runRules(doc, RULES);
      const ms = Date.now() - t0;

      const sel = issues.filter(i => i.tier !== 'red' && i.replacement !== null);
      const fixes = sel.map(i => ({ pi: i.pi, start: i.fixSpan[0], end: i.fixSpan[1], replacement: i.replacement }));
      const nx = DocEngine.applyFixes(xml, doc.paragraphs, fixes);

      // 回读：把改完的 XML 重新解析，检查 (a) 能不能解析 (b) 修复是否生效 (c) 其他文字有没有丢
      const doc2 = DocEngine.parseDocx(nx);
      const before = doc.paragraphs.map(p => p.text).join('\n');
      const after = doc2.paragraphs.map(p => p.text).join('\n');
      // 每处修复的预期长度变化 = 替换文本长度 - 被替换区间的长度（插入时为 0 长度）
      const expected = fixes.reduce((s, f) => s + f.replacement.length - (f.end - f.start), 0);
      const actual = after.length - before.length;
      const lenDelta = actual - expected;   // 不为 0 说明改了不该改的地方
      const ok = lenDelta === 0;

      zip.file('word/document.xml', nx);
      const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
      const round = await JSZip.loadAsync(buf);
      const stillHasDoc = !!round.file('word/document.xml');
      const partsOk = Object.keys(round.files).length === Object.keys(zip.files).length;

      const status = ok && stillHasDoc && partsOk ? 'OK' : `⚠ 长度差${lenDelta} 部件${partsOk}`;
      if (status !== 'OK') bad++;
      say(short.padEnd(34) + String(doc.paragraphs.length).padStart(4) +
        String(issues.length).padStart(6) + String(ms + 'ms').padStart(7) +
        String(fixes.length).padStart(6) + '   ✓    ' + status);
    } catch (e) {
      bad++;
      say(short.padEnd(34) + '   !!   ' + e.message.slice(0, 50));
    }
  }
  say('\n异常文件数：' + bad);
  fs.writeFileSync(path.join(__dirname, '..', '_batch.txt'), OUT.join('\n'), 'utf-8');
  console.log(OUT.join('\n'));
})();
