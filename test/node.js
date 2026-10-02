/* 专测一个坑：Word 会把一句话拆成好几个 <w:t>，
   反过来同一个 <w:t> 里也可能同时有好几处要改。
   这一份就是盯着「一个节点上多处修复」打的。 */
const fs = require('fs');
const path = require('path');
const DocEngine = require('../src/engine.js');
const RULES = require('../src/rules.js');

const OUT = [];
const say = s => OUT.push(s);
let pass = 0, fail = 0;
const check = (name, cond, extra) => {
  if (cond) { pass++; say('  ✓ ' + name); }
  else { fail++; say('  ✗ ' + name + (extra ? '　' + extra : '')); }
};

const para = text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const doc = (...ps) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
  ps.join('') + '</w:body></w:document>';

function run(xml, label) {
  const d = DocEngine.parseDocx(xml);
  const issues = DocEngine.runRules(d, RULES);
  const sel = issues.filter(i => i.tier !== 'red' && i.replacement !== null);
  const fixes = sel.map(i => ({ pi: i.pi, start: i.fixSpan[0], end: i.fixSpan[1], replacement: i.replacement }));
  const out = DocEngine.applyFixes(xml, d.paragraphs, fixes);
  const d2 = DocEngine.parseDocx(out);
  const before = d.paragraphs.map(p => p.text).join('\n');
  const after = d2.paragraphs.map(p => p.text).join('\n');
  const expected = fixes.reduce((s, f) => s + f.replacement.length - (f.end - f.start), 0);
  say(`\n【${label}】`);
  say(`  改前：${before}`);
  say(`  改后：${after}`);
  say(`  修复 ${fixes.length} 处，预期长度变化 ${expected}，实际 ${after.length - before.length}`);
  return { before, after, delta: after.length - before.length, expected, fixes: fixes.length, out };
}

try {
  // ---- 1. 一处节点里两个补空格 ----
  say('=== 一个 <w:t> 里多处修复 ===');
  {
    const r = run(doc(para('重点讲PRD怎么写，用QGIS和SQL。')), '一句话里补两处空格');
    check('PRD 后面补了空格', r.after.includes('讲 PRD 怎么写'), r.after);
    check('QGIS 前补了空格', r.after.includes('用 QGIS'), r.after);
    check('SQL 前补了空格', r.after.includes('和 SQL'), r.after);
    check('长度变化符合预期', r.delta === r.expected, `${r.delta} vs ${r.expected}`);
    check('没有重复字符', !/。。|，，|写写/.test(r.after), r.after);
  }

  // ---- 2. 补空格 + 合并重复标点 挤在同一个节点 ----
  {
    const r = run(doc(para('产品能力：PRD撰写；Office办公软件等。。')), '补空格和改标点在同一节点');
    check('PRD 后面补了空格', r.after.includes('PRD 撰写'), r.after);
    check('Office 后面补了空格', r.after.includes('Office 办公'), r.after);
    check('重复句号合并了', r.after.includes('等。') && !r.after.includes('。。'), r.after);
    check('长度变化符合预期', r.delta === r.expected, `${r.delta} vs ${r.expected}`);
    // 中间插空格不会产生首尾空格，这时不该加 preserve
    check('中间改动不需要 xml:space', !r.out.includes('xml:space'));
  }

  // ---- 2b. xml:space 只在首尾出现空格时才该加 ----
  // 现在的 6 条规则都改不出首尾空格，但 applyFixes 是公开接口，
  // 以后加「去掉首尾空格」这类规则就会走到这条路。
  {
    const xml = doc(para('abc'));
    const d = DocEngine.parseDocx(xml);
    const out = DocEngine.applyFixes(xml, d.paragraphs, [{ pi: 0, start: 3, end: 3, replacement: ' ' }]);
    check('尾部补空格后打开 preserve', out.includes('xml:space="preserve"'), out.match(/<w:t[^>]*>/)?.[0]);
    check('空格确实留住了', DocEngine.parseDocx(out).paragraphs[0].text === 'abc ',
      '「' + DocEngine.parseDocx(out).paragraphs[0].text + '」');

    // 已经有 preserve 的不该重复加
    const xml2 = doc('<w:p><w:r><w:t xml:space="preserve">abc</w:t></w:r></w:p>');
    const d2 = DocEngine.parseDocx(xml2);
    const out2 = DocEngine.applyFixes(xml2, d2.paragraphs, [{ pi: 0, start: 3, end: 3, replacement: ' ' }]);
    check('已有 preserve 时不重复加', (out2.match(/xml:space/g) || []).length === 1,
      (out2.match(/xml:space/g) || []).length + ' 个');
  }

  // ---- 3. Word 把一句话拆成多个 run（真实手写文档就是这样） ----
  {
    const split = '<w:p><w:r><w:t>重点讲</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>PRD</w:t></w:r>' +
      '<w:r><w:t>怎么写，用</w:t></w:r><w:r><w:t>QGIS</w:t></w:r><w:r><w:t>处理</w:t></w:r></w:p>';
    const r = run(doc(split), '一句话被拆成 5 个 run');
    check('拼接后读到的文字是对的', r.before === '重点讲PRD怎么写，用QGIS处理', r.before);
    check('跨 run 边界补上了空格', r.after === '重点讲 PRD 怎么写，用 QGIS 处理', r.after);
    check('长度变化符合预期', r.delta === r.expected, `${r.delta} vs ${r.expected}`);
  }

  // ---- 4. 多段落各改各的 ----
  {
    const r = run(doc(para('第一段有PRD。'), para('第二段用QGIS。'), para('第三段没问题。')), '三段各改一处');
    check('三段都改了', r.after === '第一段有 PRD。\n第二段用 QGIS。\n第三段没问题。', r.after);
    check('长度变化符合预期', r.delta === r.expected, `${r.delta} vs ${r.expected}`);
  }

  // ---- 5. 同一句里补空格 + 半角标点改全角 ----
  {
    const r = run(doc(para('用SQL查询,然后导出。')), '补空格 + 半角逗号');
    check('SQL 前补了空格', r.after.includes('用 SQL'), r.after);
    check('逗号变全角', r.after.includes('查询，'), r.after);
    check('长度变化符合预期', r.delta === r.expected, `${r.delta} vs ${r.expected}`);
  }

  // ---- 6. 全角字母转半角 + 补空格 ----
  {
    const r = run(doc(para('用Ｅｘｃｅｌ做表格。')), '全角字母转半角');
    check('变成半角', r.after.includes('Excel'), r.after);
    check('长度变化符合预期', r.delta === r.expected, `${r.delta} vs ${r.expected}`);
  }

  // ---- 6b. 规则不级联：改完会冒出新问题，这件事必须能被发现 ----
  {
    const xml = doc(para('用ＱＧＩＳ处理数据。'));
    const d = DocEngine.parseDocx(xml);
    const iss = DocEngine.runRules(d, RULES);
    const fx = iss.filter(i => i.tier !== 'red' && i.replacement !== null)
      .map(i => ({ pi: i.pi, start: i.fixSpan[0], end: i.fixSpan[1], replacement: i.replacement }));
    const out = DocEngine.applyFixes(xml, d.paragraphs, fx);
    const d2 = DocEngine.parseDocx(out);
    check('全角字母转成了半角', d2.paragraphs[0].text.includes('QGIS'), d2.paragraphs[0].text);
    // 第一遍扫不到「用Q」，因为原文那里是全角 Ｑ
    // 第一遍扫不到「用Q」，因为原文那里是全角 Ｑ，不匹配 [A-Za-z]
    check('第一遍没发现「用Q」这个边界',
      !iss.some(i => i.ruleId === 'cjk-latin-space' && i.start === 1),
      'cjk-latin-space 不该匹配全角字符');
    const again = DocEngine.runRules(d2, RULES).filter(i => i.tier !== 'red' && i.replacement !== null);
    check('改完再扫能查出新边界', again.length > 0, '查出 ' + again.length + ' 处');
    say('    新出现的问题：' + again.map(i => d2.paragraphs[i.pi].text.slice(i.start, i.end)).join(' / '));
  }

  // ---- 7. 带下标的段落：多段整体重排，位置不能串 ----
  {
    const r = run(doc(para('甲PRD。'), para('乙QGIS。'), para('丙SQL。'), para('丁Office。')), '四段连改');
    check('四段都改对了', r.after === '甲 PRD。\n乙 QGIS。\n丙 SQL。\n丁 Office。', r.after);
    check('长度变化符合预期', r.delta === r.expected, `${r.delta} vs ${r.expected}`);
  }

  // ---- 8. 短句内合并：只影响「几行」，不影响「几处」 ----
  // 用户的原话：「用Ｅｘｃｅｌ做数据透视表，几个字母的修改都需要报一遍太复杂」。
  // 规则是按字符报的，靠这一层把同一短句里的合成一条给人和清单看。
  say('\n【短句内合并】');
  {
    const g = text => {
      const d = DocEngine.parseDocx(doc(para(text)));
      const iss = DocEngine.runRules(d, RULES);
      return { iss, gs: DocEngine.groupIssues(iss, d.paragraphs), d };
    };

    // 逗号之内合并
    let r = g('用Ｅｘｃｅｌ做数据透视表。');
    check('5 个全角字母合成 1 行', r.gs.length === 1 && r.gs[0].items.length === 5,
      r.gs.length + ' 行 / ' + (r.gs[0] || { items: [] }).items.length + ' 项');
    check('处数照实算，还是 5 处', r.iss.length === 5, r.iss.length + ' 处');

    // 逗号断开：一个短句一行
    r = g('用Ｅｘｃｅｌ做表，统计口径是２０２６年。');
    check('逗号两侧不合并，分成 2 行', r.gs.length === 2, r.gs.length + ' 行');
    check('两行分别是 5 处和 4 处',
      r.gs.map(x => x.items.length).join('/') === '5/4',
      r.gs.map(x => x.items.length).join('/'));
    check('逗号本身没被判成断句错误', r.iss.length === 9, r.iss.length + ' 处');

    // 顿号不算断句：「Ａ、Ｂ」还是一个短句
    r = g('甲Ａ、乙Ｂ。');
    check('顿号不断句，合并成 1 行', r.gs.length === 1, r.gs.length + ' 行');

    // 不同规则绝不合并——档位不同，不能共用一个勾
    r = g('用QGIS处理。。');
    check('不同规则各自成行', r.gs.length === 2, r.gs.length + ' 行');
    check('两行规则不同', r.gs[0].ruleId !== r.gs[1].ruleId,
      r.gs.map(x => x.ruleId).join(' / '));

    // 换行也算断句
    r = g('甲Ａ\n乙Ｂ');
    check('换行处断开', r.gs.length === 2, r.gs.length + ' 行');

    // 「改后」渲染：每处只吃掉自己 fixSpan 那几个字符，不能整段替换
    const d = DocEngine.parseDocx(doc(para('这个方案可行,但是排期太紧。')));
    const iss = DocEngine.runRules(d, RULES);
    const gs = DocEngine.groupIssues(iss, d.paragraphs);
    const dd = DocEngine.groupDiff(d.paragraphs[0], gs[0], 40);
    // 命中是「行,」，实际要改的只有那个逗号，「行」必须留着
    check('划掉的范围含「行」', dd.original === '行,', JSON.stringify(dd.original));
    check('改后是「行，」而不是只剩「，」', dd.replacement === '行，', JSON.stringify(dd.replacement));
  }
} catch (e) {
  fail++;
  say('  ✗ 抛异常：' + (e && e.stack || e));
}

say(`\n结果：${pass} 通过 / ${fail} 失败`);
fs.writeFileSync(path.join(__dirname, '..', '_node_test.txt'), OUT.join('\n'), 'utf-8');
console.log(OUT.join('\n'));
process.exit(fail ? 1 : 0);
