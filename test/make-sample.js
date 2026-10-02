/* 生成「体检自检样本.docx」。
   每一节埋一类问题，同时埋一条反例——写正确的、不该被报的。
   用户拿它去核对工具查得全不全、有没有乱报。

   用法：node test/make-sample.js [输出路径] */
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
// 首尾有空格时必须开 preserve，否则 Word 打开会把空格吃掉
const sp = s => (/^\s|\s$/.test(s) ? ' xml:space="preserve"' : '');

// 把一个句子切成 n 段，模拟真实文档里「Word 把一句话拆成好几个 run」的情况
function splitRuns(text, n) {
  const step = Math.ceil(text.length / n);
  const out = [];
  for (let i = 0; i < text.length; i += step) out.push(text.slice(i, i + step));
  return out;
}

function para(text, opts = {}) {
  const rPr = opts.bold ? '<w:rPr><w:b/></w:rPr>' : '';
  if (text === '') return '<w:p><w:r><w:t></w:t></w:r></w:p>';
  const parts = opts.split ? splitRuns(text, opts.split) : [text];
  const runs = parts.map((t, i) =>
    `<w:r>${i === 0 ? rPr : ''}<w:t${sp(t)}>${esc(t)}</w:t></w:r>`).join('');
  return `<w:p>${runs}</w:p>`;
}

/* ---------------- 正文内容 ----------------
   每一节：一节标题 + 一行问题 + 一行反例。
   节标题本身必须干净——里面不能出现任何会被规则命中的字符，
   否则算出来的期望值就对不上了。 */

const DOC = [
  ['文档体检 · 自检样本', { bold: true }],
  [''],
  ['这份文档每一节都埋了一类问题，最后一行是反例。用它核对两件事：', {}],
  ['该报的有没有报全，不该报的有没有乱报。', {}],
  ['标着「反例」的那一行是故意写正确的，工具不该报它。', {}],
  [''],
  ['整份文档加起来，工具应该报出一共 34 处。', { bold: true }],
  [''],

  ['第一节 · 全角字母与数字（绿色，会自动勾上）', { bold: true }],
  ['这一行应该报 9 处，也就是 5 个全角字母加 4 个全角数字：', {}],
  ['用Ｅｘｃｅｌ做数据透视表，统计口径是２０２６年。', {}],
  ['反例，这一行不该报：', {}],
  ['用 Excel 做数据透视表，统计口径是 2026 年。', {}],
  [''],

  ['第二节 · 重复的中文标点（绿色，会自动勾上）', { bold: true }],
  ['这一行应该报 3 处：', {}],
  ['这个方案很好。。真的很好，，那就这么定了！！！', {}],
  ['反例，这一行不该报（省略号和破折号是正确写法）：', {}],
  ['省略号……和破折号——都是对的，重复的句号才不对。', {}],
  [''],

  ['第三节 · 中文里混用半角标点（黄色，要你点头）', { bold: true }],
  ['这一行应该报 3 处：', {}],
  ['这个方案可行,但是排期太紧.需要再评估;', {}],
  ['反例，这一行不该报（小数、版本号、英文缩写里的标点不算）：', {}],
  ['圆周率 3.14、版本号 v1.2.3、选项 A,B 都正常。', {}],
  [''],

  ['第四节 · 中文与英文之间缺空格（黄色，要你点头）', { bold: true }],
  ['这一行应该报 6 处（这一行的文字在文档内部被拆成了三段，顺便测跨段拼接）：', {}],
  ['我用QGIS做地图，用Excel统计数据，再写PRD文档。', { split: 3 }],
  ['反例，这一行不该报（全角标点两侧自带间距，不用加空格）：', {}],
  ['全角标点两侧不用加空格：PRD、（QGIS）、Excel；都正常。', {}],
  [''],

  ['第五节 · 全角空格（红色，只报不改）', { bold: true }],
  ['这一行应该报 1 处：', {}],
  ['第一部分　校招产品岗全景', {}],
  ['反例，这一行不该报（开头的全角空格是段首缩进）：', {}],
  ['　　这一行开头的全角空格是段首缩进。', {}],
  [''],

  ['第六节 · 连续多个空格（红色，只报不改）', { bold: true }],
  ['这一行应该报 2 处：', {}],
  ['联系方式：13800000000  |  zhangsan@163.com', {}],
  ['反例，这一行不该报（开头的半角空格是缩进）：', {}],
  ['  这一行开头的两个半角空格是缩进。', {}],
  [''],

  ['第七节 · 几种问题挤在同一段（考验修改会不会互相打架）', { bold: true }],
  ['这一段应该报 10 处，五种规则同时出现：', {}],
  ['产品能力：PRD撰写,用ＱＧＩＳ和ＳＱＬ处理数据。。就这样。', {}],
  [''],

  ['第八节 · 这一节是干净的', { bold: true }],
  ['这一整段不该报出任何问题：', {}],
  ['这一段完全是规范写法：中文标点全角，English 两侧留了空格，Excel 和 2026 也是半角。', {}],
];

/* ---------------- 打包成 docx ---------------- */

const body = DOC.map(([t, o]) => para(t, o)).join('');

const documentXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:body>' + body +
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
  '</w:body></w:document>';

const zip = new JSZip();
zip.file('[Content_Types].xml',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '</Types>');
zip.file('_rels/.rels',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>');
zip.file('word/document.xml', documentXml);

const out = process.argv[2] || path.join(__dirname, '..', 'sample.docx');

zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }).then(buf => {
  fs.writeFileSync(out, buf);
  console.log('已生成 ' + out);

  // 生成完立刻自己扫一遍，把实际结果打出来 —— 期望值不能靠嘴说
  const E = require('../src/engine.js');
  const R = require('../src/rules.js');
  const d = E.parseDocx(documentXml);
  const issues = E.runRules(d, R);

  console.log('\n实际扫出来的结果（' + issues.length + ' 处）：\n');
  console.log('段落  规则                  命中内容            改成');
  console.log('-'.repeat(72));
  for (const it of issues) {
    const p = d.paragraphs[it.pi];
    const got = p.text.slice(it.start, it.end);
    const to = it.replacement === null ? '（不改）' : (it.replacement === '' ? '（删掉）' : it.replacement);
    console.log(
      String(it.pi + 1).padStart(3) + '   ' +
      (R.find(r => r.id === it.ruleId).name + '                    ').slice(0, 20) + '  ' +
      ('「' + got.replace(/\n/g, '⏎') + '」                    ').slice(0, 20) + '  ' +
      to
    );
  }

  const byRule = {};
  for (const it of issues) byRule[it.ruleId] = (byRule[it.ruleId] || 0) + 1;
  console.log('\n按规则汇总：');
  for (const r of R) console.log('  ' + r.name + '：' + (byRule[r.id] || 0) + ' 处');
});
