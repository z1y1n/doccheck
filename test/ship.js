/* 测「出厂件」——打包出来的那个单文件 文档体检.html。
   前面三份测试跑的都是 src/ 里的源码，而用户双击的是这个打包文件。
   这里把打包文件整个丢进 jsdom 真跑一遍，确认：四个脚本都进去了、
   能执行、能读一份真 docx、能导出、导出结果能再解析。 */
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

const ROOT = path.join(__dirname, '..');
const OUT = [];
const say = s => OUT.push(s);
let pass = 0, fail = 0;
const check = (name, cond, extra) => {
  if (cond) { pass++; say('  ✓ ' + name); }
  else { fail++; say('  ✗ ' + name + (extra ? '　' + extra : '')); }
};

const html = fs.readFileSync(path.join(ROOT, '文档体检.html'), 'utf-8');

// —— 先做静态检查：出厂件不能有任何外部依赖 ——
say('【出厂件静态检查】');
check('没有外链的 script src', !/<script[^>]+src=/i.test(html),
  (html.match(/<script[^>]+src=[^>]*>/i) || [''])[0]);
check('没有外链的 link href', !/<link[^>]+href=/i.test(html),
  (html.match(/<link[^>]+href=[^>]*>/i) || [''])[0]);
check('没有 import 外部模块', !/^\s*import\s/m.test(html) && !/@import/.test(html));
const inline = (html.match(/<script(?![^>]*src)[^>]*>/gi) || []).length;
check('内联脚本是 4 块（jszip + engine + rules + ui）', inline === 4, '实际 ' + inline + ' 块');
check('引擎的导出清单在里面', html.includes("exports.parseDocx") || html.includes('parseDocx'),
  '');

(async () => {
  say('\n【出厂件真跑一份 docx】');
  const dom = new JSDOM(html, {
    url: 'http://localhost/', runScripts: 'dangerously', pretendToBeVisual: true
  });
  const { window } = dom;

  // jszip 换成 Node 那一份：打包文件里的 JSZip 属于 jsdom 这个 realm，
  // 拿 Node 的 ArrayBuffer 喂它会过不了 instanceof（这个坑前面踩过）
  window.JSZip = require('jszip');

  const got = [];
  window.URL.createObjectURL = blob => { got.push(blob); return 'blob:fake'; };
  window.URL.revokeObjectURL = () => {};
  window.HTMLAnchorElement.prototype.click = function () {
    if (got.length) got[got.length - 1]._name = this.download;
  };

  const DC = window.DocCheck;
  check('打包文件里的 ui.js 跑起来了', !!DC);
  if (!DC) { say('\n结果：' + pass + ' 通过 / ' + (fail + 1) + ' 失败'); console.log(OUT.join('\n')); process.exit(1); }

  const JSZip = require('jszip');
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
  zip.file('word/document.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    '<w:p><w:r><w:t>重点讲PRD怎么写，用QGIS和SQL。</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>等。。</w:t></w:r></w:p>' +
    '</w:body></w:document>');
  const buf = await zip.generateAsync({ type: 'nodebuffer' });

  await DC.openFile({
    name: '真件.docx',
    arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
    text: async () => buf.toString('utf-8')
  });

  check('认出是 Word 文档', DC.fileInfo().meta.includes('Word 文档'), DC.fileInfo().meta);
  check('查出问题并渲染成报告', DC.reportHtml().length > 200, DC.reportHtml().length + ' 字符');
  const boxes = [...window.document.querySelectorAll('#report input[data-role="one"]')];
  check('有可勾选的修复项', boxes.length > 0, boxes.length + ' 项');
  check('底部条显示了待改数', DC.barText().includes('将自动修改'), DC.barText());

  // 绿色规则默认勾选，直接导出。注意：const 声明不会挂到 window 上，
  // 所以这里不从 window 拿引擎，直接查导出的 XML 里有没有该有的字。
  const unxml = s => s.replace(/<[^>]+>/g, '');
  const n = DC.currentFixes().length;
  got.length = 0;
  window.document.querySelector('#exportDoc').click();
  await new Promise(r => setTimeout(r, 500));
  check('导出了文件', got.length === 1, '实际 ' + got.length);
  if (got.length) {
    check('文件名正确', got[0]._name === '真件-已修改.docx', got[0]._name);
    const z2 = await JSZip.loadAsync(Buffer.from(await got[0].arrayBuffer()));
    const missing = ['[Content_Types].xml', '_rels/.rels', 'word/document.xml'].filter(f => !z2.file(f));
    check('压缩包部件齐全', missing.length === 0, '缺 ' + missing.join(','));
    const xml2 = await z2.file('word/document.xml').async('string');
    const text = unxml(xml2.slice(xml2.indexOf('<w:body>')));
    say('    导出后正文：' + text);
    check('重复句号合并了', text.includes('等。') && !text.includes('。。'), text);
    check('绿色项确实有得改', n > 0, '实际 ' + n + ' 处');
    check('原文一个字没丢', text.includes('重点讲') && text.includes('QGIS') && text.includes('SQL'), text);
    check('黄色项没被顺手改掉（默认不勾）', text.includes('PRD怎么写'), text);
  }

  // 全勾上再导一次，这次该补空格了
  [...window.document.querySelectorAll('#report input[data-role="all"]')].forEach(b => {
    b.checked = true;
    b.dispatchEvent(new window.Event('change', { bubbles: true }));
  });
  got.length = 0;
  window.document.querySelector('#exportDoc').click();
  await new Promise(r => setTimeout(r, 500));
  check('全勾后导出了文件', got.length === 1, '实际 ' + got.length);
  if (got.length) {
    const z3 = await JSZip.loadAsync(Buffer.from(await got[0].arrayBuffer()));
    const xml3 = await z3.file('word/document.xml').async('string');
    const text3 = unxml(xml3.slice(xml3.indexOf('<w:body>')));
    say('    全勾导出后正文：' + text3);
    check('中英之间补上了空格', text3.includes('讲 PRD 怎么写'), text3);
    check('夹心结构「和SQL」也补对了', text3.includes('QGIS 和 SQL'), text3);
    check('句尾句号也合并了', text3.includes('等。') && !text3.includes('。。'), text3);
  }

  // 清单也导出一次
  got.length = 0;
  window.document.querySelector('#exportMd').click();
  await new Promise(r => setTimeout(r, 400));
  check('待处理清单也导得出', got.length === 1 && /待处理清单\.md$/.test(got[0]._name),
    got.length ? got[0]._name : '没导出');

  say(`\n结果：${pass} 通过 / ${fail} 失败`);
  fs.writeFileSync(path.join(ROOT, '_ship_test.txt'), OUT.join('\n'), 'utf-8');
  console.log(OUT.join('\n'));
  process.exit(fail ? 1 : 0);
})();
