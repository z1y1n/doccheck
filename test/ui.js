/* 界面端到端测试：在 jsdom 里真跑一遍「拖进文件 → 渲染 → 勾选 → 导出」。
   测的是 src/ui.js 本身，和浏览器里跑的是同一份代码。 */
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

const ROOT = path.join(__dirname, '..');
const OUT = [];
const say = s => OUT.push(s);
let pass = 0, fail = 0;
const check = (name, cond, extra) => {
  try { fs.appendFileSync(path.join(__dirname, '..', '_trace.txt'), '   check: ' + name + ' → ' + (cond ? '✓' : '✗') + (extra ? '　' + extra : '') + '\n'); } catch {}
  if (cond) { pass++; say('  ✓ ' + name); }
  else { fail++; say('  ✗ ' + name + (extra ? '　' + extra : '')); }
};

// 用 index.html 的真实骨架，保证选择器和真环境一致
const dom = new JSDOM(fs.readFileSync(path.join(ROOT, 'src', 'index.html'), 'utf-8'), {
  url: 'http://localhost/', runScripts: 'outside-only', pretendToBeVisual: true
});
const { window } = dom;
global.window = window;
global.document = window.document;

// 把「下载」拦下来，抓 blob
const got = [];
window.URL.createObjectURL = blob => { got.push(blob); return 'blob:fake'; };
window.URL.revokeObjectURL = () => {};
window.HTMLAnchorElement.prototype.click = function () {
  if (got.length) got[got.length - 1]._name = this.download;
};
window.alert = () => {};

// 三个文件拼成一次 eval —— window.eval 每次调用有独立的词法作用域，
// 分三次注入的话 ui.js 是看不到 engine.js 里的 const 的。
const bundle = ['engine.js', 'rules.js', 'ui.js']
  .map(f => fs.readFileSync(path.join(ROOT, 'src', f), 'utf-8'))
  .join('\n;\n');
const TRACE = m => fs.appendFileSync(path.join(ROOT, '_trace.txt'), m + '\n');
fs.writeFileSync(path.join(ROOT, '_trace.txt'), '');

window.eval(bundle + '\n;window.DocEngine = DocEngine; window.RULES = RULES;');
TRACE('1 bundle eval 完成');

// jszip 用 Node 那一份，不要 eval 进窗口 —— 两个 realm 的 ArrayBuffer
// 互相过不了 instanceof，loadAsync 会认不出参数类型
window.JSZip = require('jszip');
TRACE('2 jszip 就绪, typeof=' + typeof window.JSZip);

const { DocCheck, DocEngine } = window;
TRACE('3 DocCheck=' + typeof DocCheck);
if (!DocCheck) { console.log('ui.js 没跑起来'); process.exit(1); }

const $ = s => window.document.querySelector(s);
const wait = ms => new Promise(r => setTimeout(r, ms));

// jsdom 的 File 没有 arrayBuffer/text，补上
function fakeFile(name, content) {
  const buf = Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf-8');
  return {
    name,
    arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
    text: async () => buf.toString('utf-8')
  };
}

async function makeDocx(text) {
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
    '<w:p><w:r><w:t>' + text + '</w:t></w:r></w:p>' +
    '</w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

TRACE('4 定义完辅助函数，准备进 async');
(async () => {
  TRACE('5 已进入 async');
  try {
    // ============ 一、Markdown ============
    say('【Markdown 流程】');
    const md = '# 转行指南\n\n这份指南写给准备2027届的同学，重点讲PRD怎么写。\n\n' +
      '```js\nconst a=1;  // 中文注释English\n```\n\n' +
      '链接：[文档](https://example.com/a,b.html) 和 https://github.com/x/y.z\n';
    TRACE('6 准备调 openFile');
    const _p = DocCheck.openFile(fakeFile('sample.md', md));
    TRACE('7 拿到 promise：' + typeof _p.then);
    await _p;
    TRACE('8 openFile 回来了');

    check('文件信息条显示了文件名', DocCheck.fileInfo().name === 'sample.md');
    check('meta 显示是 Markdown', DocCheck.fileInfo().meta.includes('Markdown'));
    const html = DocCheck.reportHtml();
    check('报告渲染出来了', html.length > 200, '实际 ' + html.length + ' 字符');
    const groups = window.document.querySelectorAll('#report .group');
    check('按规则分了组', groups.length === 1, '实际 ' + groups.length + ' 组');
    const boxes = [...window.document.querySelectorAll('#report input[data-role="one"]')];
    // 「重点讲PRD怎么写」里有两个边界（讲|P 和 D|怎），同一句，合并成一行
    check('同一句里的两处合并成一行', boxes.length === 1, '实际 ' + boxes.length + ' 行');
    check('但处数照实算，还是 2 处', DocCheck.state.issues.length === 2,
      '实际 ' + DocCheck.state.issues.length + ' 处');
    check('代码块内容没出现在报告里', !html.includes('注释English'));
    check('网址没被当问题', !html.includes('a,b.html') || !html.includes('〖.b〗'));

    // 这两处是黄色规则，默认不该勾上
    check('黄色问题默认不勾选', boxes.every(b => !b.checked));
    check('底部条如实说没勾', DocCheck.barText().includes('没有勾选'), '实际「' + DocCheck.barText() + '」');

    // 没勾就点导出，应该被拦住
    got.length = 0;
    $('#exportDoc').click();
    await wait(150);
    check('没勾选时点导出会被拦下', got.length === 0, '实际生成了 ' + got.length + ' 个');

    // 一处都没勾，清单里就该列着这两处
    got.length = 0;
    $('#exportMd').click();
    await wait(200);
    check('待处理清单导出了', got.length === 1 && got[0]._name === 'sample-待处理清单.md',
      got.length ? got[0]._name : '没导出');
    if (got.length) {
      const todo = Buffer.from(await got[0].arrayBuffer()).toString('utf-8');
      check('清单里有段落定位', todo.includes('第 2 段'), todo.split('\n').filter(l => l.startsWith('- **')).join(' / '));
      check('清单里说 0 处已改', /已自动修改：\*\*\s*0\s*处/.test(todo),
        (todo.match(/已自动修改：.*/) || ['(没找到)'])[0]);
      check('清单里说 2 处待处理', /待你处理：\*\*\s*2\s*处/.test(todo),
        (todo.match(/待你处理：.*/) || ['(没找到)'])[0]);
      check('清单里给出了建议改法', todo.includes('建议改成'), '');
    }

    // 勾上再导出 —— 这是最容易出错的路径：零宽插入
    boxes.forEach(b => { b.checked = true; b.dispatchEvent(new window.Event('change', { bubbles: true })); });
    check('勾上后待改数变成 2', DocCheck.currentFixes().length === 2, '实际 ' + DocCheck.currentFixes().length);
    check('底部条跟着更新了', DocCheck.barText().includes('将自动修改'), '实际「' + DocCheck.barText() + '」');

    got.length = 0;
    $('#exportDoc').click();
    await wait(250);
    check('导出生成了文件', got.length === 1, '实际 ' + got.length);
    if (got.length) {
      check('文件名是 sample-已修改.md', got[0]._name === 'sample-已修改.md', '实际 ' + got[0]._name);
      const txt = Buffer.from(await got[0].arrayBuffer()).toString('utf-8');
      check('正文里加上了空格', txt.includes('讲 PRD 怎么写'), '实际：' + txt.split('\n')[2]);
      check('代码块原样没动', txt.includes('const a=1;  // 中文注释English'));
      check('链接地址原样没动', txt.includes('https://example.com/a,b.html'));
      check('没多改也没少改字符', txt.length - md.length === 2, '长度差 ' + (txt.length - md.length));
    }

    // 全改完之后，清单应该是空的
    got.length = 0;
    $('#exportMd').click();
    await wait(200);
    if (got.length) {
      const todo = Buffer.from(await got[0].arrayBuffer()).toString('utf-8');
      check('全部改完后清单里没有待办', /待你处理：\*\*\s*0\s*处/.test(todo),
        (todo.match(/待你处理：.*/) || ['(没找到)'])[0]);
      check('全部改完后清单不再列条目', !todo.includes('建议改成'));
    } else check('全部改完后的清单也能导出', false, '没导出');

    // ============ 二、勾选交互 ============
    say('\n【勾选交互】');
    const allBox = window.document.querySelector('#report .group input[data-role="all"]');
    check('分组标题上有全选框', !!allBox);
    if (allBox) {
      const rule = allBox.dataset.rule;
      // 界面按「修改组」勾选，所以下标是进 groups 而不是 issues
      const mine = () => [...window.document.querySelectorAll('#report input[data-role="one"]')]
        .filter(b => DocCheck.state.groups[+b.dataset.idx].ruleId === rule);
      const before = DocCheck.currentFixes().length;
      check('刚才全勾上了，全选框也是勾的', before === 2 && allBox.checked);

      allBox.checked = false;
      allBox.dispatchEvent(new window.Event('change', { bubbles: true }));
      check('取消全选后组内每一项都不勾', mine().every(b => !b.checked));
      check('取消后待改数归零', DocCheck.currentFixes().length === 0);

      allBox.checked = true;
      allBox.dispatchEvent(new window.Event('change', { bubbles: true }));
      check('重新全选后都勾上', mine().every(b => b.checked));
      check('重新全选后恢复 2 处', DocCheck.currentFixes().length === 2);
    }

    // ============ 三、docx ============
    say('\n【Word 文档流程】');
    DocCheck.reset();
    check('reset 之后回到初始状态', DocCheck.reportHtml() === '' && $('#filebar').hidden);

    const buf = await makeDocx('产品能力：PRD撰写；Office办公软件等。。');
    await DocCheck.openFile(fakeFile('简历.docx', buf));

    check('识别为 Word 文档', DocCheck.fileInfo().meta.includes('Word 文档'));
    const dboxes = [...window.document.querySelectorAll('#report input[data-role="one"]')];
    check('查出了问题', dboxes.length > 0, '实际 ' + dboxes.length + ' 处');
    check('绿色问题默认勾选', dboxes.some(b => b.checked));
    check('重复句号进了报告', DocCheck.reportHtml().includes('。。'));
    check('红色问题没有勾选框',
      [...window.document.querySelectorAll('#report .group')]
        .filter(g => g.querySelector('.pill.r'))
        .every(g => !g.querySelector('input[data-role="all"]')));

    check('黄色问题默认不勾', dboxes.some(b => !b.checked));

    const JSZip = require('jszip');
    let r1 = null;
    const unpack = async blob => {
      const z = await JSZip.loadAsync(Buffer.from(await blob.arrayBuffer()));
      const missing = ['[Content_Types].xml', '_rels/.rels', 'word/document.xml'].filter(f => !z.file(f));
      const xml = await z.file('word/document.xml').async('string');
      return { xml, missing, text: DocEngine.parseDocx(xml).paragraphs.map(p => p.text).join('|') };
    };

    // —— 第一次导出：只按默认勾选（绿色），黄色不该被动 ——
    got.length = 0;
    $('#exportDoc').click();
    await wait(400);
    check('docx 导出成功', got.length === 1, '实际 ' + got.length);
    if (got.length) {
      check('文件名是 简历-已修改.docx', got[0]._name === '简历-已修改.docx', '实际 ' + got[0]._name);
      r1 = await unpack(got[0]);
      check('压缩包里该有的部件都在', r1.missing.length === 0, '缺 ' + r1.missing.join(','));
      say('    默认导出的正文：' + r1.text);
      check('重复句号合并了', r1.text.includes('等。') && !r1.text.includes('。。'));
      check('原文一个字没丢', r1.text.includes('产品能力'));
      check('XML 结构完整', r1.xml.includes('<w:body>') && r1.xml.includes('</w:document>'));
      check('黄色问题没被动（PRD 后面还是没空格）', r1.text.includes('PRD撰写'));
    }

    // —— 第二次导出：全勾上，这次才该补空格 ——
    [...window.document.querySelectorAll('#report input[data-role="all"]')].forEach(b => {
      b.checked = true;
      b.dispatchEvent(new window.Event('change', { bubbles: true }));
    });
    check('全选后待改 3 处', DocCheck.currentFixes().length === 3, '实际 ' + DocCheck.currentFixes().length);

    got.length = 0;
    $('#exportDoc').click();
    await wait(400);
    check('全选后导出成功', got.length === 1, '实际 ' + got.length);
    if (got.length) {
      const r2 = await unpack(got[0]);
      say('    全选导出的正文：' + r2.text);
      check('PRD 后面补了空格', r2.text.includes('PRD 撰写'));
      check('Office 后面补了空格', r2.text.includes('Office 办公'));
      // 补的是句中空格，不涉及首尾，不该动 xml:space
      check('句中补空格不碰 xml:space', !r2.xml.includes('xml:space'), '意外加了 preserve');
      check('两处改动都生效了', r2.text.includes('PRD 撰写') && r2.text.includes('Office 办公'), r2.text);
      check('原文一个字没丢', r2.text.includes('产品能力'));
      check('长度只该多出 2 个空格', r2.text.length - r1.text.length === 2,
        '实际多 ' + (r2.text.length - r1.text.length));
    }

    // ============ 四、同一句里的多处修改合并成一行 ============
    // 用户的反馈：「用Ｅｘｃｅｌ做数据透视表，几个字母的修改都需要报一遍太复杂」。
    // 规则是按字符报的，但界面不能照实列——同一短句里合成一条，导出时再拆开。
    say('\n【短句内合并】');
    DocCheck.reset();
    const buf4 = await makeDocx('用Ｅｘｃｅｌ做数据透视表，统计口径是２０２６年。');
    await DocCheck.openFile(fakeFile('合并.docx', buf4));

    const mergedRows = [...window.document.querySelectorAll('#report .item')];
    check('9 处全角字符合并成 2 行（中间隔着逗号，不跨句合并）',
      mergedRows.length === 2, '实际 ' + mergedRows.length + ' 行');
    check('行上标着合并了几处',
      mergedRows.map(r => (r.querySelector('.merged') || {}).textContent).join(' / ') === '合并 5 处 / 合并 4 处',
      mergedRows.map(r => (r.querySelector('.merged') || {}).textContent).join(' / '));

    // 计数口径不能变：合并只减少「几行」，不减少「几处」
    check('组标题仍然报 9 处，不是 2 处',
      window.document.querySelector('#report .group .count').textContent === '9 处',
      window.document.querySelector('#report .group .count').textContent);
    check('顶部数字也是 9', $('#summary .stat.g .n').textContent === '9',
      $('#summary .stat.g .n').textContent);
    check('合并后问题总数没变', DocCheck.state.issues.length === 9,
      '实际 ' + DocCheck.state.issues.length);

    // 合并显示不能把「改后」渲染错。这一行画的是「原文划掉 → 新文字」，
    // 所以文字内容里旧的和新的都会出现——要分别看两个 mark 才准。
    // （这里也是半角标点那个老显示错误的防线：以前「行,」会显示成只改成「，」，把「行」吃掉。）
    const diffRow = mergedRows[0].querySelectorAll('.line')[1];
    check('划掉的是整串原文', diffRow.querySelector('mark.del').textContent === 'Ｅｘｃｅｌ',
      diffRow.querySelector('mark.del').textContent);
    check('改后是 Excel 三个字母，不多不少',
      diffRow.querySelector('mark.new').textContent === 'Excel',
      diffRow.querySelector('mark.new').textContent);

    // 勾一行 = 勾这一句里的全部修改
    const firstBox = mergedRows[0].querySelector('input[data-role="one"]');
    const secondBox = mergedRows[1].querySelector('input[data-role="one"]');
    check('绿色默认勾上', firstBox.checked && secondBox.checked);
    check('两行都是绿色，默认 9 处全在待改里', DocCheck.currentFixes().length === 9,
      '实际 ' + DocCheck.currentFixes().length);

    // 取消前一行：句内 5 处一起取消，另一句的 4 处不受影响
    firstBox.checked = false;
    firstBox.dispatchEvent(new window.Event('change', { bubbles: true }));
    check('取消这一行，句内 5 处一起取消', DocCheck.currentFixes().length === 4,
      '实际 ' + DocCheck.currentFixes().length);
    check('另一行没被连累', secondBox.checked && DocCheck.currentFixes().length === 4);
    DocCheck.state.groups[0].items.forEach(i =>
      check('句内每一处都真的取消了（' + i.original + '）', i.sel === false));

    firstBox.checked = true;
    firstBox.dispatchEvent(new window.Event('change', { bubbles: true }));
    check('重新勾上，5 处一起回来', DocCheck.currentFixes().length === 9,
      '实际 ' + DocCheck.currentFixes().length);

    // 只留前一句，去导一次
    secondBox.checked = false;
    secondBox.dispatchEvent(new window.Event('change', { bubbles: true }));
    check('只剩前一句的 5 处', DocCheck.currentFixes().length === 5,
      '实际 ' + DocCheck.currentFixes().length);

    got.length = 0;
    $('#exportDoc').click();
    await wait(300);
    if (got.length) {
      const t = (await unpack(got[0])).text;
      say('    只改前一句导出后：' + t);
      check('后一句的全角数字原样保留', t.includes('２０２６'), t);
      check('前一句整句改对了', t.includes('用Excel做数据透视表'), t);
    } else check('合并后导出能用', false, '没导出');

    // ============ 五、🔴 那类默认折叠 ============
    // 背景：拿真实简历一试，23 条里 21 条是「连续多个空格」，而且全是有意打的
    // 对齐空格。全摊开会把真该改的那两条淹掉，所以🔴默认收起来。
    say('\n【默认折叠】');
    DocCheck.reset();
    // 文本里不能有邮箱或网址：cjk-latin-space 那条规则会绕开它们，
    // 否则黄色分组根本不会出现，这一节就测不到 🟡 展开不展开了
    const buf3 = await makeDocx('联系方式：甲乙  |  丙丁。这里用PRD等。。');
    await DocCheck.openFile(fakeFile('带对齐空格.docx', buf3));

    const allGroups = [...window.document.querySelectorAll('#report .group')];
    const redG = allGroups.filter(g => g.querySelector('.pill.r'));
    const otherG = allGroups.filter(g => !g.querySelector('.pill.r'));
    check('三类问题都造出来了', redG.length > 0 && otherG.length > 0,
      '🔴 ' + redG.length + ' 组 / 其他 ' + otherG.length + ' 组');
    check('🔴 分组默认折叠', redG.every(g => g.classList.contains('collapsed')));
    check('🟢🟡 分组默认展开', otherG.every(g => !g.classList.contains('collapsed')));
    check('🟢🟡🔴 三档都出现了，这节才算测全',
      !!redG.length && otherG.some(g => g.querySelector('.pill.g')) && otherG.some(g => g.querySelector('.pill.y')),
      '分组：' + allGroups.map(g => g.dataset.rule + '/' + (g.querySelector('.pill').className.split(' ')[1] || '?')).join(' '));

    if (redG.length) {
      const g0 = redG[0];
      const note = g0.querySelector('.note');
      check('折叠后规则说明仍然露着', !!note && !note.closest('.body'),
        note ? '「' + note.textContent.slice(0, 18) + '…」' : '没有说明');
      const body = g0.querySelector('.body');
      check('被折叠的条目确实在里面藏着',
        !!body && body.querySelectorAll('.item').length > 0,
        body ? body.querySelectorAll('.item').length + ' 条' : '没有 body');

      g0.querySelector('.head').dispatchEvent(new window.Event('click', { bubbles: true }));
      check('点一下标题能展开', !g0.classList.contains('collapsed'));
      check('展开后能看到条目',
        g0.querySelectorAll('.body .item').length > 0);
      g0.querySelector('.head').dispatchEvent(new window.Event('click', { bubbles: true }));
      check('再点一下收回去', g0.classList.contains('collapsed'));
    }

    // 折叠的只是显示，不能影响计数和导出
    check('折叠不影响底部计数', DocCheck.barText().includes('将自动修改'), DocCheck.barText());
    check('折叠不影响问题总数', DocCheck.state.issues.length >= 5,
      '实际 ' + DocCheck.state.issues.length + ' 处');

    // ============ 六、出错时不能白屏 ============
    say('\n【异常处理】');
    DocCheck.reset();
    await DocCheck.openFile(fakeFile('垃圾.docx', Buffer.from('这不是一个 zip 文件')));
    check('坏文件不崩溃', true);
    check('坏文件时报告区是空的而不是残留上一个', DocCheck.reportHtml() === '');
    check('给了错误提示', $('#toast').textContent.includes('打不开'), '实际「' + $('#toast').textContent + '」');

    DocCheck.reset();
    await DocCheck.openFile(fakeFile('空.md', '# 标题\n\n这里没有任何问题。\n'));
    check('没问题的文件显示「没查出问题」', DocCheck.reportHtml().includes('没查出问题'));

    got.length = 0;
    $('#exportDoc').click();
    await wait(150);
    check('没勾任何东西时导出被拦住', got.length === 0 && $('#toast').textContent.includes('一处都没勾选'),
      '实际「' + $('#toast').textContent + '」');
  } catch (e) {
    fail++;
    say('  ✗ 抛异常：' + (e && e.stack || e));
  }

  say(`\n结果：${pass} 通过 / ${fail} 失败`);
  fs.writeFileSync(path.join(ROOT, '_ui_test.txt'), OUT.join('\n'), 'utf-8');
  console.log(OUT.join('\n'));
  process.exit(fail ? 1 : 0);
})();
