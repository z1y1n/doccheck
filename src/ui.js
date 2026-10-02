/* 文档体检 · 界面逻辑
   读文件 → 跑规则 → 渲染可勾选清单 → 导出。
   所有计算都在浏览器本地完成，文件不会离开这台电脑。 */
(function () {
  'use strict';

  const $ = s => document.querySelector(s);
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  // 把看不见的字符画出来，否则「插入一个空格」这种修复用户根本看不出来改了啥
  const vis = s => esc(s).replace(/\n/g, '⏎').replace(/\t/g, '⇥');

  const TIERS = {
    green: { label: '可自动修', cls: 'g', def: true },
    yellow: { label: '需你确认', cls: 'y', def: false },
    red: { label: '只能人工判断', cls: 'r', def: false }
  };
  const canFix = it => it.tier !== 'red' && it.replacement !== null;

  /* 引擎是按字符报问题的，所以「Ｅｘｃｅｌ」会变成 5 条记录。
     照实列出来就是 5 行，没人看得下去——用户的原话是「几个字母的修改都需要报一遍太复杂」。
     所以界面按「同一段、同一条规则、同一个短句之内」合并成一条来显示，
     勾选时整组一起勾，导出时再拆回逐字符（引擎那边要的仍是精确坐标）。
     代价是合并后不能再只改半条，这点得让用户知道。 */
  const gCanFix = g => canFix(g.items[0]);   // 同组同规则，看第一条就够
  const gSel = g => g.items.every(i => i.sel);
  const setGSel = (g, v) => g.items.forEach(i => { i.sel = v; });

  let S = null;   // 当前文件状态，null 表示还没打开文件

  /* ============ 一、读文件 ============ */

  async function openFile(file) {
    const name = file.name;
    try {
      if (/\.docx$/i.test(name)) {
        const zip = await JSZip.loadAsync(await file.arrayBuffer());
        const entry = zip.file('word/document.xml');
        if (!entry) throw new Error('这个 .docx 里没有正文（word/document.xml），可能是 .doc 改名来的，或文件已损坏。');
        const xml = await entry.async('string');
        const doc = DocEngine.parseDocx(xml);
        S = { name, kind: 'docx', zip, source: xml, doc };
      } else {
        const text = await file.text();
        const doc = DocEngine.parseMarkdown(text);
        S = { name, kind: 'md', source: text, doc };
      }

      S.issues = DocEngine.runRules(S.doc, RULES);
      S.issues.forEach(it => { it.sel = canFix(it) && TIERS[it.tier].def; });
      S.groups = DocEngine.groupIssues(S.issues, S.doc.paragraphs);
      render();
      toast(S.issues.length ? `查出 ${S.issues.length} 处问题` : '没发现问题，这份文档很干净');
    } catch (e) {
      toast('打不开这个文件：' + e.message, 5000);
      console.error(e);
    }
  }

  /* ============ 二、渲染 ============ */

  function render() {
    const has = !!S;
    $('#drop').hidden = has;
    $('#filebar').hidden = !has;
    $('#summary').hidden = !has;
    $('#bar').hidden = !has;
    if (!has) { $('#report').innerHTML = ''; return; }

    $('#filebar .name').textContent = S.name;
    $('#filebar .meta').textContent =
      `${S.kind === 'docx' ? 'Word 文档' : 'Markdown'}　·　${S.doc.paragraphs.length} 段　·　查出一共 ${S.issues.length} 处`;

    renderSummary();
    renderReport();
    renderBar();
  }

  function renderSummary() {
    const n = t => S.issues.filter(i => i.tier === t).length;
    const g = n('green'), y = n('yellow'), r = n('red');
    $('#summary').innerHTML =
      `<div class="stat g"><div class="n">${g}</div><div class="l">🟢 可安全自动修</div></div>` +
      `<div class="stat y"><div class="n">${y}</div><div class="l">🟡 能修，但要你点头</div></div>` +
      `<div class="stat r"><div class="n">${r}</div><div class="l">🔴 机器判断不了</div></div>`;
  }

  // 把问题在原文里的上下文画出来，问题本身高亮
  function marked(para, g, pad) {
    const d = DocEngine.groupDiff(para, g, pad);
    return vis(d.before) + '<mark>' + vis(d.original) + '</mark>' + vis(d.after);
  }

  // 「原文划掉 → 新文字」那一行
  function diffLine(para, g, pad) {
    const d = DocEngine.groupDiff(para, g, pad);
    const repl = d.replacement === '' ? '（删掉）' : vis(d.replacement).replace(/ /g, '␣');
    return `<span class="arrow">→</span>${vis(d.before)}` +
      `<mark class="del">${vis(d.original)}</mark><mark class="new">${repl}</mark>${vis(d.after)}`;
  }

  function renderReport() {
    if (!S.issues.length) {
      $('#report').innerHTML = '<div class="empty">✓ 这份文档没查出问题</div>';
      return;
    }
    let html = '';
    for (const rule of RULES) {
      const list = S.groups.filter(g => g.ruleId === rule.id);
      if (!list.length) continue;
      const tier = TIERS[rule.tier];
      const fixable = list.some(gCanFix);
      // 标题上仍然报「多少处」，和顶部三个数字、底部计数保持同一套口径。
      // 合并只改变「几行」，不改变「几处」——否则用户会以为工具漏查了。
      const units = list.reduce((a, g) => a + g.items.length, 0);

      // 🔴「只能人工判断」那类默认收起来。理由：它没有勾选框，你本来就不能对它做什么，
      // 而它往往是大头——简历里 23 条有 21 条是这个。全摊开的结果是真该改的那 2 条
      // 被淹掉。收起来只留一行「连续多个空格 21 处」，想看再点开。
      html += `<div class="group${rule.tier === 'red' ? ' collapsed' : ''}" data-rule="${rule.id}">`;
      html += `<div class="head">`;
      html += `<span class="caret">▼</span>`;
      html += fixable
        ? `<input type="checkbox" data-role="all" data-rule="${rule.id}">`
        : `<span class="nocheck" title="这类问题不给自动改">—</span>`;
      html += `<span class="title">${esc(rule.name)}</span>`;
      html += `<span class="pill ${tier.cls}">${tier.label}</span>`;
      html += `<span class="count">${units} 处</span>`;
      html += `</div>`;

      // 规则说明放在 .body 外面。放在里面的话，🔴 那类一折叠，
      // 用户就只剩一行「连续多个空格 21 处」，看不到「这条为什么只报不改」，
      // 容易以为工具坏了。说明本身讲的是规则，不是某一条问题，本来就该一直露着。
      if (rule.note) html += `<div class="note">${esc(rule.note)}</div>`;

      html += `<div class="body">`;
      list.forEach(g => {
        const idx = S.groups.indexOf(g);
        const many = g.items.length > 1;
        html += `<div class="item" data-idx="${idx}">`;
        html += gCanFix(g)
          ? `<input type="checkbox" data-role="one" data-idx="${idx}"${gSel(g) ? ' checked' : ''}>`
          : `<span class="nocheck">·</span>`;
        html += `<div class="body">`;
        const p = S.doc.paragraphs[g.pi];
        html += `<div class="where">第 ${g.pi + 1} 段` +
          (many ? `<span class="merged" title="同一句里的 ${g.items.length} 处修改合成了一条，勾上就一起改">合并 ${g.items.length} 处</span>` : '') +
          `</div>`;
        html += `<div class="line">${marked(p, g, 26)}</div>`;
        if (gCanFix(g)) html += `<div class="line">${diffLine(p, g, 26)}</div>`;
        html += `</div></div>`;
      });
      html += `</div></div>`;
    }
    $('#report').innerHTML = html;
    syncGroupBoxes();
  }

  // 分组的全选框要反映「全勾/全不勾/勾一半」
  function syncGroupBoxes() {
    document.querySelectorAll('input[data-role="all"]').forEach(box => {
      const list = S.groups.filter(g => g.ruleId === box.dataset.rule && gCanFix(g));
      const on = list.filter(gSel).length;
      box.checked = on === list.length && list.length > 0;
      box.indeterminate = on > 0 && on < list.length;
    });
  }

  function renderBar() {
    const n = S.issues.filter(i => i.sel && canFix(i)).length;
    $('#bar .tally').innerHTML = n
      ? `将自动修改 <b>${n}</b> 处　·　剩余 ${S.issues.length - n} 处留给你自己看`
      : `没有勾选任何自动修改　·　${S.issues.length} 处问题会全部列进清单`;
    $('#exportDoc').textContent = S.kind === 'docx' ? '导出修改后的 .docx' : '导出修改后的 .md';
  }

  /* ============ 三、导出 ============ */

  function currentFixes() {
    return S.issues.filter(i => i.sel && canFix(i))
      .map(i => ({ pi: i.pi, start: i.fixSpan[0], end: i.fixSpan[1], replacement: i.replacement }));
  }

  function download(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  const outName = suffix => S.name.replace(/(\.(docx|md|markdown|txt))$/i, '') + suffix;

  async function exportFixed() {
    if (!S) return;
    const fixes = currentFixes();
    if (!fixes.length) { toast('一处都没勾选，先把要自动改的勾上'); return; }
    const btn = $('#exportDoc');
    btn.disabled = true; btn.textContent = '正在生成…';
    try {
      let after;
      if (S.kind === 'docx') {
        const newXml = DocEngine.applyFixes(S.source, S.doc.paragraphs, fixes);
        // 自检：改完的 XML 必须还能解析，段落数不能变。不对就别给用户下载
        after = DocEngine.parseDocx(newXml);
        if (after.paragraphs.length !== S.doc.paragraphs.length) {
          throw new Error(`段落数从 ${S.doc.paragraphs.length} 变成了 ${after.paragraphs.length}，说明改坏了，已中止`);
        }
        S.zip.file('word/document.xml', newXml);
        const blob = await S.zip.generateAsync({
          type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 }
        });
        download(blob, outName('-已修改.docx'));
      } else {
        const out = DocEngine.applyTextFixes(S.source, S.doc.paragraphs, fixes);
        after = DocEngine.parseMarkdown(out);
        download(new Blob([out], { type: 'text/markdown;charset=utf-8' }), outName('-已修改.md'));
      }

      // 所有规则都跑在原文上，不会级联。所以改完之后可能又冒出新的问题——
      // 比如全角「Ｑ」转成半角「Q」之后，中文和英文就挨上了，但这一遍扫不到。
      // 与其让用户以为一次就干净了，不如直说。
      const again = DocEngine.runRules(after, RULES).filter(canFix).length;
      if (again) {
        toast(`已修改 ${fixes.length} 处。改完又出现 ${again} 处新问题，把导出的文件再查一遍就能清掉`, 7000);
      } else {
        toast(`已修改 ${fixes.length} 处，文件已下载`);
      }
    } catch (e) {
      toast('导出失败：' + e.message, 6000);
      console.error(e);
    } finally {
      btn.disabled = false;
      renderBar();
    }
  }

  function todoText() {
    return DocEngine.buildTodo({
      fileName: S.name,
      issues: S.issues,
      paragraphs: S.doc.paragraphs,
      rules: RULES,
      fixedCount: S.issues.filter(i => i.sel && canFix(i)).length
    });
  }

  function exportTodo() {
    if (!S) return;
    // 加个 BOM，用记事本打开也不会乱码
    download(new Blob(['﻿' + todoText()], { type: 'text/markdown;charset=utf-8' }), outName('-待处理清单.md'));
    toast('清单已下载');
  }

  /* ============ 四、交互 ============ */

  let toastTimer;
  function toast(msg, ms) {
    const el = $('#toast');
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.remove('show'), ms || 2600);
  }

  function reset() { S = null; $('#file').value = ''; render(); }

  // 勾选框
  $('#report').addEventListener('change', e => {
    const t = e.target;
    if (t.dataset.role === 'one') {
      // 一行可能压着好几处修改，一起勾、一起改
      setGSel(S.groups[+t.dataset.idx], t.checked);
      syncGroupBoxes(); renderBar();
    } else if (t.dataset.role === 'all') {
      S.groups.forEach(g => { if (g.ruleId === t.dataset.rule && gCanFix(g)) setGSel(g, t.checked); });
      document.querySelectorAll(`input[data-role="one"]`).forEach(b => {
        if (S.groups[+b.dataset.idx].ruleId === t.dataset.rule) b.checked = t.checked;
      });
      syncGroupBoxes(); renderBar();
    }
  });

  // 点分组标题折叠；点勾选框不要跟着折叠
  $('#report').addEventListener('click', e => {
    if (e.target.tagName === 'INPUT') return;
    const head = e.target.closest('.head');
    if (head) head.parentElement.classList.toggle('collapsed');
  });

  // 选文件。文件框藏在拖放区里面，它自己弹出的 click 会冒泡回来，
  // 不判断一下就会无限递归地把选择框开个不停。
  $('#drop').addEventListener('click', e => { if (e.target.id !== 'file') $('#file').click(); });
  $('#file').addEventListener('change', e => { if (e.target.files[0]) openFile(e.target.files[0]); });
  $('#reload').addEventListener('click', reset);

  // 拖放（整页都能接）
  ['dragenter', 'dragover'].forEach(ev => document.addEventListener(ev, e => {
    e.preventDefault(); $('#drop').classList.add('over');
  }));
  ['dragleave', 'drop'].forEach(ev => document.addEventListener(ev, e => {
    e.preventDefault();
    if (ev === 'dragleave' && e.relatedTarget) return;
    $('#drop').classList.remove('over');
  }));
  document.addEventListener('drop', e => {
    const f = e.dataTransfer && e.dataTransfer.files[0];
    if (f) openFile(f);
  });

  $('#exportDoc').addEventListener('click', exportFixed);
  $('#exportMd').addEventListener('click', exportTodo);

  // 留给自动化测试的口子（也方便以后写命令行版）
  window.DocCheck = {
    openFile, reset, currentFixes, todoText,
    reportHtml: () => $('#report').innerHTML,
    barText: () => $('#bar .tally').textContent,
    fileInfo: () => ({ name: $('#filebar .name').textContent, meta: $('#filebar .meta').textContent }),
    get state() { return S; }
  };
})();
