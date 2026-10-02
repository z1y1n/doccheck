/* 文档体检 · 核心引擎
   把 docx 拆成可检查的文本模型 → 跑规则 → 应用修复 → 重新打包。
   不碰界面，Node 和浏览器共用同一份。 */

const DocEngine = (function () {

  // ========== 一、解析 docx ==========

  // 把 document.xml 切成段落。返回 [{start, end, inTable, style}]
  function splitParagraphs(xml) {
    const out = [];
    const re = /<w:p(?:\s[^>]*)?\/>|<w:p(?:\s[^>]*)?>|<\/w:p>|<w:tbl(?:\s[^>]*)?>|<\/w:tbl>/g;
    const stack = [];
    let tbl = 0, m;
    while ((m = re.exec(xml)) !== null) {
      const tag = m[0];
      if (tag.startsWith('<w:tbl')) { tbl += tag.endsWith('/>') ? 0 : 1; continue; }
      if (tag === '</w:tbl>') { tbl = Math.max(0, tbl - 1); continue; }
      if (tag.endsWith('/>')) { out.push({ start: m.index, end: m.index + tag.length, inTable: tbl > 0 }); continue; }
      if (tag === '</w:p>') {
        const st = stack.pop();
        if (st !== undefined) out.push({ start: st, end: m.index + tag.length, inTable: tbl > 0 });
      } else stack.push(m.index);
    }
    out.sort((a, b) => a.start - b.start);
    for (const p of out) {
      const sm = /<w:pStyle\s+w:val="([^"]*)"/.exec(xml.slice(p.start, p.end));
      p.style = sm ? sm[1] : null;
    }
    return out;
  }

  // 找出被 <w:del>（修订删除）包住的区间，这些文字在 Word 里不显示，不能检查
  function findDeletedRanges(xml) {
    const ranges = [];
    const re = /<w:del(?:\s[^>]*)?>|<\/w:del>/g;
    const stack = [];
    let m;
    while ((m = re.exec(xml)) !== null) {
      if (m[0] === '</w:del>') { const st = stack.pop(); if (st !== undefined) ranges.push([st, m.index]); }
      else stack.push(m.index);
    }
    return ranges;
  }

  const inRanges = (pos, ranges) => ranges.some(r => pos >= r[0] && pos < r[1]);

  // 解析单个段落，生成扁平文本 + 字符到节点的映射
  function parseParagraph(xml, para, delRanges) {
    const seg = xml.slice(para.start, para.end);
    const items = [];

    const tRe = /<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
    let m;
    while ((m = tRe.exec(seg)) !== null) {
      const tagStart = para.start + m.index;
      if (inRanges(tagStart, delRanges)) continue;
      const attrs = m[1] || '';
      const contentStart = tagStart + 4 + attrs.length + 1;
      items.push({
        kind: 't', pos: tagStart, attrs,
        tagStart, contentStart, contentEnd: contentStart + m[2].length,
        content: m[2]
      });
    }

    const xRe = /<w:tab(?:\s[^>]*)?\/>|<w:br(?:\s[^>]*)?\/>|<w:cr(?:\s[^>]*)?\/>/g;
    while ((m = xRe.exec(seg)) !== null) {
      const pos = para.start + m.index;
      if (inRanges(pos, delRanges)) continue;
      items.push({ kind: m[0].startsWith('<w:tab') ? 'tab' : 'br', pos });
    }

    items.sort((a, b) => a.pos - b.pos);

    let text = '';
    const map = [];
    const nodes = [];
    for (const it of items) {
      if (it.kind === 't') {
        const ni = nodes.length;
        it.cStart = text.length;
        it.cEnd = it.cStart + it.content.length;
        nodes.push(it);
        for (let k = 0; k < it.content.length; k++) map.push({ n: ni, o: k });
        text += it.content;
      } else {
        const ch = it.kind === 'tab' ? '\t' : '\n';
        map.push({ n: -1, o: 0 });
        text += ch;
      }
    }
    return { ...para, text, map, nodes };
  }

  function parseDocx(xml) {
    const delRanges = findDeletedRanges(xml);
    const paras = splitParagraphs(xml).map(p => parseParagraph(xml, p, delRanges));
    return { xml, paragraphs: paras };
  }

  // ========== 二、跑规则 ==========

  // 把 docx 解析 + 逐段跑规则，返回问题列表
  function runRules(doc, rules) {
    const issues = [];
    doc.paragraphs.forEach((para, pi) => {
      if (!para.text.trim()) return;
      // checkText 是「遮罩文本」：长度和正文完全一样，但把不该检查的地方（代码块、链接地址）
      // 换成了占位符。这样规则匹配时不会碰到它们，而位置坐标仍然能对上正文。
      const target = para.checkText !== undefined ? para.checkText : para.text;
      for (const rule of rules) {
        const re = new RegExp(rule.pattern.source, rule.pattern.flags.includes('g') ? rule.pattern.flags : rule.pattern.flags + 'g');
        let m;
        while ((m = re.exec(target)) !== null) {
          if (m[0].length === 0) { re.lastIndex++; continue; }
          const ctx = { text: target, index: m.index, match: m, para, pi };
          if (rule.filter && !rule.filter(ctx)) continue;
          const r = rule.fix ? rule.fix(ctx) : null;
          issues.push({
            ruleId: rule.id, ruleName: rule.name, tier: rule.tier,
            pi, start: m.index, end: m.index + m[0].length,
            original: m[0],
            replacement: r ? r.text : null,
            fixSpan: r ? [r.from !== undefined ? r.from : m.index, r.to !== undefined ? r.to : m.index + m[0].length] : null,
            note: rule.note || ''
          });
        }
      }
    });
    return dedupeOverlaps(issues);
  }

  // 不同规则可能命中同一段文字，保留优先级高的（绿 > 黄 > 红，即先能自动修的）
  function dedupeOverlaps(issues) {
    const rank = { green: 0, yellow: 1, red: 2 };
    const sorted = [...issues].sort((a, b) => (rank[a.tier] - rank[b.tier]) || (a.start - b.start));
    const kept = [];
    for (const it of sorted) {
      if (kept.some(k => k.pi === it.pi && it.start < k.end && k.start < it.end)) continue;
      kept.push(it);
    }
    return kept.sort((a, b) => (a.pi - b.pi) || (a.start - b.start));
  }

  // ========== 三、应用修复 ==========

  // 算出要对哪些文本节点动手。零宽插入（a===b）单独处理：优先插在节点内部，其次节点末尾，最后节点开头
  function planEdits(para, a, b) {
    const plan = [];
    if (a === b) {
      for (const node of para.nodes) {
        if (a > node.cStart && a < node.cEnd) { plan.push({ node, ns: a - node.cStart, ne: a - node.cStart }); return plan; }
      }
      for (const node of para.nodes) {
        if (node.cEnd === a && node.cStart < a) { plan.push({ node, ns: node.content.length, ne: node.content.length }); return plan; }
      }
      for (const node of para.nodes) {
        if (node.cStart === a && node.cEnd > a) { plan.push({ node, ns: 0, ne: 0 }); return plan; }
      }
      return plan;
    }
    for (const node of para.nodes) {
      if (node.cEnd <= a || node.cStart >= b) continue;
      plan.push({
        node,
        ns: Math.max(a, node.cStart) - node.cStart,
        ne: Math.min(b, node.cEnd) - node.cStart
      });
    }
    return plan;
  }

  // fixes: [{pi, start, end, replacement}]，坐标是段落扁平文本上的位置
  function applyFixes(xml, paragraphs, fixes) {
    // 按段落归拢。这里有个必须踩过的坑：同一个 <w:t> 里可能同时有好几处要改
    // （一句话里既要补几个空格、又要合并重复标点）。如果每处修复各自「整节点重写」，
    // 每次都拿原始内容算，后一次就把前一次整个覆盖掉了——而且总长度可能刚好抵消，
    // 连长度校验都看不出来。所以先按节点把操作合并，再一次性重写这个节点。
    const byPara = new Map();
    for (const f of fixes) {
      if (!byPara.has(f.pi)) byPara.set(f.pi, []);
      byPara.get(f.pi).push(f);
    }

    const rewrites = [];   // {node, content}
    for (const [pi, list] of byPara) {
      const para = paragraphs[pi];
      if (!para) continue;

      const opsByNode = new Map();
      for (const f of list) {
        planEdits(para, f.start, f.end).forEach((step, i) => {
          if (!opsByNode.has(step.node)) opsByNode.set(step.node, []);
          opsByNode.get(step.node).push({ s: step.ns, e: step.ne, text: i === 0 ? f.replacement : '' });
        });
      }

      for (const [node, ops] of opsByNode) {
        // 用游标在原文上从左往右走一遍，把每处操作按位置插进去，
        // 这样各处的坐标都还是原文的坐标，不会互相打乱
        ops.sort((a, b) => a.s - b.s || a.e - b.e);
        let content = '', cursor = 0;
        for (const op of ops) {
          if (op.s < cursor) continue;   // 位置重叠的（去重后基本不会出现）直接跳过
          content += node.content.slice(cursor, op.s) + op.text;
          cursor = op.e;
        }
        content += node.content.slice(cursor);
        rewrites.push({ node, content });
      }
    }

    const edits = [];
    for (const { node, content } of rewrites) {
      edits.push({ start: node.contentStart, end: node.contentEnd, text: content });
      // 内容带首尾空格时必须打开 preserve，否则 Word 打开时会把空格吃掉
      if (/^\s|\s$/.test(content) && !/xml:space/.test(node.attrs)) {
        edits.push({ start: node.tagStart + 4, end: node.tagStart + 4, text: ' xml:space="preserve"' });
      }
    }

    edits.sort((a, b) => b.start - a.start || b.end - a.end);
    let out = xml;
    for (const e of edits) out = out.slice(0, e.start) + e.text + out.slice(e.end);
    return out;
  }

  // ========== 四、生成报告产物 ==========

  const visible = s => s.replace(/\n/g, '⏎').replace(/\t/g, '⇥');

  // ========== 四之前：把挨在一起的问题归成一组 ==========

  // 两处问题之间出现这些字符就断开，不合并。
  // 也就是「同一个短句之内合并，跨短句不合并」——「Ａ、Ｂ」这种顿号不算断句。
  const CLAUSE_BREAK = /[，。！？；：\n\t]/;

  // 同一段、同一条规则、中间只隔着普通文字的一串问题，合成一条。
  // 为什么需要：引擎是按字符报的，「Ｅｘｃｅｌ」会变成 5 条记录，界面照实列就是 5 行，
  // 没人看得下去。但引擎本身不能改——它需要逐字符的坐标才能精确改写 XML。
  // 所以归组只发生在「给人看」这一层，导出时再拆回逐字符。
  function groupIssues(issues, paragraphs) {
    const sorted = [...issues].sort((a, b) => (a.pi - b.pi) || (a.start - b.start));
    const groups = [];
    let cur = null;
    for (const it of sorted) {
      if (cur && cur.pi === it.pi && cur.ruleId === it.ruleId) {
        const gap = paragraphs[it.pi].text.slice(cur.end, it.start);
        if (!CLAUSE_BREAK.test(gap)) {
          cur.end = it.end;
          cur.items.push(it);
          continue;
        }
      }
      cur = { pi: it.pi, ruleId: it.ruleId, start: it.start, end: it.end, items: [it] };
      groups.push(cur);
    }
    return groups;
  }

  const asGroup = it => ({ pi: it.pi, ruleId: it.ruleId, start: it.start, end: it.end, items: [it] });

  // 把一个组拆成「前文 + 原文 + 改后文字 + 后文」四块。
  // 网页要把原文和改后分别上色，Markdown 报告要拼成一行文字——所以这里只给零件，不给成品。
  function groupDiff(para, g, pad) {
    const t = para.text;
    const n = pad === undefined ? 26 : pad;
    const a = Math.max(0, g.start - n);
    const b = Math.min(t.length, g.end + n);

    // 「改后」不能拿整段直接换成 replacement——每处修改只吃掉自己 fixSpan 那几个字符。
    // 「行,」这类命中范围比实际要改的范围大（真正动的只有那个逗号），
    // 整段替换会把「行」也一起吃掉。
    let fixed = '', cursor = g.start;
    for (const it of g.items) {
      const span = it.fixSpan || [it.start, it.end];
      if (span[0] < cursor) continue;
      fixed += t.slice(cursor, span[0]) +
        (it.replacement === null ? t.slice(span[0], span[1]) : it.replacement);
      cursor = span[1];
    }
    fixed += t.slice(cursor, g.end);

    return {
      before: (a > 0 ? '…' : '') + t.slice(a, g.start),
      original: t.slice(g.start, g.end),
      replacement: fixed,
      after: t.slice(g.end, b) + (b < t.length ? '…' : ''),
      count: g.items.length
    };
  }

  // 单个问题的版本，内部就是只有一项的组
  function diffOf(para, issue, pad) {
    return groupDiff(para, asGroup(issue), pad);
  }

  // 这两个既收单条问题、也收一整组：传进来的是组就照组渲染，是单条就当成只有一项的组。
  // 让调用方不用关心自己手上拿的是哪种。
  const asAnyGroup = x => (x.items ? x : asGroup(x));

  // 一行纯文本，方便用户在 Word 里 Ctrl+F 搜到
  function contextOf(para, item, pad) {
    const d = groupDiff(para, asAnyGroup(item), pad === undefined ? 14 : pad);
    return d.before + '〖' + d.original + '〗' + d.after;
  }

  // 改动后长什么样：〖新文字〗
  function fixedTextOf(para, item, pad) {
    const d = groupDiff(para, asAnyGroup(item), pad);
    return d.before + '〖' + d.replacement + '〗' + d.after;
  }

  function buildMarkdownReport(fileName, issues, rules, paragraphs) {
    const byRule = new Map();
    for (const g of groupIssues(issues, paragraphs)) {
      if (!byRule.has(g.ruleId)) byRule.set(g.ruleId, []);
      byRule.get(g.ruleId).push(g);
    }
    const tierLabel = { green: '可自动修', yellow: '需你确认', red: '只能人工判断' };
    let s = `# 文档体检报告\n\n`;
    s += `- **文件：** ${fileName}\n- **时间：** ${new Date().toLocaleString('zh-CN')}\n- **问题总数：** ${issues.length} 处\n\n`;
    s += `> 把下面每一行的「原文」复制到 Word 里 Ctrl+F 搜索，就能定位到问题位置。\n\n---\n\n`;
    for (const rule of rules) {
      const list = byRule.get(rule.id);
      if (!list) continue;
      const n = list.reduce((a, g) => a + g.items.length, 0);
      s += `## ${rule.name}　｜${tierLabel[rule.tier]}｜　${n} 处\n\n`;
      if (rule.note) s += `> ${rule.note}\n\n`;
      for (const g of list) {
        const p = paragraphs[g.pi];
        const merged = g.items.length > 1 ? `（合并 ${g.items.length} 处）` : '';
        s += `- **第 ${g.pi + 1} 段**${merged}　${visible(contextOf(p, g, 14))}\n`;
        if (g.items.some(it => it.replacement !== null)) {
          s += `  - 改后：${visible(fixedTextOf(p, g, 14))}\n`;
        }
      }
      s += `\n`;
    }
    return s;
  }

  // 导出的「待处理清单」：把没自动改的问题整理成能照着手动改的一份 Markdown
  function buildTodo(opts) {
    const { fileName, issues, paragraphs, fixedCount, rules } = opts;
    const canFix = it => it.tier !== 'red' && it.replacement !== null;
    // 这里也按短句合并，理由和网页上一样：清单是给人照着改的，
    // 「Ｅｘｃｅｌ」拆成 5 行没人愿意看。归组只按同类规则，所以一组里的档位必然相同，
    // 拿第一条判断能不能自动改就够了。
    const all = groupIssues(issues, paragraphs);
    const skipped = all.filter(g => canFix(g.items[0]) && !g.items[0].sel);
    const manual = all.filter(g => !canFix(g.items[0]));
    const count = gs => gs.reduce((a, g) => a + g.items.length, 0);

    let s = `# 文档体检 · 待处理清单\n\n`;
    s += `- **原文件：** ${fileName}\n`;
    s += `- **导出时间：** ${new Date().toLocaleString('zh-CN')}\n`;
    s += `- **已自动修改：** ${fixedCount} 处（在「已修改」那个文件里）\n`;
    s += `- **待你处理：** ${count(skipped) + count(manual)} 处\n\n`;
    s += `> 每一行的「原文」都可以直接复制到文档里 Ctrl+F 搜索定位。\n\n---\n\n`;

    const section = (title, desc, list) => {
      if (!list.length) return '';
      let out = `## ${title}　${count(list)} 处\n\n${desc}\n\n`;
      for (const rule of rules) {
        const mine = list.filter(g => g.ruleId === rule.id);
        if (!mine.length) continue;
        out += `### ${rule.name}（${count(mine)} 处）\n\n`;
        if (rule.note) out += `> ${rule.note}\n\n`;
        for (const g of mine) {
          const p = paragraphs[g.pi];
          const merged = g.items.length > 1 ? `（合并 ${g.items.length} 处）` : '';
          out += `- **第 ${g.pi + 1} 段**${merged}　${visible(contextOf(p, g, 26))}\n`;
          if (canFix(g.items[0])) out += `  - 建议改成：${visible(fixedTextOf(p, g, 26))}\n`;
        }
        out += `\n`;
      }
      return out + `---\n\n`;
    };

    s += section('一、你选择不自动改的',
      '这些机器其实能改，但你没勾选。确认没问题的话，手动改掉就行。', skipped);
    s += section('二、机器判断不了，只能你自己看的',
      '这类问题没有统一答案——可能是刻意的排版，也可能是疏漏。工具只负责把它们指出来。', manual);
    if (!skipped.length && !manual.length) s += `全部问题都已自动修改，这份清单是空的。\n`;
    return s;
  }


  // ========== 五、Markdown 支持 ==========
  // 思路和 docx 一样：把源文切成「段落」，每段附带一份等长的遮罩文本。
  // 遮罩把代码、链接地址这些不能动的地方换成占位符，规则就碰不到它们了。

  // 生成等长遮罩：反引号代码、链接地址、裸网址、邮箱 → 全换成 \u0000
  function maskInline(text) {
    const arr = text.split('');
    const zones = [];
    let m;

    // 整个匹配都要遮住的
    const wholeRe = [
      /`[^`\n]*`/g,                                       // 行内代码
      /(?:https?:\/\/|www\.)[^\s<>)\]]+/g,                // 裸网址
      /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g   // 邮箱
    ];
    for (const re of wholeRe) {
      while ((m = re.exec(text)) !== null) zones.push([m.index, m.index + m[0].length]);
    }

    // 只遮地址部分，[链接文字](地址) 里的文字照常检查
    const linkRe = /\]\(([^)\n]*)\)/g;
    while ((m = linkRe.exec(text)) !== null) zones.push([m.index + 2, m.index + 2 + m[1].length]);

    for (const [a, b] of zones) for (let k = a; k < b && k < arr.length; k++) arr[k] = '\u0000';
    return arr.join('');
  }

  // 把 md 切成段落：空行分段，``` 围起来的代码块整块跳过
  function parseMarkdown(src) {
    const paras = [];
    let inFence = false, blockStart = 0, lines = [], blockIsCode = false;

    const flush = () => {
      if (!lines.length) return;
      const text = lines.join('\n');
      paras.push({
        start: blockStart, end: blockStart + text.length,
        text, checkText: blockIsCode ? '\u0000'.repeat(text.length) : maskInline(text),
        inTable: false, style: null, isCode: blockIsCode
      });
      lines = []; blockIsCode = false;
    };

    let lineStart = 0;
    for (;;) {
      const nl = src.indexOf('\n', lineStart);
      const end = nl === -1 ? src.length : nl;
      const line = src.slice(lineStart, end);
      const fence = /^\s*(?:```|~~~)/.test(line);

      if (fence && !inFence) { flush(); inFence = true; blockStart = lineStart; blockIsCode = true; lines.push(line); }
      else if (fence && inFence) { lines.push(line); inFence = false; flush(); }
      else if (inFence) lines.push(line);
      else if (!line.trim()) flush();
      else { if (!lines.length) blockStart = lineStart; lines.push(line); }

      if (nl === -1) break;
      lineStart = nl + 1;
    }
    flush();
    return { xml: src, paragraphs: paras, isMarkdown: true };
  }

  // 纯文本版的 applyFixes：段落本身就是源文的子串，直接按全局坐标改
  function applyTextFixes(text, paragraphs, fixes) {
    const edits = [];
    for (const f of fixes) {
      const para = paragraphs[f.pi];
      if (!para) continue;
      edits.push({ start: para.start + f.start, end: para.start + f.end, text: f.replacement });
    }
    edits.sort((a, b) => b.start - a.start || b.end - a.end);
    let out = text;
    for (const e of edits) out = out.slice(0, e.start) + e.text + out.slice(e.end);
    return out;
  }

  return {
    parseDocx, parseMarkdown, runRules, applyFixes, applyTextFixes,
    diffOf, groupDiff, groupIssues, asAnyGroup,
    contextOf, fixedTextOf, buildMarkdownReport, buildTodo, splitParagraphs
  };
})();

if (typeof module !== 'undefined') module.exports = DocEngine;
