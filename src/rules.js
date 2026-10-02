/* 文档体检 · 规则库
   ------------------------------------------------------------------
   这是整套工具里你最该动手改的文件。加一条规则只需要照抄一份下面的结构。
   每条规则的字段：
     id      内部编号，别改
     name    报告里显示的名字
     tier    green  = 改错了也几乎不可能出问题（默认勾选自动修）
             yellow = 大概率对，但有例外，需要你自己拍板（默认不勾）
             red    = 机器判断不了，只能给你指出位置（不给勾选框）
     pattern 匹配用的正则
     note    一句人话，说明这条在查什么、为什么
     filter  （可选）排除例外的函数，返回 false 表示「这条不算问题」
     fix     （可选）返回 {text, from, to} 表示怎么改；不写就表示不能自动修
   ------------------------------------------------------------------ */

const RULES = [

  // ==================== 🟢 可安全自动修 ====================

  {
    id: 'fullwidth-alnum',
    name: '全角字母 / 数字',
    tier: 'green',
    pattern: /[Ａ-Ｚａ-ｚ０-９]/g,
    note: '全角字母数字只用于特殊排版，混在正文里会忽宽忽窄。改成半角。',
    fix: ctx => ({ text: String.fromCharCode(ctx.match[0].charCodeAt(0) - 0xFEE0) })
  },

  {
    id: 'dup-punct',
    name: '重复的中文标点',
    tier: 'green',
    pattern: /([，。！？；：、])\1+/g,
    note: '中文标点连打多个（「。。」「，，」）通常是手滑。注意「……」和「——」是正确写法，不会被报。',
    fix: ctx => ({ text: ctx.match[1] })
  },

  // ==================== 🟡 能修，但要你确认 ====================

  {
    id: 'halfwidth-punct',
    name: '中文里混用半角标点',
    tier: 'yellow',
    pattern: /[一-鿿、-〿＀-￯][,.;:?!]/g,
    note: '中文句子里的逗号、句号、问号应该用全角（，。？！）。已排除小数、版本号、英文缩写、网址邮箱这些情况。',
    filter: ctx => {
      const next = ctx.text[ctx.index + 2];
      if (next !== undefined && /[A-Za-z0-9]/.test(next)) return false;
      return true;
    },
    fix: ctx => {
      const table = { ',': '，', '.': '。', ';': '；', ':': '：', '?': '？', '!': '！' };
      return { from: ctx.index + 1, to: ctx.index + 2, text: table[ctx.match[0][1]] };
    }
  },

  {
    id: 'cjk-latin-space',
    name: '中文与英文之间缺空格',
    tier: 'yellow',
    // 用前瞻而不是直接匹配两个字符：正则全局扫描不会找重叠的匹配，
    // 「QGIS和SQL」这种夹心结构里，`S和` 一旦匹配就把 S 吃掉了，
    // 紧跟着的 `和S` 就永远轮不到，结果补出来是「QGIS 和SQL」。
    pattern: /[一-鿿](?=[A-Za-z])|[A-Za-z](?=[一-鿿])/g,
    note: '中文和英文之间加一个空格，排版会透气很多。全角标点（，。：（））自带间距，两侧不用加——所以「：PRD」「（QGIS）」不会被报。',
    filter: ctx => {
      const near = ctx.text.slice(Math.max(0, ctx.index - 20), ctx.index + 22);
      // 网址、邮箱附近不动，怕把链接改坏
      if (/@|[a-z0-9-]\.(com|cn|net|org|edu|gov)/i.test(near)) return false;
      return true;
    },
    fix: ctx => ({ from: ctx.index + 1, to: ctx.index + 1, text: ' ' })
  },

  // ==================== 🔴 只报不改 ====================

  {
    id: 'ideographic-space',
    name: '全角空格',
    tier: 'red',
    pattern: /　/g,
    note: '全角空格要专门切输入法才打得出来，基本都是有意的——「第一部分　小标题」这种分隔、段首缩进，都是正常用法。' +
      '所以这条只报不改。真要在中英文之间留空隙，应该用普通空格。',
    // 这条原本是 🟢 自动删，拿真实文档一试，7 处「第X部分　标题」的分隔符全被误删。
    // 改成 🟡 加过滤条件后又误报列表缩进，再改还是分不清「刻意的分隔」和「该用普通空格」。
    // 规则调两轮还调不准，就不该继续调——降级成只报不改，把判断权还给人。
    filter: ctx => ctx.text.slice(0, ctx.index).trim() !== ''  // 段首缩进不报
  },

  {
    id: 'multi-space',
    name: '连续多个空格',
    tier: 'red',
    pattern: / {2,}/g,
    note: '连续空格经常是刻意用来对齐的（比如联系方式那一行的「|」两侧）。删掉会让排版塌掉，所以这条只指出位置、不提供自动修复，你自己看着办。',
    filter: ctx => ctx.text.slice(0, ctx.index).trim() !== ''
  }

];

if (typeof module !== 'undefined') module.exports = RULES;
