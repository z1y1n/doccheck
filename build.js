/* 打包：把 index.html + jszip + engine + rules + ui 合成一个单文件 HTML。
   开发时直接开 src/index.html 就行，改完规则想发给别人时再跑这个。 */
const fs = require('fs');
const path = require('path');

const SRC = path.join(__dirname, 'src');
const read = p => fs.readFileSync(p, 'utf-8');

const html = read(path.join(SRC, 'index.html'));
const parts = {
  '../node_modules/jszip/dist/jszip.min.js': path.join(__dirname, 'node_modules', 'jszip', 'dist', 'jszip.min.js'),
  'engine.js': path.join(SRC, 'engine.js'),
  'rules.js': path.join(SRC, 'rules.js'),
  'ui.js': path.join(SRC, 'ui.js')
};

let out = html;
for (const [ref, file] of Object.entries(parts)) {
  if (!fs.existsSync(file)) {
    console.error('缺文件：' + file + '（jszip 要先 npm install）');
    process.exit(1);
  }
  const tag = new RegExp(`<script src="[^"]*${ref.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"><\\/script>`);
  if (!tag.test(out)) {
    console.error('index.html 里找不到引用：' + ref);
    process.exit(1);
  }
  out = out.replace(tag, () => '<script>\n' + read(file) + '\n</script>');
}

// 单文件版没有同目录的 CSS/JS 依赖了，把标题也点明用途
out = out.replace('<title>', '<title>单文件版 · ');
const target = path.join(__dirname, '文档体检.html');
fs.writeFileSync(target, out, 'utf-8');

const kb = (Buffer.byteLength(out, 'utf-8') / 1024).toFixed(0);
console.log(`已生成 文档体检.html（${kb} KB）—— 双击就能用，不用装任何东西`);
