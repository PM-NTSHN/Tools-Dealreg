// Kiểm thử nhanh tool Dealreg bằng Playwright (headless Chromium).
// Nạp file Excel mẫu → tìm kiếm → lọc Sale → mở chi tiết → xuất Excel; ghi lỗi console & request mạng.
//
//   LANG=C.UTF-8 node scripts/smoke_test.js "<tool.html>" "<raw.xlsx>" <out_dir> [playwright_module_path]
//
// LANG=C.UTF-8 cần thiết trong container: thiếu locale, Chromium đổi tên file tải về có dấu thành "download".
const path = require('path');
const [html, xlsx, outDir = '.', pwPath] = process.argv.slice(2);
const { chromium } = require(pwPath || 'playwright');
(async () => {
  const b = await chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {});
  const ctx = await b.newContext({ viewport: { width: 1600, height: 1000 }, acceptDownloads: true });
  const p = await ctx.newPage();
  const errs = [];
  p.on('console', m => { if (m.type() === 'error') errs.push('console: ' + m.text()); });
  p.on('pageerror', e => errs.push('pageerror: ' + e.message));
  p.on('request', r => { if (!/^(file|data|blob):/.test(r.url())) errs.push('NETWORK ' + r.url()); });
  await p.goto('file://' + path.resolve(html));
  await p.setInputFiles('#file-input', xlsx);
  await p.waitForTimeout(1500);
  await p.screenshot({ path: path.join(outDir, 'full.png'), fullPage: true });
  const n = await p.evaluate(() => DealregTool.S.rows.length);
  console.log('Số phiếu nạp:', n);
  await p.fill('#q', 'a'); await p.waitForTimeout(400);
  console.log('Tìm "a":', await p.evaluate(() => DealregTool.S.view.length));
  await p.fill('#q', ''); await p.waitForTimeout(300);
  await p.click('[data-facet="sale"] .ms-input'); await p.keyboard.press('ArrowDown'); await p.keyboard.press('Enter'); await p.waitForTimeout(300);
  console.log('Lọc Sale:', await p.evaluate(() => [DealregTool.S.view.length, [...DealregTool.S.f.sale]]));
  await p.mouse.click(5, 5);
  await p.click('#grid tbody tr >> nth=0'); await p.waitForTimeout(400);
  await p.screenshot({ path: path.join(outDir, 'detail.png') });
  await p.keyboard.press('Escape');
  const [dl] = await Promise.all([p.waitForEvent('download'), p.click('#btn-export')]);
  const f = path.join(outDir, dl.suggestedFilename()); await dl.saveAs(f);
  console.log('Đã xuất:', f);
  console.log(errs.length ? 'LỖI:\n' + errs.join('\n') : 'Không lỗi, không request mạng.');
  await b.close();
  process.exit(errs.length ? 1 : 0);
})();
