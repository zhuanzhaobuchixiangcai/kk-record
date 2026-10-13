import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';

const output = path.resolve(process.argv[2] || 'dist');
const base = new URL('https://example.com/kk-record/');
const html = await fs.readFile(path.join(output, 'index.html'), 'utf8');
const manifestHref = html.match(/<link\s+rel="manifest"\s+href="([^"]+)"/)[1];
const manifestUrl = new URL(manifestHref, base);
const manifestPath = manifestUrl.pathname.slice(base.pathname.length);
const manifest = JSON.parse(await fs.readFile(path.join(output, manifestPath), 'utf8'));
assert.equal(new URL(manifest.start_url, manifestUrl).href, base.href, '主屏幕图标必须打开应用首页');
for (const icon of manifest.icons) {
  const iconUrl = new URL(icon.src, manifestUrl);
  assert.ok(iconUrl.pathname.startsWith(base.pathname));
  await fs.access(path.join(output, iconUrl.pathname.slice(base.pathname.length)));
}
console.log('PWA manifest opens the correct site and resolves its icons');
