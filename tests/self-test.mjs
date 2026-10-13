import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { Document, Packer, Paragraph, Table, TableRow, TableCell } from 'docx';

const require = createRequire(import.meta.url);
const { saveRecords } = require('../electron/wordStore.cjs');

const source = await import('node:fs/promises').then(fs => fs.readFile(new URL('../src/main.js', import.meta.url), 'utf8'));
assert.match(source, /indexedDB/);
assert.match(source, /Packer\.toBlob/);
assert.match(source, /record\.id/);
assert.match(source, /localeCompare/);

const headers = ['日期', '领域', '发生了什么', '我的判断', '证据/原话', '下一步', '结果'];
const row = (values) => new TableRow({ children: values.map(value => new TableCell({ children: [new Paragraph(value)] })) });
const fixture = new Document({ sections: [{ children: [new Table(headers.map(h => row([h])))] }] });
const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'kk-self-test-'));
const docPath = path.join(tempDir, 'KK个人成长系统.docx');
await fs.writeFile(docPath, await Packer.toBuffer(fixture));
const result = await saveRecords(docPath, [
  { id: '1', type: '每日或事件记录', values: { 日期: '2099-10-03', 领域: '工作', 发生了什么: 'three', 我的判断: '', '证据/原话': '', 下一步: '', 结果: '' } },
  { id: '2', type: '每日或事件记录', values: { 日期: '2099-10-02', 领域: '工作', 发生了什么: 'two', 我的判断: '', '证据/原话': '', 下一步: '', 结果: '' } }
]);
assert.equal(result.added, 2);
assert.equal((await fs.readdir(path.join(tempDir, 'backups'))).length, 1);
assert.equal((await fs.readFile(`${docPath}.kk-index.json`, 'utf8')).includes('"1"'), true);
console.log('cross-platform source self-test passed');
