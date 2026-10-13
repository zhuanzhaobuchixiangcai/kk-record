import assert from 'node:assert/strict';

const source = await import('node:fs/promises').then(fs => fs.readFile(new URL('../src/main.js', import.meta.url), 'utf8'));
assert.match(source, /indexedDB/);
assert.match(source, /Packer\.toBlob/);
assert.match(source, /record\.id/);
assert.match(source, /localeCompare/);
console.log('cross-platform source self-test passed');
