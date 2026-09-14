'use strict';

const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { pipePaths } = require('./index.js');

if (process.platform !== 'win32') {
  const previous = process.env.TMPDIR;
  try {
    process.env.TMPDIR = `/${'nested/'.repeat(20)}tmp`;
    const [main, ping] = pipePaths('/tmp/officecli-pipe-test.xlsx');
    const expectedDir = `/tmp/officecli-${typeof process.getuid === 'function' ? process.getuid() : 0}`;
    assert.strictEqual(path.dirname(main), expectedDir);
    assert.strictEqual(ping, `${main}-ping`);
  } finally {
    if (previous === undefined) delete process.env.TMPDIR;
    else process.env.TMPDIR = previous;
  }
}

if (process.platform === 'darwin') {
  const realDir = fs.mkdtempSync(path.join(os.tmpdir(), 'officecli-pipe-real-'));
  const linkDir = `${realDir}-link`;
  const realFile = path.join(realDir, 'book.xlsx');
  try {
    fs.writeFileSync(realFile, '');
    fs.symlinkSync(realDir, linkDir, 'dir');
    assert.deepStrictEqual(pipePaths(path.join(linkDir, 'book.xlsx')), pipePaths(realFile));
  } finally {
    fs.rmSync(linkDir, { force: true });
    fs.rmSync(realDir, { recursive: true, force: true });
  }
}

console.log('node pipe-path test PASS');
