'use strict';

const assert = require('assert');
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

console.log('node pipe-path test PASS');
