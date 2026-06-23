// Guard against NUL bytes in tracked text-source files. A stray NUL makes git
// treat the file as binary and can silently corrupt content while still passing
// typecheck/lint/format/tests (see the echo-key NUL incident). Part of `verify`.
import { readFileSync } from 'node:fs';
import { execSync } from 'node:child_process';

const exts = [
  '.ts',
  '.tsx',
  '.js',
  '.mjs',
  '.cjs',
  '.json',
  '.md',
  '.yml',
  '.yaml',
  '.html',
  '.css',
];

const files = execSync('git ls-files -z', { encoding: 'utf8' })
  .split('\0')
  .filter((f) => f.length > 0 && exts.some((e) => f.endsWith(e)));

const offenders = files.filter((f) => readFileSync(f).includes(0));

if (offenders.length > 0) {
  console.error('check:text — NUL byte found in text source file(s):');
  for (const f of offenders) console.error(`  ${f}`);
  process.exit(1);
}

console.log(`check:text — ${String(files.length)} text files clean (no NUL bytes).`);
