import * as fs from 'fs';
import * as path from 'path';
import { spawnSync } from 'child_process';
import { generateAllSuites } from './generate-suites';

const __dirname = path.dirname(new URL(import.meta.url).pathname).replace(/^\//, '').replace(/^([a-zA-Z]:)/, '$1');
const OUTPUT_DIR = path.join(__dirname, 'test-output');
const LOG_FILE = path.join(__dirname, 'validation.log');

interface Result {
  file: string;
  status: 'PASS' | 'FAIL' | 'SKIP';
  details: string;
  timeMs: number;
}

function cleanOutput() {
  if (fs.existsSync(OUTPUT_DIR)) {
    fs.rmSync(OUTPUT_DIR, { recursive: true, force: true });
  }
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  if (fs.existsSync(LOG_FILE)) fs.unlinkSync(LOG_FILE);
}

function validateWithExcel(): { results: Result[]; elapsed: string; precheck: string } {
  console.log('\nValidating with MS Excel...');
  const ps = spawnSync('powershell', [
    '-ExecutionPolicy', 'Bypass',
    '-NoProfile',
    '-File', path.join(__dirname, 'validate-with-excel.ps1'),
    '-TestDir', OUTPUT_DIR,
    '-LogFile', LOG_FILE
  ], { encoding: 'utf-8', timeout: 300000 });

  const lines = (ps.stdout || '').trim().split('\n').filter(l => l.length > 0);
  const results: Result[] = [];
  let elapsed = '?';
  let precheck = '';

  for (const line of lines) {
    const parts = line.split('|');
    if (parts[0] === 'PRECHECK') {
      precheck = `  Precheck: ${parts[1]} — ${parts.slice(2).join('|')}`;
    } else if (parts[0] === 'PASS' || parts[0] === 'FAIL') {
      results.push({
        file: parts[1],
        status: parts[0] as 'PASS' | 'FAIL',
        details: parts[2],
        timeMs: parseInt(parts[3]) || 0,
      });
    } else if (parts[0] === 'SKIP') {
      return { results: [{ file: 'ALL', status: 'SKIP', details: 'Excel not installed', timeMs: 0 }], elapsed: '', precheck: '' };
    } else if (parts[0] === 'ELAPSED') {
      elapsed = parts[4] || '?';
    }
  }

  return { results, elapsed, precheck };
}

function printReport(results: Result[], elapsed: string, precheck: string) {
  if (results.length === 1 && results[0].status === 'SKIP') {
    console.log('\n' + '='.repeat(70));
    console.log('Excel Compatibility Tests — SKIPPED');
    console.log('='.repeat(70));
    console.log('Microsoft Excel is not installed or COM automation is unavailable.');
    console.log('These tests require a local Excel installation (Windows only).');
    console.log('='.repeat(70));
    return;
  }

  const passed = results.filter(r => r.status === 'PASS');
  const failed = results.filter(r => r.status === 'FAIL');

  console.log('\n' + '='.repeat(70));
  console.log('Excel Compatibility Test Report');
  console.log('='.repeat(70));

  if (precheck) console.log(precheck);

  // Show timing summary
  const totalMs = results.reduce((sum, r) => sum + r.timeMs, 0);
  const avgMs = results.length > 0 ? Math.round(totalMs / results.length) : 0;
  console.log(`\n  Files: ${results.length}  |  Total time: ${elapsed}  |  Avg per file: ${avgMs}ms\n`);

  console.log(`Suite                                                Format  Status  Time    Details`);
  console.log('-'.repeat(70));

  for (const r of results) {
    const name = path.basename(r.file);
    const suite = name.replace(/\.(xlsx|xlsb)$/, '');
    const fmt = name.endsWith('.xlsx') ? 'xlsx' : 'xlsb';
    const statusStr = r.status === 'PASS' ? 'PASS' : 'FAIL';
    const timeStr = `${r.timeMs}ms`.padStart(7);
    const line = `${suite.padEnd(52)} ${fmt.padEnd(6)} ${statusStr.padEnd(6)} ${timeStr} ${r.details}`;
    console.log(line);
  }

  console.log('-'.repeat(70));
  console.log(`Total: ${results.length} files | Passed: ${passed.length} | Failed: ${failed.length} | Time: ${elapsed}`);
  console.log('='.repeat(70));

  if (failed.length > 0) {
    console.log('\nFAILED FILES:');
    for (const f of failed) {
      console.log(`  FAIL  ${f.file}: ${f.details}`);
    }
  }
}

async function main() {
  console.log('='.repeat(70));
  console.log('Excel Compatibility Test Suite');
  console.log('='.repeat(70));

  // 1. Clean
  console.log('\n[1/3] Cleaning test output...');
  cleanOutput();

  // 2. Generate
  console.log('[2/3] Generating test files...');
  await generateAllSuites(OUTPUT_DIR);

  // 3. Validate with Excel
  console.log('[3/3] Validating with MS Excel...');
  const { results, elapsed, precheck } = validateWithExcel();

  // 4. Report
  printReport(results, elapsed, precheck);

  // 5. Exit code
  const failedCount = results.filter(r => r.status === 'FAIL').length;
  process.exit(failedCount);
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
