#!/usr/bin/env node
// Usage: node cli.mjs <breakdown.xlsx> <schedule.pdf> [--column "Sc No"] [--out sorted.xlsx]
import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import './core.js';

const args = process.argv.slice(2);
const opt = (name) => {
  const i = args.indexOf(name);
  if (i === -1) return null;
  const v = args[i + 1];
  args.splice(i, 2);
  return v;
};
const columnName = opt('--column');
let out = opt('--out');
const xlsxPath = args.find((a) => /\.xlsx$/i.test(a));
const pdfPath = args.find((a) => /\.pdf$/i.test(a));

if (!xlsxPath || !pdfPath) {
  console.error('Usage: node cli.mjs <breakdown.xlsx> <schedule.pdf> [--column "Sc No"] [--out sorted.xlsx]');
  process.exit(1);
}
out ??= path.join(path.dirname(xlsxPath), path.basename(xlsxPath, path.extname(xlsxPath)) + ' - shoot order.xlsx');

// pdf.js warns about missing canvas support at import; it only needs text extraction here
const { log, warn } = console;
console.log = console.warn = () => {};
const { default: pdfjs } = await import('pdfjs-dist/legacy/build/pdf.js');
Object.assign(console, { log, warn });

const { SchedCore } = globalThis;
// isEvalSupported: false closes the pdf.js 3.x malicious-PDF hole (GHSA-wgrm-67xf-hhpq)
const pdf = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(pdfPath)), verbosity: 0, isEvalSupported: false }).promise;
const sched = await SchedCore.parseSchedule(pdf);
if (!sched.sceneCount) {
  console.error(`No scenes found in ${pdfPath}. It needs to be a text (not scanned) one-line schedule with INT/EXT on each scene line.`);
  process.exit(1);
}

const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile(xlsxPath);
const report = SchedCore.processWorkbook(wb, sched, { columnName });
if (report.sheets.every((s) => s.skipped)) {
  console.error(columnName
    ? `No column titled "${columnName}" found in ${xlsxPath}.`
    : `No scene column found in ${xlsxPath}. Pass its title with --column "<title>".`);
  process.exit(1);
}
await wb.xlsx.writeFile(out);

console.log(`Schedule: ${sched.days.filter((d) => !d.unscheduled).length} shoot days, ${sched.sceneCount} scene entries`);
for (const s of report.sheets) {
  if (s.skipped) { console.log(`  ${s.name}: left as is (no scene column)`); continue; }
  console.log(`  ${s.name}: ${s.matched}/${s.scenes} matched${s.unmatched.length ? ` — not in schedule: ${s.unmatched.join(', ')}` : ''}`);
}
if (report.letterFallbacks.length) {
  console.log(`  Lettered scenes scheduled against the plain scene: ${report.letterFallbacks.map((f) => `${f.from} -> ${f.to}`).join(', ')}`);
}
if (report.unmatchedSchedule.length) {
  console.log(`  In schedule but not in breakdown: ${report.unmatchedSchedule.map((u) => u.field).join(', ')}`);
}
console.log(`Wrote ${out}`);
