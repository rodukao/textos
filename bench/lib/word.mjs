// Executa tarefas no Microsoft Word (via bench/word.ps1) e devolve os resultados.
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const script = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'word.ps1');

/** Encerra instâncias do Word abertas por automação (nunca as do usuário). */
function killAutomationWord() {
  return new Promise((resolve) => {
    execFile(
      'powershell',
      [
        '-NoProfile',
        '-Command',
        "Get-CimInstance Win32_Process -Filter \"Name='WINWORD.EXE'\" | Where-Object { $_.CommandLine -like '*/Automation*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }",
      ],
      () => resolve(),
    );
  });
}

/**
 * @param {{op: 'pdf'|'check', in: string, out?: string}[]} jobs
 * @param {number} timeoutMs tempo máximo para o lote inteiro
 */
export async function runWord(jobs, timeoutMs = 60000 + jobs.length * 30000) {
  if (jobs.length === 0) return [];
  const dir = await mkdtemp(path.join(tmpdir(), 'textos-word-'));
  const jobsFile = path.join(dir, 'jobs.json');
  const resultsFile = path.join(dir, 'results.json');
  await writeFile(jobsFile, JSON.stringify(jobs), 'utf8');
  try {
    await new Promise((resolve, reject) => {
      const child = execFile(
        'powershell',
        ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script, '-Jobs', jobsFile, '-Results', resultsFile],
        { timeout: timeoutMs, windowsHide: true },
        (err, _stdout, stderr) => {
          if (err) reject(new Error(`Word falhou: ${err.killed ? 'tempo esgotado' : err.message}\n${stderr}`));
          else resolve();
        },
      );
      child.stdin?.end();
    });
    const text = (await readFile(resultsFile, 'utf8')).replace(/^﻿/, '');
    return JSON.parse(text);
  } catch (e) {
    await killAutomationWord();
    throw e;
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}
