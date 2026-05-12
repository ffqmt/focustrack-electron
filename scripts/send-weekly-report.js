#!/usr/bin/env node
'use strict';

require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });

const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');

// ── CLI arguments ─────────────────────────────────────────────────────────────

const args = process.argv.slice(2);
const isDryRun = args.includes('--dry-run');
const startArg = (args.find(a => a.startsWith('--start=')) || '').slice('--start='.length);
const toArg = (args.find(a => a.startsWith('--to=')) || '').slice('--to='.length);

// ── Timezone / date helpers ───────────────────────────────────────────────────

const TIMEZONE = process.env.WEEKLY_REPORT_TIMEZONE || 'America/Cuiaba';

function toLocalDate(date, tz) {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: tz,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(date);
}

function ptbrDate(iso) {
  const [y, m, d] = iso.split('-');
  return `${d}/${m}/${y}`;
}

function addDays(isoDate, n) {
  // Parse as UTC noon to avoid DST edge cases
  const d = new Date(`${isoDate}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

function calculateDates() {
  if (startArg) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(startArg)) {
      throw new Error(`--start deve ser YYYY-MM-DD. Recebido: "${startArg}"`);
    }
    return { startDate: startArg, endDate: addDays(startArg, 7) };
  }

  const now = new Date();
  const endDate = toLocalDate(now, TIMEZONE);
  const startDate = toLocalDate(new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000), TIMEZONE);
  return { startDate, endDate };
}

// ── Server management ─────────────────────────────────────────────────────────

const BASE_URL = (process.env.WEEKLY_REPORT_BASE_URL || 'http://localhost:3001').replace(/\/$/, '');
const SERVER_ROOT = path.join(__dirname, '..');

async function isServerReady() {
  try {
    const res = await fetch(`${BASE_URL}/health`, { signal: AbortSignal.timeout(2500) });
    return res.ok;
  } catch {
    return false;
  }
}

async function waitForServer(maxMs = 45000) {
  const deadline = Date.now() + maxMs;
  while (Date.now() < deadline) {
    if (await isServerReady()) return true;
    await new Promise(r => setTimeout(r, 1500));
  }
  return false;
}

function startServer() {
  const proc = spawn('node', ['server.js'], {
    cwd: SERVER_ROOT,
    stdio: ['ignore', 'pipe', 'pipe'],
    env: process.env,
  });

  proc.stdout.on('data', chunk =>
    process.stdout.write(`[server] ${chunk}`)
  );
  proc.stderr.on('data', chunk =>
    process.stderr.write(`[server:err] ${chunk}`)
  );

  proc.on('exit', (code, signal) => {
    if (code !== null && code !== 0) {
      console.error(`[server] Encerrou com código ${code}`);
    }
  });

  return proc;
}

// ── Fetch report HTML ─────────────────────────────────────────────────────────

async function fetchReportHtml(startDate) {
  const url = `${BASE_URL}/api/report/week/email-html?start=${startDate}`;
  console.log(`[report] GET ${url}`);

  const res = await fetch(url, { signal: AbortSignal.timeout(120_000) });
  const body = await res.text();

  if (!res.ok) {
    throw new Error(`Endpoint retornou HTTP ${res.status}: ${body.slice(0, 300)}`);
  }

  let data;
  try {
    data = JSON.parse(body);
  } catch {
    throw new Error(`Resposta não é JSON válido: ${body.slice(0, 200)}`);
  }

  if (!data.ok) {
    throw new Error(`ok=false na resposta: ${JSON.stringify(data).slice(0, 300)}`);
  }

  return data;
}

// ── Validate HTML ─────────────────────────────────────────────────────────────

function validateHtml(html) {
  if (typeof html !== 'string') {
    throw new Error(`"html" não é string (tipo: ${typeof html})`);
  }
  if (html.length < 500) {
    throw new Error(`HTML muito curto (${html.length} chars). Verifique o endpoint.`);
  }
  if (!/<html|<table|<div/i.test(html)) {
    throw new Error('HTML não contém estrutura mínima esperada (<html>, <table> ou <div>).');
  }
}

// ── Save snapshot ─────────────────────────────────────────────────────────────

function saveSnapshot(startDate, html) {
  const dir = path.join(SERVER_ROOT, 'reports', 'email-html');
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, `weekly-report-${startDate}.html`);
  fs.writeFileSync(file, html, 'utf8');
  console.log(`[report] Snapshot salvo: ${file}`);
  return file;
}

// ── Send email ────────────────────────────────────────────────────────────────

async function sendEmail({ html, startDate, endDate, to }) {
  const from = process.env.WEEKLY_REPORT_FROM;
  const toAddr = to || process.env.WEEKLY_REPORT_TO;
  const cc = process.env.WEEKLY_REPORT_CC || '';
  const bcc = process.env.WEEKLY_REPORT_BCC || '';

  const requiredVars = {
    SMTP_HOST: process.env.SMTP_HOST,
    SMTP_USER: process.env.SMTP_USER,
    SMTP_PASS: process.env.SMTP_PASS,
    WEEKLY_REPORT_FROM: from,
    WEEKLY_REPORT_TO: toAddr,
  };
  const missing = Object.entries(requiredVars)
    .filter(([, v]) => !v)
    .map(([k]) => k);

  if (missing.length) {
    throw new Error(`Variáveis de envio ausentes: ${missing.join(', ')}`);
  }

  const subject = `Relatório semanal FocusTrack - ${ptbrDate(startDate)} a ${ptbrDate(endDate)}`;
  const text =
    'Relatório semanal FocusTrack.\n' +
    'Caso não consiga visualizar o HTML, acesse o sistema FocusTrack.';

  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: Number(process.env.SMTP_PORT || 587),
    secure: process.env.SMTP_SECURE === 'true',
    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS,
    },
  });

  const mailOptions = { from, to: toAddr, subject, text, html };
  if (cc) mailOptions.cc = cc;
  if (bcc) mailOptions.bcc = bcc;

  console.log(`[email] Para: ${toAddr}`);
  console.log(`[email] Assunto: ${subject}`);

  const info = await transporter.sendMail(mailOptions);
  console.log(`[email] Enviado. MessageId: ${info.messageId}`);
  return info;
}

// ── Main ──────────────────────────────────────────────────────────────────────

async function main() {
  const line = '='.repeat(60);
  console.log(line);
  console.log('[focustrack] Relatório semanal FocusTrack');
  console.log(`[focustrack] Modo: ${isDryRun ? 'DRY-RUN (sem envio de e-mail)' : 'REAL'}`);
  console.log(`[focustrack] Timezone: ${TIMEZONE}`);

  const { startDate, endDate } = calculateDates();
  console.log(`[focustrack] Período: ${startDate} → ${endDate}`);
  console.log(line);

  let serverProc = null;

  try {
    const alreadyRunning = await isServerReady();

    if (!alreadyRunning) {
      console.log('[server] Backend não detectado. Iniciando...');
      serverProc = startServer();
      const ready = await waitForServer(45_000);
      if (!ready) {
        throw new Error('Backend não respondeu em 45 segundos. Abortando.');
      }
      console.log('[server] Backend pronto.');
    } else {
      console.log('[server] Backend já disponível em uso externo.');
    }

    // Fetch
    const reportData = await fetchReportHtml(startDate);
    const { html } = reportData;

    // Validate
    validateHtml(html);
    console.log(`[report] HTML válido (${html.length} chars).`);

    // Snapshot
    saveSnapshot(startDate, html);

    // Send or dry-run
    if (isDryRun) {
      console.log('[email] DRY-RUN ativo: e-mail não será enviado.');
    } else {
      await sendEmail({ html, startDate, endDate, to: toArg || undefined });
    }

    console.log(line);
    console.log('[focustrack] Concluído com sucesso.');
    process.exitCode = 0;
  } catch (err) {
    console.error('[ERRO]', err.message);
    if (process.env.NODE_ENV !== 'production') {
      console.error(err.stack);
    }
    process.exitCode = 1;
  } finally {
    if (serverProc) {
      console.log('[server] Encerrando backend...');
      serverProc.kill('SIGTERM');
      // Give it a moment to flush
      await new Promise(r => setTimeout(r, 500));
    }
  }
}

if (require.main === module) {
  main().catch((error) => {
    console.error('[focustrack] Erro fatal:', error);
    process.exit(1);
  });
}
