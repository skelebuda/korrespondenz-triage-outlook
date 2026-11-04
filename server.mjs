import { readFileSync, existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import express from 'express';
import { createServerClient } from '@lunnoa/client';
import { ensureAutomation } from './triage.mjs';
import { canRestart, caseFromEntity, nextState, processOf } from './cases.mjs';

/** Which case workflow handles which category. */
const CASE_WORKFLOW = {
  transaction_advice: 'transaction',
  corporate_action: 'corporateAction',
  custodian_query: 'custodianQuery',
};

const root = path.dirname(fileURLToPath(import.meta.url));
const distDir = path.join(root, 'dist');

loadEnv(path.join(root, '.env'));

const lunnoa = createServerClient({
  baseUrl: process.env.LUNNOA_URL,
  apiKey: process.env.LUNNOA_API_KEY,
});

let automationPromise;

function automation() {
  if (!automationPromise) {
    automationPromise = ensureAutomation(lunnoa).catch((error) => {
      automationPromise = null;
      throw error;
    });
  }
  return automationPromise;
}

/** True when a real lna_ key is present. Placeholder values stay unconfigured. */
export function lunnoaConfigured(env = process.env) {
  const key = env.LUNNOA_API_KEY || '';
  const url = env.LUNNOA_URL || '';
  return Boolean(url && key.startsWith('lna_') && !key.includes('your_key'));
}

function workflowLink(workflowId) {
  return `${process.env.LUNNOA_URL}/projects/${process.env.LUNNOA_PROJECT_ID}/workflows/${workflowId}`;
}

/** Prefer the Outlook mailbox owner, then any human workspace user. */
export function userFromWorkspace(people, mailbox) {
  const rows = (Array.isArray(people) ? people : []).map((row) => row?.user).filter(Boolean);
  const user =
    rows.find((item) => mailbox && item.email === mailbox) ||
    rows.find((item) => !/service-accounts\.internal|service account/i.test(`${item.email || ''} ${item.name || ''}`));
  if (!user) return { name: mailbox || 'Angemeldet', email: mailbox || '', role: 'Korrespondenz' };
  return { name: user.name || user.email, email: user.email || mailbox || '', role: 'Korrespondenz' };
}

let sessionPromise;

function sessionUser() {
  if (!sessionPromise) {
    sessionPromise = (async () => {
      const mailbox = process.env.OUTLOOK_CONNECTION_ID
        ? (await lunnoa.connections.get(process.env.OUTLOOK_CONNECTION_ID).catch(() => null))?.name
        : '';
      const people = await lunnoa.http.get('/api/workspace-users', { query: { expansion: 'user' } }).catch(() => []);
      return userFromWorkspace(people, mailbox);
    })().catch(() => {
      sessionPromise = null;
      return { name: 'Angemeldet', email: '', role: 'Korrespondenz' };
    });
  }
  return sessionPromise;
}

async function loadBoard() {
  const { workflowIds, hasReplyDraft } = await automation();
  const page = await lunnoa.entities.list({
    objectTypeSlug: 'korrespondenz',
    pageSize: 100,
    expansion: ['attributes', 'createdAt'],
  });
  const cases = (page.data || []).map((entity) => {
    const item = caseFromEntity(entity);
    return { ...item, view: processOf(item) };
  });
  cases.sort((a, b) => String(b.createdAt || '').localeCompare(String(a.createdAt || '')));
  const links = Object.fromEntries(
    Object.entries(CASE_WORKFLOW).map(([category, key]) => [category, workflowLink(workflowIds[key])]),
  );
  return {
    cases,
    links: {
      ...links,
      inbox: workflowLink(workflowIds.inbox),
      askCustodian: workflowLink(workflowIds.askCustodian),
      notifyPm: workflowLink(workflowIds.notifyPm),
    },
    replyInThread: hasReplyDraft,
    user: await sessionUser(),
  };
}

async function runCaseAction(entity, action, pmNote) {
  const { workflowIds } = await automation();
  const category = entity.attributes?.category;
  if (action === 'restart') {
    const key = CASE_WORKFLOW[category];
    if (!key) throw new Error(`Keine Automatik für Kategorie ${category}`);
    if (entity.state !== 'eingegangen') {
      if (!canRestart(entity.state)) throw new Error('Ein erledigter Fall wird nicht neu gestartet');
      await lunnoa.entities.changeState(entity.id, { newState: 'eingegangen', reason: 'Automatik neu gestartet' });
    }
    const started = await lunnoa.executions.execute(workflowIds[key], { objectId: entity.id });
    return started.id ?? started.executionId;
  }
  if (action === 'ask-custodian') {
    if (entity.state !== 'abweichung') throw new Error('Eine Rückfrage gibt es nur bei einer Abweichung');
    const started = await lunnoa.executions.execute(workflowIds.askCustodian, { objectId: entity.id });
    return started.id ?? started.executionId;
  }
  if (action === 'match-client') {
    if (!entity.attributes?.from_email) throw new Error('Der Fall hat keine Absenderadresse');
    const started = await lunnoa.executions.execute(workflowIds.clientMatch, { objectId: entity.id });
    return started.id ?? started.executionId;
  }
  if (action === 'approve-note') {
    if (entity.state !== 'pm_freigabe') throw new Error('Freigabe gibt es nur nach dem PM-Hinweis');
    const started = await lunnoa.executions.execute(workflowIds.notifyPm, {
      objectId: entity.id,
      pmNote: typeof pmNote === 'string' ? pmNote : entity.attributes?.pm_note || '',
    });
    return started.id ?? started.executionId;
  }
  const state = nextState(entity.state, action);
  if (typeof pmNote === 'string') await lunnoa.entities.update(entity.id, { attributes: { pm_note: pmNote } });
  await lunnoa.entities.changeState(entity.id, { newState: state, reason: action });
  return null;
}

const app = express();
app.use(express.json({ limit: '1mb' }));

app.get('/api/health', (_req, res) => {
  res.json({ ok: true, lunnoa: lunnoaConfigured() });
});

app.get('/api/board', async (_req, res) => {
  res.json(await loadBoard());
});

app.get('/api/executions/:id', async (req, res) => {
  const progress = await lunnoa.executions.getProgress(req.params.id);
  res.json({
    id: progress.executionId,
    status: progress.status,
    statusMessage: progress.statusMessage,
    activeStepIndex: progress.activeStepIndex,
    steps: (progress.steps || []).map((step) => ({
      label: step.label,
      status: step.status,
      durationMs: step.durationMs,
    })),
  });
});

app.post('/api/cases/:id', async (req, res) => {
  const entity = await lunnoa.entities.get(req.params.id, { expansion: ['attributes'] });
  const executionId = await runCaseAction(entity, String(req.body?.action || ''), req.body?.pmNote);
  res.json({ ...(await loadBoard()), executionId });
});

app.use('/api', (_req, res) => {
  res.status(404).json({ error: 'Not found' });
});

app.use(express.static(distDir));
app.use((req, res) => {
  if (req.method !== 'GET' && req.method !== 'HEAD') {
    res.status(404).type('txt').send('Not found');
    return;
  }
  if (path.extname(req.path)) {
    res.status(404).type('txt').send('Not found');
    return;
  }
  res.sendFile(path.join(distDir, 'index.html'), (error) => {
    if (error) res.status(404).type('txt').send('Build the UI with npm run build');
  });
});

app.use((error, _req, res, next) => {
  if (res.headersSent) {
    next(error);
    return;
  }
  res.status(502).json({ error: error.message || 'Request failed' });
});

function loadEnv(file) {
  if (!existsSync(file)) return;
  for (const line of readFileSync(file, 'utf8').split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const eq = trimmed.indexOf('=');
    if (eq === -1) continue;
    const key = trimmed.slice(0, eq).trim();
    const value = trimmed.slice(eq + 1).trim();
    process.env[key] = value;
  }
}

if (process.argv.includes('--check')) {
  if (lunnoaConfigured({ LUNNOA_URL: 'https://automate.lunnoa.com', LUNNOA_API_KEY: 'lna_test' }) !== true) {
    throw new Error('configured key was rejected');
  }
  if (lunnoaConfigured({ LUNNOA_URL: 'https://automate.lunnoa.com', LUNNOA_API_KEY: '' }) !== false) {
    throw new Error('empty key was accepted');
  }
  if (lunnoaConfigured({ LUNNOA_URL: 'https://automate.lunnoa.com', LUNNOA_API_KEY: 'lna_your_key' }) !== false) {
    throw new Error('placeholder key was accepted');
  }
  const shown = userFromWorkspace(
    [{ user: { name: 'Sasa', email: 'sasakelebuda@lunnoalabs.ch' } }, { user: { name: 'Service account: Triage', email: 'a@service-accounts.internal' } }],
    'sasakelebuda@lunnoalabs.ch',
  );
  if (shown.name !== 'Sasa' || shown.email !== 'sasakelebuda@lunnoalabs.ch') throw new Error('session user should be the mailbox owner');
  console.log('ok');
} else if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const port = Number(process.env.PORT || 4318);
  const server = app.listen(port, () => {
    console.log(`Korrespondenz-Triage listening on http://localhost:${port}`);
  });
  server.requestTimeout = 0;
}
