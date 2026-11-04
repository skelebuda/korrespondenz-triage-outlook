/** Dev helper: `node scripts/lunnoa.mjs <command>` against the deployment in .env. */
import { readFileSync } from 'node:fs';
import { createServerClient } from '@lunnoa/client';
import { ensureAutomation } from '../triage.mjs';

for (const line of readFileSync(new URL('../.env', import.meta.url), 'utf8').split(/\r?\n/)) {
  const eq = line.indexOf('=');
  if (eq > 0 && !line.trim().startsWith('#')) process.env[line.slice(0, eq).trim()] = line.slice(eq + 1).trim();
}
const lunnoa = createServerClient({ baseUrl: process.env.LUNNOA_URL, apiKey: process.env.LUNNOA_API_KEY });
const [command, ...args] = process.argv.slice(2);

const commands = {
  async cases() {
    const page = await lunnoa.entities.list({ objectTypeSlug: 'korrespondenz', pageSize: 100, expansion: ['attributes', 'createdAt'] });
    for (const entity of page.data || []) {
      const a = entity.attributes || {};
      console.log(entity.id, entity.state.padEnd(12), String(a.category).padEnd(19), a.external_id ? 'outlook' : '-      ', entity.name);
    }
  },
  /** Creates a case without Outlook: create <category> <subject> <body>. */
  async create() {
    const [category, subject, body] = args;
    const { objectTypeId } = await ensureAutomation(lunnoa);
    const entity = await lunnoa.entities.create({
      name: subject,
      objectTypeId,
      state: 'eingegangen',
      attributes: { category, subject, body, preview: body.slice(0, 250), from_name: 'Testbank', from_email: 'test@example.com' },
    });
    console.log(entity.id);
  },
  async delete() {
    for (const id of args) await lunnoa.entities.delete(id);
  },
  /** Runs a case workflow on one object and prints every step. */
  async run() {
    const [key, objectId] = args;
    const { workflowIds } = await ensureAutomation(lunnoa);
    const started = await lunnoa.executions.execute(workflowIds[key], { objectId });
    const finished = await lunnoa.executions.waitUntilFinished(started.id ?? started.executionId);
    const full = await lunnoa.executions.get(finished.id, { expansion: ['nodes', 'status', 'statusMessage'] });
    console.log(full.status, full.statusMessage);
    for (const node of full.nodes || []) {
      if (node.output == null) continue;
      const output = node.output.response ?? node.output;
      console.log(`- ${node.name}: ${JSON.stringify(output).slice(0, 600)}`);
    }
  },
};

if (!commands[command]) throw new Error(`Commands: ${Object.keys(commands).join(', ')}`);
await commands[command]();
