function textField(id, label, order) {
  return {
    id,
    name: id,
    label,
    type: 'text',
    order,
    required: false,
    visibility: 'always',
    config: {},
  };
}

function section(id, name, order, labels) {
  const fields = Object.entries(labels).map(([fieldId, label], index) => textField(fieldId, label, index));
  return { id, name, order, collapsible: false, defaultExpanded: true, fields };
}

const transitions = {
  eingegangen: ['abgeglichen', 'abweichung', 'pm_freigabe', 'entwurf', 'freigegeben'],
  abgeglichen: ['eingegangen', 'freigegeben'],
  abweichung: ['eingegangen', 'entwurf', 'freigegeben'],
  pm_freigabe: ['eingegangen', 'freigegeben'],
  entwurf: ['eingegangen', 'freigegeben'],
};

export const correspondenceType = {
  slug: 'korrespondenz',
  name: 'Korrespondenz',
  namePlural: 'Korrespondenzen',
  description: 'Eine Depotbank- oder Kundenmail, die als Fall bearbeitet wird.',
  icon: 'mail',
  color: '#06355a',
  allowUnknownAttributes: false,
  attributeSchema: {
    version: '1',
    sections: [
      section('mail', 'Mail', 0, {
        category: 'Kategorie',
        from_name: 'Absender',
        from_email: 'Adresse',
        subject: 'Betreff',
        preview: 'Vorschau',
        body: 'Mailtext',
        external_id: 'Outlook-Id',
        mail_link: 'Mail in Outlook',
      }),
      section('work', 'Fall', 1, {
        customer: 'Kunde',
        company: 'Gesellschaft',
        portfolio: 'Portfolio',
        instrument: 'Instrument',
        isin: 'ISIN',
        mail_quantity: 'Stück Mail',
        book_quantity: 'Stück Bestand',
        amount: 'Betrag',
        currency: 'Währung',
        value_date: 'Valuta',
        match_status: 'Abgleich',
        reasoning: 'Begründung',
        event_type: 'Ereignis',
        ex_date: 'Ex-Datum',
        affected: 'Betroffene Portfolios',
        pm_note: 'Hinweis Portfolio Management',
        requested: 'Angefragt',
        documents: 'Unterlagen',
        draft: 'Antwortentwurf',
        outlook_link: 'Outlook-Entwurf',
      }),
    ],
  },
  stateSchema: {
    version: '1',
    states: [
      { id: 'eingegangen', name: 'Eingegangen', color: '#06355a', order: 0, isInitial: true },
      { id: 'abgeglichen', name: 'Abgeglichen', color: '#2e557e', order: 1 },
      { id: 'abweichung', name: 'Abweichung', color: '#e64335', order: 2 },
      { id: 'pm_freigabe', name: 'PM-Freigabe', color: '#5a7492', order: 3 },
      { id: 'entwurf', name: 'Entwurf in Outlook', color: '#5a7492', order: 4 },
      { id: 'freigegeben', name: 'Erledigt', color: '#244d6e', order: 5, isFinal: true },
    ],
    transitions: Object.entries(transitions)
      .flatMap(([fromStateId, targets]) => targets.map((toStateId) => [fromStateId, toStateId]))
      .map(([fromStateId, toStateId], index) => ({ id: `t-${index}`, fromStateId, toStateId })),
  },
};

/** Workflow-written arrays arrive as arrays; text written by hand stays text. */
function list(value) {
  if (Array.isArray(value)) return value;
  if (typeof value !== 'string' || !value.trim().startsWith('[')) return [];
  try {
    return JSON.parse(value);
  } catch {
    return [];
  }
}

export function caseFromEntity(entity) {
  const attributes = entity.attributes || {};
  return {
    ...attributes,
    id: entity.id,
    state: entity.state || 'eingegangen',
    createdAt: entity.createdAt,
    subject: attributes.subject || entity.name || '',
    affected: list(attributes.affected),
    documents: list(attributes.documents),
  };
}

export async function ensureCorrespondenceType(lunnoa) {
  const type = await lunnoa.entityTypes.upsertBySlug(correspondenceType);
  if (!type?.id) throw new Error('Korrespondenz object type was not created');
  return type;
}

/** Human steps the app may take, keyed by the state the case must be in. */
const humanSteps = {
  confirm: ['abgeglichen', 'freigegeben'],
  reviewed: ['abweichung', 'freigegeben'],
  'approve-note': ['pm_freigabe', 'freigegeben'],
  done: ['entwurf', 'freigegeben'],
};

export function nextState(current, action) {
  const step = humanSteps[action];
  if (!step) throw new Error('Unknown action');
  if (step[0] !== current) throw new Error(`Der Fall ist nicht im Status ${step[0]}`);
  return step[1];
}

export function canRestart(state) {
  return Boolean(transitions[state]?.includes('eingegangen'));
}

function filled(value) {
  if (Array.isArray(value)) return value.length > 0;
  return Boolean(value && String(value).trim());
}

/** First unfinished step becomes current; skipped stays skipped. */
function finish(steps) {
  let current = false;
  return steps.map((step) => {
    if (step.status === 'done' || step.status === 'skipped') return step;
    if (current) return { ...step, status: 'todo' };
    current = true;
    return { ...step, status: 'current' };
  });
}

const PROCESS = {
  transaction_advice: {
    title: 'Transaktionsavis abgleichen',
    summary: 'Der Leser holt die Felder aus dem Avis, eine Regel prüft den Bestand, bei Abweichung schreibt der Korrespondent die Rückfrage nach Outlook.',
  },
  corporate_action: {
    title: 'Corporate Action auswerten',
    summary: 'Der Leser erkennt das Ereignis, der Bestand liefert die betroffenen Portfolios, der Analyst schreibt den Hinweis für das Portfolio Management.',
  },
  custodian_query: {
    title: 'Depotbank-Rückfrage beantworten',
    summary: 'Der Leser versteht die Frage, das Unterlagenregister wird geprüft, der Korrespondent legt die Antwort als Entwurf in Outlook ab.',
  },
};

function transactionSteps(item) {
  const read = filled(item.instrument) || filled(item.isin) || filled(item.mail_quantity);
  const checked = filled(item.match_status);
  const matched = item.match_status === 'matched' || item.state === 'abgeglichen';
  const mismatch = item.match_status === 'mismatch' || item.match_status === 'unmatched' || item.state === 'abweichung';
  const decided = matched || mismatch || item.state === 'entwurf' || item.state === 'freigegeben';
  const replied = filled(item.outlook_link) || item.state === 'entwurf' || (item.state === 'freigegeben' && !matched);
  return finish([
    { id: 'in', name: 'Eingegangen', status: 'done' },
    { id: 'read', name: 'Avis gelesen', status: read ? 'done' : 'todo' },
    { id: 'book', name: 'Bestand geprüft', status: checked ? 'done' : 'todo' },
    { id: 'decide', name: mismatch ? 'Abweichung' : matched ? 'Abgeglichen' : 'Abgleich', status: decided ? 'done' : 'todo' },
    { id: 'reply', name: 'Rückfrage in Outlook', status: matched && !mismatch ? 'skipped' : replied ? 'done' : 'todo' },
    { id: 'done', name: 'Erledigt', status: item.state === 'freigegeben' ? 'done' : 'todo' },
  ]);
}

function corporateSteps(item) {
  const read = filled(item.event_type) || filled(item.isin) || filled(item.instrument);
  const holders = filled(item.affected) || filled(item.pm_note);
  const noted = filled(item.pm_note);
  return finish([
    { id: 'in', name: 'Eingegangen', status: 'done' },
    { id: 'read', name: 'Ereignis gelesen', status: read ? 'done' : 'todo' },
    { id: 'holders', name: 'Portfolios gefunden', status: holders ? 'done' : 'todo' },
    { id: 'note', name: 'PM-Hinweis', status: noted ? 'done' : 'todo' },
    { id: 'review', name: 'PM-Freigabe', status: item.state === 'freigegeben' ? 'done' : 'todo' },
    { id: 'done', name: 'Erledigt', status: item.state === 'freigegeben' ? 'done' : 'todo' },
  ]);
}

function querySteps(item) {
  const read = filled(item.requested) || filled(item.customer);
  const docs = filled(item.documents) || filled(item.draft);
  const written = filled(item.draft);
  const outlook = filled(item.outlook_link) || item.state === 'entwurf';
  return finish([
    { id: 'in', name: 'Eingegangen', status: 'done' },
    { id: 'read', name: 'Rückfrage verstanden', status: read ? 'done' : 'todo' },
    { id: 'docs', name: 'Unterlagen geprüft', status: docs ? 'done' : 'todo' },
    { id: 'draft', name: 'Antwort geschrieben', status: written ? 'done' : 'todo' },
    { id: 'outlook', name: 'Entwurf in Outlook', status: outlook ? 'done' : 'todo' },
    { id: 'done', name: 'Erledigt', status: item.state === 'freigegeben' ? 'done' : 'todo' },
  ]);
}

const STEPS = {
  transaction_advice: transactionSteps,
  corporate_action: corporateSteps,
  custodian_query: querySteps,
};

/** Status and steps the case view renders. Derived from state plus filled fields. */
export function processOf(item) {
  const copy = PROCESS[item.category] || { title: 'Fall', summary: '' };
  const build = STEPS[item.category];
  const steps = build
    ? build(item)
    : finish([
        { id: 'in', name: 'Eingegangen', status: 'done' },
        { id: 'done', name: 'Erledigt', status: item.state === 'freigegeben' ? 'done' : 'todo' },
      ]);
  return { ...copy, steps };
}

if (process.argv.includes('--check')) {
  const assert = (ok, message) => {
    if (!ok) throw new Error(message);
  };
  const tx = processOf({ category: 'transaction_advice', state: 'eingegangen' });
  assert(tx.steps[0].status === 'done' && tx.steps[1].status === 'current', 'new avis starts at read');
  const matched = processOf({
    category: 'transaction_advice',
    state: 'abgeglichen',
    instrument: 'Nestle',
    mail_quantity: '80',
    match_status: 'matched',
  });
  assert(matched.steps.find((step) => step.id === 'reply')?.status === 'skipped', 'matched avis skips outlook');
  assert(matched.steps.find((step) => step.id === 'done')?.status === 'current', 'matched avis waits for confirm');
  const query = processOf({
    category: 'custodian_query',
    state: 'entwurf',
    requested: 'W-8BEN',
    documents: [{ doc_type: 'W-8BEN' }],
    draft: 'Guten Tag',
    outlook_link: 'https://outlook.office.com/mail',
  });
  assert(query.steps.at(-1).status === 'current', 'query draft waits for done');
  assert(query.steps.find((step) => step.id === 'outlook')?.status === 'done', 'query outlook step');
  console.log('ok');
}
