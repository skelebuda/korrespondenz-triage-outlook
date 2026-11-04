import { ensureCorrespondenceType } from './cases.mjs';

/** Only these categories become cases. Every other mail stops at the scope decision. */
export const KEPT_CATEGORIES = ['transaction_advice', 'corporate_action', 'custodian_query'];

/** Stable trigger id so a redeploy does not reset the Outlook poll cursor. */
export const OUTLOOK_TRIGGER_ID = 'f1111111-1111-4111-8111-111111111111';

const AGENT_MODEL = 'deepseek-ai/DeepSeek-V4-Pro';
const REPLY_DRAFT_ACTION = 'microsoft-outlook_action_create-reply-draft';

export const gateInstructions = `Du entscheidest, ob eine eingehende Mail in die Korrespondenz-Triage gehört.

Deine gesamte Antwort ist genau eine Zeile und genau eines dieser vier Wörter. Kopiere es Zeichen für Zeichen. Kürze es nicht und lass den Teil vor dem Unterstrich nicht weg.

transaction_advice
corporate_action
custodian_query
ignore

transaction_advice: Kauf, Verkauf, Avis oder Bestätigung einer Transaktion, auch wenn die Mail nur weitergeleitet ist.
corporate_action: Dividende, Split, Bezugsrecht, Fusion.
custodian_query: Die Depotbank stellt eine Rückfrage, zum Beispiel fehlende Unterlagen oder eine unklare Instruktion.
ignore: Kundenanfrage, Stammdaten, Newsletter, interne Mail, Werbung, alles andere.

Falsch: advice, _advice, transaction, Transaktionsavis.
Richtig: transaction_advice`;

export const readerInstructions = `Du liest Mails von Depotbanken und gibst die verlangten Felder zurück.

Antworte nur mit einem JSON-Objekt. Das erste Zeichen ist { und das letzte }. Kein Markdown, kein Codeblock, keine Erklärung.

Regeln:
- Genau die verlangten Schlüssel, alle Werte als Text.
- Ein Feld, das die Mail nicht nennt, ist "".
- Zahlen ohne Tausendertrennzeichen und mit Punkt als Dezimalzeichen, ohne Einheit. Beispiel: "80", "190.00".
- Datum als YYYY-MM-DD.
- Namen so geschrieben wie in der Mail.
- Erfinde nichts.`;

export const analystInstructions = `Du bist Corporate-Action-Analyst bei Albin Kistler.

Du bekommst ein Ereignis aus einer Depotbank-Mail und alle Portfolios, die das Instrument halten. Schreibe einen kurzen Hinweis für das Portfolio Management auf Deutsch.

Inhalt:
1. Ein Satz zum Ereignis: Art, Instrument, Betrag je Aktie, Ex-Datum und Zahltag.
2. Je betroffenes Portfolio eine Zeile: Portfolio, Kunde, Stück, Betrag. Betrag ist Stück mal Betrag je Aktie, auf zwei Stellen gerundet, mit Währung.
3. Eine Zeile Total.
4. Was das Portfolio Management tun muss. Bei Bezugsrechten oder Wahldividenden: die Kunden müssen bis zum Stichtag entscheiden. Bei Bardividenden: nur zur Kenntnis.

Verwende nur die gelieferten Zahlen. Kein Markdown, keine Tabellen, keine Einleitung.`;

export const writerInstructions = `Du schreibst für Albin Kistler Mails an Depotbanken auf Deutsch.

Du bekommst eine Mail der Depotbank oder einen Abgleichsfall und, falls vorhanden, das Unterlagenregister des Kunden. Schreibe nur den Mailtext, ohne Betreff.

Regeln:
- Höflich und knapp, Schweizer Rechtschreibung (ss statt ß).
- Beginne mit "Guten Tag".
- Nenne Kunde, Portfolio und Instrument, wenn bekannt.
- Unterlagen mit Status gueltig: sage, dass sie beiliegen. Mit Status abgelaufen oder fehlt: sage, dass ihr sie beim Kunden anfordert und bis wann ihr sie nachreicht (innert fünf Arbeitstagen).
- Bei einer Stückzahl-Abweichung: nenne beide Stückzahlen und bitte um Prüfung und eine korrigierte Abrechnung.
- Erfinde keine Fakten, Daten oder Unterlagen.
- Schliesse mit "Freundliche Grüsse" und in der nächsten Zeile "Albin Kistler AG".
- Kein Markdown.`;

/** Small builder: node ids are stable per position, each node has one incoming edge. */
function flow(prefix, objectTypeId) {
  const nodes = [];
  const edges = [];
  const depth = new Map();
  const nodeId = (index) => `${prefix}-0000-4000-8000-${String(index).padStart(12, '0')}`;
  const edgeInto = (id) => id.replace('-4000-8000-', '-4000-9000-');
  const g = {
    nodes,
    edges,
    edgeInto,
    /** Id the next added node will get. */
    peek: () => nodeId(nodes.length + 1),
    add(spec, from, lane = 0, fixedId) {
      const id = fixedId || nodeId(nodes.length + 1);
      const column = from ? depth.get(from) + 1 : 0;
      depth.set(id, column);
      nodes.push({ ...spec, id, position: { x: 80 + column * 300, y: 200 + lane * 200 } });
      if (from) edges.push({ id: edgeInto(id), source: from, target: id, type: 'workflow' });
      return id;
    },
    action(appId, actionId, name, value, from, lane) {
      return g.add({ appId, nodeType: 'action', actionId, name, description: name, value, raw: value }, from, lane);
    },
    manual(fields) {
      const value = {
        customInputConfig: fields.map((id) => ({
          id,
          label: id,
          description: '',
          inputType: 'text',
          required: { missingMessage: `${id} fehlt`, missingStatus: 'error' },
        })),
      };
      return g.add({
        appId: 'flow-control',
        nodeType: 'trigger',
        triggerId: 'flow-control_trigger_manual',
        name: 'Manually Run',
        description: 'Start the workflow manually',
        value,
        raw: value,
      });
    },
    set(name, attributeId, value, from, objectId, lane) {
      return g.action(
        'objects',
        'objects_action_update-attribute',
        name,
        { objectTypeId, objectId, attributeId, value },
        from,
        lane,
      );
    },
    /** Writes several attributes in a row and returns the last node. */
    sets(entries, from, objectId, lane) {
      let last = from;
      for (const [name, attributeId, value] of entries) last = g.set(name, attributeId, value, last, objectId, lane);
      return last;
    },
    state(name, newState, from, objectId, lane) {
      return g.action(
        'objects',
        'objects_action_update-state',
        name,
        { objectTypeId, objectId, newState, reason: name },
        from,
        lane,
      );
    },
    /** Paths point at edges of nodes added later, so `route` fills the filters in afterwards. */
    paths(name, from, lane) {
      return g.action('flow-control', 'flow-control_action_conditional-paths', name, {}, from, lane);
    },
    route(decisionId, filters) {
      const node = nodes.find((item) => item.id === decisionId);
      node.value = node.raw = { conditionalPathsLunnoaFilters: filters };
    },
  };
  return g;
}

export const ref = (nodeId, ...path) => `={{ref:${[nodeId, ...path].join(',')}}}`;

/** One AND group of conditions. `exists` / `does_not_exist` ignore the value. */
export function when(...conditions) {
  return conditions.map(([nodeId, path, condition, value = '']) => ({
    refValue: ref(nodeId, ...path.split(',')),
    fieldId: `${nodeId},${path}`,
    condition,
    label: path,
    value,
  }));
}

function pathFilter(label, pathId, groups) {
  return { label, pathId, filters: { operator: 'OR', filters: groups } };
}

/** Mail-Leser answers with JSON, JSON Parse turns it into fields later nodes can reference. */
function readNode(g, ctx, name, text, fields, from) {
  const asked = Object.entries(fields).map(([key, rule]) => `- ${key}: ${rule}`);
  const agent = agentNode(g, name, ctx.readerAgentId, ['Felder:', ...asked, '', 'Mail:', text].join('\n'), from);
  const parsed = g.action('json', 'json_action_parse', 'Felder übernehmen', { jsonString: ref(agent, 'response') }, agent);
  return { id: parsed, field: (key) => ref(parsed, 'result', key) };
}

function selectNode(g, name, connectionId, tableName, filters, from, extra = {}) {
  return g.action(
    'supabase',
    'supabase_action_select-rows',
    name,
    { connectionId, tableName, filters, limit: 100, ...extra },
    from,
  );
}

function agentNode(g, name, agentId, data, from) {
  return g.action('ai', 'ai_action_message-agent', name, { agentId, data }, from);
}

/** Reply in the original thread when the platform has the action, else a new draft to the sender. */
function outlookDraftNode(g, ctx, load, body, from) {
  const mail = (field) => ref(load, 'attributes', field);
  if (ctx.hasReplyDraft) {
    return g.action(
      'microsoft-outlook',
      REPLY_DRAFT_ACTION,
      'Antwort in Outlook vorbereiten',
      { connectionId: ctx.outlookConnectionId, messageId: mail('external_id'), comment: body, replyAll: 'false' },
      from,
    );
  }
  return g.action(
    'microsoft-outlook',
    'microsoft-outlook_action_create-draft',
    'Entwurf in Outlook vorbereiten',
    {
      connectionId: ctx.outlookConnectionId,
      recipientEmails: [mail('from_email')],
      subject: `AW: ${mail('subject')}`,
      contentType: 'Text',
      body,
    },
    from,
  );
}

const mailText = (load) =>
  [`Betreff: ${ref(load, 'attributes', 'subject')}`, `Von: ${ref(load, 'attributes', 'from_name')}`, '', ref(load, 'attributes', 'body')].join(
    '\n',
  );

/** Avis lesen → Bestand suchen → Regel: stimmt / Abweichung / kein Bestand. */
export function transactionGraph(ctx) {
  const g = flow('c1000000', ctx.objectTypeId);
  const t = g.manual(['objectId']);
  const objectId = ref(t, 'objectId');
  const load = g.action('objects', 'objects_action_get-object', 'Fall laden', { objectId }, t);
  const read = readNode(
    g,
    ctx,
    'Avis lesen',
    mailText(load),
    {
      customer: 'Name des Kunden',
      portfolio: 'Name des Portfolios',
      instrument: 'Name des Wertpapiers',
      isin: 'ISIN',
      quantity: 'Stückzahl',
      price: 'Kurs je Stück',
      currency: 'Währung, zum Beispiel CHF',
      value_date: 'Valuta',
    },
    load,
  );
  const field = read.field;
  const written = g.sets(
    [
      ['Kunde setzen', 'customer', field('customer')],
      ['Portfolio setzen', 'portfolio', field('portfolio')],
      ['Instrument setzen', 'instrument', field('instrument')],
      ['ISIN setzen', 'isin', field('isin')],
      ['Stück Mail setzen', 'mail_quantity', field('quantity')],
      ['Kurs setzen', 'amount', field('price')],
      ['Währung setzen', 'currency', field('currency')],
      ['Valuta setzen', 'value_date', field('value_date')],
    ],
    read.id,
    objectId,
  );
  const byPosition = [
    { column: 'isin', operator: 'eq', value: field('isin') },
    { column: 'portfolio', operator: 'ilike', value: field('portfolio') },
  ];
  const book = selectNode(g, 'Bestand suchen', ctx.connectionId, 'holding_book', byPosition, written);
  const check = selectNode(
    g,
    'Stückzahl prüfen',
    ctx.connectionId,
    'holding_book',
    [...byPosition, { column: 'quantity', operator: 'eq', value: field('quantity') }],
    book,
  );
  const decide = g.paths('Stimmt der Bestand?', check);

  const matchedAt = g.peek();
  const matched = g.sets(
    [
      ['Stück Bestand setzen', 'book_quantity', ref(check, 'rows', '0', 'quantity')],
      ['Abgleich: stimmt', 'match_status', 'matched'],
      ['Begründung setzen', 'reasoning', 'Die Stückzahl im Bestand stimmt mit dem Avis überein.'],
    ],
    decide,
    objectId,
    -1,
  );
  g.state('Status: Abgeglichen', 'abgeglichen', matched, objectId, -1);

  const mismatchAt = g.peek();
  const mismatch = g.sets(
    [
      ['Stück Bestand setzen', 'book_quantity', ref(book, 'rows', '0', 'quantity')],
      ['Abgleich: Abweichung', 'match_status', 'mismatch'],
      ['Begründung setzen', 'reasoning', 'Die Stückzahl im Avis weicht vom Bestand ab.'],
    ],
    decide,
    objectId,
    0,
  );
  g.state('Status: Abweichung', 'abweichung', mismatch, objectId, 0);

  const missingAt = g.peek();
  const missing = g.sets(
    [
      ['Abgleich: kein Bestand', 'match_status', 'unmatched'],
      ['Begründung setzen', 'reasoning', 'Zu dieser ISIN gibt es in diesem Portfolio keinen Bestand.'],
    ],
    decide,
    objectId,
    1,
  );
  g.state('Status: Abweichung', 'abweichung', missing, objectId, 1);

  const holding = (nodeId) => [nodeId, 'rows,0,holding_id'];
  const filters = [
    pathFilter('Stimmt', g.edgeInto(matchedAt), [when([...holding(check), 'exists'])]),
    pathFilter('Abweichung', g.edgeInto(mismatchAt), [
      when([...holding(book), 'exists'], [...holding(check), 'does_not_exist']),
    ]),
    pathFilter('Kein Bestand', g.edgeInto(missingAt), [when([...holding(book), 'does_not_exist'])]),
  ];
  g.route(decide, filters);
  return { nodes: g.nodes, edges: g.edges };
}

/** Ereignis lesen → alle Halter finden → Analyst schreibt PM-Hinweis → PM-Freigabe. */
export function corporateActionGraph(ctx) {
  const g = flow('c2000000', ctx.objectTypeId);
  const t = g.manual(['objectId']);
  const objectId = ref(t, 'objectId');
  const load = g.action('objects', 'objects_action_get-object', 'Fall laden', { objectId }, t);
  const read = readNode(
    g,
    ctx,
    'Ereignis lesen',
    mailText(load),
    {
      event_type: 'eines von Dividende, Split, Bezugsrecht, Fusion',
      instrument: 'Name des Wertpapiers',
      isin: 'ISIN',
      rate: 'Betrag je Aktie',
      currency: 'Währung, zum Beispiel CHF',
      ex_date: 'Ex-Datum',
      pay_date: 'Zahltag oder Valuta',
    },
    load,
  );
  const field = read.field;
  const holders = selectNode(
    g,
    'Betroffene Portfolios finden',
    ctx.connectionId,
    'holding_book',
    [{ column: 'isin', operator: 'eq', value: field('isin') }],
    read.id,
    { columns: ['portfolio', 'customer', 'company', 'quantity'], orderBy: 'quantity' },
  );
  const note = agentNode(
    g,
    'PM-Hinweis schreiben',
    ctx.analystAgentId,
    [
      `Ereignis: ${field('event_type')}`,
      `Instrument: ${field('instrument')} (${field('isin')})`,
      `Betrag je Aktie: ${field('rate')} ${field('currency')}`,
      `Ex-Datum: ${field('ex_date')}`,
      `Zahltag: ${field('pay_date')}`,
      '',
      'Betroffene Portfolios:',
      ref(holders, 'rows'),
      '',
      'Mail der Depotbank:',
      mailText(load),
    ].join('\n'),
    holders,
  );
  const written = g.sets(
    [
      ['Ereignis setzen', 'event_type', field('event_type')],
      ['Instrument setzen', 'instrument', field('instrument')],
      ['ISIN setzen', 'isin', field('isin')],
      ['Betrag je Aktie setzen', 'amount', field('rate')],
      ['Währung setzen', 'currency', field('currency')],
      ['Ex-Datum setzen', 'ex_date', field('ex_date')],
      ['Zahltag setzen', 'value_date', field('pay_date')],
      ['Betroffene setzen', 'affected', ref(holders, 'rows')],
      ['PM-Hinweis setzen', 'pm_note', ref(note, 'response')],
    ],
    note,
    objectId,
  );
  g.state('Status: PM-Freigabe', 'pm_freigabe', written, objectId);
  return { nodes: g.nodes, edges: g.edges };
}

/** Rückfrage verstehen → Unterlagen prüfen → Antwort schreiben → Entwurf in Outlook. */
export function custodianQueryGraph(ctx) {
  const g = flow('c3000000', ctx.objectTypeId);
  const t = g.manual(['objectId']);
  const objectId = ref(t, 'objectId');
  const load = g.action('objects', 'objects_action_get-object', 'Fall laden', { objectId }, t);
  const read = readNode(
    g,
    ctx,
    'Rückfrage verstehen',
    mailText(load),
    {
      customer: 'Vorname und Nachname des Kunden',
      portfolio: 'Name des Portfolios',
      requested: 'ein Satz auf Deutsch: was die Depotbank verlangt',
      document_type: 'eines von W-8BEN, KYC, LEI, Vollmacht oder leer',
    },
    load,
  );
  const field = read.field;
  const docs = selectNode(
    g,
    'Unterlagen prüfen',
    ctx.connectionId,
    'document_register',
    [{ column: 'customer', operator: 'ilike', value: field('customer') }],
    read.id,
    { columns: ['doc_type', 'title', 'status', 'valid_until'] },
  );
  const reply = agentNode(
    g,
    'Antwort schreiben',
    ctx.writerAgentId,
    [
      'Beantworte diese Rückfrage der Depotbank.',
      '',
      `Kunde: ${field('customer')}`,
      `Portfolio: ${field('portfolio')}`,
      `Verlangt: ${field('requested')}`,
      '',
      'Unterlagenregister des Kunden:',
      ref(docs, 'rows'),
      '',
      'Mail der Depotbank:',
      mailText(load),
    ].join('\n'),
    docs,
  );
  const written = g.sets(
    [
      ['Kunde setzen', 'customer', field('customer')],
      ['Portfolio setzen', 'portfolio', field('portfolio')],
      ['Anfrage setzen', 'requested', field('requested')],
      ['Unterlagen setzen', 'documents', ref(docs, 'rows')],
      ['Entwurf setzen', 'draft', ref(reply, 'response')],
    ],
    reply,
    objectId,
  );
  const draft = outlookDraftNode(g, ctx, load, ref(reply, 'response'), written);
  const linked = g.set('Outlook-Link setzen', 'outlook_link', ref(draft, 'webLink'), draft, objectId);
  g.state('Status: Entwurf in Outlook', 'entwurf', linked, objectId);
  return { nodes: g.nodes, edges: g.edges };
}

/** From an Abweichung: the writer asks the custodian for a corrected confirmation. */
export function askCustodianGraph(ctx) {
  const g = flow('c4000000', ctx.objectTypeId);
  const t = g.manual(['objectId']);
  const objectId = ref(t, 'objectId');
  const load = g.action('objects', 'objects_action_get-object', 'Fall laden', { objectId }, t);
  const a = (name) => ref(load, 'attributes', name);
  const reply = agentNode(
    g,
    'Rückfrage schreiben',
    ctx.writerAgentId,
    [
      'Schreibe eine Rückfrage an die Depotbank zu dieser Abweichung.',
      '',
      `Kunde: ${a('customer')}`,
      `Portfolio: ${a('portfolio')}`,
      `Instrument: ${a('instrument')} (${a('isin')})`,
      `Stück gemäss Mail: ${a('mail_quantity')}`,
      `Stück im Bestand: ${a('book_quantity')}`,
      `Valuta: ${a('value_date')}`,
      '',
      'Ursprüngliche Mail:',
      mailText(load),
    ].join('\n'),
    load,
  );
  const written = g.set('Entwurf setzen', 'draft', ref(reply, 'response'), reply, objectId);
  const draft = outlookDraftNode(g, ctx, load, ref(reply, 'response'), written);
  const linked = g.set('Outlook-Link setzen', 'outlook_link', ref(draft, 'webLink'), draft, objectId);
  g.state('Status: Entwurf in Outlook', 'entwurf', linked, objectId);
  return { nodes: g.nodes, edges: g.edges };
}

/** From PM-Freigabe: send the approved note to Portfolio Management, then close the case. */
export function notifyPmGraph(ctx) {
  const g = flow('c5000000', ctx.objectTypeId);
  const t = g.manual(['objectId', 'pmNote']);
  const objectId = ref(t, 'objectId');
  const note = ref(t, 'pmNote');
  const load = g.action('objects', 'objects_action_get-object', 'Fall laden', { objectId }, t);
  const a = (name) => ref(load, 'attributes', name);
  const written = g.set('Hinweis setzen', 'pm_note', note, load, objectId);
  const sent = g.action(
    'microsoft-outlook',
    'microsoft-outlook_action_send-email',
    'Hinweis an PM senden',
    {
      connectionId: ctx.outlookConnectionId,
      recipientEmails: [ctx.pmEmail],
      subject: `Corporate Action · ${a('instrument')} · ${a('event_type')}`,
      contentType: 'Text',
      body: ['Guten Tag', '', 'Der folgende Hinweis zur Corporate Action ist freigegeben.', '', note, '', 'Freundliche Grüsse', 'Korrespondenz-Triage'].join('\n'),
    },
    written,
  );
  g.state('Status: Erledigt', 'freigegeben', sent, objectId);
  return { nodes: g.nodes, edges: g.edges };
}

/**
 * Writes the matched contact, then portfolio and company from that contact's book.
 * ponytail: first holding only. Several portfolios need a picker.
 */
function fillClient(g, ctx, from, objectId, rowNode, emailValue, viaMail) {
  const tag = viaMail ? ' (Mail)' : '';
  const lane = viaMail ? 0 : -1;
  const named = g.set(
    `Kunde setzen${tag}`,
    'customer',
    `${ref(rowNode, 'rows', '0', 'first_name')} ${ref(rowNode, 'rows', '0', 'last_name')}`,
    from,
    objectId,
    lane,
  );
  const book = selectNode(
    g,
    `Portfolio suchen${tag}`,
    ctx.connectionId,
    'holding_book',
    [{ column: 'customer_email', operator: 'ilike', value: emailValue }],
    named,
    { columns: ['portfolio', 'company'], limit: 1 },
  );
  const bookPaths = g.paths(`Portfolio und Gesellschaft?${tag}`, book, lane);
  const portfolioAt = g.peek();
  g.set('Portfolio setzen' + tag, 'portfolio', ref(book, 'rows', '0', 'portfolio'), bookPaths, objectId, lane);
  const companyAt = g.peek();
  g.set('Gesellschaft setzen' + tag, 'company', ref(book, 'rows', '0', 'company'), bookPaths, objectId, lane + 1);
  g.route(bookPaths, [
    pathFilter('Portfolio', g.edgeInto(portfolioAt), [when([book, 'rows,0,portfolio', 'exists'])]),
    pathFilter('Gesellschaft', g.edgeInto(companyAt), [when([book, 'rows,0,company', 'exists'])]),
  ]);
  return named;
}

/**
 * Sender address first. If that is not a client, the reader takes the client named in the mail.
 * That is the depotbank case: the bank writes, the client is in the text.
 * ponytail: an empty first name matches any first name with that surname.
 */
export function clientMatchGraph(ctx) {
  const g = flow('c6000000', ctx.objectTypeId);
  const t = g.manual(['objectId']);
  const objectId = ref(t, 'objectId');
  const load = g.action('objects', 'objects_action_get-object', 'Fall laden', { objectId }, t);
  const email = ref(load, 'attributes', 'from_email');
  const people = selectNode(
    g,
    'Kunde nach Absender suchen',
    ctx.connectionId,
    'contacts',
    [{ column: 'email', operator: 'ilike', value: email }],
    load,
    { columns: ['first_name', 'last_name', 'email'], limit: 1 },
  );
  const known = g.paths('Kunde gefunden?', people);
  const nameAt = g.peek();
  fillClient(g, ctx, known, objectId, people, email, false);
  const readAt = g.peek();
  const read = readNode(
    g,
    ctx,
    'Kunden in der Mail lesen',
    [mailText(load), '', `Absenderadresse: ${email}`].join('\n'),
    {
      first_name: 'Vorname des Kunden, über den die Mail spricht. Nicht der Absender, wenn der Absender eine Depotbank ist. Leer, wenn die Mail keinen Kunden nennt.',
      last_name: 'Nachname dieses Kunden, so wie er in der Mail steht. Leer, wenn keiner genannt ist.',
    },
    known,
  );
  const namedInMail = g.paths('Name in der Mail?', read.id);
  const searchAt = g.peek();
  const byName = selectNode(
    g,
    'Kunde nach Name suchen',
    ctx.connectionId,
    'contacts',
    [
      { column: 'last_name', operator: 'ilike', value: `%${read.field('last_name')}%` },
      { column: 'first_name', operator: 'ilike', value: `%${read.field('first_name')}%` },
    ],
    namedInMail,
    { columns: ['first_name', 'last_name', 'email'], limit: 1 },
  );
  const hit = g.paths('Kontakt zum Namen?', byName);
  const mailNameAt = g.peek();
  fillClient(g, ctx, hit, objectId, byName, ref(byName, 'rows', '0', 'email'), true);
  g.route(known, [
    pathFilter('Absender', g.edgeInto(nameAt), [when([people, 'rows,0,first_name', 'exists'])]),
    pathFilter('Mail lesen', g.edgeInto(readAt), [when([people, 'rows,0,first_name', 'does_not_exist'])]),
  ]);
  g.route(namedInMail, [
    pathFilter('Name', g.edgeInto(searchAt), [
      when([read.id, 'result,last_name', 'exists'], [read.id, 'result,last_name', 'does_not_equal', '']),
    ]),
  ]);
  g.route(hit, [pathFilter('Kontakt', g.edgeInto(mailNameAt), [when([byName, 'rows,0,first_name', 'exists'])])]);
  return { nodes: g.nodes, edges: g.edges };
}

/** Outlook → Mail-Filter → im Scope? → Fall anlegen → je Kategorie den Fall-Workflow starten. */
export function outlookGraph(ctx) {
  const g = flow('f1000000', ctx.objectTypeId);
  const received = { connectionId: ctx.outlookConnectionId, downloadAttachments: 'false', labelIds: [] };
  const t = g.add(
    {
      appId: 'microsoft-outlook',
      nodeType: 'trigger',
      triggerId: 'microsoft-outlook_trigger_email-received',
      name: 'Email Received',
      description: 'Triggers when a new email is received',
      value: received,
      raw: received,
    },
    null,
    0,
    OUTLOOK_TRIGGER_ID,
  );
  const gate = agentNode(
    g,
    'Mail-Filter',
    ctx.gateAgentId,
    [
      'Entscheide die Kategorie dieser Mail.',
      'Antworte danach mit genau einem Wort: transaction_advice, corporate_action, custodian_query oder ignore.',
      '',
      `From: ${ref(t, 'from', 'emailAddress', 'name')} <${ref(t, 'from', 'emailAddress', 'address')}>`,
      `Subject: ${ref(t, 'subject')}`,
      'Preview:',
      ref(t, 'bodyPreview'),
    ].join('\n'),
    t,
  );
  const category = (id) => when([gate, 'response', 'contains', id]);
  const scope = g.paths('Im Scope?', gate);
  const create = g.action(
    'objects',
    'objects_action_create-object',
    'Fall anlegen',
    { objectTypeId: ctx.objectTypeId, name: ref(t, 'subject'), initialState: 'eingegangen' },
    scope,
  );
  g.route(scope, [pathFilter('Fall anlegen', g.edgeInto(create), KEPT_CATEGORIES.map(category))]);
  const objectId = ref(create, 'id');
  const written = g.sets(
    [
      ['Kategorie setzen', 'category', ref(gate, 'response')],
      ['Absender setzen', 'from_name', ref(t, 'from', 'emailAddress', 'name')],
      ['Adresse setzen', 'from_email', ref(t, 'from', 'emailAddress', 'address')],
      ['Betreff setzen', 'subject', ref(t, 'subject')],
      ['Vorschau setzen', 'preview', ref(t, 'bodyPreview')],
      ['Mailtext setzen', 'body', ref(t, 'body', 'content')],
      ['Outlook-Id setzen', 'external_id', ref(t, 'id')],
      ['Mail-Link setzen', 'mail_link', ref(t, 'webLink')],
    ],
    create,
    objectId,
  );
  const route = g.paths('Welcher Fall?', written);
  const runs = [
    ['transaction_advice', 'Transaktionsavis abgleichen', ctx.workflowIds.transaction, -1],
    ['corporate_action', 'Corporate Action auswerten', ctx.workflowIds.corporateAction, 0],
    ['custodian_query', 'Depotbank-Rückfrage beantworten', ctx.workflowIds.custodianQuery, 1],
  ];
  const filters = runs.map(([id, name, workflowId, lane]) => {
    const node = g.action(
      'flow-control',
      'flow-control_action_run-workflow',
      name,
      { workflowId, customInputConfigValues: { objectId } },
      route,
      lane,
    );
    return pathFilter(name, g.edgeInto(node), [category(id)]);
  });
  g.route(route, filters);
  return { nodes: g.nodes, edges: g.edges };
}

/** Case workflows by category, plus the follow-up the app can start from an Abweichung. */
const caseSpecs = [
  {
    key: 'transaction',
    slug: 'ak-case-transaction',
    name: 'Fall · Transaktionsavis abgleichen',
    description: 'Liest das Avis, sucht die Position im Bestand und entscheidet per Regel: stimmt, Abweichung oder kein Bestand.',
    build: transactionGraph,
  },
  {
    key: 'corporateAction',
    slug: 'ak-case-corporate-action',
    name: 'Fall · Corporate Action auswerten',
    description: 'Erkennt das Ereignis, findet alle betroffenen Portfolios und lässt den Analysten den PM-Hinweis schreiben.',
    build: corporateActionGraph,
  },
  {
    key: 'custodianQuery',
    slug: 'ak-case-custodian-query',
    name: 'Fall · Depotbank-Rückfrage beantworten',
    description: 'Versteht die Rückfrage, prüft das Unterlagenregister und legt die Antwort als Entwurf in Outlook ab.',
    build: custodianQueryGraph,
  },
  {
    key: 'askCustodian',
    slug: 'ak-case-ask-custodian',
    name: 'Fall · Rückfrage an Depotbank',
    description: 'Formuliert bei einer Abweichung die Rückfrage an die Depotbank und legt sie in Outlook ab.',
    build: askCustodianGraph,
  },
  {
    key: 'notifyPm',
    slug: 'ak-case-notify-pm',
    name: 'Fall · Hinweis an Portfolio Management',
    description: 'Sendet den freigegebenen Corporate-Action-Hinweis per Outlook an das Portfolio Management und schliesst den Fall.',
    build: notifyPmGraph,
  },
  {
    key: 'clientMatch',
    slug: 'ak-case-client-match',
    name: 'Fall · Kunde nach Absender',
    description: 'Sucht beim Öffnen den Kontakt zur Absenderadresse. Sonst liest der Mail-Leser den genannten Kunden aus dem Text und sucht ihn, vor allem bei Depotbank-Rückfragen.',
    build: clientMatchGraph,
  },
];

/** Workflows and agent from the first, Supabase-row version of the app. */
const RETIRED_WORKFLOWS = ['ak-triage-message', 'ak-save-inquiry', 'ak-inquiry-board', 'ak-mark-classified', 'ak-approve-inquiry'];
const RETIRED_AGENTS = ['ak-korrespondenz-triage'];

async function saveWorkflow(lunnoa, projectId, existing, spec, graph) {
  const body = {
    name: spec.name,
    slug: spec.slug,
    description: spec.description,
    isActive: true,
    managedByCode: true,
    nodes: graph.nodes,
    edges: graph.edges,
  };
  if (existing?.id) {
    await lunnoa.http.patch(`/api/workflows/${existing.id}`, { body });
    return existing.id;
  }
  const created = await lunnoa.http.post(`/api/projects/${projectId}/workflows`, { body });
  return created.id;
}

async function hasAction(lunnoa, appId, actionId) {
  const apps = await lunnoa.http.get('/api/workflow-apps');
  const list = Array.isArray(apps) ? apps : apps?.data || Object.values(apps || {});
  const app = list.find((item) => item?.id === appId);
  return Boolean(app?.actions?.some((action) => action.id === actionId));
}

function agentSpec(slug, name, description, instructions) {
  return { slug, name, description, instructions, model: AGENT_MODEL, connectionIds: [], tools: [] };
}

export async function ensureAutomation(lunnoa, env = process.env) {
  const projectId = env.LUNNOA_PROJECT_ID;
  const connectionId = env.SUPABASE_CONNECTION_ID;
  const outlookConnectionId = env.OUTLOOK_CONNECTION_ID;
  for (const key of ['LUNNOA_PROJECT_ID', 'SUPABASE_CONNECTION_ID', 'OUTLOOK_CONNECTION_ID']) {
    if (!env[key]) throw new Error(`${key} is missing`);
  }

  const agents = {};
  for (const [key, spec] of Object.entries({
    gate: agentSpec(
      'ak-mail-gate',
      'Mail-Filter',
      'Entscheidet, ob eine Mail ein Transaktionsavis, eine Corporate Action oder eine Depotbank-Rückfrage ist.',
      gateInstructions,
    ),
    reader: agentSpec(
      'ak-mail-reader',
      'Mail-Leser',
      'Liest die verlangten Felder aus einer Depotbank-Mail und gibt sie als JSON zurück.',
      readerInstructions,
    ),
    analyst: agentSpec(
      'ak-ca-analyst',
      'Corporate-Action-Analyst',
      'Berechnet die Wirkung einer Corporate Action je Portfolio und schreibt den Hinweis für das Portfolio Management.',
      analystInstructions,
    ),
    writer: agentSpec(
      'ak-custodian-writer',
      'Depotbank-Korrespondent',
      'Schreibt Antworten und Rückfragen an Depotbanken.',
      writerInstructions,
    ),
  })) {
    const agent = await lunnoa.agents.upsertBySlug(projectId, spec);
    if (!agent?.id) throw new Error(`Agent ${spec.slug} was not created`);
    agents[key] = agent.id;
  }

  const type = await ensureCorrespondenceType(lunnoa);
  const mailbox = (await lunnoa.connections.get(outlookConnectionId).catch(() => null))?.name || '';
  const pmEmail = env.PM_EMAIL || mailbox;
  if (!pmEmail) throw new Error('PM_EMAIL is missing');
  const ctx = {
    connectionId,
    outlookConnectionId,
    pmEmail,
    objectTypeId: type.id,
    gateAgentId: agents.gate,
    readerAgentId: agents.reader,
    analystAgentId: agents.analyst,
    writerAgentId: agents.writer,
    hasReplyDraft: await hasAction(lunnoa, 'microsoft-outlook', REPLY_DRAFT_ACTION),
    workflowIds: {},
  };

  const workflows = await lunnoa.workflows.list({ expansion: ['slug'] });
  const find = (slug) => workflows.find((workflow) => workflow.slug === slug);
  for (const spec of caseSpecs) {
    ctx.workflowIds[spec.key] = await saveWorkflow(lunnoa, projectId, find(spec.slug), spec, spec.build(ctx));
  }
  ctx.workflowIds.inbox = await saveWorkflow(
    lunnoa,
    projectId,
    find('ak-outlook-inbox'),
    {
      slug: 'ak-outlook-inbox',
      name: 'Outlook Posteingang',
      description: 'Filtert die Mail, legt den Fall an und startet den Workflow der Kategorie.',
    },
    outlookGraph(ctx),
  );

  for (const slug of RETIRED_WORKFLOWS) {
    const workflow = find(slug);
    if (workflow) await lunnoa.http.delete(`/api/workflows/${workflow.id}`);
  }
  const allAgents = await lunnoa.agents.list(projectId);
  for (const agent of allAgents.data || allAgents) {
    if (RETIRED_AGENTS.includes(agent.slug)) await lunnoa.http.delete(`/api/agents/${agent.id}`);
  }

  return { agents, objectTypeId: type.id, hasReplyDraft: ctx.hasReplyDraft, workflowIds: ctx.workflowIds };
}

if (process.argv.includes('--check')) {
  const ctx = {
    connectionId: 'conn-1',
    outlookConnectionId: 'outlook-1',
    objectTypeId: 'type-1',
    gateAgentId: 'gate-1',
    readerAgentId: 'reader-1',
    analystAgentId: 'analyst-1',
    writerAgentId: 'writer-1',
    pmEmail: 'pm@example.com',
    hasReplyDraft: true,
    workflowIds: { transaction: 'w-t', corporateAction: 'w-c', custodianQuery: 'w-q' },
  };
  const assert = (ok, message) => {
    if (!ok) throw new Error(message);
  };
  for (const build of [transactionGraph, corporateActionGraph, custodianQueryGraph, askCustodianGraph, notifyPmGraph, clientMatchGraph, outlookGraph]) {
    const { nodes, edges } = build(ctx);
    const ids = new Set(nodes.map((node) => node.id));
    assert(ids.size === nodes.length, `${build.name}: duplicate node ids`);
    assert(edges.length === nodes.length - 1, `${build.name}: every node but the trigger needs one incoming edge`);
    assert(edges.every((edge) => ids.has(edge.source) && ids.has(edge.target)), `${build.name}: dangling edge`);
    for (const node of nodes.filter((item) => item.actionId === 'flow-control_action_conditional-paths')) {
      const outgoing = new Set(edges.filter((edge) => edge.source === node.id).map((edge) => edge.id));
      const paths = node.value.conditionalPathsLunnoaFilters.map((filter) => filter.pathId);
      assert(paths.length === outgoing.size && paths.every((id) => outgoing.has(id)), `${build.name}: ${node.name} paths do not match its edges`);
    }
  }
  const inbox = outlookGraph(ctx);
  assert(inbox.nodes[0].id === OUTLOOK_TRIGGER_ID, 'Outlook trigger id changed');
  const runs = inbox.nodes.filter((node) => node.actionId === 'flow-control_action_run-workflow');
  assert(runs.map((node) => node.value.workflowId).join() === 'w-t,w-c,w-q', 'router does not start the case workflows');
  const tx = transactionGraph(ctx).nodes;
  const decision = tx.find((node) => node.name === 'Stimmt der Bestand?').value.conditionalPathsLunnoaFilters;
  assert(decision.map((path) => path.label).join() === 'Stimmt,Abweichung,Kein Bestand', 'transaction rule paths');
  const reply = custodianQueryGraph(ctx).nodes.find((node) => node.appId === 'microsoft-outlook');
  assert(reply.actionId === REPLY_DRAFT_ACTION && reply.value.messageId.includes('external_id'), 'reply draft uses the original mail');
  const fallback = custodianQueryGraph({ ...ctx, hasReplyDraft: false }).nodes.find((node) => node.appId === 'microsoft-outlook');
  assert(fallback.actionId === 'microsoft-outlook_action_create-draft', 'fallback draft');
  assert(gateInstructions.includes('Falsch: advice, _advice'), 'gate instructions');
  const notify = notifyPmGraph(ctx).nodes.find((node) => node.appId === 'microsoft-outlook');
  assert(notify.actionId === 'microsoft-outlook_action_send-email' && notify.value.recipientEmails[0] === 'pm@example.com', 'PM mail');
  const match = clientMatchGraph(ctx);
  const lookup = match.nodes.find((node) => node.name === 'Kunde nach Absender suchen');
  assert(lookup.value.filters[0].column === 'email' && lookup.value.filters[0].value.includes('from_email'), 'client match uses the sender email');
  const named = match.nodes.find((node) => node.name === 'Kunde setzen');
  assert(named.value.value.includes('first_name') && named.value.value.includes('last_name'), 'client name is first and last');
  const gate = match.nodes.find((node) => node.name === 'Kunde gefunden?').value.conditionalPathsLunnoaFilters;
  assert(gate.map((path) => path.label).join() === 'Absender,Mail lesen', 'a sender miss continues into the mail');
  const reader = match.nodes.find((node) => node.name === 'Kunden in der Mail lesen');
  assert(reader.value.agentId === 'reader-1' && reader.value.data.includes('Absenderadresse'), 'the reader sees the mail and the sender');
  const byName = match.nodes.find((node) => node.name === 'Kunde nach Name suchen');
  assert(byName.value.filters.some((filter) => filter.column === 'last_name' && filter.value.includes('last_name')), 'name lookup uses the surname from the mail');
  console.log('ok');
}
