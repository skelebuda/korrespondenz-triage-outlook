import { useEffect, useMemo, useRef, useState } from 'react';

const FILTERS = [
  { id: 'transaction_advice', label: 'Transaktionsavis' },
  { id: 'corporate_action', label: 'Corporate Action' },
  { id: 'custodian_query', label: 'Depotbank-Rückfrage' },
];

const STATES: Record<string, string> = {
  eingegangen: 'Eingegangen',
  abgeglichen: 'Abgeglichen',
  abweichung: 'Abweichung',
  pm_freigabe: 'PM-Freigabe',
  entwurf: 'Entwurf in Outlook',
  freigegeben: 'Erledigt',
};

const MATCH: Record<string, string> = {
  matched: 'Bestand stimmt',
  mismatch: 'Stückzahl weicht ab',
  unmatched: 'Kein Bestand',
};

type StepStatus = 'done' | 'current' | 'todo' | 'skipped';

type ProcessStep = {
  id: string;
  name: string;
  status: StepStatus;
};

type Process = {
  title: string;
  summary: string;
  steps: ProcessStep[];
};

type RunStep = {
  label: string;
  status?: string;
  durationMs?: number | null;
};

type Run = {
  id: string;
  status?: string;
  statusMessage?: string | null;
  steps: RunStep[];
};

type Case = {
  id: string;
  state: string;
  category: string;
  view?: Process;
  from_name?: string;
  from_email?: string;
  subject?: string;
  preview?: string;
  body?: string;
  mail_link?: string;
  customer?: string;
  company?: string;
  portfolio?: string;
  instrument?: string;
  isin?: string;
  mail_quantity?: string;
  book_quantity?: string;
  amount?: string;
  currency?: string;
  value_date?: string;
  match_status?: string;
  reasoning?: string;
  event_type?: string;
  ex_date?: string;
  affected?: Record<string, unknown>[];
  pm_note?: string;
  requested?: string;
  documents?: Record<string, unknown>[];
  draft?: string;
  outlook_link?: string;
};

type SessionUser = {
  name: string;
  email: string;
  role: string;
};

type Board = {
  cases: Case[];
  links?: Record<string, string>;
  executionId?: string;
  user?: SessionUser;
};

function categoryLabel(id: string) {
  return FILTERS.find((category) => category.id === id)?.label || id;
}

function person(item: Case) {
  return item.customer || 'Kein Kundenmatch';
}

function currentStep(item: Case) {
  const current = item.view?.steps.find((step) => step.status === 'current');
  if (current?.id === 'done' && item.state !== 'freigegeben') return STATES[item.state] || current.name;
  return current?.name || STATES[item.state] || item.state;
}

const RUN_HINTS: Record<string, [RegExp, string][]> = {
  transaction_advice: [
    [/outlook|rückfrage schreiben|entwurf setzen/i, 'reply'],
    [/status:|abgleich:|stimmt der bestand/i, 'decide'],
    [/bestand|stückzahl|stück bestand/i, 'book'],
    [/avis|felder|kunde setzen|portfolio setzen|instrument|isin|stück mail|kurs|währung|valuta/i, 'read'],
  ],
  corporate_action: [
    [/hinweis an pm|senden/i, 'review'],
    [/status:|pm-freigabe|pm-hinweis setzen|erledigt/i, 'review'],
    [/pm-hinweis schreiben|analyst/i, 'note'],
    [/portfolio|halter|betroffene/i, 'holders'],
    [/ereignis|felder|instrument|isin|betrag|ex-datum|zahltag/i, 'read'],
  ],
  custodian_query: [
    [/outlook|status: entwurf/i, 'outlook'],
    [/antwort schreiben|entwurf setzen/i, 'draft'],
    [/unterlagen/i, 'docs'],
    [/rückfrage verstehen|felder|kunde setzen|anfrage setzen/i, 'read'],
  ],
};

function runningStepId(category: string, run: Run | null) {
  let reached = 'in';
  for (const node of run?.steps || []) {
    if (!['SUCCESS', 'RUNNING', 'RETRYING'].includes(String(node.status))) continue;
    const match = (RUN_HINTS[category] || []).find(([pattern]) => pattern.test(node.label || ''));
    if (match) reached = match[1];
  }
  return reached;
}

function liveSteps(item: Case, run: Run | null, running: boolean) {
  const steps = item.view?.steps || [];
  if (!running || !steps.length) return steps;
  const reached = runningStepId(item.category, run);
  const index = Math.max(0, steps.findIndex((step) => step.id === reached));
  return steps.map((step, position) => {
    if (step.status === 'skipped' || step.id === 'done') return { ...step, status: step.id === 'done' ? 'todo' : step.status };
    if (position < index) return { ...step, status: 'done' as const };
    if (position === index) return { ...step, status: 'current' as const };
    return { ...step, status: 'todo' as const };
  });
}

function activeRunLabel(run: Run | null, busy: string) {
  const active = run?.steps.find((step) => ['RUNNING', 'RETRYING', 'WAITING', 'SCHEDULED'].includes(String(step.status)));
  const last = [...(run?.steps || [])].reverse().find((step) => step.status === 'SUCCESS');
  return active?.label || last?.label || busy || 'Automatik läuft';
}

async function readJson(response: Response) {
  const body = await response.json();
  if (!response.ok) throw new Error(body.error || 'Anfrage fehlgeschlagen');
  return body;
}

function openOutlook(url?: string, fallback?: { email?: string; subject?: string; body?: string }) {
  if (url) {
    window.open(url, '_blank', 'noopener');
    return;
  }
  if (!fallback?.email) return;
  const query = new URLSearchParams();
  if (fallback.subject) query.set('subject', fallback.subject.startsWith('AW:') ? fallback.subject : `AW: ${fallback.subject}`);
  if (fallback.body) query.set('body', fallback.body);
  window.location.href = `mailto:${fallback.email}?${query}`;
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export default function App() {
  const [board, setBoard] = useState<Board | null>(null);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [filter, setFilter] = useState('all');
  const [note, setNote] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState('');
  const [run, setRun] = useState<Run | null>(null);
  const [runningId, setRunningId] = useState<string | null>(null);
  const openAfterRun = useRef(false);
  const clientLookup = useRef(new Set<string>());

  async function refresh() {
    const next = (await readJson(await fetch('/api/board'))) as Board;
    setBoard(next);
    return next;
  }

  useEffect(() => {
    refresh().catch((reason: Error) => setError(reason.message));
  }, []);

  const cases = board?.cases || [];
  const visible = cases.filter((item) =>
    filter === 'all' ? FILTERS.some((category) => category.id === item.category) : item.category === filter,
  );
  const selected = visible.find((item) => item.id === selectedId) || visible[0] || null;
  const counts = useMemo(() => {
    const tally = Object.fromEntries(FILTERS.map((category) => [category.id, 0]));
    for (const item of cases) {
      if (tally[item.category] != null) tally[item.category] += 1;
    }
    return tally;
  }, [cases]);

  useEffect(() => {
    setNote(selected?.pm_note || '');
  }, [selected?.id, selected?.pm_note]);

  useEffect(() => {
    const id = selected?.id;
    if (!id || !selected?.from_email || selected.customer || busy || clientLookup.current.has(id)) {
      if (id && selected?.customer) clientLookup.current.delete(id);
      return;
    }
    clientLookup.current.add(id);
    work(id, 'match-client');
  }, [selected?.id, selected?.from_email, selected?.customer, busy]);

  async function track(executionId: string) {
    for (;;) {
      const next = (await readJson(await fetch(`/api/executions/${executionId}`))) as Run;
      setRun(next);
      if (['SUCCESS', 'FAILED', 'CANCELLED'].includes(String(next.status))) return next;
      await sleep(1000);
    }
  }

  async function work(id: string, action: string, extra?: { pmNote?: string }) {
    const labels: Record<string, string> = {
      restart: 'Automatik läuft',
      'match-client': 'Kunde wird gesucht',
      'ask-custodian': 'Rückfrage wird nach Outlook gelegt',
      'approve-note': 'Hinweis geht an Portfolio Management',
    };
    setBusy(labels[action] || 'Fall wird gespeichert');
    setRunningId(id);
    if (action === 'restart' || action === 'ask-custodian' || action === 'approve-note') setRun({ id: '', steps: [] });
    setError('');
    openAfterRun.current = action === 'ask-custodian' || action === 'restart';
    try {
      const next = (await readJson(
        await fetch(`/api/cases/${id}`, {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ action, ...extra }),
        }),
      )) as Board;
      setBoard(next);
      if (next.executionId) {
        const finished = await track(next.executionId);
        const boardAfter = await refresh();
        const item = boardAfter.cases.find((entry) => entry.id === id);
        if (finished.status === 'FAILED') throw new Error(finished.statusMessage || 'Automatik fehlgeschlagen');
        if (openAfterRun.current && item?.outlook_link) openOutlook(item.outlook_link);
      }
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : 'Speichern fehlgeschlagen');
    } finally {
      openAfterRun.current = false;
      setBusy('');
      setRunningId(null);
      setRun(null);
    }
  }

  function openMail(kind: 'mail' | 'draft') {
    if (!selected) return;
    if (kind === 'mail') {
      openOutlook(selected.mail_link, { email: selected.from_email, subject: selected.subject });
      return;
    }
    openOutlook(selected.outlook_link, { email: selected.from_email, subject: selected.subject, body: selected.draft });
  }

  return (
    <main className="min-h-svh">
      <header className="app-bar flex items-center justify-between gap-6 px-6 py-5 text-white">
        <div>
          <p className="text-xs tracking-[0.14em] text-white/70 uppercase">Albin Kistler</p>
          <h1 className="mt-1 text-3xl">Korrespondenz-Triage</h1>
          <p className="mt-2 max-w-xl text-sm text-white/70">
            Jede Mail wird zum Fall. Transaktionsavis, Corporate Action und Depotbank-Rückfrage haben je einen eigenen Prozess.
          </p>
        </div>
        {board?.user ? <Session user={board.user} /> : null}
      </header>

      <section className="grid grid-cols-2 gap-px bg-border sm:grid-cols-4" aria-label="Kategorien">
        <button type="button" className="filter" aria-pressed={filter === 'all'} onClick={() => setFilter('all')}>
          <span className="text-xs tracking-wide text-muted-foreground uppercase">Alle</span>
          <span className="mt-1 block font-serif text-2xl">
            {FILTERS.reduce((sum, category) => sum + counts[category.id], 0)}
          </span>
        </button>
        {FILTERS.map((category) => (
          <button
            key={category.id}
            type="button"
            className="filter"
            aria-pressed={filter === category.id}
            onClick={() => setFilter(category.id)}
          >
            <span className="text-xs tracking-wide text-muted-foreground uppercase">{category.label}</span>
            <span className="mt-1 block font-serif text-2xl">{counts[category.id]}</span>
          </button>
        ))}
      </section>

      {error ? <p className="px-6 py-3 text-sm text-signal">{error}</p> : null}

      <section className="grid min-w-0 gap-6 px-6 py-6 lg:grid-cols-[360px_minmax(0,1fr)]">
        <ul className="flex flex-col gap-2">
          {visible.map((item) => (
            <li key={item.id}>
              <button
                type="button"
                className={selected?.id === item.id ? 'case case-active' : 'case'}
                onClick={() => setSelectedId(item.id)}
              >
                <span className="flex items-center justify-between gap-3">
                  <span className="text-xs tracking-wide text-muted-foreground uppercase">{categoryLabel(item.category)}</span>
                  <span className={`chip chip-${item.state}`}>{STATES[item.state] || item.state}</span>
                </span>
                <span className="mt-1 block text-left">{item.subject}</span>
                <span className="mt-1 block text-left text-sm text-muted-foreground">
                  {person(item)}
                  {item.company ? ` · ${item.company}` : ''}
                </span>
                <span className="mt-2 flex items-center gap-2 text-left text-xs text-muted-foreground">
                  {runningId === item.id ? <span className="run-spinner" aria-hidden /> : null}
                  {runningId === item.id ? busy || 'Automatik läuft' : currentStep(item)}
                </span>
              </button>
            </li>
          ))}
          {board && visible.length === 0 ? <li className="text-sm text-muted-foreground">Keine Fälle in dieser Auswahl.</li> : null}
        </ul>

        {selected ? (
          <CaseView
            item={selected}
            busy={busy}
            note={note}
            run={run}
            running={runningId === selected.id}
            onNote={setNote}
            onWork={work}
            onOpen={openMail}
          />
        ) : (
          <article className="bg-white px-6 py-10 text-sm text-muted-foreground">
            {board ? 'Keine Fälle in dieser Auswahl.' : 'Fälle werden geladen.'}
          </article>
        )}
      </section>
    </main>
  );
}

function initials(name: string) {
  return (
    name
      .split(/\s+/)
      .filter(Boolean)
      .slice(0, 2)
      .map((part) => part[0]?.toUpperCase())
      .join('') || '?'
  );
}

function Session({ user }: { user: SessionUser }) {
  return (
    <div className="session">
      <span className="session-mark" aria-hidden>
        {initials(user.name)}
      </span>
      <span>
        <span className="block text-sm">{user.name}</span>
        <span className="block text-xs text-white/70">
          {user.role}
          {user.email ? ` · ${user.email}` : ''}
        </span>
      </span>
    </div>
  );
}

function CaseView({
  item,
  busy,
  note,
  run,
  running,
  onNote,
  onWork,
  onOpen,
}: {
  item: Case;
  busy: string;
  note: string;
  run: Run | null;
  running: boolean;
  onNote: (value: string) => void;
  onWork: (id: string, action: string, extra?: { pmNote?: string }) => void;
  onOpen: (kind: 'mail' | 'draft') => void;
}) {
  return (
    <article className="bg-white px-6 py-5">
      <p className="text-xs tracking-wide text-muted-foreground uppercase">
        {categoryLabel(item.category)} · {STATES[item.state] || item.state}
      </p>
      <h2 className="mt-2 text-2xl">{item.subject}</h2>
      <p className="mt-2 text-sm text-muted-foreground">
        {item.from_name || 'Unbekannt'}
        {item.from_email ? ` · ${item.from_email}` : ''}
      </p>
      {item.view ? (
        <>
          <h3 className="mt-6 text-lg">{item.view.title}</h3>
          <p className="mt-1 text-sm text-muted-foreground">{item.view.summary}</p>
          <ol className={running ? 'steps steps-live mt-4' : 'steps mt-4'} aria-busy={running}>
            {liveSteps(item, run, running).map((step) => (
              <li key={step.id} className={`step step-${step.status}`}>
                <span className="step-dot" />
                <span className="step-name">{step.name}</span>
              </li>
            ))}
          </ol>
          {running ? (
            <p className="run-status" aria-live="polite">
              <span className="run-spinner" aria-hidden />
              {activeRunLabel(run, busy)}
            </p>
          ) : null}
        </>
      ) : null}
      <div className={running ? 'work is-busy mt-6 grid gap-6 lg:grid-cols-[1fr_1fr]' : 'mt-6 grid gap-6 lg:grid-cols-[1fr_1fr]'}>
        <div>
          <p className="text-xs tracking-wide text-muted-foreground uppercase">Mail</p>
          <p className="mt-2 text-sm leading-relaxed whitespace-pre-wrap">{item.preview || item.body || 'Keine Vorschau.'}</p>
        </div>
        <div>
          {item.category === 'corporate_action' ? (
            <CorporateWork item={item} note={note} onNote={onNote} />
          ) : item.category === 'custodian_query' ? (
            <QueryWork item={item} />
          ) : (
            <AvisWork item={item} />
          )}
        </div>
      </div>
      <CaseActions item={item} busy={busy} note={note} onWork={onWork} onOpen={onOpen} />
    </article>
  );
}

function AvisWork({ item }: { item: Case }) {
  const mismatch = item.match_status === 'mismatch' || item.match_status === 'unmatched';
  return (
    <div>
      <h3 className="text-lg">{MATCH[item.match_status || ''] || 'Abgleich'}</h3>
      <div className="mt-4 grid grid-cols-2 gap-px bg-border">
        <Fact label="Stück Mail" value={item.mail_quantity || '—'} signal={mismatch} />
        <Fact label="Stück Bestand" value={item.book_quantity || '—'} signal={mismatch} />
        <Fact label="Instrument" value={item.instrument || '—'} />
        <Fact label="ISIN" value={item.isin || '—'} />
        <Fact label="Kunde" value={person(item)} />
        <Fact label="Portfolio" value={item.portfolio || '—'} />
        <Fact label="Kurs" value={[item.amount, item.currency].filter(Boolean).join(' ') || '—'} />
        <Fact label="Valuta" value={item.value_date || '—'} />
      </div>
      <p className="mt-4 text-sm leading-relaxed">
        {item.reasoning || MATCH[item.match_status || ''] || 'Die Automatik hat den Bestand noch nicht geprüft.'}
      </p>
      {item.draft ? <DraftPreview text={item.draft} /> : null}
    </div>
  );
}

function CorporateWork({ item, note, onNote }: { item: Case; note: string; onNote: (value: string) => void }) {
  return (
    <div>
      <h3 className="text-lg">{item.event_type || 'Ereignis'}</h3>
      <div className="mt-4 grid grid-cols-2 gap-px bg-border">
        <Fact label="Instrument" value={item.instrument || '—'} />
        <Fact label="ISIN" value={item.isin || '—'} />
        <Fact label="Betrag je Aktie" value={[item.amount, item.currency].filter(Boolean).join(' ') || '—'} />
        <Fact label="Ex-Datum" value={item.ex_date || '—'} />
        <Fact label="Zahltag" value={item.value_date || '—'} />
        <Fact label="Portfolios" value={String(item.affected?.length || 0)} />
      </div>
      <p className="mt-4 text-xs tracking-wide text-muted-foreground uppercase">Betroffene Portfolios</p>
      <Rows
        rows={item.affected || []}
        columns={[
          { key: 'portfolio', label: 'Portfolio' },
          { key: 'customer', label: 'Kunde' },
          { key: 'quantity', label: 'Stück' },
        ]}
      />
      <label className="mt-4 block text-xs tracking-wide text-muted-foreground uppercase" htmlFor="pm-note">
        Hinweis für Portfolio Management
      </label>
      <textarea
        id="pm-note"
        className="mt-2 min-h-28 w-full border border-border bg-background px-3 py-2 text-sm"
        value={note}
        disabled={item.state === 'freigegeben'}
        onChange={(event) => onNote(event.target.value)}
      />
    </div>
  );
}

function QueryWork({ item }: { item: Case }) {
  return (
    <div>
      <h3 className="text-lg">Rückfrage</h3>
      <p className="mt-2 text-sm leading-relaxed">{item.requested || 'Die Automatik hat die Rückfrage noch nicht gelesen.'}</p>
      <p className="mt-4 text-xs tracking-wide text-muted-foreground uppercase">Unterlagen</p>
      <Rows
        rows={item.documents || []}
        columns={[
          { key: 'doc_type', label: 'Art' },
          { key: 'title', label: 'Titel' },
          { key: 'status', label: 'Status' },
          { key: 'valid_until', label: 'Gültig bis' },
        ]}
      />
      {item.draft ? <DraftPreview text={item.draft} /> : null}
    </div>
  );
}

function CaseActions({
  item,
  busy,
  note,
  onWork,
  onOpen,
}: {
  item: Case;
  busy: string;
  note: string;
  onWork: (id: string, action: string, extra?: { pmNote?: string }) => void;
  onOpen: (kind: 'mail' | 'draft') => void;
}) {
  const locked = Boolean(busy) || item.state === 'freigegeben';
  return (
    <div className="mt-6 flex flex-wrap gap-3">
      {busy ? (
        <p className="run-status">
          <span className="run-spinner" aria-hidden />
          {busy}
        </p>
      ) : null}
      {item.state === 'freigegeben' ? <p className="text-sm text-muted-foreground">Erledigt</p> : null}
      {item.mail_link || item.from_email ? (
        <button className="action-ghost" type="button" onClick={() => onOpen('mail')}>
          Mail in Outlook
        </button>
      ) : null}
      {item.state === 'eingegangen' ? (
        <button className="action" type="button" disabled={locked} onClick={() => onWork(item.id, 'restart')}>
          Fall automatisch bearbeiten
        </button>
      ) : null}
      {item.state === 'abgeglichen' ? (
        <button className="action" type="button" disabled={locked} onClick={() => onWork(item.id, 'confirm')}>
          Abgleich bestätigen
        </button>
      ) : null}
      {item.state === 'abweichung' ? (
        <>
          <button className="action" type="button" disabled={locked} onClick={() => onWork(item.id, 'ask-custodian')}>
            Rückfrage in Outlook öffnen
          </button>
          <button className="action-ghost" type="button" disabled={locked} onClick={() => onWork(item.id, 'reviewed')}>
            Nur protokollieren
          </button>
        </>
      ) : null}
      {item.state === 'pm_freigabe' ? (
        <button className="action" type="button" disabled={locked} onClick={() => onWork(item.id, 'approve-note', { pmNote: note })}>
          An Portfolio Management senden
        </button>
      ) : null}
      {item.outlook_link || item.state === 'entwurf' || item.draft ? (
        <button className={item.state === 'entwurf' ? 'action' : 'action-ghost'} type="button" onClick={() => onOpen('draft')}>
          Antwort in Outlook öffnen
        </button>
      ) : null}
      {item.state === 'entwurf' ? (
        <button className="action-ghost" type="button" disabled={locked} onClick={() => onWork(item.id, 'done')}>
          Als erledigt markieren
        </button>
      ) : null}
      {item.state !== 'eingegangen' && item.state !== 'freigegeben' ? (
        <button className="action-ghost" type="button" disabled={locked} onClick={() => onWork(item.id, 'restart')}>
          Automatik neu starten
        </button>
      ) : null}
    </div>
  );
}

function DraftPreview({ text }: { text: string }) {
  return (
    <>
      <p className="mt-4 text-xs tracking-wide text-muted-foreground uppercase">Antwortentwurf</p>
      <p className="mt-2 text-sm leading-relaxed whitespace-pre-wrap">{text}</p>
    </>
  );
}

function Fact({ label, value, signal }: { label: string; value: string; signal?: boolean }) {
  return (
    <div className="bg-white px-3 py-2">
      <p className="text-xs tracking-wide text-muted-foreground uppercase">{label}</p>
      <p className={signal ? 'mt-1 text-signal' : 'mt-1'}>{value}</p>
    </div>
  );
}

function Rows({ rows, columns }: { rows: Record<string, unknown>[]; columns: { key: string; label: string }[] }) {
  if (!rows.length) return <p className="mt-2 text-sm text-muted-foreground">Noch keine Einträge.</p>;
  return (
    <div className="mt-2 grid gap-px bg-border" style={{ gridTemplateColumns: `repeat(${columns.length}, minmax(0, 1fr))` }}>
      {columns.map((column) => (
        <div key={column.key} className="bg-muted px-3 py-2 text-xs tracking-wide text-muted-foreground uppercase">
          {column.label}
        </div>
      ))}
      {rows.flatMap((row, index) =>
        columns.map((column) => (
          <div key={`${column.key}-${index}`} className="bg-white px-3 py-2 text-sm">
            {String(row[column.key] ?? '—')}
          </div>
        )),
      )}
    </div>
  );
}
