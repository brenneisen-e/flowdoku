/* MassmailZielChips — v32.30. Empfänger der Info-Mail/des Reminders als
 * Mehrfachauswahl (Nutzer-Ansage 29.09.2026: „dadurch dass wir mittlerweile
 * so viele Möglichkeiten haben, wäre besser das mit Chips zu machen … damit
 * ich mehrere Gruppen auswählen kann").
 *
 * Datenmodell ohne neuen Modus: Die Teilnehmer-Gruppen sind Status-Mengen und
 * landen im bestehenden 'custom' + `massmailStatuses`. Dazu kommen additive
 * Zusätze (`MassmailExtra`): die Offenen (Erinnerung) und die internen Kreise
 * ich / Organizer / Test-Team. Nachrücker und die Erinnerung per eingefügtem
 * Verteiler bleiben eigene Modi, weil sie einen Einfüge-Schritt haben.
 *
 * Picker UND Editor rendern diese EINE Komponente und rechnen über
 * `massmailEmpfaenger` — sonst zeigt die Auswahl eine andere Zahl, als der
 * Versand anschreibt.
 */
import * as React from 'react';
import { cx } from '../../dexUi';
import { Check } from '../../Icons';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';
import { MassmailAudience, AudiencePerson } from '../adminTypes';
import { reminderRecipientsAus } from './MassmailPasteModal';
import { parsePastedRecipients, splitName } from '../../../utils/pastedRecipients';
import Modal from '../../Modal';

// v32.36: 'liste' = eigene eingefügte Adressliste (Massen-Upload), 'listeOffen'
// = davon nur, wer noch keine Zeile in der Teilnehmerliste hat.
export type MassmailExtra = 'offene' | 'ich' | 'orgs' | 'test' | 'liste' | 'listeOffen';

export const MM_ACTIVE = ['Angemeldet', 'QR versendet', 'Eingecheckt'];

interface Gruppe { key: string; de: string; en: string; stati: string[] }
const GRUPPEN: Gruppe[] = [
  { key: 'aktiv', de: 'Aktive Teilnehmer', en: 'Active participants', stati: MM_ACTIVE },
  { key: 'warte', de: 'Warteliste', en: 'Waitlist', stati: ['Warteliste'] },
  { key: 'ab', de: 'Abgemeldet', en: 'Cancelled', stati: ['Abgemeldet'] },
  { key: 'noshow', de: 'No-Show', en: 'No-show', stati: ['No-Show'] },
];

export type MassmailEmpfaenger = { ParticipantEmail: string; Vorname?: string; Nachname?: string };

function eindeutigeAdressen(list: Array<string | undefined>): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const raw of list) {
    const e = (raw || '').trim();
    const lc = e.toLowerCase();
    if (!e || lc.indexOf('@') < 0 || seen.has(lc)) continue;
    seen.add(lc);
    out.push(e);
  }
  return out;
}

export function internAdressen(ev: DeloitteEvent, myEmail: string, kreis: 'ich' | 'orgs' | 'test'): string[] {
  if (kreis === 'ich') return eindeutigeAdressen([myEmail]);
  if (kreis === 'orgs') return eindeutigeAdressen([...(ev.organizerEmails || []), ...(ev.coOrganizerEmails || [])]);
  return eindeutigeAdressen(ev.testTeamEmails || []);
}

function offeneAlsEmpfaenger(offene: AudiencePerson[]): MassmailEmpfaenger[] {
  return offene.map(x => {
    const dn = (x.displayName || '').trim();
    const komma = dn.indexOf(',');
    const nachname = komma > 0 ? dn.slice(0, komma).trim() : dn.split(' ').slice(-1)[0] || '';
    const vorname = komma > 0 ? dn.slice(komma + 1).trim() : dn.split(' ').slice(0, -1).join(' ');
    return { ParticipantEmail: x.email, Vorname: vorname, Nachname: nachname };
  });
}

/** Die EINE Empfänger-Rechnung für Auswahl und Versand. */
export function massmailEmpfaenger(a: {
  audience: MassmailAudience;
  statuses: Set<string>;
  extras: Set<MassmailExtra>;
  registrations: SPRegistration[];
  offene: AudiencePerson[] | null | undefined;
  pasteRaw: string;
  ev: DeloitteEvent;
  myEmail: string;
}): MassmailEmpfaenger[] {
  const { audience, statuses, extras, registrations, offene, pasteRaw, ev, myEmail } = a;
  // Modi mit Einfüge-Schritt: unverändert und ohne Zusätze.
  if (audience === 'nachruecker') {
    const matches = (pasteRaw || '').match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g) || [];
    const pastedSet = new Set(matches.map(m => m.toLowerCase()));
    return registrations.filter(r => MM_ACTIVE.indexOf(r.Status) >= 0 && !pastedSet.has((r.ParticipantEmail || '').toLowerCase()));
  }
  if (audience === 'reminder') {
    if (Array.isArray(offene) && offene.length > 0) return offeneAlsEmpfaenger(offene);
    return reminderRecipientsAus(pasteRaw || '', registrations);
  }
  let basis: MassmailEmpfaenger[];
  if (audience === 'custom') basis = registrations.filter(r => statuses.has(r.Status));
  else if (audience === 'waitOnly') basis = registrations.filter(r => r.Status === 'Warteliste');
  else if (audience === 'activePlusWait') basis = registrations.filter(r => MM_ACTIVE.indexOf(r.Status) >= 0 || r.Status === 'Warteliste');
  else if (audience === 'everyone') basis = registrations;
  else basis = registrations.filter(r => MM_ACTIVE.indexOf(r.Status) >= 0);
  const out: MassmailEmpfaenger[] = [];
  const seen = new Set<string>();
  const add = (r: MassmailEmpfaenger): void => {
    const lc = (r.ParticipantEmail || '').trim().toLowerCase();
    if (!lc || seen.has(lc)) return;
    seen.add(lc);
    out.push(r);
  };
  basis.forEach(add);
  if (extras.has('offene') && Array.isArray(offene)) offeneAlsEmpfaenger(offene).forEach(add);
  if (extras.has('liste')) {
    // Jede Zeile in der Teilnehmerliste zählt als Antwort (an-, abgemeldet,
    // Warteliste) — dieselbe Regel wie reminderRecipientsAus.
    const geantwortet = new Set(registrations.map(r => (r.ParticipantEmail || '').trim().toLowerCase()).filter(Boolean));
    parsePastedRecipients(pasteRaw || '')
      .filter(x => !extras.has('listeOffen') || !geantwortet.has(x.email))
      .forEach(x => { const n = splitName(x.name); add({ ParticipantEmail: x.email, Vorname: n.vorname, Nachname: n.nachname }); });
  }
  (['ich', 'orgs', 'test'] as const).forEach(k => {
    if (extras.has(k)) internAdressen(ev, myEmail, k).forEach(e => add({ ParticipantEmail: e }));
  });
  return out;
}

/** Kurzname der Auswahl für Kopfzeile und Vorschau. */
export function massmailZielLabel(isDe: boolean, audience: MassmailAudience, statuses: Set<string>, extras: Set<MassmailExtra>): string {
  if (audience === 'nachruecker') return isDe ? 'Neu Angemeldete (Abgleich mit alter Mail)' : 'New registrations (matched against old mail)';
  if (audience === 'reminder') return isDe ? 'Erinnerung an deine Einladungsliste' : 'Reminder to your invitation list';
  const teile: string[] = [];
  if (audience === 'custom') {
    const rest = new Set(statuses);
    for (const g of GRUPPEN) {
      if (g.stati.every(s => rest.has(s))) { teile.push(isDe ? g.de : g.en); g.stati.forEach(s => rest.delete(s)); }
    }
    rest.forEach(s => teile.push(s));
  } else {
    teile.push(audience === 'waitOnly' ? (isDe ? 'Warteliste' : 'Waitlist')
      : audience === 'activePlusWait' ? (isDe ? 'Aktive + Warteliste' : 'Active + waitlist')
      : audience === 'everyone' ? (isDe ? 'Alle — auch Abgemeldete' : 'Everyone — incl. cancellations')
      : (isDe ? 'Aktive Teilnehmer' : 'Active participants'));
  }
  if (extras.has('offene')) teile.push(isDe ? 'Noch nicht geantwortet' : 'Not responded yet');
  if (extras.has('liste')) teile.push(extras.has('listeOffen') ? (isDe ? 'Eigene Liste (ohne Reaktion)' : 'Own list (no response)') : (isDe ? 'Eigene Liste' : 'Own list'));
  if (extras.has('ich')) teile.push(isDe ? 'Ich' : 'Me');
  if (extras.has('orgs')) teile.push('Organizer');
  if (extras.has('test')) teile.push(isDe ? 'Test-Team' : 'Test team');
  return teile.length ? teile.join(', ') : (isDe ? 'nichts gewählt' : 'nothing selected');
}

export interface MassmailZielChipsProps {
  isDe: boolean;
  ev: DeloitteEvent;
  myEmail: string;
  registrations: SPRegistration[];
  audience: MassmailAudience;
  setAudience: (a: MassmailAudience) => void;
  statuses: Set<string>;
  setStatuses: React.Dispatch<React.SetStateAction<Set<string>>>;
  extras: Set<MassmailExtra>;
  setExtras: React.Dispatch<React.SetStateAction<Set<MassmailExtra>>>;
  offene: AudiencePerson[] | null | undefined;
  disabled?: boolean;
  /** v32.36: eingefügte Adressliste (Massen-Upload); ohne Setter keine Zeile „Eigene Liste". */
  pasteRaw?: string;
  setPasteRaw?: (v: string) => void;
}

export const MassmailZielChips: React.FC<MassmailZielChipsProps> = (p) => {
  const { isDe, ev, myEmail, registrations, audience, setAudience, statuses, setStatuses, extras, setExtras, offene, disabled, pasteRaw, setPasteRaw } = p;
  const [einzeln, setEinzeln] = React.useState(false);
  const [listeOffen, setListeOffen] = React.useState(false);
  const [entwurf, setEntwurf] = React.useState('');
  const [nurOhneReaktion, setNurOhneReaktion] = React.useState(true);
  const listeN = parsePastedRecipients(pasteRaw || '').length;
  const oeffneListe = (): void => {
    setEntwurf(pasteRaw || '');
    // Beim Reminder (Offene gewählt, aber unbekannt) ist der Abgleich die Absicht.
    setNurOhneReaktion(extras.has('listeOffen') || extras.has('offene') || !extras.has('liste'));
    setListeOffen(true);
  };
  const uebernehmeListe = (): void => {
    if (!setPasteRaw) return;
    setPasteRaw(entwurf);
    setExtras(prev => {
      const next = new Set(prev);
      if (offene === null || (Array.isArray(offene) && offene.length === 0)) next.delete('offene');
      next.add('liste');
      if (nurOhneReaktion) next.add('listeOffen'); else next.delete('listeOffen');
      return next;
    });
    if (audience === 'nachruecker' || audience === 'reminder') { setStatuses(new Set()); setAudience('custom'); }
    setListeOffen(false);
  };
  // Die Chips bedienen nur den 'custom'-Modus. Ein Alt-Modus (z. B. 'active'
  // aus einem Entwurf) wird beim ersten Klick in seine Status-Menge übersetzt,
  // damit „Warteliste dazu" nicht die Aktiven verliert.
  const stati = ((): Set<string> => {
    if (audience === 'custom') return statuses;
    if (audience === 'waitOnly') return new Set(['Warteliste']);
    if (audience === 'activePlusWait') return new Set([...MM_ACTIVE, 'Warteliste']);
    if (audience === 'everyone') return new Set([...MM_ACTIVE, 'Warteliste', 'Abgemeldet', 'No-Show']);
    if (audience === 'active') return new Set(MM_ACTIVE);
    return new Set<string>(); // Einfüge-Modi: keine Chips aktiv
  })();
  const setzeStati = (next: Set<string>): void => {
    setStatuses(next);
    if (audience !== 'custom') setAudience('custom');
  };
  const zaehle = (liste: string[]): number => registrations.filter(r => liste.indexOf(r.Status) >= 0).length;
  const toggleGruppe = (g: Gruppe): void => {
    const next = new Set(stati);
    const an = g.stati.every(s => next.has(s));
    g.stati.forEach(s => { if (an) next.delete(s); else next.add(s); });
    setzeStati(next);
  };
  const toggleStatus = (s: string): void => {
    const next = new Set(stati);
    if (next.has(s)) next.delete(s); else next.add(s);
    setzeStati(next);
  };
  const toggleExtra = (k: MassmailExtra): void => {
    setExtras(prev => {
      const next = new Set(prev);
      if (next.has(k)) next.delete(k); else next.add(k);
      return next;
    });
    // Aus einem Einfüge-Modus heraus heißt ein Chip-Klick: zurück zur Auswahl.
    if (audience === 'nachruecker' || audience === 'reminder') { setStatuses(new Set()); setAudience('custom'); }
  };
  const chip = (key: string, label: string, n: number | string, on: boolean, click: () => void, off?: boolean, title?: string): React.ReactElement => (
    <button key={key} type="button" className={cx('dex-ui-chip', on && 'is-active')} aria-pressed={on}
      disabled={disabled || off} title={title} onClick={click}>
      {on && <Check size={12} />} {label} <span style={{ opacity: 0.75, marginLeft: 2 }}>{n}</span>
    </button>
  );
  const offeneN = offene === undefined ? '…' : offene === null ? '–' : offene.length;
  const offeneTitel = offene === null
    ? (isDe ? 'Für dieses Event kennt DEX keine Eingeladenen (Sichtbarkeit nur nach Standort) — füge deine Einladungsliste unter „Eigene Liste“ ein.' : 'DEX knows no invitees for this event (location-only visibility) — paste your invitation list under “Own list”.')
    : undefined;
  const nIch = internAdressen(ev, myEmail, 'ich').length;
  const nOrgs = internAdressen(ev, myEmail, 'orgs').length;
  const nTest = internAdressen(ev, myEmail, 'test').length;
  const zeile = (titel: string, kinder: React.ReactNode): React.ReactElement => (
    <div style={{ display: 'flex', gap: 10, alignItems: 'flex-start', flexWrap: 'wrap', marginBottom: 10 }}>
      <span className="dex-ui-muted" style={{ fontSize: '0.8rem', fontWeight: 600, minWidth: 110, paddingTop: 6 }}>{titel}</span>
      <div className="dex-ui-inline" style={{ gap: 6, flexWrap: 'wrap', flex: '1 1 300px' }}>{kinder}</div>
    </div>
  );
  return (
    <div>
      {zeile(isDe ? 'Teilnehmende' : 'Participants', <>
        {GRUPPEN.map(g => chip(g.key, isDe ? g.de : g.en, zaehle(g.stati), g.stati.every(s => stati.has(s)), () => toggleGruppe(g)))}
        {chip('offene', isDe ? 'Noch nicht geantwortet' : 'Not responded yet', offeneN, extras.has('offene') && audience !== 'reminder' && audience !== 'nachruecker', () => toggleExtra('offene'), !Array.isArray(offene) || offene.length === 0, offeneTitel)}
        <button type="button" className="dex-ui-textbtn" onClick={() => setEinzeln(o => !o)} disabled={disabled}>
          {einzeln ? (isDe ? 'Einzelne Status ausblenden' : 'Hide single statuses') : (isDe ? 'Einzelne Status …' : 'Single statuses …')}
        </button>
      </>)}
      {einzeln && zeile(isDe ? 'Genauer' : 'In detail', <>
        {MM_ACTIVE.map(s => chip('st-' + s, s, zaehle([s]), stati.has(s), () => toggleStatus(s)))}
      </>)}
      {/* v32.36: Gewählt, aber nicht rechenbar — sagen, was stattdessen geht. */}
      {extras.has('offene') && offene === null && (
        <div className="dex-ui-callout dex-ui-callout--warn" style={{ margin: '0 0 10px' }}>
          <span>
            {isDe
              ? 'DEX kennt die Eingeladenen dieses Events nicht (Sichtbarkeit nur über Standorte). Füge deine Einladungsliste ein — DEX schreibt dann nur an, wer noch nicht reagiert hat.'
              : 'DEX does not know the invitees of this event (location-only visibility). Paste your invitation list — DEX then only writes to people who have not responded.'}
            {setPasteRaw && <>{' '}<button type="button" className="dex-ui-textbtn" onClick={oeffneListe}>{isDe ? 'Liste einfügen …' : 'Paste list …'}</button></>}
          </span>
        </div>
      )}
      {zeile(isDe ? 'Intern' : 'Internal', <>
        {chip('ich', isDe ? 'An mich' : 'To me', nIch, extras.has('ich') && audience !== 'reminder' && audience !== 'nachruecker', () => toggleExtra('ich'), nIch === 0)}
        {chip('orgs', isDe ? 'Organizer' : 'Organizers', nOrgs, extras.has('orgs') && audience !== 'reminder' && audience !== 'nachruecker', () => toggleExtra('orgs'), nOrgs === 0)}
        {chip('test', isDe ? 'Test-Team' : 'Test team', nTest, extras.has('test') && audience !== 'reminder' && audience !== 'nachruecker', () => toggleExtra('test'), nTest === 0,
          nTest === 0 ? (isDe ? 'Für dieses Event ist kein Test-Team eingetragen.' : 'No test team on this event.') : undefined)}
      </>)}
      {setPasteRaw && zeile(isDe ? 'Eigene Liste' : 'Own list', <>
        {listeN > 0 && chip('liste', extras.has('listeOffen') ? (isDe ? 'Eigene Liste · ohne Reaktion' : 'Own list · no response') : (isDe ? 'Eigene Liste' : 'Own list'), listeN, extras.has('liste') && audience !== 'reminder' && audience !== 'nachruecker', () => toggleExtra('liste'))}
        <button type="button" className="dex-ui-textbtn" onClick={oeffneListe} disabled={disabled}>
          {listeN > 0 ? (isDe ? 'Liste ändern …' : 'Change list …') : (isDe ? 'Adressliste einfügen …' : 'Paste address list …')}
        </button>
      </>)}
      {listeOffen && (
        <Modal open={true} onClose={() => setListeOffen(false)} maxWidth={720}
          title={isDe ? 'Eigene Liste einfügen' : 'Paste your own list'}
          subtitle={isDe ? 'Aus Outlook kopiert oder beliebig formatiert — DEX erkennt Namen und Adressen.' : 'Copied from Outlook or any format — DEX picks out names and addresses.'}
          footer={<>
            <button type="button" className="btn btn-secondary" onClick={() => setListeOffen(false)}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
            <button type="button" className="btn btn-primary" onClick={uebernehmeListe} disabled={!entwurf.trim()}>
              {isDe ? `Übernehmen (${parsePastedRecipients(entwurf).length})` : `Apply (${parsePastedRecipients(entwurf).length})`}
            </button>
          </>}>
          <textarea className="form-control" rows={10} value={entwurf} onChange={e => setEntwurf(e.target.value)}
            placeholder={isDe ? 'Max Mustermann <mmustermann@deloitte.de>; erika@deloitte.de …' : 'Jane Doe <jdoe@deloitte.com>; john@deloitte.com …'}
            style={{ width: '100%', fontFamily: 'inherit', fontSize: '0.9rem', boxSizing: 'border-box' }} />
          <label className="dex-ui-toggle-row" style={{ marginTop: 10 }}>
            <input type="checkbox" checked={nurOhneReaktion} onChange={e => setNurOhneReaktion(e.target.checked)} />
            <span className="dex-ui-toggle-row-body">
              <span className="dex-ui-toggle-row-title">{isDe ? 'Nur an Personen, die noch nicht reagiert haben' : 'Only people who have not responded yet'}</span>
              <span className="dex-ui-toggle-row-desc">{isDe ? 'Wer schon angemeldet, auf der Warteliste oder abgemeldet ist, fällt raus — typisch für den Reminder.' : 'Anyone already registered, waitlisted or cancelled is left out — typical for a reminder.'}</span>
            </span>
          </label>
        </Modal>
      )}
    </div>
  );
};

export default MassmailZielChips;
