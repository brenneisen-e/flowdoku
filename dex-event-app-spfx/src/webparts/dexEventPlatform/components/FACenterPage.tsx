/**
 * v30.5: F&A Center (Fachkonzept Abschnitte 8–11).
 *
 * Zentrale Sicht der Finance & Accounting Abteilung auf alle
 * abrechnungsrelevanten Events: Verteiler-Pflege, Dashboard-Kennzahlen,
 * durchsuchbare Tabelle und Detailansicht mit Kommunikationshistorie und
 * „Als abgerechnet markieren". Zugriff: Rolle „F&A" und Admins. Texte
 * bewusst nur deutsch — das Konzept ist deutsch, der Pilot ist intern.
 *
 * Datenhaltung: alles hängt am Piggyback `_billing` der Events (siehe
 * utils/faBilling.ts) plus der Verteiler-Zeile `_FAConfig` in
 * DEX_EmailTemplates. Die Detailansicht zeigt laut Konzept NUR den an F&A
 * ÜBERMITTELTEN Stand (Snapshots), nie den Live-Stand des Wizards.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useEvents } from '../context/EventContext';
import { useRoles } from '../context/RoleContext';
import { useCurrentUser } from '../context/UserContext';
import { useDialog } from '../context/DialogContext';
import { DeloitteEvent } from '../types';
import { BILLING_FIELDS } from '../data/billingFields';
import { deepLinkParams } from '../utils/deepLink';
import RecipientPicker from './admin/RecipientPicker';
import {
  parseBillingOf, faStatusOf, FAStatus, FA_STATUS_LABELS, FA_STATUS_COLORS,
  FA_STATUS_ORDER, FA_STATUS_SHORT,
  FAConfig, BillingLogEntry, BillingData, FAListRow, downloadFAParticipantXlsx, parsePersonValue,
  activeEmployeesLookupUrl,
} from '../utils/faBilling';
import { PersonContactHover } from './PersonContactHover';

const fmtDate = (iso: string): string => {
  const d = new Date(iso || '');
  return isFinite(d.getTime()) ? d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '—';
};
const fmtDateTime = (iso: string): string => {
  const d = new Date(iso || '');
  return isFinite(d.getTime())
    ? d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' }) + ', ' + d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })
    : '—';
};

// v30.45: `parseRecipientInput` ist entfallen. Der Verteiler wird nicht mehr
// als Freitext gepflegt, sondern als Liste (Person-Picker + Chips, siehe
// components/admin/RecipientPicker.tsx) — es gibt nichts mehr zu parsen.

/**
 * v30.50: Der Ansprechpartner eines Events als Personen-Profil.
 *
 * Fachkonzept „F&A Center — Ansprechpartner": Die Person soll „analog zu den
 * übrigen Personenanzeigen innerhalb der Plattform als Teams-/Microsoft-365-
 * Profil" erscheinen. Genau das leistet `PersonContactHover` schon überall
 * sonst (Teilnehmerliste, Organizer, Assistenz): Foto, beim Hover eine Karte
 * mit Position, Standort, klickbarer E-Mail-Adresse und Teams-Chat-Link.
 *
 * Gespeichert bleibt der Wert unverändert als `Name <email>` — dasselbe
 * Format, das `UserFieldPicker` überall schreibt. Zerlegt wird erst hier,
 * für die Anzeige. Zwei Fälle, die es wirklich gibt und die deshalb nicht
 * geraten werden: ein Alt-Wert aus der Freitext-Zeit vor v30.45 kann eine
 * nackte Adresse ODER ein reiner Name sein. Ohne Adresse gibt es kein Foto
 * und keinen Teams-Chat — dann steht der Name da, und zwar ohne einen leeren
 * Avatar-Kreis daneben, der eine Verknüpfung verspricht, die es nicht gibt.
 */
function FAContactCell(props: { value: string }): React.ReactElement {
  const { name, email } = parsePersonValue(props.value);
  if (!name && !email) return <span style={{ color: 'var(--dex-gray-400)' }}>—</span>;
  const label = name || email;
  if (!email) return <span>{label}</span>;
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8, minWidth: 0 }}>
      <PersonContactHover email={email} name={label} size={26} isDe />
      <span style={{ minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{label}</span>
    </span>
  );
}

const statusPill = (s: FAStatus): React.ReactElement => (
  <span style={{
    display: 'inline-block', padding: '3px 10px', borderRadius: 999,
    fontSize: '0.72rem', fontWeight: 700,
    background: FA_STATUS_COLORS[s].bg, color: FA_STATUS_COLORS[s].fg,
  }}>{FA_STATUS_LABELS[s]}</span>
);

export default function FACenterPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { isFA, isAdmin, searchUsers, searchUser, getEmployeeData } = useRoles();
  const { events, getFAConfig, saveFAConfig, markEventSettled, saveFAPersonalNumbers } = useEvents();
  const { confirmDialog, showAlert } = useDialog();
  const { currentUser } = useCurrentUser();

  const allowed = isFA || isAdmin;

  const [cfg, setCfg] = React.useState<FAConfig | null>(null);
  // v30.45: Empfaenger als LISTE statt als Textarea-Text. Gespeichert wird
  // weiterhin `string[]` mit nackten Adressen — nur die Bedienung wechselt vom
  // Freitext auf Person-Picker + Chips (F&A-Fachkonzept).
  const [infoAddrs, setInfoAddrs] = React.useState<string[]>([]);
  const [listAddrs, setListAddrs] = React.useState<string[]>([]);
  const [cfgBusy, setCfgBusy] = React.useState(false);
  const [cfgLogOpen, setCfgLogOpen] = React.useState(false);
  const [search, setSearch] = React.useState('');
  const [statusFilter, setStatusFilter] = React.useState<'' | FAStatus>('');
  const [selectedId, setSelectedId] = React.useState<string | null>(null);
  const [settleBusy, setSettleBusy] = React.useState(false);
  const [openMailIdx, setOpenMailIdx] = React.useState<number | null>(null);
  const [xlsxBusy, setXlsxBusy] = React.useState(false);
  // v30.60: Personalnummern-Mapping. Der Entwurf liegt lokal je E-Mail-Adresse
  // und wird erst auf Klick gespeichert — Zeichen für Zeichen zu speichern
  // hieße bei 100 Teilnehmern 100 Schreibvorgänge auf dasselbe Event-Item.
  const [pnDraft, setPnDraft] = React.useState<Record<string, { personalNr?: string; costCenter?: string }>>({});
  const [pnBusy, setPnBusy] = React.useState(false);
  const [pnEventId, setPnEventId] = React.useState<string>('');
  // v30.61: Ergebnis des automatischen Abrufs — getrennt gehalten, damit die
  // Meldung „x von y gefunden" ehrlich bleibt (leer heißt nicht „nichts da",
  // sondern kann auch „nicht erlaubt" heißen; s. SharePointService).
  const [pnAutoBusy, setPnAutoBusy] = React.useState(false);
  const [pnAutoNote, setPnAutoNote] = React.useState('');

  React.useEffect(() => {
    if (!allowed) return;
    let cancelled = false;
    getFAConfig().then(c => {
      if (cancelled) return;
      setCfg(c);
      setInfoAddrs(c.infoRecipients);
      setListAddrs(c.listRecipients);
    }).catch(() => { /* leerer Stand bleibt */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allowed]);

  // Nur Hauptevents tragen eine Abrechnung; Sub-Events und Demo bleiben draußen.
  const billingEvents = React.useMemo(() => {
    return events
      .filter(e => !e.parentEventId && !e.isDemoShowcase)
      .map(e => ({ ev: e, b: parseBillingOf(e) }))
      .filter(x => x.b && x.b.relevant)
      .map(x => ({ ...x, status: faStatusOf(x.ev, x.b) }))
      .sort((a, b2) => new Date(b2.ev.startDate || 0).getTime() - new Date(a.ev.startDate || 0).getTime());
  }, [events]);

  // v30.24: Deep-Link aus den F&A-Mails (#action=fa&event=<id>) — das Event
  // direkt in der Detailansicht öffnen. Läuft erst, wenn die Events geladen
  // sind (billingEvents leer = noch nichts zu wählen), und genau einmal:
  // danach soll die eigene Navigation im Center Vorrang haben.
  const deepLinkDone = React.useRef(false);
  React.useEffect(() => {
    if (deepLinkDone.current || !allowed || billingEvents.length === 0) return;
    let evId = '';
    try {
      const p = deepLinkParams();
      if (p.get('action') === 'fa') evId = (p.get('event') || '').trim();
    } catch { /* URL nicht lesbar — dann bleibt die Übersicht stehen */ }
    if (!evId) { deepLinkDone.current = true; return; }
    deepLinkDone.current = true;
    if (billingEvents.some(x => x.ev.id === evId)) setSelectedId(evId);
    else showAlert('Das verlinkte Event ist nicht (mehr) als abrechnungsrelevant hinterlegt. Bitte wende dich an die Organizer des Events.', { variant: 'info' });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allowed, billingEvents.length]);

  /**
   * v30.24: Den an F&A übermittelten Stand als Excel herunterladen.
   *
   * Datenquelle ist bewusst der SNAPSHOT aus `_billing.listSnapshot` und
   * nicht die Live-Teilnehmerliste: (1) Er liegt in DEX_Events und ist damit
   * auch für F&A lesbar — die Teilnehmerlisten selbst liegen auf Subsites
   * mit eigenen Berechtigungen, auf die F&A keinen Zugriff hat. (2) Er ist
   * revisionssicher genau das, was gemeldet wurde. Ersetzt den im Tenant
   * unmöglichen Datei-Anhang (s. utils/faBilling).
   */
  const downloadSnapshotXlsx = async (
    ev: DeloitteEvent,
    b: BillingData | null | undefined,
    snapshot: FAListRow[],
    sentAt?: string,
  ): Promise<void> => {
    if (xlsxBusy) return;
    setXlsxBusy(true);
    try {
      // v30.50: Aufbau der Datei liegt in utils/faBilling — dieselbe Datei,
      // die der Abrechnungs-Dialog im Organizer Center erzeugt.
      await downloadFAParticipantXlsx(ev, b, snapshot, sentAt);
    } catch (err) {
      console.warn('[DEX] F&A-Snapshot-Export fehlgeschlagen:', err);
      showAlert('Die Excel-Datei konnte nicht erzeugt werden — bitte erneut versuchen.', { variant: 'error' });
    } finally { setXlsxBusy(false); }
  };

  const counts = React.useMemo(() => {
    // v30.47: aus FA_STATUS_ORDER aufgebaut statt von Hand aufgezaehlt — eine
    // neue Stufe im Statusmodell taucht damit automatisch hier auf.
    const c = {} as Record<FAStatus, number>;
    FA_STATUS_ORDER.forEach(k => { c[k] = 0; });
    billingEvents.forEach(x => { c[x.status]++; });
    return c;
  }, [billingEvents]);

  const filtered = React.useMemo(() => {
    const q = search.trim().toLowerCase();
    return billingEvents.filter(x => {
      if (statusFilter && x.status !== statusFilter) return false;
      if (!q) return true;
      const hay = [
        String(x.ev.eventNumber || ''), x.ev.title,
        (x.b?.fields || {}).contact || '', (x.b?.fields || {}).ariba || '',
        fmtDate(x.ev.startDate), FA_STATUS_LABELS[x.status],
      ].join(' ').toLowerCase();
      return hay.indexOf(q) >= 0;
    });
  }, [billingEvents, search, statusFilter]);

  // v30.60: Der Personalnummern-Entwurf gehört zu EINEM Event. Wer die
  // Detailansicht wechselt, ohne zu speichern, würde die Eingaben sonst beim
  // nächsten Speichern dem falschen Event zuordnen. Deshalb hier verworfen —
  // sichtbar, weil die Zellen dann wieder den gespeicherten Stand zeigen.
  // Steht VOR dem `if (!allowed)`: Hooks hinter einem frühen Return reißen die
  // Hook-Reihenfolge (siehe CLAUDE.md, React #300).
  React.useEffect(() => {
    const id = selectedId || '';
    if (id === pnEventId) return;
    setPnEventId(id);
    setPnDraft({});
  }, [selectedId, pnEventId]);

  const selected = selectedId ? billingEvents.find(x => x.ev.id === selectedId) : undefined;

  if (!allowed) {
    return (
      <div className="page-container text-center">
        <div className="card" style={{ padding: '48px 32px', maxWidth: 560, margin: '0 auto' }}>
          <h2 style={{ marginTop: 0 }}>Kein Zugriff</h2>
          <p style={{ color: 'var(--dex-gray-600)' }}>
            Das F&amp;A Center steht Personen mit der Rolle &bdquo;F&amp;A&ldquo; und Admins zur Verfügung.
          </p>
          <button className="btn btn-primary mt-24" onClick={() => navigate('start')}>Zurück</button>
        </div>
      </div>
    );
  }

  const saveRecipients = async (): Promise<void> => {
    if (!cfg || cfgBusy) return;
    setCfgBusy(true);
    try {
      const nextInfo = infoAddrs.filter(Boolean);
      const nextList = listAddrs.filter(Boolean);
      const by = `${currentUser?.firstName || ''} ${currentUser?.surname || ''}`.trim() || currentUser?.email || '';
      const log = [...cfg.log];
      // Änderungen protokollieren (Konzept 8.1: „Änderungen werden protokolliert").
      if (nextInfo.join(';') !== cfg.infoRecipients.join(';')) {
        log.push({ ts: new Date().toISOString(), by, action: 'Verteiler „Abrechnungsinformationen" geändert', old: cfg.infoRecipients.join('; '), neu: nextInfo.join('; ') });
      }
      if (nextList.join(';') !== cfg.listRecipients.join(';')) {
        log.push({ ts: new Date().toISOString(), by, action: 'Verteiler „Teilnehmerlisten" geändert', old: cfg.listRecipients.join('; '), neu: nextList.join('; ') });
      }
      const next: FAConfig = { infoRecipients: nextInfo, listRecipients: nextList, log };
      const ok = await saveFAConfig(next);
      if (ok) {
        setCfg(next);
        showAlert('Verteiler gespeichert — sie gelten ab sofort für alle zukünftigen Versendungen.', { variant: 'success' });
      } else {
        showAlert('Speichern fehlgeschlagen — bitte später erneut versuchen.', { variant: 'error' });
      }
    } finally { setCfgBusy(false); }
  };

  const doSettle = async (ev: DeloitteEvent): Promise<void> => {
    if (settleBusy) return;
    const ok = await confirmDialog(
      `„${ev.title}" als abgerechnet markieren? Der Status bleibt dauerhaft bestehen; Datum und Person werden protokolliert.`,
      { confirmLabel: 'Als abgerechnet markieren' },
    );
    if (!ok) return;
    setSettleBusy(true);
    try {
      const done = await markEventSettled(ev);
      if (!done) showAlert('Markieren fehlgeschlagen — bitte später erneut versuchen.', { variant: 'error' });
    } finally { setSettleBusy(false); }
  };

  // ---------- Detailansicht (Konzept Abschnitt 10) ----------
  if (selected) {
    const { ev, b, status } = selected;
    const mails = (b?.log || []).filter(l => l.mailType);
    return (
      <div className="page-container">
        <button className="btn btn-secondary" style={{ marginBottom: 16 }} onClick={() => { setSelectedId(null); setOpenMailIdx(null); }}>
          ← Zurück zur Übersicht
        </button>
        <div className="card" style={{ padding: '24px 28px', marginBottom: 16 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 16, flexWrap: 'wrap' }}>
            <div>
              <h2 style={{ margin: '0 0 6px' }}>{ev.title}</h2>
              <div style={{ marginBottom: 10 }}>{statusPill(status)}</div>
            </div>
            {status !== 'settled' && (
              <button className="btn btn-primary" disabled={settleBusy} onClick={() => { void doSettle(ev); }}>
                {settleBusy ? 'Wird markiert…' : 'Als abgerechnet markieren'}
              </button>
            )}
          </div>
          {b?.settled && (
            <p style={{ margin: '0 0 12px', fontSize: '0.85rem', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600 }}>
              Abgerechnet am {fmtDateTime(b.settled.ts)} durch {b.settled.by}.
            </p>
          )}
          {/* Bereich 1: Eventinformationen */}
          <h3 style={{ margin: '10px 0 8px', fontSize: '0.95rem' }}>Eventinformationen</h3>
          <table style={{ borderCollapse: 'collapse', fontSize: '0.85rem' }}>
            <tbody>
              {[
                ['Event-ID', String(ev.eventNumber || ev.id)],
                ['Eventdatum', `${fmtDateTime(ev.startDate)}${ev.endDate && ev.endDate !== ev.startDate ? ' – ' + fmtDateTime(ev.endDate) : ''}`],
                ['Ort', ev.location || '—'],
                ['Organizer', (ev.organizers || []).join('; ') || '—'],
                ['Versandart', (b?.sendMode === 'auto') ? 'Automatisiert (7 Tage vor/nach dem Event)' : 'Manuell über die Organizer-Ansicht'],
              ].map(([k, v]) => (
                <tr key={k}>
                  <td style={{ padding: '4px 16px 4px 0', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap', verticalAlign: 'top' }}>{k}</td>
                  <td style={{ padding: '4px 0' }}>{v}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Bereich 2: An F&A übermittelte Abrechnungsinformationen */}
        <div className="card" style={{ padding: '20px 28px', marginBottom: 16 }}>
          <h3 style={{ margin: '0 0 10px', fontSize: '0.95rem' }}>Abrechnungsrelevante Informationen</h3>
          {b?.infoSnapshot ? (
            <>
              <p style={{ margin: '0 0 10px', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                Stand des letzten Versands an F&amp;A ({b.infoSentAt ? fmtDateTime(b.infoSentAt) : '—'}).
              </p>
              <table style={{ borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                <tbody>
                  {BILLING_FIELDS.map(f => (
                    <tr key={f.id}>
                      <td style={{ padding: '4px 16px 4px 0', color: 'var(--dex-gray-500)', verticalAlign: 'top' }}>{f.label}</td>
                      <td style={{ padding: '4px 0' }}>
                        {/* v30.50: Die Kontaktperson auch hier als Profil —
                            sonst steht dieselbe Person in der Tabelle mit Foto
                            und in der Detailansicht als Zeichenkette. */}
                        {f.type === 'user'
                          ? <FAContactCell value={(b.infoSnapshot || {})[f.id] || ''} />
                          : ((b.infoSnapshot || {})[f.id] || '—')}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </>
          ) : (
            <p style={{ margin: 0, color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>
              Es wurden bisher keine Abrechnungsinformationen an F&amp;A versendet.
            </p>
          )}
        </div>

        {/* Bereich 3: An F&A übermittelte Teilnehmerliste */}
        <div className="card" style={{ padding: '20px 28px', marginBottom: 16 }}>
          <h3 style={{ margin: '0 0 10px', fontSize: '0.95rem' }}>Teilnehmerliste</h3>
          {b?.listSnapshot && b.listSnapshot.length > 0 ? (
            <>
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap', margin: '0 0 10px' }}>
                <p style={{ margin: 0, fontSize: '0.78rem', color: 'var(--dex-gray-500)', flex: '1 1 240px' }}>
                  {b.listSnapshot.length} Personen — Stand des letzten Versands an F&amp;A ({b.listSentAt ? fmtDateTime(b.listSentAt) : '—'}).
                </p>
                {/* v30.24: Download statt Mail-Anhang — s. downloadSnapshotXlsx. */}
                <button
                  type="button"
                  className="btn btn-primary"
                  style={{ fontSize: '0.8rem', padding: '6px 16px', flexShrink: 0 }}
                  disabled={xlsxBusy}
                  onClick={() => { void downloadSnapshotXlsx(selected.ev, b, b.listSnapshot || [], b.listSentAt); }}
                >
                  {xlsxBusy ? 'Wird erzeugt…' : 'Als Excel herunterladen'}
                </button>
              </div>
              {/* v30.50: Ehrlicher Hinweis statt stiller Lücke. Snapshots aus
                  Versendungen VOR diesem Release tragen nur Name, E-Mail und
                  Status — First/Last Name, Country und Company Name gab es im
                  Datensatz noch nicht und lassen sich nicht rückwirkend
                  erfinden. Die Datei hat dann die richtige Form, aber leere
                  Spalten; wer den vollständigen Satz braucht, versendet die
                  Liste einmal neu. */}
              {!b.listSnapshot.some(p => p.firstName || p.lastName || p.company) && (
                <p style={{
                  margin: '0 0 10px', padding: '8px 12px', borderRadius: 8,
                  background: 'rgba(237,139,0,0.09)', color: 'var(--dex-orange-dark, #b35a00)',
                  fontSize: '0.78rem', lineHeight: 1.5,
                }}>
                  Dieser Stand wurde vor der Formatumstellung übermittelt und trägt nur Name, E-Mail und Status.
                  In der Excel-Datei bleiben <strong>First Name</strong>, <strong>Last Name</strong>,
                  {' '}<strong>Country</strong> und <strong>Company Name</strong> deshalb leer.
                  Für den vollständigen Satz die Teilnehmerliste einmal neu versenden.
                </p>
              )}
              {/* v30.60: Personalnummern-Mapping. F&A hat als einzige Rolle
                  Zugriff auf die Backoffice-Liste „Active Employees"; der
                  Knopf je Zeile öffnet sie in einem neuen Tab, bereits nach
                  dem Nachnamen gesucht (Parameter `k`). Die gefundene Nummer
                  wird hier eingetragen und landet in der Excel-Datei — die
                  beiden Spalten der F&A-Vorlage blieben bisher leer. */}
              <div style={{
                display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap',
                margin: '0 0 10px', padding: '8px 12px', borderRadius: 8,
                background: 'rgba(0,118,168,0.06)', fontSize: '0.78rem', lineHeight: 1.5,
              }}>
                <span style={{ flex: '1 1 260px', color: 'var(--dex-gray-600)' }}>
                  <strong>Personalnummern zuordnen:</strong> &bdquo;Automatisch füllen&ldquo; holt
                  Personalnummer und Kostenstelle für alle Personen auf einmal aus dem
                  Verzeichnis. Was dort nicht gepflegt ist, trägst du selbst ein —
                  &bdquo;Nachschlagen&ldquo; öffnet dafür die Liste <em>Active Employees</em>,
                  bereits nach dem Nachnamen gesucht.
                </span>
                {/* v30.61: Der eigentliche Weg. Die Handarbeit aus v30.60 war bei
                    100 Teilnehmern eine Stunde Abtippen — an genau der Stelle, an
                    der ein Zahlendreher auf der Rechnung landet. */}
                <button
                  type="button"
                  className="btn btn-secondary"
                  style={{ fontSize: '0.8rem', padding: '6px 16px', flexShrink: 0 }}
                  disabled={pnAutoBusy || pnBusy}
                  onClick={() => {
                    const rows = b.listSnapshot || [];
                    setPnAutoBusy(true);
                    setPnAutoNote('');
                    getEmployeeData(rows.map(r => r.email || ''))
                      .then(map => {
                        const next: Record<string, { personalNr?: string; costCenter?: string }> = { ...pnDraft };
                        let found = 0;
                        for (const r of rows) {
                          const k = (r.email || '').toLowerCase().trim();
                          const d = map[k];
                          if (!d) continue;
                          const pn = (d.employeeId || '').trim();
                          const cc = (d.costCenter || '').trim();
                          if (!pn && !cc) continue;
                          found++;
                          // Bereits Eingetragenes NICHT überschreiben: Wer von Hand
                          // korrigiert hat, hat den besseren Wert.
                          const cur = next[k] || {};
                          next[k] = {
                            personalNr: (cur.personalNr !== undefined ? cur.personalNr : r.personalNr) || pn,
                            costCenter: (cur.costCenter !== undefined ? cur.costCenter : r.costCenter) || cc,
                          };
                        }
                        setPnDraft(next);
                        setPnAutoNote(found === 0
                          ? 'Aus dem Verzeichnis kam nichts zurück. Entweder sind die Felder im Tenant nicht gepflegt, oder die Berechtigung „User.Read.All" ist im SharePoint Admin Center noch nicht freigegeben. Der manuelle Weg über „Nachschlagen" funktioniert unabhängig davon.'
                          : `${found} von ${rows.length} Personen aus dem Verzeichnis übernommen. Nicht vergessen: unten speichern.`);
                      })
                      .catch(() => setPnAutoNote('Der Abruf aus dem Verzeichnis ist fehlgeschlagen — bitte erneut versuchen oder die Nummern von Hand eintragen.'))
                      .finally(() => setPnAutoBusy(false));
                  }}
                >
                  {pnAutoBusy ? 'Wird geholt…' : 'Automatisch füllen'}
                </button>
                <button
                  type="button"
                  className="btn btn-primary"
                  style={{ fontSize: '0.8rem', padding: '6px 16px', flexShrink: 0 }}
                  disabled={pnBusy || Object.keys(pnDraft).length === 0}
                  onClick={() => {
                    setPnBusy(true);
                    saveFAPersonalNumbers(selected.ev, pnDraft)
                      .then(ok => {
                        if (ok) { setPnDraft({}); showAlert('Die Personalnummern wurden gespeichert.', { variant: 'success' }); }
                        else showAlert('Die Personalnummern konnten nicht gespeichert werden — bitte erneut versuchen.', { variant: 'error' });
                      })
                      .catch(() => showAlert('Die Personalnummern konnten nicht gespeichert werden — bitte erneut versuchen.', { variant: 'error' }))
                      .finally(() => setPnBusy(false));
                  }}
                >
                  {pnBusy ? 'Wird gespeichert…' : `Personalnummern speichern${Object.keys(pnDraft).length > 0 ? ` (${Object.keys(pnDraft).length})` : ''}`}
                </button>
                {pnAutoNote && (
                  <div style={{ flexBasis: '100%', marginTop: 4, color: 'var(--dex-gray-700)' }}>{pnAutoNote}</div>
                )}
              </div>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ borderCollapse: 'collapse', fontSize: '0.85rem', minWidth: 420 }}>
                  <thead>
                    <tr>
                      {['#', 'Name', 'E-Mail', 'Status', 'Personalnummer', 'Kostenstelle', ''].map((h, hi) => (
                        <th key={hi} style={{ textAlign: 'left', padding: '6px 12px 6px 0', borderBottom: '2px solid var(--dex-gray-200)' }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {b.listSnapshot.map((p, i) => {
                      const k = (p.email || '').toLowerCase().trim();
                      const d = pnDraft[k];
                      const pnVal = d && d.personalNr !== undefined ? d.personalNr : (p.personalNr || '');
                      const ccVal = d && d.costCenter !== undefined ? d.costCenter : (p.costCenter || '');
                      const patch = (part: { personalNr?: string; costCenter?: string }): void => setPnDraft(prev => ({
                        ...prev,
                        [k]: { personalNr: pnVal, costCenter: ccVal, ...part },
                      }));
                      const cellInput: React.CSSProperties = {
                        width: 130, padding: '3px 8px', fontSize: '0.82rem',
                        border: '1px solid var(--dex-gray-200)', borderRadius: 6,
                      };
                      return (
                        <tr key={i}>
                          <td style={{ padding: '4px 12px 4px 0', color: 'var(--dex-gray-400)' }}>{i + 1}</td>
                          <td style={{ padding: '4px 12px 4px 0' }}>{p.name}</td>
                          <td style={{ padding: '4px 12px 4px 0' }}>{p.email}</td>
                          <td style={{ padding: '4px 12px 4px 0' }}>{p.status}</td>
                          <td style={{ padding: '4px 12px 4px 0' }}>
                            <input
                              type="text" value={pnVal} style={cellInput}
                              placeholder="—"
                              onChange={e => patch({ personalNr: e.target.value })}
                            />
                          </td>
                          <td style={{ padding: '4px 12px 4px 0' }}>
                            <input
                              type="text" value={ccVal} style={cellInput}
                              placeholder="—"
                              onChange={e => patch({ costCenter: e.target.value })}
                            />
                          </td>
                          <td style={{ padding: '4px 0' }}>
                            {/* Neuer Tab + noopener: Die Backoffice-Seite liegt in
                                einer anderen Site Collection; im selben Tab wäre
                                das F&A Center samt Entwurf weg. Die Anmeldung
                                trägt der Browser mit — es ist derselbe Nutzer. */}
                            <a
                              href={activeEmployeesLookupUrl(p)}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="btn btn-secondary"
                              style={{ fontSize: '0.75rem', padding: '3px 10px', whiteSpace: 'nowrap', textDecoration: 'none' }}
                              title={`Active Employees nach „${p.lastName || p.name}“ durchsuchen`}
                            >
                              Nachschlagen
                            </a>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </>
          ) : (
            <p style={{ margin: 0, color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>
              Es wurde bisher keine Teilnehmerliste an F&amp;A versendet.
            </p>
          )}
        </div>

        {/* Bereich 4: Kommunikationshistorie */}
        <div className="card" style={{ padding: '20px 28px', marginBottom: 16 }}>
          <h3 style={{ margin: '0 0 10px', fontSize: '0.95rem' }}>Kommunikationshistorie</h3>
          {mails.length === 0 ? (
            <p style={{ margin: 0, color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>
              Über die Plattform wurden noch keine abrechnungsrelevanten E-Mails versendet.
            </p>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {mails.map((m: BillingLogEntry, i: number) => (
                <div key={i} style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: '10px 12px', fontSize: '0.82rem' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: 10, flexWrap: 'wrap' }}>
                    <strong>{m.mailType === 'info' ? 'Abrechnungsinformationen' : 'Teilnehmerliste'}</strong>
                    <span style={{ color: 'var(--dex-gray-500)' }}>{fmtDateTime(m.ts)}</span>
                  </div>
                  <div style={{ color: 'var(--dex-gray-600)', marginTop: 4, lineHeight: 1.5 }}>
                    Absender: {m.by} · Empfänger: {m.to || '—'}{m.cc ? <> · CC: {m.cc}</> : null}<br />
                    Betreff: {m.subject || '—'}
                  </div>
                  {m.body ? (
                    <button
                      type="button"
                      className="btn btn-secondary"
                      style={{ marginTop: 8, fontSize: '0.75rem', padding: '4px 12px' }}
                      onClick={() => setOpenMailIdx(openMailIdx === i ? null : i)}
                    >
                      {openMailIdx === i ? 'Inhalt ausblenden' : 'Vollständigen Inhalt anzeigen'}
                    </button>
                  ) : (
                    <div style={{ marginTop: 6, color: 'var(--dex-gray-400)', fontSize: '0.75rem' }}>
                      Inhalt aus Platzgründen nicht mehr gespeichert (nur die letzten 15 Mails behalten ihren Volltext).
                    </div>
                  )}
                  {openMailIdx === i && m.body && (
                    <div
                      style={{ marginTop: 10, border: '1px dashed var(--dex-gray-300)', borderRadius: 8, padding: 12, background: '#fff', overflowX: 'auto' }}
                      dangerouslySetInnerHTML={{ __html: m.body }}
                    />
                  )}
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    );
  }

  // ---------- Übersicht ----------
  return (
    <div className="page-container">
      <h2 style={{ marginBottom: 4 }}>F&amp;A Center</h2>
      <p style={{ color: 'var(--dex-gray-600)', marginTop: 0, marginBottom: 20, fontSize: '0.9rem' }}>
        Zentrale Verwaltung und Nachverfolgung aller abrechnungsrelevanten Veranstaltungen.
      </p>

      {/* 8.1 Verteiler */}
      <div className="card" style={{ padding: '20px 24px', marginBottom: 20 }}>
        <h3 style={{ margin: '0 0 6px', fontSize: '0.95rem' }}>Empfängeradressen</h3>
        <p style={{ margin: '0 0 12px', fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
          An diese Adressen gehen die Versendungen — getrennt für Abrechnungsinformationen und Teilnehmerlisten
          (eine Adresse pro Zeile). Änderungen wirken sofort auf alle zukünftigen Versendungen und werden protokolliert.
        </p>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: 16 }}>
          <RecipientPicker
            label="Verteiler Abrechnungsinformationen"
            hint="Personen über die Suche, Funktionspostfächer über das Feld darunter."
            value={infoAddrs}
            onChange={setInfoAddrs}
            searchUsers={searchUsers}
            searchUserByEmail={searchUser}
            disabled={cfgBusy}
          />
          <RecipientPicker
            label="Verteiler Teilnehmerlisten"
            hint="Personen über die Suche, Funktionspostfächer über das Feld darunter."
            value={listAddrs}
            onChange={setListAddrs}
            searchUsers={searchUsers}
            searchUserByEmail={searchUser}
            disabled={cfgBusy}
          />
        </div>
        <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginTop: 12, flexWrap: 'wrap' }}>
          <button className="btn btn-primary" disabled={cfgBusy || !cfg} onClick={() => { void saveRecipients(); }}>
            {cfgBusy ? 'Wird gespeichert…' : 'Verteiler speichern'}
          </button>
          {cfg && cfg.log.length > 0 && (
            <button className="btn btn-secondary" style={{ fontSize: '0.8rem' }} onClick={() => setCfgLogOpen(!cfgLogOpen)}>
              {cfgLogOpen ? 'Protokoll ausblenden' : `Änderungsprotokoll (${cfg.log.length})`}
            </button>
          )}
        </div>
        {cfgLogOpen && cfg && (
          <div style={{ marginTop: 12, fontSize: '0.78rem', color: 'var(--dex-gray-600)', display: 'flex', flexDirection: 'column', gap: 6 }}>
            {[...cfg.log].reverse().map((l, i) => (
              <div key={i} style={{ borderLeft: '3px solid var(--dex-gray-200)', paddingLeft: 10 }}>
                <strong>{fmtDateTime(l.ts)}</strong> · {l.by} · {l.action}
                {l.old !== undefined && <><br />Alt: {l.old || '—'} → Neu: {l.neu || '—'}</>}
              </div>
            ))}
          </div>
        )}
      </div>

      {/* 8.2 Dashboard */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 12, marginBottom: 20 }}>
        {/* v30.47: Kacheln aus FA_STATUS_ORDER statt aus einer handgepflegten
            Liste. Die alte Liste kannte vier der fuenf Stufen — Events im
            Zustand „Teilnehmerliste versendet, Abschluss offen" waren ueber die
            Kacheln gar nicht auffindbar. Eine Aufzaehlung, die eine Stufe
            vergisst, versteckt genau die Events, die noch jemanden brauchen. */}
        {FA_STATUS_ORDER.map(key => FA_STATUS_SHORT[key]).map((label, i) => {
          const key = FA_STATUS_ORDER[i];
          return (
          <button
            key={key}
            type="button"
            className="card"
            onClick={() => setStatusFilter(statusFilter === key ? '' : key)}
            style={{
              padding: '16px 18px', textAlign: 'left', cursor: 'pointer',
              border: statusFilter === key ? `2px solid ${FA_STATUS_COLORS[key].fg}` : '1px solid var(--dex-gray-200)',
            }}
          >
            <div style={{ fontSize: '1.8rem', fontWeight: 800, color: FA_STATUS_COLORS[key].fg }}>{counts[key]}</div>
            <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', lineHeight: 1.4 }}>{label}</div>
          </button>
          );
        })}
      </div>

      {/* 9. Tabelle */}
      <div className="card" style={{ padding: '20px 24px' }}>
        <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginBottom: 14 }}>
          <input
            className="form-input"
            style={{ flex: '1 1 260px' }}
            placeholder="Suchen: Event-ID, Name, Ansprechpartner, Ariba-Nr., Datum …"
            value={search}
            onChange={e => setSearch(e.target.value)}
          />
          <select className="form-input" style={{ width: 'auto' }} value={statusFilter} onChange={e => setStatusFilter(e.target.value as '' | FAStatus)}>
            <option value="">Alle Status</option>
            {(Object.keys(FA_STATUS_LABELS) as FAStatus[]).map(s => <option key={s} value={s}>{FA_STATUS_LABELS[s]}</option>)}
          </select>
        </div>
        {filtered.length === 0 ? (
          <p style={{ margin: 0, color: 'var(--dex-gray-500)', fontSize: '0.85rem' }}>
            {billingEvents.length === 0
              ? 'Es gibt noch keine abrechnungsrelevanten Events.'
              : 'Keine Treffer für die aktuelle Suche/Filterung.'}
          </p>
        ) : (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: '0.85rem' }}>
              <thead>
                <tr>
                  {['Event-ID', 'Eventname', 'Eventdatum', 'Ansprechpartner', 'Status'].map(h => (
                    <th key={h} style={{ textAlign: 'left', padding: '8px 14px 8px 0', borderBottom: '2px solid var(--dex-gray-200)', whiteSpace: 'nowrap' }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.map(x => (
                  <tr
                    key={x.ev.id}
                    onClick={() => setSelectedId(x.ev.id)}
                    style={{ cursor: 'pointer' }}
                    onMouseEnter={e => { (e.currentTarget as HTMLTableRowElement).style.background = 'var(--dex-gray-50, #fafafa)'; }}
                    onMouseLeave={e => { (e.currentTarget as HTMLTableRowElement).style.background = ''; }}
                  >
                    <td style={{ padding: '8px 14px 8px 0', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap' }}>{x.ev.eventNumber || x.ev.id}</td>
                    <td style={{ padding: '8px 14px 8px 0', fontWeight: 600 }}>{x.ev.title}</td>
                    <td style={{ padding: '8px 14px 8px 0', whiteSpace: 'nowrap' }}>{fmtDate(x.ev.startDate)}</td>
                    <td style={{ padding: '8px 14px 8px 0' }}>
                      {/* v30.50: Ansprechpartner als Personen-Profil (Foto +
                          Hover-Karte mit E-Mail und Teams-Chat) statt als
                          roher Text `Felten, Nils Kilian <nifelten@…>`.
                          Fachkonzept: „analog zu den übrigen Personenanzeigen
                          innerhalb der Plattform". Bis dahin war das die
                          letzte Stelle im F&A Center, an der eine Person als
                          Zeichenkette stand — und die Adresse in spitzen
                          Klammern gehört ins Datenformat, nicht auf den
                          Bildschirm. */}
                      <FAContactCell value={(x.b?.fields || {}).contact || (x.ev.organizers || [])[0] || ''} />
                    </td>
                    <td style={{ padding: '8px 0' }}>{statusPill(x.status)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
