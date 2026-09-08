/* KpiTiles — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 9403-9512 des
 * Stands vor dem Schnitt). Der Rechenweg ist unverändert; die Darstellung
 * steht seit v31.3 auf den Kennzahl-Klassen aus dexUi.ts (ui-leitfaden 5a).
 * Die Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { SPRegistration } from '../../../services/EventService';
import { DeloitteEvent } from '../../../types';
import { countConsolidatedActive } from '../logic/parentRegs';
import { cx } from '../../dexUi';
import { AlertCircle } from '../../Icons';

export interface KpiTilesProps {
  isConsolidatedMode: boolean;
  isDe: boolean;
  isSplitCapacity: boolean;
  registrations: SPRegistration[];
  /** v30.67 (Review): Teilnehmerliste nicht lesbar (`regLoadError`) — Kacheln zeigen „—". */
  regsUnknown: boolean;
  selectedEvent: DeloitteEvent;
  subEventRegsByEventId: Record<string, SPRegistration[]>;
  /** v30.67 (Review): mindestens eine Termin-Liste nicht lesbar — Summen sind Untergrenzen („≥ N"). */
  subListsIncomplete: boolean;
  t: (key: string) => string;
}

export const KpiTiles: React.FC<KpiTilesProps> = (p) => {
  const { isConsolidatedMode, isDe, isSplitCapacity, registrations, regsUnknown, selectedEvent, subEventRegsByEventId, subListsIncomplete, t } = p;
        // v30.67: Bei geteilten Kapazitäten ist `maxParticipants` per
        // Definition 0 (die Kapazität steht in durchstarter-/funstarterCapacity).
        // `> 0` sollte nur „Unbegrenzt" ausschließen, schloss aber jedes
        // Split-Event mit aus — 25 Wartende, keine Kachel.
        const hasWaitlistKPI = !!(selectedEvent?.waitlistEnabled && ((selectedEvent?.maxParticipants || 0) > 0 || isSplitCapacity));
        // v15.14: Im subEventsOnlyMode (Hauptevent ohne eigene Anmeldungen)
        // beziehen sich die Stat-Cards auf die konsolidierten Teilnehmer
        // über alle Sub-Events. Die Hauptevent-Liste selbst hat hier nur
        // Alt-Daten und würde das echte Bild verfälschen.
        const consolidatedRegs: SPRegistration[] = isConsolidatedMode
          // v22.63: Klammer-eigene Abmeldungen (Absagen auf der Klammer-Subsite)
          // in die KPI „Abgemeldet" mitzählen, damit KPI und Abmeldungs-Liste
          // übereinstimmen. Nur Abgemeldet-Zeilen, um die Aktiv-Zahlen nicht zu
          // verfälschen.
          ? ([] as SPRegistration[]).concat(
              registrations.filter(r => r.Status === 'Abgemeldet'),
              ...Object.values(subEventRegsByEventId),
            )
          : [];
        const consolidatedQRByEmail = new Set<string>();
        const consolidatedCheckedByEmail = new Set<string>();
        const consolidatedWaitlistByEmail = new Set<string>();
        const consolidatedCancelledByEmail = new Set<string>();
        const consolidatedAnyByEmail = new Set<string>();
        for (const r of consolidatedRegs) {
          // v23.3: emaillose Zeile zählt als eigener Kopf (Zeilen-Id-Fallback),
          // statt aus den KPIs zu verschwinden — sonst KPI < Tabelle.
          const key = (r.ParticipantEmail || '').toLowerCase().trim() || `__noemail#${r.Id}`;
          consolidatedAnyByEmail.add(key);
          if (r.Status === 'QR versendet') consolidatedQRByEmail.add(key);
          if (r.Status === 'Eingecheckt') consolidatedCheckedByEmail.add(key);
          if (r.Status === 'Warteliste') consolidatedWaitlistByEmail.add(key);
          if (r.Status === 'Abgemeldet') consolidatedCancelledByEmail.add(key);
        }
        const active = registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt');
        // v30.67: Über den gemeinsamen Helfer, damit die Detail-Karte
        // („Aktuell registriert") dieselbe Zahl zeigt — vorher zählte sie die
        // Matrix-Zeilen inklusive Warteliste.
        const totalActive = isConsolidatedMode ? countConsolidatedActive(subEventRegsByEventId) : active.length;
        // v30.67 (Review): „—" statt Zahl, wenn die Liste nicht lesbar war. Die
        // Kacheln stehen außerhalb des `regLoadError`-Zweigs und zeigten neben
        // „kein Zugriff" weiter „0 angemeldet" — eine Null, die keine ist
        // (CLAUDE.md: unbekannt ist keine Null). Im Klammer-Modus kommen die
        // Zahlen aus den Termin-Listen; nur „Abgemeldet" zählt die Klammer-
        // Zeilen mit und ist dann ebenfalls unbekannt. War eine Termin-Liste
        // nicht lesbar, sind die Summen nur Untergrenzen: „≥ N".
        const unknownAll = regsUnknown && !isConsolidatedMode;
        // v31.3: Der Strich sagt im Titel, was er heißt — sonst liest er sich wie 0.
        const dash = <span title={isDe ? 'Liste nicht lesbar — die Zahl ist unbekannt, nicht 0' : 'List not readable — the number is unknown, not 0'}>—</span>;
        const show = (n: number, unknown: boolean): React.ReactNode => unknown
          ? dash
          : (isConsolidatedMode && subListsIncomplete
            ? <span title={isDe ? 'Mindestens so viele — eine Termin-Liste war nicht lesbar' : 'At least this many — one date list was not readable'}>≥ {n}</span>
            : n);
        // v19.12: nach EFFEKTIVER Gruppe zählen (StarterType ODER, falls leer,
        // PreferredStarterType). Sonst fehlen angemeldete Nachrücker, deren
        // StarterType der Flow nicht gesetzt hat, in der Gruppen-Zahl — dann ist
        // „Angemeldet" gesamt ≠ Summe der beiden Gruppen (Beobachtung: 142 vs.
        // 130+11). Eine angemeldete Person mit Wunsch „Funstarter" belegt real
        // einen Funstarter-Platz und muss dort mitgezählt werden.
        const durchActive = active.filter(r => (r.StarterType || r.PreferredStarterType) === 'Durchstarter').length;
        const funActive = active.filter(r => (r.StarterType || r.PreferredStarterType) === 'Funstarter').length;
        const durchCap = selectedEvent?.durchstarterCapacity || 0;
        const funCap = selectedEvent?.funstarterCapacity || 0;
        const labelA = (selectedEvent?.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter';
        const labelB = (selectedEvent?.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter';
        const reversed = !!selectedEvent?.splitDisplayOrderReversed;
        // v31.3: Der Balken färbt sich nur, wo es etwas heißt (ab 90 % eng, darüber überbucht).
        const barMod = (n: number, cap: number): string => (n > cap ? 'dex-ui-progress-bar--red' : n / cap >= 0.9 ? 'dex-ui-progress-bar--orange' : '');
        const barW = (n: number, cap: number): string => `${Math.min(100, Math.max(0, Math.round((n / cap) * 100)))}%`;
        // Eine Gruppen-Zeile statt zweier fast gleicher Blöcke: Punkt, Name, Belegung, Auslastung.
        const grpRow = (key: string, label: string, n: number, cap: number, dot: string): React.ReactElement => (
          <div key={key}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <span className={cx('dex-ui-dot', dot)} />
              <span style={{ flex: 1, minWidth: 0, fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={label}>{label}</span>
              <strong style={{ whiteSpace: 'nowrap' }}>{unknownAll ? dash : n}<span style={{ color: 'var(--dex-gray-400)' }}>/{cap}</span></strong>
            </div>
            {!unknownAll && cap > 0 && (
              <div className="dex-ui-progress" style={{ marginTop: 5 }}>
                <div className={cx('dex-ui-progress-bar', barMod(n, cap))} style={{ width: barW(n, cap) }} />
              </div>
            )}
          </div>
        );
        const grpA = grpRow('grpA', labelA, durchActive, durchCap, 'dex-ui-dot--green');
        const grpB = grpRow('grpB', labelB, funActive, funCap, 'dex-ui-dot--orange');
        // v31.3: Auslastung nur mit Kapazität und lesbarer Liste — im Klammer-Modus
        // ist `maxParticipants` ein Alt-Wert, bei geteilten Gruppen zählt je Gruppe.
        const capTotal = selectedEvent?.maxParticipants || 0;
        // v31.3 (Nachzug): Der Schutz hing an `isConsolidatedMode` — das ist
        // im Aufrufer `subEventsOnlyMode && childEventsOf(id).length > 0`.
        // Eine Klammer, deren Kinder noch nicht geladen sind, fiel damit
        // durch und bekäme „N von M Plätzen frei" für etwas, das niemand
        // buchen kann (CLAUDE.md: `MaxParticipants` ist dort ein Alt-Wert).
        // Jetzt hängt die Aussage am Modus selbst, nicht an der Kinderzahl.
        const showCap = !isConsolidatedMode && !selectedEvent?.subEventsOnlyMode
          && !isSplitCapacity && !unknownAll && capTotal > 0;
        const over = totalActive - capTotal;
        const capText = over > 0
          ? (isDe ? `${over} über der Kapazität von ${capTotal}` : `${over} over the capacity of ${capTotal}`)
          : (isDe ? `${capTotal - totalActive} von ${capTotal} Plätzen frei` : `${capTotal - totalActive} of ${capTotal} seats free`);
        // v31.3: „—"/„≥" bekommen ihren Satz unter die Kacheln (Leitfaden 5a).
        const unknownHint = unknownAll
          ? (isDe ? 'Die Teilnehmerliste war nicht lesbar — diese Zahlen sind unbekannt, nicht 0.' : 'The participant list could not be read — these numbers are unknown, not 0.')
          : (isConsolidatedMode && subListsIncomplete
            ? (isDe ? 'Mindestens eine Termin-Liste war nicht lesbar — die Zahlen sind Untergrenzen.' : 'At least one date list could not be read — these numbers are lower bounds.')
            // v31.3 (Nachzug): Dritter Fall — im Klammer-Modus hängt die
            // Kachel „Abgemeldet" allein an `regsUnknown`. War nur die
            // Klammer-Liste gesperrt, stand ihr Strich bisher ohne jeden
            // Satz da; der Hinweis versprach eine Vollständigkeit, die er
            // nicht hatte.
            : (isConsolidatedMode && regsUnknown
              ? (isDe ? 'Die Klammer-Liste war nicht lesbar — die Zahl der Abmeldungen ist unbekannt, nicht 0.' : 'The umbrella list could not be read — the number of cancellations is unknown, not 0.')
              : ''));
        // Vier gleich gebaute Kacheln; Farbe nur mit Bedeutung (Leitfaden 5b:
        // orange Warteliste, blau eingecheckt, grau abgemeldet). Kein Hover —
        // sie filtern nichts, sie zeigen nur.
        const tile = (mod: string, value: React.ReactNode, label: string): React.ReactElement => (
          <div className={cx('dex-ui-kpi', mod)}>
            <div className="dex-ui-kpi-value">{value}</div>
            <div className="dex-ui-kpi-label">{label}</div>
          </div>
        );
        return (
          <div style={{ marginBottom: 24 }}>
            {/* v31.3: Spaltenbreiten von Hand (`2fr 1fr …`) entfallen — `dex-ui-kpi-row`
                verteilt selbst. `admin-counters` bleibt: die Mobil-Spalten haengen im SCSS daran. */}
            <div className="admin-counters dex-ui-kpi-row">
              {/* Angemeldet fuehrt — Auslastung direkt darunter. Kein `span 2`
                  fuer die Gruppen: auf einspaltigem Mobil zieht das eine zweite
                  Spalte auf; die Gruppen-Zeilen kuerzen mit Ellipse. */}
              <div className="dex-ui-kpi dex-ui-kpi--green">
                <div className="dex-ui-kpi-value">{show(totalActive, unknownAll)}</div>
                <div className="dex-ui-kpi-label">{t('status.registered')}</div>
                {showCap && (<>
                  <div className="dex-ui-progress" style={{ marginTop: 8 }}>
                    <div className={cx('dex-ui-progress-bar', barMod(totalActive, capTotal))} style={{ width: barW(totalActive, capTotal) }} />
                  </div>
                  <div className="dex-ui-kpi-sub">{capText}</div>
                </>)}
                {isSplitCapacity && (
                  <div style={{
                    marginTop: 10, paddingTop: 10, borderTop: '1px solid var(--dex-gray-200)',
                    fontSize: '0.82rem', display: 'flex', flexDirection: 'column', gap: 8,
                  }}>
                    {reversed ? <>{grpB}{grpA}</> : <>{grpA}{grpB}</>}
                  </div>
                )}
              </div>
              {/* v23.30: Warteliste direkt rechts neben „Angemeldet". */}
              {hasWaitlistKPI && tile('dex-ui-kpi--orange', show(isConsolidatedMode ? consolidatedWaitlistByEmail.size : registrations.filter(r => r.Status === 'Warteliste').length, unknownAll), t('status.waitlist'))}
              {tile('', show(isConsolidatedMode ? consolidatedQRByEmail.size : registrations.filter(r => r.Status === 'QR versendet').length, unknownAll), t('status.qrsent'))}
              {tile('dex-ui-kpi--blue', show(isConsolidatedMode ? consolidatedCheckedByEmail.size : registrations.filter(r => r.Status === 'Eingecheckt').length, unknownAll), t('status.checkedin'))}
              {tile('dex-ui-kpi--gray', show(isConsolidatedMode ? consolidatedCancelledByEmail.size : registrations.filter(r => r.Status === 'Abgemeldet').length, regsUnknown), t('status.cancelled'))}
            </div>
            {unknownHint && (
              <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm" style={{ marginTop: 10 }}>
                <span className="dex-ui-callout-icon"><AlertCircle size={14} /></span>
                <span>{unknownHint}</span>
              </div>
            )}
          </div>
        );
};

