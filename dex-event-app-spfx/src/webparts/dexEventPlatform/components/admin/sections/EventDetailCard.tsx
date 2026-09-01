/* EventDetailCard — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 6998-7623 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { formatDate, getStatusColor, localizeStatus } from '../../../utils/eventStatus';
import { formatAllDayPeriod, isEventOver } from '../../../utils/eventFormat';
import { Pencil, QrCode } from '../../Icons';
import { getCachedOrbBase64 } from '../../../services/EmailTemplates';
import { DEX_ORB_PNG } from '../../../data/brandLogos';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { groupSubEventTabs, stripGroupPrefix } from '../../../utils/subEventGroups';
import { InfoTooltip } from '../../InfoTooltip';
import OrganizerList from '../../OrganizerList';
import { eventTeamsLink, locationWithoutTeamsUrl } from '../../../utils/teamsLink';
import { TeamsJoinButton } from '../../TeamsJoinButton';
import { ActionsDropdown } from '../../admin/ActionsMenu';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';
import { ConsolidatedRow } from '../../admin/adminTypes';

export interface EventDetailCardProps {
  activeRegs: SPRegistration[];
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  consolidatedFiltered: ConsolidatedRow[];
  detailCardRef: React.MutableRefObject<HTMLDivElement>;
  events: DeloitteEvent[];
  evTabHover: string;
  handleSelectEvent: (event: DeloitteEvent) => Promise<void>;
  isAdmin: boolean;
  isConsolidatedMode: boolean;
  isDe: boolean;
  isImpersonating: boolean;
  isLoadingRegs: boolean;
  isMobile: boolean;
  isOrganizerFor: (ev: DeloitteEvent) => boolean;
  navigate: (page: import("../../../context/NavigationContext").Page, eventId?: string, intent?: import("../../../context/NavigationContext").NavIntent) => void;
  openTabGroup: string;
  registrations: SPRegistration[];
  reservedDetailHeight: number;
  reservedDetailWidth: number;
  selectedEvent: DeloitteEvent;
  setCheckInHubOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setCheckInHubStep: React.Dispatch<React.SetStateAction<"choose" | "checkin">>;
  setEvTabHover: React.Dispatch<React.SetStateAction<string>>;
  setOpenTabGroup: React.Dispatch<React.SetStateAction<string>>;
  subEventRegsByEventId: Record<string, SPRegistration[]>;
  t: (key: string) => string;
  toggleDraftStatus: () => Promise<void>;
  waitlistRegs: SPRegistration[];
}

export const EventDetailCard: React.FC<EventDetailCardProps> = (p) => {
  const { activeRegs, childEventsOf, confirmDialog, consolidatedFiltered, detailCardRef, events, evTabHover, handleSelectEvent, isAdmin, isConsolidatedMode, isDe, isImpersonating, isLoadingRegs, isMobile, isOrganizerFor, navigate, openTabGroup, registrations, reservedDetailHeight, reservedDetailWidth, selectedEvent, setCheckInHubOpen, setCheckInHubStep, setEvTabHover, setOpenTabGroup, subEventRegsByEventId, t, toggleDraftStatus, waitlistRegs } = p;
  return (
        <div ref={detailCardRef} className="card" style={{ padding: 24, minHeight: reservedDetailHeight, flex: '1 1 420px', minWidth: reservedDetailWidth || 0 }}>
          {/* Header: Event-Titel + Status-Badge + Schnellaktionen (v13.11) */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap', marginBottom: 16 }}>
            <h2 style={{ margin: 0, fontSize: '1.2rem', lineHeight: 1.2 }}>{selectedEvent.title}</h2>
            {/* v20.3: Status-Badge ist klickbar — Klick auf „Aktiv" setzt das
                Event auf Entwurf, Klick auf „Entwurf" schaltet es live
                (jeweils mit Sicherheitsabfrage). v22.15: auch Abgeschlossen/
                Abgesagt sind für Admin/Organizer klickbar und lassen sich
                wieder auf Aktiv setzen (vorher Sackgasse — z.B. wenn der
                Auto-Cleanup ein Event mit altem Testdatum auf „Abgeschlossen"
                gesetzt hatte und das Datum später korrigiert wurde). */}
            {(() => {
              const isDraft = !!selectedEvent.isFictive;
              const badgeBg = isDraft ? 'rgba(237,139,0,0.15)' : getStatusColor(selectedEvent.status) + '22';
              const badgeFg = isDraft ? 'var(--dex-orange-dark, #b35a00)' : getStatusColor(selectedEvent.status);
              const label = isDraft ? 'ENTWURF' : (isDe ? localizeStatus(selectedEvent.status) : selectedEvent.status);
              const isFinalState = !isDraft && (selectedEvent.status === 'Completed' || selectedEvent.status === 'Cancelled');
              const canToggleStatus = (isAdmin || isOrganizerFor(selectedEvent))
                && !(isImpersonating && selectedEvent.isDemoShowcase)
                && (isDraft || selectedEvent.status === 'Active' || isFinalState);
              if (!canToggleStatus) {
                return (
                  <span className="badge" style={{ background: badgeBg, color: badgeFg }}>{label}</span>
                );
              }
              return (
                <button
                  type="button"
                  className="badge"
                  onClick={() => { toggleDraftStatus().catch(() => { /* */ }); }}
                  title={isDraft
                    ? (isDe ? 'Klicken: Event live schalten (Aktiv). Alle Berechtigten sehen das Event danach und können sich anmelden.' : 'Click: publish event (Active). All eligible users will see the event and can register.')
                    : isFinalState
                      ? (isDe ? 'Klicken: Event wieder auf Aktiv setzen. Danach ist es für die Berechtigten wieder sichtbar und buchbar.' : 'Click: set event back to Active. It will be visible and bookable for eligible users again.')
                      : (isDe ? 'Klicken: Event auf Entwurf setzen. Reguläre User sehen das Event danach nicht mehr; Anmeldungen bleiben erhalten.' : 'Click: set event to draft. Regular users will no longer see the event; registrations are kept.')}
                  style={{
                    background: badgeBg, color: badgeFg,
                    border: `1px solid ${badgeFg}`,
                    cursor: 'pointer',
                    display: 'inline-flex', alignItems: 'center', gap: 5,
                  }}
                >
                  {label}
                  <span style={{ fontSize: '0.75em', opacity: 0.85 }}>⇄</span>
                </button>
              );
            })()}
            {/* v13.11: Event bearbeiten + Check-In starten als Schnell-
                Buttons direkt neben dem Status-Badge — die häufigsten
                Aktionen aus dem Aktionen-Dropdown nach oben gezogen,
                damit Organizer am Eventtag nicht erst scrollen müssen. */}
            <div style={{ display: 'flex', gap: 8, marginLeft: 'auto', flexWrap: 'wrap' }}>
              {/* v18.3: Im Demo-Modus ist das Demo-Event read-only — Edit /
                  Check-In / Aktionen sind ausgeblendet (kein SharePoint-
                  Backend), stattdessen ein Demo-Hinweis. */}
              {selectedEvent.isDemoShowcase ? (
                <span style={{
                  display: 'inline-flex', alignItems: 'center', gap: 6,
                  fontSize: '0.8rem', fontWeight: 600, color: 'var(--dex-blue, #0076a8)',
                  background: 'rgba(0,118,168,0.08)', border: '1px solid var(--dex-blue, #0076a8)',
                  borderRadius: 999, padding: '4px 12px',
                }}>
                  {isDe ? 'Demo — nur Ansicht (keine Aktionen)' : 'Demo — view only (no actions)'}
                </span>
              ) : (
                <>
                  {(isAdmin || isOrganizerFor(selectedEvent)) && (
                    <button
                      type="button"
                      className="btn btn-secondary"
                      onClick={async () => {
                        // Admins bearbeiten direkt (voller Zugriff).
                        if (isAdmin) { navigate('edit-event', selectedEvent.id); return; }
                        // v24.7 (O): abgeschlossene/vergangene Events sind für
                        // Organizer als Archivierungsschutz NICHT mehr bearbeitbar.
                        if (isEventOver(selectedEvent)) {
                          const ok = await confirmDialog(
                            isDe
                              ? 'Dieses Event ist bereits vorbei und damit abgeschlossen — Bearbeiten ist als Archivierungsschutz nicht mehr möglich. Möchtest du stattdessen ein neues Event anlegen (du kannst ein bestehendes als Vorlage nutzen)?'
                              : 'This event is over and therefore completed — editing is locked (archival protection). Would you like to create a new event instead?',
                            { confirmLabel: isDe ? 'Neues Event anlegen' : 'Create new event', cancelLabel: isDe ? 'Abbrechen' : 'Cancel' });
                          if (ok) navigate('create-event');
                          return;
                        }
                        // v24.8 (P korrigiert): aktive Events direkt bearbeiten —
                        // der „lieber neues Event?"-Hinweis erscheint NUR bei
                        // abgeschlossenen Events (siehe Zweig oben).
                        navigate('edit-event', selectedEvent.id);
                      }}
                      style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: '0.85rem', padding: '6px 12px' }}
                      title={t('admin.editbutton') || (isDe ? 'Event bearbeiten' : 'Edit event')}
                    >
                      <Pencil size={14} />
                      {isDe ? 'Event bearbeiten' : 'Edit event'}
                    </button>
                  )}
                  {/* v30.38: „Check-In starten" ist hier entfallen. Er sprang
                      direkt auf die Scan-Seite und war damit ein zweiter,
                      engerer Weg neben dem Einstieg „QR-Codes und Check-In"
                      unter dem Event-Bild — zwei Knöpfe für dieselbe Absicht,
                      von denen einer die Vorauswahl überspringt. Der Einstieg
                      unten führt in beide Richtungen (Codes verschicken /
                      Check-in) und bleibt der einzige. */}
                </>
              )}
            </div>
          </div>
          {/* Foto immer als Kreis links, Detail-Rows rechts. Layout
              unabhängig vom Bildformat (cover-Crop sorgt für den Kreis). */}
          <div style={{ display: 'flex', gap: 20, alignItems: 'flex-start' }}>
            {/* v12.6: Event-Bild jetzt prominent als großes Rechteck-
                Format (wie auf der Registrierungs-Seite) statt kleinem
                Avatar-Kreis. Hintergrund weiß für saubere Darstellung
                transparenter PNG-Logos.
                v20.2: darunter die Self-Check-in-QR-Kachel — sichtbar ab
                5 Tagen vor Event-Start ODER sobald QR-Codes versendet
                wurden (und solange das Event nicht länger als 1 Tag vorbei
                ist). Klick öffnet das Erklär-/Einstell-Modal. */}
            {(() => {
              const canManageSci = isAdmin || isOrganizerFor(selectedEvent);
              const dayMs = 24 * 60 * 60 * 1000;
              const startTs = selectedEvent.startDate ? new Date(selectedEvent.startDate).getTime() : 0;
              const endTs = selectedEvent.endDate ? new Date(selectedEvent.endDate).getTime() : startTs;
              const nowTs = Date.now();
              const notLongPast = (endTs || startTs) === 0 || nowTs <= (endTs || startTs) + dayMs;
              // v30.38: Der Knopf steht jetzt IMMER unter dem Bild, solange das
              // Event nicht lange vorbei ist. Die Fünf-Tage-Regel (`within5Days`)
              // und die QR-Phase stammen aus v20.2, als die Kachel direkt ins
              // Self-Check-in-Modal sprang — dafür war „kurz vor dem Event" der
              // richtige Zeitpunkt. Seit v30.38 führt sie in den Einstieg
              // „QR-Codes und Check-In", und den braucht man VORHER: QR-Codes
              // verschickt man in der Woche davor, nicht am Tag davor. Ein
              // Einstieg, den man erst findet, wenn es zu spät ist, ist keiner.
              const showSciTile = canManageSci && notLongPast;
              // v28.90: Ohne Event-Foto blieb die rechte Spalte leer und die
              // Detail-Zeilen liefen über die volle Breite — die Ansicht sah je
              // Event unterschiedlich aus, je nachdem ob jemand ein Bild
              // hochgeladen hatte. Statt Leerraum steht dort jetzt das
              // DEX-Logo als Platzhalter. Es ist bewusst NICHT das
              // gespeicherte Bild: Nichts wird geschrieben, Mails und
              // Anmeldeseite bleiben unverändert bildlos.
              return (
                <div style={{ flex: '0 0 auto', width: 260, maxWidth: '38%', display: 'flex', flexDirection: 'column', gap: 12 }}>
                  {selectedEvent.imageUrl ? (
                    <div
                      style={{
                        background: '#fff',
                        borderRadius: 'var(--dex-radius, 12px)',
                        overflow: 'hidden',
                        border: '1px solid var(--dex-gray-200, #e5e7eb)',
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                      }}
                    >
                      <img
                        src={selectedEvent.imageUrl}
                        alt={selectedEvent.title}
                        style={{
                          display: 'block',
                          width: '100%',
                          height: 'auto',
                          maxHeight: 240,
                          objectFit: 'contain',
                        }}
                      />
                    </div>
                  ) : (
                    // v28.91: Das BILD, nicht die animierte Canvas-Version.
                    // Ein sich drehendes Logo an der Stelle eines Event-Fotos
                    // zieht den Blick auf sich, obwohl es nur sagt „hier ist
                    // kein Bild". Darunter der Hinweis, dass es der Platzhalter
                    // ist und wo man ein eigenes Foto hinterlegt.
                    <div>
                      <div
                        style={{
                          background: '#fff',
                          borderRadius: 'var(--dex-radius, 12px)',
                          overflow: 'hidden',
                          border: '1px solid var(--dex-gray-200, #e5e7eb)',
                          display: 'flex', alignItems: 'center', justifyContent: 'center',
                          padding: 18,
                        }}
                      >
                        <img
                          src={getCachedOrbBase64() || DEX_ORB_PNG}
                          alt=""
                          style={{ display: 'block', width: '100%', height: 'auto', maxHeight: 200, objectFit: 'contain' }}
                        />
                      </div>
                      <p style={{ margin: '6px 2px 0', fontSize: '0.72rem', color: 'var(--dex-gray-500)', lineHeight: 1.4 }}>
                        {isDe
                          ? <>Standardfoto — ein eigenes Bild hinterlegst du über <strong>&bdquo;Event bearbeiten&ldquo;</strong>.</>
                          : <>Default image — set your own via <strong>&bdquo;Edit event&ldquo;</strong>.</>}
                      </p>
                    </div>
                  )}
                  {/* v30.38: Führt jetzt in denselben Einstieg wie im
                      Aktionen-Menü („QR-Codes und Check-In"), statt direkt ins
                      Self-Check-in-Modal zu springen. Vorher war das die einzige
                      Stelle, an der eine der fünf Check-in-Varianten ohne
                      Vorauswahl heraussprang — genau die Ungleichbehandlung, die
                      v30.36 im Aktionen-Menü aufgelöst hat. Wer hier klickt, will
                      „Check-in", nicht „ausgerechnet die Self-Variante". */}
                  {showSciTile && (
                    <button
                      type="button"
                      onClick={() => { setCheckInHubStep('choose'); setCheckInHubOpen(true); }}
                      title={isDe ? 'QR-Codes versenden oder Check-in am Event-Tag starten' : 'Send QR codes or start check-in on event day'}
                      style={{
                        background: '#fff',
                        border: '1px solid var(--dex-green, #86bc25)',
                        borderRadius: 'var(--dex-radius, 12px)',
                        padding: 12, cursor: 'pointer',
                        display: 'flex', alignItems: 'center', gap: 12, textAlign: 'left',
                      }}
                    >
                      <span style={{ width: 64, height: 64, flexShrink: 0, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', background: 'rgba(134,188,37,0.10)', borderRadius: 8, color: 'var(--dex-green-dark, #4a7c1f)' }}>
                        <QrCode size={32} />
                      </span>
                      <span style={{ minWidth: 0 }}>
                        <span style={{ display: 'block', fontWeight: 700, fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>
                          {isDe ? 'QR-Codes und Check-In' : 'QR codes and check-in'}
                        </span>
                        <span style={{ display: 'block', fontSize: '0.74rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                          {isDe ? 'Codes verschicken oder Check-in starten' : 'Send codes or start check-in'}
                        </span>
                      </span>
                    </button>
                  )}
                </div>
              );
            })()}
            <div style={{ flex: 1, minWidth: 0 }}>
              <h3 className="mb-16">{isDe ? 'Event-Details' : 'Event details'}</h3>
                {/* v11.28: Bookmark-Tabs statt Dropdown für schnelles Umschalten
                    zwischen Hauptevent und Sub-Events. Pro Tab wird die aktuelle
                    Teilnehmerzahl (currentParticipants aus EventContext) als
                    kleiner Badge angezeigt. */}
                {selectedEvent && (() => {
                  const isChild = !!selectedEvent.parentEventId;
                  const siblings = isChild
                    ? childEventsOf(selectedEvent.parentEventId || '')
                    : childEventsOf(selectedEvent.id);
                  if (!isChild && siblings.length === 0) return null;
                  const parent = isChild ? events.find(e => e.id === selectedEvent.parentEventId) : selectedEvent;
                  // v22.75: Der aktuell GEWÄHLTE Tab zeigt die LIVE-Zahl aus den
                  // gerade geladenen Registrierungen (registrations) — die
                  // Tab-Badges stammen sonst aus dem zwischengespeicherten
                  // Event-Zustand (letzter Listen-Load) und hinken neuen
                  // Anmeldungen hinterher (Badge 112 vs. Tabelle 126).
                  const liveSelectedActive = registrations.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt').length;
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const tabs: Array<{ id: string; label: string; count: number; isParent: boolean; ev: any }> = [];
                  if (parent) {
                    // v22.64: Im Klammer-Modus zeigt der HAUPT-Badge die ECHTE
                    // Zahl eindeutiger aktiver Personen über alle Sub-Events
                    // (live), nicht den gespeicherten Counter `currentParticipants`
                    // der Klammer — der zählt nicht buchbare Klammern unzuverlässig
                    // und kann verrutschen.
                    let parentCount = parent.currentParticipants || 0;
                    const pKids = childEventsOf(parent.id);
                    const haveSubData = parent.subEventsOnlyMode && pKids.length > 0 && pKids.every(c => subEventRegsByEventId[c.id] !== undefined);
                    if (haveSubData) {
                      const activeSet = new Set<string>();
                      for (const c of pKids) {
                        for (const r of (subEventRegsByEventId[c.id] || [])) {
                          if (r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt') {
                            // v23.3: emaillose Zeile = trotzdem ein Kopf → per
                            // Zeilen-Id mitzaehlen statt verschlucken (sonst zeigt
                            // die Klammer weniger als die Sub-Event-Tabelle).
                            const k = (r.ParticipantEmail || '').toLowerCase().trim() || `__noemail#${c.id}#${r.Id}`;
                            activeSet.add(k);
                          }
                        }
                      }
                      parentCount = activeSet.size;
                    } else if (parent.id === selectedEvent.id && !isLoadingRegs) {
                      // Normales Hauptevent ist selbst gewählt → Live-Zahl.
                      // v30.42: Während des Ladens gehört `registrations` noch
                      // dem vorher gewählten Termin — dieselbe Falle wie bei den
                      // Sub-Reitern unten. Dann lieber der eigene Zähler.
                      parentCount = liveSelectedActive;
                    }
                    tabs.push({ id: parent.id, label: parent.title || (isDe ? 'Hauptevent' : 'Main event'), count: parentCount, isParent: true, ev: parent });
                  }
                  for (const c of siblings) {
                    // v23.2: Nicht-gewählte Sub-Tabs zeigen — sofern die Liste
                    // bereits geladen ist — die LIVE-Zeilenzahl aus
                    // subEventRegsByEventId statt des veralteten Counters
                    // `currentParticipants`. Sonst „springt" der Badge je nach
                    // gewähltem Tab (gewählt = live, sonst = Cache), siehe der
                    // 188-vs-190-Effekt. Gewählter Tab bleibt die Live-Zahl der
                    // aktuell geladenen Tabelle.
                    // v30.42: …aber NUR, solange nicht gerade geladen wird.
                    // Beim Reiterwechsel gehört `registrations` noch dem ALTEN
                    // Termin; der neue Reiter zeigte deshalb kurz dessen Zahl
                    // (Nutzer-Befund: erst 55, dann 51). Während des Ladens
                    // steht die eigene Zahl des Termins da — die stimmt sofort.
                    const subRegs = subEventRegsByEventId[c.id];
                    const subLiveCount = subRegs
                      ? subRegs.filter(r => r.Status === 'Angemeldet' || r.Status === 'QR versendet' || r.Status === 'Eingecheckt').length
                      : (c.currentParticipants || 0);
                    tabs.push({ id: c.id, label: shortSubEventTitle(c.title, parent?.title) || (isDe ? 'ohne Titel' : 'untitled'), count: (c.id === selectedEvent.id && !isLoadingRegs) ? liveSelectedActive : subLiveCount, isParent: false, ev: c });
                  }
                  // v22.70: Einzelnen Tab-Button rendern (für flaches Layout
                  // UND die Sub-Event-Reihe im Klammer-Layout wiederverwendet).
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const renderTab = (t: { id: string; label: string; count: number; isParent: boolean; ev: any }): React.ReactElement => {
                    const active = t.id === selectedEvent.id;
                    return (
                      <button
                        key={t.id}
                        type="button"
                        role="tab"
                        aria-selected={active}
                        onClick={() => handleSelectEvent(t.ev).catch(() => { /* */ })}
                        onMouseEnter={() => setEvTabHover(t.id)}
                        onMouseLeave={() => setEvTabHover(prev => (prev === t.id ? null : prev))}
                        onFocus={() => setEvTabHover(t.id)}
                        onBlur={() => setEvTabHover(prev => (prev === t.id ? null : prev))}
                        style={{
                          display: 'inline-flex', alignItems: 'center', gap: 8,
                          padding: '8px 14px',
                          // v28.74: Hover-/Fokus-Effekt wie im Wizard — ohne
                          // Reaktion auf die Maus lasen sich die Reiter wie eine
                          // Beschriftung statt wie etwas Anklickbares.
                          border: `1px solid ${(evTabHover === t.id && !active) ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                          borderBottom: active ? '2px solid var(--dex-green, #86bc25)' : `1px solid ${(evTabHover === t.id) ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                          borderRadius: '8px 8px 0 0',
                          background: active ? '#fff' : ((evTabHover === t.id) ? 'rgba(134,188,37,0.14)' : 'var(--dex-gray-50, #fafafa)'),
                          color: active ? 'var(--dex-green-dark, #4a7c1f)' : ((evTabHover === t.id) ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-700)'),
                          fontWeight: active ? 700 : ((evTabHover === t.id) ? 600 : 500),
                          fontSize: '0.85rem',
                          cursor: 'pointer',
                          marginBottom: -1,
                          whiteSpace: 'nowrap',
                          maxWidth: 280,
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                          transform: (evTabHover === t.id && !active) ? 'translateY(-1px)' : 'none',
                          boxShadow: (evTabHover === t.id && !active) ? '0 -2px 6px rgba(134,188,37,0.20)' : 'none',
                          transition: 'background 0.15s, color 0.15s, border-color 0.15s, transform 0.15s, box-shadow 0.15s',
                        }}
                        title={t.label}
                      >
                        {t.isParent && (
                          <span style={{ fontSize: '0.7rem', textTransform: 'uppercase', letterSpacing: 0.4, color: active ? 'var(--dex-green-dark)' : 'var(--dex-gray-400)' }}>
                            {isDe ? 'Haupt' : 'Main'}
                          </span>
                        )}
                        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>{t.label}</span>
                        <span
                          style={{
                            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            minWidth: 24, height: 20, padding: '0 6px',
                            borderRadius: 999,
                            background: active ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)',
                            color: active ? '#fff' : 'var(--dex-gray-700)',
                            fontSize: '0.72rem', fontWeight: 700,
                          }}
                        >
                          {t.count}
                        </span>
                      </button>
                    );
                  };
                  const parentTab = tabs.find(tb => tb.isParent);
                  const childTabs = tabs.filter(tb => !tb.isParent);

                  // v30.60: Bei vielen gleich präfigierten Terminen („Day 1 - …",
                  // „Day 2 - …") erst die Gruppen zeigen, dann auf Klick die
                  // Termine dieser Gruppe. Ein Training mit 25 Sessions füllte
                  // sonst sechs Zeilen, und die Fünf-Tage-Struktur, die in den
                  // Namen steckt, war nur beim Lesen jeder Kachel zu erkennen.
                  // Ob überhaupt gruppiert wird, entscheidet utils/subEventGroups —
                  // bei wenigen Terminen bleibt die Leiste flach.
                  const grouping = groupSubEventTabs(childTabs.map(tb => tb.label));
                  const renderChildTabs = (): React.ReactNode => {
                    if (!grouping.grouped) return childTabs.map(t => renderTab(t));
                    // Voreingestellt offen: die Gruppe des gerade gewählten
                    // Termins. Sonst stünde man nach dem Wechsel vor
                    // zugeklappten Gruppen und müsste seinen eigenen Termin suchen.
                    const selIdx = childTabs.findIndex(tb => tb.id === selectedEvent.id);
                    const autoOpen = selIdx >= 0
                      ? (grouping.groups.filter(g => g.idxs.indexOf(selIdx) >= 0)[0] || null)
                      : null;
                    const openLabel = openTabGroup !== null ? openTabGroup : (autoOpen ? autoOpen.label : grouping.groups[0].label);
                    return (
                      <div style={{ width: '100%' }}>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 8 }}>
                          {grouping.groups.map(g => {
                            const on = g.label === openLabel;
                            const hasSel = selIdx >= 0 && g.idxs.indexOf(selIdx) >= 0;
                            const sum = g.idxs.reduce((n, i) => n + (childTabs[i] ? childTabs[i].count : 0), 0);
                            return (
                              <button
                                key={g.label}
                                type="button"
                                onClick={() => setOpenTabGroup(on ? '' : g.label)}
                                title={`${g.label} — ${g.idxs.length} ${isDe ? 'Termine' : 'dates'}`}
                                style={{
                                  display: 'inline-flex', alignItems: 'center', gap: 7,
                                  padding: '6px 14px', borderRadius: 999, cursor: 'pointer',
                                  border: `1px solid ${on || hasSel ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                                  background: on ? 'var(--dex-green, #86bc25)' : '#fff',
                                  color: on ? '#fff' : 'var(--dex-gray-700)',
                                  fontWeight: on || hasSel ? 700 : 500, fontSize: '0.82rem',
                                  transition: 'background 0.15s, border-color 0.15s',
                                }}
                              >
                                <span>{g.label}</span>
                                <span style={{
                                  display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                  minWidth: 22, height: 18, padding: '0 5px', borderRadius: 999,
                                  background: on ? 'rgba(255,255,255,0.28)' : 'var(--dex-gray-100)',
                                  color: on ? '#fff' : 'var(--dex-gray-600)',
                                  fontSize: '0.68rem', fontWeight: 700,
                                }}>{g.idxs.length}</span>
                                {/* Der Punkt hinter der Zahl ist die Summe der
                                    Anmeldungen dieses Tages — sie ist der Grund,
                                    warum man eine Gruppe überhaupt aufmacht. */}
                                <span style={{ fontSize: '0.72rem', opacity: 0.8 }}>· {sum}</span>
                                <span style={{ fontSize: '0.7rem' }}>{on ? '▾' : '▸'}</span>
                              </button>
                            );
                          })}
                        </div>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, alignItems: 'flex-end' }}>
                          {(grouping.groups.filter(g => g.label === openLabel)[0] || { idxs: [] }).idxs.map(i => {
                            const t = childTabs[i];
                            if (!t) return null;
                            // Innerhalb von „Day 1" heißt der Reiter „PMO", nicht
                            // „Day 1 - PMO" — das Präfix steht schon oben.
                            return renderTab({ ...t, label: stripGroupPrefix(t.label, openLabel) });
                          })}
                        </div>
                      </div>
                    );
                  };
                  // v22.70: Im Klammer-Modus die Klammer als ECHTE Klammer ÜBER
                  // den Sub-Event-Tabs darstellen (oben volle Breite, darunter die
                  // eingerückten Sub-Events). Normales Hauptevent bleibt flach.
                  const klammerLayout = !!(parentTab && parentTab.ev && parentTab.ev.subEventsOnlyMode && childTabs.length > 0);
                  if (klammerLayout && parentTab) {
                    const pActive = parentTab.id === selectedEvent.id;
                    return (
                      <div role="tablist" aria-label={isDe ? 'Event wechseln' : 'Switch event'} style={{ marginBottom: 16 }}>
                        {/* Klammer-Ebene oben — volle Breite, gefüllter Kopf. */}
                        {/* v28.75: Hover/Fokus auch auf dem Klammer-Balken —
                            der reagierte als einziger Reiter nicht auf die Maus. */}
                        <button
                          type="button"
                          role="tab"
                          aria-selected={pActive}
                          onClick={() => handleSelectEvent(parentTab.ev).catch(() => { /* */ })}
                          onMouseEnter={() => setEvTabHover(parentTab.id)}
                          onMouseLeave={() => setEvTabHover(prev => (prev === parentTab.id ? null : prev))}
                          onFocus={() => setEvTabHover(parentTab.id)}
                          onBlur={() => setEvTabHover(prev => (prev === parentTab.id ? null : prev))}
                          title={parentTab.label}
                          style={{
                            display: 'flex', alignItems: 'center', gap: 10, width: '100%',
                            padding: '10px 16px', cursor: 'pointer', textAlign: 'left',
                            border: `1.5px solid ${(pActive || evTabHover === parentTab.id) ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                            borderRadius: '10px 10px 0 0',
                            // v28.86: Ruhezustand weiß (s. Wizard).
                            background: pActive
                              ? 'var(--dex-green, #86bc25)'
                              : (evTabHover === parentTab.id ? 'rgba(134,188,37,0.14)' : '#fff'),
                            color: pActive ? '#fff' : 'var(--dex-green-dark, #4a7c1f)',
                            fontWeight: 700, fontSize: '0.9rem',
                            boxShadow: (evTabHover === parentTab.id && !pActive) ? 'inset 0 0 0 1px rgba(134,188,37,0.35)' : 'none',
                            transition: 'background 0.15s, border-color 0.15s, box-shadow 0.15s',
                          }}
                        >
                          {/* v22.73: Zahl LINKS, dann Event-Name, dann „(Klammer)"
                              + Info-Icon mit Erklärung. */}
                          <span style={{
                            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            minWidth: 26, height: 22, padding: '0 8px', borderRadius: 999,
                            background: pActive ? 'rgba(255,255,255,0.25)' : 'var(--dex-green, #86bc25)',
                            color: '#fff', fontSize: '0.74rem', fontWeight: 700, flexShrink: 0,
                          }}>{parentTab.count}</span>
                          <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', flex: 1, display: 'inline-flex', alignItems: 'center', gap: 6, minWidth: 0, color: pActive ? '#fff' : 'var(--dex-green-dark, #4a7c1f)' }}>
                            <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', color: pActive ? '#fff' : 'var(--dex-green-dark, #4a7c1f)' }}>{parentTab.label}</span>
                            <span style={{ fontWeight: 600, opacity: 0.9, flexShrink: 0, color: pActive ? '#fff' : 'var(--dex-green-dark, #4a7c1f)' }}>({isDe ? 'Klammer' : 'Bracket'})</span>
                            <span style={{ flexShrink: 0, display: 'inline-flex', color: pActive ? '#fff' : 'var(--dex-green-dark, #4a7c1f)' }} onClick={e => e.stopPropagation()}>
                              <InfoTooltip placement="bottom" text={isDe
                                ? <>Das <strong>Klammer-Event selbst wird nicht gebucht</strong> — Teilnehmer melden sich nur für die einzelnen <strong>Sub-Events</strong> an. Die Klammer fasst die Sub-Events nur zusammen. Die Zahl links zeigt, <strong>wie viele Personen sich insgesamt (kumuliert) für die Sub-Events angemeldet haben</strong>.</>
                                : <>The <strong>bracket event itself is not booked</strong> — attendees only register for the individual <strong>sub-events</strong>. The bracket just groups them. The number on the left shows <strong>how many people registered for the sub-events in total (cumulative)</strong>.</>} />
                            </span>
                          </span>
                        </button>
                        {/* Sub-Events darunter — eingerückt unter einer Klammer-Linie. */}
                        <div style={{
                          display: 'flex', flexWrap: 'wrap', gap: 6, alignItems: 'flex-end',
                          marginLeft: 18, paddingLeft: 16, paddingTop: 10,
                          borderLeft: '2px solid var(--dex-green, #86bc25)',
                          borderBottom: '1px solid var(--dex-gray-200)',
                        }}>
                          {renderChildTabs()}
                        </div>
                      </div>
                    );
                  }
                  return (
                    <div
                      role="tablist"
                      aria-label={isDe ? 'Event wechseln' : 'Switch event'}
                      style={{
                        display: 'flex', flexWrap: 'wrap', gap: 6,
                        marginBottom: 16,
                        borderBottom: '1px solid var(--dex-gray-200)',
                        paddingBottom: 0,
                      }}
                    >
                      {parentTab && renderTab(parentTab)}
                      {renderChildTabs()}
                    </div>
                  );
                })()}
              {/* Eigenes Row-Layout (zwei Spalten: Label fett, Wert links-
                  bündig). Das globale .settings-info SCSS macht stattdessen
                  space-between (also Wert rechts-bündig) — hier wollen wir
                  beide Spalten links ausgerichtet. */}
              {(() => {
                const rowStyle: React.CSSProperties = {
                  display: 'grid',
                  gridTemplateColumns: isMobile ? '1fr' : '160px 1fr',
                  gap: isMobile ? 2 : 12,
                  padding: '10px 0',
                  borderBottom: '1px solid var(--dex-gray-200)',
                  fontSize: '0.9rem',
                };
                const labelStyle: React.CSSProperties = { fontWeight: 700, color: 'var(--dex-gray-700)' };
                const valueStyle: React.CSSProperties = { fontWeight: 400, color: 'var(--dex-gray-800)' };
                return (
                  <>
                    <div style={rowStyle}>
                      <span style={labelStyle}>{isDe ? 'Zeitraum' : 'Time period'}</span>
                      {/* v29.61: Bei ganztägig nur die Daten — 00:00-23:59 ist
                          die Speicherform, nicht die Aussage. */}
                      <span style={valueStyle}>
                        {selectedEvent.allDay
                          ? formatAllDayPeriod(selectedEvent.startDate, selectedEvent.endDate, isDe)
                          : `${formatDate(selectedEvent.startDate)} - ${formatDate(selectedEvent.endDate)}`}
                      </span>
                    </div>
                    <div style={rowStyle}>
                      <span style={labelStyle}>{isDe ? 'Organizer' : 'Organizer'}</span>
                      {/* v26.23: Organizer als Foto-Chips mit Hover-Kontaktkarte
                          (Position · Standort + Teams-Chat) statt reinem Klartext —
                          gleiche Komponente wie auf der Anmeldeseite. */}
                      <span style={{ ...valueStyle, display: 'inline-flex' }}>
                        <OrganizerList
                          names={selectedEvent.organizers}
                          emails={selectedEvent.organizerEmails}
                          size="md"
                          display="chip"
                          forceIsDe={isDe}
                        />
                      </span>
                    </div>
                    {/* v29.39: Steht im Ort eine Teams-URL (so haben Organizer
                        das vor dem Teams-Feld gelöst), lief sie hier als roher
                        Text aus der Karte und war nicht klickbar. Jetzt zeigt
                        die Ort-Zeile den Ort ohne die URL, und darunter steht
                        ein Teilnahme-Knopf. */}
                    {(() => {
                      const tLink = eventTeamsLink(selectedEvent);
                      const locText = tLink ? locationWithoutTeamsUrl(selectedEvent.location) : (selectedEvent.location || '');
                      return (
                        <>
                          <div style={rowStyle}>
                            <span style={labelStyle}>{isDe ? 'Ort' : 'Location'}</span>
                            <span style={{ ...valueStyle, wordBreak: 'break-word' }}>{locText || (tLink ? (isDe ? 'Online' : 'Online') : '-')}</span>
                          </div>
                          {tLink && (
                            <div style={rowStyle}>
                              <span style={labelStyle}>{isDe ? 'Online-Teilnahme' : 'Join online'}</span>
                              <span style={valueStyle}><TeamsJoinButton url={tLink} isDe={isDe} /></span>
                            </div>
                          )}
                        </>
                      );
                    })()}
                    <div style={rowStyle}>
                      <span style={labelStyle}>{isDe ? 'Max. Teilnehmer' : 'Max. attendees'}</span>
                      <span style={valueStyle}>{(() => {
                        // v9.11: B2Run-Events nutzen Split-Kapazität statt maxParticipants —
                        // hier die Summe anzeigen statt "Unbegrenzt".
                        const split = (selectedEvent.durchstarterCapacity || 0) + (selectedEvent.funstarterCapacity || 0);
                        const eff = selectedEvent.maxParticipants && selectedEvent.maxParticipants > 0
                          ? selectedEvent.maxParticipants
                          : split;
                        return eff || (isDe ? 'Unbegrenzt' : 'Unlimited');
                      })()}</span>
                    </div>
                    <div style={rowStyle}>
                      <span style={labelStyle}>{isDe ? 'Aktuell registriert' : 'Currently registered'}</span>
                      <span style={valueStyle}>{isConsolidatedMode ? consolidatedFiltered.length : activeRegs.length}</span>
                    </div>
                    {waitlistRegs.length > 0 && (
                      <div style={rowStyle}>
                        <span style={labelStyle}>{isDe ? 'Warteliste' : 'Waitlist'}</span>
                        <span style={valueStyle}>{waitlistRegs.length}</span>
                      </div>
                    )}
                    {/* v12.7: Aktionen-Dropdown direkt unter „Aktuell
                        registriert" — alphabetisch sortiert, mit Hover-
                        Tooltip pro Action (desc-Text rechts daneben).
                        Ersetzt die separate Aktionen-Card. */}
                    <div style={{ marginTop: 14 }}>
                      <div style={{ ...labelStyle, marginBottom: 6 }}>
                        {isDe ? 'Aktionen' : 'Actions'}
                      </div>
                      <ActionsDropdown isDe={isDe} />
                    </div>
                    {/* v12.2: 'Abgefragte Felder'-Zeile entfernt — die
                        Custom-Field-Pills hier waren redundant; sie tauchen
                        ohnehin als Spalten in der Teilnehmer-Tabelle auf. */}
                  </>
                );
              })()}
            </div>
          </div>
        </div>
  );
};

