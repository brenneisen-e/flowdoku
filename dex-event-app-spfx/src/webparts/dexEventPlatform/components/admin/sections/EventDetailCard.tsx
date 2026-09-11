/* EventDetailCard — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 6998-7623 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { formatDate, localizeStatus } from '../../../utils/eventStatus';
import { formatAllDayPeriod, isEventOver } from '../../../utils/eventFormat';
import { Calendar, Pencil, Pin, QrCode, Users } from '../../Icons';
// v31.3: Gemeinsame UI-Klassen (Pillen, Chips, Zeilen, Aktions-Kachel) —
// Hover kommt aus den Klassen, nicht mehr aus `evTabHover`.
import { cx } from '../../dexUi';
import { getCachedOrbBase64 } from '../../../services/EmailTemplates';
import { DEX_ORB_PNG } from '../../../data/brandLogos';
import { groupSubEventTabs, stripGroupPrefix } from '../../../utils/subEventGroups';
import { InfoTooltip } from '../../InfoTooltip';
import OrganizerList from '../../OrganizerList';
import { eventTeamsLink, locationWithoutTeamsUrl } from '../../../utils/teamsLink';
import { TeamsJoinButton } from '../../TeamsJoinButton';
import { ActionsDropdown } from '../../admin/ActionsMenu';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';
import { countConsolidatedActive } from '../logic/parentRegs';
import { buildEventTabs } from '../logic/eventTabs';

export interface EventDetailCardProps {
  activeRegs: SPRegistration[];
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
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
  /** v30.67 (Review): Teilnehmerliste nicht lesbar (`regLoadError`) — keine Live-Zahl aus `registrations`. */
  regsUnknown: boolean;
  reservedDetailHeight: number;
  selectedEvent: DeloitteEvent;
  setCheckInHubOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setCheckInHubStep: React.Dispatch<React.SetStateAction<"choose" | "checkin">>;
  setEvTabHover: React.Dispatch<React.SetStateAction<string>>;
  setOpenTabGroup: React.Dispatch<React.SetStateAction<string>>;
  subEventRegsByEventId: Record<string, SPRegistration[]>;
  /** v30.67 (Review): mindestens eine Termin-Liste nicht lesbar — Summen sind Untergrenzen („≥ N"). */
  subListsIncomplete: boolean;
  t: (key: string) => string;
  toggleDraftStatus: () => Promise<void>;
  waitlistRegs: SPRegistration[];
  /** v30.87: „Hinweise zu diesem Event" als Zeile unter „Aktionen" (statt eigener Kachel). */
  hintsSlot?: React.ReactNode;
}

export const EventDetailCard: React.FC<EventDetailCardProps> = (p) => {
  // v31.3: `evTabHover`/`setEvTabHover`/`isMobile` bleiben in der Schnittstelle
  // (AdminPage reicht sie weiter), werden hier aber nicht mehr gelesen — der
  // Hover kommt aus den Klassen, die Label/Wert-Zeilen sind Pillen geworden.
  const { activeRegs, childEventsOf, confirmDialog, detailCardRef, events, handleSelectEvent, isAdmin, isConsolidatedMode, isDe, isImpersonating, isLoadingRegs, isOrganizerFor, navigate, openTabGroup, registrations, regsUnknown, reservedDetailHeight, selectedEvent, setCheckInHubOpen, setCheckInHubStep, setOpenTabGroup, subEventRegsByEventId, subListsIncomplete, t, toggleDraftStatus, waitlistRegs } = p;
  // v31.3: Ableitungen für den Seitenkopf — reine Berechnungen, keine Hooks.
  const isDraft = !!selectedEvent.isFictive;
  const isFinalState = !isDraft && (selectedEvent.status === 'Completed' || selectedEvent.status === 'Cancelled');
  const canToggleStatus = (isAdmin || isOrganizerFor(selectedEvent))
    && !(isImpersonating && selectedEvent.isDemoShowcase)
    && (isDraft || selectedEvent.status === 'Active' || isFinalState);
  // Farben wie `getStatusColor`: Aktiv grün, Abgeschlossen grau, Abgesagt rot,
  // alles andere (und Entwurf) orange.
  const statusPillClass = isDraft ? 'dex-ui-pill--orange'
    : selectedEvent.status === 'Active' ? 'dex-ui-pill--green'
      : selectedEvent.status === 'Completed' ? 'dex-ui-pill--gray'
        : selectedEvent.status === 'Cancelled' ? 'dex-ui-pill--red'
          : 'dex-ui-pill--orange';
  const statusLabel = isDraft ? (isDe ? 'Entwurf' : 'Draft') : (isDe ? localizeStatus(selectedEvent.status) : selectedEvent.status);
  const eventOver = isEventOver(selectedEvent);
  const tLink = eventTeamsLink(selectedEvent);
  const locText = tLink ? locationWithoutTeamsUrl(selectedEvent.location) : (selectedEvent.location || '');
  const ownChildren = selectedEvent.parentEventId ? [] : childEventsOf(selectedEvent.id);
  const deadlineRaw = (!selectedEvent.subEventsOnlyMode && selectedEvent.registrationDeadline) ? formatDate(selectedEvent.registrationDeadline) : '-';
  const deadlineText = deadlineRaw !== '-' ? deadlineRaw : '';
  // v9.11: B2Run-Events nutzen Split-Kapazität statt maxParticipants —
  // hier die Summe anzeigen statt "Unbegrenzt". v31.3: Bei reinen Sub-Event-
  // Events ist `MaxParticipants` 0 und wäre als „Unbegrenzt" eine Aussage über
  // etwas, das niemand bucht (CLAUDE.md, v29.13) — dann „Plätze je Sub-Event".
  const splitCapacity = (selectedEvent.durchstarterCapacity || 0) + (selectedEvent.funstarterCapacity || 0);
  const effCapacity = selectedEvent.maxParticipants && selectedEvent.maxParticipants > 0
    ? selectedEvent.maxParticipants
    : splitCapacity;
  const capacityText = effCapacity
    ? `${effCapacity} ${isDe ? 'Plätze' : 'seats'}`
    : (selectedEvent.subEventsOnlyMode ? (isDe ? 'Plätze je Sub-Event' : 'Seats per sub-event') : (isDe ? 'Unbegrenzt' : 'Unlimited'));
  // Zähl-Pille in Reitern und Gruppen-Chips: auf gefülltem Chip halbtransparent
  // weiß, sonst grau — reine Anzeige, deshalb inline und ohne Hover.
  const countBadgeStyle = (filled: boolean): React.CSSProperties => ({
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    minWidth: 20, height: 18, padding: '0 6px', borderRadius: 999,
    fontSize: '0.7rem', fontWeight: 700,
    background: filled ? 'rgba(255,255,255,0.28)' : 'var(--dex-gray-100, #f5f5f5)',
    color: filled ? '#fff' : 'var(--dex-gray-700, #444)',
  });
  return (
        <div ref={detailCardRef} className="card" // v31.22: `minWidth: 0` — ein Flex-Kind hat sonst `min-width: auto`
          // und laesst sich von seinem Inhalt breiter schieben als seine Zeile.
          // Genau daran wuchs die Seite beim Wechsel (s. AdminPage).
          style={{ padding: 24, minHeight: reservedDetailHeight, flex: '1 1 420px', minWidth: 0 }}>
          {/* Header: Event-Titel + Status + Schnellaktionen (v13.11).
              v31.3: als Seitenkopf nach Leitfaden 5a — Titel, Status-Pille,
              Meta-Zeile (Zeitraum · Ort · Sub-Events · Frist), rechts nur
              „Event bearbeiten", weil der Knopf WEG von der Seite führt. */}
          <div className="dex-ui-page-head" style={{ marginBottom: 14 }}>
            <div style={{ flex: '1 1 280px', minWidth: 0 }}>
              <div className="dex-ui-inline" style={{ gap: 10 }}>
                <h2 className="dex-ui-page-head-title" style={{ fontSize: '1.3rem' }}>{selectedEvent.title}</h2>
                {/* v31.3: Die Status-Pille ist reine Anzeige (kein Hover). Das
                    Umschalten sitzt daneben als Textknopf, der SAGT, was der
                    Klick tut („Live schalten") — vorher ein „⇄" auf der Pille,
                    das man erst im Tooltip verstand. */}
                <span className={cx('dex-ui-pill', statusPillClass)}>{statusLabel}</span>
                {eventOver && !isDraft && selectedEvent.status === 'Active' && (
                  <span className="dex-ui-pill dex-ui-pill--gray" title={isDe ? 'Das Enddatum liegt in der Vergangenheit.' : 'The end date is in the past.'}>{isDe ? 'Vorbei' : 'Past'}</span>
                )}
                {/* v20.3: Status umschalten — „Aktiv" → Entwurf, „Entwurf" → live
                    (jeweils mit Sicherheitsabfrage). v22.15: auch Abgeschlossen/
                    Abgesagt sind für Admin/Organizer umschaltbar und lassen sich
                    wieder auf Aktiv setzen (vorher Sackgasse — z.B. wenn der
                    Auto-Cleanup ein Event mit altem Testdatum auf „Abgeschlossen"
                    gesetzt hatte und das Datum später korrigiert wurde). */}
                {canToggleStatus && (
                  <button
                    type="button"
                    className="dex-ui-textbtn dex-ui-textbtn--muted"
                    onClick={() => { toggleDraftStatus().catch(() => { /* */ }); }}
                    title={isDraft
                      ? (isDe ? 'Klicken: Event live schalten (Aktiv). Alle Berechtigten sehen das Event danach und können sich anmelden.' : 'Click: publish event (Active). All eligible users will see the event and can register.')
                      : isFinalState
                        ? (isDe ? 'Klicken: Event wieder auf Aktiv setzen. Danach ist es für die Berechtigten wieder sichtbar und buchbar.' : 'Click: set event back to Active. It will be visible and bookable for eligible users again.')
                        : (isDe ? 'Klicken: Event auf Entwurf setzen. Reguläre User sehen das Event danach nicht mehr; Anmeldungen bleiben erhalten.' : 'Click: set event to draft. Regular users will no longer see the event; registrations are kept.')}
                  >
                    {isDraft
                      ? (isDe ? 'Live schalten' : 'Publish')
                      : isFinalState
                        ? (isDe ? 'Wieder aktivieren' : 'Set active again')
                        : (isDe ? 'Auf Entwurf setzen' : 'Set to draft')}
                  </button>
                )}
              </div>
              <div className="dex-ui-page-head-meta">
                {/* v29.61: Bei ganztägig nur die Daten — 00:00-23:59 ist
                    die Speicherform, nicht die Aussage. */}
                <span className="dex-ui-inline" style={{ gap: 5 }} title={isDe ? 'Zeitraum' : 'Time period'}>
                  <Calendar size={14} />
                  {selectedEvent.allDay
                    ? formatAllDayPeriod(selectedEvent.startDate, selectedEvent.endDate, isDe)
                    : `${formatDate(selectedEvent.startDate)} - ${formatDate(selectedEvent.endDate)}`}
                </span>
                {/* v29.39: Steht im Ort eine Teams-URL (so haben Organizer das
                    vor dem Teams-Feld gelöst), zeigt die Ort-Angabe den Ort ohne
                    die URL; der Teilnahme-Knopf steht darunter. */}
                <span className="dex-ui-inline" style={{ gap: 5, wordBreak: 'break-word' }} title={isDe ? 'Ort' : 'Location'}>
                  <Pin size={14} />
                  {locText || (tLink ? (isDe ? 'Online' : 'Online') : '-')}
                </span>
                {ownChildren.length > 0 && (
                  <span className="dex-ui-inline" style={{ gap: 5 }}>
                    <Users size={14} />
                    {ownChildren.length} {isDe ? 'Sub-Events' : 'sub-events'}
                  </span>
                )}
                {/* v31.3: Anmeldefrist als ruhige Angabe — nicht bei reinen
                    Sub-Event-Events, dort ist die Klammer-Frist ein Alt-Wert
                    (CLAUDE.md, v29.13). */}
                {deadlineText && (
                  <span className="dex-ui-inline" style={{ gap: 5 }}>
                    {isDe ? 'Anmeldung bis' : 'Registration until'} {deadlineText}
                  </span>
                )}
              </div>
              {tLink && (
                <div style={{ marginTop: 8 }}>
                  <TeamsJoinButton url={tLink} isDe={isDe} />
                </div>
              )}
            </div>
            {/* v13.11: Event bearbeiten als Schnell-Knopf oben — die häufigste
                Aktion aus dem Aktionen-Dropdown nach oben gezogen, damit
                Organizer nicht erst scrollen müssen. */}
            <div className="dex-ui-page-head-actions">
              {/* v18.3: Im Demo-Modus ist das Demo-Event read-only — Edit /
                  Check-In / Aktionen sind ausgeblendet (kein SharePoint-
                  Backend), stattdessen ein Demo-Hinweis. */}
              {selectedEvent.isDemoShowcase ? (
                <span className="dex-ui-pill dex-ui-pill--blue">
                  {isDe ? 'Demo — nur Ansicht (keine Aktionen)' : 'Demo — view only (no actions)'}
                </span>
              ) : (
                <>
                  {(isAdmin || isOrganizerFor(selectedEvent)) && (
                    <button
                      type="button"
                      className="btn btn-secondary dex-ui-btn-sm"
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
                      style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}
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
          {/* v30.87: Hinweise als Zeile in dieser Karte statt als vierte Kachel
              unter den KPI-Kacheln. v31.3: direkt unter dem Seitenkopf
              (Leitfaden 5a: Hinweise, die Handeln verlangen, kommen vor den
              Kennzahlen und Aktionen) — die Box erklärt sich selbst, eine
              eigene Überschrift davor entfällt. */}
          {p.hintsSlot && (
            <div style={{ marginBottom: 14 }}>
              {p.hintsSlot}
            </div>
          )}
          {/* Foto links, Details rechts. Layout unabhängig vom Bildformat. */}
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
                  {/* v31.3: als `dex-ui-action`-Kachel — Hover aus der Klasse,
                      Titel plus eine Zeile Folge, wie alle Aktionen im
                      Organizer Center. */}
                  {showSciTile && (
                    <button
                      type="button"
                      className="dex-ui-action"
                      onClick={() => { setCheckInHubStep('choose'); setCheckInHubOpen(true); }}
                      title={isDe ? 'QR-Codes versenden oder Check-in am Event-Tag starten' : 'Send QR codes or start check-in on event day'}
                    >
                      <span className="dex-ui-action-icon"><QrCode size={18} /></span>
                      <span className="dex-ui-action-body">
                        <span className="dex-ui-action-title">{isDe ? 'QR-Codes und Check-In' : 'QR codes and check-in'}</span>
                        <span className="dex-ui-action-desc">{isDe ? 'Codes verschicken oder Check-in starten' : 'Send codes or start check-in'}</span>
                      </span>
                    </button>
                  )}
                </div>
              );
            })()}
            <div style={{ flex: 1, minWidth: 0 }}>
              {/* v31.3: Die Zwischenüberschrift „Event-Details" entfällt — der
                  Seitenkopf oben IST die Überschrift; eine zweite darunter
                  trug keine Information. */}
                {/* v11.28: Bookmark-Tabs statt Dropdown für schnelles Umschalten
                    zwischen Hauptevent und Sub-Events. Pro Tab wird die aktuelle
                    Teilnehmerzahl (currentParticipants aus EventContext) als
                    kleiner Badge angezeigt. */}
                {selectedEvent && (() => {
                  // v31.21: Die Reiter rechnet `logic/eventTabs` — dieselbe
                  // Quelle, aus der sie jetzt auch ueber der Teilnehmerliste
                  // stehen. Zwei Rechnungen fuer dieselbe Zahl waeren die
                  // Falle, an der die Badges schon dreimal auseinandergelaufen
                  // sind (v23.2, v30.42, v30.67 — Begruendungen dort).
                  const tabs = buildEventTabs({
                    childEventsOf, events, isDe, isLoadingRegs, registrations,
                    regsUnknown, selectedEvent, subEventRegsByEventId,
                  });
                  if (tabs.length === 0) return null;
                  // v22.70: Einzelnen Tab-Button rendern (für flaches Layout
                  // UND die Sub-Event-Reihe im Klammer-Layout wiederverwendet).
                  // v28.74 → v31.3: Der Hover-Effekt kommt jetzt aus der Klasse
                  // `dex-ui-chip` (Leitfaden: kein onMouseEnter-State für reine
                  // Optik). `evTabHover`/`setEvTabHover` bleiben in der
                  // Props-Schnittstelle, werden hier aber nicht mehr gelesen.
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const renderTab = (t: { id: string; label: string; count: number; isParent: boolean; ev: any }): React.ReactElement => {
                    const active = t.id === selectedEvent.id;
                    return (
                      <button
                        key={t.id}
                        type="button"
                        role="tab"
                        aria-selected={active}
                        className={cx('dex-ui-chip', active && 'is-active')}
                        onClick={() => handleSelectEvent(t.ev).catch(() => { /* */ })}
                        style={{ maxWidth: 280, fontSize: '0.82rem', padding: '6px 12px' }}
                        title={t.label}
                      >
                        {t.isParent && (
                          <span style={{ fontSize: '0.68rem', textTransform: 'uppercase', letterSpacing: 0.4, opacity: 0.75 }}>
                            {isDe ? 'Haupt' : 'Main'}
                          </span>
                        )}
                        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{t.label}</span>
                        <span style={countBadgeStyle(active)}>{t.count}</span>
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
                            // v30.78: offen ≠ gewählt — gefüllt (`is-active`) nur mit
                            // dem gewählten Termin darin; eine offene Gruppe ohne ihn
                            // bleibt hellgrün (s. StickyTabStrip). v31.3: Hover aus
                            // `dex-ui-chip`; nur der Offen-Zustand ist inline, weil
                            // die Klasse dafür keinen dritten Zustand kennt.
                            const openOnly: React.CSSProperties | undefined = (on && !hasSel)
                              ? { borderColor: 'var(--dex-green, #86bc25)', background: 'rgba(134,188,37,0.12)', color: 'var(--dex-green-darker, #4a7c1f)', fontWeight: 700 }
                              : undefined;
                            return (
                              <button
                                key={g.label}
                                type="button"
                                aria-expanded={on}
                                className={cx('dex-ui-chip', hasSel && 'is-active')}
                                onClick={() => setOpenTabGroup(on ? '' : g.label)}
                                title={`${g.label} — ${g.idxs.length} ${isDe ? 'Termine' : 'dates'}`}
                                style={openOnly}
                              >
                                <span>{g.label}</span>
                                <span style={countBadgeStyle(hasSel)}>{g.idxs.length}</span>
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
                        {/* Klammer-Ebene oben — volle Breite. */}
                        {/* v28.75: Hover/Fokus auch auf dem Klammer-Balken —
                            der reagierte als einziger Reiter nicht auf die Maus.
                            v31.3: als `dex-ui-row`-Zeile (Hover aus der Klasse,
                            gewählt = grüne Kante und heller Grund); die Zahl
                            links ist eine Pille. */}
                        <button
                          type="button"
                          role="tab"
                          aria-selected={pActive}
                          className={cx('dex-ui-rowbtn dex-ui-row', pActive && 'is-active')}
                          onClick={() => handleSelectEvent(parentTab.ev).catch(() => { /* */ })}
                          title={parentTab.label}
                          style={{ border: '1px solid var(--dex-gray-200, #e8e8e8)', borderRadius: 12, padding: '10px 14px' }}
                        >
                          {/* v22.73: Zahl LINKS, dann Event-Name, dann „(Klammer)"
                              + Info-Icon mit Erklärung. */}
                          <span className="dex-ui-pill dex-ui-pill--green" style={{ flexShrink: 0 }}>{parentTab.count}</span>
                          <span className="dex-ui-row-main" style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                            <span className="dex-ui-row-title" style={{ color: pActive ? 'var(--dex-green-darker, #4a7c1f)' : undefined }}>{parentTab.label}</span>
                            <span style={{ fontWeight: 600, fontSize: '0.8rem', color: 'var(--dex-gray-500)', flexShrink: 0 }}>({isDe ? 'Klammer' : 'Bracket'})</span>
                            <span style={{ flexShrink: 0, display: 'inline-flex' }} onClick={e => e.stopPropagation()}>
                              <InfoTooltip placement="bottom" text={isDe
                                ? <>Das <strong>Klammer-Event selbst wird nicht gebucht</strong> — Teilnehmer melden sich nur für die einzelnen <strong>Sub-Events</strong> an. Die Klammer fasst die Sub-Events nur zusammen. Die Zahl links zeigt, <strong>wie viele Personen sich insgesamt (kumuliert) für die Sub-Events angemeldet haben</strong>.</>
                                : <>The <strong>bracket event itself is not booked</strong> — attendees only register for the individual <strong>sub-events</strong>. The bracket just groups them. The number on the left shows <strong>how many people registered for the sub-events in total (cumulative)</strong>.</>} />
                            </span>
                          </span>
                        </button>
                        {/* Sub-Events darunter — eingerückt unter einer Klammer-Linie. */}
                        <div style={{
                          display: 'flex', flexWrap: 'wrap', gap: 6, alignItems: 'center',
                          marginLeft: 18, paddingLeft: 16, paddingTop: 10, paddingBottom: 2,
                          borderLeft: '2px solid var(--dex-green, #86bc25)',
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
                      className="dex-ui-inline"
                      style={{ gap: 6, marginBottom: 16 }}
                    >
                      {parentTab && renderTab(parentTab)}
                      {renderChildTabs()}
                    </div>
                  );
                })()}
              {/* v31.3: Statt Label/Wert-Zeilen drei ruhige Abschnitte (Leitfaden
                  5a): Organizer, Plätze und Anmeldungen als Pillen, Aktionen.
                  Zeitraum, Ort, Teams-Link und Frist stehen seit v31.3 im
                  Seitenkopf oben; die Hinweise direkt darunter. */}
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'Organizer' : 'Organizer'}</div>
                {/* v26.23: Organizer als Foto-Chips mit Hover-Kontaktkarte
                    (Position · Standort + Teams-Chat) statt reinem Klartext —
                    gleiche Komponente wie auf der Anmeldeseite. */}
                <OrganizerList
                  names={selectedEvent.organizers}
                  emails={selectedEvent.organizerEmails}
                  size="md"
                  display="chip"
                  forceIsDe={isDe}
                />
              </div>
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'Plätze und Anmeldungen' : 'Seats and registrations'}</div>
                <div className="dex-ui-inline" style={{ gap: 6 }}>
                  {/* v30.67: dieselbe Zählung wie die KPI-Kachel „Angemeldet".
                      `consolidatedFiltered` enthält bewusst auch die Warteliste
                      (die Matrix zeigt Wartende als „W") — hier stand dadurch
                      120, in der Kachel darunter 100. */}
                  {/* v30.67 (Review): nicht lesbar → „—" statt „0"; im Klammer-
                      Modus mit nicht lesbarer Termin-Liste nur eine Untergrenze.
                      v31.3: Der Grund steht sichtbar in der Pille, nicht nur im
                      Tooltip — „—" allein liest sich wie ein Anzeigefehler. */}
                  {isConsolidatedMode
                    ? (subListsIncomplete
                      ? <span className="dex-ui-pill dex-ui-pill--orange" title={isDe ? 'Mindestens — eine Termin-Liste war nicht lesbar' : 'At least — one date list was not readable'}>≥ {countConsolidatedActive(subEventRegsByEventId)} {isDe ? 'angemeldet (mindestens)' : 'registered (at least)'}</span>
                      : <span className="dex-ui-pill dex-ui-pill--green">{countConsolidatedActive(subEventRegsByEventId)} {isDe ? 'angemeldet' : 'registered'}</span>)
                    : (regsUnknown
                      ? <span className="dex-ui-pill dex-ui-pill--gray" title={isDe ? 'Liste nicht lesbar' : 'List not readable'}>— {isDe ? 'angemeldet · Liste nicht lesbar' : 'registered · list not readable'}</span>
                      : <span className="dex-ui-pill dex-ui-pill--green">{activeRegs.length} {isDe ? 'angemeldet' : 'registered'}</span>)}
                  {waitlistRegs.length > 0 && (
                    <span className="dex-ui-pill dex-ui-pill--orange">{waitlistRegs.length} {isDe ? 'auf der Warteliste' : 'on the waitlist'}</span>
                  )}
                  <span className="dex-ui-pill dex-ui-pill--gray" title={isDe ? 'Maximale Teilnehmerzahl' : 'Maximum attendees'}>{capacityText}</span>
                </div>
              </div>
              {/* v12.7: Aktionen-Dropdown direkt unter den Zahlen — alphabetisch
                  sortiert, mit Beschreibung je Action. Ersetzt die separate
                  Aktionen-Card. */}
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'Aktionen' : 'Actions'}</div>
                <ActionsDropdown isDe={isDe} />
              </div>
              {/* v12.2: 'Abgefragte Felder'-Zeile entfernt — die
                  Custom-Field-Pills hier waren redundant; sie tauchen
                  ohnehin als Spalten in der Teilnehmer-Tabelle auf. */}
            </div>
          </div>
        </div>
  );
};

