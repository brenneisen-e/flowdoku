/* EventSpecificSection — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Station 3: Starter-Typ-Auswahl bei geteilter Kapazitaet, die Sub-Event-Auswahl
 * (Liste und Kalender) und die eventspezifischen Felder. Inhalt zeichengleich
 * uebernommen; die Anzeige-Bedingung ist beim Aufrufer geblieben. */
import * as React from 'react';
import { CollapsibleSection, formatDate, subEventDescHtml } from './regHelpers';
import { subEventRegDeadline } from '../../utils/eventFormat';
import { isoToLocal } from '../../utils/berlinTime';
import { Icon } from '@fluentui/react/lib/Icon';
import { Locale } from '../../context/LanguageContext';
import { DeloitteEvent, EventSpecificField } from '../../types';

/** Station 3 — Starter-Typ, Sub-Event-Auswahl und eventspezifische Felder. */
export interface EventSpecificSectionProps {
  childEvents: DeloitteEvent[];
  childOneDe: string;
  childTermPlural: string;
  childTermSingular: string;
  dayHoverKey: string;
  durchCap: number;
  event: DeloitteEvent;
  eventSpecific: Record<string, string>;
  funCap: number;
  hasStarterBlockMapping: boolean;
  hiddenChildCount: number;
  isAdmin: boolean;
  isMobile: boolean;
  isOrganizer: boolean;
  isSessionsOnlyMode: boolean;
  isSplitGroup: boolean;
  locale: Locale;
  parentAlreadyRegistered: boolean;
  parentFullNoWaitlist: boolean;
  parentRegBlocked: boolean;
  preferredStarterType: string;
  registerForOther: boolean;
  registerForParent: boolean;
  renderMainFieldsSection: () => React.ReactElement;
  renderRegField: (fRaw: EventSpecificField, store?: Record<string, string>, setStore?: (next: Record<string, string>) => void, rowIndex?: number, rowList?: EventSpecificField[]) => React.ReactElement;
  renderSubEventInlineFields: (ce: DeloitteEvent) => React.ReactElement | null;
  resolveMainEventLabel: (defaultLabel: string) => string | null;
  selectedSessions: Set<string>;
  sessionFieldValues: Record<string, Record<string, string>>;
  sessionMeta: Record<string, { count: number; wasRegistered: boolean; }>;
  setDayHoverKey: React.Dispatch<React.SetStateAction<string>>;
  setEventSpecific: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  setPendingSubEventModal: React.Dispatch<React.SetStateAction<{ subEventId: string; draftValues: Record<string, string>; }>>;
  setPreferredStarterType: React.Dispatch<React.SetStateAction<string>>;
  setRegisterForParent: React.Dispatch<React.SetStateAction<boolean>>;
  setSelectedSessions: React.Dispatch<React.SetStateAction<Set<string>>>;
  setSessionFieldValues: React.Dispatch<React.SetStateAction<Record<string, Record<string, string>>>>;
  showErrors: boolean;
  splitLabelA: string;
  splitLabelB: string;
  /** v30.67: null = Belegung nicht ermittelbar (Strich statt erfundener 0). */
  starterCounts: { durch: number; fun: number; durchWait: number; funWait: number; } | null;
  subOpenFrom: (startIso?: string) => Date | null;
  t: (key: string) => string;
  tEvent: (key: string) => string;
}
export const EventSpecificSection: React.FC<EventSpecificSectionProps> = (p) => {
  const { childEvents, childOneDe, childTermPlural, childTermSingular, dayHoverKey, durchCap, event, eventSpecific, funCap, hasStarterBlockMapping, hiddenChildCount, isAdmin, isMobile, isOrganizer, isSessionsOnlyMode, isSplitGroup, locale, parentAlreadyRegistered, parentFullNoWaitlist, parentRegBlocked, preferredStarterType, registerForOther, registerForParent, renderMainFieldsSection, renderRegField, renderSubEventInlineFields, resolveMainEventLabel, selectedSessions, sessionFieldValues, sessionMeta, setDayHoverKey, setEventSpecific, setPendingSubEventModal, setPreferredStarterType, setRegisterForParent, setSelectedSessions, setSessionFieldValues, showErrors, splitLabelA, splitLabelB, starterCounts, subOpenFrom, t, tEvent } = p;
  return (
        <div className="registration-specific">
          {/* v11.97: Section-Header + „* = Required field"-Legende in
              einer Zeile. Legende mit ROTEM Stern (vorher war der Stern
              in der Erklärung grau, jetzt im Deloitte-Rot wie alle echten
              Required-Marker). */}
          <CollapsibleSection
            isMobile={isMobile}
            icon="EditNote"
            title={t('reg.eventinfo')}
            collapsible={false}
            headerExtra={
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 12, padding: '0 12px' }}>
              <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                <span style={{ color: 'var(--dex-red, #da291c)', fontWeight: 700, marginRight: 2 }}>*</span> = {t('reg.requiredfield')}
              </span>
            </span>
            }
          >
          <div style={{ padding: '24px 20px' }}>
            {/* v11.10: Group-Selection ist ein eigener, IMMER sichtbarer
                Block (sofern das Event Split-Capacity hat). Vorher war er
                inkorrekt INNERHALB der Sub-Events-Auswahl genistet, sodass
                Events ohne Sub-Events keine Gruppen-Buttons hatten und
                Drittpersonen-Registrierungen die Gruppen-Wahl gar nicht
                anzeigten. Sub-Events erben jetzt einfach
                preferredStarterType — keine Pro-Sub-Event-Radios mehr. */}
            {isSplitGroup && (
              <div style={{ marginBottom: 20, border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 16 }}>
                <label className="form-label" style={{ fontWeight: 700, marginBottom: 6 }}>
                  <span className="required">*</span> {(event.splitSectionTitle && event.splitSectionTitle.trim()) ? event.splitSectionTitle : (locale === 'de' ? 'Gruppen-Auswahl' : 'Group selection')}
                </label>
                {/* v26.83: Organizer-eigener Hinweistext (splitHelpText) hat
                    Vorrang; sonst der Standardsatz. whiteSpace pre-wrap, damit
                    Zeilenumbrüche aus dem Wizard erhalten bleiben. */}
                <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 10, whiteSpace: 'pre-wrap' }}>
                  {(event.splitHelpText && event.splitHelpText.trim())
                    ? event.splitHelpText
                    : (locale === 'de'
                      ? `Wähle eine der zwei Gruppen aus. Ist die Wunsch-Gruppe voll, kannst du automatisch in die andere wechseln oder auf der Warteliste warten.`
                      : 'Pick one of the two groups. If your preferred group is full, you can either switch to the other or join the waitlist.')}
                </p>
                {/* v19.19: Gesamt-Kapazitäts-Zusammenfassung — Gesamtzahl der
                    Plätze, aktuell freie Plätze (geklammert ≥ 0) und die Zahl
                    der Personen auf der Warteliste. WICHTIG: hier wird NIE eine
                    Überbuchung sichtbar — freie Plätze sind bei 0 gedeckelt,
                    bei Vollbelegung steht „ausgebucht". Die echte
                    Überbuchungszahl ist ausschließlich dem Organizer/Admin
                    im Admin Center vorbehalten. */}
                {(() => {
                  const totalCap = durchCap + funCap;
                  const dActive = starterCounts?.durch ?? 0;
                  const fActive = starterCounts?.fun ?? 0;
                  const totalFree = Math.max(0, durchCap - dActive) + Math.max(0, funCap - fActive);
                  const totalWait = (starterCounts?.durchWait ?? 0) + (starterCounts?.funWait ?? 0);
                  return (
                    <div style={{
                      display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap',
                      marginBottom: 12, padding: '8px 12px', borderRadius: 8,
                      background: 'var(--dex-gray-50, #f7f7f7)', border: '1px solid var(--dex-gray-200)',
                      fontSize: '0.8rem',
                    }}>
                      <span style={{ color: 'var(--dex-gray-700)', fontWeight: 700 }}>
                        {locale === 'de' ? 'Gesamtkapazität:' : 'Total capacity:'} {totalCap} {locale === 'de' ? 'Plätze' : 'seats'}
                      </span>
                      {/* v26.72: „X frei" in der Gesamt-Zeile entfernt — die
                          Verfügbarkeit steht bereits pro Gruppe in den Karten.
                          Nur bei komplett ausgebucht bleibt ein Hinweis. */}
                      {totalFree <= 0 && (
                        <>
                          <span style={{ color: 'var(--dex-gray-300)' }}>·</span>
                          <span style={{ color: 'var(--dex-red, #c00)', fontWeight: 700 }}>
                            {locale === 'de' ? 'ausgebucht' : 'fully booked'}
                          </span>
                        </>
                      )}
                      {totalWait > 0 && (
                        <>
                          <span style={{ color: 'var(--dex-gray-300)' }}>·</span>
                          <span style={{ color: 'var(--dex-gray-600)' }}>
                            {totalWait} {locale === 'de'
                              ? (totalWait === 1 ? 'Person auf der Warteliste' : 'Personen auf der Warteliste')
                              : (totalWait === 1 ? 'person on the waitlist' : 'people on the waitlist')}
                            {event.splitSharedWaitlist ? (locale === 'de' ? ' (gemeinsam)' : ' (shared)') : ''}
                          </span>
                        </>
                      )}
                    </div>
                  );
                })()}
                <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
                  {(() => {
                    // v26.72: Beschreibung pro Gruppe frei konfigurierbar
                    // (splitDescA/B aus dem Wizard); Fallback auf den B2Run-
                    // Standardtext nur bei den Default-Labels.
                    const optA = { id: 'Durchstarter', label: splitLabelA, desc: (event?.splitDescA && event.splitDescA.trim()) || (splitLabelA === 'Durchstarter' ? t('reg.starter.durch.desc') : ''), cap: durchCap, count: starterCounts?.durch ?? 0, wait: starterCounts?.durchWait ?? 0 };
                    const optB = { id: 'Funstarter', label: splitLabelB, desc: (event?.splitDescB && event.splitDescB.trim()) || (splitLabelB === 'Funstarter' ? t('reg.starter.fun.desc') : ''), cap: funCap, count: starterCounts?.fun ?? 0, wait: starterCounts?.funWait ?? 0 };
                    // v11.25: pure UI-Reihenfolge — bei reversed wird Karte B
                    // zuerst gerendert. Interne IDs/Capacities/StarterType der
                    // Anmeldungen bleiben unangetastet.
                    return event.splitDisplayOrderReversed ? [optB, optA] : [optA, optB];
                  })().map(opt => {
                    const free = opt.cap - opt.count;
                    const isFull = free <= 0;
                    const isActive = preferredStarterType === opt.id;
                    // v26.72: gewählte Box grün, nicht-gewählte grau (vorher A grün / B orange).
                    const accent = isActive ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-500, #6b7280)';
                    return (
                      <button
                        key={opt.id}
                        type="button"
                        onClick={() => setPreferredStarterType(opt.id)}
                        style={{
                          padding: 14, textAlign: 'left',
                          borderRadius: 'var(--dex-radius, 12px)',
                          border: isActive ? `2px solid ${accent}` : '2px solid var(--dex-gray-200)',
                          // v26.88: Standard-Grün (wie die Geschlecht-/Feld-Füllung,
                          // greenFilledStyle) statt Off-Brand #f0fdf4.
                          background: isActive ? 'rgba(134,188,37,0.06)' : '#fff',
                          cursor: 'pointer', transition: 'all 0.15s',
                          position: 'relative',
                        }}
                      >
                        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
                          <strong style={{ color: accent, fontSize: '0.95rem' }}>{opt.label}</strong>
                          {isActive && <span style={{ color: accent, fontSize: '0.8rem' }}>✓</span>}
                        </div>
                        {opt.desc && <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginBottom: 6, whiteSpace: 'pre-wrap' }}>{opt.desc}</div>}
                        <div style={{ fontSize: '0.78rem' }}>
                          {isFull ? (
                            <span style={{ color: 'var(--dex-red, #c00)', fontWeight: 600 }}>{t('reg.starter.full')}</span>
                          ) : (
                            // v19.19: nie negativ — bei Überbuchung greift der
                            // isFull-Zweig oben (zeigt „Voll"), die echte
                            // Überbuchungszahl bleibt dem Organizer/Admin vorbehalten.
                            // v30.67: Ohne ermittelte Belegung (starterCounts null) keinen
                            // erfundenen Wert — „50 / 50 frei" war die Zahl, die eine volle
                            // Gruppe als frei verkaufte (Teilnehmerliste zeilenweise gesichert).
                            <span style={{ color: accent }}>{starterCounts ? `${Math.max(0, free)} / ${opt.cap} ${t('reg.starter.free')}` : `— / ${opt.cap} ${t('reg.starter.free')}`}</span>
                          )}
                          {/* v19.19: Warteliste pro Gruppe — nur bei GETRENNTEN
                              Wartelisten. Bei gemeinsamer Warteliste steht die
                              Zahl gesammelt in der Kapazitäts-Zusammenfassung. */}
                          {!event.splitSharedWaitlist && opt.wait > 0 && (
                            <span style={{ display: 'block', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                              {opt.wait} {locale === 'de'
                                ? (opt.wait === 1 ? 'Person auf der Warteliste' : 'Personen auf der Warteliste')
                                : (opt.wait === 1 ? 'person on the waitlist' : 'people on the waitlist')}
                            </span>
                          )}
                        </div>
                      </button>
                    );
                  })}
                </div>

                {event.durchstarterRequiresProof && preferredStarterType === 'Durchstarter' && (
                  <div style={{ marginTop: 12, padding: '10px 12px', background: 'rgba(237,139,0,0.06)', border: '1px solid var(--dex-orange)', borderRadius: 8 }}>
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', fontSize: '0.85rem' }}>
                      <input
                        type="checkbox"
                        checked={eventSpecific['b2run_leistungsnachweis'] === 'true'}
                        onChange={e => setEventSpecific({ ...eventSpecific, b2run_leistungsnachweis: e.target.checked ? 'true' : 'false' })}
                        style={{ marginTop: 3 }}
                      />
                      <span>
                        <strong>{t('reg.starter.proof') || 'Leistungsnachweis vorhanden'} <span className="required">*</span></strong>
                        <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                          {t('reg.starter.proof.hint') || 'Ich bestätige, dass ein entsprechender Leistungsnachweis (z.B. Wettkampfergebnis, Trainingsnachweis) vorliegt.'}
                        </span>
                      </span>
                    </label>
                    {showErrors && eventSpecific['b2run_leistungsnachweis'] !== 'true' && (
                      <div style={{ marginTop: 6, fontSize: '0.75rem', color: 'var(--dex-red)' }}>
                        {t('reg.starter.proof.required') || 'Bitte Leistungsnachweis bestätigen.'}
                      </div>
                    )}
                  </div>
                )}
                {/* v11.12: Custom-Fields mit onlyForGroup-Constraint
                    direkt INNERHALB der Gruppen-Auswahl-Box rendern.
                    Klappt erst auf, wenn der User eine Gruppe gewählt
                    hat und das Feld der gewählten Gruppe entspricht.
                    Gleicher orange-getönter Style wie der Legacy-
                    Leistungsnachweis-Block — sodass jede Gruppen-
                    spezifische Abfrage optisch klar als „Folge der
                    Gruppen-Wahl" erkennbar ist. */}
                {preferredStarterType && (() => {
                  const groupSpec = event.eventSpecificFields
                    .filter(f => f.id !== 'b2run_mobilnummer' || eventSpecific['b2run_infoservice'] === 'true')
                    .filter(f => !(f.id === 'b2run_startblock' && hasStarterBlockMapping))
                    .filter(f => {
                      if (!f.showIf || !f.showIf.fieldId) return true;
                      const raw = (eventSpecific[f.showIf.fieldId] || '').trim();
                      if (!raw) return false;
                      const answers = raw.indexOf(' | ') >= 0
                        ? raw.split(' | ').map(s => s.trim()).filter(Boolean)
                        : [raw];
                      return answers.some(a => f.showIf!.values.indexOf(a) >= 0);
                    })
                    .filter(f => {
                      const grp = f.onlyForGroup;
                      if (!grp || grp === 'all') return false;
                      if (grp === 'A') return preferredStarterType === 'Durchstarter';
                      if (grp === 'B') return preferredStarterType === 'Funstarter';
                      return false;
                    });
                  if (groupSpec.length === 0) return null;
                  const labelA = (event.splitLabelA && event.splitLabelA.trim()) || 'Durchstarter';
                  const labelB = (event.splitLabelB && event.splitLabelB.trim()) || 'Funstarter';
                  const grpLabel = preferredStarterType === 'Durchstarter' ? labelA : labelB;
                  return (
                    <div style={{
                      marginTop: 12, padding: '12px 14px',
                      background: 'rgba(237,139,0,0.06)',
                      border: '1px solid var(--dex-orange)',
                      borderRadius: 8,
                      display: 'flex', flexDirection: 'column', gap: 12,
                    }}>
                      <div style={{ fontSize: '0.78rem', fontWeight: 700, color: 'var(--dex-orange, #ed8b00)' }}>
                        {locale === 'de'
                          ? `Zusätzliche Angaben für „${grpLabel}"`
                          : `Additional details for „${grpLabel}"`}
                      </div>
                      {groupSpec.map(f => renderRegField(f))}
                    </div>
                  );
                })()}
              </div>
            )}

            {/* v29.9: Sackgasse sichtbar machen — die Klammer ist nicht
                buchbar, und von ihren Programmpunkten ist für diese Person
                keiner freigegeben. Ohne diesen Hinweis steht da eine
                Anmeldeseite ohne irgendetwas zum Anklicken, und der Grund
                zeigt sich erst beim Klick auf „Registrieren".
                Bewusst NUR in diesem Fall: Ist die Klammer selbst buchbar,
                ist ein enger gefasstes Sub-Event Absicht des Organizers und
                geht die Person nichts an. */}
            {childEvents.length === 0 && hiddenChildCount > 0 && event.subEventsOnlyMode && (
              <div style={{
                marginTop: 12, padding: '10px 14px', borderRadius: 8,
                background: 'rgba(237,139,0,0.10)', border: '1px solid var(--dex-orange, #ed8b00)',
                fontSize: '0.82rem', color: 'var(--dex-orange-dark, #b35a00)', lineHeight: 1.5,
              }}>
                {locale === 'de'
                  ? 'Die Anmeldung läuft hier ausschließlich über die einzelnen Programmpunkte. Für dich ist aktuell keiner davon freigegeben — wenn du teilnehmen möchtest, wende dich bitte an die Organizer.'
                  : 'Registration here runs exclusively via the individual programme items. None of them is currently released for you — if you would like to attend, please contact the organizers.'}
              </div>
            )}

            {/* v11.10: Sub-Events-Auswahl als eigener Block.
                v18.37: jetzt AUCH im Stellvertreter-Modus („Für andere
                registrieren") sichtbar — der Submit-Pfad meldet ausgewählte
                Sub-Events längst für die Zielperson an (registerForEvent mit
                participantEmail), nur die Auswahl-UI war versehentlich hinter
                !registerForOther versteckt. Die Belegungszahlen werden im
                Stellvertreter-Modus über getAllRegistrations geladen, ohne
                Self-Vorbelegung. Der Haupt-Event-Checkbox wird im
                Stellvertreter-Modus ausgeblendet (das Haupt-Event wird dort
                ohnehin immer mit angemeldet). Sub-Events erben
                preferredStarterType vom Group-Selection-Block oben. */}
            {childEvents.length > 0 && (
              <div style={{ marginBottom: 20, border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 16 }}>
                {/* v15.11: im subEventsOnlyMode ist die Hauptevent-Anmeldung
                    deaktiviert — Überschrift + Hinweis entsprechend
                    anpassen, sonst lesen sich „Haupt-Event und … können
                    unabhängig" widersprüchlich. */}
                <h4 style={{ marginTop: 0, marginBottom: 4, fontSize: '0.95rem' }}>
                  {event.subEventsOnlyMode
                    ? (childTermPlural
                        ? (locale === 'de' ? `${childTermPlural} auswählen` : `Select ${childTermPlural}`)
                        : (locale === 'de' ? 'Sub-Events auswählen' : 'Select sub-events'))
                    : (tEvent('reg.selection.title') || 'Wofür möchtest du dich anmelden?')}
                </h4>
                {/* v28.97: Sagen, dass nur eines geht — sonst wundert man sich,
                    warum die vorherige Auswahl verschwindet. */}
                {event.subEventSingleChoice && (
                  <p style={{ fontSize: '0.82rem', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600, marginTop: 0, marginBottom: 8 }}>
                    {locale === 'de'
                      ? `Du kannst genau ${childOneDe} auswählen — ein neuer Klick ersetzt die bisherige Wahl.`
                      : 'You can pick exactly one — a new click replaces your previous choice.'}
                  </p>
                )}
                <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 12 }}>
                  {event.subEventsOnlyMode
                    ? (locale === 'de'
                        ? `Bitte wähle mindestens ${childOneDe} aus, um dich anzumelden.`
                        : `Please pick at least one ${childTermSingular || 'sub-event'} you want to register for.`)
                    : registerForOther
                      ? (locale === 'de'
                          ? `Die Person wird für das Haupt-Event angemeldet. Wähle zusätzlich die gewünschten ${childTermPlural || 'Sub-Events'} aus.`
                          : `The person will be registered for the main event. Additionally pick the desired ${childTermPlural || 'sub-events'}.`)
                    : (childTermPlural
                        ? (locale === 'de'
                            ? `Haupt-Event und ${childTermPlural} können unabhängig voneinander an- oder abgewählt werden.`
                            : `Main event and ${childTermPlural} can be selected or deselected independently.`)
                        : (tEvent('reg.selection.hint') || 'Haupt-Event und Sessions können unabhängig voneinander an- oder abgewählt werden.'))}
                </p>

                {/* v15.7: Hauptevent-Card auch hier ausblenden bei
                    subEventsOnlyMode — gleicher Fix wie der primäre Pfad
                    weiter oben. Vorher wurde dieser Render-Pfad (Register
                    for someone else) übersehen.
                    v18.37: im Stellvertreter-Modus ebenfalls ausblenden — die
                    Person wird dort immer für das Haupt-Event angemeldet, ein
                    steuerbarer Haken wäre irreführend. */}
                {!event.subEventsOnlyMode && !registerForOther && (
                <label style={{
                  display: 'flex', alignItems: 'flex-start', gap: 10, padding: 10,
                  borderRadius: 8,
                  border: `1px solid ${registerForParent && !parentAlreadyRegistered && !parentRegBlocked ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                  background: registerForParent && !parentAlreadyRegistered && !parentRegBlocked ? 'rgba(134,188,37,0.06)' : '#fff',
                  cursor: (parentAlreadyRegistered || parentRegBlocked) ? 'default' : 'pointer',
                  opacity: parentRegBlocked ? 0.6 : 1,
                }}>
                  <input
                    type="checkbox"
                    checked={parentAlreadyRegistered ? true : (parentRegBlocked ? false : registerForParent)}
                    disabled={parentAlreadyRegistered || parentRegBlocked}
                    onChange={e => setRegisterForParent(e.target.checked)}
                    style={{ marginTop: 2 }}
                  />
                  <div style={{ flex: 1 }}>
                    {(() => { const lbl = resolveMainEventLabel(tEvent('reg.selection.mainevent') || 'Haupt-Event'); return (
                      <div style={{ fontWeight: 700 }}>{lbl ? `${lbl}: ` : ''}{event.title}</div>
                    ); })()}
                    {parentAlreadyRegistered && (
                      <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                        {tEvent('reg.selection.alreadyregistered') || 'Du bist bereits für das Haupt-Event angemeldet.'}
                      </div>
                    )}
                    {parentRegBlocked && !parentAlreadyRegistered && (
                      <div style={{ fontSize: '0.75rem', color: parentFullNoWaitlist ? 'var(--dex-red, #c00)' : 'var(--dex-orange, #ed8b00)', marginTop: 2 }}>
                        {parentFullNoWaitlist
                          ? (locale === 'de' ? 'Alle Plätze sind belegt — die Warteliste ist für dieses Event deaktiviert.' : 'All seats are taken — the waitlist is disabled for this event.')
                          : (tEvent('reg.subevents.deadlinepassed') || 'Anmeldefrist abgelaufen — nur noch die offenen Sub-Events sind wählbar.')}
                      </div>
                    )}
                  </div>
                </label>
                )}
                {/* v29.28: Die Fragen zum Haupt-Event stehen DIREKT unter
                    seiner Kachel — dort, wo sie hingehören; die Sub-Event-
                    Fragen stecken in deren Karten (v29.27). Im Stellvertreter-
                    Modus (Haupt-Kachel ausgeblendet, Anmeldung fürs Haupt-
                    Event läuft trotzdem) stehen sie an derselben Stelle über
                    den Sub-Events. Bei einer Klammer rendert der Block an der
                    alten Stelle unter der Auswahl (übergreifende Fragen). */}
                {!event.subEventsOnlyMode && (event.eventSpecificFields.length > 0 || isSplitGroup) && (
                  <div style={{ margin: '8px 0 4px' }}>{renderMainFieldsSection()}</div>
                )}

                {/* v28.91: Termin-Kalender statt Liste — nur, wenn der
                    Organizer die Sub-Events ausdrücklich als Termine angelegt
                    hat (subEventCalendar). Bei neun Tagen ist eine Liste aus
                    neun Funkbuttons kaum zu erfassen; im Kalender sieht man
                    Wochenstruktur, Lücken und freie Plätze auf einen Blick.
                    Es sind dieselben Sub-Events und dasselbe selectedSessions —
                    nur eine andere Darstellung derselben Auswahl. */}
                {!!event.subEventCalendar && (() => {
                  type DayEntry = { ce: typeof childEvents[0]; key: string };
                  // v30.67: Der Kalendertag ist der BERLINER Tag, nicht der des
                  // Browsers. Ein Termin „Di 00:00" steht als 22:00Z des Vortags
                  // in der Liste; `getDate()` lieferte auf einem UTC-Browser den
                  // Montag, die Kachel rutschte einen Tag nach links — gebucht
                  // wurde trotzdem der Dienstag. `isoToLocal` ist dieselbe
                  // Umrechnung, die der Wizard beim Anlegen benutzt.
                  const dayOf = (iso?: string): string => {
                    if (!iso) return '';
                    return isoToLocal(iso).slice(0, 10);
                  };
                  const entries: DayEntry[] = childEvents
                    .map(ce => ({ ce, key: dayOf(ce.startDate) }))
                    .filter(e => !!e.key);
                  if (entries.length === 0) return null;
                  const byDay: Record<string, DayEntry> = {};
                  entries.forEach(e => { byDay[e.key] = e; });
                  // Monate, in denen Termine liegen — jeder als eigenes Raster.
                  const monthKeys: string[] = [];
                  entries.forEach(e => {
                    const mk = e.key.slice(0, 7);
                    if (monthKeys.indexOf(mk) < 0) monthKeys.push(mk);
                  });
                  monthKeys.sort();
                  const weekdays = locale === 'de'
                    ? ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So']
                    : ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
                  const pickDay = (ce: typeof childEvents[0], isSel: boolean, disabled: boolean): void => {
                    if (disabled) return;
                    if (!isSel) {
                      // Hat der Termin eigene Abfragefelder, läuft die Auswahl
                      // über denselben Modal-Flow wie in der Listen-Ansicht.
                      if ((ce.eventSpecificFields || []).length > 0) {
                        setPendingSubEventModal({ subEventId: ce.id, draftValues: { ...(sessionFieldValues[ce.id] || {}) } });
                        return;
                      }
                      // v28.97: „Genau eines" — die neue Wahl ERSETZT die alte,
                      // statt sich danebenzulegen. Sonst müsste der Teilnehmer
                      // erst abwählen und würde bei jedem Wechsel scheitern.
                      const next = event.subEventSingleChoice ? new Set<string>() : new Set(selectedSessions);
                      next.add(ce.id);
                      setSelectedSessions(next);
                      if (event.subEventSingleChoice) {
                        setSessionFieldValues(prev => {
                          const keep: typeof prev = {};
                          if (prev[ce.id]) keep[ce.id] = prev[ce.id];
                          return keep;
                        });
                      }
                      return;
                    }
                    const next = new Set(selectedSessions);
                    next.delete(ce.id);
                    setSelectedSessions(next);
                    setSessionFieldValues(prev => { const c = { ...prev }; delete c[ce.id]; return c; });
                  };
                  return (
                    <div style={{ marginTop: 12 }}>
                      {/* v29.60: Die Zeile „Termine auswählen — angebotene Tage sind hervorgehoben"
                          stand direkt unter der Ueberschrift des Blocks (z.B.
                          „Office-Tage auswählen") und sagte dasselbe noch
                          einmal. Zwei Aufforderungen hintereinander lesen sich
                          wie zwei Schritte. Der Hinweis auf die Hervorhebung
                          ist ueberfluessig, weil man sie sieht. */}
                      {monthKeys.map(mk => {
                        const [my, mm] = mk.split('-').map(n => parseInt(n, 10));
                        const first = new Date(my, mm - 1, 1);
                        const daysInMonth = new Date(my, mm, 0).getDate();
                        // Montag als erster Wochentag (getDay: So=0).
                        const lead = (first.getDay() + 6) % 7;
                        const cells: Array<string | null> = [];
                        for (let i = 0; i < lead; i++) cells.push(null);
                        for (let d = 1; d <= daysInMonth; d++) {
                          cells.push(`${my}-${String(mm).padStart(2, '0')}-${String(d).padStart(2, '0')}`);
                        }
                        return (
                          <div key={mk} style={{ marginBottom: 14 }}>
                            <div style={{ fontWeight: 700, fontSize: '0.9rem', marginBottom: 6, color: 'var(--dex-gray-700, #444)' }}>
                              {first.toLocaleDateString(locale === 'de' ? 'de-DE' : 'en-GB', { month: 'long', year: 'numeric' })}
                            </div>
                            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: 4 }}>
                              {weekdays.map(w => (
                                <div key={w} style={{ fontSize: '0.68rem', fontWeight: 700, color: 'var(--dex-gray-400)', textAlign: 'center', padding: '2px 0' }}>{w}</div>
                              ))}
                              {cells.map((key, i) => {
                                if (!key) return <div key={`e${i}`} />;
                                const entry = byDay[key];
                                const dayNum = parseInt(key.slice(8), 10);
                                if (!entry) {
                                  return (
                                    <div key={key} style={{
                                      textAlign: 'center', padding: '8px 0', borderRadius: 8,
                                      fontSize: '0.8rem', color: 'var(--dex-gray-300, #ccc)',
                                    }}>{dayNum}</div>
                                  );
                                }
                                const ce = entry.ce;
                                const meta = sessionMeta[ce.id] || { count: null, wasRegistered: false };
                                const isSel = selectedSessions.has(ce.id);
                                const hasCap = typeof ce.maxParticipants === 'number' && ce.maxParticipants > 0;
                                // v30.62: s. Listen-Ansicht — unbekannt ist nicht voll.
                                const isFull = hasCap && meta.count !== null && meta.count >= (ce.maxParticipants || 0);
                                // v30.8: effektive Tages-Frist — materialisierte Spalte,
                    // sonst Fallback aus der rollierenden Regel (Spalte kann
                    // bei einem unter Drosselung abgebrochenen Save leer sein).
                    const effRegDeadline = subEventRegDeadline(event, ce);
                    const deadlinePassed = !!(effRegDeadline && new Date(effRegDeadline) < new Date());
                                // v29.28: Frist-Bypass für Organizer/Admins (s. Listen-Pfad).
                                const deadlineLocked = deadlinePassed && !isOrganizer && !isAdmin;
                                // v29.67: Freischalt-Regel — der Tag öffnet erst X Tage
                                // vorher (mode 'day') bzw. X Tage vor dem Montag seiner
                                // Woche (mode 'week', die ganze KW öffnet gemeinsam).
                                // Organizer/Admins dürfen vorher — derselbe Bypass wie
                                // bei der Anmeldefrist, damit sie testen können.
                                const openFrom = ((): Date | null => {
                                  const rule = event.subEventOpenRule;
                                  if (!rule) return null;
                                  // v29.76: festes Datum — gilt fuer ALLE Tage gleich.
                                  if (rule.mode === 'fixed') {
                                    const d = new Date(rule.date || '');
                                    return isFinite(d.getTime()) ? d : null;
                                  }
                                  if (!((rule.days || 0) > 0)) return null;
                                  const [oy, om, od] = key.split('-').map(n => parseInt(n, 10));
                                  const dayDate = new Date(oy, om - 1, od);
                                  if (rule.mode === 'week') {
                                    // getDay(): So=0 — auf Montag der Woche zurückrechnen.
                                    dayDate.setDate(dayDate.getDate() - ((dayDate.getDay() + 6) % 7));
                                  }
                                  dayDate.setDate(dayDate.getDate() - (rule.days || 0));
                                  dayDate.setHours(0, 0, 0, 0);
                                  return dayDate;
                                })();
                                const notYetOpen = !!openFrom && new Date() < openFrom;
                                const openLocked = notYetOpen && !isOrganizer && !isAdmin;
                                const openFromLabel = openFrom
                                  ? openFrom.toLocaleDateString(locale === 'de' ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit' })
                                  : '';
                                const disabled = (isFull && !isSel) || (deadlineLocked && !isSel) || (openLocked && !isSel);
                                // v30.62: -1 = keine Kapazität, null = Zahl unbekannt.
                                const free = (hasCap && meta.count !== null)
                                  ? Math.max(0, (ce.maxParticipants || 0) - meta.count)
                                  : (hasCap ? null : -1);
                                const title = [
                                  ce.title || '',
                                  // v30.67: Derselbe Dreiweg wie in der Beschriftung — `free`
                                  // ist bewusst null, wenn keine Quelle trägt; „null von 80
                                  // Plätzen frei" war die Erfindung, die v30.62 aus der
                                  // Kachel entfernt hat, im Tooltip aber stehen ließ.
                                  hasCap
                                    ? (free !== null
                                      ? (locale === 'de' ? `${free} von ${ce.maxParticipants} Plätzen frei` : `${free} of ${ce.maxParticipants} seats free`)
                                      : (locale === 'de' ? `${ce.maxParticipants} Plätze · Belegung nicht ermittelbar` : `${ce.maxParticipants} seats · occupancy unknown`))
                                    : (locale === 'de' ? 'Unbegrenzte Plätze' : 'Unlimited seats'),
                                  deadlinePassed
                                    ? (deadlineLocked
                                      ? (locale === 'de' ? 'Anmeldefrist abgelaufen' : 'Registration deadline passed')
                                      : (locale === 'de' ? 'Anmeldefrist abgelaufen — als Organizer/Admin trotzdem wählbar' : 'Registration deadline passed — still selectable as organizer/admin'))
                                    : '',
                                  ce.mandatoryRegistration ? (locale === 'de' ? 'Pflichttermin' : 'Mandatory date') : '',
                                  notYetOpen
                                    ? (openLocked
                                      ? (locale === 'de' ? `Anmeldung ab ${openFrom!.toLocaleDateString('de-DE')} möglich` : `Registration opens on ${openFrom!.toLocaleDateString('en-GB')}`)
                                      : (locale === 'de' ? `Anmeldung öffnet regulär am ${openFrom!.toLocaleDateString('de-DE')} — als Organizer/Admin trotzdem wählbar` : `Registration opens on ${openFrom!.toLocaleDateString('en-GB')} — still selectable as organizer/admin`))
                                    : '',
                                ].filter(Boolean).join(' · ');
                                return (
                                  <button
                                    key={key}
                                    type="button"
                                    onClick={() => pickDay(ce, isSel, disabled)}
                                    onMouseEnter={() => { if (!disabled) setDayHoverKey(key); }}
                                    onMouseLeave={() => setDayHoverKey(h => (h === key ? '' : h))}
                                    onFocus={() => { if (!disabled) setDayHoverKey(key); }}
                                    onBlur={() => setDayHoverKey(h => (h === key ? '' : h))}
                                    disabled={disabled}
                                    title={title}
                                    aria-pressed={isSel}
                                    style={{
                                      display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center',
                                      gap: 1, padding: '6px 0 5px', borderRadius: 8, minHeight: 46,
                                      // v29.62: Hover/Fokus sichtbar machen. Ein gewaehlter Tag ist
                                      // schon voll gruen — der wird beim Ueberfahren nur leicht
                                      // dunkler, damit „abwaehlen" ebenfalls als Aktion lesbar ist.
                                      border: `1px solid ${isSel
                                        ? 'var(--dex-green, #86bc25)'
                                        : (dayHoverKey === key ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)')}`,
                                      background: isSel
                                        ? (dayHoverKey === key ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-green, #86bc25)')
                                        : (dayHoverKey === key ? 'rgba(134,188,37,0.10)' : '#fff'),
                                      color: isSel ? '#fff' : (disabled ? 'var(--dex-gray-400)' : 'var(--dex-gray-800, #333)'),
                                      cursor: disabled ? 'not-allowed' : 'pointer',
                                      // v29.72: Der Freischalt-Zustand ist fuer ALLE sichtbar.
                                      // Vorher machte der Organizer/Admin-Bypass die Kachel
                                      // optisch voellig normal — die Organizerin des Events sah
                                      // beim Testen keinerlei Wirkung und meldete die Funktion
                                      // als kaputt, obwohl Teilnehmer sie korrekt gesperrt sahen.
                                      // Jetzt: Anblick gleich fuer alle (gedimmt + „ab TT.MM."),
                                      // nur das KLICKEN laesst der Bypass weiter zu.
                                      opacity: disabled ? 0.55 : (notYetOpen ? 0.65 : 1),
                                      fontWeight: 700, fontSize: '0.82rem',
                                      transition: 'background 120ms ease, border-color 120ms ease, transform 120ms ease',
                                      transform: (dayHoverKey === key && !disabled) ? 'translateY(-1px)' : 'none',
                                    }}
                                  >
                                    <span>{dayNum}</span>
                                    <span style={{ fontSize: '0.6rem', fontWeight: 600, opacity: 0.85 }}>
                                      {/* v30.6: Abgelaufene Tages-Frist nennt das Datum —
                                          „zu" liess offen, WARUM der Tag gesperrt ist und
                                          bis wann man hätte buchen können. */}
                                      {notYetOpen
                                        ? (locale === 'de' ? `ab ${openFromLabel}` : `from ${openFromLabel}`)
                                        : deadlinePassed
                                        ? (() => {
                                            const dl = new Date(effRegDeadline || '');
                                            const dlLabel = isFinite(dl.getTime())
                                              ? dl.toLocaleDateString(locale === 'de' ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit' })
                                              : '';
                                            return dlLabel
                                              ? (locale === 'de' ? `war bis ${dlLabel}` : `until ${dlLabel}`)
                                              : (locale === 'de' ? 'zu' : 'closed');
                                          })()
                                        : isFull
                                        ? (locale === 'de' ? 'voll' : 'full')
                                        : (hasCap && free !== null)
                                        ? (locale === 'de' ? `${free} frei` : `${free} free`)
                                        : '—'}
                                    </span>
                                  </button>
                                );
                              })}
                            </div>
                          </div>
                        );
                      })}
                      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                        {selectedSessions.size === 0
                          ? (locale === 'de' ? 'Noch kein Termin gewählt.' : 'No date picked yet.')
                          : (locale === 'de'
                            ? `${selectedSessions.size} ${selectedSessions.size === 1 ? 'Termin' : 'Termine'} gewählt.`
                            : `${selectedSessions.size} ${selectedSessions.size === 1 ? 'date' : 'dates'} picked.`)}
                      </div>
                    </div>
                  );
                })()}

                {/* Sessions */}
                {!event.subEventCalendar && (
                <div style={{ marginTop: 12, display: 'flex', flexDirection: 'column', gap: 8 }}>
                  <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', fontWeight: 600 }}>{childTermPlural || tEvent('reg.selection.sessions') || 'Sessions'}</div>
                  {childEvents.map(ce => {
                    const meta = sessionMeta[ce.id] || { count: null, wasRegistered: false };
                    const isSel = selectedSessions.has(ce.id);
                    const hasCap = typeof ce.maxParticipants === 'number' && ce.maxParticipants > 0;
                    // v30.62: Unbekannte Belegung sperrt NICHT. „Voll" ist eine
                    // Aussage — sie darf nicht aus einer fehlenden Zahl entstehen.
                    // Die echte Grenze zieht ohnehin serverseitig (reserveSeat).
                    const isSessionFull = hasCap && meta.count !== null && meta.count >= (ce.maxParticipants || 0);
                    // v30.8: effektive Tages-Frist — materialisierte Spalte,
                    // sonst Fallback aus der rollierenden Regel (Spalte kann
                    // bei einem unter Drosselung abgebrochenen Save leer sein).
                    const effRegDeadline = subEventRegDeadline(event, ce);
                    const deadlinePassed = !!(effRegDeadline && new Date(effRegDeadline) < new Date());
                    // v29.28: Organizer/Admins dürfen — wie beim Haupt-Event
                    // (parentRegBlocked) und wie der Wizard es ausdrücklich
                    // verspricht — auch nach der Frist anmelden. Die
                    // Kapazitäts-Sperre bleibt für alle.
                    const deadlineLocked = deadlinePassed && !isOrganizer && !isAdmin;
                    // v29.77: „Anmeldung ab" auch in dieser Liste (s. oben).
                    const seOpenFrom = subOpenFrom(ce.startDate);
                    const seNotYetOpen = !!seOpenFrom && new Date() < seOpenFrom;
                    const seOpenLocked = seNotYetOpen && !isOrganizer && !isAdmin;
                    const disabled = (isSessionFull && !isSel) || (deadlineLocked && !isSel) || (seOpenLocked && !isSel);

                    return (
                      <div key={ce.id} style={{
                        padding: 10, borderRadius: 8,
                        border: `1px solid ${isSel ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                        background: isSel ? 'rgba(134,188,37,0.06)' : '#fff',
                        opacity: seNotYetOpen && !isSel ? 0.75 : 1,
                      }}>
                        <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: disabled ? 'not-allowed' : 'pointer' }}>
                          <input
                            type="checkbox"
                            checked={isSel}
                            disabled={disabled}
                            onChange={e => {
                              if (e.target.checked) {
                                // v29.27: direkt selektieren — Fragen inline in
                                // der Karte (s. Listen-Pfad oben).
                                // v28.97: siehe Kalender — bei „genau eines"
                                // ersetzt die neue Wahl die bisherige.
                                const next = event.subEventSingleChoice ? new Set<string>() : new Set(selectedSessions);
                                next.add(ce.id);
                                setSelectedSessions(next);
                              } else {
                                const next = new Set(selectedSessions);
                                next.delete(ce.id);
                                setSelectedSessions(next);
                                setSessionFieldValues(prev => {
                                  const copy = { ...prev };
                                  delete copy[ce.id];
                                  return copy;
                                });
                              }
                            }}
                            style={{ marginTop: 2 }}
                          />
                          <div style={{ flex: 1 }}>
                            <div style={{ fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                              {ce.title || tEvent('reg.subevents.untitled')}
                              {ce.mandatoryRegistration && (
                                <span style={{ fontSize: '0.68rem', fontWeight: 700, color: '#fff', background: 'var(--dex-orange, #ed8b00)', borderRadius: 999, padding: '2px 8px' }}>
                                  {locale === 'de' ? 'Pflicht' : 'Required'}
                                </span>
                              )}
                            </div>
                          </div>
                        </label>
                        {/* v29.28: Karteninhalt linksbündig (s. Listen-Pfad). */}
                        <div style={{ marginTop: 4 }}>
                            {ce.description && (
                              // v11.97: gleiche Schriftgröße wie der Titel
                              // (Standard-Body). Vorher 0.78rem klein.
                              // v29.27: als sanitisiertes HTML statt rohem Text
                              // (s. subEventDescHtml beim Listen-Pfad oben).
                              <div style={{ color: 'var(--dex-gray-600)', marginTop: 2, wordBreak: 'break-word' }} dangerouslySetInnerHTML={{ __html: subEventDescHtml(ce.description) }} />
                            )}
                            {/* v11.94: gleiches Icon-Layout wie oben (anderer
                                Render-Pfad für Team-Modus). */}
                            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10, marginTop: 4, color: 'var(--dex-gray-600)' }}>
                              {ce.startDate && (
                                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                  {/* v11.97: Icon-Größe an Standard-Body angepasst. */}
                                  <Icon iconName="Calendar" style={{ fontSize: 15, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                                  {formatDate(ce.startDate)}
                                </span>
                              )}
                              {ce.location && (
                                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                  <Icon iconName="POI" style={{ fontSize: 15, color: '#0a3766' }} />
                                  {ce.location}
                                </span>
                              )}
                            </div>
                            <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                              {hasCap && (() => {
                                // v30.62: Ohne belastbare Belegung nur die Kapazität
                                // nennen. „0/80 belegt — 80 frei" wäre eine Zahl, die
                                // aus einem Leseverbot entsteht, nicht aus den Daten.
                                if (meta.count === null) {
                                  return <> · {ce.maxParticipants} {locale === 'de' ? 'Plätze' : 'seats'}</>;
                                }
                                const sessionFree = Math.max(0, (ce.maxParticipants || 0) - (meta.count || 0));
                                return (
                                  <> · <span style={{ color: isSessionFull ? 'var(--dex-red)' : 'inherit', fontWeight: 600 }}>
                                    {/* v19.19: belegt-Zahl bei der Kapazität deckeln (s.o.) —
                                        Überbuchung nie auf der Anmeldeseite anzeigen. */}
                                    {Math.min(meta.count, ce.maxParticipants || 0)}/{ce.maxParticipants} {tEvent('reg.subevents.taken')}
                                  </span>
                                  {!isSessionFull && (
                                    <span style={{ color: 'var(--dex-green-dark)' }}> — {sessionFree} {tEvent('reg.free')}</span>
                                  )}
                                  </>
                                );
                              })()}
                            </div>
                            {deadlinePassed && !isSel && (
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-orange)', marginTop: 2 }}>
                                {tEvent('reg.subevents.deadlinepassed')}
                              </div>
                            )}
                            {/* v29.77: „Anmeldung ab" auch hier ausweisen. */}
                            {seNotYetOpen && !isSel && (
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-orange)', marginTop: 2 }}>
                                {seOpenLocked
                                  ? (locale === 'de'
                                    ? `Anmeldung ab ${seOpenFrom!.toLocaleDateString('de-DE')} möglich.`
                                    : `Registration opens on ${seOpenFrom!.toLocaleDateString('en-GB')}.`)
                                  : (locale === 'de'
                                    ? `Anmeldung öffnet regulär am ${seOpenFrom!.toLocaleDateString('de-DE')} — als Organizer/Admin trotzdem wählbar.`
                                    : `Registration opens on ${seOpenFrom!.toLocaleDateString('en-GB')} — still selectable as organizer/admin.`)}
                              </div>
                            )}
                            {isSessionFull && !isSel && (
                              <div style={{ fontSize: '0.72rem', color: 'var(--dex-red)', marginTop: 2 }}>
                                {tEvent('reg.subevents.sessionfull')}
                              </div>
                            )}
                            {/* v11.10: Hardcoded Sub-Event-Gruppen-Radios entfernt.
                                Sub-Events erben jetzt grundsätzlich
                                preferredStarterType vom Group-Selection-Block
                                oben. Pro-Sub-Event-Gruppe ist konzeptionell
                                Quatsch — die Gruppe gehört zum Teilnehmer
                                (z.B. „Vormittag/Nachmittag"), nicht zur
                                Session. */}
                        </div>
                        {/* v29.27: Fragen inline in der Karte (s. Listen-Pfad). */}
                        {isSel && renderSubEventInlineFields(ce)}
                      </div>
                    );
                  })}
                </div>
                )}

                {isSessionsOnlyMode && selectedSessions.size > 0 && !event.subEventsOnlyMode && (
                  <div style={{
                    marginTop: 12, padding: '8px 10px', borderRadius: 6,
                    background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange)',
                    color: 'var(--dex-orange)', fontSize: '0.78rem',
                  }}>
                    {childTermPlural
                      ? (locale === 'de'
                          ? `Du meldest dich ausschließlich für ${childTermPlural} an — NICHT für das Haupt-Event.`
                          : `You are registering exclusively for ${childTermPlural} — NOT for the main event.`)
                      : (tEvent('reg.selection.sessionsonlyhint') || 'Du meldest dich ausschließlich für Sessions an — NICHT für das Haupt-Event.')}
                  </div>
                )}
              </div>
            )}
            {/* v29.28: Bei Events MIT Sub-Events rendert der Felder-Block
                oben in der Auswahl-Box direkt unter der Haupt-Event-Kachel
                (renderMainFieldsSection) — hier nur noch ohne Sub-Events
                oder bei einer Klammer (dort gelten die Fragen übergreifend
                und es gibt keine Haupt-Event-Kachel). */}
            {(childEvents.length === 0 || !!event.subEventsOnlyMode) && renderMainFieldsSection()}
          </div>
          </CollapsibleSection>
        </div>
  );
};
