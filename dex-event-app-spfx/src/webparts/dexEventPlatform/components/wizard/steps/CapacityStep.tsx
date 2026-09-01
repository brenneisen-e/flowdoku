/* CapacityStep — aus EventCreationPage.tsx ausgelagert (Zeilen 12931-14831 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; einzige
 * Aenderung ist die Anzeige-Bedingung: aus `currentStep === 3` wurde das Prop `visible`.
 * `visible` schaltet display:none statt unmount — Eingaben ueberleben den
 * Schrittwechsel genauso wie vorher. */
import * as React from 'react';
import { SubEventDraft } from '../../wizard/wizardTypes';
import { Icon } from '@fluentui/react/lib/Icon';
import { StepBadge } from '../../wizard/StepBadge';
import { LocationMultiSelect } from '../../wizard/LocationMultiSelect';
import AudiencePicker from '../../AudiencePicker';
import DatePicker from 'react-datepicker';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { InfoTooltip } from '../../InfoTooltip';
import WizardHint from '../../WizardHint';
export interface CapacityStepProps {
  visible: boolean;
  activeCapacityTabIdx: number;
  activeFrom: string;
  assistantsCanSee: boolean;
  audience: string;
  b2runStartblocks: string[];
  berlinLocalToUtcIso: (localStr: string) => string;
  cancelRuleAfter: boolean;
  cancelRuleAmount: number;
  cancelRuleEnabled: boolean;
  cancelRuleUnit: "days" | "hours";
  childTermPlural: string;
  childTermSingular: string;
  durchstarterCapacity: string;
  durchstarterStartblock: string;
  effectiveKlammerDeadline: string;
  errorBorderStyle: (fieldName: string) => React.CSSProperties;
  excludedUsers: string[];
  fieldHasError: (fieldName: string) => boolean;
  filterMode: "AND" | "OR";
  funstarterCapacity: string;
  funstarterStartblock: string;
  hauptGreyoutWrapperStyle: () => React.CSSProperties;
  isDe: boolean;
  isVisOpen: (k: string) => boolean;
  klammerDeadline: string;
  lastDeregisterDate: string;
  locationFilter: string;
  locationOptions: string[];
  maxParticipants: string;
  noCancelAfterDeadline: boolean;
  openRuleDays: number;
  openRuleEnabled: boolean;
  openRuleFixedDate: string;
  openRuleMode: "day" | "week";
  registrationDeadline: string;
  regRuleAmount: number;
  regRuleEnabled: boolean;
  regRuleUnit: "days" | "hours";
  renderHauptGreyoutBanner: () => React.ReactElement | null;
  renderKlammerVisibilityMismatch: () => React.ReactElement | null;
  renderStepIntro: (_bulletsDe: string[], _bulletsEn: string[]) => React.ReactElement | null;
  renderVisibilitySummaryBox: (locList: string[], audienceStr: string, mode: 'AND' | 'OR', excludedCount: number) => React.ReactElement;
  rollingDeadlineIso: (startIso: string, amount: number, unit: 'days' | 'hours', after?: boolean) => string;
  setActiveCapacityTabIdx: React.Dispatch<React.SetStateAction<number>>;
  setActiveFrom: React.Dispatch<React.SetStateAction<string>>;
  setAssistantsCanSee: React.Dispatch<React.SetStateAction<boolean>>;
  setAudience: React.Dispatch<React.SetStateAction<string>>;
  setCancelRuleAfter: React.Dispatch<React.SetStateAction<boolean>>;
  setCancelRuleAmount: React.Dispatch<React.SetStateAction<number>>;
  setCancelRuleEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  setCancelRuleUnit: React.Dispatch<React.SetStateAction<"days" | "hours">>;
  setDurchstarterCapacity: React.Dispatch<React.SetStateAction<string>>;
  setDurchstarterStartblock: React.Dispatch<React.SetStateAction<string>>;
  setExcludedUsers: React.Dispatch<React.SetStateAction<string[]>>;
  setFilterMode: React.Dispatch<React.SetStateAction<"AND" | "OR">>;
  setFunstarterCapacity: React.Dispatch<React.SetStateAction<string>>;
  setFunstarterStartblock: React.Dispatch<React.SetStateAction<string>>;
  setKlammerDeadline: React.Dispatch<React.SetStateAction<string>>;
  setLastDeregisterDate: React.Dispatch<React.SetStateAction<string>>;
  setLocationFilter: React.Dispatch<React.SetStateAction<string>>;
  setMaxParticipants: React.Dispatch<React.SetStateAction<string>>;
  setNoCancelAfterDeadline: React.Dispatch<React.SetStateAction<boolean>>;
  setOpenRuleDays: React.Dispatch<React.SetStateAction<number>>;
  setOpenRuleEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  setOpenRuleFixedDate: React.Dispatch<React.SetStateAction<string>>;
  setOpenRuleMode: React.Dispatch<React.SetStateAction<"day" | "week">>;
  setRegistrationDeadline: React.Dispatch<React.SetStateAction<string>>;
  setRegRuleAmount: React.Dispatch<React.SetStateAction<number>>;
  setRegRuleEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  setRegRuleUnit: React.Dispatch<React.SetStateAction<"days" | "hours">>;
  setSplitDescA: React.Dispatch<React.SetStateAction<string>>;
  setSplitDescB: React.Dispatch<React.SetStateAction<string>>;
  setSplitDisplayOrderReversed: React.Dispatch<React.SetStateAction<boolean>>;
  setSplitHelpText: React.Dispatch<React.SetStateAction<string>>;
  setSplitLabelA: React.Dispatch<React.SetStateAction<string>>;
  setSplitLabelB: React.Dispatch<React.SetStateAction<string>>;
  setSplitSectionTitle: React.Dispatch<React.SetStateAction<string>>;
  setSplitSharedWaitlist: React.Dispatch<React.SetStateAction<boolean>>;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setSubTransfer: React.Dispatch<React.SetStateAction<{ fromIdx: number; groups: string[]; targets: number[]; }>>;
  setUnlimitedParticipants: React.Dispatch<React.SetStateAction<boolean>>;
  setUserCancelAllowed: React.Dispatch<React.SetStateAction<boolean>>;
  setUseSplitCapacities: React.Dispatch<React.SetStateAction<boolean>>;
  setVisAllSubs: React.Dispatch<React.SetStateAction<boolean>>;
  setWaitlistEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  splitDescA: string;
  splitDescB: string;
  splitDisplayOrderReversed: boolean;
  splitHelpText: string;
  splitLabelA: string;
  splitLabelB: string;
  splitSectionTitle: string;
  splitSharedWaitlist: boolean;
  SUB_TRANSFER_GROUPS: { key: string; de: string; en: string; fields: string[]; }[];
  subEvents: SubEventDraft[];
  subEventsOnlyMode: boolean;
  subEventsOptIn: boolean;
  subGroupDiffCount: (srcIdx: number, fields: string[]) => number;
  t: (key: string) => string;
  title: string;
  unlimitedParticipants: boolean;
  userCancelAllowed: boolean;
  useSplitCapacities: boolean;
  visAllSubs: boolean;
  visAllSubsTouchedRef: React.MutableRefObject<boolean>;
  visHeader: (key: string, badge: React.ReactNode, title: React.ReactNode) => React.ReactElement;
  waitlistEnabled: boolean;
  zebraS3Bg: () => string;
}
export const CapacityStep: React.FC<CapacityStepProps> = (p) => {
  const { visible } = p;
  const { activeCapacityTabIdx, activeFrom, assistantsCanSee, audience, b2runStartblocks, berlinLocalToUtcIso, cancelRuleAfter, cancelRuleAmount, cancelRuleEnabled, cancelRuleUnit, childTermPlural, childTermSingular, durchstarterCapacity, durchstarterStartblock, effectiveKlammerDeadline, errorBorderStyle, excludedUsers, fieldHasError, filterMode, funstarterCapacity, funstarterStartblock, hauptGreyoutWrapperStyle, isDe, isVisOpen, klammerDeadline, lastDeregisterDate, locationFilter, locationOptions, maxParticipants, noCancelAfterDeadline, openRuleDays, openRuleEnabled, openRuleFixedDate, openRuleMode, registrationDeadline, regRuleAmount, regRuleEnabled, regRuleUnit, renderHauptGreyoutBanner, renderKlammerVisibilityMismatch, renderStepIntro, renderVisibilitySummaryBox, rollingDeadlineIso, setActiveCapacityTabIdx, setActiveFrom, setAssistantsCanSee, setAudience, setCancelRuleAfter, setCancelRuleAmount, setCancelRuleEnabled, setCancelRuleUnit, setDurchstarterCapacity, setDurchstarterStartblock, setExcludedUsers, setFilterMode, setFunstarterCapacity, setFunstarterStartblock, setKlammerDeadline, setLastDeregisterDate, setLocationFilter, setMaxParticipants, setNoCancelAfterDeadline, setOpenRuleDays, setOpenRuleEnabled, setOpenRuleFixedDate, setOpenRuleMode, setRegistrationDeadline, setRegRuleAmount, setRegRuleEnabled, setRegRuleUnit, setSplitDescA, setSplitDescB, setSplitDisplayOrderReversed, setSplitHelpText, setSplitLabelA, setSplitLabelB, setSplitSectionTitle, setSplitSharedWaitlist, setSubEvents, setSubTransfer, setUnlimitedParticipants, setUserCancelAllowed, setUseSplitCapacities, setVisAllSubs, setWaitlistEnabled, splitDescA, splitDescB, splitDisplayOrderReversed, splitHelpText, splitLabelA, splitLabelB, splitSectionTitle, splitSharedWaitlist, SUB_TRANSFER_GROUPS, subEvents, subEventsOnlyMode, subEventsOptIn, subGroupDiffCount, t, title, unlimitedParticipants, userCancelAllowed, useSplitCapacities, visAllSubs, visAllSubsTouchedRef, visHeader, waitlistEnabled, zebraS3Bg } = p;
  return (
              <div style={{ display: visible ? 'block' : 'none' }}>
              <h2 className="dex-step-head-title">
                {isDe ? 'Schritt 4 — Kapazität & Sichtbarkeit' : 'Step 4 — Capacity & Visibility'}
              </h2>
              <p className="dex-step-head-lead">
                {isDe
                  ? 'Hier legst du fest, wer das Event überhaupt sieht, wie viele Plätze es gibt und bis wann sich Teilnehmer an- bzw. abmelden können.'
                  : 'Here you decide who can see the event in the first place, how many spots there are, and the deadlines for registration and cancellation.'}
              </p>
              {renderStepIntro(
                [
                  'Sichtbarkeit: Standort-Filter und Mailverteiler/User festlegen — wer das Event in der Liste sieht',
                  'Anmelde-Deadline + letzte Abmeldemöglichkeit (vorbefüllt anhand des Event-Datums, jederzeit überschreibbar)',
                  'Maximale Teilnehmerzahl festlegen (oder Unbegrenzt)',
                  'Warteliste aktivieren — voll besetzte Events nehmen weitere Anmeldungen auf, bis ein Platz frei wird',
                  'Optional: Geteilte Kapazität — zwei frei benannte Gruppen mit eigener Platzzahl + eigener oder gemeinsamer Warteliste',
                ],
                [
                  'Visibility: configure location filter + mailing lists/individual users — who sees the event in the list',
                  'Registration deadline + last cancellation date (pre-filled from the event date, always overridable)',
                  'Set the maximum number of attendees (or Unlimited)',
                  'Enable waitlist — full events accept new registrations and promote them once a spot frees up',
                  'Optional: split capacity — two freely-named groups with own seat count + own or shared waitlist',
                ]
              )}

              {/* v15.0: pro-Sub-Event-Tabs für Kapazität. Tab 0 = Haupt-
                  Event (komplette Sichtbarkeit/Deadlines/MaxParticipants/
                  Split-UI). Tabs N>0 = schlanke MaxParticipants-only-UI pro
                  Sub-Event mit Inheritance-Toggle. Sichtbarkeit, Filter,
                  Deadlines, Split-Capacity bleiben Top-Level — pro Sub-Event
                  ist nur die Platzzahl relevant. */}
              {/* v28.78: Der Scope-Umschalter steht jetzt global unter der
                  Schritt-Leiste (renderGlobalScopeBar) — nicht mehr je Schritt. */}

              {activeCapacityTabIdx > 0 && (() => {
                const seIdx = activeCapacityTabIdx - 1;
                const se = subEvents[seIdx];
                if (!se) return null;
                const updateSub = (patch: Partial<SubEventDraft>): void => {
                  setSubEvents(prev => prev.map((x, i) => i === seIdx ? { ...x, ...patch } : x));
                };
                const seLocationFilterList = (se.locationFilter || '').split(',').map(s => s.trim()).filter(Boolean);
                return (
                  <div>
                    {/* v15.3: „Vom Hauptevent kopieren"-Button. Übernimmt
                        Kapazitäts-/Sichtbarkeits-/Deadline-/Filter-Werte
                        vom Hauptevent als Startwerte für dieses Sub-Event. */}
                    <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, flexWrap: 'wrap', marginBottom: 12 }}>
                      <button
                        type="button"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.8rem', padding: '6px 14px' }}
                        onClick={() => updateSub({
                          maxParticipants: parseInt(maxParticipants, 10) || 0,
                          registrationDeadline: registrationDeadline ? berlinLocalToUtcIso(registrationDeadline) : '',
                          lastDeregisterDate: lastDeregisterDate ? berlinLocalToUtcIso(lastDeregisterDate) : '',
                          locationFilter: locationFilter,
                          audience: audience,
                          filterMode: filterMode,
                          waitlistEnabled: waitlistEnabled,
                        })}
                        title={isDe
                          ? 'Übernimmt Teilnehmerzahl, Deadlines, Sichtbarkeit und Warteliste vom Hauptevent als Startwerte'
                          : 'Copies capacity, deadlines, visibility and waitlist from the main event as starting values'}
                      >
                        {isDe ? 'Vom Hauptevent kopieren' : 'Copy from main event'}
                      </button>
                      {/* v28.74: Gegenstück — Werte dieses Sub-Events auf die
                          anderen übertragen (was und wohin wählt der Organizer). */}
                      {subEvents.length > 1 && (
                        <button
                          type="button"
                          className="btn btn-secondary"
                          style={{ fontSize: '0.8rem', padding: '6px 14px' }}
                          onClick={() => setSubTransfer({
                            fromIdx: seIdx,
                            groups: SUB_TRANSFER_GROUPS.filter(g => g.key !== 'times' && subGroupDiffCount(seIdx, g.fields) > 0).map(g => g.key),
                            targets: subEvents.map((_, i) => i).filter(i => i !== seIdx),
                          })}
                          title={isDe
                            ? 'Überträgt ausgewählte Einstellungen dieses Sub-Events auf andere Sub-Events'
                            : 'Transfers selected settings of this sub-event to other sub-events'}
                        >
                          {isDe ? 'Einstellungen auf andere übertragen' : 'Transfer settings to others'}
                        </button>
                      )}
                    </div>

                    {/* v28.74: Abweichungs-Hinweis. Meldet von sich aus, welche
                        Einstellungen bei anderen Sub-Events anders stehen —
                        genau der Fall „Standortfilter auf einigen Tagen Berlin,
                        auf anderen leer", der sonst erst in der Zusammenfassung
                        (oder gar nicht) auffällt. */}
                    {subEvents.length > 1 && (() => {
                      const diffs = SUB_TRANSFER_GROUPS
                        .filter(g => g.key !== 'times')
                        // v30.2: Bei aktiver rollierender Regel (v29.76) hat
                        // JEDER Termin planmaessig eine ANDERE Frist — „bei 18
                        // von 18 anders" ist dann der Soll-Zustand, keine
                        // meldenswerte Abweichung.
                        .filter(g => !(g.key === 'regDeadline' && regRuleEnabled) && !(g.key === 'deregDeadline' && cancelRuleEnabled))
                        .map(g => ({ g, n: subGroupDiffCount(seIdx, g.fields) }))
                        .filter(x => x.n > 0);
                      if (diffs.length === 0) return null;
                      const others = subEvents.length - 1;
                      return (
                        <div style={{
                          margin: '0 0 14px', padding: '10px 12px', borderRadius: 8,
                          background: '#fff8e6', border: '1px solid #e0b34d', color: '#7a5a12',
                          fontSize: '0.8rem', lineHeight: 1.55,
                        }}>
                          <div style={{ fontWeight: 700, marginBottom: 4 }}>
                            {isDe ? 'Diese Einstellungen weichen von den anderen ab' : 'These settings differ from the others'}
                          </div>
                          <ul style={{ margin: '0 0 8px', paddingLeft: 18 }}>
                            {diffs.map(x => (
                              <li key={x.g.key}>
                                {isDe ? x.g.de.replace(/ —.*$/, '') : x.g.en.replace(/ —.*$/, '')}: {isDe
                                  ? <>bei <strong>{x.n} von {others}</strong> anderen {others === 1 ? 'Sub-Event' : 'Sub-Events'} anders</>
                                  : <>differs in <strong>{x.n} of {others}</strong> other sub-event{others === 1 ? '' : 's'}</>}
                              </li>
                            ))}
                          </ul>
                          <div style={{ marginBottom: 8 }}>
                            {isDe
                              ? 'Das kann so gewollt sein. Wenn nicht, übernimm die Werte dieses Sub-Events für die anderen:'
                              : 'That may be intentional. If not, apply this sub-event’s values to the others:'}
                          </div>
                          <button
                            type="button"
                            className="btn btn-primary"
                            style={{ fontSize: '0.8rem', padding: '6px 14px' }}
                            onClick={() => setSubTransfer({
                              fromIdx: seIdx,
                              groups: diffs.map(x => x.g.key),
                              targets: subEvents.map((_, i) => i).filter(i => i !== seIdx),
                            })}
                          >
                            {isDe ? 'Auf die anderen übertragen…' : 'Transfer to the others…'}
                          </button>
                        </div>
                      );
                    })()}

                    {/* v15.6: Sichtbarkeits-Sektion analog Hauptevent —
                        Header mit Auge-Icon plus erklärender Lead-Text. */}
                    <div style={{
                      paddingBottom: 12, marginBottom: 16,
                      borderBottom: '2px solid var(--dex-gray-100)',
                    }}>
                      <h3 style={{ margin: '0 0 6px', fontSize: '1.05rem', display: 'flex', alignItems: 'center', gap: 8 }}>
                        <Icon iconName="Hide3" style={{ fontSize: 18, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                        {isDe ? 'Sichtbarkeit' : 'Visibility'}
                      </h3>
                      <p style={{ margin: 0, fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                        {isDe
                          ? <>Dieses Sub-Event übernimmt <strong>standardmäßig die Sichtbarkeit {subEventsOnlyMode ? 'der Klammer' : 'des Hauptevents'}</strong> (Standortfilter + Mailverteiler sind unten vorbefüllt). Du kannst den Empfängerkreis hier aber <strong>jederzeit anpassen</strong> — oder mit „Vom Hauptevent kopieren“ oben erneut übernehmen.</>
                          : <>This sub-event <strong>inherits the visibility {subEventsOnlyMode ? 'of the bracket' : 'of the main event'}</strong> by default (location filter + mailing lists are pre-filled below). You can <strong>change the audience here at any time</strong> — or re-apply it with “Copy from main event” above.</>}
                      </p>
                      {/* v26.88: Live-Zusammenfassung wandert in die
                          AudiencePicker-Prüfzeile (summarySlot). */}
                    </div>

                    {/* v29.75: Solange „Sichtbarkeit gilt für alle Sub-Events"
                        auf der Klammer gesetzt ist, wird die Sichtbarkeit hier
                        nur ANGEZEIGT — Eingaben würden beim nächsten Klammer-
                        Edit stillschweigend überschrieben (Spiegel-Effect),
                        deshalb sperren statt verlieren lassen. */}
                    {visAllSubs && (
                      <div style={{
                        margin: '0 0 12px', padding: '10px 12px', borderRadius: 8,
                        background: 'rgba(134,188,37,0.07)', border: '1px solid var(--dex-green, #86bc25)',
                        fontSize: '0.8rem', color: 'var(--dex-gray-700)', lineHeight: 1.5,
                      }}>
                        {isDe
                          ? <>Die Sichtbarkeit wird von der <strong>{subEventsOnlyMode ? 'Klammer' : 'Hauptevent-Ebene'}</strong> vorgegeben (&bdquo;Sichtbarkeit gilt für alle {childTermPlural || 'Sub-Events'}&ldquo;). Zum Abweichen den Haken dort entfernen.</>
                          : <>Visibility is governed by the <strong>{subEventsOnlyMode ? 'bracket' : 'main event'}</strong> (“visibility applies to all {childTermPlural || 'sub-events'}”). Uncheck the box there to deviate.</>}
                      </div>
                    )}
                    <div style={visAllSubs ? { opacity: 0.55, pointerEvents: 'none' as const, userSelect: 'none' as const } : undefined}>
                    <div className="form-group" style={{ padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <StepBadge n={18} />
                        {isDe ? 'Standortfilter' : 'Location filter'}
                      </label>
                      <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: -4, marginBottom: 12, lineHeight: 1.5 }}>
                        {isDe
                          ? <>Wählst du hier einen oder mehrere Standorte aus, sehen <strong>nur Mitarbeiter mit diesen Standorten</strong> dieses Sub-Event. Leer = für alle sichtbar.</>
                          : <>If you pick one or more locations here, <strong>only employees from those locations</strong> will see this sub-event. Empty = visible to everyone.</>}
                      </p>
                      <LocationMultiSelect
                        options={locationOptions}
                        selected={seLocationFilterList}
                        onChange={list => updateSub({ locationFilter: list.join(', ') })}
                        isDe={isDe}
                      />
                    </div>

                    {/* v19.x: Mailverteiler-Auswahl + „Sichtbarkeit prüfen" +
                        „Personen ausschließen" identisch zum Hauptevent über
                        <AudiencePicker>. Die Filterverknüpfung (ODER/UND) wird
                        als middleSlot zwischen Mailverteiler-Karte und Prüfen-
                        Zeile eingeschoben — gleiche Reihenfolge wie im Hauptevent.
                        v22.10: Die Ausschluss-Liste wird jetzt AUCH pro Sub-Event
                        persistiert (Spalte ExcludedUsers), vorher nur intern. */}
                    <AudiencePicker
                      value={se.audience || ''}
                      onChange={v => updateSub({ audience: v })}
                      locationFilter={se.locationFilter || ''}
                      filterMode={(se.filterMode as 'AND' | 'OR') || 'OR'}
                      isDe={isDe}
                      // v22.10: Ausschluss-Liste pro Sub-Event jetzt persistiert
                      // (vorher interner Picker-State → ging beim Reload verloren).
                      // Der Picker liefert einen Updater (prev => next), den wir auf
                      // den aktuellen Stand des Sub-Events anwenden.
                      excludedUsers={se.excludedUsers || []}
                      onExcludedUsersChange={updater => setSubEvents(prev => prev.map((x, i) => i === seIdx ? { ...x, excludedUsers: updater(x.excludedUsers || []) } : x))}
                      stepBadge={<StepBadge n={19} />}
                      cardBgPrimary={zebraS3Bg()}
                      summarySlot={renderVisibilitySummaryBox(
                        seLocationFilterList,
                        se.audience || '',
                        (se.filterMode as 'AND' | 'OR') || 'OR',
                        (se.excludedUsers || []).length
                      )}
                      middleSlot={(seLocationFilterList.length > 0 && (se.audience || '').trim().length > 0) ? (
                        <div className="form-group" style={{ padding: '16px 20px 16px 30px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                          <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                            <StepBadge n={20} />
                            {isDe ? 'Filterverknüpfung' : 'Filter combination'}
                          </label>
                          <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', marginTop: -4, marginBottom: 12, lineHeight: 1.55 }}>
                            {isDe
                              ? <>Beide Filter sind gesetzt — bestimmt, ob für eine Person <strong>einer</strong> der Filter (ODER) oder <strong>beide</strong> (UND) zutreffen müssen, damit das Sub-Event in ihrer Liste auftaucht.</>
                              : <>Both filters are set — defines whether a person needs to match <strong>either</strong> filter (OR) or <strong>both</strong> filters (AND) for the sub-event to appear in their list.</>}
                          </p>
                          <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap' }}>
                            <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '0.9rem', cursor: 'pointer' }}>
                              <input
                                type="radio"
                                name={`subFilterMode-${se.id}`}
                                checked={(se.filterMode || 'OR') === 'OR'}
                                onChange={() => updateSub({ filterMode: 'OR' })}
                              />
                              <strong>{isDe ? 'ODER' : 'OR'}</strong>
                              <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>– {isDe ? 'Einer der Filter reicht' : 'one filter is enough'}</span>
                            </label>
                            <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '0.9rem', cursor: 'pointer' }}>
                              <input
                                type="radio"
                                name={`subFilterMode-${se.id}`}
                                checked={se.filterMode === 'AND'}
                                onChange={() => updateSub({ filterMode: 'AND' })}
                              />
                              <strong>{isDe ? 'UND' : 'AND'}</strong>
                              <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>– {isDe ? 'Beides muss zutreffen' : 'both must match'}</span>
                            </label>
                          </div>
                        </div>
                      ) : null}
                      cardBgSecondary={((se.locationFilter || '').trim().length > 0 || (se.audience || '').trim().length > 0) ? zebraS3Bg() : '#fff'}
                    />
                    </div>{/* v29.75: Ende Sichtbarkeits-Sperre bei visAllSubs */}

                    {/* Deadlines: zwei DatePicker nebeneinander, gleicher Look
                        wie im Hauptevent. */}
                    <div className="form-group" style={{ padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <StepBadge n={(seLocationFilterList.length > 0 && (se.audience || '').trim().length > 0) ? 21 : 20} />
                        {isDe ? 'Anmelde- und Abmeldefristen' : 'Registration & cancellation deadlines'}
                      </label>
                      <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: -4, marginBottom: 12, lineHeight: 1.5 }}>
                        {/* v28.20: Im Klammer-Modus das Zusammenspiel mit der
                            Klammer-Frist erklären — abweichende (frühere)
                            Sub-Fristen sind ok, spätere wirken nicht. */}
                        {subEventsOnlyMode ? (
                          klammerDeadline ? (
                            isDe
                              ? <>Frei pro {childTermSingular || 'Sub-Event'} setzbar — z.B. ein <strong>früherer</strong> Anmeldeschluss nur für dieses {childTermSingular || 'Sub-Event'}. Die Klammer-Frist (<strong>{new Date(klammerDeadline).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })}</strong>) schließt das gesamte Event — eine spätere Frist hier hat daher keine Wirkung.</>
                              : <>Settable per {childTermSingular || 'sub-event'} — e.g. an <strong>earlier</strong> cutoff just for this {childTermSingular || 'sub-event'}. The bracket deadline (<strong>{new Date(klammerDeadline).toLocaleString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })}</strong>) closes the entire event — a later deadline here has no effect.</>
                          ) : (
                            isDe
                              ? <>Frei pro {childTermSingular || 'Sub-Event'} setzbar. Ohne Klammer-Frist bleibt das Gesamt-Event offen, solange mindestens ein {childTermSingular || 'Sub-Event'} offen ist; leer = offen bis zum Ende dieses {childTermSingular || 'Sub-Events'}.</>
                              : <>Settable per {childTermSingular || 'sub-event'}. Without a bracket deadline the overall event stays open as long as at least one {childTermSingular || 'sub-event'} is open; empty = open until this {childTermSingular || 'sub-event'} ends.</>
                          )
                        ) : (isDe
                          ? <>Frei pro Sub-Event setzbar. Leer lassen → die Fristen des Hauptevents gelten.</>
                          : <>Settable per sub-event. Leave empty → the main event’s deadlines apply.</>)}
                      </p>
                      {/* v29.76: Rollierende Klammer-Regel berechnete die Felder
                          und sperrte sie. v30.6: Die Regel ist nur noch die
                          VORBELEGUNG — hier darf pro Termin ueberschrieben
                          werden (z.B. ein frueherer Schluss nur fuer diesen
                          Tag). Der Effect fuellt nur noch leere Felder;
                          „Auf Regel zuruecksetzen" holt den Regel-Wert zurueck. */}
                      {(regRuleEnabled || cancelRuleEnabled) && (
                        <p style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', marginTop: -6, marginBottom: 10, lineHeight: 1.5 }}>
                          {isDe
                            ? <>Rollierende Regel der Klammer aktiv (einstellbar in Schritt 4): {regRuleEnabled && cancelRuleEnabled ? 'beide Fristen werden' : (regRuleEnabled ? <>&bdquo;Anmeldung bis&ldquo; wird</> : <>&bdquo;Abmeldung bis&ldquo; wird</>)} aus dem Termin-Datum <strong>vorbelegt</strong>. Du kannst sie hier für diesen einzelnen Termin überschreiben — dein Wert bleibt dann stehen. Änderst du später die Regel selbst, werden wieder <strong>alle</strong> Termine neu berechnet.</>
                            : <>Rolling bracket rule active (configured in step 4): {regRuleEnabled && cancelRuleEnabled ? 'both deadlines are' : 'this deadline is'} <strong>pre-filled</strong> from the date. You can override it here for this single date — your value then sticks. Changing the rule itself recomputes <strong>all</strong> dates again.</>}
                        </p>
                      )}
                      <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                        <div className="form-group" style={{ marginBottom: 0 }}>
                          {/* v29.75: gleicher Wortlaut wie auf der Klammer. */}
                          <label className="form-label">{isDe ? 'Anmeldung bis' : 'Registration until'}</label>
                          <DatePicker
                            selected={se.registrationDeadline ? new Date(se.registrationDeadline) : null}
                            onChange={(date: Date | null) => {
                              if (!date) { updateSub({ registrationDeadline: '' }); return; }
                              const local = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
                              updateSub({ registrationDeadline: berlinLocalToUtcIso(local) });
                            }}
                            showTimeSelect
                            timeFormat="HH:mm"
                            timeIntervals={15}
                            timeCaption="Uhrzeit"
                            dateFormat="dd.MM.yyyy, HH:mm"
                            locale="de"
                            placeholderText={isDe ? 'Anmelde-Deadline' : 'Registration deadline'}
                            className="form-input"
                            wrapperClassName="dex-datepicker-wrapper"
                            calendarClassName="dex-datepicker-calendar"
                            popperPlacement="bottom-start"
                            isClearable
                            autoComplete="off"
                          />
                          {/* v30.6: Abweichung von der Regel sichtbar machen +
                              Ruecksetz-Knopf (holt den berechneten Wert zurueck). */}
                          {regRuleEnabled && regRuleAmount > 0 && (() => {
                            const ruleIso = rollingDeadlineIso(se.startDate || '', regRuleAmount, regRuleUnit);
                            if (!ruleIso) return null;
                            const cur = new Date(se.registrationDeadline || '').getTime();
                            const rt = new Date(ruleIso).getTime();
                            const matches = isFinite(cur) && Math.abs(cur - rt) < 60000;
                            return matches ? (
                              <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', margin: '4px 0 0' }}>
                                {isDe ? 'Entspricht der rollierenden Regel.' : 'Matches the rolling rule.'}
                              </p>
                            ) : (
                              <p style={{ fontSize: '0.72rem', color: '#b86700', margin: '4px 0 0' }}>
                                {isDe
                                  ? <>Manuell überschrieben — Regel wäre {new Date(ruleIso).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })}. </>
                                  : <>Manually overridden — rule would be {new Date(ruleIso).toLocaleString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })}. </>}
                                <button
                                  type="button"
                                  onClick={() => updateSub({ registrationDeadline: ruleIso })}
                                  style={{ background: 'none', border: 'none', padding: 0, color: 'var(--dex-blue, #0076a8)', cursor: 'pointer', fontSize: '0.72rem', fontWeight: 700, textDecoration: 'underline' }}
                                >
                                  {isDe ? 'Auf Regel zurücksetzen' : 'Reset to rule'}
                                </button>
                              </p>
                            );
                          })()}
                        </div>
                        <div className="form-group" style={{ marginBottom: 0 }}>
                          <label className="form-label">{isDe ? 'Abmeldung bis' : 'Cancellation until'}</label>
                          <DatePicker
                            selected={se.lastDeregisterDate ? new Date(se.lastDeregisterDate) : null}
                            onChange={(date: Date | null) => {
                              if (!date) { updateSub({ lastDeregisterDate: '' }); return; }
                              const local = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
                              updateSub({ lastDeregisterDate: berlinLocalToUtcIso(local) });
                            }}
                            showTimeSelect
                            timeFormat="HH:mm"
                            timeIntervals={15}
                            timeCaption="Uhrzeit"
                            dateFormat="dd.MM.yyyy, HH:mm"
                            locale="de"
                            placeholderText={isDe ? 'Abmeldefrist' : 'Last cancellation'}
                            className="form-input"
                            wrapperClassName="dex-datepicker-wrapper"
                            calendarClassName="dex-datepicker-calendar"
                            popperPlacement="bottom-start"
                            isClearable
                            autoComplete="off"
                          />
                          {cancelRuleEnabled && cancelRuleAmount > 0 && userCancelAllowed && (() => {
                            const ruleIso = rollingDeadlineIso(se.startDate || '', cancelRuleAmount, cancelRuleUnit, cancelRuleAfter);
                            if (!ruleIso) return null;
                            const cur = new Date(se.lastDeregisterDate || '').getTime();
                            const rt = new Date(ruleIso).getTime();
                            const matches = isFinite(cur) && Math.abs(cur - rt) < 60000;
                            return matches ? (
                              <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', margin: '4px 0 0' }}>
                                {isDe ? 'Entspricht der rollierenden Regel.' : 'Matches the rolling rule.'}
                              </p>
                            ) : (
                              <p style={{ fontSize: '0.72rem', color: '#b86700', margin: '4px 0 0' }}>
                                {isDe
                                  ? <>Manuell überschrieben — Regel wäre {new Date(ruleIso).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })}. </>
                                  : <>Manually overridden — rule would be {new Date(ruleIso).toLocaleString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })}. </>}
                                <button
                                  type="button"
                                  onClick={() => updateSub({ lastDeregisterDate: ruleIso })}
                                  style={{ background: 'none', border: 'none', padding: 0, color: 'var(--dex-blue, #0076a8)', cursor: 'pointer', fontSize: '0.72rem', fontWeight: 700, textDecoration: 'underline' }}
                                >
                                  {isDe ? 'Auf Regel zurücksetzen' : 'Reset to rule'}
                                </button>
                              </p>
                            );
                          })()}
                        </div>
                      </div>
                    </div>

                    {/* Teilnehmerzahl & Warteliste — analog Hauptevent als
                        eine kombinierte Card mit „Unbegrenzt"-Toggle und
                        Warteliste-Toggle. Split-Capacity bleibt Hauptevent-only
                        (Scope-Eingrenzung, siehe v15.6 Refactor-Plan). */}
                    <div className="form-group" style={{ padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <StepBadge n={(seLocationFilterList.length > 0 && (se.audience || '').trim().length > 0) ? 22 : 21} />
                        {isDe ? 'Teilnehmerzahl & Warteliste' : 'Capacity & waitlist'}
                      </label>
                      <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                        <div className="form-group" style={{ marginBottom: 0 }}>
                          <label className="form-label">
                            {isDe ? 'Max. Teilnehmer (0 = unbegrenzt)' : 'Max. attendees (0 = unlimited)'}
                          </label>
                          <div className="toggle-wrapper" style={{ marginTop: 4, marginBottom: 8 }}>
                            <label className="toggle">
                              <input
                                type="checkbox"
                                checked={(se.maxParticipants || 0) === 0}
                                onChange={e => {
                                  if (e.target.checked) {
                                    updateSub({ maxParticipants: 0, waitlistEnabled: false });
                                  } else {
                                    updateSub({ maxParticipants: 50 });
                                  }
                                }}
                              />
                              <span className="toggle-slider" />
                            </label>
                            <span style={{ fontSize: '0.9rem' }}>
                              {(se.maxParticipants || 0) === 0
                                ? (isDe ? 'Unbegrenzt' : 'Unlimited')
                                : (`${se.maxParticipants} ${isDe ? 'Plätze' : 'seats'}`)}
                            </span>
                          </div>
                          {(se.maxParticipants || 0) > 0 && (
                            <input
                              type="number"
                              min={0}
                              className="form-input"
                              value={se.maxParticipants || 0}
                              onChange={e => {
                                const v = parseInt(e.target.value, 10) || 0;
                                updateSub({ maxParticipants: v });
                              }}
                              placeholder={isDe ? 'Anzahl' : 'Count'}
                            />
                          )}
                        </div>
                        {(se.maxParticipants || 0) > 0 && (
                          <div className="form-group" style={{ marginBottom: 0 }}>
                            <label className="form-label">{t('create.waitlist')}</label>
                            <div className="toggle-wrapper" style={{ marginTop: 8 }}>
                              <label className="toggle">
                                <input
                                  type="checkbox"
                                  checked={typeof se.waitlistEnabled === 'boolean' ? se.waitlistEnabled : true}
                                  onChange={e => updateSub({ waitlistEnabled: e.target.checked })}
                                />
                                <span className="toggle-slider" />
                              </label>
                              <span style={{ fontSize: '0.9rem' }}>
                                {(typeof se.waitlistEnabled === 'boolean' ? se.waitlistEnabled : true) ? t('create.enabled') : t('create.disabled')}
                              </span>
                            </div>
                          </div>
                        )}
                      </div>
                      <p style={{ margin: '12px 0 0', fontSize: '0.75rem', color: 'var(--dex-gray-500)', lineHeight: 1.5 }}>
                        {isDe
                          ? <em>Hinweis: <strong>Geteilte Kapazität</strong> (zwei Gruppen mit eigener Platzzahl) ist aktuell nur auf Hauptevent-Ebene möglich — Sub-Events nutzen die einfache Gesamtkapazität.</em>
                          : <em>Note: <strong>Split capacity</strong> (two groups with separate seat counts) is currently main-event-only — sub-events use the simple total capacity.</em>}
                      </p>
                    </div>
                  </div>
                );
              })()}

              <div style={{ display: activeCapacityTabIdx === 0 ? 'block' : 'none' }}>
              {renderHauptGreyoutBanner()}

              {/* v19.27: Sichtbarkeit (Standortfilter + Mailverteiler) NICHT mehr
                  im Greyout-Wrapper — sie bleibt im „Nur Sub-Events"-Modus
                  (Klammer) editierbar, weil sie steuert, WER das ganze Event
                  sieht. Der Greyout-Wrapper startet erst ab den Fristen weiter
                  unten (die sind für die nicht-buchbare Klammer wirklich ohne
                  Wirkung). */}
              {/* v9.24: Sichtbarkeits-Steuerungen aus Step 0 hierher verschoben.
                  Die Frage 'wer darf das Event sehen' passt logisch zu Kapazität/Fristen
                  als 'Wer-Wann-Wieviel' und entlastet Step 0 (Grundlagen). */}
              {/* Zwischenüberschrift: alle Sichtbarkeits-Steuerungen
                  (Standortfilter + Mailverteiler/einzelne User) gruppieren,
                  damit der Organizer auf einen Blick versteht, dass es hier
                  um die Frage geht: wer darf das Event überhaupt sehen. */}
              <div style={{
                paddingBottom: 12, marginBottom: 16,
                borderBottom: '2px solid var(--dex-gray-100)',
              }}>
                <h3 style={{ margin: '0 0 6px', fontSize: '1.05rem', display: 'flex', alignItems: 'center', gap: 8 }}>
                  <Icon iconName="Hide3" style={{ fontSize: 18, color: 'var(--dex-green-dark, #4a7c1f)' }} />
                  {isDe ? 'Sichtbarkeit' : 'Visibility'}
                </h3>
                <p style={{ margin: 0, fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
                  {isDe ? (
                    <>
                      Dieses Event ist standardmäßig für <strong>alle Mitarbeiter von Deloitte Deutschland</strong>{' '}
                      sichtbar und buchbar. Über die folgenden beiden Bereiche — Standortfilter sowie
                      Mailverteiler / einzelne User — kannst du den Empfängerkreis gezielt einschränken.
                      Mitarbeiter außerhalb der definierten Auswahl sehen das Event nicht in ihrer
                      Übersicht und können sich entsprechend nicht anmelden.
                    </>
                  ) : (
                    <>
                      By default, this event is visible and bookable for <strong>all Deloitte Germany employees</strong>.
                      You can narrow down the audience using the two sections below — the location filter
                      and mailing lists / individual users. Employees outside the selected scope will
                      not see the event in their overview and cannot register for it.
                    </>
                  )}
                </p>
              </div>

              {/* v26.89: Live-Zusammenfassung „Aktuell eingestellt …" steht jetzt
                  GANZ OBEN — direkt über dem Standortfilter (Schritt 13), damit man
                  den aktuellen Sichtbarkeits-Stand sofort sieht. */}
              {renderVisibilitySummaryBox(
                locationFilter.split(',').map(s => s.trim()).filter(Boolean),
                audience,
                filterMode,
                (excludedUsers || []).length
              )}
              {/* v28.76: Widerspruch Klammer ↔ Sub-Events benennen (s.o.). */}
              {renderKlammerVisibilityMismatch()}

              <div className="form-group" style={{ padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                {visHeader('vis_locfilter', <StepBadge n={18} />, isDe ? 'Standortfilter' : 'Location filter')}
                {isVisOpen('vis_locfilter') && (<>
                <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: -4, marginBottom: 12, lineHeight: 1.5 }}>
                  {isDe ? (
                    <>
                      Wählst du hier einen oder mehrere Standorte aus, sehen <strong>nur Mitarbeiter mit diesen Standorten</strong> das Event.<br />
                      <em>Beispiel: &bdquo;Köln&ldquo; und &bdquo;Düsseldorf&ldquo; → Nur Mitarbeiter mit diesen Standorten sehen das Event. Alle anderen sehen es nicht.</em>
                    </>
                  ) : (
                    <>
                      If you pick one or more locations here, <strong>only employees from those locations</strong> will see the event.<br />
                      <em>Example: &bdquo;Cologne&ldquo; and &bdquo;Düsseldorf&ldquo; → only employees with one of these locations will see the event. Everyone else will not.</em>
                    </>
                  )}
                </p>
                {/* Multi-Select-Dropdown — kompakter als die alten Pillen,
                    erlaubt Suche + Mehrfachauswahl. Aktuelle Auswahl wird
                    direkt im Trigger-Button als Chip-Liste angezeigt. */}
                <LocationMultiSelect
                  options={locationOptions}
                  selected={locationFilter.split(',').map(s => s.trim()).filter(Boolean)}
                  onChange={list => setLocationFilter(list.join(', '))}
                  isDe={isDe}
                />
                {!locationFilter && (
                  <p style={{ fontSize: '0.75rem', color: 'var(--dex-green)', marginTop: 8 }}>
                    {isDe
                      ? 'Kein Standort ausgewählt → Event ist für alle sichtbar.'
                      : 'No location selected → event is visible to everyone.'}
                  </p>
                )}
                </>)}
              </div>

              {/* v19.x: Mailverteiler-Auswahl + Sichtbarkeit prüfen + Personen
                  ausschließen sind in <AudiencePicker> ausgelagert, damit
                  Hauptevent und Sub-Events exakt dieselbe UI nutzen. Die
                  Filterverknüpfung (ODER/UND) wird als middleSlot zwischen
                  Mailverteiler-Karte und Prüfen-Zeile eingeschoben, damit die
                  Reihenfolge der Sektion unverändert bleibt. Die zebraS3Bg()-
                  Aufrufe stehen in derselben Reihenfolge wie zuvor (Mailverteiler,
                  Filterverknüpfung, Prüfen-Zeile), damit die Zebra-Alternation der
                  nachfolgenden Karten gleich bleibt. */}
              <AudiencePicker
                value={audience}
                onChange={setAudience}
                locationFilter={locationFilter}
                filterMode={filterMode}
                isDe={isDe}
                excludedUsers={excludedUsers}
                onExcludedUsersChange={setExcludedUsers}
                headerSlot={visHeader('vis_audience', <StepBadge n={19} />, isDe ? 'Mailverteiler / einzelne User' : 'Mailing lists / individual users')}
                bodyOpen={isVisOpen('vis_audience')}
                cardBgPrimary={zebraS3Bg()}
                visibilityTabs={subEvents.length > 0 ? [
                  { id: 'main', title: subEventsOnlyMode ? (isDe ? 'Klammer' : 'Bracket') : (isDe ? 'Hauptevent' : 'Main event'), locationFilter, audience, filterMode },
                  ...subEvents.map(s => ({ id: s.id, title: (shortSubEventTitle(s.title, title) || (isDe ? 'Sub-Event' : 'Sub-event')).trim(), locationFilter: s.locationFilter || '', audience: s.audience || '', filterMode: (s.filterMode || 'AND') as 'AND' | 'OR' })),
                ] : undefined}
                middleSlot={(locationFilter && audience) ? (
                  /* Filterverknüpfung: nur sichtbar wenn beide Bereiche
                     (Standortfilter + Mailverteiler) Werte haben — sonst gibt
                     es nichts zu kombinieren. */
                  <div className="form-group" style={{ padding: '16px 20px 16px 30px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                    <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                      <StepBadge n={20} />
                      {isDe ? 'Filterverknüpfung' : 'Filter combination'}
                    </label>
                    <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', marginTop: -4, marginBottom: 12, lineHeight: 1.6 }}>
                      {isDe ? (
                        <>
                          <p style={{ margin: '0 0 8px' }}>
                            Bestimmt, wie der <strong>Standortfilter</strong> und der <strong>Mailverteiler / einzelne User</strong> miteinander kombiniert werden — also welche Bedingungen für eine Person erfüllt sein müssen, damit sie das Event in ihrer Liste sieht.
                          </p>
                          <ul style={{ margin: '0 0 8px 18px', padding: 0 }}>
                            <li style={{ marginBottom: 6 }}>
                              <strong>ODER (Default):</strong> <em>einer der beiden Filter reicht.</em> Beispiel: <strong>Standort = Köln</strong>, <strong>Verteiler = SAPALL</strong> → jede Person, die <strong>in Köln</strong> sitzt <strong>ODER</strong> in <strong>SAPALL</strong> ist, sieht das Event. Praktisch wenn du bewusst einen <strong>breiten Empfängerkreis</strong> willst (z.B. Standort-Mitarbeiter <strong>plus</strong> Fachgruppe).
                            </li>
                            <li>
                              <strong>UND:</strong> <em>beide Filter müssen zutreffen.</em> Beispiel: <strong>Standort = Köln</strong>, <strong>Verteiler = SAPALL</strong> → nur wer <strong>in Köln</strong> sitzt <strong>UND</strong> in <strong>SAPALL</strong> ist, sieht das Event. Praktisch wenn du den Empfängerkreis <strong>strikt eingrenzen</strong> willst (z.B. nur die SAP-Kollegen <strong>am Standort Köln</strong>).
                            </li>
                          </ul>
                          <p style={{ margin: 0, fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                            <strong>Hinweis:</strong> diese Auswahl erscheint nur, wenn <strong>beide</strong> Filter gesetzt sind — sonst gibt es nichts zu kombinieren.
                          </p>
                        </>
                      ) : (
                        <>
                          <p style={{ margin: '0 0 8px' }}>
                            Defines how the <strong>location filter</strong> and the <strong>mailing lists / individual users</strong> are combined — i.e. which conditions must be true for a person before the event shows up in their list.
                          </p>
                          <ul style={{ margin: '0 0 8px 18px', padding: 0 }}>
                            <li style={{ marginBottom: 6 }}>
                              <strong>OR (default):</strong> <em>either filter is enough.</em> Example: <strong>Location = Cologne</strong>, <strong>list = SAPALL</strong> → anyone <strong>in Cologne</strong> <strong>OR</strong> in <strong>SAPALL</strong> sees the event. Useful when you intentionally want a <strong>broad audience</strong> (e.g. location staff <strong>plus</strong> a domain group).
                            </li>
                            <li>
                              <strong>AND:</strong> <em>both filters must match.</em> Example: <strong>Location = Cologne</strong>, <strong>list = SAPALL</strong> → only people <strong>in Cologne</strong> <strong>AND</strong> in <strong>SAPALL</strong> see the event. Useful when you want to <strong>strictly narrow</strong> the audience (e.g. only the SAP colleagues <strong>at the Cologne site</strong>).
                            </li>
                          </ul>
                          <p style={{ margin: 0, fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                            <strong>Note:</strong> this selector only appears when <strong>both</strong> filters are set — otherwise there is nothing to combine.
                          </p>
                        </>
                      )}
                    </div>
                    <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap' }}>
                      <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '0.9rem', cursor: 'pointer' }}>
                        <input type="radio" name="filterMode" value="OR" checked={filterMode === 'OR'} onChange={() => setFilterMode('OR')} />
                        <strong>{isDe ? 'ODER' : 'OR'}</strong>
                        <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>
                          – {isDe ? 'Einer der Filter reicht' : 'one filter is enough'}
                        </span>
                      </label>
                      <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '0.9rem', cursor: 'pointer' }}>
                        <input type="radio" name="filterMode" value="AND" checked={filterMode === 'AND'} onChange={() => setFilterMode('AND')} />
                        <strong>{isDe ? 'UND' : 'AND'}</strong>
                        <span style={{ color: 'var(--dex-gray-500)', fontSize: '0.8rem' }}>
                          – {isDe ? 'Beides muss zutreffen' : 'both must match'}
                        </span>
                      </label>
                    </div>
                  </div>
                ) : null}
                cardBgSecondary={(locationFilter || audience) ? zebraS3Bg() : '#fff'}
              />

              {/* v30.2: Verteiler aus den Sub-Events uebernehmen — der
                  haeufige Fall: EIN Sub-Event traegt bereits den echten
                  Verteiler (z.B. 113 Eintraege), die Klammer nur eine
                  Handvoll Personen. Der Knopf zieht die VEREINIGUNG aller
                  Sub-Event-Verteiler auf die Klammer hoch (Duplikate
                  case-insensitiv gefiltert, bestehende Klammer-Eintraege
                  bleiben). Er erscheint nur, wenn es dort tatsaechlich
                  Eintraege gibt, die der Klammer fehlen. */}
              {subEvents.length > 0 && (() => {
                const split = (s2: string): string[] => (s2 || '').split(',').map(x => x.trim()).filter(Boolean);
                const parentLc = new Set(split(audience).map(x => x.toLowerCase()));
                const missing: string[] = [];
                const seen = new Set<string>();
                subEvents.forEach(sd => split(sd.audience || '').forEach(a => {
                  const lc = a.toLowerCase();
                  if (parentLc.has(lc) || seen.has(lc)) return;
                  seen.add(lc);
                  missing.push(a);
                }));
                if (missing.length === 0) return null;
                return (
                  <div style={{ margin: '0 0 12px', padding: '10px 12px', borderRadius: 8, background: '#fff8e6', border: '1px solid #e0b34d', fontSize: '0.8rem', color: '#7a5a12', lineHeight: 1.5 }}>
                    {isDe
                      ? <>In den {childTermPlural || 'Sub-Events'} stehen <strong>{missing.length} Verteiler/Personen</strong>, die der Klammer fehlen. Der Zugang läuft immer über die Klammer — wer hier fehlt, sieht das Event nicht.</>
                      : <>The {childTermPlural || 'sub-events'} contain <strong>{missing.length} lists/people</strong> missing from the bracket. Access always goes through the bracket — anyone missing here cannot see the event.</>}
                    <button
                      type="button"
                      className="btn btn-primary"
                      style={{ display: 'block', fontSize: '0.8rem', padding: '5px 12px', marginTop: 8 }}
                      onClick={() => setAudience(split(audience).concat(missing).join(', '))}
                    >
                      {isDe
                        ? `Verteiler aus den ${childTermPlural || 'Sub-Events'} übernehmen (+${missing.length})`
                        : `Adopt audience from the ${childTermPlural || 'sub-events'} (+${missing.length})`}
                    </button>
                  </div>
                );
              })()}

              {/* v29.75: „Sichtbarkeit gilt für alle Sub-Events" — der Haken
                  spiegelt Standortfilter + Verteiler + Verknüpfung der Klammer
                  laufend in alle Sub-Event-Drafts (Effect bei den States) und
                  sperrt die Sichtbarkeits-UI der Sub-Event-Reiter, solange er
                  gesetzt ist. Persistiert (Piggyback _visAllSubs), damit auch
                  ein spaeterer Edit die Regel kennt. */}
              {subEvents.length > 0 && (
                <div className="form-group" style={{ padding: '12px 20px', marginBottom: 12, background: visAllSubs ? 'rgba(134,188,37,0.07)' : zebraS3Bg(), borderRadius: 8, border: `1px solid ${visAllSubs ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-100)'}` }}>
                  <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', margin: 0 }}>
                    <input
                      type="checkbox"
                      checked={visAllSubs}
                      onChange={e => { visAllSubsTouchedRef.current = true; setVisAllSubs(e.target.checked); }}
                      style={{ width: 18, height: 18, marginTop: 1, flexShrink: 0, cursor: 'pointer', accentColor: 'var(--dex-green, #86bc25)' }}
                    />
                    <span style={{ fontSize: '0.9rem' }}>
                      <strong>{isDe ? `Sichtbarkeit gilt für alle ${childTermPlural || 'Sub-Events'}` : `Visibility applies to all ${childTermPlural || 'sub-events'}`}</strong>
                      <span style={{ display: 'block', color: 'var(--dex-gray-600)', marginTop: 2, fontWeight: 400, fontSize: '0.8rem', lineHeight: 1.5 }}>
                        {isDe
                          ? <>Standortfilter, Mailverteiler und Verknüpfung von oben werden in alle {subEvents.length} {childTermPlural || 'Sub-Events'} übernommen und bleiben synchron, solange der Haken gesetzt ist. Die Sichtbarkeit in den {childTermSingular || 'Sub-Event'}-Reitern ist dann gesperrt — Haken entfernen, um wieder je {childTermSingular || 'Sub-Event'} abzuweichen.</>
                          : <>The location filter, mailing lists and combination above are applied to all {subEvents.length} {childTermPlural || 'sub-events'} and stay in sync while the box is checked. Visibility in the {childTermSingular || 'sub-event'} tabs is locked then — uncheck to deviate per {childTermSingular || 'sub-event'} again.</>}
                      </span>
                    </span>
                  </label>
                </div>
              )}

              {/* v23.6: Assistenz-Sichtbarkeit — eigener Baustein zwischen
                  „Mailverteiler / einzelne User" und „Anmeldefristen". Bewusst
                  AUSSERHALB des Greyout-Wrappers (laufzeit-/sichtbarkeitsrelevant,
                  wie der AudiencePicker oben — auch im Klammer-Modus editierbar). */}
              <div className="form-group" style={{ padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                {visHeader('vis_assist', <StepBadge n={(locationFilter && audience) ? 21 : 20} />, isDe ? 'Sichtbarkeit für Assistenzen' : 'Visibility for assistants')}
                {isVisOpen('vis_assist') && (<>
                <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', marginTop: -4, marginBottom: 12, lineHeight: 1.55 }}>
                  {isDe
                    ? <>Standardmäßig sieht eine <strong>Assistenz</strong> (Personen mit dem Job-Title „Assistenz“) nur Events, die sie auch als normale Nutzerin/normaler Nutzer sehen würde — also nur, wenn sie selbst in den oben gesetzten Standort-/Verteiler-Kreis fällt. Aktivierst du diese Option, sehen <strong>alle Assistenzen</strong> dieses Event in ihrer Übersicht, unabhängig von den Filtern oben — damit sie z.B. stellvertretend einen Partner oder eine Direktorin anmelden können.</>
                    : <>By default an <strong>assistant</strong> (people whose job title is „Assistenz“) only sees events they would see as a regular user — i.e. only if they themselves fall within the location/distribution audience set above. Enable this option to let <strong>all assistants</strong> see this event in their overview regardless of the filters above — so they can register a partner or director on their behalf, for example.</>}
                </p>
                <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.9rem' }}>
                  <input type="checkbox" checked={assistantsCanSee} onChange={e => setAssistantsCanSee(e.target.checked)} />
                  <span>{isDe ? 'Assistenzen dürfen dieses Event generell sehen (auch außerhalb des Filterkreises)' : 'Assistants may see this event in general (even outside the filter audience)'}</span>
                </label>
                </>)}
              </div>

              {/* v28.31: Greyout-Wrapper beginnt erst NACH den Fristen. Bis v28.30
                  lag der Fristen-Block mit in der Huelle (opacity + pointer-events:
                  none) — im Klammer-Modus liess sich der Abschnitt deshalb nicht
                  einmal aufklappen, obwohl die Klammer seit v28.20 eine EIGENE,
                  wirksame Anmeldefrist haben kann. Die Teilnehmerzahl bleibt
                  ausgegraut (die Klammer hat keine eigenen Plaetze). */}
              <div className="form-group" style={{ padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                {visHeader('vis_fristen', <StepBadge n={(locationFilter && audience) ? 22 : 21} />, <>{isDe ? 'Anmelde- und Abmeldefristen' : 'Registration & cancellation deadlines'}<InfoTooltip text={isDe
                    ? 'Bis wann können sich Teilnehmer anmelden bzw. fristgerecht abmelden? Die Abmeldefrist ist die kommunizierte Deadline — abmelden geht danach standardmäßig weiterhin bis zum Event-Ende, die Organizer werden dann aber automatisch informiert. Über die Option unter den Fristen lässt sich die Selbst-Abmeldung nach der Frist auch komplett sperren. Beide Werte werden anhand des Event-Datums automatisch vorgeschlagen, du kannst sie jederzeit überschreiben.'
                    : 'Until when can attendees register or cancel within the deadline? The cancellation deadline is the communicated cutoff — by default cancelling remains possible until the event ends, but organizers are then notified automatically. The option below the deadlines can instead lock self-cancellation completely after the cutoff. Both values are auto-suggested from the event date and can be overridden at any time.'} /></>)}
                {isVisOpen('vis_fristen') && (<>
              {/* v29.75: Klartext-Zusammenfassung fuer die Klammer — WAS gilt
                  gerade fuer alle Sub-Events? Die einzelnen Regeln (Freischalt-
                  Regel, Klammer-Anmeldefrist, Abmeldefrist, Sichtbarkeit)
                  stehen verstreut in dieser Sektion und in der Sichtbarkeit
                  darueber; hier laufen sie als ein lesbarer Absatz zusammen,
                  inklusive der Zahl abweichender Sub-Events. */}
              {subEvents.length > 0 && (() => {
                const term = childTermPlural || 'Sub-Events';
                const fmt = (v: string): string => v ? new Date(v).toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '';
                const ts = (v: string): number => { const t2 = new Date(v || '').getTime(); return isFinite(t2) ? t2 : 0; };
                // Abweichungen je Frist: wie viele Sub-Events haben einen
                // ANDEREN Wert als die Klammer? (Zeit-Vergleich, nicht
                // String-Vergleich — die ISO-Formate variieren.)
                const regRef = subEventsOnlyMode ? (klammerDeadline ? (berlinLocalToUtcIso(klammerDeadline) || '') : '') : (registrationDeadline ? (berlinLocalToUtcIso(registrationDeadline) || '') : '');
                // Leer zaehlt NICHT als Abweichung: beim Hauptevent erbt ein
                // leeres Sub-Event die Frist, bei der Klammer schliesst deren
                // harter Anmeldeschluss ohnehin das gesamte Event.
                const regDev = regRef ? subEvents.filter(s => (s.registrationDeadline || '').trim() !== '' && ts(s.registrationDeadline || '') !== ts(regRef)).length : 0;
                const cancelRef = lastDeregisterDate ? (berlinLocalToUtcIso(lastDeregisterDate) || '') : '';
                // Bei der Abmeldefrist gibt es KEINEN Klammer-Durchgriff — sie
                // wirkt nur ueber die kopierten Sub-Werte. Ein leeres Sub-Event
                // hat dort also wirklich keine Frist (= Abweichung); beim
                // Hauptevent erbt es weiterhin.
                const cancelDev = cancelRef ? subEvents.filter(s => {
                  const v = (s.lastDeregisterDate || '').trim();
                  if (!v) return subEventsOnlyMode;
                  return ts(v) !== ts(cancelRef);
                }).length : 0;
                // Normalisiert vergleichen — Altbestand unterscheidet sich oft
                // nur in Leerzeichen/Reihenfolge, das ist keine Abweichung.
                const norm = (s: string): string => (s || '').split(',').map(x => x.trim().toLowerCase()).filter(Boolean).sort().join('|');
                const visDev = subEvents.filter(s => norm(s.locationFilter || '') !== norm(locationFilter) || norm(s.audience || '') !== norm(audience)).length;
                const dev = (n: number): React.ReactElement | null => n > 0
                  ? <em style={{ color: 'var(--dex-orange, #ed8b00)' }}> ({isDe ? `${n} ${n === 1 ? 'weicht' : 'weichen'} ab` : `${n} deviate${n === 1 ? 's' : ''}`})</em>
                  : null;
                const locs = locationFilter.split(',').map(s => s.trim()).filter(Boolean);
                const auds = audience.split(',').map(s => s.trim()).filter(Boolean);
                const visText = (locs.length === 0 && auds.length === 0)
                  ? (isDe ? 'alle Mitarbeiter von Deloitte Deutschland' : 'everyone at Deloitte Germany')
                  : [
                      locs.length ? (isDe ? `Standort${locs.length === 1 ? '' : 'e'} ${locs.join(', ')}` : `location${locs.length === 1 ? '' : 's'} ${locs.join(', ')}`) : '',
                      auds.length ? (isDe ? `${auds.length} Verteiler/Person${auds.length === 1 ? '' : 'en'}` : `${auds.length} list${auds.length === 1 ? '' : 's'}/people`) : '',
                    ].filter(Boolean).join(filterMode === 'AND' ? (isDe ? ' UND ' : ' AND ') : (isDe ? ' ODER ' : ' OR '));
                return (
                  <div style={{
                    marginBottom: 14, padding: '10px 14px', borderRadius: 8,
                    background: 'rgba(134,188,37,0.07)', border: '1px solid var(--dex-green, #86bc25)',
                    fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.6,
                  }}>
                    <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>
                      {isDe ? `Aktuell gilt für alle ${subEvents.length} ${term}: ` : `Currently, for all ${subEvents.length} ${term}: `}
                    </strong>
                    <ul style={{ margin: '6px 0 0', paddingLeft: 18 }}>
                      {/* v29.77: „Anmeldung ab" gilt fuer ALLE Sub-Event-Events. */}
                      {(
                        <li>{openRuleEnabled
                          ? (isDe
                            ? <>Anmeldung ab: <strong>{openRuleDays} {openRuleDays === 1 ? 'Tag' : 'Tage'} vor {openRuleMode === 'week' ? 'dem Montag der jeweiligen Woche' : 'dem jeweiligen Termin'}</strong> (rollierend)</>
                            : <>Registration opens: <strong>{openRuleDays} {openRuleDays === 1 ? 'day' : 'days'} before {openRuleMode === 'week' ? 'the Monday of each week' : 'each date'}</strong> (rolling)</>)
                          : openRuleFixedDate
                            ? (isDe
                              ? <>Anmeldung ab: <strong>{fmt(berlinLocalToUtcIso(openRuleFixedDate) || '')}</strong> (alle Termine gemeinsam)</>
                              : <>Registration opens: <strong>{fmt(berlinLocalToUtcIso(openRuleFixedDate) || '')}</strong> (all dates together)</>)
                            : (isDe
                              ? <>Anmeldung ab: <strong>sofort</strong></>
                              : <>Registration opens: <strong>immediately</strong></>)}</li>
                      )}
                      <li>{regRuleEnabled
                        ? (isDe
                          ? <>Anmeldung bis: <strong>{regRuleAmount} {regRuleUnit === 'hours' ? (regRuleAmount === 1 ? 'Stunde' : 'Stunden') : (regRuleAmount === 1 ? 'Tag' : 'Tage')} vor dem jeweiligen Termin</strong> (rollierend)</>
                          : <>Registration until: <strong>{regRuleAmount} {regRuleUnit === 'hours' ? 'hour(s)' : 'day(s)'} before each date</strong> (rolling)</>)
                        : regRef
                          ? (isDe ? <>Anmeldung bis: <strong>{fmt(regRef)}</strong>{dev(regDev)}</> : <>Registration until: <strong>{fmt(regRef)}</strong>{dev(regDev)}</>)
                          : (isDe ? <>Anmeldung bis: <strong>je {childTermSingular || 'Sub-Event'}</strong> geregelt (keine gemeinsame Frist gesetzt)</> : <>Registration until: <strong>per {childTermSingular || 'sub-event'}</strong> (no shared deadline set)</>)}</li>
                      <li>{!userCancelAllowed
                        ? (isDe ? <>Abmeldung: <strong>deaktiviert</strong> — abmelden können nur Organizer und Admins</> : <>Cancellation: <strong>disabled</strong> — only organizers and admins can cancel</>)
                        : cancelRuleEnabled
                          ? (isDe
                            ? <>Abmeldung bis: <strong>{cancelRuleAmount} {cancelRuleUnit === 'hours' ? (cancelRuleAmount === 1 ? 'Stunde' : 'Stunden') : (cancelRuleAmount === 1 ? 'Tag' : 'Tage')} vor dem jeweiligen Termin</strong> (rollierend)</>
                            : <>Cancellation until: <strong>{cancelRuleAmount} {cancelRuleUnit === 'hours' ? 'hour(s)' : 'day(s)'} before each date</strong> (rolling)</>)
                          : cancelRef
                            ? (isDe ? <>Abmeldung bis: <strong>{fmt(cancelRef)}</strong>{dev(cancelDev)}</> : <>Cancellation until: <strong>{fmt(cancelRef)}</strong>{dev(cancelDev)}</>)
                            : (isDe ? <>Abmeldung bis: <strong>je {childTermSingular || 'Sub-Event'}</strong> geregelt (keine gemeinsame Frist gesetzt)</> : <>Cancellation until: <strong>per {childTermSingular || 'sub-event'}</strong> (no shared deadline set)</>)}</li>
                      <li>{isDe ? <>sichtbar für: <strong>{visText}</strong></> : <>visible to: <strong>{visText}</strong></>}
                        {visAllSubs
                          ? <> {isDe ? '— per Haken für alle übernommen' : '— applied to all via checkbox'}{dev(0)}</>
                          : dev(visDev) || <> {isDe ? `— je ${childTermSingular || 'Sub-Event'} anpassbar` : `— adjustable per ${childTermSingular || 'sub-event'}`}</>}</li>
                    </ul>
                  </div>
                );
              })()}
              {subEventsOnlyMode && (
                <WizardHint
                  isDe={isDe}
                  title={isDe ? 'Frist ergibt sich aus den Sub-Events' : 'Deadline derives from the sub-events'}
                  defaultOpen
                  style={{ marginBottom: 12 }}
                >
                  <div>
                    {/* v30.2: Text an die rollierenden Fristen (v29.76/77)
                        angepasst — der alte Stand behauptete, Fristen wuerden
                        „weiterhin je Tab" gepflegt, und kannte weder
                        „Anmeldung ab" noch die zentrale Uebernahme. */}
                    {isDe ? (
                      <>
                        Die Klammer selbst ist <strong>nicht buchbar</strong> — hier stellst du die Anmelde-Logik für <strong>alle {childTermPlural || 'Sub-Events'}</strong> zentral ein: <strong>&bdquo;Anmeldung ab&ldquo;</strong> (festes Datum oder rollierend), <strong>&bdquo;Anmeldung bis&ldquo;</strong> und <strong>&bdquo;Abmeldung bis&ldquo;</strong> (jeweils festes Datum, das in alle {childTermPlural || 'Sub-Events'} übernommen wird, oder rollierend — dann wird die Frist <strong>je Termin automatisch ausgerechnet</strong>, auch für später hinzugefügte). Eine gesetzte feste Klammer-Anmeldefrist wirkt zusätzlich als <strong>harter Schluss fürs gesamte Event</strong>; ohne sie bleibt die Anmeldung offen, solange mindestens ein {childTermSingular || 'Sub-Event'} offen ist. Feste Fristen kannst du im jeweiligen Reiter pro {childTermSingular || 'Sub-Event'} anpassen — bei aktiver <strong>rollierender</strong> Regel sind die Felder dort gesperrt, weil die Regel sie berechnet.
                      </>
                    ) : (
                      <>
                        The bracket itself is <strong>not bookable</strong> — here you configure the registration logic for <strong>all {childTermPlural || 'sub-events'}</strong> centrally: <strong>“registration opens”</strong> (fixed date or rolling), <strong>“registration until”</strong> and <strong>“cancellation until”</strong> (a fixed date copied to all {childTermPlural || 'sub-events'}, or rolling — then the deadline is <strong>computed per date automatically</strong>, including dates added later). A fixed bracket registration deadline additionally acts as a <strong>hard cutoff for the entire event</strong>; without one, registration stays open as long as at least one {childTermSingular || 'sub-event'} is open. Fixed deadlines remain adjustable per {childTermSingular || 'sub-event'} in its tab — with an active <strong>rolling</strong> rule those fields are locked there, because the rule computes them.
                      </>
                    )}
                  </div>
                </WizardHint>
              )}
              {/* v29.75: Freischalt-Regel („Anmeldung ab") — von Schritt 1
                  (Kalender-Block) hierher gezogen: Sie ist eine Anmelde-Regel
                  und gehoert neben Anmelde- und Abmeldefrist. Nur im
                  Kalender-Modus sinnvoll, weil die Anmeldeseite sie auf den
                  Tages-Kacheln umsetzt (v29.67/v29.72). */}
              {/* v29.77: „Anmeldung ab" steht IMMER hier — bei Sub-Event-Events
                  als feste/rollierende Freischalt-Regel, bei reinen Haupt-
                  events als das Aktivierungsdatum aus den Grundlagen (bewusst
                  doppelt angeboten, damit die Sektion immer die volle Logik
                  „Anmeldung ab / Anmeldung bis / Abmeldung bis" zeigt). */}
              {subEventsOptIn ? (
                <div style={{ margin: '0 0 14px', padding: '10px 12px', borderRadius: 8, background: '#fff', border: '1px solid var(--dex-gray-200)' }}>
                  {/* v29.75: User-Wortlaut — die drei Felder heissen
                      „Anmeldung ab" / „Anmeldung bis" / „Abmeldung bis". */}
                  <label className="form-label">{isDe ? 'Anmeldung ab' : 'Registration opens'}</label>
                  {/* v30.2: Datum ZUERST (direkt unter der Ueberschrift),
                      Checkbox darunter — dieselbe Reihenfolge und derselbe
                      Wortlaut („Rollierend je Termin") wie bei „Anmeldung
                      bis"/„Abmeldung bis". Das Datumsfeld verschwindet,
                      sobald rollierend gewaehlt ist. */}
                  {!openRuleEnabled && (
                    <div style={{ maxWidth: 340 }}>
                      <DatePicker
                        selected={openRuleFixedDate ? new Date(openRuleFixedDate) : null}
                        onChange={(date: Date | null) => setOpenRuleFixedDate(date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}` : '')}
                        showTimeSelect
                        timeFormat="HH:mm"
                        timeIntervals={15}
                        timeCaption="Uhrzeit"
                        dateFormat="dd.MM.yyyy, HH:mm"
                        locale="de"
                        placeholderText={isDe ? 'Optional — leer = sofort anmeldbar' : 'Optional — empty = open immediately'}
                        className="form-input"
                        wrapperClassName="dex-datepicker-wrapper"
                        calendarClassName="dex-datepicker-calendar"
                        popperPlacement="bottom-start"
                        isClearable
                        autoComplete="off"
                      />
                      {openRuleFixedDate && (
                        <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', margin: '6px 0 0' }}>
                          {isDe
                            ? 'Alle Termine öffnen gemeinsam zu diesem Zeitpunkt.'
                            : 'All dates open together at this moment.'}
                        </p>
                      )}
                    </div>
                  )}
                  <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.85rem', margin: '8px 0' }}>
                    <input
                      type="checkbox"
                      checked={openRuleEnabled}
                      onChange={e => setOpenRuleEnabled(e.target.checked)}
                      style={{ width: 16, height: 16, cursor: 'pointer' }}
                    />
                    <span>{isDe ? 'Rollierend je Termin' : 'Rolling per date'}</span>
                  </label>
                  {openRuleEnabled && (
                    <div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', fontSize: '0.88rem' }}>
                        {isDe ? 'Anmeldung möglich ab' : 'Registration opens'}
                        <input
                          type="number"
                          min={1}
                          max={365}
                          className="form-input"
                          value={openRuleDays}
                          onChange={e => { const v = parseInt(e.target.value, 10); setOpenRuleDays(isFinite(v) && v > 0 ? Math.min(v, 365) : 1); }}
                          style={{ width: 76, padding: '4px 8px', textAlign: 'center' }}
                        />
                        {isDe ? 'Tage vor' : 'days before'}
                        <select
                          className="form-input"
                          value={openRuleMode}
                          onChange={e => setOpenRuleMode(e.target.value === 'week' ? 'week' : 'day')}
                          style={{ width: 'auto', padding: '4px 8px' }}
                        >
                          <option value="day">{isDe ? 'dem jeweiligen Termin' : 'each date'}</option>
                          <option value="week">{isDe ? 'dem Montag der jeweiligen Woche' : 'the Monday of its week'}</option>
                        </select>
                      </div>
                      <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', margin: '6px 0 0' }}>
                        {openRuleMode === 'week'
                          ? (isDe
                            ? `Alle Termine einer Kalenderwoche öffnen gemeinsam: ${openRuleDays} ${openRuleDays === 1 ? 'Tag' : 'Tage'} vor deren Montag.`
                            : `All dates of a calendar week open together: ${openRuleDays} ${openRuleDays === 1 ? 'day' : 'days'} before that week's Monday.`)
                          : (isDe
                            ? `Jeder Termin öffnet einzeln: ${openRuleDays} ${openRuleDays === 1 ? 'Tag' : 'Tage'} vor seinem Datum.`
                            : `Each date opens individually: ${openRuleDays} ${openRuleDays === 1 ? 'day' : 'days'} before its date.`)}
                        {' '}
                        {isDe
                          ? 'Noch nicht freigeschaltete Termine sind auf der Anmeldeseite ausgegraut und zeigen, ab wann die Anmeldung möglich ist.'
                          : 'Dates not yet open appear greyed out on the registration page and show when registration becomes possible.'}
                      </p>
                    </div>
                  )}
                </div>
              ) : (
                <div style={{ margin: '0 0 14px', padding: '10px 12px', borderRadius: 8, background: '#fff', border: '1px solid var(--dex-gray-200)' }}>
                  <label className="form-label">{isDe ? 'Anmeldung ab' : 'Registration opens'}</label>
                  <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)', margin: '0 0 8px', lineHeight: 1.5 }}>
                    {isDe
                      ? <>Dasselbe Feld wie die <strong>Aktivierung in Schritt 1</strong> (bewusst an beiden Stellen): Bis zu diesem Zeitpunkt ist das Event nicht anmeldbar — je nach Vorschau-Einstellung unsichtbar oder als Vorschau mit dem Hinweis &bdquo;Anmeldung ab …&ldquo;. Leer = sofort anmeldbar.</>
                      : <>The same field as the <strong>activation in step 1</strong> (deliberately in both places): until this moment the event cannot be booked — invisible or shown as a preview with a &bdquo;registration opens …&ldquo; note, depending on the preview setting. Empty = open immediately.</>}
                  </p>
                  <div style={{ maxWidth: 340 }}>
                    <DatePicker
                      selected={activeFrom ? new Date(activeFrom) : null}
                      onChange={(date: Date | null) => setActiveFrom(date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}` : '')}
                      showTimeSelect
                      timeFormat="HH:mm"
                      timeIntervals={15}
                      timeCaption="Uhrzeit"
                      dateFormat="dd.MM.yyyy, HH:mm"
                      locale="de"
                      placeholderText={isDe ? 'Optional — leer = sofort anmeldbar' : 'Optional — empty = open immediately'}
                      className="form-input"
                      wrapperClassName="dex-datepicker-wrapper"
                      calendarClassName="dex-datepicker-calendar"
                      popperPlacement="bottom-start"
                      isClearable
                      autoComplete="off"
                    />
                  </div>
                </div>
              )}
              <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                {/* v30.2: gleiche weisse Karte wie „Anmeldung ab" — die drei
                    Fristen sollen als EINE Familie lesbar sein. */}
                <div className="form-group" style={{ marginBottom: 0, background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: '10px 12px' }}>
                  <label className="form-label">
                    {/* v29.75: User-Wortlaut „Anmeldung bis" statt „Anmeldefrist". */}
                    {isDe ? 'Anmeldung bis' : 'Registration until'}
                    {/* v28.20: Im Klammer-Modus eine eigene Erklärung — was die
                        Klammer-Frist bewirkt und wie sie mit abweichenden
                        Sub-Event-Fristen zusammenspielt. */}
                    <InfoTooltip text={subEventsOnlyMode ? (isDe ? (
                      <>
                        <strong>Anmeldefrist der Klammer (optional)</strong> — die Klammer selbst ist nicht buchbar; diese Frist wirkt als <strong>harter Anmeldeschluss für das gesamte Event</strong>.<br /><br />
                        <strong>Gesetzt:</strong> Nach dem Stichtag ist die Anmeldung komplett geschlossen — auch wenn einzelne {childTermPlural || 'Sub-Events'} eine spätere oder gar keine eigene Frist haben. Beim Setzen wird der Termin automatisch <strong>in alle {childTermSingular || 'Sub-Event'}-Tabs übernommen</strong>.<br /><br />
                        <strong>Abweichung pro {childTermSingular || 'Sub-Event'}:</strong> Im jeweiligen Tab kannst du die Frist danach ändern — sinnvoll ist ein <strong>früherer</strong> Schluss (das {childTermSingular || 'Sub-Event'} macht dann eher zu). Eine <strong>spätere</strong> Frist als die Klammer hat keine Wirkung, weil die Klammer das gesamte Event zuerst schließt.<br /><br />
                        <strong>Leer:</strong> Es gelten allein die {childTermSingular || 'Sub-Event'}-Fristen — die Anmeldung bleibt offen, solange mindestens ein {childTermSingular || 'Sub-Event'} offen ist (effektiv bis zur spätesten Frist).<br /><br />
                        Organizer und Admins können wie immer auch nach Fristablauf manuell anmelden.
                      </>
                    ) : (
                      <>
                        <strong>Bracket registration deadline (optional)</strong> — the bracket itself is not bookable; this deadline acts as a <strong>hard registration cutoff for the entire event</strong>.<br /><br />
                        <strong>Set:</strong> past the cutoff, registration is fully closed — even if individual {childTermPlural || 'sub-events'} have a later or no own deadline. When you set it, the date is automatically <strong>copied to all {childTermSingular || 'sub-event'} tabs</strong>.<br /><br />
                        <strong>Deviating per {childTermSingular || 'sub-event'}:</strong> you can change the deadline in each tab afterwards — an <strong>earlier</strong> cutoff makes sense (that {childTermSingular || 'sub-event'} closes sooner). A <strong>later</strong> deadline than the bracket has no effect, because the bracket closes the whole event first.<br /><br />
                        <strong>Empty:</strong> only the {childTermSingular || 'sub-event'} deadlines apply — registration stays open as long as at least one {childTermSingular || 'sub-event'} is open (effectively until the latest deadline).<br /><br />
                        Organizers and admins can, as always, register manually after the cutoff.
                      </>
                    )) : isDe ? (
                      <>
                        <strong>Anmelde-Deadline</strong> — bis zu diesem Stichtag können sich Teilnehmer selbst registrieren.<br /><br />
                        <strong>Auswirkung für Teilnehmer:</strong> nach dem Stichtag ist der <strong>Anmelden-Button gesperrt</strong> (auch via Direktlink), reguläre User können sich nicht mehr selbst eintragen. <strong>Organizer und Co-Organizer</strong> dürfen weiterhin manuell Teilnehmer anlegen — die Deadline gilt nur für Self-Registration.<br /><br />
                        Vorbefüllt mit <strong>7 Tagen vor Event-Start</strong>, frei überschreibbar.
                      </>
                    ) : (
                      <>
                        <strong>Registration deadline</strong> — until this cutoff attendees can self-register.<br /><br />
                        <strong>Effect for attendees:</strong> past the cutoff the <strong>register button is locked</strong> (also via direct link), regular users can no longer sign themselves up. <strong>Organizers and co-organizers</strong> can still add attendees manually — the deadline only applies to self-registration.<br /><br />
                        Pre-filled with <strong>7 days before event start</strong>, freely overridable.
                      </>
                    )} />
                  </label>
                  {/* v29.76: Entweder festes Datum ODER rollierend je Termin.
                      Beim Umschalten auf rollierend wird das feste Datum
                      geleert — zwei gleichzeitig wirkende Fristen wuerde
                      niemand mehr nachvollziehen koennen.
                      v29.77: fuer ALLE Sub-Event-Events, nicht nur Kalender.
                      v30.2: Datum ZUERST, Checkbox darunter — konsistent
                      zu „Anmeldung ab". */}
                  {!(subEventsOptIn && regRuleEnabled) && (
                  <DatePicker
                    // v28.20: Im Klammer-Modus ist die Frist jetzt EDITIERBAR
                    // (eigener State klammerDeadline, Piggyback) — gesetzt +
                    // abgelaufen schließt das GESAMTE Event. Leer = wie bisher
                    // offen bis zur spätesten Sub-Event-Frist.
                    selected={subEventsOnlyMode
                      ? (klammerDeadline ? new Date(klammerDeadline) : null)
                      : (registrationDeadline ? new Date(registrationDeadline) : null)}
                    onChange={(date: Date | null) => {
                      const v = date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}` : '';
                      if (subEventsOnlyMode) {
                        setKlammerDeadline(v);
                        // v28.20: Klammer-Frist in ALLE Sub-Event-Tabs
                        // übernehmen (dort pro Sub-Event weiter anpassbar,
                        // z.B. auf einen früheren Schluss). Beim Leeren der
                        // Klammer-Frist bleiben die Sub-Fristen unberührt.
                        if (v) {
                          const iso = berlinLocalToUtcIso(v) || '';
                          setSubEvents(prev => prev.map(s => ({ ...s, registrationDeadline: iso })));
                        }
                      } else {
                        setRegistrationDeadline(v);
                      }
                    }}
                    showTimeSelect
                    timeFormat="HH:mm"
                    timeIntervals={15}
                    timeCaption="Uhrzeit"
                    dateFormat="dd.MM.yyyy, HH:mm"
                    locale="de"
                    placeholderText={subEventsOnlyMode
                      ? (isDe ? 'Optional — sonst automatisch aus Sub-Events' : 'Optional — otherwise automatic from sub-events')
                      : 'Anmelde-Deadline'}
                    className="form-input"
                    wrapperClassName="dex-datepicker-wrapper"
                    calendarClassName="dex-datepicker-calendar"
                    popperPlacement="bottom-start"
                    isClearable
                    autoComplete="off"
                  />
                  )}
                  {subEventsOptIn && (
                    <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.85rem', margin: '8px 0' }}>
                      <input
                        type="checkbox"
                        checked={regRuleEnabled}
                        onChange={e => {
                          setRegRuleEnabled(e.target.checked);
                          if (e.target.checked) {
                            if (subEventsOnlyMode) setKlammerDeadline(''); else setRegistrationDeadline('');
                          }
                        }}
                        style={{ width: 16, height: 16, cursor: 'pointer' }}
                      />
                      <span>{isDe ? 'Rollierend je Termin' : 'Rolling per date'}</span>
                    </label>
                  )}
                  {(subEventsOptIn && regRuleEnabled) && (
                    <div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', fontSize: '0.88rem' }}>
                        {isDe ? 'Anmeldung bis' : 'Registration until'}
                        <input
                          type="number"
                          min={1}
                          max={365}
                          className="form-input"
                          value={regRuleAmount}
                          onChange={e => { const v = parseInt(e.target.value, 10); setRegRuleAmount(isFinite(v) && v > 0 ? Math.min(v, 365) : 1); }}
                          style={{ width: 76, padding: '4px 8px', textAlign: 'center' }}
                        />
                        {/* v29.77: Einheit im Singular, wenn die Zahl 1 ist. */}
                        <select
                          className="form-input"
                          value={regRuleUnit}
                          onChange={e => setRegRuleUnit(e.target.value === 'hours' ? 'hours' : 'days')}
                          style={{ width: 'auto', padding: '4px 8px' }}
                        >
                          <option value="days">{isDe ? (regRuleAmount === 1 ? 'Tag' : 'Tage') : (regRuleAmount === 1 ? 'day' : 'days')}</option>
                          <option value="hours">{isDe ? (regRuleAmount === 1 ? 'Stunde' : 'Stunden') : (regRuleAmount === 1 ? 'hour' : 'hours')}</option>
                        </select>
                        {isDe ? 'vor dem jeweiligen Termin' : 'before each date'}
                      </div>
                      <p style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', margin: '6px 0 0' }}>
                        {isDe
                          ? 'Die Frist wird je Termin ausgerechnet und in dessen Einstellungen geschrieben — auch für später hinzugefügte Termine.'
                          : 'The deadline is computed per date and written into its settings — including dates added later.'}
                      </p>
                    </div>
                  )}
                  {subEventsOnlyMode && !klammerDeadline && effectiveKlammerDeadline && (
                    <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                      {isDe
                        ? <>Aktuell effektiv: <strong>{new Date(effectiveKlammerDeadline).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })}</strong> (späteste Sub-Event-Frist)</>
                        : <>Currently effective: <strong>{new Date(effectiveKlammerDeadline).toLocaleString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })}</strong> (latest sub-event deadline)</>}
                    </div>
                  )}
                  {subEventsOnlyMode && klammerDeadline && (
                    <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                      {isDe
                        ? <>Gilt als Anmeldeschluss fürs <strong>gesamte Event</strong> und wurde in alle {childTermPlural || 'Sub-Event'}-Tabs übernommen — dort pro {childTermSingular || 'Sub-Event'} anpassbar (z.B. früherer Schluss).</>
                        : <>Acts as the registration cutoff for the <strong>entire event</strong> and has been copied to all {childTermPlural || 'sub-event'} tabs — adjustable there per {childTermSingular || 'sub-event'} (e.g. an earlier cutoff).</>}
                    </div>
                  )}
                </div>
                {/* v28.31: Die Abmeldefrist gehört bei einer Klammer zu den
                    Sub-Events — hier wäre sie wirkungslos. Nur dieses eine Feld
                    ausgrauen, die Klammer-Anmeldefrist links bleibt bedienbar. */}
                {/* v29.25: Ohne Selbst-Abmeldung gibt es keine Abmeldefrist —
                    statt des Datumsfelds steht der Grund. */}
                {!userCancelAllowed ? (
                  <div className="form-group" style={{ marginBottom: 0, background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: '10px 12px' }}>
                    <label className="form-label">{isDe ? 'Abmeldung bis' : 'Cancellation until'}</label>
                    <div style={{
                      padding: '10px 12px', borderRadius: 8, fontSize: '0.8rem', lineHeight: 1.5,
                      background: 'var(--dex-gray-50, #fafafa)', border: '1px dashed var(--dex-gray-300)',
                      color: 'var(--dex-gray-600)',
                    }}>
                      {isDe
                        ? 'Entfällt — die Abmeldung durch User ist deaktiviert (Option unten). Abmelden können nur Organizer und Admins.'
                        : 'Not applicable — self-cancellation is disabled (option below). Only organizers and admins can cancel.'}
                    </div>
                  </div>
                ) : (
                /* v29.75: Bei einer Klammer war das Feld bis v29.74 ausgegraut
                   („gehoert zum einzelnen Sub-Event"). Jetzt funktioniert es
                   wie die Klammer-Anmeldefrist (v28.20): Setzen kopiert den
                   Termin in ALLE Sub-Event-Tabs, dort bleibt er pro Sub-Event
                   anpassbar. Anders als die Anmeldefrist hat die Klammer
                   KEINEN eigenen Durchgriff — die Wirkung kommt allein ueber
                   die kopierten Sub-Werte. */
                <div className="form-group" style={{ marginBottom: 0, background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: '10px 12px' }}>
                  <label className="form-label">
                    {/* v29.75: User-Wortlaut „Abmeldung bis". */}
                    {isDe ? 'Abmeldung bis' : 'Cancellation until'}
                    <InfoTooltip text={isDe ? (
                      <>
                        <strong>Letzte Abmeldemöglichkeit</strong> — der Stichtag, den du den Teilnehmern als <strong>verbindliche Abmeldefrist kommunizierst</strong>. Bis dahin gilt eine Abmeldung als unproblematisch.<br /><br />
                        <strong>Auswirkung für Teilnehmer:</strong> Eine Abmeldung bleibt bewusst <strong>bis zum Ende des Events möglich</strong> — wer kurzfristig erkrankt oder verhindert ist, kann sich also weiterhin abmelden. Nach dem Stichtag sieht die Person beim Abmelden einen <strong>deutlichen Hinweis</strong>, dass die Frist abgelaufen ist und die Organizer informiert werden. Erst <strong>nach Event-Ende</strong> ist die Selbst-Abmeldung gesperrt.<br /><br />
                        <strong>Automatismen:</strong> Bei jeder Abmeldung <strong>nach dem Stichtag</strong> bekommen die Organizer automatisch eine <strong>Info-Mail</strong> mit Name + E-Mail der Person — damit Hotel, Catering oder Transfers angepasst werden können. Zusätzliche Abmelde-Benachrichtigungen kannst du in <strong>Schritt 6 (Kommunikation)</strong> konfigurieren.<br /><br />
                        Mit der Option <strong>unter den Fristen</strong> kannst du die Selbst-Abmeldung nach dem Stichtag stattdessen <strong>komplett sperren</strong>.<br /><br />
                        Vorbefüllt mit <strong>3 Tagen vor Event-Start</strong>.
                      </>
                    ) : (
                      <>
                        <strong>Last cancellation date</strong> — the cutoff you <strong>communicate to attendees as the binding cancellation deadline</strong>. Up to this date a cancellation is considered routine.<br /><br />
                        <strong>Effect for attendees:</strong> cancelling deliberately stays <strong>possible until the event ends</strong> — anyone who falls ill or is prevented at short notice can still cancel. After the cutoff the person sees a <strong>clear notice</strong> when cancelling that the deadline has passed and the organizers will be informed. Only <strong>after the event has ended</strong> is self-cancellation locked.<br /><br />
                        <strong>Automation:</strong> for every cancellation <strong>after the cutoff</strong> the organizers automatically receive an <strong>info email</strong> with the person’s name + email — so hotel, catering or transfers can be adjusted. Additional cancellation notifications can be configured in <strong>step 6 (Communication)</strong>.<br /><br />
                        With the option <strong>below the deadlines</strong> you can instead <strong>lock self-cancellation completely</strong> after the cutoff.<br /><br />
                        Pre-filled with <strong>3 days before event start</strong>.
                      </>
                    )} />
                  </label>
                  {/* v29.76: wie „Anmeldung bis" — festes Datum ODER rollierend.
                      v30.2: Datum ZUERST, Checkbox darunter — konsistent
                      zu „Anmeldung ab". */}
                  {!(subEventsOptIn && cancelRuleEnabled) && (
                  <DatePicker
                    selected={lastDeregisterDate ? new Date(lastDeregisterDate) : null}
                    onChange={(date: Date | null) => {
                      const v = date ? `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}` : '';
                      setLastDeregisterDate(v);
                      // v29.75: Wie die Klammer-Anmeldefrist (v28.20) — in alle
                      // Sub-Event-Tabs uebernehmen. Beim Leeren bleiben die
                      // Sub-Fristen unberuehrt (bewusst gleiches Verhalten).
                      if (subEventsOnlyMode && v) {
                        const iso = berlinLocalToUtcIso(v) || '';
                        setSubEvents(prev => prev.map(s => ({ ...s, lastDeregisterDate: iso })));
                      }
                    }}
                    showTimeSelect
                    timeFormat="HH:mm"
                    timeIntervals={15}
                    timeCaption="Uhrzeit"
                    dateFormat="dd.MM.yyyy, HH:mm"
                    locale="de"
                    placeholderText="Abmeldefrist"
                    className="form-input"
                    wrapperClassName="dex-datepicker-wrapper"
                    calendarClassName="dex-datepicker-calendar"
                    popperPlacement="bottom-start"
                    isClearable
                    autoComplete="off"
                  />
                  )}
                  {subEventsOptIn && (
                    <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.85rem', margin: '8px 0' }}>
                      <input
                        type="checkbox"
                        checked={cancelRuleEnabled}
                        onChange={e => {
                          setCancelRuleEnabled(e.target.checked);
                          if (e.target.checked) setLastDeregisterDate('');
                        }}
                        style={{ width: 16, height: 16, cursor: 'pointer' }}
                      />
                      <span>{isDe ? 'Rollierend je Termin' : 'Rolling per date'}</span>
                    </label>
                  )}
                  {(subEventsOptIn && cancelRuleEnabled) && (
                    <div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', fontSize: '0.88rem' }}>
                        {isDe ? 'Abmeldung bis' : 'Cancellation until'}
                        <input
                          type="number"
                          min={1}
                          max={365}
                          className="form-input"
                          value={cancelRuleAmount}
                          onChange={e => { const v = parseInt(e.target.value, 10); setCancelRuleAmount(isFinite(v) && v > 0 ? Math.min(v, 365) : 1); }}
                          style={{ width: 76, padding: '4px 8px', textAlign: 'center' }}
                        />
                        {/* v29.77: Einheit im Singular, wenn die Zahl 1 ist. */}
                        <select
                          className="form-input"
                          value={cancelRuleUnit}
                          onChange={e => setCancelRuleUnit(e.target.value === 'hours' ? 'hours' : 'days')}
                          style={{ width: 'auto', padding: '4px 8px' }}
                        >
                          <option value="days">{isDe ? (cancelRuleAmount === 1 ? 'Tag' : 'Tage') : (cancelRuleAmount === 1 ? 'day' : 'days')}</option>
                          <option value="hours">{isDe ? (cancelRuleAmount === 1 ? 'Stunde' : 'Stunden') : (cancelRuleAmount === 1 ? 'hour' : 'hours')}</option>
                        </select>
                        {/* v29.77: Abmelden darf auch NACH dem Termin-Beginn
                            noch offen sein — Richtung waehlbar. */}
                        <select
                          className="form-input"
                          value={cancelRuleAfter ? 'after' : 'before'}
                          onChange={e => setCancelRuleAfter(e.target.value === 'after')}
                          style={{ width: 'auto', padding: '4px 8px' }}
                        >
                          <option value="before">{isDe ? 'vor dem jeweiligen Termin' : 'before each date'}</option>
                          <option value="after">{isDe ? 'nach dem jeweiligen Termin' : 'after each date starts'}</option>
                        </select>
                      </div>
                      <p style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', margin: '6px 0 0' }}>
                        {isDe
                          ? 'Die Frist wird je Termin ausgerechnet und in dessen Einstellungen geschrieben — auch für später hinzugefügte Termine.'
                          : 'The deadline is computed per date and written into its settings — including dates added later.'}
                      </p>
                    </div>
                  )}
                </div>
                )}
              </div>
              {/* v29.75: Das Feld ist bei einer Klammer nicht mehr ausgegraut —
                  erklaeren, WIE es wirkt (Kopie in alle Tabs, kein eigener
                  Klammer-Durchgriff wie bei der Anmeldefrist). */}
              {subEventsOnlyMode && userCancelAllowed && (
                <p style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', marginTop: 8, marginBottom: 0, lineHeight: 1.5 }}>
                  {isDe
                    ? <>Setzt du hier eine <strong>Abmeldefrist</strong>, wird sie in <strong>alle {childTermPlural || 'Sub-Events'}</strong> übernommen — im jeweiligen Reiter bleibt sie danach pro {childTermSingular || 'Sub-Event'} anpassbar. Die <strong>Anmeldefrist</strong> links wirkt zusätzlich als harter Anmeldeschluss für das gesamte Event.</>
                    : <>Setting a <strong>cancellation deadline</strong> here copies it to <strong>all {childTermPlural || 'sub-events'}</strong> — it stays adjustable per {childTermSingular || 'sub-event'} in its tab. The <strong>registration deadline</strong> on the left additionally acts as a hard cutoff for the entire event.</>}
                </p>
              )}
              {/* v29.25: Selbst-Abmeldung, zweistufig. Stufe 1 steht bewusst
                  DIREKT unter den Fristen und erklärt zuerst den Default —
                  sonst liest sich die Option, als wäre Abmelden heute schon
                  eingeschränkt. Stufe 2 erscheint nur, wenn sie überhaupt
                  greifen kann (Selbst-Abmeldung erlaubt + Frist gesetzt bzw.
                  Klammer mit Sub-Event-Fristen). Beides event-weit. */}
              <div style={{
                marginTop: 12, padding: '12px 14px', borderRadius: 10,
                border: `1px solid ${!userCancelAllowed ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-gray-200)'}`,
                background: !userCancelAllowed ? 'rgba(237,139,0,0.06)' : 'var(--dex-gray-50, #fafafa)',
              }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', margin: 0 }}>
                  <input
                    type="checkbox"
                    checked={userCancelAllowed}
                    onChange={e => setUserCancelAllowed(e.target.checked)}
                    style={{ marginTop: 3, width: 16, height: 16, accentColor: 'var(--dex-green, #86bc25)', flexShrink: 0 }}
                  />
                  <span style={{ fontSize: '0.85rem', lineHeight: 1.5 }}>
                    <strong>{isDe ? 'Abmeldung durch User ermöglichen' : 'Allow users to cancel themselves'}</strong>
                    <span style={{ display: 'block', marginTop: 4, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                      {isDe ? (
                        <>
                          <strong>Standard (an):</strong> Teilnehmer können sich selbst abmelden — unter „Meine Events“ und über den Link in der Bestätigungsmail.<br />
                          <strong>Deaktiviert:</strong> Teilnehmer können sich <strong>gar nicht selbst abmelden</strong>; eine Abmeldefrist entfällt. Abmelden können nur Organizer und Admins über das Organizer Center — dort steht neben den angemeldeten Teilnehmern zusätzlich der <strong>No-Show</strong>-Knopf, um Nicht-Erschienene zu markieren.
                        </>
                      ) : (
                        <>
                          <strong>Default (on):</strong> attendees can cancel themselves — under “My events” and via the link in the confirmation mail.<br />
                          <strong>Disabled:</strong> attendees <strong>cannot cancel themselves at all</strong>; there is no cancellation deadline. Only organizers and admins can cancel, via the Organizer Center — which then also shows a <strong>No-Show</strong> button next to registered attendees to mark people who did not appear.
                        </>
                      )}
                    </span>
                  </span>
                </label>
                {userCancelAllowed && (lastDeregisterDate || subEventsOnlyMode) && (
                  <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', margin: '12px 0 0', paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
                    <input
                      type="checkbox"
                      checked={!noCancelAfterDeadline}
                      onChange={e => setNoCancelAfterDeadline(!e.target.checked)}
                      style={{ marginTop: 3, width: 16, height: 16, accentColor: 'var(--dex-green, #86bc25)', flexShrink: 0 }}
                    />
                    <span style={{ fontSize: '0.85rem', lineHeight: 1.5 }}>
                      <strong>{isDe ? 'Abmeldung auch nach der Abmeldefrist erlauben' : 'Also allow cancelling after the deadline'}</strong>
                      <span style={{ display: 'block', marginTop: 4, fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                        {isDe ? (
                          <>
                            <strong>Standard (an):</strong> Teilnehmer können sich auch nach der Abmeldefrist noch abmelden — die Organizer bekommen dann automatisch eine Info-Mail („Verspätete Abmeldung“) mit Name und E-Mail der Person.<br />
                            <strong>Deaktiviert:</strong> Nach Ablauf der Frist können sich Teilnehmer <strong>nicht mehr selbst abmelden</strong> — abmelden können dann nur noch Organizer und Admins über das Organizer Center, dort ab der Frist ebenfalls mit <strong>No-Show</strong>-Knopf.
                            {subEventsOnlyMode ? <> Bei einer Klammer greift die Sperre je {childTermSingular || 'Sub-Event'} nach dessen eigener Abmeldefrist.</> : null}
                          </>
                        ) : (
                          <>
                            <strong>Default (on):</strong> attendees can still cancel after the deadline — the organizers then automatically receive an info email (“late cancellation”) with the person’s name and email.<br />
                            <strong>Disabled:</strong> once the deadline has passed, attendees <strong>can no longer cancel themselves</strong> — only organizers and admins can cancel, via the Organizer Center, which then also shows the <strong>No-Show</strong> button from the deadline on.
                            {subEventsOnlyMode ? <> For a bracket the lock applies per {childTermSingular || 'sub-event'} based on its own cancellation deadline.</> : null}
                          </>
                        )}
                      </span>
                    </span>
                  </label>
                )}
              </div>
              {fieldHasError('deadlineAfterStart') && <p style={{ color: 'var(--dex-red)', fontSize: '0.8rem', marginTop: 8, marginBottom: 0 }}>{t('create.error.deadlineAfterStart')}</p>}
              {fieldHasError('deregAfterStart') && <p style={{ color: 'var(--dex-red)', fontSize: '0.8rem', marginTop: 8, marginBottom: 0 }}>{t('create.error.deregAfterStart')}</p>}
              </>)}
              </div>

              {/* v9.17: Reihenfolge umgestellt — Standard-Teilnehmerzahl
                  steht oben, Split-Toggle wird unter dem Block subtler
                  angezeigt. Die Mehrheit der Events nutzt nur eine
                  Gesamtkapazität; der B2Run-Sonderfall ist Opt-in. */}

              {/* v28.72: Bei einer Klammer war „Teilnehmerzahl & Warteliste"
                  nur ausgegraut — ohne Begründung und ohne Weg. Organizer
                  klickten ins Leere und hielten es für kaputt. Jetzt liegt eine
                  Erklär-Box darüber, die sagt WARUM es hier nicht gilt und mit
                  einem Klick ins erste Sub-Event führt, wo die Plätze
                  tatsächlich gepflegt werden. */}
              {subEventsOnlyMode && subEvents.length > 0 ? (
                /* v28.76: Vorher lag die Erklärung als Overlay über dem
                   ausgegrauten Abschnitt. Der ist eingeklappt aber nur rund
                   50 px hoch — die Box ragte unten aus der Karte heraus und
                   überdeckte die Weiter-Knöpfe. Jetzt steht sie schlank AN
                   der Stelle des Abschnitts, statt darüber zu schweben. */
                <div style={{
                  display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap',
                  padding: '10px 14px', marginBottom: 12, borderRadius: 8,
                  background: 'var(--dex-gray-50, #fafafa)',
                  border: '1px solid var(--dex-gray-200)',
                  borderLeft: '4px solid var(--dex-green, #86bc25)',
                }}>
                  <span style={{ fontSize: '0.82rem', lineHeight: 1.5, color: 'var(--dex-gray-700)', flex: 1, minWidth: 260 }}>
                    {isDe
                      ? <><strong>Plätze &amp; Warteliste</strong> werden pro {childTermSingular || 'Sub-Event'} vergeben — bei einer Klammer hätte eine Teilnehmerzahl hier keine Wirkung.</>
                      : <><strong>Seats &amp; waitlist</strong> are set per sub-event — for a bracket a capacity here would have no effect.</>}
                  </span>
                  <button
                    type="button"
                    className="btn btn-secondary"
                    style={{ fontSize: '0.8rem', padding: '6px 12px', flexShrink: 0 }}
                    onClick={() => setActiveCapacityTabIdx(1)}
                  >
                    {isDe
                      ? `Zu „${shortSubEventTitle(subEvents[0].title, title) || (childTermSingular || 'Sub-Event')}“`
                      : `Go to „${shortSubEventTitle(subEvents[0].title, title) || 'sub-event'}“`}
                  </button>
                </div>
              ) : (
                <>
              <div style={hauptGreyoutWrapperStyle()}>
              <div className="form-group" style={{ padding: '16px 20px', marginBottom: 12, background: zebraS3Bg(), borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
                {visHeader('vis_capacity', <StepBadge n={(locationFilter && audience) ? 23 : 22} />, isDe ? 'Teilnehmerzahl & Warteliste' : 'Capacity & waitlist')}
                {isVisOpen('vis_capacity') && (<>
              {/* v10.20: Geteilte Kapazität — generisch für beliebige Events.
                  Labels werden vom Organizer frei gewählt (z.B. "Vormittag /
                  Nachmittag", "VIP / Standard", "Lauf / Walk"). Default-Fallback
                  ist 'Durchstarter' / 'Funstarter' für Backward-Compat mit
                  B2Run-Events vor v10.20. */}
              {useSplitCapacities ? (
                <div style={{ padding: 16, background: 'rgba(134,188,37,0.08)', borderRadius: 'var(--dex-radius, 12px)', border: '1px solid var(--dex-green)', marginBottom: 16 }}>
                  <label className="form-label" style={{ marginBottom: 4 }}>
                    {isDe ? 'Geteilte Kapazität' : 'Split capacity'}
                    <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> zwei getrennte Kapazitäten innerhalb des Events. Beispiele: <strong>Vormittag / Nachmittag</strong>, <strong>VIP / Standard</strong>, <strong>Lauf 5 km / Lauf 10 km</strong>. Du legst die <strong>Bezeichnungen frei fest</strong> und vergibst pro Gruppe eine eigene Platzzahl mit eigener Warteliste.<br /><br />
                      <strong>Anzeige in der App:</strong> Teilnehmer sehen auf der Anmelde-Seite <strong>zwei nebeneinanderstehende Boxen</strong> mit deinen Bezeichnungen, jeweils mit &bdquo;X / Y Plätze frei&ldquo;. Die Wahl ist <strong>verpflichtend</strong>, bevor man auf Anmelden klickt.<br /><br />
                      <strong>Automatismen:</strong> ist eine der zwei Gruppen voll, kommen weitere Anmeldungen <strong>nur in die Warteliste dieser Gruppe</strong> — nicht in die andere. Beim Nachrücken bleibt der Typ <strong>erhalten</strong> (eine VIP-Wartelisten-Person rückt nicht in einen Standard-Platz).<br /><br />
                      <strong>Empfehlung:</strong> nur verwenden, wenn die zwei Gruppen <strong>wirklich getrennt</strong> behandelt werden sollen (eigenes Catering, eigener Bus, eigener Slot beim Veranstalter). Bei einer einfachen Gesamtkapazität reicht die Standard-Teilnehmerzahl unten.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> two separate capacities within the event. Examples: <strong>morning / afternoon</strong>, <strong>VIP / standard</strong>, <strong>5 km run / 10 km run</strong>. You <strong>name the two groups freely</strong> and give each its own seat count and waitlist.<br /><br />
                      <strong>Shown in the app:</strong> attendees see <strong>two side-by-side boxes</strong> on the registration page, each labelled with your text and showing &ldquo;X / Y seats free&rdquo;. Picking one is <strong>required</strong> before submitting.<br /><br />
                      <strong>Automation:</strong> when one group is full, further sign-ups land <strong>only on that group&apos;s waitlist</strong> — not the other. When promoting from waitlist, the <strong>group is preserved</strong> (a waitlisted VIP is not auto-promoted into a standard slot).<br /><br />
                      <strong>Tip:</strong> only use this when the two groups really need to be <strong>handled separately</strong> (own catering, own bus, separate slot with the supplier). For a single overall capacity the standard attendee count below is enough.
                    </>
                  )} />
                  </label>
                  <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 0, marginBottom: 12 }}>
                    {isDe
                      ? 'Vergib pro Gruppe eine eigene Bezeichnung und Platzzahl. Die Bezeichnungen erscheinen auf der Anmeldeseite als zwei Auswahl-Boxen.'
                      : 'Give each group its own name and seat count. The names appear on the registration page as two selectable boxes.'}
                  </p>
                  {/* v26.83: frei wählbare Überschrift der Gruppen-Auswahl auf der
                      Anmeldeseite (statt „Gruppen-Auswahl"). */}
                  <div style={{ marginBottom: 12 }}>
                    <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, marginBottom: 4 }}>
                      {isDe ? 'Überschrift der Auswahl (optional)' : 'Selection heading (optional)'}
                    </label>
                    <input
                      className="form-input"
                      type="text"
                      value={splitSectionTitle}
                      onChange={e => setSplitSectionTitle(e.target.value)}
                      placeholder={isDe ? 'Leer = „Gruppen-Auswahl". z.B. „Auswahl Räume", „Auswahl Laufgruppe"' : 'Empty = "Group selection". e.g. "Choose room", "Choose run group"'}
                      maxLength={60}
                    />
                  </div>
                  {/* v10.20: zwei Text-Inputs für die frei wählbaren Bezeichnungen.
                      Wenn der Organizer nichts einträgt, fällt die Registration-
                      Seite auf 'Durchstarter' / 'Funstarter' zurück. */}
                  <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginBottom: 12 }}>
                    <div>
                      <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, marginBottom: 4 }}>
                        {isDe ? 'Bezeichnung Gruppe A' : 'Group A label'}
                      </label>
                      <input
                        className="form-input"
                        type="text"
                        value={splitLabelA}
                        onChange={e => setSplitLabelA(e.target.value)}
                        placeholder={isDe ? 'z.B. Vormittag, VIP, Durchstarter' : 'e.g. morning, VIP, starter'}
                        maxLength={40}
                      />
                    </div>
                    <div>
                      <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, marginBottom: 4 }}>
                        {isDe ? 'Bezeichnung Gruppe B' : 'Group B label'}
                      </label>
                      <input
                        className="form-input"
                        type="text"
                        value={splitLabelB}
                        onChange={e => setSplitLabelB(e.target.value)}
                        placeholder={isDe ? 'z.B. Nachmittag, Standard, Funstarter' : 'e.g. afternoon, standard, fun'}
                        maxLength={40}
                      />
                    </div>
                  </div>
                  {/* v26.72: optionale Beschreibung pro Gruppe — erscheint unter
                      dem Gruppen-Namen in der Auswahl-Karte auf der Anmeldeseite. */}
                  <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginBottom: 12 }}>
                    <div>
                      <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, marginBottom: 4 }}>
                        {isDe ? 'Beschreibung Gruppe A (optional)' : 'Group A description (optional)'}
                      </label>
                      <textarea
                        className="form-input"
                        rows={2}
                        value={splitDescA}
                        onChange={e => setSplitDescA(e.target.value)}
                        placeholder={isDe ? 'Kurzer Zusatztext, z.B. „inkl. Mittagessen"' : 'Short note, e.g. "incl. lunch"'}
                        maxLength={400}
                        style={{ resize: 'vertical' }}
                      />
                    </div>
                    <div>
                      <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, marginBottom: 4 }}>
                        {isDe ? 'Beschreibung Gruppe B (optional)' : 'Group B description (optional)'}
                      </label>
                      <textarea
                        className="form-input"
                        rows={2}
                        value={splitDescB}
                        onChange={e => setSplitDescB(e.target.value)}
                        placeholder={isDe ? 'Kurzer Zusatztext, z.B. „ohne Mittagessen"' : 'Short note, e.g. "without lunch"'}
                        maxLength={400}
                        style={{ resize: 'vertical' }}
                      />
                    </div>
                  </div>
                  {/* v26.83: frei wählbarer Hinweistext über der Gruppen-Auswahl
                      auf der Anmeldeseite. Leer = Standardsatz. */}
                  <div style={{ marginBottom: 12 }}>
                    <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, marginBottom: 4 }}>
                      {isDe ? 'Hinweistext über der Gruppen-Auswahl (optional)' : 'Help text above the group selection (optional)'}
                    </label>
                    <textarea
                      className="form-input"
                      rows={2}
                      value={splitHelpText}
                      onChange={e => setSplitHelpText(e.target.value)}
                      placeholder={isDe ? 'Leer = Standard: „Wähle eine der zwei Gruppen aus. Ist die Wunsch-Gruppe voll, kannst du automatisch in die andere wechseln oder auf der Warteliste warten."' : 'Empty = default: "Pick one of the two groups. If your preferred group is full, you can switch to the other automatically or wait on the waitlist."'}
                      maxLength={600}
                      style={{ resize: 'vertical', width: '100%' }}
                    />
                    <p style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', margin: '4px 0 0' }}>
                      {isDe ? 'Dieser Text steht auf der Anmeldeseite direkt über den beiden Gruppen-Boxen. Leer lassen für den Standardtext.' : 'This text appears on the registration page right above the two group boxes. Leave empty for the default.'}
                    </p>
                  </div>
                  <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                    <div>
                      <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, marginBottom: 4 }}>
                        <Icon iconName="People" style={{ fontSize: 14, marginRight: 6, color: 'var(--dex-green-dark, #6b9a1e)' }} />
                        {isDe ? 'Plätze' : 'Seats'} {splitLabelA.trim() || (isDe ? 'Gruppe A' : 'Group A')}
                      </label>
                      <input
                        className="form-input"
                        type="number"
                        min={0}
                        value={durchstarterCapacity}
                        onChange={e => setDurchstarterCapacity(e.target.value)}
                        placeholder="z.B. 10"
                      />
                    </div>
                    <div>
                      <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, marginBottom: 4 }}>
                        <Icon iconName="People" style={{ fontSize: 14, marginRight: 6, color: 'var(--dex-orange, #ff8c00)' }} />
                        {isDe ? 'Plätze' : 'Seats'} {splitLabelB.trim() || (isDe ? 'Gruppe B' : 'Group B')}
                      </label>
                      <input
                        className="form-input"
                        type="number"
                        min={0}
                        value={funstarterCapacity}
                        onChange={e => setFunstarterCapacity(e.target.value)}
                        placeholder="z.B. 90"
                      />
                    </div>
                  </div>
                  <div style={{ marginTop: 12, padding: '8px 12px', background: '#fff', borderRadius: 8, fontSize: '0.85rem' }}>
                    <strong>{t('create.b2runcap.total')}:</strong> {((parseInt(durchstarterCapacity, 10) || 0) + (parseInt(funstarterCapacity, 10) || 0))} {t('create.b2runcap.seats')}
                  </div>

                  {/* v10.20: Waitlist-Modus bei Split-Capacity. Default
                      'separate' (zwei Wartelisten) — entspricht dem alten
                      B2Run-Verhalten und ist die typ-bewusste Variante.
                      Alternative 'shared' = eine gemeinsame Warteliste, FIFO
                      über beide Gruppen. Sinnvoll wenn die Gruppen
                      organisatorisch fluide sind (z.B. Vormittag/Nachmittag
                      bei einem Workshop, wo der nächste freie Slot egal ist
                      welche Gruppe). Nur sichtbar wenn Warteliste aktiviert
                      ist (Toggle weiter unten). */}
                  {waitlistEnabled && (
                    <div style={{ marginTop: 12, padding: '12px 14px', background: '#fff', borderRadius: 8 }}>
                      <label className="form-label" style={{ marginBottom: 6 }}>
                        {isDe ? 'Warteliste-Verhalten' : 'Waitlist behaviour'}
                        <InfoTooltip text={isDe ? (
                          <>
                            <strong>Was du hier einstellst:</strong> ob bei einer Anmeldung über die jeweilige Kapazität hinaus eine <strong>gemeinsame</strong> oder zwei <strong>getrennte</strong> Wartelisten greifen.<br /><br />
                            <strong>Getrennt (Default):</strong> jede Gruppe hat ihre eigene Warteliste. Wer auf der {splitLabelA.trim() || 'Gruppe A'}-Warteliste landet, rückt nur in einen frei werdenden {splitLabelA.trim() || 'Gruppe A'}-Platz nach. Saubere Trennung — sinnvoll wenn die zwei Gruppen wirklich unterschiedliche Slots beim Veranstalter, eigenes Catering oder eigenen Bus haben.<br /><br />
                            <strong>Gemeinsam:</strong> alle Wartelistler stehen in einer einzigen Schlange. Wer am längsten wartet, rückt zuerst nach — egal in welche Gruppe der frei werdende Platz gehört. Sinnvoll wenn die Gruppen-Wahl nur eine UI-Komfort-Sache ist (z.B. Vormittag/Nachmittag-Slot bei einem Workshop) und der Organizer sich nicht um Typen kümmern will.<br /><br />
                            <strong>Auswirkung für Teilnehmer:</strong> bei <strong>getrennt</strong> kann es passieren, dass jemand in der einen Schlange weiter hinten steht, obwohl die andere Gruppe leer ist — dann muss man <strong>aktiv umsteigen</strong> (über den Fallback-Dialog beim nächsten Versuch). Bei <strong>gemeinsam</strong> rutscht jeder hoch sobald irgendwo ein Platz frei wird.
                          </>
                        ) : (
                          <>
                            <strong>What you set here:</strong> whether sign-ups beyond the per-group capacity land on <strong>one shared</strong> or <strong>two separate</strong> waitlists.<br /><br />
                            <strong>Separate (default):</strong> each group has its own waitlist. Someone on the {splitLabelA.trim() || 'group A'} waitlist only moves up into a freed {splitLabelA.trim() || 'group A'} seat. Clean separation — useful when the two groups have genuinely different supplier slots, own catering, own bus.<br /><br />
                            <strong>Shared:</strong> all waitlisters stand in one queue. Whoever has waited longest moves up first — regardless of which group the freed seat belongs to. Useful when the group split is just a UI convenience (e.g. morning / afternoon slot at a workshop) and the organizer does not want to manage types.<br /><br />
                            <strong>Effect for attendees:</strong> with <strong>separate</strong> someone may be further back in their queue while the other group is empty — they then have to <strong>actively switch</strong> (via the fallback dialog at next attempt). With <strong>shared</strong> everyone moves up as soon as a seat opens anywhere.
                          </>
                        )} />
                      </label>
                      <div style={{ display: 'flex', gap: 16, marginTop: 4 }}>
                        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, cursor: 'pointer', fontSize: '0.85rem' }}>
                          <input
                            type="radio"
                            name="splitWaitlistMode"
                            checked={!splitSharedWaitlist}
                            onChange={() => setSplitSharedWaitlist(false)}
                          />
                          {isDe ? 'Getrennte Wartelisten pro Gruppe' : 'Separate waitlist per group'}
                        </label>
                        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, cursor: 'pointer', fontSize: '0.85rem' }}>
                          <input
                            type="radio"
                            name="splitWaitlistMode"
                            checked={!!splitSharedWaitlist}
                            onChange={() => setSplitSharedWaitlist(true)}
                          />
                          {isDe ? 'Eine gemeinsame Warteliste' : 'One shared waitlist'}
                        </label>
                      </div>
                    </div>
                  )}

                  {/* v11.25: Display-Reihenfolge der zwei Gruppen-Karten in der
                      Registrierungs-UI umkehren. Reine Anzeige-Toggle —
                      splitLabelA/B, Kapazitäten und die internen StarterType-IDs
                      bleiben unangetastet. Nur bei aktiver Split-Capacity
                      sinnvoll. */}
                  {useSplitCapacities && (
                    <div style={{ marginTop: 12, padding: '12px 14px', background: '#fff', borderRadius: 8 }}>
                      <label style={{ display: 'inline-flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer', fontSize: '0.85rem' }}>
                        <input
                          type="checkbox"
                          checked={splitDisplayOrderReversed}
                          onChange={e => setSplitDisplayOrderReversed(e.target.checked)}
                          style={{ marginTop: 3 }}
                        />
                        <span>
                          <strong>{isDe ? 'Reihenfolge in der Registrierung umkehren' : 'Reverse order on registration page'}</strong>
                          <InfoTooltip text={isDe ? (
                            <>
                              <strong>Was du hier einstellst:</strong> ob die zwei Gruppen-Karten (&bdquo;{splitLabelA.trim() || 'Gruppe A'}&ldquo; und &bdquo;{splitLabelB.trim() || 'Gruppe B'}&ldquo;) in der Registrierungs-Maske in der <strong>aktuellen</strong> oder in <strong>umgekehrter</strong> Reihenfolge angezeigt werden.<br /><br />
                              <strong>Anzeige in der App:</strong> aus = &bdquo;{splitLabelA.trim() || 'Gruppe A'}&ldquo; links, &bdquo;{splitLabelB.trim() || 'Gruppe B'}&ldquo; rechts. An = umgekehrt: &bdquo;{splitLabelB.trim() || 'Gruppe B'}&ldquo; links, &bdquo;{splitLabelA.trim() || 'Gruppe A'}&ldquo; rechts. Gilt für Registrierung und für die Kapazitäts-Übersicht im Admin-Center.<br /><br />
                              <strong>Auswirkung für Teilnehmer:</strong> rein optisch — ändert nichts an den Plätzen, Wartelisten oder bestehenden Anmeldungen. Nur eine andere Reihenfolge der Auswahl-Buttons.<br /><br />
                              <strong>Hinweis:</strong> wenn du tatsächlich &bdquo;Gruppe 2&ldquo; zur prominenteren machen willst, ohne deine Labels und Kapazitäten zu vertauschen, ist <strong>dieses Häkchen</strong> der saubere Weg. Manuelles Vertauschen von Label A ↔ B + Kapazitäten zerschießt die Verbindung zu den existierenden Anmeldungen.
                            </>
                          ) : (
                            <>
                              <strong>What you set here:</strong> whether the two group cards (&bdquo;{splitLabelA.trim() || 'group A'}&ldquo; and &bdquo;{splitLabelB.trim() || 'group B'}&ldquo;) appear in the <strong>current</strong> or <strong>reversed</strong> order on the registration page.<br /><br />
                              <strong>Shown in the app:</strong> off = &bdquo;{splitLabelA.trim() || 'group A'}&ldquo; left, &bdquo;{splitLabelB.trim() || 'group B'}&ldquo; right. On = reversed.<br /><br />
                              <strong>Effect for attendees:</strong> purely visual — does not affect seats, waitlists or existing registrations. Just a different order of selection buttons.<br /><br />
                              <strong>Note:</strong> if you want to make &bdquo;group 2&ldquo; the more prominent one without swapping your labels and capacities, this checkbox is the clean way. Manually swapping label A ↔ B + capacities breaks the link to existing registrations.
                            </>
                          )} />
                        </span>
                      </label>
                    </div>
                  )}

                  {/* v6.15: Starter-Typ → Startblock-Zuordnung (optional).
                      Nur sinnvoll wenn Startblocks definiert sind (Reiter "Event-spezifische Felder").
                      Wenn gesetzt, wird der Startblock bei der Registrierung automatisch
                      anhand des gewählten Starter-Typs gesetzt — der User muss den Block
                      nicht extra auswählen. */}
                  {b2runStartblocks.length > 0 && (
                    <div style={{ marginTop: 12, padding: '10px 12px', background: '#fff', borderRadius: 8 }}>
                      <label className="form-label" style={{ marginBottom: 6 }}>
                        {t('create.b2runcap.starterblock.title') || 'Starter-Typ → Startblock-Zuordnung'}
                        <InfoTooltip text={isDe ? (
                          <>
                            <strong>Was du hier einstellst:</strong> eine <strong>fixe Zuordnung Starter-Typ → Startblock</strong>. Beispiel: Durchstarter immer Block A, Funstarter immer Block C.<br /><br />
                            <strong>Anzeige in der App:</strong> Teilnehmer wählen <strong>nur</strong> den Starter-Typ — der Startblock wird automatisch gesetzt, der Block-Selector verschwindet aus der Anmelde-Maske. Eine Frage weniger für den User.<br /><br />
                            <strong>Automatismen:</strong> die Startblock-Kapazität wird gegen den jeweiligen Starter-Typ gerechnet — freie Plätze pro Typ = freie Plätze im zugeordneten Block.<br /><br />
                            <strong>Leer:</strong> Teilnehmer wählen Starter-Typ <em>und</em> Startblock manuell — der Organizer hat dann gemischte Blöcke und muss bei Engpässen selbst zuteilen.
                          </>
                        ) : (
                          <>
                            <strong>What you set here:</strong> a <strong>fixed mapping starter type → start block</strong>. Example: Durchstarter always block A, Funstarter always block C.<br /><br />
                            <strong>Shown in the app:</strong> attendees only pick the <strong>starter type</strong> — the start block is set automatically, the block selector disappears from the form. One question less for the user.<br /><br />
                            <strong>Automation:</strong> the start-block capacity counts against the matching starter type — free slots per type = free slots in the assigned block.<br /><br />
                            <strong>Empty:</strong> attendees pick starter type <em>and</em> start block manually — the organizer ends up with mixed blocks and has to redistribute manually when slots get tight.
                          </>
                        )} />
                      </label>
                      <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                        <div>
                          <label style={{ display: 'block', fontSize: '0.78rem', marginBottom: 4, color: 'var(--dex-gray-600)' }}>Durchstarter →</label>
                          <select
                            className="form-select"
                            value={durchstarterStartblock}
                            onChange={e => setDurchstarterStartblock(e.target.value)}
                          >
                            <option value="">{t('create.b2runcap.starterblock.none') || '— kein automatischer Block —'}</option>
                            {b2runStartblocks.map(b => <option key={b} value={b}>{b}</option>)}
                          </select>
                        </div>
                        <div>
                          <label style={{ display: 'block', fontSize: '0.78rem', marginBottom: 4, color: 'var(--dex-gray-600)' }}>Funstarter →</label>
                          <select
                            className="form-select"
                            value={funstarterStartblock}
                            onChange={e => setFunstarterStartblock(e.target.value)}
                          >
                            <option value="">{t('create.b2runcap.starterblock.none') || '— kein automatischer Block —'}</option>
                            {b2runStartblocks.map(b => <option key={b} value={b}>{b}</option>)}
                          </select>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* v10.24: Leistungsnachweis-Pflicht-Toggle wurde entfernt.
                      Stattdessen kann der Organizer in Schritt 5 (Felder) ein
                      eigenes Pflichtfeld vom Typ Checkbox anlegen und es über
                      'Sichtbar für Teilnehmergruppe → Nur Gruppe A' gezielt
                      auf eine Split-Gruppe einschränken. Das ersetzt den
                      hartkodierten B2Run-Leistungsnachweis-Sonderfall durch
                      ein generisches Pro-Gruppe-Feld-Konzept. Hinweis steht
                      hier, damit der Organizer beim Migrieren weiß wo das
                      Feature jetzt liegt. */}
                  <div style={{ marginTop: 12, padding: '10px 12px', background: 'rgba(21,101,192,0.06)', border: '1px solid rgba(21,101,192,0.4)', borderRadius: 8, fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                    <div style={{ fontWeight: 700, color: 'var(--dex-blue, #1565c0)', marginBottom: 4, display: 'flex', alignItems: 'center', gap: 6 }}>
                      <Icon iconName="Info" style={{ fontSize: 16 }} />
                      {isDe ? 'Pflichtfelder pro Gruppe' : 'Required fields per group'}
                    </div>
                    {isDe ? (
                      <>
                        Du möchtest für eine der zwei Gruppen ein zusätzliches Pflichtfeld einblenden — z.B. eine Checkbox &bdquo;Leistungsnachweis vorhanden&ldquo; nur für die Gruppe der schnellen Läufer? Lege das Feld in <strong>Schritt 5 (Felder)</strong> an und stelle dort den Selector <strong>&bdquo;Sichtbar für Teilnehmergruppe&ldquo;</strong> auf <strong>&bdquo;Nur {(splitLabelA || '').trim() || 'Gruppe A'}&ldquo;</strong> bzw. <strong>&bdquo;Nur {(splitLabelB || '').trim() || 'Gruppe B'}&ldquo;</strong>. Das Feld wird dann in der Anmeldung dynamisch ein- oder ausgeblendet, sobald der Teilnehmer eine der zwei Boxen anklickt.
                      </>
                    ) : (
                      <>
                        Want to show an extra required field only for one of the two groups — e.g. a checkbox &ldquo;Performance proof available&rdquo; just for the fast-runner group? Add the field in <strong>step 5 (Fields)</strong> and set the <strong>&ldquo;Visible for attendee group&rdquo;</strong> selector there to <strong>&ldquo;{(splitLabelA || '').trim() || 'Group A'} only&rdquo;</strong> or <strong>&ldquo;{(splitLabelB || '').trim() || 'Group B'} only&rdquo;</strong>. The field will then be shown or hidden dynamically as the attendee picks one of the two boxes.
                      </>
                    )}
                  </div>

                  <div style={{ marginTop: 12 }}>
                    <label className="form-label" style={{ marginBottom: 4 }}>
                      {t('create.waitlist')}
                      <InfoTooltip text={isDe ? (
                        <>
                          <strong>Wartelisten für Split-Kapazitäten</strong> — bei aktivierter Warteliste hat <strong>jeder Starter-Typ seine eigene Liste</strong>. Durchstarter-Anmeldungen über der Durchstarter-Kapazität landen auf der Durchstarter-Warteliste, dasselbe für Funstarter. Beim Nachrücken wird <strong>typ-bewusst</strong> befördert — der älteste Durchstarter-Eintrag rückt in einen frei werdenden Durchstarter-Platz, kein Mix.
                        </>
                      ) : (
                        <>
                          <strong>Waitlists for split capacities</strong> — when the waitlist is on, <strong>each starter type has its own queue</strong>. Durchstarter sign-ups beyond the Durchstarter capacity go on the Durchstarter waitlist, same for Funstarter. Promotion is <strong>type-aware</strong> — the oldest Durchstarter waiting moves into a freed Durchstarter slot, no mixing.
                        </>
                      )} />
                    </label>
                    <div className="toggle-wrapper" style={{ marginTop: 4 }}>
                      <label className="toggle">
                        <input type="checkbox" checked={waitlistEnabled} onChange={e => setWaitlistEnabled(e.target.checked)} />
                        <span className="toggle-slider" />
                      </label>
                      <span style={{ fontSize: '0.9rem' }}>{waitlistEnabled ? t('create.enabled') : t('create.disabled')}</span>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                  <div className="form-group">
                    <label className="form-label">
                      {t('create.maxparticipants')}
                      <InfoTooltip text={isDe ? (
                        <>
                          <strong>Was du hier einstellst:</strong> die <strong>maximale Teilnehmerzahl</strong> oder den Modus <strong>Unbegrenzt</strong>.<br /><br />
                          <strong>Anzeige in der App:</strong> ist die Kapazität voll, sehen Teilnehmer den <strong>roten Banner Alle Plätze sind belegt</strong> auf der Anmelde-Seite. Im Admin Center wird ein Auslastungs-KPI live mitgeführt.<br /><br />
                          <strong>Automatismen:</strong> ist <strong>Warteliste</strong> aktiv und die Kapazität voll, werden neue Anmeldungen automatisch auf <strong>Status Warteliste</strong> gesetzt und bekommen die Wartelisten-Bestätigungs-Mail. Bei einer Abmeldung rückt der älteste Wartelisten-Eintrag automatisch nach und bekommt Nachrück-Mail + Outlook-Termin.<br /><br />
                          <strong>Modus Unbegrenzt:</strong> keine Auslastungs-Anzeige, keine Warteliste — wird typischerweise für interne All-Hands-Mails oder reine Info-Events verwendet.
                        </>
                      ) : (
                        <>
                          <strong>What you set here:</strong> the <strong>maximum attendee count</strong> or the <strong>Unlimited</strong> mode.<br /><br />
                          <strong>Shown in the app:</strong> when full, attendees see the <strong>red banner All spots are taken</strong> on the registration page. The admin center shows live capacity KPIs.<br /><br />
                          <strong>Automation:</strong> if <strong>waitlist</strong> is on and capacity is full, new sign-ups land in <strong>status Waitlist</strong> and get a waitlist confirmation mail. On cancellation, the oldest waitlist entry is auto-promoted and receives a promotion mail + Outlook event.<br /><br />
                          <strong>Unlimited mode:</strong> no capacity indicator, no waitlist — typically used for internal all-hands or info-only events.
                        </>
                      )} />
                    </label>
                    {/* v22.38: Checkbox statt Schieberegler (konsistent zu den
                        anderen runden Abfragen) — Frage invertiert:
                        „Teilnehmeranzahl begrenzen?" (Default: nein =
                        unbegrenzt). Bei Ja erscheint darunter das Anzahl-Feld
                        und die Warteliste springt automatisch auf an. */}
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', marginTop: 4, marginBottom: 8 }}>
                      <input
                        type="checkbox"
                        checked={!unlimitedParticipants}
                        onChange={e => {
                          const limited = e.target.checked;
                          setUnlimitedParticipants(!limited);
                          if (limited) {
                            setWaitlistEnabled(true);
                          } else {
                            setMaxParticipants('');
                            setWaitlistEnabled(false);
                          }
                        }}
                        style={{ marginTop: 3, cursor: 'pointer' }}
                      />
                      <span style={{ fontSize: '0.9rem' }}>
                        <strong>{isDe ? 'Teilnehmeranzahl begrenzen?' : 'Limit the number of participants?'}</strong>
                        <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                          {isDe
                            ? 'Default: nein — unbegrenzte Teilnehmerzahl. Wenn aktiviert, legst du unten die maximale Platzzahl fest; die Warteliste wird automatisch aktiviert.'
                            : 'Default: no — unlimited participants. When enabled, set the maximum number of seats below; the waitlist is enabled automatically.'}
                        </span>
                      </span>
                    </label>
                    {!unlimitedParticipants && (
                      <>
                        <input
                          className="form-input"
                          type="number"
                          min={0}
                          value={maxParticipants}
                          onChange={e => setMaxParticipants(e.target.value.replace(/[^0-9]/g, ''))}
                          placeholder="Anzahl"
                          style={errorBorderStyle('maxParticipants')}
                        />
                        {fieldHasError('maxParticipants') && <span style={{ color: 'var(--dex-red)', fontSize: '0.75rem' }}>{t('create.error.required')}</span>}
                      </>
                    )}
                  </div>
                  {!unlimitedParticipants && (
                    <div className="form-group">
                      <label className="form-label">
                        {t('create.waitlist')}
                        <InfoTooltip text={isDe ? (
                          <>
                            <strong>Was du hier einstellst:</strong> ob bei vollem Event eine <strong>Warteliste</strong> akzeptiert wird oder neue Anmeldungen sofort blockiert werden.<br /><br />
                            <strong>Anzeige in der App:</strong> bei aktiver Warteliste können Teilnehmer sich auch über die Kapazitäts-Grenze hinaus anmelden — bekommen Status <strong>Warteliste</strong> mit Positions-Nummer. Im Admin Center erscheint eine eigene <strong>Warteliste-Kachel</strong>.<br /><br />
                            <strong>Automatismen:</strong> Wartelisten-Anmeldungen bekommen die <strong>Wartelisten-Bestätigungs-Mail</strong>. Sobald jemand absagt, rückt der älteste Wartelisten-Eintrag automatisch nach (<strong>First-In, First-Out</strong>) — bekommt eine <strong>Nachrück-Mail</strong> mit Outlook-Termin und der Status wechselt auf Angemeldet.<br /><br />
                            <strong>Auswirkung für Teilnehmer:</strong> sie sehen ihre Position auf der Warteliste auf der Anmelde-Seite und werden automatisch informiert, wenn ein Platz frei wird.<br /><br />
                            <strong>Aus:</strong> bei vollem Event ist der Anmelde-Button gesperrt — neue Interessenten müssen den Organizer direkt kontaktieren.
                          </>
                        ) : (
                          <>
                            <strong>What you set here:</strong> whether full events accept a <strong>waitlist</strong> or new registrations are blocked immediately.<br /><br />
                            <strong>Shown in the app:</strong> when waitlist is on, attendees can register past the capacity limit — they get status <strong>Waitlist</strong> with a position number. The admin center shows a dedicated <strong>waitlist tile</strong>.<br /><br />
                            <strong>Automation:</strong> waitlist sign-ups receive the <strong>waitlist confirmation mail</strong>. As soon as someone cancels, the oldest entry is auto-promoted (<strong>first-in, first-out</strong>) — they receive a <strong>promotion mail</strong> with Outlook event and their status flips to Registered.<br /><br />
                            <strong>Effect for attendees:</strong> they see their waitlist position on the registration page and are notified automatically when a spot frees up.<br /><br />
                            <strong>Off:</strong> when capacity is full, the register button is locked — new interested people have to contact the organizer directly.
                          </>
                        )} />
                      </label>
                      <div className="toggle-wrapper" style={{ marginTop: 8 }}>
                        <label className="toggle">
                          <input type="checkbox" checked={waitlistEnabled} onChange={e => setWaitlistEnabled(e.target.checked)} />
                          <span className="toggle-slider" />
                        </label>
                        <span style={{ fontSize: '0.9rem' }}>{waitlistEnabled ? t('create.enabled') : t('create.disabled')}</span>
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* v9.17: Split-Capacity-Toggle UNTER dem Teilnehmerzahl-Block,
                  bewusst subtil — der Großteil der Events nutzt eine einzige
                  Teilnehmerzahl. Nur wer einen Lauf mit getrennten Starter-
                  Töpfen anlegt, klickt diesen Toggle.
                  v24.4 (K): erst sichtbar, wenn „Teilnehmeranzahl begrenzen"
                  aktiv ist (bzw. die geteilte Kapazität bereits an ist) —
                  ohne Begrenzung ergibt eine geteilte Kapazität keinen Sinn. */}
              {(!unlimitedParticipants || useSplitCapacities) && (
              <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, fontSize: '0.8rem', color: 'var(--dex-gray-600)', cursor: 'pointer', padding: '8px 0', marginTop: 4 }}>
                <input
                  type="checkbox"
                  checked={useSplitCapacities}
                  onChange={e => setUseSplitCapacities(e.target.checked)}
                  style={{ marginTop: 2, cursor: 'pointer' }}
                />
                <span>
                  {t('create.splitcap.label')}
                  <span style={{ display: 'block', fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 1 }}>
                    {t('create.splitcap.hint')}
                  </span>
                </span>
              </label>
              )}
              </>)}
              </div>

              </div>{/* v15.6: close hauptGreyoutWrapperStyle div (Step 4) */}
                </>
              )}{/* v28.76: Ende Klammer-Fall / Normalfall */}
              </div>{/* v15.0: close activeCapacityTabIdx===0 wrapper (Top-Level Sichtbarkeit/Deadlines/Max/Split) */}

              </div>
  );
};
