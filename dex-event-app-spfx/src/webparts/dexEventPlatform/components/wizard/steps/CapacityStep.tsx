/* CapacityStep — aus EventCreationPage.tsx ausgelagert (Zeilen 12931-14831 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; einzige
 * Aenderung ist die Anzeige-Bedingung: aus `currentStep === 3` wurde das Prop `visible`.
 * `visible` schaltet display:none statt unmount — Eingaben ueberleben den
 * Schrittwechsel genauso wie vorher.
 *
 * v31.2: Oberfläche nach docs/ui-leitfaden.md umgebaut. Reihenfolge auf
 * Klammer UND Sub-Event-Reiter jetzt Plätze & Warteliste → Fristen →
 * Sichtbarkeit (vorher umgekehrt, historisch gewachsen); Ein/Aus als
 * dex-ui-toggle-row/-switch, Alternativen (unbegrenzt/begrenzt, ODER/UND,
 * getrennte/gemeinsame Warteliste) als dex-ui-choice-Kacheln. Handler,
 * Bedingungen, Speicher-Semantik (maxParticipants bleibt bei geteilten
 * Gruppen 0) sind unverändert. Die StepBadge-Nummern folgen seit der
 * Umnummerierung (07.09.2026) wieder der Renderreihenfolge (18–23). */
import * as React from 'react';
import { SubEventDraft } from '../../wizard/wizardTypes';
import { Icon } from '@fluentui/react/lib/Icon';
import { StepBadge } from '../../wizard/StepBadge';
import { LocationMultiSelect } from '../../wizard/LocationMultiSelect';
import AudiencePicker from '../../AudiencePicker';
import DatePicker from 'react-datepicker';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { isoToLocal } from '../../../utils/berlinTime'; // v30.67
import { InfoTooltip } from '../../InfoTooltip';
// v31.2: Gemeinsame UI-Klassen (Kacheln, Schalter-Zeilen, Aufklapper) und
// die Symbole dafür — siehe docs/ui-leitfaden.md.
import { cx } from '../../dexUi';
import { AlertCircle, Check, ChevronDown, Info, Users } from '../../Icons';
import WizardHint from '../../WizardHint';

/** v30.67: Sub-Event-Fristen liegen als UTC-ISO vor. Der Picker zeigte sie
 *  über `new Date(iso)` in der BROWSER-Zeitzone, der Schreibweg (onChange →
 *  berlinLocalToUtcIso) las die Wanduhr aber als BERLINER Zeit. Auf einem
 *  UTC-/London-Rechner (Citrix, Reise) hieß das: Anzeige 21:59 statt 23:59,
 *  und jedes Anfassen des Feldes verschob die Frist um den Offset nach vorn.
 *  Hin- und Rückweg müssen dieselbe Zeitzone nehmen — wie Start/Ende in
 *  Schritt 1 (`subIsoToDate`): Berliner Wanduhr als lokales Date. */
const berlinIsoToPickerDate = (iso: string | undefined): Date | null => {
  const local = iso ? isoToLocal(iso) : '';
  if (!local) return null;
  const d = new Date(local);
  return isNaN(d.getTime()) ? null : d;
};
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
  // v31.2: Einziger lokaler Zustand des Schritts — der Aufklapper für die
  // Feineinstellungen der geteilten Gruppen (Texte, Reihenfolge, Startblöcke).
  // Steht unbedingt VOR dem return, damit die Hook-Reihenfolge fest ist.
  const [splitMoreOpen, setSplitMoreOpen] = React.useState(false);
  return (
              <div style={{ display: visible ? 'block' : 'none' }}>
              <h2 className="dex-step-head-title">
                <span className="dex-step-eyebrow">{isDe ? 'Schritt 4 von 9' : 'Step 4 of 9'}</span>
                {isDe ? 'Kapazität, Fristen & Sichtbarkeit' : 'Capacity, deadlines & visibility'}
              </h2>
              <p className="dex-step-head-lead">
                {isDe
                  ? 'Wie viele Plätze gibt es, bis wann kann man sich an- und abmelden — und wer sieht das Event überhaupt? Drei Fragen, in dieser Reihenfolge.'
                  : 'How many seats are there, until when can people register and cancel — and who gets to see the event at all? Three questions, in that order.'}
              </p>
              {/* v31.2: Reihenfolge der Einführung folgt der Seite (Plätze →
                  Warteliste → Fristen → Sichtbarkeit). */}
              {renderStepIntro(
                [
                  'Maximale Teilnehmerzahl festlegen (oder Unbegrenzt)',
                  'Warteliste aktivieren — voll besetzte Events nehmen weitere Anmeldungen auf, bis ein Platz frei wird',
                  'Optional: Geteilte Kapazität — zwei frei benannte Gruppen mit eigener Platzzahl + eigener oder gemeinsamer Warteliste',
                  'Anmelde-Deadline + letzte Abmeldemöglichkeit (vorbefüllt anhand des Event-Datums, jederzeit überschreibbar)',
                  'Sichtbarkeit: Standort-Filter und Mailverteiler/User festlegen — wer das Event in der Liste sieht',
                ],
                [
                  'Set the maximum number of attendees (or Unlimited)',
                  'Enable waitlist — full events accept new registrations and promote them once a spot frees up',
                  'Optional: split capacity — two freely-named groups with own seat count + own or shared waitlist',
                  'Registration deadline + last cancellation date (pre-filled from the event date, always overridable)',
                  'Visibility: configure location filter + mailing lists/individual users — who sees the event in the list',
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
                // v31.2: Ob die Filterverknüpfung (Badge 22) sichtbar ist — einmal
                // ausrechnen; sie ist der letzte Block, die Nummern davor hängen nicht daran.
                const seBothFilters = seLocationFilterList.length > 0 && (se.audience || '').trim().length > 0;
                const seUnlimited = (se.maxParticipants || 0) === 0;
                const seWaitlist = typeof se.waitlistEnabled === 'boolean' ? se.waitlistEnabled : true;
                return (
                  <div>
                    {/* v15.3: „Vom Hauptevent kopieren"-Button. Übernimmt
                        Kapazitäts-/Sichtbarkeits-/Deadline-/Filter-Werte
                        vom Hauptevent als Startwerte für dieses Sub-Event.
                        v28.74: Gegenstück — Werte dieses Sub-Events auf die
                        anderen übertragen (was und wohin wählt der Organizer).
                        v31.2: linksbündig statt rechts außen (Leitfaden 2a′) —
                        zwei Knöpfe allein am rechten Rand lasen sich wie eine
                        Fußzeile, nicht wie der Einstieg in den Reiter. */}
                    <div className="dex-ui-inline" style={{ gap: 10, marginBottom: 12 }}>
                      <button
                        type="button"
                        className="btn btn-secondary dex-ui-btn-sm"
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
                      {subEvents.length > 1 && (
                        <button
                          type="button"
                          className="btn btn-secondary dex-ui-btn-sm"
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
                        <div className="dex-ui-callout dex-ui-callout--warn" style={{ flexDirection: 'column', gap: 6, marginBottom: 14 }}>
                          <div style={{ fontWeight: 700 }}>
                            {isDe ? 'Diese Einstellungen weichen von den anderen ab' : 'These settings differ from the others'}
                          </div>
                          <ul style={{ margin: 0, paddingLeft: 18 }}>
                            {diffs.map(x => (
                              <li key={x.g.key}>
                                {isDe ? x.g.de.replace(/ —.*$/, '') : x.g.en.replace(/ —.*$/, '')}: {isDe
                                  ? <>bei <strong>{x.n} von {others}</strong> anderen {others === 1 ? 'Sub-Event' : 'Sub-Events'} anders</>
                                  : <>differs in <strong>{x.n} of {others}</strong> other sub-event{others === 1 ? '' : 's'}</>}
                              </li>
                            ))}
                          </ul>
                          <div>
                            {isDe
                              ? 'Das kann so gewollt sein. Wenn nicht, übernimm die Werte dieses Sub-Events für die anderen:'
                              : 'That may be intentional. If not, apply this sub-event’s values to the others:'}
                          </div>
                          <div>
                            <button
                              type="button"
                              className="btn btn-primary dex-ui-btn-sm"
                              onClick={() => setSubTransfer({
                                fromIdx: seIdx,
                                groups: diffs.map(x => x.g.key),
                                targets: subEvents.map((_, i) => i).filter(i => i !== seIdx),
                              })}
                            >
                              {isDe ? 'Auf die anderen übertragen…' : 'Transfer to the others…'}
                            </button>
                          </div>
                        </div>
                      );
                    })()}

                    {/* ===== v31.2: dieselbe Reihenfolge wie auf der Klammer —
                        Plätze → Fristen → Sichtbarkeit — mit denselben Badge-Nummern
                        wie dort (18–22; Umnummerierung 07.09.2026). ===== */}

                    {/* Teilnehmerzahl & Warteliste — Split-Capacity bleibt
                        Hauptevent-only (Scope-Eingrenzung, siehe v15.6
                        Refactor-Plan). v31.2: „Unbegrenzt"/„Begrenzt" als zwei
                        Kacheln statt Schieberegler; die Handler sind dieselben
                        (Unbegrenzt = 0 Plätze + Warteliste aus, Begrenzt = 50 als
                        Startwert), nur gegen Klick auf die aktive Kachel geschützt. */}
                    {/* v31.2 (Review): Abschnitts-Überschrift wie auf der Klammer,
                        damit Plätze, Fristen und Sichtbarkeit gleichrangig lesen. */}
                    <div className="dex-ui-section">
                      <div className="dex-ui-section-title">
                        <Users size={14} />
                        {isDe ? 'Plätze — wie viele dürfen kommen?' : 'Seats — how many may come?'}
                      </div>
                    <div className="dex-ui-card" style={{ marginBottom: 12 }}>
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <StepBadge n={18} />
                        {isDe ? 'Plätze & Warteliste' : 'Seats & waitlist'}
                      </label>
                      <div className="dex-ui-grid-2">
                        <button
                          type="button"
                          className={cx('dex-ui-choice', seUnlimited && 'is-active')}
                          aria-pressed={seUnlimited}
                          onClick={() => { if (!seUnlimited) updateSub({ maxParticipants: 0, waitlistEnabled: false }); }}
                        >
                          <span className="dex-ui-choice-body">
                            <span className="dex-ui-choice-title">{isDe ? 'Unbegrenzt' : 'Unlimited'}</span>
                            <span className="dex-ui-choice-desc">{isDe ? 'Jede Anmeldung bekommt einen Platz — keine Warteliste.' : 'Every registration gets a seat — no waitlist.'}</span>
                          </span>
                          <span className="dex-ui-choice-check">{seUnlimited && <Check size={12} />}</span>
                        </button>
                        <button
                          type="button"
                          className={cx('dex-ui-choice', !seUnlimited && 'is-active')}
                          aria-pressed={!seUnlimited}
                          onClick={() => { if (seUnlimited) updateSub({ maxParticipants: 50 }); }}
                        >
                          <span className="dex-ui-choice-body">
                            <span className="dex-ui-choice-title">{isDe ? 'Begrenzt' : 'Limited'}</span>
                            <span className="dex-ui-choice-desc">{isDe ? 'Du legst eine Platzzahl fest; ist sie erreicht, greift die Warteliste.' : 'You set a seat count; once it is reached, the waitlist takes over.'}</span>
                          </span>
                          <span className="dex-ui-choice-check">{!seUnlimited && <Check size={12} />}</span>
                        </button>
                      </div>
                      {!seUnlimited && (
                        <div className="dex-ui-grid-2" style={{ marginTop: 14 }}>
                          <div className="dex-ui-field" style={{ marginBottom: 0 }}>
                            <div className="dex-ui-label">{isDe ? 'Wie viele Plätze?' : 'How many seats?'}</div>
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
                            <p className="dex-ui-help">{isDe ? '0 = unbegrenzt' : '0 = unlimited'}</p>
                          </div>
                          <label className={cx('dex-ui-toggle-row', seWaitlist && 'is-active')} style={{ alignSelf: 'start' }}>
                            <input
                              type="checkbox"
                              checked={seWaitlist}
                              onChange={e => updateSub({ waitlistEnabled: e.target.checked })}
                            />
                            <span className="dex-ui-toggle-row-body">
                              <span className="dex-ui-toggle-row-title">{t('create.waitlist')}</span>
                              <span className="dex-ui-toggle-row-desc">
                                {seWaitlist
                                  ? (isDe ? 'An: Ist der Termin voll, landen weitere Anmeldungen auf der Warteliste und rücken automatisch nach.' : 'On: once the date is full, further sign-ups land on the waitlist and are promoted automatically.')
                                  : (isDe ? 'Aus: Ist der Termin voll, ist der Anmelde-Knopf gesperrt.' : 'Off: once the date is full, the register button is locked.')}
                              </span>
                            </span>
                          </label>
                        </div>
                      )}
                      <p className="dex-ui-help" style={{ marginTop: 12 }}>
                        {isDe
                          ? <em>Hinweis: <strong>Geteilte Kapazität</strong> (zwei Gruppen mit eigener Platzzahl) ist aktuell nur auf Hauptevent-Ebene möglich — Sub-Events nutzen die einfache Gesamtkapazität.</em>
                          : <em>Note: <strong>Split capacity</strong> (two groups with separate seat counts) is currently main-event-only — sub-events use the simple total capacity.</em>}
                      </p>
                    </div>
                    </div>{/* v31.2: Ende Abschnitt Plätze */}

                    {/* Deadlines: zwei DatePicker nebeneinander, gleicher Look
                        wie im Hauptevent. */}
                    <div className="dex-ui-section">
                      <div className="dex-ui-section-title">
                        <Icon iconName="Clock" style={{ fontSize: 14 }} />
                        {isDe ? 'Fristen — bis wann?' : 'Deadlines — until when?'}
                      </div>
                    <div className="dex-ui-card" style={{ marginBottom: 12 }}>
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <StepBadge n={19} />
                        {isDe ? 'Anmelde- und Abmeldefristen' : 'Registration & cancellation deadlines'}
                      </label>
                      <p className="dex-ui-help" style={{ margin: '0 0 12px' }}>
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
                        <div className="dex-ui-callout dex-ui-callout--neutral" style={{ marginBottom: 12 }}>
                          <span className="dex-ui-callout-icon"><Info size={16} /></span>
                          <span>
                            {isDe
                              ? <>Rollierende Regel der Klammer aktiv: {regRuleEnabled && cancelRuleEnabled ? 'beide Fristen werden' : (regRuleEnabled ? <>&bdquo;Anmeldung bis&ldquo; wird</> : <>&bdquo;Abmeldung bis&ldquo; wird</>)} aus dem Termin-Datum <strong>vorbelegt</strong>. Du kannst sie hier für diesen einzelnen Termin überschreiben — dein Wert bleibt dann stehen. Änderst du später die Regel selbst, werden wieder <strong>alle</strong> Termine neu berechnet.</>
                              : <>Rolling bracket rule active: {regRuleEnabled && cancelRuleEnabled ? 'both deadlines are' : 'this deadline is'} <strong>pre-filled</strong> from the date. You can override it here for this single date — your value then sticks. Changing the rule itself recomputes <strong>all</strong> dates again.</>}
                          </span>
                        </div>
                      )}
                      <div className="dex-ui-grid-2">
                        <div className="dex-ui-field" style={{ marginBottom: 0 }}>
                          {/* v29.75: gleicher Wortlaut wie auf der Klammer. */}
                          <div className="dex-ui-label">{isDe ? 'Anmeldung bis' : 'Registration until'}</div>
                          <DatePicker
                            selected={berlinIsoToPickerDate(se.registrationDeadline)}
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
                              <p className="dex-ui-help">
                                {isDe ? 'Entspricht der rollierenden Regel.' : 'Matches the rolling rule.'}
                              </p>
                            ) : (
                              <p className="dex-ui-help dex-ui-inline" style={{ color: '#b86700', gap: 4 }}>
                                {isDe
                                  ? <>Manuell überschrieben — Regel wäre {new Date(ruleIso).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', timeZone: 'Europe/Berlin' })}.</>
                                  : <>Manually overridden — rule would be {new Date(ruleIso).toLocaleString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', timeZone: 'Europe/Berlin' })}.</>}
                                <button
                                  type="button"
                                  className="dex-ui-textbtn"
                                  style={{ padding: '2px 6px', fontSize: '0.74rem' }}
                                  onClick={() => updateSub({ registrationDeadline: ruleIso })}
                                >
                                  {isDe ? 'Auf Regel zurücksetzen' : 'Reset to rule'}
                                </button>
                              </p>
                            );
                          })()}
                        </div>
                        <div className="dex-ui-field" style={{ marginBottom: 0 }}>
                          <div className="dex-ui-label">{isDe ? 'Abmeldung bis' : 'Cancellation until'}</div>
                          <DatePicker
                            selected={berlinIsoToPickerDate(se.lastDeregisterDate)}
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
                              <p className="dex-ui-help">
                                {isDe ? 'Entspricht der rollierenden Regel.' : 'Matches the rolling rule.'}
                              </p>
                            ) : (
                              <p className="dex-ui-help dex-ui-inline" style={{ color: '#b86700', gap: 4 }}>
                                {isDe
                                  ? <>Manuell überschrieben — Regel wäre {new Date(ruleIso).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', timeZone: 'Europe/Berlin' })}.</>
                                  : <>Manually overridden — rule would be {new Date(ruleIso).toLocaleString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', timeZone: 'Europe/Berlin' })}.</>}
                                <button
                                  type="button"
                                  className="dex-ui-textbtn"
                                  style={{ padding: '2px 6px', fontSize: '0.74rem' }}
                                  onClick={() => updateSub({ lastDeregisterDate: ruleIso })}
                                >
                                  {isDe ? 'Auf Regel zurücksetzen' : 'Reset to rule'}
                                </button>
                              </p>
                            );
                          })()}
                        </div>
                      </div>
                    </div>
                    </div>{/* v31.2: Ende Abschnitt Fristen */}

                    {/* v15.6: Sichtbarkeits-Sektion analog Hauptevent. */}
                    <div className="dex-ui-section" style={{ marginTop: 22 }}>
                      <div className="dex-ui-section-title">
                        <Icon iconName="Hide3" style={{ fontSize: 14 }} />
                        {isDe ? 'Sichtbarkeit — wer sieht diesen Termin?' : 'Visibility — who sees this date?'}
                      </div>
                      <p className="dex-ui-section-desc">
                        {isDe
                          ? <>Dieses Sub-Event übernimmt <strong>standardmäßig die Sichtbarkeit {subEventsOnlyMode ? 'der Klammer' : 'des Hauptevents'}</strong> (Standortfilter + Mailverteiler sind unten vorbefüllt). Du kannst den Empfängerkreis hier <strong>jederzeit anpassen</strong> — oder mit „Vom Hauptevent kopieren“ oben erneut übernehmen.</>
                          : <>This sub-event <strong>inherits the visibility {subEventsOnlyMode ? 'of the bracket' : 'of the main event'}</strong> by default (location filter + mailing lists are pre-filled below). You can <strong>change the audience here at any time</strong> — or re-apply it with “Copy from main event” above.</>}
                      </p>
                      {/* v26.88: Live-Zusammenfassung wandert in die
                          AudiencePicker-Prüfzeile (summarySlot). */}

                      {/* v29.75: Solange „Sichtbarkeit gilt für alle Sub-Events"
                          auf der Klammer gesetzt ist, wird die Sichtbarkeit hier
                          nur ANGEZEIGT — Eingaben würden beim nächsten Klammer-
                          Edit stillschweigend überschrieben (Spiegel-Effect),
                          deshalb sperren statt verlieren lassen. */}
                      {visAllSubs && (
                        <div className="dex-ui-callout dex-ui-callout--success" style={{ marginBottom: 12 }}>
                          <span className="dex-ui-callout-icon"><Info size={16} /></span>
                          <span>
                            {isDe
                              ? <>Die Sichtbarkeit wird von der <strong>{subEventsOnlyMode ? 'Klammer' : 'Hauptevent-Ebene'}</strong> vorgegeben (&bdquo;Sichtbarkeit gilt für alle {childTermPlural || 'Sub-Events'}&ldquo;). Zum Abweichen den Haken dort entfernen.</>
                              : <>Visibility is governed by the <strong>{subEventsOnlyMode ? 'bracket' : 'main event'}</strong> (“visibility applies to all {childTermPlural || 'sub-events'}”). Uncheck the box there to deviate.</>}
                          </span>
                        </div>
                      )}
                      <div className={visAllSubs ? 'dex-ui-card--muted' : undefined} style={visAllSubs ? { pointerEvents: 'none' as const, userSelect: 'none' as const } : undefined}>
                      <div className="dex-ui-card" style={{ marginBottom: 12 }}>
                        <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                          <StepBadge n={20} />
                          {isDe ? 'Standortfilter' : 'Location filter'}
                        </label>
                        <p className="dex-ui-help" style={{ margin: '0 0 12px' }}>
                          {isDe
                            ? <>Nur Mitarbeiter mit einem der gewählten Standorte sehen dieses Sub-Event. Leer = für alle sichtbar.</>
                            : <>Only employees from one of the selected locations see this sub-event. Empty = visible to everyone.</>}
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
                        stepBadge={<StepBadge n={21} />}
                        cardBgPrimary={zebraS3Bg()}
                        summarySlot={renderVisibilitySummaryBox(
                          seLocationFilterList,
                          se.audience || '',
                          (se.filterMode as 'AND' | 'OR') || 'OR',
                          (se.excludedUsers || []).length
                        )}
                        middleSlot={seBothFilters ? (
                          <div className="dex-ui-card" style={{ marginBottom: 12 }}>
                            <div className="dex-ui-label">
                              <StepBadge n={22} />
                              {isDe ? 'Wie greifen die beiden Filter zusammen?' : 'How do the two filters combine?'}
                            </div>
                            <p className="dex-ui-help" style={{ margin: '-2px 0 10px' }}>
                              {isDe
                                ? <>Beide Filter sind gesetzt — muss eine Person <strong>einen</strong> davon erfüllen (ODER) oder <strong>beide</strong> (UND), damit das Sub-Event in ihrer Liste auftaucht?</>
                                : <>Both filters are set — does a person need to match <strong>either</strong> filter (OR) or <strong>both</strong> (AND) for the sub-event to appear in their list?</>}
                            </p>
                            <div className="dex-ui-grid-2">
                              <button
                                type="button"
                                className={cx('dex-ui-choice', (se.filterMode || 'OR') === 'OR' && 'is-active')}
                                aria-pressed={(se.filterMode || 'OR') === 'OR'}
                                onClick={() => updateSub({ filterMode: 'OR' })}
                              >
                                <span className="dex-ui-choice-body">
                                  <span className="dex-ui-choice-title">{isDe ? 'ODER' : 'OR'}</span>
                                  <span className="dex-ui-choice-desc">{isDe ? 'Einer der Filter reicht.' : 'One filter is enough.'}</span>
                                </span>
                                <span className="dex-ui-choice-check">{(se.filterMode || 'OR') === 'OR' && <Check size={12} />}</span>
                              </button>
                              <button
                                type="button"
                                className={cx('dex-ui-choice', se.filterMode === 'AND' && 'is-active')}
                                aria-pressed={se.filterMode === 'AND'}
                                onClick={() => updateSub({ filterMode: 'AND' })}
                              >
                                <span className="dex-ui-choice-body">
                                  <span className="dex-ui-choice-title">{isDe ? 'UND' : 'AND'}</span>
                                  <span className="dex-ui-choice-desc">{isDe ? 'Beides muss zutreffen.' : 'Both must match.'}</span>
                                </span>
                                <span className="dex-ui-choice-check">{se.filterMode === 'AND' && <Check size={12} />}</span>
                              </button>
                            </div>
                          </div>
                        ) : null}
                        cardBgSecondary={((se.locationFilter || '').trim().length > 0 || (se.audience || '').trim().length > 0) ? zebraS3Bg() : '#fff'}
                      />
                      </div>{/* v29.75: Ende Sichtbarkeits-Sperre bei visAllSubs */}
                    </div>
                  </div>
                );
              })()}

              <div style={{ display: activeCapacityTabIdx === 0 ? 'block' : 'none' }}>
              {renderHauptGreyoutBanner()}

              {/* ===== v31.2: Abschnitt 1 — Plätze & Warteliste. Steht jetzt
                  ZUERST: „Wie viele dürfen kommen?" ist die Frage, die jeder
                  Organizer sofort beantworten kann; Fristen und Sichtbarkeit
                  bauen darauf auf. Die StepBadge-Nummern laufen seit der
                  Umnummerierung (07.09.2026) mit: 18 Plätze … 23 Assistenz. ===== */}
              {/* v9.17: Standard-Teilnehmerzahl zuerst, Split-Toggle darunter —
                  die Mehrheit der Events nutzt nur eine Gesamtkapazität; der
                  B2Run-Sonderfall ist Opt-in. */}
              {/* v28.72: Bei einer Klammer war „Teilnehmerzahl & Warteliste"
                  nur ausgegraut — ohne Begründung und ohne Weg. Organizer
                  klickten ins Leere und hielten es für kaputt. Jetzt steht eine
                  Erklär-Zeile an der Stelle des Abschnitts (v28.76: nicht mehr
                  als Overlay, das die Weiter-Knöpfe überdeckte), die sagt WARUM
                  es hier nicht gilt und mit einem Klick ins erste Sub-Event
                  führt, wo die Plätze tatsächlich gepflegt werden. */}
              {/* v31.2 (Review): Plätze und Fristen bekommen dieselbe Abschnitts-
                  Überschrift wie die Sichtbarkeit — drei gleichrangige Fragen,
                  drei gleich aussehende Abschnitte. Der visHeader mit Badge
                  bleibt in der Karte, weil Support auf die Nummer verweist. */}
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">
                  <Users size={14} />
                  {isDe ? 'Plätze — wie viele dürfen kommen?' : 'Seats — how many may come?'}
                </div>
              {subEventsOnlyMode && subEvents.length > 0 ? (() => {
                // v31.2 (Leitfaden 2a′): Der Kasten hat genau EINE Aktion — ins
                // erste Sub-Event springen. Also ist der ganze Kasten klickbar
                // (Hover, Enter/Leertaste), und der Knopf steht links direkt
                // hinter dem Text statt allein am rechten Rand. Derselbe
                // Handler wie bisher.
                const goFirstSub = (): void => setActiveCapacityTabIdx(1);
                return (
                <div
                  className="dex-ui-callout dex-ui-callout--neutral dex-ui-card--hover"
                  style={{ alignItems: 'center', flexWrap: 'wrap', gap: 10, marginBottom: 12, cursor: 'pointer' }}
                  role="button"
                  tabIndex={0}
                  onClick={goFirstSub}
                  onKeyDown={e => {
                    if (e.target !== e.currentTarget) return;
                    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); goFirstSub(); }
                  }}
                >
                  <span className="dex-ui-callout-icon"><Users size={18} /></span>
                  <span style={{ minWidth: 0 }}>
                    {isDe
                      ? <><strong>Plätze &amp; Warteliste</strong> werden pro {childTermSingular || 'Sub-Event'} vergeben — bei einer Klammer hätte eine Teilnehmerzahl hier keine Wirkung.</>
                      : <><strong>Seats &amp; waitlist</strong> are set per sub-event — for a bracket a capacity here would have no effect.</>}
                  </span>
                  <button
                    type="button"
                    className="btn btn-secondary dex-ui-btn-sm"
                    style={{ flexShrink: 0 }}
                    onClick={e => { e.stopPropagation(); goFirstSub(); }}
                    onKeyDown={e => e.stopPropagation()}
                  >
                    {isDe
                      ? `Zu „${shortSubEventTitle(subEvents[0].title, title) || (childTermSingular || 'Sub-Event')}“`
                      : `Go to „${shortSubEventTitle(subEvents[0].title, title) || 'sub-event'}“`}
                  </button>
                </div>
                );
              })() : (
                <>
              <div style={hauptGreyoutWrapperStyle()}>
              <div className="dex-ui-card" style={{ marginBottom: 12 }}>
                {visHeader('vis_capacity', <StepBadge n={18} />, isDe ? 'Plätze & Warteliste' : 'Seats & waitlist')}
                {isVisOpen('vis_capacity') && (<>
                {/* v22.38 stellte die Frage als Checkbox „Teilnehmeranzahl
                    begrenzen?". v31.2: zwei Kacheln „Unbegrenzt" / „Begrenzt"
                    mit je einer Zeile Folge. Die Handler sind dieselben
                    (Begrenzt schaltet die Warteliste ein, Unbegrenzt leert die
                    Zahl und schaltet sie aus) — nur mit Schutz gegen einen Klick
                    auf die schon aktive Kachel, denn eine Checkbox feuert bei
                    Gleichstand auch nicht. Bei geteilter Kapazität stellt sich
                    die Frage nicht (wie bisher). */}
                {!useSplitCapacities && (
                  <div className="dex-ui-field">
                    <div className="dex-ui-label">
                      {isDe ? 'Wie viele Plätze gibt es?' : 'How many seats are there?'}
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
                    </div>
                    <div className="dex-ui-grid-2">
                      <button
                        type="button"
                        className={cx('dex-ui-choice', unlimitedParticipants && 'is-active')}
                        aria-pressed={unlimitedParticipants}
                        onClick={() => {
                          if (!unlimitedParticipants) {
                            setUnlimitedParticipants(true);
                            setMaxParticipants('');
                            setWaitlistEnabled(false);
                          }
                        }}
                      >
                        <span className="dex-ui-choice-body">
                          <span className="dex-ui-choice-title">{isDe ? 'Unbegrenzt' : 'Unlimited'}</span>
                          <span className="dex-ui-choice-desc">
                            {isDe
                              ? 'Jede Anmeldung bekommt einen Platz — keine Auslastungsanzeige, keine Warteliste. Typisch für Info-Events und All-Hands.'
                              : 'Every registration gets a seat — no capacity indicator, no waitlist. Typical for info events and all-hands.'}
                          </span>
                        </span>
                        <span className="dex-ui-choice-check">{unlimitedParticipants && <Check size={12} />}</span>
                      </button>
                      <button
                        type="button"
                        className={cx('dex-ui-choice', !unlimitedParticipants && 'is-active')}
                        aria-pressed={!unlimitedParticipants}
                        onClick={() => {
                          if (unlimitedParticipants) {
                            setUnlimitedParticipants(false);
                            setWaitlistEnabled(true);
                          }
                        }}
                      >
                        <span className="dex-ui-choice-body">
                          <span className="dex-ui-choice-title">{isDe ? 'Begrenzt' : 'Limited'}</span>
                          <span className="dex-ui-choice-desc">
                            {isDe
                              ? 'Du legst eine Platzzahl fest. Ist sie erreicht, sehen Teilnehmer „Alle Plätze sind belegt“ — die Warteliste schaltet sich automatisch ein.'
                              : 'You set a seat count. Once it is reached, attendees see “All spots are taken” — the waitlist switches on automatically.'}
                          </span>
                        </span>
                        <span className="dex-ui-choice-check">{!unlimitedParticipants && <Check size={12} />}</span>
                      </button>
                    </div>
                  </div>
                )}
                {!useSplitCapacities && !unlimitedParticipants && (
                  <div className="dex-ui-field" style={{ maxWidth: 280 }}>
                    <div className="dex-ui-label">{t('create.maxparticipants')}</div>
                    <input
                      className="form-input"
                      type="number"
                      min={0}
                      value={maxParticipants}
                      onChange={e => setMaxParticipants(e.target.value.replace(/[^0-9]/g, ''))}
                      placeholder={isDe ? 'z.B. 50' : 'e.g. 50'}
                      style={errorBorderStyle('maxParticipants')}
                    />
                    {fieldHasError('maxParticipants') && <span style={{ color: 'var(--dex-red)', fontSize: '0.75rem' }}>{t('create.error.required')}</span>}
                  </div>
                )}

                {/* v9.17: Split-Capacity-Toggle bewusst subtil — der Großteil
                    der Events nutzt eine einzige Teilnehmerzahl.
                    v24.4 (K): erst sichtbar, wenn „Teilnehmeranzahl begrenzen"
                    aktiv ist (bzw. die geteilte Kapazität bereits an ist) —
                    ohne Begrenzung ergibt eine geteilte Kapazität keinen Sinn.
                    v31.2: direkt unter der Platzzahl statt unter der Warteliste
                    — Abhängiges folgt seinem Schalter; der lange Hinweistext
                    steht im Tooltip. */}
                {(!unlimitedParticipants || useSplitCapacities) && (
                  <label className={cx('dex-ui-toggle-row', useSplitCapacities && 'is-active')}>
                    <input
                      type="checkbox"
                      checked={useSplitCapacities}
                      onChange={e => setUseSplitCapacities(e.target.checked)}
                    />
                    <span className="dex-ui-toggle-row-body">
                      <span className="dex-ui-toggle-row-title">
                        {t('create.splitcap.label')}
                        <InfoTooltip text={isDe ? (
                          <>
                            {t('create.splitcap.hint')}<br /><br />
                            <strong>Was du hier einstellst:</strong> zwei getrennte Kapazitäten innerhalb des Events. Beispiele: <strong>Vormittag / Nachmittag</strong>, <strong>VIP / Standard</strong>, <strong>Lauf 5 km / Lauf 10 km</strong>. Du legst die <strong>Bezeichnungen frei fest</strong> und vergibst pro Gruppe eine eigene Platzzahl mit eigener Warteliste.<br /><br />
                            <strong>Anzeige in der App:</strong> Teilnehmer sehen auf der Anmelde-Seite <strong>zwei nebeneinanderstehende Boxen</strong> mit deinen Bezeichnungen, jeweils mit &bdquo;X / Y Plätze frei&ldquo;. Die Wahl ist <strong>verpflichtend</strong>, bevor man auf Anmelden klickt.<br /><br />
                            <strong>Automatismen:</strong> ist eine der zwei Gruppen voll, kommen weitere Anmeldungen <strong>nur in die Warteliste dieser Gruppe</strong> — nicht in die andere. Beim Nachrücken bleibt der Typ <strong>erhalten</strong> (eine VIP-Wartelisten-Person rückt nicht in einen Standard-Platz).<br /><br />
                            <strong>Empfehlung:</strong> nur verwenden, wenn die zwei Gruppen <strong>wirklich getrennt</strong> behandelt werden sollen (eigenes Catering, eigener Bus, eigener Slot beim Veranstalter). Bei einer einfachen Gesamtkapazität reicht die Standard-Teilnehmerzahl oben.
                          </>
                        ) : (
                          <>
                            {t('create.splitcap.hint')}<br /><br />
                            <strong>What you set here:</strong> two separate capacities within the event. Examples: <strong>morning / afternoon</strong>, <strong>VIP / standard</strong>, <strong>5 km run / 10 km run</strong>. You <strong>name the two groups freely</strong> and give each its own seat count and waitlist.<br /><br />
                            <strong>Shown in the app:</strong> attendees see <strong>two side-by-side boxes</strong> on the registration page, each labelled with your text and showing &ldquo;X / Y seats free&rdquo;. Picking one is <strong>required</strong> before submitting.<br /><br />
                            <strong>Automation:</strong> when one group is full, further sign-ups land <strong>only on that group&apos;s waitlist</strong> — not the other. When promoting from waitlist, the <strong>group is preserved</strong> (a waitlisted VIP is not auto-promoted into a standard slot).<br /><br />
                            <strong>Tip:</strong> only use this when the two groups really need to be <strong>handled separately</strong> (own catering, own bus, separate slot with the supplier). For a single overall capacity the standard attendee count above is enough.
                          </>
                        )} />
                      </span>
                      <span className="dex-ui-toggle-row-desc">
                        {isDe
                          ? 'Zwei Gruppen mit eigener Bezeichnung, Platzzahl und Warteliste — z.B. Vormittag/Nachmittag oder VIP/Standard. Aus = eine Gesamt-Teilnehmerzahl.'
                          : 'Two groups, each with its own name, seat count and waitlist — e.g. morning/afternoon or VIP/standard. Off = one overall attendee limit.'}
                      </span>
                    </span>
                  </label>
                )}

                {/* v10.20: Geteilte Kapazität — generisch für beliebige Events.
                    Labels werden vom Organizer frei gewählt (z.B. "Vormittag /
                    Nachmittag", "VIP / Standard", "Lauf / Walk"). Default-Fallback
                    ist 'Durchstarter' / 'Funstarter' für Backward-Compat mit
                    B2Run-Events vor v10.20. */}
                {useSplitCapacities && (
                  <div className="dex-ui-card dex-ui-card--soft dex-ui-card--accent" style={{ marginTop: 12 }}>
                    <p className="dex-ui-section-desc" style={{ margin: '0 0 12px' }}>
                      {isDe
                        ? 'Vergib pro Gruppe eine Bezeichnung und eine Platzzahl. Die Bezeichnungen erscheinen auf der Anmeldeseite als zwei Auswahl-Boxen.'
                        : 'Give each group a name and a seat count. The names appear on the registration page as two selectable boxes.'}
                    </p>
                    {/* v10.20: zwei Text-Inputs für die frei wählbaren Bezeichnungen.
                        Wenn der Organizer nichts einträgt, fällt die Registration-
                        Seite auf 'Durchstarter' / 'Funstarter' zurück. */}
                    <div className="dex-ui-grid-2" style={{ marginBottom: 14 }}>
                      <div className="dex-ui-field" style={{ marginBottom: 0 }}>
                        <label className="dex-ui-label">{isDe ? 'Bezeichnung Gruppe A' : 'Group A label'}</label>
                        <input
                          className="form-input"
                          type="text"
                          value={splitLabelA}
                          onChange={e => setSplitLabelA(e.target.value)}
                          placeholder={isDe ? 'z.B. Vormittag, VIP, Durchstarter' : 'e.g. morning, VIP, starter'}
                          maxLength={40}
                        />
                      </div>
                      <div className="dex-ui-field" style={{ marginBottom: 0 }}>
                        <label className="dex-ui-label">{isDe ? 'Bezeichnung Gruppe B' : 'Group B label'}</label>
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
                    <div className="dex-ui-grid-2">
                      <div className="dex-ui-field" style={{ marginBottom: 0 }}>
                        <label className="dex-ui-label">
                          <Icon iconName="People" style={{ fontSize: 14, color: 'var(--dex-green-dark, #6b9a1e)' }} />
                          <span>{isDe ? 'Plätze' : 'Seats'} {splitLabelA.trim() || (isDe ? 'Gruppe A' : 'Group A')}</span>
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
                      <div className="dex-ui-field" style={{ marginBottom: 0 }}>
                        <label className="dex-ui-label">
                          <Icon iconName="People" style={{ fontSize: 14, color: 'var(--dex-orange, #ff8c00)' }} />
                          <span>{isDe ? 'Plätze' : 'Seats'} {splitLabelB.trim() || (isDe ? 'Gruppe B' : 'Group B')}</span>
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
                    <div className="dex-ui-inline" style={{ marginTop: 12 }}>
                      <span className="dex-ui-pill dex-ui-pill--green">
                        {t('create.b2runcap.total')}: {((parseInt(durchstarterCapacity, 10) || 0) + (parseInt(funstarterCapacity, 10) || 0))} {t('create.b2runcap.seats')}
                      </span>
                    </div>

                    {/* v31.2: Feineinstellungen der Gruppen-Auswahl — Überschrift,
                        Beschreibungen, Hinweistext, Reihenfolge, Startblöcke und
                        der Pflichtfeld-Hinweis — hinter einem Aufklapper. Beim
                        Anlegen interessieren nur Namen und Plätze; sechs weitere
                        Felder auf einmal ließen die Karte wie ein Formular für
                        Experten aussehen. */}
                    <button
                      type="button"
                      className={cx('dex-ui-disclosure', splitMoreOpen && 'is-open')}
                      style={{ marginTop: 10 }}
                      aria-expanded={splitMoreOpen}
                      onClick={() => setSplitMoreOpen(o => !o)}
                    >
                      <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                      {isDe ? 'Texte & Reihenfolge auf der Anmeldeseite' : 'Texts & order on the registration page'}
                      {/* v31.2 (Review): Statt der Feldzahl steht hier, ob etwas
                          vom Standard abweicht — die Zahl sagte nur, wie viele
                          Felder drin sind, nicht ob der Organizer sie braucht. */}
                      <span className="dex-ui-disclosure-count">
                        {(splitSectionTitle.trim() || splitDescA.trim() || splitDescB.trim() || splitHelpText.trim() || splitDisplayOrderReversed || durchstarterStartblock || funstarterStartblock)
                          ? (isDe ? 'angepasst' : 'customized')
                          : (isDe ? 'Standard' : 'default')}
                      </span>
                    </button>
                    {splitMoreOpen && (
                      <div className="dex-ui-disclosure-body">
                        {/* v26.83: frei wählbare Überschrift der Gruppen-Auswahl auf der
                            Anmeldeseite (statt „Gruppen-Auswahl"). */}
                        <div className="dex-ui-field">
                          <label className="dex-ui-label">
                            {isDe ? 'Überschrift der Auswahl' : 'Selection heading'}
                            <span className="dex-ui-label-optional">(optional)</span>
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
                        {/* v26.72: optionale Beschreibung pro Gruppe — erscheint unter
                            dem Gruppen-Namen in der Auswahl-Karte auf der Anmeldeseite. */}
                        <div className="dex-ui-grid-2" style={{ marginBottom: 16 }}>
                          <div className="dex-ui-field" style={{ marginBottom: 0 }}>
                            <label className="dex-ui-label">
                              {isDe ? 'Beschreibung Gruppe A' : 'Group A description'}
                              <span className="dex-ui-label-optional">(optional)</span>
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
                          <div className="dex-ui-field" style={{ marginBottom: 0 }}>
                            <label className="dex-ui-label">
                              {isDe ? 'Beschreibung Gruppe B' : 'Group B description'}
                              <span className="dex-ui-label-optional">(optional)</span>
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
                        <div className="dex-ui-field">
                          <label className="dex-ui-label">
                            {isDe ? 'Hinweistext über der Gruppen-Auswahl' : 'Help text above the group selection'}
                            <span className="dex-ui-label-optional">(optional)</span>
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
                          <p className="dex-ui-help">
                            {isDe ? 'Dieser Text steht auf der Anmeldeseite direkt über den beiden Gruppen-Boxen. Leer lassen für den Standardtext.' : 'This text appears on the registration page right above the two group boxes. Leave empty for the default.'}
                          </p>
                        </div>
                        {/* v11.25: Display-Reihenfolge der zwei Gruppen-Karten in der
                            Registrierungs-UI umkehren. Reine Anzeige-Toggle —
                            splitLabelA/B, Kapazitäten und die internen StarterType-IDs
                            bleiben unangetastet. */}
                        <label className={cx('dex-ui-toggle-row', splitDisplayOrderReversed && 'is-active')}>
                          <input
                            type="checkbox"
                            checked={splitDisplayOrderReversed}
                            onChange={e => setSplitDisplayOrderReversed(e.target.checked)}
                          />
                          <span className="dex-ui-toggle-row-body">
                            <span className="dex-ui-toggle-row-title">
                              {isDe ? 'Reihenfolge auf der Anmeldeseite umkehren' : 'Reverse order on the registration page'}
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
                            <span className="dex-ui-toggle-row-desc">
                              {isDe
                                ? <>An = &bdquo;{splitLabelB.trim() || 'Gruppe B'}&ldquo; steht links, &bdquo;{splitLabelA.trim() || 'Gruppe A'}&ldquo; rechts. Rein optisch — Plätze, Wartelisten und Anmeldungen bleiben unberührt.</>
                                : <>On = “{splitLabelB.trim() || 'group B'}” on the left, “{splitLabelA.trim() || 'group A'}” on the right. Purely visual — seats, waitlists and registrations stay untouched.</>}
                            </span>
                          </span>
                        </label>

                        {/* v6.15: Starter-Typ → Startblock-Zuordnung (optional).
                            Nur sinnvoll wenn Startblocks definiert sind (Reiter "Event-spezifische Felder").
                            Wenn gesetzt, wird der Startblock bei der Registrierung automatisch
                            anhand des gewählten Starter-Typs gesetzt — der User muss den Block
                            nicht extra auswählen. */}
                        {b2runStartblocks.length > 0 && (
                          <div className="dex-ui-field" style={{ marginTop: 16 }}>
                            <div className="dex-ui-label">
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
                            </div>
                            <div className="dex-ui-grid-2">
                              <div>
                                <label className="dex-ui-help" style={{ display: 'block', marginTop: 0, marginBottom: 4 }}>Durchstarter →</label>
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
                                <label className="dex-ui-help" style={{ display: 'block', marginTop: 0, marginBottom: 4 }}>Funstarter →</label>
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
                            auf eine Split-Gruppe einschränken. Hinweis steht hier,
                            damit der Organizer beim Migrieren weiß wo das Feature
                            jetzt liegt. */}
                        <div className="dex-ui-callout dex-ui-callout--info" style={{ marginTop: 16 }}>
                          <span className="dex-ui-callout-icon"><Info size={16} /></span>
                          <div>
                            <strong>{isDe ? 'Pflichtfelder pro Gruppe' : 'Required fields per group'}</strong><br />
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
                        </div>
                      </div>
                    )}
                  </div>
                )}

                {/* v31.2: Warteliste als eigene Frage unter der Platzzahl —
                    vorher stand der Schalter im Normalfall rechts neben der Zahl
                    und im Gruppen-Fall ganz unten im grünen Kasten; jetzt an
                    EINER Stelle, mit dem passenden Tooltip je Fall. Sichtbar,
                    sobald es eine Grenze gibt (wie bisher: im Gruppen-Fall immer,
                    sonst nur bei begrenzter Platzzahl). */}
                {(useSplitCapacities || !unlimitedParticipants) && (
                  <div style={{ marginTop: 14 }}>
                    <label className={cx('dex-ui-toggle-row', waitlistEnabled && 'is-active')}>
                      <input type="checkbox" checked={waitlistEnabled} onChange={e => setWaitlistEnabled(e.target.checked)} />
                      <span className="dex-ui-toggle-row-body">
                        <span className="dex-ui-toggle-row-title">
                          {isDe ? 'Warteliste führen, wenn alle Plätze belegt sind' : 'Keep a waitlist once all seats are taken'}
                          <InfoTooltip text={useSplitCapacities ? (isDe ? (
                            <>
                              <strong>Wartelisten für Split-Kapazitäten</strong> — bei aktivierter Warteliste hat <strong>jeder Starter-Typ seine eigene Liste</strong>. Durchstarter-Anmeldungen über der Durchstarter-Kapazität landen auf der Durchstarter-Warteliste, dasselbe für Funstarter. Beim Nachrücken wird <strong>typ-bewusst</strong> befördert — der älteste Durchstarter-Eintrag rückt in einen frei werdenden Durchstarter-Platz, kein Mix.
                            </>
                          ) : (
                            <>
                              <strong>Waitlists for split capacities</strong> — when the waitlist is on, <strong>each starter type has its own queue</strong>. Durchstarter sign-ups beyond the Durchstarter capacity go on the Durchstarter waitlist, same for Funstarter. Promotion is <strong>type-aware</strong> — the oldest Durchstarter waiting moves into a freed Durchstarter slot, no mixing.
                            </>
                          )) : (isDe ? (
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
                          ))} />
                        </span>
                        <span className="dex-ui-toggle-row-desc">
                          {waitlistEnabled
                            ? (isDe
                              ? 'An: Weitere Anmeldungen landen mit Positionsnummer auf der Warteliste, bekommen die Wartelisten-Mail und rücken automatisch nach, sobald ein Platz frei wird — wer zuerst kam, zuerst.'
                              : 'On: further sign-ups land on the waitlist with a position number, get the waitlist mail and are promoted automatically as soon as a seat frees up — first in, first out.')
                            : (isDe
                              ? 'Aus: Ist das Event voll, ist der Anmelde-Knopf gesperrt — Interessenten müssen dich direkt kontaktieren.'
                              : 'Off: once the event is full the register button is locked — interested people have to contact you directly.')}
                        </span>
                      </span>
                    </label>

                    {/* v10.20: Waitlist-Modus bei Split-Capacity. Default
                        'separate' (zwei Wartelisten) — entspricht dem alten
                        B2Run-Verhalten und ist die typ-bewusste Variante.
                        Alternative 'shared' = eine gemeinsame Warteliste, FIFO
                        über beide Gruppen. Sinnvoll wenn die Gruppen
                        organisatorisch fluide sind (z.B. Vormittag/Nachmittag
                        bei einem Workshop, wo der nächste freie Slot egal ist
                        welche Gruppe). Nur sichtbar wenn Warteliste aktiviert ist. */}
                    {useSplitCapacities && waitlistEnabled && (
                      <div className="dex-ui-field" style={{ marginTop: 12 }}>
                        <div className="dex-ui-label">
                          {isDe ? 'Wie verhält sich die Warteliste bei zwei Gruppen?' : 'How does the waitlist behave with two groups?'}
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
                        </div>
                        <div className="dex-ui-grid-2">
                          <button
                            type="button"
                            className={cx('dex-ui-choice', !splitSharedWaitlist && 'is-active')}
                            aria-pressed={!splitSharedWaitlist}
                            onClick={() => setSplitSharedWaitlist(false)}
                          >
                            <span className="dex-ui-choice-body">
                              <span className="dex-ui-choice-title">{isDe ? 'Getrennt je Gruppe' : 'Separate per group'}</span>
                              <span className="dex-ui-choice-desc">
                                {isDe
                                  ? <>Wer auf der Warteliste von &bdquo;{splitLabelA.trim() || 'Gruppe A'}&ldquo; steht, rückt nur in einen freien Platz dieser Gruppe nach. Standard.</>
                                  : <>Someone on the “{splitLabelA.trim() || 'group A'}” waitlist only moves into a freed seat of that group. Default.</>}
                              </span>
                            </span>
                            <span className="dex-ui-choice-check">{!splitSharedWaitlist && <Check size={12} />}</span>
                          </button>
                          <button
                            type="button"
                            className={cx('dex-ui-choice', !!splitSharedWaitlist && 'is-active')}
                            aria-pressed={!!splitSharedWaitlist}
                            onClick={() => setSplitSharedWaitlist(true)}
                          >
                            <span className="dex-ui-choice-body">
                              <span className="dex-ui-choice-title">{isDe ? 'Eine gemeinsame Warteliste' : 'One shared waitlist'}</span>
                              <span className="dex-ui-choice-desc">
                                {isDe
                                  ? 'Wer am längsten wartet, rückt zuerst nach — egal, in welcher Gruppe der Platz frei wird.'
                                  : 'Whoever has waited longest moves up first — no matter which group the seat frees up in.'}
                              </span>
                            </span>
                            <span className="dex-ui-choice-check">{!!splitSharedWaitlist && <Check size={12} />}</span>
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                )}
                </>)}
              </div>

              </div>{/* v15.6: close hauptGreyoutWrapperStyle div (Step 4) */}
                </>
              )}{/* v28.76: Ende Klammer-Fall / Normalfall */}
              </div>{/* v31.2: Ende Abschnitt Plätze */}

              {/* ===== v31.2: Abschnitt 2 — Fristen. „Anmeldung ab / bis" und
                  „Abmeldung bis" als eine Familie, direkt darunter die
                  Selbst-Abmeldung (Frist und Selbst-Abmeldung gehören
                  zusammen). ===== */}
              {/* v28.31: Greyout-Wrapper beginnt erst NACH den Fristen. Bis v28.30
                  lag der Fristen-Block mit in der Huelle (opacity + pointer-events:
                  none) — im Klammer-Modus liess sich der Abschnitt deshalb nicht
                  einmal aufklappen, obwohl die Klammer seit v28.20 eine EIGENE,
                  wirksame Anmeldefrist haben kann. */}
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">
                  <Icon iconName="Clock" style={{ fontSize: 14 }} />
                  {isDe ? 'Fristen — bis wann?' : 'Deadlines — until when?'}
                </div>
              <div className="dex-ui-card" style={{ marginBottom: 12 }}>
                {visHeader('vis_fristen', <StepBadge n={19} />, <>{isDe ? 'Anmelde- und Abmeldefristen' : 'Registration & cancellation deadlines'}<InfoTooltip text={isDe
                    ? 'Bis wann können sich Teilnehmer anmelden bzw. fristgerecht abmelden? Die Abmeldefrist ist die kommunizierte Deadline — abmelden geht danach standardmäßig weiterhin bis zum Event-Ende, die Organizer werden dann aber automatisch informiert. Über die Option unter den Fristen lässt sich die Selbst-Abmeldung nach der Frist auch komplett sperren. Beide Werte werden anhand des Event-Datums automatisch vorgeschlagen, du kannst sie jederzeit überschreiben.'
                    : 'Until when can attendees register or cancel within the deadline? The cancellation deadline is the communicated cutoff — by default cancelling remains possible until the event ends, but organizers are then notified automatically. The option below the deadlines can instead lock self-cancellation completely after the cutoff. Both values are auto-suggested from the event date and can be overridden at any time.'} /></>)}
                {isVisOpen('vis_fristen') && (<>
              {/* v29.75: Klartext-Zusammenfassung fuer die Klammer — WAS gilt
                  gerade fuer alle Sub-Events? Die einzelnen Regeln (Freischalt-
                  Regel, Klammer-Anmeldefrist, Abmeldefrist, Sichtbarkeit)
                  stehen verstreut in dieser Sektion und in der Sichtbarkeit;
                  hier laufen sie als ein lesbarer Absatz zusammen, inklusive
                  der Zahl abweichender Sub-Events. */}
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
                  <div className="dex-ui-callout dex-ui-callout--success" style={{ flexDirection: 'column', gap: 4, marginBottom: 14 }}>
                    <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>
                      {isDe ? `Aktuell gilt für alle ${subEvents.length} ${term}: ` : `Currently, for all ${subEvents.length} ${term}: `}
                    </strong>
                    <ul style={{ margin: 0, paddingLeft: 18 }}>
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
              {/* v31.2: Der Hinweis ist standardmäßig zu (vorher defaultOpen) —
                  der Absatz füllte einen halben Bildschirm, bevor das erste
                  Feld kam; die Zusammenfassung darüber sagt schon, was gilt.
                  Text an v30.6 angepasst: die rollierende Regel SPERRT die
                  Felder je Termin nicht mehr, sie belegt sie nur vor. */}
              {subEventsOnlyMode && (
                <WizardHint
                  isDe={isDe}
                  title={isDe ? 'Frist ergibt sich aus den Sub-Events' : 'Deadline derives from the sub-events'}
                  style={{ marginBottom: 12 }}
                >
                  <div>
                    {isDe ? (
                      <>
                        Die Klammer selbst ist <strong>nicht buchbar</strong> — hier stellst du die Anmelde-Logik für <strong>alle {childTermPlural || 'Sub-Events'}</strong> zentral ein: <strong>&bdquo;Anmeldung ab&ldquo;</strong> (festes Datum oder rollierend), <strong>&bdquo;Anmeldung bis&ldquo;</strong> und <strong>&bdquo;Abmeldung bis&ldquo;</strong> (jeweils festes Datum, das in alle {childTermPlural || 'Sub-Events'} übernommen wird, oder rollierend — dann wird die Frist <strong>je Termin automatisch ausgerechnet</strong>, auch für später hinzugefügte). Eine gesetzte feste Klammer-Anmeldefrist wirkt zusätzlich als <strong>harter Schluss fürs gesamte Event</strong>; ohne sie bleibt die Anmeldung offen, solange mindestens ein {childTermSingular || 'Sub-Event'} offen ist. Im jeweiligen Reiter kannst du die Frist pro {childTermSingular || 'Sub-Event'} anpassen — bei aktiver <strong>rollierender</strong> Regel ist der Regel-Wert dort nur vorbelegt und bleibt überschreibbar.
                      </>
                    ) : (
                      <>
                        The bracket itself is <strong>not bookable</strong> — here you configure the registration logic for <strong>all {childTermPlural || 'sub-events'}</strong> centrally: <strong>“registration opens”</strong> (fixed date or rolling), <strong>“registration until”</strong> and <strong>“cancellation until”</strong> (a fixed date copied to all {childTermPlural || 'sub-events'}, or rolling — then the deadline is <strong>computed per date automatically</strong>, including dates added later). A fixed bracket registration deadline additionally acts as a <strong>hard cutoff for the entire event</strong>; without one, registration stays open as long as at least one {childTermSingular || 'sub-event'} is open. You can adjust the deadline per {childTermSingular || 'sub-event'} in its tab — with an active <strong>rolling</strong> rule the rule value is only pre-filled there and stays overridable.
                      </>
                    )}
                  </div>
                </WizardHint>
              )}
              {/* v29.75: Freischalt-Regel („Anmeldung ab") — von Schritt 1
                  (Kalender-Block) hierher gezogen: Sie ist eine Anmelde-Regel
                  und gehoert neben Anmelde- und Abmeldefrist.
                  v29.77: „Anmeldung ab" steht IMMER hier — bei Sub-Event-Events
                  als feste/rollierende Freischalt-Regel, bei reinen Haupt-
                  events als das Aktivierungsdatum aus den Grundlagen (bewusst
                  doppelt angeboten, damit die Sektion immer die volle Logik
                  „Anmeldung ab / Anmeldung bis / Abmeldung bis" zeigt).
                  v29.75: User-Wortlaut — die drei Felder heissen „Anmeldung
                  ab" / „Anmeldung bis" / „Abmeldung bis"; die Namen bleiben. */}
              {subEventsOptIn ? (
                <div className="dex-ui-card dex-ui-card--soft" style={{ marginBottom: 14 }}>
                  <div className="dex-ui-label">{isDe ? 'Anmeldung ab' : 'Registration opens'}</div>
                  <p className="dex-ui-help" style={{ margin: '-2px 0 10px' }}>
                    {isDe ? 'Ab wann können sich Teilnehmer anmelden? Leer = sofort.' : 'From when can attendees register? Empty = immediately.'}
                  </p>
                  {/* v30.2: Datum ZUERST, Schalter darunter — dieselbe
                      Reihenfolge wie bei „Anmeldung bis"/„Abmeldung bis". Das
                      Datumsfeld verschwindet, sobald rollierend gewaehlt ist. */}
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
                        <p className="dex-ui-help">
                          {isDe
                            ? 'Alle Termine öffnen gemeinsam zu diesem Zeitpunkt.'
                            : 'All dates open together at this moment.'}
                        </p>
                      )}
                    </div>
                  )}
                  <label className="dex-ui-switch" style={{ marginTop: 10 }}>
                    <input
                      type="checkbox"
                      checked={openRuleEnabled}
                      onChange={e => setOpenRuleEnabled(e.target.checked)}
                    />
                    <span className="dex-ui-switch-track" />
                    <span className="dex-ui-switch-label">{isDe ? 'Rollierend je Termin' : 'Rolling per date'}</span>
                  </label>
                  {openRuleEnabled && (
                    <div style={{ marginTop: 10 }}>
                      <div className="dex-ui-inline" style={{ fontSize: '0.88rem' }}>
                        {isDe ? 'Anmeldung möglich ab' : 'Registration opens'}
                        <input
                          type="number"
                          min={1}
                          max={365}
                          className="dex-ui-input dex-ui-input--sm"
                          value={openRuleDays}
                          onChange={e => { const v = parseInt(e.target.value, 10); setOpenRuleDays(isFinite(v) && v > 0 ? Math.min(v, 365) : 1); }}
                          style={{ width: 76, textAlign: 'center' }}
                        />
                        {isDe ? 'Tage vor' : 'days before'}
                        <select
                          className="dex-ui-select"
                          value={openRuleMode}
                          onChange={e => setOpenRuleMode(e.target.value === 'week' ? 'week' : 'day')}
                          style={{ width: 'auto' }}
                        >
                          <option value="day">{isDe ? 'dem jeweiligen Termin' : 'each date'}</option>
                          <option value="week">{isDe ? 'dem Montag der jeweiligen Woche' : 'the Monday of its week'}</option>
                        </select>
                      </div>
                      <p className="dex-ui-help">
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
                <div className="dex-ui-card dex-ui-card--soft" style={{ marginBottom: 14 }}>
                  <div className="dex-ui-label">{isDe ? 'Anmeldung ab' : 'Registration opens'}</div>
                  <p className="dex-ui-help" style={{ margin: '-2px 0 10px' }}>
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
              <div className="dex-ui-grid-2">
                {/* v30.2: gleiche Karte wie „Anmeldung ab" — die drei
                    Fristen sollen als EINE Familie lesbar sein. */}
                <div className="dex-ui-card dex-ui-card--soft">
                  <div className="dex-ui-label">
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
                  </div>
                  <p className="dex-ui-help" style={{ margin: '-2px 0 10px' }}>
                    {subEventsOnlyMode
                      ? (isDe ? 'Harter Anmeldeschluss fürs gesamte Event. Leer = offen, solange ein Termin offen ist.' : 'Hard cutoff for the entire event. Empty = open as long as one date is open.')
                      : (isDe ? 'Danach ist der Anmelden-Knopf gesperrt; Organizer können weiterhin manuell anmelden.' : 'After this the register button is locked; organizers can still add attendees manually.')}
                  </p>
                  {/* v29.76: Entweder festes Datum ODER rollierend je Termin.
                      Beim Umschalten auf rollierend wird das feste Datum
                      geleert — zwei gleichzeitig wirkende Fristen wuerde
                      niemand mehr nachvollziehen koennen.
                      v29.77: fuer ALLE Sub-Event-Events, nicht nur Kalender.
                      v30.2: Datum ZUERST, Schalter darunter — konsistent
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
                    <label className="dex-ui-switch" style={{ marginTop: 10 }}>
                      <input
                        type="checkbox"
                        checked={regRuleEnabled}
                        onChange={e => {
                          setRegRuleEnabled(e.target.checked);
                          if (e.target.checked) {
                            if (subEventsOnlyMode) setKlammerDeadline(''); else setRegistrationDeadline('');
                          }
                        }}
                      />
                      <span className="dex-ui-switch-track" />
                      <span className="dex-ui-switch-label">{isDe ? 'Rollierend je Termin' : 'Rolling per date'}</span>
                    </label>
                  )}
                  {(subEventsOptIn && regRuleEnabled) && (
                    <div style={{ marginTop: 10 }}>
                      <div className="dex-ui-inline" style={{ fontSize: '0.88rem' }}>
                        {isDe ? 'Anmeldung bis' : 'Registration until'}
                        <input
                          type="number"
                          min={1}
                          max={365}
                          className="dex-ui-input dex-ui-input--sm"
                          value={regRuleAmount}
                          onChange={e => { const v = parseInt(e.target.value, 10); setRegRuleAmount(isFinite(v) && v > 0 ? Math.min(v, 365) : 1); }}
                          style={{ width: 76, textAlign: 'center' }}
                        />
                        {/* v29.77: Einheit im Singular, wenn die Zahl 1 ist. */}
                        <select
                          className="dex-ui-select"
                          value={regRuleUnit}
                          onChange={e => setRegRuleUnit(e.target.value === 'hours' ? 'hours' : 'days')}
                          style={{ width: 'auto' }}
                        >
                          <option value="days">{isDe ? (regRuleAmount === 1 ? 'Tag' : 'Tage') : (regRuleAmount === 1 ? 'day' : 'days')}</option>
                          <option value="hours">{isDe ? (regRuleAmount === 1 ? 'Stunde' : 'Stunden') : (regRuleAmount === 1 ? 'hour' : 'hours')}</option>
                        </select>
                        {isDe ? 'vor dem jeweiligen Termin' : 'before each date'}
                      </div>
                      <p className="dex-ui-help">
                        {isDe
                          ? 'Die Frist wird je Termin ausgerechnet und in dessen Einstellungen geschrieben — auch für später hinzugefügte Termine.'
                          : 'The deadline is computed per date and written into its settings — including dates added later.'}
                      </p>
                    </div>
                  )}
                  {subEventsOnlyMode && !klammerDeadline && effectiveKlammerDeadline && (
                    <p className="dex-ui-help">
                      {isDe
                        ? <>Aktuell effektiv: <strong>{new Date(effectiveKlammerDeadline).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })}</strong> (späteste Sub-Event-Frist)</>
                        : <>Currently effective: <strong>{new Date(effectiveKlammerDeadline).toLocaleString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })}</strong> (latest sub-event deadline)</>}
                    </p>
                  )}
                  {subEventsOnlyMode && klammerDeadline && (
                    <p className="dex-ui-help">
                      {isDe
                        ? <>Gilt als Anmeldeschluss fürs <strong>gesamte Event</strong> und wurde in alle {childTermPlural || 'Sub-Event'}-Tabs übernommen — dort pro {childTermSingular || 'Sub-Event'} anpassbar (z.B. früherer Schluss).</>
                        : <>Acts as the registration cutoff for the <strong>entire event</strong> and has been copied to all {childTermPlural || 'sub-event'} tabs — adjustable there per {childTermSingular || 'sub-event'} (e.g. an earlier cutoff).</>}
                    </p>
                  )}
                </div>
                {/* v28.31: Die Abmeldefrist gehört bei einer Klammer zu den
                    Sub-Events — hier wäre sie wirkungslos. Nur dieses eine Feld
                    ausgrauen, die Klammer-Anmeldefrist links bleibt bedienbar. */}
                {/* v29.25: Ohne Selbst-Abmeldung gibt es keine Abmeldefrist —
                    statt des Datumsfelds steht der Grund. */}
                {!userCancelAllowed ? (
                  <div className="dex-ui-card dex-ui-card--soft dex-ui-card--muted">
                    <div className="dex-ui-label">{isDe ? 'Abmeldung bis' : 'Cancellation until'}</div>
                    <div className="dex-ui-callout dex-ui-callout--neutral" style={{ borderStyle: 'dashed' }}>
                      {isDe
                        ? 'Entfällt — die Abmeldung durch Teilnehmer ist unten ausgeschaltet. Abmelden können nur Organizer und Admins.'
                        : 'Not applicable — self-cancellation is switched off below. Only organizers and admins can cancel.'}
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
                <div className="dex-ui-card dex-ui-card--soft">
                  <div className="dex-ui-label">
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
                  </div>
                  <p className="dex-ui-help" style={{ margin: '-2px 0 10px' }}>
                    {isDe ? 'Die kommunizierte Frist. Danach bekommst du bei jeder Abmeldung eine Info-Mail.' : 'The communicated cutoff. After it you get an info mail for every cancellation.'}
                  </p>
                  {/* v29.76: wie „Anmeldung bis" — festes Datum ODER rollierend.
                      v30.2: Datum ZUERST, Schalter darunter. */}
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
                    <label className="dex-ui-switch" style={{ marginTop: 10 }}>
                      <input
                        type="checkbox"
                        checked={cancelRuleEnabled}
                        onChange={e => {
                          setCancelRuleEnabled(e.target.checked);
                          if (e.target.checked) setLastDeregisterDate('');
                        }}
                      />
                      <span className="dex-ui-switch-track" />
                      <span className="dex-ui-switch-label">{isDe ? 'Rollierend je Termin' : 'Rolling per date'}</span>
                    </label>
                  )}
                  {(subEventsOptIn && cancelRuleEnabled) && (
                    <div style={{ marginTop: 10 }}>
                      <div className="dex-ui-inline" style={{ fontSize: '0.88rem' }}>
                        {isDe ? 'Abmeldung bis' : 'Cancellation until'}
                        <input
                          type="number"
                          min={1}
                          max={365}
                          className="dex-ui-input dex-ui-input--sm"
                          value={cancelRuleAmount}
                          onChange={e => { const v = parseInt(e.target.value, 10); setCancelRuleAmount(isFinite(v) && v > 0 ? Math.min(v, 365) : 1); }}
                          style={{ width: 76, textAlign: 'center' }}
                        />
                        {/* v29.77: Einheit im Singular, wenn die Zahl 1 ist. */}
                        <select
                          className="dex-ui-select"
                          value={cancelRuleUnit}
                          onChange={e => setCancelRuleUnit(e.target.value === 'hours' ? 'hours' : 'days')}
                          style={{ width: 'auto' }}
                        >
                          <option value="days">{isDe ? (cancelRuleAmount === 1 ? 'Tag' : 'Tage') : (cancelRuleAmount === 1 ? 'day' : 'days')}</option>
                          <option value="hours">{isDe ? (cancelRuleAmount === 1 ? 'Stunde' : 'Stunden') : (cancelRuleAmount === 1 ? 'hour' : 'hours')}</option>
                        </select>
                        {/* v29.77: Abmelden darf auch NACH dem Termin-Beginn
                            noch offen sein — Richtung waehlbar. */}
                        <select
                          className="dex-ui-select"
                          value={cancelRuleAfter ? 'after' : 'before'}
                          onChange={e => setCancelRuleAfter(e.target.value === 'after')}
                          style={{ width: 'auto' }}
                        >
                          <option value="before">{isDe ? 'vor dem jeweiligen Termin' : 'before each date'}</option>
                          <option value="after">{isDe ? 'nach dem jeweiligen Termin' : 'after each date starts'}</option>
                        </select>
                      </div>
                      <p className="dex-ui-help">
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
                <p className="dex-ui-help" style={{ marginTop: 8 }}>
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
                  Klammer mit Sub-Event-Fristen). Beides event-weit.
                  v31.2: als Schalter-Zeilen; die ausgeschaltete Stufe 1 bleibt
                  orange umrandet, weil sie Teilnehmern etwas wegnimmt. */}
              <div className="dex-ui-stack" style={{ marginTop: 12 }}>
                <label
                  className={cx('dex-ui-toggle-row', userCancelAllowed && 'is-active')}
                  style={!userCancelAllowed ? { borderColor: 'var(--dex-orange, #ed8b00)', background: 'rgba(237,139,0,0.06)' } : undefined}
                >
                  <input
                    type="checkbox"
                    checked={userCancelAllowed}
                    onChange={e => setUserCancelAllowed(e.target.checked)}
                  />
                  <span className="dex-ui-toggle-row-body">
                    <span className="dex-ui-toggle-row-title">{isDe ? 'Teilnehmer dürfen sich selbst abmelden' : 'Attendees may cancel themselves'}</span>
                    <span className="dex-ui-toggle-row-desc">
                      {isDe ? (
                        <>
                          <strong>An (Standard):</strong> Abmelden geht unter „Meine Events“ und über den Link in der Bestätigungsmail.<br />
                          <strong>Aus:</strong> Teilnehmer können sich <strong>gar nicht selbst abmelden</strong>; eine Abmeldefrist entfällt. Abmelden können nur Organizer und Admins über das Organizer Center — dort steht neben den Angemeldeten zusätzlich der <strong>No-Show</strong>-Knopf für Nicht-Erschienene.
                        </>
                      ) : (
                        <>
                          <strong>On (default):</strong> attendees cancel under “My events” or via the link in the confirmation mail.<br />
                          <strong>Off:</strong> attendees <strong>cannot cancel themselves at all</strong>; there is no cancellation deadline. Only organizers and admins can cancel, via the Organizer Center — which then also shows a <strong>No-Show</strong> button next to registered attendees.
                        </>
                      )}
                    </span>
                  </span>
                </label>
                {userCancelAllowed && (lastDeregisterDate || subEventsOnlyMode) && (
                  <label className={cx('dex-ui-toggle-row', !noCancelAfterDeadline && 'is-active')}>
                    <input
                      type="checkbox"
                      checked={!noCancelAfterDeadline}
                      onChange={e => setNoCancelAfterDeadline(!e.target.checked)}
                    />
                    <span className="dex-ui-toggle-row-body">
                      <span className="dex-ui-toggle-row-title">{isDe ? 'Abmeldung auch nach der Abmeldefrist erlauben' : 'Also allow cancelling after the deadline'}</span>
                      <span className="dex-ui-toggle-row-desc">
                        {isDe ? (
                          <>
                            <strong>An (Standard):</strong> Teilnehmer können sich auch nach der Frist abmelden — die Organizer bekommen dann automatisch eine Info-Mail („Verspätete Abmeldung“) mit Name und E-Mail der Person.<br />
                            <strong>Aus:</strong> Nach der Frist können sich Teilnehmer <strong>nicht mehr selbst abmelden</strong> — nur noch Organizer und Admins über das Organizer Center, dort ab der Frist ebenfalls mit <strong>No-Show</strong>-Knopf.
                            {subEventsOnlyMode ? <> Bei einer Klammer greift die Sperre je {childTermSingular || 'Sub-Event'} nach dessen eigener Abmeldefrist.</> : null}
                          </>
                        ) : (
                          <>
                            <strong>On (default):</strong> attendees can still cancel after the deadline — the organizers then automatically receive an info email (“late cancellation”) with the person’s name and email.<br />
                            <strong>Off:</strong> once the deadline has passed, attendees <strong>can no longer cancel themselves</strong> — only organizers and admins can, via the Organizer Center, which then also shows the <strong>No-Show</strong> button from the deadline on.
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
              </div>{/* v31.2: Ende Abschnitt Fristen */}

              {/* ===== v31.2: Abschnitt 3 — Sichtbarkeit. Jetzt ZULETZT: erst
                  „wie viele", dann „bis wann", dann „wer sieht es". ===== */}
              {/* v19.27: Sichtbarkeit (Standortfilter + Mailverteiler) NICHT
                  im Greyout-Wrapper — sie bleibt im „Nur Sub-Events"-Modus
                  (Klammer) editierbar, weil sie steuert, WER das ganze Event
                  sieht. */}
              {/* v9.24: Sichtbarkeits-Steuerungen aus Step 0 hierher verschoben.
                  Die Frage 'wer darf das Event sehen' passt logisch zu Kapazität/Fristen
                  als 'Wer-Wann-Wieviel' und entlastet Step 0 (Grundlagen). */}
              <div className="dex-ui-section" style={{ marginTop: 22 }}>
                <div className="dex-ui-section-title">
                  <Icon iconName="Hide3" style={{ fontSize: 14 }} />
                  {isDe ? 'Sichtbarkeit — wer sieht das Event?' : 'Visibility — who sees the event?'}
                </div>
                <p className="dex-ui-section-desc">
                  {isDe
                    ? <>Ohne Filter sehen <strong>alle Mitarbeiter von Deloitte Deutschland</strong> das Event und können sich anmelden. Mit Standortfilter und Mailverteilern grenzt du den Kreis ein — wer außerhalb liegt, findet das Event nicht in seiner Übersicht und kann sich nicht anmelden.</>
                    : <>Without filters, <strong>all Deloitte Germany employees</strong> see the event and can register. Use the location filter and mailing lists to narrow the audience — anyone outside will not find the event in their overview and cannot register.</>}
                </p>

                {/* v26.89: Live-Zusammenfassung „Aktuell eingestellt …" steht
                    GANZ OBEN — direkt über dem Standortfilter, damit man den
                    aktuellen Sichtbarkeits-Stand sofort sieht. */}
                {renderVisibilitySummaryBox(
                  locationFilter.split(',').map(s => s.trim()).filter(Boolean),
                  audience,
                  filterMode,
                  (excludedUsers || []).length
                )}
                {/* v28.76: Widerspruch Klammer ↔ Sub-Events benennen (s.o.). */}
                {renderKlammerVisibilityMismatch()}

                <div className="dex-ui-card" style={{ marginBottom: 12 }}>
                  {visHeader('vis_locfilter', <StepBadge n={20} />, isDe ? 'Standortfilter' : 'Location filter')}
                  {isVisOpen('vis_locfilter') && (<>
                  <p className="dex-ui-help" style={{ margin: '0 0 12px' }}>
                    {isDe
                      ? <>Nur Mitarbeiter mit einem der gewählten Standorte sehen das Event. <em>Beispiel: &bdquo;Köln&ldquo; und &bdquo;Düsseldorf&ldquo; → alle anderen Standorte sehen es nicht.</em></>
                      : <>Only employees from one of the selected locations see the event. <em>Example: &bdquo;Cologne&ldquo; and &bdquo;Düsseldorf&ldquo; → every other location will not see it.</em></>}
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
                    <p className="dex-ui-help" style={{ color: 'var(--dex-green-darker, #4a7c1f)', marginTop: 8 }}>
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
                    Reihenfolge der Sektion unverändert bleibt. */}
                <AudiencePicker
                  value={audience}
                  onChange={setAudience}
                  locationFilter={locationFilter}
                  filterMode={filterMode}
                  isDe={isDe}
                  excludedUsers={excludedUsers}
                  onExcludedUsersChange={setExcludedUsers}
                  headerSlot={visHeader('vis_audience', <StepBadge n={21} />, isDe ? 'Mailverteiler / einzelne User' : 'Mailing lists / individual users')}
                  bodyOpen={isVisOpen('vis_audience')}
                  cardBgPrimary={zebraS3Bg()}
                  visibilityTabs={subEvents.length > 0 ? [
                    { id: 'main', title: subEventsOnlyMode ? (isDe ? 'Klammer' : 'Bracket') : (isDe ? 'Hauptevent' : 'Main event'), locationFilter, audience, filterMode },
                    ...subEvents.map(s => ({ id: s.id, title: (shortSubEventTitle(s.title, title) || (isDe ? 'Sub-Event' : 'Sub-event')).trim(), locationFilter: s.locationFilter || '', audience: s.audience || '', filterMode: (s.filterMode || 'AND') as 'AND' | 'OR' })),
                  ] : undefined}
                  middleSlot={(locationFilter && audience) ? (
                    /* Filterverknüpfung: nur sichtbar wenn beide Bereiche
                       (Standortfilter + Mailverteiler) Werte haben — sonst gibt
                       es nichts zu kombinieren. v31.2: als zwei Kacheln mit je
                       einer Zeile Folge; die Beispiele stehen im Tooltip. */
                    <div className="dex-ui-card" style={{ marginBottom: 12 }}>
                      <div className="dex-ui-label">
                        <StepBadge n={22} />
                        {isDe ? 'Wie greifen die beiden Filter zusammen?' : 'How do the two filters combine?'}
                        <InfoTooltip text={isDe ? (
                          <>
                            Bestimmt, wie der <strong>Standortfilter</strong> und der <strong>Mailverteiler / einzelne User</strong> miteinander kombiniert werden — also welche Bedingungen für eine Person erfüllt sein müssen, damit sie das Event in ihrer Liste sieht.<br /><br />
                            <strong>ODER (Default):</strong> <em>einer der beiden Filter reicht.</em> Beispiel: <strong>Standort = Köln</strong>, <strong>Verteiler = SAPALL</strong> → jede Person, die <strong>in Köln</strong> sitzt <strong>ODER</strong> in <strong>SAPALL</strong> ist, sieht das Event. Praktisch wenn du bewusst einen <strong>breiten Empfängerkreis</strong> willst (z.B. Standort-Mitarbeiter <strong>plus</strong> Fachgruppe).<br /><br />
                            <strong>UND:</strong> <em>beide Filter müssen zutreffen.</em> Beispiel: <strong>Standort = Köln</strong>, <strong>Verteiler = SAPALL</strong> → nur wer <strong>in Köln</strong> sitzt <strong>UND</strong> in <strong>SAPALL</strong> ist, sieht das Event. Praktisch wenn du den Empfängerkreis <strong>strikt eingrenzen</strong> willst (z.B. nur die SAP-Kollegen <strong>am Standort Köln</strong>).<br /><br />
                            <strong>Hinweis:</strong> diese Auswahl erscheint nur, wenn <strong>beide</strong> Filter gesetzt sind — sonst gibt es nichts zu kombinieren.
                          </>
                        ) : (
                          <>
                            Defines how the <strong>location filter</strong> and the <strong>mailing lists / individual users</strong> are combined — i.e. which conditions must be true for a person before the event shows up in their list.<br /><br />
                            <strong>OR (default):</strong> <em>either filter is enough.</em> Example: <strong>Location = Cologne</strong>, <strong>list = SAPALL</strong> → anyone <strong>in Cologne</strong> <strong>OR</strong> in <strong>SAPALL</strong> sees the event. Useful when you intentionally want a <strong>broad audience</strong> (e.g. location staff <strong>plus</strong> a domain group).<br /><br />
                            <strong>AND:</strong> <em>both filters must match.</em> Example: <strong>Location = Cologne</strong>, <strong>list = SAPALL</strong> → only people <strong>in Cologne</strong> <strong>AND</strong> in <strong>SAPALL</strong> see the event. Useful when you want to <strong>strictly narrow</strong> the audience (e.g. only the SAP colleagues <strong>at the Cologne site</strong>).<br /><br />
                            <strong>Note:</strong> this selector only appears when <strong>both</strong> filters are set — otherwise there is nothing to combine.
                          </>
                        )} />
                      </div>
                      <p className="dex-ui-help" style={{ margin: '-2px 0 10px' }}>
                        {isDe ? 'Beide Filter sind gesetzt — muss eine Person einen davon erfüllen oder beide?' : 'Both filters are set — does a person need to match one of them or both?'}
                      </p>
                      <div className="dex-ui-grid-2">
                        <button
                          type="button"
                          className={cx('dex-ui-choice', filterMode === 'OR' && 'is-active')}
                          aria-pressed={filterMode === 'OR'}
                          onClick={() => setFilterMode('OR')}
                        >
                          <span className="dex-ui-choice-body">
                            <span className="dex-ui-choice-title">{isDe ? 'ODER — einer der Filter reicht' : 'OR — one filter is enough'}</span>
                            <span className="dex-ui-choice-desc">
                              {isDe ? 'Breiter Kreis: z.B. alle in Köln plus die Fachgruppe. Standard.' : 'Broad audience: e.g. everyone in Cologne plus the domain group. Default.'}
                            </span>
                          </span>
                          <span className="dex-ui-choice-check">{filterMode === 'OR' && <Check size={12} />}</span>
                        </button>
                        <button
                          type="button"
                          className={cx('dex-ui-choice', filterMode === 'AND' && 'is-active')}
                          aria-pressed={filterMode === 'AND'}
                          onClick={() => setFilterMode('AND')}
                        >
                          <span className="dex-ui-choice-body">
                            <span className="dex-ui-choice-title">{isDe ? 'UND — beides muss zutreffen' : 'AND — both must match'}</span>
                            <span className="dex-ui-choice-desc">
                              {isDe ? 'Enger Kreis: z.B. nur die SAP-Kollegen am Standort Köln.' : 'Narrow audience: e.g. only the SAP colleagues at the Cologne site.'}
                            </span>
                          </span>
                          <span className="dex-ui-choice-check">{filterMode === 'AND' && <Check size={12} />}</span>
                        </button>
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
                    /* v31.2 (Leitfaden 2a′): Knopf links direkt hinter dem Text
                       statt rechts außen. Der Kasten selbst bleibt bewusst
                       NICHT klickbar — die Aktion schreibt in den Verteiler,
                       ein Fehlklick auf die Fläche soll das nicht auslösen. */
                    <div className="dex-ui-callout dex-ui-callout--warn" style={{ alignItems: 'center', flexWrap: 'wrap', gap: 10, marginBottom: 12 }}>
                      <span className="dex-ui-callout-icon"><AlertCircle size={18} /></span>
                      <span style={{ minWidth: 0 }}>
                        {isDe
                          ? <>In den {childTermPlural || 'Sub-Events'} stehen <strong>{missing.length} Verteiler/Personen</strong>, die der Klammer fehlen. Der Zugang läuft immer über die Klammer — wer hier fehlt, sieht das Event nicht.</>
                          : <>The {childTermPlural || 'sub-events'} contain <strong>{missing.length} lists/people</strong> missing from the bracket. Access always goes through the bracket — anyone missing here cannot see the event.</>}
                      </span>
                      <button
                        type="button"
                        className="btn btn-primary dex-ui-btn-sm"
                        style={{ flexShrink: 0 }}
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
                  <label className={cx('dex-ui-toggle-row', visAllSubs && 'is-active')} style={{ marginBottom: 12 }}>
                    <input
                      type="checkbox"
                      checked={visAllSubs}
                      onChange={e => { visAllSubsTouchedRef.current = true; setVisAllSubs(e.target.checked); }}
                    />
                    <span className="dex-ui-toggle-row-body">
                      <span className="dex-ui-toggle-row-title">{isDe ? `Sichtbarkeit gilt für alle ${childTermPlural || 'Sub-Events'}` : `Visibility applies to all ${childTermPlural || 'sub-events'}`}</span>
                      <span className="dex-ui-toggle-row-desc">
                        {isDe
                          ? <>Standortfilter, Mailverteiler und Verknüpfung von oben werden in alle {subEvents.length} {childTermPlural || 'Sub-Events'} übernommen und bleiben synchron, solange der Haken gesetzt ist. Die Sichtbarkeit in den {childTermSingular || 'Sub-Event'}-Reitern ist dann gesperrt — Haken entfernen, um wieder je {childTermSingular || 'Sub-Event'} abzuweichen.</>
                          : <>The location filter, mailing lists and combination above are applied to all {subEvents.length} {childTermPlural || 'sub-events'} and stay in sync while the box is checked. Visibility in the {childTermSingular || 'sub-event'} tabs is locked then — uncheck to deviate per {childTermSingular || 'sub-event'} again.</>}
                      </span>
                    </span>
                  </label>
                )}

                {/* v23.6: Assistenz-Sichtbarkeit — eigener Baustein. Bewusst
                    AUSSERHALB des Greyout-Wrappers (laufzeit-/sichtbarkeitsrelevant,
                    wie der AudiencePicker oben — auch im Klammer-Modus editierbar). */}
                <div className="dex-ui-card" style={{ marginBottom: 12 }}>
                  {visHeader('vis_assist', <StepBadge n={(locationFilter && audience) ? 23 : 22} />, isDe ? 'Sichtbarkeit für Assistenzen' : 'Visibility for assistants')}
                  {isVisOpen('vis_assist') && (
                    <label className={cx('dex-ui-toggle-row', assistantsCanSee && 'is-active')}>
                      <input type="checkbox" checked={assistantsCanSee} onChange={e => setAssistantsCanSee(e.target.checked)} />
                      <span className="dex-ui-toggle-row-body">
                        <span className="dex-ui-toggle-row-title">
                          {isDe ? 'Alle Assistenzen sehen dieses Event — auch außerhalb des Filterkreises' : 'All assistants see this event — even outside the filter audience'}
                          <InfoTooltip text={isDe
                            ? <>Standardmäßig sieht eine <strong>Assistenz</strong> (Personen mit dem Job-Title „Assistenz“) nur Events, die sie auch als normale Nutzerin/normaler Nutzer sehen würde — also nur, wenn sie selbst in den oben gesetzten Standort-/Verteiler-Kreis fällt. Aktivierst du diese Option, sehen <strong>alle Assistenzen</strong> dieses Event in ihrer Übersicht, unabhängig von den Filtern oben — damit sie z.B. stellvertretend einen Partner oder eine Direktorin anmelden können.</>
                            : <>By default an <strong>assistant</strong> (people whose job title is „Assistenz“) only sees events they would see as a regular user — i.e. only if they themselves fall within the location/distribution audience set above. Enable this option to let <strong>all assistants</strong> see this event in their overview regardless of the filters above — so they can register a partner or director on their behalf, for example.</>} />
                        </span>
                        <span className="dex-ui-toggle-row-desc">
                          {isDe
                            ? 'Aus (Standard): Eine Assistenz (Job-Title „Assistenz“) sieht nur, was sie auch als normale Nutzerin sähe. An: Alle Assistenzen sehen das Event, um z.B. stellvertretend einen Partner oder eine Direktorin anzumelden.'
                            : 'Off (default): an assistant (job title “Assistenz”) only sees what they would see as a regular user. On: all assistants see the event, e.g. to register a partner or director on their behalf.'}
                        </span>
                      </span>
                    </label>
                  )}
                </div>
              </div>
              </div>{/* v15.0: close activeCapacityTabIdx===0 wrapper (Top-Level Sichtbarkeit/Deadlines/Max/Split) */}

              </div>
  );
};
