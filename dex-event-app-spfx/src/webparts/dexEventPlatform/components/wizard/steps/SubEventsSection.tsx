/* SubEventsSection — aus EventCreationPage.tsx ausgelagert (Zeilen 11784-12928 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; einzige
 * Aenderung ist die Anzeige-Bedingung: aus `currentStep === 0 && activeScopeIdx === 0` wurde das Prop `visible`.
 * `visible` schaltet display:none statt unmount — Eingaben ueberleben den
 * Schrittwechsel genauso wie vorher. */
import * as React from 'react';
import { StepBadge } from '../../wizard/StepBadge';
import WizardHint from '../../WizardHint';
import { InfoTooltip } from '../../InfoTooltip';
import DatePicker from 'react-datepicker';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { Plus, X } from '../../Icons';
import { SubEventDraft } from '../../wizard/wizardTypes';
import ImageCropModal from '../../ImageCropModal';
export interface SubEventsSectionProps {
  visible: boolean;
  activeScopeIdx: number;
  audience: string;
  berlinLocalToUtcIso: (localStr: string) => string;
  childGender: "" | "m" | "f" | "n";
  childTermPlural: string;
  childTermSingular: string;
  confirmDialog: (message: React.ReactNode, opts?: import("../../../context/DialogContext").ConfirmOptions) => Promise<boolean>;
  currentStep: number;
  customTermMode: boolean;
  dayKeyOfSub: (se: SubEventDraft) => string;
  endDate: string;
  filterMode: "AND" | "OR";
  goToScopeBar: () => void;
  isDe: boolean;
  isoToLocal: (iso: string) => string;
  klammerDeadline: string;
  locationFilter: string;
  mainEventLabel: string;
  mainEventLabelMode: "none" | "default" | "custom";
  openRuleDays: number;
  openRuleEnabled: boolean;
  openRuleMode: "day" | "week";
  orgGetsSubInvites: boolean;
  orgInvitesTouchedRef: React.MutableRefObject<boolean>;
  removedSavedSubs: SubEventDraft[];
  removeSubEventDraft: (se: SubEventDraft) => void;
  requireSubEventSelection: boolean;
  setAllSubsAllDay: (v: boolean) => void;
  setAllSubsShowAsFree: (v: boolean) => void;
  setChildGender: React.Dispatch<React.SetStateAction<"" | "m" | "f" | "n">>;
  setChildTermPlural: React.Dispatch<React.SetStateAction<string>>;
  setChildTermSingular: React.Dispatch<React.SetStateAction<string>>;
  setCustomTermMode: React.Dispatch<React.SetStateAction<boolean>>;
  setEndDate: React.Dispatch<React.SetStateAction<string>>;
  setMainEventLabel: React.Dispatch<React.SetStateAction<string>>;
  setMainEventLabelMode: React.Dispatch<React.SetStateAction<"none" | "default" | "custom">>;
  setOrgGetsSubInvites: React.Dispatch<React.SetStateAction<boolean>>;
  setRemovedSavedSubs: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setRequireSubEventSelection: React.Dispatch<React.SetStateAction<boolean>>;
  setScope: (idx: number) => void;
  setStartDate: React.Dispatch<React.SetStateAction<string>>;
  setSubEventCalendar: React.Dispatch<React.SetStateAction<boolean>>;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setSubEventSingleChoice: React.Dispatch<React.SetStateAction<boolean>>;
  setSubEventsOnlyMode: React.Dispatch<React.SetStateAction<boolean>>;
  setSubEventsOptIn: React.Dispatch<React.SetStateAction<boolean>>;
  setSubImageCropIdx: React.Dispatch<React.SetStateAction<number>>;
  setTerminListOpen: React.Dispatch<React.SetStateAction<boolean>>;
  startDate: string;
  subEventCalendar: boolean;
  subEvents: SubEventDraft[];
  subEventSingleChoice: boolean;
  subEventsOnlyMode: boolean;
  subEventsOptIn: boolean;
  subImageCropIdx: number;
  t: (key: string) => string;
  terminListOpen: boolean;
  title: string;
  toggleDaySubEvent: (d: Date | null) => void;
}
export const SubEventsSection: React.FC<SubEventsSectionProps> = (p) => {
  const { visible } = p;
  const { activeScopeIdx, audience, berlinLocalToUtcIso, childGender, childTermPlural, childTermSingular, confirmDialog, currentStep, customTermMode, dayKeyOfSub, endDate, filterMode, goToScopeBar, isDe, isoToLocal, klammerDeadline, locationFilter, mainEventLabel, mainEventLabelMode, openRuleDays, openRuleEnabled, openRuleMode, orgGetsSubInvites, orgInvitesTouchedRef, removedSavedSubs, removeSubEventDraft, requireSubEventSelection, setAllSubsAllDay, setAllSubsShowAsFree, setChildGender, setChildTermPlural, setChildTermSingular, setCustomTermMode, setEndDate, setMainEventLabel, setMainEventLabelMode, setOrgGetsSubInvites, setRemovedSavedSubs, setRequireSubEventSelection, setScope, setStartDate, setSubEventCalendar, setSubEvents, setSubEventSingleChoice, setSubEventsOnlyMode, setSubEventsOptIn, setSubImageCropIdx, setTerminListOpen, startDate, subEventCalendar, subEvents, subEventSingleChoice, subEventsOnlyMode, subEventsOptIn, subImageCropIdx, t, terminListOpen, title, toggleDaySubEvent } = p;
  return (
              <div style={{ display: visible ? 'block' : 'none' }}>
              {/* v28.90: Die Frage stand ganz oben in Schritt 1, noch vor
                  dem Event-Titel — man musste sie beantworten, bevor klar
                  war, worum es überhaupt geht. Sie steht jetzt hinter den
                  Grundlagen (Titel, Datum, Beschreibung, Bild) und direkt
                  bei dem, was von ihr abhaengt: Erklaerung, Bezeichnung,
                  Anmelde-Modus und die Liste der Sub-Events. */}
              {/* v28.83: Die Opt-in-Frage steht jetzt in den GRUNDLAGEN statt in
                  Schritt 3. Ob ein Event ueberhaupt aus mehreren Teilen besteht,
                  ist eine Grundsatzfrage wie Titel und Datum — nicht etwas, das
                  man erst in einem spaeteren Schritt entdeckt. Abschalten mit
                  vorhandenen Sub-Events fragt nach (Soft-Disable, s. v28.2). */}
              {/* v22.36: Opt-in-Frage — Default nein; erst bei „ja" erscheint
                  die gesamte Sub-Event-Konfiguration. Abschalten mit
                  vorhandenen Sub-Events fragt nach und verwirft sie dann. */}
              {/* v28.91: Sichtbarer Schnitt zwischen den Grundlagen (1-5) und
                  dem Aufbau des Events. Ohne ihn schloss die Sub-Event-Frage
                  direkt an das Bild-Feld an und las sich wie noch ein
                  Grundlagen-Feld — dabei beginnt hier ein anderes Thema:
                  nicht mehr WAS das Event ist, sondern WORAUS es besteht. */}
              {/* v29.7: …und dieser Schnitt ist jetzt derselbe grüne Balken wie
                  der Schritt-Kopf. Als graue Haarlinie mit Kleinschrift war er
                  optisch schwächer als die Feld-Beschriftungen darüber und
                  darunter — das Auge nahm ihn als Beschriftung wahr, nicht als
                  Themenwechsel. */}
              <h3 className="dex-step-sub-head">
                {isDe ? 'Nutzung von Sub-Events' : 'Using sub-events'}
              </h3>
              {/* v29.8: Als Frage formuliert, die der Organizer beantworten
                  kann, statt als Beschreibung dessen, was er hier „festlegt".
                  Ein Beispiel dazu — abstrakte Begriffe wie „Aufbau" oder
                  „woraus es besteht" sagen niemandem, ob das den eigenen Fall
                  betrifft. Die Definition bleibt in der Erklär-Box unter dem
                  Schalter; hier steht nur die Frage. */}
              <p className="dex-step-sub-lead">
                {isDe
                  ? 'Besteht dein Event aus mehreren Teilen — zum Beispiel Workshops, einem Abend-Dinner oder mehreren Terminen?'
                  : 'Does your event consist of several parts — for example workshops, an evening dinner or several dates?'}
              </p>

              {/* v28.91: Schalter und Erklaerung gehoeren zusammen — die Frage
                  und die Antwort darauf, was ein Sub-Event ueberhaupt ist.
                  Zwei Kaesten uebereinander lasen sich wie zwei Themen. */}
              <div style={{ background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12, padding: '12px 16px', marginBottom: 12, border: '1px solid var(--dex-gray-200)' }}>
                {/* v30.4: Checkbox statt Toggle-Slider — der Schalter war das
                    einzige Slider-Element der Seite; alle anderen Optionen
                    sind Checkboxen, und zwei Bedienformen für dieselbe Art
                    Entscheidung lesen sich wie zwei verschiedene Dinge. */}
                <label style={{ display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer', margin: 0 }}>
                  <input
                    type="checkbox"
                    checked={subEventsOptIn}
                    style={{ width: 18, height: 18, flexShrink: 0, cursor: 'pointer' }}
                    onChange={e => {
                      const on = e.target.checked;
                      if (!on && subEvents.length > 0) {
                        (async () => {
                          // v28.2 SOFT-DISABLE: Der Toggle löscht NICHTS mehr.
                          // Die Sub-Events (inkl. Teilnehmerlisten und
                          // Anmeldungen) bleiben gespeichert und werden beim
                          // Speichern nur per _subEventsDisabled-Flag von der
                          // Anmeldeseite genommen. Wieder-Einschalten stellt
                          // alles unverändert wieder her — auch über ein
                          // Speichern hinweg. Endgültig löschen geht weiterhin
                          // über das X an der einzelnen Sub-Event-Karte.
                          const ok = await confirmDialog(
                            isDe
                              ? `Sub-Events deaktivieren? Deine ${subEvents.length} Sub-Event(s) bleiben mit allen Eingaben und Anmeldungen gespeichert, werden Teilnehmern nach dem Speichern aber nicht mehr angeboten. Beim Wieder-Einschalten ist alles unverändert da.`
                              : `Deactivate sub-events? Your ${subEvents.length} sub-event(s) remain stored with all input and registrations, but after saving they are no longer offered to attendees. Re-enabling restores everything unchanged.`,
                            { confirmLabel: isDe ? 'Deaktivieren' : 'Deactivate' }
                          );
                          if (ok) {
                            setSubEventsOptIn(false);
                            // Modi zurücksetzen, die ohne sichtbare Sub-Events
                            // keinen Sinn ergeben (sonst wäre das Event für
                            // Teilnehmer unbuchbar).
                            if (subEventsOnlyMode) setSubEventsOnlyMode(false);
                            if (requireSubEventSelection) setRequireSubEventSelection(false);
                          }
                        })().catch(() => { /* */ });
                        return;
                      }
                      setSubEventsOptIn(on);
                    }}
                  />
                  {/* v28.90: „Sub-Events nutzen?" mit der Antwort „— nein
                      (Standard)" dahinter las sich wie ein halb ausgefülltes
                      Formularfeld. Jetzt eine schlichte Handlung — was ein
                      Sub-Event ist, steht in der Erklär-Box direkt darunter,
                      die muss der Schalter nicht wiederholen. Ein Zusatz
                      erscheint nur, wenn es etwas zu sagen gibt: dass die
                      Sub-Events unten angelegt werden, bzw. dass deaktivierte
                      erhalten bleiben. */}
                  {/* v28.91: Die Sub-Event-Angaben sind Teil von Schritt 1 und
                      zaehlen deshalb in dessen Nummerierung weiter (1-5 oben). */}
                  <StepBadge n={6} />
                  <span style={{ fontSize: '0.9rem' }}>
                    <strong>{isDe ? 'Sub-Events aktivieren' : 'Enable sub-events'}</strong>
                    {subEventsOptIn
                      // v28.89: Die einzelnen Sub-Events werden weiter unten
                      // auf dieser Seite angelegt und über die Reiter oben
                      // bearbeitet — der Verweis auf „Schritt 3" stimmt seit
                      // v28.87 nicht mehr.
                      ? (isDe ? ' — die einzelnen Sub-Events legst du weiter unten an; bearbeitet werden sie über die Reiter oben.' : ' — you create the individual sub-events further down; edit them via the tabs above.')
                      : (subEvents.length > 0
                        ? (isDe
                          ? ` — deaktiviert. ${subEvents.length} Sub-Event(s) bleiben mit allen Anmeldungen gespeichert, sind für Teilnehmer aber unsichtbar. Einschalten stellt alles wieder her.`
                          : ` — deactivated. ${subEvents.length} sub-event(s) remain stored with all registrations but are hidden from attendees. Re-enable to restore.`)
                        : '')}
                  </span>
                </label>

                {/* v22.36: Erklärung, was ein Sub-Event ist (graue Beschreibungs-Box).
                    v28.96: Abstand zum Schalter darüber — die Box klebte direkt
                    daran und die Kachel wirkte gequetscht. */}
                <WizardHint
                  isDe={isDe}
                  variant="description"
                  title={isDe ? 'Was ist ein Sub-Event?' : 'What is a sub-event?'}
                  style={{ marginTop: 14, marginBottom: 0 }}
                >
                {isDe
                  ? <>Ein Sub-Event ist ein <strong>eigenständiger Programmbaustein</strong> innerhalb deines Events — z.&nbsp;B. ein Workshop, eine Session, ein Networking-Dinner oder eine Lauf-Distanz. Jedes Sub-Event hat <strong>eigene Plätze, einen eigenen Termin und eine eigene Teilnehmerliste</strong>, auf Wunsch auch eigene Abfrage-Felder; Teilnehmer wählen ihre Sub-Events direkt im Anmeldeformular. Typische Beispiele: eine Konferenz mit wählbaren Workshops oder ein Sommerfest mit optionalem Abendprogramm. Einfache Events (Meeting, Lunch, Feier) brauchen <strong>keine</strong> Sub-Events.</>
                  : <>A sub-event is a <strong>separate programme building block</strong> inside your event — e.g. a workshop, a session, a networking dinner or a run distance. Each sub-event has <strong>its own seats, its own schedule and its own attendee list</strong>, optionally its own custom fields; attendees pick their sub-events directly in the registration form. Typical examples: a conference with selectable workshops or a summer party with an optional evening programme. Simple events (meeting, lunch, celebration) do <strong>not</strong> need sub-events.</>}
              </WizardHint>
              </div>{/* Ende Kachel: Schalter + Erklaerung */}

              {/* v28.84: Bezeichnung und Anmelde-Modus gehoeren zur
                  Grundsatzfrage aus dem Schalter darüber — nicht in einen
                  spaeteren Schritt. Beide standen bisher in Schritt 3.
                  Sichtbar nur, wenn Sub-Events aktiv sind. */}
              {subEventsOptIn && (<>
              {/* Bezeichnungs-Dropdown */}
              <div style={{
                background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                padding: '14px 16px', marginBottom: 16,
                border: '1px solid var(--dex-gray-200)',
              }}>
                <label className="form-label" style={{ fontSize: '0.95rem', fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={7} />
                  {isDe ? 'Bezeichnung der Sub-Events' : 'Sub-event naming'}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> wie die untergeordneten Bausteine in der App genannt werden. Die <strong>Standardbezeichnung</strong> ist &bdquo;Sub-Event&ldquo;. Du kannst aber z.B. &bdquo;Workshop&ldquo;, &bdquo;Session&ldquo;, &bdquo;Programmpunkt&ldquo; oder &bdquo;Event-Section&ldquo; auswählen — oder eine eigene Bezeichnung (Singular + Plural) eintippen.<br /><br />
                      <strong>Anzeige in der App:</strong> der gewählte Begriff erscheint überall dort, wo bisher &bdquo;Sub-Event(s)&ldquo; stand — z.B. im Anmeldeformular der Teilnehmer (&bdquo;Verfügbare Workshops&ldquo;), in &bdquo;Meine Events&ldquo; und im Admin Center.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> how the child building blocks are named throughout the app. The <strong>default</strong> is &bdquo;Sub-event&ldquo;. You can pick e.g. &bdquo;Workshop&ldquo;, &bdquo;Session&ldquo;, &bdquo;Programmpunkt&ldquo; or &bdquo;Event section&ldquo; — or type your own singular + plural.<br /><br />
                      <strong>Shown in the app:</strong> the chosen term replaces &bdquo;Sub-event(s)&ldquo; everywhere — e.g. in the attendee registration form (&bdquo;Available workshops&ldquo;), in &bdquo;My events&ldquo; and in the admin center.
                    </>
                  )} />
                </label>
                {(() => {
                  const presets = [
                    { key: 'subevent',     singular: isDe ? 'Sub-Event' : 'Sub-event',         plural: isDe ? 'Sub-Events' : 'Sub-events' },
                    { key: 'workshop',     singular: 'Workshop',                                plural: 'Workshops' },
                    { key: 'session',      singular: 'Session',                                 plural: 'Sessions' },
                    { key: 'programmpunkt', singular: isDe ? 'Programmpunkt' : 'Program item', plural: isDe ? 'Programmpunkte' : 'Program items' },
                    { key: 'section',      singular: isDe ? 'Event-Section' : 'Event section', plural: isDe ? 'Event-Sections' : 'Event sections' },
                  ];
                  const matchKey = (() => {
                    // v15.9: customTermMode hat Priorität — wer in den
                    // Custom-Modus geklickt hat bleibt dort, auch wenn
                    // beide Inputs noch leer sind.
                    if (customTermMode) return 'custom';
                    const s = (childTermSingular || '').trim();
                    const p = (childTermPlural || '').trim();
                    if (!s && !p) return 'subevent';
                    const hit = presets.find(x => x.singular === s && x.plural === p);
                    return hit ? hit.key : 'custom';
                  })();
                  return (
                    <>
                      <select
                        className="form-input"
                        value={matchKey}
                        onChange={e => {
                          const k = e.target.value;
                          if (k === 'custom') {
                            // v15.9: Custom-Modus sticky machen, auch ohne
                            // initiale Werte — sonst kippt der Dropdown
                            // sofort zurück auf 'subevent'.
                            setCustomTermMode(true);
                            return;
                          }
                          // Preset gewählt → Custom-Modus aufheben + Werte
                          // aus dem Preset übernehmen.
                          setCustomTermMode(false);
                          const preset = presets.find(x => x.key === k);
                          if (preset) {
                            setChildTermSingular(preset.singular);
                            setChildTermPlural(preset.plural);
                          }
                        }}
                        style={{ marginTop: 6, maxWidth: 360 }}
                      >
                        {presets.map(p => (
                          <option key={p.key} value={p.key}>
                            {p.plural}
                          </option>
                        ))}
                        <option value="custom">{isDe ? 'Eigene Bezeichnung…' : 'Custom term…'}</option>
                      </select>
                      {matchKey === 'custom' && (
                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginTop: 10, maxWidth: 480 }}>
                          <div>
                            <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>{isDe ? 'Singular' : 'Singular'}</label>
                            <input
                              type="text"
                              className="form-input"
                              value={childTermSingular}
                              onChange={e => setChildTermSingular(e.target.value)}
                              placeholder={isDe ? 'z.B. Modul' : 'e.g. Module'}
                              style={{ padding: '6px 10px', fontSize: '0.9rem' }}
                            />
                          </div>
                          <div>
                            <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>{isDe ? 'Plural' : 'Plural'}</label>
                            <input
                              type="text"
                              className="form-input"
                              value={childTermPlural}
                              onChange={e => setChildTermPlural(e.target.value)}
                              placeholder={isDe ? 'z.B. Module' : 'e.g. Modules'}
                              style={{ padding: '6px 10px', fontSize: '0.9rem' }}
                            />
                          </div>
                        </div>
                      )}
                      {/* v29.60: Der unbestimmte Artikel wurde bisher aus dem
                          Wort geraten — und stand im Nominativ. Auf der
                          Anmeldeseite heisst es aber „wähle mindestens" …,
                          also Akkusativ: bei maskulinen Begriffen „einen",
                          nicht „ein". Das laesst sich nicht ableiten,
                          deshalb wird es hier gefragt. */}
                      {isDe && (
                        <div style={{ marginTop: 10, maxWidth: 480 }}>
                          <label style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', display: 'block', marginBottom: 3 }}>
                            Artikel (nur für die deutschen Texte)
                          </label>
                          <select
                            className="form-input"
                            value={childGender}
                            onChange={e => setChildGender(e.target.value as '' | 'm' | 'f' | 'n')}
                            style={{ padding: '6px 10px', fontSize: '0.9rem', maxWidth: 280 }}
                          >
                            <option value="">Automatisch erkennen</option>
                            <option value="m">der … (z.B. der Office-Tag)</option>
                            <option value="f">die … (z.B. die Session)</option>
                            <option value="n">das … (z.B. das Modul)</option>
                          </select>
                          <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', margin: '4px 0 0' }}>
                            {(() => {
                              const term = (childTermSingular || 'Sub-Event').trim() || 'Sub-Event';
                              const art = childGender === 'm' ? 'einen' : childGender === 'f' ? 'eine' : childGender === 'n' ? 'ein' : '';
                              return art
                                ? `Auf der Anmeldeseite steht dann: „Bitte wähle mindestens ${art} ${term} aus."`
                                : 'Ohne Auswahl rät DEX anhand der Endung — das geht bei ungewöhnlichen Bezeichnungen schief.';
                            })()}
                          </p>
                        </div>
                      )}
                    </>
                  );
                })()}

              {/* v28.91: Anmelde-Modus in derselben Kachel wie die
                  Bezeichnung — beides beschreibt, WIE die Sub-Events
                  auftreten. */}
              <div style={{ marginTop: 16, paddingTop: 14, borderTop: '1px solid var(--dex-gray-200)' }}>
                <label className="form-label" style={{ fontSize: '0.95rem', fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={8} />
                  {isDe ? 'Anmelde-Modus' : 'Registration mode'}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> ob sich Teilnehmer zusätzlich zum Hauptevent für Sub-Events anmelden können (Standard) oder ob es <strong>überhaupt kein Hauptevent-Anmelden</strong> mehr gibt und die Anmeldung ausschließlich über einzelne Sub-Events läuft.<br /><br />
                      <strong>Anzeige in der App:</strong> im Modus &bdquo;Nur Sub-Events&ldquo; ist die Anmelde-Checkbox für das Hauptevent im Teilnehmerformular ausgeblendet — der Teilnehmer muss zwingend mindestens einen Sub-Event auswählen. Im Schritt 6 (Kommunikation) wird der Haupt-Event-Tab ausgegraut, weil die Hauptevent-Kommunikation in diesem Modus nicht greift.<br /><br />
                      <strong>Empfehlung:</strong> nutze &bdquo;Nur Sub-Events&ldquo; für Mehrtages-Programme, in denen jeder Teilnehmer aus einem Pool von Slots wählt und ein Hauptevent-Slot keinen Sinn ergibt.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> whether attendees can additionally register for sub-events alongside the main event (default) or whether there is <strong>no main-event registration at all</strong> and registration runs exclusively via individual sub-events.<br /><br />
                      <strong>Shown in the app:</strong> in &bdquo;Sub-events only&ldquo; mode the main-event registration checkbox in the attendee form is hidden — the attendee must pick at least one sub-event. In step 7 (Communication) the main-event tab is greyed out, because main-event communication does not apply in this mode.<br /><br />
                      <strong>Tip:</strong> use &bdquo;Sub-events only&ldquo; for multi-day programmes where every attendee picks from a pool of slots and a main-event slot makes no sense.
                    </>
                  )} />
                </label>
                {/* v15.3.1: Custom-styled Radio-Cards in echtem Deloitte-Grün —
                    weil Edge mit accentColor uneinheitlich rendert (mal solid,
                    mal hollow), hier explizite Visual-Komposition: nativer
                    Input visually-hidden + eigener grüner Outer-Border-Kreis
                    mit grünem Inner-Dot wenn ausgewählt. Funktioniert 1:1 in
                    allen Browsern + matched die Optik der Registration-Page. */}
                {/* v28.73: Anker für den Quicklink aus dem Klammer-Info-Tooltip. */}
                <div id="dex-subevents-mode" style={{ display: 'flex', flexDirection: 'column', gap: 10, marginTop: 8 }}>
                  {[false, true].map(modeVal => {
                    const selected = !!subEventsOnlyMode === modeVal;
                    return (
                      <label
                        key={String(modeVal)}
                        style={{
                          display: 'flex', alignItems: 'flex-start', gap: 10,
                          padding: '10px 14px', borderRadius: 8,
                          border: `1px solid ${selected ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                          background: selected ? 'rgba(134,188,37,0.06)' : '#fff',
                          cursor: 'pointer',
                          transition: 'border-color 0.15s, background 0.15s',
                        }}
                      >
                        {/* Hidden native radio for accessibility + form-state */}
                        <input
                          type="radio"
                          name="subEventsMode"
                          checked={selected}
                          onChange={() => setSubEventsOnlyMode(modeVal)}
                          style={{
                            position: 'absolute', opacity: 0, pointerEvents: 'none',
                            width: 1, height: 1, margin: -1, padding: 0,
                            border: 0, overflow: 'hidden', clip: 'rect(0 0 0 0)',
                          }}
                        />
                        {/* Visual custom radio */}
                        <span
                          aria-hidden="true"
                          style={{
                            display: 'inline-block',
                            width: 18, height: 18, borderRadius: '50%',
                            border: `2px solid ${selected ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-400, #9aa0a6)'}`,
                            background: '#fff',
                            position: 'relative', flexShrink: 0,
                            marginTop: 2,
                            transition: 'border-color 0.15s',
                          }}
                        >
                          {selected && (
                            <span style={{
                              position: 'absolute', inset: 3,
                              borderRadius: '50%',
                              background: 'var(--dex-green, #86bc25)',
                            }} />
                          )}
                        </span>
                        <span style={{ fontSize: '0.88rem', flex: 1 }}>
                          {modeVal === false
                            ? (isDe
                                ? <>Anmeldung für <strong>Hauptevent + {(childTermPlural || 'Sub-Events').trim() || 'Sub-Events'}</strong> (Standard)</>
                                : <>Registration for <strong>main event + {(childTermPlural || 'sub-events').trim() || 'sub-events'}</strong> (default)</>)
                            : (isDe
                                ? <>Nur für <strong>{(childTermPlural || 'Sub-Events').trim() || 'Sub-Events'}</strong> (kein Hauptevent-Anmelden)</>
                                : <>Only for <strong>{(childTermPlural || 'sub-events').trim() || 'sub-events'}</strong> (no main-event registration)</>)}
                        </span>
                      </label>
                    );
                  })}
                </div>
              </div>

              </div>{/* Ende Kachel: Bezeichnung + Anmelde-Modus */}
              </>)}

              {subEventsOptIn && (<>
              {/* v28.97: Wie VIELE Sub-Events darf jemand waehlen? Das ist eine
                  eigene Frage neben „welche Ebenen sind buchbar" (Anmelde-Modus
                  darüber): Bei einer Workshop-Reihe soll man sich oft für
                  genau einen Termin entscheiden, bei einem Programm mit
                  Abendessen für beliebig viele. Aendert ausschliesslich die
                  Auswahl auf der Anmeldeseite — Teilnehmerlisten, Kapazitaeten
                  und Fristen der Sub-Events bleiben, wie sie sind. */}
              <div style={{
                background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                padding: '14px 16px', marginBottom: 16,
                border: '1px solid var(--dex-gray-200)',
              }}>
                <label className="form-label" style={{ fontSize: '0.95rem', fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8 }}>
                  {isDe ? 'Wie viele darf man auswählen?' : 'How many may be picked?'}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> ob ein Teilnehmer <strong>mehrere</strong> {(childTermPlural || 'Sub-Events')} gleichzeitig buchen darf oder sich für <strong>genau eines</strong> entscheiden muss.<br /><br />
                      <strong>Anzeige in der App:</strong> Bei &bdquo;mehrere&ldquo; sind es Kästchen zum Ankreuzen, bei &bdquo;genau eines&ldquo; verhält sich die Auswahl wie ein Radio-Knopf — ein neuer Klick ersetzt die bisherige Wahl. Im Kalender genauso.<br /><br />
                      <strong>Wichtig:</strong> An den {(childTermPlural || 'Sub-Events')} selbst ändert sich nichts — jedes behält seine eigene Teilnehmerliste, Kapazität und Frist. Es geht nur um die Auswahl.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> whether an attendee may book <strong>several</strong> sub-events at once or has to pick <strong>exactly one</strong>.<br /><br />
                      <strong>Shown in the app:</strong> &bdquo;several&ldquo; renders checkboxes, &bdquo;exactly one&ldquo; behaves like a radio button — a new click replaces the previous choice. Same in the calendar.<br /><br />
                      <strong>Note:</strong> nothing changes on the sub-events themselves — each keeps its own attendee list, capacity and deadline.
                    </>
                  )} />
                </label>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginTop: 8 }}>
                  {([
                    { val: false, label: isDe
                      ? <>Mehrere <strong>{(childTermPlural || 'Sub-Events')}</strong> gleichzeitig <span style={{ color: 'var(--dex-gray-500)' }}>(Standard)</span></>
                      : <>Several <strong>sub-events</strong> at once <span style={{ color: 'var(--dex-gray-500)' }}>(default)</span></> },
                    { val: true, label: isDe
                      ? <>Genau <strong>eines</strong> — die Auswahl ersetzt die vorherige</>
                      : <>Exactly <strong>one</strong> — a new pick replaces the previous</> },
                  ]).map(opt => {
                    const selected = subEventSingleChoice === opt.val;
                    return (
                      <label
                        key={String(opt.val)}
                        style={{
                          display: 'flex', alignItems: 'flex-start', gap: 10,
                          padding: '10px 14px', borderRadius: 8,
                          border: `1px solid ${selected ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                          background: selected ? 'rgba(134,188,37,0.06)' : '#fff',
                          cursor: 'pointer', transition: 'border-color 0.15s, background 0.15s',
                        }}
                      >
                        <input
                          type="radio"
                          name="subEventSingleChoice"
                          checked={selected}
                          onChange={() => setSubEventSingleChoice(opt.val)}
                          style={{
                            position: 'absolute', opacity: 0, pointerEvents: 'none',
                            width: 1, height: 1, margin: -1, padding: 0,
                            border: 0, overflow: 'hidden', clip: 'rect(0 0 0 0)',
                          }}
                        />
                        <span aria-hidden="true" style={{
                          display: 'inline-block', width: 18, height: 18, borderRadius: '50%',
                          border: `2px solid ${selected ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-400, #9aa0a6)'}`,
                          background: '#fff', position: 'relative', flexShrink: 0, marginTop: 2,
                        }}>
                          {selected && <span style={{ position: 'absolute', inset: 3, borderRadius: '50%', background: 'var(--dex-green, #86bc25)' }} />}
                        </span>
                        <span style={{ fontSize: '0.88rem', flex: 1 }}>{opt.label}</span>
                      </label>
                    );
                  })}
                </div>
                {/* Pflichttermine und „genau eines" widersprechen sich, sobald es
                    mehr als einen Pflichttermin gibt — dann kaeme niemand durch. */}
                {subEventSingleChoice && subEvents.filter(x => x.mandatory).length > 1 && (
                  <div style={{
                    margin: '12px 0 0', padding: '9px 11px', borderRadius: 8,
                    background: '#fff8e6', border: '1px solid #e0b34d', color: '#7a5a12',
                    fontSize: '0.78rem', lineHeight: 1.55,
                  }}>
                    {isDe
                      ? <>Es sind <strong>{subEvents.filter(x => x.mandatory).length} Pflichttermine</strong> markiert, obwohl nur eines gewählt werden darf — damit kann sich niemand anmelden. Nimm die Pflicht bei allen bis auf höchstens einem heraus (je Termin über die Reiter oben).</>
                      : <>There are <strong>{subEvents.filter(x => x.mandatory).length} mandatory dates</strong> although only one may be picked — nobody could register. Remove the mandatory flag from all but at most one (per date via the tabs above).</>}
                  </div>
                )}
              </div>

              {/* v24.58: Anzeige-Name des Haupt-Events in der Sub-Event-Auswahl.
                  Nur relevant, wenn das Hauptevent mit-buchbar ist (also NICHT
                  im „Nur Sub-Events"-Modus, wo es keine Hauptevent-Zeile gibt). */}
              {!subEventsOnlyMode && (
                <div style={{
                  background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                  padding: '14px 16px', marginBottom: 16,
                  border: '1px solid var(--dex-gray-200)',
                }}>
                  <label className="form-label" style={{ fontSize: '0.95rem', fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8 }}>
                    {isDe ? 'Bezeichnung des Haupt-Events in der Auswahl' : 'Main-event label in the selection'}
                    <InfoTooltip text={isDe ? (
                      <>
                        <strong>Was du hier einstellst:</strong> Wenn dein Event {(childTermPlural || 'Sub-Events')} hat, sieht der Teilnehmer auf der Anmeldeseite mehrere wählbare Bereiche. Standardmäßig steht über dem Hauptevent das Wort &bdquo;Haupt-Event&ldquo; — hier kannst du das umbenennen oder ganz weglassen.<br /><br />
                        <strong>Anzeige in der App:</strong> die gewählte Bezeichnung erscheint als kleine Überschrift über der Hauptevent-Auswahl, vor dem Event-Titel. Bei &bdquo;Kein Name&ldquo; wird nur der Event-Titel gezeigt.
                      </>
                    ) : (
                      <>
                        <strong>What you set here:</strong> When your event has {(childTermPlural || 'sub-events')}, attendees see several selectable areas on the registration page. By default the main event is prefixed with the word &bdquo;Main event&ldquo; — here you can rename it or drop it entirely.<br /><br />
                        <strong>Shown in the app:</strong> the chosen label appears as a small heading above the main-event option, before the event title. With &bdquo;No label&ldquo; only the event title is shown.
                      </>
                    )} />
                  </label>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginTop: 8 }}>
                    {([
                      { val: 'default' as const, label: isDe ? <>Standard: <strong>„Haupt-Event“</strong></> : <>Default: <strong>„Main event“</strong></> },
                      { val: 'custom' as const, label: isDe ? <>Eigener Name (z.B. <strong>„Konferenz“</strong>, <strong>„Hauptprogramm“</strong>)</> : <>Custom name (e.g. <strong>„Conference“</strong>, <strong>„Main programme“</strong>)</> },
                      { val: 'none' as const, label: isDe ? <>Kein Name (nur der Event-Titel)</> : <>No label (just the event title)</> },
                    ]).map(opt => {
                      const selected = mainEventLabelMode === opt.val;
                      return (
                        <label
                          key={opt.val}
                          style={{
                            display: 'flex', alignItems: 'flex-start', gap: 10,
                            padding: '10px 14px', borderRadius: 8,
                            border: `1px solid ${selected ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                            background: selected ? 'rgba(134,188,37,0.06)' : '#fff',
                            cursor: 'pointer', transition: 'border-color 0.15s, background 0.15s',
                          }}
                        >
                          <input
                            type="radio"
                            name="mainEventLabelMode"
                            checked={selected}
                            onChange={() => setMainEventLabelMode(opt.val)}
                            style={{
                              position: 'absolute', opacity: 0, pointerEvents: 'none',
                              width: 1, height: 1, margin: -1, padding: 0,
                              border: 0, overflow: 'hidden', clip: 'rect(0 0 0 0)',
                            }}
                          />
                          <span aria-hidden="true" style={{
                            display: 'inline-block', width: 18, height: 18, borderRadius: '50%',
                            border: `2px solid ${selected ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-400, #9aa0a6)'}`,
                            background: '#fff', position: 'relative', flexShrink: 0, marginTop: 2,
                            transition: 'border-color 0.15s',
                          }}>
                            {selected && (
                              <span style={{ position: 'absolute', inset: 3, borderRadius: '50%', background: 'var(--dex-green, #86bc25)' }} />
                            )}
                          </span>
                          <span style={{ fontSize: '0.88rem', flex: 1 }}>{opt.label}</span>
                        </label>
                      );
                    })}
                  </div>
                  {mainEventLabelMode === 'custom' && (
                    <input
                      type="text"
                      className="form-input"
                      value={mainEventLabel}
                      onChange={e => setMainEventLabel(e.target.value)}
                      placeholder={isDe ? 'z.B. Konferenz' : 'e.g. Conference'}
                      style={{ marginTop: 10, padding: '8px 12px', fontSize: '0.9rem' }}
                    />
                  )}
                  {/* Live-Vorschau, wie es der Teilnehmer sieht. */}
                  <div style={{ marginTop: 12, fontSize: '0.82rem', color: 'var(--dex-gray-500)' }}>
                    {isDe ? 'Vorschau:' : 'Preview:'}{' '}
                    <span style={{ fontWeight: 700, color: 'var(--dex-gray-700, #444)' }}>
                      {mainEventLabelMode === 'none'
                        ? (title || (isDe ? 'Event-Titel' : 'Event title'))
                        : `${mainEventLabelMode === 'custom' ? (mainEventLabel.trim() || (isDe ? 'Eigener Name' : 'Custom name')) : (isDe ? 'Haupt-Event' : 'Main event')}: ${title || (isDe ? 'Event-Titel' : 'Event title')}`}
                    </span>
                  </div>
                </div>
              )}

                {/* v28.91: Termine per Kalender. Bei einer Reihe über neun
                    Tage bedeutete das bisher: neunmal „Hinzufügen", neunmal
                    Titel tippen, neunmal zwei Datumsfelder. Hier klickt man
                    die Tage an. Erzeugt werden ganz normale Sub-Events — nur
                    schneller, und die Anmeldeseite zeigt sie dann als
                    Kalender statt als Liste aus neun Funkbuttons. */}
                <div style={{
                  background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                  padding: '14px 16px', marginBottom: 16,
                  border: '1px solid var(--dex-gray-200)',
                }}>
                  <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                    <input
                      type="checkbox"
                      checked={subEventCalendar}
                      onChange={e => setSubEventCalendar(e.target.checked)}
                      style={{ width: 18, height: 18, marginTop: 1, flexShrink: 0, cursor: 'pointer' }}
                    />
                    <span style={{ fontSize: '0.9rem' }}>
                      <StepBadge n={9} /> <strong>{isDe ? 'Die Sub-Events sind Termine (ein Tag je Sub-Event)' : 'The sub-events are dates (one day each)'}</strong>

                      <span style={{ display: 'block', color: 'var(--dex-gray-600)', marginTop: 2, fontWeight: 400 }}>
                        {isDe
                          ? 'Dann legst du sie unten im Kalender an — Tag anklicken, fertig — und Teilnehmer wählen ihre Tage auf der Anmeldeseite ebenfalls im Kalender statt aus einer langen Liste. Jeder Tag bleibt ein vollwertiges Sub-Event mit eigenen Plätzen, eigener Frist und eigenem Outlook-Termin.'
                          : 'You then create them in the calendar below — click a day, done — and attendees pick their days in a calendar instead of a long list. Each day stays a full sub-event with its own seats, deadline and Outlook entry.'}
                      </span>
                    </span>
                  </label>
                  {/* v29.56: Sammel-Schalter. Der Einzel-Haken je Termin bleibt
                      (Schritt 1, Reiter des Tages) — hier setzt man alle auf
                      einmal, weil niemand einundzwanzig Reiter durchklickt.
                      Dreizustand: an = alle, aus = keiner, gestrichelt =
                      gemischt (dann hat jemand einzelne abweichend gesetzt). */}
                  {subEventsOptIn && subEvents.length > 0 && (() => {
                    const nAll = subEvents.filter(se => se.allDay).length;
                    const nBusy = subEvents.filter(se => !se.showAsFree).length;
                    const n = subEvents.length;
                    const row = (
                      key: string, checked: boolean, mixed: boolean,
                      onChange: (v: boolean) => void, title: string, desc: React.ReactNode,
                    ): JSX.Element => (
                      <label key={key} style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', marginTop: 12, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
                        <input
                          type="checkbox"
                          checked={checked}
                          ref={el => { if (el) el.indeterminate = mixed; }}
                          onChange={e => onChange(e.target.checked)}
                          style={{ width: 18, height: 18, marginTop: 1, flexShrink: 0, cursor: 'pointer' }}
                        />
                        <span style={{ fontSize: '0.9rem' }}>
                          <strong>{title}</strong>
                          {mixed && (
                            <span style={{ marginLeft: 8, fontSize: '0.75rem', color: 'var(--dex-orange, #ed8b00)', fontWeight: 600 }}>
                              {isDe ? 'gemischt' : 'mixed'}
                            </span>
                          )}
                          <span style={{ display: 'block', color: 'var(--dex-gray-600)', marginTop: 2, fontWeight: 400 }}>{desc}</span>
                        </span>
                      </label>
                    );
                    return (
                      <>
                        {row('bulk-allday', nAll === n, nAll > 0 && nAll < n, setAllSubsAllDay,
                          isDe ? 'Alle Termine sind Ganztagestermine' : 'All dates are all-day entries',
                          isDe
                            ? <>Setzt den Ganztags-Status für <strong>alle {n} Termine</strong> auf einmal. Einzelne Tage kannst du danach über ihren Reiter abweichend einstellen.</>
                            : <>Sets the all-day flag for <strong>all {n} dates</strong> at once. You can still change individual days on their own tab afterwards.</>)}
                        {row('bulk-busy', nBusy === n, nBusy > 0 && nBusy < n, (v) => setAllSubsShowAsFree(!v),
                          isDe ? 'Alle Termine blockieren den Kalender' : 'All dates block the calendar',
                          isDe
                            ? <>Setzt <strong>alle {n} Termine</strong> auf Beschäftigt. Ohne Haken erscheinen sie als Frei — bei ganztägigen Terminen meist die bessere Wahl, sonst gilt jeder Tag als komplett belegt.</>
                            : <>Marks <strong>all {n} dates</strong> as busy. Without it they show as free — usually the better choice for all-day entries, otherwise every day counts as fully booked.</>)}
                      </>
                    );
                  })()}
                  {/* v29.55: Bekommen die Organizer die Outlook-Termine ALLER
                      Termine? Der Flow trägt sie als Teilnehmer ein (aus
                      OrganizerEmail, und die steht auf jeder Sub-Event-Zeile) —
                      bei einer Reihe über zwanzig Tage sind das zwanzig
                      Blocker für Tage ohne eigene Buchung. Genau die
                      Beschwerde aus der Rückmeldung. */}
                  {subEventsOptIn && subEvents.length > 0 && (
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', marginTop: 12, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
                      <input
                        type="checkbox"
                        checked={orgGetsSubInvites}
                        onChange={e => { orgInvitesTouchedRef.current = true; setOrgGetsSubInvites(e.target.checked); }}
                        style={{ width: 18, height: 18, marginTop: 1, flexShrink: 0, cursor: 'pointer' }}
                      />
                      <span style={{ fontSize: '0.9rem' }}>
                        <strong>{isDe ? 'Organizer bekommen alle Outlook-Termine' : 'Organizers receive every Outlook entry'}</strong>
                        <span style={{ display: 'block', color: 'var(--dex-gray-600)', marginTop: 2, fontWeight: 400 }}>
                          {isDe
                            ? <>Ohne Haken stehen die Termine nur in den Kalendern der <strong>angemeldeten Teilnehmer</strong>. Mit Haken landet <strong>jeder einzelne Termin</strong> auch bei euch als Organizer — bei {subEvents.length} {subEvents.length === 1 ? 'Termin' : 'Terminen'} also {subEvents.length} {subEvents.length === 1 ? 'Eintrag' : 'Einträge'}, auch für Tage, an denen ihr nicht dabei seid.</>
                            : <>Without it, the entries only appear in the calendars of <strong>registered attendees</strong>. With it, <strong>every single entry</strong> also lands with you as organizers — {subEvents.length} {subEvents.length === 1 ? 'entry' : 'entries'} for {subEvents.length} {subEvents.length === 1 ? 'date' : 'dates'}, including days you are not attending.</>}
                        </span>
                      </span>
                    </label>
                  )}
                  {subEventCalendar && (() => {
                    // v29.22: DREI Zustände, drei Farben. Vorher war alles
                    // ein Grün — und ein abgewählter Tag bekam obendrein die
                    // keyboard-selected-Färbung des DatePickers (dunkelgrün),
                    // sah also aus, als wäre er noch an.
                    const keyToDate = (k: string): Date => {
                      const [y, m, d] = k.split('-').map(n => parseInt(n, 10));
                      return new Date(y, m - 1, d);
                    };
                    const savedDays = subEvents.filter(se => se.dbId).map(se => dayKeyOfSub(se)).filter(Boolean).map(keyToDate);
                    const newDays = subEvents.filter(se => !se.dbId).map(se => dayKeyOfSub(se)).filter(Boolean).map(keyToDate);
                    const removedDays = removedSavedSubs.map(se => dayKeyOfSub(se)).filter(Boolean).map(keyToDate);
                    const marked = [...savedDays, ...newDays];
                    const openTo = startDate ? new Date(startDate) : undefined;
                    return (
                      <div style={{ marginTop: 12 }}>
                        {/* v29.22: Legende — drei Farben, drei Zustände. Der
                            Organizer muss VOR dem Speichern sehen, was das
                            Speichern tun wird. */}
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px 16px', margin: '0 0 8px', fontSize: '0.78rem', color: 'var(--dex-gray-700)' }}>
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                            <span style={{ width: 12, height: 12, borderRadius: '50%', background: 'var(--dex-green, #86bc25)', flexShrink: 0 }} />
                            {isDe ? 'gespeicherter Termin' : 'saved date'}
                          </span>
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                            <span style={{ width: 12, height: 12, borderRadius: '50%', background: 'var(--dex-blue, #0076a8)', flexShrink: 0 }} />
                            {isDe ? 'neu — wird beim Speichern angelegt' : 'new — created on save'}
                          </span>
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                            <span style={{ width: 12, height: 12, borderRadius: '50%', background: 'var(--dex-orange, #ed8b00)', flexShrink: 0 }} />
                            {isDe ? 'abgewählt — wird beim Speichern endgültig gelöscht (samt Teilnehmerliste); erneut anklicken stellt ihn wieder her' : 'deselected — permanently deleted on save (incl. attendee list); click again to restore'}
                          </span>
                        </div>
                        {/* v30.2: Erklärtext RECHTS neben dem Kalender statt
                            eingeklemmt darüber — rechts vom Raster war bisher
                            nur Leerraum. Auf schmalen Spalten bricht der Text
                            per flexWrap wieder unter den Kalender. */}
                        <div style={{ display: 'flex', alignItems: 'flex-start', gap: 16, flexWrap: 'wrap' }}>
                          <div style={{ flexShrink: 0 }}>
                            <DatePicker
                              inline
                              selected={null}
                              onChange={toggleDaySubEvent}
                              highlightDates={[
                                { 'dex-day-picked': savedDays },
                                { 'dex-day-new': newDays },
                                { 'dex-day-removed': removedDays },
                              ]}
                              openToDate={openTo}
                              locale="de"
                              // v29.17: dex-termin-calendar blendet die Nachbarmonats-
                              // Tage aus (CSS). In diesem Kalender ist ein Klick eine
                              // AKTION (Termin anlegen/entfernen), kein Datums-Pick —
                              // der Klick auf den „01.10." unten im September-Raster
                              // legte den Tag an UND sprang in den Oktober um. Wer
                              // einen Oktober-Tag will, blättert mit dem Pfeil.
                              calendarClassName="dex-datepicker-calendar dex-termin-calendar"
                            />
                          </div>
                          <p style={{ flex: '1 1 260px', minWidth: 240, fontSize: '0.8rem', color: 'var(--dex-gray-600)', margin: 0, lineHeight: 1.6 }}>
                            {isDe
                              ? <>Klick auf einen Tag legt ihn als Termin an, ein erneuter Klick nimmt ihn zurück. Titel und Zeiten werden gesetzt — jeder Tag übernimmt die <strong>Uhrzeit des Hauptevents</strong> (ohne Uhrzeit dort: ganztägig). Das ist nur die Vorbelegung: In der <strong>Liste unter dem Kalender</strong> öffnest du einen Termin mit &bdquo;Bearbeiten&ldquo; und änderst <strong>Start und Ende genau dort</strong> — ebenso Titel, Beschreibung und Bild; Plätze und Frist in Schritt 4.</>
                              : <>Clicking a day creates it as a date, clicking again removes it. Title and times are filled in — each day takes the <strong>main event&rsquo;s time of day</strong> (all-day if none is set there). That is only the starting point: in the <strong>list below the calendar</strong> open a date via &ldquo;Edit&rdquo; and change <strong>its start and end right there</strong> — as well as title, description and image; seats and deadline in step 4.</>}
                          </p>
                        </div>
                        {/* v29.75: Die Freischalt-Regel („Termine erst kurz
                            vorher zur Anmeldung freischalten") stand bis
                            v29.74 HIER im Kalender-Block. Sie ist aber eine
                            Anmelde-Regel, keine Termin-Anlage — sie wohnt
                            jetzt in Schritt 4 bei den Anmelde- und
                            Abmeldefristen der Klammer. Der Hinweis bleibt,
                            damit Organizer sie nicht an der alten Stelle
                            suchen. */}
                        <p style={{ margin: '10px 0 0', fontSize: '0.78rem', color: 'var(--dex-gray-500)', lineHeight: 1.5 }}>
                          {isDe
                            ? <>{openRuleEnabled
                              ? <>&bdquo;Anmeldung ab&ldquo; aktiv: je Termin <strong>{openRuleDays} {openRuleDays === 1 ? 'Tag' : 'Tage'} vor {openRuleMode === 'week' ? 'dem Montag der jeweiligen Woche' : 'dem jeweiligen Termin'}</strong>. </>
                              : <></>}&bdquo;Anmeldung ab&ldquo;, &bdquo;Anmeldung bis&ldquo; und &bdquo;Abmeldung bis&ldquo; für alle Termine stellst du in <strong>Schritt 4</strong> unter &bdquo;Anmelde- und Abmeldefristen&ldquo; ein.</>
                            : <>{openRuleEnabled
                              ? <>“Registration opens” active: <strong>{openRuleDays} {openRuleDays === 1 ? 'day' : 'days'} before {openRuleMode === 'week' ? 'the Monday of each week' : 'each date'}</strong>. </>
                              : <></>}You configure “registration opens/until” and “cancellation until” for all dates in <strong>step 4</strong> under “Registration & cancellation deadlines”.</>}
                        </p>
                        {/* v29.48: Termine außerhalb des Event-Zeitraums benennen.
                            Der Kalender lässt jeden Tag zu, der Zeitraum des
                            Hauptevents wandert aber nicht mit — in der Rückmeldung
                            standen deshalb der 28.09. und der 30.09. in der Liste,
                            während oben „Zeitraum bis 25.09." stand. Der Organizer
                            sieht beide Angaben nie nebeneinander; hier tun sie es. */}
                        {(() => {
                          const dayOfLocal = (v: string): string => (v || '').slice(0, 10);
                          const from = dayOfLocal(startDate);
                          const to = dayOfLocal(endDate) || from;
                          if (!from) return null;
                          const outside = subEvents
                            .map(se => dayKeyOfSub(se))
                            .filter(Boolean)
                            .filter(k => k < from || k > to)
                            .sort();
                          if (outside.length === 0) return null;
                          const fmt = (k: string): string => {
                            const [y, m, d] = k.split('-');
                            return `${d}.${m}.${y}`;
                          };
                          const newFrom = [from, ...outside].sort()[0];
                          const newTo = [to, ...outside].sort()[outside.length];
                          return (
                            <div style={{
                              marginTop: 10, padding: '10px 12px', borderRadius: 8,
                              background: 'rgba(237,139,0,0.10)', border: '1px solid var(--dex-orange, #ed8b00)',
                              fontSize: '0.82rem', color: 'var(--dex-gray-800)',
                            }}>
                              {isDe
                                ? <>{outside.length === 1 ? 'Ein Termin liegt' : `${outside.length} Termine liegen`} <strong>außerhalb des Event-Zeitraums</strong> ({fmt(from)} – {fmt(to)}): {outside.map(fmt).join(', ')}. Teilnehmer sehen sie trotzdem im Anmelde-Kalender. Entweder den Tag oben abwählen oder den Zeitraum erweitern.</>
                                : <>{outside.length === 1 ? 'One date is' : `${outside.length} dates are`} <strong>outside the event period</strong> ({fmt(from)} – {fmt(to)}): {outside.map(fmt).join(', ')}. Attendees still see them in the registration calendar. Either deselect the day above or extend the period.</>}
                              <button
                                type="button"
                                onClick={() => {
                                  // Nur den Tagesteil ersetzen — die Uhrzeiten des
                                  // Hauptevents bleiben, wie der Organizer sie gesetzt hat.
                                  setStartDate(`${newFrom}${(startDate || '').slice(10) || 'T00:00'}`);
                                  setEndDate(`${newTo}${(endDate || '').slice(10) || 'T23:59'}`);
                                }}
                                style={{
                                  display: 'block', marginTop: 8, background: 'transparent',
                                  border: 'none', padding: 0, cursor: 'pointer', fontFamily: 'inherit',
                                  color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700, fontSize: '0.82rem',
                                }}
                              >
                                {isDe
                                  ? `Zeitraum auf ${fmt(newFrom)} – ${fmt(newTo)} erweitern`
                                  : `Extend period to ${fmt(newFrom)} – ${fmt(newTo)}`}
                              </button>
                            </div>
                          );
                        })()}
                        <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap', margin: '10px 0 0' }}>
                          <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)', margin: 0 }}>
                            {marked.length === 0 && removedDays.length === 0
                              ? (isDe ? 'Noch kein Termin angelegt.' : 'No date created yet.')
                              : (isDe
                                ? `${marked.length} ${marked.length === 1 ? 'Termin' : 'Termine'} angelegt${removedDays.length > 0 ? ` · ${removedDays.length} zum Löschen vorgemerkt` : ''}.`
                                : `${marked.length} ${marked.length === 1 ? 'date' : 'dates'} created${removedDays.length > 0 ? ` · ${removedDays.length} marked for deletion` : ''}.`)}
                          </p>
                          {(marked.length > 0 || removedDays.length > 0) && (
                            <button
                              type="button"
                              onClick={() => setTerminListOpen(v => !v)}
                              style={{
                                display: 'inline-flex', alignItems: 'center', gap: 5,
                                background: 'transparent', border: 'none', cursor: 'pointer',
                                fontFamily: 'inherit', color: 'var(--dex-green-dark, #4a7c1f)',
                                fontWeight: 700, fontSize: '0.8rem', padding: 0,
                              }}
                            >
                              {terminListOpen
                                ? (isDe ? '▾ Liste einklappen' : '▾ Collapse list')
                                : (isDe ? '▸ Liste anzeigen' : '▸ Show list')}
                            </button>
                          )}
                        </div>
                        {/* v29.13: Die angelegten Termine stehen hier als Liste.
                            v28.96 hatte sie entfernt, weil sie „dieselben Tage
                            noch einmal" zeigt — das stimmt für das ANLEGEN, aber
                            der Kalender kann nur anlegen und wegnehmen. Zum
                            Bearbeiten musste man wissen, dass die Reiter GANZ
                            OBEN dafür da sind; ein Knopf „Sub-Events bearbeiten"
                            allein sagt nicht, welcher Tag welcher ist. Jetzt
                            führt jede Zeile direkt auf ihren Termin — Titel,
                            Zeiten, Beschreibung und Bild stehen dann oben in den
                            Feldern, Plätze und Frist in Schritt 4. */}
                        {terminListOpen && (() => {
                          const rows = subEvents
                            .map((se, idx) => ({ se, idx, key: dayKeyOfSub(se) }))
                            .filter(r => !!r.key)
                            .sort((a, b) => a.key.localeCompare(b.key));
                          const removedRows = removedSavedSubs
                            .map(se => ({ se, key: dayKeyOfSub(se) }))
                            .filter(r => !!r.key)
                            .sort((a, b) => a.key.localeCompare(b.key));
                          if (rows.length === 0 && removedRows.length === 0) return null;
                          const timeOfIso = (v: string): string => {
                            const local = isoToLocal(v || '');
                            const tt = (local || '').slice(11, 16);
                            return /^\d{2}:\d{2}$/.test(tt) ? tt : '';
                          };
                          return (
                            <div style={{ marginTop: 10, display: 'flex', flexDirection: 'column', gap: 6 }}>
                              {rows.map(({ se, idx }) => {
                                const active = activeScopeIdx === idx + 1;
                                const st = timeOfIso(se.startDate);
                                const en = timeOfIso(se.endDate);
                                // v29.52: Der Haken schlägt die Uhrzeiten — sonst
                                // steht in der Liste „00:00–23:59" bei einem
                                // Termin, der ganztägig gebucht wird.
                                const span = se.allDay
                                  ? (isDe ? 'ganztägig' : 'all day')
                                  : (st && en) ? `${st}–${en}` : (isDe ? 'ganztägig' : 'all day');
                                const cap = (se.maxParticipants || 0) > 0
                                  ? `${se.maxParticipants} ${isDe ? 'Plätze' : 'seats'}`
                                  : (isDe ? 'unbegrenzt' : 'unlimited');
                                return (
                                  <div
                                    key={se.id}
                                    style={{
                                      display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap',
                                      padding: '8px 10px', borderRadius: 8, background: '#fff',
                                      border: `1px solid ${active ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                                      borderLeft: `3px solid ${active ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                                    }}
                                  >
                                    <div style={{ flex: 1, minWidth: 160 }}>
                                      <div style={{ fontSize: '0.86rem', fontWeight: 700, color: 'var(--dex-gray-800, #333)' }}>
                                        {shortSubEventTitle(se.title, title) || (isDe ? 'Ohne Titel' : 'Untitled')}
                                      </div>
                                      <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                                        {span} · {cap}
                                        {se.registrationDeadline
                                          ? ` · ${isDe ? 'Frist' : 'deadline'} ${(isoToLocal(se.registrationDeadline) || '').slice(0, 10).split('-').reverse().join('.')}`
                                          : ''}
                                      </div>
                                    </div>
                                    <button
                                      type="button"
                                      className="btn btn-secondary"
                                      style={{ fontSize: '0.78rem', padding: '5px 12px' }}
                                      onClick={() => { setScope(idx + 1); goToScopeBar(); }}
                                      title={isDe
                                        ? 'Öffnet diesen Termin oben im Reiter — Titel, Zeiten, Beschreibung und Bild stehen dann in den Feldern darüber, Plätze und Frist in Schritt 4.'
                                        : 'Opens this date in the tab above — title, times, description and image then live in the fields above, seats and deadline in step 4.'}
                                    >
                                      {isDe ? 'Bearbeiten' : 'Edit'}
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() => removeSubEventDraft(se)}
                                      style={{
                                        background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)',
                                        padding: 4, lineHeight: 1,
                                      }}
                                      title={t('create.subevents.remove')}
                                    >
                                      <X size={16} />
                                    </button>
                                  </div>
                                );
                              })}
                              {/* v29.22: zum Löschen vorgemerkte Termine —
                                  orange, mit Rückholknopf. */}
                              {removedRows.map(({ se }) => (
                                <div
                                  key={se.id}
                                  style={{
                                    display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap',
                                    padding: '8px 10px', borderRadius: 8,
                                    background: 'rgba(237,139,0,0.07)',
                                    border: '1px solid var(--dex-orange, #ed8b00)',
                                    borderLeft: '3px solid var(--dex-orange, #ed8b00)',
                                  }}
                                >
                                  <div style={{ flex: 1, minWidth: 160 }}>
                                    <div style={{ fontSize: '0.86rem', fontWeight: 700, color: 'var(--dex-gray-800, #333)', textDecoration: 'line-through' }}>
                                      {shortSubEventTitle(se.title, title) || (isDe ? 'Ohne Titel' : 'Untitled')}
                                    </div>
                                    <div style={{ fontSize: '0.74rem', color: 'var(--dex-orange, #b35a00)', marginTop: 2, fontWeight: 600 }}>
                                      {isDe
                                        ? 'Wird beim Speichern endgültig gelöscht — samt Teilnehmerliste und Anmeldungen.'
                                        : 'Will be permanently deleted on save — including its attendee list and registrations.'}
                                    </div>
                                  </div>
                                  <button
                                    type="button"
                                    className="btn btn-secondary"
                                    style={{ fontSize: '0.78rem', padding: '5px 12px' }}
                                    onClick={() => {
                                      setRemovedSavedSubs(prev => prev.filter(x => x.id !== se.id));
                                      setSubEvents(prev => prev.some(x => x.id === se.id) ? prev : [...prev, se]);
                                    }}
                                  >
                                    {isDe ? 'Wiederherstellen' : 'Restore'}
                                  </button>
                                </div>
                              ))}
                            </div>
                          );
                        })()}
                      </div>
                    );
                  })()}
                </div>

                {/* ===== Sub-Events (z.B. Workshop-Tage, Networking-Dinner, Kick-off-Sessions) ===== */}
                {/* v28.96: Im Termin-Modus ist diese Liste redundant und
                    verwirrend — sie zeigt dieselben Tage noch einmal, die
                    direkt darüber im Kalender stehen. Entfernen geht dort per
                    Klick, Bearbeiten über die Reiter oben. */}
                {!subEventCalendar && (
                <div className="form-group" style={{ marginTop: 0 }}>
                  <label className="form-label" style={{ fontSize: '1rem', fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8 }}>
                    <StepBadge n={10} />
                    {/* v15.5: dynamische Bezeichnung — verwendet den oben
                        gewählten Plural-Term statt fix „Sub-Events".
                        v28.96: Im Termin-Modus ist die Liste NICHT optional und
                        auch nichts Zusätzliches — sie zeigt genau die Tage, die
                        oben im Kalender angeklickt wurden. „Sub-Events
                        (optional)" darüber liest sich, als käme hier noch eine
                        zweite, freiwillige Sorte. */}
                    {subEventCalendar
                      ? (isDe ? 'Angelegte Termine' : 'Created dates')
                      : <>{(childTermPlural || (isDe ? 'Sub-Events' : 'Sub-events'))} {isDe ? '(optional)' : '(optional)'}</>}
                    <InfoTooltip text={isDe ? (
                      <>
                        <strong>Was du hier einstellst:</strong> <strong>zusätzliche Sessions</strong> zum Hauptevent — z.B. eine Trainingsreihe, optionale Workshops, Side-Events am Vortag. Pro Session ein eigener Eintrag mit Titel, Ort, Start/Ende, Anmeldeschluss und Kapazität.<br /><br />
                        <strong>Anzeige in der App:</strong> Teilnehmer sehen die Sessions auf der Anmelde-Seite als <strong>eigene Anmelde-Bereiche</strong> — Haupt-Event und Sessions können <strong>unabhängig voneinander</strong> an- und abgewählt werden. Niemand muss zwingend das Haupt-Event mitbuchen, um sich für eine Session anzumelden.<br /><br />
                        <strong>Automatismen:</strong> pro Session-An-/Abmeldung gibt es eine <strong>eigene Bestätigungs-Mail</strong> und einen <strong>eigenen Outlook-Termin</strong> (im Deloitte-Layout). Pro Session optional ein- oder ausschaltbar (Mails / Outlook).<br /><br />
                        <strong>Empfehlung:</strong> nutze Sub-Events für mehrtägige Trainings, optionale Add-on-Workshops, oder regelmäßige Side-Sessions zu einem Haupt-Event. Bei einfachen Single-Day-Events nicht nötig.
                      </>
                    ) : (
                      <>
                        <strong>What you set here:</strong> <strong>additional sessions</strong> attached to the main event — e.g. a training series, optional workshops, side events on the day before. One entry per session with title, location, start/end, registration cutoff and capacity.<br /><br />
                        <strong>Shown in the app:</strong> attendees see the sessions on the registration page as <strong>independent registration blocks</strong> — main event and sessions can be picked or skipped <strong>independently</strong>. Nobody has to book the main event to sign up for a session.<br /><br />
                        <strong>Automation:</strong> each session registration / cancellation triggers its <strong>own confirmation mail</strong> and its <strong>own Outlook event</strong> (in the Deloitte template). Per session optionally toggleable (mails / Outlook).<br /><br />
                        <strong>Tip:</strong> use sub-events for multi-day trainings, optional add-on workshops, or recurring side sessions of a main event. Not needed for simple single-day events.
                      </>
                    )} />
                  </label>
                  {subEvents.length === 0 && (
                    <div style={{
                      padding: 10, border: '1px dashed var(--dex-gray-300)', borderRadius: 'var(--dex-radius)',
                      color: 'var(--dex-gray-500)', fontSize: '0.82rem', marginBottom: 8, marginTop: 4,
                      textAlign: 'center', background: 'var(--dex-gray-50, #fafafa)',
                    }}>
                      {t('create.subevents.empty')}
                    </div>
                  )}
                  {subEvents.map((se, idx) => {
                    // SubEvent-Daten werden intern als UTC-ISO gespeichert, für die
                    // react-datepicker-Komponenten brauchen wir Date-Objekte mit den
                    // richtigen Berlin-Lokalzeiten. Wir parsen via isoToLocal, was den
                    // Berlin-Wert als "YYYY-MM-DDTHH:MM" liefert, und bauen daraus ein
                    // JavaScript-Date-Objekt mit lokalen Werten.
                    const startDateObj = se.startDate ? (() => {
                      const local = isoToLocal(se.startDate); // "YYYY-MM-DDTHH:MM" in Berlin
                      if (!local) return null;
                      const [dp, tp] = local.split('T');
                      const [y, mo, da] = dp.split('-').map(n => parseInt(n, 10));
                      const [h, mi] = (tp || '00:00').split(':').map(n => parseInt(n, 10));
                      return new Date(y, mo - 1, da, h, mi, 0, 0);
                    })() : null;
                    const endDateObj = se.endDate ? (() => {
                      const local = isoToLocal(se.endDate);
                      if (!local) return null;
                      const [dp, tp] = local.split('T');
                      const [y, mo, da] = dp.split('-').map(n => parseInt(n, 10));
                      const [h, mi] = (tp || '00:00').split(':').map(n => parseInt(n, 10));
                      return new Date(y, mo - 1, da, h, mi, 0, 0);
                    })() : null;
                    // v15: deadlineObj entfernt — der Anmeldeschluss-Editor
                    // wandert nach Schritt 4 (Kapazität) in den Sub-Event-Tab.
                    return (
                      <div key={se.id} style={{
                        padding: '14px 16px', marginBottom: 10, marginTop: 4,
                        background: 'var(--dex-gray-50, #fafafa)', borderRadius: 'var(--dex-radius)',
                        border: '1px solid var(--dex-gray-200)', borderLeft: '3px solid var(--dex-green, #86bc25)',
                      }}>
                        {/* v28.89: Die Karte ist kein Editor mehr. Titel,
                            Zeiten, Beschreibung und Bild eines Sub-Events
                            werden oben in DENSELBEN Feldern gepflegt wie beim
                            Hauptevent — Reiter wählen, bearbeiten. Zwei Orte
                            für dieselbe Angabe waren der Grund, warum niemand
                            wusste, welcher gilt. Hier bleibt, was es sonst
                            nirgends gibt: die Liste selbst (anlegen,
                            umschalten, entfernen) und die Pflichtanmeldung. */}
                        <div style={{ display: 'flex', gap: 10, alignItems: 'center', marginBottom: 10, flexWrap: 'wrap' }}>
                          <div style={{ flex: 1, minWidth: 180 }}>
                            <div style={{ fontSize: '0.92rem', fontWeight: 700, color: 'var(--dex-gray-800, #333)' }}>
                              {shortSubEventTitle(se.title, title) || (isDe ? 'Ohne Titel' : 'Untitled')}
                            </div>
                            <div style={{ fontSize: '0.76rem', color: 'var(--dex-gray-500)', marginTop: 2 }}>
                              {(() => {
                                const fmt = (d: Date | null): string => (d
                                  ? d.toLocaleString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })
                                  : '');
                                const s = fmt(startDateObj);
                                const e = fmt(endDateObj);
                                if (!s && !e) return isDe ? 'Zeiten wie Hauptevent' : 'Times as main event';
                                return `${s || '—'} – ${e || '—'}`;
                              })()}
                            </div>
                          </div>
                          <button
                            type="button"
                            className="btn btn-secondary"
                            style={{ fontSize: '0.8rem', padding: '5px 14px' }}
                            onClick={() => setScope(idx + 1)}
                            title={isDe
                              ? 'Öffnet dieses Sub-Event oben im Reiter — Titel, Zeiten, Beschreibung und Bild stehen dann in den Feldern darüber.'
                              : 'Opens this sub-event in the tab above — title, times, description and image then live in the fields above.'}
                          >
                            {isDe ? 'Bearbeiten' : 'Edit'}
                          </button>
                          <button
                            type="button"
                            // v27.11: Entfernen bestätigen lassen — vorher
                            // löschte EIN Klick den Draft sofort; bei bereits
                            // gespeicherten Sub-Events wurden beim nächsten
                            // Speichern still Teilnehmerliste + Anmeldungen
                            // mitgelöscht. v29.13: gemeinsamer Handler.
                            onClick={() => removeSubEventDraft(se)}
                            style={{
                              background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)',
                              fontSize: '1.1rem', padding: '4px', lineHeight: 1,
                            }}
                            title={t('create.subevents.remove')}
                          >
                            <X size={16} />
                          </button>
                        </div>
                        {/* v28.89: Ende-vor-Start wird jetzt oben am Feld
                            gemeldet (dort wird korrigiert) — hier bleibt nur
                            der Hinweis, dass etwas nicht stimmt. */}
                        {startDateObj && endDateObj && endDateObj <= startDateObj && (
                          <p style={{ color: 'var(--dex-red, #c00)', fontSize: '0.78rem', margin: '-4px 0 8px' }}>
                            {isDe
                              ? 'Das Enddatum liegt vor dem Startdatum — über „Bearbeiten" korrigieren.'
                              : 'The end date is before the start date — fix it via „Edit".'}
                          </p>
                        )}

                        {/* v28.90: Die Pflichtanmeldung stand in JEDER Karte —
                            bei neun Sub-Events also neunmal derselbe
                            Erklaertext, der die Liste unlesbar machte. Sie
                            ist eine Einstellung DIESES Sub-Events und steht
                            deshalb jetzt bei seinen Grundlagen, auf seinem
                            eigenen Reiter — nicht in der Klammer-Liste. */}

                        {/* v15: Mail- und Outlook-Toggles raus aus der Sub-Event-
                            Card — sie leben jetzt ausschliesslich in Schritt 6
                            (Kommunikation) pro Sub-Event-Tab. Hier bleibt nur
                            das absolut Notwendige zur Anlage des Sub-Events. */}
                      </div>
                    );
                  })}
                  {/* v27.11: Zuschnitt-Modal für Sub-Event-Bilder — ein
                      gemeinsames Modal für alle Karten, Ziel via subImageCropIdx. */}
                  <button
                    type="button"
                    className="btn btn-outline"
                    style={{ fontSize: '0.85rem', padding: '6px 16px', marginTop: 4 }}
                    onClick={() => {
                      const newSub: SubEventDraft = {
                        id: (typeof crypto !== 'undefined' && crypto.randomUUID) ? crypto.randomUUID() : `se_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
                        title: '',
                        description: '',
                        location: '',
                        startDate: '',
                        endDate: '',
                        maxParticipants: 0,
                        disableEmails: false,
                        disableOutlook: false,
                        // v15.3: neue Sub-Events starten leer und vollwertig.
                        // v22.6: Sichtbarkeit (Standortfilter + Mailverteiler +
                        // Verknüpfung) wird standardmäßig vom Hauptevent/der
                        // Klammer vorbelegt — kann pro Sub-Event geändert werden.
                        // „Vom Hauptevent kopieren" übernimmt zusätzlich
                        // Kapazität/Deadlines/Warteliste.
                        locationAddress: { street: '', houseNo: '', zip: '', city: '' },
                        agenda: [],
                        transferTimes: [],
                        // v28.20: Hat die Klammer eine explizite Frist, starten
                        // neue Sub-Events mit demselben Anmeldeschluss.
                        ...(subEventsOnlyMode && klammerDeadline
                          ? { registrationDeadline: berlinLocalToUtcIso(klammerDeadline) || '' }
                          : {}),
                        lastDeregisterDate: '',
                        locationFilter: locationFilter,
                        audience: audience,
                        filterMode: filterMode,
                        excludedUsers: [],
                        waitlistEnabled: true,
                        askSalutation: false,
                      };
                      setSubEvents([...subEvents, newSub]);
                    }}
                  >
                    <Plus size={14} /> {t('create.subevents.add')}
                  </button>
                </div>
                )}
              </>)}{/* close subEventsOptIn */}
              {/* v28.96: Das Zuschnitt-Modal für Sub-Event-Bilder lag IM
                  Listen-Block. Der wird auf einem Sub-Event-Reiter gar nicht
                  gerendert (activeScopeIdx === 0) — „Bild editieren" dort
                  setzte also den Index, ohne dass je ein Dialog aufging.
                  Jetzt hängt es am Wrapper und ist von beiden Wegen aus da. */}
                {subImageCropIdx !== null && subEvents[subImageCropIdx] && (
                  <ImageCropModal
                    open
                    src={subEvents[subImageCropIdx].imagePreview || ''}
                    isDe={isDe}
                    onClose={() => setSubImageCropIdx(null)}
                    onApply={(dataUrl, file) => {
                      const i = subImageCropIdx;
                      setSubEvents(prev => prev.map((x, xi) => xi === i ? { ...x, imagePreview: dataUrl, imageFile: file, imageRemoved: false } : x));
                      setSubImageCropIdx(null);
                    }}
                  />
                )}
              </div>
  );
};
