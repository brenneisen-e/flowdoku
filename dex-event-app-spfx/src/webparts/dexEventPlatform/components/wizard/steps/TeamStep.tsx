/**
 * v30.13 — Modularisierung Stufe 3: Schritt „Team-Anmeldung" (Step 7,
 * Index 6) als eigene Komponente. JSX 1:1 aus EventCreationPage; der
 * Schritt ist reine Konfiguration und hängt an neun State-Paaren des
 * Wizards (Toggle, Größe, Team-Name, eigener Begriff, Erstell-Sperre,
 * Beitritts-Modus, Slot-Sichtbarkeit, Approval). `visible` ersetzt
 * `currentStep === 6` — display:none statt unmount, damit Eingaben beim
 * Schrittwechsel erhalten bleiben.
 */
import * as React from 'react';
import { InfoTooltip } from '../../InfoTooltip';
import { useLanguage } from '../../../context/LanguageContext';
// v31.2: Gemeinsame UI-Klassen (Schalter, Kacheln, Aufklapper) statt
// Inline-Styles — nur so gibt es Hover, siehe docs/ui-leitfaden.md.
import { cx } from '../../dexUi';
import { Check, ChevronDown, Settings, Users } from '../../Icons';

export interface TeamStepProps {
  visible: boolean;
  teamRegistrationEnabled: boolean;
  setTeamRegistrationEnabled: (v: boolean) => void;
  teamSize: number;
  setTeamSize: (v: number) => void;
  askTeamName: boolean;
  setAskTeamName: (v: boolean) => void;
  teamTermSingular: string;
  setTeamTermSingular: (v: string) => void;
  teamTermPlural: string;
  setTeamTermPlural: (v: string) => void;
  teamMembersCannotCreate: boolean;
  setTeamMembersCannotCreate: (v: boolean) => void;
  teamPartialAllowed: boolean;
  setTeamPartialAllowed: (v: boolean) => void;
  teamOpenSlotsVisible: boolean;
  setTeamOpenSlotsVisible: (v: boolean) => void;
  teamJoinRequiresApproval: boolean;
  setTeamJoinRequiresApproval: (v: boolean) => void;
}

export const TeamStep: React.FC<TeamStepProps> = ({
  visible,
  teamRegistrationEnabled, setTeamRegistrationEnabled,
  teamSize, setTeamSize,
  askTeamName, setAskTeamName,
  teamTermSingular, setTeamTermSingular,
  teamTermPlural, setTeamTermPlural,
  teamMembersCannotCreate, setTeamMembersCannotCreate,
  teamPartialAllowed, setTeamPartialAllowed,
  teamOpenSlotsVisible, setTeamOpenSlotsVisible,
  teamJoinRequiresApproval, setTeamJoinRequiresApproval,
}) => {
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  // v31.2: Aufklapper „Weitere Einstellungen" (eigene Bezeichnung). Standard
  // zu — der Zähler daneben zeigt, ob dort etwas gesetzt ist.
  const [moreOpen, setMoreOpen] = React.useState(false);
  const hasCustomTerm = !!((teamTermSingular || '').trim() || (teamTermPlural || '').trim());
  return (
    <div style={{ display: visible ? 'block' : 'none' }}>
    <h2 className="dex-step-head-title">
      <span className="dex-step-eyebrow">{isDe ? 'Schritt 7 von 9' : 'Step 7 of 9'}</span>
      {isDe ? 'Team-Anmeldung' : 'Team registration'}
    </h2>
    <p className="dex-step-head-lead">
      {isDe
        ? <><strong>Optional.</strong> Eine Person meldet ein ganzes Team auf einmal an — praktisch für Lauf-Teams, Workshop-Gruppen oder Tische bei einer Abendveranstaltung. Standard: aus.</>
        : <><strong>Optional.</strong> One person registers an entire team in one go — handy for running teams, workshop groups or tables at an evening event. Default: off.</>}
    </p>

    {/* Haupt-Ein/Aus des Schritts.
        v31.2: Schalter statt Checkbox (Leitfaden 2b: Ein/Aus für einen
        ganzen Bereich); die grüne Kante zeigt an, dass darunter etwas gilt.
        Der Satz darunter nennt die Folge für BEIDE Zustände. */}
    <div className="dex-ui-section">
      <div className={cx('dex-ui-card', teamRegistrationEnabled && 'dex-ui-card--accent')}>
        <div className="dex-ui-inline">
          <label className="dex-ui-switch">
            <input
              type="checkbox"
              checked={teamRegistrationEnabled}
              onChange={e => setTeamRegistrationEnabled(e.target.checked)}
            />
            <span className="dex-ui-switch-track" />
            <span className="dex-ui-switch-label">{isDe ? 'Team-Anmeldung erlauben' : 'Allow team registration'}</span>
          </label>
          <InfoTooltip text={isDe
            ? <>
                <strong>Was du hier einstellst:</strong> ob eine Person ein <strong>ganzes Team</strong> über das Anmeldeformular anmelden darf — statt sich nur selbst einzutragen.<br /><br />
                <strong>Anzeige in der App:</strong> der Team-Lead sieht nach Eingabe seiner eigenen Daten ein zusätzliches Formularfeld pro weiterem Team-Mitglied (Name, E-Mail). Default: aus — dann verhält sich das Event wie gewohnt (eine Person meldet sich selbst an).<br /><br />
                <strong>Auswirkung für Teilnehmer:</strong> die mit angemeldeten Personen bekommen automatisch eine eigene Bestätigungsmail und (sofern Outlook aktiv ist) eigene Kalender-Einladung — sie müssen sich nicht selber registrieren.
              </>
            : <>
                <strong>What this controls:</strong> whether one person can register an <strong>entire team</strong> via the registration form — instead of only registering themselves.<br /><br />
                <strong>Where you see it:</strong> after entering their own details, the team lead sees an additional form block per team member (name, email). Default: off — the event behaves as usual (one person registers themselves).<br /><br />
                <strong>For attendees:</strong> co-registered members automatically receive their own confirmation email and (if Outlook is enabled) their own calendar invite — they do not have to register themselves.
              </>
          } />
        </div>
        <div className="dex-ui-help" style={{ marginTop: 8 }}>
          {isDe
            ? 'An: Der Team-Lead trägt im Anmeldeformular je Mitglied Name und E-Mail ein; jedes Mitglied bekommt Bestätigungsmail und Outlook-Termin automatisch. Aus: Jede Person meldet sich selbst an — wie gewohnt.'
            : 'On: The team lead enters name and email per member in the registration form; every member automatically gets a confirmation mail and Outlook invite. Off: Everyone registers themselves — as usual.'}
        </div>
      </div>
    </div>

    {/* Team-Größe + Team-Name-Frage — gedämpft, wenn Team-Anmeldung aus.
        v31.2: Die Beschriftung ist die Frage, die Grenzen (2–20, Vorgabe 4)
        stehen sichtbar im Hilfetext statt nur im Tooltip. */}
    <div className="dex-ui-section">
      <div className="dex-ui-section-title">{isDe ? 'Größe und Name' : 'Size and name'}</div>
      <div className={cx('dex-ui-card', !teamRegistrationEnabled && 'dex-ui-card--muted')}>
        <div className="dex-ui-field">
          <label className="dex-ui-label" htmlFor="dex-wizard-team-size">
            {isDe ? 'Wie viele Personen passen in ein Team?' : 'How many people fit in a team?'}
            <InfoTooltip text={isDe
            ? <>
                <strong>Was du hier einstellst:</strong> die maximale Anzahl Personen pro Team (inkl. Team-Lead). Min. 2, Max. 20. Default 4.<br /><br />
                <strong>Anzeige in der App:</strong> der Team-Lead sieht so viele Mitglied-Slots wie hier gesetzt; einzelne Slots können leer bleiben, ein Team ist also nicht zwingend voll.<br /><br />
                <strong>Auswirkung für Teilnehmer:</strong> ein Team kann maximal so groß werden — versucht der Team-Lead, mehr Mitglieder einzutragen, wird er gestoppt.
              </>
            : <>
                <strong>What this controls:</strong> the maximum number of people per team (including the team lead). Min. 2, max. 20. Default 4.<br /><br />
                <strong>Where you see it:</strong> the team lead sees as many member slots as configured here; slots can stay empty, so teams are not required to be full.<br /><br />
                <strong>For attendees:</strong> a team caps at this size — attempting to add more members is blocked.
              </>
            } />
          </label>
          <div className="dex-ui-inline">
            <input
              id="dex-wizard-team-size"
              type="number"
              className="form-input"
              min={2}
              max={20}
              value={teamSize}
              disabled={!teamRegistrationEnabled}
              onChange={e => {
                const v = parseInt(e.target.value, 10);
                if (isNaN(v)) { setTeamSize(2); return; }
                setTeamSize(Math.max(2, Math.min(20, v)));
              }}
              style={{ maxWidth: 120 }}
            />
            <span className="dex-ui-muted">{isDe ? 'Personen, Team-Lead eingeschlossen' : 'people, including the team lead'}</span>
          </div>
          <div className="dex-ui-help">
            {isDe
              ? 'Mindestens 2, höchstens 20 (Vorgabe 4). Plätze dürfen leer bleiben — ein Team muss nicht voll sein; mehr Mitglieder als hier gesetzt lässt das Formular nicht zu.'
              : 'At least 2, at most 20 (default 4). Seats may stay empty — a team does not have to be full; the form does not allow more members than set here.'}
          </div>
        </div>
        <label className={cx('dex-ui-toggle-row', askTeamName && 'is-active', !teamRegistrationEnabled && 'is-disabled')}>
          <input
            type="checkbox"
            checked={askTeamName}
            disabled={!teamRegistrationEnabled}
            onChange={e => setAskTeamName(e.target.checked)}
          />
          <span className="dex-ui-toggle-row-body">
            <span className="dex-ui-toggle-row-title">
              {isDe ? 'Der Team-Lead gibt dem Team einen Namen' : 'The team lead gives the team a name'}
              {/* v31.2: Das Tooltip-Symbol liegt im <label> — ohne preventDefault
                  schaltet ein Klick auf das „i" die Checkbox mit. */}
              <span onClick={e => e.preventDefault()}>
              <InfoTooltip text={isDe
            ? <>
                <strong>Was du hier einstellst:</strong> ob der Team-Lead beim Anmelden zusätzlich einen <strong>frei wählbaren Team-Namen</strong> eingeben muss (z.B. &bdquo;Die schnellen Sieben&ldquo;).<br /><br />
                <strong>Anzeige in der App:</strong> der Team-Name erscheint auf der Seite &bdquo;Meine Events&ldquo; beim Team-Lead und allen Mitgliedern. Bei offenen Slots (Team noch nicht voll) wird der Team-Name in der Slot-Liste angezeigt, damit andere Teilnehmer bei Interesse beitreten können.<br /><br />
                <strong>Auswirkung für Teilnehmer:</strong> macht das Team identifizierbar. Bleibt die Option aus, wird das Team nur intern über den Namen des Team-Leads referenziert.
              </>
            : <>
                <strong>What this controls:</strong> whether the team lead has to enter a <strong>freely chosen team name</strong> during registration (e.g. &ldquo;The Fast Seven&rdquo;).<br /><br />
                <strong>Where you see it:</strong> the team name appears on &ldquo;My Events&rdquo; for the team lead and all members. For open slots (team not full yet), the name is displayed in the slot list so other attendees can join.<br /><br />
                <strong>For attendees:</strong> makes the team identifiable. If turned off, teams are referenced internally only via the team lead&apos;s name.
              </>
              } />
              </span>
            </span>
            <span className="dex-ui-toggle-row-desc">
              {isDe
                ? <>Dann fragt das Formular einen frei wählbaren Namen ab (z.B. &bdquo;Die schnellen Sieben&ldquo;); er steht bei allen Mitgliedern unter &bdquo;Meine Events&ldquo; und in der Liste offener Teams. Aus: das Team heißt intern nach dem Team-Lead.</>
                : <>Then the form asks for a freely chosen name (e.g. &bdquo;The Fast Seven&ldquo;); it appears for all members under &bdquo;My Events&ldquo; and in the list of open teams. Off: the team is referenced internally by its lead.</>}
            </span>
          </span>
        </label>
      </div>
    </div>

    {/* v22.78: „Teilnehmer dürfen keine neuen Teams erstellen".
        v31.2: Als Frage mit zwei Kacheln statt einer verneinten Checkbox —
        „dürfen keine … erstellen" musste man zweimal lesen. Dieselbe
        boolesche Bindung: rechte Kachel = true. Steht VOR dem Beitritts-
        Block, weil der davon abhängt. */}
    <div className="dex-ui-section">
      <div className="dex-ui-section-title">{isDe ? 'Wer stellt die Teams zusammen?' : 'Who puts the teams together?'}</div>
      <div className={cx('dex-ui-card', !teamRegistrationEnabled && 'dex-ui-card--muted')}>
        <div className="dex-ui-grid-2" role="radiogroup" aria-label={isDe ? 'Wer stellt die Teams zusammen?' : 'Who puts the teams together?'}>
          <button type="button" role="radio" aria-checked={!teamMembersCannotCreate} disabled={!teamRegistrationEnabled}
            className={cx('dex-ui-choice', !teamMembersCannotCreate && 'is-active')} onClick={() => setTeamMembersCannotCreate(false)}>
            <span className="dex-ui-choice-icon"><Users size={18} /></span>
            <span className="dex-ui-choice-body">
              <span className="dex-ui-choice-title">{isDe ? 'Die Teilnehmer selbst' : 'The participants themselves'}</span>
              <span className="dex-ui-choice-desc">
                {isDe
                  ? 'Der Team-Lead meldet sein Team im Formular an. Wer noch kein Team hat, kann später einem offenen Team beitreten.'
                  : 'The team lead registers their team in the form. Anyone without a team can join an open team later.'}
              </span>
            </span>
            <span className="dex-ui-choice-check"><Check size={12} /></span>
          </button>
          <button type="button" role="radio" aria-checked={teamMembersCannotCreate} disabled={!teamRegistrationEnabled}
            className={cx('dex-ui-choice', teamMembersCannotCreate && 'is-active')} onClick={() => setTeamMembersCannotCreate(true)}>
            <span className="dex-ui-choice-icon"><Settings size={18} /></span>
            <span className="dex-ui-choice-body">
              <span className="dex-ui-choice-title">{isDe ? 'Du als Organizer' : 'You as the organizer'}</span>
              <span className="dex-ui-choice-desc">
                {isDe
                  ? 'Alle melden sich einzeln an; du verteilst sie im Organizer Center per Drag & Drop auf die Teams. Empfohlen für Break-Out-Sessions.'
                  : 'Everyone registers individually; you assign them to teams in the Organizer Center via drag & drop. Recommended for break-out sessions.'}
              </span>
            </span>
            <span className="dex-ui-choice-check"><Check size={12} /></span>
          </button>
        </div>
      </div>
    </div>

    {/* v11.81: Beitritts-Modus — Modus + Sichtbarkeit + Approval.
        v31.2: Als Abschnitt mit Frage-Überschrift; die beiden Radios sind
        Auswahl-Kacheln (zwei Alternativen mit Folge), die Checkboxen
        Schalter-Zeilen. v22.78: Beitritt ist irrelevant, wenn der Organizer
        selbst zuordnet — dann gedämpft, aber weiter bedienbar (wie zuvor),
        und ein Satz sagt WARUM statt nur auszugrauen. */}
    <div className="dex-ui-section">
      <div className="dex-ui-section-title">{isDe ? 'Beitritt zu offenen Teams' : 'Joining open teams'}</div>
      <div className={cx('dex-ui-card', (!teamRegistrationEnabled || teamMembersCannotCreate) && 'dex-ui-card--muted')}>
        {teamMembersCannotCreate && (
          <div className="dex-ui-callout dex-ui-callout--neutral" style={{ marginBottom: 14 }}>
            {isDe
              ? 'Du ordnest die Teams selbst zu — niemand tritt einem Team bei. Die Einstellungen hier greifen dann nicht.'
              : 'You assign the teams yourself — nobody joins a team. These settings then have no effect.'}
          </div>
        )}
        <div className="dex-ui-field">
          <div className="dex-ui-label">{isDe ? 'Muss ein Team beim Anmelden vollständig sein?' : 'Does a team have to be complete at registration?'}</div>
          <div className="dex-ui-grid-2" role="radiogroup" aria-label={isDe ? 'Vollständige oder Teil-Teams' : 'Complete or partial teams'}>
            <button type="button" role="radio" aria-checked={!teamPartialAllowed} disabled={!teamRegistrationEnabled}
              className={cx('dex-ui-choice', !teamPartialAllowed && 'is-active')} onClick={() => setTeamPartialAllowed(false)}>
              <span className="dex-ui-choice-body">
                <span className="dex-ui-choice-title">{isDe ? 'Ja, nur komplette Teams' : 'Yes, only complete teams'}</span>
                <span className="dex-ui-choice-desc">
                  {isDe
                    ? `Der Team-Lead trägt alle ${teamSize} Mitglieder beim Anmelden ein. Halbe Teams gibt es nicht.`
                    : `The team lead enters all ${teamSize} members at registration. Partial teams are not possible.`}
                </span>
              </span>
              <span className="dex-ui-choice-check"><Check size={12} /></span>
            </button>
            <button type="button" role="radio" aria-checked={teamPartialAllowed} disabled={!teamRegistrationEnabled}
              className={cx('dex-ui-choice', teamPartialAllowed && 'is-active')} onClick={() => setTeamPartialAllowed(true)}>
              <span className="dex-ui-choice-body">
                <span className="dex-ui-choice-title">{isDe ? 'Nein, Teil-Teams sind erlaubt' : 'No, partial teams are allowed'}</span>
                <span className="dex-ui-choice-desc">
                  {isDe
                    ? `Der Team-Lead meldet z.B. 2 von ${teamSize} an; die übrigen Plätze bleiben offen und andere können später beitreten (Schalter unten).`
                    : `The team lead registers e.g. 2 of ${teamSize}; the remaining seats stay open and others can join later (switches below).`}
                </span>
              </span>
              <span className="dex-ui-choice-check"><Check size={12} /></span>
            </button>
          </div>
        </div>

        <div className="dex-ui-stack">
          {/* Sichtbarkeit offener Slots */}
          <label className={cx('dex-ui-toggle-row', teamOpenSlotsVisible && 'is-active', !teamRegistrationEnabled && 'is-disabled')}>
            <input
              type="checkbox"
              checked={teamOpenSlotsVisible}
              disabled={!teamRegistrationEnabled}
              onChange={e => {
                const v = e.target.checked;
                setTeamOpenSlotsVisible(v);
                if (!v) setTeamJoinRequiresApproval(false);
              }}
            />
            <span className="dex-ui-toggle-row-body">
              <span className="dex-ui-toggle-row-title">
                {isDe ? 'Offene Teams stehen auf der Anmeldeseite' : 'Open teams are listed on the registration page'}
                <span onClick={e => e.preventDefault()}>
                <InfoTooltip text={isDe
            ? <>
                <strong>Was du hier einstellst:</strong> ob andere Teilnehmer Teams mit offenen Slots in der Anmeldeseite sehen und beitreten können.<br /><br />
                <strong>Anzeige in der App:</strong> auf der Anmeldeseite erscheint eine Liste &bdquo;Teams mit freien Plätzen&ldquo; — pro Team mit der Anzahl freier Slots und (falls aktiviert) dem Team-Namen, aber <strong>ohne</strong> die Namen der bereits angemeldeten Mitglieder (Privatsphäre).<br /><br />
                <strong>Auswirkung für Teilnehmer:</strong> wer noch in keinem Team ist, kann mit einem Klick einem offenen Slot beitreten — entweder sofort gültig oder erst nach Bestätigung durch den Team-Lead (siehe nächste Option).
              </>
            : <>
                <strong>What this controls:</strong> whether other attendees see and can join teams with open slots on the registration page.<br /><br />
                <strong>Where you see it:</strong> the registration page shows a list &ldquo;teams with free seats&rdquo; — per team with the count of free slots and (if enabled) the team name, but <strong>without</strong> the names of already-registered members (privacy).<br /><br />
                <strong>For attendees:</strong> anyone not yet in a team can join an open slot with one click — either immediately or only after lead approval (see next option).
              </>
                } />
                </span>
              </span>
              <span className="dex-ui-toggle-row-desc">
                {isDe
                  ? <>Dann sehen andere &bdquo;Team mit X freien Plätzen&ldquo; und treten mit einem Klick bei — <strong>ohne</strong> die Namen der schon angemeldeten Mitglieder (Privatsphäre).</>
                  : <>Then others see &bdquo;team with X free seats&ldquo; and join with one click — <strong>without</strong> the names of already-registered members (privacy).</>}
              </span>
            </span>
          </label>

          {/* Approval-Pflicht durch Team-Lead — nur sinnvoll, wenn offene Teams sichtbar sind */}
          <label className={cx('dex-ui-toggle-row', teamJoinRequiresApproval && 'is-active', (!teamRegistrationEnabled || !teamOpenSlotsVisible) && 'is-disabled')}>
            <input
              type="checkbox"
              checked={teamJoinRequiresApproval}
              disabled={!teamRegistrationEnabled || !teamOpenSlotsVisible}
              onChange={e => setTeamJoinRequiresApproval(e.target.checked)}
            />
            <span className="dex-ui-toggle-row-body">
              <span className="dex-ui-toggle-row-title">
                {isDe ? 'Der Team-Lead bestätigt jeden Beitritt' : 'The team lead approves every join'}
                <span onClick={e => e.preventDefault()}>
                <InfoTooltip text={isDe
            ? <>
                <strong>Was du hier einstellst:</strong> ob jede Beitrittsanfrage zu einem offenen Team-Slot erst vom Team-Lead bestätigt werden muss.<br /><br />
                <strong>Anzeige in der App:</strong> der Team-Lead bekommt eine Mail mit <strong>&bdquo;Bestätigen&ldquo;</strong>- und <strong>&bdquo;Ablehnen&ldquo;</strong>-Buttons pro Anfrage. Bis zur Bestätigung steht der Beitretende in einer Approve-Queue und ist noch nicht offiziell im Team.<br /><br />
                <strong>Auswirkung für Teilnehmer:</strong> wenn aktiv, wird der Beitritt erst nach Bestätigung gültig — und der Beitretende bekommt erst dann seine Bestätigungsmail und (falls Outlook aktiv) den Kalendertermin. Wenn aus: Beitritt ist sofort gültig.
              </>
            : <>
                <strong>What this controls:</strong> whether every join request to an open team slot has to be confirmed by the team lead first.<br /><br />
                <strong>Where you see it:</strong> the team lead receives an email with <strong>&ldquo;Confirm&rdquo;</strong> and <strong>&ldquo;Reject&rdquo;</strong> buttons per request. Until confirmed, the joiner sits in an approve queue and is not yet officially in the team.<br /><br />
                <strong>For attendees:</strong> if active, the join only becomes valid after confirmation — and the joiner receives their confirmation mail and (if Outlook is enabled) the calendar invite only at that point. If off: join is immediately valid.
              </>
                } />
                </span>
              </span>
              <span className="dex-ui-toggle-row-desc">
                {isDe
                  ? <>Dann bekommt der Team-Lead je Anfrage eine Mail mit &bdquo;Bestätigen / Ablehnen&ldquo;; erst nach seiner Zusage ist die Person im Team und erhält Mail und Termin. Aus: der Beitritt gilt sofort.</>
                  : <>Then the team lead gets an email with &bdquo;Confirm / Reject&ldquo; per request; only after approval is the person in the team and receives mail and invite. Off: the join is valid immediately.</>}
              </span>
            </span>
          </label>
        </div>
      </div>
    </div>

    {/* v22.78: Eigener Team-Begriff (frei benennbar wie Event-Sections).
        v31.2: Feinschliff, deshalb im Aufklapper — die Seite endet mit den
        vier Kernfragen. Der Zähler verrät, ob etwas gesetzt ist, auch wenn
        der Aufklapper zu ist; die Eingaben sind wie zuvor nur bei aktiver
        Team-Anmeldung bedienbar. */}
    <div className="dex-ui-section">
      <button type="button" className={cx('dex-ui-disclosure', moreOpen && 'is-open')} aria-expanded={moreOpen} onClick={() => setMoreOpen(o => !o)}>
        <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
        {isDe ? 'Weitere Einstellungen' : 'More settings'}
        <span className="dex-ui-disclosure-count">
          {hasCustomTerm ? (isDe ? '1 angepasst' : '1 customised') : (isDe ? 'Standard' : 'default')}
        </span>
      </button>
      {moreOpen && (
        <div className="dex-ui-disclosure-body">
          <div className={cx('dex-ui-card', !teamRegistrationEnabled && 'dex-ui-card--muted')}>
            <div className="dex-ui-field">
              <label className="dex-ui-label" htmlFor="dex-wizard-team-term-singular">
                {isDe ? <>Wie sollen die Teams heißen? <span className="dex-ui-label-optional">(optional)</span></> : <>What should the teams be called? <span className="dex-ui-label-optional">(optional)</span></>}
                <InfoTooltip text={isDe
                  ? <><strong>Was du hier einstellst:</strong> einen eigenen Begriff für die Teams — z.B. <strong>„Break-Out Session“</strong>, „Gruppe“ oder „Tisch“. Leer = Standard „Team“.<br /><br /><strong>Anzeige in der App:</strong> ersetzt das Wort „Team“ überall (Organizer Center, „Meine Events“, Anmeldeformular).</>
                  : <><strong>What this controls:</strong> a custom term for the teams — e.g. <strong>“Break-Out session”</strong>, “group” or “table”. Empty = default “Team”.<br /><br /><strong>Where you see it:</strong> replaces the word “Team” everywhere (organizer center, “My Events”, registration form).</>} />
              </label>
              <div className="dex-ui-grid-2" style={{ gap: 10 }}>
                <input
                  id="dex-wizard-team-term-singular"
                  type="text" className="form-input"
                  value={teamTermSingular}
                  disabled={!teamRegistrationEnabled}
                  onChange={e => setTeamTermSingular(e.target.value)}
                  placeholder={isDe ? 'Einzahl, z.B. Break-Out Session' : 'Singular, e.g. Break-out session'}
                />
                <input
                  type="text" className="form-input"
                  value={teamTermPlural}
                  disabled={!teamRegistrationEnabled}
                  onChange={e => setTeamTermPlural(e.target.value)}
                  placeholder={isDe ? 'Mehrzahl, z.B. Break-Out Sessions' : 'Plural, e.g. Break-out sessions'}
                  aria-label={isDe ? 'Bezeichnung Mehrzahl' : 'Label plural'}
                />
              </div>
              <div className="dex-ui-help">
                {isDe
                  ? <>Leer heißt &bdquo;Team&ldquo;. Der Begriff ersetzt das Wort überall: Anmeldeformular, &bdquo;Meine Events&ldquo;, Organizer Center.</>
                  : <>Empty means &bdquo;Team&ldquo;. The term replaces the word everywhere: registration form, &bdquo;My Events&ldquo;, Organizer Center.</>}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>

    {/* v15: alter Hinweis „Logik folgt mit v11.82+" entfernt —
        die komplette Team-Anmelde-Logik (Multi-Person-Form,
        Mails, Outlook, Slot-Beitritt, Lead-Approval, Admin-Center-
        Team-Management) ist seit v11.82–v11.86 live. */}

    {/* v20.2: Der Self-Check-in-Block (v18.33) ist aus dem Wizard
        ausgezogen — Self-Check-in ist jetzt grundsätzlich immer
        verfügbar und wird beim ersten Klick auf die Aktionen
        (Check-in-Seite, Admin Center, QR-Kachel im Event-Detail)
        automatisch aktiviert. Zeitfenster + Deaktivieren: im
        Kachel-Modal des Admin Centers. */}

    </div>
  );
};
