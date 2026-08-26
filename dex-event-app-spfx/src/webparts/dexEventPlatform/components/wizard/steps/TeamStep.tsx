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
  return (
    <div style={{ display: visible ? 'block' : 'none' }}>
    <h2 className="dex-step-head-title">
      {isDe ? 'Schritt 7 — Team-Anmeldung' : 'Step 7 — Team Registration'}
    </h2>
    <p className="dex-step-head-lead">
      {isDe
        ? <><strong>Optional</strong> — erlaube einer Person, ein ganzes Team gleichzeitig anzumelden. Praktisch z.B. für Lauf-Teams, Workshop-Gruppen oder Tische bei einer Abendveranstaltung. Default: aus.</>
        : <><strong>Optional</strong> — let a single person register an entire team in one go. Handy e.g. for running teams, workshop groups or tables at an evening event. Default: off.</>}
    </p>

    {/* Toggle Team-Anmeldung erlauben */}
    <div style={{
      background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
      padding: '14px 16px', marginBottom: 12,
      border: '1px solid var(--dex-gray-200)',
    }}>
      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
        <input
          type="checkbox"
          checked={teamRegistrationEnabled}
          onChange={e => setTeamRegistrationEnabled(e.target.checked)}
          style={{ marginTop: 3, cursor: 'pointer' }}
        />
        <span style={{ flex: 1 }}>
          <strong>{isDe ? 'Team-Anmeldung erlauben' : 'Allow team registration'}</strong>
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
          <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
            {isDe
              ? 'Wenn aktiviert, kann eine Person ein ganzes Team anmelden — die anderen Mitglieder bekommen Bestätigungsmail + Outlook-Termin automatisch.'
              : 'When enabled, one person can register an entire team — the other members automatically receive a confirmation mail + Outlook invite.'}
          </span>
        </span>
      </label>
    </div>

    {/* Team-Größe + Team-Name-Frage — ausgegraut wenn Team-Anmeldung aus */}
    <div style={{
      background: teamRegistrationEnabled ? '#ffffff' : 'var(--dex-gray-50, #fafafa)',
      borderRadius: 12, padding: '14px 16px', marginBottom: 12,
      border: '1px solid var(--dex-gray-200)',
      opacity: teamRegistrationEnabled ? 1 : 0.55,
      transition: 'opacity 0.2s ease',
    }}>
      <div style={{ marginBottom: 14 }}>
        <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <strong>{isDe ? 'Team-Größe' : 'Team size'}</strong>
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
        <input
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
      </div>
      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}>
        <input
          type="checkbox"
          checked={askTeamName}
          disabled={!teamRegistrationEnabled}
          onChange={e => setAskTeamName(e.target.checked)}
          style={{ marginTop: 3, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}
        />
        <span style={{ flex: 1 }}>
          <strong>{isDe ? 'Team-Namen abfragen' : 'Ask for team name'}</strong>
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
          <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
            {isDe
              ? 'Wenn aktiv, gibt der Team-Lead bei der Anmeldung einen Team-Namen ein, der dann auf der MyEvents-Seite und in offenen Slots angezeigt wird.'
              : 'When enabled, the team lead enters a team name during registration which is shown on the MyEvents page and in open slots.'}
          </span>
        </span>
      </label>
    </div>

    {/* v22.78: Eigener Team-Begriff (frei benennbar wie Event-Sections)
        + „Teilnehmer dürfen keine neuen Teams erstellen". */}
    <div style={{
      background: teamRegistrationEnabled ? '#ffffff' : 'var(--dex-gray-50, #fafafa)',
      borderRadius: 12, padding: '14px 16px', marginBottom: 12,
      border: '1px solid var(--dex-gray-200)',
      opacity: teamRegistrationEnabled ? 1 : 0.55, transition: 'opacity 0.2s ease',
    }}>
      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
        <strong>{isDe ? 'Bezeichnung (statt „Team")' : 'Label (instead of “Team”)'}</strong>
        <InfoTooltip text={isDe
          ? <><strong>Was du hier einstellst:</strong> einen eigenen Begriff für die Teams — z.B. <strong>„Break-Out Session“</strong>, „Gruppe“ oder „Tisch“. Leer = Standard „Team“.<br /><br /><strong>Anzeige in der App:</strong> ersetzt das Wort „Team“ überall (Organizer Center, „Meine Events“, Anmeldeformular).</>
          : <><strong>What this controls:</strong> a custom term for the teams — e.g. <strong>“Break-Out session”</strong>, “group” or “table”. Empty = default “Team”.<br /><br /><strong>Where you see it:</strong> replaces the word “Team” everywhere (organizer center, “My Events”, registration form).</>} />
      </label>
      <div className="form-grid-2col" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
        <input
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
        />
      </div>
      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginTop: 14, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}>
        <input
          type="checkbox"
          checked={teamMembersCannotCreate}
          disabled={!teamRegistrationEnabled}
          onChange={e => setTeamMembersCannotCreate(e.target.checked)}
          style={{ marginTop: 3, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}
        />
        <span style={{ flex: 1 }}>
          <strong>{isDe ? 'Teilnehmer dürfen keine neuen Teams erstellen' : 'Participants cannot create new teams'}</strong>
          <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
            {isDe
              ? 'Empfohlen für Break-Out-Sessions: Die Teilnehmer melden sich normal an, die Zuordnung in die Teams/Break-outs nimmst DU als Organizer vor (per Drag & Drop im Organizer Center).'
              : 'Recommended for break-out sessions: participants register normally, and YOU assign them to teams/break-outs as the organizer (drag & drop in the Organizer Center).'}
          </span>
        </span>
      </label>
    </div>

    {/* v11.81: Beitritts-Modus — Sub-Box mit Modus + Sichtbarkeit + Approval */}
    <div style={{
      background: teamRegistrationEnabled ? '#ffffff' : 'var(--dex-gray-50, #fafafa)',
      borderRadius: 12, padding: '14px 16px', marginBottom: 12,
      border: '1px solid var(--dex-gray-200)',
      opacity: teamRegistrationEnabled ? 1 : 0.55,
      transition: 'opacity 0.2s ease',
      // v22.78: Beitritts-Modus ist irrelevant, wenn Teilnehmer keine
      // Teams erstellen/beitreten (Organizer ordnet zu) — dann ausgrauen.
      ...(teamMembersCannotCreate ? { opacity: 0.45 } : {}),
    }}>
      <div style={{ fontWeight: 700, fontSize: '0.95rem', marginBottom: 10, color: 'var(--dex-gray-800)' }}>
        {isDe ? 'Beitritts-Modus' : 'Join mode'}
      </div>

      {/* Radio-Group: komplette vs. Teil-Teams */}
      <div style={{ marginBottom: 14 }}>
        <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginBottom: 8, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}>
          <input
            type="radio"
            name="teamPartialMode"
            checked={!teamPartialAllowed}
            disabled={!teamRegistrationEnabled}
            onChange={() => setTeamPartialAllowed(false)}
            style={{ marginTop: 3, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}
          />
          <span style={{ flex: 1 }}>
            <strong>{isDe ? 'Nur komplette Teams' : 'Only complete teams'}</strong>
            <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
              {isDe
                ? 'Der Team-Lead muss alle N Mitglieder beim Anmelden eintragen. Halbe Teams sind nicht möglich.'
                : 'The team lead must enter all N members during registration. Partial teams are not possible.'}
            </span>
          </span>
        </label>
        <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}>
          <input
            type="radio"
            name="teamPartialMode"
            checked={teamPartialAllowed}
            disabled={!teamRegistrationEnabled}
            onChange={() => setTeamPartialAllowed(true)}
            style={{ marginTop: 3, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}
          />
          <span style={{ flex: 1 }}>
            <strong>{isDe ? 'Auch Teil-Teams erlaubt' : 'Partial teams allowed'}</strong>
            <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
              {isDe
                ? 'Der Team-Lead kann z.B. 2 von 4 Mitgliedern anmelden, die restlichen 2 Slots bleiben offen — andere Personen können später beitreten (siehe nächste Option).'
                : 'The team lead can register e.g. 2 of 4 members; the remaining 2 slots stay open — others can join later (see next option).'}
            </span>
          </span>
        </label>
      </div>

      {/* Checkbox: Sichtbarkeit offener Slots */}
      <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, marginBottom: 12, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}>
        <input
          type="checkbox"
          checked={teamOpenSlotsVisible}
          disabled={!teamRegistrationEnabled}
          onChange={e => {
            const v = e.target.checked;
            setTeamOpenSlotsVisible(v);
            if (!v) setTeamJoinRequiresApproval(false);
          }}
          style={{ marginTop: 3, cursor: teamRegistrationEnabled ? 'pointer' : 'not-allowed' }}
        />
        <span style={{ flex: 1 }}>
          <strong>{isDe ? 'Unvollständige Teams öffentlich für Beitritt sichtbar' : 'Open teams publicly visible for joining'}</strong>
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
          <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
            {isDe
              ? <>Wenn aktiv: andere Teilnehmer sehen offene Slots in der Registrierungsseite als &bdquo;Team mit X freien Plätzen&ldquo; — <strong>ohne</strong> die Namen der bereits angemeldeten Mitglieder (Privatsphäre).</>
              : <>When active: other attendees see open slots on the registration page as &ldquo;team with X free seats&rdquo; — <strong>without</strong> the names of already-registered members (privacy).</>}
          </span>
        </span>
      </label>

      {/* Checkbox: Approval-Pflicht durch Team-Lead */}
      <label style={{
        display: 'flex', alignItems: 'flex-start', gap: 10,
        cursor: (teamRegistrationEnabled && teamOpenSlotsVisible) ? 'pointer' : 'not-allowed',
        opacity: (teamRegistrationEnabled && teamOpenSlotsVisible) ? 1 : 0.55,
        transition: 'opacity 0.2s ease',
      }}>
        <input
          type="checkbox"
          checked={teamJoinRequiresApproval}
          disabled={!teamRegistrationEnabled || !teamOpenSlotsVisible}
          onChange={e => setTeamJoinRequiresApproval(e.target.checked)}
          style={{ marginTop: 3, cursor: (teamRegistrationEnabled && teamOpenSlotsVisible) ? 'pointer' : 'not-allowed' }}
        />
        <span style={{ flex: 1 }}>
          <strong>{isDe ? 'Beitritt erfordert Bestätigung durch Team-Kapitän' : 'Joining requires team captain approval'}</strong>
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
          <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
            {isDe
              ? 'Wenn aktiv: jeder Beitritt zu einem offenen Team geht erst in eine Approve-Queue. Der Team-Lead bekommt eine Mail mit „Bestätigen / Ablehnen"-Buttons. Erst nach Bestätigung ist die Person im Team. Wenn aus: Beitritt ist sofort gültig.'
              : 'When active: every join to an open team enters an approve queue. The team lead gets an email with "Confirm / Reject" buttons. Only after confirmation is the person in the team. When off: joins are immediately valid.'}
          </span>
        </span>
      </label>
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
