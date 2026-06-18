/**
 * Flowchart-Seite - Visualisierung der Prozesse
 *
 * Zeigt die einzelnen Abläufe (Anmeldung, Abmeldung, ID-Reorder, Nachrücken)
 * als verständliche Flowcharts für Dritte.
 */

import * as React from 'react';

// ==================== Flowchart-Bausteine ====================

interface FlowNodeProps {
  type: 'start' | 'end' | 'process' | 'decision' | 'data' | 'subprocess';
  label: string;
  color?: string;
  /** Optionale Prosa-Erklärung unter dem Knoten: was passiert & warum. */
  details?: string;
}

function FlowNode({ type, label, color, details }: FlowNodeProps): React.ReactElement {
  const baseStyle: React.CSSProperties = {
    padding: '10px 16px',
    fontSize: '0.8rem',
    textAlign: 'center',
    lineHeight: 1.3,
    fontWeight: 500,
    maxWidth: 320,
    margin: '0 auto',
  };

  const styles: Record<string, React.CSSProperties> = {
    start: {
      ...baseStyle,
      background: color || 'var(--dex-green)',
      color: '#fff',
      borderRadius: 24,
    },
    end: {
      ...baseStyle,
      background: color || 'var(--dex-gray-400)',
      color: '#fff',
      borderRadius: 24,
    },
    process: {
      ...baseStyle,
      background: color || '#e3f2fd',
      color: '#1565c0',
      borderRadius: 8,
      border: '2px solid #90caf9',
    },
    decision: {
      ...baseStyle,
      background: color || '#fff3e0',
      color: '#e65100',
      borderRadius: 8,
      border: '2px solid #ffcc80',
      transform: 'rotate(0deg)',
      fontStyle: 'italic',
    },
    data: {
      ...baseStyle,
      background: color || '#f3e5f5',
      color: '#6a1b9a',
      borderRadius: 8,
      border: '2px solid #ce93d8',
    },
    subprocess: {
      ...baseStyle,
      background: color || '#e8f5e9',
      color: '#2e7d32',
      borderRadius: 8,
      border: '2px solid #a5d6a7',
      borderLeft: '6px solid #66bb6a',
    },
  };

  // v23.44: BPMN-konforme Form pro Typ — Kreise für Ereignisse (Start/Ende),
  // Raute mit ✕ für exklusive Verzweigungen (Entscheidung), abgerundete
  // Rechtecke für Aufgaben. Bei Kreis/Raute steht der Text darunter (wie in BPMN).
  const renderChip = (): React.ReactElement => {
    if (type === 'start' || type === 'end') {
      const circle: React.CSSProperties = type === 'start'
        ? { background: color || 'var(--dex-green)', color: '#fff', border: '2px solid transparent' }
        : { background: '#fff', color: color || 'var(--dex-red, #c00)', border: `3px solid ${color || 'var(--dex-red, #c00)'}` };
      return (
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6, maxWidth: 320, margin: '0 auto' }}>
          <span style={{ width: 48, height: 48, borderRadius: '50%', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: '1.1rem', flexShrink: 0, ...circle }}>
            {type === 'start' ? '▶' : '◼'}
          </span>
          <span style={{ fontSize: '0.82rem', fontWeight: 700, textAlign: 'center', color: 'var(--dex-gray-700)' }}>{label}</span>
        </div>
      );
    }
    if (type === 'decision') {
      return (
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6, maxWidth: 360, margin: '0 auto' }}>
          <span style={{ width: 42, height: 42, transform: 'rotate(45deg)', background: color || '#fff3e0', border: '2px solid #ffcc80', borderRadius: 4, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
            <span style={{ transform: 'rotate(-45deg)', color: '#e65100', fontWeight: 800 }}>✕</span>
          </span>
          <span style={{ fontSize: '0.82rem', fontStyle: 'italic', textAlign: 'center', color: '#e65100' }}>{label}</span>
        </div>
      );
    }
    return <div style={styles[type] || styles.process}>{label}</div>;
  };

  if (!details) {
    return renderChip();
  }
  // v9.19: Details werden als Bullet-List gerendert (statt Fließtext).
  // Sätze werden an ". " aufgespalten, jeder Satz wird ein <li>. Schlagwörter
  // werden automatisch fett markiert: Status-Namen, DEX-Entitäten, Tech-Begriffe,
  // sowie alles in `code` oder **bold** Markdown-Schreibweise.
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', maxWidth: 540, margin: '0 auto' }}>
      {renderChip()}
      <div style={{
        marginTop: 6,
        padding: '8px 14px',
        fontSize: '0.78rem',
        lineHeight: 1.55,
        color: 'var(--dex-gray-700)',
        background: 'var(--dex-gray-50, #fafafa)',
        border: '1px solid var(--dex-gray-200)',
        borderRadius: 8,
        textAlign: 'left',
        maxWidth: 520,
        width: '100%',
      }}>
        {renderBulletDetails(details)}
      </div>
    </div>
  );
}

/**
 * v9.19: Details-Renderer mit Auto-Bold + Bullet-Split.
 * Sätze werden bei ". " getrennt. Schlagwörter werden in <strong> gewickelt.
 */
function renderBulletDetails(details: string): React.ReactElement {
  // v23.44: NUR noch anwenderfreundliche Schlagwörter fett markieren — bewusst
  // KEINE technischen Begriffe (Listen-Namen, Hintergrund-Prozesse, Tech-Jargon)
  // mehr, damit die Erklärungen für Nicht-ITler lesbar bleiben.
  const BOLD_KEYWORDS = [
    // Anmelde-Status
    'Angemeldet', 'Warteliste', 'Eingecheckt', 'Abgemeldet', 'QR versendet', 'No-Show', 'Entwurf',
    // Rollen / Personen
    'Admin', 'Admins', 'Organizer', 'Co-Organizer', 'Teilnehmer', 'Team-Kapitän',
    // anwenderfreundliche Kernbegriffe
    'automatisch', 'Teilnehmernummer', 'Wartelisten-Platz', 'Bestätigungs-Mail', 'Outlook-Termin',
  ];
  // Sortierung: längste Keywords zuerst (verhindert dass "die Event-Liste" in
  // "die Event-ListeListItem" das längere matched). Plus String-Escape für Regex.
  const escaped = BOLD_KEYWORDS
    .slice()
    .sort((a, b) => b.length - a.length)
    .map(k => k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
    .join('|');
  const keywordRegex = new RegExp(`(${escaped})`, 'g');
  // Sätze splitten: an ". " ABER nicht an "z.B.", "u.a." etc. (Heuristik: kein Punkt
  // direkt vor einem Buchstaben in derselben Wortgruppe).
  const sentences = details
    .split(/(?<=\.\s)(?=[A-ZÄÖÜ])/g)
    .map(s => s.trim())
    .filter(Boolean);
  const renderSentence = (s: string, idx: number): React.ReactElement => {
    const parts = s.split(keywordRegex);
    return (
      <li key={idx} style={{ marginBottom: 4 }}>
        {parts.map((p, i) => keywordRegex.test(p) ? <strong key={i}>{p}</strong> : <React.Fragment key={i}>{p}</React.Fragment>)}
      </li>
    );
  };
  // Wenn nur 1 Satz: kein Bullet, einfach Absatz.
  if (sentences.length <= 1) {
    const parts = details.split(keywordRegex);
    return <span>{parts.map((p, i) => keywordRegex.test(p) ? <strong key={i}>{p}</strong> : <React.Fragment key={i}>{p}</React.Fragment>)}</span>;
  }
  return (
    <ul style={{ margin: 0, paddingLeft: 18 }}>
      {sentences.map(renderSentence)}
    </ul>
  );
}

function Arrow({ label }: { label?: string }): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', margin: '4px 0' }}>
      <div style={{ width: 2, height: label ? 8 : 16, background: 'var(--dex-gray-300)' }} />
      {label && (
        <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)', fontWeight: 600, margin: '2px 0' }}>{label}</span>
      )}
      {label && <div style={{ width: 2, height: 8, background: 'var(--dex-gray-300)' }} />}
      <div style={{ width: 0, height: 0, borderLeft: '6px solid transparent', borderRight: '6px solid transparent', borderTop: '8px solid var(--dex-gray-300)' }} />
    </div>
  );
}

function BranchContainer({ children }: { children: React.ReactNode }): React.ReactElement {
  return (
    <div style={{ display: 'flex', gap: 24, justifyContent: 'center', margin: '4px 0' }}>
      {children}
    </div>
  );
}

function Branch({ label, children }: { label: string; children: React.ReactNode }): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', minWidth: 140 }}>
      <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)', fontWeight: 700, marginBottom: 4, textTransform: 'uppercase' }}>{label}</span>
      <div style={{ width: 2, height: 12, background: 'var(--dex-gray-300)' }} />
      <div style={{ width: 0, height: 0, borderLeft: '6px solid transparent', borderRight: '6px solid transparent', borderTop: '8px solid var(--dex-gray-300)' }} />
      {children}
    </div>
  );
}

// ==================== Flowcharts ====================

function RegistrationFlow(): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <FlowNode
        type="start"
        label="User klickt 'Registrieren'"
        details="Einstieg: der User öffnet ein Event und klickt auf den grünen Register-Button. Events, bei denen der User nicht im Audience-Filter steht oder die Registration-Deadline abgelaufen ist, zeigen statt des Buttons einen Sperr-Banner."
      />
      <Arrow />
      <FlowNode
        type="process"
        label="Basis-Daten eingeben"
        details="Im Registrierungs-Formular trägt der User die Stamm-Daten ein (Anrede, Vorname, Nachname, E-Mail), plus event-spezifische Felder (Zusatzfelder wie Travel-Zustimmung, Zimmerpartner-Picker, Reisekosten etc.). Bei B2Run-Events mit Split-Kapazitäten wählt er zusätzlich den Wunsch-Starter-Typ. Name und E-Mail werden aus dem SPFx-Profil vorausgefüllt, können aber überschrieben werden (nur mit Berechtigung für 'andere Person anmelden')."
      />
      <Arrow />
      <FlowNode
        type="decision"
        label="Bereits registriert für dieses Event?"
        details="Der Service prüft per die eigene Anmeldung in der Teilnehmerliste, ob schon ein Eintrag zu dieser E-Mail existiert. Wird verhindern dass ein User doppelt gezählt wird — und gleichzeitig erlauben, dass jemand, der sich zuvor abgemeldet hat, sich wieder anmelden kann."
      />
      <BranchContainer>
        <Branch label="Ja — zuvor abgemeldet">
          <FlowNode
            type="process"
            label="Reaktivierung"
            details="Der bestehende Abgemeldet-Eintrag wird wiederverwendet statt neu angelegt: Status zurück auf 'Angemeldet' oder 'Warteliste' (je nach Kapazität), CancellationDate gelöscht, RegistrationDate auf jetzt gesetzt. Außerdem wird — analog zur Erst-Anmeldung — atomar eine neue TeilnehmerID am Counter gezogen: wer mal #12 war und reaktiviert, bekommt jetzt z.B. die #87, also die nächst-freie ID am Ende der Liste. So bleibt die Änderungs-Historie (ChangeLog) erhalten und der Eintrag hat sofort eine sichtbare ID inklusive QR-Code, ohne auf den nächsten Cancel + IDReorder-Flow warten zu müssen."
          />
        </Branch>
        <Branch label="Ja — aktiv">
          <FlowNode
            type="end"
            color="var(--dex-red)"
            label="Fehler: Bereits angemeldet"
            details="Die Anmeldung bricht ab und der User sieht eine Fehlermeldung. So wird Doppel-Zählung in das Teilnehmer-Register verhindert. Wenn der User sich von der Warteliste auf 'aktiv' promoten lassen möchte, muss er sich zuerst abmelden und neu anmelden."
          />
        </Branch>
        <Branch label="Nein — neu">
          <FlowNode
            type="decision"
            label="B2Run-Split-Kapazitäten aktiv?"
            details="Die App prüft DurchstarterCapacity > 0 UND FunstarterCapacity > 0 im Event-Record. Nur bei aktivem Split greift die typ-bewusste Logik darunter; bei normalen Events läuft der Weg mit einer einzigen MaxParticipants-Schwelle."
          />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode
        type="decision"
        label="Wunsch-Starter-Typ hat freie Plätze?"
        details="Die App zählt aktive Teilnehmer pro Typ (Status = Angemeldet / QR versendet / Eingecheckt) und vergleicht mit der jeweiligen Kapazität. Der Wunsch-Typ hat Vorrang: nur wenn er voll ist, wird der Alt-Typ oder die Warteliste angeboten."
      />
      <BranchContainer>
        <Branch label="Ja — Wunsch frei">
          <FlowNode
            type="process"
            color="#e8f5e9"
            label="Direkt angemeldet im Wunsch-Typ"
            details="Status = 'Angemeldet', StarterType = Wunsch (z.B. 'Durchstarter'), PreferredStarterType = Wunsch. Keine weitere Interaktion nötig."
          />
        </Branch>
        <Branch label="Nein — Wunsch voll, Alt frei">
          <FlowNode
            type="process"
            color="#fff3e0"
            label="Fallback-Dialog"
            details="Die App zeigt ein Modal: 'Der Wunsch-Typ ist voll. Möchtest du stattdessen als <Alt-Typ> starten, oder auf die <Wunsch>-Warteliste?' Der User entscheidet explizit — kein stiller Auto-Fallback mehr. Das war vor v6.5 anders: damals wurde still umgeschaltet, was zu Überraschungen führte."
          />
        </Branch>
        <Branch label="Nein — beide voll">
          <FlowNode
            type="process"
            color="#ffebee"
            label="Warteliste für Wunsch-Typ"
            details="Status = 'Warteliste', StarterType = leer (wird erst beim Nachrücken gesetzt), PreferredStarterType = Wunsch. So weiß der IDReorder-Flow später, welcher Typ-Warteliste der User angehört."
          />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode
        type="subprocess"
        label="Atomare TeilnehmerID-Vergabe (v7.28+, bei Erst-Anmeldung UND Reaktivierung)"
        details="Sowohl bei einer echt neuen Anmeldung als auch bei der Reaktivierung eines Abgemeldet-Items holt der Service die nächste TeilnehmerID atomar aus der Counter-Liste den Zähler (eine Liste pro Event-Bereich mit genau einem Item, NextValue=N). Vorgehen: GET counter-item mit konfliktsicher → PATCH NextValue=N+1 mit IF-MATCH=etag → bei HTTP 412 (jemand war schneller) wird mit kurzem Jitter retried (max 8x). Damit können mehrere User parallel anmelden, ohne dass IDs doppelt vergeben werden. Reaktivierte bekommen die nächst-freie ID am Ende der Liste — sie hängen sich also funktional wie ein Neuzugang hinten an, ihre alte ID bekommen sie nicht zurück. v7.31: Nach Cancel und nach 'IDs neu vergeben' wird der Counter auf den aktuellen Max-Wert gesynct (syncCounterToMax), damit er nicht über die echten IDs hinaus 'davonrast'. Im die Nummerierungs-Automatik-Flow wird am Ende der Counter ebenfalls auf den frischen Max-Wert gepatcht. Fallback bei alten Events ohne Counter-Liste: max(TeilnehmerID)+1 — dann muss der Admin einmalig 'Spalten fixen' klicken, damit die Counter-Liste angelegt + mit dem aktuellen Max-Wert geseedet wird."
      />
      <Arrow />
      <FlowNode
        type="subprocess"
        label="Eintrag in Teilnehmerliste"
        details="Der Service schreibt ein neues Item in die 'Teilnehmer'-Liste der Event-Event-Bereich: TeilnehmerID (eindeutig vom Counter — sowohl bei Erst-Anmeldung als auch bei Reaktivierung), Anrede, Vorname, Nachname, E-Mail, Status, StarterType, PreferredStarterType, RegistrationDate, RegisteredByName/Email (Audit), Antworten (JSON der event-spezifischen Felder). Lücken in der ID-Sequenz, die durch Abmeldungen entstehen, werden später vom die Nummerierungs-Aufträge-Flow geschlossen — er sortiert Aktive (1..N) und Warteliste (N+1..N+M) lückenlos um."
      />
      <Arrow />
      <FlowNode
        type="subprocess"
        label="das Teilnehmer-Register aktualisieren"
        details="In der zentralen das Teilnehmer-Register-Liste (auf Site-Ebene, listet alle User und ihre Event-Registrierungen) wird die EventNumber in das Feld EventRegistered (Angemeldet) bzw. EventOnWaitlist (Warteliste) eingetragen. Diese Liste ist für den schnellen My-Events-Pfad — damit muss nicht jede Event-Event-Bereich einzeln nach dem User durchsucht werden."
      />
      <Arrow />
      <FlowNode
        type="decision"
        label="Event-Flag 'DisableEmails' aktiv?"
        details="Organizer kann im Event-Editor pro Event die automatischen Mails unterdrücken — z.B. für Test-Events oder wenn manuelle Einladungen geplant sind."
      />
      <BranchContainer>
        <Branch label="Nein — Mails aktiv">
          <FlowNode
            type="data"
            label="die Mail-Warteschlange: Bestätigungs-Mail in Queue"
            details="Das Template-Type je nach Status: 'Anmeldung' oder 'Warteliste'. Das SP-Template wird aus die Mail-Vorlagen geladen (pro EmailLanguage DE/EN), Platzhalter {{Name}}, {{EventTitle}}, {{Organizer}}, {{WaitlistPosition}} werden aufgelöst. Event-spezifische Template-Overrides (z.B. eigener Anmelde-Text pro Event) werden vorher angewendet. Der Body wird im Deloitte-Layout gewrappt und in die Mail-Warteschlange mit Status='Pending' geschrieben. Der den Mail-Versand-Flow versendet innerhalb von ~1 Minute."
          />
        </Branch>
        <Branch label="Ja — Mails deaktiviert">
          <FlowNode
            type="process"
            color="#f5f5f5"
            label="Kein Mail-Eintrag"
            details="Der Status-Eintrag in der Teilnehmer-Liste bleibt natürlich, nur die Mail wird nicht gequeued."
          />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode
        type="decision"
        label="'Outlook-Einladung' aktiv UND Status = Angemeldet?"
        details="Wartelisten-Teilnehmer bekommen keine Kalender-Einladung — erst beim Nachrücken. Und falls das Event 'DisableOutlook' gesetzt hat, wird gar keiner eingeladen."
      />
      <BranchContainer>
        <Branch label="Ja">
          <FlowNode
            type="data"
            label="die Kalender-Warteschlange: 'Einladen' in Queue"
            details="Ein Item mit ActionType='Einladen', EventId und Attendee (E-Mail) wird in die Kalender-Warteschlange gepushed. Der die Outlook-Automatik-Flow liest den Parent-Termin aus dem Shared Mailbox (das zentrale Event-Postfach) über Termin-ID (Termin-Verknüpfung im Event), fügt den Attendee per Graph PATCH hinzu und Outlook sendet dem User die Einladung."
          />
        </Branch>
        <Branch label="Nein">
          <FlowNode
            type="process"
            color="#f5f5f5"
            label="Kein Outlook-Eintrag"
          />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode
        type="end"
        label="Registrierung abgeschlossen"
        details="Der User sieht die Erfolgs-Seite. Bei Events mit Sub-Events (z.B. B2Run-Trainingssessions): Seit v6.14 wählt der User auf der Registrierungsseite direkt per Checkbox aus, wofür er sich anmelden möchte — Haupt-Event und/oder einzelne Sessions. Pro ausgewählter Session wird eine eigene Registrierung angelegt (separate Bestätigungsmail + Outlook-Termin). Bei B2Run-Parents kommt pro Session eine Durchstarter-/Funstarter-Auswahl hinzu; meldet sich der User gleichzeitig für das Haupt-Event an, wird der dort gewählte Starter-Typ automatisch auf die Session-TN-Listen übernommen. Wenn der User nur Sessions wählt (Haupt-Event abgehakt), zeigt der Success-Screen einen Sessions-Only-Hinweis und in 'My Events' erscheint der Parent-Eintrag mit dem orangefarbenen Badge 'Nur Sessions'."
      />
      <Arrow />
      {/* v11.82: Team-Anmeldung als paralleler Branch zur Solo-Anmeldung */}
      <FlowNode
        type="subprocess"
        color="#f3f8ec"
        label="Alternativ: Team-Anmeldung (v11.82)"
        details={'Wenn der Organizer in Schritt 4 die Team-Anmeldung aktiviert hat und der User den Toggle „Ich melde mich + mein Team an" setzt, läuft ein paralleler Submit-Pfad: Der Lead wählt N-1 weitere Teilnehmer per People-Picker aus und bestätigt vorab schriftlich, dass alle Mitglieder zugestimmt haben. Beim Submit reserviert die App atomar N Plätze auf einmal (reserveSeat mit count=N — entweder alle N oder das gesamte Team landet gemeinsam auf der Warteliste; kein Teil-Anmelden). Danach werden N parallele Teilnehmer-Items mit identischer TeamId (UUID) angelegt — genau ein Eintrag (Lead) hat TeamLead=true. Jedes Mitglied bekommt eine eigene Bestätigungs-Mail (inkl. „Du wurdest als Teil eines Teams angemeldet"-Hinweis-Box plus klarem Hinweis, falls kein Einverständnis vorlag, sich beim Organizer zu melden oder selbst über „Meine Events" abzumelden) und einen eigenen Outlook-Termin. In „Meine Events" sieht jedes Mitglied ein Team-Badge mit Belegungs-Anzahl, Lead-Markierung und Namen der anderen Mitglieder. Beitritts-/Approve-Flows für offene Team-Slots folgen in einer späteren Iteration. v23.0: Der Begriff „Team" ist pro Event frei benennbar (z.B. „Break-Out Session") — der gewählte Begriff erscheint überall in der App.'}
      />
      <Arrow />
      {/* v23.0: Organizer-gesteuerte Zuordnung + Per-Team-Mail */}
      <FlowNode
        type="subprocess"
        color="#f3f8ec"
        label="Alternativ: Organizer ordnet selbst zu (v23.0)"
        details={'Hat der Organizer in Schritt 4 den Schalter „Teilnehmer dürfen keine neuen Teams/Gruppen erstellen" aktiviert, melden sich alle einzeln an — die Team-Anmelde-Karte ist ausgeblendet. Die Aufteilung übernimmt dann der Organizer im Organizer Center: In der Teams-Sektion zieht er Personen per Drag & Drop zwischen den Teams/Break-Out-Sessions und der „ohne Team"-Box. Das Loslassen ruft assignRegistrationToTeam (setzt/leert TeamId, TeamName, TeamLead), promotet bei Bedarf einen neuen Lead und schreibt einen Änderungsprotokoll-Eintrag. Sind die Gruppen fertig eingeteilt, verschickt der Button „Mail an <Teams>" pro aktivem Mitglied eine eigene Mail im Deloitte-Layout: Der Organizer trägt pro Team eine eigene Info ein (z.B. einen dedizierten Microsoft-Teams-Einwahllink), die den Platzhalter {{TeamInfo}} ersetzt. So bekommt jede Break-Out-Session in einem Rutsch ihren eigenen Treffpunkt/Link. Gated über den E-Mail-Schalter aus Schritt 6.'}
      />
    </div>
  );
}

function CancellationFlow(): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <FlowNode
        type="start"
        label="Abmeldung ausgelöst"
        details={'Einstiegspunkt. Die Abmeldung kann von drei Seiten kommen: der Teilnehmer selbst (Self-Cancel aus My-Events), ein Admin/Organizer (aus dem Admin-Center) oder ein Team-Lead, der stellvertretend ein Team-Mitglied aus dem Modal „Team verwalten" entfernt (v11.86). Alle drei Pfade konvergieren auf denselben Cancel-Step in der Teilnehmerliste — unterscheiden sich aber im UX und in den Audit-Spalten CancelledByName/CancelledByEmail.'}
      />
      <Arrow />
      <FlowNode
        type="decision"
        label="Wer meldet ab?"
        details="Die App unterscheidet den Pfad, weil der Admin sofortiges Feedback über den Nachrücker braucht (damit er den nächsten Schritt planen kann), während ein einzelner User nur seinen eigenen Eintrag sieht und die kurze Verzögerung durch den Flow nicht wahrnimmt. Der Team-Lead-Pfad (v11.86) ähnelt dem User-Pfad, schreibt aber den Lead in die Audit-Felder."
      />
      <BranchContainer>
        <Branch label="User (Self-Cancel)">
          <FlowNode
            type="process"
            label="User klickt in My-Events auf 'Abmelden'"
            details="Zwei-Klick-Bestätigung (erster Klick zeigt 'Wirklich abmelden?', zweiter führt aus). Das verhindert versehentliche Abmeldungen. Zugriffsschutz sorgt dafür dass der User eh nur seinen Eintrag sieht — kein Info-Toast über den Nachrücker nötig. CancelledByName/CancelledByEmail = der eingeloggte User."
          />
        </Branch>
        <Branch label="Admin/Organizer">
          <FlowNode
            type="process"
            color="#fff3e0"
            label="Admin klickt im Admin-Center auf 'Abmelden'"
            details="Single-Click mit Confirm-Dialog. Nach Bestätigung erscheint sofort ein orange Toast oben rechts mit Spinner: 'Abmeldung von X wird verarbeitet…'. Nachrücker wird direkt client-seitig promoted (siehe unten). Admin sieht am Ende den Erfolgs-Toast mit Namen des Nachrückers. CancelledByName/CancelledByEmail = der Admin/Organizer."
          />
        </Branch>
        <Branch label="Team-Lead löscht Team-Mitglied (v11.86)">
          <FlowNode
            type="process"
            color="#e8f5e9"
            label="Lead klickt im Team-verwalten-Modal auf den Trash-Button"
            details={'In „Meine Events" öffnet der Team-Lead über „Team bearbeiten" das Modal „Team verwalten". Pro aktivem Mitglied (außer ihm selbst) gibt es einen roten Trash-Button. Klick öffnet eine zweite Confirm-Ebene mit dem Namen der Person und dem Hinweis, dass die Abmeldung stellvertretend erfolgt. Bestätigung ruft cancelTeamMember im EventContext auf. CancelledByName/CancelledByEmail = der eingeloggte Lead — die abgemeldete Person sieht im Audit also den Lead, nicht sich selbst. Der Rest des Cancel-Pfads (Status auf Abgemeldet, das Teilnehmer-Register, Team-Nachlauf, Mail, Outlook, IDReorder) ist identisch zum Self-Cancel.'}
          />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode
        type="subprocess"
        label="Teilnehmerliste: Status → 'Abgemeldet'"
        details="Per Aktualisierung-Request auf das Teilnehmer-Item: Status wird auf 'Abgemeldet' gesetzt, CancellationDate = jetzt, CancelledByName/Email werden zur Audit-Nachverfolgung gespeichert (wer hat die Abmeldung ausgelöst). TeilnehmerID bleibt vorerst — der IDReorder-Flow setzt sie später auf null für diesen Abgemeldeten."
      />
      <Arrow />
      <FlowNode
        type="subprocess"
        label="das Teilnehmer-Register aktualisieren"
        details="Die zentrale das Teilnehmer-Register-Liste enthält pro User die Liste der EventNumbers, für die er angemeldet oder auf der Warteliste ist. Bei Abmeldung wird die EventNumber aus EventRegistered bzw. EventOnWaitlist entfernt. So weiß die My-Events-Seite auf einen Blick, dass der User für dieses Event nicht mehr aktiv ist."
      />
      <Arrow />
      <FlowNode
        type="decision"
        label="War es eine Team-Anmeldung? (v11.83)"
        details="Wenn die abgemeldete Registrierung eine TeamId hat, läuft parallel zum Standard-Cancel-Pfad ein Team-Nachlauf: verbleibende Mitglieder werden ermittelt, ggf. der frühere Lead durch das früheste verbleibende Mitglied ersetzt (Promote), und alle bekommen eine Info-Mail."
      />
      <BranchContainer>
        <Branch label="Ja — Team-Cancel">
          <FlowNode
            type="subprocess"
            color="#e8f5e9"
            label="Verbleibende Mitglieder laden"
            details="getTeamMembers(subsiteUrl, TeamId) liefert alle Einträge derselben Team-Id. Davon filtert die App den gerade abgemeldeten Eintrag und alle Status='Abgemeldet'-Einträge heraus."
          />
          <Arrow />
          <FlowNode
            type="decision"
            label="War der Abgemeldete Team-Lead UND gibt es verbleibende Mitglieder?"
            details="Beide Bedingungen müssen wahr sein. Bei genau einer 1-Person-Anmeldung (Cancel des Solo-Leads) löst sich das Team einfach auf — kein Promote, keine Info-Mails."
          />
          <BranchContainer>
            <Branch label="Ja">
              <FlowNode
                type="subprocess"
                color="#e8f5e9"
                label="Auto-Promote des neuen Leads"
                details={'Sortier-Kriterium: kleinste TeilnehmerID, sonst früheste RegistrationDate, sonst kleinste Item-Id. Der Treffer bekommt per Aktualisierung TeamLead=true. In der Info-Mail an diesen Member steht ein extra Hinweis „Du bist jetzt der neue Team-Lead".'}
              />
            </Branch>
            <Branch label="Nein">
              <FlowNode type="process" color="#f5f5f5" label="Kein Promote nötig" />
            </Branch>
          </BranchContainer>
          <Arrow />
          <FlowNode
            type="data"
            label="die Mail-Warteschlange: Info-Mail pro Mitglied"
            details={'Pro verbleibendem Member wird eine Info-Mail im Deloitte-Layout in die Queue gelegt. Inhalt: Name der abgemeldeten Person, aktuelle Belegung „N/Team-Size", drei Handlungs-Optionen (nichts tun, Lead fügt neu hinzu, andere belegen den Slot via Anmelde-Seite). Best-effort — Fehler beim Queueing brechen den Cancel nicht ab.'}
          />
        </Branch>
        <Branch label="Nein — Solo-Cancel">
          <FlowNode type="process" color="#f5f5f5" label="Kein Team-Nachlauf" />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode
        type="decision"
        label="Event hat 'E-Mails versenden' aktiv?"
        details="Pro Event kann der Organizer die automatischen Mails unterdrücken (z.B. für manuell verwaltete Events)."
      />
      <BranchContainer>
        <Branch label="Ja">
          <FlowNode
            type="data"
            label="die Mail-Warteschlange: Abmelde-Mail in Queue"
            details="Template-Type 'Abmeldung' aus die Mail-Vorlagen wird geladen (pro EmailLanguage), Platzhalter {{Name}}, {{EventTitle}} aufgelöst, in Deloitte-Layout gewrappt. Status='Pending', den Mail-Versand-Flow verschickt binnen ~1 Minute. Empfänger ist der abgemeldete Teilnehmer selbst — als Bestätigung."
          />
        </Branch>
        <Branch label="Nein">
          <FlowNode type="process" color="#f5f5f5" label="Kein Mail-Eintrag" />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode
        type="decision"
        label="Event hat 'Outlook-Einladung' aktiv?"
        details="Wenn der User im Shared-Mailbox-Kalender als Attendee stand, muss er dort entfernt werden — damit sein Outlook-Termin verschwindet."
      />
      <BranchContainer>
        <Branch label="Ja">
          <FlowNode
            type="data"
            label="die Kalender-Warteschlange: 'Ausladen' in Queue"
            details="Ein die Kalender-Warteschlange-Item mit ActionType='Ausladen', EventId und Attendee-Email wird gepushed. Der die Outlook-Automatik-Flow liest den Kalender-Termin (Termin-ID aus dem Event), filtert den Attendee raus und PATCHt den Termin mit der neuen Attendee-Liste. Outlook benachrichtigt den User, dass sein Termin zurückgezogen wurde."
          />
        </Branch>
        <Branch label="Nein">
          <FlowNode type="process" color="#f5f5f5" label="Kein Outlook-Eintrag" />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode
        type="decision"
        label="Abmeldung durch Admin/Organizer?"
        details="Die zentrale Weiche: nur beim Admin-Pfad läuft der client-seitige Promote sofort. Bei Self-Cancel übernimmt der Flow die Arbeit (siehe IDReorder-Flow) — der User wartet nicht, weil er die Veränderung eh nicht sieht."
      />
      <BranchContainer>
        <Branch label="Ja — sofort promoten">
          <FlowNode
            type="subprocess"
            color="#e8f5e9"
            label="Client-seitiger Promote"
            details="Die App ruft promoteFirstWaitlistItem auf. Bei B2Run-Split-Kapazitäten wird der Filter auf PreferredStarterType eingeschränkt (onlyWithPreferredType = cancelledStarterType) — damit wird ein abgemeldeter Durchstarter nur durch einen anderen Durchstarter aus der Warteliste ersetzt, nicht durch einen Funstarter. Der Nachrücker bekommt Status='Angemeldet' und StarterType wird auf den freigewordenen Typ gesetzt."
          />
          <Arrow />
          <FlowNode
            type="data"
            label="Nachrück-Mail + Outlook-Einladung"
            details="Für den Nachrücker wird ein die Mail-Warteschlange-Item mit Template-Type 'Nachruecken' und ein die Kalender-Warteschlange-Item mit ActionType='Einladen' in die Queues geschrieben. Der Nachrücker erhält dadurch automatisch eine freundliche Nachrück-Mail und eine Kalender-Einladung — beides im Deloitte-Layout. Der Admin sieht im Toast: 'Nachgerückt: <Name> <E-Mail> (<Starter-Typ>)'."
          />
        </Branch>
        <Branch label="Nein — User-Self-Cancel">
          <FlowNode
            type="process"
            color="#f5f5f5"
            label="Kein Client-Promote"
            details="Der Flow übernimmt den Promote beim nächsten Reorder-Durchlauf (innerhalb ~1 Minute). So wird verhindert, dass der User seinen Client-Promote ausführt (der nur seine eigene Sicht kennt) und es zu Race-Conditions kommt."
          />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode
        type="data"
        color="#ffebee"
        label="die Nummerierungs-Aufträge: Reorder-Auftrag in Queue"
        details="Ein neues Item in der die Nummerierungs-Aufträge-Liste (Pending) löst den die Nummerierungs-Automatik-Flow aus. Der Flow läuft sequenziell (Concurrency=1) — bei mehreren Abmeldungen werden die Einträge nacheinander abgearbeitet. Er sortiert die TeilnehmerIDs neu (Zwei-Pass: Angemeldete zuerst, Warteliste danach) und übernimmt bei User-Self-Cancels auch das Nachrücken."
      />
      <Arrow />
      <FlowNode
        type="end"
        label="Abmeldung abgeschlossen"
        details="Bei Admin-Pfad: Toast bleibt sichtbar bis der Admin ihn schließt, der Reorder läuft im Hintergrund. Bei User-Pfad: die Teilnehmer-Liste ist nach dem Flow-Run (~1 Min) neu sortiert, Nachrücker (falls vorhanden) wurde benachrichtigt."
      />
    </div>
  );
}

function IDReorderFlow(): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <FlowNode
        type="start"
        color="#6a1b9a"
        label="die Hintergrund-Automatik Trigger"
        details="Der die Nummerierungs-Automatik-Flow pollt die die Nummerierungs-Aufträge-Liste im Minuten-Takt. Bei jedem neuen Eintrag mit Status='Pending' startet eine Flow-Instanz. Concurrency=1 stellt sicher, dass immer nur eine Instanz gleichzeitig läuft — bei mehreren kurz aufeinanderfolgenden Abmeldungen werden die Einträge sequenziell (in Reihenfolge der Erstellung) abgearbeitet. Maximal 100 Einträge können warten."
      />
      <Arrow />
      <FlowNode
        type="process"
        label="Status → 'Processing'"
        details="Sofort beim Start: das die Nummerierungs-Aufträge-Item wird auf 'Processing' gesetzt. Damit sieht der nächste Polling-Durchlauf das Item nicht mehr als 'Pending' und startet keinen Parallel-Run auf dem gleichen Item."
      />
      <Arrow />
      <FlowNode
        type="subprocess"
        label="Get_Enrolled_Participants"
        details="Aus der 'Teilnehmer'-Liste der Event-Event-Bereich werden alle Items geladen, deren Status NICHT 'Abgemeldet' ist. Das ergibt die zusammengefasste Liste von Angemeldeten (Status='Angemeldet' / 'QR versendet' / 'Eingecheckt') und Warteliste-Teilnehmern ('Warteliste'). Die Sortierung ist RegistrationDate asc — ältere Registrierungen zuerst."
      />
      <Arrow />
      <FlowNode
        type="process"
        label="Count_Active zählen"
        details="Mit einer Filter_Non_Waitlist-Query wird die Anzahl der bereits belegten Plätze berechnet (Items mit Status ≠ 'Warteliste'). Diese Zahl wird später benötigt, um zu prüfen, ob ein Nachrücken nötig ist. Bei Events mit MaxParticipants=0 (unbegrenzt) greift diese Prüfung nicht."
      />
      <Arrow />
      <FlowNode
        type="process"
        color="#e8f5e9"
        label="Sort_ByStatusPriority (Zwei-Pass)"
        details="Seit v6.6 der Kernkniff: die Enrolled-Liste wird in eine 'Status-Priority'-Reihenfolge gebracht — alle Angemeldeten zuerst (in RegistrationDate-Reihenfolge), dann alle Warteliste-Teilnehmer dahinter. So bekommen Angemeldete die niedrigen TeilnehmerIDs (1..N) und die Warteliste hängt lückenlos hinten an (N+1..N+M). Vor v6.6 wurden alle Enrolled in RegistrationDate-Reihenfolge durchnummeriert, wodurch Wartelisten-Items mitten zwischen Angemeldeten landeten — ein klassisches Sortier-Problem, das diese Compose-Action behebt."
      />
      <Arrow />
      <FlowNode
        type="process"
        label="Generate_Indices + GenerateSPData"
        details="Aus der sortierten Liste werden Index-basierte Paare erzeugt: jedem Item wird seine zukünftige TeilnehmerID zugewiesen (Position+1). Da die Liste bereits sortiert ist (Angemeldete vorne, Warteliste hinten), stimmen die zugewiesenen Zahlen automatisch mit dem gewünschten Schema überein."
      />
      <Arrow />
      <FlowNode
        type="subprocess"
        label="$batch-Update an "
        details="Per OData-$batch-API werden alle TeilnehmerIDs in einem einzigen Request-Paket (Chunks à 250) aktualisiert. Das ist deutlich schneller als einzelne PATCH-Calls pro Item (bei 100 Teilnehmern: 1 Call statt 100). WICHTIG: nur das Feld TeilnehmerID wird geschrieben — Status, Vorname, Nachname usw. bleiben unverändert. Warteliste-Items behalten ihren Warteliste-Status, nur die TID ändert sich."
      />
      <Arrow />
      <FlowNode
        type="subprocess"
        label="Get_EventDetails"
        details="Nach dem Reorder werden die Event-Metadaten aus die Event-Liste geladen (Title, MaxParticipants, EmailLanguage und — relevant für B2Run-Split — DurchstarterCapacity und FunstarterCapacity). Diese Daten bestimmen, wie der Nachrück-Prozess weiterläuft."
      />
      <Arrow />
      <FlowNode
        type="decision"
        label="Is_B2RunSplit?"
        details="Geprüft wird: DurchstarterCapacity > 0 UND FunstarterCapacity > 0. Nur dann arbeitet der Flow typ-bewusst. Bei klassischen Events (eine MaxParticipants-Zahl, keine Typ-Trennung) greift der einfachere 'Check_Nachrücken'-Pfad."
      />
      <BranchContainer>
        <Branch label="Ja — B2Run-Split">
          <FlowNode
            type="process"
            color="#e3f2fd"
            label="Durchstarter-Pass"
            details="Count_Active_Durchstarter zählt aktive Teilnehmer mit StarterType='Durchstarter'. Wenn diese Zahl < DurchstarterCapacity ist, holt der Flow per GET den ersten Warteliste-Eintrag mit PreferredStarterType='Durchstarter' (sortiert nach RegistrationDate). Existiert einer → Aktualisierung: Status=Angemeldet, StarterType=Durchstarter. Nachrück-Mail (Template 'Nachruecken') + Outlook-'Einladen' werden in die Queues geschrieben."
          />
          <Arrow />
          <FlowNode
            type="process"
            color="#fff3e0"
            label="Funstarter-Pass"
            details="Identisch wie oben, nur für Funstarter: Count_Active_Funstarter < FunstarterCapacity → Warteliste-Eintrag mit PreferredStarterType='Funstarter' promoten. Beide Pässe laufen nacheinander im gleichen Flow-Run — sie können also in einem einzigen Run BEIDE Typen befüllen, wenn beide Plätze frei haben und passende Warteliste-Teilnehmer existieren."
          />
        </Branch>
        <Branch label="Nein — klassisches Event">
          <FlowNode
            type="decision"
            label="Check_Nachrücken"
            details="Bei klassischen Events: wenn Count_Active < MaxParticipants UND MaxParticipants > 0. Der ==-Fall (z.B. 100/100) greift nicht — nur bei einer echten Unterbesetzung wird nachgerückt."
          />
          <Arrow />
          <FlowNode
            type="subprocess"
            label="Get_Waitlist_First → Promote"
            details="Der erste Warteliste-Eintrag (älteste RegistrationDate) wird per Aktualisierung auf Status='Angemeldet' gesetzt. KEINE TeilnehmerID-Änderung — die Zwei-Pass-Sortierung hat die korrekte TID bereits zugewiesen. Anschließend: Nachrück-Mail + Outlook-Einladung in Queue."
          />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode
        type="data"
        label="die Mail-Warteschlange: Nachrück-Mail in Queue"
        details="Template aus die Mail-Vorlagen (TemplateType='Nachruecken', Language passend zum Event), Platzhalter {{Name}} und {{EventTitle}} werden aufgelöst. Status='Pending' — den Mail-Versand-Flow versendet die Mail im nächsten Durchlauf (~1 Min)."
      />
      <Arrow />
      <FlowNode
        type="data"
        label="die Kalender-Warteschlange: Einladen in Queue"
        details="die Kalender-Warteschlange-Item mit ActionType='Einladen' und Attendee = E-Mail des Nachrückers. Der die Outlook-Automatik-Flow fügt den Attendee per Graph PATCH zum Shared-Mailbox-Termin hinzu — Outlook schickt automatisch die Kalender-Einladung."
      />
      <Arrow />
      <FlowNode
        type="process"
        label="Status → 'Done'"
        details="Das die Nummerierungs-Aufträge-Item wird final auf 'Done' gesetzt. Falls im Laufe des Flows ein Fehler auftritt (Graph-Timeout, 429 Rate-Limit, 404), führt der Error-Handler den Status stattdessen auf 'Failed' — so kann der Admin im Admin-Center oder direkt in der SP-Liste prüfen, welche Runs erfolglos waren."
      />
      <Arrow />
      <FlowNode
        type="end"
        color="#6a1b9a"
        label="Flow-Instanz beendet"
        details="Bei weiteren die Nummerierungs-Aufträge-Einträgen startet die nächste Instanz. Concurrency=1 garantiert sequentielle Abarbeitung — keine Race-Conditions bei paralleler Bearbeitung."
      />
      <Arrow />
      <FlowNode
        type="process"
        color="#f5f5f5"
        label="Hinweis: Admin-Cancels"
        details="Seit v6.8 promotet die App bei Admin/Organizer-Cancels den Nachrücker sofort client-seitig — inklusive Nachrück-Mail und Outlook-Einladung. Wenn der Flow dann später läuft, sieht er Count_Active = MaxParticipants und überspringt den Promote-Zweig. Er macht dann 'nur' noch die Zwei-Pass-Sortierung der TeilnehmerIDs — das bleibt essentiell, damit die Liste sauber bleibt."
      />
    </div>
  );
}

function EventCreationFlow(): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <FlowNode
        type="start"
        label="Organizer/Admin erstellt Event (8 Schritte: Grundlagen → Zeit&Ort → Kapazität → Team-Anmeldung → Felder → Kommunikation → Dokumente → Quiz)"
        details="v11.80 + erweitert v11.81: Zwischen Schritt 3 (Kapazität) und dem bisherigen Schritt 4 (Felder) ist ein neuer Schritt 4 'Team-Anmeldung' eingezogen — Basis-Settings: Toggle 'Team-Anmeldung erlauben' (Default aus), Team-Größe (Default 4, Min 2, Max 20), Toggle 'Team-Name abfragen' (Default aus). Beitritts-Modus (v11.81): Radio 'Nur komplette Teams' vs. 'Auch Teil-Teams erlaubt', Checkbox 'Unvollständige Teams öffentlich sichtbar', Checkbox 'Beitritt erfordert Lead-Bestätigung'. Aktuell wird nur die Konfiguration persistiert; die tatsächliche Multi-Person-Anmelde-Logik (Multi-Person-Form, automatische Mails an Mitglieder, Outlook-Einladungen, Slot-Beitritt, Approve-Queue) folgt mit v11.82+."
      />
      <Arrow />
      <FlowNode
        type="process"
        label="Schritt 4 — Team-Anmeldung konfigurieren (v11.80 + v11.81)"
        details="TeamRegistrationEnabled, TeamSize, AskTeamName (v11.80) sowie TeamPartialAllowed, TeamOpenSlotsVisible, TeamJoinRequiresApproval (v11.81) werden in die Event-Liste persistiert. Die UI gruppiert die sechs Settings in zwei Sub-Boxen: Basis (Toggle, Größe, Team-Name) + Beitritts-Modus (Radio komplett/teilweise + Sichtbarkeit + Approval). Der gesamte Schritt rendert direkt nach 'Kapazität & Sichtbarkeit' und vor 'Felder'. Die alten Steps Felder/Kommunikation/Dokumente/Fun-Zone rücken in der UI je um eins nach hinten — alle Step-Headlines, alle Tooltip-Verweise und das Handbuch sind entsprechend renumbered."
      />
      <Arrow />
      <FlowNode type="process" label="Nächste EventNumber ermitteln (max + 1)" />
      <Arrow />
      <FlowNode type="subprocess" label="Event-Bereich erstellen (URL aus Titel generiert)" />
      <Arrow />
      <FlowNode type="subprocess" label="Teilnehmerliste 'Teilnehmer' auf Event-Bereich erstellen (TeilnehmerID, Anrede, Vorname, Nachname, Status, ... + Zusatzfelder)" />
      <Arrow />
      <FlowNode type="process" label="Zugriffsschutz + Berechtigungen setzen (Owners=FullControl, Visitors=Contribute+ILS)" />
      <Arrow />
      <FlowNode type="subprocess" label="Event in die Event-Liste eintragen (mit EventNumber, Event-BereichUrl, DisableEmails, DisableOutlook, Zusatzfelder)" />
      <Arrow />
      <FlowNode type="decision" label="Bild oder Dokumente vorhanden?" />
      <BranchContainer>
        <Branch label="Ja">
          <FlowNode type="subprocess" label="Bild komprimieren + als __eventimage__-Attachment an die Event-Liste-Item anhängen" />
          <Arrow />
          <FlowNode type="subprocess" label="Dokumente als Attachments an die Event-Liste-Item anhängen" />
          <Arrow />
          <FlowNode type="process" label="EventImageUrl auf Attachment-URL patchen" />
        </Branch>
        <Branch label="Nein">
          <FlowNode type="process" color="#f5f5f5" label="Keine Anhänge" />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="data" label="die Mail-Warteschlange: Event-Erstellt Mail an alle Organizer (immer, unabhängig von DisableEmails)" />
      <Arrow />
      <FlowNode type="data" color="#fce4ec" label="Automatischer Hintergrund-Prozess: Outlook-Termin anlegen: Erstellt initialen Outlook-Termin im Kalender das zentrale Event-Postfach + speichert Termin-Verknüpfung zurück" />
      <Arrow />
      <FlowNode
        type="decision"
        label="Bearbeitung eines bestehenden Events?"
        details="Beim Edit prüft der Wizard pro Event (Hauptevent + jedes Sub-Event), ob Outlook-relevante Felder (Titel, Start, Ende, OutlookBody) gegenüber dem Mount-Snapshot verändert wurden. Für jedes betroffene Event mit Outlook-Termin erscheint im Confirm-Modal eine eigene Checkbox — der Organizer entscheidet pro Termin einzeln, ob die Teilnehmer eine 'Aktualisierter Termin'-Benachrichtigung bekommen sollen (v11.63)."
      />
      <BranchContainer>
        <Branch label="Pro Event: Haken gesetzt">
          <FlowNode
            type="data"
            color="#fce4ec"
            label="die Kalender-Warteschlange: UpdateEvent + OutlookDirty=false (nur dieses Event)"
            details="Für jedes angehakte Event (Hauptevent ODER Sub-Event): queueOutlookEvent(eventId, 'UpdateEvent') + updateEvent({OutlookDirty:false}). Der die Outlook-Automatik-Flow PATCHt den Termin im Shared-Mailbox-Kalender; Outlook schickt den Teilnehmern dieses Termins eine 'Aktualisierter Termin'-Mail. Andere Termine im selben Event-Verbund bleiben unangetastet."
          />
        </Branch>
        <Branch label="Pro Event: Haken aus">
          <FlowNode
            type="process"
            color="#fff8e1"
            label="updateEvent + OutlookDirty=true (nur dieses Event)"
            details="Für jedes nicht angehakte Event (das aber im Detect-Lauf als geändert erkannt wurde): Outlook-Termin wird NICHT angefasst. Damit beim nächsten Wizard-Lauf der Hinweis 'Outlook-Synchronisation steht aus' für dieses Event erscheint, schreibt die SPFx-App OutlookDirty=true zurück. Die Hinweisbox in Schritt 1 nennt Hauptevent und/oder die Anzahl der dirty-Sub-Events explizit."
          />
        </Branch>
        <Branch label="Nein (Create / keine Outlook-Änderung)">
          <FlowNode type="process" color="#f5f5f5" label="Direkter Save ohne Modal" />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="end" label="Event bereit für Registrierungen" />
    </div>
  );
}

function MassEmailFlow(): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <FlowNode type="start" label="Organizer/Admin klickt 'E-Mail an alle Teilnehmer' im Admin Center" />
      <Arrow />
      <FlowNode type="process" label="Modal öffnet sich: Betreff, Titel, RichText-Body eingeben" />
      <Arrow />
      <FlowNode type="process" label="Vorschau anzeigen (Body wird in Deloitte-Template gewrappt: Logo, grüner Header, Footer)" />
      <Arrow />
      <FlowNode type="decision" label="Senden bestätigen?" />
      <BranchContainer>
        <Branch label="Ja">
          <FlowNode type="process" label="Empfänger-Liste sammeln: alle aktiven Teilnehmer (Angemeldet, QR versendet, Eingecheckt)" />
        </Branch>
        <Branch label="Nein">
          <FlowNode type="end" color="var(--dex-orange)" label="Abbruch" />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="process" label="In Batches á max ~250 Zeichen Recipient-String aufteilen (semicolon-getrennt)" />
      <Arrow />
      <FlowNode type="data" label="Pro Batch: Ein die Mail-Warteschlange Eintrag mit Recipient='email1;email2;email3...' und EmailType='Massenmail'" />
      <Arrow />
      <FlowNode type="subprocess" color="#fce4ec" label="Automatischer Mail-Versand im Hintergrund: Lädt Logo + Default-Bild aus Config, ersetzt Platzhalter, sendet via Shared Mailbox das zentrale Event-Postfach" />
      <Arrow />
      <FlowNode type="end" label="Versand abgeschlossen, Status auf 'Sent'" />
    </div>
  );
}

function SelfCheckInFlow(): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <FlowNode type="start" label="Organizer öffnet eine Self-Check-in-Aktion (Check-in-Seite, Admin Center oder QR-Kachel unter dem Event-Bild)" details="Self-Check-in ist immer möglich — es gibt nichts zu konfigurieren (v20.2: der frühere Wizard-Schalter ist entfallen). Beim ersten Öffnen wird automatisch ein geheimer Token für das Event erzeugt (Schlüssel für den statischen Link UND für den rotierenden HMAC-QR-Code). Das Check-in-Zeitfenster (Von/Bis) ist im Kachel-Modal einstellbar — Standard: 2 Stunden vor Event-Start bis Event-Ende (keine verfrühten, keine nachträglichen Check-ins)." />
      <Arrow />
      <FlowNode type="decision" label="Welcher QR-Modus am Eingang?" />
      <BranchContainer>
        <Branch label="Statisches PDF">
          <FlowNode type="process" label="Admin Center → 'Self-Check-in: QR-PDF' → A4-PDF mit QR + Anleitung herunterladen, ausdrucken, aushängen" details="Der QR enthält den statischen Link ?action=selfcheckin&token=<Token>. Bequem, aber per Foto teilbar — daher mit Zeitfenster kombinieren." />
        </Branch>
        <Branch label="Rotierende Live-Anzeige">
          <FlowNode type="process" label="Admin Center → 'Self-Check-in: Live-Anzeige' → Bildschirm am Eingang zeigt QR, der alle ~45s automatisch wechselt" details="Der QR enthält ?action=selfcheckin&event=<Nr>&code=<HMAC>&t=<Fenster>. Ein abfotografierter Code verfällt nach max. 2 Fenstern — foto-sicher." />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="process" label="Teilnehmer scannt den QR-Code mit der NATIVEN Handy-Kamera" details="Bewusst kein In-App-Scanner: der native Kamera-Scan funktioniert zuverlässig auch in der Mobile-App, die getUserMedia im WebView blockiert. Der Link öffnet die DEX-App." />
      <Arrow />
      <FlowNode type="subprocess" label="App liest URL-Parameter (?action=selfcheckin…) und ruft EventContext.selfCheckIn() auf" />
      <Arrow />
      <FlowNode type="decision" label="Token gültig + (bei Live) Code frisch + im Zeitfenster?" />
      <BranchContainer>
        <Branch label="Nein">
          <FlowNode type="end" color="var(--dex-orange)" label="Ergebnis-Seite zeigt Grund: 'Code abgelaufen' / 'Check-in noch nicht offen' / 'Event nicht gefunden'" />
        </Branch>
        <Branch label="Ja">
          <FlowNode type="subprocess" label="Eigene Registrierung auf der Event-Bereich per E-Mail des eingeloggten Users finden" details="Jeder checkt ausschließlich sich selbst ein — Zugriffsschutz erlaubt nur das Schreiben des eigenen Eintrags." />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="decision" label="Status der eigenen Anmeldung?" />
      <BranchContainer>
        <Branch label="Angemeldet / QR versendet">
          <FlowNode type="data" label="Aktualisierung Status='Eingecheckt' + CheckedInDate/By — grüne Bestätigung 'Eingecheckt!'" />
        </Branch>
        <Branch label="Bereits eingecheckt">
          <FlowNode type="end" color="var(--dex-green)" label="Hinweis 'Bereits eingecheckt' — kein doppelter Counter" />
        </Branch>
        <Branch label="Warteliste / keine Anmeldung">
          <FlowNode type="end" color="var(--dex-orange)" label="Hinweis: kein Check-in möglich" />
        </Branch>
      </BranchContainer>
    </div>
  );
}

function IDReorderManualFlow(): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <FlowNode type="start" label="Admin klickt 'IDs neu vergeben' im Admin Center" />
      <Arrow />
      <FlowNode type="process" label="Bestätigung anzeigen (Sortierung nach SP-Item-ID = Erstellungsreihenfolge)" />
      <Arrow />
      <FlowNode type="subprocess" label="Alle Teilnehmer der Event-Bereich-Liste laden ($orderby=Id asc)" />
      <Arrow />
      <FlowNode type="process" label="Aktive (Angemeldet, QR versendet, Eingecheckt, Warteliste): TeilnehmerID = 1, 2, 3 ... N" />
      <Arrow />
      <FlowNode type="process" label="Inaktive (Abgemeldet): TeilnehmerID = null" />
      <Arrow />
      <FlowNode type="subprocess" label="Pro Item ein Aktualisierung auf TeilnehmerID (nur wenn sich die ID ändert)" />
      <Arrow />
      <FlowNode type="end" label="Erfolgs-Hinweis: 'X aktualisiert, Y Fehler' + Reload der Liste" />
    </div>
  );
}

function TeamJoinFlow(): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <FlowNode
        type="start"
        label="Beitritts-Wunsch ausgelöst (v11.83 / v11.84)"
        details={'Drei Einstiege: (1) User klickt auf der Event-Anmeldeseite einen \'Beitreten\'-/ \'Beitritt anfragen\'-Button in der \'Offene Teams\'-Box (User-Pfad, v11.83). (2) Team-Lead klickt in „Meine Events" auf „+ Mitglied hinzufügen" am Team-Badge (Lead-Pfad, v11.83). (3) Organizer/Admin klickt im Admin Center in der Teams-Sektion auf „Person hinzufügen" (Admin-Pfad, v11.84). Alle drei Pfade laufen anschliessend durch denselben Doppel-Anmelde-Schutz.'}
      />
      <Arrow />
      <FlowNode
        type="subprocess"
        label="Doppel-Anmelde-Schutz"
        details="isUserAlreadyOnEvent(subsiteUrl, email) prüft per OData-Filter (E-Mail + Status in Angemeldet/QR versendet/Eingecheckt/Warteliste). Treffer = Abbruch mit klarer Fehlermeldung."
      />
      <Arrow />
      <FlowNode
        type="decision"
        label="Welcher Pfad?"
        details="Selbst-Beitritt vs. Lead-Add vs. Admin-Center-Add. Bei Selbst-Beitritt: prüfe ob Approval erforderlich ist. Bei Lead-Add und Admin-Center-Add: kein Approval — die hinzufügende Person trägt die Verantwortung der Zustimmung (Pflicht-Checkbox)."
      />
      <BranchContainer>
        <Branch label="Lead-Add (MyEvents) oder Admin-Center-Add">
          <FlowNode
            type="subprocess"
            color="#e8f5e9"
            label="addTeamMember()-Pfad direkt"
            details="reserveSeat + registerTeamMember + Bestätigungs-Mail + Outlook + Info-Mail an alle Mitglieder. Kein Approval-Detour. Im Admin-Center-Pfad ist das die einzige Möglichkeit, im MyEvents-Pfad bedient sich nur der Lead daran (UI versteckt den Button für Nicht-Leads)."
          />
        </Branch>
        <Branch label="Selbst-Beitritt (Anmelde-Seite) — direkt (kein Approval)">
          <FlowNode
            type="subprocess"
            color="#e8f5e9"
            label="reserveSeat() atomisch"
            details="Split-aware: bei aktivem PreferredStarterType wird die entsprechende Gruppe inkrementiert. Bei Vollbelegung landet der User auf der Warteliste (kein Hard-Fail)."
          />
          <Arrow />
          <FlowNode
            type="subprocess"
            color="#e8f5e9"
            label="registerTeamMember() — neuer Eintrag"
            details="Inserts ein Teilnehmer-Item mit gleicher TeamId wie das offene Team, TeamLead=false, TeamName geerbt."
          />
          <Arrow />
          <FlowNode
            type="data"
            label="die Mail-Warteschlange + die Kalender-Warteschlange + Info-Mails"
            details="Bestätigungs-Mail an den Beitretenden, Outlook-Einladung in seine Queue, Info-Mail an die anderen aktiven Mitglieder '<Name> ist eurem Team beigetreten'."
          />
        </Branch>
        <Branch label="Selbst-Beitritt mit Approval-Queue">
          <FlowNode
            type="data"
            color="#ffebee"
            label="DEX_TeamJoinRequests: neuer Pending-Eintrag"
            details="Eine Zeile mit EventId, TeamId, RequesterEmail, RequesterDisplayName, Status='Pending'. Liegt global auf der Site-Collection-Ebene, nicht pro Event-Bereich."
          />
          <Arrow />
          <FlowNode
            type="data"
            label="Notification-Mail an den Team-Lead"
            details="Mail mit Approve-/Reject-Buttons (HTML-Links auf die App mit Query-Parametern ?action=teamjoin&request=<id>&decision=approve|reject). Die App-UI in MyEvents zeigt aber sowieso den 'Beitritts-Anfragen (N)'-Block."
          />
          <Arrow />
          <FlowNode
            type="decision"
            label="Team-Lead entscheidet (in MyEvents)"
            details="Im 'Beitritts-Anfragen'-Block klickt der Lead pro Anfrage auf 'Bestätigen' oder 'Ablehnen'."
          />
          <BranchContainer>
            <Branch label="Bestätigen">
              <FlowNode
                type="subprocess"
                color="#e8f5e9"
                label="addTeamMember()-Pfad"
                details="reserveSeat + registerTeamMember + Bestätigungs-Mail + Outlook + Info-Mails an alle, identisch zum Direkt-Beitritts-Pfad."
              />
              <Arrow />
              <FlowNode
                type="data"
                label="DEX_TeamJoinRequests: Status='Approved'"
                details="DecidedDate + DecidedByEmail werden gesetzt — Audit-Spur, wer wann genehmigt hat."
              />
            </Branch>
            <Branch label="Ablehnen">
              <FlowNode
                type="data"
                color="#ffebee"
                label="DEX_TeamJoinRequests: Status='Rejected'"
                details="Plus kurze Absage-Mail im Deloitte-Layout an den Anfragenden: 'Deine Anfrage wurde abgelehnt — du kannst dich gerne einzeln anmelden, falls die Kapazität reicht.'"
              />
            </Branch>
          </BranchContainer>
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode
        type="end"
        label="Fertig"
        details="Im Direkt-Pfad sieht der User sofort die Erfolgs-Box auf der Anmeldeseite und das Event in 'Meine Events'. Im Approval-Pfad bekommt er eine Mail mit Ergebnis, sobald der Lead entschieden hat."
      />
    </div>
  );
}

function ColumnFixFlow(): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <FlowNode type="start" label="Admin klickt 'Spalten fixen' im Admin Center" />
      <Arrow />
      <FlowNode type="subprocess" label="Bestehende Felder der Teilnehmerliste laden" />
      <Arrow />
      <FlowNode type="decision" label="Pflicht-Spalten fehlen? (z.B. Anrede)" />
      <BranchContainer>
        <Branch label="Ja">
          <FlowNode type="subprocess" label="Fehlende Spalten anlegen (Anrede als Choice Frau/Herr/Divers)" />
        </Branch>
        <Branch label="Nein">
          <FlowNode type="process" color="#f5f5f5" label="Keine neuen Spalten nötig" />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="subprocess" label="Default-View: alle View-Felder entfernen, dann in korrekter Reihenfolge wieder hinzufügen (TeilnehmerID > Anrede > Vorname > Nachname > Email > ... > Zusatzfelder)" />
      <Arrow />
      <FlowNode type="end" label="Erfolgs-Hinweis: 'Spalten hinzugefügt: X | View-Reihenfolge korrigiert'" />
    </div>
  );
}

// v23.41: Archivierung und Löschung — bewusst nicht-technisch erklärt.
function ArchiveDeleteFlow(): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <FlowNode type="start" label="Ein Admin öffnet die App"
        details="Im Hintergrund laufen viele Arbeitslisten voll (verschickte Mails, Kalender-Aufgaben, Protokolle). Damit sie nicht endlos wachsen, gibt es zwei Aufräum-Stufen: Archivieren und danach Löschen." />
      <Arrow />
      <FlowNode type="decision" label="Gibt es alte Einträge zum Aufräumen?"
        details="Beim App-Start zählt die App im Hintergrund, wie viele Zeilen aufgeräumt werden können. Nur dann erscheint oben rechts (und in der Admin-Kachel) ein Hinweis." />
      <BranchContainer>
        <Branch label="Nein">
          <FlowNode type="process" color="#f5f5f5" label="Nichts zu tun — kein Hinweis" />
        </Branch>
        <Branch label="Ja">
          <FlowNode type="process" label="Admin sieht den Aufräum-Hinweis (mit Anzahl)"
            details="Zwei getrennte Aktionen, jeweils mit Anzahl: 1) Archivieren (verschieben), 2) altes Archiv löschen." />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="process" label="STUFE 1 — Archivieren"
        details="Was passiert: Zeilen von Events, die schon vorbei ODER bereits gelöscht sind, werden aus den Arbeitslisten in eine eigene Archiv-Ablage verschoben. Wichtig: Noch nicht verschickte Mails (Status offen) bleiben unangetastet." />
      <Arrow />
      <FlowNode type="data" label="Verschieben statt kopieren — Schritt für Schritt"
        details="Jede Zeile wird zuerst sicher ins Archiv geschrieben und ERST DANN aus der Arbeitsliste entfernt. So kann nichts verloren gehen. Ein Fortschrittsbalken zeigt den Stand; abbrechen ist jederzeit möglich (schon Verschobenes bleibt im Archiv)." />
      <Arrow />
      <FlowNode type="end" color="var(--dex-green)" label="Arbeitslisten sind wieder schlank, alte Einträge liegen im Archiv" />
      <Arrow />
      <FlowNode type="process" label="STUFE 2 — Altes Archiv löschen"
        details="Das Archiv ist die letzte Ablage. Rund einen Monat nach Ablauf des Events braucht die Daten in der Regel niemand mehr — sie können endgültig entfernt werden." />
      <Arrow />
      <FlowNode type="decision" label="Einträge älter als 1 Monat vorhanden?" />
      <BranchContainer>
        <Branch label="Nein">
          <FlowNode type="process" color="#f5f5f5" label="Archiv ist aktuell — nichts löschen" />
        </Branch>
        <Branch label="Ja">
          <FlowNode type="process" label="Admin bestätigt das Löschen (Sicherheitsabfrage)"
            details="Weil Löschen endgültig ist, kommt eine rot markierte Sicherheitsabfrage mit der genauen Anzahl." />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="end" color="var(--dex-green)" label="Alte Archiv-Einträge sind entfernt — das Archiv bleibt dauerhaft handhabbar" />
    </div>
  );
}

// ==================== Hauptkomponente ====================

export default function FlowchartPage(): React.ReactElement {
  const [activeFlow, setActiveFlow] = React.useState<string>('registration');

  const flows = [
    { id: 'registration', label: 'Anmeldung', icon: '→' },
    { id: 'cancellation', label: 'Abmeldung', icon: '←' },
    { id: 'reorder', label: 'ID-Korrektur & Nachrücken', icon: '↻' },
    { id: 'creation', label: 'Event-Erstellung', icon: '+' },
    { id: 'massemail', label: 'Massenmail', icon: '✉' },
    { id: 'teamjoin', label: 'Team-Beitritt', icon: '+' },
    { id: 'selfcheckin', label: 'Self-Check-in', icon: '✓' },
    { id: 'idmanual', label: 'IDs neu vergeben (Admin)', icon: '#' },
    { id: 'columnfix', label: 'Spalten fixen (Admin)', icon: '⚙' },
    { id: 'archive', label: 'Archivierung und Löschung', icon: '🗄' },
  ];

  const renderFlow = (): React.ReactElement => {
    switch (activeFlow) {
      case 'registration': return <RegistrationFlow />;
      case 'cancellation': return <CancellationFlow />;
      case 'reorder': return <IDReorderFlow />;
      case 'creation': return <EventCreationFlow />;
      case 'massemail': return <MassEmailFlow />;
      case 'teamjoin': return <TeamJoinFlow />;
      case 'selfcheckin': return <SelfCheckInFlow />;
      case 'idmanual': return <IDReorderManualFlow />;
      case 'columnfix': return <ColumnFixFlow />;
      case 'archive': return <ArchiveDeleteFlow />;
      default: return <RegistrationFlow />;
    }
  };

  return (
    <div className="page-container">
      <h2 style={{ marginBottom: 8 }}>Prozess-Übersicht</h2>
      <p style={{ color: 'var(--dex-gray-600)', fontSize: '0.88rem', marginBottom: 16, lineHeight: 1.5, maxWidth: 760 }}>
        Hier siehst du Schritt für Schritt, was bei den wichtigsten Abläufen passiert — verständlich erklärt, ohne Technik-Begriffe.
        Wähle oben einen Ablauf und fahre mit der Maus über einen Schritt, um mehr zu erfahren.
      </p>

      {/* Tab-Navigation */}
      <div className="flowchart-tabs" style={{ display: 'flex', gap: 8, marginBottom: 24, flexWrap: 'wrap' }}>
        {flows.map(f => (
          <button
            key={f.id}
            className={activeFlow === f.id ? 'btn btn-primary' : 'btn btn-secondary'}
            onClick={() => setActiveFlow(f.id)}
            style={{ fontSize: '0.85rem' }}
          >
            <span style={{ marginRight: 6 }}>{f.icon}</span>
            {f.label}
          </button>
        ))}
      </div>

      {/* Legende — BPMN-Formen (Kreis = Ereignis, Rechteck = Aufgabe, Raute = Entscheidung). */}
      <div className="card" style={{ padding: 12, marginBottom: 16 }}>
        <p style={{ margin: '0 0 8px', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
          Die Diagramme folgen der BPMN-Notation: <strong>Kreis</strong> = Start/Ende, <strong>abgerundetes Rechteck</strong> = Aufgabe/Schritt, <strong>Raute</strong> = Entscheidung, <strong>Pfeil</strong> = Reihenfolge (von oben nach unten).
        </p>
        <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap', fontSize: '0.75rem', alignItems: 'center' }}>
          <span><span style={{ display: 'inline-block', width: 14, height: 14, borderRadius: '50%', background: 'var(--dex-green)', marginRight: 5, verticalAlign: 'middle' }} /> Start</span>
          <span><span style={{ display: 'inline-block', width: 14, height: 14, borderRadius: '50%', background: '#fff', border: '3px solid var(--dex-red, #c00)', marginRight: 5, verticalAlign: 'middle', boxSizing: 'border-box' }} /> Ende</span>
          <span><span style={{ display: 'inline-block', width: 14, height: 14, borderRadius: 3, background: '#e3f2fd', border: '1px solid #90caf9', marginRight: 5, verticalAlign: 'middle' }} /> Schritt in der App</span>
          <span><span style={{ display: 'inline-block', width: 14, height: 14, borderRadius: 3, background: '#e8f5e9', border: '1px solid #a5d6a7', borderLeft: '3px solid #66bb6a', marginRight: 5, verticalAlign: 'middle' }} /> Wird gespeichert</span>
          <span><span style={{ display: 'inline-block', width: 12, height: 12, borderRadius: 2, background: '#fff3e0', border: '1px solid #ffcc80', marginRight: 8, verticalAlign: 'middle', transform: 'rotate(45deg)' }} /> Entscheidung</span>
          <span><span style={{ display: 'inline-block', width: 14, height: 14, borderRadius: 3, background: '#f3e5f5', border: '1px solid #ce93d8', marginRight: 5, verticalAlign: 'middle' }} /> Läuft automatisch im Hintergrund</span>
        </div>
      </div>

      {/* Flowchart */}
      <div className="card" style={{ padding: 24, overflowX: 'auto' }}>
        {renderFlow()}
      </div>
    </div>
  );
}
