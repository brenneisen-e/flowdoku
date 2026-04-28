/**
 * Flowchart-Seite - Visualisierung der Prozesse
 *
 * Zeigt die einzelnen Ablaeufe (Anmeldung, Abmeldung, ID-Reorder, Nachrücken)
 * als verstaendliche Flowcharts fuer Dritte.
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

  if (!details) {
    return <div style={styles[type] || styles.process}>{label}</div>;
  }
  // Mit Prosa-Details: Knoten bündig zentriert, darunter linksbündiger Info-Text in grau.
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', maxWidth: 540, margin: '0 auto' }}>
      <div style={styles[type] || styles.process}>{label}</div>
      <div style={{
        marginTop: 6,
        padding: '8px 14px',
        fontSize: '0.78rem',
        lineHeight: 1.55,
        color: 'var(--dex-gray-600)',
        background: 'var(--dex-gray-50, #fafafa)',
        border: '1px solid var(--dex-gray-200)',
        borderRadius: 8,
        textAlign: 'left',
        maxWidth: 520,
      }}>
        {details}
      </div>
    </div>
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
        details="Im Registrierungs-Formular trägt der User die Stamm-Daten ein (Anrede, Vorname, Nachname, E-Mail), plus event-spezifische Felder (Custom Fields wie Travel-Zustimmung, Zimmerpartner-Picker, Reisekosten etc.). Bei B2Run-Events mit Split-Kapazitäten wählt er zusätzlich den Wunsch-Starter-Typ. Name und E-Mail werden aus dem SPFx-Profil vorausgefüllt, können aber überschrieben werden (nur mit Berechtigung für 'andere Person anmelden')."
      />
      <Arrow />
      <FlowNode
        type="decision"
        label="Bereits registriert für dieses Event?"
        details="Der Service prüft per getMyRegistration in der Subsite-Teilnehmer-Liste, ob schon ein Eintrag zu dieser E-Mail existiert. Wird verhindern dass ein User doppelt gezählt wird — und gleichzeitig erlauben, dass jemand, der sich zuvor abgemeldet hat, sich wieder anmelden kann."
      />
      <BranchContainer>
        <Branch label="Ja — zuvor abgemeldet">
          <FlowNode
            type="process"
            label="Reaktivierung"
            details="Der bestehende Abgemeldet-Eintrag wird wiederverwendet statt neu angelegt: Status zurück auf 'Angemeldet' oder 'Warteliste' (je nach Kapazität), CancellationDate gelöscht, RegistrationDate auf jetzt gesetzt. Die TeilnehmerID bleibt zunächst null — sie wird beim nächsten ID-Reorder-Run vergeben. So bleibt die Änderungs-Historie (ChangeLog) erhalten."
          />
        </Branch>
        <Branch label="Ja — aktiv">
          <FlowNode
            type="end"
            color="var(--dex-red)"
            label="Fehler: Bereits angemeldet"
            details="Die Anmeldung bricht ab und der User sieht eine Fehlermeldung. So wird Doppel-Zählung in DEX_Participants verhindert. Wenn der User sich von der Warteliste auf 'aktiv' promoten lassen möchte, muss er sich zuerst abmelden und neu anmelden."
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
        label="Atomare TeilnehmerID-Vergabe (v7.28+, nur bei Erst-Anmeldung)"
        details="Bei einer ECHT neuen Anmeldung (kein bestehendes Abgemeldet-Item) holt der Service die nächste TeilnehmerID atomar aus der Counter-Liste DEX_TeilnehmerCounter (eine Liste pro Subsite mit genau einem Item, NextValue=N). Vorgehen: GET counter-item mit ETag → PATCH NextValue=N+1 mit IF-MATCH=etag → bei HTTP 412 (jemand war schneller) wird mit kurzem Jitter retried (max 8x). Damit können mehrere User parallel anmelden, ohne dass IDs doppelt vergeben werden. Bei Reaktivierung eines Abgemeldet-Items wird der Counter NICHT angefasst — die TeilnehmerID bleibt null und der DEX_IDReorder-Flow vergibt sie sequentiell zurück. v7.31: Nach Cancel und nach 'IDs neu vergeben' wird der Counter auf den aktuellen Max-Wert gesynct (syncCounterToMax), damit er nicht über die echten IDs hinaus 'davonrast'. Empfohlen ist zusätzlich, im DEX_IDReorder_TeilnehmerIDs-Flow am Ende den Counter ebenfalls zu setzen (Patch DEX_TeilnehmerCounter/items(1) mit NextValue=max). Fallback bei alten Events ohne Counter-Liste: max(TeilnehmerID)+1 — dann muss der Admin einmalig 'Spalten fixen' klicken, damit die Counter-Liste angelegt + mit dem aktuellen Max-Wert geseedet wird."
      />
      <Arrow />
      <FlowNode
        type="subprocess"
        label="Eintrag in Subsite-Teilnehmerliste"
        details="Der Service schreibt ein neues Item in die 'Teilnehmer'-Liste der Event-Subsite: TeilnehmerID (eindeutig vom Counter, oder null bei Reaktivierung), Anrede, Vorname, Nachname, ParticipantEmail, Status, StarterType, PreferredStarterType, RegistrationDate, RegisteredByName/Email (Audit), CustomData (JSON der event-spezifischen Felder). Bei Late-Cancel-Reaktivierungen oder Storno-Nachrückern (siehe DEX_IDReorder-Flow) wird die TeilnehmerID nachträglich vom Flow vergeben, um Lücken zu schließen."
      />
      <Arrow />
      <FlowNode
        type="subprocess"
        label="DEX_Participants aktualisieren"
        details="In der zentralen DEX_Participants-Liste (auf Site-Ebene, listet alle User und ihre Event-Registrierungen) wird die EventNumber in das Feld EventRegistered (Angemeldet) bzw. EventOnWaitlist (Warteliste) eingetragen. Diese Liste ist für den schnellen My-Events-Pfad — damit muss nicht jede Event-Subsite einzeln nach dem User durchsucht werden."
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
            label="DEX_Emails: Bestätigungs-Mail in Queue"
            details="Das Template-Type je nach Status: 'Anmeldung' oder 'Warteliste'. Das SP-Template wird aus DEX_EmailTemplates geladen (pro EmailLanguage DE/EN), Platzhalter {{Name}}, {{EventTitle}}, {{Organizer}}, {{WaitlistPosition}} werden aufgelöst. Event-spezifische Template-Overrides (z.B. eigener Anmelde-Text pro Event) werden vorher angewendet. Der Body wird im Deloitte-Layout gewrappt und in DEX_Emails mit Status='Pending' geschrieben. Der DEX_SEND_MAIL-Flow versendet innerhalb von ~1 Minute."
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
            label="DEX_Outlook: 'Einladen' in Queue"
            details="Ein Item mit ActionType='Einladen', EventId und Attendee (E-Mail) wird in DEX_Outlook gepushed. Der DEX_Outlook_Einladungen-Flow liest den Parent-Termin aus dem Shared Mailbox (no_reply.events@deloitte.de) über iCalUId (CalendarLink im Event), fügt den Attendee per Graph PATCH hinzu und Outlook sendet dem User die Einladung."
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
    </div>
  );
}

function CancellationFlow(): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <FlowNode
        type="start"
        label="Abmeldung ausgelöst"
        details="Einstiegspunkt. Die Abmeldung kann von zwei Seiten kommen: der Teilnehmer selbst (Self-Cancel aus My-Events) oder ein Admin/Organizer (aus dem Admin-Center). Beide Pfade sind fast identisch, unterscheiden sich aber im UX und beim Nachrücker-Timing."
      />
      <Arrow />
      <FlowNode
        type="decision"
        label="Wer meldet ab?"
        details="Die App unterscheidet den Pfad, weil der Admin sofortiges Feedback über den Nachrücker braucht (damit er den nächsten Schritt planen kann), während ein einzelner User nur seinen eigenen Eintrag sieht und die kurze Verzögerung durch den Flow nicht wahrnimmt."
      />
      <BranchContainer>
        <Branch label="User (Self-Cancel)">
          <FlowNode
            type="process"
            label="User klickt in My-Events auf 'Abmelden'"
            details="Zwei-Klick-Bestätigung (erster Klick zeigt 'Wirklich abmelden?', zweiter führt aus). Das verhindert versehentliche Abmeldungen. Item-Level-Security sorgt dafür dass der User eh nur seinen Eintrag sieht — kein Info-Toast über den Nachrücker nötig."
          />
        </Branch>
        <Branch label="Admin/Organizer">
          <FlowNode
            type="process"
            color="#fff3e0"
            label="Admin klickt im Admin-Center auf 'Abmelden'"
            details="Single-Click mit Confirm-Dialog. Nach Bestätigung erscheint sofort ein orange Toast oben rechts mit Spinner: 'Abmeldung von X wird verarbeitet…'. Nachrücker wird direkt client-seitig promoted (siehe unten). Admin sieht am Ende den Erfolgs-Toast mit Namen des Nachrückers."
          />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode
        type="subprocess"
        label="Teilnehmerliste: Status → 'Abgemeldet'"
        details="Per MERGE-Request auf das Teilnehmer-Item: Status wird auf 'Abgemeldet' gesetzt, CancellationDate = jetzt, CancelledByName/Email werden zur Audit-Nachverfolgung gespeichert (wer hat die Abmeldung ausgelöst). TeilnehmerID bleibt vorerst — der IDReorder-Flow setzt sie später auf null für diesen Abgemeldeten."
      />
      <Arrow />
      <FlowNode
        type="subprocess"
        label="DEX_Participants aktualisieren"
        details="Die zentrale DEX_Participants-Liste enthält pro User die Liste der EventNumbers, für die er angemeldet oder auf der Warteliste ist. Bei Abmeldung wird die EventNumber aus EventRegistered bzw. EventOnWaitlist entfernt. So weiß die My-Events-Seite auf einen Blick, dass der User für dieses Event nicht mehr aktiv ist."
      />
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
            label="DEX_Emails: Abmelde-Mail in Queue"
            details="Template-Type 'Abmeldung' aus DEX_EmailTemplates wird geladen (pro EmailLanguage), Platzhalter {{Name}}, {{EventTitle}} aufgelöst, in Deloitte-Layout gewrappt. Status='Pending', DEX_SEND_MAIL-Flow verschickt binnen ~1 Minute. Empfänger ist der abgemeldete Teilnehmer selbst — als Bestätigung."
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
            label="DEX_Outlook: 'Ausladen' in Queue"
            details="Ein DEX_Outlook-Item mit ActionType='Ausladen', EventId und Attendee-Email wird gepushed. Der DEX_Outlook_Einladungen-Flow liest den Kalender-Termin (iCalUId aus dem Event), filtert den Attendee raus und PATCHt den Termin mit der neuen Attendee-Liste. Outlook benachrichtigt den User, dass sein Termin zurückgezogen wurde."
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
            details="Für den Nachrücker wird ein DEX_Emails-Item mit Template-Type 'Nachruecken' und ein DEX_Outlook-Item mit ActionType='Einladen' in die Queues geschrieben. Der Nachrücker erhält dadurch automatisch eine freundliche Nachrück-Mail und eine Kalender-Einladung — beides im Deloitte-Layout. Der Admin sieht im Toast: 'Nachgerückt: <Name> <E-Mail> (<Starter-Typ>)'."
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
        label="DEX_IDReorder: Reorder-Auftrag in Queue"
        details="Ein neues Item in der DEX_IDReorder-Liste (Pending) löst den DEX_IDReorder_TeilnehmerIDs-Flow aus. Der Flow läuft sequenziell (Concurrency=1) — bei mehreren Abmeldungen werden die Einträge nacheinander abgearbeitet. Er sortiert die TeilnehmerIDs neu (Zwei-Pass: Angemeldete zuerst, Warteliste danach) und übernimmt bei User-Self-Cancels auch das Nachrücken."
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
        label="Power Automate Trigger"
        details="Der DEX_IDReorder_TeilnehmerIDs-Flow pollt die DEX_IDReorder-Liste im Minuten-Takt. Bei jedem neuen Eintrag mit Status='Pending' startet eine Flow-Instanz. Concurrency=1 stellt sicher, dass immer nur eine Instanz gleichzeitig läuft — bei mehreren kurz aufeinanderfolgenden Abmeldungen werden die Einträge sequenziell (in Reihenfolge der Erstellung) abgearbeitet. Maximal 100 Einträge können warten."
      />
      <Arrow />
      <FlowNode
        type="process"
        label="Status → 'Processing'"
        details="Sofort beim Start: das DEX_IDReorder-Item wird auf 'Processing' gesetzt. Damit sieht der nächste Polling-Durchlauf das Item nicht mehr als 'Pending' und startet keinen Parallel-Run auf dem gleichen Item."
      />
      <Arrow />
      <FlowNode
        type="subprocess"
        label="Get_Enrolled_Participants"
        details="Aus der 'Teilnehmer'-Liste der Event-Subsite werden alle Items geladen, deren Status NICHT 'Abgemeldet' ist. Das ergibt die zusammengefasste Liste von Angemeldeten (Status='Angemeldet' / 'QR versendet' / 'Eingecheckt') und Warteliste-Teilnehmern ('Warteliste'). Die Sortierung ist RegistrationDate asc — ältere Registrierungen zuerst."
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
        label="$batch-Update an SharePoint"
        details="Per OData-$batch-API werden alle TeilnehmerIDs in einem einzigen Request-Paket (Chunks à 250) aktualisiert. Das ist deutlich schneller als einzelne PATCH-Calls pro Item (bei 100 Teilnehmern: 1 Call statt 100). WICHTIG: nur das Feld TeilnehmerID wird geschrieben — Status, Vorname, Nachname usw. bleiben unverändert. Warteliste-Items behalten ihren Warteliste-Status, nur die TID ändert sich."
      />
      <Arrow />
      <FlowNode
        type="subprocess"
        label="Get_EventDetails"
        details="Nach dem Reorder werden die Event-Metadaten aus DEX_Events geladen (Title, MaxParticipants, EmailLanguage und — relevant für B2Run-Split — DurchstarterCapacity und FunstarterCapacity). Diese Daten bestimmen, wie der Nachrück-Prozess weiterläuft."
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
            details="Count_Active_Durchstarter zählt aktive Teilnehmer mit StarterType='Durchstarter'. Wenn diese Zahl < DurchstarterCapacity ist, holt der Flow per SharePoint GET den ersten Warteliste-Eintrag mit PreferredStarterType='Durchstarter' (sortiert nach RegistrationDate). Existiert einer → MERGE: Status=Angemeldet, StarterType=Durchstarter. Nachrück-Mail (Template 'Nachruecken') + Outlook-'Einladen' werden in die Queues geschrieben."
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
            details="Der erste Warteliste-Eintrag (älteste RegistrationDate) wird per MERGE auf Status='Angemeldet' gesetzt. KEINE TeilnehmerID-Änderung — die Zwei-Pass-Sortierung hat die korrekte TID bereits zugewiesen. Anschließend: Nachrück-Mail + Outlook-Einladung in Queue."
          />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode
        type="data"
        label="DEX_Emails: Nachrück-Mail in Queue"
        details="Template aus DEX_EmailTemplates (TemplateType='Nachruecken', Language passend zum Event), Platzhalter {{Name}} und {{EventTitle}} werden aufgelöst. Status='Pending' — DEX_SEND_MAIL-Flow versendet die Mail im nächsten Durchlauf (~1 Min)."
      />
      <Arrow />
      <FlowNode
        type="data"
        label="DEX_Outlook: Einladen in Queue"
        details="DEX_Outlook-Item mit ActionType='Einladen' und Attendee = E-Mail des Nachrückers. Der DEX_Outlook_Einladungen-Flow fügt den Attendee per Graph PATCH zum Shared-Mailbox-Termin hinzu — Outlook schickt automatisch die Kalender-Einladung."
      />
      <Arrow />
      <FlowNode
        type="process"
        label="Status → 'Done'"
        details="Das DEX_IDReorder-Item wird final auf 'Done' gesetzt. Falls im Laufe des Flows ein Fehler auftritt (Graph-Timeout, 429 Rate-Limit, 404), führt der Error-Handler den Status stattdessen auf 'Failed' — so kann der Admin im Admin-Center oder direkt in der SP-Liste prüfen, welche Runs erfolglos waren."
      />
      <Arrow />
      <FlowNode
        type="end"
        color="#6a1b9a"
        label="Flow-Instanz beendet"
        details="Bei weiteren DEX_IDReorder-Einträgen startet die nächste Instanz. Concurrency=1 garantiert sequentielle Abarbeitung — keine Race-Conditions bei paralleler Bearbeitung."
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
      <FlowNode type="start" label="Organizer/Admin erstellt Event (7 Schritte: Grundlagen → Zeit&Ort → Kapazität → Felder → Kommunikation → Dokumente → Quiz)" />
      <Arrow />
      <FlowNode type="process" label="Nächste EventNumber ermitteln (max + 1)" />
      <Arrow />
      <FlowNode type="subprocess" label="SharePoint Subsite erstellen (URL aus Titel generiert)" />
      <Arrow />
      <FlowNode type="subprocess" label="Teilnehmerliste 'Teilnehmer' auf Subsite erstellen (TeilnehmerID, Anrede, Vorname, Nachname, Status, ... + Custom Fields)" />
      <Arrow />
      <FlowNode type="process" label="Item-Level Security + Berechtigungen setzen (Owners=FullControl, Visitors=Contribute+ILS)" />
      <Arrow />
      <FlowNode type="subprocess" label="Event in DEX_Events eintragen (mit EventNumber, SubsiteUrl, DisableEmails, DisableOutlook, CustomFields)" />
      <Arrow />
      <FlowNode type="decision" label="Bild oder Dokumente vorhanden?" />
      <BranchContainer>
        <Branch label="Ja">
          <FlowNode type="subprocess" label="Bild komprimieren + als __eventimage__-Attachment an DEX_Events-Item anhängen" />
          <Arrow />
          <FlowNode type="subprocess" label="Dokumente als Attachments an DEX_Events-Item anhängen" />
          <Arrow />
          <FlowNode type="process" label="EventImageUrl auf Attachment-URL patchen" />
        </Branch>
        <Branch label="Nein">
          <FlowNode type="process" color="#f5f5f5" label="Keine Anhänge" />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="data" label="DEX_Emails: Event-Erstellt Mail an alle Organizer (immer, unabhängig von DisableEmails)" />
      <Arrow />
      <FlowNode type="data" color="#fce4ec" label="Power Automate Trigger DEX_CreateOutlookEvent: Erstellt initialen Outlook-Termin im Kalender no_reply.events@deloitte.de + speichert CalendarLink (iCalUId) zurück" />
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
      <FlowNode type="data" label="Pro Batch: Ein DEX_Emails Eintrag mit Recipient='email1;email2;email3...' und EmailType='Massenmail'" />
      <Arrow />
      <FlowNode type="subprocess" color="#fce4ec" label="Power Automate Trigger DEX_SEND_MAIL: Lädt Logo + Default-Bild aus Config, ersetzt Platzhalter, sendet via Shared Mailbox no_reply.events@deloitte.de" />
      <Arrow />
      <FlowNode type="end" label="Versand abgeschlossen, Status auf 'Sent'" />
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
      <FlowNode type="subprocess" label="Alle Teilnehmer der Subsite-Liste laden ($orderby=Id asc)" />
      <Arrow />
      <FlowNode type="process" label="Aktive (Angemeldet, QR versendet, Eingecheckt, Warteliste): TeilnehmerID = 1, 2, 3 ... N" />
      <Arrow />
      <FlowNode type="process" label="Inaktive (Abgemeldet): TeilnehmerID = null" />
      <Arrow />
      <FlowNode type="subprocess" label="Pro Item ein MERGE-Update auf TeilnehmerID (nur wenn sich die ID ändert)" />
      <Arrow />
      <FlowNode type="end" label="Erfolgs-Hinweis: 'X aktualisiert, Y Fehler' + Reload der Liste" />
    </div>
  );
}

function ColumnFixFlow(): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <FlowNode type="start" label="Admin klickt 'Spalten fixen' im Admin Center" />
      <Arrow />
      <FlowNode type="subprocess" label="Bestehende Felder der Subsite-Teilnehmerliste laden" />
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
      <FlowNode type="subprocess" label="Default-View: alle View-Felder entfernen, dann in korrekter Reihenfolge wieder hinzufügen (TeilnehmerID > Anrede > Vorname > Nachname > Email > ... > Custom Fields)" />
      <Arrow />
      <FlowNode type="end" label="Erfolgs-Hinweis: 'Spalten hinzugefügt: X | View-Reihenfolge korrigiert'" />
    </div>
  );
}

// ==================== Hauptkomponente ====================

export default function FlowchartPage(): React.ReactElement {
  const [activeFlow, setActiveFlow] = React.useState<string>('registration');

  const flows = [
    { id: 'registration', label: 'Anmeldung', icon: '→' },
    { id: 'cancellation', label: 'Abmeldung', icon: '←' },
    { id: 'reorder', label: 'ID-Korrektur (Power Automate)', icon: '↻' },
    { id: 'creation', label: 'Event-Erstellung', icon: '+' },
    { id: 'massemail', label: 'Massenmail', icon: '✉' },
    { id: 'idmanual', label: 'IDs neu vergeben (Admin)', icon: '#' },
    { id: 'columnfix', label: 'Spalten fixen (Admin)', icon: '⚙' },
  ];

  const renderFlow = (): React.ReactElement => {
    switch (activeFlow) {
      case 'registration': return <RegistrationFlow />;
      case 'cancellation': return <CancellationFlow />;
      case 'reorder': return <IDReorderFlow />;
      case 'creation': return <EventCreationFlow />;
      case 'massemail': return <MassEmailFlow />;
      case 'idmanual': return <IDReorderManualFlow />;
      case 'columnfix': return <ColumnFixFlow />;
      default: return <RegistrationFlow />;
    }
  };

  return (
    <div className="page-container">
      <h2 style={{ marginBottom: 8 }}>Prozess-Übersicht</h2>
      <p style={{ color: 'var(--dex-gray-500)', fontSize: '0.85rem', marginBottom: 16 }}>
        Visualisierung der Abläufe in der DEX Event Experience Platform.
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

      {/* Legende */}
      <div className="card" style={{ padding: 12, marginBottom: 16 }}>
        <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap', fontSize: '0.75rem' }}>
          <span><span style={{ display: 'inline-block', width: 12, height: 12, borderRadius: 12, background: 'var(--dex-green)', marginRight: 4, verticalAlign: 'middle' }} /> Start / Ende</span>
          <span><span style={{ display: 'inline-block', width: 12, height: 12, borderRadius: 3, background: '#e3f2fd', border: '1px solid #90caf9', marginRight: 4, verticalAlign: 'middle' }} /> App-Aktion</span>
          <span><span style={{ display: 'inline-block', width: 12, height: 12, borderRadius: 3, background: '#e8f5e9', border: '1px solid #a5d6a7', borderLeft: '3px solid #66bb6a', marginRight: 4, verticalAlign: 'middle' }} /> SharePoint-Schreibvorgang</span>
          <span><span style={{ display: 'inline-block', width: 12, height: 12, borderRadius: 3, background: '#fff3e0', border: '1px solid #ffcc80', marginRight: 4, verticalAlign: 'middle' }} /> Entscheidung</span>
          <span><span style={{ display: 'inline-block', width: 12, height: 12, borderRadius: 3, background: '#f3e5f5', border: '1px solid #ce93d8', marginRight: 4, verticalAlign: 'middle' }} /> Queue / Async</span>
        </div>
      </div>

      {/* Flowchart */}
      <div className="card" style={{ padding: 24, overflowX: 'auto' }}>
        {renderFlow()}
      </div>
    </div>
  );
}
