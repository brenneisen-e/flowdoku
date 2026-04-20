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
}

function FlowNode({ type, label, color }: FlowNodeProps): React.ReactElement {
  const baseStyle: React.CSSProperties = {
    padding: '10px 16px',
    fontSize: '0.8rem',
    textAlign: 'center',
    lineHeight: 1.3,
    fontWeight: 500,
    maxWidth: 220,
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

  return <div style={styles[type] || styles.process}>{label}</div>;
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
      <FlowNode type="start" label="User klickt 'Registrieren'" />
      <Arrow />
      <FlowNode type="process" label="Vorname, Nachname, Anrede, E-Mail, Event-spezifische Felder eingeben" />
      <Arrow />
      <FlowNode type="decision" label="Bereits registriert?" />
      <BranchContainer>
        <Branch label="Ja (abgemeldet)">
          <FlowNode type="process" label="Registrierung reaktivieren (Status zurücksetzen)" />
        </Branch>
        <Branch label="Ja (aktiv)">
          <FlowNode type="end" color="var(--dex-red)" label="Fehler: Bereits angemeldet" />
        </Branch>
        <Branch label="Nein">
          <FlowNode type="decision" label="Plätze frei?" />
          <Arrow />
          <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)', textAlign: 'center', marginBottom: 4 }}>
            Ja → Angemeldet / Nein → Warteliste
          </div>
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="subprocess" label="Eintrag in Subsite-Teilnehmerliste (TeilnehmerID, Anrede, Vorname, Nachname, Status, CustomData)" />
      <Arrow />
      <FlowNode type="subprocess" label="DEX_Participants aktualisieren (EventNumber in EventRegistered oder EventOnWaitlist)" />
      <Arrow />
      <FlowNode type="decision" label="Event hat 'E-Mails versenden' aktiv?" />
      <BranchContainer>
        <Branch label="Ja">
          <FlowNode type="data" label="DEX_Emails: Bestätigungs-E-Mail in Queue" />
        </Branch>
        <Branch label="Nein">
          <FlowNode type="process" color="#f5f5f5" label="Kein Mail-Eintrag - geskippt" />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="decision" label="Event hat 'Outlook-Einladung' aktiv UND Status = Angemeldet?" />
      <BranchContainer>
        <Branch label="Ja">
          <FlowNode type="data" label="DEX_Outlook: Kalender-Einladung in Queue ('Einladen')" />
        </Branch>
        <Branch label="Nein">
          <FlowNode type="process" color="#f5f5f5" label="Kein Outlook-Eintrag - geskippt" />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="end" label="Registrierung abgeschlossen" />
    </div>
  );
}

function CancellationFlow(): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <FlowNode type="start" label="User klickt 'Abmelden'" />
      <Arrow />
      <FlowNode type="decision" label="Zweiter Klick zur Bestätigung?" />
      <BranchContainer>
        <Branch label="Nein">
          <FlowNode type="end" color="var(--dex-orange)" label="Abbruch — Anmeldung behalten" />
        </Branch>
        <Branch label="Ja">
          <FlowNode type="process" label="Weiter mit Abmeldung" />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="subprocess" label="Teilnehmerliste: Status → 'Abgemeldet', CancellationDate setzen, TeilnehmerID → null" />
      <Arrow />
      <FlowNode type="subprocess" label="DEX_Participants: EventNumber aus EventRegistered/EventOnWaitlist entfernen" />
      <Arrow />
      <FlowNode type="decision" label="Event hat 'E-Mails versenden' aktiv?" />
      <BranchContainer>
        <Branch label="Ja">
          <FlowNode type="data" label="DEX_Emails: Abmeldungs-E-Mail in Queue" />
        </Branch>
        <Branch label="Nein">
          <FlowNode type="process" color="#f5f5f5" label="Kein Mail-Eintrag - geskippt" />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="decision" label="Event hat 'Outlook-Einladung' aktiv?" />
      <BranchContainer>
        <Branch label="Ja">
          <FlowNode type="data" label="DEX_Outlook: Kalender-Einladung zurückziehen ('Ausladen')" />
        </Branch>
        <Branch label="Nein">
          <FlowNode type="process" color="#f5f5f5" label="Kein Outlook-Eintrag - geskippt" />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="data" color="#ffebee" label="DEX_IDReorder: Reorder-Auftrag in Queue (Pending)" />
      <Arrow />
      <FlowNode type="end" label="Abmeldung abgeschlossen → Power Automate übernimmt ID-Korrektur" />
    </div>
  );
}

function IDReorderFlow(): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <FlowNode type="start" color="#6a1b9a" label="Power Automate Trigger: Neuer Eintrag in DEX_IDReorder (Status = Pending)" />
      <Arrow />
      <FlowNode type="process" label="Concurrency = 1 (nur eine Instanz gleichzeitig, weitere warten)" />
      <Arrow />
      <FlowNode type="process" label="Status → 'Processing'" />
      <Arrow />
      <FlowNode type="subprocess" label="Get_Enrolled_Participants: alle Teilnehmer mit Status ≠ 'Abgemeldet' laden (Aktive + Warteliste), sortiert nach RegistrationDate asc" />
      <Arrow />
      <FlowNode type="process" label="Count_Active: Anzahl Enrolled-Einträge mit Status ≠ 'Warteliste' zählen (= aktuell belegte Plätze)" />
      <Arrow />
      <FlowNode type="process" label="TeilnehmerIDs neu vergeben: 1, 2, 3, ... N (lückenlos, Aktive zuerst wegen RegistrationDate, Warteliste dahinter)" />
      <Arrow />
      <FlowNode type="subprocess" label="$batch-Update an SharePoint: nur TeilnehmerID schreiben (Status bleibt unangetastet — Warteliste bleibt Warteliste)" />
      <Arrow />
      <FlowNode type="decision" label="Platz frei? (Count_Active < MaxParticipants AND MaxParticipants > 0)" />
      <BranchContainer>
        <Branch label="Ja">
          <FlowNode type="subprocess" label="Get_Waitlist_First: ersten Warteliste-Eintrag laden (Status eq 'Warteliste', sortiert nach RegistrationDate asc)" />
          <Arrow />
          <FlowNode type="process" label="Promote_Waitlist: MERGE Status → 'Angemeldet' (KEINE TeilnehmerID-Änderung — der Batch hat die neue TID bereits korrekt auf Count_Active + 1 gesetzt)" />
          <Arrow />
          <FlowNode type="data" label="DEX_Emails: Nachrücker-E-Mail mit Template 'Nachruecken' (pro EmailLanguage)" />
          <Arrow />
          <FlowNode type="data" label="DEX_Outlook: Kalender-Einladung in Queue ('Einladen')" />
        </Branch>
        <Branch label="Nein">
          <FlowNode type="process" label="Kein Nachrücken nötig — Event voll oder keine Warteliste" />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="process" label="Status → 'Done'" />
      <Arrow />
      <FlowNode type="end" color="#6a1b9a" label="Flow-Instanz beendet. Bei weiteren Queue-Einträgen startet die nächste Instanz." />
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
