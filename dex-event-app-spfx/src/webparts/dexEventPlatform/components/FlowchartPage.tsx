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
      <FlowNode type="process" label="Vorname, Nachname, E-Mail, Event-spezifische Felder eingeben" />
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
      <FlowNode type="subprocess" label="Eintrag in Subsite-Teilnehmerliste (Vorname, Nachname, Status, CustomData)" />
      <Arrow />
      <FlowNode type="subprocess" label="DEX_Participants aktualisieren (EventNumber in EventRegistered oder EventOnWaitlist)" />
      <Arrow />
      <FlowNode type="data" label="DEX_Emails: Bestätigungs-E-Mail in Queue" />
      <Arrow />
      <FlowNode type="data" label="DEX_Outlook: Kalender-Einladung in Queue (nur bei Anmeldung, nicht Warteliste)" />
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
      <FlowNode type="data" label="DEX_Emails: Abmeldungs-E-Mail in Queue" />
      <Arrow />
      <FlowNode type="data" label="DEX_Outlook: Kalender-Einladung zurückziehen" />
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
      <FlowNode type="subprocess" label="Alle aktiven Teilnehmer laden (Status = Angemeldet oder Eingecheckt) von SubsiteUrl" />
      <Arrow />
      <FlowNode type="process" label="Sortieren nach RegistrationDate (älteste zuerst)" />
      <Arrow />
      <FlowNode type="process" label="TeilnehmerIDs vergeben: 1, 2, 3, ... N (lückenlos)" />
      <Arrow />
      <FlowNode type="subprocess" label="$batch Updates an SharePoint (50 Items pro Batch, 1-2 Sek Pause)" />
      <Arrow />
      <FlowNode type="decision" label="Platz frei? (Angemeldet < MaxParticipants)" />
      <BranchContainer>
        <Branch label="Ja">
          <FlowNode type="process" label="Ersten Warteliste-Eintrag nachrücken → Status = Angemeldet" />
          <Arrow />
          <FlowNode type="data" label="DEX_Participants: EventNumber von Waitlist → Registered" />
          <Arrow />
          <FlowNode type="data" label="DEX_Emails: Nachrücker-E-Mail" />
          <Arrow />
          <FlowNode type="data" label="DEX_Outlook: Kalender-Einladung" />
        </Branch>
        <Branch label="Nein">
          <FlowNode type="process" label="Kein Nachrücken nötig" />
        </Branch>
      </BranchContainer>
      <Arrow />
      <FlowNode type="process" label="Status → 'Done'" />
      <Arrow />
      <FlowNode type="end" color="#6a1b9a" label="Flow-Instanz beendet. Falls weitere Queue-Einträge: Nächste Instanz startet." />
    </div>
  );
}

function EventCreationFlow(): React.ReactElement {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
      <FlowNode type="start" label="Organizer/Admin erstellt Event" />
      <Arrow />
      <FlowNode type="process" label="Nächste EventNumber ermitteln (max + 1)" />
      <Arrow />
      <FlowNode type="subprocess" label="SharePoint Subsite erstellen (URL aus Titel generiert)" />
      <Arrow />
      <FlowNode type="subprocess" label="Teilnehmerliste 'Teilnehmer' auf Subsite erstellen (Basis- + Custom-Felder)" />
      <Arrow />
      <FlowNode type="process" label="Item-Level Security + Berechtigungen setzen" />
      <Arrow />
      <FlowNode type="subprocess" label="Event in DEX_Events eintragen (mit EventNumber, SubsiteUrl, CustomFields)" />
      <Arrow />
      <FlowNode type="end" label="Event bereit für Registrierungen" />
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
  ];

  const renderFlow = (): React.ReactElement => {
    switch (activeFlow) {
      case 'registration': return <RegistrationFlow />;
      case 'cancellation': return <CancellationFlow />;
      case 'reorder': return <IDReorderFlow />;
      case 'creation': return <EventCreationFlow />;
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
