/**
 * Wiederverwendbare Mockup- und Live-Demo-Komponenten fuer das Handbuch.
 *
 * - PlaceholderShot: gestylter "Screenshot folgt"-Slot
 * - MiniBrowser: simuliert ein Browser-Frame um echte App-Komponenten herum
 * - ClickPath: SVG-Mockup mit Click-Cursor + Highlight + Pfeil
 * - DemoCard: vereinfachter EventCard-Mockup
 * - DemoQuizQuestion: Quiz-Frage-Mockup mit Answer-Buttons
 * - DemoParticipantTable: Mini-Teilnehmer-Tabelle mit Beispieldaten
 *
 * Beispielnamen sind durchweg Mustermann/Musterfrau, NIE echte Kollegen.
 */
import * as React from 'react';

// ============================================================================
// PlaceholderShot — Slot fuer echten Screenshot (bis User Bild liefert)
// ============================================================================

export const PlaceholderShot: React.FC<{ caption?: string; height?: number }> = ({ caption, height = 240 }) => (
  <div style={{
    border: '2px dashed var(--dex-gray-300)',
    background: 'repeating-linear-gradient(45deg, var(--dex-gray-50) 0 10px, transparent 10px 20px)',
    borderRadius: 8,
    height,
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: 6,
    padding: 16,
    textAlign: 'center',
    color: 'var(--dex-gray-500)',
    fontSize: '0.85rem',
  }}>
    <div style={{ fontWeight: 700, fontSize: '0.95rem' }}>Screenshot folgt</div>
    {caption && <div style={{ maxWidth: 480, lineHeight: 1.4 }}>{caption}</div>}
  </div>
);

// ============================================================================
// MiniBrowser — ein Frame der einen Bereich der App "wie im Browser" zeigt
// ============================================================================

export const MiniBrowser: React.FC<{ title?: string; children: React.ReactNode }> = ({ title = 'DEX Event Experience Platform', children }) => (
  <div style={{
    border: '1px solid var(--dex-gray-300)',
    borderRadius: 10,
    overflow: 'hidden',
    boxShadow: '0 4px 14px rgba(0,0,0,0.08)',
    background: '#fff',
  }}>
    <div style={{
      background: '#f3f3f3',
      borderBottom: '1px solid var(--dex-gray-200)',
      padding: '8px 14px',
      display: 'flex',
      alignItems: 'center',
      gap: 10,
      fontSize: '0.78rem',
      color: 'var(--dex-gray-500)',
    }}>
      <span style={{ display: 'inline-flex', gap: 5 }}>
        <span style={{ width: 10, height: 10, borderRadius: '50%', background: '#ff5f57' }} />
        <span style={{ width: 10, height: 10, borderRadius: '50%', background: '#ffbd2e' }} />
        <span style={{ width: 10, height: 10, borderRadius: '50%', background: '#28ca42' }} />
      </span>
      <span style={{ flex: 1, textAlign: 'center', fontWeight: 500 }}>{title}</span>
    </div>
    <div style={{ background: '#fff', padding: 16 }}>{children}</div>
  </div>
);

// ============================================================================
// ClickPath — SVG-Mockup mit Cursor-Pfeil + Highlight um ein Click-Target
// ============================================================================

export const ClickPath: React.FC<{
  label: string;
  /** Optional zweiter Step im selben Bild */
  label2?: string;
  hint?: string;
}> = ({ label, label2, hint }) => (
  <div style={{
    background: '#fff',
    border: '1px solid var(--dex-gray-200)',
    borderRadius: 8,
    padding: 16,
    minHeight: 140,
    position: 'relative',
  }}>
    <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
      <div style={{
        padding: '10px 18px',
        borderRadius: 8,
        background: 'rgba(134,188,37,0.15)',
        border: '2px solid var(--dex-green)',
        fontSize: '0.85rem',
        fontWeight: 600,
        color: 'var(--dex-green-dark)',
        position: 'relative',
      }}>
        {label}
        <span style={{ position: 'absolute', top: -10, right: -10, background: 'var(--dex-green-dark)', color: '#fff', borderRadius: '50%', width: 22, height: 22, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.75rem', fontWeight: 700 }}>1</span>
      </div>
      {label2 && (
        <>
          <span style={{ color: 'var(--dex-gray-400)', fontSize: '1.4rem' }}>→</span>
          <div style={{
            padding: '10px 18px',
            borderRadius: 8,
            background: 'rgba(0,118,168,0.12)',
            border: '2px solid var(--dex-blue, #0076a8)',
            fontSize: '0.85rem',
            fontWeight: 600,
            color: 'var(--dex-blue, #0076a8)',
            position: 'relative',
          }}>
            {label2}
            <span style={{ position: 'absolute', top: -10, right: -10, background: 'var(--dex-blue, #0076a8)', color: '#fff', borderRadius: '50%', width: 22, height: 22, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.75rem', fontWeight: 700 }}>2</span>
          </div>
        </>
      )}
    </div>
    {hint && (
      <div style={{ marginTop: 12, fontSize: '0.78rem', color: 'var(--dex-gray-500)', fontStyle: 'italic' }}>
        💡 {hint}
      </div>
    )}
  </div>
);

// ============================================================================
// DemoEventCard — vereinfachte EventCard mit Mock-Daten
// ============================================================================

export const DemoEventCard: React.FC<{ title?: string; date?: string; location?: string; status?: string }> = ({
  title = 'B2Run Frankfurt 2026 (Beispiel)',
  date = '12.06.2026 — 18:30 Uhr',
  location = 'Commerzbank-Arena, Frankfurt',
  status = '34 / 100 Plätze frei',
}) => (
  <div style={{
    border: '1px solid var(--dex-gray-200)',
    borderRadius: 10,
    overflow: 'hidden',
    maxWidth: 360,
    background: '#fff',
  }}>
    <div style={{
      height: 120,
      background: 'linear-gradient(135deg, #86bc25 0%, #0076a8 100%)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      color: '#fff',
      fontWeight: 700,
      fontSize: '1.2rem',
      textShadow: '0 1px 3px rgba(0,0,0,0.3)',
      padding: 16,
      textAlign: 'center',
    }}>
      {title}
    </div>
    <div style={{ padding: 14 }}>
      <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>{date}</div>
      <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>{location}</div>
      <div style={{ marginTop: 8, fontSize: '0.78rem', color: 'var(--dex-green-dark)', fontWeight: 600 }}>{status}</div>
      <button style={{
        marginTop: 10, width: '100%', padding: '8px 14px', borderRadius: 6,
        background: 'var(--dex-green)', color: '#fff', border: 'none',
        fontSize: '0.85rem', fontWeight: 600, cursor: 'pointer',
      }}>Anmelden</button>
    </div>
  </div>
);

// ============================================================================
// DemoQuizQuestion — eine Quiz-Frage mit Answer-Buttons
// ============================================================================

export const DemoQuizQuestion: React.FC<{ showAnswered?: boolean }> = ({ showAnswered = false }) => (
  <div style={{
    border: '1px solid var(--dex-gray-200)',
    borderRadius: 10,
    padding: 20,
    background: '#fff',
    maxWidth: 540,
  }}>
    <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginBottom: 6 }}>Frage 2 von 5</div>
    <div style={{ fontSize: '1rem', fontWeight: 600, marginBottom: 14 }}>
      In welcher Stadt findet das Event statt?
    </div>
    {['München', 'Frankfurt', 'Berlin', 'Köln'].map((opt, idx) => {
      const isCorrect = opt === 'Frankfurt';
      const isSelected = showAnswered && opt === 'Frankfurt';
      return (
        <button
          key={opt}
          style={{
            display: 'block', width: '100%', textAlign: 'left',
            padding: '10px 14px', marginBottom: 6,
            borderRadius: 8,
            border: `2px solid ${showAnswered && isCorrect ? 'var(--dex-green)' : 'var(--dex-gray-300)'}`,
            background: showAnswered && isSelected ? 'rgba(134,188,37,0.15)' : '#fff',
            fontSize: '0.85rem', cursor: 'pointer',
            color: showAnswered && isCorrect ? 'var(--dex-green-dark)' : 'inherit',
            fontWeight: showAnswered && isCorrect ? 600 : 400,
          }}
        >
          <span style={{ fontWeight: 700, marginRight: 10 }}>{String.fromCharCode(65 + idx)}</span>
          {opt}
          {showAnswered && isCorrect && <span style={{ float: 'right' }}>✓</span>}
        </button>
      );
    })}
  </div>
);

// ============================================================================
// DemoParticipantTable — Mini-Teilnehmertabelle mit Mock-Eintraegen
// ============================================================================

export const DemoParticipantTable: React.FC<{ withQuizScore?: boolean }> = ({ withQuizScore = false }) => {
  const rows = [
    { id: 1, anrede: 'Herr', vorname: 'Max', nachname: 'Mustermann', email: 'max.mustermann@deloitte.de', status: 'Angemeldet', score: '4/5' },
    { id: 2, anrede: 'Frau', vorname: 'Erika', nachname: 'Musterfrau', email: 'erika.musterfrau@deloitte.de', status: 'Angemeldet', score: '5/5' },
    { id: 3, anrede: 'Frau', vorname: 'Marie', nachname: 'Musterfrau', email: 'marie.musterfrau@deloitte.de', status: 'Eingecheckt', score: '3/5' },
    { id: 4, anrede: 'Herr', vorname: 'Otto', nachname: 'Mustermann', email: 'otto.mustermann@deloitte.de', status: 'Warteliste', score: '—' },
  ];
  return (
    <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 8, overflow: 'hidden', background: '#fff' }}>
      <table style={{ width: '100%', fontSize: '0.78rem', borderCollapse: 'collapse' }}>
        <thead style={{ background: 'var(--dex-gray-50)', textAlign: 'left' }}>
          <tr>
            <th style={{ padding: '8px 10px', borderBottom: '1px solid var(--dex-gray-200)' }}>ID</th>
            <th style={{ padding: '8px 10px', borderBottom: '1px solid var(--dex-gray-200)' }}>Name</th>
            <th style={{ padding: '8px 10px', borderBottom: '1px solid var(--dex-gray-200)' }}>E-Mail</th>
            <th style={{ padding: '8px 10px', borderBottom: '1px solid var(--dex-gray-200)' }}>Status</th>
            {withQuizScore && <th style={{ padding: '8px 10px', borderBottom: '1px solid var(--dex-gray-200)' }}>Quiz</th>}
          </tr>
        </thead>
        <tbody>
          {rows.map(r => (
            <tr key={r.id}>
              <td style={{ padding: '8px 10px', borderBottom: '1px solid var(--dex-gray-100)' }}>{r.id}</td>
              <td style={{ padding: '8px 10px', borderBottom: '1px solid var(--dex-gray-100)' }}>
                <strong>{r.anrede} {r.vorname} {r.nachname}</strong>
              </td>
              <td style={{ padding: '8px 10px', borderBottom: '1px solid var(--dex-gray-100)', color: 'var(--dex-gray-500)' }}>{r.email}</td>
              <td style={{ padding: '8px 10px', borderBottom: '1px solid var(--dex-gray-100)' }}>
                <span style={{
                  fontSize: '0.7rem', fontWeight: 600,
                  padding: '2px 8px', borderRadius: 10,
                  background: r.status === 'Eingecheckt' ? 'rgba(134,188,37,0.18)' : r.status === 'Warteliste' ? 'rgba(237,139,0,0.15)' : 'rgba(0,118,168,0.12)',
                  color: r.status === 'Eingecheckt' ? 'var(--dex-green-dark)' : r.status === 'Warteliste' ? '#b86700' : 'var(--dex-blue, #0076a8)',
                }}>{r.status}</span>
              </td>
              {withQuizScore && <td style={{ padding: '8px 10px', borderBottom: '1px solid var(--dex-gray-100)', fontFamily: 'monospace' }}>{r.score}</td>}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

// ============================================================================
// DemoStepper — vertikale Step-Liste (zeigt Wizard-Steps des Event-Edit)
// ============================================================================

export const DemoStepper: React.FC<{ activeStep?: number }> = ({ activeStep = 0 }) => {
  const steps = ['Grundinformationen', 'Zeit & Ort', 'Kapazität', 'Felder', 'Kommunikation', 'Dokumente', 'FunZone (Quiz)'];
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 6, background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 14, maxWidth: 320 }}>
      {steps.map((s, i) => (
        <div key={s} style={{
          display: 'flex', alignItems: 'center', gap: 10,
          padding: '6px 10px', borderRadius: 6,
          background: i === activeStep ? 'rgba(134,188,37,0.15)' : 'transparent',
          fontWeight: i === activeStep ? 700 : 400,
          color: i === activeStep ? 'var(--dex-green-dark)' : 'var(--dex-gray-700)',
          fontSize: '0.85rem',
        }}>
          <span style={{
            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
            width: 22, height: 22, borderRadius: '50%',
            background: i === activeStep ? 'var(--dex-green)' : 'var(--dex-gray-200)',
            color: i === activeStep ? '#fff' : 'var(--dex-gray-600)',
            fontSize: '0.7rem', fontWeight: 700,
          }}>{i + 1}</span>
          {s}
        </div>
      ))}
    </div>
  );
};
