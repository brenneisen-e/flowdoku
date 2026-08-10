/**
 * Wiederverwendbare Mockup- und Live-Demo-Komponenten für das Handbuch.
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
// PlaceholderShot — Slot für echten Screenshot (bis User Bild liefert)
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
// DemoParticipantTable — Mini-Teilnehmertabelle mit Mock-Einträgen
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

// ============================================================================
// Callout — farbige Hinweis-Box (info | tip | warning | success)
// ============================================================================

export const Callout: React.FC<{ variant?: 'info' | 'tip' | 'warning' | 'success'; title?: string; children: React.ReactNode }> = ({ variant = 'info', title, children }) => {
  const styles: Record<string, { bg: string; border: string; color: string; icon: string }> = {
    info: { bg: 'rgba(0,118,168,0.06)', border: '#0076a8', color: '#0b3c52', icon: 'ⓘ' },
    tip: { bg: 'rgba(255,193,7,0.08)', border: '#f0ad2e', color: '#8a5d00', icon: '💡' },
    warning: { bg: 'rgba(200,30,30,0.06)', border: '#c9302c', color: '#7a1b1b', icon: '⚠' },
    success: { bg: 'rgba(134,188,37,0.10)', border: '#86bc25', color: '#4a7c1f', icon: '✓' },
  };
  const s = styles[variant];
  return (
    <div style={{ background: s.bg, borderLeft: `4px solid ${s.border}`, borderRadius: 6, padding: '10px 14px', color: s.color, fontSize: '0.85rem', lineHeight: 1.5 }}>
      <div style={{ display: 'flex', gap: 8 }}>
        <span style={{ fontWeight: 700 }}>{s.icon}</span>
        <div>
          {title && <div style={{ fontWeight: 700, marginBottom: 4 }}>{title}</div>}
          <div>{children}</div>
        </div>
      </div>
    </div>
  );
};

// ============================================================================
// DemoFormField — ein visualisiertes Formularfeld mit Label, Input und Info-Icon
// ============================================================================

export const DemoFormField: React.FC<{ label: string; value?: string; placeholder?: string; required?: boolean; info?: string; type?: 'text' | 'select' | 'textarea' }> = ({ label, value, placeholder, required, info, type = 'text' }) => (
  <div style={{ marginBottom: 10 }}>
    <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '0.82rem', fontWeight: 600, color: 'var(--dex-gray-700)', marginBottom: 4 }}>
      {required && <span style={{ color: '#c9302c' }}>*</span>}
      {label}
      {info && (
        <span title={info} style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: 18, height: 18, borderRadius: '50%', background: 'var(--dex-gray-200)', color: 'var(--dex-gray-600)', fontSize: '0.7rem', fontWeight: 700, cursor: 'help' }}>i</span>
      )}
    </label>
    {type === 'textarea' ? (
      <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 6, padding: '8px 10px', minHeight: 60, fontSize: '0.85rem', color: value ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)', background: '#fff' }}>
        {value || placeholder || ''}
      </div>
    ) : (
      <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 6, padding: '8px 10px', fontSize: '0.85rem', color: value ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)', background: '#fff', display: 'flex', alignItems: 'center', gap: 6 }}>
        <span style={{ flex: 1 }}>{value || placeholder || ''}</span>
        {type === 'select' && <span style={{ color: 'var(--dex-gray-400)' }}>▾</span>}
      </div>
    )}
  </div>
);

// ============================================================================
// DemoToggle — simulierter Toggle-Switch mit Label
// ============================================================================

export const DemoToggle: React.FC<{ label: string; on?: boolean; hint?: string }> = ({ label, on = false, hint }) => (
  <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: hint ? 4 : 0 }}>
    <span style={{
      display: 'inline-block', width: 40, height: 22, borderRadius: 999,
      background: on ? 'var(--dex-green)' : 'var(--dex-gray-300)',
      position: 'relative', transition: 'background 0.2s',
    }}>
      <span style={{
        position: 'absolute', top: 2, left: on ? 20 : 2,
        width: 18, height: 18, borderRadius: '50%', background: '#fff',
        transition: 'left 0.2s', boxShadow: '0 1px 3px rgba(0,0,0,0.2)',
      }} />
    </span>
    <span style={{ fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>{label}</span>
    {hint && <span style={{ fontSize: '0.75rem', color: 'var(--dex-gray-500)', fontStyle: 'italic', marginLeft: 'auto' }}>{hint}</span>}
  </div>
);

// ============================================================================
// DemoPillToggles — moderner Standortfilter (Pill-Buttons mit Check)
// ============================================================================

export const DemoPillToggles: React.FC<{ options: string[]; selected?: string[] }> = ({ options, selected = [] }) => (
  <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
    {options.map(o => {
      const isOn = selected.indexOf(o) >= 0;
      return (
        <span key={o} style={{
          display: 'inline-flex', alignItems: 'center', gap: 6,
          padding: '5px 12px', borderRadius: 999,
          border: `1.5px solid ${isOn ? 'var(--dex-green)' : 'var(--dex-gray-300)'}`,
          background: isOn ? 'var(--dex-green)' : '#fff',
          color: isOn ? '#fff' : 'var(--dex-gray-700)',
          fontSize: '0.78rem', fontWeight: isOn ? 500 : 400,
        }}>
          {isOn && <span>✓</span>}
          {o}
        </span>
      );
    })}
  </div>
);

// ============================================================================
// DemoOrganizerChips — der Chip-Picker für Organizer mit Remove + Reorder
// ============================================================================

export const DemoOrganizerChips: React.FC<{ names?: string[] }> = ({ names = ['Erika Musterfrau', 'Max Mustermann'] }) => (
  <div>
    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 10 }}>
      {names.map((n, i) => (
        <span key={n} style={{
          display: 'inline-flex', alignItems: 'center', gap: 6,
          padding: '4px 4px 4px 12px', background: 'var(--dex-green)', color: '#fff',
          borderRadius: 999, fontSize: '0.8rem', fontWeight: 500,
        }}>
          {n}
          {i > 0 && <span style={{ background: 'rgba(255,255,255,0.25)', width: 20, height: 20, borderRadius: '50%', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.7rem' }}>◀</span>}
          {i < names.length - 1 && <span style={{ background: 'rgba(255,255,255,0.25)', width: 20, height: 20, borderRadius: '50%', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.7rem' }}>▶</span>}
          <span style={{ background: 'rgba(255,255,255,0.25)', width: 20, height: 20, borderRadius: '50%', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.85rem' }}>×</span>
        </span>
      ))}
    </div>
    <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 6, padding: '8px 10px', fontSize: '0.85rem', color: 'var(--dex-gray-400)', background: '#fff' }}>
      Name oder E-Mail eingeben und aus der Liste auswählen…
    </div>
  </div>
);

// ============================================================================
// DemoWizardProgress — horizontaler Fortschritt durch Event-Erstellung
// ============================================================================

export const DemoWizardProgress: React.FC<{ activeStep?: number }> = ({ activeStep = 0 }) => {
  const steps = ['Grundlagen', 'Zeit & Ort', 'Kapazität', 'Felder', 'Kommunikation', 'Dokumente', 'Fun-Zone'];
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 4, flexWrap: 'wrap', padding: '8px 0' }}>
      {steps.map((s, i) => {
        const done = i < activeStep;
        const active = i === activeStep;
        return (
          <React.Fragment key={s}>
            <div style={{
              display: 'inline-flex', alignItems: 'center', gap: 6,
              padding: '4px 10px', borderRadius: 999,
              background: active ? 'var(--dex-green)' : done ? 'rgba(134,188,37,0.15)' : 'var(--dex-gray-100)',
              color: active ? '#fff' : done ? 'var(--dex-green-dark)' : 'var(--dex-gray-600)',
              fontSize: '0.78rem', fontWeight: active || done ? 600 : 400,
            }}>
              <span style={{
                display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                width: 18, height: 18, borderRadius: '50%',
                background: active ? 'rgba(255,255,255,0.25)' : done ? 'var(--dex-green)' : 'var(--dex-gray-300)',
                color: done && !active ? '#fff' : active ? '#fff' : 'var(--dex-gray-600)',
                fontSize: '0.68rem', fontWeight: 700,
              }}>{done ? '✓' : i + 1}</span>
              {s}
            </div>
            {i < steps.length - 1 && <span style={{ color: 'var(--dex-gray-300)' }}>›</span>}
          </React.Fragment>
        );
      })}
    </div>
  );
};

// ============================================================================
// DemoAdminCenter — Admin Center Überblick mit Kacheln + Metriken
// ============================================================================

export const DemoAdminCenter: React.FC = () => (
  <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: 18, background: '#fff' }}>
    <div style={{ fontSize: '0.95rem', fontWeight: 700, marginBottom: 12 }}>B2Run Frankfurt 2026 — Admin</div>
    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 10, marginBottom: 14 }}>
      {[
        { label: 'Angemeldet', value: '67', color: 'var(--dex-green)' },
        { label: 'Warteliste', value: '14', color: '#f0ad2e' },
        { label: 'Eingecheckt', value: '52', color: 'var(--dex-blue, #0076a8)' },
        { label: 'Abgemeldet', value: '3', color: 'var(--dex-gray-500)' },
      ].map(m => (
        <div key={m.label} style={{ background: 'var(--dex-gray-50)', borderRadius: 8, padding: 10, textAlign: 'center' }}>
          <div style={{ fontSize: '1.4rem', fontWeight: 700, color: m.color }}>{m.value}</div>
          <div style={{ fontSize: '0.7rem', color: 'var(--dex-gray-600)', textTransform: 'uppercase', letterSpacing: 0.5 }}>{m.label}</div>
        </div>
      ))}
    </div>
    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 10 }}>
      {['QR senden', 'E-Mails kopieren', 'Massenmail', 'IDs neu vergeben', 'Spalten fixen'].map(a => (
        <span key={a} style={{ padding: '5px 10px', borderRadius: 6, border: '1px solid var(--dex-gray-300)', background: '#fff', fontSize: '0.75rem', color: 'var(--dex-gray-700)' }}>{a}</span>
      ))}
    </div>
  </div>
);

// ============================================================================
// DemoCheckInScanner — QR-Scanner UI
// ============================================================================

export const DemoCheckInScanner: React.FC<{ scanned?: boolean }> = ({ scanned = false }) => (
  <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: 14, background: '#fff', maxWidth: 340 }}>
    <div style={{ fontSize: '0.85rem', fontWeight: 700, marginBottom: 8 }}>Check-In — B2Run Frankfurt</div>
    <div style={{
      width: '100%', aspectRatio: '1 / 1', background: '#000', borderRadius: 8,
      position: 'relative', overflow: 'hidden',
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      marginBottom: 10,
    }}>
      <div style={{
        position: 'absolute', inset: 28, border: '2px solid rgba(134,188,37,0.8)',
        borderRadius: 12,
      }} />
      <div style={{
        position: 'absolute', left: 28, right: 28,
        top: scanned ? '70%' : '30%',
        height: 2, background: 'var(--dex-green)', boxShadow: '0 0 10px var(--dex-green)',
        transition: 'top 0.4s',
      }} />
      <span style={{ color: '#fff', fontSize: '0.75rem', opacity: 0.6, zIndex: 2 }}>Kamera-Vorschau</span>
    </div>
    <div style={{
      padding: '8px 12px', borderRadius: 6,
      background: scanned ? 'rgba(134,188,37,0.15)' : 'var(--dex-gray-50)',
      color: scanned ? 'var(--dex-green-dark)' : 'var(--dex-gray-600)',
      fontSize: '0.82rem', fontWeight: 500, textAlign: 'center',
    }}>
      {scanned ? '✓ Max Mustermann eingecheckt' : 'Warte auf QR-Code …'}
    </div>
  </div>
);

// ============================================================================
// DemoRolesTable — Rollen-Übersicht
// ============================================================================

export const DemoRolesTable: React.FC = () => {
  const rows = [
    { name: 'Erika Musterfrau', email: 'erika.musterfrau@deloitte.de', role: 'Admin', color: '#2e7d32', bg: '#e8f5e9' },
    { name: 'Max Mustermann', email: 'max.mustermann@deloitte.de', role: 'Organizer', color: '#1565c0', bg: '#e3f2fd' },
    { name: 'Marie Musterfrau', email: 'marie.musterfrau@deloitte.de', role: 'Organizer', color: '#1565c0', bg: '#e3f2fd' },
  ];
  return (
    <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 8, overflow: 'hidden', background: '#fff' }}>
      <table style={{ width: '100%', fontSize: '0.8rem', borderCollapse: 'collapse' }}>
        <thead style={{ background: 'var(--dex-gray-50)', textAlign: 'left' }}>
          <tr>
            <th style={{ padding: '8px 10px' }}>Name</th>
            <th style={{ padding: '8px 10px' }}>E-Mail</th>
            <th style={{ padding: '8px 10px' }}>Rolle</th>
            <th style={{ padding: '8px 10px' }} />
          </tr>
        </thead>
        <tbody>
          {rows.map(r => (
            <tr key={r.email}>
              <td style={{ padding: '8px 10px', borderTop: '1px solid var(--dex-gray-100)' }}><strong>{r.name}</strong></td>
              <td style={{ padding: '8px 10px', borderTop: '1px solid var(--dex-gray-100)', color: 'var(--dex-gray-500)' }}>{r.email}</td>
              <td style={{ padding: '8px 10px', borderTop: '1px solid var(--dex-gray-100)' }}>
                <span style={{ padding: '2px 10px', borderRadius: 12, background: r.bg, color: r.color, fontSize: '0.72rem', fontWeight: 500 }}>{r.role}</span>
              </td>
              <td style={{ padding: '8px 10px', borderTop: '1px solid var(--dex-gray-100)', textAlign: 'right' }}>
                <span style={{ padding: '3px 8px', border: '1px solid var(--dex-gray-300)', borderRadius: 4, fontSize: '0.72rem', color: 'var(--dex-gray-600)' }}>Ändern</span>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

// ============================================================================
// DemoEmailPreview — Vorschau einer System-Mail
// ============================================================================

export const DemoEmailPreview: React.FC<{ subject?: string; heading?: string; body?: React.ReactNode }> = ({
  subject = 'Deine Anmeldung für B2Run Frankfurt 2026',
  heading = 'Anmeldebestätigung',
  body,
}) => (
  <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 10, overflow: 'hidden', background: '#fff', maxWidth: 480 }}>
    <div style={{ background: 'var(--dex-gray-50)', borderBottom: '1px solid var(--dex-gray-200)', padding: '8px 14px', fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
      <div><strong>Betreff:</strong> {subject}</div>
      <div><strong>An:</strong> max.mustermann@deloitte.de</div>
    </div>
    <div style={{ padding: 0 }}>
      <div style={{ background: 'var(--dex-green)', color: '#fff', padding: '16px 18px', fontSize: '1.05rem', fontWeight: 700 }}>
        {heading}
      </div>
      <div style={{ padding: 18, fontSize: '0.85rem', color: 'var(--dex-gray-700)', lineHeight: 1.6 }}>
        {body || (
          <>
            <p style={{ margin: '0 0 10px 0' }}>Hallo Max,</p>
            <p style={{ margin: '0 0 10px 0' }}>deine Anmeldung zu <strong>B2Run Frankfurt 2026</strong> ist bestätigt. Wir freuen uns auf dich!</p>
            <p style={{ margin: '0 0 10px 0' }}><strong>Wann:</strong> 12.06.2026 um 18:30 Uhr</p>
            <p style={{ margin: 0 }}><strong>Wo:</strong> Commerzbank-Arena, Frankfurt</p>
          </>
        )}
      </div>
      <div style={{ padding: 12, background: '#fafafa', borderTop: '1px solid var(--dex-gray-100)', textAlign: 'center', fontSize: '0.7rem', color: 'var(--dex-gray-400)' }}>
        Deloitte · Event Experience Platform
      </div>
    </div>
  </div>
);

// ============================================================================
// DemoQRCode — stilisierter QR-Code als SVG
// ============================================================================

export const DemoQRCode: React.FC<{ size?: number; label?: string }> = ({ size = 120, label = 'Dein Check-In-Code' }) => {
  const cells = 21;
  const pattern = React.useMemo(() => {
    // Deterministisches Pseudo-Muster
    const arr: boolean[][] = [];
    for (let y = 0; y < cells; y++) {
      arr[y] = [];
      for (let x = 0; x < cells; x++) {
        arr[y][x] = ((x * 7 + y * 11 + (x ^ y)) % 3) === 0;
      }
    }
    // Finder-Pattern-Ecken
    const setBlock = (ox: number, oy: number): void => {
      for (let y = 0; y < 7; y++) for (let x = 0; x < 7; x++) {
        const edge = y === 0 || y === 6 || x === 0 || x === 6;
        const inner = y >= 2 && y <= 4 && x >= 2 && x <= 4;
        arr[oy + y][ox + x] = edge || inner;
      }
    };
    setBlock(0, 0); setBlock(cells - 7, 0); setBlock(0, cells - 7);
    return arr;
  }, []);
  const cell = size / cells;
  return (
    <div style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: 6 }}>
      <svg width={size} height={size} style={{ background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 6 }}>
        {pattern.map((row, y) => row.map((on, x) => on && (
          <rect key={`${x}-${y}`} x={x * cell} y={y * cell} width={cell} height={cell} fill="#111" />
        )))}
      </svg>
      {label && <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)' }}>{label}</div>}
    </div>
  );
};

// ============================================================================
// DemoRoommatePicker — Person-Suche mit Benachrichtigungs-Hinweis
// ============================================================================

export const DemoRoommatePicker: React.FC = () => (
  <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: 14, background: '#fff', maxWidth: 440 }}>
    <div style={{ fontSize: '0.82rem', fontWeight: 600, marginBottom: 6 }}>Präferierter Zimmerpartner</div>
    <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 6, padding: '8px 10px', fontSize: '0.85rem', color: 'var(--dex-gray-800)', background: '#fff', marginBottom: 6 }}>
      Marie Musterfrau &lt;marie.musterfrau@deloitte.de&gt;
    </div>
    <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', fontStyle: 'italic' }}>
      Die ausgewählte Person wird per E-Mail über deine Auswahl informiert.
    </div>
  </div>
);

// ============================================================================
// DemoDateTimePicker — Datums- und Uhrzeit-Picker
// ============================================================================

export const DemoDateTimePicker: React.FC<{ label: string; value?: string; afterLabel?: string; afterValue?: string }> = ({ label, value = '12.06.2026, 18:30', afterLabel, afterValue }) => (
  <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
    <div style={{ flex: 1, minWidth: 200 }}>
      <label style={{ fontSize: '0.8rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>{label}</label>
      <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 6, padding: '8px 10px', fontSize: '0.85rem', background: '#fff', marginTop: 4 }}>
        📅 {value}
      </div>
    </div>
    {afterLabel && (
      <div style={{ flex: 1, minWidth: 200 }}>
        <label style={{ fontSize: '0.8rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>{afterLabel}</label>
        <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 6, padding: '8px 10px', fontSize: '0.85rem', background: '#fff', marginTop: 4 }}>
          📅 {afterValue}
        </div>
      </div>
    )}
  </div>
);

// ============================================================================
// DemoQuizEditor — Editor für eine Quiz-Frage (Form-View)
// ============================================================================

export const DemoQuizEditor: React.FC = () => (
  <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: 16, background: '#fff', maxWidth: 540 }}>
    <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', fontWeight: 700, letterSpacing: 0.5, marginBottom: 10 }}>Quiz-Editor · Frage 1</div>
    <div style={{ marginBottom: 10 }}>
      <label style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', display: 'block', marginBottom: 4 }}>Fragetext</label>
      <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 6, padding: '8px 10px', fontSize: '0.85rem', background: '#fff' }}>
        In welcher Stadt findet das Event statt?
      </div>
    </div>
    <div style={{ marginBottom: 8, fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)' }}>
      Antwortoptionen (eine als richtig markieren)
    </div>
    {[
      { text: 'München', correct: false },
      { text: 'Frankfurt', correct: true },
      { text: 'Berlin', correct: false },
      { text: 'Köln', correct: false },
    ].map((o, idx) => (
      <div key={o.text} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
        <span style={{
          width: 22, height: 22, borderRadius: '50%',
          display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
          border: `2px solid ${o.correct ? 'var(--dex-green)' : 'var(--dex-gray-300)'}`,
          background: o.correct ? 'var(--dex-green)' : '#fff',
          color: '#fff', fontSize: '0.7rem',
        }}>{o.correct ? '✓' : ''}</span>
        <div style={{ flex: 1, border: '1px solid var(--dex-gray-300)', borderRadius: 6, padding: '6px 10px', fontSize: '0.85rem', background: '#fff' }}>
          <span style={{ color: 'var(--dex-gray-400)', marginRight: 8, fontSize: '0.7rem' }}>{String.fromCharCode(65 + idx)}</span>
          {o.text}
        </div>
        <span style={{ cursor: 'pointer', color: 'var(--dex-gray-400)', fontSize: '0.85rem' }}>×</span>
      </div>
    ))}
    <div style={{ marginTop: 10, display: 'flex', gap: 8 }}>
      <span style={{ padding: '4px 10px', border: '1px dashed var(--dex-gray-300)', borderRadius: 6, fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>+ Option hinzufügen</span>
      <span style={{ padding: '4px 10px', border: '1px dashed var(--dex-gray-300)', borderRadius: 6, fontSize: '0.75rem', color: 'var(--dex-gray-500)' }}>+ Nächste Frage</span>
    </div>
  </div>
);

// ============================================================================
// DemoQuizResult — End-Screen nach Quiz-Abschluss
// ============================================================================

export const DemoQuizResult: React.FC<{ score?: number; total?: number }> = ({ score = 4, total = 5 }) => {
  const percent = Math.round((score / total) * 100);
  return (
    <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: 24, background: '#fff', maxWidth: 420, textAlign: 'center' }}>
      <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', textTransform: 'uppercase', fontWeight: 700, letterSpacing: 0.5, marginBottom: 14 }}>
        Dein Ergebnis
      </div>
      <div style={{
        width: 120, height: 120, margin: '0 auto 16px', borderRadius: '50%',
        background: `conic-gradient(var(--dex-green) ${percent * 3.6}deg, var(--dex-gray-100) 0deg)`,
        display: 'flex', alignItems: 'center', justifyContent: 'center', position: 'relative',
      }}>
        <div style={{ width: 96, height: 96, borderRadius: '50%', background: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', flexDirection: 'column' }}>
          <div style={{ fontSize: '1.6rem', fontWeight: 700, color: 'var(--dex-green-dark)' }}>{score}/{total}</div>
          <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>{percent}%</div>
        </div>
      </div>
      <div style={{ fontSize: '0.95rem', fontWeight: 600, marginBottom: 4 }}>Super gemacht, Max!</div>
      <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', marginBottom: 14 }}>
        Nur dein bester Versuch zählt — du kannst das Quiz beliebig oft wiederholen.
      </div>
      <div style={{ display: 'flex', gap: 8, justifyContent: 'center' }}>
        <span style={{ padding: '6px 16px', background: 'var(--dex-green)', color: '#fff', borderRadius: 6, fontSize: '0.82rem', fontWeight: 600 }}>Noch mal</span>
        <span style={{ padding: '6px 16px', background: 'var(--dex-gray-100)', color: 'var(--dex-gray-700)', borderRadius: 6, fontSize: '0.82rem' }}>Zurück zum Event</span>
      </div>
    </div>
  );
};

// ============================================================================
// DemoQuizScoreboard — Admin-Übersicht der Quiz-Ergebnisse
// ============================================================================

export const DemoQuizScoreboard: React.FC = () => {
  const rows = [
    { rank: 1, name: 'Erika Musterfrau', score: 5, time: '48s' },
    { rank: 2, name: 'Max Mustermann', score: 4, time: '1:12' },
    { rank: 3, name: 'Marie Musterfrau', score: 4, time: '1:54' },
    { rank: 4, name: 'Otto Mustermann', score: 3, time: '2:03' },
  ];
  return (
    <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 10, overflow: 'hidden', background: '#fff', maxWidth: 520 }}>
      <div style={{ padding: '12px 14px', borderBottom: '1px solid var(--dex-gray-100)', fontSize: '0.85rem', fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <span>Quiz-Ranking</span>
        <span style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', fontWeight: 400 }}>67 Teilnehmer · Ø 3,8 Punkte</span>
      </div>
      <table style={{ width: '100%', fontSize: '0.8rem', borderCollapse: 'collapse' }}>
        <thead style={{ background: 'var(--dex-gray-50)', textAlign: 'left' }}>
          <tr>
            <th style={{ padding: '8px 12px' }}>#</th>
            <th style={{ padding: '8px 12px' }}>Name</th>
            <th style={{ padding: '8px 12px' }}>Punkte</th>
            <th style={{ padding: '8px 12px' }}>Zeit</th>
          </tr>
        </thead>
        <tbody>
          {rows.map(r => (
            <tr key={r.name}>
              <td style={{ padding: '8px 12px', borderTop: '1px solid var(--dex-gray-100)', fontWeight: 700, color: r.rank === 1 ? 'var(--dex-green-dark)' : 'var(--dex-gray-700)' }}>
                {r.rank === 1 ? '🏆' : r.rank}
              </td>
              <td style={{ padding: '8px 12px', borderTop: '1px solid var(--dex-gray-100)' }}>{r.name}</td>
              <td style={{ padding: '8px 12px', borderTop: '1px solid var(--dex-gray-100)' }}>
                <span style={{ padding: '2px 10px', borderRadius: 12, background: r.score === 5 ? 'rgba(134,188,37,0.18)' : 'var(--dex-gray-100)', color: r.score === 5 ? 'var(--dex-green-dark)' : 'var(--dex-gray-700)', fontSize: '0.75rem', fontWeight: 600 }}>
                  {r.score}/5
                </span>
              </td>
              <td style={{ padding: '8px 12px', borderTop: '1px solid var(--dex-gray-100)', color: 'var(--dex-gray-500)', fontFamily: 'monospace', fontSize: '0.78rem' }}>{r.time}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

// ============================================================================
// DemoQuizStart — Einstiegs-Screen vor Quiz-Start
// ============================================================================

export const DemoQuizStart: React.FC = () => (
  <div style={{ border: '1px solid var(--dex-gray-200)', borderRadius: 10, padding: 24, background: 'linear-gradient(135deg, rgba(134,188,37,0.08), rgba(0,118,168,0.06))', maxWidth: 420, textAlign: 'center' }}>
    <div style={{ fontSize: '2.4rem', marginBottom: 6 }}>🎯</div>
    <div style={{ fontSize: '1.05rem', fontWeight: 700, marginBottom: 4 }}>Bereit für das Quiz?</div>
    <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, marginBottom: 16 }}>
      5 Fragen · keine Zeitbegrenzung · beliebig oft wiederholbar · nur dein bester Versuch wird gezählt
    </div>
    <div style={{ display: 'inline-block', padding: '10px 24px', background: 'var(--dex-green)', color: '#fff', borderRadius: 8, fontSize: '0.9rem', fontWeight: 600 }}>
      Quiz starten
    </div>
  </div>
);

// ============================================================================
// NumberedList — strukturierte nummerierte Schritte
// ============================================================================

export const NumberedList: React.FC<{ items: React.ReactNode[] }> = ({ items }) => (
  <ol style={{ paddingLeft: 0, listStyle: 'none', margin: 0 }}>
    {items.map((item, i) => (
      <li key={i} style={{ display: 'flex', gap: 10, marginBottom: 8, fontSize: '0.88rem', lineHeight: 1.5 }}>
        <span style={{ flexShrink: 0, width: 22, height: 22, borderRadius: '50%', background: 'var(--dex-green)', color: '#fff', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.75rem', fontWeight: 700, marginTop: 1 }}>{i + 1}</span>
        <span>{item}</span>
      </li>
    ))}
  </ol>
);
