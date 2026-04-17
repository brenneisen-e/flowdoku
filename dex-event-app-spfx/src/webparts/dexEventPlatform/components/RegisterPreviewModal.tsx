/**
 * Preview-Modal, das dem Organizer zeigt, wie die Registrierungsseite des
 * aktuellen Events den Teilnehmern erscheinen wird — rendered aus dem
 * aktuellen Form-State (noch nicht gespeichert).
 *
 * Einfache Variante der RegistrationPage ohne Absend-Logik: Event-Card,
 * Organizer-Liste, Standard-Felder (Anrede/Vorname/Nachname/E-Mail) und
 * alle konfigurierten Custom-Fields.
 */

import * as React from 'react';
import { X } from './Icons';
import OrganizerList from './OrganizerList';

export interface RegisterPreviewData {
  title: string;
  description: string;
  location: string;
  locationAddress: { street?: string; houseNo?: string; zip?: string; city?: string };
  startDate: string;   // "YYYY-MM-DDTHH:mm" local
  endDate: string;
  imagePreview?: string;
  organizers: string[];
  organizerEmails: string[];
  maxParticipants: number;
  unlimitedParticipants: boolean;
  customFields: Array<{
    id: string;
    label: string;
    type: string;
    required: boolean;
    visible: boolean;
    options?: string[];
  }>;
  isFictive?: boolean;
}

export interface RegisterPreviewModalProps {
  open: boolean;
  onClose: () => void;
  data: RegisterPreviewData;
}

function formatLocal(v: string): string {
  if (!v) return '—';
  const d = new Date(v);
  if (isNaN(d.getTime())) return v;
  return d.toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
}

export const RegisterPreviewModal: React.FC<RegisterPreviewModalProps> = ({ open, onClose, data }) => {
  if (!open) return null;

  const addr = data.locationAddress || {};
  const hasAddr = !!(addr.street || addr.houseNo || addr.zip || addr.city);
  const visibleFields = data.customFields.filter(f => f.visible !== false && f.label && f.label.trim());

  return (
    <div
      style={{
        position: 'fixed', inset: 0, zIndex: 1300,
        background: 'rgba(0,0,0,0.55)',
        display: 'flex', alignItems: 'stretch', justifyContent: 'center',
        padding: '24px 16px', overflowY: 'auto',
      }}
      onClick={onClose}
    >
      <div
        style={{
          width: '100%', maxWidth: 1100,
          background: '#fff', borderRadius: 'var(--dex-radius, 12px)',
          boxShadow: '0 20px 60px rgba(0,0,0,0.3)',
          display: 'flex', flexDirection: 'column', overflow: 'hidden',
          maxHeight: 'calc(100vh - 48px)',
        }}
        onClick={e => e.stopPropagation()}
      >
        {/* Header */}
        <div style={{
          display: 'flex', alignItems: 'center', justifyContent: 'space-between',
          padding: '14px 20px', borderBottom: '1px solid var(--dex-gray-200)',
        }}>
          <h3 style={{ margin: 0, fontSize: '1.05rem' }}>
            Registrierungsseite — Vorschau
            {data.isFictive && (
              <span style={{ marginLeft: 10, padding: '2px 10px', borderRadius: 999, background: 'var(--dex-orange, #ed8b00)', color: '#fff', fontSize: '0.7rem', fontWeight: 700, letterSpacing: 0.5 }}>
                TEST-EVENT
              </span>
            )}
          </h3>
          <button
            type="button"
            onClick={onClose}
            style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 4 }}
            aria-label="Schließen"
          >
            <X size={20} />
          </button>
        </div>

        <div style={{ padding: 24, overflowY: 'auto', background: 'var(--dex-gray-50, #f8f9fa)' }}>
          <div style={{
            display: 'grid',
            gridTemplateColumns: 'minmax(0, 1.2fr) minmax(0, 1.5fr)',
            gap: 20,
          }}>
            {/* LINKS: Event-Info */}
            <section>
              <div style={{ fontSize: '0.7rem', fontWeight: 700, textTransform: 'uppercase', letterSpacing: 0.5, color: 'var(--dex-red, #da291c)', marginBottom: 8 }}>
                Ausgewähltes Event
              </div>
              <div style={{
                background: '#fff', borderRadius: 10, overflow: 'hidden',
                border: '1px solid var(--dex-gray-200)',
              }}>
                {data.imagePreview ? (
                  <img src={data.imagePreview} alt={data.title} style={{ display: 'block', width: '100%', maxHeight: 260, objectFit: 'cover' }} />
                ) : (
                  <div style={{ height: 140, background: 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)' }} />
                )}
                <div style={{ padding: 16 }}>
                  <h4 style={{ margin: 0, fontSize: '1.05rem', fontWeight: 700 }}>{data.title || '(kein Titel)'}</h4>
                  <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', marginTop: 6 }}>
                    📅 {formatLocal(data.startDate)} – {formatLocal(data.endDate)}
                  </div>
                  {(data.location || hasAddr) && (
                    <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 8 }}>
                      {data.location && <div style={{ fontWeight: 700, color: 'var(--dex-gray-700)' }}>📍 {data.location}</div>}
                      {hasAddr && (
                        <div style={{ paddingLeft: 18 }}>
                          {[addr.street, addr.houseNo].filter(Boolean).join(' ')}
                          {(addr.zip || addr.city) && <br />}
                          {[addr.zip, addr.city].filter(Boolean).join(' ')}
                        </div>
                      )}
                    </div>
                  )}
                  {data.organizers.length > 0 && (
                    <div style={{ marginTop: 10 }}>
                      <div style={{ fontSize: '0.7rem', fontWeight: 700, color: 'var(--dex-gray-500)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 4 }}>Organizer</div>
                      <OrganizerList names={data.organizers} emails={data.organizerEmails} size="sm" />
                    </div>
                  )}
                  {!data.unlimitedParticipants && data.maxParticipants > 0 && (
                    <div style={{ marginTop: 10, fontSize: '0.82rem', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600 }}>
                      {data.maxParticipants} / {data.maxParticipants} Plätze frei
                    </div>
                  )}
                  {data.description && (
                    <div style={{ marginTop: 12, padding: 12, background: 'var(--dex-gray-50)', borderRadius: 8, fontSize: '0.85rem', lineHeight: 1.5, color: 'var(--dex-gray-700)' }}>
                      {data.description}
                    </div>
                  )}
                </div>
              </div>
            </section>

            {/* RECHTS: Anmeldeformular (disabled) */}
            <section>
              <div style={{ fontSize: '0.7rem', fontWeight: 700, textTransform: 'uppercase', letterSpacing: 0.5, color: 'var(--dex-green-dark, #4a7c1f)', marginBottom: 8 }}>
                Anmeldeformular
              </div>
              <div style={{
                background: '#fff', borderRadius: 10, padding: 16,
                border: '1px solid var(--dex-gray-200)',
              }}>
                {/* Standard-Felder */}
                <div style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-600)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 8 }}>
                  Persönliche Informationen
                </div>
                {[
                  { label: 'Anrede', value: 'Bitte wählen', required: true, type: 'select' },
                  { label: 'Vorname', value: 'Max', required: true, disabled: true },
                  { label: 'Nachname', value: 'Mustermann', required: true, disabled: true },
                  { label: 'E-Mail', value: 'max.mustermann@deloitte.de', required: true, disabled: true },
                ].map(f => (
                  <div key={f.label} style={{ marginBottom: 10 }}>
                    <label style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', display: 'block', marginBottom: 3 }}>
                      {f.required && <span style={{ color: 'var(--dex-red, #c9302c)', marginRight: 3 }}>*</span>}
                      {f.label}
                    </label>
                    <div style={{
                      border: '1px solid var(--dex-gray-300)', borderRadius: 6,
                      padding: '6px 10px', fontSize: '0.85rem',
                      color: 'var(--dex-gray-500)',
                      background: f.disabled ? 'var(--dex-gray-50)' : '#fff',
                    }}>
                      {f.value}
                    </div>
                  </div>
                ))}

                {/* Custom Fields */}
                {visibleFields.length > 0 && (
                  <>
                    <div style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-600)', textTransform: 'uppercase', letterSpacing: 0.5, marginTop: 16, marginBottom: 8 }}>
                      Event-spezifische Informationen
                    </div>
                    {visibleFields.map(f => (
                      <div key={f.id} style={{ marginBottom: 10 }}>
                        <label style={{ fontSize: '0.78rem', fontWeight: 600, color: 'var(--dex-gray-700)', display: 'block', marginBottom: 3 }}>
                          {f.required && <span style={{ color: 'var(--dex-red, #c9302c)', marginRight: 3 }}>*</span>}
                          {f.label}
                        </label>
                        {f.type === 'select' ? (
                          <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 6, padding: '6px 10px', fontSize: '0.85rem', color: 'var(--dex-gray-400)', display: 'flex', justifyContent: 'space-between' }}>
                            <span>Bitte wählen {f.options && f.options.length > 0 ? `(${f.options.length} Optionen)` : ''}</span>
                            <span>▾</span>
                          </div>
                        ) : f.type === 'checkbox' ? (
                          <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
                            <input type="checkbox" disabled style={{ width: 16, height: 16 }} />
                            <span>Zustimmen</span>
                          </div>
                        ) : f.type === 'textarea' ? (
                          <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 6, padding: '6px 10px', fontSize: '0.85rem', color: 'var(--dex-gray-400)', minHeight: 60 }}>
                            …
                          </div>
                        ) : f.type === 'date' ? (
                          <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 6, padding: '6px 10px', fontSize: '0.85rem', color: 'var(--dex-gray-400)' }}>
                            📅 TT.MM.JJJJ
                          </div>
                        ) : f.type === 'user' ? (
                          <>
                            <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 6, padding: '6px 10px', fontSize: '0.85rem', color: 'var(--dex-gray-400)' }}>
                              Name oder E-Mail suchen…
                            </div>
                            <div style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', marginTop: 4, fontStyle: 'italic' }}>
                              Die ausgewählte Person wird per E-Mail über deine Auswahl informiert.
                            </div>
                          </>
                        ) : (
                          <div style={{ border: '1px solid var(--dex-gray-300)', borderRadius: 6, padding: '6px 10px', fontSize: '0.85rem', color: 'var(--dex-gray-400)' }}>
                            …
                          </div>
                        )}
                      </div>
                    ))}
                  </>
                )}

                <button
                  type="button"
                  disabled
                  style={{
                    marginTop: 16, width: '100%', padding: '10px 16px',
                    background: 'var(--dex-green, #86bc25)', color: '#fff',
                    border: 'none', borderRadius: 8, fontSize: '0.95rem', fontWeight: 600,
                    cursor: 'not-allowed', opacity: 0.75,
                  }}
                >
                  Anmelden (Vorschau)
                </button>
              </div>
            </section>
          </div>
        </div>

        <div style={{
          display: 'flex', justifyContent: 'flex-end', gap: 8,
          padding: '12px 20px', borderTop: '1px solid var(--dex-gray-200)',
        }}>
          <button type="button" className="btn btn-primary" onClick={onClose}>Schließen</button>
        </div>
      </div>
    </div>
  );
};
