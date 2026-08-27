/**
 * v30.13 — Modularisierung Stufe 3, Pilot: Schritt „Abrechnung" (v29.66,
 * F&A-Pilot) als eigene Komponente. Der JSX-Block ist 1:1 aus
 * EventCreationPage übernommen; der Schritt hängt an genau sechs
 * State-Werten und ist damit der kleinste Props-Vertrag des Wizards.
 * `billingMissing` ist abgeleitet und wandert mit hierher (einzige
 * Nutzung). Die Sichtbarkeit steuert weiter der Wizard: das adminLike-
 * Gate bleibt am Aufrufer, `visible` ersetzt `currentStep === 9` —
 * display:none statt unmount, damit Eingaben beim Schrittwechsel
 * erhalten bleiben (gleiches Muster wie alle anderen Schritte).
 * Texte bewusst nur deutsch — der Pilot ist auf Admins begrenzt (isDe
 * kommt mit der Freischaltung für Organizer).
 */
import * as React from 'react';
import { BILLING_FIELDS } from '../../../data/billingFields';

export interface BillingStepProps {
  visible: boolean;
  billingRelevant: boolean | null;
  setBillingRelevant: (v: boolean) => void;
  billingSendMode: 'auto' | 'manual';
  setBillingSendMode: (v: 'auto' | 'manual') => void;
  billingFields: Record<string, string>;
  setBillingFields: React.Dispatch<React.SetStateAction<Record<string, string>>>;
}

export const BillingStep: React.FC<BillingStepProps> = ({
  visible, billingRelevant, setBillingRelevant,
  billingSendMode, setBillingSendMode, billingFields, setBillingFields,
}) => {
  const billingMissing = BILLING_FIELDS.filter(f => !(billingFields[f.id] || '').trim());
  return (
    <div style={{ display: visible ? 'block' : 'none' }}>
      <div style={{ background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12, padding: '16px 18px', marginBottom: 16, border: '1px solid var(--dex-gray-200)' }}>
        <label className="form-label" style={{ fontSize: '0.95rem', fontWeight: 700, marginBottom: 6 }}>
          Handelt es sich um ein abrechnungsrelevantes Event?
        </label>
        <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', margin: '0 0 10px' }}>
          Abrechnungsrelevante Events sind Veranstaltungen, deren Kosten oder
          Bewirtungsaufwendungen gegenüber Finance &amp; Accounting dokumentiert
          oder abgerechnet werden müssen. Das ist der Fall, wenn im Nachgang
          <strong> Rechnungen über die Kreditorenbuchhaltung eingereicht
          werden</strong> — etwa für Catering, eine externe Raumbuchung oder
          Anmeldegebühren (z.B. Startgelder für Läufer) — oder wenn für das
          Event <strong>Ariba-Bestellungen</strong> ausgelöst werden.
        </p>
        <div style={{ display: 'flex', gap: 18 }}>
          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 7, cursor: 'pointer', fontSize: '0.9rem' }}>
            <input type="radio" name="dexBillingRelevant" checked={billingRelevant === true} onChange={() => setBillingRelevant(true)} />
            Ja
          </label>
          <label style={{ display: 'inline-flex', alignItems: 'center', gap: 7, cursor: 'pointer', fontSize: '0.9rem' }}>
            <input type="radio" name="dexBillingRelevant" checked={billingRelevant === false} onChange={() => setBillingRelevant(false)} />
            Nein
          </label>
        </div>
      </div>

      {billingRelevant === true && (
        <>
          {/* Status — systemseitig aus den Pflichtfeldern abgeleitet,
              nie gespeichert und nie von Hand setzbar. */}
          <div style={{
            padding: '10px 14px', borderRadius: 8, marginBottom: 16, fontSize: '0.85rem',
            background: billingMissing.length > 0 ? 'rgba(237,139,0,0.10)' : 'rgba(134,188,37,0.12)',
            border: `1px solid ${billingMissing.length > 0 ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green, #86bc25)'}`,
            color: 'var(--dex-gray-800)',
          }}>
            {billingMissing.length > 0
              ? <><strong>Status: Abrechnungsrelevante Informationen unvollständig</strong> — {billingMissing.length} von {BILLING_FIELDS.length} Pflichtfeldern fehlen noch. Speichern ist trotzdem möglich.</>
              : <><strong>Status: Vollständig</strong> — alle {BILLING_FIELDS.length} Pflichtangaben sind gepflegt.</>}
          </div>

          <div style={{ background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12, padding: '16px 18px', marginBottom: 16, border: '1px solid var(--dex-gray-200)' }}>
            <label className="form-label" style={{ fontSize: '0.95rem', fontWeight: 700, marginBottom: 6 }}>
              Informationen zur Abrechnung
            </label>
            <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', margin: '0 0 12px' }}>
              Abrechnungsrelevante Informationen müssen an die Finance &amp; Accounting
              Abteilung gemeldet werden. Dies beinhaltet insbesondere allgemeine
              Eventinformationen, Teilnehmerlisten sowie Rechnungen und Belege. Die
              folgenden Einstellungen unterstützen die standardisierte und teilweise
              automatisierte Übermittlung dieser Informationen.
            </p>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer', marginBottom: 10 }}>
              <input type="radio" name="dexBillingSend" checked={billingSendMode === 'auto'} onChange={() => setBillingSendMode('auto')} style={{ marginTop: 3 }} />
              <span style={{ fontSize: '0.88rem' }}>
                <strong>Automatisierter Versand</strong>
                <span style={{ display: 'block', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                  Abrechnungsinformationen 7 Kalendertage vor dem Event (bei kurzfristiger
                  Erstellung: sofort nach Aktivierung), finale Teilnehmerliste 7 Kalendertage
                  danach — jeweils an F&amp;A, Organizer in CC.
                  <em style={{ display: 'block', marginTop: 2, color: 'var(--dex-orange, #b96a00)' }}>
                    Pilot: Die Auswahl wird bereits gespeichert, der Automatik-Flow existiert noch nicht.
                  </em>
                </span>
              </span>
            </label>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
              <input type="radio" name="dexBillingSend" checked={billingSendMode === 'manual'} onChange={() => setBillingSendMode('manual')} style={{ marginTop: 3 }} />
              <span style={{ fontSize: '0.88rem' }}>
                <strong>Manueller Versand</strong> <span style={{ color: 'var(--dex-gray-500)', fontWeight: 400 }}>(Standard)</span>
                <span style={{ display: 'block', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                  Kein automatischer Versand — Abrechnungsinformationen und Teilnehmerliste
                  werden über das Organizer Center aktiv an F&amp;A gesendet.
                </span>
              </span>
            </label>
          </div>

          <div style={{ background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12, padding: '16px 18px', marginBottom: 16, border: '1px solid var(--dex-gray-200)' }}>
            <label className="form-label" style={{ fontSize: '0.95rem', fontWeight: 700, marginBottom: 2 }}>
              Abrechnungsrelevante Informationen
            </label>
            {/* v30.4: Legende — die Sternchen standen unerklärt im Raum. */}
            <p style={{ fontSize: '0.74rem', color: 'var(--dex-gray-500)', margin: '0 0 10px' }}>
              <span className="required">*</span> Pflichtangabe — ohne sie gilt die Abrechnungsmeldung an Finance &amp; Accounting als unvollständig. Speichern kannst du trotzdem jederzeit.
            </p>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '12px 16px' }}>
              {BILLING_FIELDS.map(f => {
                const val = billingFields[f.id] || '';
                const empty = !val.trim();
                const setVal = (v: string): void => setBillingFields(prev => ({ ...prev, [f.id]: v }));
                return (
                  // v30.4: Flex-Spalte, Label wächst — die Eingabefelder
                  // einer Zeile stehen damit auf gleicher Höhe, auch wenn
                  // ein Label („Name der Veranstaltung bzw. Anlass …")
                  // zweizeilig umbricht.
                  <div key={f.id} style={{ display: 'flex', flexDirection: 'column' }}>
                    <label style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', display: 'block', marginBottom: 3, flexGrow: 1 }}>
                      {f.label} <span className="required">*</span>
                    </label>
                    {f.type === 'select' ? (
                      <select
                        className="form-input"
                        value={val}
                        onChange={e => setVal(e.target.value)}
                        style={{ borderColor: empty ? 'var(--dex-orange, #ed8b00)' : undefined }}
                      >
                        <option value="">Bitte wählen…</option>
                        {(f.options || []).map(o => <option key={o} value={o}>{o}</option>)}
                      </select>
                    ) : (
                      <input
                        className="form-input"
                        type={f.type === 'date' ? 'date' : 'text'}
                        value={val}
                        onChange={e => setVal(e.target.value)}
                        style={{ borderColor: empty ? 'var(--dex-orange, #ed8b00)' : undefined }}
                      />
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        </>
      )}
    </div>
  );
};
