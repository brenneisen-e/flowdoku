/* BillingStatusStrip — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 7892-7971 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { FA_STATUS_COLORS, FA_STATUS_LABELS, FA_STATUS_NEXT, faStatusOf, missingBillingFields, parseBillingOf } from '../../../utils/faBilling';
import { Icon } from '@fluentui/react/lib/Icon';
import { BILLING_FIELDS, canEditBilling } from '../../../data/billingFields';
import { DeloitteEvent } from '../../../types';

export interface BillingStatusStripProps {
  isAdmin: boolean;
  isDe: boolean;
  isFA: boolean;
  isOrganizerFor: (ev: DeloitteEvent) => boolean;
  isQRScannerOnlyForSelected: boolean;
  navigate: (page: import("../../../context/NavigationContext").Page, eventId?: string, intent?: import("../../../context/NavigationContext").NavIntent) => void;
  selectedEvent: DeloitteEvent;
  setBillingPanelOpen: React.Dispatch<React.SetStateAction<boolean>>;
}

export const BillingStatusStrip: React.FC<BillingStatusStripProps> = (p) => {
  const { isAdmin, isDe, isFA, isOrganizerFor, isQRScannerOnlyForSelected, navigate, selectedEvent, setBillingPanelOpen } = p;
          const bb = parseBillingOf(selectedEvent);
          if (!bb) return null;
          const miss = missingBillingFields(bb);
          const incomplete = miss.length > 0;
          const st = faStatusOf(selectedEvent, bb);
          const stColors = FA_STATUS_COLORS[st];
          return (
            <div style={{
              marginBottom: 16, padding: '12px 16px', borderRadius: 10,
              border: `1px solid ${incomplete ? 'var(--dex-orange, #ed8b00)' : 'var(--dex-green, #86bc25)'}`,
              background: incomplete ? 'rgba(237,139,0,0.07)' : 'rgba(134,188,37,0.07)',
              display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap',
            }}>
              <Icon iconName={incomplete ? 'Warning' : 'CheckMark'} style={{ fontSize: 16, color: incomplete ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-green-dark, #4a7c1f)' }} />
              <div style={{ flex: 1, minWidth: 240, fontSize: '0.85rem', lineHeight: 1.5, color: 'var(--dex-gray-800)' }}>
                <strong>{isDe ? 'Abrechnungsrelevantes Event' : 'Billing-relevant event'}</strong>
                <span style={{
                  marginLeft: 8, padding: '1px 8px', borderRadius: 999, fontSize: '0.7rem', fontWeight: 700,
                  background: stColors.bg, color: stColors.fg,
                }}>{FA_STATUS_LABELS[st]}</span>
                {/* v30.47: Was als NAECHSTES zu tun ist — direkt unter dem
                    Status. Ein Status allein sagt, wo man steht; er sagt nicht,
                    ob man selbst dran ist oder wartet. Genau das war bei
                    „vollstaendig, Versendung ausstehend" die Frage. */}
                {FA_STATUS_NEXT[st] && (
                  <div style={{ marginTop: 4, fontSize: '0.8rem', fontWeight: 600, color: stColors.fg }}>
                    {FA_STATUS_NEXT[st]}
                  </div>
                )}
                <div style={{ marginTop: 4, color: 'var(--dex-gray-700)' }}>
                  {incomplete
                    ? (isDe
                      ? `Es fehlen noch ${miss.length} von ${BILLING_FIELDS.length} Pflichtangaben für die Meldung an Finance & Accounting: ${miss.map(f => f.label).slice(0, 4).join(', ')}${miss.length > 4 ? ` +${miss.length - 4} weitere` : ''}. Solange sie fehlen, kannst du die Abrechnungsinformationen nicht versenden — Speichern und Aktivieren des Events sind davon nicht betroffen.`
                      : `${miss.length} of ${BILLING_FIELDS.length} required details for the Finance & Accounting report are still missing: ${miss.map(f => f.label).slice(0, 4).join(', ')}${miss.length > 4 ? ` +${miss.length - 4} more` : ''}. Until they are filled in, the billing information cannot be sent — saving and activating the event are unaffected.`)
                    : (isDe
                      ? `Alle ${BILLING_FIELDS.length} Pflichtangaben sind gepflegt. Über „Event-Abrechnung" schickst du Abrechnungsinformationen oder die Teilnehmerliste an F&A.`
                      : `All ${BILLING_FIELDS.length} required details are filled in. Use „Event billing" to send the billing information or the participant list to F&A.`)}
                </div>
              </div>
              <div style={{ display: 'inline-flex', gap: 8, flexWrap: 'wrap' }}>
                {incomplete && (
                  <button
                    type="button"
                    className="btn btn-outline"
                    style={{ fontSize: '0.78rem', padding: '4px 12px', color: 'var(--dex-orange, #ed8b00)', borderColor: 'var(--dex-orange, #ed8b00)' }}
                    onClick={() => {
                      // v30.44: Direkt in Schritt 10 „Abrechnung" statt auf
                      // Schritt 1. Wer auf „Angaben ergänzen" klickt, weiß
                      // genau, welche Angaben gemeint sind — ihn erst durch
                      // neun Schritte zu schicken ist kein Weg, sondern eine
                      // Suchaufgabe. Der Wizard liest die Marke EINMAL beim
                      // Mount und räumt sie danach selbst ab.
                      // v30.46: Nur setzen, wenn der Wizard den Schritt fuer
                      // diese Person ueberhaupt rendert — sonst landet sie auf
                      // einem Index, den es fuer sie nicht gibt, und sieht ein
                      // leeres Formular statt einer Meldung. Dieselbe Ableitung
                      // wie im Wizard (`canEditBilling`), damit beide Seiten
                      // nicht auseinanderlaufen koennen: Ein Schalter,
                      // FA_BILLING_STEP_FOR_ORGANIZERS, oeffnet beides zugleich.
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      if (canEditBilling(isAdmin, isOrganizerFor(selectedEvent), isFA)) { try { (window as any).__dexPreviewInitialStep = 9; } catch { /* */ } }
                      navigate('edit-event', selectedEvent.id);
                    }}
                  >
                    {isDe ? 'Angaben ergänzen' : 'Complete details'}
                  </button>
                )}
                <button
                  type="button"
                  className="btn btn-secondary"
                  style={{ fontSize: '0.78rem', padding: '4px 12px' }}
                  onClick={() => setBillingPanelOpen(true)}
                >
                  {isDe ? 'Event-Abrechnung öffnen' : 'Open event billing'}
                </button>
              </div>
            </div>
          );
};

