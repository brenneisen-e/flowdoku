/* BillingStatusStrip — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 7892-7971 des
 * Stands vor dem Schnitt). Die Ableitung des Status ist unverändert; seit
 * v31.3 ist die Anzeige eine Zeile (Punkt · Pille · Knöpfe daneben) statt
 * eines Kastens mit Absatz. Die Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { FA_STATUS_COLORS, FA_STATUS_LABELS, FA_STATUS_NEXT, FA_STATUS_SHORT, faStatusOf, missingBillingFields, parseBillingOf } from '../../../utils/faBilling';
import { Icon } from '@fluentui/react/lib/Icon';
import { BILLING_FIELDS, canEditBilling } from '../../../data/billingFields';
import { DeloitteEvent } from '../../../types';
import { cx } from '../../dexUi';
import { InfoTooltip } from '../../InfoTooltip';

export interface BillingStatusStripProps {
  isAdmin: boolean;
  isDe: boolean;
  isFA: boolean;
  isOrganizerFor: (ev: DeloitteEvent) => boolean;
  navigate: (page: import("../../../context/NavigationContext").Page, eventId?: string, intent?: import("../../../context/NavigationContext").NavIntent) => void;
  selectedEvent: DeloitteEvent;
  setBillingPanelOpen: React.Dispatch<React.SetStateAction<boolean>>;
}

export const BillingStatusStrip: React.FC<BillingStatusStripProps> = (p) => {
  const { isAdmin, isDe, isFA, isOrganizerFor, navigate, selectedEvent, setBillingPanelOpen } = p;
          const bb = parseBillingOf(selectedEvent);
          if (!bb) return null;
          const miss = missingBillingFields(bb);
          const incomplete = miss.length > 0;
          const st = faStatusOf(selectedEvent, bb);
          const stColors = FA_STATUS_COLORS[st];
          // v31.3: Sichtbar bleibt, was fehlt; die Folge („dann geht kein
          // Versand, Speichern aber schon") steht im Tooltip — sonst sind es
          // vier Zeilen Fließtext über dem Knopf, der sie auflöst.
          const detail = incomplete
            ? (isDe
              ? `Es fehlen noch ${miss.length} von ${BILLING_FIELDS.length} Pflichtangaben für die Meldung an Finance & Accounting: ${miss.map(f => f.label).slice(0, 4).join(', ')}${miss.length > 4 ? ` +${miss.length - 4} weitere` : ''}.`
              : `${miss.length} of ${BILLING_FIELDS.length} required details for the Finance & Accounting report are still missing: ${miss.map(f => f.label).slice(0, 4).join(', ')}${miss.length > 4 ? ` +${miss.length - 4} more` : ''}.`)
            : (isDe
              ? `Alle ${BILLING_FIELDS.length} Pflichtangaben sind gepflegt. Über „Event-Abrechnung" schickst du Abrechnungsinformationen oder die Teilnehmerliste an F&A.`
              : `All ${BILLING_FIELDS.length} required details are filled in. Use „Event billing" to send the billing information or the participant list to F&A.`);
          return (
            <div
              className={cx('dex-ui-callout', incomplete ? 'dex-ui-callout--warn' : 'dex-ui-callout--success')}
              style={{ marginBottom: 16, flexDirection: 'column', gap: 5 }}
            >
              {/* Zeile 1: worum geht es, und wo steht das Event? Die Kurzform
                  in der Pille, der volle Konzept-Wortlaut im Titel. */}
              <div className="dex-ui-inline">
                <Icon iconName={incomplete ? 'Warning' : 'CheckMark'} style={{ fontSize: 15, color: incomplete ? 'var(--dex-orange-dark, #b35a00)' : 'var(--dex-green-dark, #4a7c1f)' }} />
                <strong style={{ fontSize: '0.88rem', color: 'var(--dex-gray-800)' }}>{isDe ? 'Abrechnungsrelevantes Event' : 'Billing-relevant event'}</strong>
                <span className="dex-ui-pill" style={{ background: stColors.bg, color: stColors.fg }} title={FA_STATUS_LABELS[st]}>
                  <span className="dex-ui-dot" style={{ background: stColors.fg }} />{FA_STATUS_SHORT[st]}
                </span>
                {incomplete && (
                  <InfoTooltip text={isDe
                    ? 'Solange Pflichtangaben fehlen, kannst du die Abrechnungsinformationen nicht versenden — Speichern und Aktivieren des Events sind davon nicht betroffen.'
                    : 'While required details are missing, the billing information cannot be sent — saving and activating the event are unaffected.'} />
                )}
              </div>
              {/* v30.47: Was als NAECHSTES zu tun ist — direkt unter dem
                  Status. Ein Status allein sagt, wo man steht; er sagt nicht,
                  ob man selbst dran ist oder wartet. Genau das war bei
                  „vollstaendig, Versendung ausstehend" die Frage. */}
              {FA_STATUS_NEXT[st] && (
                <div style={{ fontSize: '0.82rem', fontWeight: 600, color: stColors.fg }}>{FA_STATUS_NEXT[st]}</div>
              )}
              <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-700)' }}>{detail}</div>
              {/* v31.3: Die Knoepfe stehen links beim Text, nicht am rechten
                  Rand — sie gehoeren zu dem, was daneben steht (Leitfaden 2a′). */}
              <div className="dex-ui-inline" style={{ marginTop: 3 }}>
                {incomplete && (
                  <button
                    type="button"
                    className="btn btn-outline dex-ui-btn-sm"
                    style={{ color: 'var(--dex-orange, #ed8b00)', borderColor: 'var(--dex-orange, #ed8b00)' }}
                    onClick={() => {
                      // v30.44: Direkt in Schritt 10 „Abrechnung" statt auf
                      // Schritt 1. Wer auf „Angaben ergänzen" klickt, weiß
                      // genau, welche Angaben gemeint sind — ihn erst durch
                      // neun Schritte zu schicken ist kein Weg, sondern eine
                      // Suchaufgabe. Der Wizard liest die Marke EINMAL beim
                      // Mount und räumt sie danach selbst ab.
                      // v30.46: Nur setzen, wenn der Wizard den Schritt für
                      // diese Person ueberhaupt rendert — sonst landet sie auf
                      // einem Index, den es für sie nicht gibt, und sieht ein
                      // leeres Formular statt einer Meldung. Dieselbe Ableitung
                      // wie im Wizard (`canEditBilling`), damit beide Seiten
                      // nicht auseinanderlaufen können: Ein Schalter,
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
                  className="btn btn-secondary dex-ui-btn-sm"
                  onClick={() => setBillingPanelOpen(true)}
                >
                  {isDe ? 'Event-Abrechnung öffnen' : 'Open event billing'}
                </button>
              </div>
            </div>
          );
};

