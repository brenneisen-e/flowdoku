/* AdminToast — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 6783-6842 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich übernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.3: Nach docs/ui-leitfaden.md umgebaut — ruhige weiße Karte statt farbiger
 * Kante und farbiger Überschrift; die Art der Meldung trägt ein `dex-ui-dot`
 * (bzw. der Spinner, solange die Abmeldung läuft), Schließen ist ein
 * `dex-ui-iconbtn` statt eines nackten „×". Position und isMobile-Zweig bleiben
 * unverändert. Die Sprache kommt über `useLocaleSafe`, weil die Props
 * unverändert bleiben müssen (derselbe Weg wie im OverbookDecisionModal) — bis
 * v31.2 war der Text nur deutsch, obwohl die App zweisprachig ist.
 */
import * as React from 'react';
import { cx, ensureDexUiStyles } from '../../dexUi';
import { X } from '../../Icons';
import { useLocaleSafe } from '../../../context/LanguageContext';

export interface AdminToastProps {
  adminToast: { kind: "cancelling"; name: string; } | { kind: "promoted"; name: string; email: string; type?: string; } | { kind: "no-promote"; name: string; };
  isMobile: boolean;
  setAdminToast: React.Dispatch<React.SetStateAction<{ kind: "cancelling"; name: string; } | { kind: "promoted"; name: string; email: string; type?: string; } | { kind: "no-promote"; name: string; }>>;
}

export const AdminToast: React.FC<AdminToastProps> = (p) => {
  const { adminToast, isMobile, setAdminToast } = p;
  const isDe = useLocaleSafe() === 'de';
  // Idempotent — die Meldung schwebt über der Seite, ohne dass ein Modal oder
  // die WizardFormShell das Stylesheet schon eingehängt hätte.
  ensureDexUiStyles();
  const t = (de: string, en: string): string => (isDe ? de : en);
  // Solange die Abmeldung läuft, darf sie nicht weggeklickt werden (unverändert).
  const closable = adminToast.kind !== 'cancelling';
  // v31.3: Die Art der Meldung trägt der Punkt — grün „nachgerückt", grau
  // „niemand nachgerückt"; während der Abmeldung steht der Spinner an seiner
  // Stelle. Die Karte bleibt weiß (Leitfaden 1.1: Farbe nur, wo sie etwas
  // bedeutet) — vorher waren Rahmen, Kante und Überschrift eingefärbt.
        return (
          <div className="dex-ui-card dex-ui-fade-in" role="status" aria-live="polite" style={{
            position: 'fixed', top: 80, zIndex: 1000,
            ...(isMobile
              ? { left: 12, right: 12, maxWidth: 'min(460px, calc(100vw - 24px))' }
              : { right: 20, maxWidth: 460 }),
            boxShadow: '0 10px 30px rgba(0,0,0,0.14)',
            display: 'flex', alignItems: 'flex-start', gap: 12,
          }}>
            {adminToast.kind === 'cancelling' ? (
              <div style={{
                width: 18, height: 18, marginTop: 2, flexShrink: 0,
                border: `3px solid var(--dex-gray-200)`,
                borderTopColor: 'var(--dex-orange, #ed8b00)',
                borderRadius: '50%',
                animation: 'dex-spin 0.8s linear infinite',
              }} />
            ) : (
              <span className={cx('dex-ui-dot', adminToast.kind === 'promoted' && 'dex-ui-dot--green')} style={{ marginTop: 7 }} aria-hidden="true" />
            )}
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontWeight: 700, fontSize: '0.9rem', color: 'var(--dex-gray-800)' }}>
                {adminToast.kind === 'cancelling' && t(`Abmeldung von ${adminToast.name} wird verarbeitet…`, `Cancelling ${adminToast.name}…`)}
                {adminToast.kind === 'promoted' && t(`Nachgerückt: ${adminToast.name}`, `Moved up from the waitlist: ${adminToast.name}`)}
                {adminToast.kind === 'no-promote' && t(`Abmeldung von ${adminToast.name} verarbeitet`, `Cancellation for ${adminToast.name} done`)}
              </div>
              <div style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.5, marginTop: 3 }}>
                {adminToast.kind === 'cancelling' && t('Teilnehmer wird abgemeldet, Warteliste wird geprüft und ggf. ein Nachrücker informiert.', 'The attendee is being removed, the waitlist is checked and the next person is notified if there is one.')}
                {adminToast.kind === 'promoted' && (
                  <>
                    <strong>{adminToast.email}</strong>{adminToast.type ? ` (${adminToast.type})` : ''}
                    {t(' ist automatisch von der Warteliste nachgerückt. Nachrück-Mail und Outlook-Einladung sind raus.', ' was moved up from the waitlist automatically. The confirmation mail and the Outlook invitation have been sent.')}
                  </>
                )}
                {adminToast.kind === 'no-promote' && t('Aktuell ist niemand auf der Warteliste (bzw. kein passender Starter-Typ). Der Platz bleibt frei.', 'Nobody is on the waitlist right now (or nobody with a matching starter type). The spot stays free.')}
              </div>
            </div>
            {closable && (
              <button type="button" className="dex-ui-iconbtn" onClick={() => setAdminToast(null)}
                aria-label={t('Schließen', 'Close')} title={t('Schließen', 'Close')}
                style={{ marginTop: -4, marginRight: -6 }}><X size={16} /></button>
            )}
          </div>
        );
};

