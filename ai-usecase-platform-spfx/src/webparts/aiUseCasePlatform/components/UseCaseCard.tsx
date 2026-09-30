/**
 * Eine Use-Case-Kachel.
 *
 * Aufbau nach dem DEX-Muster: Bild oben, Titel, eine Zeile Folge, darunter
 * die Meta-Angaben als Pillen. Der Hover kommt aus `dex-ui-card--hover` —
 * ein Inline-Style kann kein `:hover`, das ist der Grund, warum es den
 * Klassensatz ueberhaupt gibt.
 *
 * Die ganze Kachel ist EIN Knopf. Zwei Klickziele nebeneinander („Details"
 * und „Starten") waeren zwei Bedienwege fuer dieselbe Sache; die Detailseite
 * traegt den Start-Knopf.
 */

import * as React from 'react';
import { cx } from './dexUi';
import { UseCase, Bewertung } from '../types';
import { useLanguage } from '../context/LanguageContext';
import { sichereUrl } from '../utils/sicher';
import { statusText } from '../utils/anzeige';

/** Kuerzel aus dem Titel, wenn kein Bild hinterlegt ist. */
function kuerzel(titel: string): string {
  const worte = (titel || '?').split(/[\s/-]+/).filter(Boolean);
  // Ein Titel wie „-" oder „/" bleibt ohne Wort — der Zugriff darauf warf und ließ die ganze
  // Kachelwand weiß (Review 29.09.2026).
  if (worte.length === 0) return '?';
  if (worte.length === 1) return worte[0].slice(0, 2).toUpperCase();
  return (worte[0][0] + worte[1][0]).toUpperCase();
}

/** Die Bewertung als Punkte — wie im Konzept-Deck (●●● / ●●○ / ●○○). */
export function bewertungPunkte(b: Bewertung): string {
  if (b === 'Hoch') return '●●●';
  if (b === 'Mittel') return '●●○';
  if (b === 'Niedrig') return '●○○';
  return '○○○';
}

export interface UseCaseCardProps {
  useCase: UseCase;
  onOpen: (id: number) => void;
}

export default function UseCaseCard(props: UseCaseCardProps): React.ReactElement {
  const { useCase: uc, onOpen } = props;
  const { t, isDe } = useLanguage();

  const istLive = uc.status === 'Live';
  const statusLabel = statusText(uc.status, t);
  // Die Adresse kommt aus der Liste, die auch direkt beschreibbar ist: nur http(s).
  const bild = sichereUrl(uc.bildUrl);

  return (
    <button
      type="button"
      // `--muted` daempft, was noch nicht aufrufbar ist — sichtbar, aber
      // erkennbar noch nicht so weit. Die Kachel bleibt klickbar: Auch ein
      // geplanter Use Case hat eine Beschreibung, die jemanden interessiert.
      className={cx('dex-ui-card', 'dex-ui-card--hover', !istLive && 'dex-ui-card--muted')}
      style={{ padding: 0, overflow: 'hidden', textAlign: 'left', cursor: 'pointer', display: 'flex', flexDirection: 'column' }}
      onClick={() => onOpen(uc.id)}
      aria-label={`${uc.titel} — ${statusLabel}`}
    >
      {/* Kopf: ein KLEINES abgerundetes Quadrat (Bildschirmfoto bzw. Kürzel), daneben Bereich und Titel,
          rechts der Status. Bis v1.3 stand oben ein großes 16:9-Bild — bei 18 Kacheln erschlug es die
          Wand (Nutzer-Ansage 30.09.2026: „klein, rund oder quadratisch"). */}
      <span style={{ display: 'flex', alignItems: 'flex-start', gap: 12, padding: '14px 14px 0' }}>
        <span
          aria-hidden="true"
          style={{
            display: 'flex', alignItems: 'center', justifyContent: 'center', flex: '0 0 auto',
            width: 64, height: 64, borderRadius: 16, overflow: 'hidden',
            border: '1px solid var(--dex-gray-200, #e8e8e8)',
            // Anführungszeichen in der Adresse würden den `url("…")`-Wert aufbrechen.
            background: bild
              ? `center/cover no-repeat url("${bild.replace(/"/g, '%22')}")`
              : 'linear-gradient(135deg, rgba(134,188,37,0.16), rgba(134,188,37,0.05))',
            fontSize: '1.2rem', fontWeight: 700, color: 'var(--dex-green-dark, #6b9a1e)', letterSpacing: '0.04em',
          }}
        >{!bild && kuerzel(uc.titel)}</span>
        <span style={{ display: 'block', flex: 1, minWidth: 0 }}>
          {uc.bereich && (
            <span className="dex-ui-muted" style={{ display: 'block', fontSize: '0.72rem', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 4 }}>
              {uc.bereich}
            </span>
          )}
          <span className="dex-ui-tile-title" style={{ display: 'block', fontSize: '1rem' }}>{uc.titel}</span>
        </span>
        <span className={cx('dex-ui-pill', 'dex-ui-pill--sm', istLive ? 'dex-ui-pill--green' : 'dex-ui-pill--gray')} style={{ flex: '0 0 auto' }}>{statusLabel}</span>
      </span>

      <span style={{ display: 'block', padding: '10px 14px 14px', flex: 1 }}>
        <span className="dex-ui-tile-desc" style={{ display: 'block' }}>{uc.kurzbeschreibung}</span>

        {/* Die drei Bewertungen aus der Use-Case-Analyse. Mit Titel-Attribut
            UND sichtbarer Beschriftung — auf dem Handy gibt es kein
            Ueberfahren, ein Wert nur im `title` waere dort unerreichbar. */}
        <span className="dex-ui-meta" style={{ marginTop: 10 }}>
          <span className="dex-ui-meta-item" title={t('Sales-Relevanz', 'Sales relevance')}>
            {t('Sales', 'Sales')} {bewertungPunkte(uc.salesRelevanz)}
          </span>
          <span className="dex-ui-meta-item" title={t('Machbarkeit', 'Feasibility')}>
            {t('Machbarkeit', 'Feasibility')} {bewertungPunkte(uc.machbarkeit)}
          </span>
          <span className="dex-ui-meta-item" title={t('Demo-Tauglichkeit', 'Demo readiness')}>
            {t('Demo', 'Demo')} {bewertungPunkte(uc.demoTauglichkeit)}
          </span>
        </span>

        {!istLive && (
          <span className="dex-ui-help" style={{ display: 'block', marginTop: 8 }}>
            {uc.status === 'InArbeit'
              ? t('Wird gerade gebaut — noch nicht aufrufbar.', 'Being built — not callable yet.')
              : uc.status === 'Geplant'
                ? t('Geplant — es gibt noch keine Demo.', 'Planned — there is no demo yet.')
                : t('Archiviert.', 'Archived.')}
          </span>
        )}
        {istLive && !uc.ressourcen.deployment && (
          // Ehrlich benennen statt eine tote Kachel anbieten: „Live" ohne
          // Link ist ein Pflegefehler, kein Zustand, den der Nutzer versteht.
          <span className="dex-ui-help" style={{ display: 'block', marginTop: 8 }}>
            {isDe ? 'Steht auf „Live", aber es ist kein Deployment-Link hinterlegt.' : 'Marked "Live", but no deployment link is stored.'}
          </span>
        )}
      </span>
    </button>
  );
}
