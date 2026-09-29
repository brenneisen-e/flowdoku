/**
 * „Über die App" — was die Plattform ist und wie man sie benutzt.
 *
 * Vorbild ist `LandingInfoModal` in DEX (Wofür ist die App da, wie läuft es ab,
 * wer darf was). Das Tutorial dort ist hier nicht übernommen: Eine Plattform
 * mit einer Kachelwand und einem Start-Knopf braucht keine geführte Tour, und
 * ein Tutorial, das über die Oberfläche hinwegführt, wäre die schwerste
 * Komponente des Projekts für den geringsten Nutzen.
 *
 * Der Dialog beantwortet drei Fragen in der Reihenfolge, in der man sie
 * stellt: Was ist das? Wie funktioniert es? Was darf ich hier?
 */

import * as React from 'react';
import Modal from './Modal';
import { useLanguage } from '../context/LanguageContext';
import { useRoles } from '../context/RoleContext';
import { rolleLabel } from '../utils/rollen';
import { APP_NAME, APP_SUBTITLE_DE, APP_SUBTITLE_EN } from '../constants';
import { APP_VERSION } from '../version';
import { UserRole } from '../types';

export default function AboutModal(props: { open: boolean; onClose: () => void; onAsk: () => void }): React.ReactElement | null {
  const { open, onClose, onAsk } = props;
  const { t, isDe } = useLanguage();
  const { currentUserRole, isRolesLoading } = useRoles();

  const SCHRITTE: Array<{ titel: string; text: string }> = [
    { titel: t('Use Case finden', 'Find a use case'), text: t('Unter „Use Cases" stehen alle Agent-Demos als Kacheln. Suche und Bereichsfilter helfen, wenn es viele werden.', 'Under "Use cases" all agent demos appear as tiles. Search and area filter help once there are many.') },
    { titel: t('Kurzbeschreibung lesen', 'Read the summary'), text: t('Die Detailseite sagt, was die Demo zeigt, wie ausgereift sie ist und was du davor wissen solltest.', 'The detail page says what the demo shows, how mature it is and what to know beforehand.') },
    { titel: t('Demo starten', 'Start the demo'), text: t('Ein Klick öffnet die laufende Demo. Dazu gibt es Source Code, Deployment-Guide, Wiki und Video, soweit das Team sie geliefert hat.', 'One click opens the running demo. Source code, deployment guide, wiki and video are there as far as the team supplied them.') },
  ];

  const ROLLEN: Array<{ r: UserRole; text: string }> = [
    { r: 'User', text: t('Use Cases ansehen und Demos starten.', 'View use cases and start demos.') },
    { r: 'Organizer', text: t('Zusätzlich Use Cases anlegen und pflegen — mit Bild, Links und Bewertung.', 'Also create and maintain use cases — with image, links and assessment.') },
    { r: 'Admin', text: t('Zusätzlich Rollen vergeben und Rechte prüfen.', 'Also assign roles and check rights.') },
  ];

  return (
    <Modal
      open={open}
      onClose={onClose}
      maxWidth={600}
      ariaLabel={t('Über die App', 'About the app')}
      title={t('Über die App', 'About the app')}
      subtitle={`${APP_NAME} · ${isDe ? APP_SUBTITLE_DE : APP_SUBTITLE_EN}`}
      footer={<>
        <button type="button" className="btn btn-secondary" onClick={() => { onClose(); onAsk(); }}>{t('Frage stellen', 'Ask a question')}</button>
        <button type="button" className="btn btn-primary" onClick={onClose}>{t('Schließen', 'Close')}</button>
      </>}
    >
      <div className="dex-ui-section" style={{ margin: 0 }}>
        <div className="dex-ui-section-title">{t('Wofür ist die Plattform da?', 'What is the platform for?')}</div>
        <p style={{ margin: 0 }}>
          {t('Alle Agent-Demos für den Banking-Sektor an einer Stelle: von der Idee über die Demo im Bau bis zur laufenden Anwendung. Entstanden aus der Use-Case-Analyse und dem internen Hackathon.',
            'All agent demos for the banking sector in one place: from the idea through the demo under construction to the running application. Born from the use case analysis and the internal hackathon.')}
        </p>
      </div>

      <div className="dex-ui-section" style={{ margin: 0 }}>
        <div className="dex-ui-section-title">{t('So funktioniert es', 'How it works')}</div>
        <div className="dex-ui-stack">
          {SCHRITTE.map((s, i) => (
            <div key={s.titel} className="dex-ui-step">
              <span className="dex-ui-step-num">{i + 1}</span>
              <span className="dex-ui-step-body">
                <span className="dex-ui-step-title" style={{ display: 'block' }}>{s.titel}</span>
                <span className="dex-ui-step-hint" style={{ display: 'block' }}>{s.text}</span>
              </span>
            </div>
          ))}
        </div>
      </div>

      <div className="dex-ui-section" style={{ margin: 0 }}>
        <div className="dex-ui-section-title">{t('Wer darf was?', 'Who may do what?')}</div>
        <div className="dex-ui-stack">
          {ROLLEN.map(x => (
            <div key={x.r} className="dex-ui-row dex-ui-row--static">
              <span className="dex-ui-row-main">
                <span className="dex-ui-row-title dex-ui-row-title--wrap">
                  {rolleLabel(x.r)}
                  {!isRolesLoading && currentUserRole === x.r && (
                    <span className="dex-ui-pill dex-ui-pill--sm dex-ui-pill--green">{t('deine Rolle', 'your role')}</span>
                  )}
                </span>
                <span className="dex-ui-row-sub">{x.text}</span>
              </span>
            </div>
          ))}
        </div>
      </div>

      <p className="dex-ui-help" style={{ margin: 0 }}>{APP_NAME} v{APP_VERSION}</p>
    </Modal>
  );
}
