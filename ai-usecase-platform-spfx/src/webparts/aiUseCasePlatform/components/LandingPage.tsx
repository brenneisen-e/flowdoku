/**
 * Der Startbildschirm — Aufbau wie `LandingPage` in DEX.
 *
 * Von oben nach unten: der Orb, die Begrüßung mit Vorname, ein Willkommens-
 * text, EINE Hinweiskarte, der Knopf „Start", für Nutzer ohne Organizer-Rolle
 * die Einladung dazu.
 *
 * Nutzer-Ansage 29.09.2026: „Nimm nochmal DEX als Vorlage mit einer
 * Landingpage wie hier" (Screenshot der DEX-Startseite). Die Hinweiskarte ist
 * dort „Du bist für <Event> angemeldet" mit Countdown; hier steht an ihrer
 * Stelle, was hinter „Start" wartet: wie viele Use Cases es gibt und wie
 * viele davon sofort laufen. Dieselbe Rolle — ein Satz Wirklichkeit vor dem
 * Knopf —, nur mit den Daten dieser Plattform.
 *
 * „Start" ist die Handlung dieser Seite und damit der einzige Primär-Knopf.
 * Die Pflege- und Rollen-Wege, die hier bis v1.2 als Blasen standen, sind
 * Kacheln der Start-Übersicht; auf der Landing Page wären sie ein zweiter Weg
 * dorthin.
 *
 * Der Knopf ist NICHT gesperrt, solange geladen wird. Wer auf einen toten
 * Knopf klickt, hält die App für kaputt; die Übersicht dahinter zeigt ihren
 * eigenen Ladezustand.
 */

import * as React from 'react';
import DexLogo from './DexLogo';
import { LayoutGrid, Plus, AlertCircle } from './Icons';
import { ensureDexUiStyles } from './dexUi';
import { useLanguage } from '../context/LanguageContext';
import { useNavigation } from '../context/NavigationContext';
import { useRoles } from '../context/RoleContext';
import { useUseCases } from '../context/UseCaseContext';
import { useCurrentUser } from '../context/UserContext';
import { useHilfe } from '../context/HilfeContext';
import { useIsMobile } from '../utils/useIsMobile';
import { APP_NAME } from '../constants';

export default function LandingPage(): React.ReactElement {
  ensureDexUiStyles();
  const { t, isDe } = useLanguage();
  const { navigate } = useNavigation();
  const { isOrganizer, isRolesLoading } = useRoles();
  const { useCases, ladeStatus } = useUseCases();
  const { currentUser } = useCurrentUser();
  const { openAbout, openKontakt } = useHilfe();
  const isMobile = useIsMobile();

  const stunde = new Date().getHours();
  const gruss = stunde < 11
    ? t('Guten Morgen', 'Good morning')
    : stunde < 18 ? t('Guten Tag', 'Good afternoon') : t('Guten Abend', 'Good evening');

  // Archivierte zählen nicht: Für normale Nutzer stehen sie nicht auf der
  // Wand, und eine Zahl, die man dort nicht nachzählen kann, ist keine.
  const sichtbar = useCases.filter(u => u.status !== 'Archiviert');
  const gesamt = sichtbar.length;
  const live = sichtbar.filter(u => u.status === 'Live').length;

  // Die Einladung zur Rolle nur für den, der sie nicht hat — und erst, wenn
  // die Rollen geladen sind: Sonst blitzt sie bei jedem Organizer kurz auf.
  const zeigeOrganizerEinladung = !isOrganizer && !isRolesLoading && !isMobile;

  const kartenSymbol = (
    <span
      aria-hidden="true"
      style={{
        width: 66, height: 66, flexShrink: 0, borderRadius: 8, border: '1px solid var(--dex-gray-200)',
        background: 'rgba(134,188,37,0.12)', color: 'var(--dex-green-dark, #4a7c1f)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
      }}
    >
      <LayoutGrid size={30} strokeWidth={1.6} />
    </span>
  );

  return (
    <div className="landing">
      <div className="landing__hero">
        <div className="landing__card" style={{ position: 'relative' }}>
          <div className="landing__orb">
            <DexLogo title={APP_NAME} motion="oscillate" style={{ width: '100%' }} />
          </div>

          <div className="landing__text">
            <h1>{gruss}{currentUser.firstName ? <>, <strong>{currentUser.firstName}</strong></> : ''}.</h1>
            <p>
              {(() => {
                // Das Wort, das etwas öffnet, steht im Satz, nicht als Knopf
                // daneben — wie „Self-Service-App" in DEX.
                const link = (
                  <button
                    type="button"
                    className="dex-ui-textlink"
                    onClick={openAbout}
                    title={t('Was die Plattform ist und wie sie läuft', 'What the platform is and how it works')}
                  >
                    Agentic Banking Demo Hub
                  </button>
                );
                return isDe
                  ? <>Willkommen bei der <strong>{APP_NAME}</strong>. Unser {link} für Agent-Demos im Banking-Sektor. Von der Idee bis zur laufenden Demo. Alles an einer Stelle.</>
                  : <>Welcome to the <strong>{APP_NAME}</strong>. Our {link} for agent demos in the banking sector. From idea to running demo. Everything in one place.</>;
              })()}
            </p>
          </div>

          {/* Was hinter „Start" wartet. Drei Zustände, drei Aussagen: „0 Use
              Cases" bei einem Lesefehler wäre eine Aussage über die Plattform,
              die niemand geprüft hat. */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10, width: '100%' }}>
            {ladeStatus === 'fehler' && (
              <div className="dex-ui-callout dex-ui-callout--warn" role="status">
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <span>
                  {t('Wie viele Use Cases es gerade gibt, konnte nicht gelesen werden — die Übersicht sagt dir den Grund.',
                    'How many use cases there are right now could not be read — the overview tells you why.')}
                </span>
              </div>
            )}

            {ladeStatus === 'laedt' && (
              <div className="dex-ui-card" style={{ display: 'flex', alignItems: 'center', gap: 14, width: '100%' }}>
                {kartenSymbol}
                <span style={{ minWidth: 0, flex: 1 }}>
                  <span style={{ display: 'block', fontWeight: 700, fontSize: '0.92rem', color: 'var(--dex-gray-800)' }}>
                    {t('Die Use Cases werden geladen …', 'Loading the use cases …')}
                  </span>
                  <div className="dex-ui-progress dex-ui-progress--indeterminate" style={{ marginTop: 8 }}><div className="dex-ui-progress-bar" /></div>
                </span>
              </div>
            )}

            {ladeStatus === 'ok' && (
              <button
                type="button"
                onClick={() => navigate('usecases')}
                className="dex-ui-card dex-ui-card--hover"
                style={{ display: 'flex', alignItems: 'center', gap: 14, width: '100%', textAlign: 'left', cursor: 'pointer', fontFamily: 'inherit' }}
                title={t('Zu den Use Cases', 'Go to the use cases')}
              >
                {kartenSymbol}
                <span style={{ minWidth: 0, flex: 1 }}>
                  <span style={{ display: 'block', fontWeight: 700, fontSize: '0.92rem', color: 'var(--dex-gray-800)' }}>
                    {gesamt === 0
                      ? t('Noch keine Use Cases', 'No use cases yet')
                      : isDe
                        ? <><span style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>{gesamt} {gesamt === 1 ? 'Use Case' : 'Use Cases'}</span> {gesamt === 1 ? 'steht' : 'stehen'} bereit</>
                        : <><span style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>{gesamt} {gesamt === 1 ? 'use case' : 'use cases'}</span> {gesamt === 1 ? 'is' : 'are'} ready</>}
                  </span>
                  <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 2 }}>
                    {gesamt === 0
                      ? (isOrganizer
                        ? t('Leg den ersten im Use Case Studio an.', 'Create the first one in the Use Case Studio.')
                        : t('Sobald die ersten Agent-Demos stehen, erscheinen sie hier.', 'As soon as the first agent demos exist, they appear here.'))
                      : t(`${live} davon kannst du sofort starten`, `you can start ${live} of them right away`)}
                  </span>
                </span>
                {gesamt > 0 && (
                  <span className="dex-ui-pill dex-ui-pill--green" style={{ flexShrink: 0 }}>
                    {live} live
                  </span>
                )}
              </button>
            )}
          </div>

          <button className="btn btn-lg btn-block btn-primary" onClick={() => navigate('start')} style={{ maxWidth: 360 }}>
            {t('Start', 'Start')}
          </button>

          {/* Der Weg zur Rolle — wie „DEX für dein Event nutzen". Nur für den,
              der sie nicht hat; auf dem Handy ausgeblendet (spart Platz, und
              die Anfrage bleibt in der Start-Übersicht erreichbar). */}
          {zeigeOrganizerEinladung && (
            <div className="landing__actions" style={{ display: 'flex', gap: 12, justifyContent: 'center', alignItems: 'center', flexWrap: 'wrap' }}>
              <button
                type="button"
                onClick={() => openKontakt('organizer')}
                className="landing__bubble dex-ui-card dex-ui-card--hover"
                style={{ display: 'flex', alignItems: 'center', gap: 12, cursor: 'pointer', textAlign: 'left', fontFamily: 'inherit', width: '100%', maxWidth: 400 }}
                title={t('Use Case Organizer werden', 'Become a Use Case Organizer')}
              >
                <span className="dex-ui-choice-icon" aria-hidden="true" style={{ width: 38, height: 38, borderRadius: '50%' }}>
                  <Plus size={20} />
                </span>
                <span style={{ minWidth: 0 }}>
                  <span style={{ display: 'block', fontWeight: 800, fontSize: '0.98rem', lineHeight: 1.25, color: 'var(--dex-gray-800)' }}>
                    {t('Eigenen Use Case einstellen', 'Add your own use case')}
                  </span>
                  <span style={{ display: 'block', fontSize: '0.82rem', lineHeight: 1.3, marginTop: 2, color: 'var(--dex-gray-600)' }}>
                    {t('Werde Use Case Organizer und pflege eigene Demos.', 'Become a Use Case Organizer and maintain your own demos.')}
                  </span>
                </span>
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
