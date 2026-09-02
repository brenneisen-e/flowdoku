/* WizardTermsModal — aus EventCreationPage.tsx ausgelagert (Zeilen 5989-6322 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; die Komponente
 * gibt ein Fragment zurueck, damit die Geschwister-Reihenfolge im Elternbaum
 * exakt bleibt. */
import * as React from 'react';
import { Check } from '../Icons';

export interface WizardTermsModalProps {
  canBilling: boolean;
  goBack: () => void;
  internalCheckbox: boolean;
  isDe: boolean;
  setBillingPromptOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setInternalCheckbox: React.Dispatch<React.SetStateAction<boolean>>;
  setTcAccepted: React.Dispatch<React.SetStateAction<boolean>>;
  setTcCheckbox: React.Dispatch<React.SetStateAction<boolean>>;
  setTcExpanded: React.Dispatch<React.SetStateAction<boolean>>;
  showTermsModal: boolean;
  tcCheckbox: boolean;
  tcExpanded: boolean;
}

export const WizardTermsModal: React.FC<WizardTermsModalProps> = (p) => {
  const { canBilling, goBack, internalCheckbox, isDe, setBillingPromptOpen, setInternalCheckbox, setTcAccepted, setTcCheckbox, setTcExpanded, showTermsModal, tcCheckbox, tcExpanded } = p;
  return (
    <>
      {showTermsModal && (
        <div
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1200,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
          }}
        >
          <div
            className="card"
            style={{
              width: '100%', maxWidth: 720, maxHeight: '90vh', overflow: 'auto',
              padding: 28, borderRadius: 16, background: '#fff',
              boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
            }}
          >
            <h2 style={{ margin: '0 0 4px', fontSize: '1.3rem' }}>
              {isDe
                ? 'Deloitte Event Experience Platform — Nutzungsbedingungen (Deutschland)'
                : 'Deloitte Event Experience Platform — Terms of Use (Germany)'}
            </h2>
            <p style={{ margin: '0 0 16px', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
              {isDe ? 'Letzte Überarbeitung: 05.08.2026' : 'Last revised: 5 August 2026'}
            </p>

            {/* Eingeklappte Kurzfassung — die volle Fassung kann der Nutzer
                über den Toggle ausklappen. Die Checkbox-Bestätigung ist
                trotzdem Pflicht (siehe weiter unten). */}
            <div
              style={{
                background: 'var(--dex-gray-50, #f8f9fa)',
                border: '1px solid var(--dex-gray-200, #e5e7eb)',
                borderRadius: 10,
                padding: '12px 14px',
                fontSize: '0.88rem',
                lineHeight: 1.5,
                color: 'var(--dex-gray-700)',
              }}
            >
              {isDe
                ? <>Bitte gehe sorgfältig mit personenbezogenen Daten der Teilnehmer um, sammle nur das absolut Nötige, nutze die Daten ausschließlich für den vereinbarten Event-Zweck und beachte die Datenschutzregeln von Deloitte Deutschland. Volltext über den Button unten einsehen.</>
                : <>Please handle attendees&apos; personal data with care, collect only what is absolutely necessary, use the data exclusively for the agreed event purpose, and follow Deloitte Germany&apos;s data-protection rules. Use the button below to read the full text.</>}
            </div>

            <button
              type="button"
              onClick={() => setTcExpanded(v => !v)}
              style={{
                marginTop: 10,
                background: 'none',
                border: 'none',
                color: 'var(--dex-green-dark, #4a7c1f)',
                fontWeight: 600,
                fontSize: '0.85rem',
                cursor: 'pointer',
                padding: '4px 0',
                display: 'inline-flex',
                alignItems: 'center',
                gap: 6,
              }}
            >
              {tcExpanded
                ? (isDe ? '▲ Vollständige Bedingungen einklappen' : '▲ Hide full terms')
                : (isDe ? '▼ Vollständige Bedingungen anzeigen' : '▼ Show full terms')}
            </button>

            {tcExpanded && (
              <div style={{ fontSize: '0.88rem', lineHeight: 1.55, color: 'var(--dex-gray-800)', marginTop: 12 }}>
                {isDe ? (
                  <>
                    <p>
                      Der Zugang zur Event Experience Platform wird dir als Mitarbeiter von Deloitte Deutschland gewährt,
                      damit du das Teilnehmermanagement für Veranstaltungen, Events, Workshops oder andere Termine
                      organisieren kannst.
                    </p>

                    <p style={{ marginBottom: 6 }}>Die Plattform dient zur Koordination von:</p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Internen Deloitte Veranstaltungen</li>
                      <li>Externen Veranstaltungen, bei denen das Teilnehmermanagement für Deloitte-Mitarbeiter organisiert wird (bspw. Laufveranstaltungen wie B2Run, oder JPMorgan)</li>
                    </ul>

                    <p>
                      <strong>Wichtiger Hinweis:</strong> Für externe Events mit externen Teilnehmern ist die Plattform
                      nicht vorgesehen. Externe Nicht-Deloitte-Mitarbeiter werden über dieses Tool
                      nicht koordiniert und erhalten keinen Zugang zur Plattform. Alles zu solchen Veranstaltungen findest du im{' '}
                      <a href="https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx" target="_blank" rel="noopener noreferrer">Event Management im DeloitteNet</a>.
                    </p>

                    <p>Jedes Event, das du erstellst, muss den nachfolgenden Richtlinien folgen.</p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Wichtige Datenschutzhinweise</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Die Teilnahme an Events ist immer freiwillig und darf nicht erzwungen werden.</li>
                      <li>Vermeide die Sammlung personenbezogener Daten so weit wie möglich.</li>
                      <li>Sammle nur die Daten, die du unbedingt benötigst, um den Zweck des Events zu erreichen.</li>
                      <li>Reduziere Freitextfelder auf das absolute Minimum, um individuelle Informationen zur Identifizierung von Personen zu vermeiden.</li>
                      <li>Verwende gesammelte Daten ausschließlich für den definierten und genehmigten Zweck. Falls Abweichungen notwendig sind, wende dich im Voraus an das Datenschutz-Team.</li>
                    </ul>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Berechtigungen und Datenzugriff</h3>
                    <p style={{ marginTop: 0, marginBottom: 6 }}><strong>Als Event-Ersteller / Administrator:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Du erhältst Admin-Funktionalitäten für dein spezifisches Event.</li>
                      <li>Du kannst auf die gesamte Teilnehmerliste deines Events zugreifen.</li>
                      <li>Diese Berechtigung gilt ausschließlich für das von dir erstellte Event.</li>
                      <li>Du darfst Teilnehmerinformationen nicht mit anderen teilen oder für andere Zwecke verwenden.</li>
                    </ul>

                    <p style={{ marginBottom: 6 }}><strong>Als Event-Teilnehmer:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Du kannst dich für Events an- oder abmelden.</li>
                      <li>Deine Anmeldung ist freiwillig.</li>
                      <li>Du erhältst Informationen zum jeweiligen Event.</li>
                      <li>Du hast keinen Zugriff auf die Teilnehmerliste oder Informationen über andere Teilnehmer.</li>
                      <li>Du siehst nur deine eigenen Event-Anmeldungen und -Daten.</li>
                    </ul>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Datenschutzbestimmungen im Detail</h3>
                    <p style={{ marginTop: 0, marginBottom: 6 }}><strong>Beschränkung der Sammlung personenbezogener und vertraulicher Daten:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Nur was unbedingt erforderlich ist, um den beabsichtigten Zweck zu erreichen.</li>
                      <li>Offene Fragen auf das Minimum reduzieren (um die Sammlung unnötiger oder nicht autorisierter Daten zu vermeiden).</li>
                    </ul>

                    <p>
                      <strong>Sammle keine sensiblen personenbezogenen Daten</strong> — das heißt: keine Daten bezüglich
                      Rasse oder ethnischer Herkunft, religiöser oder philosophischer Überzeugungen,
                      Gewerkschaftsmitgliedschaft, politischer Meinungen, medizinischer oder gesundheitlicher Zustände
                      oder Informationen über das Sexualleben oder die sexuelle Orientierung einer Person. Falls sensible
                      personenbezogene Daten gesammelt werden müssen, kontaktiere zuerst das Team unter
                      {' '}<a href="mailto:privacy@deloitte.de">privacy@deloitte.de</a>.
                    </p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Besondere Bestimmungen für das Teilnehmermanagement</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Teilnehmerdaten dürfen nur für das spezifische Event verwendet werden, für das sie gesammelt wurden.</li>
                      <li>Die Weitergabe von Teilnehmerlisten an Dritte ist untersagt.</li>
                      <li>Teilnehmerdaten anderer Events sind nicht einsehbar.</li>
                      <li>Nach Abschluss des Events sind Teilnehmerdaten gemäß den Deloitte-Richtlinien zu behandeln.</li>
                    </ul>

                    <p>
                      Ermögliche anonyme Antworten, wann immer möglich. Verwende personenbezogene und vertrauliche Daten,
                      die in einem Event gesammelt wurden, nicht für andere Zwecke als den ursprünglich angegebenen.
                      Sprich dich mit dem Datenschutz-Team ab, falls eine andere Nutzung der Daten beabsichtigt ist
                      (du benötigst die vorherige schriftliche Einwilligung der betroffenen Personen / Teilnehmer
                      unter Verwendung einer entsprechenden Vorlage).
                    </p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Kontaktinformationen</h3>
                    <ul style={{ marginTop: 0 }}>
                      {/* v29.43: Funktionspostfach statt persönlichem Konto. */}
                      <li>Kontakt: DEX-Team (<a href="mailto:dex.event@deloitte.de">dex.event@deloitte.de</a>)</li>
                    </ul>

                    <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                      Diese Richtlinien gelten für alle Arten von Events, einschließlich Workshops, Seminare,
                      Webinare, Konferenzen und andere Veranstaltungen, deren Teilnehmermanagement für
                      Deloitte-Mitarbeiter über die Event Experience Platform organisiert wird.
                    </p>
                  </>
                ) : (
                  <>
                    <p>
                      Access to the Event Experience Platform is granted to you as an employee of Deloitte Germany so
                      that you can organise attendee management for events, workshops or other appointments.
                    </p>

                    <p style={{ marginBottom: 6 }}>The platform is used to coordinate:</p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Internal Deloitte events</li>
                      <li>External events for which attendee management is organised on behalf of Deloitte employees (e.g. running events such as B2Run or JPMorgan)</li>
                    </ul>

                    <p>
                      <strong>Important note:</strong> The platform is not intended for external events with external
                      attendees. External non-Deloitte employees are not coordinated through this
                      tool and will not be granted access to the platform. Everything about such events is on{' '}
                      <a href="https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx" target="_blank" rel="noopener noreferrer">Event Management on DeloitteNet</a>.
                    </p>

                    <p>Every event you create must follow the guidelines below.</p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Key data-protection guidance</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Attending events is always voluntary and must never be enforced.</li>
                      <li>Avoid collecting personal data wherever possible.</li>
                      <li>Only collect data that is strictly necessary to achieve the event&apos;s purpose.</li>
                      <li>Keep free-text fields to an absolute minimum to avoid collecting individual information that could identify people.</li>
                      <li>Use collected data exclusively for the defined and approved purpose. If you need to deviate, contact the data-protection team in advance.</li>
                    </ul>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Permissions and data access</h3>
                    <p style={{ marginTop: 0, marginBottom: 6 }}><strong>As event creator / administrator:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>You receive admin functionality for your specific event.</li>
                      <li>You can access the entire attendee list of your event.</li>
                      <li>This permission is limited to the event you created.</li>
                      <li>You may not share attendee information with others or use it for other purposes.</li>
                    </ul>

                    <p style={{ marginBottom: 6 }}><strong>As event attendee:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>You can register for or unregister from events.</li>
                      <li>Your registration is voluntary.</li>
                      <li>You receive information about the relevant event.</li>
                      <li>You have no access to the attendee list or information about other attendees.</li>
                      <li>You only see your own event registrations and data.</li>
                    </ul>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Data-protection rules in detail</h3>
                    <p style={{ marginTop: 0, marginBottom: 6 }}><strong>Restricting the collection of personal and confidential data:</strong></p>
                    <ul style={{ marginTop: 0 }}>
                      <li>Only what is strictly necessary to achieve the intended purpose.</li>
                      <li>Reduce open-ended questions to a minimum (to avoid collecting unnecessary or unauthorised data).</li>
                    </ul>

                    <p>
                      <strong>Do not collect sensitive personal data</strong> — that is, no data on race or ethnic origin,
                      religious or philosophical beliefs, trade-union membership, political opinions, medical or health
                      conditions, or information about a person&apos;s sex life or sexual orientation. If sensitive personal
                      data must be collected, contact the team first at
                      {' '}<a href="mailto:privacy@deloitte.de">privacy@deloitte.de</a>.
                    </p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Specific rules for attendee management</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Attendee data may only be used for the specific event for which it was collected.</li>
                      <li>Sharing attendee lists with third parties is prohibited.</li>
                      <li>Attendee data of other events is not accessible.</li>
                      <li>After the event, attendee data must be handled in line with Deloitte policy.</li>
                    </ul>

                    <p>
                      Allow anonymous responses wherever possible. Do not use personal or confidential data collected for
                      one event for purposes other than the originally stated one. Coordinate with the data-protection
                      team if you intend to use the data differently (you will need prior written consent from the
                      affected individuals / attendees, using an appropriate template).
                    </p>

                    <h3 style={{ fontSize: '1rem', marginTop: 20, marginBottom: 8 }}>Contact</h3>
                    <ul style={{ marginTop: 0 }}>
                      <li>Contact: DEX team (<a href="mailto:dex.event@deloitte.de">dex.event@deloitte.de</a>)</li>
                    </ul>

                    <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                      These guidelines apply to all types of events including workshops, seminars, webinars, conferences
                      and any other events whose attendee management for Deloitte employees is organised through the
                      Event Experience Platform.
                    </p>
                  </>
                )}
              </div>
            )}

            <label
              style={{
                display: 'flex', alignItems: 'flex-start', gap: 10,
                marginTop: 20, padding: 14,
                background: tcCheckbox ? 'rgba(134,188,37,0.08)' : 'var(--dex-gray-50, #f8f9fa)',
                border: `1px solid ${tcCheckbox ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200, #e5e7eb)'}`,
                borderRadius: 10, cursor: 'pointer',
                transition: 'all 0.15s ease',
              }}
            >
              <input
                type="checkbox"
                checked={tcCheckbox}
                onChange={e => setTcCheckbox(e.target.checked)}
                style={{ marginTop: 2, width: 18, height: 18, accentColor: 'var(--dex-green, #86bc25)', cursor: 'pointer' }}
              />
              <span style={{ fontSize: '0.9rem', lineHeight: 1.4 }}>
                {isDe
                  ? 'Ich habe die Nutzungs- und Datenschutzbedingungen gelesen und akzeptiere sie. Ich bestätige, dass ich mich beim Anlegen und Verwalten dieses Events an die Datenschutzbestimmungen halten werde.'
                  : 'I have read and accept the terms of use and data-protection rules. I confirm that I will follow the data-protection rules when creating and managing this event.'}
              </span>
            </label>

            <label
              style={{
                display: 'flex', alignItems: 'flex-start', gap: 10,
                marginTop: 12, padding: 14,
                background: internalCheckbox ? 'rgba(134,188,37,0.08)' : 'var(--dex-gray-50, #f8f9fa)',
                border: `1px solid ${internalCheckbox ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200, #e5e7eb)'}`,
                borderRadius: 10, cursor: 'pointer',
                transition: 'all 0.15s ease',
              }}
            >
              <input
                type="checkbox"
                checked={internalCheckbox}
                onChange={e => setInternalCheckbox(e.target.checked)}
                style={{ marginTop: 2, width: 18, height: 18, accentColor: 'var(--dex-green, #86bc25)', cursor: 'pointer' }}
              />
              <span style={{ fontSize: '0.9rem', lineHeight: 1.4 }}>
                {isDe
                  ? <>Ich bestätige, dass dies ein <strong>internes Deloitte Event</strong> ist oder die <strong>Deloitte-Teilnahme an einer externen Veranstaltung</strong> koordiniert.
                      <span style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-600)', marginTop: 4 }}>
                        Für externe Events mit externen Teilnehmern ist DEX nicht vorgesehen — alles dazu findest du im <a href="https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx" target="_blank" rel="noopener noreferrer" style={{ color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600 }}>Event Management im DeloitteNet</a>.
                      </span></>
                  : <>I confirm that this is a <strong>Deloitte-internal event</strong> or coordinates <strong>Deloitte participation in an external event</strong>.
                      <span style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-600)', marginTop: 4 }}>
                        DEX is not intended for external events with external attendees — everything about those is on <a href="https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx" target="_blank" rel="noopener noreferrer" style={{ color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600 }}>Event Management on DeloitteNet</a>.
                      </span></>}
              </span>
            </label>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 12, marginTop: 20 }}>
              <button
                type="button"
                className="btn btn-outline"
                onClick={() => goBack()}
              >
                {isDe ? 'Abbrechen' : 'Cancel'}
              </button>
              <button
                type="button"
                className="btn btn-primary"
                disabled={!tcCheckbox || !internalCheckbox}
                onClick={() => {
                  setTcAccepted(true);
                  // v29.66: F&A-Pilot — direkt nach dem Akzeptieren fragt der
                  // Dialog nach der Abrechnungsrelevanz (nur Admins, nur beim
                  // Anlegen; im Edit-Modus erscheinen die Bedingungen nicht).
                  if (canBilling) setBillingPromptOpen(true);
                }}
                style={{ opacity: (tcCheckbox && internalCheckbox) ? 1 : 0.5, cursor: (tcCheckbox && internalCheckbox) ? 'pointer' : 'not-allowed' }}
              >
                <Check size={16} /> {isDe ? 'Akzeptieren & weiter' : 'Accept & continue'}
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
};
