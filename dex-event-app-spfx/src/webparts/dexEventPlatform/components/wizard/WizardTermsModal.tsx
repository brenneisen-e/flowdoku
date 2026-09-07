/* WizardTermsModal — aus EventCreationPage.tsx ausgelagert (Zeilen 5989-6322 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; die Komponente
 * gibt ein Fragment zurueck, damit die Geschwister-Reihenfolge im Elternbaum
 * exakt bleibt.
 *
 * v31.2: Auf die gemeinsame `Modal`-Komponente (Kopf/Fuß) und die dex-ui-Klassen
 * umgestellt. Der Text ist rechtlich relevant und bleibt Wort für Wort — nur die
 * Gliederung ändert sich: Kurzfassung als Hinweiskasten, Volltext als Aufklapper
 * mit Abschnitten, die zwei Pflicht-Bestätigungen als Schalter-Zeilen mit Titel
 * und Folge-Zeile. Der Dialog ist eine Pflicht-Entscheidung: kein Schließen über
 * Backdrop, Escape oder X (`dismissable={false}` + `hideClose`), genau wie vorher
 * — verlassen wird er nur über „Abbrechen" (goBack) oder „Akzeptieren". */
import * as React from 'react';
import Modal from '../Modal';
import { cx } from '../dexUi';
import { Check, ChevronDown, FileText } from '../Icons';

// v31.2: Ein Link steht dreimal im Text — einmal deklarieren, damit ein Tippfehler
// nicht drei Stellen auseinanderlaufen lässt.
const EVENT_MGMT_URL = 'https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx';
const LINK_STYLE: React.CSSProperties = { color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600 };
const TEXT_STYLE: React.CSSProperties = { fontSize: '0.88rem', lineHeight: 1.55, color: 'var(--dex-gray-800)' };
const UL: React.CSSProperties = { marginTop: 0 };
const P_LEAD: React.CSSProperties = { marginTop: 0, marginBottom: 6 };

/** v31.2: Ein Abschnitt des Volltexts — Versal-Überschrift mit Linie statt sechs gleich gestylter <h3>. */
const TermsSection: React.FC<{ title: string; children: React.ReactNode }> = ({ title, children }) => (
  <div className="dex-ui-section">
    <h3 className="dex-ui-section-title">{title}</h3>
    {children}
  </div>
);

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
  const bothChecked = tcCheckbox && internalCheckbox;

  // v31.2: Genau ein Primär-Knopf (Akzeptieren), Abbrechen sekundär. Das
  // Deaktivieren läuft über `disabled` — die Modal-Styles dämpfen es selbst.
  const footer = (
    <>
      <button type="button" className="btn btn-secondary" onClick={() => goBack()}>
        {isDe ? 'Abbrechen' : 'Cancel'}
      </button>
      <button
        type="button"
        className="btn btn-primary"
        disabled={!bothChecked}
        title={bothChecked ? undefined : (isDe ? 'Bitte bestätige zuerst beide Punkte.' : 'Please confirm both points first.')}
        onClick={() => {
          setTcAccepted(true);
          // v29.66: F&A-Pilot — direkt nach dem Akzeptieren fragt der
          // Dialog nach der Abrechnungsrelevanz (nur Admins, nur beim
          // Anlegen; im Edit-Modus erscheinen die Bedingungen nicht).
          if (canBilling) setBillingPromptOpen(true);
        }}
      >
        <Check size={16} /> {isDe ? 'Akzeptieren & weiter' : 'Accept & continue'}
      </button>
    </>
  );

  return (
    <Modal
      open={showTermsModal}
      onClose={() => goBack()}
      maxWidth={720}
      dismissable={false}
      hideClose
      ariaLabel={isDe ? 'Nutzungsbedingungen' : 'Terms of use'}
      icon={<FileText size={20} />}
      title={isDe
        ? 'Deloitte Event Experience Platform — Nutzungsbedingungen (Deutschland)'
        : 'Deloitte Event Experience Platform — Terms of Use (Germany)'}
      subtitle={isDe
        ? 'Letzte Überarbeitung: 05.08.2026 · Lies die Kurzfassung, bestätige die zwei Punkte unten — dann geht es zum Event.'
        : 'Last revised: 5 August 2026 · Read the summary, confirm the two points below — then you continue to the event.'}
      footer={footer}
    >
      {/* Eingeklappte Kurzfassung — die volle Fassung kann der Nutzer
          über den Toggle ausklappen. Die Checkbox-Bestätigung ist
          trotzdem Pflicht (siehe weiter unten). */}
      {/* v31.2: marginTop 0 — die Modal-Karte hält die Abstände schon über ihr gap. */}
      <div className="dex-ui-section" style={{ marginTop: 0 }}>
        <h4 className="dex-ui-section-title">{isDe ? 'Das Wichtigste in Kürze' : 'The essentials'}</h4>
        <div className="dex-ui-callout dex-ui-callout--neutral" style={{ fontSize: '0.88rem', color: 'var(--dex-gray-700)' }}>
          <span>
            {isDe
              ? <>Geh sorgfältig mit personenbezogenen Daten der Teilnehmer um, sammle nur das absolut Nötige, nutze die Daten ausschließlich für den vereinbarten Event-Zweck und beachte die Datenschutzregeln von Deloitte Deutschland. Den Volltext kannst du direkt darunter aufklappen.</>
              : <>Handle attendees&apos; personal data with care, collect only what is absolutely necessary, use the data exclusively for the agreed event purpose, and follow Deloitte Germany&apos;s data-protection rules. You can expand the full text right below.</>}
          </span>
        </div>

        <button
          type="button"
          className={cx('dex-ui-disclosure', tcExpanded && 'is-open')}
          aria-expanded={tcExpanded}
          onClick={() => setTcExpanded(v => !v)}
          style={{ marginTop: 8 }}
        >
          <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
          {tcExpanded
            ? (isDe ? 'Vollständige Bedingungen einklappen' : 'Hide full terms')
            : (isDe ? 'Vollständige Bedingungen anzeigen' : 'Show full terms')}
        </button>
      </div>

      {tcExpanded && (
        <div className="dex-ui-disclosure-body dex-ui-fade-in" style={TEXT_STYLE}>
          {isDe ? (
            <>
              <p style={{ marginTop: 0 }}>
                Der Zugang zur Event Experience Platform wird dir als Mitarbeiter von Deloitte Deutschland gewährt,
                damit du das Teilnehmermanagement für Veranstaltungen, Events, Workshops oder andere Termine
                organisieren kannst.
              </p>

              <p style={P_LEAD}>Die Plattform dient zur Koordination von:</p>
              <ul style={UL}>
                <li>Internen Deloitte Veranstaltungen</li>
                <li>Externen Veranstaltungen, bei denen das Teilnehmermanagement für Deloitte-Mitarbeiter organisiert wird (bspw. Laufveranstaltungen wie B2Run, oder JPMorgan)</li>
              </ul>

              <div className="dex-ui-callout dex-ui-callout--warn">
                <span>
                  <strong>Wichtiger Hinweis:</strong> Für externe Events mit externen Teilnehmern ist die Plattform
                  nicht vorgesehen. Externe Nicht-Deloitte-Mitarbeiter werden über dieses Tool
                  nicht koordiniert und erhalten keinen Zugang zur Plattform. Alles zu solchen Veranstaltungen findest du im{' '}
                  <a href={EVENT_MGMT_URL} target="_blank" rel="noopener noreferrer">Event Management im DeloitteNet</a>.
                </span>
              </div>

              <p>Jedes Event, das du erstellst, muss den nachfolgenden Richtlinien folgen.</p>

              <TermsSection title="Wichtige Datenschutzhinweise">
                <ul style={UL}>
                  <li>Die Teilnahme an Events ist immer freiwillig und darf nicht erzwungen werden.</li>
                  <li>Vermeide die Sammlung personenbezogener Daten so weit wie möglich.</li>
                  <li>Sammle nur die Daten, die du unbedingt benötigst, um den Zweck des Events zu erreichen.</li>
                  <li>Reduziere Freitextfelder auf das absolute Minimum, um individuelle Informationen zur Identifizierung von Personen zu vermeiden.</li>
                  <li>Verwende gesammelte Daten ausschließlich für den definierten und genehmigten Zweck. Falls Abweichungen notwendig sind, wende dich im Voraus an das Datenschutz-Team.</li>
                </ul>
              </TermsSection>

              <TermsSection title="Berechtigungen und Datenzugriff">
                <p style={P_LEAD}><strong>Als Event-Ersteller / Administrator:</strong></p>
                <ul style={UL}>
                  <li>Du erhältst Admin-Funktionalitäten für dein spezifisches Event.</li>
                  <li>Du kannst auf die gesamte Teilnehmerliste deines Events zugreifen.</li>
                  <li>Diese Berechtigung gilt ausschließlich für das von dir erstellte Event.</li>
                  <li>Du darfst Teilnehmerinformationen nicht mit anderen teilen oder für andere Zwecke verwenden.</li>
                </ul>
                <p style={P_LEAD}><strong>Als Event-Teilnehmer:</strong></p>
                <ul style={UL}>
                  <li>Du kannst dich für Events an- oder abmelden.</li>
                  <li>Deine Anmeldung ist freiwillig.</li>
                  <li>Du erhältst Informationen zum jeweiligen Event.</li>
                  <li>Du hast keinen Zugriff auf die Teilnehmerliste oder Informationen über andere Teilnehmer.</li>
                  <li>Du siehst nur deine eigenen Event-Anmeldungen und -Daten.</li>
                </ul>
              </TermsSection>

              <TermsSection title="Datenschutzbestimmungen im Detail">
                <p style={P_LEAD}><strong>Beschränkung der Sammlung personenbezogener und vertraulicher Daten:</strong></p>
                <ul style={UL}>
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
              </TermsSection>

              <TermsSection title="Besondere Bestimmungen für das Teilnehmermanagement">
                <ul style={UL}>
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
              </TermsSection>

              <TermsSection title="Kontaktinformationen">
                <ul style={UL}>
                  {/* v29.43: Funktionspostfach statt persönlichem Konto. */}
                  <li>Kontakt: DEX-Team (<a href="mailto:dex.event@deloitte.de">dex.event@deloitte.de</a>)</li>
                </ul>
                <p className="dex-ui-muted">
                  Diese Richtlinien gelten für alle Arten von Events, einschließlich Workshops, Seminare,
                  Webinare, Konferenzen und andere Veranstaltungen, deren Teilnehmermanagement für
                  Deloitte-Mitarbeiter über die Event Experience Platform organisiert wird.
                </p>
              </TermsSection>
            </>
          ) : (
            <>
              <p style={{ marginTop: 0 }}>
                Access to the Event Experience Platform is granted to you as an employee of Deloitte Germany so
                that you can organise attendee management for events, workshops or other appointments.
              </p>

              <p style={P_LEAD}>The platform is used to coordinate:</p>
              <ul style={UL}>
                <li>Internal Deloitte events</li>
                <li>External events for which attendee management is organised on behalf of Deloitte employees (e.g. running events such as B2Run or JPMorgan)</li>
              </ul>

              <div className="dex-ui-callout dex-ui-callout--warn">
                <span>
                  <strong>Important note:</strong> The platform is not intended for external events with external
                  attendees. External non-Deloitte employees are not coordinated through this
                  tool and will not be granted access to the platform. Everything about such events is on{' '}
                  <a href={EVENT_MGMT_URL} target="_blank" rel="noopener noreferrer">Event Management on DeloitteNet</a>.
                </span>
              </div>

              <p>Every event you create must follow the guidelines below.</p>

              <TermsSection title="Key data-protection guidance">
                <ul style={UL}>
                  <li>Attending events is always voluntary and must never be enforced.</li>
                  <li>Avoid collecting personal data wherever possible.</li>
                  <li>Only collect data that is strictly necessary to achieve the event&apos;s purpose.</li>
                  <li>Keep free-text fields to an absolute minimum to avoid collecting individual information that could identify people.</li>
                  <li>Use collected data exclusively for the defined and approved purpose. If you need to deviate, contact the data-protection team in advance.</li>
                </ul>
              </TermsSection>

              <TermsSection title="Permissions and data access">
                <p style={P_LEAD}><strong>As event creator / administrator:</strong></p>
                <ul style={UL}>
                  <li>You receive admin functionality for your specific event.</li>
                  <li>You can access the entire attendee list of your event.</li>
                  <li>This permission is limited to the event you created.</li>
                  <li>You may not share attendee information with others or use it for other purposes.</li>
                </ul>
                <p style={P_LEAD}><strong>As event attendee:</strong></p>
                <ul style={UL}>
                  <li>You can register for or unregister from events.</li>
                  <li>Your registration is voluntary.</li>
                  <li>You receive information about the relevant event.</li>
                  <li>You have no access to the attendee list or information about other attendees.</li>
                  <li>You only see your own event registrations and data.</li>
                </ul>
              </TermsSection>

              <TermsSection title="Data-protection rules in detail">
                <p style={P_LEAD}><strong>Restricting the collection of personal and confidential data:</strong></p>
                <ul style={UL}>
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
              </TermsSection>

              <TermsSection title="Specific rules for attendee management">
                <ul style={UL}>
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
              </TermsSection>

              <TermsSection title="Contact">
                <ul style={UL}>
                  <li>Contact: DEX team (<a href="mailto:dex.event@deloitte.de">dex.event@deloitte.de</a>)</li>
                </ul>
                <p className="dex-ui-muted">
                  These guidelines apply to all types of events including workshops, seminars, webinars, conferences
                  and any other events whose attendee management for Deloitte employees is organised through the
                  Event Experience Platform.
                </p>
              </TermsSection>
            </>
          )}
        </div>
      )}

      {/* v31.2: Die zwei Pflicht-Bestätigungen als Schalter-Zeilen mit Titel und
          Folge-Zeile — vorher zwei gleich aussehende Kästen mit je einem
          Fließtext-Satz, bei denen man erst beim Lesen merkte, dass es zwei
          verschiedene Fragen sind. */}
      <div className="dex-ui-section" style={{ marginTop: 0 }}>
        <h4 className="dex-ui-section-title">{isDe ? 'Deine Bestätigung' : 'Your confirmation'}</h4>
        <p className="dex-ui-section-desc">
          {isDe
            ? 'Beide Punkte sind Pflicht — erst dann geht es zum Event.'
            : 'Both points are required — only then can you continue to the event.'}
        </p>
        <div className="dex-ui-stack">
          <label className={cx('dex-ui-toggle-row', tcCheckbox && 'is-active')}>
            <input type="checkbox" checked={tcCheckbox} onChange={e => setTcCheckbox(e.target.checked)} />
            <span className="dex-ui-toggle-row-body">
              <span className="dex-ui-toggle-row-title">
                {isDe
                  ? 'Ich habe die Nutzungs- und Datenschutzbedingungen gelesen und akzeptiere sie.'
                  : 'I have read and accept the terms of use and data-protection rules.'}
              </span>
              <span className="dex-ui-toggle-row-desc">
                {isDe
                  ? 'Ich bestätige, dass ich mich beim Anlegen und Verwalten dieses Events an die Datenschutzbestimmungen halten werde.'
                  : 'I confirm that I will follow the data-protection rules when creating and managing this event.'}
              </span>
            </span>
          </label>
          <label className={cx('dex-ui-toggle-row', internalCheckbox && 'is-active')}>
            <input type="checkbox" checked={internalCheckbox} onChange={e => setInternalCheckbox(e.target.checked)} />
            <span className="dex-ui-toggle-row-body">
              <span className="dex-ui-toggle-row-title">
                {isDe
                  ? <>Ich bestätige, dass dies ein <strong>internes Deloitte Event</strong> ist oder die <strong>Deloitte-Teilnahme an einer externen Veranstaltung</strong> koordiniert.</>
                  : <>I confirm that this is a <strong>Deloitte-internal event</strong> or coordinates <strong>Deloitte participation in an external event</strong>.</>}
              </span>
              <span className="dex-ui-toggle-row-desc">
                {isDe
                  ? <>Für externe Events mit externen Teilnehmern ist DEX nicht vorgesehen — alles dazu findest du im <a href={EVENT_MGMT_URL} target="_blank" rel="noopener noreferrer" style={LINK_STYLE}>Event Management im DeloitteNet</a>.</>
                  : <>DEX is not intended for external events with external attendees — everything about those is on <a href={EVENT_MGMT_URL} target="_blank" rel="noopener noreferrer" style={LINK_STYLE}>Event Management on DeloitteNet</a>.</>}
              </span>
            </span>
          </label>
        </div>
      </div>
    </Modal>
  );
};
