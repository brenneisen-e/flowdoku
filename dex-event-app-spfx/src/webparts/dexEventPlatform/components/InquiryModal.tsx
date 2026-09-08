/**
 * Geteiltes Inquiry-Modal — "DEX App für dein Event anfragen".
 * Wird aus LandingPage und StartPage (Organizer-Tile-Overlay) verwendet.
 * Sendet eine Nachricht an die DEX-Maintainer; nimmt Name (vorbefüllt mit
 * Display-Name des eingeloggten Users), Event-Name und Freitext-Beschreibung
 * entgegen.
 */
import * as React from 'react';
import { useEvents } from '../context/EventContext';
import { useCurrentUser } from '../context/UserContext';
import { useLanguage } from '../context/LanguageContext';
import Modal from './Modal';
import LandingInfoModal from './LandingInfoModal';
import { Info, Check, AlertCircle, CaptainHat, MessageSquare, Send } from './Icons';
import { InfoTooltip } from './InfoTooltip';
// v31.2: Gemeinsame UI-Klassen (Auswahl-Kacheln, Toggle-Zeilen, Callouts) —
// Modal.tsx injiziert das Stylesheet, hier braucht es nur `cx`.
import { cx } from './dexUi';

/**
 * v29.45: Was der Organizer für sein Event braucht — ankreuzen statt frei
 * beschreiben. Das Freitextfeld war die einzige Angabe; wer DEX noch nicht
 * kennt, weiß dort nicht, wonach zu fragen wäre („Anzahl Teilnehmer, Termin,
 * gewünschte Funktionen …" nennt keine einzige Funktion). Die Liste sagt
 * zugleich, was die App überhaupt kann, und wir sehen vor dem ersten Gespräch,
 * worauf es hinausläuft.
 *
 * Bewusst kurz gehalten und in der Sprache der Organizer formuliert — keine
 * Feature-Namen aus dem Wizard. Das Freitextfeld bleibt für alles andere.
 */
const NEED_OPTIONS: Array<{ id: string; de: string; en: string; infoDe: string; infoEn: string }> = [
  {
    id: 'visibility',
    de: 'Nur für bestimmte Personen sichtbar (Verteiler, Standort)',
    en: 'Visible only to certain people (distribution list, location)',
    infoDe: 'Kann DEX schon: Du hinterlegst Mailverteiler, einzelne Personen und/oder Standorte — nur wer dazu passt, sieht das Event überhaupt. Einzelne Personen lassen sich gezielt ausschließen. Im Organizer Center siehst du jederzeit, wie viele Personen dahinterstehen, und kannst die, die noch nicht geantwortet haben, mit einem Klick erinnern.',
    infoEn: 'DEX already does this: you add distribution lists, individual people and/or locations — only matching people see the event at all, and individuals can be excluded. The Organizer Center shows how many people that is and lets you remind those who have not responded.',
  },
  {
    id: 'capacity',
    de: 'Begrenzte Plätze mit Warteliste und Nachrücken',
    en: 'Limited seats with waiting list and auto-promotion',
    infoDe: 'Kann DEX schon: Du setzt eine Teilnehmerzahl; ist sie erreicht, landen weitere Anmeldungen automatisch auf der Warteliste — in der Reihenfolge der Anmeldung. Meldet sich jemand ab, rückt die erste Person automatisch nach und bekommt ihre Mail. Freie Plätze kannst du auch selbst per Knopfdruck auffüllen.',
    infoEn: 'DEX already does this: set a capacity and further registrations go to the waiting list in order of registration. If someone cancels, the first person is promoted automatically and notified. You can also fill open seats manually.',
  },
  {
    id: 'fields',
    de: 'Eigene Fragen im Anmeldeformular (Essen, Größe, Zustimmung …)',
    en: 'Own questions in the registration form (food, size, consent …)',
    infoDe: 'Kann DEX schon: Du baust dir die Fragen selbst — Auswahl, Freitext, Zahl, Datum, Ja/Nein-Haken oder Personensuche, jeweils pflicht oder optional. Pflicht-Haken („Ich bestätige …") blockieren die Anmeldung, solange sie fehlen. Alle Antworten stehen in der Teilnehmerliste und im Excel-Export.',
    infoEn: 'DEX already does this: build your own questions — choice, free text, number, date, yes/no or people picker, each optional or required. Required consent boxes block submission until ticked. All answers appear in the attendee list and the Excel export.',
  },
  {
    id: 'subevents',
    de: 'Mehrere Termine oder Sessions zur Auswahl',
    en: 'Several dates or sessions to choose from',
    infoDe: 'Kann DEX schon: Ein Event kann mehrere Termine (Sessions, Workshops, Tage) enthalten, aus denen die Teilnehmer wählen — mit eigener Kapazität, eigenen Fragen, eigenen Mails und eigenem Outlook-Termin je Termin. Du siehst eine gemeinsame Teilnehmerliste über alle Termine hinweg.',
    infoEn: 'DEX already does this: an event can hold several sessions or days to choose from — each with its own capacity, questions, emails and calendar entry — with one combined attendee list across all of them.',
  },
  {
    id: 'documents',
    de: 'Dokumente bereitstellen oder von Teilnehmern einfordern',
    en: 'Share documents or require uploads from attendees',
    infoDe: 'Kann DEX schon: Du hängst Dateien ans Event (Agenda, Anfahrt, Programm) — sie stehen auf der Anmeldeseite und unter „Meine Events". Umgekehrt kannst du einen Upload verlangen (PDF oder Bild, auf Wunsch Pflicht); die Datei hängt an der Teilnehmerzeile und ist im Organizer Center einsehbar.',
    infoEn: 'DEX already does this: attach files to the event (agenda, directions) — visible on the registration page and in My Events. You can also require an upload (PDF or image, optionally mandatory) that is attached to the attendee row.',
  },
  {
    id: 'checkin',
    de: 'Check-in vor Ort mit QR-Code',
    en: 'On-site check-in with QR code',
    infoDe: 'Kann DEX schon: Jeder Teilnehmer bekommt auf Wunsch einen persönlichen QR-Code per Mail. Vor Ort scannt ihr ihn mit dem Handy — oder die Gäste checken sich über einen Aushang selbst ein. Du siehst live, wer da ist, und kannst auch von Hand ein- und auschecken.',
    infoEn: 'DEX already does this: attendees receive a personal QR code by email; scan it on site with a phone, or let guests check themselves in. You see live who has arrived and can check people in or out manually.',
  },
  {
    id: 'hotel',
    de: 'Hotel- und Übernachtungsplanung',
    en: 'Hotel and accommodation planning',
    infoDe: 'Kann DEX schon: Hotels mit Kontingenten und Zeiträumen hinterlegen, Teilnehmer automatisch verteilen lassen (auf Wunsch mit Zimmerpartner-Wunsch aus dem Formular) und die Belegung als Liste exportieren.',
    infoEn: 'DEX already does this: set up hotels with room contingents and stay periods, distribute attendees automatically (optionally honouring roommate wishes from the form) and export the allocation.',
  },
  {
    id: 'teams',
    de: 'Anmeldung als Team oder Gruppe',
    en: 'Registration as a team or group',
    infoDe: 'Kann DEX schon: Teilnehmer melden sich als Team an — mit Teamname, fester Teamgröße und einer Person, die das Team führt. Offene Plätze können sichtbar sein, damit andere dazustoßen; auf Wunsch bestätigt die Team-Leitung neue Mitglieder.',
    infoEn: 'DEX already does this: attendees can register as a team — team name, fixed size and a lead. Open slots can be visible so others may join, optionally with approval by the team lead.',
  },
  {
    id: 'external',
    de: 'Gäste von außerhalb Deloitte',
    en: 'Guests from outside Deloitte',
    infoDe: 'Teilweise: Externe ohne Deloitte-Konto kannst du über den Organizer eintragen lassen; die Anmeldeseite selbst ist auf Deloitte-Konten ausgelegt. Geht es um eine Veranstaltung, bei der sich überwiegend Externe selbst anmelden, ist das Event-Management-Team der richtige Weg — sprich uns an, wir sagen dir, was passt.',
    infoEn: 'Partly: guests without a Deloitte account can be added by the organizer; the registration page itself is built for Deloitte accounts. For events where mostly external guests register themselves, the event management team is the better route — talk to us.',
  },
  {
    id: 'online',
    de: 'Online- oder Hybrid-Teilnahme (Teams)',
    en: 'Online or hybrid attendance (Teams)',
    infoDe: 'Kann DEX schon: Du legst die Teams-Besprechung wie gewohnt selbst an und hinterlegst den Teilnahme-Link am Event. Er steht dann im Outlook-Termin der Teilnehmer, im Organizer Center und unter „Meine Events" als Knopf. DEX erzeugt keine Teams-Besprechungen selbst.',
    infoEn: 'DEX already does this: create the Teams meeting yourself and store the join link with the event. It then appears in the attendees\' calendar entry, in the Organizer Center and in My Events. DEX does not create Teams meetings itself.',
  },
];

/**
 * v29.46: Die zehn Punkte in drei Gruppen — untereinander gelistet lasen sie
 * sich wie eine Wunschliste ohne Ordnung. Die Gruppen folgen dem, worüber der
 * Organizer ohnehin nachdenkt: WER darf teilnehmen, WIE läuft die Anmeldung,
 * was passiert DRUMHERUM.
 */
const NEED_GROUPS: Array<{ id: string; de: string; en: string; ids: string[] }> = [
  { id: 'access', de: 'Wer teilnehmen darf', en: 'Who can attend', ids: ['visibility', 'external'] },
  { id: 'signup', de: 'Anmeldung', en: 'Registration', ids: ['capacity', 'subevents', 'fields', 'teams'] },
  { id: 'around', de: 'Rund um das Event', en: 'Around the event', ids: ['documents', 'checkin', 'hotel', 'online'] },
];

interface InquiryModalProps {
  open: boolean;
  onClose: () => void;
  /** v23.37: „Organizer werden"-Antrag statt allgemeiner DEX-Anfrage.
   *  Legt einen nachverfolgbaren Antrag an (Admins bestätigen ihn in der App). */
  organizerMode?: boolean;
}

export default function InquiryModal({ open, onClose, organizerMode }: InquiryModalProps): React.ReactElement | null {
  const { sendAdminInquiry, requestOrganizerRole } = useEvents();
  const { currentUser, photoUrl } = useCurrentUser();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const userFullName = `${currentUser.firstName} ${currentUser.surname}`.trim();
  // v24.24: Standort + Position des eingeloggten Users (read-only Anzeige +
  // gehen mit in die Anfrage-Mail an die Admins).
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const userLocation = ((currentUser as any).location || '').trim();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const userJobTitle = ((currentUser as any).jobTitle || '').trim();
  const userInitials = `${(currentUser.firstName || '')[0] || ''}${(currentUser.surname || '')[0] || ''}`.toUpperCase() || '?';
  const [eventName, setEventName] = React.useState('');
  const [message, setMessage] = React.useState('');
  const [sending, setSending] = React.useState(false);
  const [status, setStatus] = React.useState<'' | 'success' | 'error'>('');
  const [showInfo, setShowInfo] = React.useState(false);
  // v28.41: Art des Events VOR allem anderen abfragen. Bei „extern" ist DEX
  // das falsche Werkzeug — dann statt eines Formulars der Verweis auf die
  // Event-Management-Seite im DeloitteNet, und der Absende-Knopf bleibt gesperrt.
  const [eventScope, setEventScope] = React.useState<'' | 'internal' | 'external'>('');
  // v29.45: angekreuzte Bedarfe (Ids aus NEED_OPTIONS).
  const [needs, setNeeds] = React.useState<string[]>([]);
  // v29.46: „Etwas anderes" — die Liste kann nicht alles kennen, und ein
  // Bedarf, der nicht draufsteht, soll nicht durchs Raster fallen.
  const [otherNeed, setOtherNeed] = React.useState('');
  const toggleNeed = (id: string): void =>
    setNeeds(prev => (prev.indexOf(id) >= 0 ? prev.filter(x => x !== id) : [...prev, id]));

  // v23.37: im Organizer-Modus reicht der Name (Nachricht optional, kein
  // Event-Name) — die allgemeine Anfrage braucht Event-Name + Nachricht.
  // v28.41: Ohne Angabe der Event-Art laesst sich nichts absenden, und bei
  // „extern" bleibt der Knopf gesperrt — DEX ist dafür schlicht das falsche
  // Werkzeug, eine Anfrage wäre für beide Seiten verlorene Zeit.
  const canSubmit = organizerMode
    ? true
    : (!!eventName.trim() && !!message.trim() && eventScope === 'internal');

  async function handleSubmit(): Promise<void> {
    if (!canSubmit || sending) return;
    setSending(true);
    setStatus('');
    let ok = false;
    if (organizerMode) {
      const res = await requestOrganizerRole(currentUser.email || '', userFullName, userLocation, message.trim());
      ok = res.ok;
    } else {
      // v29.45: Die Auswahl geht als lesbare Liste mit in die Anfrage — so
      // steht sie in der Mail an das DEX-Team, ohne dass dort ein neues Feld
      // ausgewertet werden muss.
      const needLines = needs
        .map(id => NEED_OPTIONS.filter(o => o.id === id)[0])
        .filter(Boolean)
        .map(o => `• ${isDe ? o.de : o.en}`);
      if (needs.indexOf('other') >= 0 && otherNeed.trim()) {
        needLines.push(`• ${isDe ? 'Sonstiges' : 'Other'}: ${otherNeed.trim()}`);
      }
      const needBlock = needLines.length > 0
        ? `${isDe ? 'Benötigte Funktionen:' : 'Needed features:'}\n${needLines.join('\n')}\n\n`
        : '';
      ok = await sendAdminInquiry(userFullName, currentUser.email || '', eventName.trim(), `${needBlock}${message.trim()}`, userLocation, userJobTitle);
    }
    setSending(false);
    if (ok) {
      setStatus('success');
      setEventName('');
      setMessage('');
      setNeeds([]);
      setOtherNeed('');
      setTimeout(() => { onClose(); setStatus(''); }, 1800);
    } else {
      setStatus('error');
    }
  }

  // v31.2: Bei „extern" bleibt der Rest des Formulars sichtbar, aber gedämpft —
  // der Kasten darüber sagt, dass DEX hier nicht passt; die Felder darunter
  // sollen dann nicht so aussehen, als lohne sich das Ausfüllen noch.
  const dimmed = !organizerMode && eventScope === 'external';
  const submitHint = !organizerMode && !canSubmit && !sending
    ? (eventScope === 'external'
        ? (isDe ? 'Für externe Events ist DEX nicht nutzbar.' : 'DEX cannot be used for external events.')
        : (isDe ? 'Wähle die Event-Art und fülle Name und Beschreibung aus.' : 'Choose the event type and fill in name and description.'))
    : undefined;

  // v13.1: Modal-Wrapper-Komponente — kapselt Backdrop/Escape/Padding.
  // v31.2: Kopf und Fuß über die Modal-Props (title/subtitle/icon/footer) statt
  // eigenem <h2> und eigener Knopfzeile; der Verweis auf die Info-Seite sitzt
  // links im Fuß als Textknopf — vorher eine große Farbfläche über dem Formular.
  return (
    <Modal
      open={open}
      onClose={onClose}
      dismissable={!sending}
      maxWidth={580}
      ariaLabel={organizerMode ? (isDe ? 'Organizer werden' : 'Become an organizer') : (isDe ? 'DEX-Anfrage' : 'DEX inquiry')}
      icon={organizerMode ? <CaptainHat size={20} /> : <MessageSquare size={20} />}
      title={organizerMode
        ? (isDe ? 'Organizer werden' : 'Become an organizer')
        : (isDe ? 'DEX App für dein Event anfragen' : 'Request the DEX App for your event')}
      subtitle={organizerMode
        ? (isDe
            ? 'Stelle einen Antrag, Organizer zu werden — dann kannst du eigene Events anlegen und verwalten. Die Admins prüfen deinen Antrag und schalten dich frei.'
            : 'Request to become an organizer — you can then create and manage your own events. The admins review your request and grant access.')
        : (isDe
            ? 'Wir melden uns kurz bei dir und besprechen, wie wir dein Event unterstützen können.'
            : 'We will get back to you and discuss how we can support your event.')}
      footer={<>
        {!organizerMode && (
          <span className="dex-ui-modal-foot-left">
            <button
              type="button"
              className="dex-ui-textbtn"
              onClick={() => setShowInfo(true)}
              title={isDe ? 'Erst mal mehr erfahren? Alle Funktionen, Zielgruppen und Beispiel-Events im Überblick.' : 'Learn more first? All features, audiences and example events at a glance.'}
            >
              <Info size={15} />
              {isDe ? 'Was kann die DEX App?' : 'What can the DEX App do?'}
            </button>
          </span>
        )}
        <button type="button" className="btn btn-secondary" onClick={onClose} disabled={sending}>
          {isDe ? 'Abbrechen' : 'Cancel'}
        </button>
        <button
          type="button"
          className="btn btn-primary"
          onClick={handleSubmit}
          disabled={sending || !canSubmit}
          title={submitHint}
        >
          <Send size={16} />
          {sending
            ? (isDe ? 'Wird gesendet …' : 'Sending …')
            : organizerMode
              ? (isDe ? 'Antrag senden' : 'Send request')
              : (isDe ? 'Anfrage senden' : 'Send inquiry')}
        </button>
      </>}
    >
        {/* v28.40: Einsatzbereich klarstellen. Bis hierhin stand in der
            Anfrage-Strecke nirgends, für welche Art von Events DEX gedacht
            ist — die einzige Erwaehnung von „extern" war die technische
            Aussage „keine externen APIs" in der Info-Box, die man sogar
            falsch herum lesen kann.
            v31.2: Jetzt die ERSTE Frage im Dialog, als zwei Auswahl-Kacheln
            statt Radio-Paar im grauen Kasten — sie entscheidet, ob der Rest
            überhaupt Sinn hat (bei „extern" bleibt Senden gesperrt), und eine
            Kachel mit Hover und Häkchen liest sich als Entscheidung, nicht als
            Fußnote. Semantik unverändert: eventScope 'internal' | 'external'. */}
        {!organizerMode && (
          <div role="radiogroup" aria-label={isDe ? 'Um was für ein Event geht es?' : 'What kind of event is it?'}>
            <div className="dex-ui-label">{isDe ? 'Um was für ein Event geht es?' : 'What kind of event is it?'}</div>
            <div className="dex-ui-grid-2" style={{ gap: 10 }}>
              {([
                {
                  id: 'internal' as const,
                  title: isDe ? 'Internes Event' : 'Internal event',
                  desc: isDe
                    ? 'Ein internes Deloitte Event — oder die Koordination der Deloitte-Teilnahme an einer externen Veranstaltung (z.B. B2Run).'
                    : 'A Deloitte-internal event — or coordinating Deloitte participation in an external event (e.g. B2Run).',
                },
                {
                  id: 'external' as const,
                  title: isDe ? 'Externes Event mit externen Teilnehmern' : 'External event with external attendees',
                  desc: isDe ? 'Gäste außerhalb von Deloitte melden sich selbst an.' : 'Guests from outside Deloitte register themselves.',
                },
              ]).map(o => (
                <button
                  key={o.id}
                  type="button"
                  role="radio"
                  aria-checked={eventScope === o.id}
                  className={cx('dex-ui-choice', eventScope === o.id && 'is-active')}
                  onClick={() => setEventScope(o.id)}
                  disabled={sending}
                >
                  <span className="dex-ui-choice-body">
                    <span className="dex-ui-choice-title" style={{ display: 'block' }}>{o.title}</span>
                    <span className="dex-ui-choice-desc" style={{ display: 'block' }}>{o.desc}</span>
                  </span>
                  <span className="dex-ui-choice-check" aria-hidden="true">{eventScope === o.id && <Check size={12} />}</span>
                </button>
              ))}
            </div>
            {eventScope === 'external' && (
              <div className="dex-ui-callout dex-ui-callout--danger" style={{ marginTop: 10 }}>
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <span>
                  {isDe
                    ? <><strong>Dafür ist DEX nicht nutzbar.</strong> Die Plattform ist auf interne Deloitte Events ausgelegt; externe Gäste bekommen keinen Zugang und können sich nicht selbst anmelden.<br /><br />Alles zu externen Veranstaltungen findest du hier:<br /><a href="https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx" target="_blank" rel="noopener noreferrer" style={{ color: 'inherit', fontWeight: 700 }}>Event Management im DeloitteNet</a></>
                    : <><strong>DEX cannot be used for this.</strong> The platform is built for Deloitte-internal events; external guests get no access and cannot register themselves.<br /><br />Everything about external events can be found here:<br /><a href="https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx" target="_blank" rel="noopener noreferrer" style={{ color: 'inherit', fontWeight: 700 }}>Event Management on DeloitteNet</a></>}
                </span>
              </div>
            )}
          </div>
        )}
        {!organizerMode && (
          <div className={dimmed ? 'dex-ui-card--muted' : undefined}>
            <label className="dex-ui-label" htmlFor="dexInquiryEventName">{isDe ? 'Wie heißt dein Event?' : 'What is your event called?'}</label>
            <input
              id="dexInquiryEventName"
              type="text"
              className="dex-ui-input"
              value={eventName}
              onChange={e => setEventName(e.target.value)}
              disabled={sending}
              placeholder={isDe ? 'z.B. Summer Party 2026' : 'e.g. Summer Party 2026'}
            />
          </div>
        )}
        {/* v29.45: Bedarfs-Checkliste. Steht VOR dem Freitext: erst ankreuzen,
            was es an Funktionen braucht, dann alles Übrige beschreiben. */}
        {!organizerMode && eventScope === 'internal' && (
          <div>
            <div className="dex-ui-label">
              {isDe ? 'Was brauchst du für dein Event?' : 'What do you need for your event?'}
              <span className="dex-ui-label-optional">(optional)</span>
            </div>
            <div className="dex-ui-help" style={{ margin: '-2px 0 10px' }}>
              {isDe
                ? 'Mehrfachauswahl — hilft uns, das Gespräch vorzubereiten. Das „i" zeigt, was DEX dafür schon kann. Unsicher? Einfach frei lassen.'
                : 'Multiple choice — helps us prepare. The „i" shows what DEX already offers. Not sure? Just leave it empty.'}
            </div>
            {/* v29.46: gruppiert statt einer langen Liste.
                v31.2: Toggle-Zeilen im Zweispalter statt nackter Checkboxen —
                jede Zeile hat damit Hover und Rahmen, die Auswahl ist auf einen
                Blick zu sehen; die Gruppen-Titel als Abschnitts-Zeile. */}
            {NEED_GROUPS.map(group => (
              <div key={group.id} className="dex-ui-section" style={{ marginTop: 12 }}>
                <div className="dex-ui-section-title">{isDe ? group.de : group.en}</div>
                <div className="dex-ui-grid-2" style={{ gap: 8 }}>
                  {group.ids
                    .map(id => NEED_OPTIONS.filter(o => o.id === id)[0])
                    .filter(Boolean)
                    .map(opt => {
                      const checked = needs.indexOf(opt.id) >= 0;
                      return (
                        <label key={opt.id} className={cx('dex-ui-toggle-row', checked && 'is-active', sending && 'is-disabled')}>
                          <input type="checkbox" checked={checked} onChange={() => toggleNeed(opt.id)} disabled={sending} />
                          <span className="dex-ui-toggle-row-body">
                            <span className="dex-ui-toggle-row-title" style={{ fontWeight: 500, fontSize: '0.84rem' }}>
                              {isDe ? opt.de : opt.en}
                              {/* v29.46: Was DEX dafür schon mitbringt — die Liste
                                  soll nicht nur abfragen, sondern auch
                                  beantworten, was die App an dieser Stelle kann. */}
                              <InfoTooltip text={isDe ? opt.infoDe : opt.infoEn} />
                            </span>
                          </span>
                        </label>
                      );
                    })}
                </div>
              </div>
            ))}
            {/* v29.46: Auffangbecken für alles, was die Liste nicht kennt. */}
            <div className="dex-ui-section" style={{ marginTop: 12 }}>
              <label className={cx('dex-ui-toggle-row', needs.indexOf('other') >= 0 && 'is-active', sending && 'is-disabled')}>
                <input type="checkbox" checked={needs.indexOf('other') >= 0} onChange={() => toggleNeed('other')} disabled={sending} />
                <span className="dex-ui-toggle-row-body">
                  <span className="dex-ui-toggle-row-title" style={{ fontWeight: 500, fontSize: '0.84rem' }}>
                    {isDe ? 'Sonstiges — etwas anderes' : 'Other — something else'}
                  </span>
                </span>
              </label>
              {needs.indexOf('other') >= 0 && (
                <input
                  type="text"
                  className="dex-ui-input"
                  value={otherNeed}
                  onChange={e => setOtherNeed(e.target.value)}
                  disabled={sending}
                  placeholder={isDe ? 'Was brauchst du noch?' : 'What else do you need?'}
                  style={{ marginTop: 8 }}
                />
              )}
            </div>
          </div>
        )}
        {/* v31.2: Beschriftung als Frage statt „Sonst noch etwas?" — das las
            sich optional, obwohl die Beschreibung in der Anfrage Pflicht ist
            (canSubmit). Im Organizer-Modus bleibt sie optional und sagt es. */}
        <div className={dimmed ? 'dex-ui-card--muted' : undefined}>
          <label className="dex-ui-label" htmlFor="dexInquiryMessage">
            {organizerMode
              ? (isDe ? 'Warum möchtest du Organizer werden?' : 'Why do you want to become an organizer?')
              : (isDe ? 'Was sollen wir über dein Event wissen?' : 'What should we know about your event?')}
            {organizerMode && <span className="dex-ui-label-optional">(optional)</span>}
          </label>
          <textarea
            id="dexInquiryMessage"
            className="dex-ui-textarea"
            value={message}
            onChange={e => setMessage(e.target.value)}
            disabled={sending}
            rows={organizerMode ? 3 : 5}
            placeholder={organizerMode
              ? (isDe ? 'Optional: kurz, worum es geht …' : 'Optional: briefly what it is about …')
              : (isDe
                ? 'Anzahl Teilnehmer, Termin, Besonderheiten …'
                : 'Number of participants, date, anything special …')}
          />
        </div>
        {status === 'success' && (
          <div className="dex-ui-callout dex-ui-callout--success" role="status">
            <span className="dex-ui-callout-icon"><Check size={16} /></span>
            <span>
              {organizerMode
                ? (isDe ? 'Antrag gesendet — die Admins prüfen ihn und schalten dich frei.' : 'Request sent — the admins will review and grant access.')
                : (isDe ? 'Anfrage gesendet — wir melden uns!' : 'Request sent — we will get back to you!')}
            </span>
          </div>
        )}
        {status === 'error' && (
          <div className="dex-ui-callout dex-ui-callout--danger" role="alert">
            <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
            <span>{isDe ? 'Senden fehlgeschlagen — versuch es später noch einmal.' : 'Sending failed — please try again later.'}</span>
          </div>
        )}
        {/* v24.24: „Dein Name" ist nicht mehr frei editierbar — stattdessen eine
            read-only Personen-Karte des eingeloggten Users (Foto, Name, Position,
            Standort), analog zur Teilnehmerliste. Diese Infos gehen mit in die
            Anfrage-Mail an die Admins.
            v31.2: Steht jetzt am Ende statt vor den Fragen — der Organizer
            beantwortet erst, worum es geht, und sieht zuletzt, mit welchen
            Daten die Anfrage rausgeht (wie ein Absender unter einem Brief). */}
        <div>
          <div className="dex-ui-muted" style={{ fontWeight: 600, marginBottom: 6 }}>
            {isDe ? 'Du fragst an als' : 'Requesting as'}
          </div>
          <div className="dex-ui-card dex-ui-card--soft" style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '10px 12px' }}>
            {photoUrl
              ? <img src={photoUrl} alt="" className="dex-ui-avatar dex-ui-avatar--lg" />
              : <span className="dex-ui-avatar dex-ui-avatar--lg" style={{ background: 'var(--dex-green, #86bc25)', color: '#fff' }}>{userInitials}</span>}
            <span style={{ minWidth: 0 }}>
              <span style={{ display: 'block', fontWeight: 600, color: 'var(--dex-gray-800)', fontSize: '0.92rem' }}>
                {userFullName || (currentUser.email || '')}
              </span>
              {(userJobTitle || userLocation) && (
                <span className="dex-ui-muted" style={{ display: 'block', fontSize: '0.78rem' }}>
                  {[userJobTitle, userLocation].filter(Boolean).join(' · ')}
                </span>
              )}
              {currentUser.email && (
                <span style={{ display: 'block', fontSize: '0.75rem', color: 'var(--dex-gray-400)' }}>
                  {currentUser.email}
                </span>
              )}
            </span>
          </div>
        </div>
      <LandingInfoModal
        open={showInfo}
        locale={locale === 'de' ? 'de' : 'en'}
        onClose={() => setShowInfo(false)}
      />
    </Modal>
  );
}
