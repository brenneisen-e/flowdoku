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
import { Info } from './Icons';
import { InfoTooltip } from './InfoTooltip';

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

  // v13.1: Modal-Wrapper-Komponente — kapselt Backdrop/Escape/Padding.
  return (
    <Modal
      open={open}
      onClose={onClose}
      dismissable={!sending}
      ariaLabel={organizerMode ? (isDe ? 'Organizer werden' : 'Become an organizer') : (isDe ? 'DEX-Anfrage' : 'DEX inquiry')}
    >
        <h2 style={{ margin: 0, fontSize: '1.3rem', color: 'var(--dex-gray-800)' }}>
          {organizerMode
            ? (isDe ? 'Organizer werden' : 'Become an organizer')
            : (isDe ? 'DEX App für dein Event anfragen' : 'Request the DEX App for your event')}
        </h2>
        <p style={{ margin: 0, fontSize: '0.85rem', color: 'var(--dex-gray-500)' }}>
          {organizerMode
            ? (isDe
                ? 'Stelle einen Antrag, Organizer zu werden — dann kannst du eigene Events anlegen und verwalten. Die Admins prüfen deinen Antrag und schalten dich frei.'
                : 'Request to become an organizer — you can then create and manage your own events. The admins review your request and grant access.')
            : (isDe
                ? 'Wir melden uns kurz bei dir und besprechen, wie wir dein Event unterstützen können.'
                : 'We will get back to you and discuss how we can support your event.')}
        </p>
        {!organizerMode && <button
          type="button"
          onClick={() => setShowInfo(true)}
          style={{
            display: 'flex', alignItems: 'center', gap: 10,
            padding: '12px 14px',
            background: 'linear-gradient(135deg, rgba(134,188,37,0.10), rgba(0,118,168,0.06))',
            border: '1px solid var(--dex-green, #86bc25)',
            borderRadius: 10,
            cursor: 'pointer',
            textAlign: 'left',
            fontFamily: 'inherit',
            transition: 'all 0.15s',
          }}
          onMouseEnter={e => { e.currentTarget.style.background = 'linear-gradient(135deg, rgba(134,188,37,0.18), rgba(0,118,168,0.10))'; }}
          onMouseLeave={e => { e.currentTarget.style.background = 'linear-gradient(135deg, rgba(134,188,37,0.10), rgba(0,118,168,0.06))'; }}
        >
          <span style={{
            width: 32, height: 32, borderRadius: '50%', flexShrink: 0,
            background: 'var(--dex-green, #86bc25)', color: '#fff',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>
            <Info size={16} />
          </span>
          <span style={{ fontSize: '0.85rem', lineHeight: 1.35, color: 'var(--dex-gray-700)' }}>
            <strong style={{ display: 'block', color: 'var(--dex-gray-800)', marginBottom: 2 }}>
              {locale === 'de' ? 'Erst mal mehr über die DEX App erfahren?' : 'First learn more about the DEX App?'}
            </strong>
            {locale === 'de'
              ? 'Hier klicken — alle Funktionen, Zielgruppen und Beispiel-Events im Überblick.'
              : 'Click here — all features, audiences and example events at a glance.'}
          </span>
        </button>}
        {/* v24.24: „Dein Name" ist nicht mehr frei editierbar — stattdessen eine
            read-only Personen-Karte des eingeloggten Users (Foto, Name, Position,
            Standort), analog zur Teilnehmerliste. Diese Infos gehen mit in die
            Anfrage-Mail an die Admins. */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
          {locale === 'de' ? 'Du fragst an als' : 'Requesting as'}
          <div style={{
            display: 'flex', alignItems: 'center', gap: 12,
            padding: '10px 12px', borderRadius: 10,
            border: '1px solid var(--dex-gray-200)', background: 'var(--dex-gray-50, #fafafa)',
          }}>
            {photoUrl
              ? <img src={photoUrl} alt="" style={{ width: 44, height: 44, borderRadius: '50%', objectFit: 'cover', flexShrink: 0 }} />
              : <span style={{
                  width: 44, height: 44, borderRadius: '50%', flexShrink: 0,
                  background: 'var(--dex-green, #86bc25)', color: '#fff',
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  fontWeight: 700, fontSize: '0.95rem',
                }}>{userInitials}</span>}
            <span style={{ minWidth: 0 }}>
              <span style={{ display: 'block', fontWeight: 600, color: 'var(--dex-gray-800)', fontSize: '0.92rem' }}>
                {userFullName || (currentUser.email || '')}
              </span>
              {(userJobTitle || userLocation) && (
                <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
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
        {/* v28.40: Einsatzbereich klarstellen. Bis hierhin stand in der
            Anfrage-Strecke nirgends, für welche Art von Events DEX gedacht
            ist — die einzige Erwaehnung von „extern" war die technische
            Aussage „keine externen APIs" in der Info-Box, die man sogar
            falsch herum lesen kann. */}
        {!organizerMode && (
          <div style={{
            padding: '12px 14px', borderRadius: 8,
            background: 'var(--dex-gray-50, #f7f7f5)',
            border: '1px solid var(--dex-gray-200)',
            fontSize: '0.82rem', color: 'var(--dex-gray-700)', lineHeight: 1.5,
          }}>
            <div style={{ fontWeight: 700, marginBottom: 8 }}>
              {isDe ? 'Um was für ein Event geht es?' : 'What kind of event is it?'}
            </div>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, marginBottom: 8, cursor: 'pointer' }}>
              <input type="radio" name="dexEventScope" checked={eventScope === 'internal'}
                onChange={() => setEventScope('internal')} disabled={sending} style={{ marginTop: 3 }} />
              <span>
                <strong>{isDe ? 'Internes Event' : 'Internal event'}</strong>
                <span style={{ display: 'block', fontSize: '0.76rem', color: 'var(--dex-gray-600)' }}>
                  {isDe
                    ? 'Ein internes Deloitte Event — oder die Koordination der Deloitte-Teilnahme an einer externen Veranstaltung (z.B. B2Run).'
                    : 'A Deloitte-internal event — or coordinating Deloitte participation in an external event (e.g. B2Run).'}
                </span>
              </span>
            </label>
            <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, cursor: 'pointer' }}>
              <input type="radio" name="dexEventScope" checked={eventScope === 'external'}
                onChange={() => setEventScope('external')} disabled={sending} style={{ marginTop: 3 }} />
              <span>
                <strong>{isDe ? 'Externes Event mit externen Teilnehmern' : 'External event with external attendees'}</strong>
                <span style={{ display: 'block', fontSize: '0.76rem', color: 'var(--dex-gray-600)' }}>
                  {isDe ? 'Gäste außerhalb von Deloitte melden sich selbst an.' : 'Guests from outside Deloitte register themselves.'}
                </span>
              </span>
            </label>
            {eventScope === 'external' && (
              <div style={{
                marginTop: 10, padding: '10px 12px', borderRadius: 8,
                background: '#fef3f2', border: '1px solid var(--dex-red, #c00)',
                color: '#7a1f1c', fontSize: '0.8rem', lineHeight: 1.55,
              }}>
                {isDe
                  ? <><strong>Dafür ist DEX nicht nutzbar.</strong> Die Plattform ist auf interne Deloitte Events ausgelegt; externe Gäste bekommen keinen Zugang und können sich nicht selbst anmelden.<br /><br />Alles zu externen Veranstaltungen findest du hier:<br /><a href="https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx" target="_blank" rel="noopener noreferrer" style={{ color: '#7a1f1c', fontWeight: 700 }}>Event Management im DeloitteNet</a></>
                  : <><strong>DEX cannot be used for this.</strong> The platform is built for Deloitte-internal events; external guests get no access and cannot register themselves.<br /><br />Everything about external events can be found here:<br /><a href="https://mydeloittenet.de.deloitte.com/sites/CEO/Pages/Event-Management.aspx" target="_blank" rel="noopener noreferrer" style={{ color: '#7a1f1c', fontWeight: 700 }}>Event Management on DeloitteNet</a></>}
              </div>
            )}
          </div>
        )}
        {!organizerMode && (
          <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
            {locale === 'de' ? 'Event-Name' : 'Event name'}
            <input
              type="text"
              className="form-input"
              value={eventName}
              onChange={e => setEventName(e.target.value)}
              disabled={sending}
              placeholder={locale === 'de' ? 'z.B. Summer Party 2026' : 'e.g. Summer Party 2026'}
            />
          </label>
        )}
        {/* v29.45: Bedarfs-Checkliste. Steht VOR dem Freitext: erst ankreuzen,
            was es an Funktionen braucht, dann alles Übrige beschreiben. */}
        {!organizerMode && eventScope === 'internal' && (
          <div style={{
            border: '1px solid var(--dex-gray-200)', borderRadius: 8, padding: '10px 12px',
            background: 'var(--dex-gray-50, #fafafa)',
          }}>
            <div style={{ fontSize: '0.8rem', fontWeight: 700, color: 'var(--dex-gray-800, #333)' }}>
              {isDe ? 'Was brauchst du für dein Event?' : 'What do you need for your event?'}
            </div>
            <div style={{ fontSize: '0.75rem', color: 'var(--dex-gray-600)', margin: '2px 0 8px' }}>
              {isDe
                ? 'Mehrfachauswahl, alles optional — es hilft uns, das Gespräch vorzubereiten. Das „i" neben jedem Punkt erklärt, was DEX dafür schon mitbringt. Unsicher? Einfach frei lassen.'
                : 'Multiple choice, all optional — it helps us prepare. The „i" next to each item explains what DEX already offers. Not sure? Just leave it empty.'}
            </div>
            {/* v29.46: gruppiert statt einer langen Liste. */}
            {NEED_GROUPS.map(group => (
              <div key={group.id} style={{ marginBottom: 10 }}>
                <div style={{
                  fontSize: '0.72rem', fontWeight: 700, letterSpacing: 0.3,
                  color: 'var(--dex-green-dark, #4a7c1f)', textTransform: 'uppercase',
                  margin: '0 0 4px',
                }}>
                  {isDe ? group.de : group.en}
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  {group.ids
                    .map(id => NEED_OPTIONS.filter(o => o.id === id)[0])
                    .filter(Boolean)
                    .map(opt => {
                      const checked = needs.indexOf(opt.id) >= 0;
                      return (
                        <label
                          key={opt.id}
                          style={{
                            display: 'flex', alignItems: 'flex-start', gap: 8, fontSize: '0.82rem',
                            color: 'var(--dex-gray-700)', cursor: sending ? 'default' : 'pointer',
                            padding: '3px 4px', borderRadius: 6,
                            background: checked ? 'rgba(134,188,37,0.10)' : 'transparent',
                          }}
                        >
                          <input
                            type="checkbox"
                            checked={checked}
                            onChange={() => toggleNeed(opt.id)}
                            disabled={sending}
                            style={{ marginTop: 2, accentColor: 'var(--dex-green, #86bc25)' }}
                          />
                          <span>
                            {isDe ? opt.de : opt.en}
                            {/* v29.46: Was DEX dafür schon mitbringt — die Liste
                                soll nicht nur abfragen, sondern auch
                                beantworten, was die App an dieser Stelle kann. */}
                            <InfoTooltip text={isDe ? opt.infoDe : opt.infoEn} />
                          </span>
                        </label>
                      );
                    })}
                </div>
              </div>
            ))}
            {/* v29.46: Auffangbecken für alles, was die Liste nicht kennt. */}
            <div>
              <label
                style={{
                  display: 'flex', alignItems: 'flex-start', gap: 8, fontSize: '0.82rem',
                  color: 'var(--dex-gray-700)', cursor: sending ? 'default' : 'pointer',
                  padding: '3px 4px', borderRadius: 6,
                  background: needs.indexOf('other') >= 0 ? 'rgba(134,188,37,0.10)' : 'transparent',
                }}
              >
                <input
                  type="checkbox"
                  checked={needs.indexOf('other') >= 0}
                  onChange={() => toggleNeed('other')}
                  disabled={sending}
                  style={{ marginTop: 2, accentColor: 'var(--dex-green, #86bc25)' }}
                />
                <span>{isDe ? 'Sonstiges — etwas anderes' : 'Other — something else'}</span>
              </label>
              {needs.indexOf('other') >= 0 && (
                <input
                  type="text"
                  className="form-input"
                  value={otherNeed}
                  onChange={e => setOtherNeed(e.target.value)}
                  disabled={sending}
                  placeholder={isDe ? 'Was brauchst du noch?' : 'What else do you need?'}
                  style={{ marginTop: 4, fontSize: '0.82rem' }}
                />
              )}
            </div>
          </div>
        )}
        <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
          {organizerMode
            ? (isDe ? 'Warum möchtest du Organizer werden? (optional)' : 'Why do you want to become an organizer? (optional)')
            : (isDe ? 'Sonst noch etwas?' : 'Anything else?')}
          <textarea
            className="form-textarea"
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
        </label>
        {status === 'success' && (
          <div style={{ color: 'var(--dex-green)', fontSize: '0.85rem' }}>
            {organizerMode
              ? (isDe ? 'Antrag gesendet — die Admins prüfen ihn und schalten dich frei.' : 'Request sent — the admins will review and grant access.')
              : (isDe ? 'Anfrage gesendet — wir melden uns!' : 'Request sent — we will get back to you!')}
          </div>
        )}
        {status === 'error' && (
          <div style={{ color: 'var(--dex-red)', fontSize: '0.85rem' }}>
            {locale === 'de' ? 'Senden fehlgeschlagen. Bitte später erneut versuchen.' : 'Sending failed. Please try again later.'}
          </div>
        )}
        <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 4 }}>
          <button type="button" className="btn btn-secondary" onClick={onClose} disabled={sending}>
            {locale === 'de' ? 'Abbrechen' : 'Cancel'}
          </button>
          <button
            type="button"
            className="btn btn-primary"
            onClick={handleSubmit}
            disabled={sending || !canSubmit}
          >
            {sending
              ? (isDe ? 'Wird gesendet...' : 'Sending...')
              : organizerMode
                ? (isDe ? 'Antrag senden' : 'Send request')
                : (isDe ? 'Anfrage senden' : 'Send inquiry')}
          </button>
        </div>
      <LandingInfoModal
        open={showInfo}
        locale={locale === 'de' ? 'de' : 'en'}
        onClose={() => setShowInfo(false)}
      />
    </Modal>
  );
}
