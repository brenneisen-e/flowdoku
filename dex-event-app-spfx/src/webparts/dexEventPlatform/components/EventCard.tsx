/**
 * Event-Karte für die Übersichtsseite
 *
 * Zeigt Gradient-Hintergrund, Event-Infos und freie Plätze.
 * Die Gradient-Farben rotieren basierend auf dem Index.
 *
 * v31.9: Die Kachel beantwortet vier Fragen, und der Körper (`__body`) steht
 * seither in genau dieser Rangfolge: Bild/Titel/Ort (im Bild-Overlay) →
 * Zeitraum → Zustand (frei · voll · Warteliste · offene Termine) →
 * Ansprechpartner → Knopf. Die Geometrie (`.event-card__*`) liegt im
 * geteilten SCSS-Modul und wird zentral gepflegt — hier wird deshalb
 * ausschließlich INNERHALB von `__body` umsortiert; alles Weitere kommt über
 * die `dex-ui-`-Klassen (docs/ui-leitfaden.md). `ensureDexUiStyles()` ruft
 * die Seite (`EventListPage`), nicht diese Unterkomponente.
 */

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { useNavigation } from '../context/NavigationContext';
import { useLanguage } from '../context/LanguageContext';
import { useRoles } from '../context/RoleContext';
import { useEvents } from '../context/EventContext';
import { DeloitteEvent } from '../types';
import { useCachedImageWithFallback } from '../utils/imageCache';
import { isRegistrationFullyClosed, isRegistrationOpen } from '../utils/eventFormat';
import OrganizerList from './OrganizerList';
import { AlertCircle, Calendar, Check } from './Icons';
// v24.91: Portal-Popover (Organizer-Kontakt im „Registration closed"-Overlay)
// wird an document.body gerendert — daher mit `styles.dexApp` wrappen, damit
// die gescopten CSS-Variablen greifen (analog Modal-Fix v24.65).
import styles from './DexEventPlatform.module.scss';

// Deutsches Datumsformat
function formatDate(iso: string): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleDateString('de-DE', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  }) + ' ' + d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
}

// Nur Datum (ohne Uhrzeit) - wird z.B. für die Registration Deadline verwendet,
// die im Formular ausschliesslich als Datum gepflegt wird.
function formatDateOnly(iso: string): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleDateString('de-DE', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  });
}

// 4 verschiedene Farbverläufe, rotieren durch
// type wird aktuell nicht genutzt, könnte man später für typspezifische Farben verwenden
function getEventGradient(_type: string, index: number): string {
  const gradients = [
    'linear-gradient(135deg, #0a2e1a 0%, #1a6b3c 40%, #00ff88 100%)',
    'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
    'linear-gradient(135deg, #2d5016 0%, #86bc25 50%, #c5e63c 100%)',
    'linear-gradient(135deg, #1a0a2e 0%, #4a1a6b 40%, #8800ff 100%)',
  ];
  return gradients[index % gradients.length];
}

interface Props {
  event: DeloitteEvent;
  index: number;
  isRegistered?: boolean;
  isWaitlisted?: boolean;
  /** v11.90: eingeloggter User ist Haupt- oder Co-Organizer dieses Events. */
  isOwnOrganizer?: boolean;
}

export default function EventCard({ event, index, isRegistered, isWaitlisted, isOwnOrganizer }: Props): React.ReactElement {
  const { navigate } = useNavigation();
  const { t, locale } = useLanguage();
  const isDe = locale === 'de';
  const { canCreateEvents } = useRoles();
  const { childEventsOf } = useEvents();
  // v19.22: Event-Bild über den IndexedDB-Cache — beim zweiten App-Aufruf sofort
  // da, ohne SharePoint-Roundtrip.
  // v28.11: Für den Kachel-Hintergrund (cover) das UNBESCHNITTENE Querformat-
  // Original bevorzugen, falls vorhanden — ein Kreis-Zuschnitt wirkt als
  // Kachel-Hintergrund verloren, das Original füllt die Fläche.
  // v29.34: … aber mit Rückfall. Das Original ist ein zweites Attachment, und
  // seine URL kann tot sein (fehlgeschlagener Upload, späterer Bildwechsel).
  // Ohne Rückfall blieb die Kachel weiß, obwohl das Event ein Bild hat.
  const cachedImage = useCachedImageWithFallback(event.imageOrigUrl, event.imageUrl);
  // v29.35: Kein Event-Bild, aber ein Mail-Logo? Dann zeigt die Anmeldeseite
  // seit v29.13 dieses Logo — die Kachel blieb beim Farbverlauf. Dasselbe
  // Event hatte damit an einer Stelle ein Foto und an der anderen keins.
  // Das Logo kommt hier NICHT als Cover-Hintergrund (dafür wurde es v29.13
  // bewusst ausgelassen: ein Logo zerfällt im Beschnitt), sondern vollständig
  // über dem Verlauf — der Verlauf bleibt der Rahmen, das Bild bleibt heil.
  const mailLogoFallback = (!event.imageUrl && event.mailImageBase64) ? event.mailImageBase64 : '';
  // v9.8: B2Run-Events haben maxParticipants=0, weil die Kapazität auf
  // Durchstarter + Funstarter aufgeteilt ist. Die Summe gilt als
  // Gesamtkapazität — sonst zeigt die Karte fälschlich "Unbegrenzt", obwohl
  // es z.B. 140 Plätze gibt.
  const splitCapacity = (event.durchstarterCapacity || 0) + (event.funstarterCapacity || 0);
  const effectiveMax = event.maxParticipants && event.maxParticipants > 0
    ? event.maxParticipants
    : splitCapacity;
  const isUnlimited = !effectiveMax || effectiveMax === 0;
  // v29.13: Besteht ein Event ausschließlich aus Sub-Events, ist das Hauptevent
  // gar nicht buchbar — die Plätze liegen einzeln bei den Sub-Events. Die
  // Kachel zeigte trotzdem den Zähler des Hauptevents, und weil dort keine
  // Kapazität gepflegt wird, stand da „Unbegrenzt". Das ist nicht nur
  // nichtssagend, es ist das Gegenteil der Wahrheit: jedes Sub-Event hat seine
  // eigene Obergrenze. Dasselbe gilt für die Anmeldefrist — bei
  // unterschiedlichen Sub-Event-Fristen gibt es keine gültige Frist für die
  // Klammer. Beides bleibt deshalb weg; die Zahlen stehen auf der Anmeldeseite
  // pro Sub-Event.
  const subOnly = !!event.subEventsOnlyMode;
  // v24.72: Wartelisten-Anzahl abziehen — ein frei gewordener Platz geht IMMER
  // zuerst an die Warteliste und ist daher nicht „frei" für neue Anmeldungen.
  // Verhindert auch das kurze, falsche „1 frei" während des Nachrückens.
  const freePlaces = isUnlimited ? Infinity : effectiveMax - event.currentParticipants - (event.waitlistCount || 0);
  const isFull = !isUnlimited && freePlaces <= 0;
  const alreadySignedUp = isRegistered || isWaitlisted;
  const childEvents = childEventsOf(event.id);
  // v22.54: Die Anmeldung bleibt offen, solange das Hauptevent ODER mindestens
  // ein Sub-Event noch offen ist — eine abgelaufene Klammer-/Hauptevent-Frist
  // sperrt nicht mehr das ganze Event, wenn die Sub-Events noch laufen.
  const isDeadlinePassed = isRegistrationFullyClosed(event, childEvents);
  // v31.9: Bei `subEventsOnlyMode` schweigt die Kachel seit v29.13 zu Plätzen
  // und Frist — zu Recht, beides gilt dort nur je Termin. Nur stand an der
  // Stelle danach GAR NICHTS: Die Kachel sagte kein Wort darüber, dass die
  // Entscheidung eine Ebene tiefer liegt. An ihre Stelle treten deshalb die
  // Zahl der noch offenen Termine und ein Satz dazu.
  // Gefiltert wie die Anmeldeseite: soft-deaktivierte Sub-Events zählen nie,
  // Entwürfe (isFictive) nur für Organizer. Der Zielgruppen-Filter der
  // Anmeldeseite (`isEventVisibleForUser`) fehlt hier bewusst — er bräuchte
  // Nutzerdaten, die die Kachel nicht hat. Deshalb sagt der Text „offen" und
  // nicht „für dich buchbar".
  const openChildCount = event.subEventsDisabled
    ? 0
    : childEvents.filter(ce => (!ce.isFictive || canCreateEvents || isOwnOrganizer) && isRegistrationOpen(ce)).length;
  // v31.9: Bezeichnung immer aus den Term-Konstanten (Leitfaden 6d) — im
  // `subEventsOnlyMode` heißen die Kinder für Teilnehmer nicht „Sub-Events",
  // sondern schlicht „Events" (v29.13); ein eigener Begriff des Organizers
  // („Office-Tage", „Sessions") gewinnt davor.
  const childTermSingular = event.childEventTermSingular || (isDe ? 'Event' : 'event');
  const childTermPlural = event.childEventTermPlural || (isDe ? 'Events' : 'events');
  // Nur normale User bekommen den Deadline-Overlay. Organizer/Admins dürfen
  // trotzdem reinklicken, um ggf. manuell zu registrieren.
  // v24.90: isOwnOrganizer ergänzt — per-Event-Co-Organizer (im Wizard zum
  // Event hinzugefügt, OHNE globale Organizer-/Admin-Rolle, also
  // canCreateEvents=false) konnten ein nach Frist geschlossenes Event sonst
  // nicht öffnen und niemanden mehr anmelden. Sie sind aber vollwertige
  // Organizer DIESES Events und müssen reinklicken können.
  const showDeadlineOverlay = isDeadlinePassed && !canCreateEvents && !isOwnOrganizer && !alreadySignedUp;
  // v23.14: Vorschau vor Aktivierung („Aktiv ab" in der Zukunft + previewBeforeActive).
  // Reguläre User sehen die Karte, dürfen aber NICHT in die Anmeldeseite —
  // blockierender Overlay „Anmeldung ab …". Organizer/Admins dürfen weiterhin
  // rein (zum Vorbereiten) und bekommen stattdessen nur einen Hinweis-Badge.
  const activeFromTs = event.activeFrom ? new Date(event.activeFrom).getTime() : 0;
  const notYetActive = activeFromTs > 0 && activeFromTs > Date.now();
  const showPreviewOverlay = notYetActive && !canCreateEvents && !isOwnOrganizer && !alreadySignedUp;
  const showOrganizerActiveBadge = notYetActive && (canCreateEvents || isOwnOrganizer);
  const blockClick = showDeadlineOverlay || showPreviewOverlay;

  // v24.91: Organizer-Kontakt im „Registration closed"-Overlay — das Wort
  // „Organizer" wird unterstrichen; Hover/Klick blendet die Kontaktkarte(n)
  // der Veranstalter ein (Foto, E-Mail, Teams). Portal an document.body, weil
  // die Karte overflow:hidden + Hover-transform hat (würde ein In-Card-
  // Popover clippen).
  const orgNames = (event.organizers || []).reduce<string[]>((acc, o) => [...acc, ...o.split(';')], []).map(o => o.trim()).filter(Boolean);
  const hiddenOrgEmails = (event.hideOrganizer && event.hideOrganizerIndividualOnly) ? (event.hiddenOrganizerEmails || []) : [];
  const allOrgsHidden = !!event.hideOrganizer && !event.hideOrganizerIndividualOnly;
  const hasOrgContacts = !allOrgsHidden && orgNames.length > 0;
  // v31.9: Für die Kontaktzeile im Karten-Körper reicht `hasOrgContacts` nicht:
  // Sind ALLE Organizer einzeln ausgeblendet, liefert `OrganizerList` null —
  // übrig bliebe die Überschrift „Fragen?" ohne eine einzige Person.
  const hasVisibleOrgs = hasOrgContacts && orgNames.length > hiddenOrgEmails.length;
  const orgTriggerRef = React.useRef<HTMLSpanElement>(null);
  const [orgOpen, setOrgOpen] = React.useState(false);
  const [orgCoords, setOrgCoords] = React.useState<{ x: number; y: number } | null>(null);
  const orgCloseTimer = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const cancelOrgClose = (): void => { if (orgCloseTimer.current) { clearTimeout(orgCloseTimer.current); orgCloseTimer.current = null; } };
  const openOrg = (): void => { cancelOrgClose(); const r = orgTriggerRef.current?.getBoundingClientRect(); if (r) setOrgCoords({ x: r.left + r.width / 2, y: r.bottom + 6 }); setOrgOpen(true); };
  const scheduleOrgClose = (): void => { cancelOrgClose(); orgCloseTimer.current = setTimeout(() => setOrgOpen(false), 220); };
  React.useEffect(() => () => cancelOrgClose(), []);

  // v31.9: Der dunkle Karten-Overlay stand dreimal wortgleich als Inline-Style
  // in dieser Datei (Frist, Vorschau, angemeldet). Eine `dex-ui-`-Klasse dafür
  // gibt es nicht — bis sie im Leitfaden steht, ist EIN Objekt in dieser Datei
  // die kleinere Doppelung.
  const overlayStyle: React.CSSProperties = {
    position: 'absolute', inset: 0, zIndex: 10, borderRadius: 'var(--dex-radius)',
    background: 'linear-gradient(180deg, rgba(0,0,0,0.72) 0%, rgba(0,0,0,0.62) 100%)',
    display: 'flex', flexDirection: 'column', gap: 4,
    alignItems: 'center', justifyContent: 'center', padding: 24, textAlign: 'center',
  };
  const overlayTitleStyle: React.CSSProperties = { color: '#fff', fontWeight: 700, fontSize: '1rem', display: 'inline-flex', alignItems: 'center', gap: 7 };
  const overlaySubtitleStyle: React.CSSProperties = { color: 'rgba(255,255,255,0.6)', fontSize: '0.8rem' };

  // v31.9: Ein Eintages-Event bekommt EINE Zeile, kein Fragment. Vorher stand
  // hier stur „<Start> bis" + Zeilenumbruch + „<Ende>" — bei leerem `endDate`
  // (Altbestand; der Rückfall auf `startDate` greift erst seit v22 beim
  // Speichern) endete die Kachel mitten im Satz, gefolgt von einer Leerzeile.
  // Liegen Start und Ende am selben Tag, steht das Datum einmal und dahinter
  // die Zeitspanne — „01.02.2026 09:00 – 17:00" statt zweier voller Zeilen.
  const startTxt = formatDate(event.startDate);
  const endTxt = formatDate(event.endDate);
  const dateText = !startTxt
    ? endTxt
    : (!endTxt || endTxt === startTxt)
      ? startTxt
      : (startTxt.slice(0, 10) === endTxt.slice(0, 10)
        ? `${startTxt} – ${endTxt.slice(11)}`
        : `${startTxt} ${t('events.until')} ${endTxt}`);

  return (
    <div className="event-card" style={{ position: 'relative', cursor: blockClick ? 'not-allowed' : 'pointer', ...(event.isDemoShowcase ? { outline: '2px dashed var(--dex-blue, #0076a8)', outlineOffset: 2 } : {}) }} onClick={() => (!alreadySignedUp && !blockClick) ? navigate('registration', event.id) : undefined}>
      {/* v17.25: Demo-Showcase-Event deutlich markieren (blaues DEMO-Badge
          oben rechts, gestrichelter Rahmen). Nur im Demo-Impersonation-Modus
          überhaupt in der Liste. */}
      {event.isDemoShowcase && (
        <div style={{
          position: 'absolute', top: 10, right: 10, zIndex: 6,
          padding: '3px 10px', borderRadius: 999,
          background: 'var(--dex-blue, #0076a8)', color: '#fff',
          fontSize: '0.68rem', fontWeight: 700, letterSpacing: 1,
          boxShadow: '0 1px 4px rgba(0,0,0,0.15)',
        }}>
          DEMO
        </div>
      )}
      {event.isFictive && !event.isDemoShowcase && (
        <div style={{
          position: 'absolute', top: 10, right: 10, zIndex: 5,
          padding: '3px 10px', borderRadius: 999,
          background: 'var(--dex-orange, #ed8b00)', color: '#fff',
          fontSize: '0.68rem', fontWeight: 700, letterSpacing: 0.5,
          boxShadow: '0 1px 4px rgba(0,0,0,0.15)',
        }}>
          {t('create.fictive.badge')}
        </div>
      )}
      {/* v11.90: Eigene Events (User selbst Haupt- oder Co-Organizer)
          bekommen ein grünes „Organizer"-Badge oben links — auch in
          Kombination mit dem orangen „Entwurf"-Badge oben rechts. */}
      {isOwnOrganizer && (
        <div style={{
          position: 'absolute', top: 10, left: 10, zIndex: 5,
          padding: '3px 10px', borderRadius: 999,
          background: 'var(--dex-green, #86bc25)', color: '#fff',
          fontSize: '0.68rem', fontWeight: 700, letterSpacing: 0.5,
          boxShadow: '0 1px 4px rgba(0,0,0,0.15)',
        }}>
          Organizer
        </div>
      )}
      {/* v23.14: Organizer/Admin sehen bei „Aktiv ab" in der Zukunft einen
          Hinweis-Banner „Anmeldung ab …" (nicht blockierend — sie dürfen rein,
          um das Event vorzubereiten). Regulären Usern wird stattdessen der
          blockierende Vorschau-Overlay gezeigt. */}
      {showOrganizerActiveBadge && (
        <div style={{
          position: 'absolute', top: 0, left: 0, right: 0, zIndex: 6,
          padding: '5px 10px',
          background: 'var(--dex-green)', color: '#fff',
          fontSize: '0.72rem', fontWeight: 700, textAlign: 'center',
          borderTopLeftRadius: 'var(--dex-radius)', borderTopRightRadius: 'var(--dex-radius)',
        }}>
          {t('events.regfrom')} {formatDate(event.activeFrom || '')}
        </div>
      )}
      {showDeadlineOverlay && (
        <div style={overlayStyle}>
          <div style={overlayTitleStyle}>
            <AlertCircle size={18} />
            {t('events.deadlinepassed')}
          </div>
          <div style={{ color: 'rgba(255,255,255,0.85)', fontSize: '0.85rem', maxWidth: 320, lineHeight: 1.4 }}>
            {!hasOrgContacts ? t('events.deadlinepassed.hint') : (
              // v31.8: Der Satz steht als drei Schlüssel da (pre/word/post),
              // damit das anklickbare Wort ein eigener Text ist. Vorher wurde
              // die übersetzte Zeichenkette mit /(organizer)/i zerlegt — wer
              // den Satz umformuliert hätte, ohne das Wort „Organizer" zu
              // treffen, hätte die Kontaktkarte lautlos entfernt.
              <>
                {t('events.deadlinepassed.hint.pre')}
                <span
                  ref={orgTriggerRef}
                  onMouseEnter={openOrg}
                  onMouseLeave={scheduleOrgClose}
                  onClick={(e) => { e.stopPropagation(); if (orgOpen) setOrgOpen(false); else openOrg(); }}
                  style={{ textDecoration: 'underline', textUnderlineOffset: 3, cursor: 'pointer', fontWeight: 700, color: '#fff' }}
                >{t('events.deadlinepassed.hint.word')}</span>
                {t('events.deadlinepassed.hint.post')}
              </>
            )}
          </div>
          <div style={overlaySubtitleStyle}>
            {event.title}
          </div>
        </div>
      )}
      {/* v24.91: Organizer-Kontaktkarte (Portal an document.body, mit
          styles.dexApp-Scope für die CSS-Variablen). */}
      {orgOpen && orgCoords && hasOrgContacts && ReactDOM.createPortal(
        <div
          className={styles.dexApp}
          onMouseEnter={openOrg}
          onMouseLeave={scheduleOrgClose}
          onClick={(e) => e.stopPropagation()}
          style={{ position: 'fixed', top: orgCoords.y, left: (() => { const hw = Math.min(180, window.innerWidth * 0.45); return Math.min(Math.max(orgCoords.x, hw), window.innerWidth - hw); })(), transform: 'translateX(-50%)', zIndex: 11000, background: '#fff', borderRadius: 12, padding: 14, boxShadow: '0 10px 30px rgba(0,0,0,0.28)', border: '1px solid var(--dex-gray-200, #e1e1e1)', maxWidth: 'min(360px, 90vw)' }}
        >
          <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600, #666)', marginBottom: 8, textAlign: 'center' }}>
            {isDe ? 'Bei Fragen wende dich gerne an:' : 'For questions, feel free to reach out to:'}
          </div>
          <OrganizerList names={orgNames} emails={event.organizerEmails} hiddenEmails={hiddenOrgEmails} display="card" size="sm" forceIsDe={isDe} />
        </div>,
        document.body
      )}
      {/* v23.14: Vorschau-Overlay für reguläre User — sichtbar, aber Anmeldung
          erst ab dem Aktivierungszeitpunkt (Anmeldeseite nicht öffenbar). */}
      {showPreviewOverlay && (
        <div style={overlayStyle}>
          <div style={overlayTitleStyle}>
            <Calendar size={18} strokeWidth={2} />
            {t('events.previewsoon')}
          </div>
          <div style={{ color: 'rgba(255,255,255,0.9)', fontSize: '0.9rem' }}>
            {t('events.regfrom')} {formatDate(event.activeFrom || '')}
          </div>
          <div style={overlaySubtitleStyle}>
            {event.title}
          </div>
        </div>
      )}
      {alreadySignedUp && (
        <div style={overlayStyle}>
          <div style={overlayTitleStyle}>
            {/* v31.9: Haken bzw. Uhr-Hinweis vor dem Wort — auf dem Handy ist
                die Kachel oft das Einzige, was jemand von seinem Status sieht. */}
            {(isWaitlisted && !isRegistered) ? <AlertCircle size={18} /> : <Check size={18} />}
            {/* v30.2: „Angemeldet" gewinnt. Bei einer Termin-Reihe ist man oft
                für viele Tage angemeldet und nur bei EINEM auf der Warteliste —
                die Kachel behauptete dann pauschal „Warteliste". */}
            {(isWaitlisted && !isRegistered) ? t('status.waitlist') : t('status.registered')}
          </div>
          <div style={{ color: 'rgba(255,255,255,0.7)', fontSize: '0.8rem', marginBottom: 12 }}>
            {event.title}
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8, alignItems: 'center' }}>
            <button
              className="btn btn-primary"
              style={{ fontSize: '0.85rem', padding: '8px 20px' }}
              onClick={(e: React.MouseEvent) => { e.stopPropagation(); navigate('my-events'); }}
            >
              {t('myevents.title')}
            </button>
            {/* v30.2: auch per-Event-Organizer (isOwnOrganizer) — der Listen-
                View kann das seit v24.90, das Kachel-Overlay sperrte sie aus:
                Ein Organizer auf der Warteliste kam von hier aus nicht mehr
                an „Für andere Person registrieren". */}
            {(canCreateEvents || isOwnOrganizer) && (
              <button
                className="btn btn-secondary"
                style={{ fontSize: '0.8rem', padding: '6px 16px' }}
                onClick={(e: React.MouseEvent) => {
                  e.stopPropagation();
                  navigate('registration', event.id, 'register-other');
                }}
              >
                {t('reg.registerother')}
              </button>
            )}
          </div>
        </div>
      )}
      <div className="event-card__image" style={{
        // v23.19/v23.20: Bei eigener Karten-Darstellung weißer Hintergrund, damit
        // ein Minus-Zoom (Bild kleiner) drumherum WEISS zeigt statt des
        // Cover-Bildes. Sonst wie bisher (cover bzw. Gradient-Fallback).
        background: (event.imageUrl && event.imageDisplay?.card)
          ? '#fff'
          : event.imageUrl
            ? `url(${cachedImage}) center/cover no-repeat`
            : getEventGradient(event.type, index),
        position: 'relative',
      }}>
        {/* v23.19: Optionale Pro-Ansicht-Darstellung (Karte) — überlagert das
            Hintergrund-Bild mit individuellem Zoom/Position. Nur wenn gesetzt. */}
        {event.imageUrl && event.imageDisplay?.card && (
          <img
            src={cachedImage}
            alt=""
            style={{
              // v23.23: contain statt cover — der (Kreis-)Bildinhalt wird nie
              // abgeschnitten; „Größe" < 1 verkleinert mit weißem Rand drumherum.
              position: 'absolute', inset: 0, width: '100%', height: '100%', objectFit: 'contain',
              transform: `scale(${event.imageDisplay.card.zoom})`,
              transformOrigin: 'center center',
            }}
          />
        )}
        {/* v29.35: Mail-Logo als Rückfall — contain, damit nichts abgeschnitten
            wird, mit etwas Luft und Platz für den Titel-Streifen unten. */}
        {mailLogoFallback && (
          <img
            src={mailLogoFallback}
            alt=""
            style={{
              position: 'absolute', left: 0, right: 0, top: 0, bottom: 0,
              width: '100%', height: '100%', objectFit: 'contain',
              padding: '10px 10px 46px 10px', boxSizing: 'border-box',
            }}
          />
        )}
        {/* v23.24: position+z-index, damit der Titel-Overlay IMMER über dem
            absolut positionierten (angepassten) Bild liegt — sonst überdeckt
            das Bild den Event-Namen. Freie Plätze stehen nur noch im grünen
            Badge im Karten-Body (kein doppelter Zähler mehr). */}
        <div className="event-card__overlay" style={{ position: 'relative', zIndex: 1 }}>
          <h3 className="event-card__title">{event.title}</h3>
          <div className="event-card__meta">
            <span>{event.location}</span>
          </div>
        </div>
      </div>
      <div className="event-card__body" style={{ display: 'flex', flexDirection: 'column', flex: 1 }}>
        {/* v31.9: 1. Wann. Eigene Zeile statt „Datum links, Badge rechts" —
            der Zeitraum ist die Frage, die nach Titel und Ort kommt, und ein
            Badge daneben drängt ihn auf schmalen Kacheln in den Umbruch.
            `dex-ui-meta` liefert Symbol, Abstand und Umbruch (v31.8). */}
        {dateText && (
          <div className="event-card__dates dex-ui-meta">
            <span className="dex-ui-meta-item">
              <Calendar size={14} strokeWidth={2} />
              {dateText}
            </span>
          </div>
        )}
        {/* v31.9: 2. Kann ich noch buchen? Alles zum Zustand in EINER Zeile —
            freie Plätze bzw. offene Termine, daneben die Frist. Die Zeile
            entfällt ganz, wenn sie nichts zu sagen hat — eine leere Flex-Zeile
            zieht sonst den 8-px-Abstand des `__body` ins Leere. */}
        {(!subOnly || openChildCount > 0) && (
        <div className="dex-ui-inline">
          {!subOnly && (
            <span className={isFull ? 'dex-ui-pill dex-ui-pill--red' : 'dex-ui-pill dex-ui-pill--green'}>
              {isFull ? t('status.waitlist') : (isUnlimited ? t('reg.unlimited') : `${freePlaces} ${t('reg.free')}`)}
            </span>
          )}
          {/* v31.9: Im `subEventsOnlyMode` entfallen Plätze und Frist (v29.13).
              Die Zahl der noch offenen Termine tritt an ihre Stelle — aber nur,
              wenn sie größer als 0 ist: Eine 0 wäre hier keine Aussage über das
              Event, sondern über eine womöglich gar nicht gelesene Liste
              (CLAUDE.md: ein Lesefehler ist keine Null). */}
          {subOnly && openChildCount > 0 && (
            <span className="dex-ui-pill dex-ui-pill--blue">
              {isDe
                ? `noch ${openChildCount} ${openChildCount === 1 ? childTermSingular : childTermPlural} offen`
                : `${openChildCount} ${openChildCount === 1 ? childTermSingular : childTermPlural} still open`}
            </span>
          )}
          {!subOnly && event.registrationDeadline && formatDateOnly(event.registrationDeadline) && (
            <span className="event-card__deadline">
              {t('events.regopen')} {formatDateOnly(event.registrationDeadline)}
            </span>
          )}
        </div>
        )}
        {/* v31.9: … und der Satz dazu. Ohne ihn schwieg die Kachel genau dort,
            wo die Entscheidung liegt: eine Ebene tiefer, je Termin.
            Bezeichnung über die Term-Konstanten, nie fest verdrahtet. */}
        {subOnly && (
          <div className="dex-ui-muted">
            {isDe
              ? `Du meldest dich je ${childTermSingular} einzeln an.`
              : `You register for each ${childTermSingular} separately.`}
          </div>
        )}
        {/* Der rote Kasten ist KEINE Dublette zum Overlay: Er rendert bei
            abgelaufener Frist für ALLE — auch für Organizer und bereits
            Angemeldete, die den Overlay nie sehen. */}
        {isDeadlinePassed && (
          <div className="dex-ui-callout dex-ui-callout--danger dex-ui-callout--sm">
            <span className="dex-ui-callout-icon"><AlertCircle size={14} /></span>
            <span className="dex-ui-callout-body" style={{ fontWeight: 600 }}>{t('events.deadlinepassed')}</span>
          </div>
        )}
        {/* v31.9: 3. Wen frage ich? Bis hierher waren die Organizer NUR im
            Frist-Overlay erreichbar — also erst, wenn die Anmeldung zu ist,
            und über ein Hover-Portal, das es auf dem Handy nicht gibt
            (Leitfaden 6b). Die Datenschutz-Schalter wandern mit: `allOrgsHidden`
            (hideOrganizer ohne Einzelauswahl) blendet die Zeile ganz aus,
            `hiddenOrgEmails` nimmt einzeln versteckte Personen heraus. Die
            Chips stoppen die Klick-Weitergabe selbst — die Karte navigiert
            also nicht, wenn jemand einen Organizer antippt. */}
        {hasVisibleOrgs && (
          <div className="dex-ui-inline">
            <span className="dex-ui-muted">{isDe ? 'Fragen?' : 'Questions?'}</span>
            <OrganizerList names={orgNames} emails={event.organizerEmails} hiddenEmails={hiddenOrgEmails} size="sm" compact forceIsDe={isDe} />
          </div>
        )}
        <div style={{ marginTop: 'auto', paddingTop: 12 }}>
          <button
            className="btn btn-primary event-card__register-btn"
            onClick={(e: React.MouseEvent) => {
              e.stopPropagation();
              navigate('registration', event.id);
            }}
          >
            {/* v29.33: „Registrierung starten" — der Klick öffnet nur die
                Anmeldeseite, angemeldet ist man erst nach dem Absenden dort. */}
            {t('reg.registerstart')}
          </button>
        </div>
      </div>
    </div>
  );
}
