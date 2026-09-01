/* EventCard — aus RegistrationPage.tsx ausgelagert (v30.66).
 * Station 1 der Anmeldeseite: Hero-Bild (inkl. Kreis-/Zoom-Logik aus v28.9),
 * Datum, Ort, Organizer und Beschreibung. Inhalt zeichengleich uebernommen. */
import * as React from 'react';
import { getCachedOrbBase64, replacePlaceholders } from '../../services/EmailTemplates';
import { DEX_ORB_PNG } from '../../data/brandLogos';
import { Icon } from '@fluentui/react/lib/Icon';
import { formatAllDayPeriod } from '../../utils/eventFormat';
import { formatDateRange } from './regHelpers';
import { Mail } from '../Icons';
import OrganizerList from '../OrganizerList';
import { formatOrganizerList } from '../../context/EventContext';
import { Locale } from '../../context/LanguageContext';
import { DeloitteEvent } from '../../types';

/** Station 1 — Dein Event: Bild, Eckdaten, Organizer, Beschreibung. */
export interface EventCardProps {
  cachedImage: string;
  cachedZoomImage: string;
  /** 140 auf dem Handy, sonst 170 — im Props-Objekt des Aufrufers weitet der
   *  Literal-Typ ohnehin auf `number` auf. */
  circleSize: number;
  currentUser: import("../../types/index").User;
  event: DeloitteEvent;
  heroImgUrl: string;
  imgAspectReady: boolean;
  imgCircleNotch: boolean;
  imgHovered: boolean;
  imgSlotH: 210 | 260 | 300;
  imgSlotW: 280 | 420 | 210 | 240;
  imgZoomed: boolean;
  isMobile: boolean;
  locale: Locale;
  setImgHovered: React.Dispatch<React.SetStateAction<boolean>>;
  setImgZoomed: React.Dispatch<React.SetStateAction<boolean>>;
  showOrbPlaceholder: boolean;
  usesMailImage: boolean;
}
export const EventCard: React.FC<EventCardProps> = (p) => {
  const { cachedImage, cachedZoomImage, circleSize, currentUser, event, heroImgUrl, imgAspectReady, imgCircleNotch, imgHovered, imgSlotH, imgSlotW, imgZoomed, isMobile, locale, setImgHovered, setImgZoomed, showOrbPlaceholder, usesMailImage } = p;
  return (
        <div
          className="registration-event"
          // v28.7: Kreis-Notch — der Kreis ragt über die Oberkante der Karte
          // hinaus. Dafür muss das overflow:hidden der Karte weichen und
          // oben Platz für den überstehenden Halbkreis geschaffen werden.
          // v28.98: DAS hier war der Grund, warum der Foto-Platzhalter oben
          // abgeschnitten wurde — er nutzt seit v28.91 dasselbe Notch-Layout,
          // stand aber nicht in dieser Bedingung. Ohne den zusätzlichen
          // Platz und mit dem overflow:hidden der Karte wird der überstehende
          // Halbkreis schlicht weggeschnitten. (Meine früheren Erklärungen —
          // negativer Rand, Vorschau-Banner — waren beide falsch.)
          style={(imgCircleNotch || showOrbPlaceholder) ? { overflow: 'visible', marginTop: circleSize / 2 } : undefined}
        >
          <div
            className="registration-event__card"
            style={{
              display: 'flex',
              // v28.3 („Geführte Schritte"): Desktop-Standard = Bild kompakt
              // links + Inhalt rechts. v28.5: Der Organizer kann im Wizard
              // stattdessen „Banner"-Layout wählen (event.imageBanner) — dann
              // liegt das Bild in voller Kartenbreite ÜBER den Infos (gut für
              // breite Querformat-Fotos). Handy: immer Bild oben.
              // v28.6: Infos LINKS, Bild RECHTS (row-reverse — das Bild steht
              // im DOM zuerst, wird aber rechts gerendert). Banner/Mobil: Bild oben.
              // v28.7: Kreis-Bilder → Spalte, der Kreis sitzt oben mittig.
              flexDirection: (isMobile || event.imageBanner || imgCircleNotch || showOrbPlaceholder) ? 'column' : 'row-reverse',
              gap: 16,
              alignItems: (isMobile || event.imageBanner || imgCircleNotch || showOrbPlaceholder) ? 'stretch' : 'flex-start',
            }}
          >
            {/* v28.3: Bild-Slot nur rendern, wenn das Event ein Bild hat —
                sonst stünde links ein leerer 300px-Block.
                v28.19: … und erst, wenn die Bildform-Analyse fertig ist
                (imgAspectReady) — sonst startete das Bild kurz im Seiten-Slot
                rechts und sprang dann in den Kreis. Banner-Layout hängt nicht
                von der Form ab und rendert sofort. */}
            {/* v28.90: Ohne Event-Foto blieb der Bild-Slot leer und die Karte
                sah anders aus als bei Events mit Bild — der erste Eindruck der
                Anmeldeseite hing damit daran, ob jemand ein Foto hochgeladen
                hat. Statt Leerraum steht dort jetzt das DEX-Logo. Nur auf dem
                Desktop: Auf dem Handy liegt das Bild ÜBER den Infos und würde
                Titel und Datum nach unten drücken. Kein Zoom-Knopf — es gibt
                nichts zu vergrößern. */}
            {showOrbPlaceholder && (
              <div
                className="registration-event__image"
                title={locale === 'de' ? 'Für dieses Event ist kein Bild hinterlegt.' : 'No image is set for this event.'}
                style={{
                  background: '#fff',
                  position: 'relative',
                  width: circleSize, height: circleSize, flex: '0 0 auto',
                  borderRadius: '50%',
                  border: '1px solid var(--dex-gray-200)',
                  boxShadow: '0 2px 12px rgba(0,0,0,0.10)',
                  alignSelf: 'center',
                  // v28.97: Exakt dasselbe Layout wie ein rundes EVENT-Bild
                  // (imgCircleNotch): der Kreis hängt zur Haelfte in der
                  // Kartenkante. In v28.95 hatte ich den negativen Rand
                  // herausgenommen, weil der Kreis oben abgeschnitten wirkte —
                  // damit sah der Platzhalter aber als EINZIGER anders aus als
                  // alle anderen Kreis-Bilder, mit einer Luecke darunter. Zwei
                  // Darstellungen für dieselbe Stelle sind schlechter als
                  // eine; deshalb zurueck auf das gemeinsame Layout. Sollte
                  // der Zuschnitt wieder auftreten, liegt die Ursache im
                  // Container darüber und gehört dort behoben, nicht hier.
                  marginTop: -(circleSize / 2 + 16),
                  display: 'flex', alignItems: 'center', justifyContent: 'center',
                  padding: 14,
                  boxSizing: 'border-box',
                }}
              >
                <img
                  src={getCachedOrbBase64() || DEX_ORB_PNG}
                  alt=""
                  style={{ display: 'block', width: '100%', height: '100%', objectFit: 'contain' }}
                />
              </div>
            )}
            {heroImgUrl && (event.imageBanner || imgAspectReady) && (
            <div
              className="registration-event__image"
              // v28.12: Hover zeigt das Lupen-Icon; die Großansicht öffnet
              // erst der Klick darauf.
              onMouseEnter={() => setImgHovered(true)}
              onMouseLeave={() => setImgHovered(false)}
              style={{
                position: 'relative',
                // v11.91: Hintergrund auf Weiß gesetzt — PNGs mit Transparenz
                // zeigten vorher den hellgrauen Hintergrund durch, was wie ein
                // unsauberer „grauer Rand" um Logos aussah.
                background: '#fff',
                borderRadius: 'var(--dex-radius)',
                overflow: 'hidden',
                // v28.3/v28.4: Desktop = fester Bild-Slot links, contain (kein
                // Crop — Event-Bilder sind oft Poster mit Text). Querformat
                // bekommt den breiteren 420er-Slot und sitzt vertikal mittig
                // neben den Infos; Hochkant/Quadrat den kompakten 300er-Slot.
                // Handy = volle Breite mit begrenzter Höhe.
                // v28.7: Kreis-/Quadrat-Bilder = eigener Kreis oben mittig,
                // ragt zur Hälfte über die Oberkante der Karte hinaus
                // (negativer marginTop gegen Karten-Padding + Halbkreis).
                ...(imgCircleNotch
                  ? {
                    width: circleSize, height: circleSize, flex: '0 0 auto',
                    borderRadius: '50%',
                    border: '1px solid var(--dex-gray-200)',
                    boxShadow: '0 2px 12px rgba(0,0,0,0.10)',
                    alignSelf: 'center',
                    marginTop: -(circleSize / 2 + 16),
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                  }
                  : isMobile
                  ? { width: '100%', maxHeight: 200, display: 'flex', justifyContent: 'center' }
                  : event.imageBanner
                  // v28.5: Banner-Layout — volle Kartenbreite, Höhe begrenzt,
                  // contain (kein Crop bei Postern mit Text).
                  ? { width: '100%', maxHeight: 320, display: 'flex', alignItems: 'center', justifyContent: 'center' }
                  : {
                    flex: `0 0 ${imgSlotW}px`,
                    maxWidth: imgSlotW,
                    maxHeight: imgSlotH,
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    alignSelf: 'center',
                  }),
              }}
            >
              {heroImgUrl && (
                <img
                  src={cachedImage}
                  alt={event.title}
                  // v11.56: 'contain', damit das Bild vollständig sichtbar
                  // bleibt (kein Crop). v28.3: Desktop einheitlich auf den
                  // kompakten 300er-Slot begrenzt; die Pro-Ansicht-Hero-
                  // Einstellung (Zoom/Höhe) wirkt weiter, gedeckelt auf den
                  // Slot, damit Station 1 kompakt bleibt.
                  // v28.6: borderRadius direkt am <img> — bei contain liegt die
                  // sichtbare Bildkante INNERHALB des Containers, dessen
                  // overflow-Rundung griff daher nicht. Kreis-PNGs (transparente
                  // Ecken) bleiben unverändert rund.
                  // v28.7: Im Kreis-Notch füllt das Bild den Kreis komplett
                  // (cover) — beim typischen Kreis-Zuschnitt (Quadrat mit
                  // transparenten Ecken) liegt die Bildkante exakt am Rand.
                  // v29.13: Ein als Rückfall gezeigtes MAIL-LOGO wird im Kreis
                  // nicht beschnitten (contain + Innenabstand) — Logos haben
                  // Ränder und Schrift, die ein cover-Zuschnitt anschneidet.
                  style={imgCircleNotch
                    ? (usesMailImage
                      ? { width: '100%', height: '100%', objectFit: 'contain', display: 'block', padding: 14, boxSizing: 'border-box' }
                      : { width: '100%', height: '100%', objectFit: 'cover', display: 'block' })
                    : isMobile
                    ? { width: '100%', maxHeight: 200, height: 'auto', objectFit: 'cover', display: 'block', borderRadius: 'var(--dex-radius)' }
                    : event.imageBanner
                    ? { maxWidth: '100%', maxHeight: 320, width: 'auto', height: 'auto', objectFit: 'contain', display: 'block', margin: '0 auto', borderRadius: 'var(--dex-radius)' }
                    : event.imageDisplay?.hero
                    ? { display: 'block', margin: '0 auto', maxWidth: '100%', maxHeight: Math.min(event.imageDisplay.hero.height ?? imgSlotH, imgSlotH), width: 'auto', height: 'auto', objectFit: 'contain', transform: `scale(${Math.min(event.imageDisplay.hero.zoom || 1, 1.5)})`, transformOrigin: 'center center', borderRadius: 'var(--dex-radius)' }
                    : { maxWidth: '100%', maxHeight: imgSlotH, width: 'auto', height: 'auto', objectFit: 'contain', display: 'block', borderRadius: 'var(--dex-radius)' }
                  }
                />
              )}
              {/* v11.91: Info-Button entfernt — die Beschreibung ist jetzt
                  immer ausgeklappt, kein Toggle mehr nötig. */}
              {/* v28.12: Lupen-Icon beim Hover (auf dem Handy immer sichtbar,
                  dort gibt es kein Hover) — Klick öffnet die Lightbox.
                  Mittig unten platziert, damit es auch im Kreis-Notch
                  (borderRadius 50% + overflow hidden) sichtbar bleibt. */}
              {(imgHovered || isMobile) && (
                <button
                  type="button"
                  onClick={() => setImgZoomed(true)}
                  title={locale === 'de' ? 'Bild vergrößern' : 'Enlarge image'}
                  aria-label={locale === 'de' ? 'Bild vergrößern' : 'Enlarge image'}
                  style={{
                    position: 'absolute', bottom: 8, left: '50%', transform: 'translateX(-50%)',
                    width: 30, height: 30, borderRadius: '50%', border: 'none',
                    background: 'rgba(0,0,0,0.55)', color: '#fff',
                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                    cursor: 'pointer', zIndex: 2,
                  }}
                >
                  <Icon iconName="ZoomIn" style={{ fontSize: 14 }} />
                </button>
              )}
              {/* v28.12: Lightbox — Klick irgendwo (oder aufs X) schließt.
                  fixed, entkommt dem overflow:hidden des Containers. */}
              {imgZoomed && (
                <div
                  onClick={() => setImgZoomed(false)}
                  style={{
                    position: 'fixed', inset: 0, zIndex: 3000,
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    background: 'rgba(0,0,0,0.6)', cursor: 'zoom-out', padding: 24,
                  }}
                >
                  <img
                    src={cachedZoomImage}
                    alt={event.title}
                    style={{
                      maxWidth: '82vw', maxHeight: '80vh', width: 'auto', height: 'auto',
                      objectFit: 'contain', background: '#fff', borderRadius: 12,
                      boxShadow: '0 16px 56px rgba(0,0,0,0.4)', padding: 8,
                    }}
                  />
                  <button
                    type="button"
                    onClick={() => setImgZoomed(false)}
                    aria-label={locale === 'de' ? 'Schließen' : 'Close'}
                    style={{
                      position: 'absolute', top: 18, right: 22,
                      width: 36, height: 36, borderRadius: '50%', border: 'none',
                      background: 'rgba(255,255,255,0.92)', color: 'var(--dex-gray-800)',
                      display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                      cursor: 'pointer', fontSize: '1.1rem', fontWeight: 700,
                    }}
                  >×</button>
                </div>
              )}
            </div>
            )}
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 10, padding: '4px 4px 4px 0' }}>
              <h4 style={{ fontSize: '1rem', margin: 0 }}>{event.title}</h4>
              {/* v11.91: Datum + Ort als prominente Badges mit Icon.
                  v11.93: Datum einzeilig (nowrap) — Box wächst auf
                  natürliche Breite. Der Ort-Kasten streckt sich auf
                  dieselbe Breite, damit beide Boxen visuell aligniert
                  sind. inline-flex + alignItems:stretch sorgt für gleiche
                  Breite ohne festen Wert. */}
              {/* v11.94: alignSelf:stretch + maxWidth:100% damit die Box
                  nicht über den Card-Rand rausragt; gleichzeitig wächst
                  sie auf die natürliche Breite des längeren Inhalts und
                  beide Boxen sind gleich breit. */}
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'stretch', gap: 6, maxWidth: '100%' }}>
                <div style={{
                  display: 'flex', alignItems: 'center', gap: 8,
                  padding: '8px 12px', borderRadius: 8,
                  background: 'rgba(134,188,37,0.10)', color: 'var(--dex-green-dark, #4a7c1f)',
                  fontSize: '0.88rem', fontWeight: 600,
                }}>
                  <Icon iconName="Calendar" style={{ fontSize: 16, flexShrink: 0 }} />
                  <span>
                    {/* v11.94: kompaktes Range-Format („–" statt „until"),
                        bei gleichem Tag nur einmal Datum + „HH:MM - HH:MM". */}
                    {/* v29.61: ganztägig zeigt keine Uhrzeiten. */}
                    {event.allDay
                      ? formatAllDayPeriod(event.startDate, event.endDate, locale === 'de')
                      : formatDateRange(event.startDate, event.endDate)}
                  </span>
                </div>
                {(event.location || (event.locationAddress && (event.locationAddress.street || event.locationAddress.city))) && (() => {
                  const addr = event.locationAddress;
                  const hasAddr = !!(addr && (addr.street || addr.city));
                  const hasStreet = !!(addr && addr.street);
                  const cityLine = addr ? [addr.zip, addr.city].filter(Boolean).join(' ') : '';
                  // v26.88: Nur Name + Stadt (KEINE Straße) → einzeilig „Name, Stadt"
                  // (z.B. „RheinEnergieSTADION, Köln") statt Name/Stadt untereinander.
                  const nameCityInline = !!(event.location && cityLine && !hasStreet);
                  // v26.82: Pin-Icon nur bei echter Mehrzeiler-Adresse oben
                  // ausrichten; bei einer Zeile (inkl. „Name, Stadt") zentrieren.
                  const multiLine = hasAddr && !nameCityInline;
                  return (
                  <div style={{
                    display: 'flex', alignItems: multiLine ? 'flex-start' : 'center', gap: 8,
                    padding: '8px 12px', borderRadius: 8,
                    background: 'rgba(0,86,166,0.08)', color: '#0a3766',
                    fontSize: '0.88rem',
                  }}>
                    <Icon iconName="POI" style={{ fontSize: 16, marginTop: multiLine ? 2 : 0, flexShrink: 0 }} />
                    <span>
                      {nameCityInline ? (
                        <>
                          <span style={{ fontWeight: 700 }}>{event.location}</span>
                          <span style={{ fontWeight: 400 }}>{`, ${cityLine}`}</span>
                        </>
                      ) : (
                        <>
                          {event.location && (
                            <span style={{ fontWeight: 700 }}>{event.location}</span>
                          )}
                          {addr && (addr.street || addr.city) && (
                            <>
                              <br />
                              <span style={{ fontWeight: 400 }}>
                                {[addr.street, addr.houseNo].filter(Boolean).join(' ')}
                                {(addr.zip || addr.city) && <br />}
                                {[addr.zip, addr.city].filter(Boolean).join(' ')}
                              </span>
                            </>
                          )}
                        </>
                      )}
                    </span>
                  </div>
                  );
                })()}
              </div>
              {/* v26.89: Reihenfolge getauscht — ANSPRECHPARTNER steht jetzt VOR
                  dem ORGANIZER-Block (vorher umgekehrt). */}
              {/* v10.16: Optionaler Ansprechpartner — frei eingegebene Person
                  außerhalb des App-User-Pools. Reines Anzeige-Feld; Mailto-Link
                  wenn Email gesetzt. Wird nur gerendert wenn mindestens Name
                  oder Email gepflegt sind. */}
              {(event.contactName || event.contactEmail || event.contactInfo) && (
                <div style={{ marginTop: 12 }}>
                  {/* v28.4: Überschrift AUSSERHALB der Box — gleiche Optik und
                      Position wie das ORGANIZER-Label darunter. */}
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6, color: 'var(--dex-gray-600)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 6, fontWeight: 600, fontSize: '0.85rem' }}>
                    <span style={{ display: 'inline-flex', flexShrink: 0 }}><Mail size={15} /></span>
                    <span>{locale === 'de' ? 'Ansprechpartner' : 'Contact'}</span>
                  </div>
                  <div style={{ padding: '10px 12px', background: 'var(--dex-gray-50, #f7f7f7)', borderRadius: 8, border: '1px solid var(--dex-gray-200)' }}>
                  {/* v28.5: Schriftgrößen wie in den Datums-/Ort-Boxen (0.88rem). */}
                  {event.contactName && (
                    <div style={{ fontSize: '0.88rem', fontWeight: 700, color: 'var(--dex-gray-800)' }}>{event.contactName}</div>
                  )}
                  {event.contactEmail && (
                    <div style={{ fontSize: '0.88rem', marginTop: 2 }}>
                      <a href={`mailto:${event.contactEmail}`} style={{ color: 'var(--dex-green, #86bc25)', textDecoration: 'none' }}>{event.contactEmail}</a>
                    </div>
                  )}
                  {event.contactInfo && (
                    <div style={{ fontSize: '0.88rem', color: 'var(--dex-gray-700)', marginTop: 4, whiteSpace: 'pre-wrap', lineHeight: 1.4 }}>{event.contactInfo}</div>
                  )}
                  </div>
                </div>
              )}
              {(() => {
                // v24.15: „Organizer ausblenden" ohne Einzel-Modus = ALLE aus.
                if (event.hideOrganizer && !event.hideOrganizerIndividualOnly) return null;
                // Organizer als Chips mit Foto (Hover-Enlarge). Namen werden von "Nachname, Vorname"
                // in "Vorname Nachname" normalisiert. v11.91: Label + Chip größer für bessere Lesbarkeit.
                const orgs = event.organizers.reduce<string[]>((acc, o) => [...acc, ...o.split(';')], []).map(o => {
                  const trimmed = o.trim();
                  const parts = trimmed.split(',').map(s => s.trim());
                  return parts.length === 2 ? `${parts[1]} ${parts[0]}` : trimmed;
                }).filter(Boolean);
                if (orgs.length === 0) return null;
                // v26.89: Gibt es einen expliziten Ansprechpartner, blenden wir den
                // „Bei Fragen wende dich gerne an:"-Kopf im Organizer-Hover aus —
                // für Rückfragen ist dann ausdrücklich der Ansprechpartner zuständig.
                const hasExplicitContact = !!(event.contactName || event.contactEmail || event.contactInfo);
                return (
                  <div style={{ marginTop: 6 }}>
                    <div style={{ fontSize: '0.85rem', color: 'var(--dex-gray-600)', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 6, fontWeight: 600 }}>Organizer</div>
                    <OrganizerList names={orgs} emails={event.organizerEmails} hiddenEmails={(event.hideOrganizer && event.hideOrganizerIndividualOnly) ? event.hiddenOrganizerEmails : []} forceIsDe={locale === 'de'} size="md" display={event.organizerDisplayLarge ? 'card' : 'chip'} nameFontSize="1.05rem" hideContactPrompt={hasExplicitContact} fullWidth contactEmail={event.contactOrganizerEmail || undefined} />
                  </div>
                );
              })()}
              {/* v23.25: Die „X / Y Plätze frei"-Anzeige steht jetzt direkt
                  über dem Registrieren-Button (siehe registration-actions). */}
            </div>
          </div>
          {/* v11.91: Beschreibung immer ausgeklappt — kein Toggle mehr. */}
          {event.description && (
            // v9.25: Beschreibung darf HTML enthalten (RichText-Editor im
            // EventCreation/Edit). Wir rendern als HTML statt Plain-Text,
            // damit Formatierung wie Listen, Links, Fett etc. funktioniert.
            // Die Description kommt aus dem eigenen Tenant — sicherer Origin.
            <div
              className="dex-event-desc"
              style={{
                padding: '12px 16px', color: 'var(--dex-gray-700)',
                background: 'var(--dex-gray-50)', borderRadius: '0 0 var(--dex-radius) var(--dex-radius)',
                borderTop: '1px solid var(--dex-gray-200)',
                wordBreak: 'break-word',
              }}
              dangerouslySetInnerHTML={{
                __html: (() => {
                  // v11.91: Email-Adressen in der Beschreibung automatisch
                  // in mailto-Links umwandeln. Funktioniert sowohl für
                  // Plain-Text als auch für HTML — Emails in bereits
                  // verlinktem Text (innerhalb von href="...") werden
                  // übersprungen.
                  // v29.59 BUG-FIX: Die Beschreibung wurde ROH gerendert — die
                  // Variablen, die der Editor anbietet ({{EventTitle}},
                  // {{Organizer}}, {{Name}}, {{AppUrl}}, {{ContactEmail}}),
                  // standen auf der Anmeldeseite als Text da. In der
                  // Live-Vorschau des Editors waren sie ersetzt, also sah der
                  // Organizer beim Schreiben nie, was Teilnehmer bekommen.
                  //
                  // {{WaitlistPosition}} bleibt bewusst aussen vor: Vor der
                  // Anmeldung gibt es keine Position, und eine erfundene Zahl
                  // waere schlimmer als der sichtbare Platzhalter. Der Editor
                  // bietet die Variable fuer die MAIL-Texte an; in der
                  // Beschreibung ergibt sie keinen Sinn.
                  const raw = ((): string => {
                    const src = event.description || '';
                    if (src.indexOf('{{') < 0) return src;
                    const orgNames = (event.organizers || [])
                      .reduce<string[]>((acc, o) => [...acc, ...o.split(';')], [])
                      .map(o => o.trim()).filter(Boolean);
                    const orgHidden = !!event.hideOrganizer && !event.hideOrganizerIndividualOnly;
                    return replacePlaceholders(src, {
                      EventTitle: event.title || '',
                      // Ausgeblendete Organizer duerfen ueber die Beschreibung
                      // nicht doch wieder sichtbar werden (v29.48-Logik).
                      Organizer: orgHidden ? '' : formatOrganizerList(orgNames, locale === 'de' ? 'DE' : 'EN'),
                      Name: `${currentUser.firstName || ''} ${currentUser.surname || ''}`.trim(),
                      AppUrl: `${window.location.origin}${window.location.pathname}`,
                      ContactEmail: event.contactEmail || '',
                    });
                  })();
                  const isHtml = /<[a-z][\s\S]*>/i.test(raw);
                  const base = isHtml
                    ? raw
                    : raw.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
                  const EMAIL_RE = /\b([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})\b/g;
                  // Skip emails that already sit inside href="..." (already linked)
                  return base.replace(EMAIL_RE, (match, _email, offset, full) => {
                    const beforeWindow = full.slice(Math.max(0, offset - 80), offset);
                    if (/href\s*=\s*["'][^"']*$/.test(beforeWindow)) return match;
                    if (/>$/.test(beforeWindow) && /<a [^>]*$/i.test(beforeWindow)) return match;
                    // v11.95: Deloitte-Grün (--dex-green #86bc25) statt dem
                    // dunkleren Olive — gleiche Farbe wie die Card-Header.
                    return `<a href="mailto:${match}" style="color:#86bc25;text-decoration:underline">${match}</a>`;
                  });
                })(),
              }}
            />
          )}
          {/* v24.59: Der frühere rote „Alle Plätze belegt …"-Text unter der
              Event-Karte ist entfernt — die Info steht jetzt im Badge über den
              Buttons (Status „Alle Plätze belegt") und in der Button-Beschriftung
              (Warteliste + aktuelle Anzahl). */}
          {/* v30.66: Der hier bis dahin stehende, seit v10.20 per {false && …}
              abgeschaltete Auswahl-Block (Hauptevent-/Sessions-Checkboxen) ist entfernt —
              die aktive Auswahl steht in der rechten Spalte (registration-specific). */}
        </div>
  );
};
