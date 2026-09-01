/* ActiveEventHintsBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 9886-10275 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { looksEnglishText, stripHtmlToText } from '../../../utils/eventStatus';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { isEventOver } from '../../../utils/eventFormat';
import { ChevronDown, ChevronUp, Info, QrCode } from '../../Icons';
import { DeloitteEvent } from '../../../types';

export interface ActiveEventHintsBoxProps {
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
  expandedHintIds: Set<string>;
  hintLangBusy: boolean;
  hintsDismissTick: number;
  isDe: boolean;
  parentEventForSelected: DeloitteEvent;
  refreshEvents: () => Promise<void>;
  selectedEvent: DeloitteEvent;
  setExpandedHintIds: React.Dispatch<React.SetStateAction<Set<string>>>;
  setHintLangBusy: React.Dispatch<React.SetStateAction<boolean>>;
  setHintsDismissTick: React.Dispatch<React.SetStateAction<number>>;
  setQrSendModalOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setSelectedEvent: React.Dispatch<React.SetStateAction<DeloitteEvent>>;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
  updateEvent: (eventId: string, updates: Record<string, unknown>, opts?: { skipReload?: boolean; }) => Promise<boolean>;
}

export const ActiveEventHintsBox: React.FC<ActiveEventHintsBoxProps> = (p) => {
  const { childEventsOf, expandedHintIds, hintLangBusy, hintsDismissTick, isDe, parentEventForSelected, refreshEvents, selectedEvent, setExpandedHintIds, setHintLangBusy, setHintsDismissTick, setQrSendModalOpen, setSelectedEvent, showAlert, updateEvent } = p;
          void hintsDismissTick; // erzwingt Re-Render nach „Ausblenden"
          const dismissKey = (id: string): string => `dex_hint_dismiss_${selectedEvent.id}_${id}`;
          const isDismissed = (id: string): boolean => {
            try { return window.localStorage.getItem(dismissKey(id)) === '1'; } catch { return false; }
          };
          const hints: Array<{ id: string; title: string; body: React.ReactNode; action?: React.ReactNode }> = [];
          // 1) Englischer Inhalt, aber Anmeldesprache nicht fest auf Englisch.
          const fieldsText = (selectedEvent.eventSpecificFields || [])
            .map(f => [f.label, f.helpText, (f.options || []).join(' ')].filter(Boolean).join(' '))
            .join(' ');
          const contentText = `${stripHtmlToText(selectedEvent.description || '')} ${fieldsText}`;
          if ((selectedEvent.registrationLanguage || '') !== 'en' && looksEnglishText(contentText)) {
            hints.push({
              id: 'lang-en',
              title: isDe ? 'Anmeldesprache auf Englisch festlegen?' : 'Fix registration language to English?',
              body: isDe
                ? 'Beschreibung und Felder dieses Events sind offenbar auf Englisch — die Anmeldeseite folgt aber der App-Sprache des Teilnehmers. Bei deutscher App-Einstellung mischt das Formular dann Deutsch (Buttons, Hinweise, Datenschutz) und Englisch (Inhalte). Empfehlung: die Anmeldesprache fest auf Englisch stellen. (Auch im Wizard änderbar: Schritt 5 „Felder" → „Sprache des Anmeldeformulars".)'
                : 'The description and fields of this event appear to be in English — but the registration page follows each participant\'s app language. With a German app setting the form then mixes German (buttons, hints, privacy note) and English (content). Recommendation: fix the registration language to English. (Also changeable in the wizard: step 5 “Fields” → “Registration form language”.)',
              action: (
                <button
                  type="button"
                  className="btn btn-primary"
                  disabled={hintLangBusy}
                  style={{ fontSize: '0.78rem', padding: '6px 12px' }}
                  onClick={() => {
                    (async () => {
                      setHintLangBusy(true);
                      const ok = await updateEvent(selectedEvent.id, { 'RegistrationLanguage': 'en' });
                      setHintLangBusy(false);
                      if (ok) {
                        setSelectedEvent(prev => prev ? { ...prev, registrationLanguage: 'en' } : prev);
                        await refreshEvents();
                        showAlert(isDe
                          ? 'Anmeldesprache auf Englisch festgelegt — das Anmeldeformular erscheint jetzt für alle Teilnehmer durchgängig auf Englisch.'
                          : 'Registration language fixed to English — the registration form now appears consistently in English for everyone.', { variant: 'success' });
                      } else {
                        showAlert(isDe ? 'Anmeldesprache konnte nicht gespeichert werden.' : 'Could not save the registration language.', { variant: 'error' });
                      }
                    })().catch(() => { /* */ });
                  }}
                >
                  {hintLangBusy ? (isDe ? 'Speichert…' : 'Saving…') : (isDe ? 'Auf Englisch festlegen' : 'Fix to English')}
                </button>
              ),
            });
          }
          // 2) Beschreibung fehlt oder ist sehr kurz.
          if (stripHtmlToText(selectedEvent.description || '').length < 20) {
            hints.push({
              id: 'no-desc',
              title: isDe ? 'Beschreibung ergänzen' : 'Add a description',
              body: isDe
                ? 'Das Event hat (fast) keine Beschreibung — Teilnehmer sehen auf der Anmeldeseite dann kaum, worum es geht. Über „Event bearbeiten" → Schritt 1 (Grundlagen) ergänzen.'
                : 'The event has (almost) no description — participants see very little about it on the registration page. Add one via “Edit event” → step 1 (Basics).',
            });
          }
          // 3) Event-Bild fehlt. v23.6: Bei einem Sub-Event NICHT meckern, wenn
          // die Klammer/das Hauptevent bereits ein Bild hat — Sub-Events nutzen
          // den Bild-/Hero-Kontext des Parents, ein eigenes Bild ist optional.
          if (!selectedEvent.imageUrl && !(parentEventForSelected && parentEventForSelected.imageUrl)) {
            hints.push({
              id: 'no-image',
              title: isDe ? 'Event-Bild hochladen' : 'Upload an event image',
              body: isDe
                ? 'Ohne Bild wirkt die Event-Karte in der Übersicht und der Mail-Kopf deutlich weniger einladend. Über „Event bearbeiten" → Schritt 1 (Grundlagen) hochladen.'
                : 'Without an image the event card in the list and the email header look much less inviting. Upload one via “Edit event” → step 1 (Basics).',
            });
          }
          // v23.21: Ansprechpartner-Freitext wiederholt Titel/Datum/Ort —
          // gleicher Hinweis wie im Wizard, hier auch im Organizer Center, damit
          // der Organizer es ohne Öffnen des Wizards sieht.
          {
            const ci = (selectedEvent.contactInfo || '').replace(/\s+/g, ' ').toLowerCase();
            if (ci.trim().length >= 4) {
              const tl = (selectedEvent.title || '').trim().toLowerCase();
              const locl = (selectedEvent.location || '').trim().toLowerCase();
              let redundant = (tl.length >= 4 && ci.indexOf(tl) >= 0) || (locl.length >= 4 && ci.indexOf(locl) >= 0);
              if (!redundant && selectedEvent.startDate) {
                const d = new Date(selectedEvent.startDate);
                if (!isNaN(d.getTime())) {
                  const dd = String(d.getDate()).padStart(2, '0');
                  const mm = String(d.getMonth() + 1).padStart(2, '0');
                  if ([`${dd}.${mm}.${d.getFullYear()}`, `${dd}.${mm}.`].some(p => ci.indexOf(p) >= 0)) redundant = true;
                }
              }
              if (redundant) {
                hints.push({
                  id: 'contact-redundant',
                  title: isDe ? 'Ansprechpartner-Text kürzen' : 'Shorten the contact text',
                  body: isDe
                    ? 'Beim Ansprechpartner steht offenbar der Event-Titel, das Datum oder der Ort — die werden bereits separat auf der Anmeldeseite angezeigt. Das Feld ist nur für die Erreichbarkeit gedacht. Über „Event bearbeiten" → Schritt 1 (Grundlagen) kürzen.'
                    : 'The contact text appears to repeat the event title, date or location — these are already shown separately on the registration page. The field is only for availability. Shorten it via “Edit event” → step 1 (Basics).',
                });
              }
            }
          }
          // 3b) v22.63: Der frühere allgemeine „Sichtbarkeit gilt fürs ganze
          // Event"-Hinweis ist entfallen — er erschien immer und war reines
          // Erklär-Rauschen. Hinweise kommen jetzt nur noch bei echten
          // möglichen Inkonsistenzen (3c/3d/3e).
          // 3c–e) v24.26: Sichtbarkeits-Hinweise KONSOLIDIERT. Früher liefen drei
          // sehr ähnliche Einzelkarten (kleine Zielgruppe / Sub-Event mit eigener
          // kleiner Zielgruppe / Sub-Event weiter geöffnet als das Event). Jetzt
          // eine gemeinsame „Sichtbarkeit prüfen"-Karte, die nur die tatsächlich
          // zutreffenden Punkte als Liste zeigt — ein Hinweis, ein Ausblenden.
          {
            const aud = selectedEvent.audienceFilter || [];
            const loc = selectedEvent.locationAudience || [];
            const hasAllPattern = aud.some(a => { const f = (a || '').toLowerCase(); return f === 'all' || f === 'deall'; });
            const resolved = (selectedEvent.audienceResolvedEmails || []).map(s => (s || '').trim()).filter(Boolean);
            const atCount = aud.filter(a => a.indexOf('@') >= 0).length;
            const effCount = resolved.length > 0 ? resolved.length : atCount;
            // (1) Hauptevent sehr kleine Zielgruppe
            const mainTiny = aud.length > 0 && loc.length === 0 && !hasAllPattern && effCount > 0 && effCount < 10;

            const parentAudKey = aud.join('|');
            const parentLocKey = loc.join('|');
            const parentShowsAll = (aud.length === 0 && loc.length === 0) || hasAllPattern;
            // (2) Sub-Events mit eigener sehr kleiner Sichtbarkeit
            // (3) Sub-Events weiter/anders geöffnet als die Klammer
            const smallSubs: string[] = [];
            const riskySubs: string[] = [];
            for (const ch of childEventsOf(selectedEvent.id)) {
              const cAud = ch.audienceFilter || [];
              const cLoc = ch.locationAudience || [];
              if (cAud.length === 0 && cLoc.length === 0) continue; // erbt die Klammer
              if (cAud.join('|') === parentAudKey && cLoc.join('|') === parentLocKey) continue; // identisch zur Klammer
              const cHasAll = cAud.some(a => { const f = (a || '').toLowerCase(); return f === 'all' || f === 'deall'; });
              const cResolved = (ch.audienceResolvedEmails || []).map(s => (s || '').trim()).filter(Boolean);
              const cAt = cAud.filter(a => a.indexOf('@') >= 0).length;
              const cEff = cResolved.length > 0 ? cResolved.length : cAt;
              if (cAud.length > 0 && cLoc.length === 0 && !cHasAll && cEff > 0 && cEff < 10) {
                smallSubs.push(`${shortSubEventTitle(ch.title, selectedEvent.title)} (${cEff})`);
              }
              if (!parentShowsAll) {
                riskySubs.push(shortSubEventTitle(ch.title, selectedEvent.title));
              }
            }

            if (mainTiny || smallSubs.length > 0 || riskySubs.length > 0) {
              hints.push({
                id: 'visibility-check',
                title: isDe ? 'Sichtbarkeit — bitte prüfen' : 'Visibility — please check',
                body: isDe ? (
                  <>
                    Schau dir kurz die Zielgruppe an — folgendes ist aufgefallen:
                    <ul style={{ margin: '6px 0 0', paddingLeft: 18, lineHeight: 1.5 }}>
                      {mainTiny && <li>Das <strong>Event</strong> ist nur für <strong>{effCount} {effCount === 1 ? 'Person' : 'Personen'}</strong> sichtbar (einzelne Adressen — kein Standort/Verteiler, nicht „alle Mitarbeiter“). Falls du mehr Leute erreichen willst, ergänze einen Standort oder Verteiler.</li>}
                      {smallSubs.length > 0 && <li>Diese <strong>Sub-Events</strong> haben eine eigene, sehr kleine Zielgruppe — wer dort nicht gelistet ist, kann sich nicht anmelden:
                        <ul style={{ margin: '3px 0 0', paddingLeft: 16 }}>{smallSubs.map(s => <li key={s}>{s}</li>)}</ul>
                      </li>}
                      {riskySubs.length > 0 && <li>Diese <strong>Sub-Events</strong> sind für andere/mehr Leute geöffnet als das Event — wer nur beim Sub-Event steht, aber nicht beim Event, <strong>kann das Event nicht öffnen</strong> und sieht es nie (beim Event sollten alle dabei sein, die irgendein Sub-Event sehen sollen):
                        <ul style={{ margin: '3px 0 0', paddingLeft: 16 }}>{riskySubs.map(s => <li key={s}>{s}</li>)}</ul>
                      </li>}
                    </ul>
                    <div style={{ marginTop: 6 }}>Prüfen/anpassen über „Event bearbeiten“ → Schritt 4 (ggf. Sub-Event-Tab) → „Sichtbarkeit prüfen“.</div>
                  </>
                ) : (
                  <>
                    Take a quick look at the audience — the app noticed:
                    <ul style={{ margin: '6px 0 0', paddingLeft: 18, lineHeight: 1.5 }}>
                      {mainTiny && <li>The <strong>event</strong> is visible to only <strong>{effCount} {effCount === 1 ? 'person' : 'people'}</strong> (individual addresses — no location/list, not “all employees”). If you want to reach more people, add a location or distribution list.</li>}
                      {smallSubs.length > 0 && <li>These <strong>sub-events</strong> have their own, very small audience — anyone not listed there cannot register:
                        <ul style={{ margin: '3px 0 0', paddingLeft: 16 }}>{smallSubs.map(s => <li key={s}>{s}</li>)}</ul>
                      </li>}
                      {riskySubs.length > 0 && <li>These <strong>sub-events</strong> are open to other/more people than the event — anyone listed only on the sub-event but not on the event <strong>cannot open the event</strong> and never sees it (the event should include everyone who should see any sub-event):
                        <ul style={{ margin: '3px 0 0', paddingLeft: 16 }}>{riskySubs.map(s => <li key={s}>{s}</li>)}</ul>
                      </li>}
                    </ul>
                    <div style={{ marginTop: 6 }}>Check/adjust via “Edit event” → step 4 (sub-event tab if needed) → “Check visibility”.</div>
                  </>
                ),
              });
            }
          }
          // 3f) v24.27/v24.28: Feld-Tipps — passendere Feldart (Datum/Person)
          // ODER Feld, das ohnehin schon automatisch aus dem Profil erfasst wird
          // (z.B. Abteilung, Standort, Firma) und deshalb überflüssig sein kann.
          // Zugang zum Ändern läuft über „Event bearbeiten" → Schritt „Felder".
          {
            const looksDate = (s: string): boolean => /(datum|date|check[\s-]?in|check[\s-]?out|anreise|abreise|geburtstag|birthday|deadline|frist|termin|ankunft|abfahrt|arrival|departure)/i.test(s || '');
            const looksName = (s: string): boolean => /(\bname\b|vorname|nachname|ansprechpartner|counselor|kolleg|mitarbeiter|\bmentor\b|\bpate\b|\bbuddy\b|begleitung|\bgast\b)/i.test(s || '');
            // Felder, die i.d.R. schon automatisch aus dem Deloitte-Profil
            // kommen (Anrede/Vorname/Nachname/E-Mail/Abteilung/Standort/Position/
            // „name" allein bewusst NICHT — das ist zu mehrdeutig (z.B. „Name of
            // counselor"). v26.83: Telefon/Mobil/Handy UND Adresse RAUS — die
            // werden NICHT automatisch aus dem Profil erfasst (Fehlalarm: der
            // Hinweis empfahl fälschlich, eine Mobilnummer-Abfrage wegzulassen).
            const looksProfile = (s: string): boolean => /(vorname|nachname|first ?name|last ?name|e-?mail|abteilung|department|standort|location|\boffice\b|\bbüro\b|firma|company|unternehmen|arbeitgeber|gesellschaft|\bgmbh\b|legal ?entity|\bentity\b|rechtsträger|member ?firm|job ?title)/i.test(s || '');
            // Eindeutige Feld-Namen sammeln (gleiches Label in Haupt- + mehreren
            // Sub-Events nur EINMAL nennen). Profil-Felder haben Vorrang (sie
            // sollen ganz entfallen, nicht nur die Feldart wechseln).
            const profileSet = new Map<string, string>();
            const dateSet = new Map<string, string>();
            const nameSet = new Map<string, string>();
            const scan = (fields: { type?: string; label?: string }[] | undefined): void => {
              for (const f of (fields || [])) {
                const lbl = (f.label || '').trim();
                if (!lbl) continue;
                const key = lbl.toLowerCase();
                if (looksProfile(lbl)) { if (!profileSet.has(key)) profileSet.set(key, lbl); }
                else if ((f.type === 'text' || f.type === 'number') && looksDate(lbl)) { if (!dateSet.has(key)) dateSet.set(key, lbl); }
                else if (f.type === 'text' && looksName(lbl)) { if (!nameSet.has(key)) nameSet.set(key, lbl); }
              }
            };
            scan(selectedEvent.eventSpecificFields);
            for (const ch of childEventsOf(selectedEvent.id)) scan(ch.eventSpecificFields);
            const dateLike = Array.from(dateSet.values());
            const nameLike = Array.from(nameSet.values());
            const profileLike = Array.from(profileSet.values());
            if (dateLike.length > 0 || nameLike.length > 0 || profileLike.length > 0) {
              // v24.56: Feldnamen als Badges/Chips darstellen, damit klar ist,
              // welche Felder gemeint sind. Format pro Tipp: „Du hast …" +
              // Empfehlung + Erklärung.
              const badges = (arr: string[]): React.ReactNode => (
                <span style={{ display: 'inline-flex', flexWrap: 'wrap', gap: 6, verticalAlign: 'middle' }}>
                  {arr.map((f, i) => (
                    <span key={`${f}-${i}`} style={{ display: 'inline-block', background: 'rgba(237,139,0,0.14)', border: '1px solid var(--dex-orange, #ed8b00)', color: 'var(--dex-orange-dark, #b35a00)', borderRadius: 999, padding: '1px 9px', fontSize: '0.74rem', fontWeight: 700, whiteSpace: 'nowrap' }}>{f}</span>
                  ))}
                </span>
              );
              const tip = (intro: React.ReactNode, empf: string, erkl: string): React.ReactElement => (
                <div style={{ marginTop: 4 }}>
                  <div>{intro}</div>
                  <div style={{ marginTop: 3 }}><strong style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>{isDe ? 'Empfehlung:' : 'Recommendation:'}</strong> {empf}</div>
                  <div style={{ marginTop: 2, color: 'var(--dex-gray-600)' }}><strong>{isDe ? 'Erklärung:' : 'Why:'}</strong> {erkl}</div>
                </div>
              );
              hints.push({
                id: 'fieldtype-suggestion',
                title: isDe ? 'Tipps zu deinen Feldern' : 'Tips for your fields',
                body: (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                    {dateLike.length > 0 && tip(
                      isDe ? <>Du hast {dateLike.length === 1 ? 'das Feld' : 'die Felder'} {badges(dateLike)} als <strong>Freitext</strong> eingestellt.</> : <>You set {dateLike.length === 1 ? 'the field' : 'the fields'} {badges(dateLike)} as <strong>free text</strong>.</>,
                      isDe ? 'Nutze besser den Feldtyp „Datum".' : 'Better use the „Date" field type.',
                      isDe ? 'Teilnehmer wählen dann ein Datum im Kalender (optional mit Uhrzeit) — keine Tippfehler, einheitliches Format.' : 'Attendees then pick a date from a calendar (optionally with time) — no typos, consistent format.'
                    )}
                    {nameLike.length > 0 && tip(
                      isDe ? <>Du hast {nameLike.length === 1 ? 'das Feld' : 'die Felder'} {badges(nameLike)} als <strong>Freitext</strong> eingestellt.</> : <>You set {nameLike.length === 1 ? 'the field' : 'the fields'} {badges(nameLike)} as <strong>free text</strong>.</>,
                      isDe ? 'Nutze besser den Feldtyp „Person".' : 'Better use the „Person" field type.',
                      isDe ? 'Teilnehmer suchen die Person direkt (mit Foto und Standort) — eindeutig statt frei getippt.' : 'Attendees search the person directly (with photo and location) — unambiguous instead of free text.'
                    )}
                    {profileLike.length > 0 && tip(
                      isDe ? <>Du fragst {profileLike.length === 1 ? 'das Feld' : 'die Felder'} {badges(profileLike)} ab.</> : <>You ask for {profileLike.length === 1 ? 'the field' : 'the fields'} {badges(profileLike)}.</>,
                      isDe ? 'Diese Felder kannst du weglassen.' : 'You can drop these fields.',
                      isDe ? 'Angaben wie Abteilung, Standort oder Firma kommen automatisch aus dem Profil — du musst sie nicht extra abfragen.' : 'Details like department, location or company come automatically from the profile — no need to ask for them.'
                    )}
                    <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.76rem' }}>
                      {isDe ? 'Anpassen über „Event bearbeiten" → Schritt „Felder". Bestehende Antworten bleiben erhalten, werden aber nicht automatisch ins neue Format umgewandelt.' : 'Adjust via „Edit event" → step „Fields". Existing answers are kept but not automatically converted.'}
                    </div>
                  </div>
                ),
              });
            }
          }
          // 4) v22.34: End-Datum fehlt (Hauptevent oder Sub-Event) — ohne Ende
          // kann der Outlook-Termin nicht angelegt werden (der Kalendereintrag
          // braucht Start UND Ende; das Sub-Event bekommt dann nie eine
          // OutlookEventId). Praxisfall: Organizerin vergaß beim Anlegen das
          // End-Datum eines Sub-Events → kein Outlook-Termin für die Teilnehmer.
          {
            const noEndChildren = childEventsOf(selectedEvent.id)
              .filter(c => !!(c.startDate || '').trim() && !(c.endDate || '').trim());
            const mainNoEnd = !!(selectedEvent.startDate || '').trim() && !(selectedEvent.endDate || '').trim();
            if (mainNoEnd || noEndChildren.length > 0) {
              const names: string[] = [];
              if (mainNoEnd) names.push(isDe ? 'das Hauptevent' : 'the main event');
              for (const c of noEndChildren) {
                names.push(`„${shortSubEventTitle(c.title, selectedEvent.title) || c.title}"`);
              }
              hints.push({
                id: 'no-enddate',
                title: isDe
                  ? 'End-Datum fehlt — Outlook-Termin kann nicht erstellt werden'
                  : 'End date missing — Outlook invite cannot be created',
                body: isDe
                  ? <>Ohne End-Datum kann für die Teilnehmer <strong>kein Outlook-Termin</strong> angelegt werden (ein Kalendereintrag braucht Start UND Ende) — betroffen: <strong>{names.join(', ')}</strong>. Bitte über „Event bearbeiten“ das End-Datum nachtragen (Hauptevent: Schritt 1 „Grundlagen“, Sub-Events: Schritt 2 „Sub-Events“). Beim Speichern fragt die App dann, ob der Outlook-Termin angelegt bzw. aktualisiert werden soll.</>
                  : <>Without an end date <strong>no Outlook invite</strong> can be created for attendees (a calendar entry needs a start AND an end) — affected: <strong>{names.join(', ')}</strong>. Please add the end date via “Edit event” (main event: step 1 “Basics”, sub-events: step 2 “Sub-events”). When saving, the app then asks whether the Outlook invite should be created or updated.</>,
              });
            }
          }
          // 5) v22.69: Hauptevent/Klammer ist live, aber ein Sub-Event steht
          // noch auf Entwurf — Entwurf-Sub-Events sind für reguläre Teilnehmer
          // NICHT buchbar (seit v22.68). Der Organizer denkt sonst, alles sei
          // buchbar.
          if (!selectedEvent.isFictive) {
            const draftKids = childEventsOf(selectedEvent.id).filter(c => c.isFictive);
            if (draftKids.length > 0) {
              const draftNames = draftKids.map(c => shortSubEventTitle(c.title, selectedEvent.title)).join(', ');
              hints.push({
                id: 'draft-subevent-live-parent',
                title: isDe ? 'Sub-Event noch im Entwurf — nicht buchbar' : 'Sub-event still a draft — not bookable',
                body: isDe
                  ? <>Das Event ist live, aber diese Sub-Events stehen noch auf <strong>Entwurf</strong>: <strong>{draftNames}</strong>. Entwurf-Sub-Events sind für reguläre Teilnehmer <strong>nicht sichtbar und nicht buchbar</strong>. Wenn sie buchbar sein sollen, schalte sie über den Status-Badge oben (Entwurf ⇄ Aktiv) auf den jeweiligen Sub-Event-Tab live.</>
                  : <>The event is live, but these sub-events are still in <strong>draft</strong>: <strong>{draftNames}</strong>. Draft sub-events are <strong>not visible and not bookable</strong> for regular attendees. If they should be bookable, publish them via the status badge (draft ⇄ active) on the respective sub-event tab.</>,
              });
            }
          }
          // 6) v24.83: Ab 5 Tagen vor Event-Start ein Hinweis GANZ OBEN, dass
          // jetzt die persönlichen Check-in-QR-Codes versendet werden können
          // (optional — „falls du möchtest"). Nur im Hauptevent-Detail (nicht
          // im Sub-Event-Detail) und nur solange das Event nicht vorbei ist.
          if (!parentEventForSelected) {
            const startMs = selectedEvent.startDate ? new Date(selectedEvent.startDate).getTime() : 0;
            const daysUntilStart = startMs ? (startMs - Date.now()) / 86400000 : Infinity;
            if (daysUntilStart <= 5 && !isEventOver(selectedEvent)) {
              hints.unshift({
                id: 'qr-send-window',
                title: isDe ? 'QR-Codes versenden möglich' : 'QR codes can be sent now',
                body: isDe ? (
                  <>
                    Das Event startet in den nächsten Tagen. Du kannst jetzt — wenn du möchtest — die persönlichen <strong>Check-in-QR-Codes</strong> an alle angemeldeten Teilnehmer verschicken. Jede Person bekommt ihren Code per E-Mail; wer sich danach noch anmeldet, erhält ihn automatisch. Am Veranstaltungstag scannst du die Codes am Eingang (oder die Teilnehmer checken sich per Self-Check-in selbst ein).
                    <div style={{ marginTop: 8 }}>
                      <button type="button" className="btn btn-primary" style={{ fontSize: '0.8rem', padding: '6px 14px', display: 'inline-flex', alignItems: 'center', gap: 6 }} onClick={() => setQrSendModalOpen(true)}>
                        <QrCode size={14} /> QR-Codes versenden
                      </button>
                    </div>
                  </>
                ) : (
                  <>
                    The event starts within the next few days. You can now — if you like — send the personal <strong>check-in QR codes</strong> to all registered attendees. Each person gets their code by email; anyone registering afterwards receives it automatically. On the event day you scan the codes at the entrance (or attendees self-check-in).
                    <div style={{ marginTop: 8 }}>
                      <button type="button" className="btn btn-primary" style={{ fontSize: '0.8rem', padding: '6px 14px', display: 'inline-flex', alignItems: 'center', gap: 6 }} onClick={() => setQrSendModalOpen(true)}>
                        <QrCode size={14} /> Send QR codes
                      </button>
                    </div>
                  </>
                ),
              });
            }
          }
          const visible = hints.filter(h => !isDismissed(h.id));
          if (visible.length === 0) return null;
          return (
            // v26.77: Hinweise-Box jetzt in VOLLER BREITE direkt über der
            // Teilnehmerliste (vorher schmale rechte Spalte) — so fällt sie
            // deutlicher auf. Text durchgängig linksbündig.
            <div style={{ marginBottom: 24 }}>
              <div className="card" style={{ padding: 20, background: 'rgba(237,139,0,0.06)', border: '1px solid var(--dex-orange, #ed8b00)', textAlign: 'left' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
                  <span style={{ color: 'var(--dex-orange, #ed8b00)', display: 'inline-flex' }}><Info size={18} /></span>
                  <h3 style={{ margin: 0, fontSize: '1rem', color: 'var(--dex-orange-dark, #b35a00)' }}>{isDe ? 'Hinweise zu diesem Event' : 'Hints for this event'}</h3>
                </div>
                <p style={{ margin: '0 0 14px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                  {isDe ? 'Der App sind ein paar Dinge aufgefallen, die du dir kurz anschauen solltest (zum Aufklappen auf die Überschrift tippen):' : 'The app noticed a few things worth a quick look (tap a heading to expand):'}
                </p>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                  {visible.map(h => {
                    const open = expandedHintIds.has(h.id);
                    return (
                    <div key={h.id} style={{ borderTop: '1px solid rgba(237,139,0,0.2)', paddingTop: 10 }}>
                      {/* v24.50: Überschrift = Klappschalter (Default eingeklappt). */}
                      <button
                        type="button"
                        onClick={() => setExpandedHintIds(prev => { const n = new Set(prev); if (n.has(h.id)) n.delete(h.id); else n.add(h.id); return n; })}
                        style={{ width: '100%', textAlign: 'left', background: 'none', border: 'none', padding: 0, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 8 }}
                      >
                        <span style={{ color: 'var(--dex-orange, #ed8b00)', display: 'inline-flex', flexShrink: 0 }}>
                          {open ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
                        </span>
                        <span style={{ fontWeight: 700, fontSize: '0.85rem', color: 'var(--dex-gray-800)' }}>{h.title}</span>
                      </button>
                      {open && (
                        <div style={{ marginTop: 6, paddingLeft: 22 }}>
                          <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>{h.body}</div>
                          <div style={{ marginTop: 8, display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                            {h.action}
                            <button
                              type="button"
                              onClick={() => {
                                try { window.localStorage.setItem(dismissKey(h.id), '1'); } catch { /* */ }
                                setHintsDismissTick(t => t + 1);
                              }}
                              style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', color: 'var(--dex-gray-500)', fontSize: '0.74rem', textDecoration: 'underline' }}
                            >
                              {isDe ? 'Hinweis ausblenden' : 'Dismiss hint'}
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                    );
                  })}
                </div>
              </div>
            </div>
          );
};

