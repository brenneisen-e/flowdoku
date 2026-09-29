/* Einstieg des Screenshot-Harness — mountet die ECHTE App ohne SharePoint.
 *
 * Nichts an der App ist ersetzt: `AiUseCasePlatform` läuft mit ihren echten
 * Providern (Sprache, Dialoge, Nutzer, Rollen, Navigation, Use Cases, Suche,
 * Hilfe). Ersetzt ist nur, womit sie spricht — der `WebPartContext`, dessen
 * `spHttpClient` ein SharePoint im Speicher ist (`fakeSharePoint.ts`).
 *
 * Adressparameter (alle optional):
 *   role=admin|organizer|user|first   Rolle der angemeldeten Person (first = Rollenliste leer → Erstinstallation)
 *   lang=de|en                        Sprache der Oberfläche
 *   mobile=1                          Handy-Zweige auch in einem breiten Fenster (sonst genügt ein schmales)
 *   state=ok|forbidden|error|empty|leer|roles403|logerror|fresh
 *   delay=<ms>                        künstliche Antwortzeit (Vorgabe 25)
 *   data=variety|start                Startdaten mit Status-/Bild-Abweichungen (Vorgabe) oder unverändert
 *   uc=<id>                           Deep-Link — liest die App selbst, das ist ein echter Test
 */
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import AiUseCasePlatform from '../../src/webparts/aiUseCasePlatform/components/AiUseCasePlatform';
import { baueKontext, parseParams } from './fakeSharePoint';

const params = parseParams(window.location.search);

/* Merker aus früheren Läufen würden die Bilder verfälschen (Sprache!). Der
 * Harness startet immer leer — nur die Sprache aus der Adresse bleibt gesetzt.
 * Der Schlüssel steht in context/LanguageContext.tsx. */
try { window.localStorage.clear(); window.sessionStorage.clear(); } catch { /* privates Fenster */ }
try { window.localStorage.setItem('aiuc_locale', params.lang); } catch { /* */ }

/* `useIsMobile` fragt `window.matchMedia('(max-width: 768px)')`. Mit einem
 * schmalen Fenster (Shot-Lauf: 390 × 844) stimmt das von selbst. `mobile=1`
 * erzwingt es zusätzlich für ein breites Fenster — nur für die Abfrage der App,
 * nicht für das CSS (dessen Media-Queries folgen weiter der echten Breite). */
if (params.mobile && window.innerWidth > 768) {
  const echt = window.matchMedia ? window.matchMedia.bind(window) : null;
  window.matchMedia = ((q: string): MediaQueryList => {
    if (/max-width:\s*768px/.test(q)) {
      return {
        matches: true, media: q, onchange: null,
        addListener: () => undefined, removeListener: () => undefined,
        addEventListener: () => undefined, removeEventListener: () => undefined,
        dispatchEvent: () => false,
      } as MediaQueryList;
    }
    return echt ? echt(q) : ({ matches: false, media: q, addListener() { /* */ }, removeListener() { /* */ }, addEventListener() { /* */ }, removeEventListener() { /* */ } } as any);
  }) as typeof window.matchMedia;
}

const context = baueKontext(params);
ReactDOM.render(<AiUseCasePlatform context={context} />, document.getElementById('root'));
