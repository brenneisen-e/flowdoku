/**
 * v30.52: Das Bild im Mail-Kopf — EIN Zustand, EINE Auswahl, EINE Umrechnung.
 *
 * Vorher lag dieselbe Sache dreimal unterschiedlich in `AdminPage`:
 *
 *  - **Massenmail** hatte `massmailHero` + `massmailImageLayout`, eine eigene
 *    Reiter-Reihe „DEX-Logo / Event-Foto", einen eigenen Zuschneide-Knopf und
 *    `applyMassmailHero`.
 *  - **Einladungsmail** hatte dasselbe noch einmal, mit `invite`-Präfix und
 *    einer zweiten, wortgleichen Kopie der Reiter-Reihe.
 *  - **QR-Mail** hatte **gar nichts** — kein Bildwechsel, keine Breite, keine
 *    „Volle Breite". Genau das war der gemeldete Fall: derselbe Dialog-Typ,
 *    aber in der QR-Mail fehlten die Knöpfe.
 *
 * Zwei Kopien laufen irgendwann auseinander, drei Verhaltensweisen für
 * dieselbe Frage sind schon auseinandergelaufen. Deshalb liegt hier alles
 * beisammen: der Zustandstyp, die Umrechnung in die Kopf-Maße, das Einsetzen
 * des Bildes ins fertige HTML und die Bedienung.
 *
 * Die Bedienung (Auswahl-Reiter) steht in
 * `components/admin/MailHeaderImageChooser.tsx` — hier liegt nur die reine
 * Logik, damit auch `services/EmailTemplates` sie nutzen kann, ohne React
 * mitzuziehen.
 *
 * **Was NICHT hierher gehört:** Die Felder für Breite und Innenabstand stehen
 * weiterhin in `HtmlEditorModal` („HEADER-BILD"). Sie waren nie doppelt — sie
 * wurden von der QR-Mail nur nicht angefordert. Dort stehen sie direkt neben
 * der Live-Vorschau, die sie sofort zeigt; sie hier hochzuziehen würde die
 * Einstellung von ihrer Wirkung trennen. Beide Stellen arbeiten jetzt auf
 * demselben `MailHeaderImage`-Objekt.
 */

export interface MailHeaderImage {
  /** `logo` = Standard (DEX-Orb bzw. das Mail-Logo des Events, vom Flow
   *  eingesetzt), `event` = das Event-Foto fest eingebacken. */
  hero: 'logo' | 'event';
  width: number;
  paddingV: number;
  paddingH: number;
}

export const MAIL_HEADER_IMAGE_DEFAULT: MailHeaderImage = {
  hero: 'logo', width: 180, paddingV: 30, paddingH: 30,
};

/**
 * Orb-Schutz (aus `headerOptsFor`, v29.37; im Wizard `headerLayoutFor`, v28.29).
 *
 * Ohne eigenes Bild setzt der Flow das Standard-DEX-Logo in den Kopf — 600 px
 * breit wäre das ein bildschirmfüllender, unten abgeschnittener Orb. Deshalb
 * gilt „Volle Breite" nur, wenn wirklich ein eigenes Bild im Kopf steht.
 */
export function mailHeaderOpts(
  img: MailHeaderImage,
  hasOwnImage: boolean
): { imageWidth: number; imagePaddingV: number; imagePaddingH: number } {
  return {
    imageWidth: hasOwnImage ? img.width : Math.min(img.width, 180),
    imagePaddingV: hasOwnImage ? img.paddingV : Math.max(img.paddingV, 20),
    imagePaddingH: hasOwnImage ? img.paddingH : Math.max(img.paddingH, 20),
  };
}

/**
 * Steht im Kopf ein EIGENES Bild? Entweder das eingebackene Event-Foto oder —
 * wenn `{{ORB_URL}}` stehen bleibt — das Mail-Logo des Events, das der Flow
 * einsetzt. Nur dann darf die volle Breite gelten.
 */
export function hasOwnHeaderImage(
  img: MailHeaderImage,
  eventPhotoB64: string,
  eventMailLogo: string | undefined | null
): boolean {
  return (img.hero === 'event' && !!eventPhotoB64) || !!eventMailLogo;
}

/**
 * Ersetzt `{{ORB_URL}}` im gewickelten Mail-HTML durch das gewählte Kopf-Bild.
 *
 * Bei „Event-Foto" (und geladenem Foto) wird das Bild als Base64 fest
 * eingebacken; sonst bleibt der Platzhalter stehen, damit der Flow wie gehabt
 * das Standard-Bild einsetzt (DEX-Logo bzw. das konfigurierte Mail-Logo des
 * Events). Ersetzt `applyMassmailHero` + `applyInviteHero`, die Zeile für
 * Zeile identisch waren.
 */
export function applyHeroImage(
  wrappedHtml: string,
  img: MailHeaderImage,
  eventPhotoB64: string
): string {
  return (img.hero === 'event' && eventPhotoB64)
    ? wrappedHtml.replace(/\{\{ORB_URL\}\}/g, eventPhotoB64)
    : wrappedHtml;
}

/**
 * Liest einen gespeicherten Stand (QR-Mail-Override) defensiv ein.
 *
 * Alles, was nicht plausibel ist, fällt auf den Standard zurück: Der Wert
 * kommt aus einem JSON-Feld, das auch von Flows geschrieben wird — eine
 * fehlende Zahl darf hier keine `NaN`-Breite in die Mail tragen.
 */
export function normalizeMailHeaderImage(raw: unknown): MailHeaderImage {
  const o = (raw && typeof raw === 'object') ? raw as Record<string, unknown> : {};
  const num = (v: unknown, def: number, max: number): number => {
    const n = typeof v === 'number' ? v : parseInt(String(v ?? ''), 10);
    if (!isFinite(n) || n < 0) return def;
    return Math.min(n, max);
  };
  return {
    hero: o.hero === 'event' ? 'event' : 'logo',
    width: num(o.width, MAIL_HEADER_IMAGE_DEFAULT.width, 600),
    paddingV: num(o.paddingV, MAIL_HEADER_IMAGE_DEFAULT.paddingV, 80),
    paddingH: num(o.paddingH, MAIL_HEADER_IMAGE_DEFAULT.paddingH, 80),
  };
}

/**
 * v30.87: Kopf-Maße für eine Mail ZU EINEM EVENT — aus dessen Overrides.
 *
 * Befund (Nutzer, 07.09.2026, Organizer-Mail „Verspätete Abmeldung"): Das
 * Event-Bild stand klein und zentriert im Kopf, obwohl der Vollbild-Kopf seit
 * v29.29 der Standard ist. Grund: Rund zwanzig App-Mails (Team-, Zimmer-,
 * Hotel-, Abrechnungs-, Erinnerungs- und Organizer-Mails) riefen
 * `wrapTemplate` ohne Bildmaße auf — also mit dem alten 180-px-Default. Der
 * Vollbild-Kopf griff nur dort, wo der Wizard ihn beim Speichern
 * ausdrücklich in `_headerImageLayout` ablegte.
 *
 * Regel: Ein gespeichertes `_headerImageLayout` gewinnt (der Organizer hat es
 * eingestellt). Ohne gespeichertes Layout gilt der Vollbild-Kopf (600/0/0),
 * sobald das Event ein eigenes Mail-Logo hat (`_eventLogo` im Blob bzw. die
 * Spalte EmailImageBase64, vom Flow für {{ORB_URL}} eingesetzt). Ohne
 * eigenes Bild bleibt es beim kleinen Kopf — der DEX-Orb in 600 px wäre ein
 * bildschirmfüllender, abgeschnittener Kreis (Orb-Schutz, s. oben).
 */
export function eventHeaderImageOpts(
  overridesJson: string | undefined | null,
  mailLogoB64?: string | null
): { imageWidth: number; imagePaddingV: number; imagePaddingH: number } {
  let il: { width?: unknown; paddingV?: unknown; paddingH?: unknown } = {};
  let logo = '';
  try {
    const o = JSON.parse(overridesJson || '{}') || {};
    il = (o._headerImageLayout && typeof o._headerImageLayout === 'object') ? o._headerImageLayout : {};
    logo = typeof o._eventLogo === 'string' ? o._eventLogo : '';
  } catch { /* kein Blob → Defaults */ }
  const own = !!(mailLogoB64 || logo);
  const stored = typeof il.width === 'number' && il.width > 0;
  const base: MailHeaderImage = stored
    ? {
      hero: 'logo',
      width: il.width as number,
      paddingV: (typeof il.paddingV === 'number' && il.paddingV >= 0) ? il.paddingV : 30,
      paddingH: (typeof il.paddingH === 'number' && il.paddingH >= 0) ? il.paddingH : 30,
    }
    : (own ? { hero: 'logo', width: 600, paddingV: 0, paddingH: 0 } : { ...MAIL_HEADER_IMAGE_DEFAULT });
  return mailHeaderOpts(base, own);
}

export function isDefaultMailHeaderImage(img: MailHeaderImage): boolean {
  return img.hero === MAIL_HEADER_IMAGE_DEFAULT.hero
    && img.width === MAIL_HEADER_IMAGE_DEFAULT.width
    && img.paddingV === MAIL_HEADER_IMAGE_DEFAULT.paddingV
    && img.paddingH === MAIL_HEADER_IMAGE_DEFAULT.paddingH;
}

