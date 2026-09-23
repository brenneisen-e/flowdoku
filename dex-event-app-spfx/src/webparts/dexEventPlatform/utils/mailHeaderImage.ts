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
   *  eingesetzt), `event` = das Event-Foto fest eingebacken, `custom` = ein
   *  für DIESE Mail hochgeladenes Bild (v31.9.7).
   *
   *  Das Bild selbst steht NICHT in diesem Objekt: Der Typ wird als JSON in
   *  SharePoint gespeichert (QR-Mail-Overrides), ein Base64-Bild würde die
   *  Spalte aufblähen. Es wird — genau wie das Event-Foto — als eigener
   *  Parameter durchgereicht. */
  hero: 'logo' | 'event' | 'custom';
  width: number;
  paddingV: number;
  paddingH: number;
}

export const MAIL_HEADER_IMAGE_DEFAULT: MailHeaderImage = {
  hero: 'logo', width: 180, paddingV: 30, paddingH: 30,
};

/**
 * v31.76: Kopf-Maße für ein FOTO (Event-Foto oder hochgeladenes Bild) nach
 * seiner Form — Nutzer-Regel 22.09.2026: „wenn es rund ist, dann 300 px,
 * sonst volle Breite."
 *
 * Anlass: In der QR-Mail stand der runde Event-Kreis bildschirmfüllend im
 * Kopf. Die Regel „volle Breite ohne Rand" (600/0/0) stammt aus v29.29/
 * v30.87 und meint das Mail-LOGO, ein Banner; beim Wechsel auf „Event-Foto"
 * wurden dieselben Maße übernommen, gemessen hat die Form niemand.
 *
 * Rund, quadratisch, hochkant (Breite höchstens 1,3 × Höhe) → 300 px, mit
 * Rand, mittig. Alles deutlich Breitere ist ein Banner → 600/0/0. Ohne
 * messbare Maße (0/0) gilt der sichere Fall: 300 px. Breite und Abstand
 * bleiben im Editor danach frei einstellbar; das hier ist der Startwert.
 */
export const KOPF_VOLLE_BREITE = { width: 600, paddingV: 0, paddingH: 0 };
export const KOPF_RUND = { width: 300, paddingV: 24, paddingH: 24 };
type KopfMasse = { width: number; paddingV: number; paddingH: number };

export function kopfMasseFuerBild(width: number, height: number): Pick<MailHeaderImage, 'width' | 'paddingV' | 'paddingH'> {
  const banner = width > 0 && height > 0 && width / height > 1.3;
  const m = banner ? { ...KOPF_VOLLE_BREITE } : { ...KOPF_RUND };
  // v31.80: nie über die eigene Bildbreite hinaus — ein 400-px-Bild auf 600
  // aufgezogen ist die Unschärfe, die der Nutzer am 23.09.2026 gemeldet hat.
  if (width > 0 && width < m.width) m.width = width;
  return m;
}

/**
 * v31.80: Logos werden bis 1200 px breit gespeichert (vorher 600). Mail-
 * Clients auf Bildschirmen mit doppelter Pixeldichte zeigen den 600-px-Kopf
 * mit 1200 Gerätepixeln — ein 600-px-Bild war dort immer weich, ein Logo
 * mit Schrift (Hofbräu-Beispiel, 23.09.2026) sichtbar matschig. Die Größe
 * hält `shrinkLogoB64` mit einer Qualitäts-Leiter unter 250 KB.
 */
export const LOGO_MAX_BREITE = 1200;

/** Trägt ein Layout genau die automatische Vollbreiten-Vorgabe (600/0/0)? */
export function istVolleBreite(l: KopfMasse): boolean {
  return l.width === 600 && l.paddingV === 0 && l.paddingH === 0;
}

/**
 * v31.78: Maße einer Data-URL SYNCHRON aus dem Dateikopf (PNG, JPEG, GIF) —
 * ohne `Image()`, damit die Regel auch dort greift, wo Mails ohne Oberfläche
 * gebaut werden (Bestätigung, Warteliste, Organizer-Mails, Outlook-Body).
 * `null`, wenn das Format nicht lesbar ist (WebP, SVG, kein Base64).
 */
export function bildMasseSync(dataUrl: string): { width: number; height: number } | null {
  try {
    if (!dataUrl || dataUrl.indexOf('data:image/') !== 0) return null;
    const comma = dataUrl.indexOf(',');
    if (comma < 0 || dataUrl.slice(0, comma).indexOf(';base64') < 0) return null;
    // 128 KB Base64 reichen: Der Bildkopf steht vorn; Canvas-Ausgaben tragen
    // weder EXIF noch ICC-Profil, der JPEG-SOF-Marker kommt früh.
    let b64 = dataUrl.slice(comma + 1, comma + 1 + 131072);
    b64 = b64.slice(0, b64.length - (b64.length % 4));
    const bin = atob(b64);
    const n = bin.length;
    const at = (i: number): number => bin.charCodeAt(i) & 0xff;
    if (n > 24 && at(0) === 0x89 && at(1) === 0x50 && at(2) === 0x4e && at(3) === 0x47) {
      return { width: ((at(16) << 24) | (at(17) << 16) | (at(18) << 8) | at(19)) >>> 0, height: ((at(20) << 24) | (at(21) << 16) | (at(22) << 8) | at(23)) >>> 0 };
    }
    if (n > 10 && bin.slice(0, 3) === 'GIF') {
      return { width: at(6) | (at(7) << 8), height: at(8) | (at(9) << 8) };
    }
    if (n > 4 && at(0) === 0xff && at(1) === 0xd8) {
      let i = 2;
      while (i + 9 < n) {
        if (at(i) !== 0xff) { i++; continue; }
        const m = at(i + 1);
        if (m === 0xff) { i++; continue; }
        if (m === 0xd8 || m === 0x01 || (m >= 0xd0 && m <= 0xd7)) { i += 2; continue; }
        const len = (at(i + 2) << 8) | at(i + 3);
        const sof = m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc;
        if (sof) return { height: (at(i + 5) << 8) | at(i + 6), width: (at(i + 7) << 8) | at(i + 8) };
        i += 2 + Math.max(2, len);
      }
    }
    return null;
  } catch { return null; }
}

/**
 * v31.78: DIE Formregel für das Mail-LOGO eines Events — Nutzer-Ansage
 * 22.09.2026: „bei E-Mail-Versand auch darauf achten: wenn das Bild ca. rund
 * ist (also nicht komplett quer), dann wäre 300 px besser als Standardgröße."
 *
 * Gilt nur für die AUTOMATISCHE Vollbreiten-Vorgabe 600/0/0 (v29.29-Default
 * neuer Events, v30.87-Regel „eigenes Logo → volle Breite"). Ist das Logo
 * nicht deutlich breiter als hoch (Breite ≤ 1,3 × Höhe), wird daraus 300 px
 * mit Rand. Jede andere, selbst gesetzte Breite bleibt unangetastet — wer
 * 400 px eingestellt hat, hat das gewollt. Ohne lesbare Maße bleibt es beim
 * Layout, wie es ist.
 */
export function formRegelKopf<T extends KopfMasse>(layout: T, logoB64?: string | null): T {
  if (!logoB64) return layout;
  const m = bildMasseSync(logoB64);
  if (!m || !m.width || !m.height) return layout;
  let out: T = layout;
  if (istVolleBreite(layout) && m.width / m.height <= 1.3) out = { ...layout, ...KOPF_RUND };
  // v31.80: Nie breiter rendern, als das Bild Pixel hat. Ein altes Logo, das
  // der Wizard auf 360 oder 480 px verkleinert hatte, wurde vom Kopf auf 600
  // px aufgezogen — genau die weiche Schrift aus dem Nutzer-Befund vom
  // 23.09.2026 („das Bild im Header ist echt ganz schön schlechte Qualität").
  if (m.width < out.width) out = { ...out, width: m.width };
  return out;
}

/** Das Mail-Logo (`_eventLogo`) aus dem Overrides-Blob — '' ohne. */
export function eventLogoAus(overridesJson: string | undefined | null): string {
  try {
    const o = JSON.parse(overridesJson || '{}') || {};
    return typeof o._eventLogo === 'string' ? o._eventLogo : '';
  } catch { return ''; }
}

/** Maße einer Data-URL messen — 0/0, wenn das Bild nicht dekodierbar ist. */
export function bildMasse(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise(resolve => {
    try {
      const i = new Image();
      i.onload = () => resolve({ width: i.naturalWidth || 0, height: i.naturalHeight || 0 });
      i.onerror = () => resolve({ width: 0, height: 0 });
      i.src = dataUrl;
    } catch { resolve({ width: 0, height: 0 }); }
  });
}

/** v31.76: Trägt ein FOTO noch die Logo-Vorgabe 600/0/0? Dann hat es die
 *  Maße nur geerbt (vor v31.76 gab es keine Formprüfung) — beim Öffnen des
 *  Editors wird die Form nachgemessen. */
export function fotoMitLogoMassen(img: MailHeaderImage): boolean {
  return (img.hero === 'event' || img.hero === 'custom') && img.width === 600 && img.paddingV === 0 && img.paddingH === 0;
}

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
  eventMailLogo: string | undefined | null,
  customB64?: string
): boolean {
  return (img.hero === 'event' && !!eventPhotoB64)
    || (img.hero === 'custom' && !!customB64)
    || !!eventMailLogo;
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
  eventPhotoB64: string,
  customB64?: string
): string {
  // v31.9.7: `custom` = ein für diese eine Mail hochgeladenes Bild. Wie beim
  // Event-Foto wird es fest eingebacken; ohne Bild bleibt der Platzhalter
  // stehen und der Flow setzt wie gehabt das Standard-Bild ein.
  if (img.hero === 'custom' && customB64) {
    return wrappedHtml.replace(/\{\{ORB_URL\}\}/g, customB64);
  }
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
export function normalizeMailHeaderImage(raw: unknown, allowCustom?: boolean): MailHeaderImage {
  const o = (raw && typeof raw === 'object') ? raw as Record<string, unknown> : {};
  const num = (v: unknown, def: number, max: number): number => {
    const n = typeof v === 'number' ? v : parseInt(String(v ?? ''), 10);
    if (!isFinite(n) || n < 0) return def;
    return Math.min(n, max);
  };
  return {
    // v31.9.7: `custom` kommt bewusst NICHT aus dem gespeicherten JSON
    // zurueck, solange kein Bild dazu vorliegt — ein gespeichertes `custom`
    // ohne Bild waere eine Auswahl, die still auf den Platzhalter faellt.
    // v31.74: Der Aufrufer sagt mit `allowCustom`, dass er das Bild hat
    // (QR-Mail: `headerCustomB64` im Override).
    hero: o.hero === 'event' ? 'event' : (o.hero === 'custom' && allowCustom ? 'custom' : 'logo'),
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
  // v31.78: rundes Logo → 300 px statt der automatischen vollen Breite.
  return mailHeaderOpts(own ? formRegelKopf(base, mailLogoB64 || logo) : base, own);
}

/** v31.0: Hat das Event ein eigenes Mail-Logo (`_eventLogo` im Blob oder
 *  die Spalte EmailImageBase64)? Nur dann darf „Volle Breite" ohne Event-Foto
 *  gelten — sonst wäre es der Orb in 600 px. */
export function hasOwnMailLogo(overridesJson: string | undefined | null, mailLogoB64?: string | null): boolean {
  if (mailLogoB64 && mailLogoB64.trim()) return true;
  try {
    const o = JSON.parse(overridesJson || '{}') || {};
    return typeof o._eventLogo === 'string' && o._eventLogo.trim().length > 0;
  } catch { return false; }
}

/**
 * v31.0: Kopf-Maße einer Mail mit eigenem Override (QR-Mail): gespeicherte
 * Werte gewinnen; OHNE gespeicherte Werte gilt die v30.87-Regel aus
 * `eventHeaderImageOpts` — gespeichertes `_headerImageLayout`, sonst volle
 * Breite bei eigenem Mail-Logo, sonst der kleine Standard. Vorher lieferte
 * `normalizeMailHeaderImage(undefined)` hier immer 180/30/30, und die QR-Mail
 * zeigte ein eigenes Logo klein, während alle anderen Mails es voll zeigten
 * (Nutzer 07.09.2026: „warum ist die Vorschau vom QR-Code schon wieder so
 * klein?").
 */
export function resolveMailHeaderImage(raw: unknown, overridesJson: string | undefined | null, mailLogoB64?: string | null, allowCustom?: boolean): MailHeaderImage {
  // v31.78: Steht das Logo im Kopf, gilt die Formregel (rund → 300 px) für
  // die automatische Vollbreite — an jedem der drei Ausgänge unten.
  const mitFormregel = (r: MailHeaderImage): MailHeaderImage =>
    r.hero === 'logo' ? formRegelKopf(r, mailLogoB64 || eventLogoAus(overridesJson)) : r;
  const o = (raw && typeof raw === 'object') ? raw as Record<string, unknown> : null;
  const storedWidth = o ? (typeof o.width === 'number' ? o.width : parseInt(String(o.width ?? ''), 10)) : NaN;
  if (o && isFinite(storedWidth) && storedWidth > 0) return mitFormregel(normalizeMailHeaderImage(o, allowCustom));
  let il: { width?: unknown; paddingV?: unknown; paddingH?: unknown } = {};
  try {
    const ov = JSON.parse(overridesJson || '{}') || {};
    il = (ov._headerImageLayout && typeof ov._headerImageLayout === 'object') ? ov._headerImageLayout : {};
  } catch { /* Defaults */ }
  if (typeof il.width === 'number' && il.width > 0) {
    return mitFormregel({
      hero: 'logo',
      width: il.width,
      paddingV: (typeof il.paddingV === 'number' && il.paddingV >= 0) ? il.paddingV : 30,
      paddingH: (typeof il.paddingH === 'number' && il.paddingH >= 0) ? il.paddingH : 30,
    });
  }
  return hasOwnMailLogo(overridesJson, mailLogoB64) ? mitFormregel({ hero: 'logo', width: 600, paddingV: 0, paddingH: 0 }) : { ...MAIL_HEADER_IMAGE_DEFAULT };
}

export function isDefaultMailHeaderImage(img: MailHeaderImage): boolean {
  return img.hero === MAIL_HEADER_IMAGE_DEFAULT.hero
    && img.width === MAIL_HEADER_IMAGE_DEFAULT.width
    && img.paddingV === MAIL_HEADER_IMAGE_DEFAULT.paddingV
    && img.paddingH === MAIL_HEADER_IMAGE_DEFAULT.paddingH;
}

