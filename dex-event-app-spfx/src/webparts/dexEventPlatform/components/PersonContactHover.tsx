/**
 * v24.56: Avatar mit Hover-Kontaktkarte (E-Mail + MS-Teams-Chat) — analog zur
 * Organizer-Hover-Karte auf der Anmeldeseite, aber als wiederverwendbare
 * Komponente für die Teilnehmerliste im Organizer Center.
 *
 * Mouse-Over auf das Foto öffnet ein kleines Popover (fixed positioniert, damit
 * Tabellen-Overflow nichts abschneidet) mit großem Foto, Name, klickbarer
 * E-Mail-Adresse und Teams-Chat-Link + (falls vorhanden) Position/Standort/Firma.
 */
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { useRoles } from '../context/RoleContext';

// v26.8: Modul-globaler Cache für lazy nachgeladene Profil-Infos (Position +
// Standort) — analog zur Organizer-Hover-Karte auf der Anmeldeseite
// (OrganizerList.tsx), damit dieselbe Person nicht mehrfach abgefragt wird.
const profileCache = new Map<string, { jobTitle: string; location: string }>();

/*
 * v31.27 — Foto-Gedaechtnis fuer die Sitzung.
 *
 * Nutzer-Frage 11.09.2026: „kann das Foto der Teilnehmer und der Assistenzen
 * nicht besser wirklich in die TN-Liste im SP rein? Dann wird es nicht immer
 * wie neu geladen. … Oder ist der Live-Abruf ueber die SST genauso gut und
 * schlau?"
 *
 * Der Live-Abruf ist besser, und zwar aus vier Gruenden:
 *
 *  1. **SharePoint Online cached Profilbilder ohnehin 24 Stunden** im Browser.
 *     Ein zweiter Aufruf holt sie nicht neu, er holt sie aus der Platte.
 *  2. **Ein Foto in der Liste waere eine Kopie personenbezogener Daten** — in
 *     JEDER Teilnehmerliste, also dutzendfach ueber alle Subsites, jede mit
 *     eigener Aufbewahrung. DEX loescht Teilnehmerlisten nach drei Monaten;
 *     Gesichter zusaetzlich zu verteilen ist kein Performance-Thema, sondern
 *     ein Datenschutz-Thema.
 *  3. **Es waere ein Standbild.** Wer sein Profilbild wechselt, haette in DEX
 *     fuer immer das alte — und niemand wuesste, warum.
 *  4. **Es waere langsamer, nicht schneller.** 400 Zeilen mal rund 4 KB
 *     base64 sind 1,6 MB, die bei JEDEM Lesen der Teilnehmerliste mitkaemen —
 *     genau die Datenmenge, die v31.24/v31.26 herausgenommen haben. Und zwar
 *     auch fuer Zeilen, zu denen niemand scrollt.
 *
 * Was wirklich zu oft passiert, ist etwas anderes, und das steht hier: Hat
 * eine Person KEIN Foto, antwortet SharePoint mit einem Platzhalter bzw.
 * Fehler — und bisher fragte jede Zeile, jedes Popover und jeder
 * Reiterwechsel erneut, weil `failed` nur im Zustand der einzelnen Komponente
 * lag. Bei vierhundert Zeilen ohne Bild sind das vierhundert Anfragen, die
 * garantiert nichts liefern.
 *
 * Beide Mengen leben nur fuer die Sitzung. Ein dauerhafter Merker waere die
 * Standbild-Falle aus Grund 3 in klein.
 */
const photoFailed = new Set<string>();
const photoOk = new Set<string>();

function getInitials(name: string): string {
  const parts = (name || '').includes(',')
    ? name.split(',').reverse().map(s => s.trim())
    : (name || '').split(/\s+/).filter(Boolean);
  return parts.slice(0, 2).map(p => (p[0] || '').toUpperCase()).join('');
}

function photoUrl(email: string, size: 'S' | 'M' | 'L'): string {
  if (!email) return '';
  return `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=${size}`;
}

function teamsChatUrl(email: string): string {
  return `msteams:/l/chat/0/0?users=${encodeURIComponent(email)}`;
}

export interface PersonContactHoverProps {
  email: string;
  name: string;
  /** Avatar-Durchmesser in px (Default 30). */
  size?: number;
  subline?: string; // "Position • Standort • Firma"
  isDe?: boolean;
}

export function PersonContactHover(props: PersonContactHoverProps): React.ReactElement {
  const { email, name, size = 30, subline, isDe = true } = props;
  const { searchUser } = useRoles();
  const [failed, setFailed] = React.useState<boolean>(() => photoFailed.has((props.email || '').toLowerCase()));
  const [coords, setCoords] = React.useState<{ x: number; y: number; above: boolean } | null>(null);
  const [open, setOpen] = React.useState(false);
  // v26.8: Wird keine subline übergeben, laden wir Position + Standort beim
  // ersten Hover live nach (gleicher Mechanismus wie die Anmeldeseite).
  const emailLc = (email || '').toLowerCase();
  const [fetchedSub, setFetchedSub] = React.useState<string>(
    emailLc && profileCache.has(emailLc)
      ? [profileCache.get(emailLc)!.jobTitle, profileCache.get(emailLc)!.location].filter(Boolean).join(' · ')
      : ''
  );
  const wrapperRef = React.useRef<HTMLSpanElement>(null);
  const closeTimer = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const initials = getInitials(name);
  // v29.29: Foto erst laden, wenn der Avatar in den sichtbaren Bereich kommt.
  // Bei einer Teilnehmerliste mit mehreren hundert Zeilen feuerte die Seite
  // sonst ebenso viele userphoto.aspx-Anfragen auf einmal ab — die Liste
  // scrollt in einem eigenen Container, sichtbar sind aber immer nur wenige
  // Zeilen. Bis zum Laden steht der Initialen-Kreis (gleiche Größe, also
  // kein Springen des Layouts). Vorlauf von 300 px, damit beim Scrollen
  // nichts nachzieht.
  // v31.27: Ist das Foto in dieser Sitzung schon einmal geladen worden, liegt
  // es im Browser-Cache — dann gleich zeigen statt erst auf den Scroll zu
  // warten. Das nimmt der Liste das Nachblitzen beim Reiterwechsel.
  const [visible, setVisible] = React.useState<boolean>(() => photoOk.has((props.email || '').toLowerCase()));
  React.useEffect(() => {
    if (visible || !email) return undefined;
    const el = wrapperRef.current;
    // Ohne IntersectionObserver (alte Browser) sofort laden — lieber die
    // Anfragen als gar kein Foto.
    if (!el || typeof IntersectionObserver === 'undefined') { setVisible(true); return undefined; }
    const io = new IntersectionObserver(entries => {
      if (entries.some(e => e.isIntersecting)) { setVisible(true); io.disconnect(); }
    }, { rootMargin: '300px' });
    io.observe(el);
    return () => io.disconnect();
  }, [visible, email]);
  const popoverHeight = 200;
  const effectiveSub = subline || fetchedSub || undefined;

  const cancelClose = (): void => { if (closeTimer.current) { clearTimeout(closeTimer.current); closeTimer.current = null; } };
  const scheduleClose = (): void => { cancelClose(); closeTimer.current = setTimeout(() => setOpen(false), 200); };
  React.useEffect(() => () => cancelClose(), []);

  const openPopover = (): void => {
    if (!email) return;
    const r = wrapperRef.current?.getBoundingClientRect();
    if (!r) return;
    const spaceBelow = window.innerHeight - r.bottom;
    const above = spaceBelow < popoverHeight + 16;
    setCoords({ x: r.left + r.width / 2, y: above ? r.top - 8 : r.bottom + 8, above });
    setOpen(true);
    // Position + Standort lazy nachladen (nur wenn keine subline gesetzt ist).
    if (!subline && !fetchedSub && emailLc) {
      searchUser(email).then((res) => {
        if (!res) return;
        const entry = { jobTitle: res.jobTitle || '', location: res.location || '' };
        profileCache.set(emailLc, entry);
        const sub = [entry.jobTitle, entry.location].filter(Boolean).join(' · ');
        if (sub) setFetchedSub(sub);
      }).catch(() => { /* best-effort */ });
    }
  };

  // v26.36: Auf Touch-/Handy-Geräten gibt es kein Hover — ein TAP auf den Avatar
  // öffnet/schließt die Kontaktkarte (E-Mail + Teams-Chat sonst unerreichbar).
  // Desktop-Hover bleibt unverändert.
  const toggleOnTap = (e: React.MouseEvent): void => {
    if (!email) return;
    e.stopPropagation();
    if (open) { cancelClose(); setOpen(false); } else { cancelClose(); openPopover(); }
  };

  // v26.36: Tap außerhalb schließt die (per Tap geöffnete) Karte wieder.
  React.useEffect(() => {
    if (!open) return;
    const onDocClick = (ev: MouseEvent): void => {
      if (wrapperRef.current && !wrapperRef.current.contains(ev.target as Node)) setOpen(false);
    };
    document.addEventListener('click', onDocClick, true);
    return () => document.removeEventListener('click', onDocClick, true);
  }, [open]);

  return (
    <span ref={wrapperRef} style={{ display: 'inline-flex', flexShrink: 0 }}
      onMouseEnter={() => { cancelClose(); openPopover(); }}
      onMouseLeave={scheduleClose}
      onClick={toggleOnTap}
    >
      {!failed && email && visible ? (
        <img
          /*
           * v31.26 — zwei Bremsen auf einmal (Nutzer-Vermutung 11.09.2026:
           * „ich glaube es hat auch was mit den ganzen Fotos zu tun … kann man
           * das ggf. auch noch optimieren?" — sie stimmte).
           *
           * (1) GRÖSSE. Hier stand fest `'L'`. Der Avatar in der
           *     Teilnehmerliste ist 34 px groß und lud das GROSSE Bild
           *     (SharePoint liefert dafür rund 200 px Kantenlänge). Das ist
           *     ein Vielfaches der Bytes für Pixel, die niemand sieht. Die
           *     Größe richtet sich jetzt nach der tatsächlichen Darstellung;
           *     `M` ab 40 px, damit Retina-Displays nicht unscharf werden
           *     (dort ist ein 34-px-Kreis physisch 68 px).
           *
           * (2) PRIORITÄT. Die Fotos gehen an DENSELBEN Host wie die
           *     Teilnehmerlisten, und ein Browser öffnet je Host nur rund
           *     sechs Verbindungen. Ohne Angabe stehen vierhundert Fotos
           *     gleichrangig VOR der nächsten Listen-Abfrage in der
           *     Warteschlange — gemessen: eine Liste, die sonst 320 ms
           *     braucht, kam einmal erst nach 1353 ms. `fetchpriority="low"`
           *     stellt sie hinten an; ein Gesicht darf warten, eine
           *     Teilnehmerliste nicht.
           *
           * Kleingeschrieben, weil React 17 das camelCase-`fetchPriority`
           * noch nicht kennt und es als unbekanntes Attribut verwerfen würde.
           * Unbekannte KLEINGESCHRIEBENE Attribute reicht React unverändert
           * ans DOM durch — genau das wollen wir hier.
           */
          src={photoUrl(email, size <= 36 ? 'S' : (size <= 64 ? 'M' : 'L'))}
          alt={name}
          loading="lazy"
          decoding="async"
          // Der React-17-Typ kennt `fetchpriority` nicht; das Attribut selbst
          // reicht React unverändert ans DOM durch. Der Spread ist die einzige
          // Stelle, an der TypeScript nicht im Weg steht — ohne `any`-Cast auf
          // das ganze Element.
          {...({ fetchpriority: 'low' } as Record<string, string>)}
          onError={() => { photoFailed.add((email || '').toLowerCase()); setFailed(true); }}
          onLoad={() => { photoOk.add((email || '').toLowerCase()); }}
          style={{ width: size, height: size, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0, cursor: 'default' }}
        />
      ) : (
        <span style={{ width: size, height: size, borderRadius: '50%', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', background: 'linear-gradient(135deg, #86bc25, #0076a8)', color: '#fff', fontSize: size * 0.4, fontWeight: 700, flexShrink: 0 }}>{initials}</span>
      )}

      {open && email && coords && (
        <span
          onMouseEnter={cancelClose}
          onMouseLeave={scheduleClose}
          style={{
            position: 'fixed',
            top: coords.above ? undefined : coords.y,
            bottom: coords.above ? window.innerHeight - coords.y : undefined,
            // v26.36: x-Position in den Viewport clampen, damit die (auf 90vw
            // gedeckelte) Karte auf schmalen Handy-Screens nie halb aus dem Bild
            // ragt. halfWidth = halbe maximale Kartenbreite (90vw).
            left: (() => { const hw = Math.min(180, window.innerWidth * 0.45); return Math.min(Math.max(coords.x, hw), window.innerWidth - hw); })(),
            transform: 'translateX(-50%)', zIndex: 3000,
            background: '#fff', borderRadius: 10, padding: 12,
            boxShadow: '0 8px 24px rgba(0,0,0,0.18)', border: '1px solid var(--dex-gray-200)',
            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6, minWidth: 180, maxWidth: '90vw',
          }}
        >
          {!failed && (
            <img src={photoUrl(email, 'L')} alt={name} onError={() => setFailed(true)}
              style={{ width: 96, height: 96, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-200)' }} />
          )}
          <span style={{ fontSize: '0.82rem', fontWeight: 700, color: 'var(--dex-gray-800)', textAlign: 'center', whiteSpace: 'normal', wordBreak: 'break-word' }}>{name}</span>
          {effectiveSub && <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)', textAlign: 'center' }}>{effectiveSub}</span>}
          <a href={`mailto:${email}`}
            style={{ fontSize: '0.75rem', color: 'var(--dex-green, #86bc25)', textDecoration: 'none', fontWeight: 600, whiteSpace: 'normal', wordBreak: 'break-all', textAlign: 'center' }}
            onMouseEnter={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'underline'; }}
            onMouseLeave={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'none'; }}
          >{email}</a>
          <a href={teamsChatUrl(email)}
            style={{ display: 'inline-flex', alignItems: 'center', gap: 4, fontSize: '0.72rem', color: '#6264A7', textDecoration: 'none', fontWeight: 600, whiteSpace: 'nowrap' }}
            onMouseEnter={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'underline'; }}
            onMouseLeave={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'none'; }}
          >
            <Icon iconName="TeamsLogo" style={{ fontSize: 13 }} /> {isDe ? 'Teams-Chat' : 'Teams chat'}
          </a>
        </span>
      )}
    </span>
  );
}

export default PersonContactHover;
