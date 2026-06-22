/**
 * Gemeinsame Komponente für die Anzeige von Organisatoren mit Foto.
 *
 * Verwendet die SharePoint-Userphoto-URL:
 *   /_layouts/15/userphoto.aspx?accountname=<email>&size=<S|M|L>
 *
 * - Default: kleines Avatar (28px) plus Name als Chip
 * - Mouse-Over: größeres Bild (96px) plus Name in Tooltip-Box
 * - Fallback: Initialen-Avatar, falls Foto nicht geladen werden kann
 */

import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';

export interface OrganizerListProps {
  names: string[];
  emails: string[];
  size?: 'sm' | 'md';
  compact?: boolean;
  /** v24.8 (J): E-Mails (lowercase), die NICHT angezeigt werden sollen
   *  (einzeln ausgeblendete Organizer). Nur für teilnehmer-seitige Anzeigen
   *  übergeben — in Verwaltungs-Ansichten weglassen. */
  hiddenEmails?: string[];
  /** v23.25: 'card' = Organizer dauerhaft groß (Foto + Name + Mail + Rolle
   *  direkt sichtbar). Default 'chip' = klein mit Hover-Popup. */
  display?: 'chip' | 'card';
  /** v24.11: Sprache erzwingen (z.B. Anmeldeseite mit fest englischer
   *  Formularsprache). undefined = App-Sprache. */
  forceIsDe?: boolean;
}

// v11.95: pro Email einmalig profile-lookup, Ergebnis App-weit gecached
// (nur Tab-Lebensdauer). Verhindert dass jeder Hover einen neuen REST-Call
// triggert wenn der User mehrfach drüberfährt.
const profileCache = new Map<string, { jobTitle: string; location: string }>();

function getInitials(name: string): string {
  const parts = name.includes(',')
    ? name.split(',').reverse().map(s => s.trim())
    : name.split(/\s+/).filter(Boolean);
  return parts.slice(0, 2).map(p => (p[0] || '').toUpperCase()).join('');
}

function photoUrl(email: string, size: 'S' | 'M' | 'L'): string {
  if (!email) return '';
  return `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}&size=${size}`;
}

// v23.26/v23.27: Deep-Link in einen 1:1-MS-Teams-Chat mit der Person.
// Wir nutzen das `msteams:`-Protokoll statt der `https://teams.microsoft.com/l/…`-
// URL — letztere zwingt den Browser über die „Stay better connected …"-
// Launcher-Zwischenseite. Das Protokoll öffnet die installierte Teams-App
// direkt (im Deloitte-Tenant überall vorhanden).
function teamsChatUrl(email: string): string {
  return `msteams:/l/chat/0/0?users=${encodeURIComponent(email)}`;
}

/** v23.26: Mail- + MS-Teams-Chat-Link, in Popup und großer Karte gleich. */
function ContactLinks({ email, isDe }: { email: string; isDe: boolean }): React.ReactElement | null {
  if (!email) return null;
  return (
    <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: 3 }}>
      <a
        href={`mailto:${email}`}
        style={{ fontSize: '0.75rem', color: 'var(--dex-green, #86bc25)', textDecoration: 'none', fontWeight: 600, whiteSpace: 'nowrap' }}
        onMouseEnter={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'underline'; }}
        onMouseLeave={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'none'; }}
      >{email}</a>
      <a
        href={teamsChatUrl(email)}
        style={{ display: 'inline-flex', alignItems: 'center', gap: 4, fontSize: '0.72rem', color: '#6264A7', textDecoration: 'none', fontWeight: 600, whiteSpace: 'nowrap' }}
        onMouseEnter={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'underline'; }}
        onMouseLeave={e => { (e.currentTarget as HTMLElement).style.textDecoration = 'none'; }}
      >
        <Icon iconName="TeamsLogo" style={{ fontSize: 13 }} /> {isDe ? 'Teams-Chat' : 'Teams chat'}
      </a>
    </span>
  );
}

/**
 * v23.26: Eine Person als großer Eintrag (Foto + Name + Kontakt + Rolle) —
 * für den „card"-Modus, wo mehrere Organizer NEBENEINANDER in EINER Kachel
 * stehen. Lädt Rolle + Standort sofort (ohne Hover).
 */
function OrganizerCardEntry({ name, email, forceIsDe }: { name: string; email: string; forceIsDe?: boolean }): React.ReactElement {
  const [failed, setFailed] = React.useState(false);
  const [retryAttempt, setRetryAttempt] = React.useState(0);
  const [profile, setProfile] = React.useState<{ jobTitle: string; location: string } | null>(
    email && profileCache.has(email.toLowerCase()) ? profileCache.get(email.toLowerCase()) as { jobTitle: string; location: string } : null
  );
  const { searchUser } = useRoles();
  const { locale } = useLanguage();
  const isDe = forceIsDe !== undefined ? forceIsDe : locale === 'de';
  const size = 96;
  const initials = getInitials(name);
  const cacheBust = retryAttempt > 0 ? `&_r=${retryAttempt}` : '';

  React.useEffect(() => {
    if (!email || profile) return;
    const cacheKey = email.toLowerCase();
    const cached = profileCache.get(cacheKey);
    if (cached) { setProfile(cached); return; }
    searchUser(email).then(res => {
      if (res) {
        const entry = { jobTitle: res.jobTitle || '', location: res.location || '' };
        profileCache.set(cacheKey, entry);
        setProfile(entry);
      }
    }).catch(() => { /* silent */ });
  }, [email, profile, searchUser]);

  return (
    <div style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: 5, textAlign: 'center', minWidth: 150, maxWidth: 200 }}>
      {!failed && email ? (
        <img
          src={`${photoUrl(email, 'L')}${cacheBust}`}
          alt={name}
          onError={() => { if (retryAttempt < 1) setRetryAttempt(retryAttempt + 1); else setFailed(true); }}
          style={{ width: size, height: size, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-200)' }}
        />
      ) : (
        <span style={{ width: size, height: size, borderRadius: '50%', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', background: 'linear-gradient(135deg, #86bc25, #0076a8)', color: '#fff', fontSize: size * 0.36, fontWeight: 700 }}>{initials}</span>
      )}
      <span style={{ fontSize: '0.88rem', fontWeight: 700, color: 'var(--dex-gray-800)' }}>{name}</span>
      <ContactLinks email={email} isDe={isDe} />
      {profile && (profile.jobTitle || profile.location) && (
        <span style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
          {[profile.jobTitle, profile.location].filter(Boolean).join(' · ')}
        </span>
      )}
    </div>
  );
}

function OrganizerChip({ name, email, sizeClass, isOpen, onOpen, onScheduleClose, onCancelClose, forceIsDe }: {
  name: string; email: string; sizeClass: 'sm' | 'md'; forceIsDe?: boolean;
  // v23.47: Auf/Zu-Status wird vom Eltern-Container gesteuert, damit IMMER nur
  // EIN Popover gleichzeitig offen ist (schnelles Wechseln zwischen Organizern
  // zeigte vorher kurz beide — der verzögerte Close des ersten überlappte mit
  // dem Open des zweiten).
  isOpen: boolean; onOpen: () => void; onScheduleClose: () => void; onCancelClose: () => void;
}): React.ReactElement {
  const [failed, setFailed] = React.useState(false);
  // SP /_layouts/15/userphoto.aspx liefert sporadisch transient 404 — z.B. wenn die
  // Mailbox gerade geprovisioned wird oder die Anfrage parallel zu vielen anderen
  // läuft. Ein einmaliger Retry mit Cache-Bust-Suffix recovered das oft. Erst nach
  // dem zweiten Fehlschlag fallen wir auf das Initialen-Avatar zurück.
  const [retryAttempt, setRetryAttempt] = React.useState(0);
  const [coords, setCoords] = React.useState<{ x: number; y: number; above: boolean } | null>(null);
  // v11.95: lazy-loaded JobTitle + Standort. Wird beim ersten Hover über
  // searchUser nachgeladen und im modul-globalen profileCache gemerkt.
  const [profile, setProfile] = React.useState<{ jobTitle: string; location: string } | null>(
    email && profileCache.has(email.toLowerCase()) ? profileCache.get(email.toLowerCase()) as { jobTitle: string; location: string } : null
  );
  const { searchUser } = useRoles();
  const { locale } = useLanguage();
  const isDe = forceIsDe !== undefined ? forceIsDe : locale === 'de';
  const wrapperRef = React.useRef<HTMLSpanElement>(null);
  const avatarSize = sizeClass === 'sm' ? 24 : 32;
  const enlargedSize = 120;
  const popoverHeight = 180; // ungefähre Popover-Höhe für Flip-Entscheidung
  const initials = getInitials(name);
  const cacheBust = retryAttempt > 0 ? `&_r=${retryAttempt}` : '';

  const openPopover = (): void => {
    const r = wrapperRef.current?.getBoundingClientRect();
    if (!r) return;
    const spaceBelow = window.innerHeight - r.bottom;
    const above = spaceBelow < popoverHeight + 16;
    const y = above ? r.top - 8 : r.bottom + 8;
    const x = r.left + r.width / 2;
    setCoords({ x, y, above });
    // v23.47: Eltern-Container öffnet dieses (und schließt jedes andere) Popover.
    onOpen();
    // v11.95: beim ersten Hover JobTitle + Standort lazy nachladen.
    // Cache pro Email — beim wiederholten Hover sofort verfügbar.
    if (email && !profile) {
      const cacheKey = email.toLowerCase();
      const cached = profileCache.get(cacheKey);
      if (cached) {
        setProfile(cached);
      } else {
        searchUser(email).then(res => {
          if (res) {
            const entry = { jobTitle: res.jobTitle || '', location: res.location || '' };
            profileCache.set(cacheKey, entry);
            setProfile(entry);
          }
        }).catch(() => { /* silent — falls Lookup fehlschlägt einfach nur Name+Email zeigen */ });
      }
    }
  };

  return (
    <span
      ref={wrapperRef}
      style={{
        display: 'inline-flex', alignItems: 'center', gap: 8,
        padding: sizeClass === 'sm' ? '2px 10px 2px 2px' : '3px 12px 3px 3px',
        background: 'var(--dex-gray-100, #f2f2f2)',
        borderRadius: 999,
        fontSize: sizeClass === 'sm' ? '0.78rem' : '0.85rem',
        color: 'var(--dex-gray-800)',
        position: 'relative',
        cursor: 'default',
      }}
      onMouseEnter={() => { onCancelClose(); openPopover(); }}
      onMouseLeave={onScheduleClose}
    >
      {!failed && email ? (
        <img
          src={`${photoUrl(email, 'S')}${cacheBust}`}
          alt={name}
          onError={() => {
            if (retryAttempt < 1) setRetryAttempt(retryAttempt + 1);
            else setFailed(true);
          }}
          style={{
            width: avatarSize, height: avatarSize, borderRadius: '50%',
            objectFit: 'cover', background: 'var(--dex-gray-200)',
            flexShrink: 0,
          }}
        />
      ) : (
        <span
          style={{
            width: avatarSize, height: avatarSize, borderRadius: '50%',
            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
            background: 'linear-gradient(135deg, #86bc25, #0076a8)',
            color: '#fff', fontSize: avatarSize * 0.42, fontWeight: 700,
            flexShrink: 0,
          }}
        >{initials}</span>
      )}
      <span style={{ whiteSpace: 'nowrap' }}>{name}</span>

      {/* Hover-Vergrößerung: fixed positioning damit Container-Overflow nichts abschneidet */}
      {isOpen && email && !failed && coords && (
        <span
          onMouseEnter={onCancelClose}
          onMouseLeave={onScheduleClose}
          style={{
            position: 'fixed',
            top: coords.above ? undefined : coords.y,
            bottom: coords.above ? window.innerHeight - coords.y : undefined,
            left: coords.x,
            transform: 'translateX(-50%)',
            zIndex: 2000,
            background: '#fff', borderRadius: 10, padding: 12,
            boxShadow: '0 8px 24px rgba(0,0,0,0.18)',
            border: '1px solid var(--dex-gray-200)',
            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6,
            // v23.25: pointerEvents aktiv, damit die Maus in die Karte fahren
            // und auf die Mail-Adresse klicken kann.
          }}
        >
          {/* v23.25: freundlicher Hinweis-Kopf. */}
          <span style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap', marginBottom: 2 }}>
            {isDe ? 'Bei Fragen wende dich gerne an:' : 'If you have any questions, feel free to reach out:'}
          </span>
          <img
            src={photoUrl(email, 'L')}
            alt={name}
            style={{
              width: enlargedSize, height: enlargedSize, borderRadius: '50%',
              objectFit: 'cover', background: 'var(--dex-gray-200)',
            }}
          />
          <span style={{ fontSize: '0.8rem', fontWeight: 600, whiteSpace: 'nowrap' }}>{name}</span>
          {/* v23.25/v23.26: klickbarer Mailto- + Teams-Chat-Link. */}
          <ContactLinks email={email} isDe={isDe} />
          {/* v11.95: JobTitle + Standort aus dem SP-Profil — lazy beim
              ersten Hover geladen, danach gecached. */}
          {profile && (profile.jobTitle || profile.location) && (
            <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)', whiteSpace: 'nowrap', textAlign: 'center' }}>
              {[profile.jobTitle, profile.location].filter(Boolean).join(' · ')}
            </span>
          )}
        </span>
      )}
    </span>
  );
}

/**
 * Extrahiert Lastname-Tokens aus einem Display-Namen.
 *  "von Albedyll, Benedikt" → ["von", "albedyll"]
 *  "Benedikt von Albedyll" → ["von", "albedyll"]  (letzte 2 Worte als Lastname-Heuristik bei Adelspräfix)
 *  "Heymann, Thorsten"     → ["heymann"]
 *  "Eike Brenneisen"       → ["brenneisen"]
 */
function lastnameTokens(name: string): string[] {
  const trimmed = (name || '').trim();
  if (!trimmed) return [];
  let lastnamePart: string;
  if (trimmed.indexOf(',') >= 0) {
    lastnamePart = trimmed.split(',')[0].trim();
  } else {
    // Heuristik: letztes Wort ist Lastname; falls vorletztes Wort ein Adelspräfix ist,
    // nimm die letzten beiden.
    const words = trimmed.split(/\s+/).filter(Boolean);
    if (words.length === 0) return [];
    if (words.length === 1) return [words[0].toLowerCase()];
    const PREFIXES = new Set(['von', 'van', 'de', 'der', 'den', 'di', 'da', 'du', 'la', 'le']);
    if (words.length >= 2 && PREFIXES.has(words[words.length - 2].toLowerCase())) {
      lastnamePart = words.slice(-2).join(' ');
    } else {
      lastnamePart = words[words.length - 1];
    }
  }
  return lastnamePart.toLowerCase().split(/\s+/).filter(t => t.length >= 3);
}

/**
 * Defensives Pairing: bevorzugt Index-Match wenn die Email-Local-Part den Lastname
 * der Person enthält, sonst sucht den ersten passenden Email-Eintrag aus dem Pool.
 *
 * Hintergrund: bei legacy oder closure-bug-betroffenen Events kann es vorkommen,
 * dass `organizers[]` und `organizerEmails[]` aus dem SP-Storage out-of-sync sind.
 * Reines `emails[i]`-Pairing zeigt dann das falsche Foto neben dem Namen. Mit dem
 * Lastname-Matcher korrigiert sich das visuell, auch wenn die Storage-Reihenfolge
 * gedreht ist.
 */
function pairNamesEmails(names: string[], emails: string[]): Array<{ name: string; email: string }> {
  const used = new Set<number>();
  const result: Array<{ name: string; email: string }> = [];
  const localParts = emails.map(e => (e || '').toLowerCase().split('@')[0]);

  const matchesLastname = (emailIdx: number, tokens: string[]): boolean => {
    if (emailIdx < 0 || emailIdx >= localParts.length) return false;
    const local = localParts[emailIdx];
    if (!local) return false;
    return tokens.some(t => local.indexOf(t) >= 0);
  };

  for (let i = 0; i < names.length; i++) {
    const tokens = lastnameTokens(names[i]);
    let chosen = -1;
    if (!used.has(i) && matchesLastname(i, tokens)) {
      chosen = i;
    } else {
      for (let j = 0; j < emails.length; j++) {
        if (!used.has(j) && matchesLastname(j, tokens)) { chosen = j; break; }
      }
    }
    if (chosen < 0) {
      // Kein Lastname-Match — Index-Fallback wenn Slot frei, sonst nächster freier Slot.
      if (!used.has(i) && i < emails.length) chosen = i;
      else for (let j = 0; j < emails.length; j++) if (!used.has(j)) { chosen = j; break; }
    }
    if (chosen >= 0) used.add(chosen);
    result.push({ name: names[i].trim(), email: chosen >= 0 ? (emails[chosen] || '').trim() : '' });
  }
  return result;
}

function OrganizerCardTile({ items, forceIsDe }: { items: Array<{ name: string; email: string }>; forceIsDe?: boolean }): React.ReactElement {
  const { locale } = useLanguage();
  const isDe = forceIsDe !== undefined ? forceIsDe : locale === 'de';
  // v23.26: EINE Kachel mit allen Organizern nebeneinander (statt einzeln
  // beim Mouse-Over). Ein Hinweis-Kopf, dann die Personen in einer Reihe.
  return (
    <div
      style={{
        display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: 12,
        background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 12,
        padding: '14px 20px', boxShadow: '0 2px 8px rgba(0,0,0,0.06)', maxWidth: '100%',
      }}
    >
      <span style={{ fontSize: '0.72rem', color: 'var(--dex-gray-500)' }}>
        {isDe ? 'Bei Fragen wende dich gerne an:' : 'If you have any questions, feel free to reach out:'}
      </span>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 20, justifyContent: 'center', alignItems: 'flex-start' }}>
        {items.map((o, i) => (
          <OrganizerCardEntry key={`${o.email || o.name}-${i}`} name={o.name} email={o.email} forceIsDe={forceIsDe} />
        ))}
      </div>
    </div>
  );
}

/**
 * v23.47: Container für die Chip-Variante. Hält EINEN gemeinsamen „offen"-Index
 * + EINEN gemeinsamen Close-Timer. Dadurch ist immer nur ein Hover-Popover
 * sichtbar: Fährt die Maus von einem Organizer-Chip zum nächsten, schließt das
 * vorige sofort (statt erst nach dem 200ms-Verzögerungs-Close — der vorher kurz
 * BEIDE Popover gleichzeitig zeigte).
 */
function OrganizerChipRow({ items, sizeClass, compact, forceIsDe }: {
  items: Array<{ name: string; email: string }>; sizeClass: 'sm' | 'md'; compact: boolean; forceIsDe?: boolean;
}): React.ReactElement {
  const [openIdx, setOpenIdx] = React.useState<number | null>(null);
  // Verzögertes Schließen, damit die Maus über die 8px-Lücke vom Chip in die
  // große Karte wandern kann, ohne dass sie sofort zuklappt.
  const closeTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const cancelClose = React.useCallback((): void => {
    if (closeTimerRef.current) { clearTimeout(closeTimerRef.current); closeTimerRef.current = null; }
  }, []);
  const scheduleClose = React.useCallback((): void => {
    cancelClose();
    closeTimerRef.current = setTimeout(() => setOpenIdx(null), 200);
  }, [cancelClose]);
  React.useEffect(() => () => cancelClose(), [cancelClose]);
  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', gap: compact ? 4 : 6 }}>
      {items.map((o, i) => (
        <OrganizerChip
          key={`${o.email || o.name}-${i}`}
          name={o.name}
          email={o.email}
          sizeClass={sizeClass}
          forceIsDe={forceIsDe}
          isOpen={openIdx === i}
          onOpen={() => { cancelClose(); setOpenIdx(i); }}
          onScheduleClose={scheduleClose}
          onCancelClose={cancelClose}
        />
      ))}
    </div>
  );
}

export default function OrganizerList({ names, emails, size = 'md', compact = false, display = 'chip', hiddenEmails, forceIsDe }: OrganizerListProps): React.ReactElement | null {
  let items = pairNamesEmails(names, emails).filter(o => !!o.name);
  // v24.8 (J): einzeln ausgeblendete Organizer aus der Anzeige nehmen (Rechte bleiben).
  if (hiddenEmails && hiddenEmails.length > 0) {
    const hidden = new Set(hiddenEmails.map(e => (e || '').toLowerCase()));
    items = items.filter(o => !(o.email && hidden.has(o.email.toLowerCase())));
  }
  if (items.length === 0) return null;
  // v23.26: Großer Modus = EINE gemeinsame Kachel mit allen Organizern.
  if (display === 'card') return <OrganizerCardTile items={items} forceIsDe={forceIsDe} />;
  return <OrganizerChipRow items={items} sizeClass={size} compact={compact} forceIsDe={forceIsDe} />;
}
