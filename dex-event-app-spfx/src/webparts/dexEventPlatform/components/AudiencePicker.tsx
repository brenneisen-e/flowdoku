/**
 * AudiencePicker — wiederverwendbare „Mailverteiler / einzelne User"-Auswahl
 * inkl. „Sichtbarkeit prüfen" und „Personen ausschließen".
 *
 * Extrahiert aus EventCreationPage, damit Hauptevent UND Sub-Events exakt
 * dieselbe Sichtbarkeits-UI nutzen. Jede Instanz hält ihren eigenen State
 * (Suche, Ergebnisse, international-Toggle, Chip-Suche/„Alle anzeigen",
 * Verteiler-Cache, Prüfen-Modal, Ausschließen-Modal, Massenimport-Modal),
 * sodass Hauptevent und jedes Sub-Event voneinander unabhängig sind.
 *
 * Datenformat unverändert: `value` ist ein kommaseparierter String aus
 * E-Mails / Gruppen, `onChange` liefert denselben String zurück.
 *
 * Der Ausschließen-Teil (`excludedUsers`) wird beim Hauptevent in das Event
 * persistiert — dafür gibt es die optionalen Props `excludedUsers` +
 * `onExcludedUsersChange`. Werden sie nicht übergeben (z.B. bei Sub-Events,
 * die kein eigenes Ausschluss-Feld speichern), hält die Komponente den
 * Ausschluss-State intern, damit die UI identisch bleibt.
 */

import * as React from 'react';
import { X, Users } from './Icons';
import InternationalSearchToggle from './InternationalSearchToggle';
import BulkUserImportModal from './BulkUserImportModal';
import { useRoles } from '../context/RoleContext';

const EXCLUDE_PAGE_SIZE = 200;

interface ResolvedUser {
  email: string;
  displayName: string;
  firstName: string;
  lastName: string;
  jobTitle: string;
  location: string;
  source: string;
}

interface Props {
  /** Kommaseparierter Audience-String (E-Mails / Gruppen). */
  value: string;
  /** Liefert den geänderten kommaseparierten String zurück. */
  onChange: (next: string) => void;
  /** Aktuell gewählte Standorte (kommasepariert) — für den kombinierten Sichtbarkeits-Test. */
  locationFilter: string;
  /** Verknüpfung zwischen Standortfilter und Audience. */
  filterMode: 'AND' | 'OR';
  isDe: boolean;
  /** Optional: extern persistierte Ausschluss-Liste (Hauptevent). */
  excludedUsers?: string[];
  /** Optional: Setter für die externe Ausschluss-Liste. */
  onExcludedUsersChange?: (updater: (prev: string[]) => string[]) => void;
  /** Optional: zwischen Mailverteiler-Karte und „Sichtbarkeit prüfen"-Zeile
   *  eingeschobener Inhalt (z.B. die Filterverknüpfung ODER/UND), damit die
   *  Reihenfolge der Sichtbarkeits-Sektion identisch bleibt. */
  middleSlot?: React.ReactNode;
  /** Hintergrundfarbe der Mailverteiler-Karte (Zebra-Alternation des Wizards). */
  cardBgPrimary?: string;
  /** Hintergrundfarbe der „Sichtbarkeit prüfen"-Zeile (Zebra-Alternation). */
  cardBgSecondary?: string;
  /** Optionales Schritt-Badge (Zahl) im Mailverteiler-Label — Hauptevent zeigt es. */
  stepBadge?: React.ReactNode;
}

export default function AudiencePicker({
  value,
  onChange,
  locationFilter,
  filterMode,
  isDe,
  excludedUsers: excludedUsersProp,
  onExcludedUsersChange,
  middleSlot,
  cardBgPrimary = '#fff',
  cardBgSecondary = '#fff',
  stepBadge,
}: Props): React.ReactElement {
  const { searchUsers, searchGroups, getGroupMembers, searchUsersByLocation } = useRoles();

  const audience = value;
  const setAudience = (updater: string | ((prev: string) => string)): void => {
    const next = typeof updater === 'function' ? (updater as (prev: string) => string)(audience) : updater;
    onChange(next);
  };

  // Audience-Suche (Personen + Verteiler/Security-Groups)
  const [audienceSearch, setAudienceSearch] = React.useState('');
  const [audienceResults, setAudienceResults] = React.useState<Array<{ kind: 'user' | 'group'; email: string; displayName: string }>>([]);
  const [isSearchingAudience, setIsSearchingAudience] = React.useState(false);
  const [audienceIncludeIntl, setAudienceIncludeIntl] = React.useState(false);
  const audienceTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  // Audience-Chip-Pagination + Inline-Suche
  const [audienceShowAll, setAudienceShowAll] = React.useState(false);
  const [audienceChipSearch, setAudienceChipSearch] = React.useState('');
  // Massenimport
  const [bulkAudienceOpen, setBulkAudienceOpen] = React.useState(false);

  // Modal: Members einer Gruppe anzeigen
  const [memberModalOpen, setMemberModalOpen] = React.useState(false);
  const [memberModalGroupName, setMemberModalGroupName] = React.useState('');
  const [memberModalLoading, setMemberModalLoading] = React.useState(false);
  const [memberModalMembers, setMemberModalMembers] = React.useState<Array<{ email: string; displayName: string }>>([]);
  const [memberModalError, setMemberModalError] = React.useState('');

  // Sichtbarkeit-Prüfen-Modal
  const [showEmailModal, setShowEmailModal] = React.useState(false);
  const [visibilityAudienceCache, setVisibilityAudienceCache] = React.useState<Set<string>>(new Set());
  const [visibilityCacheLoading, setVisibilityCacheLoading] = React.useState(false);
  const [emailSearch, setEmailSearch] = React.useState('');
  const [emailSearchResults, setEmailSearchResults] = React.useState<Array<{ email: string; displayName: string; location: string }>>([]);
  const [emailSearchIncludeIntl, setEmailSearchIncludeIntl] = React.useState(false);
  const [isSearchingEmails, setIsSearchingEmails] = React.useState(false);

  // Personen-ausschließen-Modal. Wenn keine externe Persistenz übergeben wird,
  // hält die Komponente die Ausschluss-Liste intern (UI bleibt identisch).
  const [internalExcluded, setInternalExcluded] = React.useState<string[]>([]);
  const excludedUsers = excludedUsersProp !== undefined ? excludedUsersProp : internalExcluded;
  const setExcludedUsers = React.useCallback((updater: (prev: string[]) => string[]): void => {
    if (onExcludedUsersChange) onExcludedUsersChange(updater);
    else setInternalExcluded(updater);
  }, [onExcludedUsersChange]);

  const [excludeModalOpen, setExcludeModalOpen] = React.useState(false);
  const [excludeResolvedUsers, setExcludeResolvedUsers] = React.useState<ResolvedUser[]>([]);
  const [excludeResolving, setExcludeResolving] = React.useState(false);
  const [excludeSearch, setExcludeSearch] = React.useState('');
  const [excludeIncludeIntl, setExcludeIncludeIntl] = React.useState(false);
  const [excludeFilters, setExcludeFilters] = React.useState<{ email: string; lastName: string; firstName: string; jobTitle: string; location: string }>({
    email: '', lastName: '', firstName: '', jobTitle: '', location: '',
  });
  const [excludeSortBy, setExcludeSortBy] = React.useState<'email' | 'lastName' | 'firstName' | 'jobTitle' | 'location'>('lastName');
  const [excludeSortDir, setExcludeSortDir] = React.useState<'asc' | 'desc'>('asc');
  const [excludePage, setExcludePage] = React.useState(0);

  const addAudienceItem = (val: string): void => {
    const list = audience.split(',').map(s => s.trim()).filter(Boolean);
    if (list.indexOf(val) >= 0) return;
    list.push(val);
    setAudience(list.join(', '));
  };
  const removeAudienceItem = (val: string): void => {
    const list = audience.split(',').map(s => s.trim()).filter(Boolean).filter(x => x !== val);
    setAudience(list.join(', '));
  };

  const openMembersModal = async (groupEmail: string): Promise<void> => {
    setMemberModalOpen(true);
    setMemberModalGroupName(groupEmail);
    setMemberModalMembers([]);
    setMemberModalError('');
    setMemberModalLoading(true);
    const res = await getGroupMembers(groupEmail);
    setMemberModalLoading(false);
    if (!res) {
      setMemberModalError('Mitglieder konnten nicht geladen werden. Vermutlich fehlt die Berechtigung "Group.Read.All" im SharePoint App Catalog (Admin Consent erforderlich).');
      return;
    }
    setMemberModalGroupName(res.groupName || groupEmail);
    setMemberModalMembers(res.members);
  };

  return (
    <>
      <div className="form-group" style={{ position: 'relative', padding: '16px 20px', marginBottom: 12, background: cardBgPrimary, borderRadius: 8, border: '1px solid var(--dex-gray-100)' }}>
        <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          {stepBadge}
          {isDe ? 'Mailverteiler / einzelne User' : 'Mailing lists / individual users'}
        </label>
        <p style={{ fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginTop: -4, marginBottom: 12, lineHeight: 1.5 }}>
          {isDe
            ? <>Wähle <strong>einzelne Personen</strong> oder ganze <strong>Mailverteiler bzw. Security-Gruppen aus Entra</strong> aus. Wenn auch ein Standortfilter gesetzt ist, kannst du unten festlegen, ob beide Bedingungen (UND) oder eine davon (ODER) reichen.</>
            : <>Pick <strong>individual people</strong> or entire <strong>mailing lists / security groups from Entra</strong>. If you also set a location filter, you can decide below whether both conditions (AND) or either of them (OR) is enough.</>}
        </p>
        {/* v16.4: Hinweis fuer den Organizer, dass Mitglieder zum
            Save-Zeitpunkt eingefroren werden und das Event bei DL-
            Mitglieder-Aenderungen einmal neu gespeichert werden muss,
            damit die neuen Mitglieder die Sichtbarkeit bekommen. */}
        <div style={{
          fontSize: '0.78rem', color: 'var(--dex-orange-dark, #b35a00)',
          background: 'rgba(237,139,0,0.08)', border: '1px dashed var(--dex-orange, #ed8b00)',
          borderRadius: 6, padding: '8px 12px', marginBottom: 12, lineHeight: 1.5,
        }}>
          {isDe
            ? <><strong>Hinweis:</strong> Die Mitglieder der ausgewählten Mailverteiler werden beim Speichern des Events einmal aufgelöst und gespeichert — das ist der schnelle Pfad für den Sichtbarkeits-Check. Wenn sich später Mitglieder eines Verteilers ändern (z.B. neue Person zur DL hinzugefügt), <strong>speichere das Event einmal neu</strong>, damit die App den frischen Stand bekommt.</>
            : <><strong>Note:</strong> The members of the selected distribution lists are resolved and cached when the event is saved — this is the fast path for the visibility check. If list members change later (e.g. new person added to a DL), <strong>re-save the event once</strong> to refresh the cache.</>}
        </div>
        {/* Chip-Liste der bereits ausgewaehlten Audience-Eintraege.
            Bei vielen Eintraegen: Inline-Suche + Pagination (nur 10 sichtbar, 'Mehr anzeigen'-Button). */}
        {audience.trim().length > 0 && (() => {
          const allEntries = audience.split(',').map(s => s.trim()).filter(Boolean);
          const chipSearchLc = audienceChipSearch.trim().toLowerCase();
          const filtered = chipSearchLc
            ? allEntries.filter(e => e.toLowerCase().indexOf(chipSearchLc) >= 0)
            : allEntries;
          const visibleLimit = 10;
          const visible = audienceShowAll || chipSearchLc ? filtered : filtered.slice(0, visibleLimit);
          const hiddenCount = filtered.length - visible.length;
          return (
            <div style={{ marginBottom: 8 }}>
              {/* Meta-Zeile mit Anzahl + Such-Input (nur wenn viele Eintraege) */}
              {allEntries.length > visibleLimit && (
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6, fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
                  <span>{allEntries.length} Einträge{chipSearchLc && ` — ${filtered.length} Treffer`}</span>
                  <input
                    type="text"
                    className="form-input"
                    value={audienceChipSearch}
                    onChange={e => setAudienceChipSearch(e.target.value)}
                    placeholder="In Zielgruppe suchen..."
                    style={{ flex: 1, maxWidth: 260, fontSize: '0.75rem', padding: '4px 8px' }}
                  />
                  {audienceChipSearch && (
                    <button
                      type="button"
                      onClick={() => setAudienceChipSearch('')}
                      style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 2, color: 'var(--dex-gray-400)' }}
                      title="Suche löschen"
                    >
                      <X size={14} />
                    </button>
                  )}
                </div>
              )}
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                {visible.map((entry, i) => {
                  const isEmail = entry.indexOf('@') >= 0;
                  return (
                    <span key={`${entry}-${i}`} style={{
                      display: 'inline-flex', alignItems: 'center', gap: 6,
                      padding: '4px 10px', borderRadius: 999,
                      background: isEmail ? 'rgba(0,118,168,0.10)' : 'rgba(134,188,37,0.12)',
                      color: isEmail ? 'var(--dex-blue, #0076a8)' : 'var(--dex-green-dark)',
                      fontSize: '0.8rem', fontWeight: 600,
                    }}>
                      {entry}
                      <button
                        type="button"
                        title="Mitglieder anzeigen"
                        onClick={() => openMembersModal(entry)}
                        style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 0, display: isEmail ? 'inline-flex' : 'none', color: 'inherit' }}
                      >
                        <Users size={12} />
                      </button>
                      <button
                        type="button"
                        onClick={() => removeAudienceItem(entry)}
                        title="Entfernen"
                        style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 0, color: 'inherit', fontWeight: 700 }}
                      >
                        <X size={12} />
                      </button>
                    </span>
                  );
                })}
              </div>
              {/* Mehr / Weniger Button */}
              {!chipSearchLc && allEntries.length > visibleLimit && (
                <button
                  type="button"
                  className="btn btn-secondary"
                  style={{ fontSize: '0.75rem', padding: '4px 10px', marginTop: 6 }}
                  onClick={() => setAudienceShowAll(!audienceShowAll)}
                >
                  {audienceShowAll ? `Weniger anzeigen (${allEntries.length - visibleLimit} ausblenden)` : `Alle anzeigen (+${hiddenCount} weitere)`}
                </button>
              )}
              {chipSearchLc && filtered.length === 0 && (
                <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400)', marginTop: 4, fontStyle: 'italic' }}>
                  Kein Treffer für &bdquo;{audienceChipSearch}&ldquo; in der Zielgruppe.
                </div>
              )}
            </div>
          );
        })()}
        {/* Such-Input */}
        <input
          className="form-input"
          value={audienceSearch}
          onChange={e => {
            const val = e.target.value;
            setAudienceSearch(val);
            if (audienceTimerRef.current) clearTimeout(audienceTimerRef.current);
            if (val.trim().length >= 2) {
              audienceTimerRef.current = setTimeout(async () => {
                setIsSearchingAudience(true);
                try {
                  const [users, groups] = await Promise.all([
                    searchUsers(val.trim(), audienceIncludeIntl),
                    searchGroups(val.trim()),
                  ]);
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const u: Array<{ kind: 'user' | 'group'; email: string; displayName: string }> = users.map((x: any) => ({ kind: 'user' as const, email: x.email, displayName: x.displayName }));
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  const g: Array<{ kind: 'user' | 'group'; email: string; displayName: string }> = groups.map((x: any) => ({ kind: 'group' as const, email: x.email, displayName: x.displayName }));
                  setAudienceResults([...g, ...u]); // Gruppen zuerst anzeigen
                } catch { setAudienceResults([]); }
                setIsSearchingAudience(false);
              }, 300);
            } else {
              setAudienceResults([]);
            }
          }}
          placeholder="Personen oder Gruppen suchen (z.B. SAPAlliance, max@deloitte.de, DEKOELN)"
        />
        <InternationalSearchToggle
          checked={audienceIncludeIntl}
          onChange={async next => {
            setAudienceIncludeIntl(next);
            const q = audienceSearch.trim();
            if (q.length >= 2) {
              setIsSearchingAudience(true);
              try {
                const [users, groups] = await Promise.all([
                  searchUsers(q, next),
                  searchGroups(q),
                ]);
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const u: Array<{ kind: 'user' | 'group'; email: string; displayName: string }> = users.map((x: any) => ({ kind: 'user' as const, email: x.email, displayName: x.displayName }));
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const g: Array<{ kind: 'user' | 'group'; email: string; displayName: string }> = groups.map((x: any) => ({ kind: 'group' as const, email: x.email, displayName: x.displayName }));
                setAudienceResults([...g, ...u]);
              } catch { setAudienceResults([]); }
              setIsSearchingAudience(false);
            }
          }}
          isDe={isDe}
        />
        {isSearchingAudience && (
          <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-400)', marginTop: 4 }}>Suche...</div>
        )}
        {audienceResults.length > 0 && (
          <div style={{
            position: 'absolute', left: 0, right: 0, top: '100%', zIndex: 100,
            background: '#fff', border: '1px solid var(--dex-gray-200)',
            borderRadius: 'var(--dex-radius)', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
            maxHeight: 280, overflowY: 'auto',
          }}>
            {audienceResults.map((r, i) => (
              <div
                key={`${r.kind}-${r.email}-${i}`}
                style={{
                  padding: '8px 12px', cursor: 'pointer', fontSize: '0.85rem',
                  borderBottom: '1px solid var(--dex-gray-100)',
                  display: 'flex', alignItems: 'center', gap: 8,
                }}
                onMouseDown={() => {
                  addAudienceItem(r.email);
                  setAudienceSearch('');
                  setAudienceResults([]);
                }}
              >
                <span style={{
                  fontSize: '0.7rem', fontWeight: 700,
                  padding: '2px 8px', borderRadius: 4,
                  background: r.kind === 'group' ? 'rgba(134,188,37,0.18)' : 'rgba(0,118,168,0.14)',
                  color: r.kind === 'group' ? 'var(--dex-green-dark)' : 'var(--dex-blue, #0076a8)',
                }}>
                  {r.kind === 'group' ? 'GRUPPE' : 'USER'}
                </span>
                <strong>{r.displayName}</strong>
                <span style={{ color: 'var(--dex-gray-400)', marginLeft: 'auto' }}>{r.email}</span>
              </div>
            ))}
          </div>
        )}
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: 6 }}>
          <button
            type="button"
            className="btn btn-secondary"
            style={{ fontSize: '0.75rem', padding: '6px 12px', whiteSpace: 'nowrap' }}
            onClick={() => setBulkAudienceOpen(true)}
          >
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}><Users size={12} /> Massenimport (Liste einfügen)</span>
          </button>
          <p style={{ fontSize: '0.75rem', color: 'var(--dex-gray-400)', margin: 0, flex: 1 }}>
            Klicke einen Treffer an um ihn hinzuzufügen. Bei Gruppen kannst du im Chip per <Users size={11} /> die Mitglieder einsehen.
            Statt zu suchen kannst du auch direkt die Verteiler-Mail eintippen (z.B. SAPAlliance@deloitte.com) oder Sondergruppen wie <code>DEALL</code>, <code>DEKOELN</code>.
          </p>
        </div>
      </div>

      {middleSlot}

      {/* v22.5: „Sichtbarkeit prüfen" ist IMMER verfügbar — auch wenn KEIN
          Filter gesetzt ist (zeigt dann „für alle sichtbar"). „Personen
          ausschließen" erscheint weiterhin nur, sobald ein Standortfilter
          ODER eine Audience gesetzt ist (sonst gibt es nichts auszuschließen). */}
      {(
        <div className="form-group" style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '16px 20px', marginBottom: 12, background: cardBgSecondary, borderRadius: 8, border: '1px solid var(--dex-gray-100)', flexWrap: 'wrap' }}>
          <button
            className="btn btn-outline"
            style={{ fontSize: '0.8rem', padding: '6px 14px', whiteSpace: 'nowrap' }}
            onClick={async () => {
              setShowEmailModal(true);
              // v8.9: Verteiler einmal aufloesen und cachen, damit
              // jeder Such-Treffer in O(1) gegen die Mailgruppen-
              // Mitgliedschaft geprueft werden kann.
              setVisibilityCacheLoading(true);
              const cache = new Set<string>();
              const audItems = audience.split(',').map(s => s.trim()).filter(Boolean);
              for (const item of audItems) {
                if (item.indexOf('@') < 0) continue;
                // Direkter User-Eintrag → in den Cache
                cache.add(item.toLowerCase());
                // Verteiler/Gruppe → Members aufloesen
                try {
                  const grp = await getGroupMembers(item).catch(() => null);
                  if (grp && grp.members) {
                    for (const m of grp.members) {
                      if (m.email) cache.add(m.email.toLowerCase());
                    }
                  }
                } catch { /* */ }
              }
              setVisibilityAudienceCache(cache);
              setVisibilityCacheLoading(false);
            }}
            type="button"
          >
            <Users size={14} /> {isDe ? 'Sichtbarkeit prüfen' : 'Check visibility'}
          </button>
          {(locationFilter || audience) && (
          <button
            className="btn btn-outline"
            style={{ fontSize: '0.8rem', padding: '6px 14px', whiteSpace: 'nowrap' }}
            onClick={async () => {
              // Resolver: Mailgruppen-Members via Graph aufloesen,
              // einzelne E-Mails direkt durchreichen, Mitglieds-Quelle
              // markieren (z.B. 'SAPALL@deloitte.com').
              setExcludeModalOpen(true);
              setExcludeResolving(true);
              const resolved: ResolvedUser[] = [];
              const seen = new Set<string>();
              // v8.9: Standorte zuerst aufloesen — alle User des
              // Standorts werden via Graph geholt (officeLocation
              // exact match, Fallback startsWith).
              const locItems = locationFilter.split(',').map(s => s.trim()).filter(Boolean);
              for (const loc of locItems) {
                try {
                  const users = await searchUsersByLocation(loc).catch(() => []);
                  for (const u of users) {
                    const k = (u.email || '').toLowerCase();
                    if (!k || seen.has(k)) continue;
                    seen.add(k);
                    resolved.push({
                      email: u.email,
                      displayName: u.displayName,
                      firstName: u.firstName || '',
                      lastName: u.lastName || '',
                      jobTitle: u.jobTitle || '',
                      location: u.location || loc,
                      source: loc,
                    });
                  }
                } catch { /* skip */ }
              }
              const audItems = audience.split(',').map(s => s.trim()).filter(Boolean);
              for (const item of audItems) {
                try {
                  if (item.indexOf('@') >= 0) {
                    // Wenn sichtbar wie eine Gruppe (z.B. SAPALL@), getGroupMembers; sonst direkt
                    const grp = await getGroupMembers(item).catch(() => null);
                    if (grp && grp.members && grp.members.length > 0) {
                      for (const m of grp.members) {
                        const k = (m.email || '').toLowerCase();
                        if (!k || seen.has(k)) continue;
                        seen.add(k);
                        resolved.push({
                          email: m.email,
                          displayName: m.displayName,
                          firstName: m.firstName || '',
                          lastName: m.lastName || '',
                          jobTitle: m.jobTitle || '',
                          location: m.location || '',
                          source: item,
                        });
                      }
                    } else {
                      // Direkter User-Eintrag
                      const k = item.toLowerCase();
                      if (!seen.has(k)) {
                        seen.add(k);
                        resolved.push({ email: item, displayName: item, firstName: '', lastName: '', jobTitle: '', location: '', source: 'direkt' });
                      }
                    }
                  }
                } catch { /* skip */ }
              }
              setExcludeResolvedUsers(resolved);
              setExcludeResolving(false);
            }}
            type="button"
          >
            <Users size={14} /> {isDe ? 'Personen ausschließen' : 'Exclude users'}
            {excludedUsers.length > 0 && (
              <span style={{ marginLeft: 6, padding: '1px 7px', background: 'var(--dex-red, #c00)', color: '#fff', borderRadius: 999, fontSize: '0.7rem', fontWeight: 700 }}>
                {excludedUsers.length}
              </span>
            )}
          </button>
          )}
          <p style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)', margin: 0, lineHeight: 1.5, flex: 1, minWidth: 200 }}>
            {isDe
              ? 'Öffnet eine Vorschau, mit der du anhand einer Testperson verifizieren kannst, ob das Event wirklich für den gewünschten Personenkreis sichtbar ist — ohne Filter ist es für alle sichtbar.'
              : 'Opens a preview where you can use a test person to verify whether the event is really visible to the intended audience — with no filter it is visible to everyone.'}
          </p>
        </div>
      )}

      {/* Massenimport-Modal — schreibt in die eigene Audience-Liste dieser Instanz. */}
      <BulkUserImportModal
        open={bulkAudienceOpen}
        onClose={() => setBulkAudienceOpen(false)}
        title="Massenimport — Sichtbarkeit"
        description={(
          <p style={{ marginTop: 0 }}>
            Trag Personen, Verteilergruppen oder Email-Adressen direkt in den
            <strong> Sichtbarkeits-Filter</strong> ein. Externe (kein
            <code style={{ margin: '0 4px' }}>@deloitte.de</code>) werden zwar
            geschrieben, sehen das Event aber NICHT — die Plattform ist DEALL-only.
          </p>
        )}
        existingEmails={audience.split(',').map(s => s.trim()).filter(Boolean)}
        searchUsers={searchUsers}
        onAdd={({ email }) => {
          // Audience ist `,`-separierte String-Liste — Email anhängen wenn
          // noch nicht drin (Doppel-Check zur Sicherheit, das Modal filtert
          // schon vor).
          setAudience(prev => {
            const list = (prev || '').split(',').map(s => s.trim()).filter(Boolean);
            if (list.indexOf(email) < 0) list.push(email);
            return list.join(', ');
          });
        }}
      />

      {/* Gruppen-Mitglieder-Modal */}
      {memberModalOpen && (
        <div
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 1100,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
          }}
          onClick={() => setMemberModalOpen(false)}
        >
          <div
            className="card"
            style={{ width: '90%', maxWidth: 560, maxHeight: '80vh', overflow: 'auto', padding: 24, background: '#fff', borderRadius: 16 }}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0 }}>
                <Users size={18} /> {memberModalGroupName}
              </h3>
              <button
                style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-600)' }}
                onClick={() => setMemberModalOpen(false)}
              >
                <X size={20} />
              </button>
            </div>
            {memberModalLoading && (
              <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', textAlign: 'center', padding: 16 }}>
                {isDe ? 'Mitglieder werden geladen…' : 'Loading members…'}
              </p>
            )}
            {memberModalError && (
              <p style={{ fontSize: '0.82rem', color: 'var(--dex-red, #c00)', lineHeight: 1.5 }}>{memberModalError}</p>
            )}
            {!memberModalLoading && !memberModalError && (
              <>
                <p style={{ fontSize: '0.82rem', color: 'var(--dex-gray-500)', marginBottom: 8 }}>
                  {memberModalMembers.length} {isDe ? 'Mitglieder' : 'members'}
                </p>
                <div style={{ maxHeight: 420, overflowY: 'auto' }}>
                  {memberModalMembers.map(m => (
                    <div key={m.email} style={{ display: 'flex', flexDirection: 'column', padding: '6px 0', borderBottom: '1px solid var(--dex-gray-100)' }}>
                      <span style={{ fontSize: '0.85rem', color: 'var(--dex-gray-800)' }}>{m.displayName}</span>
                      <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>{m.email}</span>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {/* Personen-ausschließen-Modal */}
      {excludeModalOpen && (
        <div
          style={{
            position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1200,
            display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
          }}
          onClick={() => setExcludeModalOpen(false)}
        >
          <div
            className="card"
            style={{
              width: '100%', maxWidth: 1100, maxHeight: '90vh', overflow: 'auto',
              padding: 24, borderRadius: 16, background: '#fff',
              boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
            }}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0 }}>
                <Users size={18} /> {isDe ? 'Personen ausschließen' : 'Exclude users'}
              </h3>
              <button
                style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-600)' }}
                onClick={() => setExcludeModalOpen(false)}
              >
                <X size={20} />
              </button>
            </div>

            <p style={{ margin: '0 0 12px', fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
              {isDe
                ? 'Hier kannst du einzelne Personen explizit ausschließen — sie sehen das Event NICHT, auch wenn sie über Standortfilter oder Mailverteiler eigentlich Sichtbarkeit hätten. Die Mitglieder der oben gewählten Mailverteiler werden automatisch aufgelistet (per Microsoft Graph). Standortfilter-User sind nicht direkt aus der App auflistbar — die kannst du über die Suche unten gezielt finden und ausschließen.'
                : 'Here you can explicitly exclude individuals — they will NOT see the event, even if they would otherwise have visibility via location filter or mailing list. Members of the mailing lists chosen above are listed automatically (via Microsoft Graph). Users matched only by location filter cannot be listed directly — use the search below to find and exclude them.'}
            </p>

            {/* Suchfeld — filtert die Tabelle global ueber Email/Vor-/
                Nachname/Position, und ergaenzt bei Bedarf neue User via
                Directory-Suche (z.B. wenn der Gesuchte nicht im Verteiler
                ist, aber explizit ausgeschlossen werden soll). */}
            <div style={{ marginBottom: 12 }}>
              <input
                type="text"
                value={excludeSearch}
                onChange={async e => {
                  const v = e.target.value;
                  setExcludeSearch(v);
                  // Bei Such-Eingabe immer auf Seite 0 zuruecksetzen,
                  // damit der Treffer auch sichtbar ist.
                  setExcludePage(0);
                  if (v.trim().length < 2) return;
                  try {
                    const found = await searchUsers(v.trim(), excludeIncludeIntl);
                    // Nur User, die noch nicht in der resolved-Liste stecken,
                    // anhaengen — sonst Duplikate. seen-Set baut sich durch
                    // resolved + bereits in der Suche gefundene auf.
                    setExcludeResolvedUsers(prev => {
                      const seen = new Set(prev.map(u => u.email.toLowerCase()));
                      const next = [...prev];
                      for (const u of found) {
                        const k = (u.email || '').toLowerCase();
                        if (k && !seen.has(k)) {
                          seen.add(k);
                          // displayName splitten zu first/last falls noetig (Format
                          // 'Nachname, Vorname' oder 'Vorname Nachname').
                          let fn = '';
                          let ln = '';
                          const dn = (u.displayName || '').trim();
                          if (dn.indexOf(',') >= 0) {
                            const parts = dn.split(',').map(s => s.trim());
                            ln = parts[0] || '';
                            fn = parts[1] || '';
                          } else {
                            const parts = dn.split(/\s+/);
                            fn = parts[0] || '';
                            ln = parts.slice(1).join(' ');
                          }
                          next.push({
                            email: u.email,
                            displayName: u.displayName,
                            firstName: fn,
                            lastName: ln,
                            jobTitle: u.jobTitle || '',
                            location: u.location || '',
                            source: isDe ? 'Suche' : 'search',
                          });
                        }
                      }
                      return next;
                    });
                  } catch { /* */ }
                }}
                placeholder={isDe ? 'Person suchen (Name oder E-Mail)' : 'Search person (name or email)'}
                className="form-input"
                style={{ width: '100%' }}
              />
              <InternationalSearchToggle
                checked={excludeIncludeIntl}
                onChange={async next => {
                  setExcludeIncludeIntl(next);
                  const v = excludeSearch.trim();
                  if (v.length < 2) return;
                  try {
                    const found = await searchUsers(v, next);
                    setExcludeResolvedUsers(prev => {
                      const seen = new Set(prev.map(u => u.email.toLowerCase()));
                      const acc = [...prev];
                      for (const u of found) {
                        const k = (u.email || '').toLowerCase();
                        if (k && !seen.has(k)) {
                          seen.add(k);
                          let fn = '';
                          let ln = '';
                          const dn = (u.displayName || '').trim();
                          if (dn.indexOf(',') >= 0) {
                            const parts = dn.split(',').map(s => s.trim());
                            ln = parts[0] || '';
                            fn = parts[1] || '';
                          } else {
                            const parts = dn.split(/\s+/);
                            fn = parts[0] || '';
                            ln = parts.slice(1).join(' ');
                          }
                          acc.push({
                            email: u.email,
                            displayName: u.displayName,
                            firstName: fn,
                            lastName: ln,
                            jobTitle: u.jobTitle || '',
                            location: u.location || '',
                            source: isDe ? 'Suche' : 'search',
                          });
                        }
                      }
                      return acc;
                    });
                  } catch { /* */ }
                }}
                isDe={isDe}
              />
            </div>

            {excludeResolving && (
              <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', textAlign: 'center', padding: 16 }}>
                {isDe ? 'Verteiler werden aufgelöst…' : 'Resolving distribution lists…'}
              </p>
            )}

            {!excludeResolving && excludeResolvedUsers.length === 0 && (
              <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', padding: 16, textAlign: 'center' }}>
                {isDe
                  ? 'Keine Personen aufgelöst. Nutze die Suche, um einzelne Personen hinzuzufügen.'
                  : 'No people resolved. Use the search to add individuals.'}
              </p>
            )}

            {/* v8.8: Tabelle mit Spalten Email / Nachname / Vorname / Position
                / Standort + Filter-Inputs in der Header-Zeile + Sortier-Klick.
                Damit kann der Organizer z.B. nach jobTitle 'Partner' filtern
                und alle nicht-Partner per 'Alle ausschliessen'-Aktion
                excluden. */}
            {(() => {
              const f = excludeFilters;
              // v8.12: Globaler Such-Filter (excludeSearch) wirkt
              // zusaetzlich zur Spalten-Filter-Logik. Match ueber Email,
              // Vor-/Nachname und Position. Damit landet ein Such-Treffer
              // (z.B. 'brenneisen') sofort als einziger sichtbarer Eintrag
              // in der Tabelle.
              const gs = excludeSearch.trim().toLowerCase();
              const filtered = excludeResolvedUsers.filter(u =>
                (!f.email || u.email.toLowerCase().indexOf(f.email.toLowerCase()) >= 0) &&
                (!f.lastName || u.lastName.toLowerCase().indexOf(f.lastName.toLowerCase()) >= 0) &&
                (!f.firstName || u.firstName.toLowerCase().indexOf(f.firstName.toLowerCase()) >= 0) &&
                (!f.jobTitle || u.jobTitle.toLowerCase().indexOf(f.jobTitle.toLowerCase()) >= 0) &&
                (!f.location || u.location.toLowerCase().indexOf(f.location.toLowerCase()) >= 0) &&
                (!gs ||
                  u.email.toLowerCase().indexOf(gs) >= 0 ||
                  u.lastName.toLowerCase().indexOf(gs) >= 0 ||
                  u.firstName.toLowerCase().indexOf(gs) >= 0 ||
                  u.jobTitle.toLowerCase().indexOf(gs) >= 0 ||
                  u.displayName.toLowerCase().indexOf(gs) >= 0)
              );
              const sorted = [...filtered].sort((a, b) => {
                const av = (a[excludeSortBy] || '').toLowerCase();
                const bv = (b[excludeSortBy] || '').toLowerCase();
                if (av < bv) return excludeSortDir === 'asc' ? -1 : 1;
                if (av > bv) return excludeSortDir === 'asc' ? 1 : -1;
                return 0;
              });
              const headerSort = (col: typeof excludeSortBy): void => {
                if (excludeSortBy === col) setExcludeSortDir(d => d === 'asc' ? 'desc' : 'asc');
                else { setExcludeSortBy(col); setExcludeSortDir('asc'); }
              };
              const sortIcon = (col: typeof excludeSortBy): string => excludeSortBy === col ? (excludeSortDir === 'asc' ? ' ▲' : ' ▼') : '';
              const headerStyle: React.CSSProperties = { padding: '8px 6px', textAlign: 'left', fontSize: '0.78rem', fontWeight: 700, color: 'var(--dex-gray-700)', cursor: 'pointer', userSelect: 'none', whiteSpace: 'nowrap' };
              const cellStyle: React.CSSProperties = { padding: '6px', fontSize: '0.82rem', borderBottom: '1px solid var(--dex-gray-100)' };
              const filterInputStyle: React.CSSProperties = { width: '100%', padding: '4px 6px', border: '1px solid var(--dex-gray-200)', borderRadius: 4, fontSize: '0.75rem' };
              const filterActive = !!(f.email || f.lastName || f.firstName || f.jobTitle || f.location);
              const totalPages = Math.max(1, Math.ceil(sorted.length / EXCLUDE_PAGE_SIZE));
              const safePage = Math.min(excludePage, totalPages - 1);
              const pageStart = safePage * EXCLUDE_PAGE_SIZE;
              const pageEnd = Math.min(pageStart + EXCLUDE_PAGE_SIZE, sorted.length);
              const pageItems = sorted.slice(pageStart, pageEnd);
              return (
                <>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 6, flexWrap: 'wrap', gap: 8 }}>
                    <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                      {filterActive
                        ? (isDe
                          ? <><strong>{filtered.length}</strong> von <strong>{excludeResolvedUsers.length}</strong> Personen passen zum Filter</>
                          : <><strong>{filtered.length}</strong> of <strong>{excludeResolvedUsers.length}</strong> people match the filter</>)
                        : (isDe
                          ? <><strong>{excludeResolvedUsers.length}</strong> Personen aufgelöst</>
                          : <><strong>{excludeResolvedUsers.length}</strong> people resolved</>)
                      }
                    </span>
                    {filtered.length > 0 && (
                      <div style={{ display: 'flex', gap: 6 }}>
                        <button
                          type="button"
                          onClick={() => {
                            setExcludedUsers(prev => {
                              const seen = new Set(prev.map(e => e.toLowerCase()));
                              const next = [...prev];
                              for (const u of filtered) {
                                const k = u.email.toLowerCase();
                                if (!seen.has(k)) {
                                  seen.add(k);
                                  next.push(u.email);
                                }
                              }
                              return next;
                            });
                          }}
                          style={{ fontSize: '0.72rem', padding: '4px 10px', background: 'rgba(218,41,28,0.08)', border: '1px solid var(--dex-red, #c00)', color: 'var(--dex-red, #c00)', borderRadius: 6, cursor: 'pointer', fontWeight: 600 }}
                        >
                          {isDe ? 'Alle gefilterten ausschließen' : 'Exclude all filtered'}
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            const filterEmails = new Set(filtered.map(u => u.email.toLowerCase()));
                            setExcludedUsers(prev => prev.filter(e => !filterEmails.has(e.toLowerCase())));
                          }}
                          style={{ fontSize: '0.72rem', padding: '4px 10px', background: '#fff', border: '1px solid var(--dex-gray-300)', color: 'var(--dex-gray-700)', borderRadius: 6, cursor: 'pointer' }}
                        >
                          {isDe ? 'Alle wieder einschließen' : 'Include all again'}
                        </button>
                      </div>
                    )}
                  </div>
                  <div style={{ maxHeight: 420, overflowY: 'auto', border: '1px solid var(--dex-gray-200)', borderRadius: 6 }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                      <thead style={{ position: 'sticky', top: 0, background: 'var(--dex-gray-50, #fafafa)', zIndex: 1 }}>
                        <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                          <th style={{ ...headerStyle, width: 36, cursor: 'default' }} />
                          <th style={headerStyle} onClick={() => headerSort('email')}>E-Mail{sortIcon('email')}</th>
                          <th style={headerStyle} onClick={() => headerSort('lastName')}>{isDe ? 'Nachname' : 'Last name'}{sortIcon('lastName')}</th>
                          <th style={headerStyle} onClick={() => headerSort('firstName')}>{isDe ? 'Vorname' : 'First name'}{sortIcon('firstName')}</th>
                          <th style={headerStyle} onClick={() => headerSort('jobTitle')}>{isDe ? 'Position' : 'Position'}{sortIcon('jobTitle')}</th>
                          <th style={headerStyle} onClick={() => headerSort('location')}>{isDe ? 'Standort' : 'Location'}{sortIcon('location')}</th>
                        </tr>
                        <tr style={{ borderBottom: '1px solid var(--dex-gray-200)', background: '#fff' }}>
                          <th style={{ padding: 4 }} />
                          <th style={{ padding: 4 }}><input style={filterInputStyle} value={f.email} onChange={e => { setExcludeFilters(p => ({ ...p, email: e.target.value })); setExcludePage(0); }} placeholder={isDe ? 'filtern…' : 'filter…'} /></th>
                          <th style={{ padding: 4 }}><input style={filterInputStyle} value={f.lastName} onChange={e => { setExcludeFilters(p => ({ ...p, lastName: e.target.value })); setExcludePage(0); }} placeholder={isDe ? 'filtern…' : 'filter…'} /></th>
                          <th style={{ padding: 4 }}><input style={filterInputStyle} value={f.firstName} onChange={e => { setExcludeFilters(p => ({ ...p, firstName: e.target.value })); setExcludePage(0); }} placeholder={isDe ? 'filtern…' : 'filter…'} /></th>
                          <th style={{ padding: 4 }}><input style={filterInputStyle} value={f.jobTitle} onChange={e => { setExcludeFilters(p => ({ ...p, jobTitle: e.target.value })); setExcludePage(0); }} placeholder={isDe ? 'z.B. Partner' : 'e.g. Partner'} /></th>
                          <th style={{ padding: 4 }}><input style={filterInputStyle} value={f.location} onChange={e => { setExcludeFilters(p => ({ ...p, location: e.target.value })); setExcludePage(0); }} placeholder={isDe ? 'z.B. Köln' : 'e.g. Cologne'} /></th>
                        </tr>
                      </thead>
                      <tbody>
                        {pageItems.map(u => {
                          const emailLc = u.email.toLowerCase();
                          const isExcluded = excludedUsers.some(e => e.toLowerCase() === emailLc);
                          const toggle = (): void => {
                            if (isExcluded) setExcludedUsers(prev => prev.filter(e => e.toLowerCase() !== emailLc));
                            else setExcludedUsers(prev => [...prev, u.email]);
                          };
                          return (
                            <tr
                              key={u.email}
                              onClick={toggle}
                              style={{
                                cursor: 'pointer',
                                background: isExcluded ? 'rgba(218,41,28,0.06)' : 'transparent',
                              }}
                            >
                              <td style={{ ...cellStyle, textAlign: 'center' }}>
                                <input
                                  type="checkbox"
                                  checked={!isExcluded}
                                  onChange={toggle}
                                  onClick={e => e.stopPropagation()}
                                  style={{ accentColor: 'var(--dex-green)', width: 14, height: 14, cursor: 'pointer' }}
                                />
                              </td>
                              <td style={{ ...cellStyle, color: isExcluded ? 'var(--dex-red, #c00)' : 'var(--dex-gray-700)', fontFamily: 'monospace', fontSize: '0.78rem' }}>{u.email}</td>
                              <td style={{ ...cellStyle, fontWeight: 500, color: isExcluded ? 'var(--dex-red, #c00)' : 'var(--dex-gray-800)' }}>{u.lastName || '—'}</td>
                              <td style={{ ...cellStyle, color: isExcluded ? 'var(--dex-red, #c00)' : 'var(--dex-gray-700)' }}>{u.firstName || '—'}</td>
                              <td style={{ ...cellStyle, color: 'var(--dex-gray-600)', fontSize: '0.78rem' }}>{u.jobTitle || '—'}</td>
                              <td style={{ ...cellStyle, color: 'var(--dex-gray-600)', fontSize: '0.78rem' }}>{u.location || '—'}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                  {/* v8.9: Pagination — 200 pro Seite. Wird nur angezeigt
                      wenn es mehr als eine Seite gibt. */}
                  {totalPages > 1 && (
                    <div style={{
                      display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                      marginTop: 8, padding: '6px 4px', flexWrap: 'wrap', gap: 8,
                    }}>
                      <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)' }}>
                        {isDe
                          ? <>Zeige <strong>{pageStart + 1}</strong>–<strong>{pageEnd}</strong> von <strong>{sorted.length}</strong></>
                          : <>Showing <strong>{pageStart + 1}</strong>–<strong>{pageEnd}</strong> of <strong>{sorted.length}</strong></>}
                      </span>
                      <div style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                        <button type="button" disabled={safePage === 0} onClick={() => setExcludePage(0)} style={{ padding: '4px 8px', fontSize: '0.78rem', border: '1px solid var(--dex-gray-300)', background: '#fff', borderRadius: 4, cursor: safePage === 0 ? 'not-allowed' : 'pointer', opacity: safePage === 0 ? 0.4 : 1 }}>«</button>
                        <button type="button" disabled={safePage === 0} onClick={() => setExcludePage(p => Math.max(0, p - 1))} style={{ padding: '4px 8px', fontSize: '0.78rem', border: '1px solid var(--dex-gray-300)', background: '#fff', borderRadius: 4, cursor: safePage === 0 ? 'not-allowed' : 'pointer', opacity: safePage === 0 ? 0.4 : 1 }}>‹ {isDe ? 'Zurück' : 'Prev'}</button>
                        <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-700)', padding: '0 8px' }}>
                          {isDe ? 'Seite' : 'Page'} <strong>{safePage + 1}</strong> / {totalPages}
                        </span>
                        <button type="button" disabled={safePage >= totalPages - 1} onClick={() => setExcludePage(p => Math.min(totalPages - 1, p + 1))} style={{ padding: '4px 8px', fontSize: '0.78rem', border: '1px solid var(--dex-gray-300)', background: '#fff', borderRadius: 4, cursor: safePage >= totalPages - 1 ? 'not-allowed' : 'pointer', opacity: safePage >= totalPages - 1 ? 0.4 : 1 }}>{isDe ? 'Weiter' : 'Next'} ›</button>
                        <button type="button" disabled={safePage >= totalPages - 1} onClick={() => setExcludePage(totalPages - 1)} style={{ padding: '4px 8px', fontSize: '0.78rem', border: '1px solid var(--dex-gray-300)', background: '#fff', borderRadius: 4, cursor: safePage >= totalPages - 1 ? 'not-allowed' : 'pointer', opacity: safePage >= totalPages - 1 ? 0.4 : 1 }}>»</button>
                      </div>
                    </div>
                  )}
                </>
              );
            })()}

            {/* Bereits ausgeschlossene User die NICHT in der resolved-Liste sind
                (z.B. weil sie nur ueber Standortfilter sichtbar waeren und
                ueber die Suche manuell ausgeschlossen wurden in einer
                frueheren Session) — separat darstellen. */}
            {excludedUsers.filter(e => !excludeResolvedUsers.some(u => u.email.toLowerCase() === e.toLowerCase())).length > 0 && (
              <div style={{ marginTop: 16, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
                <h4 style={{ margin: '0 0 8px', fontSize: '0.85rem', color: 'var(--dex-gray-700)' }}>
                  {isDe ? 'Weitere ausgeschlossene Personen' : 'Other excluded users'}
                </h4>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                  {excludedUsers
                    .filter(e => !excludeResolvedUsers.some(u => u.email.toLowerCase() === e.toLowerCase()))
                    .map(e => (
                      <span key={e} style={{
                        display: 'inline-flex', alignItems: 'center', gap: 6,
                        padding: '3px 4px 3px 10px',
                        background: 'rgba(218,41,28,0.08)',
                        border: '1px solid var(--dex-red, #c00)',
                        color: 'var(--dex-red, #c00)',
                        borderRadius: 999, fontSize: '0.78rem',
                      }}>
                        {e}
                        <button
                          type="button"
                          onClick={() => setExcludedUsers(prev => prev.filter(x => x.toLowerCase() !== e.toLowerCase()))}
                          style={{
                            width: 18, height: 18, borderRadius: '50%',
                            border: 'none', background: 'rgba(218,41,28,0.2)',
                            color: 'var(--dex-red, #c00)', cursor: 'pointer',
                            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                            fontSize: '0.85rem', lineHeight: 1,
                          }}
                          title={isDe ? 'Ausschluss aufheben' : 'Remove exclusion'}
                        >×</button>
                      </span>
                    ))}
                </div>
              </div>
            )}

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 20, paddingTop: 12, borderTop: '1px solid var(--dex-gray-200)' }}>
              <button
                type="button"
                className="btn btn-primary"
                onClick={() => setExcludeModalOpen(false)}
              >
                {isDe ? 'Fertig' : 'Done'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Sichtbarkeit-prüfen-Modal */}
      {showEmailModal && (
        <div style={{
          position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 1000,
          display: 'flex', alignItems: 'center', justifyContent: 'center',
        }} onClick={() => setShowEmailModal(false)}>
          <div
            className="card"
            style={{ width: '90%', maxWidth: 600, maxHeight: '80vh', overflow: 'auto', padding: 24 }}
            onClick={e => e.stopPropagation()}
          >
            <div className="flex-between mb-16">
              <h3 style={{ margin: 0 }}>
                <Users size={18} /> {isDe ? 'Sichtbarkeit prüfen' : 'Check visibility'}
              </h3>
              <button
                style={{ background: 'none', border: 'none', fontSize: '1.2rem', cursor: 'pointer', color: 'var(--dex-gray-600)' }}
                onClick={() => setShowEmailModal(false)}
              >
                <X size={20} />
              </button>
            </div>

            <p style={{ margin: '0 0 12px', fontSize: '0.82rem', color: 'var(--dex-gray-600)', lineHeight: 1.55 }}>
              {isDe
                ? 'Hier kannst du verifizieren, ob die kombinierte Sichtbarkeit (Standortfilter + Mailverteiler / einzelne User + Verknüpfung) wirklich zu der Person passt, die das Event sehen soll. Tippe Name oder E-Mail einer Testperson ein und klick „Suchen" — die Tabelle darunter zeigt, ob sie das Event in ihrer Übersicht sieht und woher die Sichtbarkeit kommt (Standort-Match oder Mitgliedschaft in einem Verteiler).'
                : 'Use this to verify whether the combined visibility (location filter + mailing lists / individual users + the AND/OR mode) actually matches the person you want to reach. Type a test person\'s name or email and click "Search" — the table below shows whether they can see the event and where the visibility comes from (location match or membership in a list).'}
            </p>

            <div style={{ marginBottom: 16, padding: '10px 14px', background: 'var(--dex-gray-100)', borderRadius: 'var(--dex-radius)', fontSize: '0.85rem' }}>
              <div style={{ marginBottom: 6 }}>
                <strong>{isDe ? 'Standortfilter:' : 'Location filter:'}</strong>{' '}
                {locationFilter ? locationFilter.split(',').map(s => s.trim()).map(s => (
                  <span key={s} className="badge badge-green" style={{ marginRight: 6 }}>{s}</span>
                )) : <span style={{ color: 'var(--dex-gray-400)' }}>{isDe ? 'Keine' : 'None'}</span>}
              </div>
              <div style={{ marginBottom: 6 }}>
                <strong>{isDe ? 'Mailverteiler / einzelne User:' : 'Mailing lists / individual users:'}</strong>{' '}
                {audience ? audience.split(',').map(s => s.trim()).map(s => (
                  <span key={s} className="badge badge-orange" style={{ marginRight: 6 }}>{s}</span>
                )) : <span style={{ color: 'var(--dex-gray-400)' }}>{isDe ? 'Keine' : 'None'}</span>}
              </div>
              {locationFilter && audience && (
                <div>
                  <strong>{isDe ? 'Verknüpfung:' : 'Combination:'}</strong>{' '}
                  <span className={`badge ${filterMode === 'AND' ? 'badge-red' : 'badge-green'}`}>
                    {filterMode === 'AND'
                      ? (isDe ? 'UND (beide müssen zutreffen)' : 'AND (both must match)')
                      : (isDe ? 'ODER (eines reicht)' : 'OR (one is enough)')}
                  </span>
                </div>
              )}
            </div>

            <div className="form-group">
              <label className="form-label">{isDe ? 'User suchen (Name oder E-Mail)' : 'Search user (name or email)'}</label>
              <div style={{ display: 'flex', gap: 8 }}>
                <input
                  className="form-input"
                  value={emailSearch}
                  onChange={e => setEmailSearch(e.target.value)}
                  placeholder="z.B. Max Mustermann oder mmustermann@"
                  onKeyDown={async e => {
                    if (e.key === 'Enter' && emailSearch.length >= 2) {
                      setIsSearchingEmails(true);
                      const results = await searchUsers(emailSearch, emailSearchIncludeIntl);
                      setEmailSearchResults(results);
                      setIsSearchingEmails(false);
                    }
                  }}
                />
                <button
                  className="btn btn-primary"
                  style={{ whiteSpace: 'nowrap' }}
                  disabled={emailSearch.length < 2 || isSearchingEmails}
                  onClick={async () => {
                    setIsSearchingEmails(true);
                    const results = await searchUsers(emailSearch, emailSearchIncludeIntl);
                    setEmailSearchResults(results);
                    setIsSearchingEmails(false);
                  }}
                >
                  {isSearchingEmails ? '...' : (isDe ? 'Suchen' : 'Search')}
                </button>
              </div>
              <InternationalSearchToggle
                checked={emailSearchIncludeIntl}
                onChange={async next => {
                  setEmailSearchIncludeIntl(next);
                  if (emailSearch.length >= 2) {
                    setIsSearchingEmails(true);
                    try {
                      const results = await searchUsers(emailSearch, next);
                      setEmailSearchResults(results);
                    } catch { /* */ }
                    setIsSearchingEmails(false);
                  }
                }}
                isDe={isDe}
              />
            </div>

            {emailSearchResults.length > 0 && (
              <div style={{ marginTop: 12 }}>
                <p style={{ fontSize: '0.85rem', color: 'var(--dex-gray-500)', marginBottom: 8 }}>
                  {emailSearchResults.length} Ergebnis{emailSearchResults.length !== 1 ? 'se' : ''}:
                </p>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                  <thead>
                    <tr style={{ borderBottom: '2px solid var(--dex-gray-200)' }}>
                      <th style={{ textAlign: 'left', padding: 6 }}>Name</th>
                      <th style={{ textAlign: 'left', padding: 6 }}>E-Mail</th>
                      <th style={{ textAlign: 'left', padding: 6 }}>Standort</th>
                      <th style={{ textAlign: 'center', padding: 6 }}>{isDe ? 'Sichtbar?' : 'Visible?'}</th>
                      <th style={{ textAlign: 'left', padding: 6 }}>{isDe ? 'Begründung' : 'Reason'}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {visibilityCacheLoading && (
                      <tr><td colSpan={5} style={{ padding: 12, textAlign: 'center', fontSize: '0.82rem', color: 'var(--dex-gray-500)' }}>
                        {isDe ? 'Verteiler werden aufgelöst…' : 'Resolving distribution lists…'}
                      </td></tr>
                    )}
                    {!visibilityCacheLoading && emailSearchResults.map(u => {
                      // v8.9: Volle isEventVisibleForUser-Logik nachbauen
                      // — Exclude > Standort + Audience + UND/ODER.
                      const emailLc = (u.email || '').toLowerCase();
                      const loc = (u.location || '').toLowerCase();
                      const locationFilters = locationFilter ? locationFilter.split(',').map(s => s.trim().toLowerCase()).filter(Boolean) : [];
                      const hasLocFilter = locationFilters.length > 0 && locationFilters.indexOf('all') < 0;
                      const hasAudFilter = visibilityAudienceCache.size > 0;
                      const isExcluded = excludedUsers.some(e => e.toLowerCase() === emailLc);
                      const matchedLoc = hasLocFilter
                        ? locationFilters.find(f => {
                            const norm = f.replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ä/g, 'ae');
                            return loc.indexOf(f) >= 0 || loc.indexOf(norm) >= 0;
                          })
                        : null;
                      const locMatch = !hasLocFilter || !!matchedLoc;
                      const audMatch = !hasAudFilter || visibilityAudienceCache.has(emailLc);
                      let visible: boolean;
                      let reasonParts: string[] = [];
                      if (!hasLocFilter && !hasAudFilter) {
                        visible = true;
                        reasonParts.push(isDe ? 'Keine Filter gesetzt — für alle sichtbar' : 'No filters set — visible to everyone');
                      } else if (filterMode === 'OR') {
                        visible = (hasLocFilter && locMatch) || (hasAudFilter && audMatch);
                        if (locMatch && hasLocFilter) reasonParts.push(isDe ? `Standort-Match (${matchedLoc})` : `location match (${matchedLoc})`);
                        if (audMatch && hasAudFilter) reasonParts.push(isDe ? 'in Mailverteiler/User-Liste' : 'in mailing list / user');
                        if (!visible) reasonParts.push(isDe ? 'kein Filter passt' : 'no filter matches');
                      } else {
                        // AND
                        visible = (!hasLocFilter || locMatch) && (!hasAudFilter || audMatch);
                        if (hasLocFilter) reasonParts.push(locMatch ? (isDe ? `Standort ✓ (${matchedLoc})` : `location ✓ (${matchedLoc})`) : (isDe ? 'Standort ✗' : 'location ✗'));
                        if (hasAudFilter) reasonParts.push(audMatch ? (isDe ? 'Verteiler ✓' : 'mailing list ✓') : (isDe ? 'Verteiler ✗' : 'mailing list ✗'));
                      }
                      // Exclude hat Vorrang
                      if (isExcluded && visible) {
                        visible = false;
                        reasonParts = [(isDe ? `wäre sichtbar (${reasonParts.join(', ')}), aber explizit ausgeschlossen` : `would be visible (${reasonParts.join(', ')}), but explicitly excluded`)];
                      } else if (isExcluded) {
                        reasonParts.push(isDe ? '+ ausgeschlossen' : '+ excluded');
                      }
                      return (
                        <tr key={u.email} style={{ borderBottom: '1px solid var(--dex-gray-100)' }}>
                          <td style={{ padding: 6 }}>{u.displayName}</td>
                          <td style={{ padding: 6, color: 'var(--dex-gray-600)' }}>{u.email}</td>
                          <td style={{ padding: 6, color: 'var(--dex-gray-500)' }}>{u.location || '-'}</td>
                          <td style={{ padding: 6, textAlign: 'center' }}>
                            {visible
                              ? <span style={{ color: '#22c55e', fontWeight: 700, fontSize: '1.1rem' }}>&#10003;</span>
                              : <span style={{ color: '#ef4444', fontWeight: 700, fontSize: '1.1rem' }}>&#10007;</span>}
                          </td>
                          <td style={{ padding: 6, color: 'var(--dex-gray-600)', fontSize: '0.78rem' }}>
                            {reasonParts.join(', ')}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      )}
    </>
  );
}
