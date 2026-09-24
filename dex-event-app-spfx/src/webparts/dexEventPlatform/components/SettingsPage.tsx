/**
 * Settings-Seite
 * Zeigt User-Infos, Rollenmanagement (Admin), Admin-Aktionen und Download-Link.
 */

import * as React from 'react';
import { useNavigation } from '../context/NavigationContext';
import { useCurrentUser } from '../context/UserContext';
import { useRoles } from '../context/RoleContext';
import { useEvents } from '../context/EventContext';
import { useLanguage } from '../context/LanguageContext';
// v20.4: moderne Confirm-Modals statt window.confirm.
import { useDialog } from '../context/DialogContext';
import { UserRole } from '../types';
import { Plus, Trash2, X, Mail } from './Icons';
import Modal from './Modal';
import { EventService } from '../services/EventService';
import type { PersonRenameResult, PersonRenameBereich as PersonRenameBereichLike } from '../services/events/personRename';
import InternationalSearchToggle from './InternationalSearchToggle';
import { PersonContactHover } from './PersonContactHover';

/**
 * v26.34: „Coordinated Events"-Zelle — zeigt standardmäßig nur die ANZAHL der
 * betreuten Events als klickbaren Chip; ein Klick klappt die Titel-Badges aus.
 * Hält die Rollen-Tabelle bei Organizern mit vielen Events kompakt.
 * Modul-Level-Komponente (stabile Identität) → Aufklapp-Zustand bleibt pro Zeile
 * erhalten, auch wenn die Elternkomponente neu rendert.
 */
function CoordinatedEventsCell(props: { titles: string[]; isDe: boolean }): React.ReactElement {
  const { titles, isDe } = props;
  const [open, setOpen] = React.useState(false);
  if (titles.length === 0) return <span style={{ color: 'var(--dex-gray-300)' }}>—</span>;
  const word = isDe ? (titles.length === 1 ? 'Event' : 'Events') : (titles.length === 1 ? 'event' : 'events');
  return (
    <div>
      <button
        type="button"
        onClick={() => setOpen(o => !o)}
        title={open ? (isDe ? 'Zuklappen' : 'Collapse') : (isDe ? 'Events anzeigen' : 'Show events')}
        style={{ display: 'inline-flex', alignItems: 'center', gap: 6, padding: '2px 10px', borderRadius: 999, border: '1px solid rgba(134,188,37,0.35)', background: 'rgba(134,188,37,0.14)', color: 'var(--dex-green-dark)', fontSize: '0.74rem', fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap' }}
      >
        {titles.length} {word}
        <span aria-hidden="true" style={{ fontSize: '0.7rem' }}>{open ? '▾' : '▸'}</span>
      </button>
      {open && (
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, marginTop: 6 }}>
          {titles.map((t2, idx) => (
            <span key={idx} style={{ display: 'inline-block', padding: '2px 8px', borderRadius: 10, background: 'rgba(134,188,37,0.14)', color: 'var(--dex-green-dark)', fontSize: '0.72rem', fontWeight: 600 }}>{t2}</span>
          ))}
        </div>
      )}
    </div>
  );
}

export default function SettingsPage(): React.ReactElement {
  const { navigate } = useNavigation();
  const { currentUser } = useCurrentUser();
  const {
    roles, isAdmin, originalIsAdmin, canCreateEvents, rolesReadStatus,
    addRole, updateRole, setPowerUser, removeRole, hadRoleRightsIssue, isRolesLoading, siteUrl, searchUsers, searchUser,
    auditRolesAccess, getBasicProfiles, lastRoleRightsMissing, lastRightsAudit, refreshRoles,
  } = useRoles();
  const { events, sendOrganizerOnboarding, repairAllOrganizerPermissions } = useEvents();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  // v20.4: App-Modal statt window.confirm.
  const { confirmDialog, showAlert } = useDialog();
  // v13.0: Settings/Rollenverwaltung ist Admin-only. Vorher konnte ein
  // Demo-User die Seite öffnen — Admin-Controls waren zwar versteckt,
  // aber der Seitenzugriff selbst war frei. Wir nutzen originalIsAdmin
  // damit der Admin-im-Demo-Modus seine eigene Einstellungen weiterhin
  // testen kann.
  React.useEffect(() => {
    if (!originalIsAdmin) navigate('start', undefined, undefined, { replace: true }); // v31.59: kein Rücksprungziel
  }, [originalIsAdmin, navigate]);

  /**
   * Map: organizer-email-lowercase -> Liste von Event-Titeln, die diese Person koordiniert.
   * Nutzt den gleichen Substring-Match wie die AdminPage Filter-Logik:
   * Organizer-Eintrag enthält entweder den Vor+Nachnamen oder nur den Nachnamen.
   */
  const organizerEventMap = React.useMemo<Record<string, string[]>>(() => {
    const map: Record<string, string[]> = {};
    for (const role of roles) {
      // v30.60: F&A koordiniert seit dem Rollen-Zuschnitt eigene Events mit.
      if (role.role !== 'Organizer' && role.role !== 'F&A' && role.role !== 'Admin' && role.role !== 'IT-Admin') continue;
      const emailLc = (role.userEmail || '').toLowerCase();
      if (!emailLc) continue;
      const matched: string[] = [];
      for (const evt of events) {
        // v30.15: Nur MUTTER-Events zählen — Sub-Events/Kalender-Tage tragen
        // seit v29.20 dieselbe OrganizerEmail, eine Office-Tage-Reihe mit 22
        // Terminen blähte die Zelle sonst auf „22 Events" auf.
        if (evt.parentEventId) continue;
        // v6.20: strikt per E-Mail — kein Namens-Substring-Match mehr.
        // Der alte Heuristik-Match per Nachname hat bei häufigen Nachnamen
        // False-Positives produziert (z.B. Assistentin mit gleichem Nachnamen
        // wie ein Organizer bekam Zugriff zugeordnet).
        if (evt.organizerEmails && evt.organizerEmails.some(e => e.toLowerCase() === emailLc)) {
          matched.push(evt.title);
        }
      }
      if (matched.length > 0) map[emailLc] = matched;
    }
    return map;
  }, [roles, events]);

  /**
   * v10.16: Per-Event-Co-Organizers — Personen die im Wizard-Picker eines Events
   * als Organizer eingetragen sind, aber KEINEN globalen Eintrag in DEX_Roles
   * haben. Sie haben für DAS jeweilige Event vollen Zugriff (über die Event-
   * Subsite-Permissions), tauchen aber nicht in der regulären roles-Tabelle auf.
   * Damit der Admin sieht „wer hat per Event Zugriff ohne globalen Status",
   * werden sie hier eingesammelt: pro Email die Liste der Events + ein
   * Display-Name (aus dem ersten Event-Match abgeleitet).
   */
  const coOrganizersList = React.useMemo<Array<{ email: string; name: string; events: string[] }>>(() => {
    const knownEmails = new Set(roles.map(r => (r.userEmail || '').toLowerCase()));
    const accumulator: Record<string, { name: string; events: string[] }> = {};
    for (const evt of events) {
      // v30.15: nur Mutter-Events — s. organizerEventMap (22-Kalendertage-Falle).
      if (evt.parentEventId) continue;
      const orgEmails = evt.organizerEmails || [];
      const orgNames = evt.organizers || [];
      for (let i = 0; i < orgEmails.length; i++) {
        const emailLc = (orgEmails[i] || '').toLowerCase();
        if (!emailLc) continue;
        if (knownEmails.has(emailLc)) continue;  // schon in der Roles-Tabelle
        const name = orgNames[i] || emailLc;
        if (!accumulator[emailLc]) accumulator[emailLc] = { name, events: [] };
        accumulator[emailLc].events.push(evt.title);
      }
    }
    // v24.84: alphabetisch nach Name sortieren (vorher nach E-Mail-Schlüssel).
    return Object.keys(accumulator).map(emailLc => ({
      email: emailLc,
      name: accumulator[emailLc].name,
      events: accumulator[emailLc].events,
    })).sort((a, b) => (a.name || '').localeCompare(b.name || '', 'de'));
  }, [roles, events]);
  // v19.26: Die Per-Event-Organizer-Namen stammen aus dem index-basierten
  // Pairing von Organizer-Namen ↔ -E-Mails im Event. Driftet dieses Pairing
  // (z.B. weil ein Feld einen leeren Slot hatte und unabhängig gefiltert wurde),
  // steht der Name neben der falschen E-Mail — und das Foto (lädt per E-Mail)
  // zeigt die falsche Person. Wir lösen den Anzeige-Namen daher zuverlässig über
  // die E-Mail auf (gleicher Schlüssel wie das Foto); Fallback bleibt der
  // Index-Name. So passen Name + Foto immer zusammen.
  const [resolvedOrgNames, setResolvedOrgNames] = React.useState<Record<string, string>>({});
  const orgNameAttemptedRef = React.useRef<Set<string>>(new Set());
  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      for (const co of coOrganizersList) {
        if (orgNameAttemptedRef.current.has(co.email)) continue;
        orgNameAttemptedRef.current.add(co.email);
        try {
          // includeInternational=true → deckt @deloitte.de UND @deloitte.com ab.
          const results = await searchUsers(co.email, true);
          const match = results.find(r => (r.email || '').toLowerCase() === co.email);
          if (match && match.displayName && !cancelled) {
            setResolvedOrgNames(prev => ({ ...prev, [co.email]: match.displayName }));
          }
        } catch { /* Fallback: Index-Name aus co.name */ }
      }
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [coOrganizersList]);
  // Formular-State für neue Rolle
  const [newEmail, setNewEmail] = React.useState('');
  const [newName, setNewName] = React.useState('');
  const [newRole, setNewRole] = React.useState<UserRole>('Organizer');
  const [newLocation, setNewLocation] = React.useState('');
  const [isAdding, setIsAdding] = React.useState(false);
  const [showAddForm, setShowAddForm] = React.useState(false);
  const [statusMsg, setStatusMsg] = React.useState('');
  // v31.85: Rechte auf den Teilnehmerlisten ALLER Events (Organizer,
  // Co-Organizer, Check-in-Team) — aus dem Admin Hub hierher verschoben
  // (Nutzer-Ansage 23.09.2026: „kann Organizer-Berechtigungen nicht besser in
  // die Rollenverwaltung?"). Hier stehen die Rollen, hier gehört die Frage
  // „wer darf welche Liste sehen" hin; im Hub war es eine von zwölf Kacheln.
  const [listPerm, setListPerm] = React.useState<{ running: boolean; done: number; total: number; label: string; result: string; failed: boolean }>({ running: false, done: 0, total: 0, label: '', result: '', failed: false });
  const runListPermissions = async (): Promise<void> => {
    if (listPerm.running) return;
    if (!(await confirmDialog(
      isDe
        ? 'Für ALLE Events prüfen, ob jeder Organizer, Co-Organizer und jedes Check-in-Team-Mitglied Zugriff auf die Teilnehmerlisten hat — auf dem Haupt-Event UND auf jedem Termin? Fehlende Rechte werden ergänzt. Es wird nichts entzogen und nichts gelöscht. Je nach Anzahl der Events kann das einige Minuten dauern.'
        : 'Check for ALL events whether every organizer, co-organizer and check-in team member has access to the participant lists — on the main event AND on every date? Missing permissions are added. Nothing is revoked and nothing is deleted. This may take a few minutes depending on the number of events.',
      { confirmLabel: isDe ? 'Jetzt prüfen' : 'Check now' }
    ))) return;
    setListPerm({ running: true, done: 0, total: 0, label: '', result: '', failed: false });
    try {
      const r = await repairAllOrganizerPermissions((done, total, label) => setListPerm(prev => ({ ...prev, done, total, label })));
      // Die Zahl der Zuweisungen ist bewusst NICHT als „so viele waren kaputt"
      // formuliert: SharePoint meldet bei addroleassignment nicht, ob das Recht
      // neu ist. Deshalb steht dort, was getan wurde, nicht was gefehlt hat.
      // v31.89: `unresolved` mischt drei Dinge (maintenance.ts): E-Mail-Adressen
      // ohne SharePoint-Konto, Subsites mit HTTP 404 (Teilnehmerliste gelöscht
      // oder recycelt — unkritisch, das Event ist vorbei) und echte Ablehnungen.
      // Der Screenshot vom 24.09.2026 zeigte 13 Subsite-URLs als „Adressen
      // konnten nicht zugeordnet werden — meist ehemalige Kolleg:innen".
      const siteName = (s: string): string => s.replace(/\s*\(.*$/, '').replace(/\/+$/, '').split('/').pop() || s;
      const weg = r.unresolved.filter(u => /HTTP 404\)/.test(u));
      const abgelehnt = r.unresolved.filter(u => /\(.*HTTP \d+\)/.test(u) && !/HTTP 404\)/.test(u));
      const adressen = r.unresolved.filter(u => !/\(.*HTTP \d+\)/.test(u));
      const wegSites = Array.from(new Set(weg.map(siteName)));
      const teile: string[] = [];
      if (adressen.length) teile.push(isDe
        ? `${adressen.length} Adresse(n) ohne SharePoint-Konto (meist ehemalige Kolleg:innen): ${adressen.join(', ')}.`
        : `${adressen.length} address(es) without a SharePoint account (usually former colleagues): ${adressen.join(', ')}.`);
      if (wegSites.length) teile.push(isDe
        ? `${wegSites.length} Teilnehmerliste(n) gibt es nicht mehr (Subsite gelöscht oder recycelt — unkritisch): ${wegSites.join(', ')}.`
        : `${wegSites.length} participant list(s) no longer exist (subsite deleted or recycled — harmless): ${wegSites.join(', ')}.`);
      if (abgelehnt.length) teile.push(isDe
        ? `${abgelehnt.length} Zuweisung(en) hat SharePoint abgelehnt: ${abgelehnt.join(', ')}.`
        : `${abgelehnt.length} assignment(s) were rejected by SharePoint: ${abgelehnt.join(', ')}.`);
      const un = teile.length ? ' ' + teile.join(' ') : '';
      const echteFehler = abgelehnt.length;
      const msg = isDe
        ? `Fertig: ${r.trees} Event(s) mit insgesamt ${r.sites} Teilnehmerliste(n) durchlaufen, ${r.grants} Zuweisung(en) gesetzt${echteFehler ? `, ${echteFehler} abgelehnt` : ''}.${un}`
        : `Done: ${r.trees} event(s) with ${r.sites} participant list(s) processed, ${r.grants} assignment(s) applied${echteFehler ? `, ${echteFehler} rejected` : ''}.${un}`;
      setListPerm({ running: false, done: 0, total: 0, label: '', result: msg, failed: echteFehler > 0 });
    } catch {
      setListPerm({ running: false, done: 0, total: 0, label: '', result: isDe ? 'Berechtigungs-Prüfung fehlgeschlagen.' : 'Permission check failed.', failed: true });
    }
  };

  // v31.91: Person umbenennen (Heirat, neue Adresse) — alte gegen neue
  // Adresse/Namen an allen Stellen tauschen (services/events/personRename).
  // Anlass: Inga Fuhr → Guenther, 24.09.2026; das Rechte-Audit meldete
  // „gleiche Rechte, andere Schreibweise". Vorschau (Trockenlauf) vor dem
  // Schreiben — die Zahl je Stelle sagt, was passieren wird.
  const [aliasHints, setAliasHints] = React.useState<Array<{ name?: string; email: string; spEmail: string }>>([]);
  const [renameForm, setRenameForm] = React.useState<{ oldEmail: string; newEmail: string; displayName: string; first: string; last: string }>({ oldEmail: '', newEmail: '', displayName: '', first: '', last: '' });
  const [renameBusy, setRenameBusy] = React.useState<'' | 'plan' | 'run'>('');
  const [renameLabel, setRenameLabel] = React.useState('');
  const [renamePlan, setRenamePlan] = React.useState<PersonRenameResult | null>(null);
  const [renameResult, setRenameResult] = React.useState<PersonRenameResult | null>(null);
  const renameSvc = React.useMemo((): EventService | null => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ctx = (window as any).__dexSpfxContext;
    return ctx ? new EventService(ctx) : null;
  }, []);
  const renameArgs = (): { oldEmail: string; newEmail: string; newDisplayName: string; newFirstName: string; newLastName: string } => ({
    oldEmail: renameForm.oldEmail.trim(), newEmail: renameForm.newEmail.trim(),
    newDisplayName: renameForm.displayName.trim(), newFirstName: renameForm.first.trim(), newLastName: renameForm.last.trim(),
  });
  const renameGueltig = (): boolean => {
    const a = renameArgs();
    return a.oldEmail.indexOf('@') > 0 && a.newEmail.indexOf('@') > 0 && a.oldEmail.toLowerCase() !== a.newEmail.toLowerCase();
  };
  const runRenamePlan = async (): Promise<void> => {
    if (!renameSvc || renameBusy || !renameGueltig()) return;
    setRenameBusy('plan'); setRenamePlan(null); setRenameResult(null);
    try { setRenamePlan(await renameSvc.renamePerson(renameArgs(), { dryRun: true }, setRenameLabel)); }
    catch { showAlert(isDe ? 'Vorschau fehlgeschlagen.' : 'Preview failed.', { variant: 'error' }); }
    setRenameBusy(''); setRenameLabel('');
  };
  const runRename = async (): Promise<void> => {
    if (!renameSvc || renameBusy || !renamePlan) return;
    const a = renameArgs();
    const summe = renamePlan.roles.hits + renamePlan.eventFields.hits + renamePlan.eventPiggybacks.hits + renamePlan.registry.hits + renamePlan.registrations.hits + renamePlan.registeredBy.hits;
    const ok = await confirmDialog(
      isDe
        ? `${a.oldEmail} → ${a.newEmail}${a.newDisplayName ? ` (${a.newDisplayName})` : ''}: ${summe} Stelle(n) jetzt umschreiben?\n\nRollen ${renamePlan.roles.hits}, Event-Felder ${renamePlan.eventFields.hits}, Teams ${renamePlan.eventPiggybacks.hits}, Register ${renamePlan.registry.hits}, Anmeldungen ${renamePlan.registrations.hits}, Registriert-von ${renamePlan.registeredBy.hits}. Es geht keine Mail raus. SharePoint-Rechte bleiben — das Konto ist dasselbe.`
        : `${a.oldEmail} → ${a.newEmail}${a.newDisplayName ? ` (${a.newDisplayName})` : ''}: rewrite ${summe} place(s) now?\n\nRoles ${renamePlan.roles.hits}, event fields ${renamePlan.eventFields.hits}, teams ${renamePlan.eventPiggybacks.hits}, registry ${renamePlan.registry.hits}, registrations ${renamePlan.registrations.hits}, registered-by ${renamePlan.registeredBy.hits}. No email is sent. SharePoint rights stay — it is the same account.`,
      { danger: true, confirmLabel: isDe ? 'Umbenennen' : 'Rename' },
    );
    if (!ok) return;
    setRenameBusy('run');
    try {
      const r = await renameSvc.renamePerson(a, { dryRun: false }, setRenameLabel);
      setRenameResult(r);
      setRenamePlan(null);
      try { await refreshRoles(); } catch { /* */ }
    } catch { showAlert(isDe ? 'Umbenennung fehlgeschlagen.' : 'Rename failed.', { variant: 'error' }); }
    setRenameBusy(''); setRenameLabel('');
  };

  // v30.81: Leserechte auf DEX_Roles prüfen — Fortschritt und Ergebnis.
  const [accessAudit, setAccessAudit] = React.useState<{ running: boolean; done: number; total: number; result: string }>({ running: false, done: 0, total: 0, result: '' });
  const runAccessAudit = async (): Promise<void> => {
    if (accessAudit.running) return;
    setAccessAudit({ running: true, done: 0, total: 0, result: '' });
    try {
      const r = await auditRolesAccess((done, total) => setAccessAudit(prev => ({ ...prev, done, total })));
      let msg: string;
      if (r.readFailed) {
        msg = isDe
          ? 'Die Berechtigungen der Rollenliste konnten nicht gelesen werden — bitte später erneut versuchen.'
          : 'The permissions of the roles list could not be read — please try again later.';
      } else if (r.missing.length === 0) {
        msg = isDe
          ? `${r.checked} Rollen-Einträge geprüft — alle haben ihre drei Rechte (Rollenliste lesen, Event-Liste, Site).`
          : `${r.checked} role entries checked — all have their three rights (read roles list, events list, site).`;
      } else {
        const fixed = r.fixed.length ? (isDe ? ` Nachgesetzt: ${r.fixed.join('; ')}.` : ` Granted: ${r.fixed.join('; ')}.`) : '';
        const failed = r.failed.length ? (isDe ? ` NICHT setzbar (Konto nicht auflösbar oder Drosselung): ${r.failed.join('; ')}.` : ` Could NOT be granted (account not resolvable or throttling): ${r.failed.join('; ')}.`) : '';
        msg = isDe
          ? `${r.checked} Rollen-Einträge geprüft, bei ${r.missing.length} fehlte mindestens ein Recht — ohne Leserecht fehlt die Kachel, ohne Site-Vollzugriff scheitert das Anlegen eines Events („Subsite konnte nicht erstellt werden").${fixed}${failed} Betroffene müssen die App einmal neu laden.`
          : `${r.checked} role entries checked, ${r.missing.length} lacked at least one right — without read access the tile is missing, without site full control creating an event fails.${fixed}${failed} Affected people need to reload the app once.`;
      }
      // v31.67: Adress-Abweichungen getrennt melden — die Rechte sind da, nur
      // die E-Mail in DEX_Roles ist eine andere Schreibweise als am Konto.
      // Vorher wurden diese Personen bei JEDEM Lauf als „Lücke" gemeldet und
      // „nachgesetzt" (Nutzer-Befund 16.09.2026).
      // v31.91: Die Abweichungen merken — die Karte „Person umbenennen" unten
      // bietet sie als Vorbelegung an.
      setAliasHints(!r.readFailed && r.aliases ? r.aliases.map(a => ({ name: a.name, email: a.email, spEmail: a.spEmail })) : []);
      if (!r.readFailed && r.aliases && r.aliases.length > 0) {
        const list = r.aliases.map(a => `${a.name || a.email} (DEX_Roles: ${a.email} · SharePoint: ${a.spEmail})`).join('; ');
        msg += isDe
          ? ` Hinweis: ${r.aliases.length} ${r.aliases.length === 1 ? 'Eintrag hat' : 'Einträge haben'} alle Rechte, aber unter einer anderen E-Mail-Schreibweise als in DEX_Roles — ${list}. Am besten die Zeile in der Rollenverwaltung auf die SharePoint-Adresse ändern, dann verschwindet der Hinweis.`
          : ` Note: ${r.aliases.length} ${r.aliases.length === 1 ? 'entry has' : 'entries have'} all rights, but under a different e-mail spelling than in DEX_Roles — ${list}. Change the row in the role management to the SharePoint address to clear this note.`;
      }
      setAccessAudit({ running: false, done: 0, total: 0, result: msg });
    } catch (e) {
      console.warn('[DEX] auditRolesAccess failed:', e);
      setAccessAudit({ running: false, done: 0, total: 0, result: isDe ? 'Prüfung abgebrochen — bitte erneut versuchen.' : 'Check aborted — please try again.' });
    }
  };
  const [isSearching, setIsSearching] = React.useState(false);
  const [userFound, setUserFound] = React.useState<boolean | null>(null);
  const [suggestions, setSuggestions] = React.useState<Array<{ email: string; displayName: string; location: string }>>([]);
  const [showSuggestions, setShowSuggestions] = React.useState(false);
  const [includeIntl, setIncludeIntl] = React.useState(false);
  const searchTimer = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  // Timer cleanup bei Unmount
  React.useEffect(() => {
    return () => { if (searchTimer.current) clearTimeout(searchTimer.current); };
  }, []);

  // Email-Eingabe: User-Autocomplete mit Debounce
  const handleEmailChange = (query: string): void => {
    setNewEmail(query);
    setUserFound(null);

    if (searchTimer.current) clearTimeout(searchTimer.current);

    if (query.length >= 2) {
      searchTimer.current = setTimeout(async () => {
        setIsSearching(true);
        const results = await searchUsers(query, includeIntl);
        setSuggestions(results);
        setShowSuggestions(results.length > 0);
        setIsSearching(false);
      }, 400);
    } else {
      setSuggestions([]);
      setShowSuggestions(false);
    }
  };

  React.useEffect(() => {
    if (newEmail.length >= 2 && userFound !== true) {
      (async () => {
        setIsSearching(true);
        try {
          const results = await searchUsers(newEmail, includeIntl);
          setSuggestions(results);
          setShowSuggestions(results.length > 0);
        } catch {
          setSuggestions([]);
        }
        setIsSearching(false);
      })().catch(() => {});
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [includeIntl]);

  const selectSuggestion = (suggestion: { email: string; displayName: string; location: string }): void => {
    setNewEmail(suggestion.email);
    setNewName(suggestion.displayName);
    setNewLocation(suggestion.location);
    setUserFound(true);
    setShowSuggestions(false);
    setSuggestions([]);
  };

  // Onboarding-Mail-Prompt: erscheint nach erfolgreicher Zuweisung einer
  // Organizer- oder Admin-Rolle. Ein Klick auf "Mail senden" verschickt eine
  // Begrüßungsmail mit Links zur App und zum Handbuch (siehe
  // organizerOnboardingEmail in EmailTemplates.ts) — Cc geht automatisch an
  // die DEX-Verantwortlichen.
  const [onboardingPrompt, setOnboardingPrompt] = React.useState<
    { email: string; name: string; role: 'Organizer' | 'Admin' } | null
  >(null);
  const [isSendingOnboarding, setIsSendingOnboarding] = React.useState(false);

  const handleAddRole = async (): Promise<void> => {
    if (!newEmail || !newName) return;
    setIsAdding(true);
    setStatusMsg('');
    const assignedEmail = newEmail;
    const assignedName = newName;
    const assignedRole = newRole;
    // v24.84: Duplikaterkennung — hat die Person schon eine Rolle, nicht erneut
    // hinzufügen (sonst doppelte Zeile in DEX_Roles). Stattdessen Hinweis, die
    // bestehende Rolle in der Liste unten zu ändern.
    const existing = roles.find(r => (r.userEmail || '').toLowerCase() === assignedEmail.toLowerCase());
    if (existing) {
      setStatusMsg(`Error: ${assignedName} hat bereits eine Rolle (${existing.role}). Bitte ändere die bestehende Rolle in der Liste unten, statt die Person erneut hinzuzufügen.`);
      setIsAdding(false);
      setTimeout(() => setStatusMsg(''), 6000);
      return;
    }
    const success = await addRole(assignedEmail, assignedName, assignedRole, newLocation);
    if (success) {
      // v30.85: Zeile gesetzt heißt nicht Rechte gesetzt — das war der Weg,
      // auf dem 18 Organizer ohne Site-Recht entstanden sind.
      const missing = hadRoleRightsIssue() ? lastRoleRightsMissing() : [];
      setStatusMsg(missing.length > 0
        ? (isDe
          ? `Rolle zugewiesen — aber diese Rechte konnten trotz Wiederholung NICHT gesetzt werden: ${missing.join(', ')}. Bitte in ein paar Minuten „Rechte prüfen" ausführen, sonst fehlt der Person die Kachel bzw. das Anlegen scheitert.`
          : `Role assigned — but these rights could NOT be granted despite retries: ${missing.join(', ')}. Please run "Check rights" in a few minutes, otherwise the tile is missing or creating events fails.`)
        : 'Role assigned successfully.');
      setNewEmail('');
      setNewName('');
      setNewLocation('');
      setShowAddForm(false);
      // v28.44: Die Onboarding-Mail geht jetzt AUTOMATISCH raus — vorher war sie
      // ein Angebot per Rückfrage, das man wegklicken konnte, und dann startete
      // der neue Organizer ohne Links, Handbuch und Einsatzbereich-Hinweis.
      // Die User-Rolle bleibt aussen vor: Die Mail erklärt Organizer-/Admin-
      // Funktionen, die Standard-User gar nicht haben.
      // v31.42: F&A dazu. Die Rolle ist Organizer PLUS Abrechnung — sie bekam
      // bisher gar keine Mail und erfuhr damit auch nichts vom Pilotbetrieb.
      if ((assignedRole === 'Organizer' || assignedRole === 'Admin' || assignedRole === 'F&A') && missing.length === 0) {
        void sendOrganizerOnboarding(assignedEmail, assignedName, assignedRole)
          .then(sent => setStatusMsg(sent
            ? 'Rolle zugewiesen — Onboarding-Mail wurde verschickt.'
            : 'Rolle zugewiesen. Die Onboarding-Mail konnte nicht verschickt werden — bitte über den Briefumschlag in der Liste erneut senden.'))
          .catch(() => { /* Status bleibt bei der Zuweisungs-Meldung */ });
      }
    } else {
      setStatusMsg('Error: Could not assign role. Please try again.');
    }
    setIsAdding(false);
    setTimeout(() => setStatusMsg(''), 4000);
  };

  const handleSendOnboarding = async (): Promise<void> => {
    if (!onboardingPrompt) return;
    setIsSendingOnboarding(true);
    const ok = await sendOrganizerOnboarding(onboardingPrompt.email, onboardingPrompt.name, onboardingPrompt.role);
    setIsSendingOnboarding(false);
    setOnboardingPrompt(null);
    setStatusMsg(ok ? 'Onboarding-Mail wurde verschickt.' : 'Onboarding-Mail konnte nicht versendet werden.');
    setTimeout(() => setStatusMsg(''), 4000);
  };

  const [isRemoving, setIsRemoving] = React.useState<number | null>(null);

  const [onboardingResendId, setOnboardingResendId] = React.useState<number | null>(null);
  const resendOnboarding = async (
    itemId: number, email: string, name: string, role: 'Organizer' | 'Admin' | 'F&A',
  ): Promise<void> => {
    const ok = await confirmDialog(isDe
      ? `Onboarding-Mail an ${name || email} senden?\n\nSie enthält die wichtigsten Links, eine Anleitung für das erste Test-Event und den Hinweis, dass DEX für interne Deloitte Events gedacht ist.`
      : `Send the onboarding mail to ${name || email}?\n\nIt contains the key links, a walkthrough for a first test event and the note that DEX is meant for internal Deloitte events.`,
      { confirmLabel: isDe ? 'Senden' : 'Send' });
    if (!ok) return;
    setOnboardingResendId(itemId);
    const sent = await sendOrganizerOnboarding(email, name, role).catch(() => false);
    setOnboardingResendId(null);
    showAlert(
      sent
        ? (isDe ? 'Onboarding-Mail wurde in die Warteschlange eingetragen.' : 'Onboarding mail has been queued.')
        : (isDe ? 'Die Onboarding-Mail konnte nicht verschickt werden.' : 'Could not send the onboarding mail.'),
      { variant: sent ? 'success' : 'error' },
    );
  };

  const handleRemoveRole = async (itemId: number, userName: string): Promise<void> => {
    const confirmed = await confirmDialog(`Rolle für "${userName}" entfernen?`, { danger: true, confirmLabel: isDe ? 'Entfernen' : 'Remove' });
    if (!confirmed) return;
    setIsRemoving(itemId);
    const success = await removeRole(itemId);
    // v30.67 (Review): Zwei Fehlerarten, zwei Meldungen. `success` sagt nur,
    // ob die DEX_Roles-Zeile weg ist; ob die SharePoint-Rechte mitgingen,
    // sagt hadRoleRightsIssue(). Als EIN boolean war die Meldung im Fall
    // „Zeile nicht gelöscht" falsch („entfernt — aber …", während die Zeile
    // weiter in der Tabelle stand).
    const rightsIssue = hadRoleRightsIssue();
    setIsRemoving(null);
    if (success && !rightsIssue) {
      setStatusMsg(isDe ? `Rolle für ${userName} entfernt.` : `Role for ${userName} removed.`);
    } else if (success) {
      // Die DEX_Roles-Zeile ist weg; was blieb, ist ein Recht auf dem Web
      // oder einer Liste. Das muss der Admin wissen, sonst hält er die
      // Person für ausgesperrt, die es nicht ist.
      setStatusMsg(isDe
        ? `Rolle für ${userName} entfernt — aber die SharePoint-Rechte konnten nicht vollständig entzogen werden. Bitte die Berechtigungen der Person auf der Site prüfen (Details in der Browser-Konsole unter [DEX]).`
        : `Role for ${userName} removed — but the SharePoint permissions could not be fully revoked. Please check the person's permissions on the site (details in the browser console under [DEX]).`);
    } else {
      setStatusMsg(isDe
        ? `Rolle für ${userName} konnte nicht entfernt werden — die Zeile steht weiterhin in der Rollenliste. Bitte erneut versuchen.`
        : `Role for ${userName} could not be removed — the row is still in the roles list. Please try again.`);
    }
    setTimeout(() => setStatusMsg(''), (success && !rightsIssue) ? 4000 : 12000);
  };

  const handleChangeRole = async (itemId: number, role: UserRole): Promise<void> => {
    // v30.67: Rueckgabewert auswerten — vorher blieb ein nicht entzogenes Recht
    // beim Herabstufen unsichtbar.
    const ok = await updateRole(itemId, role);
    // v30.67 (Review): s. handleRemoveRole — Zeile und Rechte getrennt melden.
    const rightsIssue = hadRoleRightsIssue();
    if (!ok) {
      setStatusMsg(isDe
        ? 'Rolle konnte nicht geändert werden — die Zeile trägt weiterhin die alte Rolle. Bitte erneut versuchen.'
        : 'Role could not be changed — the row still carries the old role. Please try again.');
      setTimeout(() => setStatusMsg(''), 12000);
    } else if (rightsIssue) {
      setStatusMsg(isDe
        ? 'Rolle geändert — aber die SharePoint-Rechte konnten nicht vollständig angepasst werden. Bitte die Berechtigungen der Person auf der Site prüfen (Details in der Browser-Konsole unter [DEX]).'
        : 'Role changed — but the SharePoint permissions could not be fully adjusted. Please check the person\'s permissions on the site (details in the browser console under [DEX]).');
      setTimeout(() => setStatusMsg(''), 12000);
    }
  };

  // v24.87: roleBadge entfernt (war nur für die gelöschte „User Information"-Karte).

  // ===================== v24.85: Role-Management nach Kategorien =====================
  // Test-Team + Check-in-Team über ALLE Events aggregieren (read-only, analog
  // zu den Per-Event-Co-Organizern). Datenwerte: event.testTeamEmails/-Names
  // bzw. event.qrScannerEmails/-Names.
  const aggregateTeam = React.useCallback((pick: (e: (typeof events)[number]) => { emails?: string[]; names?: string[] }) => {
    const acc: Record<string, { name: string; events: string[] }> = {};
    for (const evt of events) {
      const sel = pick(evt); const emails = sel.emails || []; const names = sel.names || [];
      for (let i = 0; i < emails.length; i++) {
        const lc = (emails[i] || '').toLowerCase(); if (!lc) continue;
        if (!acc[lc]) acc[lc] = { name: names[i] || lc, events: [] };
        if (acc[lc].events.indexOf(evt.title) < 0) acc[lc].events.push(evt.title);
      }
    }
    return Object.keys(acc).map(lc => ({ email: lc, name: acc[lc].name, events: acc[lc].events }))
      .sort((a, b) => (a.name || '').localeCompare(b.name || '', 'de'));
  }, [events]);
  const testersList = React.useMemo(() => aggregateTeam(e => ({ emails: e.testTeamEmails, names: e.testTeamNames })), [aggregateTeam]);
  const checkinList = React.useMemo(() => aggregateTeam(e => ({ emails: e.qrScannerEmails, names: e.qrScannerNames })), [aggregateTeam]);

  // Position (Job Title) + Standort pro Person live nachladen — DEX_Roles
  // speichert die Position nicht. Best-effort.
  // v30.82: Vorher EIN searchUser je Person, sequentiell — je Person bis zu
  // vier SharePoint-/Graph-Aufrufe, bei 130 Zeilen 130 bis 500 Requests bei
  // jedem Öffnen der Seite (Nutzer-Frage 07.09.2026: „ist das so sinnvoll
  // wegen Throttle?"). Jetzt: 24-h-Cache in localStorage, dann EIN Graph-
  // Batch für alle noch fehlenden (20 je Request), und nur für die, die Graph
  // nicht kennt, der alte Einzelweg — gedrosselt mit Pause.
  const [profiles, setProfiles] = React.useState<Record<string, { jobTitle?: string; location?: string }>>(() => {
    try {
      const raw = window.localStorage.getItem('dex_role_profiles_v1');
      if (!raw) return {};
      const parsed = JSON.parse(raw) as { ts: number; data: Record<string, { jobTitle?: string; location?: string }> };
      if (!parsed || !parsed.data || (Date.now() - (parsed.ts || 0)) > 24 * 60 * 60 * 1000) return {};
      return parsed.data;
    } catch { return {}; }
  });
  const profileAttemptedRef = React.useRef<Set<string>>(new Set());
  React.useEffect(() => {
    const emails = new Set<string>();
    roles.forEach(r => emails.add((r.userEmail || '').toLowerCase()));
    coOrganizersList.forEach(c => emails.add(c.email));
    testersList.forEach(c => emails.add(c.email));
    checkinList.forEach(c => emails.add(c.email));
    const todo = Array.from(emails).filter(em => !!em && !profileAttemptedRef.current.has(em) && !profiles[em]);
    if (todo.length === 0) return;
    todo.forEach(em => profileAttemptedRef.current.add(em));
    let cancelled = false;
    (async () => {
      const found: Record<string, { jobTitle?: string; location?: string }> = {};
      try {
        const batch = await getBasicProfiles(todo);
        for (const em of Object.keys(batch)) {
          const b = batch[em];
          if (b.jobTitle || b.location) found[em] = { jobTitle: b.jobTitle, location: b.location };
        }
      } catch { /* Graph nicht erreichbar → Einzelweg unten */ }
      if (cancelled) return;
      if (Object.keys(found).length > 0) setProfiles(prev => ({ ...prev, ...found }));
      // Rest über den Einzelweg — mit Pause, damit die Seite SharePoint nicht
      // in die Drosselung treibt; typischerweise nur Aliase und Externe.
      const rest = todo.filter(em => !found[em]);
      for (let i = 0; i < rest.length; i++) {
        if (cancelled) return;
        if (i > 0) await new Promise<void>(resolve => setTimeout(resolve, 400));
        try {
          const u = await searchUser(rest[i]);
          if (u && !cancelled && (u.jobTitle || u.location)) {
            setProfiles(prev => ({ ...prev, [rest[i]]: { jobTitle: u.jobTitle, location: u.location } }));
          }
        } catch { /* best-effort */ }
      }
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [roles, coOrganizersList, testersList, checkinList]);
  // Cache schreiben — verschmerzbar, wenn localStorage gesperrt ist.
  React.useEffect(() => {
    if (Object.keys(profiles).length === 0) return;
    try { window.localStorage.setItem('dex_role_profiles_v1', JSON.stringify({ ts: Date.now(), data: profiles })); } catch { /* */ }
  }, [profiles]);

  // Klapp-Status pro Kategorie (Default: alle offen).
  const [openSections, setOpenSections] = React.useState<Set<string>>(() => new Set(['admins', 'organizer', 'coorg', 'tester', 'checkin', 'user']));
  const toggleSection = (k: string): void => setOpenSections(prev => { const n = new Set(prev); if (n.has(k)) n.delete(k); else n.add(k); return n; });
  // v24.87: Freitext-Suche über alle Kategorien (Name/E-Mail/Position/Standort).
  const [roleSearch, setRoleSearch] = React.useState('');
  // Deloitte-Displayname „Nachname, Vorname" → { first, last }.
  const splitName = (full: string): { first: string; last: string } => {
    const n = (full || '').trim();
    if (!n) return { first: '', last: '' };
    const c = n.indexOf(',');
    if (c >= 0) return { last: n.substring(0, c).trim(), first: n.substring(c + 1).trim() };
    const parts = n.split(/\s+/);
    return parts.length > 1 ? { first: parts[0], last: parts.slice(1).join(' ') } : { first: n, last: '' };
  };

  const renderRoleSections = (): React.ReactElement => {
    const thS: React.CSSProperties = { textAlign: 'left', padding: 8, color: 'var(--dex-gray-500)', fontSize: '0.76rem', fontWeight: 600, whiteSpace: 'nowrap' };
    const tdS: React.CSSProperties = { padding: 8, verticalAlign: 'middle' };
    const byName = (a: { userName: string }, b: { userName: string }): number => (a.userName || '').localeCompare(b.userName || '', 'de');
    // v24.87: Freitext-Filter (Name/E-Mail/Position/Standort) über alle Kategorien.
    const q = roleSearch.trim().toLowerCase();
    const hit = (parts: Array<string | undefined>): boolean => !q || parts.some(s => (s || '').toLowerCase().indexOf(q) >= 0);
    const matchRole = (r: typeof roles[number]): boolean => { const p = profiles[(r.userEmail || '').toLowerCase()] || {}; return hit([r.userName, r.userEmail, p.jobTitle, p.location, r.location]); };
    const matchAgg = (x: { email: string; name: string }): boolean => { const p = profiles[x.email] || {}; return hit([x.name, x.email, p.jobTitle, p.location]); };
    // v26.33: IT-Admins zählen zur Admin-Gruppe (gleiche Rechte).
    const admins = [...roles].filter(r => r.role === 'Admin' || r.role === 'IT-Admin').filter(matchRole).sort(byName);
    // v30.60: F&A steht in der Organizer-Gruppe, nicht bei den Teilnehmern.
    // Seit die Rolle Organizer-Rechte trägt, wäre sie unter „User" eine
    // Falschaussage über die Rechte dieser Person; das F&A-Pill in der
    // Status-Spalte unterscheidet sie weiterhin.
    const organizers = [...roles].filter(r => r.role === 'Organizer' || r.role === 'F&A').filter(matchRole).sort(byName);
    const usersLeft = [...roles].filter(r => r.role !== 'Admin' && r.role !== 'IT-Admin' && r.role !== 'Organizer' && r.role !== 'F&A').filter(matchRole).sort(byName);
    const eventBadges = (titles: string[]): React.ReactNode => titles.length === 0
      ? <span style={{ color: 'var(--dex-gray-300)' }}>—</span>
      : <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>{titles.map((t2, idx) => (
          <span key={idx} style={{ display: 'inline-block', padding: '2px 8px', borderRadius: 10, background: 'rgba(134,188,37,0.14)', color: 'var(--dex-green-dark)', fontSize: '0.72rem', fontWeight: 600 }}>{t2}</span>
        ))}</div>;
    const catPill = (label: string, bg: string, fg: string): React.ReactNode => (
      <span style={{ display: 'inline-block', padding: '2px 10px', borderRadius: 999, background: bg, color: fg, fontSize: '0.74rem', fontWeight: 600, whiteSpace: 'nowrap' }}>{label}</span>
    );
    const headRow = (
      <tr style={{ borderBottom: '2px solid var(--dex-gray-200, #eee)' }}>
        <th style={thS}>{isDe ? 'Vorname' : 'First name'}</th>
        <th style={thS}>{isDe ? 'Nachname' : 'Last name'}</th>
        <th style={thS}>Email</th>
        <th style={thS}>{isDe ? 'Position' : 'Position'}</th>
        <th style={thS}>{isDe ? 'Standort' : 'Location'}</th>
        <th style={thS}>Role</th>
        <th style={thS}>Status</th>
        <th style={thS}>Power User</th>
        <th style={thS}>Coordinated Events</th>
        <th style={{ ...thS, textAlign: 'right' }} />
      </tr>
    );
    // v24.87: Status-Badges (Admin/Organizer/User + Co-Organizer/Tester/Check-in).
    const adminPill = catPill('Admin', '#e8f5e9', '#2e7d32');
    // v26.33: IT-Admin — volle Admin-Rechte, aber kein Empfänger der Mails.
    const itAdminPill = catPill('IT-Admin', '#ede7f6', '#5e35b1');
    const organizerPill = catPill('Organizer', 'rgba(0,118,168,0.10)', 'var(--dex-blue, #0076a8)');
    const userPill = catPill('User', '#f5f5f5', '#666');
    // v30.5/v30.60: F&A — alles wie Organizer, plus Zugriff aufs F&A Center.
    const faPill = catPill('F&A', 'rgba(237,139,0,0.12)', '#b86700');
    const coOrgPill = catPill('Co-Organizer', 'rgba(237,139,0,0.15)', 'var(--dex-orange-dark, #b35a00)');
    // v31.86: Organizer sehen dieselben Abschnitte, aber ohne Bedienung —
    // kein Rollen-Select, kein Power-User-Schalter, kein Löschen/Onboarding.
    const readOnly = !isAdmin;
    // Editierbare Zeile (DEX_Roles: Admins / Organizer / User)
    const editableRow = (r: typeof roles[number]): React.ReactElement => {
      const { first, last } = splitName(r.userName);
      const emailLc = (r.userEmail || '').toLowerCase();
      const prof = profiles[emailLc] || {};
      const pos = prof.jobTitle || '';
      const isSelf = emailLc === currentUser.email.toLowerCase();
      const evts = organizerEventMap[emailLc] || [];
      return (
        <tr key={r.id} style={{ borderBottom: '1px solid var(--dex-gray-100, #f0f0f0)' }}>
          <td style={tdS}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 0 }}>
              <PersonContactHover email={r.userEmail} name={r.userName} size={30} subline={pos} isDe={isDe} />
              <span style={{ fontWeight: 500 }}>{first || '-'}</span>
            </div>
          </td>
          <td style={{ ...tdS, fontWeight: 500 }}>{last || '-'}</td>
          <td style={{ ...tdS, color: 'var(--dex-gray-600)' }}>{r.userEmail}</td>
          <td style={{ ...tdS, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{pos || '—'}</td>
          {/* v24.87: Standort ist NICHT editierbar — er ergibt sich aus dem
              Profil zur E-Mail-Adresse (live nachgeladen, Fallback DEX_Roles). */}
          <td style={{ ...tdS, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{prof.location || r.location || '—'}</td>
          <td style={tdS}>
            {readOnly ? (
              <span style={{ fontSize: '0.8rem', color: 'var(--dex-gray-700)' }}>{r.role}</span>
            ) : (
            <select
              value={r.role}
              onChange={e => handleChangeRole(r.id, e.target.value as UserRole)}
              style={{ padding: '4px 8px', borderRadius: 6, border: '1px solid var(--dex-gray-200, #ddd)', fontSize: '0.8rem', background: '#fff' }}
              disabled={isSelf}
            >
              <option value="Admin">Admin</option>
              <option value="IT-Admin">IT-Admin</option>
              <option value="Organizer">Organizer</option>
              <option value="F&A">F&A</option>
              <option value="User">User</option>
            </select>
            )}
          </td>
          <td style={tdS}>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
              {r.role === 'Admin' && adminPill}
              {r.role === 'IT-Admin' && itAdminPill}
              {r.role === 'Organizer' && organizerPill}
              {r.role === 'F&A' && faPill}
              {r.role === 'User' && userPill}
              {/* v24.87: globaler Organizer, der zusätzlich Events betreut → auch Co-Organizer. */}
              {r.role === 'Organizer' && evts.length > 0 && coOrgPill}
            </div>
          </td>
          <td style={tdS}>
            {readOnly ? (
              <span style={{ fontSize: '0.78rem', color: r.isPowerUser ? '#b35a00' : 'var(--dex-gray-300)' }}>{r.isPowerUser ? '★ Power User' : '—'}</span>
            ) : (r.role === 'Organizer' || r.role === 'Admin') ? (
              <button
                type="button"
                disabled={isSelf}
                onClick={() => setPowerUser(r.id, !r.isPowerUser)}
                title={r.isPowerUser ? (isDe ? 'Power-User-Status entfernen' : 'Remove power-user status') : (isDe ? 'Als Power User markieren' : 'Mark as power user')}
                style={{ display: 'inline-flex', alignItems: 'center', gap: 4, padding: '4px 10px', borderRadius: 999, cursor: isSelf ? 'default' : 'pointer', fontSize: '0.74rem', fontWeight: 600, whiteSpace: 'nowrap', border: `1px solid ${r.isPowerUser ? '#b35a00' : 'var(--dex-gray-300)'}`, background: r.isPowerUser ? '#fff4e5' : '#fff', color: r.isPowerUser ? '#b35a00' : 'var(--dex-gray-600)' }}
              >
                <span aria-hidden="true">{r.isPowerUser ? '★' : '☆'}</span>
                {r.isPowerUser ? 'Power User' : (isDe ? 'Power User?' : 'Power user?')}
              </button>
            ) : <span style={{ color: 'var(--dex-gray-300)' }}>—</span>}
          </td>
          <td style={{ ...tdS, fontSize: '0.78rem', color: 'var(--dex-gray-600)', maxWidth: 280 }}>
            {r.role === 'User' ? <span style={{ color: 'var(--dex-gray-300)' }}>—</span> : <CoordinatedEventsCell titles={evts} isDe={isDe} />}
          </td>
          <td style={{ ...tdS, textAlign: 'right', whiteSpace: 'nowrap' }}>
            {/* v28.44: Onboarding-Mail nachträglich verschicken — für alle,
                die vor dieser Version Organizer wurden (und damals keine
                bekommen haben) oder die sie nicht mehr finden. */}
            {!readOnly && (r.role === 'Organizer' || r.role === 'Admin' || r.role === 'F&A') && (
              <button
                onClick={() => { void resendOnboarding(r.id, r.userEmail, r.userName, r.role as 'Organizer' | 'Admin' | 'F&A'); }}
                disabled={onboardingResendId === r.id}
                style={{ border: 'none', background: 'none', cursor: onboardingResendId === r.id ? 'wait' : 'pointer', color: 'var(--dex-green-dark, #4a7c1f)', padding: 4, opacity: onboardingResendId === r.id ? 0.4 : 1 }}
                title={isDe ? 'Onboarding-Mail (erneut) senden' : 'Send onboarding mail (again)'}
              >
                {onboardingResendId === r.id ? '...' : <Mail size={16} />}
              </button>
            )}
            {!readOnly && !isSelf && (
              <button onClick={() => handleRemoveRole(r.id, r.userName)} disabled={isRemoving === r.id} style={{ border: 'none', background: 'none', cursor: isRemoving === r.id ? 'wait' : 'pointer', color: 'var(--dex-danger, #e53935)', padding: 4, opacity: isRemoving === r.id ? 0.4 : 1 }} title="Rolle entfernen">
                {isRemoving === r.id ? '...' : <Trash2 size={16} />}
              </button>
            )}
          </td>
        </tr>
      );
    };
    // Read-only Zeile (aggregiert: Co-Organizer / Tester / Check-in)
    const aggRow = (p: { email: string; name: string; events: string[] }, rolePill: React.ReactNode): React.ReactElement => {
      const { first, last } = splitName(p.name);
      const prof = profiles[p.email] || {};
      const displayName = resolvedOrgNames[p.email] || p.name;
      const dn = splitName(displayName);
      return (
        <tr key={p.email} style={{ borderBottom: '1px solid var(--dex-gray-100, #f0f0f0)' }}>
          <td style={tdS}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 0 }}>
              <PersonContactHover email={p.email} name={displayName} size={30} subline={prof.jobTitle || ''} isDe={isDe} />
              <span style={{ fontWeight: 500 }}>{dn.first || first || '-'}</span>
            </div>
          </td>
          <td style={{ ...tdS, fontWeight: 500 }}>{dn.last || last || '-'}</td>
          <td style={{ ...tdS, color: 'var(--dex-gray-600)' }}>{p.email}</td>
          <td style={{ ...tdS, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{prof.jobTitle || '—'}</td>
          <td style={{ ...tdS, color: 'var(--dex-gray-600)', fontSize: '0.8rem' }}>{prof.location || '—'}</td>
          <td style={tdS}><span style={{ color: 'var(--dex-gray-300)' }}>—</span></td>
          <td style={tdS}>{rolePill}</td>
          <td style={tdS}><span style={{ color: 'var(--dex-gray-300)' }}>—</span></td>
          <td style={{ ...tdS, fontSize: '0.78rem', color: 'var(--dex-gray-600)', maxWidth: 280 }}>{eventBadges(p.events)}</td>
          <td style={{ ...tdS, textAlign: 'right' }} />
        </tr>
      );
    };
    const coorgF = coOrganizersList.filter(matchAgg);
    const testersF = testersList.filter(matchAgg);
    const checkinF = checkinList.filter(matchAgg);
    // v24.87: Organizer + Co-Organizer in EINEM Abschnitt — globale Organizer
    // (editierbar) zuerst, danach die reinen Per-Event-Co-Organizer (read-only).
    // Die Status-Spalte zeigt „Organizer" und/oder „Co-Organizer".
    const organizerSectionBody = [
      ...organizers.map(editableRow),
      ...coorgF.map(p => aggRow(p, coOrgPill)),
    ];
    const sections: Array<{ key: string; title: string; body: React.ReactNode; count: number }> = [
      { key: 'admins', title: 'Admins', count: admins.length, body: admins.map(editableRow) },
      { key: 'organizer', title: isDe ? 'Organizer & Co-Organizer' : 'Organizers & co-organizers', count: organizers.length + coorgF.length, body: organizerSectionBody },
      { key: 'tester', title: isDe ? 'Tester' : 'Testers', count: testersF.length, body: testersF.map(p => aggRow(p, catPill(isDe ? 'Tester' : 'Tester', 'rgba(0,118,168,0.10)', 'var(--dex-blue, #0076a8)'))) },
      { key: 'checkin', title: 'Check-in', count: checkinF.length, body: checkinF.map(p => aggRow(p, catPill('Check-in', 'rgba(134,188,37,0.16)', 'var(--dex-green-dark, #4a7c1f)'))) },
    ];
    if (usersLeft.length > 0) sections.push({ key: 'user', title: isDe ? 'Weitere (User)' : 'Other (users)', count: usersLeft.length, body: usersLeft.map(editableRow) });
    return (
      <div>
        {sections.map(sec => {
          const open = q ? true : openSections.has(sec.key);
          return (
            <div key={sec.key} style={{ marginBottom: 12, border: '1px solid var(--dex-gray-200, #eee)', borderRadius: 8, overflow: 'hidden' }}>
              <button
                type="button"
                onClick={() => toggleSection(sec.key)}
                style={{ width: '100%', textAlign: 'left', display: 'flex', alignItems: 'center', gap: 8, padding: '10px 14px', background: 'var(--dex-gray-50, #fafafa)', border: 'none', cursor: 'pointer', fontWeight: 700, fontSize: '0.95rem', color: 'var(--dex-gray-800)' }}
              >
                <span style={{ color: 'var(--dex-gray-500)' }}>{open ? '▾' : '▸'}</span>
                {sec.title} <span style={{ color: 'var(--dex-gray-400)', fontWeight: 500 }}>({sec.count})</span>
              </button>
              {open && (
                <div style={{ overflowX: 'auto', padding: '0 6px 6px' }}>
                  {sec.count === 0 ? (
                    <p style={{ padding: '8px 10px', margin: 0, color: 'var(--dex-gray-400)', fontStyle: 'italic', fontSize: '0.82rem' }}>{isDe ? 'keine' : 'none'}</p>
                  ) : (
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem' }}>
                      <thead>{headRow}</thead>
                      <tbody>{sec.body}</tbody>
                    </table>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
    );
  };

  return (
    <div className="page-container">
      <div>

        {/* v24.87: „User Information"-Karte entfernt — diese Seite ist reine
            Rollenverwaltung (die eigenen User-Infos stehen im Header-Avatar-Menü). */}

        {/* v24.84: „Admin Actions"-Karte (Create New Event / View All Events /
            Rollen-Matrix) entfernt — diese Wege gibt es bereits über die
            Start-/Admin-Kacheln. Diese Seite ist jetzt reine Rollenverwaltung. */}

        {/* v9.21: Rollenmanagement collapsible — Admin kann die ganze Liste
            zusammenklappen wenn er sie gerade nicht braucht. Default: zu. */}
        {isAdmin && (
          <div className="card">
            {/* v24.85: nicht mehr die ganze Karte einklappbar — stattdessen je
                Kategorie ein eigener aufklappbarer Abschnitt (renderRoleSections). */}
            <h2 style={{ margin: '0 0 12px', fontSize: '1.1rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Rollenverwaltung' : 'Role management'}
            </h2>
            <div style={{ marginTop: 4 }}>
            <p style={{ color: 'var(--dex-gray-500, #888)', fontSize: '0.85rem', marginBottom: 16 }}>
              Manage who can create events. Roles are stored in the SharePoint list &ldquo;DEX_Roles&rdquo;.
              {' '}
              <a
                href={`${siteUrl}/Lists/DEX_Roles/AllItems.aspx`}
                target="_blank"
                rel="noopener noreferrer"
                style={{ color: 'var(--dex-green-dark, #6b9a1e)', fontWeight: 600, textDecoration: 'underline' }}
              >
                Open SharePoint List
              </a>
            </p>

            {/* v29.63: Doppelte Rollen-Eintraege beim Aufruf benennen.
                `addRole` prueft nicht auf Bestand — jede Zuweisung legt eine
                neue Zeile in DEX_Roles an. Zwei Zeilen fuer dieselbe Person
                sind kein Schoenheitsfehler: `refreshRoles` nimmt per `find`
                die ERSTE, also entscheidet die Reihenfolge in der Liste
                darueber, welche Rolle gilt — und eine spaetere Herabstufung
                kann von einer alten Zeile ueberstimmt werden. */}
            {(() => {
              const byMail = new Map<string, typeof roles>();
              for (const r of roles) {
                const k = (r.userEmail || '').trim().toLowerCase();
                if (!k) continue;
                const arr = byMail.get(k) || [];
                arr.push(r);
                byMail.set(k, arr);
              }
              const dupes = Array.from(byMail.values()).filter(a => a.length > 1);
              if (dupes.length === 0) return null;
              return (
                <div style={{
                  padding: '10px 14px', borderRadius: 8, marginBottom: 16, fontSize: '0.85rem',
                  background: 'rgba(237,139,0,0.10)', border: '1px solid var(--dex-orange, #ed8b00)',
                  color: 'var(--dex-gray-800)',
                }}>
                  <strong>
                    {isDe
                      ? `${dupes.length} ${dupes.length === 1 ? 'Person hat' : 'Personen haben'} mehrere Rollen-Einträge`
                      : `${dupes.length} ${dupes.length === 1 ? 'person has' : 'people have'} multiple role entries`}
                  </strong>
                  <div style={{ marginTop: 4 }}>
                    {isDe
                      ? 'Gültig ist der erste Eintrag der Liste — eine spätere Änderung kann dadurch wirkungslos bleiben. Bitte die überzähligen Zeilen entfernen.'
                      : 'The first entry in the list wins — a later change can therefore have no effect. Please remove the surplus rows.'}
                  </div>
                  <ul style={{ margin: '8px 0 0', paddingLeft: 18 }}>
                    {dupes.slice(0, 10).map(a => (
                      <li key={a[0].userEmail} style={{ marginBottom: 2 }}>
                        <strong>{a[0].userName || a[0].userEmail}</strong>
                        {' — '}
                        {a.map(x => x.role).join(', ')}
                        {' '}
                        <span style={{ color: 'var(--dex-gray-500)' }}>({a[0].userEmail})</span>
                      </li>
                    ))}
                  </ul>
                  {dupes.length > 10 && (
                    <div style={{ marginTop: 4, color: 'var(--dex-gray-500)' }}>
                      {isDe ? `… und ${dupes.length - 10} weitere` : `… and ${dupes.length - 10} more`}
                    </div>
                  )}
                </div>
              );
            })()}

            {statusMsg && (
              <div style={{
                padding: '8px 12px', borderRadius: 8, marginBottom: 16, fontSize: '0.85rem',
                background: statusMsg.startsWith('Error') ? '#fce4ec' : '#e8f5e9',
                color: statusMsg.startsWith('Error') ? '#c62828' : '#2e7d32',
              }}>
                {statusMsg}
              </div>
            )}

            {/* v30.81: Leserechte prüfen. Eine Rolle wirkt nur, wenn die Person
                die Rollenliste lesen darf — das Recht wird beim Zuweisen
                best-effort gesetzt und kann an der Drosselung scheitern; wer
                die Zeile direkt in SharePoint bekam, hat es nie. Die Person
                steht dann in der Liste und sieht trotzdem „Organizer werden?"
                (Befund 07.09.2026). */}
            <div style={{
              display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap',
              padding: '10px 14px', borderRadius: 8, marginBottom: 16, fontSize: '0.85rem',
              background: 'var(--dex-gray-50, #f7f7f7)', border: '1px solid var(--dex-gray-200)',
            }}>
              <div style={{ flex: 1, minWidth: 240, color: 'var(--dex-gray-700)', lineHeight: 1.45 }}>
                <strong>{isDe ? 'SharePoint-Rechte der Rollen' : 'SharePoint rights of the roles'}</strong>
                <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
                  {isDe
                    ? 'Eine Rolle braucht drei Rechte, die beim Zuweisen gesetzt werden, aber an der Drosselung scheitern können: Lesen auf DEX_Roles (sonst „Organizer werden?" statt Kachel), Vollzugriff auf DEX_Events und auf die Site (sonst „Subsite konnte nicht erstellt werden" beim Anlegen). Prüft alle Einträge gegen die tatsächlichen Berechtigungen und setzt fehlende nach.'
                    : 'A role needs three rights that are set on assignment but can fail under throttling: read on DEX_Roles (otherwise "Want to become an organizer?" instead of the tile), full control on DEX_Events and on the site (otherwise "Subsite could not be created"). Checks all entries against the actual permissions and grants what is missing.'}
                </div>
                {lastRightsAudit && (
                  <div style={{ marginTop: 4, fontSize: '0.76rem', color: 'var(--dex-gray-500)' }}>
                    {isDe
                      ? `Letzte automatische Prüfung (läuft bei Admins einmal täglich beim Start): ${new Date(lastRightsAudit.ts).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })} — ${lastRightsAudit.checked} geprüft, ${lastRightsAudit.missing} mit Lücke, ${lastRightsAudit.fixed} nachgesetzt${lastRightsAudit.failed > 0 ? `, ${lastRightsAudit.failed} NICHT setzbar` : ''}.`
                      : `Last automatic check (runs once a day for admins at start): ${new Date(lastRightsAudit.ts).toLocaleString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })} — ${lastRightsAudit.checked} checked, ${lastRightsAudit.missing} with gaps, ${lastRightsAudit.fixed} granted${lastRightsAudit.failed > 0 ? `, ${lastRightsAudit.failed} NOT grantable` : ''}.`}
                  </div>
                )}
                {accessAudit.result && (
                  <div style={{ marginTop: 6, color: accessAudit.result.indexOf('NICHT') >= 0 || accessAudit.result.indexOf('NOT') >= 0 || accessAudit.result.indexOf('nicht gelesen') >= 0 ? 'var(--dex-red, #c00)' : 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600 }}>
                    {accessAudit.result}
                  </div>
                )}
              </div>
              <button
                type="button"
                className="btn btn-secondary"
                disabled={accessAudit.running || isRolesLoading}
                onClick={() => { void runAccessAudit(); }}
                style={{ fontSize: '0.82rem', whiteSpace: 'nowrap' }}
              >
                {accessAudit.running
                  ? (accessAudit.total > 0 ? `${isDe ? 'Prüft' : 'Checking'} ${accessAudit.done}/${accessAudit.total} …` : (isDe ? 'Liest Berechtigungen …' : 'Reading permissions …'))
                  : (isDe ? 'Rechte prüfen' : 'Check rights')}
              </button>
            </div>

            {/* v31.85: Rechte auf den Teilnehmerlisten — aus dem Admin Hub
                hierher verschoben. Ein Lauf geht über ALLE Events, Klammer
                und jeden Termin, für Organizer, Co-Organizer und Check-in-
                Team. Additiv: nichts wird entzogen. */}
            <div style={{
              display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap',
              padding: '10px 14px', borderRadius: 8, marginBottom: 16, fontSize: '0.85rem',
              background: 'var(--dex-gray-50, #f7f7f7)', border: '1px solid var(--dex-gray-200)',
            }}>
              <div style={{ flex: 1, minWidth: 240, color: 'var(--dex-gray-700)', lineHeight: 1.45 }}>
                <strong>{isDe ? 'Rechte auf den Teilnehmerlisten (alle Events)' : 'Rights on the participant lists (all events)'}</strong>
                <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)' }}>
                  {isDe
                    ? 'Organizer und Co-Organizer brauchen Vollzugriff, das Check-in-Team das Recht „Design“ (liest alle Zeilen trotz Zeilen-Sicherheit) auf jeder Teilnehmerliste — auf dem Haupt-Event UND auf jedem Termin. Beides wird beim Speichern gesetzt, kann aber scheitern (Drosselung, Person hatte die App noch nie geöffnet, Termin im selben Speichern angelegt). Läuft einmal über alle Events und ergänzt, was fehlt; es wird nichts entzogen.'
                    : 'Organizers and co-organizers need full control, the check-in team the "Design" right (reads all rows despite item-level security) on every participant list — on the main event AND on every date. Both are set on save but can fail (throttling, person had never opened the app, date created in the same save). Runs once over all events and adds what is missing; nothing is revoked.'}
                </div>
                {listPerm.running && (
                  <div style={{ marginTop: 4, fontSize: '0.76rem', color: 'var(--dex-gray-500)' }}>
                    {listPerm.total > 0
                      ? `${listPerm.done}/${listPerm.total}${listPerm.label ? ` · ${listPerm.label}` : ''}`
                      : (isDe ? 'Events werden gelesen …' : 'Reading events …')}
                  </div>
                )}
                {listPerm.result && (
                  <div style={{ marginTop: 6, color: listPerm.failed ? 'var(--dex-red, #c00)' : 'var(--dex-green-dark, #4a7c1f)', fontWeight: 600 }}>
                    {listPerm.result}
                  </div>
                )}
              </div>
              <button
                type="button"
                className="btn btn-secondary"
                disabled={listPerm.running || accessAudit.running}
                onClick={() => { void runListPermissions(); }}
                style={{ fontSize: '0.82rem', whiteSpace: 'nowrap' }}
              >
                {listPerm.running
                  ? (isDe ? 'Läuft …' : 'Running …')
                  : (isDe ? 'Alle Events prüfen' : 'Check all events')}
              </button>
            </div>

            {/* v31.91: Person umbenennen — Heirat, neue Adresse. Vorschau
                zuerst (Trockenlauf mit Zahl je Stelle), dann Schreiben. Die
                Alias-Treffer aus „Rechte prüfen" stehen als Vorbelegung bereit. */}
            <div style={{
              padding: '10px 14px', borderRadius: 8, marginBottom: 16, fontSize: '0.85rem',
              background: 'var(--dex-gray-50, #f7f7f7)', border: '1px solid var(--dex-gray-200)',
            }}>
              <strong>{isDe ? 'Person umbenennen (neue E-Mail-Adresse, neuer Name)' : 'Rename a person (new email address, new name)'}</strong>
              <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)', margin: '2px 0 10px', lineHeight: 1.45 }}>
                {isDe
                  ? 'Nach Heirat oder Adresswechsel: tauscht die alte gegen die neue Adresse in DEX_Roles, in den Organizer-, Co-Organizer-, Check-in- und Test-Team-Feldern aller Events, im Teilnehmer-Register und in den Teilnehmerlisten (eigene Anmeldungen sowie „Registriert von"). Erst die Vorschau, dann das Schreiben. SharePoint-Rechte bleiben — es ist dasselbe Konto.'
                  : 'After marriage or an address change: swaps the old for the new address in DEX_Roles, in the organizer, co-organizer, check-in and test-team fields of all events, in the participant registry and in the participant lists (own registrations and “registered by”). Preview first, then write. SharePoint rights stay — it is the same account.'}
              </div>
              {aliasHints.length > 0 && (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginBottom: 10 }}>
                  {aliasHints.map(a => (
                    <button key={a.email} type="button" className="btn btn-outline dex-ui-btn-sm"
                      onClick={() => {
                        const nm = (a.name || '').trim();
                        const parts = nm.indexOf(',') >= 0 ? nm.split(',').map(s => s.trim()) : [];
                        setRenameForm({ oldEmail: a.email, newEmail: a.spEmail, displayName: nm, first: parts[1] || '', last: parts[0] || '' });
                        setRenamePlan(null); setRenameResult(null);
                      }}>
                      {isDe ? 'Übernehmen: ' : 'Use: '}{a.name || a.email} → {a.spEmail}
                    </button>
                  ))}
                </div>
              )}
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 8, marginBottom: 10 }}>
                <input className="form-input" placeholder={isDe ? 'Alte E-Mail' : 'Old email'} value={renameForm.oldEmail} onChange={e => { setRenameForm(f => ({ ...f, oldEmail: e.target.value })); setRenamePlan(null); }} disabled={!!renameBusy} />
                <input className="form-input" placeholder={isDe ? 'Neue E-Mail' : 'New email'} value={renameForm.newEmail} onChange={e => { setRenameForm(f => ({ ...f, newEmail: e.target.value })); setRenamePlan(null); }} disabled={!!renameBusy} />
                <input className="form-input" placeholder={isDe ? 'Anzeigename (Nachname, Vorname)' : 'Display name (Last, First)'} value={renameForm.displayName} onChange={e => { setRenameForm(f => ({ ...f, displayName: e.target.value })); setRenamePlan(null); }} disabled={!!renameBusy} />
                <input className="form-input" placeholder={isDe ? 'Vorname' : 'First name'} value={renameForm.first} onChange={e => { setRenameForm(f => ({ ...f, first: e.target.value })); setRenamePlan(null); }} disabled={!!renameBusy} />
                <input className="form-input" placeholder={isDe ? 'Nachname (neu)' : 'Last name (new)'} value={renameForm.last} onChange={e => { setRenameForm(f => ({ ...f, last: e.target.value })); setRenamePlan(null); }} disabled={!!renameBusy} />
              </div>
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
                <button type="button" className="btn btn-secondary" disabled={!!renameBusy || !renameGueltig()} onClick={() => { void runRenamePlan(); }} style={{ fontSize: '0.82rem' }}>
                  {renameBusy === 'plan' ? (isDe ? `Vorschau … ${renameLabel}` : `Preview … ${renameLabel}`) : (isDe ? 'Vorschau (nichts wird geschrieben)' : 'Preview (nothing is written)')}
                </button>
                {renamePlan && (
                  <button type="button" className="btn btn-primary" disabled={!!renameBusy} onClick={() => { void runRename(); }} style={{ fontSize: '0.82rem' }}>
                    {renameBusy === 'run' ? (isDe ? `Schreibt … ${renameLabel}` : `Writing … ${renameLabel}`) : (isDe ? 'Jetzt umbenennen' : 'Rename now')}
                  </button>
                )}
              </div>
              {(renamePlan || renameResult) && (() => {
                const r = (renameResult || renamePlan) as PersonRenameResult;
                const zeile = (label: string, b: PersonRenameBereichLike): string =>
                  r.dryRun
                    ? `${label}: ${b.hits}${b.unreadable ? (isDe ? ` (${b.unreadable} nicht lesbar)` : ` (${b.unreadable} unreadable)`) : ''}`
                    : `${label}: ${b.written}/${b.hits}${b.failed ? (isDe ? `, ${b.failed} fehlgeschlagen` : `, ${b.failed} failed`) : ''}${b.unreadable ? (isDe ? `, ${b.unreadable} nicht lesbar` : `, ${b.unreadable} unreadable`) : ''}`;
                const fehler = r.roles.failed + r.eventFields.failed + r.eventPiggybacks.failed + r.registry.failed + r.registrations.failed + r.registeredBy.failed;
                return (
                  <div style={{ marginTop: 10, fontSize: '0.8rem', color: (!r.dryRun && fehler > 0) ? 'var(--dex-red, #c00)' : 'var(--dex-gray-700)', lineHeight: 1.5 }}>
                    <strong>{r.dryRun ? (isDe ? 'Vorschau — gefundene Stellen:' : 'Preview — places found:') : (isDe ? 'Umbenannt:' : 'Renamed:')}</strong>{' '}
                    {[
                      zeile(isDe ? 'Rollen' : 'Roles', r.roles),
                      zeile(isDe ? 'Event-Felder' : 'Event fields', r.eventFields),
                      zeile(isDe ? 'Teams in Events' : 'Teams in events', r.eventPiggybacks),
                      zeile(isDe ? 'Teilnehmer-Register' : 'Participant registry', r.registry),
                      zeile(isDe ? 'Anmeldungen' : 'Registrations', r.registrations),
                      zeile(isDe ? 'Registriert/Abgemeldet von' : 'Registered/cancelled by', r.registeredBy),
                    ].join(' · ')}
                    {r.eventTitles.length > 0 && <div>{isDe ? 'Events: ' : 'Events: '}{r.eventTitles.slice(0, 12).join(', ')}{r.eventTitles.length > 12 ? ` … (+${r.eventTitles.length - 12})` : ''}</div>}
                    {r.listSites.length > 0 && <div>{isDe ? 'Teilnehmerlisten: ' : 'Participant lists: '}{r.listSites.slice(0, 12).join(', ')}{r.listSites.length > 12 ? ` … (+${r.listSites.length - 12})` : ''}</div>}
                    {!r.dryRun && (
                      <div style={{ marginTop: 4 }}>
                        {isDe ? 'Betroffene Person bitte die App einmal neu laden. Danach „Rechte prüfen" ausführen — der Hinweis zur Schreibweise sollte weg sein.' : 'Ask the person to reload the app once. Then run “Check rights” — the spelling note should be gone.'}
                      </div>
                    )}
                  </div>
                );
              })()}
            </div>

            {/* Neue Rolle hinzufügen — v11.72: nach OBEN verschoben, damit der
                Admin nicht erst durch die Tabelle scrollen muss. Form ist auf
                EINEN People-Picker reduziert (analog zum Organizer-Picker im
                EventCreation-Wizard). */}
            {!showAddForm ? (
              <button
                className="btn btn-primary"
                onClick={() => setShowAddForm(true)}
                style={{ fontSize: '0.85rem', marginBottom: 16 }}
              >
                <Plus size={16} /> Add User Role
              </button>
            ) : (
              <div style={{
                marginBottom: 16, padding: 16, background: 'var(--dex-gray-50, #fafafa)',
                borderRadius: 'var(--dex-radius, 8px)', border: '1px solid var(--dex-gray-200, #eee)',
              }}>
                <h4 style={{ margin: '0 0 12px', fontSize: '0.9rem' }}>Assign New Role</h4>

                {/* People-Picker — Such-Input ODER ausgewählter Chip */}
                <div style={{ marginBottom: 12 }}>
                  <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 4 }}>
                    Person <span style={{ color: 'var(--dex-danger, red)' }}>*</span>
                  </label>
                  {userFound === true && newEmail && newName ? (
                    <div style={{
                      display: 'inline-flex', alignItems: 'center', gap: 10,
                      padding: '6px 10px 6px 6px', background: '#fff',
                      border: '1px solid var(--dex-green)', borderRadius: 999,
                      maxWidth: '100%',
                    }}>
                      <img
                        src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(newEmail)}&size=L`}
                        alt={newName}
                        onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                        style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                      />
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 600, fontSize: '0.85rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {newName}
                        </div>
                        <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.75rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {newEmail}{newLocation ? ` · ${newLocation}` : ''}
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={() => { setNewEmail(''); setNewName(''); setNewLocation(''); setUserFound(null); }}
                        title="Auswahl entfernen"
                        style={{
                          border: 'none', background: 'var(--dex-gray-100)', cursor: 'pointer',
                          width: 24, height: 24, borderRadius: '50%',
                          display: 'flex', alignItems: 'center', justifyContent: 'center',
                          color: 'var(--dex-gray-600)', flexShrink: 0,
                        }}
                      >
                        <X size={16} />
                      </button>
                    </div>
                  ) : (
                    <div style={{ position: 'relative' }}>
                      <input
                        className="form-input"
                        value={newEmail}
                        onChange={e => handleEmailChange(e.target.value)}
                        onFocus={() => { if (suggestions.length > 0) setShowSuggestions(true); }}
                        onBlur={() => setTimeout(() => setShowSuggestions(false), 200)}
                        placeholder="Name oder Email eingeben..."
                        style={{ fontSize: '0.85rem' }}
                        autoFocus
                        autoComplete="off"
                      />
                      {isSearching && (
                        <span style={{ position: 'absolute', right: 12, top: '50%', transform: 'translateY(-50%)', fontSize: '0.8rem', color: 'var(--dex-gray-400)' }}>
                          Suche...
                        </span>
                      )}
                      <div style={{ marginTop: 2 }}>
                        <InternationalSearchToggle query={newEmail} checked={includeIntl} onChange={setIncludeIntl} isDe={isDe} />
                      </div>
                      {/* v11.75: explizite „Keine Treffer"-Box wenn die Suche
                          fertig ist und 0 Treffer hat — sonst wirkt der Picker
                          stumm und der Admin weiß nicht, ob die Suche lief. */}
                      {!isSearching && newEmail && newEmail.length >= 2 && suggestions.length === 0 && (
                        <div style={{
                          marginTop: 6, padding: '8px 12px', borderRadius: 'var(--dex-radius)',
                          border: '1px dashed var(--dex-gray-300)', background: 'var(--dex-gray-50)',
                          color: 'var(--dex-gray-600)', fontSize: '0.78rem',
                        }}>
                          Keine Treffer für &bdquo;{newEmail}&ldquo;. Versuche es mit dem vollen Namen oder der E-Mail-Adresse.
                        </div>
                      )}
                      {showSuggestions && suggestions.length > 0 && (
                        <div style={{
                          position: 'absolute', top: '100%', left: 0, right: 0, zIndex: 200,
                          background: '#fff', border: '1px solid var(--dex-gray-200)', borderRadius: 'var(--dex-radius)',
                          boxShadow: 'var(--dex-shadow-hover)', maxHeight: 320, overflowY: 'auto', marginTop: 2,
                        }}>
                          {suggestions.map((s, i) => (
                            <div
                              key={i}
                              onMouseDown={() => selectSuggestion(s)}
                              style={{
                                display: 'flex', alignItems: 'center', gap: 10,
                                padding: '8px 12px', cursor: 'pointer',
                                borderBottom: i < suggestions.length - 1 ? '1px solid var(--dex-gray-100)' : 'none',
                              }}
                              onMouseEnter={e => { (e.currentTarget as HTMLDivElement).style.background = 'var(--dex-gray-100)'; }}
                              onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = '#fff'; }}
                            >
                              <img
                                src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(s.email)}&size=L`}
                                alt={s.displayName}
                                onError={e => { (e.currentTarget as HTMLImageElement).style.visibility = 'hidden'; }}
                                style={{ width: 32, height: 32, borderRadius: '50%', objectFit: 'cover', background: 'var(--dex-gray-100)', flexShrink: 0 }}
                              />
                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ fontWeight: 600, fontSize: '0.85rem' }}>{s.displayName}</div>
                                <div style={{ color: 'var(--dex-gray-500)', fontSize: '0.78rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                  {s.email}{s.location ? ` · ${s.location}` : ''}
                                </div>
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                </div>

                <div style={{ marginTop: 12 }}>
                  <label style={{ display: 'block', fontSize: '0.8rem', color: 'var(--dex-gray-500)', marginBottom: 4 }}>Role</label>
                  <select
                    className="form-select"
                    value={newRole}
                    onChange={e => setNewRole(e.target.value as UserRole)}
                    style={{ fontSize: '0.85rem', maxWidth: 200 }}
                  >
                    <option value="Organizer">Organizer</option>
                    <option value="Admin">Admin</option>
                    <option value="IT-Admin">IT-Admin (volle Rechte, keine Mails)</option>
                    {/* v30.5: F&A — Teilnehmer-Rechte + F&A Center, keine Organizer-/Admin-Rechte. */}
                    {/* v31.47: Die Beschriftung stammte aus v30.5, als F&A eine
                        reine Lese-Rolle war („Teilnehmer-Rechte plus das
                        Center"). Der Zuschnitt ist seit v30.60 ein anderer —
                        F&A kann alles, was ein Organizer kann, UND hat das F&A
                        Center. Der Code macht das längst so (`isOrganizer` und
                        `canCreateEvents` sind true, `grantOrganizerPermissions`
                        läuft, seit v31.42 auch die Onboarding-Mail); nur diese
                        Zeile behauptete weiter „nur F&A Center" — und wer eine
                        Rolle vergibt, liest genau sie. */}
                    <option value="F&A">F&A (Organizer + F&A Center)</option>
                  </select>
                </div>

                <div style={{ display: 'flex', gap: 8, marginTop: 16, justifyContent: 'flex-end' }}>
                  <button
                    className="btn btn-secondary"
                    onClick={() => { setShowAddForm(false); setNewEmail(''); setNewName(''); setNewLocation(''); setUserFound(null); }}
                    style={{ fontSize: '0.85rem' }}
                  >
                    Cancel
                  </button>
                  <button
                    className="btn btn-primary"
                    onClick={handleAddRole}
                    disabled={isAdding || !newEmail || !newName}
                    style={{ fontSize: '0.85rem' }}
                  >
                    {isAdding ? 'Saving...' : 'Assign Role'}
                  </button>
                </div>
              </div>
            )}

            {/* v24.87: Freitext-Suche über alle Kategorien. */}
            {!isRolesLoading && (
              <div style={{ position: 'relative', marginBottom: 14, maxWidth: 360 }}>
                <input
                  className="form-input"
                  value={roleSearch}
                  onChange={e => setRoleSearch(e.target.value)}
                  placeholder={isDe ? 'Suchen (Name, E-Mail, Position, Standort)…' : 'Search (name, email, position, location)…'}
                  style={{ fontSize: '0.85rem', paddingRight: roleSearch ? 30 : undefined }}
                />
                {roleSearch && (
                  <button
                    type="button"
                    onClick={() => setRoleSearch('')}
                    title={isDe ? 'Suche zurücksetzen' : 'Clear search'}
                    style={{ position: 'absolute', right: 6, top: '50%', transform: 'translateY(-50%)', border: 'none', background: 'var(--dex-gray-100)', cursor: 'pointer', width: 22, height: 22, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--dex-gray-600)' }}
                  >
                    <X size={14} />
                  </button>
                )}
              </div>
            )}

            {/* v24.85: Rollen je Kategorie als eigene aufklappbare Abschnitte. */}
            {isRolesLoading ? (
              <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic' }}>{isDe ? 'Rollen werden geladen…' : 'Loading roles...'}</p>
            ) : renderRoleSections()}

            </div>
          </div>
        )}

        {/* v31.86: Organizer sehen die Rollen — nur lesen. Nutzer-Ansage
            23.09.2026: „die Rollenliste, also welcher User bereits Organizer
            ist oder Check-in, für alle Organizer einsehbar — damit Leute, die
            schon Organizer sind und Co-Organizer für ein weiteres Event
            werden, keine Mails mehr erzeugen." Dieselben Abschnitte wie beim
            Admin (renderRoleSections mit readOnly), ohne Zuweisen, Rechte-
            Kästen und Dubletten-Hinweis. Ist DEX_Roles nicht lesbar, steht
            das hier statt einer leeren Tabelle. */}
        {!isAdmin && canCreateEvents && (
          <div className="card">
            <h2 style={{ margin: '0 0 12px', fontSize: '1.1rem', color: 'var(--dex-gray-800)' }}>
              {isDe ? 'Rollen — wer ist schon Organizer?' : 'Roles — who is already an organizer?'}
            </h2>
            <p style={{ color: 'var(--dex-gray-500, #888)', fontSize: '0.85rem', marginBottom: 12, lineHeight: 1.5 }}>
              {isDe
                ? 'Nur zum Nachsehen: Wer hier als Admin, Organizer oder F&A steht, kann Events bearbeiten und braucht keine Freigabe, wenn du sie als Co-Organizer benennst. Wer fehlt, bekommt beim Speichern automatisch einen Freigabe-Antrag bei den Admins. Check-in und Tester sind je Event vergeben und stehen hier nur zur Übersicht.'
                : 'Read-only: anyone listed as Admin, Organizer or F&A can edit events and needs no approval when you name them as co-organizer. Anyone missing gets an approval request to the admins automatically on save. Check-in and testers are assigned per event and are listed here for reference only.'}
            </p>
            {rolesReadStatus !== 'ok' ? (
              <div style={{ padding: '10px 14px', borderRadius: 8, fontSize: '0.85rem', background: '#fce4ec', color: '#c62828' }}>
                {isDe
                  ? 'Die Rollenliste konnte nicht gelesen werden. Bitte einen Admin bitten, in der Rollenverwaltung „Rechte prüfen" auszuführen.'
                  : 'The roles list could not be read. Please ask an admin to run "Check rights" in role management.'}
              </div>
            ) : (
              <>
                <input
                  type="text"
                  value={roleSearch}
                  onChange={e => setRoleSearch(e.target.value)}
                  placeholder={isDe ? 'Name, E-Mail, Position oder Standort …' : 'Name, email, position or location …'}
                  className="form-input"
                  style={{ marginBottom: 12, maxWidth: 420 }}
                />
                {renderRoleSections()}
              </>
            )}
          </div>
        )}

        {/* v9.16/v9.21: Test-Team war hier global, ist jetzt per-Event
            (im EventCreation-Wizard). TestTeamManager entfernt. */}

        {/* v24.86: „Default-Mail-Templates re-seed" und „Wochenbericht — Test-
            Versand" sind ins Admin Center (Aktionen-Dropdown) gewandert.
            v24.84: „Listen-Berechtigungen"-Übersicht entfernt. */}

      </div>

      {/* Onboarding-Mail-Prompt nach erfolgreicher Rollen-Zuweisung */}
      {/* v13.4: Onboarding-Prompt jetzt über das <Modal>-Wrapper-Component. */}
      <Modal
        open={!!onboardingPrompt}
        onClose={() => setOnboardingPrompt(null)}
        dismissable={!isSendingOnboarding}
        maxWidth={460}
        ariaLabel="Onboarding-Mail senden"
      >
        {onboardingPrompt && (
          <>
            <h3 style={{ margin: '0 0 12px', fontSize: '1.05rem' }}>
              Onboarding-Mail an {onboardingPrompt.name}?
            </h3>
            <p style={{ margin: '0 0 12px', fontSize: '0.9rem', lineHeight: 1.5, color: 'var(--dex-gray-700)' }}>
              <strong>{onboardingPrompt.name}</strong> wurde als <strong>{onboardingPrompt.role}</strong> hinzugefügt.
              Möchtest du eine Begrüßungsmail mit Link zur App, zum Handbuch und einer kurzen
              Anleitung zum ersten Test-Event verschicken?
            </p>
            <p style={{ margin: '0 0 20px', fontSize: '0.8rem', color: 'var(--dex-gray-500)' }}>
              Empfänger: {onboardingPrompt.email}<br />
              Cc: dex.event@deloitte.de
            </p>
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
              <button
                className="btn btn-secondary"
                onClick={() => setOnboardingPrompt(null)}
                disabled={isSendingOnboarding}
                style={{ fontSize: '0.85rem' }}
              >
                Nicht senden
              </button>
              <button
                className="btn btn-primary"
                onClick={handleSendOnboarding}
                disabled={isSendingOnboarding}
                style={{ fontSize: '0.85rem' }}
              >
                {isSendingOnboarding ? 'Sende...' : 'Mail senden'}
              </button>
            </div>
          </>
        )}
      </Modal>
    </div>
  );
}

