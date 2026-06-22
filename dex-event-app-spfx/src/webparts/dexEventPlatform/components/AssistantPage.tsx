// v24.36: Assistenz-Seite — Verwaltung von Anmeldungen, die der eingeloggte
// User STELLVERTRETEND für andere Personen durchgeführt hat (einsehen,
// Angaben anpassen, ab-/anmelden bei Sub-Events). Sichtbar nur für „Assistenz"
// (wer mind. eine Fremd-Anmeldung getätigt hat) und Admins. Filter läuft über
// `RegisteredByEmail = ich` UND `ParticipantEmail ≠ ich`.

import * as React from 'react';
import { useEvents } from '../context/EventContext';
import { useNavigation } from '../context/NavigationContext';
import { useLanguage } from '../context/LanguageContext';
import { useDialog } from '../context/DialogContext';
import { useRoles } from '../context/RoleContext';
import { UserFieldPicker } from './UserFieldPicker';
import { DeloitteEvent, EventSpecificField } from '../types';
import { SPRegistration } from '../services/EventService';
import { isEventOver } from '../utils/eventFormat';
import { CachedImg } from './CachedImage';
import Modal from './Modal';
import { Icon } from '@fluentui/react/lib/Icon';

interface ProxyItem { event: DeloitteEvent; registration: SPRegistration; }

interface Group {
  pEmail: string;
  pName: string;
  pFirst: string;
  pLast: string;
  pLocation: string;
  pCompany: string;
  rootEvent: DeloitteEvent;
  // eventId → ProxyItem (Haupt- + Sub-Event-Anmeldungen dieser Person)
  regs: Map<string, ProxyItem>;
}

const ACTIVE_STATUSES = ['Angemeldet', 'QR versendet', 'Eingecheckt', 'Warteliste'];

function fmtDate(iso?: string, de = true): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleDateString(de ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

// showIf-Sichtbarkeit eines Feldes anhand der aktuellen Antworten prüfen.
function fieldVisible(f: EventSpecificField, data: Record<string, string>): boolean {
  if (!f.showIf || !f.showIf.fieldId) return true;
  const raw = (data[f.showIf.fieldId] || '').trim();
  if (!raw) return false;
  const answers = raw.indexOf(' | ') >= 0 ? raw.split(' | ').map(s => s.trim()) : [raw];
  return answers.some(a => f.showIf!.values.indexOf(a) >= 0);
}

export default function AssistantPage(): React.ReactElement {
  const { events, getMyProxyRegistrations, cancelProxyRegistration, updateProxyRegistration, registerForEvent, childEventsOf } = useEvents();
  const { navigate } = useNavigation();
  const { locale } = useLanguage();
  const { confirmDialog, showAlert } = useDialog();
  const { searchUsers, searchUser } = useRoles();
  const isDe = locale === 'de';

  const [loading, setLoading] = React.useState(true);
  const [groups, setGroups] = React.useState<Group[]>([]);
  const [busyKey, setBusyKey] = React.useState<string | null>(null);

  // Edit-/Register-Modal-State.
  const [fieldModal, setFieldModal] = React.useState<{
    mode: 'edit' | 'register';
    event: DeloitteEvent;          // Ziel-Event (Haupt- oder Sub-Event)
    group: Group;
    registration?: SPRegistration; // nur im edit-Modus
    data: Record<string, string>;
  } | null>(null);
  const [modalBusy, setModalBusy] = React.useState(false);

  // Vollständige Event-Lookup-Map (Haupt- UND Sub-Events) — damit das Parent-/
  // Klammer-Event eines Sub-Events sicher aufgelöst wird, auch wenn die
  // Assistenz NUR ein Sub-Event (nicht das Hauptevent) gebucht hat.
  const eventById = React.useMemo(() => {
    const m = new Map<string, DeloitteEvent>();
    for (const e of events) m.set(e.id, e);
    return m;
  }, [events]);

  const buildGroups = React.useCallback((items: ProxyItem[]): Group[] => {
    const map = new Map<string, Group>();
    for (const it of items) {
      const ev = it.event;
      const rootId = ev.parentEventId || ev.id;
      const pEmail = (it.registration.ParticipantEmail || '').toLowerCase();
      const key = `${pEmail}__${rootId}`;
      let g = map.get(key);
      if (!g) {
        // root-Event = das Klammer-/Hauptevent. Über die volle Event-Map
        // auflösen (nicht nur über die Treffer-Liste — sonst würde ein
        // Sub-Event ohne Parent-Anmeldung fälschlich selbst zum root).
        const rootEvent = eventById.get(rootId) || ev;
        g = {
          pEmail,
          pName: it.registration.ParticipantName || `${it.registration.Vorname || ''} ${it.registration.Nachname || ''}`.trim() || pEmail,
          pFirst: it.registration.Vorname || '',
          pLast: it.registration.Nachname || '',
          pLocation: it.registration.Location || '',
          pCompany: it.registration.Company || '',
          rootEvent,
          regs: new Map(),
        };
        map.set(key, g);
      }
      g.regs.set(ev.id, it);
    }
    // Gruppen sortieren: nach Teilnehmer-Name, dann Event-Titel.
    return Array.from(map.values()).sort((a, b) =>
      a.pName.localeCompare(b.pName) || a.rootEvent.title.localeCompare(b.rootEvent.title));
  }, [eventById]);

  const reload = React.useCallback(async () => {
    setLoading(true);
    try {
      const items = await getMyProxyRegistrations();
      setGroups(buildGroups(items));
    } catch {
      setGroups([]);
    } finally {
      setLoading(false);
    }
  }, [getMyProxyRegistrations, buildGroups]);

  React.useEffect(() => { void reload(); }, [reload]);

  const statusLabel = (s: string): string => {
    if (s === 'Angemeldet') return isDe ? 'Angemeldet' : 'Registered';
    if (s === 'Warteliste') return isDe ? 'Warteliste' : 'Waitlist';
    if (s === 'QR versendet') return isDe ? 'QR versendet' : 'QR sent';
    if (s === 'Eingecheckt') return isDe ? 'Eingecheckt' : 'Checked in';
    if (s === 'Abgemeldet') return isDe ? 'Abgemeldet' : 'Cancelled';
    return s;
  };
  const statusColor = (s: string): string => {
    if (s === 'Warteliste') return '#ed8b00';
    if (s === 'Abgemeldet') return '#999';
    return 'var(--dex-green, #86bc25)';
  };

  async function doCancel(group: Group, item: ProxyItem): Promise<void> {
    const ev = item.event;
    if (isEventOver(ev)) {
      await showAlert(isDe ? 'Dieses Event ist vorbei — eine Abmeldung ist nicht mehr möglich.' : 'This event is over — cancellation is no longer possible.');
      return;
    }
    const ok = await confirmDialog(
      isDe
        ? `${group.pName} von „${ev.title}" abmelden? Die Person erhält eine Abmelde-Bestätigung.`
        : `Cancel ${group.pName} from „${ev.title}"? The person will receive a cancellation confirmation.`,
      { danger: true, confirmLabel: isDe ? 'Abmelden' : 'Cancel registration' }
    );
    if (!ok) return;
    const key = `cancel_${group.pEmail}_${ev.id}`;
    setBusyKey(key);
    try {
      const success = await cancelProxyRegistration(ev.id, item.registration);
      if (success) {
        await showAlert(isDe ? 'Abmeldung durchgeführt.' : 'Cancellation done.', { variant: 'success' });
        await reload();
      } else {
        await showAlert(isDe ? 'Abmeldung fehlgeschlagen.' : 'Cancellation failed.', { variant: 'error' });
      }
    } finally {
      setBusyKey(null);
    }
  }

  function openEdit(group: Group, item: ProxyItem): void {
    let data: Record<string, string> = {};
    try { if (item.registration.CustomData) data = JSON.parse(item.registration.CustomData); } catch { /* */ }
    setFieldModal({ mode: 'edit', event: item.event, group, registration: item.registration, data });
  }

  function openRegister(group: Group, childEvent: DeloitteEvent): void {
    setFieldModal({ mode: 'register', event: childEvent, group, data: {} });
  }

  async function submitFieldModal(): Promise<void> {
    if (!fieldModal) return;
    const { mode, event, group, registration, data } = fieldModal;
    // Pflichtfelder prüfen (nur sichtbare).
    const fields = (event.eventSpecificFields || []).filter(f => f.label);
    for (const f of fields) {
      if (!f.required) continue;
      if (!fieldVisible(f, data)) continue;
      if (f.type === 'document') continue; // Dokument-Upload hier nicht bearbeitbar
      const v = (data[f.id] || '').trim();
      if (!v || (f.type === 'checkbox' && v !== 'true')) {
        await showAlert(isDe ? `Bitte das Pflichtfeld „${f.label}" ausfüllen.` : `Please fill the required field „${f.label}".`);
        return;
      }
    }
    setModalBusy(true);
    try {
      if (mode === 'edit' && registration) {
        const ok = await updateProxyRegistration(event.id, registration, data);
        if (ok) {
          await showAlert(isDe ? 'Angaben gespeichert.' : 'Details saved.', { variant: 'success' });
          setFieldModal(null);
          await reload();
        } else {
          await showAlert(isDe ? 'Speichern fehlgeschlagen.' : 'Saving failed.', { variant: 'error' });
        }
      } else if (mode === 'register') {
        if (isEventOver(event)) {
          await showAlert(isDe ? 'Dieses Sub-Event ist vorbei — eine Anmeldung ist nicht mehr möglich.' : 'This sub-event is over.');
          setModalBusy(false);
          return;
        }
        const res = await registerForEvent(
          event.id, data, group.pFirst, group.pLast, group.pEmail, undefined,
          { proxyConsentConfirmed: true, actorAllowedAsAssistant: true }
        );
        if (res.ok) {
          await showAlert(
            res.status === 'Warteliste'
              ? (isDe ? `${group.pName} wurde auf die Warteliste gesetzt.` : `${group.pName} was put on the waitlist.`)
              : (isDe ? `${group.pName} wurde angemeldet.` : `${group.pName} was registered.`),
            { variant: 'success' }
          );
          setFieldModal(null);
          await reload();
        } else {
          await showAlert(isDe ? 'Anmeldung fehlgeschlagen.' : 'Registration failed.', { variant: 'error' });
        }
      }
    } finally {
      setModalBusy(false);
    }
  }

  const renderField = (f: EventSpecificField): React.ReactElement | null => {
    if (!fieldModal) return null;
    if (!fieldVisible(f, fieldModal.data)) return null;
    const setVal = (v: string): void => setFieldModal(prev => prev ? { ...prev, data: { ...prev.data, [f.id]: v } } : prev);
    const val = fieldModal.data[f.id] || '';
    const labelEl = (
      <label style={{ display: 'block', fontWeight: 600, fontSize: '0.85rem', marginBottom: 4 }}>
        {f.label}{f.required ? ' *' : ''}
      </label>
    );
    const baseInput: React.CSSProperties = {
      width: '100%', padding: '9px 11px', borderRadius: 8,
      border: `1px solid ${val ? 'var(--dex-green, #86bc25)' : '#ccc'}`,
      background: val ? 'rgba(134,188,37,0.06)' : '#fff', fontSize: '0.9rem', boxSizing: 'border-box',
    };
    if (f.type === 'user' || f.type === 'roommate') {
      return (
        <div key={f.id} style={{ marginBottom: 14 }}>
          {labelEl}
          {f.helpText && f.helpTextStyle === 'inline' && (
            <div style={{ fontSize: '0.78rem', color: '#777', marginBottom: 5 }}>{f.helpText}</div>
          )}
          <UserFieldPicker
            value={val}
            onChange={setVal}
            searchUsers={searchUsers}
            searchUserByEmail={searchUser}
            placeholder={isDe ? 'Name oder E-Mail eingeben…' : 'Type a name or email…'}
            errorStyle={{}}
            forcedIsDe={isDe}
          />
        </div>
      );
    }
    if (f.type === 'document') {
      return (
        <div key={f.id} style={{ marginBottom: 14 }}>
          {labelEl}
          <div style={{ fontSize: '0.8rem', color: '#888', fontStyle: 'italic' }}>
            {isDe ? 'Dokument-Uploads kannst du hier nicht bearbeiten — die Person erledigt das über „Meine Events".' : 'Document uploads cannot be edited here — the person handles that via „My Events".'}
          </div>
        </div>
      );
    }
    let control: React.ReactElement;
    if (f.type === 'checkbox') {
      control = (
        <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.9rem' }}>
          <input type="checkbox" checked={val === 'true'} onChange={e => setVal(e.target.checked ? 'true' : 'false')} />
          <span>{f.confirmLabel || (isDe ? 'Ja, bestätigen' : 'Yes, confirm')}</span>
        </label>
      );
    } else if (f.type === 'select') {
      if (f.multi) {
        const selected = val ? val.split(' | ').map(s => s.trim()) : [];
        control = (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
            {(f.options || []).map(opt => {
              const on = selected.indexOf(opt) >= 0;
              return (
                <label key={opt} style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: '0.9rem' }}>
                  <input type="checkbox" checked={on} onChange={() => {
                    const next = on ? selected.filter(s => s !== opt) : [...selected, opt];
                    setVal(next.join(' | '));
                  }} />
                  <span>{opt}</span>
                </label>
              );
            })}
          </div>
        );
      } else {
        control = (
          <select value={val} onChange={e => setVal(e.target.value)} style={baseInput}>
            <option value="">{isDe ? '— bitte wählen —' : '— please select —'}</option>
            {(f.options || []).map(opt => <option key={opt} value={opt}>{opt}</option>)}
          </select>
        );
      }
    } else if (f.type === 'date') {
      control = <input type={f.withTime ? 'datetime-local' : 'date'} value={val} onChange={e => setVal(e.target.value)} style={baseInput} />;
    } else if (f.type === 'number') {
      control = <input type="number" value={val} onChange={e => setVal(e.target.value)} style={baseInput} />;
    } else {
      control = <input type="text" value={val} onChange={e => setVal(e.target.value)} style={baseInput} />;
    }
    return (
      <div key={f.id} style={{ marginBottom: 14 }}>
        {labelEl}
        {f.helpText && f.helpTextStyle === 'inline' && (
          <div style={{ fontSize: '0.78rem', color: '#777', marginBottom: 5 }}>{f.helpText}</div>
        )}
        {control}
      </div>
    );
  };

  // v24.37: Vergangene Events ausblenden — nach Event-Ende sind Angaben-
  // Anpassung / An-/Abmeldung ohnehin gesperrt (Aufbewahrungs-/Löschkonzept).
  // Eine Gruppe bleibt sichtbar, solange das Klammer-/Hauptevent ODER mind. ein
  // Sub-Event noch nicht vorbei ist (dort kann die Assistenz noch handeln).
  const visibleGroups = React.useMemo(() => groups.filter(g => {
    if (!isEventOver(g.rootEvent)) return true;
    return childEventsOf(g.rootEvent.id).some(c => !isEventOver(c));
  }), [groups, childEventsOf]);

  return (
    <div className="page-container">
      <button
        type="button"
        onClick={() => navigate('start')}
        style={{ background: 'none', border: 'none', color: 'var(--dex-green, #86bc25)', cursor: 'pointer', fontSize: '0.9rem', fontWeight: 600, marginBottom: 12, display: 'inline-flex', alignItems: 'center', gap: 6 }}
      >
        <Icon iconName="Back" /> {isDe ? 'Zurück' : 'Back'}
      </button>

      <h1 style={{ marginTop: 0, display: 'flex', alignItems: 'center', gap: 10 }}>
        <Icon iconName="People" style={{ fontSize: 28, color: 'var(--dex-green, #86bc25)' }} />
        {isDe ? 'Assistenz' : 'Assistant'}
      </h1>
      <p style={{ color: '#666', marginTop: -4, maxWidth: 720 }}>
        {isDe
          ? 'Hier verwaltest du Anmeldungen, die du stellvertretend für andere Personen durchgeführt hast: Angaben einsehen und anpassen, bei Sub-Events an- und abmelden.'
          : 'Manage registrations you made on behalf of other people: view and edit details, register or cancel for sub-events.'}
      </p>

      {/* Hinweis-Box: nur selbst durchgeführte Anmeldungen sind sichtbar. */}
      <div style={{
        display: 'flex', gap: 10, alignItems: 'flex-start',
        background: 'rgba(237,139,0,0.08)', border: '1px solid var(--dex-orange, #ed8b00)',
        borderRadius: 10, padding: '12px 14px', maxWidth: 760, margin: '4px 0 8px',
      }}>
        <Icon iconName="Info" style={{ fontSize: 18, color: 'var(--dex-orange, #ed8b00)', marginTop: 1, flexShrink: 0 }} />
        <div style={{ fontSize: '0.84rem', color: '#5a4a20', lineHeight: 1.5 }}>
          {isDe ? (
            <>
              <strong>Wichtig:</strong> Du siehst hier ausschließlich Anmeldungen, die du <strong>selbst stellvertretend</strong> für eine andere Person durchgeführt hast. Hat sich die Person (z.B. dein Partner oder Director) <strong>selbst angemeldet</strong>, kannst du diese Anmeldung hier <strong>aktuell nicht</strong> einsehen oder bearbeiten — in dem Fall muss die Person ihre Anmeldung selbst über „Meine Events“ verwalten.
            </>
          ) : (
            <>
              <strong>Important:</strong> You only see registrations you made <strong>on behalf of</strong> another person yourself. If the person (e.g. your partner or director) <strong>registered themselves</strong>, you currently <strong>cannot</strong> view or manage that registration here — in that case the person needs to manage it themselves via „My events“.
            </>
          )}
        </div>
      </div>

      {loading ? (
        <div style={{ padding: 48, textAlign: 'center', color: '#999' }}>{isDe ? 'Lade …' : 'Loading …'}</div>
      ) : visibleGroups.length === 0 ? (
        <div style={{ padding: 40, textAlign: 'center', color: '#888', background: '#f7f7f7', borderRadius: 12, marginTop: 16 }}>
          <Icon iconName="People" style={{ fontSize: 40, color: '#ccc', display: 'block', marginBottom: 10 }} />
          {groups.length === 0
            ? (isDe
              ? 'Du hast aktuell keine Anmeldungen für andere Personen durchgeführt.'
              : 'You currently have no registrations made on behalf of others.')
            : (isDe
              ? 'Du hast nur stellvertretende Anmeldungen für bereits vergangene Events — die lassen sich nicht mehr anpassen.'
              : 'You only have registrations on behalf of others for past events — those can no longer be changed.')}
        </div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 18, marginTop: 16 }}>
          {visibleGroups.map(group => {
            const root = group.rootEvent;
            const children = childEventsOf(root.id);
            const rootItem = group.regs.get(root.id);
            // Alle Sub-Events zeigen (mit Anmeldung = verwalten, ohne = anmelden).
            const subItems = children.map(c => ({ child: c, item: group.regs.get(c.id) }));
            const subLine = [group.pLocation, group.pCompany].filter(Boolean).join(' · ');
            return (
              <div key={`${group.pEmail}__${root.id}`} className="card" style={{ padding: 0, overflow: 'hidden' }}>
                {/* Kopf: Teilnehmer + Event */}
                <div style={{ display: 'flex', gap: 14, alignItems: 'center', padding: '16px 18px', borderBottom: '1px solid #eee', flexWrap: 'wrap' }}>
                  {root.imageUrl && (
                    <CachedImg
                      src={root.imageUrl}
                      alt=""
                      style={{ width: 64, height: 64, borderRadius: 10, objectFit: 'cover', flexShrink: 0 }}
                    />
                  )}
                  <div style={{ flex: 1, minWidth: 200 }}>
                    <div style={{ fontWeight: 700, fontSize: '1.05rem' }}>{root.title}</div>
                    <div style={{ fontSize: '0.85rem', color: '#555', marginTop: 2 }}>
                      {fmtDate(root.startDate, isDe)}{root.location ? ` · ${root.location}` : ''}
                    </div>
                  </div>
                  <div style={{ background: '#f3f7ec', borderRadius: 10, padding: '8px 12px', minWidth: 180 }}>
                    <div style={{ fontSize: '0.7rem', textTransform: 'uppercase', letterSpacing: 0.5, color: '#86bc25', fontWeight: 700 }}>
                      {isDe ? 'Angemeldete Person' : 'Registered person'}
                    </div>
                    <div style={{ fontWeight: 600, fontSize: '0.92rem' }}>{group.pName}</div>
                    <div style={{ fontSize: '0.78rem', color: '#666' }}>{group.pEmail}</div>
                    {subLine && <div style={{ fontSize: '0.78rem', color: '#888', marginTop: 2 }}>{subLine}</div>}
                  </div>
                </div>

                {/* Haupt-Event-Anmeldung */}
                <div style={{ padding: '12px 18px' }}>
                  {rootItem ? (
                    <RegRow
                      title={isDe ? 'Hauptanmeldung' : 'Main registration'}
                      eventTitle={root.title}
                      item={rootItem}
                      group={group}
                      isDe={isDe}
                      statusLabel={statusLabel}
                      statusColor={statusColor}
                      busyKey={busyKey}
                      onEdit={() => openEdit(group, rootItem)}
                      onCancel={() => doCancel(group, rootItem)}
                    />
                  ) : (
                    <div style={{ fontSize: '0.85rem', color: '#888', fontStyle: 'italic' }}>
                      {isDe ? 'Keine eigene Anmeldung auf das Hauptevent (nur Sub-Events).' : 'No own registration on the main event (sub-events only).'}
                    </div>
                  )}
                </div>

                {/* Sub-Events */}
                {children.length > 0 && (
                  <div style={{ padding: '6px 18px 16px', borderTop: '1px solid #f0f0f0' }}>
                    <div style={{ fontSize: '0.8rem', fontWeight: 700, color: '#666', textTransform: 'uppercase', letterSpacing: 0.5, margin: '8px 0' }}>
                      {isDe ? 'Sub-Events' : 'Sub-events'}
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                      {subItems.map(({ child, item }) => (
                        <RegRow
                          key={child.id}
                          title={child.title.indexOf('|') >= 0 ? child.title.split('|').pop()!.trim() : child.title}
                          eventTitle={child.title}
                          item={item}
                          group={group}
                          isDe={isDe}
                          statusLabel={statusLabel}
                          statusColor={statusColor}
                          busyKey={busyKey}
                          childEvent={child}
                          onEdit={item ? () => openEdit(group, item) : undefined}
                          onCancel={item ? () => doCancel(group, item) : undefined}
                          onRegister={!item ? () => openRegister(group, child) : undefined}
                        />
                      ))}
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {/* Felder-Modal (Angaben anpassen / Sub-Event anmelden) */}
      {fieldModal && (
        <Modal open onClose={() => { if (!modalBusy) setFieldModal(null); }} maxWidth={560} dismissable={!modalBusy} ariaLabel="Angaben">
          <h2 style={{ marginTop: 0, fontSize: '1.2rem' }}>
            {fieldModal.mode === 'register'
              ? (isDe ? `Für „${fieldModal.event.title.indexOf('|') >= 0 ? fieldModal.event.title.split('|').pop()!.trim() : fieldModal.event.title}" anmelden` : `Register for „${fieldModal.event.title}"`)
              : (isDe ? 'Angaben anpassen' : 'Edit details')}
          </h2>
          <div style={{ fontSize: '0.85rem', color: '#666', marginBottom: 14 }}>
            {isDe ? 'Person: ' : 'Person: '}<strong>{fieldModal.group.pName}</strong> · {fieldModal.group.pEmail}
          </div>
          <div style={{ maxHeight: '55vh', overflowY: 'auto', paddingRight: 4 }}>
            {(fieldModal.event.eventSpecificFields || []).filter(f => f.label).length === 0 ? (
              <div style={{ color: '#888', fontStyle: 'italic', fontSize: '0.9rem' }}>
                {isDe ? 'Dieses Event hat keine zusätzlichen Abfragen.' : 'This event has no additional questions.'}
              </div>
            ) : (
              (fieldModal.event.eventSpecificFields || []).filter(f => f.label).map(renderField)
            )}
          </div>
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 18 }}>
            <button
              type="button"
              onClick={() => { if (!modalBusy) setFieldModal(null); }}
              disabled={modalBusy}
              style={{ padding: '9px 16px', borderRadius: 8, border: '1px solid #ccc', background: '#fff', cursor: 'pointer', fontWeight: 600 }}
            >
              {isDe ? 'Abbrechen' : 'Cancel'}
            </button>
            <button
              type="button"
              onClick={submitFieldModal}
              disabled={modalBusy}
              style={{ padding: '9px 18px', borderRadius: 8, border: 'none', background: 'var(--dex-green, #86bc25)', color: '#fff', cursor: 'pointer', fontWeight: 700 }}
            >
              {modalBusy ? '…' : fieldModal.mode === 'register' ? (isDe ? 'Anmelden' : 'Register') : (isDe ? 'Speichern' : 'Save')}
            </button>
          </div>
        </Modal>
      )}
    </div>
  );
}

interface RegRowProps {
  title: string;
  eventTitle: string;
  item?: ProxyItem;
  group: Group;
  isDe: boolean;
  statusLabel: (s: string) => string;
  statusColor: (s: string) => string;
  busyKey: string | null;
  childEvent?: DeloitteEvent;
  onEdit?: () => void;
  onCancel?: () => void;
  onRegister?: () => void;
}

function RegRow(props: RegRowProps): React.ReactElement {
  const { title, item, group, isDe, statusLabel, statusColor, busyKey, childEvent, onEdit, onCancel, onRegister } = props;
  const ev = item?.event || childEvent;
  const over = ev ? isEventOver(ev) : false;
  const status = item?.registration.Status || '';
  const active = item && ACTIVE_STATUSES.indexOf(status) >= 0;
  const hasFields = ev ? (ev.eventSpecificFields || []).some(f => f.label) : false;
  const cancelBusy = busyKey === `cancel_${group.pEmail}_${ev?.id || ''}`;
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap', padding: '8px 0' }}>
      <div style={{ flex: 1, minWidth: 160 }}>
        <div style={{ fontWeight: 600, fontSize: '0.92rem' }}>{title}</div>
        {item ? (
          <span style={{ display: 'inline-block', marginTop: 3, fontSize: '0.74rem', fontWeight: 700, color: '#fff', background: statusColor(status), borderRadius: 999, padding: '2px 10px' }}>
            {statusLabel(status)}
          </span>
        ) : (
          <span style={{ fontSize: '0.78rem', color: '#999' }}>{isDe ? 'Nicht angemeldet' : 'Not registered'}</span>
        )}
        {over && <span style={{ fontSize: '0.74rem', color: '#999', marginLeft: 8 }}>{isDe ? '(Event vorbei)' : '(event over)'}</span>}
      </div>
      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
        {item && active && hasFields && onEdit && !over && (
          <button type="button" onClick={onEdit} style={btnSecondary}>
            <Icon iconName="Edit" style={{ fontSize: 13, marginRight: 5 }} />{isDe ? 'Angaben anpassen' : 'Edit details'}
          </button>
        )}
        {item && active && onCancel && !over && (
          <button type="button" onClick={onCancel} disabled={cancelBusy} style={btnDanger}>
            {cancelBusy ? '…' : (isDe ? 'Abmelden' : 'Cancel')}
          </button>
        )}
        {!item && onRegister && !over && (
          <button type="button" onClick={onRegister} style={btnPrimary}>
            <Icon iconName="Add" style={{ fontSize: 13, marginRight: 5 }} />{isDe ? 'Anmelden' : 'Register'}
          </button>
        )}
      </div>
    </div>
  );
}

const btnSecondary: React.CSSProperties = {
  padding: '7px 13px', borderRadius: 8, border: '1px solid #ccc', background: '#fff',
  cursor: 'pointer', fontWeight: 600, fontSize: '0.82rem', whiteSpace: 'nowrap',
};
const btnPrimary: React.CSSProperties = {
  padding: '7px 13px', borderRadius: 8, border: 'none', background: 'var(--dex-green, #86bc25)',
  color: '#fff', cursor: 'pointer', fontWeight: 700, fontSize: '0.82rem', whiteSpace: 'nowrap',
};
const btnDanger: React.CSSProperties = {
  padding: '7px 13px', borderRadius: 8, border: '1px solid #d64545', background: '#fff',
  color: '#d64545', cursor: 'pointer', fontWeight: 700, fontSize: '0.82rem', whiteSpace: 'nowrap',
};
