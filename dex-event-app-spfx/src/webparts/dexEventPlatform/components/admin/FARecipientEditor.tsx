/**
 * v30.45: Empfänger-Verwaltung im F&A Center (Fachkonzept „F&A Center —
 * Verwaltung der Empfänger").
 *
 * Vorher: eine `<textarea>`, eine Adresse pro Zeile. Das war die einzige
 * Stelle in der App, an der Personen als roher Text standen — überall sonst
 * (Organizer, Co-Organizer, Assistenz, Teilnehmer) stehen sie als Profil mit
 * Foto. Für die Person, die den Verteiler pflegt, ist der Unterschied nicht
 * kosmetisch: Bei `m.mustermann@deloitte.de` sieht man nicht, ob das die
 * richtige Person ist; bei Foto, Name und Job Title sieht man es sofort.
 *
 * **Das gespeicherte Format ändert sich NICHT.** `FAConfig.infoRecipients` /
 * `listRecipients` bleiben `string[]` mit nackten E-Mail-Adressen — genau das,
 * was `sendFAMail` in die Mail-Queue schreibt. Die Profil-Darstellung ist reine
 * Anzeige und wird zur Laufzeit über `searchUserByEmail` aufgelöst. Hätte man
 * stattdessen `Name <email>` gespeichert, müsste der Versandpfad das wieder
 * auseinandernehmen — ein Formatwechsel auf dem Weg zur Mail ist genau die
 * Sorte Änderung, die man erst bemerkt, wenn eine Abrechnung nicht ankommt.
 *
 * Die Unterscheidung Person vs. Gruppenadresse fällt damit ebenfalls zur
 * Laufzeit: Was sich zu einem Profil auflösen lässt, ist eine Person; alles
 * andere (Funktionspostfächer wie `fa-abrechnung@deloitte.de`) bleibt ein
 * E-Mail-Eintrag mit voller Adresse. Das Konzept verlangt genau diese zwei
 * Darstellungen — und es ist keine Raterei, sondern eine Auskunft des
 * Verzeichnisses.
 */
import * as React from 'react';
import { UserFieldPicker } from '../UserFieldPicker';
import { X, Plus } from '../Icons';

type Profile = { displayName: string; location: string; jobTitle: string };

export interface FARecipientEditorProps {
  label: string;
  hint?: string;
  value: string[];
  onChange: (next: string[]) => void;
  searchUsers: (q: string, includeIntl?: boolean) => Promise<Array<{ email: string; displayName: string; location: string; jobTitle: string }>>;
  searchUserByEmail: (email: string) => Promise<Profile | null>;
  disabled?: boolean;
}

const chipBase: React.CSSProperties = {
  display: 'inline-flex', alignItems: 'center', gap: 8,
  border: '1px solid var(--dex-gray-200, #e5e7eb)', borderRadius: 999,
  padding: '4px 6px 4px 4px', background: '#fff', maxWidth: '100%',
};

export default function FARecipientEditor(props: FARecipientEditorProps): React.ReactElement {
  const { value, onChange, searchUsers, searchUserByEmail } = props;
  // Aufgelöste Profile je Adresse. `null` = geprüft und KEINE Person
  // (Gruppenadresse); fehlender Schlüssel = noch nicht geprüft.
  const [profiles, setProfiles] = React.useState<Record<string, Profile | null>>({});
  const [pickerValue, setPickerValue] = React.useState('');
  const [groupInput, setGroupInput] = React.useState('');
  const [groupError, setGroupError] = React.useState('');
  const [removeHover, setRemoveHover] = React.useState<string | null>(null);

  // Profile nacheinander auflösen, nicht parallel: Ein Verteiler hat eine
  // Handvoll Einträge, und die Suche läuft gegen dieselbe Schnittstelle, die
  // auch der People-Picker benutzt — es gibt keinen Grund, sie zu bündeln.
  React.useEffect(() => {
    let cancelled = false;
    const todo = value.filter(a => !(a in profiles));
    if (todo.length === 0) return undefined;
    (async () => {
      for (const addr of todo) {
        let p: Profile | null = null;
        try { p = await searchUserByEmail(addr); } catch { p = null; }
        if (cancelled) return;
        setProfiles(prev => ({ ...prev, [addr]: p && p.displayName ? p : null }));
      }
    })().catch(() => { /* best-effort — dann eben als E-Mail-Eintrag */ });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value.join('|')]);

  const addAddress = (raw: string): boolean => {
    const addr = (raw || '').trim().toLowerCase();
    if (!addr || addr.indexOf('@') < 0 || addr.indexOf('.') < 0) {
      setGroupError('Bitte eine vollständige E-Mail-Adresse eingeben.');
      return false;
    }
    if (value.some(v => v.toLowerCase() === addr)) {
      setGroupError('Diese Adresse steht bereits im Verteiler.');
      return false;
    }
    setGroupError('');
    onChange(value.concat(addr));
    return true;
  };

  const removeAddress = (addr: string): void => {
    onChange(value.filter(v => v !== addr));
  };

  return (
    <div>
      <label className="form-label" style={{ fontSize: '0.8rem' }}>{props.label}</label>
      {props.hint && (
        <p style={{ margin: '0 0 8px', fontSize: '0.75rem', color: 'var(--dex-gray-500)', lineHeight: 1.45 }}>{props.hint}</p>
      )}

      {/* Aktuelle Empfänger */}
      {value.length === 0 ? (
        <p style={{ margin: '0 0 10px', fontSize: '0.8rem', color: 'var(--dex-orange-dark, #b35a00)' }}>
          Noch keine Empfänger — an diesen Verteiler kann nichts versendet werden.
        </p>
      ) : (
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 10 }}>
          {value.map(addr => {
            const prof = profiles[addr];
            const isPerson = !!prof;
            return (
              <span key={addr} style={chipBase}>
                {isPerson ? (
                  <img
                    src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(addr)}&size=S`}
                    alt=""
                    style={{ width: 28, height: 28, borderRadius: '50%', objectFit: 'cover', flexShrink: 0 }}
                  />
                ) : (
                  // Gruppenadresse/Funktionspostfach: bewusst KEIN Personen-
                  // Kreis, sonst behauptet die Anzeige einen Menschen, den es
                  // nicht gibt.
                  <span style={{
                    width: 28, height: 28, borderRadius: 6, flexShrink: 0,
                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                    background: 'var(--dex-gray-100, #f0f0f0)', color: 'var(--dex-gray-600)',
                    fontSize: '0.7rem', fontWeight: 700,
                  }}>@</span>
                )}
                <span style={{ minWidth: 0, lineHeight: 1.25 }}>
                  <span style={{ display: 'block', fontSize: '0.8rem', fontWeight: 600, color: 'var(--dex-gray-800)' }}>
                    {isPerson ? prof!.displayName : addr}
                  </span>
                  <span style={{ display: 'block', fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>
                    {isPerson
                      ? [prof!.jobTitle, prof!.location].filter(Boolean).join(' · ') || addr
                      : 'Gruppenadresse'}
                  </span>
                </span>
                <button
                  type="button"
                  disabled={props.disabled}
                  onClick={() => removeAddress(addr)}
                  onMouseEnter={() => setRemoveHover(addr)}
                  onMouseLeave={() => setRemoveHover(prev => (prev === addr ? null : prev))}
                  title={`${addr} aus dem Verteiler entfernen`}
                  aria-label={`${addr} entfernen`}
                  style={{
                    border: 'none', background: removeHover === addr ? 'rgba(218,41,28,0.12)' : 'transparent',
                    color: removeHover === addr ? 'var(--dex-red, #da291c)' : 'var(--dex-gray-500)',
                    borderRadius: '50%', width: 22, height: 22, flexShrink: 0,
                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                    cursor: props.disabled ? 'default' : 'pointer',
                    transition: 'background 120ms ease, color 120ms ease',
                  }}
                >
                  <X size={13} />
                </button>
              </span>
            );
          })}
        </div>
      )}

      {/* Person hinzufügen */}
      <div style={{ marginBottom: 8 }}>
        <UserFieldPicker
          value={pickerValue}
          onChange={v => {
            setPickerValue(v);
            const m = (v || '').match(/<([^>]+@[^>]+)>/);
            if (m && addAddress(m[1])) setPickerValue('');
          }}
          searchUsers={searchUsers}
          searchUserByEmail={searchUserByEmail}
          placeholder="Person suchen (Name oder E-Mail)…"
          errorStyle={{}}
          forcedIsDe
        />
      </div>

      {/* Gruppenadresse hinzufügen */}
      <div style={{ display: 'flex', gap: 8, alignItems: 'flex-start', flexWrap: 'wrap' }}>
        <input
          className="form-input"
          style={{ flex: '1 1 220px', minWidth: 0, fontSize: '0.82rem' }}
          value={groupInput}
          disabled={props.disabled}
          placeholder="Gruppenadresse / Funktionspostfach…"
          onChange={e => { setGroupInput(e.target.value); if (groupError) setGroupError(''); }}
          onKeyDown={e => {
            if (e.key === 'Enter') { e.preventDefault(); if (addAddress(groupInput)) setGroupInput(''); }
          }}
        />
        <button
          type="button"
          className="btn btn-secondary"
          style={{ fontSize: '0.8rem', padding: '6px 14px', whiteSpace: 'nowrap' }}
          disabled={props.disabled || !groupInput.trim()}
          onClick={() => { if (addAddress(groupInput)) setGroupInput(''); }}
        >
          <Plus size={13} /> Hinzufügen
        </button>
      </div>
      {groupError && (
        <p style={{ margin: '6px 0 0', fontSize: '0.75rem', color: 'var(--dex-red, #da291c)' }}>{groupError}</p>
      )}
    </div>
  );
}
