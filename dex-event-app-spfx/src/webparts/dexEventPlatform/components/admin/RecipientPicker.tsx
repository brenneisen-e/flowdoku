/**
 * v30.45: Empfänger-Verwaltung mit Profil-Chips (zuerst für das F&A Center,
 * seit v30.51 auch für die CC-Felder der Organizer-Mails).
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
 *
 * v31.3: Optik auf den Klassensatz aus `dexUi.ts` umgestellt (Organizer-Center-
 * Runde). Aus den frei schwebenden Chips ist eine Liste aus `dex-ui-row`-Zeilen
 * geworden — in der schmalen Spalte des F&A Centers brach ein Chip mit Foto,
 * Name und Job Title ohnehin um, und als Zeile steht der Name jeder Person
 * untereinander an derselben Stelle. Der Hover liegt jetzt dort, wo ein Klick
 * etwas tut: auf dem Entfernen-Knopf (`dex-ui-iconbtn--danger`), nicht mehr in
 * einem `onMouseEnter`-State.
 */
import * as React from 'react';
import { UserFieldPicker } from '../UserFieldPicker';
import { X, Plus, AlertCircle } from '../Icons';
import { ensureDexUiStyles } from '../dexUi';
import { useLocaleSafe } from '../../context/LanguageContext';

type Profile = { displayName: string; location: string; jobTitle: string };

export interface RecipientPickerProps {
  /** Text, wenn noch niemand eingetragen ist. Default: F&A-Verteiler-Wortlaut. */
  emptyText?: string;
  label: string;
  hint?: string;
  value: string[];
  onChange: (next: string[]) => void;
  searchUsers: (q: string, includeIntl?: boolean) => Promise<Array<{ email: string; displayName: string; location: string; jobTitle: string }>>;
  searchUserByEmail: (email: string) => Promise<Profile | null>;
  disabled?: boolean;
}

export default function RecipientPicker(props: RecipientPickerProps): React.ReactElement {
  const { value, onChange, searchUsers, searchUserByEmail } = props;
  // v31.3: Die vier eigenen Texte der Komponente waren als einzige im
  // Verteiler nur deutsch (label/hint/emptyText kommen zweisprachig von den
  // Aufrufern herein). Die Sprache kommt aus dem Context; ohne Provider
  // faellt sie auf Deutsch zurueck — die Props bleiben unveraendert.
  const isDe = useLocaleSafe() === 'de';
  // v31.3: Der Verteiler steht auch außerhalb eines Modals (F&A Center) —
  // dort injiziert das gemeinsame Stylesheet sonst niemand.
  ensureDexUiStyles();
  // Aufgelöste Profile je Adresse. `null` = geprüft und KEINE Person
  // (Gruppenadresse); fehlender Schlüssel = noch nicht geprüft.
  const [profiles, setProfiles] = React.useState<Record<string, Profile | null>>({});
  const [pickerValue, setPickerValue] = React.useState('');
  const [groupInput, setGroupInput] = React.useState('');
  // v30.51.1: Meldung für BEIDE Wege (Personensuche und Gruppenadresse) —
  // die Personensuche kann dieselbe Ablehnung erzeugen wie das Textfeld.
  const [addError, setAddError] = React.useState('');

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
      setAddError(isDe ? 'Bitte eine vollständige E-Mail-Adresse eingeben.' : 'Please enter a complete email address.');
      return false;
    }
    if (value.some(v => v.toLowerCase() === addr)) {
      setAddError(isDe ? 'Diese Adresse steht bereits im Verteiler.' : 'This address is already in the list.');
      return false;
    }
    setAddError('');
    onChange(value.concat(addr));
    return true;
  };

  const removeAddress = (addr: string): void => {
    onChange(value.filter(v => v !== addr));
  };

  return (
    <div>
      {/* 1. Wofür ist der Verteiler — und wie viele stehen drin? Die Zahl stand
          bisher nirgends; wer prüft, ob eine Person schon eingetragen ist,
          musste die Chips zählen. */}
      <div className="dex-ui-label" style={{ marginBottom: props.hint ? 4 : 8 }}>
        <span>{props.label}</span>
        {value.length > 0 && <span className="dex-ui-pill dex-ui-pill--gray">{value.length}</span>}
      </div>
      {props.hint && (
        <p className="dex-ui-help" style={{ margin: '0 0 10px' }}>{props.hint}</p>
      )}

      {/* 2. Aktuelle Empfänger. Leer ist kein Nebensatz, sondern eine Folge:
          an einen leeren Verteiler geht nichts raus — deshalb ein Warnkasten
          statt einer orangen Textzeile. */}
      {value.length === 0 ? (
        <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm" style={{ marginBottom: 12 }}>
          <span className="dex-ui-callout-icon"><AlertCircle size={14} /></span>
          <span>{props.emptyText || (isDe ? 'Noch keine Empfänger — an diesen Verteiler kann nichts versendet werden.' : 'No recipients yet — nothing can be sent to this list.')}</span>
        </div>
      ) : (
        <div className="dex-ui-card dex-ui-card--list" style={{ marginBottom: 12 }}>
          {value.map(addr => {
            const prof = profiles[addr];
            const isPerson = !!prof;
            return (
              // Die Zeile selbst ist keine Aktion — ihr Hover hebt nur den
              // Entfernen-Knopf hervor (dex-ui-row-actions), der einzige
              // Klick, den es hier gibt.
              <div key={addr} className="dex-ui-row">
                {isPerson ? (
                  <img
                    src={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(addr)}&size=S`}
                    alt=""
                    className="dex-ui-avatar"
                  />
                ) : (
                  // Gruppenadresse/Funktionspostfach: bewusst KEIN Personen-
                  // Kreis, sonst behauptet die Anzeige einen Menschen, den es
                  // nicht gibt.
                  <span className="dex-ui-avatar" aria-hidden="true" style={{ borderRadius: 8 }}>@</span>
                )}
                {/* v31.3: Die volle Adresse steht im Titel der Zeile — bei einer
                    Person verdrängen Job Title und Standort sie sonst ganz. */}
                <div className="dex-ui-row-main" title={addr}>
                  <div className="dex-ui-row-title">{isPerson ? prof!.displayName : addr}</div>
                  <div className="dex-ui-row-sub" style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {isPerson
                      ? [prof!.jobTitle, prof!.location].filter(Boolean).join(' · ') || addr
                      : (isDe ? 'Gruppenadresse' : 'Group address')}
                  </div>
                </div>
                <span className="dex-ui-row-actions">
                  <button
                    type="button"
                    className="dex-ui-iconbtn dex-ui-iconbtn--danger"
                    disabled={props.disabled}
                    onClick={() => removeAddress(addr)}
                    title={`${addr} aus dem Verteiler entfernen`}
                    aria-label={`${addr} entfernen`}
                  >
                    <X size={14} />
                  </button>
                </span>
              </div>
            );
          })}
        </div>
      )}

      {/* 3. Hinzufügen — beide Wege unter einer Überschrift, damit die Trennung
          zwischen „steht drin“ und „kommt dazu“ sichtbar ist. */}
      <div className="dex-ui-section-title">{isDe ? 'Empfänger hinzufügen' : 'Add recipients'}</div>
      {/* Die Meldung steht ÜBER beiden Feldern: sie kann aus der Personensuche
          genauso kommen wie aus dem Adressfeld (v30.51.1), und unter dem
          Adressfeld hätte sie der Suchende nicht gesehen. */}
      {addError && (
        <div className="dex-ui-callout dex-ui-callout--danger dex-ui-callout--sm" role="alert" style={{ marginBottom: 8 }}>
          <span className="dex-ui-callout-icon"><AlertCircle size={14} /></span>
          <span>{addError}</span>
        </div>
      )}

      {/* Person hinzufügen */}
      <div>
        <UserFieldPicker
          value={pickerValue}
          onChange={v => {
            const m = (v || '').match(/<([^>]+@[^>]+)>/);
            if (!m) { setPickerValue(v); return; }
            // v30.51.1: Der Picker wird IMMER geleert, auch wenn die Adresse
            // abgelehnt wurde.
            //
            // `UserFieldPicker` hält genau EINE Person: Sobald eine gewählt
            // ist, zeigt er ihre Karte statt des Suchfelds. Vorher wurde er
            // nur bei Erfolg geleert — wer dieselbe Person versehentlich ein
            // zweites Mal wählte, bekam „steht bereits im Verteiler" und
            // gleichzeitig einen Picker, der die Person festhielt: Ab da ließ
            // sich hier NIEMAND mehr eintragen, und der einzige Ausweg war,
            // das × an der Karte zu treffen. Genau das war der gemeldete Fall.
            addAddress(m[1]);
            setPickerValue('');
          }}
          searchUsers={searchUsers}
          searchUserByEmail={searchUserByEmail}
          placeholder="Person suchen (Name oder E-Mail)…"
          errorStyle={{}}
          forcedIsDe
        />
      </div>

      {/* Gruppenadresse hinzufügen — Feld und Knopf in EINER Zeile: der Knopf
          gehört zur Eingabe, nicht an den rechten Rand. */}
      <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap', marginTop: 10 }}>
        <input
          className="dex-ui-input"
          style={{ flex: '1 1 200px', minWidth: 0 }}
          value={groupInput}
          disabled={props.disabled}
          aria-label="Gruppenadresse oder Funktionspostfach"
          placeholder="Gruppenadresse, z.B. fa-abrechnung@deloitte.de"
          onChange={e => { setGroupInput(e.target.value); if (addError) setAddError(''); }}
          onKeyDown={e => {
            if (e.key === 'Enter') { e.preventDefault(); if (addAddress(groupInput)) setGroupInput(''); }
          }}
        />
        <button
          type="button"
          className="btn btn-secondary dex-ui-btn-sm"
          style={{ whiteSpace: 'nowrap' }}
          disabled={props.disabled || !groupInput.trim()}
          onClick={() => { if (addAddress(groupInput)) setGroupInput(''); }}
        >
          <Plus size={13} /> Hinzufügen
        </button>
      </div>
    </div>
  );
}
