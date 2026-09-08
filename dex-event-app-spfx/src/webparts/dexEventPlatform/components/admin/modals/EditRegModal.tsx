/* EditRegModal — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 14077-14370 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.2: Auf das gemeinsame `Modal` (Kopf, Fuß, Escape, Backdrop) und die
 * dex-ui-Klassen umgestellt. Die Reihenfolge folgt jetzt dem, was der
 * Organizer hier tut: Person korrigieren → Antworten zum Event anpassen →
 * Profil nachsehen. Das M365-Profil ist nicht änderbar und steht deshalb im
 * Aufklapper statt als sechs graue Kästen zwischen den Eingaben. Der alte
 * Erklär-Absatz (sieben Zeilen) steht als je ein Satz dort, wo er gilt:
 * Tenant-Prüfung unter der E-Mail, Profil-Herkunft im Aufklapper,
 * Protokollierung im Untertitel. Gespeichert wird exakt dasselbe wie vorher.
 *
 * v31.4 (Nachtrag): Dazu kommt „Ausgegebenes Trikot". Nutzer-Wunsch vom
 * 08.09.2026 (B2Run Köln, Lauftag am nächsten Tag): „Zudem soll man im
 * Nachhinein unter der TN-Liste auch über Bearbeiten das rausgegebene Shirt
 * ändern und speichern können." Der Wert steht NICHT im Zeilen-Patch, sondern
 * in der eigenen Spalte `ShirtIssued` — geschrieben wird er deshalb im
 * Save-Pfad über `setShirtIssued`/`clearShirtIssued` (s.
 * `logic/useEditModalHandlers`), der Dialog hält ihn nur unter
 * `SHIRT_ISSUED_FORM_KEY` im Formular.
 */
import * as React from 'react';
import Modal from '../../Modal';
import { AlertCircle, ChevronDown, Pencil, Shirt } from '../../Icons';
import { cx } from '../../dexUi';
import { MultiSelectDropdown } from '../../MultiSelectDropdown';
import { DeloitteEvent } from '../../../types';
import { SPRegistration } from '../../../services/EventService';
import {
  SHIRT_ISSUED_FORM_KEY, parseShirtIssue, parseShirtStock, shirtAllocate, shirtFieldOf,
} from '../../../utils/checkInExtras';
import { FieldSelectInput, fieldVisibleByShowIf } from './FieldSelectInput';

export interface EditRegModalProps {
  closeEditModal: () => void;
  editError: string;
  editForm: Record<string, string>;
  /** v31.4: Die bearbeitete Zeile — für den gespeicherten Ausgabe-Eintrag. */
  editingReg: SPRegistration;
  isDe: boolean;
  isSavingEdit: boolean;
  /**
   * v31.4: Alle Zeilen des Events. Daraus kommen das Größenfeld
   * (`shirtFieldOf` entscheidet über die Antworten, nicht über die
   * Feld-Reihenfolge) und die Größen, die es an diesem Event überhaupt gibt.
   */
  registrations: SPRegistration[];
  saveEdit: () => Promise<void>;
  selectedEvent: DeloitteEvent;
  /**
   * v31.4 (Review): Das Eltern-Event, wenn ein Termin gewählt ist. Die
   * Größenfrage kann auf der Klammer stehen und wird von den Terminen nicht
   * geerbt — dann fand `shirtFieldOf` hier nichts und der ganze Abschnitt
   * „Trikot-Ausgabe" fiel ersatzlos weg, obwohl das Check-in-Team am Termin
   * eine Größe festgehalten hat. Der Tisch (`CheckInPage.shirtFieldsFor`) und
   * die Bestellliste (`ShirtSizeModal`) lösen längst über BEIDE Ebenen auf.
   */
  parentEvent?: DeloitteEvent | null;
  setEditForm: React.Dispatch<React.SetStateAction<Record<string, string>>>;
}

export const EditRegModal: React.FC<EditRegModalProps> = (p) => {
  const { closeEditModal, editError, editForm, editingReg, isDe, isSavingEdit, registrations, saveEdit, selectedEvent, parentEvent, setEditForm } = p;
  // v31.2: Der Aufrufer mountet den Dialog je Öffnen neu — der Aufklapper
  // startet deshalb verlässlich geschlossen.
  const [profileOpen, setProfileOpen] = React.useState(false);

  // Vorname / Nachname / E-Mail sind seit v9.7 editierbar (mit
  // Deloitte-Domain- und Tenant-Existenz-Check beim Speichern).
  // Die uebrigen Profil-Felder bleiben read-only — sie kommen
  // aus dem M365-Profil und werden bei einer Mail-Änderung
  // mit den Profil-Daten der neuen Person überschrieben.
  const readOnlyFields: Array<{ key: string; label: string }> = [
    { key: 'Anrede', label: isDe ? 'Anrede' : 'Salutation' },
    { key: 'Phone', label: isDe ? 'Telefon' : 'Phone' },
    { key: 'Department', label: 'Department' },
    { key: 'Location', label: isDe ? 'Standort' : 'Location' },
    { key: 'JobTitle', label: 'Job Title' },
    { key: 'Status', label: 'Status' },
  ];

  // B2Run-Starter-Typ (Funstarter / Durchstarter). Hardcoded SP-Spalte auf
  // der Teilnehmerliste (kein regulärer Custom-Field-Eintrag), daher explizit
  // hier gerendert. Updates BEIDE intern getrackten Felder zugleich
  // (StarterType + PreferredStarterType) — die getrennte Speicherung von
  // „aktuell vs. Wunsch" ist Implementierungs-Detail für die Warteliste-
  // Nachrück-Logik und braucht im Edit-Modal keine UI-Komplexität. v10.15+
  const hasStarterType = selectedEvent.durchstarterCapacity !== undefined
    && selectedEvent.funstarterCapacity !== undefined
    && (selectedEvent.durchstarterCapacity > 0 || selectedEvent.funstarterCapacity > 0);
  // v31.2: Die Gruppen heißen je Event anders (splitLabelA/B) — der Wert in
  // der Spalte bleibt „Durchstarter"/„Funstarter", nur die Beschriftung folgt
  // dem Event, wie in Tabelle und KPI-Kacheln.
  const starterOptions = [
    { value: 'Durchstarter', label: (selectedEvent.splitLabelA && selectedEvent.splitLabelA.trim()) || 'Durchstarter' },
    { value: 'Funstarter', label: (selectedEvent.splitLabelB && selectedEvent.splitLabelB.trim()) || 'Funstarter' },
  ];
  const customFields = selectedEvent.eventSpecificFields || [];
  const hasAnswers = hasStarterType || customFields.length > 0;

  /**
   * v31.4 (Nachtrag): Das ausgegebene Trikot — nur an Events mit Größenfeld.
   *
   * Die Größen kommen aus derselben Rechnung wie am Check-in-Tisch
   * (`shirtAllocate`), damit hier keine Größe auswählbar ist, die es am Event
   * gar nicht gibt — und keine fehlt, die im Karton liegt. `null` heißt: Das
   * Event hat kein Größenfeld, dann gibt es auch nichts auszugeben.
   *
   * `assumedWish` ist die Annahme aus dem Check-in (s. `shirtAllocate`): Sie
   * wird bewusst NICHT als Wert vorgetragen — sonst schriebe ein Klick auf
   * „Speichern" eine Vermutung in die Spalte und machte sie zur Tatsache.
   *
   * v31.4 (Review): Gesucht wird über Termin UND Klammer (das Größenfeld darf
   * auf beiden Ebenen stehen), und der Abschnitt bleibt sichtbar, sobald die
   * Zeile einen `ShirtIssued`-Wert trägt — auch ohne gefundenes Feld. Sonst
   * gäbe es für eine am Tisch festgehaltene, falsche Größe keinen einzigen
   * Weg zurück.
   */
  const shirt = React.useMemo(() => {
    const flds = (selectedEvent.eventSpecificFields || [])
      .concat((parentEvent && parentEvent.eventSpecificFields) || []);
    const field = shirtFieldOf(flds, registrations);
    const savedIssue = parseShirtIssue(editingReg && editingReg.ShirtIssued);
    if (!field) {
      // Kein Größenfeld, aber ein festgehaltener Wert: nur der Wert selbst als
      // Auswahl — mehr weiß die App an dieser Stelle ehrlicherweise nicht.
      if (!savedIssue || !savedIssue.size) return null;
      return { sizes: [savedIssue.size], assumedWish: '', noField: true };
    }
    // Der Bestand hängt am Hauptevent; bei einem Termin ohne eigenen Bestand
    // bleibt die Liste trotzdem vollständig — `shirtAllocate` nimmt auch die
    // gewünschten und die bereits ausgegebenen Größen auf. v31.4 (Review):
    // Deshalb zuerst die Klammer fragen, dann das gewählte Event.
    const alloc = shirtAllocate(flds, registrations, parseShirtStock(
      (parentEvent && parentEvent.emailTemplateOverrides) || selectedEvent.emailTemplateOverrides
    ));
    const em = ((editingReg && editingReg.ParticipantEmail) || '').toLowerCase().trim();
    const a = em ? alloc.byEmail[em] : undefined;
    return {
      // Doppelte Beschriftungen wären doppelte React-Keys — zwei Schlüssel
      // können dieselbe Anzeige-Schreibweise tragen.
      sizes: alloc.rows.map(r => r.size).filter((s, i, arr) => !!s && arr.indexOf(s) === i),
      assumedWish: (a && a.issued && a.issuedAssumed) ? a.issued : '',
      noField: false,
    };
  }, [selectedEvent, parentEvent, registrations, editingReg]);
  // „andere Größe" bleibt offen, sobald der Nutzer sie gewählt hat; von selbst
  // geht sie auf, wenn der gespeicherte Wert in keiner Kachel steckt.
  const [shirtOther, setShirtOther] = React.useState(false);
  const shirtIssued = editForm[SHIRT_ISSUED_FORM_KEY] || '';
  const setShirtIssued = (v: string): void => setEditForm(prev => ({ ...prev, [SHIRT_ISSUED_FORM_KEY]: v }));
  const shirtOtherOpen = shirtOther || (!!shirtIssued && !!shirt && shirt.sizes.indexOf(shirtIssued) < 0);
  const shirtSaved = parseShirtIssue(editingReg && editingReg.ShirtIssued);

  const spOf = (fieldId: string): string => {
    const src = (selectedEvent.eventSpecificFields || []).find(f => f.id === fieldId);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return src ? ((src as any).spInternalName || '') : '';
  };

  const renderEditable = (key: string, label: string, type?: string, help?: React.ReactNode): React.ReactNode => (
    <div className="dex-ui-field">
      <label className="dex-ui-label" htmlFor={`editreg-${key}`}>{label}</label>
      <input
        id={`editreg-${key}`}
        className="dex-ui-input"
        type={type || 'text'}
        value={editForm[key] || ''}
        onChange={e => setEditForm(prev => ({ ...prev, [key]: e.target.value }))}
      />
      {help && <p className="dex-ui-help">{help}</p>}
    </div>
  );
  const renderReadOnly = (label: string, value: string): React.ReactNode => (
    <div className="dex-ui-field">
      <span className="dex-ui-label" style={{ color: 'var(--dex-gray-500)' }}>{label}</span>
      <div style={{
        padding: '8px 12px', background: 'var(--dex-gray-50, #fafafa)',
        border: '1px solid var(--dex-gray-200, #e8e8e8)', borderRadius: 10,
        fontSize: '0.88rem', lineHeight: 1.5, minHeight: 38, boxSizing: 'border-box',
        color: value ? 'var(--dex-gray-800)' : 'var(--dex-gray-400)',
      }}>
        {value || (isDe ? '— nicht gesetzt —' : '— not set —')}
      </div>
    </div>
  );

  return (
    <Modal
      open
      onClose={closeEditModal}
      dismissable={!isSavingEdit}
      maxWidth={920}
      ariaLabel={isDe ? 'Teilnehmer bearbeiten' : 'Edit attendee'}
      icon={<Pencil size={20} />}
      title={<>
        {isDe ? 'Teilnehmer bearbeiten' : 'Edit attendee'}
        {' — '}
        <span style={{ color: 'var(--dex-green-dark)' }}>{editForm.Vorname} {editForm.Nachname}</span>
        {editForm.Status && (
          <span className="dex-ui-pill dex-ui-pill--gray" style={{ marginLeft: 10, verticalAlign: 'middle' }}>{editForm.Status}</span>
        )}
      </>}
      subtitle={isDe
        ? 'Name, E-Mail und die Antworten dieser Person korrigieren. Jede Änderung landet mit Datum und deinem Namen im Audit-Log und im ChangeLog der Person.'
        : 'Fix this person’s name, email and answers. Every change is logged in the audit log and in the attendee’s ChangeLog with date and your name.'}
      footer={<>
        <button type="button" className="btn btn-secondary" onClick={closeEditModal} disabled={isSavingEdit}>
          {isDe ? 'Abbrechen' : 'Cancel'}
        </button>
        <button type="button" className="btn btn-primary" onClick={saveEdit} disabled={isSavingEdit}>
          {isSavingEdit ? (isDe ? 'Speichert…' : 'Saving…') : (isDe ? 'Speichern' : 'Save')}
        </button>
      </>}
    >
      {/* 1) Person — die Felder, die man hier am häufigsten korrigiert */}
      <div className="dex-ui-section">
        <div className="dex-ui-section-title">{isDe ? 'Person' : 'Person'}</div>
        <p className="dex-ui-section-desc">
          {isDe ? 'Zum Beispiel nach einem Tippfehler bei der manuellen Anlage.' : 'For example after a typo during manual creation.'}
        </p>
        <div className="dex-ui-grid-2">
          {renderEditable('Vorname', isDe ? 'Vorname' : 'First name')}
          {renderEditable('Nachname', isDe ? 'Nachname' : 'Last name')}
        </div>
        <div style={{ marginTop: 16 }}>
          {renderEditable('ParticipantEmail', 'E-Mail', 'email', isDe
            ? 'Nur Deloitte-Adressen (@deloitte.de / @deloitte.com). Beim Speichern prüft die App, ob die Person im Tenant existiert — externe Adressen sind nicht erlaubt. Bei neuer Adresse zieht die App Telefon, Department, Standort und Job Title der neuen Person automatisch nach.'
            : 'Only Deloitte addresses (@deloitte.de / @deloitte.com). On save the app verifies that the person exists in the tenant — external addresses are not allowed. With a new address, phone, department, location and job title of the new person are refreshed automatically.')}
        </div>
      </div>

      {/* 2) Antworten zum Event — DAS ist der editierbare Teil. Renderer
          abhängig vom Field-Type (text/number/select/checkbox). Multi-Select
          speichert Werte als " | "-getrennten String, identisch zum
          Registrierungs-Pfad. */}
      {hasAnswers && (
        <div className="dex-ui-section">
          <div className="dex-ui-section-title">{isDe ? 'Antworten zu diesem Event' : 'Answers for this event'}</div>
          <p className="dex-ui-section-desc">
            {isDe
              ? 'Diese Antworten werden gespeichert — das M365-Profil unten nicht.'
              : 'These answers are saved — the M365 profile below is not.'}
          </p>
          <div className="dex-ui-grid-2">
            {hasStarterType && (
              <div className="dex-ui-field" style={{ gridColumn: '1 / -1' }}>
                <span className="dex-ui-label" id="editreg-startertype-label">
                  {isDe ? 'Zu welcher Gruppe gehört die Person?' : 'Which group does this person belong to?'}
                </span>
                <div className="dex-ui-inline" role="group" aria-labelledby="editreg-startertype-label">
                  {starterOptions.map(o => {
                    const active = (editForm.StarterType || '') === o.value;
                    return (
                      <button
                        key={o.value}
                        type="button"
                        className={cx('dex-ui-chip', active && 'is-active')}
                        aria-pressed={active}
                        onClick={() => {
                          // Beide Felder synchron halten — der Aktuelle wechselt
                          // mit, der Wunsch ebenfalls (User-Erwartung: „ich ändere
                          // den Starter-Typ" = beides ändert sich). v31.2: Nochmal
                          // anklicken leert die Auswahl — das war vorher die
                          // Option „— bitte wählen —" im Dropdown.
                          const v = active ? '' : o.value;
                          setEditForm(prev => ({ ...prev, StarterType: v, PreferredStarterType: v }));
                        }}
                      >
                        {o.label}
                      </button>
                    );
                  })}
                </div>
                <p className="dex-ui-help">
                  {isDe
                    ? 'Starter-Typ für die Zuteilung: aktueller Typ und Wunsch wechseln zusammen. Noch einmal anklicken setzt die Auswahl zurück.'
                    : 'Starter type for the allocation: current type and preference change together. Click again to clear the selection.'}
                </p>
              </div>
            )}

            {customFields.map(cf => {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const sp = (cf as any).spInternalName || '';
              if (!sp) return null;
              // v30.95: Sichtbarkeitsregel wie auf der Anmeldeseite —
              // „Mobilnummer" hing als Pflichtfeld mit Stern im Dialog,
              // obwohl der Infoservice auf Nein stand.
              if (!fieldVisibleByShowIf(cf, fid => editForm[spOf(fid)] || '')) return null;
              const value = editForm[sp] || '';
              const setVal = (v: string): void => setEditForm(prev => ({ ...prev, [sp]: v }));
              const inputId = `editreg-cf-${cf.id}`;
              const labelEl = (
                <label className="dex-ui-label" htmlFor={inputId}>
                  {cf.label}{cf.required && <span style={{ color: 'var(--dex-red, #c00)' }} aria-hidden="true">*</span>}
                </label>
              );

              // Single-Select-Dropdown. v30.95: dieselbe Kombibox wie
              // die Anmeldeseite (Kategorien, Wert „Kategorie Option").
              if (cf.type === 'select' && !cf.multi && cf.options && cf.options.length > 0) {
                return (
                  <div key={cf.id} className="dex-ui-field">
                    {labelEl}
                    <FieldSelectInput field={cf} value={value} onChange={setVal} isDe={isDe} />
                  </div>
                );
              }

              // v11.89: Multi-Select-Dropdown (vorher Checkbox-Liste).
              // Werte werden weiterhin ' | '-getrennt gespeichert.
              if (cf.type === 'select' && cf.multi && cf.options && cf.options.length > 0) {
                const selected = value.split(' | ').map(s => s.trim()).filter(Boolean);
                return (
                  <div key={cf.id} className="dex-ui-field" style={{ gridColumn: '1 / -1' }}>
                    {labelEl}
                    <MultiSelectDropdown
                      options={cf.options}
                      value={selected}
                      onChange={next => setVal(next.join(' | '))}
                      placeholder={isDe ? '— bitte wählen —' : '— please choose —'}
                    />
                  </div>
                );
              }

              // Checkbox (true/false). v31.2: als Schalter mit Ja/Nein daneben —
              // eine nackte Checkbox sagt nicht, was der Haken bedeutet.
              if (cf.type === 'checkbox') {
                const isChecked = value === 'true' || value === '1';
                return (
                  <div key={cf.id} className="dex-ui-field">
                    {labelEl}
                    <label className="dex-ui-switch" style={{ marginTop: 6 }}>
                      <input
                        id={inputId}
                        type="checkbox"
                        checked={isChecked}
                        onChange={e => setVal(e.target.checked ? 'true' : 'false')}
                      />
                      <span className="dex-ui-switch-track" />
                      <span className="dex-ui-switch-label">{isChecked ? (isDe ? 'Ja' : 'Yes') : (isDe ? 'Nein' : 'No')}</span>
                    </label>
                  </div>
                );
              }

              // Number
              if (cf.type === 'number') {
                return (
                  <div key={cf.id} className="dex-ui-field">
                    {labelEl}
                    <input id={inputId} className="dex-ui-input" type="number" value={value} onChange={e => setVal(e.target.value)} />
                  </div>
                );
              }

              // Default: text-Input (auch für 'text', 'user', 'roommate')
              return (
                <div key={cf.id} className="dex-ui-field">
                  {labelEl}
                  <input id={inputId} className="dex-ui-input" value={value} onChange={e => setVal(e.target.value)} />
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* 3) Trikot-Ausgabe — was die Person am Tisch WIRKLICH bekommen hat.
          v31.4 (Nachtrag): Keine Antwort aus dem Formular, sondern ein
          Protokoll des Organisators — deshalb ein eigener Abschnitt und ein
          eigener Schreibvorgang (Spalte `ShirtIssued`). */}
      {shirt && (
        <div className="dex-ui-section">
          <div className="dex-ui-section-title">{isDe ? 'Trikot-Ausgabe' : 'Shirt handout'}</div>
          <p className="dex-ui-section-desc">
            {isDe
              ? 'Korrigiere hier, was am Ausgabetisch wirklich rausgegangen ist. Der Wert zieht den Bestand ab und schlägt jeden Vorschlag der App.'
              : 'Correct here what really went out at the handout desk. The value is deducted from the stock and beats every proposal the app makes.'}
          </p>
          <div className="dex-ui-field">
            <span className="dex-ui-label" id="editreg-shirt-label">
              {isDe ? 'Welches Trikot hat die Person bekommen?' : 'Which shirt did this person get?'}
            </span>
            <div className="dex-ui-inline" role="group" aria-labelledby="editreg-shirt-label">
              <button
                type="button"
                className={cx('dex-ui-chip', !shirtOtherOpen && !shirtIssued && 'is-active')}
                aria-pressed={!shirtOtherOpen && !shirtIssued}
                onClick={() => { setShirtOther(false); setShirtIssued(''); }}
              >
                {isDe ? 'nichts ausgegeben' : 'nothing handed out'}
              </button>
              {shirt.sizes.map(s => {
                const active = !shirtOtherOpen && shirtIssued === s;
                return (
                  <button
                    key={s}
                    type="button"
                    className={cx('dex-ui-chip', active && 'is-active')}
                    aria-pressed={active}
                    onClick={() => { setShirtOther(false); setShirtIssued(active ? '' : s); }}
                  >
                    <Shirt size={13} /> {s}
                  </button>
                );
              })}
              <button
                type="button"
                className={cx('dex-ui-chip', shirtOtherOpen && 'is-active')}
                aria-pressed={shirtOtherOpen}
                onClick={() => setShirtOther(true)}
              >
                {isDe ? 'andere Größe' : 'other size'}
              </button>
            </div>
            {shirtOtherOpen && (
              <input
                id="editreg-shirt-other"
                className="dex-ui-input dex-ui-input--sm"
                style={{ marginTop: 8, maxWidth: 260 }}
                value={shirtIssued}
                onChange={e => setShirtIssued(e.target.value)}
                placeholder={isDe ? 'z.B. Herrengröße XXL' : 'e.g. Men XXL'}
                aria-label={isDe ? 'Andere ausgegebene Größe' : 'Other size handed out'}
              />
            )}
            <p className="dex-ui-help">
              {isDe
                ? 'Gespeichert wird mit „Speichern" unten — zusammen mit den übrigen Änderungen. Fehlt auf dieser Teilnehmerliste die Spalte, sagt die Meldung es dir.'
                : 'Saved with “Save” below, together with the other changes. If the column is missing on this attendee list, the message will say so.'}
            </p>
            {/* v31.4 (Review): „Nichts ausgegeben" lässt sich speichern, aber
                nicht durchhalten: `shirtAllocate` zählt jede eingecheckte
                Person OHNE Spalteneintrag wieder mit ihrer Wunschgröße als
                abgeholt. Eine gelöschte Ausgabe fällt damit sofort in die
                Annahme zurück — und der Hinweiskasten darunter erscheint
                gerade dann NICHT, weil `assumedWish` bei einem erfassten Wert
                leer ist. Der Satz gehört also genau hierhin. */}
            {shirtSaved && !shirtIssued && (
              <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm" style={{ marginTop: 8 }}>
                <span className="dex-ui-callout-icon"><AlertCircle size={14} /></span>
                <span>
                  {isDe
                    ? <>Der bisherige Eintrag <strong>{shirtSaved.size}</strong> wird beim Speichern gelöscht. Ist die Person eingecheckt,
                      zählt sie danach wieder mit ihrer Wunschgröße als abgeholt (Annahme) — in der Bestellliste ändert sich die Zahl
                      dann nicht. Wer wirklich keins bekommen hat, muss dafür ausgecheckt oder als No-Show markiert werden.</>
                    : <>The existing entry <strong>{shirtSaved.size}</strong> is deleted when you save. If the person is checked in, they
                      count as collected with their wished size again (assumption) — the number in the order list will not change.
                      For someone who really got none, check them out or mark them as a no-show.</>}
                </span>
              </div>
            )}
            {/* Die Annahme aus der Bestellliste benennen, statt sie als Wert
                vorzutragen: Ein vorbelegtes Feld, das nur eine Vermutung ist,
                wird beim nächsten Speichern zur Tatsache. */}
            {!shirtSaved && !shirtIssued && !!shirt.assumedWish && (
              <div className="dex-ui-callout dex-ui-callout--info dex-ui-callout--sm" style={{ marginTop: 8 }}>
                <span className="dex-ui-callout-icon"><Shirt size={14} /></span>
                <span>
                  {isDe
                    ? <>Zählt aktuell als abgeholt, weil die Person eingecheckt ist — Wunschgröße <strong>{shirt.assumedWish}</strong>.
                      Sobald du hier eine Größe auswählst, gilt deine Angabe. &bdquo;Nichts ausgegeben&ldquo; lässt die Annahme bestehen.</>
                    : <>Currently counts as collected because the person is checked in — wished size <strong>{shirt.assumedWish}</strong>.
                      As soon as you pick a size here, your entry counts. &ldquo;Nothing handed out&rdquo; leaves the assumption in place.</>}
                </span>
              </div>
            )}
          </div>
        </div>
      )}

      {/* 4) Profil aus M365 — nur lesen, deshalb zu */}
      <div className="dex-ui-section">
        <button
          type="button"
          className={cx('dex-ui-disclosure', profileOpen && 'is-open')}
          onClick={() => setProfileOpen(o => !o)}
          aria-expanded={profileOpen}
        >
          <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
          {isDe ? 'Profil aus M365 — nicht änderbar' : 'M365 profile — read-only'}
          <span className="dex-ui-disclosure-count">{readOnlyFields.length}</span>
        </button>
        {profileOpen && (
          <div className="dex-ui-disclosure-body dex-ui-fade-in">
            <p className="dex-ui-section-desc" style={{ margin: '0 0 12px' }}>
              {isDe
                ? 'Anrede, Telefon, Department, Standort und Job Title kommen aus dem M365-Profil und werden bei einem Mail-Wechsel automatisch nachgezogen. Den Status änderst du über die Aktions-Knöpfe in der Liste.'
                : 'Salutation, phone, department, location and job title come from the M365 profile and are refreshed automatically when the email changes. The status is changed via the action buttons in the list.'}
            </p>
            <div className="dex-ui-grid-3">
              {readOnlyFields.map(f => (
                <React.Fragment key={f.key}>{renderReadOnly(f.label, editForm[f.key] || '')}</React.Fragment>
              ))}
            </div>
          </div>
        )}
      </div>

      {editError && (
        <div className="dex-ui-callout dex-ui-callout--danger" role="alert">
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span>{editError}</span>
        </div>
      )}
    </Modal>
  );
};
