import * as React from 'react';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { StickyTabStrip } from '../../wizard/StickyTabStrip';
import { InfoTooltip } from '../../InfoTooltip';
import { SubEventDraft } from '../../wizard/wizardTypes';
import { Send, Trash2 } from '../../Icons';
import { CustomFieldInput } from '../../wizard/customFieldInput';

/* renderHeaderSizeControl — aus EventCreationPage.tsx ausgelagert (Zeilen 1318-1369 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface RenderHeaderSizeControlCtx {
  headerImageLayout: { width: number; paddingV: number; paddingH: number; };
  isDe: boolean;
  setHeaderImageLayout: React.Dispatch<React.SetStateAction<{ width: number; paddingV: number; paddingH: number; }>>;
}

export function renderHeaderSizeControlImpl(ctx: RenderHeaderSizeControlCtx, previewSrc: string, note: string): React.ReactElement {
  const { headerImageLayout, isDe, setHeaderImageLayout } = ctx;
    const PREV_W = 260; const sc = PREV_W / 600;
    const isFullWidthPreset = headerImageLayout.width === 600 && headerImageLayout.paddingV === 0 && headerImageLayout.paddingH === 0;
    const isDefaultPreset = headerImageLayout.width === 180 && headerImageLayout.paddingV === 30 && headerImageLayout.paddingH === 30;
    const numInput = (val: number, min: number, max: number, def: number, set: (n: number) => void): React.ReactElement => (
      <input type="number" min={min} max={max} step={min === 80 ? 10 : 2} value={val}
        onChange={e => set(Math.max(min, Math.min(max, parseInt(e.target.value, 10) || def)))}
        style={{ width: 78, height: 28, fontSize: '0.82rem', borderRadius: 4, border: '1px solid var(--dex-gray-300)', padding: '0 8px' }} />
    );
    const lbl: React.CSSProperties = { display: 'flex', flexDirection: 'column', gap: 3, fontSize: '0.7rem', color: 'var(--dex-gray-600)' };
    return (
      <div style={{ marginTop: 12, padding: '10px 12px', background: 'var(--dex-gray-50, #f7f7f5)', border: '1px solid var(--dex-gray-200)', borderRadius: 8 }}>
        <div style={{ fontSize: '0.7rem', fontWeight: 700, color: 'var(--dex-green-dark, #4a7c1f)', letterSpacing: 0.3, marginBottom: 8 }}>
          {isDe ? 'BILDGRÖSSE IM KOPF — gilt für Mail & Outlook-Termin' : 'HEADER IMAGE SIZE — applies to mail & Outlook invite'}
        </div>
        <div style={{ display: 'flex', gap: 18, flexWrap: 'wrap', alignItems: 'flex-start' }}>
          <div style={{ display: 'flex', alignItems: 'flex-end', gap: 12, flexWrap: 'wrap' }}>
            <label style={lbl}>{isDe ? 'Breite (px)' : 'Width (px)'}{numInput(headerImageLayout.width, 80, 600, 180, n => setHeaderImageLayout(p => ({ ...p, width: n })))}</label>
            <label style={lbl}>{isDe ? 'Abstand seitl.' : 'Padding sides'}{numInput(headerImageLayout.paddingH, 0, 80, 0, n => setHeaderImageLayout(p => ({ ...p, paddingH: n })))}</label>
            <label style={lbl}>{isDe ? 'Abstand ob./unt.' : 'Padding top/bot.'}{numInput(headerImageLayout.paddingV, 0, 80, 0, n => setHeaderImageLayout(p => ({ ...p, paddingV: n })))}</label>
            {/* v28.31: Beide Voreinstellungen zeigen jetzt an, WELCHE gerade
                aktiv ist. Vorher war „Volle Breite" immer gruen und „Standard"
                immer grau — auch wenn tatsaechlich 180/30/30 (= Standard) stand. */}
            <button type="button" onClick={() => setHeaderImageLayout({ width: 600, paddingV: 0, paddingH: 0 })}
              title={isDe ? 'Bild füllt den Kopf über die volle Breite' : 'Image fills the header full width'}
              style={{ height: 28, padding: '0 12px', fontSize: '0.72rem', fontWeight: isFullWidthPreset ? 700 : 600, cursor: 'pointer', background: isFullWidthPreset ? 'var(--dex-green, #86bc25)' : 'transparent', color: isFullWidthPreset ? '#fff' : 'var(--dex-gray-600)', border: isFullWidthPreset ? 'none' : '1px solid var(--dex-gray-300)', borderRadius: 6 }}>
              {isFullWidthPreset ? '✓ ' : ''}{isDe ? 'Volle Breite' : 'Full width'}
            </button>
            <button type="button" onClick={() => setHeaderImageLayout({ width: 180, paddingV: 30, paddingH: 30 })}
              style={{ height: 28, padding: '0 12px', fontSize: '0.72rem', fontWeight: isDefaultPreset ? 700 : 600, cursor: 'pointer', background: isDefaultPreset ? 'var(--dex-green, #86bc25)' : 'transparent', color: isDefaultPreset ? '#fff' : 'var(--dex-gray-600)', border: isDefaultPreset ? 'none' : '1px solid var(--dex-gray-300)', borderRadius: 6 }}>
              {isDefaultPreset ? '✓ ' : ''}{isDe ? 'Standard' : 'Default'}
            </button>
          </div>
          {previewSrc && (
            <div style={{ width: PREV_W, flexShrink: 0, border: '1px solid var(--dex-gray-200)', borderRadius: 4, overflow: 'hidden', background: '#fff' }}>
              <div style={{ textAlign: 'center', padding: `${Math.round(headerImageLayout.paddingV * sc)}px ${Math.round(headerImageLayout.paddingH * sc)}px` }}>
                <img src={previewSrc} alt="" style={{ display: 'inline-block', width: '100%', maxWidth: Math.max(20, Math.round(headerImageLayout.width * sc)), height: 'auto' }} />
              </div>
              <div style={{ borderTop: '2px solid var(--dex-green, #86bc25)' }} />
              <div style={{ fontSize: '0.6rem', color: 'var(--dex-gray-400)', textAlign: 'center', padding: '2px 0' }}>{isDe ? 'So groß im Mail-Kopf (verkleinert)' : 'Size in the mail header (scaled)'}</div>
            </div>
          )}
        </div>
        {/* v28.29: sagt, WOHER das gezeigte Bild kommt (eigenes / vom Hauptevent
            geerbt / Standardlogo). Vorher zeigte die Vorschau kommentarlos das
            Event-Foto, obwohl gespeichert etwas anderes wurde. */}
        {note && (
          <div style={{ marginTop: 8, fontSize: '0.72rem', color: 'var(--dex-gray-600)', lineHeight: 1.45 }}>
            {note}
          </div>
        )}
      </div>
    );
}

/* renderOutlookUpdateButton — aus EventCreationPage.tsx ausgelagert (Zeilen 1940-2026 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface RenderOutlookUpdateButtonCtx {
  activeCommTabIdx: number;
  childTermPlural: string;
  childTermSingular: string;
  createMissingOutlookAppointments: () => Promise<void>;
  editEvent: import("../../../types/index").DeloitteEvent;
  isDe: boolean;
  outlookMissingTargets: () => { id: string; title: string; }[];
  outlookUpdateBusy: boolean;
  outlookUpdateDone: string;
  outlookUpdateTargets: () => { id: string; title: string; }[];
  subEvents: SubEventDraft[];
  title: string;
  triggerOutlookUpdateAll: () => Promise<void>;
  triggerOutlookUpdateNow: () => Promise<void>;
}

export function renderOutlookUpdateButtonImpl(ctx: RenderOutlookUpdateButtonCtx): React.ReactNode {
  const { activeCommTabIdx, childTermPlural, childTermSingular, createMissingOutlookAppointments, editEvent, isDe, outlookMissingTargets, outlookUpdateBusy, outlookUpdateDone, outlookUpdateTargets, subEvents, title, triggerOutlookUpdateAll, triggerOutlookUpdateNow } = ctx;
    if (!editEvent) return null; // nur beim Bearbeiten sinnvoll (Neu-Event hat noch keinen Termin)
    // v28.28: Der Kasten war orange umrandet und wurde dadurch als Warnung
    // („da steht noch was aus") gelesen — obwohl er nur ein dauerhaft
    // verfügbares Werkzeug ist und nach dem Klick natürlich stehen bleibt.
    // Jetzt neutral, mit Ziel-Angabe und sichtbarer Erfolgsmeldung.
    const tabTitle = activeCommTabIdx > 0
      ? (subEvents[activeCommTabIdx - 1]?.title || (childTermSingular || 'Sub-Event'))
      : (title || editEvent.title || (isDe ? 'Hauptevent' : 'main event'));
    const allTargets = outlookUpdateTargets();
    const showAll = allTargets.length > 1;
    // v28.67: fehlende Termine benennen (s. outlookMissingTargets).
    const missingTargets = outlookMissingTargets();
    const totalTargets = allTargets.length + missingTargets.length;
    // v28.69: nachanlegbar sind nur Sub-Events — das Hauptevent nicht, seine
    // Item-Id steht in ParentEventId aller Kinder (s. createMissingOutlookAppointments).
    const missingSubIds = missingTargets.filter(m => m.id && m.id !== (editEvent?.id || ''));
    return (
      <div style={{ marginTop: 14, padding: 12, borderRadius: 8, background: 'var(--dex-gray-50, #f8f9fa)', border: '1px solid var(--dex-gray-200)' }}>
        <div style={{ fontSize: '0.8rem', fontWeight: 600, color: 'var(--dex-gray-800)', marginBottom: 6 }}>
          {isDe ? 'Outlook-Termin manuell nachschicken (optional)' : 'Re-send the Outlook appointment manually (optional)'}
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <button
            type="button"
            className="btn btn-secondary"
            disabled={outlookUpdateBusy}
            onClick={() => { void triggerOutlookUpdateNow(); }}
            style={{ fontSize: '0.82rem', padding: '7px 14px' }}
          >
            {outlookUpdateBusy
              ? (isDe ? 'Wird aktualisiert…' : 'Updating…')
              : (isDe ? `Termin von „${tabTitle}" aktualisieren` : `Update appointment of „${tabTitle}"`)}
          </button>
          {showAll && (
            <button
              type="button"
              className="btn btn-secondary"
              disabled={outlookUpdateBusy}
              onClick={() => { void triggerOutlookUpdateAll(); }}
              style={{ fontSize: '0.82rem', padding: '7px 14px' }}
            >
              {isDe
                ? `Alle ${allTargets.length} Termine aktualisieren`
                : `Update all ${allTargets.length} appointments`}
            </button>
          )}
        </div>
        {missingTargets.length > 0 && (
          <div style={{
            marginTop: 8, padding: '8px 10px', borderRadius: 6, fontSize: '0.76rem', lineHeight: 1.5,
            background: '#fff8e6', border: '1px solid #e0b34d', color: '#7a5a12',
          }}>
            {isDe
              ? <>Für <strong>{missingTargets.length} von {totalTargets}</strong> Terminen dieses Events gibt es noch <strong>keinen</strong> Kalendereintrag — deshalb steht oben nur „{allTargets.length}“: {missingTargets.map(m => m.title || '?').join(', ')}. Häufigste Ursache: das {childTermSingular || 'Sub-Event'} wurde ohne Start-/Endzeit gespeichert, dann kann kein Termin erzeugt werden. {missingSubIds.length > 0 ? <>Der Knopf unten legt die fehlenden Termine jetzt an — Sub-Events ohne eigene Zeiten übernehmen dabei die Zeiten des Hauptevents. Anmeldungen und Teilnehmerlisten bleiben unverändert.</> : <>Für das Hauptevent selbst lässt sich das hier nicht nachholen — bitte beim Support melden.</>}</>
              : <>There is <strong>no</strong> calendar entry yet for <strong>{missingTargets.length} of {totalTargets}</strong> appointments of this event — that is why it says „{allTargets.length}“ above: {missingTargets.map(m => m.title || '?').join(', ')}. Most common cause: the sub-event was saved without a start/end time, so no appointment can be created. {missingSubIds.length > 0 ? <>The button below creates the missing appointments now — sub-events without their own times inherit the main event&apos;s times. Registrations and attendee lists stay untouched.</> : <>This cannot be repaired here for the main event itself — please contact support.</>}</>}
            {missingSubIds.length > 0 && (
              <div style={{ marginTop: 8 }}>
                <button
                  type="button"
                  className="btn btn-primary"
                  disabled={outlookUpdateBusy}
                  onClick={() => { void createMissingOutlookAppointments(); }}
                  style={{ fontSize: '0.82rem', padding: '7px 14px' }}
                >
                  {outlookUpdateBusy
                    ? (isDe ? 'Wird angelegt…' : 'Creating…')
                    : (isDe ? `${missingSubIds.length} fehlende Termine jetzt anlegen` : `Create ${missingSubIds.length} missing appointments now`)}
                </button>
              </div>
            )}
          </div>
        )}
        {outlookUpdateDone && (
          <div style={{
            marginTop: 8, padding: '6px 10px', borderRadius: 6, fontSize: '0.76rem', fontWeight: 600,
            background: '#f1f7e8', border: '1px solid var(--dex-green, #86bc25)', color: 'var(--dex-green-dark, #4a7c1f)',
          }}>
            ✓ {outlookUpdateDone}
          </div>
        )}
        <div style={{ fontSize: '0.74rem', color: 'var(--dex-gray-600)', marginTop: 8, lineHeight: 1.5 }}>
          {isDe
            ? <>Nur nötig, wenn der Kalendereintrag der Teilnehmer noch veraltet ist — der Termin wird dann mit dem zuletzt <strong>gespeicherten</strong> Stand neu verschickt. Erst speichern, dann klicken. <strong>Wichtig:</strong> {childTermPlural || 'Sub-Events'} haben eigene Termine; dieser Knopf betrifft nur „{tabTitle}“{showAll ? ' — für alle auf einmal den zweiten Knopf nutzen' : ''}. Der Kasten bleibt dauerhaft stehen, er ist keine Fehlermeldung.</>
            : <>Only needed if the attendees’ calendar entry is still outdated — the appointment is re-sent with the last <strong>saved</strong> state. Save first, then click. <strong>Important:</strong> sub-events have their own appointments; this button only affects „{tabTitle}“{showAll ? ' — use the second button for all at once' : ''}. This box is always here, it is not an error message.</>}
        </div>
      </div>
    );
}

/* renderShowIfConfig — aus EventCreationPage.tsx ausgelagert (Zeilen 2471-2626 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface RenderShowIfConfigCtx {
  isDe: boolean;
  showAlert: (message: React.ReactNode, opts?: import("../../../context/DialogContext").AlertOptions) => void;
}

export function renderShowIfConfigImpl(ctx: RenderShowIfConfigCtx, field: CustomFieldInput, idx: number, allFields: CustomFieldInput[], onUpdate: (u: Partial<CustomFieldInput>) => void): React.ReactElement {
  const { isDe, showAlert } = ctx;
                      const candidateSources = allFields.slice(0, idx).filter(other =>
                        (other.type === 'select' || other.type === 'checkbox') && (other.label || '').trim().length > 0
                      );
                      const sourceField = field.showIf
                        ? allFields.find(o => o.id === field.showIf!.fieldId)
                        : null;
                      const removeShowIf = (): void => {
                        // showIf gezielt löschen: updateCustomField macht ein
                        // shallow-merge, also setzen wir undefined und filtern
                        // beim Save raus.
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        onUpdate({ showIf: undefined as any });
                      };
                      return (
                        <div style={{ marginLeft: 32, marginTop: 10, padding: '10px 12px', background: 'rgba(21,101,192,0.04)', border: '1px dashed var(--dex-gray-300)', borderRadius: 8 }}>
                          {!field.showIf ? (
                            <span style={{ display: 'inline-flex', alignItems: 'center' }}>
                              <button
                                type="button"
                                onClick={() => {
                                  if (candidateSources.length === 0) {
                                    showAlert(isDe
                                      ? 'Es gibt noch kein Dropdown- oder Checkbox-Feld VOR diesem hier, an das die Sichtbarkeit geknüpft werden könnte. Lege zuerst ein passendes Feld weiter oben an.'
                                      : 'There is no dropdown or checkbox field BEFORE this one yet that visibility could depend on. Please add a suitable field above first.');
                                    return;
                                  }
                                  const first = candidateSources[0];
                                  onUpdate({
                                    showIf: {
                                      fieldId: first.id,
                                      values: first.type === 'checkbox' ? ['true'] : (first.options[0] ? [first.options[0]] : []),
                                    },
                                  });
                                }}
                                style={{
                                  background: 'none', border: 'none', padding: 0,
                                  color: 'var(--dex-green-dark, #4a7c1f)',
                                  fontSize: '0.8rem', fontWeight: 600, cursor: 'pointer',
                                }}
                              >
                                + {isDe ? 'Sichtbarkeitsbedingung hinzufügen' : 'Add visibility condition'}
                              </button>
                              <InfoTooltip
                                text={isDe
                                  ? 'Dieses Feld wird nur angezeigt, wenn die Antwort auf eine andere (zuvor angelegte) Frage einem von dir festgelegten Wert entspricht. Beispiel: „Roommate" wird nur gefragt, wenn die Frage „Zimmerart" mit „Doppelzimmer" beantwortet wurde. Andernfalls bleibt das Feld komplett verborgen — und blockiert auch nicht die Pflichtfeld-Validierung.'
                                  : 'This field is shown only when the answer to another (previously added) question matches a value you specify. Example: "Roommate" is only asked when the question "Room type" is answered with "Double room". Otherwise the field stays fully hidden — and does not block the required-field validation either.'}
                              />
                            </span>
                          ) : (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                              <div style={{ fontSize: '0.78rem', color: 'var(--dex-gray-600)', fontWeight: 600, display: 'flex', alignItems: 'center' }}>
                                {isDe ? 'Diese Frage nur anzeigen wenn:' : 'Only show this question when:'}
                                <InfoTooltip
                                  text={isDe
                                    ? 'Dieses Feld wird nur angezeigt, wenn die Antwort auf die Quell-Frage einem der gewählten Werte entspricht. Bei Mehrfachauswahl-Quellen reicht ein Treffer. Pflichtfeld-Validierung wird übersprungen, solange das Feld verborgen ist.'
                                    : 'This field is shown only when the answer to the source question matches one of the chosen values. With multi-select sources a single match is enough. Required-field validation is skipped as long as the field stays hidden.'}
                                />
                              </div>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                                <select
                                  className="form-select"
                                  value={field.showIf.fieldId}
                                  onChange={e => {
                                    const newSrc = allFields.find(o => o.id === e.target.value);
                                    if (!newSrc) return;
                                    onUpdate({
                                      showIf: {
                                        fieldId: newSrc.id,
                                        values: newSrc.type === 'checkbox' ? ['true'] : (newSrc.options[0] ? [newSrc.options[0]] : []),
                                      },
                                    });
                                  }}
                                  style={{ fontSize: '0.82rem', padding: '4px 8px', minWidth: 180, maxWidth: 320 }}
                                >
                                  {candidateSources.map(o => (
                                    <option key={o.id} value={o.id}>
                                      {allFields.findIndex(c => c.id === o.id) + 1}. {o.label}
                                    </option>
                                  ))}
                                  {/* fallback wenn die ausgewählte Quelle hinter dem Feld gelandet
                                      ist (z.B. nach einem Move) — option in der Liste anzeigen,
                                      aber als ungültig markiert lassen. */}
                                  {sourceField && !candidateSources.find(c => c.id === sourceField.id) && (
                                    <option value={sourceField.id} disabled>
                                      ⚠ {sourceField.label} ({isDe ? 'liegt hinter diesem Feld' : 'is positioned after this field'})
                                    </option>
                                  )}
                                </select>
                                <span style={{ fontSize: '0.82rem', color: 'var(--dex-gray-600)' }}>
                                  {isDe ? '=' : '='}
                                </span>
                                {sourceField && sourceField.type === 'checkbox' ? (
                                  <select
                                    className="form-select"
                                    value={field.showIf.values[0] || 'true'}
                                    onChange={e => onUpdate({
                                      showIf: { fieldId: field.showIf!.fieldId, values: [e.target.value] },
                                    })}
                                    style={{ fontSize: '0.82rem', padding: '4px 8px', minWidth: 130 }}
                                  >
                                    <option value="true">{isDe ? 'angehakt' : 'checked'}</option>
                                    <option value="false">{isDe ? 'nicht angehakt' : 'unchecked'}</option>
                                  </select>
                                ) : sourceField ? (
                                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                                    {(sourceField.options || []).filter(Boolean).map(opt => {
                                      const checked = field.showIf!.values.indexOf(opt) >= 0;
                                      return (
                                        <label
                                          key={opt}
                                          style={{
                                            display: 'inline-flex', alignItems: 'center', gap: 4,
                                            padding: '4px 10px', borderRadius: 999,
                                            fontSize: '0.78rem', cursor: 'pointer',
                                            border: `1px solid ${checked ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                                            background: checked ? 'rgba(134,188,37,0.10)' : '#fff',
                                            color: checked ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-600)',
                                            fontWeight: 600,
                                          }}
                                        >
                                          <input
                                            type="checkbox"
                                            checked={checked}
                                            onChange={() => {
                                              const next = checked
                                                ? field.showIf!.values.filter(v => v !== opt)
                                                : [...field.showIf!.values, opt];
                                              onUpdate({
                                                showIf: { fieldId: field.showIf!.fieldId, values: next },
                                              });
                                            }}
                                            style={{ display: 'none' }}
                                          />
                                          {checked ? '✓' : '○'} {opt}
                                        </label>
                                      );
                                    })}
                                  </div>
                                ) : null}
                                <button
                                  type="button"
                                  onClick={removeShowIf}
                                  title={isDe ? 'Bedingung entfernen' : 'Remove condition'}
                                  style={{
                                    background: 'none', border: 'none', cursor: 'pointer',
                                    color: 'var(--dex-red, #c00)', fontSize: '0.8rem',
                                    padding: '4px 6px', marginLeft: 'auto',
                                  }}
                                >
                                  ✕ {isDe ? 'entfernen' : 'remove'}
                                </button>
                              </div>
                            </div>
                          )}
                        </div>
                      );
}

/* renderGlobalScopeBar — aus EventCreationPage.tsx ausgelagert (Zeilen 3857-3919 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface RenderGlobalScopeBarCtx {
  activeScopeIdx: number;
  childTermPlural: string;
  childTermSingular: string;
  currentStep: number;
  isDe: boolean;
  renderPerEventTabStrip: (activeIdx: number, onChange: (idx: number) => void, mainLabel: string, ariaLabel: string) => React.ReactElement | null;
  SCOPE_AWARE_STEPS: number[];
  setScope: (idx: number) => void;
  subEvents: SubEventDraft[];
  subEventsOnlyMode: boolean;
  title: string;
}

export function renderGlobalScopeBarImpl(ctx: RenderGlobalScopeBarCtx): React.ReactElement | null {
  const { activeScopeIdx, childTermPlural, childTermSingular, currentStep, isDe, renderPerEventTabStrip, SCOPE_AWARE_STEPS, setScope, subEvents, subEventsOnlyMode, title } = ctx;
    if (subEvents.length === 0) return null;
    const named = subEvents.filter(s => (s.title || '').trim());
    // v29.21 (Audit): Nicht mehr verstecken, wenn ein Sub-Reiter aktiv ist.
    // Sequenz vorher: „Hinzufügen" (Draft ohne Titel) → „Bearbeiten"
    // (setScope(1)) → die Leiste war null, die Sub-Event-Liste hängt an
    // activeScopeIdx === 0 — keine Bedienung mehr, um zurück auf die Klammer
    // zu kommen. Die Reiter tragen für unbenannte Drafts den Fallback
    // „Sub-Event ohne Titel".
    if (named.length === 0 && activeScopeIdx === 0) return null;
    const applies = SCOPE_AWARE_STEPS.indexOf(currentStep) >= 0;
    const scopeIdx = Math.min(activeScopeIdx, subEvents.length);
    const mainLabel = `${subEventsOnlyMode ? (isDe ? 'Klammer' : 'Bracket') : (isDe ? 'Haupt-Event' : 'Main event')}: ${title || (isDe ? 'Ohne Titel' : 'Untitled')}`;
    // v28.90: Die Karte war grün getönt und mit grünem Rand abgesetzt — sie las
    // sich dadurch wie ein Status („hier stimmt etwas") statt wie das, was sie
    // ist: eine Navigation. Grün bleibt der aktiven Auswahl vorbehalten.
    // Ausserdem `overflow: hidden` (plus `minWidth: 0` weiter innen): Die
    // Reiter-Reihe schob sich bei vielen Sub-Events über den rechten Kartenrand
    // hinaus — Flex-Kinder haben `min-width: auto`, die Scroll-Fläche konnte
    // ihren Container also aufblähen.
    // v28.91: …und ganz ohne eigene Fläche. Die weiße Karte auf grauem Grund
    // war immer noch ein Kasten, der um Aufmerksamkeit konkurriert; die Reiter
    // selbst tragen ihre Form bereits. Transparent, nur Abstand.
    return (
      <div id="dex-scope-bar" style={{
        margin: '18px 0 0', padding: '12px 0 14px', borderRadius: 0,
        background: 'transparent',
        border: 'none',
        overflow: 'hidden',
      }}>
        {applies ? (
          renderPerEventTabStrip(
            scopeIdx,
            setScope,
            mainLabel,
            isDe ? 'Event-Ebene wechseln' : 'Switch event level',
          )
        ) : (
          <>
            <div style={{
              fontSize: '0.78rem', fontWeight: 700, letterSpacing: '0.03em',
              textTransform: 'uppercase', color: 'var(--dex-gray-500)', marginBottom: 6,
            }}>
              {isDe ? 'Welches (Sub-)Event bearbeitest du gerade?' : 'Which (sub-)event are you editing?'}
            </div>
            <div style={{
              display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap',
              padding: '9px 14px', borderRadius: 10,
              background: '#fff', border: '1px dashed var(--dex-gray-300)',
              fontSize: '0.84rem', color: 'var(--dex-gray-700)',
            }}>
              <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>
                {isDe ? 'Dieser Schritt gilt für das gesamte Event' : 'This step applies to the entire event'}
              </strong>
              <span style={{ color: 'var(--dex-gray-600)' }}>
                {isDe
                  ? `— ${subEventsOnlyMode ? 'Klammer' : 'Haupt-Event'} und alle ${named.length} ${named.length === 1 ? (childTermSingular || 'Sub-Event') : (childTermPlural || 'Sub-Events')} gemeinsam. Eine Auswahl gibt es hier nicht.`
                  : `— ${subEventsOnlyMode ? 'bracket' : 'main event'} and all ${named.length} sub-events together. There is nothing to pick here.`}
              </span>
            </div>
          </>
        )}
      </div>
    );
}

/* renderVisibilitySummaryBox — aus EventCreationPage.tsx ausgelagert (Zeilen 3968-4012 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface RenderVisibilitySummaryBoxCtx {
  isDe: boolean;
}

export function renderVisibilitySummaryBoxImpl(ctx: RenderVisibilitySummaryBoxCtx, locList: string[], audienceStr: string, mode: "AND" | "OR", excludedCount: number): React.ReactElement {
  const { isDe } = ctx;
    const locs = (locList || []).filter(Boolean);
    const auds = (audienceStr || '').split(',').map(s => s.trim()).filter(Boolean);
    // v28.76: Klartext statt Stichworten. Vorher stand hier „Sichtbar für
    // 1 Verteiler/Personen." — grammatisch schief und inhaltlich unklar
    // (1 Verteiler? 1 Person? beides?). Jetzt ein ganzer Satz, der sagt, WER
    // das Event sieht.
    let text: string;
    if (locs.length === 0 && auds.length === 0) {
      text = isDe
        ? 'Das Event sehen alle Mitarbeiter von Deloitte Deutschland.'
        : 'Everyone at Deloitte Germany can see this event.';
    } else {
      const parts: string[] = [];
      if (locs.length) {
        parts.push(isDe
          ? (locs.length === 1 ? `Mitarbeiter am Standort ${locs[0]}` : `Mitarbeiter an den Standorten ${locs.join(', ')}`)
          : (locs.length === 1 ? `employees at location ${locs[0]}` : `employees at the locations ${locs.join(', ')}`));
      }
      if (auds.length) {
        parts.push(isDe
          ? (auds.length === 1 ? 'die Mitglieder des hinterlegten Verteilers bzw. die hinterlegte Person' : `die Mitglieder der ${auds.length} hinterlegten Verteiler bzw. Personen`)
          : (auds.length === 1 ? 'the members of the selected distribution list or the selected person' : `the members of the ${auds.length} selected distribution lists / people`));
      }
      const joiner = parts.length > 1
        ? (mode === 'AND' ? (isDe ? ' und gleichzeitig ' : ' and at the same time ') : (isDe ? ' oder ' : ' or '))
        : '';
      text = (isDe ? 'Das Event sehen nur ' : 'Only ') + parts.join(joiner) + (isDe ? '.' : ' can see this event.');
    }
    return (
      <div style={{
        marginTop: 10, padding: '8px 12px', borderRadius: 8,
        background: 'rgba(134,188,37,0.07)', border: '1px solid var(--dex-green, #86bc25)',
        fontSize: '0.78rem', color: 'var(--dex-gray-700)', lineHeight: 1.5,
      }}>
        <strong style={{ color: 'var(--dex-green-dark, #4a7c1f)' }}>
          {isDe ? 'Aktuell eingestellt: ' : 'Currently configured: '}
        </strong>
        {text}
        {excludedCount > 0 && (
          <> {isDe
            ? `${excludedCount} Person${excludedCount === 1 ? ' ist' : 'en sind'} ausgeschlossen.`
            : `${excludedCount} ${excludedCount === 1 ? 'person is' : 'people are'} excluded.`}</>
        )}
      </div>
    );
}

/* renderKlammerVisibilityMismatch — aus EventCreationPage.tsx ausgelagert (Zeilen 4032-4103 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface RenderKlammerVisibilityMismatchCtx {
  audience: string;
  isDe: boolean;
  locationFilter: string;
  setLocationFilter: React.Dispatch<React.SetStateAction<string>>;
  subEvents: SubEventDraft[];
  subEventsOnlyMode: boolean;
}

export function renderKlammerVisibilityMismatchImpl(ctx: RenderKlammerVisibilityMismatchCtx): React.ReactElement | null {
  const { audience, isDe, locationFilter, setLocationFilter, subEvents, subEventsOnlyMode } = ctx;
    if (!subEventsOnlyMode || subEvents.length === 0) return null;
    const split = (s: string): string[] => (s || '').split(',').map(x => x.trim()).filter(Boolean);
    const parentLocs = split(locationFilter);
    const parentAuds = split(audience);
    const parentOpen = parentLocs.length === 0 && parentAuds.length === 0;
    const childLocSets = subEvents.map(s => split(s.locationFilter || ''));

    // (a) Klammer offen, aber JEDES Sub-Event schränkt ein.
    if (parentOpen && childLocSets.every(l => l.length > 0)) {
      const union = Array.from(new Set(childLocSets.reduce((a, b) => a.concat(b), [])));
      return (
        <div style={{
          marginTop: 10, padding: '10px 12px', borderRadius: 8,
          background: '#fff8e6', border: '1px solid #e0b34d', color: '#7a5a12',
          fontSize: '0.78rem', lineHeight: 1.55,
        }}>
          <strong>{isDe ? 'Die Klammer lässt mehr zu als ihre Sub-Events' : 'The bracket is broader than its sub-events'}</strong>
          <div style={{ marginTop: 4 }}>
            {isDe
              ? <>Hier ist <strong>kein Standort</strong> gesetzt, das Event ist also für alle sichtbar — aber <strong>alle {subEvents.length} Sub-Events</strong> sind auf {union.length === 1 ? <>den Standort <strong>{union[0]}</strong></> : <>die Standorte <strong>{union.join(', ')}</strong></>} beschränkt. Wer nicht dazugehört, sieht das Event in der Übersicht, findet darin aber <strong>nichts, wofür er sich anmelden kann</strong>.</>
              : <>No location is set here, so the event is visible to everyone — but <strong>all {subEvents.length} sub-events</strong> are restricted to {union.join(', ')}. People outside see the event but find nothing they can register for.</>}
          </div>
          <button
            type="button"
            className="btn btn-primary"
            style={{ fontSize: '0.78rem', padding: '5px 12px', marginTop: 8 }}
            onClick={() => setLocationFilter(union.join(', '))}
          >
            {isDe
              ? `Klammer ebenfalls auf ${union.join(', ')} setzen`
              : `Restrict the bracket to ${union.join(', ')} as well`}
          </button>
        </div>
      );
    }

    // (b) Ein Sub-Event lässt mehr zu, als die Klammer durchlässt.
    if (!parentOpen && parentLocs.length > 0) {
      const lc = (s: string): string => s.toLowerCase();
      const parentLc = parentLocs.map(lc);
      const offenders = subEvents
        .map((s, i) => ({ s, extra: childLocSets[i].filter(l => parentLc.indexOf(lc(l)) < 0) }))
        .filter(x => x.extra.length > 0);
      if (offenders.length > 0) {
        const extras = Array.from(new Set(offenders.reduce<string[]>((a, b) => a.concat(b.extra), [])));
        return (
          <div style={{
            marginTop: 10, padding: '10px 12px', borderRadius: 8,
            background: '#fff8e6', border: '1px solid #e0b34d', color: '#7a5a12',
            fontSize: '0.78rem', lineHeight: 1.55,
          }}>
            <strong>{isDe ? 'Einstellungen in Sub-Events, die nicht greifen können' : 'Sub-event settings that cannot take effect'}</strong>
            <div style={{ marginTop: 4 }}>
              {isDe
                ? <>{offenders.length === 1 ? 'Ein Sub-Event lässt' : `${offenders.length} Sub-Events lassen`} {extras.length === 1 ? <>den Standort <strong>{extras[0]}</strong></> : <>die Standorte <strong>{extras.join(', ')}</strong></>} zu — die Klammer aber nicht. Der Zugang läuft immer über die Klammer, deshalb bleiben diese Personen <strong>trotzdem draußen</strong>. Entweder hier ergänzen oder im Sub-Event entfernen.</>
                : <>{offenders.length} sub-event(s) allow {extras.join(', ')}, but the bracket does not. Access always goes through the bracket, so those people stay out anyway.</>}
            </div>
            <button
              type="button"
              className="btn btn-primary"
              style={{ fontSize: '0.78rem', padding: '5px 12px', marginTop: 8 }}
              onClick={() => setLocationFilter(Array.from(new Set(parentLocs.concat(extras))).join(', '))}
            >
              {isDe
                ? `${extras.join(', ')} hier ergänzen`
                : `Add ${extras.join(', ')} here`}
            </button>
          </div>
        );
      }
    }
    return null;
}

/* renderPreviewSection — aus EventCreationPage.tsx ausgelagert (Zeilen 4625-4699 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface RenderPreviewSectionCtx {
  customFields: CustomFieldInput[];
  endDate: string;
  eventImageUrl: string;
  formatPreviewDate: (val: string) => string;
  startDate: string;
  title: string;
}

export function renderPreviewSectionImpl(ctx: RenderPreviewSectionCtx, sectionId: string): React.ReactElement | null {
  const { customFields, endDate, eventImageUrl, formatPreviewDate, startDate, title } = ctx;
    switch (sectionId) {
      case 'event':
        return (
          <div className="registration-event" style={{ borderRadius: 'var(--dex-radius-lg)' }}>
            <div className="section-header section-header--red">Selected Event</div>
            <div className="registration-event__card">
              <div className="registration-event__image" style={{
                background: eventImageUrl
                  ? `url(${eventImageUrl}) center/cover`
                  : 'linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)',
              }}>
                <div className="registration-event__overlay">
                  <h4>{title || 'Event Titel'}</h4>
                  <p>{formatPreviewDate(startDate)} until<br />{formatPreviewDate(endDate)}</p>
                </div>
              </div>
            </div>
          </div>
        );
      case 'personal':
        return (
          <div className="registration-form" style={{ borderRadius: 'var(--dex-radius-lg)' }}>
            <div className="section-header">Personal Information</div>
            <div style={{ padding: '16px 20px' }}>
              <div className="form-group"><label className="form-label"><span className="required">*</span> Salutation</label><select className="form-select" disabled><option>Please select</option></select></div>
              <div className="form-group"><label className="form-label"><span className="required">*</span> First Name</label><input className="form-input" disabled placeholder="First Name" /></div>
              <div className="form-group"><label className="form-label"><span className="required">*</span> Surname</label><input className="form-input" disabled placeholder="Surname" /></div>
              <div className="form-group"><label className="form-label"><span className="required">*</span> E-Mail</label><input className="form-input" disabled placeholder="email@deloitte.de" /></div>
            </div>
          </div>
        );
      case 'specific':
        return (
          <div className="registration-specific" style={{ borderRadius: 'var(--dex-radius-lg)' }}>
            <div className="section-header">Event specific Information</div>
            <div style={{ padding: '16px 20px' }}>
              {customFields.filter(f => f.label).length === 0 ? (
                <p style={{ color: 'var(--dex-gray-400)', fontStyle: 'italic', fontSize: '0.9rem' }}>No additional information required.</p>
              ) : (
                customFields.filter(f => f.label).map(field => (
                  <div className="form-group" key={field.id}>
                    <label className="form-label">{field.required && <span className="required">*</span>}{field.label}</label>
                    {field.type === 'select' && field.multi ? (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 4, padding: 8, border: '1px solid var(--dex-gray-200)', borderRadius: 6, background: '#fff' }}>
                        {field.options.map(o => o.trim()).filter(Boolean).map(opt => (
                          <label key={opt} style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '0.85rem', color: 'var(--dex-gray-600)' }}>
                            <input type="checkbox" disabled />
                            <span>{opt}</span>
                          </label>
                        ))}
                        <span style={{ fontSize: '0.7rem', color: 'var(--dex-gray-400)', marginTop: 2 }}>Mehrere Auswahl möglich</span>
                      </div>
                    ) : field.type === 'select' ? (
                      <select className="form-select" disabled><option>Please select</option>{field.options.map(o => o.trim()).filter(Boolean).map(opt => <option key={opt}>{opt}</option>)}</select>
                    ) : field.type === 'checkbox' ? (
                      <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '0.9rem' }}><input type="checkbox" disabled /> {field.label}</label>
                    ) : (
                      <input className="form-input" disabled placeholder={field.label} type={field.type === 'number' ? 'number' : 'text'} />
                    )}
                  </div>
                ))
              )}
            </div>
          </div>
        );
      case 'actions':
        return (
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 16 }}>
            <button className="btn btn-danger" disabled style={{ opacity: 0.5 }}><Trash2 size={16} /> Delete</button>
            <button className="btn btn-primary" disabled style={{ opacity: 0.5 }}><Send size={16} /> Register</button>
          </div>
        );
      default:
        return null;
    }
}

/* renderPerEventTabStrip — aus EventCreationPage.tsx ausgelagert (Zeilen 5370-5507 des
 * urspruenglichen Stands). Der Funktionskoerper ist zeichengleich uebernommen;
 * alles, was er aus dem Komponenten-Scope liest, kommt jetzt ueber `ctx` —
 * dasselbe Muster wie `svc` bei den EventService-Modulen. Das Objekt wird beim
 * Aufruf gebaut, nicht memoisiert: damit sieht die Funktion exakt die Werte des
 * laufenden Renders, wie die Closure vorher auch. */
export interface RenderPerEventTabStripCtx {
  ariaLabel: string;
  childTermPlural: string;
  childTermSingular: string;
  goToSubEventsMode: () => void;
  isDe: boolean;
  mainLabel: string;
  subEventCalendar: boolean;
  subEvents: SubEventDraft[];
  subEventsOnlyMode: boolean;
  title: string;
}

export function renderPerEventTabStripImpl(ctx: RenderPerEventTabStripCtx, activeIdx: number, onChange: (idx: number) => void): React.ReactElement | null {
  const { ariaLabel, childTermPlural, childTermSingular, goToSubEventsMode, isDe, mainLabel, subEventCalendar, subEvents, subEventsOnlyMode, title } = ctx;
    if (subEvents.length === 0) return null;
    // v22.5: Der „Haupt"/„Klammer"-Badge links im Tab trägt die Rolle bereits —
    // deshalb das doppelte „Klammer: …"/„Haupt-Event: …"-Präfix aus dem Label
    // strippen (sonst stand „KLAMMER  Klammer: …" doppelt da). Sub-Event-Tabs
    // zeigen nur den reinen Sub-Namen (ohne „<Hauptevent> | "-Präfix).
    const strippedMain = mainLabel.replace(/^(Klammer|Bracket|Haupt-Event|Main event):\s*/i, '').trim();
    const tabs: Array<{ label: string; isMain: boolean }> = [
      { label: strippedMain || mainLabel, isMain: true },
      ...subEvents.map(s => ({
        label: (shortSubEventTitle(s.title, title) || (isDe ? 'Sub-Event ohne Titel' : 'Untitled sub-event')).trim(),
        isMain: false,
      })),
    ];
    // v28.72: Geltungsbereich benennen. Die Reiter standen bisher ohne
    // Erklärung da — sie sehen aus wie eine Beschriftung („dieses Event hat
    // 5 Teile"), nicht wie eine Umschaltung. Organizer stellten deshalb alles
    // am ersten Reiter ein und wunderten sich, dass es für die anderen nicht
    // galt; manche merkten gar nicht, dass ihr Event eine Klammer ist. Zwei
    // Ergänzungen, beide am Blick des Nutzers ausgerichtet:
    //  - eine Frage ÜBER den Reitern, die die Bedienung benennt,
    //  - ein Hinweis UNTER den Reitern, direkt über den Feldern: für wen die
    //    Einstellungen gerade gelten und wo die anderen zu finden sind. Der
    //    steht bewusst bei den Feldern, weil dort hingeschaut wird — nicht
    //    oben in der Leiste.
    const subCount = subEvents.length;
    const activeIsMain = activeIdx === 0;
    const activeLabel = (tabs[activeIdx] || tabs[0]).label;
    const mainWord = subEventsOnlyMode ? (isDe ? 'Klammer' : 'bracket') : (isDe ? 'Haupt-Event' : 'main event');
    const otherSubs = activeIsMain ? subCount : subCount - 1;
    const scopeText = ((): React.ReactNode => {
      if (activeIsMain) {
        if (isDe) {
          return (
            <>Du bearbeitest gerade {subEventsOnlyMode ? <>die <strong>Klammer</strong></> : <>das <strong>Haupt-Event</strong></>} „{activeLabel}“.{' '}
              {subEventsOnlyMode
                ? <>Zur Klammer meldet sich niemand direkt an — Teilnehmer wählen eines der Sub-Events. </>
                : null}
              Die Einstellungen auf dieser Seite gelten <strong>ausschließlich für {subEventsOnlyMode ? 'die Klammer' : 'das Haupt-Event'}</strong>. {otherSubs === 1 ? 'Das andere Sub-Event stellst du' : `Die ${otherSubs} Sub-Events stellst du`} oben über {otherSubs === 1 ? 'seinen Reiter' : 'ihre Reiter'} <strong>separat</strong> ein.</>
          );
        }
        return (
          <>You are editing the <strong>{subEventsOnlyMode ? 'bracket' : 'main event'}</strong> „{activeLabel}“.{' '}
            {subEventsOnlyMode ? <>Nobody registers for the bracket itself — attendees pick one of the sub-events. </> : null}
            These settings apply <strong>only to it</strong>. The {otherSubs === 1 ? 'other sub-event' : `${otherSubs} sub-events`} are configured <strong>separately</strong> via {otherSubs === 1 ? 'its tab' : 'their tabs'} above.</>
        );
      }
      if (isDe) {
        return (
          <>Du bearbeitest gerade das <strong>Sub-Event</strong> „{activeLabel}“. Die Einstellungen auf dieser Seite gelten <strong>ausschließlich für dieses Sub-Event</strong>
            {otherSubs > 0
              ? <> — {mainWord === 'Klammer' ? 'die Klammer' : 'das Haupt-Event'} und {otherSubs === 1 ? 'das weitere Sub-Event' : `die ${otherSubs} weiteren Sub-Events`} stellst du oben über die Reiter separat ein.</>
              : <> — {mainWord === 'Klammer' ? 'die Klammer' : 'das Haupt-Event'} stellst du oben über den Reiter separat ein.</>}</>
        );
      }
      return (
        <>You are editing the <strong>sub-event</strong> „{activeLabel}“. These settings apply <strong>only to it</strong>
          {otherSubs > 0
            ? <> — the {subEventsOnlyMode ? 'bracket' : 'main event'} and the {otherSubs === 1 ? 'other sub-event' : `${otherSubs} other sub-events`} are configured separately via the tabs above.</>
            : <> — the {subEventsOnlyMode ? 'bracket' : 'main event'} is configured separately via its tab above.</>}</>
      );
    })();
    // v22.30: Rendering + Sticky-Pin + gefüllter Aktiv-Tab leben in der
    // Modul-Komponente StickyTabStrip (Hooks pro Instanz).
    return (
      <>
        <div style={{
          fontSize: '0.78rem', fontWeight: 700, letterSpacing: '0.03em',
          textTransform: 'uppercase', color: 'var(--dex-gray-500)', marginBottom: 6,
        }}>
          {isDe ? 'Welches (Sub-)Event bearbeitest du gerade?' : 'Which (sub-)event are you editing?'}
        </div>
        <StickyTabStrip
          tabs={tabs}
          activeIdx={activeIdx}
          onChange={onChange}
          ariaLabel={ariaLabel}
          mainBadge={subEventsOnlyMode ? (isDe ? 'Klammer' : 'Bracket') : (isDe ? 'Haupt' : 'Main')}
          klammer={subEventsOnlyMode}
          klammerWord={isDe ? 'Klammerevent' : 'bracket event'}
          // v29.23: Zähl-Badge rechts in der Klammer-Zeile — ersetzt die frei
          // schwebende Zahl neben den umbrechenden Reitern. Im Kalender-Modus
          // sind die Kinder „Termine", sonst gilt die Event-Bezeichnung.
          countBadge={subCount >= 2 ? `${subCount} ${subEventCalendar
            ? (isDe ? 'Termine' : 'dates')
            : (childTermPlural || (isDe ? 'Sub-Events' : 'sub-events'))}` : undefined}
          klammerInfo={
            <InfoTooltip
              placement="bottom"
              interactive
              text={isDe ? (
                <>
                  <strong>Klammerevent</strong> — zu diesem Event selbst meldet sich <strong>niemand</strong> an. Teilnehmer sehen nur die {childTermPlural || 'Sub-Events'} darunter und melden sich <strong>dort</strong> an. Der Eventname ist die Klammer darüber: Er erscheint in der Übersicht und fasst die {childTermPlural || 'Sub-Events'} zusammen.<br /><br />
                  Deshalb gibt es hier keine eigene Teilnehmerzahl und keine eigene Warteliste — beides pflegst du je {childTermSingular || 'Sub-Event'}.<br /><br />
                  <strong>Du willst, dass man sich auch zum Hauptevent anmelden kann?</strong> Dann stell die Anmeldung in Schritt 1 (&bdquo;Grundlagen&ldquo;) um —{' '}
                  <button
                    type="button"
                    onClick={() => goToSubEventsMode()}
                    style={{
                      background: 'none', border: 'none', padding: 0, font: 'inherit',
                      color: 'var(--dex-green-dark, #4a7c1f)', textDecoration: 'underline',
                      cursor: 'pointer', fontWeight: 700,
                    }}
                  >
                    hier direkt hinspringen
                  </button>.
                </>
              ) : (
                <>
                  <strong>Bracket event</strong> — <strong>nobody</strong> registers for this event itself. Attendees only see the sub-events below and register <strong>there</strong>. The event name is the bracket around them: it appears in the overview and groups the sub-events.<br /><br />
                  That is why there is no capacity and no waitlist at this level — you set both per sub-event.<br /><br />
                  <strong>Want people to be able to register for the main event too?</strong> Then change the registration mode in step 1 —{' '}
                  <button
                    type="button"
                    onClick={() => goToSubEventsMode()}
                    style={{
                      background: 'none', border: 'none', padding: 0, font: 'inherit',
                      color: 'var(--dex-green-dark, #4a7c1f)', textDecoration: 'underline',
                      cursor: 'pointer', fontWeight: 700,
                    }}
                  >
                    jump there directly
                  </button>.
                </>
              )}
            />
          }
        />
        <div style={{
          margin: '-6px 0 16px', padding: '10px 12px', borderRadius: 8,
          background: 'var(--dex-gray-50, #f8f9fa)',
          border: '1px solid var(--dex-gray-200)',
          borderLeft: '4px solid var(--dex-green, #86bc25)',
          fontSize: '0.82rem', lineHeight: 1.55, color: 'var(--dex-gray-700)',
        }}>
          {scopeText}
        </div>
      </>
    );
}

