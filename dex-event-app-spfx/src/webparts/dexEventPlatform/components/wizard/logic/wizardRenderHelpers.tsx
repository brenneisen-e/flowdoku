import * as React from 'react';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { StickyTabStrip } from '../../wizard/StickyTabStrip';
import { InfoTooltip } from '../../InfoTooltip';
import { SubEventDraft } from '../../wizard/wizardTypes';
import { AlertCircle, Check, Info, Plus, RefreshCw, Send, Trash2, Users, X } from '../../Icons';
import { CustomFieldInput } from '../../wizard/customFieldInput';
// v31.2: gemeinsame UI-Klassen (Callouts, Chips, Textknöpfe) — Hover kommt aus
// dem Stylesheet, nicht mehr aus Inline-Styles ohne Hover.
import { cx } from '../../dexUi';

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
        className="dex-ui-input dex-ui-input--sm"
        onChange={e => set(Math.max(min, Math.min(max, parseInt(e.target.value, 10) || def)))}
        // v31.2: fontSize inline, weil `.dexApp input` (0.95rem) die Klasse überstimmt.
        style={{ width: 84, fontSize: '0.84rem' }} />
    );
    const lbl: React.CSSProperties = { display: 'flex', flexDirection: 'column', gap: 4, fontSize: '0.74rem', fontWeight: 600, color: 'var(--dex-gray-600)' };
    // v31.2: Reihenfolge „Voreinstellung → Feinwerte → Vorschau". Die zwei
    // Voreinstellungen sind das, was fast jeder wählt; sie stehen deshalb
    // zuerst und als Chips (mit Hover), die drei Zahlenfelder folgen als
    // Feinjustierung. Grenzen stehen im Hilfetext statt nur im min/max.
    return (
      <div className="dex-ui-card dex-ui-card--soft" style={{ marginTop: 12, padding: '12px 14px' }}>
        <div className="dex-ui-section-title" style={{ marginBottom: 4 }}>
          {isDe ? 'Bildgröße im Kopf' : 'Header image size'}
        </div>
        <div className="dex-ui-help" style={{ marginTop: 0, marginBottom: 10 }}>
          {isDe ? 'Gilt für Mail und Outlook-Termin.' : 'Applies to the mail and the Outlook invite.'}
        </div>
        <div style={{ display: 'flex', gap: 18, flexWrap: 'wrap', alignItems: 'flex-start' }}>
          <div className="dex-ui-stack" style={{ flex: 1, minWidth: 240 }}>
            {/* v28.31: Beide Voreinstellungen zeigen jetzt an, WELCHE gerade
                aktiv ist. Vorher war „Volle Breite" immer gruen und „Standard"
                immer grau — auch wenn tatsaechlich 180/30/30 (= Standard) stand. */}
            <div className="dex-ui-inline">
              <button type="button" className={cx('dex-ui-chip', isFullWidthPreset && 'is-active')}
                onClick={() => setHeaderImageLayout({ width: 600, paddingV: 0, paddingH: 0 })}
                title={isDe ? 'Bild füllt den Kopf über die volle Breite' : 'Image fills the header full width'}>
                {isFullWidthPreset && <Check size={12} />}{isDe ? 'Volle Breite' : 'Full width'}
              </button>
              <button type="button" className={cx('dex-ui-chip', isDefaultPreset && 'is-active')}
                onClick={() => setHeaderImageLayout({ width: 180, paddingV: 30, paddingH: 30 })}
                title={isDe ? '180 px breit, 30 px Abstand rundum' : '180 px wide, 30 px padding all around'}>
                {isDefaultPreset && <Check size={12} />}{isDe ? 'Standard' : 'Default'}
              </button>
              <span className="dex-ui-muted">{isDe ? 'oder selbst einstellen:' : 'or set your own:'}</span>
            </div>
            <div style={{ display: 'flex', alignItems: 'flex-end', gap: 12, flexWrap: 'wrap' }}>
              <label style={lbl}>{isDe ? 'Breite (px)' : 'Width (px)'}{numInput(headerImageLayout.width, 80, 600, 180, n => setHeaderImageLayout(p => ({ ...p, width: n })))}</label>
              <label style={lbl}>{isDe ? 'Abstand seitlich' : 'Padding sides'}{numInput(headerImageLayout.paddingH, 0, 80, 0, n => setHeaderImageLayout(p => ({ ...p, paddingH: n })))}</label>
              <label style={lbl}>{isDe ? 'Abstand oben/unten' : 'Padding top/bottom'}{numInput(headerImageLayout.paddingV, 0, 80, 0, n => setHeaderImageLayout(p => ({ ...p, paddingV: n })))}</label>
            </div>
            <div className="dex-ui-help" style={{ marginTop: 0 }}>
              {isDe ? 'Breite 80–600 px, Abstände 0–80 px.' : 'Width 80–600 px, padding 0–80 px.'}
            </div>
          </div>
          {previewSrc && (
            <div style={{ width: PREV_W, flexShrink: 0, border: '1px solid var(--dex-gray-200)', borderRadius: 10, overflow: 'hidden', background: '#fff' }}>
              <div style={{ textAlign: 'center', padding: `${Math.round(headerImageLayout.paddingV * sc)}px ${Math.round(headerImageLayout.paddingH * sc)}px` }}>
                <img src={previewSrc} alt="" style={{ display: 'inline-block', width: '100%', maxWidth: Math.max(20, Math.round(headerImageLayout.width * sc)), height: 'auto' }} />
              </div>
              <div style={{ borderTop: '2px solid var(--dex-green, #86bc25)' }} />
              <div style={{ fontSize: '0.68rem', color: 'var(--dex-gray-500)', textAlign: 'center', padding: '3px 0' }}>{isDe ? 'So groß im Mail-Kopf (verkleinert)' : 'Size in the mail header (scaled)'}</div>
            </div>
          )}
        </div>
        {/* v28.29: sagt, WOHER das gezeigte Bild kommt (eigenes / vom Hauptevent
            geerbt / Standardlogo). Vorher zeigte die Vorschau kommentarlos das
            Event-Foto, obwohl gespeichert etwas anderes wurde. */}
        {note && (
          <div className="dex-ui-help" style={{ marginTop: 10 }}>
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
    // v31.2: Reihenfolge „Was macht das? → Knöpfe → Ergebnis". Der Erklärtext
    // schrumpft auf die zwei Sätze, die man vor dem Klick braucht (gespeicherter
    // Stand, erst speichern); der Rest (eigene Termine je Sub-Event, kein
    // Fehler-Kasten) steht im Tooltip. Kein Satz gestrichen.
    const missingNames = missingTargets.map(m => m.title || '?').join(', ');
    return (
      <div className="dex-ui-card dex-ui-card--soft" style={{ marginTop: 14 }}>
        <div className="dex-ui-label" style={{ marginBottom: 2 }}>
          <RefreshCw size={16} />
          {isDe ? 'Outlook-Termin neu verschicken' : 'Re-send the Outlook appointment'}
          <span className="dex-ui-label-optional">(optional)</span>
          <InfoTooltip
            text={isDe
              ? <><strong>Wichtig:</strong> {childTermPlural || 'Sub-Events'} haben eigene Termine; der erste Knopf betrifft nur &bdquo;{tabTitle}&ldquo;{showAll ? ' — für alle auf einmal den zweiten Knopf nutzen' : ''}. Der Kasten bleibt dauerhaft stehen, er ist keine Fehlermeldung.</>
              : <><strong>Important:</strong> sub-events have their own appointments; the first button only affects &bdquo;{tabTitle}&ldquo;{showAll ? ' — use the second button for all at once' : ''}. This box is always here, it is not an error message.</>}
          />
        </div>
        <div className="dex-ui-help" style={{ marginTop: 0, marginBottom: 10 }}>
          {isDe
            ? <>Nur nötig, wenn der Kalendereintrag der Teilnehmer noch veraltet ist: Der Termin geht mit dem zuletzt <strong>gespeicherten</strong> Stand neu raus — erst speichern, dann klicken.</>
            : <>Only needed if the attendees&rsquo; calendar entry is still outdated: the appointment is re-sent with the last <strong>saved</strong> state — save first, then click.</>}
        </div>
        <div className="dex-ui-inline">
          <button
            type="button"
            className="btn btn-secondary dex-ui-btn-sm"
            disabled={outlookUpdateBusy}
            onClick={() => { void triggerOutlookUpdateNow(); }}
          >
            {outlookUpdateBusy
              ? (isDe ? 'Wird aktualisiert…' : 'Updating…')
              : (isDe ? `Termin von „${tabTitle}“ aktualisieren` : `Update appointment of „${tabTitle}“`)}
          </button>
          {showAll && (
            <button
              type="button"
              className="btn btn-secondary dex-ui-btn-sm"
              disabled={outlookUpdateBusy}
              onClick={() => { void triggerOutlookUpdateAll(); }}
            >
              {isDe
                ? `Alle ${allTargets.length} Termine aktualisieren`
                : `Update all ${allTargets.length} appointments`}
            </button>
          )}
        </div>
        {missingTargets.length > 0 && (
          <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginTop: 10 }}>
            <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
            <div style={{ flex: 1, minWidth: 0 }}>
              <strong>
                {isDe
                  ? `${missingTargets.length} von ${totalTargets} Terminen haben noch keinen Kalendereintrag`
                  : `${missingTargets.length} of ${totalTargets} appointments have no calendar entry yet`}
              </strong>
              <div style={{ marginTop: 4 }}>
                {isDe
                  ? <>Deshalb steht oben nur &bdquo;{allTargets.length}&ldquo;: {missingNames}. Häufigste Ursache: das {childTermSingular || 'Sub-Event'} wurde ohne Start-/Endzeit gespeichert, dann kann kein Termin erzeugt werden. {missingSubIds.length > 0 ? <>Der Knopf legt die fehlenden Termine jetzt an — Sub-Events ohne eigene Zeiten übernehmen die Zeiten des Hauptevents. Anmeldungen und Teilnehmerlisten bleiben unverändert.</> : <>Für das Hauptevent selbst lässt sich das hier nicht nachholen — bitte beim Support melden.</>}</>
                  : <>That is why it says &bdquo;{allTargets.length}&ldquo; above: {missingNames}. Most common cause: the sub-event was saved without a start/end time, so no appointment can be created. {missingSubIds.length > 0 ? <>The button creates the missing appointments now — sub-events without their own times inherit the main event&apos;s times. Registrations and attendee lists stay untouched.</> : <>This cannot be repaired here for the main event itself — please contact support.</>}</>}
              </div>
              {missingSubIds.length > 0 && (
                <button
                  type="button"
                  className="btn btn-primary dex-ui-btn-sm"
                  disabled={outlookUpdateBusy}
                  onClick={() => { void createMissingOutlookAppointments(); }}
                  style={{ marginTop: 8 }}
                >
                  {outlookUpdateBusy
                    ? (isDe ? 'Wird angelegt…' : 'Creating…')
                    : (isDe ? `${missingSubIds.length} fehlende Termine jetzt anlegen` : `Create ${missingSubIds.length} missing appointments now`)}
                </button>
              )}
            </div>
          </div>
        )}
        {outlookUpdateDone && (
          <div className="dex-ui-callout dex-ui-callout--success" style={{ marginTop: 10, fontWeight: 600 }}>
            <span className="dex-ui-callout-icon"><Check size={16} /></span>
            <div>{outlookUpdateDone}</div>
          </div>
        )}
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
                      // v31.2: Ohne Bedingung nur eine Textknopf-Zeile (kein Kasten um
                      // einen einzelnen Link); mit Bedingung ein gestrichelter Kasten,
                      // der sich wie ein Satz liest: „Nur anzeigen, wenn [Frage] ist [Wert]".
                      // Werte als dex-ui-chip statt Eigenbau-Pillen — dieselbe Optik wie
                      // überall, und endlich ein Hover.
                      const srcSelectStyle: React.CSSProperties = { width: 'auto', minWidth: 180, maxWidth: 320, padding: '6px 32px 6px 10px', fontSize: '0.84rem' };
                      return (
                        <div style={!field.showIf
                          ? { marginLeft: 32, marginTop: 6 }
                          : { marginLeft: 32, marginTop: 10, padding: '10px 14px', background: 'var(--dex-gray-50, #fafafa)', border: '1px dashed var(--dex-gray-300)', borderRadius: 12 }}>
                          {!field.showIf ? (
                            <span style={{ display: 'inline-flex', alignItems: 'center' }}>
                              <button
                                type="button"
                                className="dex-ui-textbtn"
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
                              >
                                <Plus size={14} /> {isDe ? 'Nur unter einer Bedingung anzeigen' : 'Show only under a condition'}
                              </button>
                              <InfoTooltip
                                text={isDe
                                  ? 'Dieses Feld wird nur angezeigt, wenn die Antwort auf eine andere (zuvor angelegte) Frage einem von dir festgelegten Wert entspricht. Beispiel: „Roommate" wird nur gefragt, wenn die Frage „Zimmerart" mit „Doppelzimmer" beantwortet wurde. Andernfalls bleibt das Feld komplett verborgen — und blockiert auch nicht die Pflichtfeld-Validierung.'
                                  : 'This field is shown only when the answer to another (previously added) question matches a value you specify. Example: "Roommate" is only asked when the question "Room type" is answered with "Double room". Otherwise the field stays fully hidden — and does not block the required-field validation either.'}
                              />
                            </span>
                          ) : (
                            <div className="dex-ui-stack" style={{ gap: 8 }}>
                              <div className="dex-ui-label" style={{ marginBottom: 0, fontSize: '0.82rem' }}>
                                {isDe ? 'Diese Frage erscheint nur, wenn' : 'This question only appears when'}
                                <InfoTooltip
                                  text={isDe
                                    ? 'Dieses Feld wird nur angezeigt, wenn die Antwort auf die Quell-Frage einem der gewählten Werte entspricht. Bei Mehrfachauswahl-Quellen reicht ein Treffer. Pflichtfeld-Validierung wird übersprungen, solange das Feld verborgen ist.'
                                    : 'This field is shown only when the answer to the source question matches one of the chosen values. With multi-select sources a single match is enough. Required-field validation is skipped as long as the field stays hidden.'}
                                />
                              </div>
                              <div className="dex-ui-inline">
                                <select
                                  className="dex-ui-select"
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
                                  style={srcSelectStyle}
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
                                {/* v31.2: „ist" statt „=" — die Zeile liest sich als Satz:
                                    „Zimmerart ist Doppelzimmer". */}
                                <span className="dex-ui-muted" style={{ fontWeight: 600 }}>
                                  {isDe ? 'ist' : 'is'}
                                </span>
                                {sourceField && sourceField.type === 'checkbox' ? (
                                  <select
                                    className="dex-ui-select"
                                    value={field.showIf.values[0] || 'true'}
                                    onChange={e => onUpdate({
                                      showIf: { fieldId: field.showIf!.fieldId, values: [e.target.value] },
                                    })}
                                    style={{ ...srcSelectStyle, minWidth: 130 }}
                                  >
                                    <option value="true">{isDe ? 'angehakt' : 'checked'}</option>
                                    <option value="false">{isDe ? 'nicht angehakt' : 'unchecked'}</option>
                                  </select>
                                ) : sourceField ? (
                                  <div className="dex-ui-inline" style={{ gap: 4 }}>
                                    {(sourceField.options || []).filter(Boolean).map(opt => {
                                      const checked = field.showIf!.values.indexOf(opt) >= 0;
                                      return (
                                        <label key={opt} className={cx('dex-ui-chip', checked && 'is-active')}>
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
                                          {checked && <Check size={12} />}{opt}
                                        </label>
                                      );
                                    })}
                                  </div>
                                ) : null}
                                <button
                                  type="button"
                                  className="dex-ui-textbtn dex-ui-textbtn--danger"
                                  onClick={removeShowIf}
                                  title={isDe ? 'Bedingung entfernen' : 'Remove condition'}
                                  style={{ marginLeft: 'auto' }}
                                >
                                  <X size={14} /> {isDe ? 'Bedingung entfernen' : 'Remove condition'}
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
            {/* v31.2: Die Frage über dem Kasten passt zur Antwort darunter.
                Vorher fragte die Zeile „Welches (Sub-)Event bearbeitest du?",
                und der Kasten antwortete „hier gibt es keine Auswahl" — zwei
                Aussagen, die sich widersprechen. */}
            <div className="dex-ui-section-title" style={{ marginBottom: 6 }}>
              {isDe ? 'Für welches (Sub-)Event gilt dieser Schritt?' : 'Which (sub-)event does this step apply to?'}
            </div>
            <div className="dex-ui-callout dex-ui-callout--neutral">
              <span className="dex-ui-callout-icon"><Info size={16} /></span>
              <div>
                <strong style={{ color: 'var(--dex-gray-800)' }}>
                  {isDe ? 'Für das gesamte Event' : 'The entire event'}
                </strong>
                {isDe
                  ? ` — ${subEventsOnlyMode ? 'Klammer' : 'Haupt-Event'} und alle ${named.length} ${named.length === 1 ? (childTermSingular || 'Sub-Event') : (childTermPlural || 'Sub-Events')} gemeinsam. Eine Auswahl gibt es hier nicht.`
                  : ` — ${subEventsOnlyMode ? 'bracket' : 'main event'} and all ${named.length} sub-events together. There is nothing to pick here.`}
              </div>
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
    // v31.2: Info-Callout statt grüner Fläche — der Kasten fasst zusammen, er
    // bestätigt nichts. Grün bleibt der aktiven Auswahl darüber vorbehalten.
    return (
      <div className="dex-ui-callout dex-ui-callout--info" style={{ marginTop: 10 }}>
        <span className="dex-ui-callout-icon"><Users size={16} /></span>
        <div>
          <strong>{isDe ? 'So ist es eingestellt: ' : 'Current setting: '}</strong>
          {text}
          {excludedCount > 0 && (
            <> {isDe
              ? `${excludedCount} Person${excludedCount === 1 ? ' ist' : 'en sind'} ausgeschlossen.`
              : `${excludedCount} ${excludedCount === 1 ? 'person is' : 'people are'} excluded.`}</>
          )}
        </div>
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

    // v31.2: Beide Fälle als Warn-Callout (Titel als Aussage, eine Zeile Folge,
    // der Fix-Knopf direkt darunter). Der Knopf ist `btn-outline`, nicht
    // `btn-primary` — der einzige Primär-Knopf des Schritts bleibt „Weiter".
    const warnBox = (titleTxt: string, body: React.ReactNode, btnLabel: string, onFix: () => void): React.ReactElement => (
      <div className="dex-ui-callout dex-ui-callout--warn" style={{ marginTop: 10 }}>
        <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
        <div style={{ flex: 1, minWidth: 0 }}>
          <strong>{titleTxt}</strong>
          <div style={{ marginTop: 4 }}>{body}</div>
          <button type="button" className="btn btn-outline dex-ui-btn-sm" style={{ marginTop: 8 }} onClick={onFix}>
            {btnLabel}
          </button>
        </div>
      </div>
    );

    // (a) Klammer offen, aber JEDES Sub-Event schränkt ein.
    if (parentOpen && childLocSets.every(l => l.length > 0)) {
      const union = Array.from(new Set(childLocSets.reduce((a, b) => a.concat(b), [])));
      const unionTxt = union.length === 1
        ? <>{isDe ? 'den Standort' : 'the location'} <strong>{union[0]}</strong></>
        : <>{isDe ? 'die Standorte' : 'the locations'} <strong>{union.join(', ')}</strong></>;
      return warnBox(
        isDe ? 'Die Klammer lässt mehr zu als ihre Sub-Events' : 'The bracket is broader than its sub-events',
        isDe
          ? <>Hier ist <strong>kein Standort</strong> gesetzt, das Event sehen also alle — aber <strong>alle {subEvents.length} Sub-Events</strong> sind auf {unionTxt} beschränkt. Wer nicht dazugehört, sieht das Event, findet darin aber <strong>nichts zum Anmelden</strong>.</>
          : <>No location is set here, so everyone sees the event — but <strong>all {subEvents.length} sub-events</strong> are restricted to {unionTxt}. People outside see the event but find <strong>nothing to register for</strong>.</>,
        isDe ? `Klammer ebenfalls auf ${union.join(', ')} setzen` : `Restrict the bracket to ${union.join(', ')} as well`,
        () => setLocationFilter(union.join(', ')),
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
        const extrasTxt = extras.length === 1
          ? <>{isDe ? 'den Standort' : 'the location'} <strong>{extras[0]}</strong></>
          : <>{isDe ? 'die Standorte' : 'the locations'} <strong>{extras.join(', ')}</strong></>;
        return warnBox(
          isDe ? 'Ein Sub-Event lässt Standorte zu, die die Klammer sperrt' : 'A sub-event allows locations the bracket blocks',
          isDe
            ? <>{offenders.length === 1 ? 'Ein Sub-Event lässt' : `${offenders.length} Sub-Events lassen`} {extrasTxt} zu — die Klammer nicht. Der Zugang läuft immer über die Klammer, diese Personen bleiben also <strong>trotzdem draußen</strong>. Entweder hier ergänzen oder im Sub-Event entfernen.</>
            : <>{offenders.length === 1 ? 'One sub-event allows' : `${offenders.length} sub-events allow`} {extrasTxt} — the bracket does not. Access always goes through the bracket, so those people <strong>stay out anyway</strong>. Either add them here or remove them in the sub-event.</>,
          isDe ? `${extras.join(', ')} hier ergänzen` : `Add ${extras.join(', ')} here`,
          () => setLocationFilter(Array.from(new Set(parentLocs.concat(extras))).join(', ')),
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
    // v31.2: Zwei Sätze statt vier — erst „was du bearbeitest", dann „wo der
    // Rest ist". Der Klammer-Satz („niemand meldet sich zur Klammer an") bleibt,
    // weil er die häufigste Rückfrage vorwegnimmt. Kein Inhalt gestrichen.
    const scopeText = ((): React.ReactNode => {
      if (activeIsMain) {
        if (isDe) {
          return (
            <>Du bearbeitest {subEventsOnlyMode ? <>die <strong>Klammer</strong></> : <>das <strong>Haupt-Event</strong></>} &bdquo;{activeLabel}&ldquo;{subEventsOnlyMode ? <> — zur Klammer meldet sich niemand direkt an, Teilnehmer wählen eines der Sub-Events</> : null}.{' '}
              Alles auf dieser Seite gilt <strong>nur dafür</strong>; {otherSubs === 1 ? 'das andere Sub-Event' : `die ${otherSubs} Sub-Events`} stellst du oben über {otherSubs === 1 ? 'seinen Reiter' : 'ihre Reiter'} <strong>separat</strong> ein.</>
          );
        }
        return (
          <>You are editing the <strong>{subEventsOnlyMode ? 'bracket' : 'main event'}</strong> &bdquo;{activeLabel}&ldquo;{subEventsOnlyMode ? <> — nobody registers for the bracket itself, attendees pick one of the sub-events</> : null}.{' '}
            Everything on this page applies <strong>only to it</strong>; the {otherSubs === 1 ? 'other sub-event is' : `${otherSubs} sub-events are`} configured <strong>separately</strong> via {otherSubs === 1 ? 'its tab' : 'their tabs'} above.</>
        );
      }
      if (isDe) {
        return (
          <>Du bearbeitest das <strong>Sub-Event</strong> &bdquo;{activeLabel}&ldquo;. Alles auf dieser Seite gilt <strong>nur dafür</strong>
            {otherSubs > 0
              ? <> — {mainWord === 'Klammer' ? 'die Klammer' : 'das Haupt-Event'} und {otherSubs === 1 ? 'das weitere Sub-Event' : `die ${otherSubs} weiteren Sub-Events`} stellst du oben über die Reiter separat ein.</>
              : <> — {mainWord === 'Klammer' ? 'die Klammer' : 'das Haupt-Event'} stellst du oben über den Reiter separat ein.</>}</>
        );
      }
      return (
        <>You are editing the <strong>sub-event</strong> &bdquo;{activeLabel}&ldquo;. Everything on this page applies <strong>only to it</strong>
          {otherSubs > 0
            ? <> — the {subEventsOnlyMode ? 'bracket' : 'main event'} and the {otherSubs === 1 ? 'other sub-event' : `${otherSubs} other sub-events`} are configured separately via the tabs above.</>
            : <> — the {subEventsOnlyMode ? 'bracket' : 'main event'} is configured separately via its tab above.</>}</>
      );
    })();
    // v22.30: Rendering + Sticky-Pin + gefüllter Aktiv-Tab leben in der
    // Modul-Komponente StickyTabStrip (Hooks pro Instanz).
    return (
      <>
        <div className="dex-ui-section-title" style={{ marginBottom: 6 }}>
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
        {/* v31.2: neutraler Hinweiskasten statt grüner Kante — Grün bleibt dem
            aktiven Reiter darüber vorbehalten, der Text ist Orientierung. */}
        <div className="dex-ui-callout dex-ui-callout--neutral" style={{ margin: '-6px 0 16px' }}>
          <span className="dex-ui-callout-icon"><Info size={16} /></span>
          <div>{scopeText}</div>
        </div>
      </>
    );
}

