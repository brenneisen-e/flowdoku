/**
 * v31.2: Gemeinsame UI-Klassen für Wizard, Organizer Center und alle Modale —
 * EINMAL als <style> in document.head.
 *
 * Warum ein globales Stylesheet und kein SCSS-Modul:
 *
 *  1. Inline-Styles können kein `:hover` (CLAUDE.md). Bis v31.1 hatte deshalb
 *     jede zweite Karte ihren eigenen `hoverIdx`-State oder gar keinen
 *     Hover — und ein Element ohne Hover liest sich als Beschriftung, nicht
 *     als Knopf.
 *  2. Modale hängen per Portal an `document.body`, also AUSSERHALB des
 *     gescopten SCSS-Moduls (`.dexApp`). `Modal.tsx` injiziert aus demselben
 *     Grund seit v24.64 die Button-Styles selbst. Diese Datei ist die
 *     Verallgemeinerung davon: eine Quelle für beide Welten.
 *  3. 60 Dateien wurden in v31.2 parallel modernisiert. Ohne gemeinsamen
 *     Klassensatz hätte jede Datei ihre eigene Karte, ihren eigenen Chip und
 *     ihren eigenen Hover erfunden — und die Oberfläche wäre nach dem Umbau
 *     uneinheitlicher als vorher.
 *
 * Alle Klassen tragen das Präfix `dex-ui-`. Zustände: `.is-active`,
 * `.is-open`, `.is-disabled`, `.is-done`. Farben kommen aus den `--dex-*`-
 * Variablen (auf `:root` gespiegelt, s. DexEventPlatform.module.scss), mit
 * festen Fallbacks — dieselbe Regel wie bei `.btn`.
 *
 * Die Bedienungsanleitung für diese Klassen steht in `docs/ui-leitfaden.md`.
 * Wer eine Klasse ergänzt: dort eintragen, sonst benutzt sie niemand.
 */

export const DEX_UI_STYLE_ID = 'dex-ui-global-styles';

const G = 'var(--dex-green, #86bc25)';
const GD = 'var(--dex-green-dark, #6b9a1e)';
const GDT = 'var(--dex-green-darker, #4a7c1f)';
const G100 = 'var(--dex-gray-100, #f5f5f5)';
const G200 = 'var(--dex-gray-200, #e8e8e8)';
const G300 = 'var(--dex-gray-300, #d1d1d1)';
const G400 = 'var(--dex-gray-400, #a0a0a0)';
const G500 = 'var(--dex-gray-500, #808080)';
const G600 = 'var(--dex-gray-600, #666666)';
const G700 = 'var(--dex-gray-700, #444444)';
const G800 = 'var(--dex-gray-800, #333333)';
const SOFT = 'var(--dex-gray-50, #fafafa)';
const EASE = '0.18s ease';

export const DEX_UI_CSS = `
/* ---- Bewegung ------------------------------------------------------- */
@keyframes dexUiFadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: none; } }
@keyframes dexUiModalIn { from { opacity: 0; transform: translateY(10px) scale(0.98); } to { opacity: 1; transform: none; } }
@keyframes dexUiPulse { 0% { box-shadow: 0 0 0 0 rgba(134,188,37,0.55); } 70% { box-shadow: 0 0 0 10px rgba(134,188,37,0); } 100% { box-shadow: 0 0 0 0 rgba(134,188,37,0); } }
.dex-ui-fade-in { animation: dexUiFadeIn 0.25s ease-out both; }
.dex-ui-modal-card { animation: dexUiModalIn 0.22s cubic-bezier(0.2, 0.8, 0.2, 1) both; }
.dex-ui-pulse { animation: dexUiPulse 1.8s ease-out infinite; }
@media (prefers-reduced-motion: reduce) {
  .dex-ui-fade-in, .dex-ui-modal-card, .dex-ui-pulse { animation: none !important; }
}

/* ---- Karten --------------------------------------------------------- */
.dex-ui-card {
  background: #fff; border: 1px solid ${G200}; border-radius: 14px; padding: 16px 18px;
  box-sizing: border-box;
  transition: border-color ${EASE}, box-shadow ${EASE}, transform ${EASE}, background ${EASE};
}
.dex-ui-card--soft { background: ${SOFT}; }
.dex-ui-card--accent { border-left: 4px solid ${G}; }
.dex-ui-card--hover:hover { border-color: rgba(134,188,37,0.55); box-shadow: 0 6px 20px rgba(0,0,0,0.07); transform: translateY(-1px); }
.dex-ui-card--muted { opacity: 0.6; }
.dex-ui-card--muted:hover { opacity: 1; }

/* ---- Auswahl-Kachel (eine von mehreren Optionen) --------------------- */
.dex-ui-choice {
  display: flex; align-items: flex-start; gap: 12px; width: 100%; text-align: left;
  cursor: pointer; padding: 14px 16px; border-radius: 12px; box-sizing: border-box;
  border: 1.5px solid ${G200}; background: #fff; color: inherit;
  font: inherit; font-family: inherit; line-height: 1.45;
  transition: border-color ${EASE}, background ${EASE}, box-shadow ${EASE}, transform ${EASE};
}
.dex-ui-choice:hover { border-color: rgba(134,188,37,0.6); background: rgba(134,188,37,0.05); box-shadow: 0 4px 14px rgba(0,0,0,0.06); }
.dex-ui-choice.is-active { border-color: ${G}; background: rgba(134,188,37,0.09); box-shadow: 0 0 0 3px rgba(134,188,37,0.15); }
.dex-ui-choice:focus-visible { outline: 2px solid ${G}; outline-offset: 2px; }
.dex-ui-choice:disabled, .dex-ui-choice.is-disabled { opacity: 0.55; cursor: not-allowed; box-shadow: none; transform: none; background: #fff; border-color: ${G200}; }
.dex-ui-choice-icon { flex-shrink: 0; width: 36px; height: 36px; border-radius: 10px; display: inline-flex; align-items: center; justify-content: center; background: rgba(134,188,37,0.12); color: ${GDT}; }
.dex-ui-choice.is-active .dex-ui-choice-icon { background: ${G}; color: #fff; }
.dex-ui-choice-body { flex: 1; min-width: 0; }
.dex-ui-choice-title { font-weight: 700; font-size: 0.9rem; color: ${G800}; }
.dex-ui-choice-desc { font-size: 0.8rem; color: ${G600}; margin-top: 3px; line-height: 1.5; }
.dex-ui-choice-check { flex-shrink: 0; width: 20px; height: 20px; border-radius: 50%; border: 2px solid ${G300}; display: inline-flex; align-items: center; justify-content: center; color: #fff; transition: all ${EASE}; }
.dex-ui-choice.is-active .dex-ui-choice-check { border-color: ${G}; background: ${G}; }

/* ---- Chips (umschaltbar) und Pills (nur Anzeige) ---------------------- */
.dex-ui-chip {
  display: inline-flex; align-items: center; gap: 6px; padding: 5px 12px; border-radius: 999px;
  border: 1px solid ${G200}; background: #fff; color: ${G600};
  font-size: 0.78rem; font-weight: 600; line-height: 1.2; cursor: pointer; font-family: inherit;
  transition: border-color ${EASE}, background ${EASE}, color ${EASE}, transform ${EASE};
}
.dex-ui-chip:hover { border-color: ${G}; color: ${GDT}; background: rgba(134,188,37,0.07); }
.dex-ui-chip.is-active { background: ${G}; border-color: ${G}; color: #fff; }
.dex-ui-chip.is-active:hover { background: ${GD}; border-color: ${GD}; }
.dex-ui-chip:disabled, .dex-ui-chip.is-disabled { opacity: 0.55; cursor: not-allowed; }
.dex-ui-chip:focus-visible { outline: 2px solid ${G}; outline-offset: 2px; }
.dex-ui-pill { display: inline-flex; align-items: center; gap: 5px; padding: 4px 10px; border-radius: 999px; font-size: 0.76rem; font-weight: 600; line-height: 1.2; white-space: nowrap; }
.dex-ui-pill--green { background: rgba(134,188,37,0.14); color: ${GDT}; }
.dex-ui-pill--gray { background: ${G100}; color: ${G600}; }
.dex-ui-pill--orange { background: rgba(237,139,0,0.14); color: var(--dex-orange-dark, #b35a00); }
.dex-ui-pill--red { background: var(--dex-red-light, #fce8e6); color: #b3261e; }
.dex-ui-pill--blue { background: rgba(56,96,178,0.12); color: var(--dex-blue, #3860b2); }

/* ---- Knöpfe ohne Fläche --------------------------------------------- */
.dex-ui-iconbtn {
  display: inline-flex; align-items: center; justify-content: center; width: 32px; height: 32px;
  border-radius: 50%; border: none; background: transparent; color: ${G600}; cursor: pointer; padding: 0;
  flex-shrink: 0; transition: background ${EASE}, color ${EASE}, transform ${EASE};
}
.dex-ui-iconbtn:hover { background: ${G100}; color: ${G800}; }
.dex-ui-iconbtn:active { transform: scale(0.94); }
.dex-ui-iconbtn--danger:hover { background: var(--dex-red-light, #fce8e6); color: var(--dex-red, #da291c); }
.dex-ui-iconbtn--green:hover { background: rgba(134,188,37,0.14); color: ${GDT}; }
.dex-ui-iconbtn:disabled { opacity: 0.45; cursor: not-allowed; background: transparent; }
.dex-ui-iconbtn:focus-visible { outline: 2px solid ${G}; outline-offset: 1px; }
.dex-ui-textbtn {
  display: inline-flex; align-items: center; gap: 6px; background: none; border: none; padding: 5px 8px; border-radius: 8px;
  color: ${GDT}; font-weight: 600; font-size: 0.82rem; cursor: pointer; font-family: inherit; line-height: 1.3;
  transition: background ${EASE}, color ${EASE};
}
.dex-ui-textbtn:hover { background: rgba(134,188,37,0.10); }
.dex-ui-textbtn--muted { color: ${G600}; }
.dex-ui-textbtn--muted:hover { background: ${G100}; color: ${G800}; }
.dex-ui-textbtn--danger { color: var(--dex-red, #da291c); }
.dex-ui-textbtn--danger:hover { background: var(--dex-red-light, #fce8e6); }
.dex-ui-textbtn:disabled { opacity: 0.5; cursor: not-allowed; background: none; }
.dex-ui-textbtn:focus-visible { outline: 2px solid ${G}; outline-offset: 1px; }
/* Kompakte Variante der bestehenden .btn-Klassen — braucht !important, weil
   .dex-modal-overlay .btn seine Maße ebenfalls mit !important setzt. */
.dex-ui-btn-sm { padding: 6px 14px !important; font-size: 0.82rem !important; border-radius: 10px !important; }
.btn:focus-visible { outline: 2px solid ${GD}; outline-offset: 2px; }
.btn:active:not(:disabled) { transform: translateY(1px); }

/* ---- Zeilen und Listen ---------------------------------------------- */
.dex-ui-row { display: flex; align-items: center; gap: 12px; padding: 10px 12px; border-radius: 10px; transition: background 0.15s ease; }
.dex-ui-row:hover { background: ${SOFT}; }
.dex-ui-row--bordered { border-bottom: 1px solid ${G100}; border-radius: 0; }
.dex-ui-row--bordered:last-child { border-bottom: none; }
.dex-ui-row-main { flex: 1; min-width: 0; }
.dex-ui-row-title { font-weight: 600; font-size: 0.88rem; color: ${G800}; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.dex-ui-row-sub { font-size: 0.76rem; color: ${G500}; margin-top: 2px; }
.dex-ui-row-actions { display: inline-flex; align-items: center; gap: 4px; flex-shrink: 0; opacity: 0.7; transition: opacity ${EASE}; }
.dex-ui-row:hover .dex-ui-row-actions { opacity: 1; }
.dex-ui-drag-handle { cursor: grab; color: ${G400}; display: inline-flex; align-items: center; padding: 4px; border-radius: 6px; transition: color ${EASE}, background ${EASE}; }
.dex-ui-drag-handle:hover { color: ${G700}; background: ${G100}; }
.dex-ui-drag-handle:active { cursor: grabbing; }

/* ---- Schalter-Zeile (Checkbox + Text) ------------------------------- */
.dex-ui-toggle-row {
  display: flex; align-items: flex-start; gap: 12px; padding: 12px 14px; border-radius: 12px; box-sizing: border-box;
  border: 1px solid ${G200}; background: #fff; cursor: pointer;
  transition: border-color ${EASE}, background ${EASE}, box-shadow ${EASE};
}
.dex-ui-toggle-row:hover { border-color: rgba(134,188,37,0.55); background: rgba(134,188,37,0.04); }
.dex-ui-toggle-row.is-active { border-color: rgba(134,188,37,0.7); background: rgba(134,188,37,0.07); }
.dex-ui-toggle-row.is-disabled { opacity: 0.55; cursor: not-allowed; background: #fff; border-color: ${G200}; }
.dex-ui-toggle-row > input[type='checkbox'], .dex-ui-toggle-row > input[type='radio'] { margin-top: 3px; cursor: pointer; flex-shrink: 0; accent-color: #86bc25; }
.dex-ui-toggle-row-body { flex: 1; min-width: 0; }
.dex-ui-toggle-row-title { font-weight: 600; font-size: 0.88rem; color: ${G800}; display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
.dex-ui-toggle-row-desc { font-size: 0.78rem; color: ${G500}; margin-top: 4px; line-height: 1.5; }

/* ---- Schalter (Switch) statt Checkbox für Ein/Aus ------------------- */
.dex-ui-switch { display: inline-flex; align-items: center; gap: 10px; cursor: pointer; user-select: none; }
.dex-ui-switch > input { position: absolute; opacity: 0; width: 0; height: 0; margin: 0; }
.dex-ui-switch-track { position: relative; width: 40px; height: 22px; border-radius: 999px; background: ${G300}; flex-shrink: 0; transition: background ${EASE}; }
.dex-ui-switch-track::after { content: ''; position: absolute; top: 2px; left: 2px; width: 18px; height: 18px; border-radius: 50%; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,0.25); transition: transform ${EASE}; }
.dex-ui-switch > input:checked + .dex-ui-switch-track { background: ${G}; }
.dex-ui-switch > input:checked + .dex-ui-switch-track::after { transform: translateX(18px); }
.dex-ui-switch > input:focus-visible + .dex-ui-switch-track { box-shadow: 0 0 0 3px rgba(134,188,37,0.3); }
.dex-ui-switch:hover .dex-ui-switch-track { background: ${G400}; }
.dex-ui-switch:hover > input:checked + .dex-ui-switch-track { background: ${GD}; }
.dex-ui-switch > input:disabled + .dex-ui-switch-track { opacity: 0.5; }
.dex-ui-switch.is-disabled { cursor: not-allowed; opacity: 0.7; }
.dex-ui-switch-label { font-weight: 600; font-size: 0.88rem; color: ${G800}; }

/* ---- Abschnitte, Felder, Raster ------------------------------------ */
.dex-ui-section { margin: 22px 0 0; }
.dex-ui-section:first-child { margin-top: 0; }
.dex-ui-section-title {
  display: flex; align-items: center; gap: 10px; margin: 0 0 10px;
  font-size: 0.72rem; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; color: ${G500};
}
.dex-ui-section-title::after { content: ''; flex: 1; height: 1px; background: ${G200}; }
.dex-ui-section-desc { font-size: 0.82rem; color: ${G500}; margin: -4px 0 12px; line-height: 1.5; }
.dex-ui-field { margin-bottom: 16px; }
.dex-ui-field:last-child { margin-bottom: 0; }
.dex-ui-label { display: flex; align-items: center; gap: 6px; margin-bottom: 6px; font-weight: 600; font-size: 0.88rem; color: ${G800}; }
.dex-ui-label-optional { font-weight: 500; font-size: 0.74rem; color: ${G400}; }
.dex-ui-help { font-size: 0.78rem; color: ${G500}; margin-top: 5px; line-height: 1.45; }
.dex-ui-grid-2 { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }
.dex-ui-grid-3 { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 14px; }
.dex-ui-grid-auto { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; }
@media (max-width: 768px) { .dex-ui-grid-2, .dex-ui-grid-3 { grid-template-columns: 1fr; } }
.dex-ui-divider { height: 1px; background: ${G200}; margin: 16px 0; border: none; }
.dex-ui-muted { font-size: 0.8rem; color: ${G500}; }
.dex-ui-stack { display: flex; flex-direction: column; gap: 10px; }
.dex-ui-inline { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }

/* ---- Kompakte Eingaben (Modale, Zeilen) ----------------------------- */
.dex-ui-input, .dex-ui-select, .dex-ui-textarea {
  width: 100%; padding: 9px 12px; border: 1.5px solid ${G200}; border-radius: 10px; box-sizing: border-box;
  font-size: 0.9rem; font-family: inherit; line-height: 1.4; background: #fff; color: ${G800}; outline: none;
  transition: border-color ${EASE}, box-shadow ${EASE}, background ${EASE};
}
.dex-ui-input:hover:not(:focus):not(:disabled), .dex-ui-select:hover:not(:focus):not(:disabled), .dex-ui-textarea:hover:not(:focus):not(:disabled) { border-color: ${G300}; }
.dex-ui-input:focus, .dex-ui-select:focus, .dex-ui-textarea:focus { border-color: ${G}; box-shadow: 0 0 0 3px rgba(134,188,37,0.15); }
.dex-ui-input:disabled, .dex-ui-select:disabled, .dex-ui-textarea:disabled { background: ${G100}; color: ${G500}; cursor: not-allowed; }
.dex-ui-input::placeholder, .dex-ui-textarea::placeholder { color: ${G400}; }
.dex-ui-select { appearance: none; -webkit-appearance: none; padding-right: 32px; background-image: url("data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3e%3cpath fill='%23666' d='M6 8L1 3h10z'/%3e%3c/svg%3e"); background-repeat: no-repeat; background-position: right 12px center; }
.dex-ui-textarea { min-height: 90px; resize: vertical; }
.dex-ui-input--sm { padding: 6px 10px; font-size: 0.84rem; border-radius: 8px; }
.form-input:hover:not(:focus):not(:disabled), .form-select:hover:not(:focus):not(:disabled), .form-textarea:hover:not(:focus):not(:disabled) { border-color: ${G300}; }

/* ---- Aufklapper ------------------------------------------------------ */
.dex-ui-disclosure {
  display: flex; align-items: center; gap: 8px; width: 100%; text-align: left; background: none; border: none;
  padding: 8px 6px; margin: 0 -6px; cursor: pointer; color: ${G700}; font: inherit; font-family: inherit;
  font-weight: 600; font-size: 0.84rem; border-radius: 8px; transition: background ${EASE}, color ${EASE};
}
.dex-ui-disclosure:hover { color: ${GDT}; background: rgba(134,188,37,0.06); }
.dex-ui-disclosure-chevron { display: inline-flex; align-items: center; color: ${G500}; transition: transform 0.2s ease; flex-shrink: 0; }
/* Ruhezustand ist ChevronDown (18 von 19 Aufrufern) — geöffnet zeigt er nach
   oben. Die 90°-Drehung des ersten Entwurfs zeigte bei ChevronDown nach links. */
.dex-ui-disclosure.is-open .dex-ui-disclosure-chevron { transform: rotate(180deg); }
.dex-ui-disclosure-count { margin-left: auto; font-weight: 600; font-size: 0.74rem; color: ${G500}; }
.dex-ui-disclosure-body { padding: 4px 0 8px 2px; }

/* ---- Hinweiskästen --------------------------------------------------- */
.dex-ui-callout { display: flex; gap: 10px; align-items: flex-start; padding: 10px 14px; border-radius: 12px; font-size: 0.82rem; line-height: 1.5; border: 1px solid; box-sizing: border-box; }
.dex-ui-callout--info { background: #f4f9fd; border-color: #cfe3f3; color: #22516f; }
.dex-ui-callout--success { background: rgba(134,188,37,0.10); border-color: rgba(134,188,37,0.45); color: #3f6b17; }
.dex-ui-callout--warn { background: #fff7e6; border-color: #f5c77a; color: #7a4a00; }
.dex-ui-callout--danger { background: var(--dex-red-light, #fce8e6); border-color: #f2b1ab; color: #9b2018; }
.dex-ui-callout--neutral { background: ${SOFT}; border-color: ${G200}; color: ${G600}; }
.dex-ui-callout-icon { flex-shrink: 0; display: inline-flex; margin-top: 1px; }

/* ---- Kennzahlen ------------------------------------------------------ */
.dex-ui-kpi { padding: 12px 14px; border-radius: 12px; background: ${SOFT}; border: 1px solid ${G200}; min-width: 0; transition: border-color ${EASE}, box-shadow ${EASE}; }
.dex-ui-kpi:hover { border-color: ${G300}; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
.dex-ui-kpi-value { font-size: 1.35rem; font-weight: 800; color: ${G800}; line-height: 1.1; letter-spacing: -0.01em; }
.dex-ui-kpi-label { font-size: 0.72rem; font-weight: 600; color: ${G500}; text-transform: uppercase; letter-spacing: 0.05em; margin-top: 4px; }
.dex-ui-kpi--green .dex-ui-kpi-value { color: ${GDT}; }
.dex-ui-kpi--orange .dex-ui-kpi-value { color: var(--dex-orange-dark, #b35a00); }

/* ---- Nummerierte Schritte (Dialoge mit Ablauf) ---------------------- */
.dex-ui-step { display: flex; gap: 12px; align-items: center; padding: 12px 14px; border-radius: 12px; border: 1px solid ${G200}; background: #fff; box-sizing: border-box; transition: border-color ${EASE}, box-shadow ${EASE}; }
.dex-ui-step:hover { border-color: rgba(134,188,37,0.5); box-shadow: 0 4px 14px rgba(0,0,0,0.05); }
.dex-ui-step-num { flex-shrink: 0; width: 28px; height: 28px; border-radius: 50%; background: ${G}; color: #fff; font-weight: 700; font-size: 0.8rem; display: inline-flex; align-items: center; justify-content: center; line-height: 1; }
.dex-ui-step.is-done .dex-ui-step-num { background: ${GDT}; }
.dex-ui-step.is-pending .dex-ui-step-num { background: ${G300}; }
.dex-ui-step-body { flex: 1; min-width: 0; }
.dex-ui-step-title { font-size: 0.88rem; font-weight: 600; color: ${G800}; }
.dex-ui-step-hint { font-size: 0.76rem; color: ${G500}; line-height: 1.45; margin-top: 2px; }
.dex-ui-step-action { flex-shrink: 0; }

/* ---- Segment-Reiter -------------------------------------------------- */
.dex-ui-tabs { display: inline-flex; padding: 3px; background: ${G100}; border-radius: 999px; gap: 2px; max-width: 100%; overflow-x: auto; }
.dex-ui-tab { padding: 6px 14px; border-radius: 999px; border: none; background: transparent; color: ${G600}; font-weight: 600; font-size: 0.8rem; cursor: pointer; font-family: inherit; white-space: nowrap; line-height: 1.3; transition: background ${EASE}, color ${EASE}, box-shadow ${EASE}; }
.dex-ui-tab:hover { color: ${G800}; background: rgba(255,255,255,0.7); }
.dex-ui-tab.is-active { background: #fff; color: ${G800}; box-shadow: 0 1px 3px rgba(0,0,0,0.12); }
.dex-ui-tab:focus-visible { outline: 2px solid ${G}; outline-offset: 1px; }
.dex-ui-tab:disabled { opacity: 0.5; cursor: not-allowed; }

/* ---- Tabellen -------------------------------------------------------- */
.dex-ui-table { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 0.84rem; }
.dex-ui-table th { text-align: left; font-size: 0.72rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; color: ${G500}; padding: 8px 10px; border-bottom: 1px solid ${G200}; background: ${SOFT}; white-space: nowrap; }
.dex-ui-table th:first-child { border-top-left-radius: 10px; }
.dex-ui-table th:last-child { border-top-right-radius: 10px; }
.dex-ui-table td { padding: 9px 10px; border-bottom: 1px solid ${G100}; vertical-align: middle; color: ${G800}; }
.dex-ui-table tbody tr { transition: background 0.15s ease; }
.dex-ui-table tbody tr:hover td { background: ${SOFT}; }
.dex-ui-table tbody tr:last-child td { border-bottom: none; }
.dex-ui-table-wrap { overflow-x: auto; border: 1px solid ${G200}; border-radius: 12px; }
.dex-ui-table-wrap .dex-ui-table th:first-child, .dex-ui-table-wrap .dex-ui-table th:last-child { border-radius: 0; }

/* ---- Avatare --------------------------------------------------------- */
.dex-ui-avatar { width: 32px; height: 32px; border-radius: 50%; object-fit: cover; background: ${G200}; flex-shrink: 0; display: inline-flex; align-items: center; justify-content: center; font-weight: 700; font-size: 0.78rem; color: ${G700}; overflow: hidden; }
.dex-ui-avatar--lg { width: 44px; height: 44px; font-size: 0.95rem; }
.dex-ui-avatar-stack { display: inline-flex; }
.dex-ui-avatar-stack .dex-ui-avatar { border: 2px solid #fff; margin-left: -8px; }
.dex-ui-avatar-stack .dex-ui-avatar:first-child { margin-left: 0; }

/* ---- Modal-Kopf und -Fuß --------------------------------------------- */
.dex-ui-modal-head { display: flex; align-items: flex-start; gap: 12px; padding-bottom: 12px; border-bottom: 1px solid ${G100}; }
.dex-ui-modal-head-icon { flex-shrink: 0; width: 40px; height: 40px; border-radius: 12px; display: inline-flex; align-items: center; justify-content: center; background: rgba(134,188,37,0.12); color: ${GDT}; }
.dex-ui-modal-title { margin: 0; font-size: 1.15rem; font-weight: 700; color: ${G800}; letter-spacing: -0.01em; line-height: 1.3; }
.dex-ui-modal-subtitle { margin: 4px 0 0; font-size: 0.85rem; color: ${G600}; line-height: 1.5; }
.dex-ui-modal-body { display: flex; flex-direction: column; gap: 14px; }
.dex-ui-modal-foot { display: flex; justify-content: flex-end; align-items: center; gap: 10px; flex-wrap: wrap; padding-top: 14px; border-top: 1px solid ${G100}; margin-top: 4px; }
.dex-ui-modal-foot--split { justify-content: space-between; }
.dex-ui-modal-foot-left { display: inline-flex; align-items: center; gap: 8px; flex-wrap: wrap; margin-right: auto; }

/* ---- Leere Zustände -------------------------------------------------- */
.dex-ui-empty { text-align: center; padding: 28px 16px; color: ${G500}; font-size: 0.86rem; border: 1.5px dashed ${G200}; border-radius: 14px; background: ${SOFT}; }
.dex-ui-empty-icon { display: inline-flex; width: 44px; height: 44px; border-radius: 50%; background: ${G100}; color: ${G500}; align-items: center; justify-content: center; margin-bottom: 10px; }
.dex-ui-empty-title { font-weight: 700; color: ${G700}; font-size: 0.92rem; margin-bottom: 4px; }

/* ---- Nachzug aus der v31.2-Runde (Wünsche der Umbau-Agenten) ---------- */
.dex-ui-choice-title, .dex-ui-choice-desc { display: block; }
.dex-ui-row.is-active { background: rgba(134,188,37,0.08); box-shadow: inset 3px 0 0 ${G}; }
.dex-ui-row.is-done { opacity: 0.6; }
.dex-ui-row.is-done:hover { opacity: 1; }
.dex-ui-rowbtn { width: 100%; text-align: left; background: transparent; border: none; padding: 0; margin: 0; font: inherit; color: inherit; cursor: pointer; }
.dex-ui-btn-reset { background: none; border: none; padding: 0; margin: 0; font: inherit; color: inherit; cursor: pointer; }
.dex-ui-card--list { padding: 6px; }
.dex-ui-select--sm { padding: 6px 28px 6px 10px; font-size: 0.84rem; border-radius: 8px; background-position: right 10px center; }
.dex-ui-input--error, .dex-ui-select--error, .dex-ui-textarea--error { border-color: var(--dex-red, #da291c) !important; }
.dex-ui-input--error:focus, .dex-ui-select--error:focus, .dex-ui-textarea--error:focus { box-shadow: 0 0 0 3px rgba(218,41,28,0.15) !important; }
.dex-ui-label-required { color: var(--dex-red, #da291c); margin-left: 2px; }
.dex-ui-range { width: 100%; accent-color: #86bc25; cursor: pointer; }
.dex-ui-dropzone { display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 6px; text-align: center; padding: 22px 16px; border: 1.5px dashed ${G300}; border-radius: 14px; background: ${SOFT}; color: ${G600}; font-size: 0.86rem; cursor: pointer; transition: border-color ${EASE}, background ${EASE}; }
.dex-ui-dropzone:hover, .dex-ui-dropzone.is-over { border-color: ${G}; background: rgba(134,188,37,0.06); color: ${GDT}; }
.dex-ui-callout--sm { padding: 6px 10px; font-size: 0.76rem; border-radius: 8px; }
.dex-ui-callout--flush { border-radius: 0; border-width: 0 0 1px 0; }
.dex-ui-tab.is-off { opacity: 0.55; }
.dex-ui-iconbtn.is-active { background: rgba(134,188,37,0.14); color: ${GDT}; }
.dex-ui-grid-2 > .dex-ui-field, .dex-ui-grid-3 > .dex-ui-field, .dex-ui-grid-auto > .dex-ui-field { margin-bottom: 0; }
.dex-ui-grid-3-1 { display: grid; grid-template-columns: 3fr 1fr; gap: 10px; }
.dex-ui-grid-1-3 { display: grid; grid-template-columns: 1fr 3fr; gap: 10px; }
@media (max-width: 768px) { .dex-ui-grid-3-1, .dex-ui-grid-1-3 { grid-template-columns: 1fr; } }
.dex-ui-disclosure-chevron.is-open { transform: rotate(180deg); }
.dex-ui-chip-remove { display: inline-flex; align-items: center; justify-content: center; width: 16px; height: 16px; border-radius: 50%; margin-left: 2px; margin-right: -4px; background: rgba(0,0,0,0.08); color: inherit; font-size: 0.7rem; line-height: 1; transition: background ${EASE}; }
.dex-ui-chip:hover .dex-ui-chip-remove { background: rgba(0,0,0,0.16); }
.dex-ui-chip.is-active .dex-ui-chip-remove { background: rgba(255,255,255,0.25); }

/* ---- Barrierefreiheit ---------------------------------------------- */
.dex-ui-sr-only { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0,0,0,0); white-space: nowrap; border: 0; }
`;

/**
 * Stellt sicher, dass das Stylesheet genau einmal im Dokument steht.
 * Idempotent und ohne Wirkung auf dem Server (kein `document`).
 */
export function ensureDexUiStyles(): void {
  if (typeof document === 'undefined') return;
  if (document.getElementById(DEX_UI_STYLE_ID)) return;
  const el = document.createElement('style');
  el.id = DEX_UI_STYLE_ID;
  el.textContent = DEX_UI_CSS;
  document.head.appendChild(el);
}

/**
 * Klassen zusammensetzen — `cx('dex-ui-chip', active && 'is-active')`.
 * Bewusst winzig; eine Abhängigkeit wie `classnames` wäre für sieben Zeilen
 * zu viel Bundle.
 */
export function cx(...parts: Array<string | false | null | undefined>): string {
  return parts.filter(Boolean).join(' ');
}
