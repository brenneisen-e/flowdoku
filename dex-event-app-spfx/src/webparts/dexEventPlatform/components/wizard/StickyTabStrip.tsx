/**
 * v28.94: Aus `EventCreationPage` herausgeloest (dort 380 der 18.198 Zeilen).
 * Reiter-Leiste fuer die Ebene, auf der gerade gearbeitet wird — Klammer/
 * Haupt-Event plus die Sub-Events. Die Komponente kennt nichts vom Wizard-
 * State: Sie bekommt Tabs, aktiven Index und einen onChange.
 */
import * as React from 'react';

/**
 * v22.30: Tab-Leiste der pro-Sub-Event-Schritte (Klammer/Haupt + Sub-Events).
 * - Sitzt ganz OBEN in der weißen Schritt-Karte (über der Überschrift).
 * - Bleibt beim Scrollen sichtbar: JS-Pin unter dem (ebenfalls gepinnten)
 *   Header — CSS-sticky wird im SP-Canvas durch Overflow-Vorfahren
 *   ausgehebelt (gleiche Falle wie beim Header, siehe Header.tsx v22.28).
 *   Ein Platzhalter hält die Höhe, die Leiste wechselt auf position:fixed.
 * - Der AKTIVE Tab ist gefüllt grün mit weißer Schrift.
 * Versteckte Schritte (display:none) pinnen nicht (offsetParent-Check).
 */
export function StickyTabStrip(props: {
  tabs: Array<{ label: string; isMain: boolean }>;
  activeIdx: number;
  onChange: (idx: number) => void;
  ariaLabel: string;
  mainBadge: string;
  /** v22.71: Klammer-Modus — der Haupt-Tab wird als echte Klammer ÜBER den
   *  Sub-Event-Tabs dargestellt (oben volle Breite, darunter eingerückt). */
  klammer?: boolean;
  /** v22.71: Haupt-/Klammer-Tab nicht anklickbar (z.B. Schritt 6: im
   *  „Nur Sub-Events"-Modus ist die Klammer-Kommunikation nicht relevant). */
  mainDisabled?: boolean;
  /** Optionaler Hinweis-Text neben dem deaktivierten Klammer-Tab. */
  mainDisabledNote?: string;
  /** v28.73: Wort in Klammern hinter dem Eventnamen, z.B. „Klammerevent". */
  klammerWord?: string;
  /** v28.73: Info-Icon (Tooltip) rechts im Klammer-Reiter — erklärt, dass die
   *  Anmeldung über die Sub-Events läuft, mit Quicklink zur Umstellung. */
  klammerInfo?: React.ReactNode;
}): React.ReactElement {
  const phRef = React.useRef<HTMLDivElement | null>(null);
  const [pin, setPin] = React.useState<null | { top: number; left: number; width: number; height: number }>(null);
  // v28.72: Hover-/Fokus-Index für die Reiter (Inline-Styles können kein :hover).
  const [hoverIdx, setHoverIdx] = React.useState<number | null>(null);

  /**
   * v28.89: Die Reiter-Reihe bei vielen Sub-Events bedienbar machen.
   *
   * Seit v28.85 bricht die Reihe nicht mehr um, sondern scrollt — damit war
   * das Umbruch-Problem weg, aber ein neues da: Bei neun Terminen liegt die
   * Hälfte außerhalb des Sichtfelds, ohne dass man es sieht. Wer per Schritt
   * oder Validierung auf einen Reiter geschickt wird, findet ihn nicht wieder;
   * und ein waagerechter Scrollbalken ist auf dem Trackpad ein Zufallsfund.
   *
   * Drei Ergänzungen, bewusst am etablierten Tab-Muster orientiert (Pfeile
   * links/rechts, aktiver Reiter wird sichtbar gehalten, Pfeiltasten):
   *  - `ovf` sagt, ob links/rechts noch etwas liegt → nur dann ein Pfeil,
   *  - der aktive Reiter wird nach jedem Wechsel in den sichtbaren Bereich
   *    gescrollt (auch wenn der Wechsel von außen kommt, z.B. aus setScope),
   *  - Pfeil links/rechts blättert von Reiter zu Reiter (Rolle „tablist"
   *    verspricht das ohnehin — bisher tat es nichts).
   */
  const scrollRef = React.useRef<HTMLDivElement | null>(null);
  const [ovf, setOvf] = React.useState<{ left: boolean; right: boolean }>({ left: false, right: false });
  const updateOvf = React.useCallback((): void => {
    const el = scrollRef.current;
    if (!el) return;
    const max = el.scrollWidth - el.clientWidth;
    setOvf(prev => {
      const next = { left: el.scrollLeft > 2, right: el.scrollLeft < max - 2 };
      return (prev.left === next.left && prev.right === next.right) ? prev : next;
    });
  }, []);
  React.useEffect(() => {
    updateOvf();
    const el = scrollRef.current;
    if (!el) return undefined;
    el.addEventListener('scroll', updateOvf, { passive: true });
    window.addEventListener('resize', updateOvf);
    return () => {
      el.removeEventListener('scroll', updateOvf);
      window.removeEventListener('resize', updateOvf);
    };
  }, [updateOvf, props.tabs.length]);
  React.useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const btn = el.querySelector(`[data-tabidx="${props.activeIdx}"]`) as HTMLElement | null;
    if (btn) {
      const left = btn.offsetLeft;
      const right = left + btn.offsetWidth;
      if (left < el.scrollLeft) el.scrollLeft = Math.max(0, left - 24);
      else if (right > el.scrollLeft + el.clientWidth) el.scrollLeft = right - el.clientWidth + 24;
    }
    updateOvf();
  }, [props.activeIdx, props.tabs.length, updateOvf]);
  const nudge = (dir: -1 | 1): void => {
    const el = scrollRef.current;
    if (!el) return;
    el.scrollLeft += dir * Math.max(160, Math.round(el.clientWidth * 0.7));
  };
  /** Pfeiltasten blättern durch die Sub-Event-Reiter (Index 1..n). */
  const onTabKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
    const last = props.tabs.length - 1;
    const cur = props.activeIdx;
    const next = e.key === 'ArrowLeft' ? cur - 1 : cur + 1;
    if (next < 0 || next > last) return;
    if (next === 0 && props.mainDisabled) return;
    e.preventDefault();
    props.onChange(next);
  };
  /** Kleiner runder Blätter-Knopf — nur sichtbar, wenn es dort etwas gibt. */
  const arrowBtn = (dir: -1 | 1): React.ReactElement => (
    <button
      type="button"
      onClick={() => nudge(dir)}
      aria-label={dir === -1
        ? (props.ariaLabel ? 'Reiter nach links' : 'Scroll tabs left')
        : (props.ariaLabel ? 'Reiter nach rechts' : 'Scroll tabs right')}
      tabIndex={-1}
      style={{
        flexShrink: 0, width: 26, height: 26, borderRadius: '50%',
        border: '1px solid var(--dex-gray-300)', background: '#fff',
        color: 'var(--dex-green-dark, #4a7c1f)', cursor: 'pointer',
        display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
        fontSize: '0.9rem', lineHeight: 1, padding: 0, alignSelf: 'center',
        boxShadow: '0 1px 3px rgba(0,0,0,0.10)',
      }}
    >
      {dir === -1 ? '‹' : '›'}
    </button>
  );
  React.useEffect(() => {
    const update = (): void => {
      // v28.82: Fixieren beim Scrollen abgeschaltet. Seit v28.78 sitzt der
      // Umschalter in einer getoenten Scope-Karte, deren Beschriftung und
      // Erklaerung AUSSERHALB dieser Komponente liegen — beim Fixieren riss
      // der Reiter-Block als eigener Kasten heraus, die Karte blieb zurueck
      // und der gruene Schritt-Kopf wurde angeschnitten. Die Karte steht
      // ohnehin ganz oben, direkt unter der Schritt-Leiste, und nach jedem
      // Schrittwechsel ist man wieder dort. Der Nutzen des Mitscrollens wiegt
      // die zerrissene Darstellung nicht auf.
      setPin(null);
      return;
      // eslint-disable-next-line no-unreachable
      const ph = phRef.current;
      if (!ph || ph.offsetParent === null) { setPin(null); return; }
      const headerEl = document.querySelector('.header');
      const topEdge = headerEl ? Math.max(0, headerEl.getBoundingClientRect().bottom) : 0;
      const r = ph.getBoundingClientRect();
      if (r.top < topEdge) {
        const height = ph.offsetHeight || 48;
        setPin(prev => {
          const next = { top: topEdge, left: r.left, width: r.width, height: prev ? prev.height : height };
          if (prev && prev.top === next.top && prev.left === next.left && prev.width === next.width) return prev;
          return next;
        });
      } else {
        setPin(null);
      }
    };
    update();
    window.addEventListener('scroll', update, true);
    window.addEventListener('resize', update);
    return () => {
      window.removeEventListener('scroll', update, true);
      window.removeEventListener('resize', update);
    };
  }, []);
  return (
    <div ref={phRef} style={pin ? { height: pin.height } : undefined}>
      <div style={{
        display: 'flex', alignItems: 'center', gap: 8, marginBottom: 16, minWidth: 0,
        // v28.82: Im fixierten Zustand loeste sich der Reiter-Block bisher als
        // WEISSER Kasten aus der getoenten Scope-Karte — die Karte blieb
        // zurueck, der Block schwebte darueber und schnitt den gruenen
        // Schritt-Kopf an. Jetzt traegt der fixierte Zustand dieselbe Flaeche
        // wie die Karte (deckend, damit nichts durchscheint), sitzt buendig
        // ueber die volle Breite und schliesst mit einem klaren Rand ab.
        ...(pin ? {
          position: 'fixed', top: pin.top, left: pin.left, width: pin.width,
          zIndex: 800, background: '#f2f7e8', boxSizing: 'border-box',
          padding: '10px 16px 10px', marginBottom: 0,
          borderRadius: '0 0 12px 12px',
          boxShadow: '0 8px 18px rgba(0,0,0,0.10)',
          borderBottom: '1px solid rgba(134,188,37,0.45)',
        } : {}),
      }}>
        {(() => {
          const renderTabBtn = (tab: { label: string; isMain: boolean }, tabIdx: number): React.ReactElement => {
            const active = tabIdx === props.activeIdx;
            // v28.72: Hover-Effekt. Die Reiter waren statische graue Chips —
            // ohne Reaktion auf die Maus las sich das wie eine Beschriftung
            // statt wie etwas Anklickbares. Jetzt heben sie sich beim
            // Darüberfahren an (grüner Rand + Tönung + 1px nach oben).
            const hovered = hoverIdx === tabIdx && !active;
            return (
              <button
                key={tabIdx}
                type="button"
                role="tab"
                data-tabidx={tabIdx}
                aria-selected={active}
                onClick={() => props.onChange(tabIdx)}
                onMouseEnter={() => setHoverIdx(tabIdx)}
                onMouseLeave={() => setHoverIdx(prev => (prev === tabIdx ? null : prev))}
                onFocus={() => setHoverIdx(tabIdx)}
                onBlur={() => setHoverIdx(prev => (prev === tabIdx ? null : prev))}
                style={{
                  display: 'inline-flex', alignItems: 'center', gap: 8,
                  padding: '8px 14px',
                  border: active
                    ? '1px solid var(--dex-green, #86bc25)'
                    : `1px solid ${hovered ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-200)'}`,
                  borderRadius: '8px 8px 0 0',
                  background: active
                    ? 'var(--dex-green, #86bc25)'
                    : (hovered ? 'rgba(134,188,37,0.14)' : 'var(--dex-gray-50, #fafafa)'),
                  color: active ? '#fff' : (hovered ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-700)'),
                  fontWeight: active ? 700 : (hovered ? 600 : 500),
                  fontSize: '0.85rem',
                  cursor: 'pointer',
                  marginBottom: -1,
                  whiteSpace: 'nowrap',
                  maxWidth: 280,
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  // v29.17: NIE mitschrumpfen. In der (frueheren) nowrap-Zeile
                  // setzte overflow:hidden die min-content-Breite auf ~0 —
                  // Flexbox quetschte bei 20+ Terminen jeden Reiter auf EINEN
                  // Buchstaben („D M D. F.“), statt die Zeile scrollen zu
                  // lassen. Ein Reiter ist so breit wie sein Text, Punkt.
                  flexShrink: 0,
                  transform: hovered ? 'translateY(-1px)' : 'none',
                  boxShadow: hovered ? '0 -2px 6px rgba(134,188,37,0.20)' : 'none',
                  transition: 'background 0.15s, color 0.15s, border-color 0.15s, transform 0.15s, box-shadow 0.15s',
                }}
                title={tab.label}
              >
                {tab.isMain && (
                  <span style={{ fontSize: '0.7rem', textTransform: 'uppercase', letterSpacing: 0.4, color: active ? 'rgba(255,255,255,0.85)' : 'var(--dex-gray-400)' }}>
                    {props.mainBadge}
                  </span>
                )}
                <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>{tab.label}</span>
              </button>
            );
          };
          const klammerLayout = !!props.klammer && props.tabs.length > 1 && !!props.tabs[0] && props.tabs[0].isMain;
          if (klammerLayout) {
            const dis = !!props.mainDisabled;
            const pActive = !dis && props.activeIdx === 0;
            return (
              <div role="tablist" aria-label={props.ariaLabel} style={{ display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0, gap: 0 }}>
                {/* Klammer-Ebene oben — volle Breite.
                    v28.73: Vorne steht jetzt der EVENTNAME, dahinter der
                    dezente Zusatz „(Klammerevent)" und ein Info-Icon mit der
                    Erklärung. Vorher stand „⟦ KLAMMER ⟧" als Praefix davor —
                    ein Fachbegriff an der auffälligsten Stelle, der weder
                    sagt, was er bedeutet, noch dass er eine Entscheidung war.
                    Das Icon liegt bewusst NEBEN dem Reiter-Button (nicht
                    darin), damit der Quicklink im Tooltip nicht in einem
                    Button verschachtelt ist. */}
                {/* v28.74: Rahmen/Fläche liegen auf der ZEILE, nicht auf dem
                    Button — dadurch kann der Button auf Textbreite schrumpfen
                    und das Info-Icon sitzt direkt hinter „(Klammerevent)"
                    statt am rechten Rand, ohne dass die grüne Leiste ihre
                    volle Breite verliert. */}
                {/* v28.75: Die ganze Zeile ist der Reiter — nicht nur der Text.
                    Vorher war der Bereich rechts vom Info-Icon tot: Der Button
                    schrumpfte auf Textbreite, der Rest der grünen Leiste nahm
                    keine Klicks an. Jetzt trägt die Zeile selbst role/onClick
                    (inkl. Tastatur) und den Hover — ein <button> geht hier
                    nicht, weil das Info-Icon mit seinem Link darin läge. */}
                <div
                  role="tab"
                  aria-selected={pActive}
                  aria-disabled={dis}
                  tabIndex={dis ? -1 : 0}
                  onClick={() => { if (!dis) props.onChange(0); }}
                  onKeyDown={e => {
                    if (dis) return;
                    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); props.onChange(0); }
                  }}
                  onMouseEnter={() => setHoverIdx(0)}
                  onMouseLeave={() => setHoverIdx(prev => (prev === 0 ? null : prev))}
                  onFocus={() => setHoverIdx(0)}
                  onBlur={() => setHoverIdx(prev => (prev === 0 ? null : prev))}
                  title={dis ? (props.mainDisabledNote || props.tabs[0].label) : props.tabs[0].label}
                  style={{
                    display: 'flex', alignItems: 'center', gap: 8, width: '100%', minWidth: 0,
                    padding: '10px 12px 10px 16px', cursor: dis ? 'not-allowed' : 'pointer',
                    border: `1.5px solid ${(pActive || (hoverIdx === 0 && !dis)) ? 'var(--dex-green, #86bc25)' : 'var(--dex-gray-300)'}`,
                    borderRadius: '10px 10px 0 0',
                    // v28.86: Nicht ausgewaehlt = WEISS, wie die Sub-Event-Reiter
                    // auch. Die gruene Toenung im Ruhezustand liess die Klammer
                    // wie dauerhaft aktiv aussehen — neben einem wirklich
                    // aktiven Sub-Event-Reiter waren dann zwei Elemente gruen.
                    background: pActive
                      ? 'var(--dex-green, #86bc25)'
                      : ((hoverIdx === 0 && !dis) ? 'rgba(134,188,37,0.14)' : '#fff'),
                    color: pActive ? '#fff' : 'var(--dex-green-dark, #4a7c1f)',
                    fontWeight: 700, fontSize: '0.9rem',
                    opacity: dis ? 0.55 : 1,
                    boxShadow: (hoverIdx === 0 && !pActive && !dis) ? 'inset 0 0 0 1px rgba(134,188,37,0.35)' : 'none',
                    transition: 'background 0.15s, border-color 0.15s, box-shadow 0.15s',
                  }}
                >
                  <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', color: pActive ? '#fff' : 'var(--dex-green-dark, #4a7c1f)' }}>{props.tabs[0].label}</span>
                  <span style={{ fontSize: '0.76rem', fontWeight: 600, whiteSpace: 'nowrap', flexShrink: 0, opacity: 0.85, color: pActive ? 'rgba(255,255,255,0.9)' : 'var(--dex-green-dark, #4a7c1f)' }}>
                    ({props.klammerWord || 'Klammerevent'})
                  </span>
                  <span
                    style={{ display: 'inline-flex', alignItems: 'center', flexShrink: 0 }}
                    onClick={e => e.stopPropagation()}
                  >
                    {props.klammerInfo}
                  </span>
                  <span style={{ flex: 1 }} />
                  {dis && props.mainDisabledNote && (
                    <span style={{ fontSize: '0.65rem', fontWeight: 600, padding: '2px 6px', borderRadius: 8, background: 'var(--dex-gray-200, #e0e0e0)', color: 'var(--dex-gray-600)', flexShrink: 0 }}>
                      {props.mainDisabledNote}
                    </span>
                  )}
                </div>
                {/* Sub-Events darunter — eingerückt unter einer Klammer-Linie. */}
                {/* v29.17: Wieder UMBRECHEN statt scrollen — dieselbe
                    Darstellung wie die Reiter im Organizer Center (Event-
                    Details), wo genau dieses Layout mit vielen Terminen gut
                    lesbar ist. Die v28.85-Scroll-Zeile war doppelt gescheitert:
                    Ein Flexbox-Fehler quetschte die Reiter auf einen
                    Buchstaben statt zu scrollen (overflow:hidden am Button →
                    min-content ~0), und selbst korrekt gescrollt blieben von
                    21 Terminen nur wenige sichtbar. Beim Umbruch tragen die
                    Klammer-Linie links und die gemeinsame Grundlinie die
                    Zusammengehörigkeit. Die Pfeil-Knöpfe bleiben im Code und
                    erscheinen nur, wenn tatsächlich horizontal gescrollt wird
                    (ovf) — beim Umbruch also nie. */}
                <div style={{
                  display: 'flex', alignItems: 'stretch', gap: 6, minWidth: 0,
                  marginLeft: 18, paddingLeft: 16, paddingTop: 10, paddingRight: 2,
                  borderLeft: '2px solid var(--dex-green, #86bc25)',
                  borderBottom: pin ? 'none' : '1px solid var(--dex-gray-200)',
                }}>
                  {ovf.left && arrowBtn(-1)}
                  <div
                    ref={scrollRef}
                    onKeyDown={onTabKeyDown}
                    style={{
                      display: 'flex', flexWrap: 'wrap', gap: 6, alignItems: 'flex-end',
                      flex: 1, minWidth: 0,
                      paddingBottom: 2,
                    }}
                  >
                    {props.tabs.slice(1).map((tab, i) => renderTabBtn(tab, i + 1))}
                  </div>
                  {ovf.right && arrowBtn(1)}
                  {/* v28.89: Standortanzeige „3 / 21" — beim Umbruch weiter
                      nützlich als Orientierung, welcher Termin gerade offen ist. */}
                  {props.tabs.length - 1 >= 5 && (
                    <span style={{
                      alignSelf: 'center', flexShrink: 0, fontSize: '0.72rem', fontWeight: 700,
                      color: 'var(--dex-gray-500)', whiteSpace: 'nowrap', paddingLeft: 2,
                    }}>
                      {props.activeIdx > 0 ? `${props.activeIdx} / ${props.tabs.length - 1}` : `${props.tabs.length - 1}`}
                    </span>
                  )}
                </div>
              </div>
            );
          }
          return (
            <div
              role="tablist"
              aria-label={props.ariaLabel}
              style={{
                display: 'flex', flexWrap: 'wrap', gap: 6, flex: 1,
                borderBottom: pin ? 'none' : '1px solid var(--dex-gray-200)',
                paddingBottom: 0,
              }}
            >
              {props.tabs.map((tab, tabIdx) => renderTabBtn(tab, tabIdx))}
            </div>
          );
        })()}
      </div>
    </div>
  );
}
