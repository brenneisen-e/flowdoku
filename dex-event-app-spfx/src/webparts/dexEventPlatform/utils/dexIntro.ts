/**
 * v32.17: Ist das ein Einführungs-Event zu DEX selbst?
 *
 * Nutzer-Ansage 29.09.2026: Beim Event, in dem DEX vorgestellt wird, soll auf
 * der Anmeldeseite und unter „Meine Events" das animierte Logo der Landing
 * Page stehen — und nur dort. Bewusst kein Schalter im Assistenten: Ein
 * weiteres Piggyback-Flag müsste beim Laden gestrippt und mitgetragen werden,
 * für eine Handvoll Events im Jahr. Erkannt wird am Titel: „DEX" als eigenes
 * Wort plus ein Einführungs-Begriff. Ein Termin (Sub-Event) erbt die
 * Einordnung vom Hauptevent, weil sein Titel oft nur „Session 1" lautet.
 */
const EINFUEHRUNG = /(einf(ü|ue)hrung|vorstellung|intro|kick-?off|onboarding|launch|walk-?through)/;

const titelPasst = (titel?: string): boolean => {
  const t = (titel || '').toLowerCase();
  return /(^|[^a-z0-9_])dex([^a-z0-9_]|$)/.test(t) && EINFUEHRUNG.test(t);
};

export const istDexEinfuehrung = (
  ev?: { title?: string } | null,
  hauptevent?: { title?: string } | null,
): boolean => titelPasst(ev?.title) || titelPasst(hauptevent?.title);
