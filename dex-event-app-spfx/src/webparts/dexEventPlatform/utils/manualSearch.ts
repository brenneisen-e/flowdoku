/**
 * v26.0.0 — Leichtgewichtige Handbuch-Suche fürs Ticketsystem.
 *
 * Lädt die Handbuch-Sektionen LAZY (`await import` von handbookContent), damit
 * der schwere Handbuch-Code NICHT ins Boot-Bundle wandert. Genutzt von:
 *  - dem „Hast du Fragen?"-Modal (Live-Vorschläge zur eingegebenen Frage =
 *    Selbsthilfe, bevor ein Ticket rausgeht),
 *  - dem Antwort-Composer (Artikel-Picker für Power-User/Organizer).
 */

export interface ManualArticle {
  id: string;
  title: string;
  description: string;
}

interface IndexedArticle extends ManualArticle {
  hay: string;
  /** v32.0.12: Der Artikeltext selbst. Bis v32.0.11 durchsuchte das
   *  Fragen-Fenster nur Titel, Beschreibung und Stichwörter — „Reminder" fand
   *  nichts, obwohl „E-Mail an Teilnehmer" den Reminder ausdrücklich erklärt
   *  (Nutzer-Befund 28.09.2026). Treffer hier zählen schwächer. */
  body: string;
}

/** Alle Texte aus einer Handbuch-Sektion einsammeln — auch aus React-Knoten
 *  (Überschriften, Absätze, Listen). Funktionen und Zyklen werden übergangen. */
function textOf(node: unknown, seen: WeakSet<object>, depth: number, out: string[]): void {
  if (node === null || node === undefined || typeof node === 'boolean' || depth > 40) return;
  if (typeof node === 'string') { out.push(node); return; }
  if (typeof node === 'number') return;
  if (typeof node !== 'object') return;
  if (seen.has(node as object)) return;
  seen.add(node as object);
  if (Array.isArray(node)) { for (const n of node) textOf(n, seen, depth + 1, out); return; }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const o = node as any;
  if (o.$$typeof) { if (o.props) textOf(o.props, seen, depth + 1, out); return; }
  for (const k of Object.keys(o)) {
    if (k === 'id' || k === 'icon' || k === 'style' || k === 'className' || k === 'key' || k === 'visibleFor') continue;
    textOf(o[k], seen, depth + 1, out);
  }
}

let _cache: { locale: string; items: IndexedArticle[] } | null = null;

function norm(s: string): string {
  return (s || '')
    .toLowerCase()
    .replace(/ä/g, 'ae').replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ß/g, 'ss')
    .replace(/[-_/]/g, ' ')
    .replace(/[^a-z0-9 ]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

async function loadIndex(locale: 'de' | 'en'): Promise<IndexedArticle[]> {
  if (_cache && _cache.locale === locale) return _cache.items;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const mod: any = await import('../components/manual/handbookContent');
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const sections: any[] = mod.getManualSections(locale) || [];
  const items: IndexedArticle[] = sections.map((s) => ({
    id: s.id,
    title: s.title || s.id,
    description: s.description || '',
    hay: norm(`${s.title || ''} ${s.description || ''} ${s.keywords || ''}`),
    body: (() => { const out: string[] = []; try { textOf(s.perspectives, new WeakSet(), 0, out); } catch { /* ohne Text */ } return norm(out.join(' ')); })(),
  }));
  _cache = { locale, items };
  return items;
}

/** Live-Suche zur eingegebenen Frage — beste Treffer zuerst. */
export async function searchManual(query: string, locale: 'de' | 'en', limit = 5): Promise<ManualArticle[]> {
  const items = await loadIndex(locale);
  const words = norm(query).split(' ').filter((w) => w.length >= 3);
  if (words.length === 0) return [];
  const scored = items
    .map((it) => {
      let score = 0;
      for (const w of words) {
        if (it.hay.indexOf(w) >= 0) score += 2;
        if (norm(it.title).indexOf(w) >= 0) score += 3; // Treffer im Titel höher gewichten
        else if (it.hay.indexOf(w) < 0 && it.body.indexOf(w) >= 0) score += 1; // nur im Text: schwächer
      }
      return { it, score };
    })
    .filter((x) => x.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, limit);
  return scored.map((x) => ({ id: x.it.id, title: x.it.title, description: x.it.description }));
}

/** Alle Handbuch-Artikel (für den Picker im Antwort-Composer). */
export async function listManualArticles(locale: 'de' | 'en'): Promise<ManualArticle[]> {
  const items = await loadIndex(locale);
  return items.map((it) => ({ id: it.id, title: it.title, description: it.description }));
}

/** Vollständige Handbuch-Sektion (mit Inhalten) lazy laden — für die
 *  Inline-Vorschau im „Hast du Fragen?"-Modal (Antwort ausklappen, ohne ins
 *  Handbuch zu springen). Rückgabe ist eine ManualSection (enthält React-Nodes). */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export async function getManualSection(id: string, locale: 'de' | 'en'): Promise<any | null> {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mod: any = await import('../components/manual/handbookContent');
    const sections = mod.getManualSections(locale) || [];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return sections.find((s: any) => s.id === id) || null;
  } catch { return null; }
}

/** Einen Handbuch-Artikel öffnen (gleicher Mechanismus wie die globale Suche). */
export function openManualArticle(id: string, navigate: (page: 'manual') => void): void {
  try { window.localStorage.setItem('dex_open_manual_section', id); } catch { /* */ }
  try { window.dispatchEvent(new CustomEvent('dex-open-manual-section', { detail: id })); } catch { /* */ }
  navigate('manual');
}
