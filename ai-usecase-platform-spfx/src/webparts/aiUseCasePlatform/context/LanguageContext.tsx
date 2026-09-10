/**
 * Sprache der Oberflaeche.
 *
 * Aufbau wie in DEX, aber schlank: DEX traegt ein Woerterbuch mit mehreren
 * hundert Schluesseln, weil dort Mailtexte, Formularfelder und Assistent
 * zweisprachig sind. Hier reicht der Rahmen — Kacheln, Detailseite und
 * Verwaltung. Wer einen Text ergaenzt, traegt beide Sprachen ein.
 *
 * `useLocaleSafe` ist bewusst der Zugriff der Wahl: Er funktioniert AUCH
 * ausserhalb des Providers. In DEX war die Verwechslung von
 * `navigator.language` und der App-Sprache ein wiederkehrender Fehler — bei
 * englischem Browser und deutscher App standen zwei Sprachen untereinander
 * (v31.9.5). Eine Quelle, ueberall dieselbe.
 */

import * as React from 'react';

export type Locale = 'de' | 'en';

interface LanguageContextType {
  locale: Locale;
  isDe: boolean;
  setLocale: (l: Locale) => void;
  /** Text nach Sprache. `t('Speichern', 'Save')` */
  t: (de: string, en: string) => string;
}

const LanguageContext = React.createContext<LanguageContextType | undefined>(undefined);

const STORAGE_KEY = 'aiuc_locale';

function initialLocale(): Locale {
  try {
    const saved = window.localStorage.getItem(STORAGE_KEY);
    if (saved === 'de' || saved === 'en') return saved;
  } catch { /* privates Fenster */ }
  // Erstbesuch: die Browsersprache entscheidet EINMAL. Danach zaehlt nur noch
  // die Wahl der Person.
  const nav = (typeof navigator !== 'undefined' && navigator.language) || 'de';
  return nav.toLowerCase().indexOf('de') === 0 ? 'de' : 'en';
}

export function LanguageProvider(props: { children: React.ReactNode }): React.ReactElement {
  const [locale, setLocaleState] = React.useState<Locale>(initialLocale);

  const setLocale = React.useCallback((l: Locale): void => {
    setLocaleState(l);
    try { window.localStorage.setItem(STORAGE_KEY, l); } catch { /* privates Fenster */ }
  }, []);

  const value = React.useMemo<LanguageContextType>(() => ({
    locale,
    isDe: locale === 'de',
    setLocale,
    t: (de: string, en: string) => (locale === 'de' ? de : en),
  }), [locale, setLocale]);

  return React.createElement(LanguageContext.Provider, { value }, props.children);
}

export function useLanguage(): LanguageContextType {
  const ctx = React.useContext(LanguageContext);
  if (!ctx) throw new Error('useLanguage muss innerhalb des LanguageProvider stehen');
  return ctx;
}

/**
 * Die Sprache, auch ausserhalb des Providers.
 *
 * `Modal` und andere uebernommene Bausteine rufen das; sie sollen nicht
 * abstuerzen, wenn sie einmal ohne Provider gerendert werden (Vorschau,
 * Test, Portal ausserhalb des Baums).
 */
export function useLocaleSafe(): Locale {
  const ctx = React.useContext(LanguageContext);
  return ctx ? ctx.locale : 'de';
}
