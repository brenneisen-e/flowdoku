/**
 * Der Suchtext — geteilt zwischen dem Suchfeld in der Kopfzeile und der
 * Kachelwand.
 *
 * Ein eigener Kontext und kein Feld im NavigationContext: Navigation wechselt
 * Seiten, die Suche filtert eine. Beides zu mischen hieße, dass jede Eingabe
 * einen Render aller Seitenwechsel-Konsumenten auslöst.
 */

import * as React from 'react';

interface SucheContextType {
  suche: string;
  setSuche: (s: string) => void;
}

const SucheContext = React.createContext<SucheContextType | undefined>(undefined);

export function SucheProvider(props: { children: React.ReactNode }): React.ReactElement {
  const [suche, setSuche] = React.useState('');
  const value = React.useMemo<SucheContextType>(() => ({ suche, setSuche }), [suche]);
  return React.createElement(SucheContext.Provider, { value }, props.children);
}

export function useSuche(): SucheContextType {
  const ctx = React.useContext(SucheContext);
  if (!ctx) throw new Error('useSuche muss innerhalb des SucheProvider stehen');
  return ctx;
}
