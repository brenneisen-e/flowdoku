/**
 * Die zwei Hilfe-Dialoge — „Über die App" und „Fragen / Organizer werden" —
 * EINMAL gehalten, von vier Stellen geöffnet.
 *
 * Geöffnet werden sie aus der Kopfzeile („Hast du Fragen?", Burger-Menü), von
 * der Landing Page („Organizer werden") und von der Start-Übersicht (Kachel
 * „Fragen & Feedback"). Hätte jede Stelle ihren eigenen Dialog, stünde derselbe
 * Text viermal im Bundle und liefe bei der ersten Textänderung auseinander.
 */

import * as React from 'react';
import ContactModal, { KontaktArt } from '../components/ContactModal';
import AboutModal from '../components/AboutModal';

interface HilfeContextType {
  openKontakt: (art?: KontaktArt) => void;
  openAbout: () => void;
}

const HilfeContext = React.createContext<HilfeContextType | undefined>(undefined);

export function HilfeProvider(props: { children: React.ReactNode }): React.ReactElement {
  const [kontakt, setKontakt] = React.useState<KontaktArt | null>(null);
  const [about, setAbout] = React.useState(false);

  const value = React.useMemo<HilfeContextType>(() => ({
    openKontakt: (art?: KontaktArt) => setKontakt(art || 'frage'),
    openAbout: () => setAbout(true),
  }), []);

  return (
    <HilfeContext.Provider value={value}>
      {props.children}
      <ContactModal open={kontakt !== null} art={kontakt || 'frage'} onClose={() => setKontakt(null)} />
      <AboutModal open={about} onClose={() => setAbout(false)} onAsk={() => setKontakt('frage')} />
    </HilfeContext.Provider>
  );
}

export function useHilfe(): HilfeContextType {
  const ctx = React.useContext(HilfeContext);
  if (!ctx) throw new Error('useHilfe muss innerhalb des HilfeProvider stehen');
  return ctx;
}
