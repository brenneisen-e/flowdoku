/* v32.0.6: Countdown vor dem Anlegen eines Events.
 *
 * Nutzer-Ansage 28.09.2026: „diesen Button will ich die ersten 3–5 Sekunden
 * abbrechen können — bei Event-Erstellung, nicht beim Ändern." Das Anlegen
 * selbst ist nicht abbrechbar (Subsite, Liste und Rechte entstehen in
 * Schritten; ein Abbruch in der Mitte hinterließe Reste). Deshalb wartet die
 * App VOR dem ersten Schreibvorgang ein paar Sekunden — bis dahin ist noch
 * nichts passiert, und „Abbrechen" ist sauber. */
import * as React from 'react';
import Modal from '../Modal';
import { Send } from '../Icons';

export const ANLEGE_SEKUNDEN = 4;

export function AnlegeCountdown(p: {
  sekunden: number | null;
  isDe: boolean;
  onTick: () => void;
  onAbbrechen: () => void;
  onJetzt: () => void;
}): React.ReactElement | null {
  const { sekunden, isDe, onTick, onAbbrechen, onJetzt } = p;
  // Ein Zeitgeber je Sekunde; bei 0 genau einmal starten.
  const gestartetRef = React.useRef(false);
  React.useEffect(() => {
    if (sekunden === null) { gestartetRef.current = false; return; }
    if (sekunden <= 0) {
      if (!gestartetRef.current) { gestartetRef.current = true; onJetzt(); }
      return;
    }
    const t = window.setTimeout(onTick, 1000);
    return () => window.clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sekunden]);
  if (sekunden === null || sekunden <= 0) return null;
  const anteil = Math.round(((ANLEGE_SEKUNDEN - sekunden) / ANLEGE_SEKUNDEN) * 100);
  return (
    <Modal
      open
      onClose={onAbbrechen}
      maxWidth={440}
      icon={<Send size={18} />}
      title={isDe ? `Event wird in ${sekunden} Sekunde${sekunden === 1 ? '' : 'n'} angelegt` : `Creating the event in ${sekunden} second${sekunden === 1 ? '' : 's'}`}
      subtitle={isDe ? 'Bis dahin ist noch nichts gespeichert — du kannst abbrechen.' : 'Nothing is saved yet — you can still cancel.'}
      footer={<>
        <button type="button" className="btn btn-secondary" onClick={onAbbrechen}>{isDe ? 'Abbrechen' : 'Cancel'}</button>
        <button type="button" className="btn btn-primary" onClick={onJetzt}>{isDe ? 'Jetzt anlegen' : 'Create now'}</button>
      </>}
    >
      <div className="dex-ui-progress" aria-hidden="true">
        <div className="dex-ui-progress-bar" style={{ width: `${anteil}%`, transition: 'width 1s linear' }} />
      </div>
    </Modal>
  );
}
