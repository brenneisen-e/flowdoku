/**
 * Meldungen und Rueckfragen an EINER Stelle.
 *
 * Wie in DEX: Kein `window.alert`, kein `window.confirm`. Beide sehen in
 * SharePoint fremd aus, lassen sich nicht gestalten und blockieren den
 * Renderer. Stattdessen ein Dialog im Haus-Design ueber `Modal`.
 *
 * Die Rueckfrage nennt die FOLGE, nicht nur die Frage (UI-Leitfaden 2b):
 * „Use Case loeschen? Die Kachel verschwindet fuer alle." — nicht
 * „Sind Sie sicher?".
 *
 * Geschrieben als JSX, nicht ueber `React.createElement`: `Modal` verlangt
 * `children` als Pflicht-Prop, und beim Aufruf per `createElement` kennt die
 * Typpruefung die Kinder aus dem dritten Parameter nicht — sie ins
 * Props-Objekt zu schreiben ist der Ausweg, den die ESLint-Regel
 * `react/no-children-prop` zu Recht verbietet. In JSX stellt sich die Frage
 * nicht.
 */

import * as React from 'react';
import Modal from '../components/Modal';
import { useLocaleSafe } from './LanguageContext';

export interface AlertOptions {
  variant?: 'info' | 'success' | 'error';
  title?: string;
}

export interface ConfirmOptions {
  confirmLabel?: string;
  cancelLabel?: string;
  title?: string;
  /** true = der Bestaetigen-Knopf ist die gefaehrliche Aktion. */
  danger?: boolean;
}

interface DialogContextType {
  showAlert: (message: React.ReactNode, opts?: AlertOptions) => void;
  confirmDialog: (message: React.ReactNode, opts?: ConfirmOptions) => Promise<boolean>;
}

const DialogContext = React.createContext<DialogContextType | undefined>(undefined);

interface AlertState { message: React.ReactNode; opts: AlertOptions }
interface ConfirmState { message: React.ReactNode; opts: ConfirmOptions; resolve: (v: boolean) => void }

export function DialogProvider(props: { children: React.ReactNode }): React.ReactElement {
  const [alertState, setAlertState] = React.useState<AlertState | null>(null);
  const [confirmState, setConfirmState] = React.useState<ConfirmState | null>(null);
  const locale = useLocaleSafe();
  const isDe = locale === 'de';

  const showAlert = React.useCallback((message: React.ReactNode, opts?: AlertOptions): void => {
    setAlertState({ message, opts: opts || {} });
  }, []);

  const confirmDialog = React.useCallback((message: React.ReactNode, opts?: ConfirmOptions): Promise<boolean> => {
    return new Promise<boolean>(resolve => {
      setConfirmState({ message, opts: opts || {}, resolve });
    });
  }, []);

  const value = React.useMemo<DialogContextType>(() => ({ showAlert, confirmDialog }), [showAlert, confirmDialog]);

  const closeConfirm = (answer: boolean): void => {
    if (confirmState) confirmState.resolve(answer);
    setConfirmState(null);
  };

  const alertTitel = alertState
    ? (alertState.opts.title || (
      alertState.opts.variant === 'error' ? (isDe ? 'Das hat nicht geklappt' : 'That did not work')
        : alertState.opts.variant === 'success' ? (isDe ? 'Erledigt' : 'Done')
          : (isDe ? 'Hinweis' : 'Notice')
    ))
    : '';

  return (
    <DialogContext.Provider value={value}>
      {props.children}

      {alertState && (
        <Modal
          open
          onClose={() => setAlertState(null)}
          maxWidth={460}
          ariaLabel={alertTitel}
          title={alertTitel}
          footer={
            <button type="button" className="btn btn-primary" onClick={() => setAlertState(null)}>
              {isDe ? 'Alles klar' : 'Got it'}
            </button>
          }
        >
          <div className="dex-ui-callout-body">{alertState.message}</div>
        </Modal>
      )}

      {confirmState && (
        <Modal
          open
          onClose={() => closeConfirm(false)}
          maxWidth={460}
          ariaLabel={confirmState.opts.title || (isDe ? 'Bitte bestätigen' : 'Please confirm')}
          title={confirmState.opts.title || (isDe ? 'Bitte bestätigen' : 'Please confirm')}
          footer={<>
            <button type="button" className="btn btn-secondary" onClick={() => closeConfirm(false)}>
              {confirmState.opts.cancelLabel || (isDe ? 'Abbrechen' : 'Cancel')}
            </button>
            {/* Der gefaehrliche Knopf bleibt grau (DEX v24.63) — die Warnung
                traegt der Text, nicht die Farbe. */}
            <button
              type="button"
              className={confirmState.opts.danger ? 'btn btn-secondary' : 'btn btn-primary'}
              onClick={() => closeConfirm(true)}
            >
              {confirmState.opts.confirmLabel || (isDe ? 'Ja, weiter' : 'Yes, continue')}
            </button>
          </>}
        >
          <div className="dex-ui-callout-body">{confirmState.message}</div>
        </Modal>
      )}
    </DialogContext.Provider>
  );
}

export function useDialog(): DialogContextType {
  const ctx = React.useContext(DialogContext);
  if (!ctx) throw new Error('useDialog muss innerhalb des DialogProvider stehen');
  return ctx;
}
