/**
 * v20.4 — Zentrale App-Dialoge: moderne Modal-Ersätze für die nativen
 * Browser-Boxen window.confirm() / window.alert() / window.prompt().
 *
 * Motivation: Die nativen Dialoge („deudeloitte.sharepoint.com says …")
 * sehen nach Browser-Chrome aus statt nach App, sind nicht stylebar und
 * zeigen die SharePoint-Domain als Absender. Dieser Provider stellt zwei
 * Funktionen bereit, die überall in der App die nativen Boxen ersetzen:
 *
 *   const { confirmDialog, showAlert } = useDialog();
 *   if (!(await confirmDialog('Event wirklich löschen?'))) return;
 *   showAlert('Gespeichert.', { variant: 'success' });
 *
 * - confirmDialog(message, opts?) → Promise<boolean> (OK = true).
 *   WICHTIG: anders als window.confirm NICHT blockierend — Aufrufstellen
 *   müssen `await` nutzen (Handler ggf. async machen).
 * - showAlert(message, opts?) → void (fire-and-forget, eine OK-Schaltfläche).
 * - Mehrere Dialoge hintereinander werden über eine FIFO-Queue serialisiert.
 * - Gerendert wird über das shared <Modal> (z-index 9999); der Provider
 *   rendert den Dialog NACH den children, damit er im DOM hinter allen
 *   Seiten-Modals liegt und damit über ihnen gezeichnet wird.
 * - useDialog() funktioniert auch OHNE Provider (z.B. Handbuch-Previews):
 *   dann fällt es auf die nativen Browser-Dialoge zurück.
 */
import * as React from 'react';
import Modal from '../components/Modal';
import { useLanguage } from './LanguageContext';

export interface ConfirmOptions {
  /** Überschrift; Default „Bitte bestätigen" / "Please confirm". */
  title?: string;
  /** Label des Bestätigen-Buttons; Default „OK". */
  confirmLabel?: string;
  /** Label des Abbrechen-Buttons; Default „Abbrechen" / "Cancel". */
  cancelLabel?: string;
  /** true = Bestätigen-Button in Rot (destruktive Aktionen wie Löschen). */
  danger?: boolean;
}

export interface AlertOptions {
  /** Überschrift; Default abhängig von variant. */
  title?: string;
  /** Färbt Kopfzeile/Akzent: info (Default), error (rot), success (grün). */
  variant?: 'info' | 'error' | 'success';
}

export interface PromptOptions {
  title?: string;
  defaultValue?: string;
  placeholder?: string;
  confirmLabel?: string;
}

interface DialogContextType {
  confirmDialog: (message: React.ReactNode, opts?: ConfirmOptions) => Promise<boolean>;
  showAlert: (message: React.ReactNode, opts?: AlertOptions) => void;
  /** Ersatz für window.prompt: Text-Eingabe-Modal. null = abgebrochen. */
  promptDialog: (message: React.ReactNode, opts?: PromptOptions) => Promise<string | null>;
}

const DialogContext = React.createContext<DialogContextType | null>(null);

type PendingDialog =
  | { kind: 'confirm'; message: React.ReactNode; opts: ConfirmOptions; resolve: (ok: boolean) => void }
  | { kind: 'alert'; message: React.ReactNode; opts: AlertOptions; resolve: () => void }
  | { kind: 'prompt'; message: React.ReactNode; opts: PromptOptions; resolve: (value: string | null) => void };

export function DialogProvider(props: { children: React.ReactNode }): React.ReactElement {
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const [queue, setQueue] = React.useState<PendingDialog[]>([]);
  const current = queue.length > 0 ? queue[0] : null;
  // Eingabewert des aktuell offenen Prompt-Dialogs.
  const [promptValue, setPromptValue] = React.useState('');
  React.useEffect(() => {
    if (current && current.kind === 'prompt') {
      setPromptValue((current.opts as PromptOptions).defaultValue || '');
    }
  }, [current]);

  const push = React.useCallback((d: PendingDialog): void => {
    setQueue(prev => [...prev, d]);
  }, []);
  const closeCurrent = React.useCallback((ok: boolean, promptResult?: string): void => {
    setQueue(prev => {
      const head = prev[0];
      if (head) {
        if (head.kind === 'confirm') head.resolve(ok);
        else if (head.kind === 'prompt') head.resolve(ok ? (promptResult ?? '') : null);
        else head.resolve();
      }
      return prev.slice(1);
    });
  }, []);

  const confirmDialog = React.useCallback(
    (message: React.ReactNode, opts: ConfirmOptions = {}): Promise<boolean> =>
      new Promise<boolean>(resolve => push({ kind: 'confirm', message, opts, resolve })),
    [push]
  );
  const showAlert = React.useCallback(
    (message: React.ReactNode, opts: AlertOptions = {}): void => {
      push({ kind: 'alert', message, opts, resolve: () => { /* fire-and-forget */ } });
    },
    [push]
  );

  const promptDialog = React.useCallback(
    (message: React.ReactNode, opts: PromptOptions = {}): Promise<string | null> =>
      new Promise<string | null>(resolve => push({ kind: 'prompt', message, opts, resolve })),
    [push]
  );

  const value = React.useMemo<DialogContextType>(
    () => ({ confirmDialog, showAlert, promptDialog }),
    [confirmDialog, showAlert, promptDialog]
  );

  // Darstellung des aktuell offenen Dialogs.
  let dialogNode: React.ReactNode = null;
  if (current) {
    const isConfirm = current.kind === 'confirm';
    const isPrompt = current.kind === 'prompt';
    const variant = current.kind === 'alert' ? ((current.opts as AlertOptions).variant || 'info') : undefined;
    const accent = isConfirm
      ? ((current.opts as ConfirmOptions).danger ? 'var(--dex-red, #da291c)' : 'var(--dex-green-dark, #4a7c1f)')
      : isPrompt
        ? 'var(--dex-green-dark, #4a7c1f)'
        : (variant === 'error' ? 'var(--dex-red, #da291c)' : variant === 'success' ? 'var(--dex-green-dark, #4a7c1f)' : 'var(--dex-gray-800, #333)');
    const defaultTitle = isConfirm
      ? (isDe ? 'Bitte bestätigen' : 'Please confirm')
      : isPrompt
        ? (isDe ? 'Eingabe' : 'Input')
        : (variant === 'error'
          ? (isDe ? 'Fehler' : 'Error')
          : variant === 'success'
            ? (isDe ? 'Erledigt' : 'Done')
            : (isDe ? 'Hinweis' : 'Note'));
    const title = current.opts.title || defaultTitle;
    const confirmOpts = current.opts as ConfirmOptions;
    const promptOpts = current.opts as PromptOptions;
    dialogNode = (
      <Modal
        open={true}
        onClose={() => closeCurrent(false)}
        maxWidth={460}
        ariaLabel={title}
      >
        <h3 style={{ margin: 0, fontSize: '1.02rem', color: accent }}>{title}</h3>
        <div style={{ fontSize: '0.9rem', color: 'var(--dex-gray-700, #444)', lineHeight: 1.55, whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>
          {current.message}
        </div>
        {isPrompt && (
          <input
            className="form-input"
            value={promptValue}
            placeholder={promptOpts.placeholder || ''}
            onChange={e => setPromptValue(e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') closeCurrent(true, promptValue); }}
            autoFocus
            style={{ fontSize: '0.9rem', padding: '9px 12px' }}
          />
        )}
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, marginTop: 4 }}>
          {(isConfirm || isPrompt) && (
            <button
              type="button"
              className="btn btn-secondary"
              onClick={() => closeCurrent(false)}
              style={{ fontSize: '0.88rem', padding: '9px 18px' }}
            >
              {(isConfirm && confirmOpts.cancelLabel) || (isDe ? 'Abbrechen' : 'Cancel')}
            </button>
          )}
          <button
            type="button"
            className="btn btn-primary"
            onClick={() => closeCurrent(true, isPrompt ? promptValue : undefined)}
            autoFocus={!isPrompt}
            style={{
              fontSize: '0.88rem', padding: '9px 18px',
              ...(isConfirm && confirmOpts.danger
                ? { background: 'var(--dex-red, #da291c)', borderColor: 'var(--dex-red, #da291c)' }
                : {}),
            }}
          >
            {isConfirm ? (confirmOpts.confirmLabel || 'OK') : isPrompt ? (promptOpts.confirmLabel || 'OK') : 'OK'}
          </button>
        </div>
      </Modal>
    );
  }

  return (
    <DialogContext.Provider value={value}>
      {props.children}
      {dialogNode}
    </DialogContext.Provider>
  );
}

export function useDialog(): DialogContextType {
  const ctx = React.useContext(DialogContext);
  // Fallback ohne Provider (z.B. Handbuch-Preview-Bäume): native Dialoge,
  // damit nichts hart bricht. ReactNode-Messages degradieren dabei zu ''.
  if (ctx) return ctx;
  return {
    confirmDialog: async (message: React.ReactNode): Promise<boolean> =>
      window.confirm(typeof message === 'string' ? message : ''),
    showAlert: (message: React.ReactNode): void => {
      window.alert(typeof message === 'string' ? message : '');
    },
    promptDialog: async (message: React.ReactNode, opts: PromptOptions = {}): Promise<string | null> =>
      window.prompt(typeof message === 'string' ? message : '', opts.defaultValue || ''),
  };
}
