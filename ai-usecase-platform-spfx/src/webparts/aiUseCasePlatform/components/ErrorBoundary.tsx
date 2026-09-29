/**
 * Fängt Fehler beim Rendern einer Seite ab — statt das ganze Webpart zu leeren.
 *
 * Es gab bis v1.3 keine Grenze: Ein Lazy-Chunk, der nach einem Update nicht mehr
 * da ist (SharePoint liefert den alten Dateinamen mit 404), oder jede Ausnahme im
 * Render ließ das gesamte Webpart weiß, ohne Wort. Sichtprüfung/Review
 * 29.09.2026: Schon ein Titel wie „-" warf in `kuerzel()`.
 *
 * Eine Klasse, weil React Grenzen nur so kennt. Die Sprache kommt als Eigenschaft
 * (eine Klasse kann `useLanguage` nicht rufen), der Text ist bewusst kurz und sagt,
 * was man tun kann.
 */

import * as React from 'react';

interface Props {
  isDe: boolean;
  /** Ändert sich der Schlüssel (Seitenwechsel), wird die Grenze zurückgesetzt. */
  resetKey: string;
  children: React.ReactNode;
}

interface State {
  fehler: Error | null;
  key: string;
}

export default class ErrorBoundary extends React.Component<Props, State> {
  state: State = { fehler: null, key: this.props.resetKey };

  static getDerivedStateFromError(fehler: Error): Partial<State> {
    return { fehler };
  }

  static getDerivedStateFromProps(props: Props, state: State): Partial<State> | null {
    // Wer auf eine andere Seite wechselt, soll nicht am Fehler der vorigen hängen.
    return props.resetKey !== state.key ? { fehler: null, key: props.resetKey } : null;
  }

  componentDidCatch(fehler: Error): void {
    console.error('[AIUC] Seite konnte nicht dargestellt werden:', fehler);
  }

  render(): React.ReactNode {
    if (!this.state.fehler) return this.props.children;
    const de = this.props.isDe;
    return (
      <div className="dex-ui-callout dex-ui-callout--danger" role="alert" style={{ margin: '24px 0' }}>
        <span className="dex-ui-callout-body">
          <strong>{de ? 'Diese Seite konnte nicht angezeigt werden.' : 'This page could not be displayed.'}</strong>
          <br />
          {de
            ? 'Meist hilft es, die Seite neu zu laden — etwa wenn die App gerade aktualisiert wurde. Bleibt es dabei, melde dich über „Hast du Fragen?".'
            : 'Reloading the page usually helps — for example right after the app was updated. If it stays this way, use "Questions?" in the header.'}
          <br />
          <button type="button" className="btn btn-primary" style={{ marginTop: 10 }} onClick={() => window.location.reload()}>
            {de ? 'Seite neu laden' : 'Reload page'}
          </button>
        </span>
      </div>
    );
  }
}
