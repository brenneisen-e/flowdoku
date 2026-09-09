/* CancelledEventsCollapsible — aus MyEventsPage.tsx ausgelagert (Zeilen
 * 3597-3674 des urspruenglichen Stands, v30.65). Einklappbare Liste der
 * abgemeldeten Events.
 *
 * v31.8: Optik nach `docs/ui-leitfaden.md` (Abschnitt 6). `ensureDexUiStyles()`
 * wird hier NICHT gerufen — das macht einmal die Seiten-Komponente
 * (`MyEventsPage`), Unterkomponenten nie.
 */
import * as React from 'react';
import { ChevronDown, Calendar, Pin, Info } from '../Icons';
import { cx } from '../dexUi';
import { CachedImg } from '../CachedImage';
import { extractTeamsUrl, locationWithoutTeamsUrl } from '../../utils/teamsLink';
import { formatDateTimeRange } from './myEventsHelpers';
// v31.8 (Nachzug): Die Komponente bekommt keine Sprache als Prop (die Signatur
// bleibt, `MyEventCard`/`MyEventsPage` rufen sie unverändert). Die
// Anzeigesprache kommt deshalb aus dem Context — sonst steht ein deutsches
// Datum in der englischen Oberfläche.
import { useLocaleSafe } from '../../context/LanguageContext';

// v11.97: Einklappbare Liste der abgemeldeten Events. Default eingeklappt,
// damit lange Cancelled-Listen die Hauptliste nicht überlagern. Header
// zeigt Count + Chevron, Klick togglt die Liste.
export default function CancelledEventsCollapsible(props: {
  count: number;
  title: string;
  locale: string;
  // v31.8: `event` trägt jetzt zusätzlich Datum, Ort und Bild — alle drei
  // optional, damit der Aufrufer unverändert seine `MyEventEntry[]` übergibt.
  // Grund: Die Zeile verlor bisher alles außer Titel und Abmeldedatum, womit
  // „wann war das noch mal?" nicht mehr zu beantworten war (Leitfaden 6c).
  entries: Array<{
    event: { id: string; title: string; startDate?: string; endDate?: string; location?: string; imageUrl?: string };
    registration: { CancellationDate?: string };
  }>;
  formatDate: (iso: string) => string;
  statusLabel: string;
  // v18.62: Hinweis + Button, dass man sich für eine erneute Teilnahme
  // im Anmelde-Bereich neu registrieren muss (Abmeldung ist endgültig).
  hintText: string;
  reRegisterLabel: string;
  onReRegister: () => void;
}): React.ReactElement {
  // v31.8 (Nachzug): Hooks stehen zusammen ganz oben — die Komponente hat heute
  // keinen frühen Return, und ein Hook hinter einem später ergänzten Return
  // reisst die Hook-Reihenfolge (React #300, siehe CLAUDE.md zu v30.3).
  const isDe = useLocaleSafe() === 'de';
  const [open, setOpen] = React.useState(false);
  return (
    <div className="mt-24">
      {/* v31.8: Vorher ein handgebauter Knopf mit ▶ im Text — ein Zeichen ist
          kein Bedienhinweis. Jetzt der gemeinsame Aufklapper mit ChevronDown,
          Hover und Tastaturfokus; die Zahl steht rechts, wie überall. */}
      <button
        type="button"
        className={cx('dex-ui-disclosure', open && 'is-open')}
        onClick={() => setOpen(o => !o)}
        aria-expanded={open}
      >
        <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
        <span>{props.title}</span>
        <span className="dex-ui-disclosure-count">{props.count}</span>
      </button>
      {open && (
        <div className="dex-ui-disclosure-body">
          {/* v31.8: Der Knopf hing per `flex: 1 1 auto` am rechten Rand. Er
              gehört zum Satz, also steht er linksbündig darunter (2a′). */}
          <div className="dex-ui-callout dex-ui-callout--neutral" style={{ marginBottom: 12 }}>
            <span className="dex-ui-callout-icon"><Info size={16} /></span>
            <div>
              <div>{props.hintText}</div>
              <button
                type="button"
                className="btn btn-outline dex-ui-btn-sm"
                style={{ marginTop: 8 }}
                onClick={props.onReRegister}
              >{props.reRegisterLabel}</button>
            </div>
          </div>
          {/* Die aktiven Karten stehen im DOM VOR dieser Liste — der
              Tour-Selektor `.my-event-card` greift das erste Vorkommen. Die
              Klasse bleibt deshalb auch hier stehen. */}
          <div className="my-events-list">
            {props.entries.map(({ event, registration }) => {
              // v31.8: Ort ohne Teams-URL — sonst läuft bei Online-Events ein
              // roher Link durch die Zeile (gleiche Aufbereitung wie auf der
              // aktiven Karte).
              // v31.8 (Nachzug): `locationWithoutTeamsUrl` ist dafür gebaut,
              // dass der Teams-Link DANEBEN als Knopf steht. In dieser Zeile
              // gibt es keinen — steht im Ort-Feld nur die URL, bliebe die
              // Ortsangabe sonst leer. Dann „Online", wortgleich zur aktiven
              // Karte (`MyEventCard`, v29.39).
              const bareLocation = locationWithoutTeamsUrl(event.location);
              const where = bareLocation || (extractTeamsUrl(event.location) ? 'Online' : '');
              // v31.8 (Nachzug): Ohne drittes Argument formatiert der Helfer
              // immer 'de-DE' (Default `isDe = true`) — im englischen UI stand
              // damit ein deutsches Datum.
              const when = formatDateTimeRange(event.startDate, event.endDate, isDe);
              return (
                <div
                  key={event.id}
                  // `dex-ui-card--muted`: gedämpft im Ruhezustand, beim
                  // Überfahren wieder voll lesbar — der Zustand „abgemeldet"
                  // versteckt die Daten nicht, er nimmt ihnen nur das Gewicht.
                  className="card my-event-card dex-ui-card--muted"
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: 12,
                    padding: '10px 16px',
                    flexWrap: 'wrap',
                  }}
                >
                  {event.imageUrl && (
                    <div
                      className="my-event-card__thumb"
                      style={{
                        flexShrink: 0,
                        width: 88,
                        height: 64,
                        borderRadius: 'var(--dex-radius, 12px)',
                        background: 'var(--dex-gray-50, #fafafa)',
                        border: '1px solid var(--dex-gray-200)',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        overflow: 'hidden',
                      }}
                    >
                      <CachedImg
                        src={event.imageUrl}
                        alt={event.title}
                        loading="lazy"
                        decoding="async"
                        style={{ maxWidth: '100%', maxHeight: '100%', objectFit: 'contain', display: 'block' }}
                      />
                    </div>
                  )}
                  <div style={{ flex: '1 1 220px', minWidth: 0 }}>
                    <strong style={{ fontSize: '0.95rem' }}>{event.title}</strong>
                    {(when || where) && (
                      <div style={{
                        marginTop: 4, display: 'flex', flexWrap: 'wrap',
                        gap: '2px 18px', fontSize: '0.78rem', color: 'var(--dex-gray-600)',
                      }}>
                        {when && (
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                            <Calendar size={13} strokeWidth={2} />{when}
                          </span>
                        )}
                        {where && (
                          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, minWidth: 0 }}>
                            <Pin size={13} strokeWidth={2} />
                            <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{where}</span>
                          </span>
                        )}
                      </div>
                    )}
                  </div>
                  <span style={{ fontSize: '0.78rem', color: 'var(--dex-gray-500)' }}>
                    {props.locale}: {registration.CancellationDate ? props.formatDate(registration.CancellationDate) : '-'}
                  </span>
                  {/* Reine Anzeige, also Pille statt Chip — die Farbe bleibt
                      rot wie auf der aktiven Karte (`getStatusBadgeClass`). */}
                  <span className="dex-ui-pill dex-ui-pill--red">{props.statusLabel}</span>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}
