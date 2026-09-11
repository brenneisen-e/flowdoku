/**
 * v31.12 — Auswertung der Event-Umfrage.
 *
 * Aufbau wie die anderen Auswertungs-Karten (docs/ui-leitfaden.md 5a Punkt 8):
 * eingeklappt, Kopf mit Zähler und Chevron rechts, darin Kennzahlen und die
 * Balken je Antwort.
 *
 * Zwei Dinge, die diese Karte anders macht als eine übliche Statistik:
 *
 *  - **Ein Lesefehler ist keine Null.** Konnte die Antwortliste nicht gelesen
 *    werden, steht hier „unbekannt" und der Grund — nicht „0 Antworten".
 *    Das ist die Lehre, die DEX rund sechzig Audit-Befunde gekostet hat
 *    (v30.66/v30.67).
 *  - **Bei einer anonymen Umfrage steht hier kein Name**, auch nicht in der
 *    Reihenfolge der Freitexte. Und der Kasten sagt, wo die Anonymität endet,
 *    statt sie zu versprechen: Wer Vollzugriff auf die Website hat, sieht in
 *    SharePoint selbst weiterhin, wer eine Zeile geschrieben hat.
 */

import * as React from 'react';
import { ChevronDown, AlertCircle } from '../../Icons';
import { cx, ensureDexUiStyles } from '../../dexUi';
import { DeloitteEvent } from '../../../types';
import { EventService } from '../../../services/EventService';
import type { PollConfig, PollStats } from '../../../services/events/poll';

export interface PollStatsSectionProps {
  selectedEvent: DeloitteEvent;
  eventServiceRef: EventService;
  isDe: boolean;
}

export const PollStatsSection: React.FC<PollStatsSectionProps> = ({ selectedEvent, eventServiceRef, isDe }) => {
  ensureDexUiStyles();
  const t = (de: string, en: string): string => (isDe ? de : en);

  const [open, setOpen] = React.useState(false);
  const [poll, setPoll] = React.useState<PollConfig | null>(null);
  const [stats, setStats] = React.useState<PollStats | null>(null);
  const [laedt, setLaedt] = React.useState(false);
  const [ladefehler, setLadefehler] = React.useState(false);
  const [schliesst, setSchliesst] = React.useState(false);

  const eventNumber = selectedEvent.eventNumber || 0;

  const laden = React.useCallback(async (): Promise<void> => {
    if (!eventServiceRef || !eventNumber) return;
    setLaedt(true);
    setLadefehler(false);
    try {
      const cfg = await eventServiceRef.getPoll(eventNumber);
      if (!cfg.ok) { setLadefehler(true); return; }
      setPoll(cfg.poll);
      if (!cfg.poll) { setStats(null); return; }
      setStats(await eventServiceRef.getPollStats(eventNumber, cfg.poll.anonym));
    } finally {
      setLaedt(false);
    }
  }, [eventServiceRef, eventNumber]);

  React.useEffect(() => { void laden(); }, [laden]);

  const umfrageSchliessen = React.useCallback(async (aktiv: boolean): Promise<void> => {
    if (!poll) return;
    setSchliesst(true);
    try {
      const ok = await eventServiceRef.savePoll({ ...poll, aktiv });
      if (ok) setPoll({ ...poll, aktiv });
    } finally {
      setSchliesst(false);
    }
  }, [poll, eventServiceRef]);

  // Keine Umfrage, kein Kasten — die Karte erscheint erst, wenn eine
  // Umfrage angelegt wurde (über „E-Mail versenden").
  if (!ladefehler && !poll) return null;

  const zaehler = stats ? stats.zaehler : {};
  const maxWert = Math.max(1, ...Object.keys(zaehler).map(k => zaehler[k]));
  const unlesbar = !!stats && stats.nichtLesbar;

  return (
    <div className="dex-ui-card" style={{ marginTop: 16 }}>
      <button
        type="button"
        className={cx('dex-ui-disclosure', open && 'is-open')}
        onClick={() => setOpen(o => !o)}
        aria-expanded={open}
      >
        <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
        {t('Umfrage auswerten', 'Poll results')}
        {stats && !unlesbar && <span className="dex-ui-disclosure-count">{stats.geantwortet}</span>}
      </button>

      {open && (
        <div className="dex-ui-disclosure-body">
          {ladefehler && (
            <div className="dex-ui-callout dex-ui-callout--danger" role="alert">
              <span className="dex-ui-callout-icon"><AlertCircle size={15} /></span>
              <span className="dex-ui-callout-body">
                {t('Die Umfrage konnte nicht gelesen werden. Ob es eine gibt und wie viele geantwortet haben, ist damit unbekannt — nicht null.',
                  'The poll could not be read. Whether one exists and how many answered is therefore unknown — not zero.')}
                <br />
                <button type="button" className="dex-ui-textbtn" style={{ marginTop: 6 }} onClick={() => { void laden(); }}>
                  {t('Erneut versuchen', 'Try again')}
                </button>
              </span>
            </div>
          )}

          {poll && (
            <>
              <div className="dex-ui-muted" style={{ fontSize: '0.85rem', marginBottom: 10 }}>{poll.frage}</div>

              {laedt && (
                <div className="dex-ui-progress dex-ui-progress--indeterminate"><div className="dex-ui-progress-bar" /></div>
              )}

              {unlesbar && (
                <div className="dex-ui-callout dex-ui-callout--warn" role="alert">
                  <span className="dex-ui-callout-icon"><AlertCircle size={15} /></span>
                  <span className="dex-ui-callout-body">
                    {t('Die Antwortliste war nicht lesbar. Die Zahlen unten wären keine Aussage über die Umfrage, deshalb stehen sie nicht da.',
                      'The answer list was not readable. The numbers below would say nothing about the poll, so they are not shown.')}
                    <br />
                    <button type="button" className="dex-ui-textbtn" style={{ marginTop: 6 }} onClick={() => { void laden(); }}>
                      {t('Erneut versuchen', 'Try again')}
                    </button>
                  </span>
                </div>
              )}

              {stats && !unlesbar && (
                <>
                  <div className="dex-ui-kpi-row">
                    <div className="dex-ui-kpi dex-ui-kpi--green">
                      <div className="dex-ui-kpi-value">{stats.geantwortet}</div>
                      <div className="dex-ui-kpi-label">{t('Antworten', 'Answers')}</div>
                    </div>
                    <div className="dex-ui-kpi dex-ui-kpi--gray">
                      <div className="dex-ui-kpi-value">{stats.freitexte.length}</div>
                      <div className="dex-ui-kpi-label">{t('mit Kommentar', 'with a comment')}</div>
                    </div>
                    <div className={cx('dex-ui-kpi', poll.aktiv ? 'dex-ui-kpi--blue' : 'dex-ui-kpi--gray')}>
                      <div className="dex-ui-kpi-value" style={{ fontSize: '1rem' }}>
                        {poll.aktiv ? t('offen', 'open') : t('beendet', 'closed')}
                      </div>
                      <div className="dex-ui-kpi-label">{t('Status', 'Status')}</div>
                    </div>
                  </div>

                  {stats.geantwortet === 0 ? (
                    <div className="dex-ui-empty" style={{ padding: '18px 0' }}>
                      <div className="dex-ui-empty-title">{t('Noch keine Antwort', 'No answers yet')}</div>
                      <div className="dex-ui-empty-desc">
                        {t('Sobald jemand ein Kästchen in der Mail anklickt, steht die Antwort hier.',
                          'As soon as somebody clicks a box in the email, the answer appears here.')}
                      </div>
                    </div>
                  ) : (
                    <div className="dex-ui-stack" style={{ gap: 10, marginTop: 12 }}>
                      {poll.optionen.map((opt, i) => {
                        const n = zaehler[opt] || 0;
                        const anteil = Math.round((n / stats.geantwortet) * 100);
                        return (
                          <div key={i}>
                            <div className="dex-ui-inline" style={{ justifyContent: 'space-between', fontSize: '0.85rem' }}>
                              <span>{opt}</span>
                              <span className="dex-ui-muted">{n} · {anteil}%</span>
                            </div>
                            <div className="dex-ui-progress">
                              <div className="dex-ui-progress-bar" style={{ width: `${Math.round((n / maxWert) * 100)}%` }} />
                            </div>
                          </div>
                        );
                      })}
                      {/* Antworten, die es in der Umfrage nicht mehr gibt — etwa
                          weil eine Antwortmöglichkeit nachträglich umbenannt
                          wurde. Sie einfach wegzulassen hiesse, Stimmen
                          verschwinden zu lassen. */}
                      {Object.keys(zaehler).filter(k => poll.optionen.indexOf(k) < 0).map(k => (
                        <div key={`alt-${k}`} className="dex-ui-inline" style={{ justifyContent: 'space-between', fontSize: '0.85rem' }}>
                          <span className="dex-ui-muted">{k} ({t('nicht mehr in der Umfrage', 'no longer in the poll')})</span>
                          <span className="dex-ui-muted">{zaehler[k]}</span>
                        </div>
                      ))}
                    </div>
                  )}

                  {stats.freitexte.length > 0 && (
                    <div style={{ marginTop: 16 }}>
                      <div className="dex-ui-section-title">{t('Was dazugeschrieben wurde', 'What people added')}</div>
                      <div className="dex-ui-stack" style={{ gap: 6 }}>
                        {stats.freitexte.map((f, i) => (
                          <div key={i} className="dex-ui-card dex-ui-card--sm dex-ui-card--soft" style={{ fontSize: '0.85rem' }}>{f}</div>
                        ))}
                      </div>
                    </div>
                  )}

                  {!poll.anonym && stats.emails.length > 0 && (
                    <div style={{ marginTop: 16 }}>
                      <div className="dex-ui-section-title">{t('Wer geantwortet hat', 'Who answered')}</div>
                      <div className="dex-ui-help" style={{ wordBreak: 'break-word' }}>{stats.emails.join(', ')}</div>
                    </div>
                  )}

                  <div className="dex-ui-muted" style={{ fontSize: '0.76rem', marginTop: 16, lineHeight: 1.5 }}>
                    {poll.anonym
                      ? t('Diese Umfrage ist als anonym angelegt: DEX zeigt dir keine Namen und sortiert die Kommentare, damit auch die Reihenfolge nichts verrät. Gespeichert bleibt, WER geantwortet hat — sonst könntest du niemanden erinnern. Wer Vollzugriff auf die SharePoint-Website hat, kann dort weiterhin sehen, wer eine Zeile geschrieben hat.',
                        'This poll is set to anonymous: DEX shows you no names and sorts the comments so that even the order gives nothing away. It does record WHO answered — otherwise you could not remind anyone. Anyone with full control of the SharePoint site can still see who wrote a row.')
                      : t('Diese Umfrage läuft mit Namen — du siehst oben, wer geantwortet hat.',
                        'This poll runs with names — you can see above who answered.')}
                  </div>

                  <div style={{ marginTop: 10 }}>
                    <button
                      type="button"
                      className="dex-ui-textbtn dex-ui-textbtn--muted"
                      style={{ marginLeft: -8 }}
                      disabled={schliesst}
                      onClick={() => { void umfrageSchliessen(!poll.aktiv); }}
                    >
                      {poll.aktiv
                        ? t('Umfrage beenden — Links in der Mail nehmen dann nichts mehr an', 'Close the poll — the links in the email then accept nothing')
                        : t('Umfrage wieder öffnen', 'Reopen the poll')}
                    </button>
                  </div>
                </>
              )}
            </>
          )}
        </div>
      )}
    </div>
  );
};
