/**
 * v31.32 — Die Feedback-Übersicht für Admins.
 *
 * Nutzer-Frage 14.09.2026: „und wo kann ich jetzt als Admin das Feedback
 * aufrufen?" Antwort bis dahin: nirgends. Das Feedback ging als Mail an
 * `dex.event@deloitte.de`, und seit v31.31 liegt der letzte Stand am Event —
 * aber nur als Wiedervorlage für die Person selbst. Diese Seite macht daraus
 * eine Ansicht.
 *
 * ## Woher die Daten kommen — und was das begrenzt
 *
 * Aus den Events selbst: Piggyback `_feedback` in `EmailTemplateOverrides`
 * (siehe `utils/dexFeedback`). Kein zusätzlicher Abruf, keine neue Liste — die
 * Events liegen ohnehin im `EventContext`.
 *
 * Daraus folgen zwei Grenzen, und sie stehen ausdrücklich AUF der Seite, statt
 * dass der Leser sie sich erschließen muss:
 *
 *  - **Erst ab v31.31.** Ältere Rückmeldungen existieren ausschließlich als
 *    Mail im Postfach des DEX-Teams. Eine Seite, die „3 Rückmeldungen" zeigt,
 *    obwohl zwanzig gegeben wurden, ist schlimmer als keine Seite.
 *  - **Nur sichtbare Events.** Wer ein Event nicht sehen darf, sieht auch sein
 *    Feedback nicht. Für Admins ist das in der Regel alles, aber „in der Regel"
 *    ist keine Zusage.
 *
 * ## Warum die Auszählung neben den Einzeltexten steht
 *
 * Angehakte Antworten und Freitext beantworten verschiedene Fragen: Die Haken
 * sagen, WIE OFT etwas auftritt, der Freitext sagt, WAS genau. Nur die
 * Auszählung zu zeigen verliert die Begründung; nur die Einzeltexte zu zeigen
 * macht aus einer Häufung eine Anekdote.
 */

import * as React from 'react';
import { BarChart3, Download, MessageSquare } from './Icons';
import { cx, ensureDexUiStyles } from './dexUi';
import { useLanguage } from '../context/LanguageContext';
import { useEvents } from '../context/EventContext';
import {
  FEEDBACK_BESSER, FEEDBACK_GUT, FeedbackOption, FeedbackStand, readFeedbackMap,
} from '../utils/dexFeedback';

interface Zeile extends FeedbackStand {
  eventId: string;
  eventTitle: string;
  eventStart?: string;
  email: string;
}

/** Ein Datum kurz und in der Sprache der Oberfläche. */
function datum(iso: string | undefined, isDe: boolean): string {
  if (!iso) return '—';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '—';
  return d.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

const FeedbackOverviewPage: React.FC = () => {
  ensureDexUiStyles();
  const { locale } = useLanguage();
  const isDe = locale === 'de';
  const t = (de: string, en: string): string => (isDe ? de : en);
  const { events, isEventsLoading } = useEvents();
  const [suche, setSuche] = React.useState('');

  const zeilen = React.useMemo<Zeile[]>(() => {
    const out: Zeile[] = [];
    for (const ev of events || []) {
      const map = readFeedbackMap(ev.emailTemplateOverrides);
      for (const email of Object.keys(map)) {
        const st = map[email];
        if (!st || !Array.isArray(st.gut)) continue;
        out.push({
          ...st,
          eventId: String(ev.id),
          eventTitle: ev.title || '',
          eventStart: ev.startDate,
          email,
        });
      }
    }
    // Neueste zuerst — wer die Seite öffnet, will wissen, was zuletzt kam.
    out.sort((a, b) => String(b.at || '').localeCompare(String(a.at || '')));
    return out;
  }, [events]);

  const gefiltert = React.useMemo<Zeile[]>(() => {
    const q = suche.trim().toLowerCase();
    if (!q) return zeilen;
    return zeilen.filter(z => [z.eventTitle, z.email, z.name || '', z.freitextGut, z.freitextBesser]
      .join(' ').toLowerCase().indexOf(q) >= 0);
  }, [zeilen, suche]);

  /** Wie oft wurde welche Antwort angehakt — über die GEFILTERTE Menge. */
  const zaehlung = (katalog: FeedbackOption[], feld: 'gut' | 'besser'): Array<{ o: FeedbackOption; n: number }> => {
    const zahlen: Record<string, number> = {};
    for (const z of gefiltert) for (const id of (z[feld] || [])) zahlen[id] = (zahlen[id] || 0) + 1;
    return katalog.map(o => ({ o, n: zahlen[o.id] || 0 })).sort((a, b) => b.n - a.n);
  };

  const csv = (): void => {
    const kopf = ['Event', 'Event-Id', 'Von', 'E-Mail', 'Datum', 'Was lief gut', 'Was verbessern', 'Text gut', 'Text verbessern'];
    const txt = (ids: string[], katalog: FeedbackOption[]): string =>
      ids.map(id => (katalog.filter(o => o.id === id)[0] || { de: id }).de).join(' | ');
    const q = (s: string): string => `"${String(s || '').replace(/"/g, '""')}"`;
    const lines = [kopf.map(q).join(';')].concat(gefiltert.map(z => [
      z.eventTitle, z.eventId, z.name || '', z.email, z.at || '',
      txt(z.gut || [], FEEDBACK_GUT), txt(z.besser || [], FEEDBACK_BESSER),
      z.freitextGut || '', z.freitextBesser || '',
    ].map(q).join(';')));
    // BOM voran, sonst zeigt Excel die Umlaute kaputt an.
    const blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'DEX_Feedback.csv';
    a.click();
    URL.revokeObjectURL(a.href);
  };

  const balken = (eintraege: Array<{ o: FeedbackOption; n: number }>, farbe: string): React.ReactElement => {
    const max = Math.max(1, ...eintraege.map(e => e.n));
    return (
      <div className="dex-ui-stack" style={{ gap: 6 }}>
        {eintraege.map(({ o, n }) => (
          <div key={o.id} style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <span style={{ flex: '1 1 auto', fontSize: '0.82rem', opacity: n > 0 ? 1 : 0.45 }}>{isDe ? o.de : o.en}</span>
            <span style={{ flex: '0 0 120px', background: 'var(--dex-gray-100, #f3f4f6)', borderRadius: 999, height: 8, overflow: 'hidden' }}>
              <span style={{ display: 'block', width: `${(n / max) * 100}%`, background: farbe, height: '100%' }} />
            </span>
            <span style={{ flex: '0 0 28px', textAlign: 'right', fontSize: '0.82rem', fontWeight: 700, opacity: n > 0 ? 1 : 0.45 }}>{n}</span>
          </div>
        ))}
      </div>
    );
  };

  return (
    <div className="page-container">
      <div className="dex-ui-card-head" style={{ marginBottom: 10 }}>
        <h1 className="dex-ui-card-head-title" style={{ margin: 0, display: 'inline-flex', alignItems: 'center', gap: 8 }}>
          <MessageSquare size={20} /> {t('Feedback der Organizer', 'Organizer feedback')}
        </h1>
        <span className="dex-ui-card-head-meta">
          {gefiltert.length === zeilen.length
            ? t(`${zeilen.length} Rückmeldungen`, `${zeilen.length} responses`)
            : t(`${gefiltert.length} von ${zeilen.length} Rückmeldungen`, `${gefiltert.length} of ${zeilen.length} responses`)}
        </span>
        <span className="dex-ui-card-head-actions">
          <button type="button" className="dex-ui-chip" onClick={csv} disabled={gefiltert.length === 0}>
            <Download size={14} /> CSV
          </button>
        </span>
      </div>

      {/* Die zwei Grenzen der Ansicht — oben, nicht im Kleingedruckten. */}
      <div className="dex-ui-callout" role="note" style={{ marginBottom: 14 }}>
        <BarChart3 size={16} />
        <span>
          {t('Gezeigt wird, was seit Version 31.31 gespeichert wurde — ältere Rückmeldungen liegen nur als Mail im Postfach des DEX-Teams. Und nur Events, die du sehen darfst.',
            'Shown is what has been stored since version 31.31 — older responses exist only as email in the DEX team mailbox. And only events you are allowed to see.')}
        </span>
      </div>

      {isEventsLoading && zeilen.length === 0 ? (
        <div className="dex-ui-empty" role="status" aria-live="polite">
          <div className="dex-ui-progress dex-ui-progress--indeterminate"><div className="dex-ui-progress-bar" /></div>
          <div className="dex-ui-empty-title" style={{ marginTop: 14 }}>{t('Einen Moment …', 'One moment …')}</div>
        </div>
      ) : zeilen.length === 0 ? (
        <div className="dex-ui-empty">
          <div className="dex-ui-empty-title">{t('Noch keine Rückmeldung', 'No responses yet')}</div>
          <div className="dex-ui-empty-desc">
            {t('Organizer bekommen den Feedback-Knopf in der Mail nach dem Event und noch einmal kurz vor dem Löschen der Teilnehmerliste.',
              'Organizers get the feedback button in the email after the event and again shortly before the participant list is deleted.')}
          </div>
        </div>
      ) : (
        <>
          <div className="dex-ui-searchbar" style={{ marginBottom: 16, maxWidth: 420 }}>
            <input
              className="dex-ui-input"
              value={suche}
              onChange={e => setSuche(e.target.value)}
              placeholder={t('Event, Person oder Text suchen', 'Search event, person or text')}
            />
          </div>

          <div className="dex-ui-section">
            <div className="dex-ui-section-title">{t('Wie oft wurde was angehakt?', 'How often was each answer ticked?')}</div>
            <div className="dex-ui-section-desc">
              {t('Über die aktuell angezeigten Rückmeldungen.', 'Across the responses currently shown.')}
            </div>
            <div className="dex-ui-card" style={{ padding: 16, marginBottom: 10 }}>
              <div className="dex-ui-label" style={{ marginBottom: 8 }}>{t('Was lief gut', 'What went well')}</div>
              {balken(zaehlung(FEEDBACK_GUT, 'gut'), 'var(--dex-green, #86bc25)')}
            </div>
            <div className="dex-ui-card" style={{ padding: 16 }}>
              <div className="dex-ui-label" style={{ marginBottom: 8 }}>{t('Was wir verbessern können', 'What we can improve')}</div>
              {balken(zaehlung(FEEDBACK_BESSER, 'besser'), 'var(--dex-orange, #ed8b00)')}
            </div>
          </div>

          <div className="dex-ui-section">
            <div className="dex-ui-section-title">{t('Die einzelnen Rückmeldungen', 'The individual responses')}</div>
            <div className="dex-ui-stack">
              {gefiltert.map(z => (
                <div key={`${z.eventId}|${z.email}`} className="dex-ui-card" style={{ padding: 16 }}>
                  <div className="dex-ui-card-head" style={{ marginBottom: 8 }}>
                    <span className="dex-ui-card-head-title" style={{ fontSize: '0.95rem' }}>{z.eventTitle}</span>
                    <span className="dex-ui-card-head-meta">
                      {z.name || z.email} · {datum(z.at, isDe)}
                    </span>
                  </div>
                  <div style={{ fontSize: '0.8rem', color: 'var(--dex-gray-600)', marginBottom: 8 }}>{z.email}</div>
                  {(z.gut || []).length > 0 && (
                    <div style={{ marginBottom: 6 }}>
                      {(z.gut || []).map(id => {
                        const o = FEEDBACK_GUT.filter(x => x.id === id)[0];
                        return o ? <span key={id} className={cx('dex-ui-pill', 'dex-ui-pill--green')} style={{ marginRight: 6, marginBottom: 4, display: 'inline-block' }}>{isDe ? o.de : o.en}</span> : null;
                      })}
                    </div>
                  )}
                  {(z.besser || []).length > 0 && (
                    <div style={{ marginBottom: 6 }}>
                      {(z.besser || []).map(id => {
                        const o = FEEDBACK_BESSER.filter(x => x.id === id)[0];
                        return o ? <span key={id} className={cx('dex-ui-pill', 'dex-ui-pill--orange')} style={{ marginRight: 6, marginBottom: 4, display: 'inline-block' }}>{isDe ? o.de : o.en}</span> : null;
                      })}
                    </div>
                  )}
                  {!!(z.freitextGut || '').trim() && (
                    <p style={{ margin: '6px 0 0', fontSize: '0.86rem', borderLeft: '3px solid var(--dex-green, #86bc25)', paddingLeft: 10, whiteSpace: 'pre-wrap' }}>
                      {z.freitextGut}
                    </p>
                  )}
                  {!!(z.freitextBesser || '').trim() && (
                    <p style={{ margin: '6px 0 0', fontSize: '0.86rem', borderLeft: '3px solid var(--dex-orange, #ed8b00)', paddingLeft: 10, whiteSpace: 'pre-wrap' }}>
                      {z.freitextBesser}
                    </p>
                  )}
                </div>
              ))}
            </div>
          </div>
        </>
      )}
    </div>
  );
};

export default FeedbackOverviewPage;
