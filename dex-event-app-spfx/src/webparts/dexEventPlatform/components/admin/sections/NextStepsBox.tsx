/* NextStepsBox — „Nächste Schritte" für ein Event im Entwurf.
 * Ausgelagert aus AdminPage.tsx (Zeilen 7632-7875 des Stands vor dem Schnitt);
 * die Anzeige-Bedingung bleibt beim Aufrufer.
 *
 * v31.3: Nach docs/ui-leitfaden.md 5a (Punkt 4) umgebaut — jeder Schritt ist
 * eine `dex-ui-step`-Zeile mit dem Knopf darin (vorher hing die einzige Aktion
 * als unterstrichener Textlink mitten im Fließtext). Erklärtext auf zwei Zeilen
 * gekürzt, alles Längere (fehlende Angaben, Sichtbarkeit, Abweichungen) als
 * Hinweiskasten darunter; die grüne Flächenfarbe ist weg (Leitfaden 1.1).
 */
import * as React from 'react';
import { AlertCircle, ChevronDown, Info, Mail } from '../../Icons';
import { cx, ensureDexUiStyles } from '../../dexUi';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { DeloitteEvent } from '../../../types';

export interface NextStepsBoxProps {
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
  isDe: boolean;
  openInviteModal: () => void;
  selectedEvent: DeloitteEvent;
  setVisListOpen: React.Dispatch<React.SetStateAction<boolean>>;
  visListOpen: boolean;
}

export const NextStepsBox: React.FC<NextStepsBoxProps> = (p) => {
  const { childEventsOf, isDe, openInviteModal, selectedEvent, setVisListOpen, visListOpen } = p;
  // Idempotent — Modal und WizardFormShell rufen es ebenfalls; hier nötig, weil
  // die Box ohne offenes Modal in der Event-Seite steht.
  ensureDexUiStyles();
  return (
          <aside style={{ flex: '1 1 360px', minWidth: 320 }}>
            <div className="dex-ui-card dex-ui-card--accent">
              <div className="dex-ui-card-head" style={{ marginBottom: 4 }}>
                <h3 className="dex-ui-card-head-title">
                  <span style={{ color: 'var(--dex-green-dark, #4a7c1f)', display: 'inline-flex' }}><Info size={18} /></span>
                  {isDe ? 'Nächste Schritte' : 'Next steps'}
                </h3>
              </div>
              <p style={{ margin: '0 0 12px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {isDe ? 'Dein Event ist angelegt — so machst du es startklar:' : 'Your event is created — here is how to get it ready:'}
              </p>
              <ol className="dex-ui-stack" style={{ margin: 0, padding: 0, listStyle: 'none' }}>
                {(() => {
                  // v22.6: Sichtbarkeit eines Events als Klartext (für Haupt-
                  // event UND je Sub-Section).
                  const visText = (lc: string[], au: string[]): string => {
                    if (lc.length === 0 && au.length === 0) {
                      return isDe ? 'alle Mitarbeiter von Deloitte Deutschland' : 'all Deloitte Germany employees';
                    }
                    // v28.76: „1 Verteiler/Personen" war grammatisch schief —
                    // Singular und Plural in einem Wort. Jetzt sauber gebeugt.
                    const parts: string[] = [];
                    if (lc.length) parts.push((isDe ? (lc.length === 1 ? 'Standort ' : 'Standorte ') : (lc.length === 1 ? 'location ' : 'locations ')) + lc.join(', '));
                    if (au.length) {
                      parts.push(isDe
                        ? (au.length === 1 ? '1 Verteiler bzw. Person' : `${au.length} Verteiler bzw. Personen`)
                        : (au.length === 1 ? '1 distribution list or person' : `${au.length} distribution lists / people`));
                    }
                    return parts.join(isDe ? ' und ' : ' and ');
                  };
                  const locs = (selectedEvent.locationAudience || []).filter(Boolean);
                  const auds = (selectedEvent.audienceFilter || []).filter(Boolean);
                  const children = childEventsOf(selectedEvent.id);
                  const hasChildren = children.length > 0;
                  const parentVisText = visText(locs, auds);
                  const visSummary = (isDe ? 'Sichtbar für ' : 'Visible to ') + parentVisText + '.';
                  // Pro Sub-Section die Sichtbarkeit; wenn alle gleich → nur einmal.
                  // v22.22: Eine Sub-Section OHNE eigene Filter ist zur Laufzeit
                  // NICHT für „alle Mitarbeiter" sichtbar — der Zugang läuft immer
                  // über das Gesamt-Event (dessen Sichtbarkeit gilt dann auch für
                  // die Sub-Section). Das auch so benennen, statt irreführend
                  // „alle Mitarbeiter von Deloitte Deutschland" anzuzeigen.
                  const parentRestricted = locs.length > 0 || auds.length > 0;
                  const childVis = children.map(c => {
                    const cl = (c.locationAudience || []).filter(Boolean);
                    const ca = (c.audienceFilter || []).filter(Boolean);
                    const inherits = cl.length === 0 && ca.length === 0 && parentRestricted;
                    return {
                      title: shortSubEventTitle(c.title, selectedEvent.title) || c.title,
                      inherits,
                      text: inherits
                        ? (isDe ? 'wie das Gesamt-Event (keine eigene Einschränkung)' : 'same as the overall event (no own restriction)')
                        : visText(cl, ca),
                    };
                  });
                  const allChildrenSame = childVis.length > 0 && childVis.every(c => c.text === childVis[0].text);
                  // v22.8: Wenn Gesamt-Event UND alle Sub-Sections dieselbe
                  // Sichtbarkeit haben, ist die Unterscheidung überflüssig — dann
                  // nur EINE Aussage zeigen. (v22.22: gilt auch, wenn alle
                  // Sub-Sections die Sichtbarkeit des Gesamt-Events erben.)
                  const everythingSame = hasChildren && allChildrenSame && (childVis[0].text === parentVisText || childVis[0].inherits);
                  // v26.43: Konkrete Fehlt-Hinweise statt nur generischem Text —
                  // welche Grundangaben sind noch leer? (Organizer-Feedback: „warum
                  // sagt mir die Box nicht, dass die Beschreibung fehlt?")
                  const missingBits: string[] = [];
                  // v28.79: „Keine Beschreibung nutzen" ist eine Entscheidung,
                  // kein Versaeumnis — der Wizard legt dafuer das Flag
                  // `_noDescription` ab. Ohne diese Ausnahme meldete die Box
                  // die Beschreibung dauerhaft als fehlend, obwohl der
                  // Organizer sie bewusst weggelassen hatte.
                  const descriptionWaived = ((): boolean => {
                    try {
                      const ov = JSON.parse(selectedEvent.emailTemplateOverrides || '{}');
                      return !!(ov && ov._noDescription);
                    } catch { return false; }
                  })();
                  if (!(selectedEvent.description || '').trim() && !descriptionWaived) missingBits.push(isDe ? 'Beschreibung' : 'description');
                  const locationMissing = !(selectedEvent.location || '').trim();
                  if (locationMissing) missingBits.push(isDe ? 'Ort' : 'location');
                  if (!(selectedEvent.imageUrl || '').trim()) missingBits.push(isDe ? 'Event-Bild' : 'event image');
                  // v31.3: Die Abweichungs-Warnung steckte im selben grauen
                  // Kasten wie die Sichtbarkeits-Aufzählung — sie ist aber das
                  // Wichtigste an dieser Stelle und bekommt jetzt einen eigenen
                  // Warn-Kasten. Die Bedingung ist unverändert (Sub-Sections
                  // vorhanden, nicht alles gleich, Kinder untereinander
                  // verschieden); die Zählung darin ist wortgleich übernommen.
                  const varianceWarn: React.ReactNode = (hasChildren && !everythingSame && !allChildrenSame) ? (() => {
                    const groups = new Map<string, string[]>();
                    childVis.forEach(c => {
                      const arr = groups.get(c.text) || [];
                      arr.push(c.title);
                      groups.set(c.text, arr);
                    });
                    if (groups.size < 2) return null;
                    const sorted = Array.from(groups.entries()).sort((a, b) => b[1].length - a[1].length);
                    const minority = sorted.slice(1);
                    const minCount = minority.reduce((s, e) => s + e[1].length, 0);
                    return (
                      <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm">
                        <span className="dex-ui-callout-icon"><AlertCircle size={14} /></span>
                        <div style={{ minWidth: 0 }}>
                          {isDe
                            ? <>
                                <strong>Achtung — die Sub-Sections sind unterschiedlich sichtbar</strong> ({groups.size} Varianten).{' '}
                                {sorted[0][1].length} {sorted[0][1].length === 1 ? 'Sub-Section hat' : 'Sub-Sections haben'} &bdquo;{sorted[0][0]}&ldquo;, {minCount} {minCount === 1 ? 'weicht ab' : 'weichen ab'}: {minority.map(e => e[1].join(', ')).join('; ')}.{' '}
                                Wenn das nicht gewollt ist: in &bdquo;Event bearbeiten&ldquo; den Schritt <strong>Kapazität &amp; Sichtbarkeit</strong> öffnen, die passende Sub-Section wählen und dort auf <strong>&bdquo;Einstellungen auf andere übertragen&ldquo;</strong> klicken.
                              </>
                            : <>
                                <strong>Careful — the sub-sections have different visibility</strong> ({groups.size} variants).{' '}
                                {sorted[0][1].length} of them use &bdquo;{sorted[0][0]}&ldquo;, {minCount} differ: {minority.map(e => e[1].join(', ')).join('; ')}.{' '}
                                If that is not intended: open &bdquo;Edit event&ldquo; &rarr; step <strong>Capacity &amp; visibility</strong>, pick the right sub-section and use <strong>&bdquo;Transfer settings to others&ldquo;</strong>.
                              </>}
                        </div>
                      </div>
                    );
                  })() : null;
                  // v31.3: Ein Schritt = Titel (was zu tun ist) + höchstens zwei
                  // Zeilen Folge + optional ein Knopf und ein Hinweiskasten.
                  // `done` dämpft nur die Nummer — erledigte Schritte bleiben
                  // sichtbar (Leitfaden 5a).
                  const steps: Array<{ title: string; hint: React.ReactNode; done?: boolean; doneLabel?: string; action?: React.ReactNode; extra?: React.ReactNode }> = [
                    {
                      title: isDe ? 'Event finalisieren' : 'Finalize the event',
                      hint: isDe
                        ? 'Ergänze über „Event bearbeiten" die Felder, das Bild und die Texte.'
                        : 'Use “Edit event” to complete fields, image and texts.',
                      // v31.3: Ein Schritt, dem nichts mehr fehlt, ist erledigt —
                      // das ist dieselbe Prüfung wie der Fehlt-noch-Kasten.
                      done: missingBits.length === 0,
                      doneLabel: isDe ? 'Angaben vollständig' : 'All details filled in',
                      extra: missingBits.length > 0 ? (
                        <div className="dex-ui-callout dex-ui-callout--warn dex-ui-callout--sm">
                          <span className="dex-ui-callout-icon"><AlertCircle size={14} /></span>
                          <div style={{ minWidth: 0 }}>
                            {isDe
                              ? <>Fehlt noch: <strong>{missingBits.join(', ')}</strong>.</>
                              : <>Still missing: <strong>{missingBits.join(', ')}</strong>.</>}
                            {/* v28.79: Beim Ort die Folge benennen — ohne ihn
                                steht im Outlook-Termin und in den Mails nur
                                das Wort „Veranstaltungsort" als Platzhalter. */}
                            {locationMissing && (
                              <div style={{ marginTop: 4 }}>
                                {isDe
                                  ? <>Ohne Ort erscheint im Outlook-Termin und in den Mails nur das Wort &bdquo;Veranstaltungsort&ldquo; &mdash; die Teilnehmer wissen dann nicht, wohin sie kommen sollen.</>
                                  : <>Without a location the Outlook invite and the emails only show the word &bdquo;venue&ldquo; &mdash; attendees will not know where to go.</>}
                              </div>
                            )}
                          </div>
                        </div>
                      ) : undefined,
                    },
                    {
                      title: isDe ? 'Test-An- und Abmeldung' : 'Test registration & cancellation',
                      hint: isDe
                        ? 'Melde dich einmal selbst an und wieder ab — so siehst du, ob Bestätigungs-Mail, Outlook-Termin und Abmelde-Mail richtig ankommen.'
                        : 'Register and cancel yourself once — that shows you whether the confirmation email, Outlook invite and cancellation email arrive correctly.',
                    },
                    {
                      title: isDe ? 'Event live schalten' : 'Publish the event',
                      hint: isDe
                        ? 'Schalte oben am Status-Häkchen von „Entwurf" auf „Aktiv" — danach sehen es die berechtigten Gruppen.'
                        : 'Switch the status badge above from “Draft” to “Active” — it is then visible to the eligible groups.',
                      extra: (
                        <div className="dex-ui-stack">
                          <div className="dex-ui-callout dex-ui-callout--neutral dex-ui-callout--sm">
                            <span className="dex-ui-callout-icon"><Info size={14} /></span>
                            <div style={{ minWidth: 0 }}>
                              {everythingSame ? (
                                // Gesamt-Event und alle Sub-Sections gleich → eine Aussage.
                                <>{visSummary} {isDe ? `(Gesamt-Event und alle ${childVis.length} Sub-Section${childVis.length === 1 ? '' : 's'}.)` : `(Overall event and all ${childVis.length} sub-section${childVis.length === 1 ? '' : 's'}.)`}</>
                              ) : !hasChildren ? (
                                <>{visSummary}</>
                              ) : (
                                <>
                                  <strong>{isDe ? 'Gesamt-Event: ' : 'Overall event: '}</strong>{visSummary}
                                  <div style={{ marginTop: 4 }}>
                                    {allChildrenSame ? (
                                      <>
                                        <strong>
                                          {isDe ? `Für alle ${childVis.length} Sub-Sections gilt: ` : `For all ${childVis.length} sub-sections: `}
                                        </strong>
                                        {childVis[0].text}.
                                      </>
                                    ) : (
                                      <>
                                        {/* v28.79: Neun Zeilen Sichtbarkeit haben die Box
                                            erschlagen und die Warnung darunter verdeckt.
                                            Die Liste ist jetzt eingeklappt — was zaehlt
                                            (die Abweichung) steht sichtbar darunter.
                                            v31.3: als dex-ui-disclosure, damit der Pfeil
                                            sich dreht und der Knopf einen Hover hat. */}
                                        <button
                                          type="button"
                                          onClick={() => setVisListOpen(v => !v)}
                                          aria-expanded={visListOpen}
                                          className={cx('dex-ui-disclosure', visListOpen && 'is-open')}
                                          style={{ width: 'auto', margin: 0, padding: '2px 6px 2px 0', color: 'inherit', fontSize: '0.78rem' }}
                                        >
                                          <span className="dex-ui-disclosure-chevron"><ChevronDown size={13} /></span>
                                          {isDe
                                            ? `Sub-Sections einzeln (${childVis.length})`
                                            : `Sub-sections in detail (${childVis.length})`}
                                        </button>
                                        {visListOpen && childVis.map((c, ci) => (
                                          <div key={ci} style={{ paddingLeft: 8 }}>• <strong>{c.title}:</strong> {c.text}</div>
                                        ))}
                                      </>
                                    )}
                                  </div>
                                </>
                              )}
                            </div>
                          </div>
                          {/* v28.75: Die Liste ZEIGTE die Abweichung bisher nur —
                              benannt wurde sie nicht. Wer neun Zeilen
                              untereinander liest, übersieht, dass vier davon einen
                              Standortfilter tragen und fünf nicht. Also die
                              Varianten zählen und beim Namen nennen. */}
                          {varianceWarn}
                        </div>
                      ),
                    },
                    {
                      title: isDe ? 'Einladungsmail verschicken' : 'Send the invitation email',
                      hint: isDe
                        ? 'Optional: Einladung mit Anmelde-Link über DEX verschicken — an dich zum Weiterleiten oder an den Mailverteiler.'
                        : 'Optional: send the invitation with the registration link via DEX — to yourself for forwarding or to the mail distribution.',
                      // v31.3: Der Knopf steht in der Schritt-Zeile statt als
                      // unterstrichener Link mitten im Satz (Leitfaden 1.4).
                      action: (
                        <button
                          type="button"
                          className="btn btn-secondary dex-ui-btn-sm"
                          onClick={openInviteModal}
                        >
                          <Mail size={14} />
                          {isDe ? 'Einladungsmail öffnen' : 'Open invitation email'}
                        </button>
                      ),
                    },
                    {
                      title: isDe ? 'Anmeldungen verfolgen' : 'Track registrations',
                      hint: isDe
                        ? 'Sobald sich Teilnehmer anmelden, siehst du hier Anzahl, Status und die komplette Teilnehmerliste.'
                        : 'As soon as participants register, you see the count, the status and the full participant list here.',
                    },
                  ];
                  return steps.map((s, i) => (
                    <li key={i}>
                      {/* v31.3: Knopf und Hinweiskasten stehen linksbündig im
                          Schritt-Körper — rechts außen (Leitfaden 2a′) bliebe in
                          dieser schmalen Spalte kaum Platz für den Text. */}
                      <div className={cx('dex-ui-step', s.done && 'is-done')} style={{ alignItems: 'flex-start' }}>
                        <span className="dex-ui-step-num">{i + 1}</span>
                        <div className="dex-ui-step-body">
                          <div className="dex-ui-step-title">
                            {s.title}
                            {s.done && s.doneLabel && (
                              <span className="dex-ui-pill dex-ui-pill--green" style={{ marginLeft: 8 }}>{s.doneLabel}</span>
                            )}
                          </div>
                          <div className="dex-ui-step-hint">{s.hint}</div>
                          {s.action && <div style={{ marginTop: 8 }}>{s.action}</div>}
                          {s.extra && <div style={{ marginTop: 8 }}>{s.extra}</div>}
                        </div>
                      </div>
                    </li>
                  ));
                })()}
              </ol>
            </div>
          </aside>
  );
};

