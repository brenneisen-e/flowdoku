/* NextStepsBox — 1:1 aus AdminPage.tsx ausgelagert (Zeilen 7632-7875 des
 * Stands vor dem Schnitt). Der Inhalt ist zeichengleich uebernommen; die
 * Anzeige-Bedingung bleibt beim Aufrufer.
 */
import * as React from 'react';
import { Info } from '../../Icons';
import { shortSubEventTitle } from '../../../utils/subEventTitle';
import { DeloitteEvent } from '../../../types';

export interface NextStepsBoxProps {
  childEventsOf: (parentEventId: string) => DeloitteEvent[];
  isAdmin: boolean;
  isDe: boolean;
  isOrganizerFor: (ev: DeloitteEvent) => boolean;
  openInviteModal: () => void;
  selectedEvent: DeloitteEvent;
  setVisListOpen: React.Dispatch<React.SetStateAction<boolean>>;
  visListOpen: boolean;
}

export const NextStepsBox: React.FC<NextStepsBoxProps> = (p) => {
  const { childEventsOf, isAdmin, isDe, isOrganizerFor, openInviteModal, selectedEvent, setVisListOpen, visListOpen } = p;
  return (
          <aside style={{ flex: '1 1 360px', minWidth: 320 }}>
            <div className="card" style={{ padding: 20, background: 'rgba(134,188,37,0.05)', border: '1px solid var(--dex-green, #86bc25)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
                <span style={{ color: 'var(--dex-green-dark, #4a7c1f)', display: 'inline-flex' }}><Info size={18} /></span>
                <h3 style={{ margin: 0, fontSize: '1rem', color: 'var(--dex-green-dark, #4a7c1f)' }}>{isDe ? 'Nächste Schritte' : 'Next steps'}</h3>
              </div>
              <p style={{ margin: '0 0 14px', fontSize: '0.8rem', color: 'var(--dex-gray-600)', lineHeight: 1.5 }}>
                {isDe ? 'Dein Event ist angelegt — so machst du es startklar:' : 'Your event is created — here is how to get it ready:'}
              </p>
              <ol style={{ margin: 0, padding: 0, listStyle: 'none', display: 'flex', flexDirection: 'column', gap: 13 }}>
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
                  const steps: Array<{ title: string; body: React.ReactNode }> = [
                    {
                      title: isDe ? 'Event finalisieren' : 'Finalize the event',
                      body: (
                        <>
                          {isDe
                            ? 'Über „Event bearbeiten" Felder, Bild und Texte vervollständigen.'
                            : 'Use “Edit event” to complete fields, image and texts.'}
                          {missingBits.length > 0 && (
                            <span style={{ display: 'block', marginTop: 5, padding: '6px 8px', borderRadius: 6, background: '#fff4e5', border: '1px solid #ed8b00', color: '#8a4b00', fontSize: '0.74rem', lineHeight: 1.45 }}>
                              {isDe
                                ? <>Fehlt noch: <strong>{missingBits.join(', ')}</strong>.</>
                                : <>Still missing: <strong>{missingBits.join(', ')}</strong>.</>}
                              {/* v28.79: Beim Ort die Folge benennen — ohne ihn
                                  steht im Outlook-Termin und in den Mails nur
                                  das Wort „Veranstaltungsort" als Platzhalter. */}
                              {locationMissing && (
                                <span style={{ display: 'block', marginTop: 4 }}>
                                  {isDe
                                    ? <>Ohne Ort erscheint im Outlook-Termin und in den Mails nur das Wort &bdquo;Veranstaltungsort&ldquo; &mdash; die Teilnehmer wissen dann nicht, wohin sie kommen sollen.</>
                                    : <>Without a location the Outlook invite and the emails only show the word &bdquo;venue&ldquo; &mdash; attendees will not know where to go.</>}
                                </span>
                              )}
                            </span>
                          )}
                        </>
                      ),
                    },
                    {
                      title: isDe ? 'Test-An- und Abmeldung' : 'Test registration & cancellation',
                      body: isDe
                        ? 'Melde dich einmal selbst an und wieder ab, um zu prüfen, ob die automatische Kommunikation (Bestätigungs-Mail, Outlook-Termin, Abmelde-Mail) richtig ankommt.'
                        : 'Register and cancel yourself once to check that the automatic communication (confirmation email, Outlook invite, cancellation email) works correctly.',
                    },
                    {
                      title: isDe ? 'Event live schalten' : 'Publish the event',
                      body: (
                        <>
                          {isDe
                            ? 'Oben über das Status-Häkchen „Entwurf → Aktiv" schalten. Danach ist es für die berechtigten Gruppen sichtbar.'
                            : 'Switch the status badge above from “Draft → Active”. It is then visible to the eligible groups.'}
                          <span style={{ display: 'block', marginTop: 5, padding: '6px 8px', borderRadius: 6, background: '#fff', border: '1px solid var(--dex-gray-200)', color: 'var(--dex-gray-600)', fontSize: '0.74rem', lineHeight: 1.45 }}>
                            {everythingSame ? (
                              // Gesamt-Event und alle Sub-Sections gleich → eine Aussage.
                              <>{visSummary} {isDe ? `(Gesamt-Event und alle ${childVis.length} Sub-Section${childVis.length === 1 ? '' : 's'}.)` : `(Overall event and all ${childVis.length} sub-section${childVis.length === 1 ? '' : 's'}.)`}</>
                            ) : !hasChildren ? (
                              <>{visSummary}</>
                            ) : (
                              <>
                                <strong style={{ color: 'var(--dex-gray-700)' }}>{isDe ? 'Gesamt-Event: ' : 'Overall event: '}</strong>{visSummary}
                                <span style={{ display: 'block', marginTop: 4 }}>
                                  {allChildrenSame ? (
                                    <>
                                      <strong style={{ color: 'var(--dex-gray-700)' }}>
                                        {isDe ? `Für alle ${childVis.length} Sub-Sections gilt: ` : `For all ${childVis.length} sub-sections: `}
                                      </strong>
                                      {childVis[0].text}.
                                    </>
                                  ) : (
                                    <>
                                      {/* v28.79: Neun Zeilen Sichtbarkeit haben die Box
                                          erschlagen und die Warnung darunter verdeckt.
                                          Die Liste ist jetzt eingeklappt — was zaehlt
                                          (die Abweichung) steht sichtbar darunter. */}
                                      <button
                                        type="button"
                                        onClick={() => setVisListOpen(v => !v)}
                                        style={{
                                          background: 'none', border: 'none', padding: 0, cursor: 'pointer',
                                          font: 'inherit', color: 'var(--dex-gray-700)', fontWeight: 700,
                                          display: 'inline-flex', alignItems: 'center', gap: 5,
                                        }}
                                      >
                                        <span style={{ fontSize: '0.7rem' }}>{visListOpen ? '▾' : '▸'}</span>
                                        {isDe
                                          ? `Sub-Sections einzeln (${childVis.length})`
                                          : `Sub-sections in detail (${childVis.length})`}
                                      </button>
                                      {visListOpen && childVis.map((c, ci) => (
                                        <span key={ci} style={{ display: 'block', paddingLeft: 8 }}>• <strong>{c.title}:</strong> {c.text}</span>
                                      ))}
                                      {/* v28.75: Die Liste ZEIGTE die Abweichung bisher nur —
                                          benannt wurde sie nicht. Wer neun Zeilen
                                          untereinander liest, übersieht, dass vier davon einen
                                          Standortfilter tragen und fünf nicht. Also die
                                          Varianten zählen und beim Namen nennen. */}
                                      {(() => {
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
                                          <span style={{
                                            display: 'block', marginTop: 8, padding: '7px 9px', borderRadius: 6,
                                            background: '#fff8e6', border: '1px solid #e0b34d', color: '#7a5a12',
                                            fontSize: '0.74rem', lineHeight: 1.5,
                                          }}>
                                            {isDe
                                              ? <>
                                                  <strong>Achtung — die Sub-Sections sind unterschiedlich sichtbar</strong> ({groups.size} Varianten).{' '}
                                                  {sorted[0][1].length} {sorted[0][1].length === 1 ? 'Sub-Section hat' : 'Sub-Sections haben'} „{sorted[0][0]}“, {minCount} {minCount === 1 ? 'weicht ab' : 'weichen ab'}: {minority.map(e => e[1].join(', ')).join('; ')}.{' '}
                                                  Wenn das nicht gewollt ist: in &bdquo;Event bearbeiten&ldquo; den Schritt <strong>Kapazität &amp; Sichtbarkeit</strong> öffnen, die passende Sub-Section wählen und dort auf <strong>&bdquo;Einstellungen auf andere übertragen&ldquo;</strong> klicken.
                                                </>
                                              : <>
                                                  <strong>Careful — the sub-sections have different visibility</strong> ({groups.size} variants).{' '}
                                                  {sorted[0][1].length} of them use „{sorted[0][0]}“, {minCount} differ: {minority.map(e => e[1].join(', ')).join('; ')}.{' '}
                                                  If that is not intended: open „Edit event“ → step <strong>Capacity &amp; visibility</strong>, pick the right sub-section and use <strong>„Transfer settings to others“</strong>.
                                                </>}
                                          </span>
                                        );
                                      })()}
                                    </>
                                  )}
                                </span>
                              </>
                            )}
                          </span>
                        </>
                      ),
                    },
                    {
                      title: isDe ? 'Einladungsmail verschicken' : 'Send the invitation email',
                      body: (
                        <>
                          {isDe
                            ? 'Optional kannst du die Einladung mit Anmelde-Link direkt über DEX verschicken — an dich zum Weiterleiten oder an den Mailverteiler.'
                            : 'Optionally send the invitation with the registration link directly via DEX — to yourself for forwarding or to the mail distribution.'}
                          {' '}
                          <button
                            type="button"
                            onClick={openInviteModal}
                            style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', color: 'var(--dex-green-dark, #4a7c1f)', fontWeight: 700, fontSize: '0.78rem', textDecoration: 'underline' }}
                          >
                            {isDe ? 'Einladungsmail öffnen' : 'Open invitation email'}
                          </button>
                        </>
                      ),
                    },
                    {
                      title: isDe ? 'Anmeldungen verfolgen' : 'Track registrations',
                      body: isDe
                        ? 'Sobald sich Teilnehmer anmelden, siehst du hier im Admin-Panel alle Infos — Anzahl, Status und die komplette Teilnehmerliste.'
                        : 'As soon as participants register, you see everything here in the admin panel — count, status and the full participant list.',
                    },
                  ];
                  return steps.map((s, i) => (
                    <li key={i} style={{ display: 'flex', gap: 10, alignItems: 'flex-start' }}>
                      <span style={{ flexShrink: 0, width: 24, height: 24, borderRadius: '50%', background: 'var(--dex-green, #86bc25)', color: '#fff', fontWeight: 700, fontSize: '0.82rem', display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}>{i + 1}</span>
                      <span style={{ fontSize: '0.82rem', lineHeight: 1.5 }}>
                        <strong style={{ color: 'var(--dex-gray-800)' }}>{s.title}.</strong>{' '}
                        <span style={{ color: 'var(--dex-gray-600)' }}>{s.body}</span>
                      </span>
                    </li>
                  ));
                })()}
              </ol>
            </div>
          </aside>
  );
};

