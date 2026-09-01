/* LocationProgramStep — aus EventCreationPage.tsx ausgelagert (Zeilen 11155-11768 des
 * urspruenglichen Stands). Das JSX ist unveraendert uebernommen; einzige
 * Aenderung ist die Anzeige-Bedingung: aus `currentStep === 2` wurde das Prop `visible`.
 * `visible` schaltet display:none statt unmount — Eingaben ueberleben den
 * Schrittwechsel genauso wie vorher. */
import * as React from 'react';
import { SubEventDraft } from '../../wizard/wizardTypes';
import { AgendaItem } from '../../../types';
import { StepBadge } from '../../wizard/StepBadge';
import { buildOutlookLocation } from '../../../utils/eventFormat';
import { Plus, X } from '../../Icons';
import { InfoTooltip } from '../../InfoTooltip';
export interface LocationProgramStepProps {
  visible: boolean;
  activeLocationTabIdx: number;
  addAgendaItem: () => void;
  addrCity: string;
  addrHouseNo: string;
  addrStreet: string;
  addrZip: string;
  agenda: AgendaItem[];
  isDe: boolean;
  isMobile: boolean;
  isoToLocal: (iso: string) => string;
  location: string;
  locationOptions: string[];
  onlineMeetingMode: "none" | "own" | "auto";
  outlookLocationOverride: string;
  removeAgendaItem: (id: string) => void;
  renderStepIntro: (_bulletsDe: string[], _bulletsEn: string[]) => React.ReactElement | null;
  setAddrCity: React.Dispatch<React.SetStateAction<string>>;
  setAddrHouseNo: React.Dispatch<React.SetStateAction<string>>;
  setAddrStreet: React.Dispatch<React.SetStateAction<string>>;
  setAddrZip: React.Dispatch<React.SetStateAction<string>>;
  setLocation: React.Dispatch<React.SetStateAction<string>>;
  setOnlineMeetingMode: React.Dispatch<React.SetStateAction<"none" | "own" | "auto">>;
  setOutlookLocationOverride: React.Dispatch<React.SetStateAction<string>>;
  setSubEvents: React.Dispatch<React.SetStateAction<SubEventDraft[]>>;
  setTeamsLink: React.Dispatch<React.SetStateAction<string>>;
  setTransferTimes: React.Dispatch<React.SetStateAction<{ id: string; location: string; meetingPoint: string; address: string; date: string; departureTime: string; arrivalTime: string; description: string; }[]>>;
  startDate: string;
  subEvents: SubEventDraft[];
  t: (key: string) => string;
  teamsLink: string;
  transferTimes: { id: string; location: string; meetingPoint: string; address: string; date: string; departureTime: string; arrivalTime: string; description: string; }[];
  updateAgendaItem: (id: string, updates: Partial<AgendaItem>) => void;
}
export const LocationProgramStep: React.FC<LocationProgramStepProps> = (p) => {
  const { visible } = p;
  const { activeLocationTabIdx, addAgendaItem, addrCity, addrHouseNo, addrStreet, addrZip, agenda, isDe, isMobile, isoToLocal, location, locationOptions, onlineMeetingMode, outlookLocationOverride, removeAgendaItem, renderStepIntro, setAddrCity, setAddrHouseNo, setAddrStreet, setAddrZip, setLocation, setOnlineMeetingMode, setOutlookLocationOverride, setSubEvents, setTeamsLink, setTransferTimes, startDate, subEvents, t, teamsLink, transferTimes, updateAgendaItem } = p;
  return (
              <div style={{ display: visible ? 'block' : 'none' }}>
              <h2 className="dex-step-head-title">
                {isDe ? 'Schritt 3 — Ort & Programm' : 'Step 3 — Location & Programme'}
              </h2>
              <p className="dex-step-head-lead">
                {isDe
                  ? 'Hier sagst du, wo das Event stattfindet, wie der Tagesablauf aussieht und wie Teilnehmer hinkommen — alle Eingaben (Veranstaltungsort, Adresse, Agenda, Transferzeiten) sehen die Teilnehmer direkt auf der Anmelde-Seite und später unter „Meine Events".'
                  : 'Here you say where the event takes place, what the schedule looks like and how attendees get there — all inputs (venue, address, agenda, transfers) are shown to attendees directly on the registration page and later under "My Events".'}
              </p>
              {renderStepIntro(
                [
                  'Veranstaltungsort und Adresse erfassen',
                  'Agenda pflegen — Tagesablauf für die Teilnehmer',
                  'Optional: Transferzeiten (Bus/Bahn/Treffpunkt) hinterlegen',
                ],
                [
                  'Set event location and address',
                  'Maintain the agenda — schedule shown to participants',
                  'Optional: add transfer times (bus/train/meeting point)',
                ]
              )}
              {/* v22.38: Der frühere blaue Info-Banner ist in die
                  Schritt-Beschreibung (grüner Header-Lead) gewandert. */}

              {/* v15.3: pro-Sub-Event-Tabs für den Ort. Tab 0 = Haupt-Event
                  (komplette Ort/Adresse/Agenda/Transferzeiten-UI bleibt
                  unverändert). Tabs N>0 = vollwertige Ort/Adresse/Agenda/
                  Transferzeiten-UI pro Sub-Event — kein Inheritance-Toggle
                  mehr, jedes Sub-Event hat eigene Werte. Per
                  „Vom Hauptevent kopieren"-Button kann der Organizer die
                  Hauptevent-Werte als Startpunkt übernehmen. */}
              {/* v28.78: Der Scope-Umschalter steht jetzt global unter der
                  Schritt-Leiste (renderGlobalScopeBar) — nicht mehr je Schritt. */}

              {activeLocationTabIdx > 0 && (() => {
                const seIdx = activeLocationTabIdx - 1;
                const se = subEvents[seIdx];
                if (!se) return null;
                const seAddr = se.locationAddress || { street: '', houseNo: '', zip: '', city: '' };
                const seAgenda = se.agenda || [];
                const seTransfers = se.transferTimes || [];
                const updateSub = (patch: Partial<SubEventDraft>): void => {
                  setSubEvents(prev => prev.map((x, i) => i === seIdx ? { ...x, ...patch } : x));
                };
                const updateSubAgendaItem = (id: string, patch: Partial<AgendaItem>): void => {
                  updateSub({ agenda: seAgenda.map(a => a.id === id ? { ...a, ...patch } : a) });
                };
                return (
                  <div>
                    {/* v15.3: „Vom Hauptevent kopieren"-Button. Übernimmt
                        Ort, Adresse, Agenda und Transferzeiten vom Hauptevent
                        als Startwerte für dieses Sub-Event. */}
                    <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 12 }}>
                      <button
                        type="button"
                        className="btn btn-secondary"
                        style={{ fontSize: '0.8rem', padding: '6px 14px' }}
                        onClick={() => updateSub({
                          location: location,
                          locationAddress: { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity },
                          agenda: agenda.map(a => ({ ...a, id: `ag-${Date.now()}-${Math.random().toString(36).slice(2, 6)}` })),
                          transferTimes: transferTimes.map(tt => ({ ...tt, id: `tr-${Date.now()}-${Math.random().toString(36).slice(2, 6)}` })),
                        })}
                        title={isDe
                          ? 'Übernimmt Ort, Adresse, Agenda und Transferzeiten vom Hauptevent als Startwerte'
                          : 'Copies location, address, agenda and transfer times from the main event as starting values'}
                      >
                        {isDe ? 'Vom Hauptevent kopieren' : 'Copy from main event'}
                      </button>
                    </div>
                    <div className="form-group">
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <StepBadge n={14} />
                        {t('create.location')}
                      </label>
                      <input
                        className="form-input"
                        value={se.location || ''}
                        onChange={e => updateSub({ location: e.target.value })}
                        placeholder={isDe ? 'z.B. RheinEnergieStadion, Köln' : 'e.g. RheinEnergieStadion, Cologne'}
                      />
                    </div>
                    <div className="form-group">
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <StepBadge n={15} />
                        {isDe ? 'Adresse' : 'Address'}
                      </label>
                      <div style={{ display: 'grid', gridTemplateColumns: '3fr 1fr', gap: 8, marginBottom: 8 }}>
                        <input className="form-input" value={seAddr.street} onChange={e => updateSub({ locationAddress: { ...seAddr, street: e.target.value } })} placeholder="Straße" />
                        <input className="form-input" value={seAddr.houseNo} onChange={e => updateSub({ locationAddress: { ...seAddr, houseNo: e.target.value } })} placeholder="Hausnr." />
                      </div>
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 3fr', gap: 8 }}>
                        <input className="form-input" value={seAddr.zip} onChange={e => updateSub({ locationAddress: { ...seAddr, zip: e.target.value } })} placeholder="PLZ" />
                        <input className="form-input" value={seAddr.city} onChange={e => updateSub({ locationAddress: { ...seAddr, city: e.target.value } })} placeholder="Ort" />
                      </div>
                    </div>
                    {/* v18.44: Outlook-Ort pro Sub-Event überschreibbar (auch hier, nicht nur im Outlook-Editor). */}
                    <div className="form-group" style={{ marginTop: 16 }}>
                      <label className="form-label">
                        {isDe ? 'Ort im Outlook-Termin' : 'Location in the Outlook event'}
                      </label>
                      <input
                        className="form-input"
                        value={se.outlookLocation || ''}
                        onChange={e => updateSub({ outlookLocation: e.target.value })}
                        placeholder={buildOutlookLocation(se.location, seAddr) || (isDe ? 'z.B. Mezzomar, Harffstraße 110a, Düsseldorf' : 'e.g. Mezzomar, Harffstraße 110a, Düsseldorf')}
                      />
                      <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                        {isDe
                          ? 'Leer lassen = automatisch aus Veranstaltungsort + Adresse dieses Sub-Events.'
                          : 'Leave empty = automatic from this sub-event\'s venue + address.'}
                      </span>
                    </div>
                    <div className="form-group" style={{ marginTop: 24 }}>
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '1rem', fontWeight: 700 }}>
                        <StepBadge n={16} />
                        {t('create.agenda')}
                      </label>
                      {seAgenda
                        .slice()
                        .sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time))
                        .map(item => (
                        <div key={item.id} style={{
                          display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'flex-start',
                          padding: '10px 12px', marginBottom: 8,
                          background: 'var(--dex-gray-50, #fafafa)', borderRadius: 'var(--dex-radius)',
                          border: '1px solid var(--dex-gray-200)',
                        }}>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2, minWidth: 120, flexBasis: isMobile ? '100%' : undefined }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.date')}</label>
                            <input type="date" className="form-input" value={item.date} onChange={e => updateSubAgendaItem(item.id, { date: e.target.value })} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                          </div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2, minWidth: 80, flexBasis: isMobile ? '100%' : undefined }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.time')}</label>
                            <input type="time" className="form-input" value={item.time} onChange={e => updateSubAgendaItem(item.id, { time: e.target.value })} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                          </div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2, minWidth: 80, flexBasis: isMobile ? '100%' : undefined }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.endtime')}</label>
                            <input type="time" className="form-input" value={item.endTime || ''} onChange={e => updateSubAgendaItem(item.id, { endTime: e.target.value })} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                          </div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2, flex: 1, minWidth: 150, flexBasis: isMobile ? '100%' : undefined }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.title')}</label>
                            <input type="text" className="form-input" value={item.title} onChange={e => updateSubAgendaItem(item.id, { title: e.target.value })} placeholder={t('create.agenda.title')} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                          </div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 2, flex: 1, minWidth: 150, flexBasis: isMobile ? '100%' : undefined }}>
                            <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.desc')}</label>
                            <input type="text" className="form-input" value={item.description || ''} onChange={e => updateSubAgendaItem(item.id, { description: e.target.value })} placeholder={t('create.agenda.desc')} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                          </div>
                          <div style={{ display: 'flex', alignItems: 'flex-end', paddingBottom: 2 }}>
                            <button type="button" onClick={() => updateSub({ agenda: seAgenda.filter(a => a.id !== item.id) })} style={{
                              background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)',
                              fontSize: '1.1rem', padding: '4px', lineHeight: 1,
                            }} title={t('general.delete')}>
                              <X size={16} />
                            </button>
                          </div>
                        </div>
                      ))}
                      <button type="button" className="btn btn-outline" onClick={() => updateSub({
                        agenda: [...seAgenda, {
                          id: `ag-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
                          // v29.21 (Audit): Berliner Tag statt UTC-Tag —
                          // se.startDate ist UTC-ISO; slice(0,10) lieferte bei
                          // Startzeiten 00:00-01:59 Berlin den VORTAG.
                          date: se.startDate ? (isoToLocal(se.startDate) || '').slice(0, 10) : '',
                          time: '',
                          endTime: '',
                          icon: 'Calendar',
                          title: '',
                          description: '',
                        }],
                      })} style={{ fontSize: '0.85rem', padding: '6px 16px', marginTop: 4 }}>
                        <Plus size={14} /> {t('create.agenda.add')}
                      </button>
                    </div>
                    <div className="form-group" style={{ marginTop: 24 }}>
                      <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '1rem', fontWeight: 700 }}>
                        <StepBadge n={17} />
                        {t('create.transfers')}
                      </label>
                      {seTransfers.map(tt => (
                        <div key={tt.id} style={{
                          padding: '12px 14px', marginBottom: 8,
                          background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                          border: '1px solid var(--dex-gray-200)',
                        }}>
                          <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr 1fr auto', gap: 8, marginBottom: 8 }}>
                            <div>
                              <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.location')}</label>
                              <input type="text" className="form-input" value={tt.location} onChange={e => updateSub({ transferTimes: seTransfers.map(x => x.id === tt.id ? { ...x, location: e.target.value } : x) })} placeholder="Stadt..." style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                            </div>
                            <div>
                              <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.meetingpoint')}</label>
                              <input type="text" className="form-input" value={tt.meetingPoint || ''} onChange={e => updateSub({ transferTimes: seTransfers.map(x => x.id === tt.id ? { ...x, meetingPoint: e.target.value } : x) })} placeholder="z.B. Hbf..." style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                            </div>
                            <div>
                              <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.address')}</label>
                              <input type="text" className="form-input" value={tt.address || ''} onChange={e => updateSub({ transferTimes: seTransfers.map(x => x.id === tt.id ? { ...x, address: e.target.value } : x) })} placeholder="Straße, PLZ Ort" style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                            </div>
                            <div style={{ display: 'flex', alignItems: 'flex-end', paddingBottom: 4 }}>
                              <button type="button" onClick={() => updateSub({ transferTimes: seTransfers.filter(x => x.id !== tt.id) })} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)', padding: '4px', lineHeight: 1 }} title={t('general.delete')}>
                                <X size={16} />
                              </button>
                            </div>
                          </div>
                          <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr 1fr 2fr', gap: 8 }}>
                            <div>
                              <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.date')}</label>
                              <input type="date" className="form-input" value={tt.date} onChange={e => updateSub({ transferTimes: seTransfers.map(x => x.id === tt.id ? { ...x, date: e.target.value } : x) })} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                            </div>
                            <div>
                              <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.departure')}</label>
                              <input type="time" className="form-input" value={tt.departureTime} onChange={e => updateSub({ transferTimes: seTransfers.map(x => x.id === tt.id ? { ...x, departureTime: e.target.value } : x) })} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                            </div>
                            <div>
                              <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.arrival')}</label>
                              <input type="time" className="form-input" value={tt.arrivalTime} onChange={e => updateSub({ transferTimes: seTransfers.map(x => x.id === tt.id ? { ...x, arrivalTime: e.target.value } : x) })} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                            </div>
                            <div>
                              <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.desc')}</label>
                              <input type="text" className="form-input" value={tt.description || ''} onChange={e => updateSub({ transferTimes: seTransfers.map(x => x.id === tt.id ? { ...x, description: e.target.value } : x) })} placeholder={t('create.transfers.desc')} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                            </div>
                          </div>
                        </div>
                      ))}
                      <button type="button" className="btn btn-outline" onClick={() => updateSub({
                        transferTimes: [...seTransfers, {
                          id: `tr-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
                          location: '',
                          meetingPoint: '',
                          address: '',
                          // v29.21 (Audit): Berliner Tag statt UTC-Tag —
                          // se.startDate ist UTC-ISO; slice(0,10) lieferte bei
                          // Startzeiten 00:00-01:59 Berlin den VORTAG.
                          date: se.startDate ? (isoToLocal(se.startDate) || '').slice(0, 10) : '',
                          departureTime: '',
                          arrivalTime: '',
                          description: '',
                        }],
                      })} style={{ fontSize: '0.85rem', padding: '6px 16px', marginTop: 4 }}>
                        <Plus size={14} /> {t('create.transfers.add')}
                      </button>
                    </div>
                  </div>
                );
              })()}

              <div style={{ display: activeLocationTabIdx === 0 ? 'block' : 'none' }}>
              {/* v15.7: Step 3 NICHT mehr ausgrauen wenn subEventsOnlyMode —
                  Ort/Adresse/Agenda/Transferzeiten sind übergreifende
                  Event-Infos (gehen auf die Event-Detail-Seite, in „Meine
                  Events"-Card usw.). Sie bleiben immer relevant, auch wenn
                  das Hauptevent nicht anmeldbar ist. */}
              <div>
              <div className="form-group">
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={14} />
                  {t('create.location')}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> den <strong>Namen des Veranstaltungsortes</strong> (z.B. RheinEnergieStadion, Köln oder Deloitte Düsseldorf, Schwannstraße 6).<br /><br />
                      <strong>Anzeige in der App:</strong> erscheint auf der <strong>Anmelde-Seite</strong>, in der <strong>Eventliste</strong>, in <strong>Meine Events</strong> als Ort-Information.<br /><br />
                      <strong>Automatismen:</strong> wandert in den <strong>Outlook-Termin der Teilnehmer</strong> als Ort-Feld — so sehen sie auf einen Blick, wo sie hin müssen. Falls Bing Maps den Ort findet, blendet Outlook automatisch eine Karte ein.<br /><br />
                      <strong>Empfehlung:</strong> sprechender Name + Stadt — verwende die Adresse für die strukturierte Detail-Adresse darunter, hier reicht der Veranstaltungsort.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>name of the venue</strong> (e.g. RheinEnergieStadion, Cologne or Deloitte Düsseldorf, Schwannstraße 6).<br /><br />
                      <strong>Shown in the app:</strong> appears on the <strong>registration page</strong>, in the <strong>event list</strong> and in <strong>My Events</strong> as the location.<br /><br />
                      <strong>Automation:</strong> goes into the attendee Outlook event as the location field — so they immediately see where to go. If Bing Maps recognises it, Outlook auto-inserts a map.<br /><br />
                      <strong>Tip:</strong> descriptive name + city — use the structured address below for full details, here just the venue name.
                    </>
                  )} />
                </label>
                <input className="form-input" value={location} onChange={e => setLocation(e.target.value)} placeholder="z.B. RheinEnergieStadion, Köln" />
              </div>
              <div className="form-group">
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <StepBadge n={15} />
                  Adresse
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> die <strong>strukturierte Adresse</strong> (Straße, Hausnr., PLZ, Ort) — getrennt eingegeben, damit die Adresse einheitlich aussieht.<br /><br />
                      <strong>Anzeige in der App:</strong> wird auf der <strong>Anmelde-Seite</strong> und in <strong>Meine Events</strong> sauber formatiert angezeigt — Teilnehmer können auf die Adresse klicken und sie z.B. in Google Maps öffnen.<br /><br />
                      <strong>Automatismen:</strong> wird in <strong>Bestätigungs-Mails</strong> und im <strong>Outlook-Termin-Body</strong> als Klartext-Adresse mitgegeben — z.B. zum Kopieren ins Navi.<br /><br />
                      <strong>Optional:</strong> wenn der Veranstaltungsort oben schon eindeutig genug ist, kannst du die strukturierte Adresse leer lassen.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>structured address</strong> (street, number, ZIP, city) — entered field by field for consistent formatting.<br /><br />
                      <strong>Shown in the app:</strong> shown nicely on the <strong>registration page</strong> and under <strong>My Events</strong> — attendees can click it and e.g. open it in Google Maps.<br /><br />
                      <strong>Automation:</strong> goes into the <strong>confirmation mails</strong> and the <strong>Outlook event body</strong> as plain text — handy to copy into a navigation device.<br /><br />
                      <strong>Optional:</strong> if the venue name above is already explicit enough, you can leave the structured address empty.
                    </>
                  )} />
                </label>
                <div style={{ display: 'grid', gridTemplateColumns: '3fr 1fr', gap: 8, marginBottom: 8 }}>
                  <input className="form-input" value={addrStreet} onChange={e => setAddrStreet(e.target.value)} placeholder="Straße" />
                  <input className="form-input" value={addrHouseNo} onChange={e => setAddrHouseNo(e.target.value)} placeholder="Hausnr." />
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 3fr', gap: 8 }}>
                  <input className="form-input" value={addrZip} onChange={e => setAddrZip(e.target.value)} placeholder="PLZ" />
                  <input className="form-input" value={addrCity} onChange={e => setAddrCity(e.target.value)} placeholder="Ort" />
                </div>
              </div>

              {/* v18.40: Ort im Outlook-Termin (überschreibbar) */}
              <div className="form-group" style={{ marginTop: 16 }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  {isDe ? 'Ort im Outlook-Termin' : 'Location in the Outlook event'}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> den Text, der im <strong>&bdquo;Ort&ldquo;-Feld des Outlook-Termins</strong> der Teilnehmer steht.<br /><br />
                      <strong>Standard:</strong> wird automatisch aus <strong>Veranstaltungsort + Adresse</strong> oben zusammengebaut (siehe Platzhalter). Lässt du das Feld <strong>leer</strong>, wird immer dieser aktuelle Standard verwendet.<br /><br />
                      <strong>Überschreiben:</strong> Trägst du hier etwas ein, wird genau dieser Text als Termin-Ort genommen — z.&nbsp;B. ein abweichender Raum, ein Online-Link oder ein Kurzname.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the text shown in the <strong>&bdquo;Location&ldquo; field of attendees&apos; Outlook event</strong>.<br /><br />
                      <strong>Default:</strong> built automatically from <strong>venue + address</strong> above (see placeholder). Leave it <strong>empty</strong> to always use that current default.<br /><br />
                      <strong>Override:</strong> type something here to use exactly that text as the event location — e.g. a different room, an online link or a short name.
                    </>
                  )} />
                </label>
                <input
                  className="form-input"
                  value={outlookLocationOverride}
                  onChange={e => setOutlookLocationOverride(e.target.value)}
                  placeholder={buildOutlookLocation(location, { street: addrStreet, houseNo: addrHouseNo, zip: addrZip, city: addrCity }) || (isDe ? 'z.B. Mezzomar, Harffstraße 110a, Düsseldorf' : 'e.g. Mezzomar, Harffstraße 110a, Düsseldorf')}
                />
                <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                  {isDe
                    ? 'Leer lassen = automatisch aus Veranstaltungsort + Adresse. Eingabe überschreibt den Termin-Ort.'
                    : 'Leave empty = automatic from venue + address. Any input overrides the event location.'}
                </span>
              </div>

              {/* v30.26: Online-Meeting — Checkbox bei Ort, darunter die Wahl
                  zwischen eigenem Link und automatischer Teams-Besprechung.
                  Der Unterschied ist eine echte Entscheidung (Besprechungs-
                  optionen behalten oder Bequemlichkeit), deshalb steht die
                  Konsequenz direkt an der Auswahl und nicht im Tooltip. */}
              <div className="form-group" style={{ marginTop: 16, padding: '14px 16px', borderRadius: 12, border: '1px solid var(--dex-gray-200)', background: 'var(--dex-gray-50, #fafafa)' }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={onlineMeetingMode !== 'none'}
                    onChange={e => setOnlineMeetingMode(e.target.checked ? 'own' : 'none')}
                    style={{ marginTop: 3, cursor: 'pointer' }}
                  />
                  <span style={{ flex: 1 }}>
                    <strong>{isDe ? 'Online-Meeting (Microsoft Teams)' : 'Online meeting (Microsoft Teams)'}</strong>
                    <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                      {isDe
                        ? 'Für Events, an denen man per Teams teilnimmt — auch zusätzlich zu einem Präsenz-Ort (hybrid).'
                        : 'For events attended via Teams — also possible in addition to a physical location (hybrid).'}
                    </span>
                  </span>
                </label>
                {onlineMeetingMode !== 'none' && (
                  <div style={{ marginTop: 12, display: 'flex', flexDirection: 'column', gap: 10 }}>
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                      <input
                        type="radio"
                        name="dexOnlineMeetingMode"
                        checked={onlineMeetingMode === 'own'}
                        onChange={() => setOnlineMeetingMode('own')}
                        style={{ marginTop: 3, cursor: 'pointer' }}
                      />
                      <span style={{ flex: 1, fontSize: '0.88rem' }}>
                        <strong>{isDe ? 'Ich stelle den Teams-Link selbst' : 'I provide the Teams link myself'}</strong>
                        <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 3, lineHeight: 1.5 }}>
                          {isDe
                            ? <>Du legst die Besprechung wie gewohnt in Outlook oder Teams an und trägst den Link unten ein. <strong>Empfohlen</strong>, wenn du die Besprechungsoptionen brauchst: Lobby, Aufzeichnung, Referenten-Rollen — die kannst du nur an deiner eigenen Besprechung ändern.</>
                            : <>You create the meeting in Outlook or Teams and paste the link below. <strong>Recommended</strong> if you need the meeting options: lobby, recording, presenter roles — those can only be changed on your own meeting.</>}
                        </span>
                      </span>
                    </label>
                    <label style={{ display: 'flex', alignItems: 'flex-start', gap: 10, cursor: 'pointer' }}>
                      <input
                        type="radio"
                        name="dexOnlineMeetingMode"
                        checked={onlineMeetingMode === 'auto'}
                        onChange={() => setOnlineMeetingMode('auto')}
                        style={{ marginTop: 3, cursor: 'pointer' }}
                      />
                      <span style={{ flex: 1, fontSize: '0.88rem' }}>
                        <strong>{isDe ? 'DEX erzeugt den Teams-Link automatisch' : 'DEX creates the Teams link automatically'}</strong>
                        <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-600)', marginTop: 3, lineHeight: 1.5 }}>
                          {isDe
                            ? <>Der Termin wird als echte Teams-Besprechung angelegt — mit &bdquo;Teilnehmen&ldquo;-Knopf direkt im Kalender, ohne dass du etwas vorbereiten musst.<br /><strong style={{ color: 'var(--dex-orange-dark, #b35a00)' }}>Wichtig:</strong> Die Besprechung gehört dem Gruppenpostfach (no_reply.events). <strong>Du kannst danach keine Besprechungsoptionen mehr ändern</strong> — keine Lobby-Einstellung, keine Aufzeichnung, keine Referenten-Rollen.</>
                            : <>The event is created as a real Teams meeting — with a &bdquo;Join&ldquo; button right in the calendar, with nothing to prepare.<br /><strong style={{ color: 'var(--dex-orange-dark, #b35a00)' }}>Important:</strong> the meeting belongs to the group mailbox (no_reply.events). <strong>You cannot change any meeting options afterwards</strong> — no lobby settings, no recording, no presenter roles.</>}
                        </span>
                      </span>
                    </label>
                  </div>
                )}
              </div>

              {/* v29.39: Teams-Link. Steht hier bei Ort und Adresse, weil er
                  dieselbe Frage beantwortet („wo findet es statt?") — und
                  bewusst NUR hier: Ein zweites Feld im Outlook-Editor wären
                  zwei Bedienwege für denselben Wert.
                  v30.26: nur noch im Modus „eigener Link" sichtbar. */}
              {onlineMeetingMode === 'own' && (
              <div className="form-group" style={{ marginTop: 16 }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  {isDe ? 'Teams-Link (optional)' : 'Teams link (optional)'}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> den <strong>Teilnahme-Link deiner eigenen Teams-Besprechung</strong>. Lege die Besprechung wie gewohnt in Outlook oder Teams an und kopiere den Link hierher — DEX erzeugt selbst keine Teams-Meetings.<br /><br />
                      <strong>Anzeige in der App:</strong> im <strong>Outlook-Termin</strong> als Knopf &bdquo;An Microsoft-Teams-Besprechung teilnehmen&ldquo;, im <strong>Organizer Center</strong> und in <strong>Meine Events</strong> als Teilnahme-Knopf.<br /><br />
                      <strong>Wichtig:</strong> Der Link steht im <strong>Text</strong> des Termins. Outlook kennt den Termin dadurch <strong>nicht</strong> als Online-Besprechung — es gibt also keinen &bdquo;Teilnehmen&ldquo;-Knopf in der Kalenderleiste und keinen Direktaufruf aus Teams heraus. Die Teilnehmer klicken den Link im Termin bzw. in der App.<br /><br />
                      <strong>Gilt für:</strong> das ganze Event, also auch für die Termine der Sub-Events.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>join link of your own Teams meeting</strong>. Create the meeting in Outlook or Teams as usual and paste the link here — DEX does not create Teams meetings itself.<br /><br />
                      <strong>Shown in the app:</strong> in the <strong>Outlook event</strong> as a &bdquo;Join the Microsoft Teams meeting&ldquo; button, and in the <strong>Organizer Center</strong> and <strong>My Events</strong> as a join button.<br /><br />
                      <strong>Important:</strong> The link sits in the <strong>body</strong> of the event. Outlook therefore does <strong>not</strong> treat it as an online meeting — there is no &bdquo;Join&ldquo; button in the calendar bar and no direct join from Teams. Attendees click the link in the event or in the app.<br /><br />
                      <strong>Applies to:</strong> the whole event, including the sub-event calendar entries.
                    </>
                  )} />
                </label>
                <input
                  className="form-input"
                  value={teamsLink}
                  onChange={e => setTeamsLink(e.target.value)}
                  placeholder="https://teams.microsoft.com/l/meetup-join/..."
                />
                <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-gray-500)', marginTop: 4 }}>
                  {isDe
                    ? 'Der Link erscheint im Outlook-Termin, im Organizer Center und in „Meine Events". Hinweis: Der Termin ist damit für Outlook keine Online-Besprechung — es gibt keinen „Teilnehmen"-Knopf im Kalender und keinen Direktaufruf aus Teams, sondern den Link im Termin.'
                    : 'The link appears in the Outlook event, in the Organizer Center and in „My Events". Note: Outlook does not treat the event as an online meeting — there is no „Join" button in the calendar and no direct join from Teams, just the link inside the event.'}
                </span>
                {teamsLink.trim() && !/^https?:\/\//i.test(teamsLink.trim()) && (
                  <span style={{ display: 'block', fontSize: '0.78rem', color: 'var(--dex-red, #da291c)', fontWeight: 600, marginTop: 4 }}>
                    {isDe
                      ? 'Das sieht nicht nach einem Link aus — er muss mit https:// beginnen, sonst wird er nicht übernommen.'
                      : 'That does not look like a link — it must start with https://, otherwise it is ignored.'}
                  </span>
                )}
              </div>
              )}

              {/* ===== Agenda Editor ===== */}
              <div className="form-group" style={{ marginTop: 24 }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '1rem', fontWeight: 700 }}>
                  <StepBadge n={16} />
                  {t('create.agenda')}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> den <strong>Programmablauf des Events</strong> als Liste — pro Punkt: Datum, Start- und Endzeit, Titel, optionale Beschreibung und ein Icon (z.B. Kaffee, Vortrag, Pause).<br /><br />
                      <strong>Anzeige in der App:</strong> erscheint als <strong>schöner Timeline-Block</strong> auf der Anmelde-Seite und in Meine Events — Punkte werden automatisch nach Datum + Uhrzeit sortiert. Mehrtägige Events werden tageweise gruppiert.<br /><br />
                      <strong>Automatismen:</strong> die Agenda landet <strong>nicht</strong> automatisch im Outlook-Termin-Body (dafür gibt es das eigene Feld <strong>Text im Outlook-Termin</strong> in Schritt 6).<br /><br />
                      <strong>Empfehlung:</strong> hilft Teilnehmern, sich auf den Tag einzustellen — bei Tagungen oder Auswärtsterminen sehr empfohlen, bei kurzen Office-Events optional.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> the <strong>event programme</strong> as a list — per item: date, start/end time, title, optional description and an icon (e.g. coffee, talk, break).<br /><br />
                      <strong>Shown in the app:</strong> shown as a <strong>nice timeline block</strong> on the registration page and in My Events — items are auto-sorted by date + time. Multi-day events are grouped per day.<br /><br />
                      <strong>Automation:</strong> the agenda is <strong>not</strong> automatically pulled into the Outlook event body (there is a dedicated field <strong>Text in the Outlook event</strong> in step 7 for that).<br /><br />
                      <strong>Tip:</strong> helps attendees plan their day — strongly recommended for off-site events / conferences, optional for short office events.
                    </>
                  )} />
                </label>
                {agenda
                  .slice()
                  .sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time))
                  .map((item, agendaIdx) => (
                  <div key={item.id} style={{
                    display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'flex-start',
                    padding: '10px 12px', marginBottom: 8,
                    background: 'var(--dex-gray-50, #fafafa)', borderRadius: 'var(--dex-radius)',
                    border: '1px solid var(--dex-gray-200)',
                  }}>
                    {/* v22.36: Laufende Nummer statt Icon-Picker.
                        v22.38: kleiner (24px), vertikal mittig zur Zeile
                        (alignSelf center) und im Header-Grün (--dex-green). */}
                    <span style={{
                      alignSelf: 'center', flexShrink: 0,
                      width: 24, height: 24, display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                      borderRadius: '50%', background: 'var(--dex-green, #86bc25)', color: '#fff',
                      fontWeight: 700, fontSize: '0.78rem', lineHeight: 1,
                    }}>{agendaIdx + 1}</span>

                    {/* Date */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 2, minWidth: 120 }}>
                      <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.date')}</label>
                      <input type="date" className="form-input" value={item.date} onChange={e => updateAgendaItem(item.id, { date: e.target.value })} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                    </div>

                    {/* Start Time */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 2, minWidth: 80 }}>
                      <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.time')}</label>
                      <input type="time" className="form-input" value={item.time} onChange={e => updateAgendaItem(item.id, { time: e.target.value })} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                    </div>

                    {/* End Time */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 2, minWidth: 80 }}>
                      <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.endtime')}</label>
                      <input type="time" className="form-input" value={item.endTime || ''} onChange={e => updateAgendaItem(item.id, { endTime: e.target.value })} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                    </div>

                    {/* Title */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 2, flex: 1, minWidth: 150 }}>
                      <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.title')}</label>
                      <input type="text" className="form-input" value={item.title} onChange={e => updateAgendaItem(item.id, { title: e.target.value })} placeholder={t('create.agenda.title')} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                    </div>

                    {/* Description */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 2, flex: 1, minWidth: 150 }}>
                      <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.agenda.desc')}</label>
                      <input type="text" className="form-input" value={item.description || ''} onChange={e => updateAgendaItem(item.id, { description: e.target.value })} placeholder={t('create.agenda.desc')} style={{ padding: '4px 8px', fontSize: '0.85rem' }} />
                    </div>

                    {/* Delete */}
                    <div style={{ display: 'flex', alignItems: 'flex-end', paddingBottom: 2 }}>
                      <button type="button" onClick={() => removeAgendaItem(item.id)} style={{
                        background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)',
                        fontSize: '1.1rem', padding: '4px', lineHeight: 1,
                      }} title={t('general.delete')}>
                        <X size={16} />
                      </button>
                    </div>
                  </div>
                ))}
                <button type="button" className="btn btn-outline" onClick={addAgendaItem} style={{ fontSize: '0.85rem', padding: '6px 16px', marginTop: 4 }}>
                  <Plus size={14} /> {t('create.agenda.add')}
                </button>
              </div>

              {/* ===== Transferzeiten Editor ===== */}
              <div className="form-group" style={{ marginTop: 24 }}>
                <label className="form-label" style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '1rem', fontWeight: 700 }}>
                  <StepBadge n={17} />
                  {t('create.transfers')}
                  <InfoTooltip text={isDe ? (
                    <>
                      <strong>Was du hier einstellst:</strong> <strong>An- und Abreise-Infos</strong> für Teilnehmer — z.B. Bus-/Shuttle-/Bahn-Treffpunkte mit Datum, Abfahrt, Ankunft und optionaler Zusatzinfo (Bus-Kennzeichen, Treffpunkt-Schild, Wagen-Nr.). Pro Stadt ein eigener Eintrag möglich.<br /><br />
                      <strong>Anzeige in der App:</strong> erscheint als <strong>eigener Block</strong> auf der Anmelde-Seite und in Meine Events mit allen Details auf einen Blick.<br /><br />
                      <strong>Automatismen:</strong> Transferzeiten gehen <strong>nicht</strong> in den Outlook-Termin (sonst würde der Termin Bus als Konkurrenz-Termin im Kalender blocken). Sie sind nur in der App sichtbar.<br /><br />
                      <strong>Empfehlung:</strong> bei Auswärtsterminen mit organisierter Anreise sehr empfohlen — bei rein lokalen Office-Events nicht nötig.
                    </>
                  ) : (
                    <>
                      <strong>What you set here:</strong> <strong>arrival and departure info</strong> for attendees — e.g. bus/shuttle/train pickups with date, departure, arrival and an optional note (bus number, meeting-point sign, carriage no.). One entry per city.<br /><br />
                      <strong>Shown in the app:</strong> shown as a <strong>dedicated block</strong> on the registration page and in My Events with all details at a glance.<br /><br />
                      <strong>Automation:</strong> transfer times do <strong>not</strong> end up in the Outlook event (otherwise the bus trip would clash with the actual event in the calendar). They live only in the app.<br /><br />
                      <strong>Tip:</strong> strongly recommended for off-site events with organised travel — not needed for local office events.
                    </>
                  )} />
                </label>
                {transferTimes.map((tt) => (
                  <div key={tt.id} style={{
                    padding: '12px 14px', marginBottom: 8,
                    background: 'var(--dex-gray-50, #fafafa)', borderRadius: 12,
                    border: '1px solid var(--dex-gray-200)',
                  }}>
                    {/* Zeile 1: Stadt + Treffpunkt + Adresse + Löschen */}
                    <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr 1fr auto', gap: 8, marginBottom: 8 }}>
                      <div>
                        <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.location')}</label>
                        <input type="text" className="form-input" list={`transfer-locations-${tt.id}`} value={tt.location} onChange={e => setTransferTimes(transferTimes.map(x => x.id === tt.id ? { ...x, location: e.target.value } : x))} placeholder="Stadt eingeben..." style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                        <datalist id={`transfer-locations-${tt.id}`}>
                          {locationOptions.filter(o => o !== 'All').map(opt => (
                            <option key={opt} value={opt} />
                          ))}
                        </datalist>
                      </div>
                      <div>
                        <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.meetingpoint')}</label>
                        <input type="text" className="form-input" value={tt.meetingPoint || ''} onChange={e => setTransferTimes(transferTimes.map(x => x.id === tt.id ? { ...x, meetingPoint: e.target.value } : x))} placeholder="z.B. Flughafen, Hbf..." style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                      </div>
                      <div>
                        <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.address')}</label>
                        <input type="text" className="form-input" value={tt.address || ''} onChange={e => setTransferTimes(transferTimes.map(x => x.id === tt.id ? { ...x, address: e.target.value } : x))} placeholder="Straße, PLZ Ort" style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                      </div>
                      <div style={{ display: 'flex', alignItems: 'flex-end', paddingBottom: 4 }}>
                        <button type="button" onClick={() => setTransferTimes(transferTimes.filter(x => x.id !== tt.id))} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--dex-red, #c00)', padding: '4px', lineHeight: 1 }} title={t('general.delete')}>
                          <X size={16} />
                        </button>
                      </div>
                    </div>
                    {/* Zeile 2: Datum + Abfahrt + Ankunft + Beschreibung */}
                    <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr 1fr 2fr', gap: 8 }}>
                      <div>
                        <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.date')}</label>
                        <input type="date" className="form-input" value={tt.date} onChange={e => setTransferTimes(transferTimes.map(x => x.id === tt.id ? { ...x, date: e.target.value } : x))} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                      </div>
                      <div>
                        <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.departure')}</label>
                        <input type="time" className="form-input" value={tt.departureTime} onChange={e => setTransferTimes(transferTimes.map(x => x.id === tt.id ? { ...x, departureTime: e.target.value } : x))} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                      </div>
                      <div>
                        <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.arrival')}</label>
                        <input type="time" className="form-input" value={tt.arrivalTime} onChange={e => setTransferTimes(transferTimes.map(x => x.id === tt.id ? { ...x, arrivalTime: e.target.value } : x))} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                      </div>
                      <div>
                        <label style={{ fontSize: '0.7rem', color: 'var(--dex-gray-500)' }}>{t('create.transfers.desc')}</label>
                        <input type="text" className="form-input" value={tt.description || ''} onChange={e => setTransferTimes(transferTimes.map(x => x.id === tt.id ? { ...x, description: e.target.value } : x))} placeholder={t('create.transfers.desc')} style={{ padding: '6px 8px', fontSize: '0.85rem' }} />
                      </div>
                    </div>
                  </div>
                ))}
                <button type="button" className="btn btn-outline" onClick={() => setTransferTimes([...transferTimes, { id: `tr-${Date.now()}`, location: '', meetingPoint: '', address: '', date: startDate ? startDate.slice(0, 10) : '', departureTime: '', arrivalTime: '', description: '' }])} style={{ fontSize: '0.85rem', padding: '6px 16px', marginTop: 4 }}>
                  <Plus size={14} /> {t('create.transfers.add')}
                </button>
              </div>

              </div>{/* v15.6: close hauptGreyoutWrapperStyle div (Step 3) */}
              </div>{/* v15.0: close activeLocationTabIdx===0 wrapper (Top-Level Ort/Adresse/Agenda/Transfer) */}

              </div>
  );
};
