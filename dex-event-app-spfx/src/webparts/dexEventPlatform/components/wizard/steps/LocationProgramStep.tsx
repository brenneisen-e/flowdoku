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
import { AlertCircle, Check, ChevronDown, Plus, X } from '../../Icons';
import { InfoTooltip } from '../../InfoTooltip';
import { AgendaEditor } from '../AgendaEditor';
import { cx } from '../../dexUi';
export interface LocationProgramStepProps {
  visible: boolean;
  activeLocationTabIdx: number;
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
  /** v30.86: Programmpunkte mit Check-in — die Agenda ist dann die Check-in-Liste. */
  agendaCheckIn: boolean;
  agendaTermPlural: string;
  agendaTermSingular: string;
  setAgenda: React.Dispatch<React.SetStateAction<AgendaItem[]>>;
}

type TransferTime = LocationProgramStepProps['transferTimes'][number];

/* v31.2: Adresse als vier beschriftete Felder — vorher standen die
   Bezeichnungen nur im Platzhalter und verschwanden beim Tippen; und sie
   waren auch auf Englisch deutsch. Einmal gebaut, zweimal genutzt
   (Hauptevent und Sub-Event), damit beide Zweige gleich aussehen. */
const AddressFields: React.FC<{
  isDe: boolean; street: string; houseNo: string; zip: string; city: string;
  onStreet: (_v: string) => void; onHouseNo: (_v: string) => void; onZip: (_v: string) => void; onCity: (_v: string) => void;
}> = ({ isDe, street, houseNo, zip, city, onStreet, onHouseNo, onZip, onCity }) => {
  const field = (label: string, value: string, onChange: (v: string) => void, placeholder: string): React.ReactElement => (
    <div>
      <span className="dex-ui-muted" style={{ display: 'block', fontSize: '0.72rem', marginBottom: 3 }}>{label}</span>
      <input className="form-input" value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder} aria-label={label} />
    </div>
  );
  return (
    <div className="dex-ui-stack" style={{ gap: 8 }}>
      <div style={{ display: 'grid', gridTemplateColumns: '3fr 1fr', gap: 8 }}>
        {field(isDe ? 'Straße' : 'Street', street, onStreet, isDe ? 'z.B. Schwannstraße' : 'e.g. Schwannstraße')}
        {field(isDe ? 'Hausnr.' : 'No.', houseNo, onHouseNo, '6')}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 3fr', gap: 8 }}>
        {field(isDe ? 'PLZ' : 'ZIP', zip, onZip, '40476')}
        {field(isDe ? 'Ort' : 'City', city, onCity, isDe ? 'Düsseldorf' : 'Düsseldorf')}
      </div>
    </div>
  );
};

/* v31.2: Transferzeiten — eine Karte je Eintrag, Löschen als runder Symbol-
   Knopf mit Hover, leerer Zustand statt leerer Fläche. Ein Helfer für beide
   Zweige (Hauptevent schreibt über setTransferTimes, Sub-Event über
   updateSub) — vorher stand derselbe 45-Zeilen-Block zweimal in der Datei,
   mit leicht abweichenden Platzhaltern. Die nativen Datums-/Zeitfelder
   bleiben (kein neues, aber auch kein Umbau — s. Bericht). */
const TransfersSection: React.FC<{
  items: TransferTime[]; onChange: (_next: TransferTime[]) => void;
  isDe: boolean; isMobile: boolean; t: (_key: string) => string;
  defaultDate: string; locationOptions?: string[]; tooltip?: React.ReactNode;
}> = ({ items, onChange, isDe, isMobile, t, defaultDate, locationOptions, tooltip }) => {
  const patch = (id: string, pt: Partial<TransferTime>): void => onChange(items.map(x => x.id === id ? { ...x, ...pt } : x));
  const small = (label: string, input: React.ReactElement): React.ReactElement => (
    <div>
      <label className="dex-ui-muted" style={{ display: 'block', fontSize: '0.72rem', marginBottom: 3 }}>{label}</label>
      {input}
    </div>
  );
  const inStyle: React.CSSProperties = { padding: '6px 8px', fontSize: '0.85rem' };
  return (
    <div className="dex-ui-field">
      <label className="dex-ui-label" style={{ fontSize: '0.95rem' }}>
        <StepBadge n={17} />
        {isDe ? 'Gibt es organisierte Anreisen (Bus, Shuttle, Bahn)?' : 'Is there organised travel (bus, shuttle, train)?'}
        <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
        {tooltip && <InfoTooltip text={tooltip} />}
      </label>
      <div className="dex-ui-help" style={{ margin: '-2px 0 10px' }}>
        {isDe
          ? 'Je Stadt ein Eintrag mit Treffpunkt, Abfahrt und Ankunft. Teilnehmer sehen die Transfers auf der Anmeldeseite und in „Meine Events" — nicht im Outlook-Termin.'
          : 'One entry per city with meeting point, departure and arrival. Attendees see transfers on the registration page and in “My Events” — not in the Outlook event.'}
      </div>
      {items.length === 0 && (
        <div className="dex-ui-empty" style={{ padding: '18px 16px', marginBottom: 10 }}>
          <div className="dex-ui-empty-title">{isDe ? 'Noch keine Transferzeiten' : 'No transfer times yet'}</div>
          {isDe ? 'Bei rein lokalen Office-Events kannst du das leer lassen.' : 'For purely local office events you can leave this empty.'}
        </div>
      )}
      {items.map(tt => (
        <div key={tt.id} className="dex-ui-card dex-ui-card--soft" style={{ padding: '12px 14px', marginBottom: 8 }}>
          {/* Zeile 1: Stadt + Treffpunkt + Adresse + Löschen */}
          <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr 1fr auto', gap: 8, marginBottom: 8 }}>
            {small(t('create.transfers.location'), (
              <>
                <input type="text" className="form-input" list={locationOptions ? `transfer-locations-${tt.id}` : undefined} value={tt.location} onChange={e => patch(tt.id, { location: e.target.value })} placeholder={isDe ? 'Stadt, z.B. Köln' : 'City, e.g. Cologne'} style={inStyle} />
                {locationOptions && (
                  <datalist id={`transfer-locations-${tt.id}`}>
                    {locationOptions.filter(o => o !== 'All').map(opt => (
                      <option key={opt} value={opt} />
                    ))}
                  </datalist>
                )}
              </>
            ))}
            {small(t('create.transfers.meetingpoint'), <input type="text" className="form-input" value={tt.meetingPoint || ''} onChange={e => patch(tt.id, { meetingPoint: e.target.value })} placeholder={isDe ? 'z.B. Flughafen, Hbf Gleis 4' : 'e.g. airport, main station platform 4'} style={inStyle} />)}
            {small(t('create.transfers.address'), <input type="text" className="form-input" value={tt.address || ''} onChange={e => patch(tt.id, { address: e.target.value })} placeholder={isDe ? 'Straße, PLZ Ort' : 'Street, ZIP City'} style={inStyle} />)}
            <div style={{ display: 'flex', alignItems: 'flex-end' }}>
              <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" onClick={() => onChange(items.filter(x => x.id !== tt.id))} title={t('general.delete')} aria-label={t('general.delete')}>
                <X size={16} />
              </button>
            </div>
          </div>
          {/* Zeile 2: Datum + Abfahrt + Ankunft + Beschreibung */}
          <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr 1fr 2fr', gap: 8 }}>
            {small(t('create.transfers.date'), <input type="date" className="form-input" value={tt.date} onChange={e => patch(tt.id, { date: e.target.value })} style={inStyle} />)}
            {small(t('create.transfers.departure'), <input type="time" className="form-input" value={tt.departureTime} onChange={e => patch(tt.id, { departureTime: e.target.value })} style={inStyle} />)}
            {small(t('create.transfers.arrival'), <input type="time" className="form-input" value={tt.arrivalTime} onChange={e => patch(tt.id, { arrivalTime: e.target.value })} style={inStyle} />)}
            {small(t('create.transfers.desc'), <input type="text" className="form-input" value={tt.description || ''} onChange={e => patch(tt.id, { description: e.target.value })} placeholder={isDe ? 'z.B. Bus-Kennzeichen, Schild am Treffpunkt' : 'e.g. bus number, sign at the meeting point'} style={inStyle} />)}
          </div>
        </div>
      ))}
      <button type="button" className="btn btn-outline dex-ui-btn-sm" onClick={() => onChange([...items, {
        id: `tr-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        location: '', meetingPoint: '', address: '', date: defaultDate, departureTime: '', arrivalTime: '', description: '',
      }])}>
        <Plus size={14} /> {t('create.transfers.add')}
      </button>
    </div>
  );
};

export const LocationProgramStep: React.FC<LocationProgramStepProps> = (p) => {
  const { visible } = p;
  const { agendaCheckIn, agendaTermPlural, agendaTermSingular, setAgenda } = p;
  // v31.2: Aufklapper „Weitere Einstellungen zum Ort" (Outlook-Ort) je Zweig —
  // reiner Anzeige-State, standardmäßig zu (Leitfaden: Seltenes ist zu).
  const [moreOpenMain, setMoreOpenMain] = React.useState(false);
  const [moreOpenSub, setMoreOpenSub] = React.useState(false);
  // v30.94: „Letzten Tag kopieren" (v30.86) ist in den AgendaEditor gewandert
  // (je Cluster „Kopieren", Name zählt hoch).
  const agendaPlural = agendaTermPlural.trim() || (p.isDe ? 'Programmpunkte' : 'Agenda items');
  const { activeLocationTabIdx, addrCity, addrHouseNo, addrStreet, addrZip, agenda, isDe, isMobile, isoToLocal, location, locationOptions, onlineMeetingMode, outlookLocationOverride, renderStepIntro, setAddrCity, setAddrHouseNo, setAddrStreet, setAddrZip, setLocation, setOnlineMeetingMode, setOutlookLocationOverride, setSubEvents, setTeamsLink, setTransferTimes, startDate, subEvents, t, teamsLink, transferTimes } = p;
  return (
              <div style={{ display: visible ? 'block' : 'none' }}>
              <h2 className="dex-step-head-title">
                <span className="dex-step-eyebrow">{isDe ? 'Schritt 3 von 9' : 'Step 3 of 9'}</span>
                {isDe ? 'Ort & Programm' : 'Location & Programme'}
              </h2>
              <p className="dex-step-head-lead">
                {isDe
                  ? <>Wo findet das Event statt, wie kommen die Teilnehmer hin, was passiert wann? Alles hier sehen sie auf der <strong>Anmeldeseite</strong> und in <strong>&bdquo;Meine Events&ldquo;</strong>; Ort und Adresse landen auch im Outlook-Termin.</>
                  : <>Where does the event take place, how do attendees get there, what happens when? Everything here is shown on the <strong>registration page</strong> and in <strong>&bdquo;My Events&ldquo;</strong>; venue and address also go into the Outlook event.</>}
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
                // v29.21 (Audit): Berliner Tag statt UTC-Tag als Vorbelegung —
                // se.startDate ist UTC-ISO; slice(0,10) lieferte bei
                // Startzeiten 00:00-01:59 Berlin den VORTAG.
                const seDefaultDate = se.startDate ? (isoToLocal(se.startDate) || '').slice(0, 10) : '';
                return (
                  <div>
                    {/* v15.3: „Vom Hauptevent kopieren"-Button. Übernimmt
                        Ort, Adresse, Agenda und Transferzeiten vom Hauptevent
                        als Startwerte für dieses Sub-Event.
                        v31.2: als Hinweiszeile mit dem Knopf daran — vorher
                        stand er allein rechts, ohne zu sagen, wofür. */}
                    <div className="dex-ui-callout dex-ui-callout--neutral" style={{ alignItems: 'center', marginBottom: 16 }}>
                      <span style={{ flex: 1 }}>
                        {isDe
                          ? 'Dieser Termin hat eigene Angaben zu Ort, Anreise und Programm. Du kannst die Werte des Hauptevents als Startpunkt übernehmen.'
                          : 'This date has its own venue, travel and programme details. You can take the main event’s values as a starting point.'}
                      </span>
                      <button
                        type="button"
                        className="btn btn-secondary dex-ui-btn-sm"
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
                    <div className="dex-ui-section">
                      <div className="dex-ui-section-title">{isDe ? 'Wo findet dieser Termin statt?' : 'Where does this date take place?'}</div>
                      <div className="dex-ui-field">
                        <label className="dex-ui-label">
                          <StepBadge n={14} />
                          {isDe ? 'Wie heißt der Veranstaltungsort?' : 'What is the venue called?'}
                        </label>
                        <input
                          className="form-input"
                          value={se.location || ''}
                          onChange={e => updateSub({ location: e.target.value })}
                          placeholder={isDe ? 'z.B. RheinEnergieStadion, Köln' : 'e.g. RheinEnergieStadion, Cologne'}
                        />
                        <div className="dex-ui-help">
                          {isDe
                            ? 'Erscheint auf der Anmeldeseite, in „Meine Events" und als Ort im Outlook-Termin dieses Termins.'
                            : 'Appears on the registration page, in “My Events” and as the location of this date’s Outlook event.'}
                        </div>
                      </div>
                      <div className="dex-ui-field">
                        <label className="dex-ui-label">
                          <StepBadge n={15} />
                          {isDe ? 'Wie lautet die Adresse?' : 'What is the address?'}
                          <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
                        </label>
                        <AddressFields
                          isDe={isDe}
                          street={seAddr.street} houseNo={seAddr.houseNo} zip={seAddr.zip} city={seAddr.city}
                          onStreet={v => updateSub({ locationAddress: { ...seAddr, street: v } })}
                          onHouseNo={v => updateSub({ locationAddress: { ...seAddr, houseNo: v } })}
                          onZip={v => updateSub({ locationAddress: { ...seAddr, zip: v } })}
                          onCity={v => updateSub({ locationAddress: { ...seAddr, city: v } })}
                        />
                      </div>
                      {/* v18.44: Outlook-Ort pro Sub-Event überschreibbar (auch hier, nicht nur im Outlook-Editor).
                          v31.2: im Aufklapper, wie beim Hauptevent. */}
                      <button type="button" className={cx('dex-ui-disclosure', moreOpenSub && 'is-open')} onClick={() => setMoreOpenSub(o => !o)} aria-expanded={moreOpenSub}>
                        <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                        {isDe ? 'Weitere Einstellungen zum Ort' : 'More location settings'}
                        {(se.outlookLocation || '').trim() && <span className="dex-ui-disclosure-count">{isDe ? '1 angepasst' : '1 customised'}</span>}
                      </button>
                      {moreOpenSub && (
                        <div className="dex-ui-disclosure-body">
                          <div className="dex-ui-field">
                            <label className="dex-ui-label">
                              {isDe ? 'Was soll im Outlook-Termin als Ort stehen?' : 'What should the Outlook event show as location?'}
                            </label>
                            <input
                              className="form-input"
                              value={se.outlookLocation || ''}
                              onChange={e => updateSub({ outlookLocation: e.target.value })}
                              placeholder={buildOutlookLocation(se.location, seAddr) || (isDe ? 'z.B. Mezzomar, Harffstraße 110a, Düsseldorf' : 'e.g. Mezzomar, Harffstraße 110a, Düsseldorf')}
                            />
                            <div className="dex-ui-help">
                              {isDe
                                ? 'Leer lassen = automatisch aus Veranstaltungsort + Adresse dieses Sub-Events (siehe Platzhalter).'
                                : 'Leave empty = automatic from this sub-event\'s venue + address (see placeholder).'}
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                    <div className="dex-ui-section">
                      <div className="dex-ui-section-title">{isDe ? 'Wie kommen die Teilnehmer hin?' : 'How do attendees get there?'}</div>
                      <TransfersSection
                        items={seTransfers}
                        onChange={next => updateSub({ transferTimes: next })}
                        isDe={isDe}
                        isMobile={isMobile}
                        t={t}
                        defaultDate={seDefaultDate}
                      />
                    </div>
                    <div className="dex-ui-section">
                      <div className="dex-ui-section-title">{isDe ? 'Was passiert wann?' : 'What happens when?'}</div>
                      <label className="dex-ui-label" style={{ fontSize: '0.95rem' }}>
                        <StepBadge n={16} />
                        {isDe ? 'Wie sieht das Programm aus?' : 'What does the programme look like?'}
                        <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
                      </label>
                      {/* v30.94: gemeinsamer Tages-Editor (AgendaEditor) — s. dort. */}
                      <AgendaEditor
                        items={seAgenda}
                        onChange={upd => updateSub({ agenda: upd(seAgenda) })}
                        isDe={isDe}
                        isMobile={isMobile}
                        termSingular={agendaCheckIn ? agendaTermSingular : (isDe ? 'Programmpunkt' : 'Agenda item')}
                        termPlural={agendaCheckIn ? agendaTermPlural : (isDe ? 'Programmpunkte' : 'Agenda items')}
                        defaultDate={seDefaultDate}
                      />
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
              <div className="dex-ui-section">
              <div className="dex-ui-section-title">{isDe ? 'Wo findet das Event statt?' : 'Where does the event take place?'}</div>
              <div className="dex-ui-field">
                <label className="dex-ui-label">
                  <StepBadge n={14} />
                  {isDe ? 'Wie heißt der Veranstaltungsort?' : 'What is the venue called?'}
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
                <input className="form-input" value={location} onChange={e => setLocation(e.target.value)} placeholder={isDe ? 'z.B. RheinEnergieStadion, Köln' : 'e.g. RheinEnergieStadion, Cologne'} />
                <div className="dex-ui-help">
                  {isDe
                    ? 'Sprechender Name + Stadt reicht — er erscheint auf der Anmeldeseite, in der Eventliste und als Ort im Outlook-Termin der Teilnehmer.'
                    : 'A descriptive name + city is enough — it appears on the registration page, in the event list and as the location of the attendees’ Outlook event.'}
                </div>
              </div>
              <div className="dex-ui-field">
                <label className="dex-ui-label">
                  <StepBadge n={15} />
                  {isDe ? 'Wie lautet die Adresse?' : 'What is the address?'}
                  <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
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
                <AddressFields
                  isDe={isDe}
                  street={addrStreet} houseNo={addrHouseNo} zip={addrZip} city={addrCity}
                  onStreet={setAddrStreet} onHouseNo={setAddrHouseNo} onZip={setAddrZip} onCity={setAddrCity}
                />
                <div className="dex-ui-help">
                  {isDe
                    ? 'Teilnehmer können sie anklicken und z.B. in Google Maps öffnen; sie steht auch in der Bestätigungsmail und im Termin-Text.'
                    : 'Attendees can click it and open it e.g. in Google Maps; it also goes into the confirmation mail and the event text.'}
                </div>
              </div>

              {/* v30.26: Online-Meeting — Schalter bei Ort, darunter die Wahl
                  zwischen eigenem Link und automatischer Teams-Besprechung.
                  Der Unterschied ist eine echte Entscheidung (Besprechungs-
                  optionen behalten oder Bequemlichkeit), deshalb steht die
                  Konsequenz direkt an der Auswahl und nicht im Tooltip.
                  v31.2: Switch + zwei Auswahl-Kacheln statt Checkbox + Radios;
                  die Warnung zum Gruppenpostfach steht als eigener Kasten unter
                  der gewählten Kachel, nicht mehr im Kleingedruckten. */}
              <div className="dex-ui-card" style={{ marginTop: 16 }}>
                <label className="dex-ui-switch">
                  <input
                    type="checkbox"
                    checked={onlineMeetingMode !== 'none'}
                    onChange={e => setOnlineMeetingMode(e.target.checked ? 'own' : 'none')}
                  />
                  <span className="dex-ui-switch-track" />
                  <span className="dex-ui-switch-label">{isDe ? 'Teilnahme per Microsoft Teams möglich' : 'Attendance via Microsoft Teams possible'}</span>
                </label>
                <div className="dex-ui-help" style={{ marginLeft: 50 }}>
                  {isDe
                    ? 'Dann bekommt der Outlook-Termin einen Teams-Link — auch zusätzlich zu einem Präsenz-Ort (hybrid).'
                    : 'Then the Outlook event gets a Teams link — also in addition to a physical location (hybrid).'}
                </div>
                {onlineMeetingMode !== 'none' && (
                  <div className="dex-ui-stack" style={{ marginTop: 14 }}>
                    <div className="dex-ui-grid-2" role="radiogroup" aria-label={isDe ? 'Woher kommt der Teams-Link?' : 'Where does the Teams link come from?'}>
                      <button type="button" role="radio" aria-checked={onlineMeetingMode === 'own'} className={cx('dex-ui-choice', onlineMeetingMode === 'own' && 'is-active')} onClick={() => setOnlineMeetingMode('own')}>
                        <span className="dex-ui-choice-body">
                          <span className="dex-ui-choice-title">{isDe ? 'Ich stelle den Teams-Link selbst' : 'I provide the Teams link myself'}</span>
                          <span className="dex-ui-choice-desc">
                            {isDe
                              ? <>Du legst die Besprechung wie gewohnt in Outlook oder Teams an und trägst den Link unten ein. <strong>Empfohlen</strong>, wenn du Lobby, Aufzeichnung oder Referenten-Rollen brauchst — die kannst du nur an deiner eigenen Besprechung ändern.</>
                              : <>You create the meeting in Outlook or Teams as usual and paste the link below. <strong>Recommended</strong> if you need lobby, recording or presenter roles — those can only be changed on your own meeting.</>}
                          </span>
                        </span>
                        <span className="dex-ui-choice-check"><Check size={12} /></span>
                      </button>
                      <button type="button" role="radio" aria-checked={onlineMeetingMode === 'auto'} className={cx('dex-ui-choice', onlineMeetingMode === 'auto' && 'is-active')} onClick={() => setOnlineMeetingMode('auto')}>
                        <span className="dex-ui-choice-body">
                          <span className="dex-ui-choice-title">{isDe ? 'DEX erzeugt den Teams-Link automatisch' : 'DEX creates the Teams link automatically'}</span>
                          <span className="dex-ui-choice-desc">
                            {isDe
                              ? <>Der Termin wird als echte Teams-Besprechung angelegt — mit &bdquo;Teilnehmen&ldquo;-Knopf direkt im Kalender, ohne dass du etwas vorbereiten musst.</>
                              : <>The event is created as a real Teams meeting — with a &bdquo;Join&ldquo; button right in the calendar, with nothing to prepare.</>}
                          </span>
                        </span>
                        <span className="dex-ui-choice-check"><Check size={12} /></span>
                      </button>
                    </div>
                    {onlineMeetingMode === 'auto' && (
                      <div className="dex-ui-callout dex-ui-callout--warn">
                        <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                        <span>
                          {isDe
                            ? <><strong>Wichtig:</strong> Die Besprechung gehört dem Gruppenpostfach (no_reply.events). <strong>Du kannst danach keine Besprechungsoptionen mehr ändern</strong> — keine Lobby-Einstellung, keine Aufzeichnung, keine Referenten-Rollen.</>
                            : <><strong>Important:</strong> the meeting belongs to the group mailbox (no_reply.events). <strong>You cannot change any meeting options afterwards</strong> — no lobby settings, no recording, no presenter roles.</>}
                        </span>
                      </div>
                    )}
                    {/* v29.39: Teams-Link. Steht hier bei Ort und Adresse, weil er
                        dieselbe Frage beantwortet („wo findet es statt?") — und
                        bewusst NUR hier: Ein zweites Feld im Outlook-Editor wären
                        zwei Bedienwege für denselben Wert.
                        v30.26: nur noch im Modus „eigener Link" sichtbar. */}
                    {onlineMeetingMode === 'own' && (
                      <div className="dex-ui-field">
                        <label className="dex-ui-label">
                          {isDe ? 'Wie lautet der Teilnahme-Link?' : 'What is the join link?'}
                          <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
                          <InfoTooltip text={isDe ? (
                            <>
                              <strong>Was du hier einstellst:</strong> den <strong>Teilnahme-Link deiner eigenen Teams-Besprechung</strong>. Lege die Besprechung wie gewohnt in Outlook oder Teams an und kopiere den Link hierher — in diesem Modus erzeugt DEX selbst keine Teams-Besprechung.<br /><br />
                              <strong>Anzeige in der App:</strong> im <strong>Outlook-Termin</strong> als Knopf &bdquo;An Microsoft-Teams-Besprechung teilnehmen&ldquo;, im <strong>Organizer Center</strong> und in <strong>Meine Events</strong> als Teilnahme-Knopf.<br /><br />
                              <strong>Wichtig:</strong> Der Link steht im <strong>Text</strong> des Termins. Outlook kennt den Termin dadurch <strong>nicht</strong> als Online-Besprechung — es gibt also keinen &bdquo;Teilnehmen&ldquo;-Knopf in der Kalenderleiste und keinen Direktaufruf aus Teams heraus. Die Teilnehmer klicken den Link im Termin bzw. in der App.<br /><br />
                              <strong>Gilt für:</strong> das ganze Event, also auch für die Termine der Sub-Events.
                            </>
                          ) : (
                            <>
                              <strong>What you set here:</strong> the <strong>join link of your own Teams meeting</strong>. Create the meeting in Outlook or Teams as usual and paste the link here — in this mode DEX does not create a Teams meeting itself.<br /><br />
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
                        <div className="dex-ui-help">
                          {isDe
                            ? 'Erscheint im Outlook-Termin, im Organizer Center und in „Meine Events" — für das ganze Event samt Sub-Events. Für Outlook ist der Termin damit keine Online-Besprechung: kein „Teilnehmen"-Knopf im Kalender, Teilnehmer klicken den Link im Termin.'
                            : 'Appears in the Outlook event, in the Organizer Center and in “My Events” — for the whole event including sub-events. Outlook does not treat it as an online meeting: no “Join” button in the calendar, attendees click the link inside the event.'}
                        </div>
                        {teamsLink.trim() && !/^https?:\/\//i.test(teamsLink.trim()) && (
                          <div className="dex-ui-help" style={{ color: 'var(--dex-red, #da291c)', fontWeight: 600 }}>
                            {isDe
                              ? 'Das sieht nicht nach einem Link aus — er muss mit https:// beginnen, sonst wird er nicht übernommen.'
                              : 'That does not look like a link — it must start with https://, otherwise it is ignored.'}
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                )}
              </div>

              {/* v31.2: Aufklapper „Weitere Einstellungen zum Ort" — der
                  Outlook-Ort ist eine Feineinstellung, die fast niemand
                  anfasst; offen stehend las sie sich als drittes Pflichtfeld. */}
              <button type="button" className={cx('dex-ui-disclosure', moreOpenMain && 'is-open')} onClick={() => setMoreOpenMain(o => !o)} aria-expanded={moreOpenMain}>
                <span className="dex-ui-disclosure-chevron"><ChevronDown size={16} /></span>
                {isDe ? 'Weitere Einstellungen zum Ort' : 'More location settings'}
                {outlookLocationOverride.trim() && <span className="dex-ui-disclosure-count">{isDe ? '1 angepasst' : '1 customised'}</span>}
              </button>
              {moreOpenMain && (
                <div className="dex-ui-disclosure-body">
                  {/* v18.40: Ort im Outlook-Termin (überschreibbar) */}
                  <div className="dex-ui-field">
                    <label className="dex-ui-label">
                      {isDe ? 'Was soll im Outlook-Termin als Ort stehen?' : 'What should the Outlook event show as location?'}
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
                    <div className="dex-ui-help">
                      {isDe
                        ? 'Leer lassen = automatisch aus Veranstaltungsort + Adresse (siehe Platzhalter). Eine Eingabe überschreibt den Termin-Ort.'
                        : 'Leave empty = automatic from venue + address (see placeholder). Any input overrides the event location.'}
                    </div>
                  </div>
                </div>
              )}
              </div>

              {/* ===== Transferzeiten ===== v31.2: vor dem Programm — die
                  Geschichte des Schritts ist „Wo? → Wie kommt man hin? → Was
                  passiert wann?". Die Badge-Nummer 17 bleibt (Support). */}
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'Wie kommen die Teilnehmer hin?' : 'How do attendees get there?'}</div>
                <TransfersSection
                  items={transferTimes}
                  onChange={setTransferTimes}
                  isDe={isDe}
                  isMobile={isMobile}
                  t={t}
                  locationOptions={locationOptions}
                  defaultDate={startDate ? startDate.slice(0, 10) : ''}
                  tooltip={isDe ? (
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
                  )}
                />
              </div>

              {/* ===== Agenda Editor ===== */}
              <div className="dex-ui-section">
                <div className="dex-ui-section-title">{isDe ? 'Was passiert wann?' : 'What happens when?'}</div>
                <label className="dex-ui-label" style={{ fontSize: '0.95rem' }}>
                  <StepBadge n={16} />
                  {agendaCheckIn ? agendaPlural : (isDe ? 'Wie sieht das Programm aus?' : 'What does the programme look like?')}
                  <span className="dex-ui-label-optional">{isDe ? '(optional)' : '(optional)'}</span>
                  {agendaCheckIn && (
                    <span className="dex-ui-pill dex-ui-pill--green">
                      {isDe ? 'Check-in je Punkt' : 'Check-in per item'}
                    </span>
                  )}
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
                <div className="dex-ui-help" style={{ margin: '-2px 0 10px' }}>
                  {isDe
                    ? 'Teilnehmer sehen den Ablauf als Zeitleiste auf der Anmeldeseite und in „Meine Events" — nach Tag und Uhrzeit sortiert.'
                    : 'Attendees see the schedule as a timeline on the registration page and in “My Events” — sorted by day and time.'}
                </div>
                {/* v30.86: Im Programmpunkte-Modus (Schritt 1) ist diese Liste
                    mehr als Anzeige: Jeder Punkt ist eine Check-in-Station.
                    Ohne diesen Satz sucht der Organizer die Punkte in Schritt 1,
                    wo der Schalter steht. */}
                {agendaCheckIn && (
                  <div className="dex-ui-callout dex-ui-callout--success" style={{ marginBottom: 12 }}>
                    <span>
                      {isDe
                        ? <>Du hast in Schritt 1 <strong>Programmpunkte mit Anwesenheits-Check-in</strong> gewählt. Jede Zeile hier ist ein Punkt, an dem später eingecheckt wird — Titel, Datum, Start, Ende und Raum reichen. Teilnehmer sehen die Liste auf der Anmeldeseite und in &bdquo;Meine Events&ldquo;; angemeldet wird nur fürs Event.</>
                        : <>In step 1 you chose <strong>agenda items with attendance check-in</strong>. Every row here is an item people check in at — title, date, start, end and room are enough. Attendees see the list on the registration page and in “My events”; they register for the event only.</>}
                    </span>
                  </div>
                )}
                {/* v30.94: Tages-Editor statt einer Karte je Punkt (Nutzer-
                    Befund 07.09.2026: 27 Karten mit je sechs Labels und
                    US-Datumsanzeige aus den nativen Feldern). Datum einmal je
                    Tag über den DatePicker (dd.MM.yyyy), Zeiten als HH:MM-
                    Text, Beschreibung aufklappbar. Tag kopieren/löschen und
                    „Weiterer Tag" liegen im Editor — `copyLastAgendaDay`
                    (v30.86) ist damit abgelöst. */}
                <AgendaEditor
                  items={agenda}
                  onChange={setAgenda}
                  isDe={isDe}
                  isMobile={isMobile}
                  termSingular={agendaCheckIn ? agendaTermSingular : (isDe ? 'Programmpunkt' : 'Agenda item')}
                  termPlural={agendaCheckIn ? agendaTermPlural : (isDe ? 'Programmpunkte' : 'Agenda items')}
                  defaultDate={startDate ? startDate.slice(0, 10) : ''}
                />
              </div>

              </div>{/* v15.6: close hauptGreyoutWrapperStyle div (Step 3) */}
              </div>{/* v15.0: close activeLocationTabIdx===0 wrapper (Top-Level Ort/Adresse/Agenda/Transfer) */}

              </div>
  );
};
