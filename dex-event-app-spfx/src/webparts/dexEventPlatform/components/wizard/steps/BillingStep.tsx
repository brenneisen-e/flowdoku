/**
 * v30.13 — Modularisierung Stufe 3, Pilot: Schritt „Abrechnung" (v29.66,
 * F&A-Pilot) als eigene Komponente. Der JSX-Block ist 1:1 aus
 * EventCreationPage übernommen; der Schritt hängt an genau sechs
 * State-Werten und ist damit der kleinste Props-Vertrag des Wizards.
 * `billingMissing` ist abgeleitet und wandert mit hierher (einzige
 * Nutzung). Die Sichtbarkeit steuert weiter der Wizard: das adminLike-
 * Gate bleibt am Aufrufer, `visible` ersetzt `currentStep === 9` —
 * display:none statt unmount, damit Eingaben beim Schrittwechsel
 * erhalten bleiben (gleiches Muster wie alle anderen Schritte).
 * Texte bewusst nur deutsch — der Pilot ist auf Admins begrenzt (isDe
 * kommt mit der Freischaltung für Organizer).
 *
 * v31.2: Umbau nach docs/ui-leitfaden.md. Beide Fragen des Schritts (relevant
 * ja/nein, Versand manuell/automatisch) sind Auswahl-Kacheln mit einer Zeile
 * Folge statt nackter Radios; die Pflichtangaben stehen VOR dem Versandmodus
 * (Pflicht vor Optional), und der Status-Kasten sitzt direkt über den Feldern,
 * die er zählt. Die langen Erklärabsätze stecken in InfoTooltips — nichts
 * davon ist gestrichen, nur nachrangig. Bindungen und Setter sind unverändert.
 */
import * as React from 'react';
import { BILLING_FIELDS } from '../../../data/billingFields';
import { UserFieldPicker } from '../../UserFieldPicker';
import { useRoles } from '../../../context/RoleContext';
import { cx } from '../../dexUi';
import { InfoTooltip } from '../../InfoTooltip';
import { AlertCircle, Check } from '../../Icons';

export interface BillingStepProps {
  visible: boolean;
  billingRelevant: boolean | null;
  setBillingRelevant: (v: boolean) => void;
  billingSendMode: 'auto' | 'manual';
  setBillingSendMode: (v: 'auto' | 'manual') => void;
  billingFields: Record<string, string>;
  setBillingFields: React.Dispatch<React.SetStateAction<Record<string, string>>>;
}

export const BillingStep: React.FC<BillingStepProps> = ({
  visible, billingRelevant, setBillingRelevant,
  billingSendMode, setBillingSendMode, billingFields, setBillingFields,
}) => {
  // v30.45: Der Person-Picker braucht die Suche aus dem RoleContext. Wie `t`
  // und `confirmDialog` in den anderen Schritten holt sich die Komponente das
  // selbst, statt es durch den Props-Vertrag zu schleifen.
  const { searchUsers, searchUser } = useRoles();
  const billingMissing = BILLING_FIELDS.filter(f => !(billingFields[f.id] || '').trim());
  const complete = billingMissing.length === 0;
  // v31.2: EINE Kachel-Form für beide Fragen des Schritts — Titel, eine Zeile
  // Folge, Häkchen rechts. Radios ohne Konsequenz-Zeile ließen den Organizer
  // raten, was „Ja" nach sich zieht; hier steht es in der Kachel.
  const choice = (active: boolean, onPick: () => void, title: React.ReactNode, desc: React.ReactNode): React.ReactElement => (
    <button type="button" role="radio" aria-checked={active} className={cx('dex-ui-choice', active && 'is-active')} onClick={onPick}>
      <div className="dex-ui-choice-body">
        <div className="dex-ui-choice-title">{title}</div>
        <div className="dex-ui-choice-desc">{desc}</div>
      </div>
      <span className="dex-ui-choice-check" aria-hidden="true">{active && <Check size={12} />}</span>
    </button>
  );
  return (
    <div style={{ display: visible ? 'block' : 'none' }}>
      {/* v30.28: Der grüne Schritt-Kopf fehlte seit der Extraktion in v30.13 —
          jeder andere Schritt bringt ihn mit, hier blieb an seiner Stelle ein
          leerer weißer Streifen (die Form-Karte kompensiert per marginBottom
          die negative Top-Margin des Kopfs, der nie kam). */}
      <h2 className="dex-step-head-title">
        <span className="dex-step-eyebrow">Schritt 10</span>
        Abrechnung
      </h2>
      <p className="dex-step-head-lead">
        Nur für Admins: Ist das Event abrechnungsrelevant, sammelt DEX hier die
        Angaben für Finance &amp; Accounting und verschickt sie nach dem Event.
      </p>

      <div className="dex-ui-section">
        <div className="dex-ui-section-title">Abrechnungsrelevanz</div>
        <div className="dex-ui-label" style={{ fontSize: '0.95rem' }}>
          Muss dieses Event gegenüber Finance &amp; Accounting abgerechnet werden?
          <InfoTooltip text={<>
            Abrechnungsrelevante Events sind Veranstaltungen, deren Kosten oder
            Bewirtungsaufwendungen gegenüber Finance &amp; Accounting dokumentiert
            oder abgerechnet werden müssen. Das ist der Fall, wenn im Nachgang
            <strong> Rechnungen über die Kreditorenbuchhaltung eingereicht
            werden</strong> — etwa für Catering, eine externe Raumbuchung oder
            Anmeldegebühren (z.B. Startgelder für Läufer) — oder wenn für das
            Event <strong>Ariba-Bestellungen</strong> ausgelöst werden.
          </>} />
        </div>
        <p className="dex-ui-section-desc">
          Ja, wenn danach Rechnungen über die Kreditorenbuchhaltung laufen oder
          Ariba-Bestellungen ausgelöst werden — etwa Catering, Raumbuchung, Startgelder.
        </p>
        <div className="dex-ui-grid-2" role="radiogroup" aria-label="Abrechnungsrelevant">
          {choice(billingRelevant === true, () => setBillingRelevant(true),
            'Ja, abrechnungsrelevant',
            <>Es fallen Kosten an, die F&amp;A dokumentiert oder abrechnet. Du trägst unten die Pflichtangaben ein, DEX meldet sie an F&amp;A.</>)}
          {choice(billingRelevant === false, () => setBillingRelevant(false),
            'Nein, nicht abrechnungsrelevant',
            <>Keine Rechnungen, keine Ariba-Bestellungen. Es gibt nichts an F&amp;A zu melden — der Schritt ist damit erledigt.</>)}
        </div>
      </div>

      {billingRelevant === true && (
        <>
          <div className="dex-ui-section">
            <div className="dex-ui-section-title">Angaben für Finance &amp; Accounting</div>
            {/* v30.4: Legende — die Sternchen standen unerklärt im Raum. */}
            <p className="dex-ui-section-desc">
              <span className="required">*</span> Pflichtangabe — ohne sie gilt die Abrechnungsmeldung an Finance &amp; Accounting als unvollständig. Speichern kannst du trotzdem jederzeit.
            </p>
            {/* Status — systemseitig aus den Pflichtfeldern abgeleitet,
                nie gespeichert und nie von Hand setzbar. v31.2: steht direkt
                über den Feldern, die er zählt, statt drei Kästen weiter oben. */}
            <div className={cx('dex-ui-callout', complete ? 'dex-ui-callout--success' : 'dex-ui-callout--warn')} role="status" style={{ marginBottom: 14 }}>
              <span className="dex-ui-callout-icon">{complete ? <Check size={16} /> : <AlertCircle size={16} />}</span>
              <span>
                {complete
                  ? <><strong>Vollständig</strong> — alle {BILLING_FIELDS.length} Pflichtangaben sind gepflegt.</>
                  : <><strong>Noch unvollständig</strong> — {billingMissing.length} von {BILLING_FIELDS.length} Pflichtangaben fehlen (orange Rahmen). Speichern ist trotzdem möglich.</>}
              </span>
            </div>
            <div className="dex-ui-grid-2">
              {BILLING_FIELDS.map(f => {
                const val = billingFields[f.id] || '';
                const empty = !val.trim();
                const setVal = (v: string): void => setBillingFields(prev => ({ ...prev, [f.id]: v }));
                return (
                  // v30.4: Flex-Spalte, Label wächst — die Eingabefelder
                  // einer Zeile stehen damit auf gleicher Höhe, auch wenn
                  // ein Label („Name der Veranstaltung bzw. Anlass …")
                  // zweizeilig umbricht.
                  <div key={f.id} className="dex-ui-field" style={{ display: 'flex', flexDirection: 'column', marginBottom: 0 }}>
                    <label className="dex-ui-label" style={{ flexGrow: 1, alignItems: 'flex-start' }}>
                      {f.label} <span className="required">*</span>
                    </label>
                    {f.type === 'user' ? (
                      // v30.45: Kontaktperson als Person-Picker statt Freitext
                      // (F&A-Fachkonzept). Derselbe Picker wie überall sonst —
                      // Chip mit Profilfoto, Job Title und Standort, Auswahl nur
                      // aus dem Tenant. Der gespeicherte Wert bleibt der String
                      // `Name <email>`, deshalb ändert sich an
                      // `missingBillingFields`, am F&A-Mailtext und an allem,
                      // was `_billing` liest, nichts.
                      <UserFieldPicker
                        value={val}
                        onChange={setVal}
                        searchUsers={searchUsers}
                        searchUserByEmail={searchUser}
                        placeholder="Name oder E-Mail der Kontaktperson…"
                        errorStyle={empty ? { borderColor: 'var(--dex-orange, #ed8b00)' } : {}}
                        forcedIsDe
                      />
                    ) : f.type === 'select' ? (
                      <select
                        className="form-input"
                        value={val}
                        onChange={e => setVal(e.target.value)}
                        style={{ borderColor: empty ? 'var(--dex-orange, #ed8b00)' : undefined }}
                      >
                        <option value="">Bitte wählen…</option>
                        {(f.options || []).map(o => <option key={o} value={o}>{o}</option>)}
                      </select>
                    ) : (
                      <input
                        className="form-input"
                        type={f.type === 'date' ? 'date' : 'text'}
                        value={val}
                        onChange={e => setVal(e.target.value)}
                        style={{ borderColor: empty ? 'var(--dex-orange, #ed8b00)' : undefined }}
                      />
                    )}
                  </div>
                );
              })}
            </div>
          </div>

          {/* v31.2: Der Versandmodus ist optional (Vorgabe „manuell") und steht
              deshalb HINTER den Pflichtangaben; die Vorgabe steht links. */}
          <div className="dex-ui-section">
            <div className="dex-ui-section-title">Übermittlung an Finance &amp; Accounting</div>
            <div className="dex-ui-label" style={{ fontSize: '0.95rem' }}>
              Wie sollen die Abrechnungsinformationen an F&amp;A gehen?
              <InfoTooltip text={<>
                Abrechnungsrelevante Informationen müssen an die Finance &amp; Accounting
                Abteilung gemeldet werden. Dies beinhaltet insbesondere allgemeine
                Eventinformationen, Teilnehmerlisten sowie Rechnungen und Belege. Die
                Einstellung hier unterstützt die standardisierte und teilweise
                automatisierte Übermittlung dieser Informationen.
              </>} />
            </div>
            <div className="dex-ui-grid-2" role="radiogroup" aria-label="Versand an Finance & Accounting">
              {choice(billingSendMode === 'manual', () => setBillingSendMode('manual'),
                <>Manuell senden <span className="dex-ui-pill dex-ui-pill--gray" style={{ marginLeft: 6 }}>Standard</span></>,
                <>Kein automatischer Versand. Du schickst Abrechnungsinformationen und Teilnehmerliste selbst aus dem Organizer Center an F&amp;A.</>)}
              {choice(billingSendMode === 'auto', () => setBillingSendMode('auto'),
                'Automatisch senden',
                <>
                  DEX schickt die Abrechnungsinformationen 7 Kalendertage vor dem Event
                  (bei kurzfristiger Erstellung: sofort nach Aktivierung) und die finale
                  Teilnehmerliste 7 Kalendertage danach — jeweils an F&amp;A, Organizer in CC.
                  <em style={{ display: 'block', marginTop: 4, color: 'var(--dex-orange-dark, #b35a00)' }}>
                    Pilot: Die Auswahl wird bereits gespeichert, der Automatik-Flow existiert noch nicht.
                  </em>
                </>)}
            </div>
          </div>
        </>
      )}
    </div>
  );
};
