/**
 * Rollenverwaltung — zwei Reiter: „Personen" und „Rechte je Rolle".
 *
 * Aufbau nach der Rollenverwaltung in DEX (`SettingsPage`, `RoleMatrixPage`),
 * aber auf das verkleinert, was diese Plattform hat: drei Rollen, drei Listen.
 *
 * Fünf Regeln aus echten Vorfällen prägen die Seite — sie stehen hier, weil man
 * sie beim Umbauen leicht wieder herausnimmt:
 *
 * 1. **Ein Lesefehler ist keine Null.** Personensuche und Rechte-Prüfung
 *    liefern bei einem Fehler `null`. Die Seite sagt dann „Suche nicht möglich"
 *    bzw. „Rechte konnten nicht gelesen werden" — nie „Keine Treffer" und nie
 *    „alles in Ordnung". In DEX hat genau diese Verwechslung Admins gesagt, eine
 *    Person gebe es nicht, weil eine Anfrage gedrosselt war.
 * 2. **Ein Eintrag in der Rollenliste ist noch keine Rolle.** Sie wirkt erst,
 *    wenn die Rechte auf den Listen stehen (Leserecht auf der Rollenliste,
 *    Bearbeiten auf den Use Cases, Beitragen im Protokoll). Jeder Schreibvorgang
 *    meldet, was davon NICHT gesetzt werden konnte; „Rechte prüfen" liest es
 *    nach.
 * 3. **Nie implizit herabstufen.** Wer schon eine Rolle hat, wird über die
 *    Auswahl in seiner Zeile geändert — mit Bestätigung —, nie über „Person
 *    hinzufügen".
 * 4. **Die eigene Rolle ändert man nicht selbst.** Wer sich herabstuft, käme
 *    nicht mehr an diese Seite; der Entzug der Rechte läuft für die angemeldete
 *    Person ohnehin nicht (Selbstschutz im Service).
 * 5. **Die Adresse ist der einzige Schlüssel — und sie ist nicht eindeutig.**
 *    Verglichen wird über `toLowerCase().trim()`; der Audit meldet, wenn die
 *    Rechte unter einer anderen Schreibweise stehen als die Rollenzeile.
 *
 * Alle Hooks stehen vor dem ersten frühen `return` (ESLint
 * `react-hooks/rules-of-hooks` ist hier ein Fehler, kein Hinweis).
 */

import * as React from 'react';
import { cx, ensureDexUiStyles } from './dexUi';
import { Plus, Trash2, AlertCircle, Search, X, Check, RefreshCw, Users, Info } from './Icons';
import { useRoles } from '../context/RoleContext';
import { useLanguage } from '../context/LanguageContext';
import { useNavigation } from '../context/NavigationContext';
import { useDialog } from '../context/DialogContext';
import { RoleAssignment, UserRole } from '../types';
import { ROLLEN_ALLE, rolleLabel, rolleRang } from '../utils/rollen';
import { RECHTE_MATRIX, MATRIX_SPALTEN, MatrixWert } from '../data/rollenMatrix';
import { PersonenTreffer, RechteBericht, RechteListe } from '../services/SharePointService';

// ---------------------------------------------------------------------------
// Kleinkram
// ---------------------------------------------------------------------------

/** Ein Schild mit Haken (lucide „shield-check") — Icons.tsx gehört jemand anderem. */
function ShieldCheck(props: { size?: number }): React.ReactElement {
  const size = props.size || 16;
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z" />
      <path d="m9 12 2 2 4-4" />
    </svg>
  );
}

/** Zwei Adressen gleich? Die Adresse ist der Schlüssel — verglichen wird klein und getrimmt. */
function gleich(a: string | undefined | null, b: string | undefined | null): boolean {
  return (a || '').trim().toLowerCase() === (b || '').trim().toLowerCase();
}

/** Initialen für den Kreis. „Nachname, Vorname" (so heißen Deloitte-Konten) wird gedreht. */
function initialenVon(name: string): string {
  const roh = (name || '').trim();
  if (!roh) return '?';
  const worte = (roh.indexOf(',') >= 0 ? roh.split(',').reverse().join(' ') : roh).trim().split(/\s+/);
  const erster = (worte[0] || '').charAt(0);
  const letzter = worte.length > 1 ? (worte[worte.length - 1] || '').charAt(0) : '';
  return ((erster + letzter) || roh.charAt(0)).toUpperCase();
}

function Avatar(props: { name: string }): React.ReactElement {
  // Bewusst keine Fotos: Ein Bild je Zeile wäre eine Anfrage je Zeile, und die
  // Drosselung ist das Risiko dieser Seite.
  return <span className="dex-ui-avatar" aria-hidden="true">{initialenVon(props.name)}</span>;
}

function datumText(iso: string, isDe: boolean): string {
  if (!iso) return '';
  try {
    const d = new Date(iso);
    if (isNaN(d.getTime())) return '';
    return d.toLocaleDateString(isDe ? 'de-DE' : 'en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' });
  } catch {
    return '';
  }
}

/**
 * Die Rechte-Texte, die der `RoleContext` und der Service als Klartext liefern
 * (ASCII, ohne Umlaute — sie sind älter als die zweisprachige Oberfläche).
 * Unbekanntes bleibt, wie es ist: Ein roher Satz ist besser als ein
 * verschwundener.
 */
const RECHTE_TEXTE: Record<string, { de: string; en: string }> = {
  'Rollenliste (Vollzugriff)': { de: 'Rollenliste (Vollzugriff)', en: 'Roles list (full control)' },
  'Rollenliste (Lesen)': { de: 'Rollenliste (Lesen)', en: 'Roles list (read)' },
  'Rollenliste (Vollzugriff bleibt bestehen)': { de: 'Rollenliste (Vollzugriff ließ sich nicht entziehen)', en: 'Roles list (full control could not be revoked)' },
  'Use-Case-Liste (Bearbeiten)': { de: 'Use-Case-Liste (Bearbeiten)', en: 'Use case list (edit)' },
  'Protokoll (Beitragen)': { de: 'Protokoll (Beitragen)', en: 'Log (contribute)' },
  'Person nicht aufloesbar': { de: 'Person nicht auflösbar', en: 'person could not be resolved' },
  'Entzug unvollstaendig': { de: 'Entzug unvollständig', en: 'revocation incomplete' },
};

function rechteText(roh: string, isDe: boolean): string {
  return Object.prototype.hasOwnProperty.call(RECHTE_TEXTE, roh) ? (isDe ? RECHTE_TEXTE[roh].de : RECHTE_TEXTE[roh].en) : roh;
}

/** Was eine Rolle auf einer Liste braucht — in Worten, für den Bericht. */
function listeMitRecht(key: RechteListe, rolle: UserRole, isDe: boolean): string {
  if (key === 'roles') return rolle === 'Admin' ? (isDe ? 'Rollenliste (Vollzugriff)' : 'Roles list (full control)') : (isDe ? 'Rollenliste (Lesen)' : 'Roles list (read)');
  if (key === 'useCases') return isDe ? 'Use-Case-Liste (Bearbeiten)' : 'Use case list (edit)';
  return isDe ? 'Protokoll (Beitragen)' : 'Log (contribute)';
}

/** Der Name der Liste ohne Recht — für „hat mehr als vorgesehen". */
function listenName(key: RechteListe, isDe: boolean): string {
  if (key === 'roles') return isDe ? 'Rollenliste' : 'Roles list';
  if (key === 'useCases') return isDe ? 'Use-Case-Liste' : 'Use case list';
  return isDe ? 'Protokoll' : 'Log';
}

function rollePille(rolle: UserRole): string {
  return cx('dex-ui-pill', 'dex-ui-pill--sm', rolle === 'Admin' ? 'dex-ui-pill--orange' : rolle === 'Organizer' ? 'dex-ui-pill--green' : 'dex-ui-pill--gray');
}

// ---------------------------------------------------------------------------
// Die Seite
// ---------------------------------------------------------------------------

type Reiter = 'personen' | 'rechte';

export default function RolePage(): React.ReactElement {
  ensureDexUiStyles();
  const { roles, isAdmin, isRolesLoading, rolesReadStatus, refreshRoles, service } = useRoles();
  const { t } = useLanguage();
  const { navigate } = useNavigation();

  const [reiter, setReiter] = React.useState<Reiter>('personen');
  // Klartext, wenn das LETZTE Lesen der Rollenliste gescheitert ist. `refreshRoles`
  // lässt in dem Fall den alten Stand stehen und meldet nichts — die Seite
  // würde also einen Stand zeigen, der nicht mehr stimmt, ohne es zu sagen.
  const [veraltet, setVeraltet] = React.useState('');
  const [neuLaedt, setNeuLaedt] = React.useState(false);

  const pruefeStand = React.useCallback((): void => {
    const st = service.lastRolesReadStatus;
    setVeraltet(st >= 200 && st < 300 ? '' : (st ? `HTTP ${st}` : 'keine Antwort'));
  }, [service]);

  const neuLaden = async (): Promise<void> => {
    setNeuLaedt(true);
    try {
      await refreshRoles();
      pruefeStand();
    } finally {
      setNeuLaedt(false);
    }
  };

  if (!isAdmin) {
    return (
      <div className="dex-ui-empty">
        <div className="dex-ui-empty-title">{t('Nur für Admins', 'Admins only')}</div>
        <div className="dex-ui-empty-desc">{t('Rollen vergeben dürfen nur Admins der Plattform.', 'Only platform admins may assign roles.')}</div>
        <button type="button" className="dex-ui-empty-action" onClick={() => navigate('landing')}>{t('Zum Startbildschirm', 'Back to the start screen')}</button>
      </div>
    );
  }

  return (
    <div>
      <div className="dex-ui-page-head">
        <div>
          <h1 className="dex-ui-page-head-title">{t('Rollenverwaltung', 'Role management')}</h1>
          <p className="dex-ui-page-head-meta">
            {isRolesLoading ? t('Rollen werden gelesen …', 'Reading roles …') : `${roles.length} ${roles.length === 1 ? t('Eintrag', 'entry') : t('Einträge', 'entries')}`}
          </p>
        </div>
        <div className="dex-ui-page-head-actions">
          <button type="button" className="dex-ui-textbtn" disabled={neuLaedt || isRolesLoading} onClick={() => { void neuLaden(); }}>
            <RefreshCw size={14} /> {neuLaedt ? t('Lädt …', 'Loading …') : t('Neu laden', 'Reload')}
          </button>
        </div>
      </div>

      {rolesReadStatus === 'forbidden' && (
        <div className="dex-ui-callout dex-ui-callout--danger" role="alert" style={{ marginBottom: 16 }}>
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span>{t('Die Rollenliste ist für dich nicht lesbar. Was hier steht, ist deshalb unvollständig.', 'The roles list is not readable for you. What you see here is therefore incomplete.')}</span>
        </div>
      )}
      {rolesReadStatus === 'error' && (
        <div className="dex-ui-callout dex-ui-callout--warn" role="alert" style={{ marginBottom: 16 }}>
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span>{t('Die Rollenliste konnte nicht gelesen werden (Netzwerk oder Drosselung). Was hier steht, kann unvollständig sein — versuch es mit „Neu laden“ noch einmal.', 'The roles list could not be read (network or throttling). What you see here may be incomplete — try “Reload” again.')}</span>
        </div>
      )}
      {veraltet && (
        <div className="dex-ui-callout dex-ui-callout--warn" role="alert" style={{ marginBottom: 16 }}>
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span>
            {t(`Das letzte Lesen der Rollenliste ist gescheitert (${veraltet}). Angezeigt wird der Stand von vorhin — eine Änderung kann schon durch sein, ohne dass sie hier steht. „Neu laden“ versucht es noch einmal.`,
              `The last read of the roles list failed (${veraltet}). You are seeing the earlier state — a change may already be through without showing here. “Reload” tries again.`)}
          </span>
        </div>
      )}

      <div style={{ marginBottom: 18 }}>
        <div className="dex-ui-tabs" role="tablist" aria-label={t('Bereiche der Rollenverwaltung', 'Sections of the role management')}>
          <button type="button" role="tab" aria-selected={reiter === 'personen'} className={cx('dex-ui-tab', reiter === 'personen' && 'is-active')} onClick={() => setReiter('personen')}>
            {t('Personen', 'People')}
          </button>
          <button type="button" role="tab" aria-selected={reiter === 'rechte'} className={cx('dex-ui-tab', reiter === 'rechte' && 'is-active')} onClick={() => setReiter('rechte')}>
            {t('Rechte je Rolle', 'Rights per role')}
          </button>
        </div>
      </div>

      {reiter === 'personen' ? <PersonenTab pruefeStand={pruefeStand} /> : <RechteMatrix />}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Reiter „Personen"
// ---------------------------------------------------------------------------

type Meldung = { ton: 'ok' | 'warn' | 'fehler'; text: string } | null;

function PersonenTab(props: { pruefeStand: () => void }): React.ReactElement {
  const { lastRightsMissing } = useRoles();
  const { t, isDe } = useLanguage();

  // Ein Schreibvorgang oder Audit läuft — sperrt alle anderen Knöpfe. Zwei
  // gleichzeitige Rechte-Vergaben an dieselbe Liste sind genau das Muster, das
  // SharePoint drosselt.
  const [laeuft, setLaeuft] = React.useState(false);
  const [meldung, setMeldung] = React.useState<Meldung>(null);
  const [bericht, setBericht] = React.useState<RechteBericht | null>(null);
  const meldeTimer = React.useRef<number | undefined>(undefined);

  React.useEffect(() => () => window.clearTimeout(meldeTimer.current), []);

  const melde = (ton: 'ok' | 'warn' | 'fehler', text: string): void => {
    window.clearTimeout(meldeTimer.current);
    setMeldung({ ton, text });
    // Erfolg räumt sich selbst weg; alles andere bleibt stehen, bis es jemand
    // schließt oder die nächste Aktion es ersetzt.
    if (ton === 'ok') meldeTimer.current = window.setTimeout(() => setMeldung(null), 7000);
  };

  /**
   * Nach jedem Schreiben: melden, was an Rechten fehlt. Nie verschweigen — „Zeile
   * gesetzt" heißt nicht „Rechte gesetzt" (in DEX fehlten so bei 18 von 126
   * Einträgen Rechte, und jede Zuweisung hatte Erfolg gemeldet).
   */
  const meldeRechte = (erfolg: boolean, erledigt: string, fehler: string, folge?: string): void => {
    if (!erfolg) { melde('fehler', fehler); return; }
    const fehlt = lastRightsMissing();
    if (fehlt.length > 0) {
      // `folge` ersetzt den Standardsatz dort, wo er nicht stimmt: Nach dem
      // Entfernen gibt es keine Zeile mehr, die „Rechte prüfen" nachsetzen könnte.
      const satz = folge || (isDe
        ? 'Solange das so ist, wirkt die Rolle nicht vollständig. „Rechte prüfen“ unten zeigt es und setzt nach.'
        : 'While that is so, the role does not fully apply. “Check rights” below shows it and re-grants.');
      melde('warn', isDe
        ? `${erledigt} Aber bei den SharePoint-Rechten hat nicht alles geklappt: ${fehlt.map(f => rechteText(f, true)).join(', ')}. ${satz}`
        : `${erledigt} But not everything worked with the SharePoint rights: ${fehlt.map(f => rechteText(f, false)).join(', ')}. ${satz}`);
      return;
    }
    melde('ok', erledigt);
  };

  // Ein Bericht gilt für den Stand, in dem er gelesen wurde. Nach jeder Änderung
  // wäre „Rechte ok" neben einer Zeile, die es nicht mehr ist, eine Falschaussage.
  const verwerfeBericht = (): void => setBericht(null);

  return (
    <div>
      {meldung && (
        <div
          className={cx('dex-ui-callout', meldung.ton === 'ok' ? 'dex-ui-callout--success' : meldung.ton === 'warn' ? 'dex-ui-callout--warn' : 'dex-ui-callout--danger')}
          role={meldung.ton === 'ok' ? 'status' : 'alert'}
          style={{ marginBottom: 16 }}
        >
          <span className="dex-ui-callout-icon">{meldung.ton === 'ok' ? <Check size={16} /> : <AlertCircle size={16} />}</span>
          <span className="dex-ui-callout-body">{meldung.text}</span>
          <button type="button" className="dex-ui-iconbtn" style={{ width: 24, height: 24 }} onClick={() => setMeldung(null)} aria-label={t('Meldung schließen', 'Dismiss message')}>
            <X size={13} />
          </button>
        </div>
      )}

      <RolleVergeben laeuft={laeuft} setLaeuft={setLaeuft} melde={melde} meldeRechte={meldeRechte} pruefeStand={props.pruefeStand} verwerfeBericht={verwerfeBericht} />
      <RollenListe laeuft={laeuft} setLaeuft={setLaeuft} melde={melde} meldeRechte={meldeRechte} pruefeStand={props.pruefeStand} verwerfeBericht={verwerfeBericht} bericht={bericht} />
      <RechtePruefung laeuft={laeuft} setLaeuft={setLaeuft} bericht={bericht} setBericht={setBericht} />
    </div>
  );
}

interface AktionsProps {
  laeuft: boolean;
  setLaeuft: (b: boolean) => void;
  melde: (ton: 'ok' | 'warn' | 'fehler', text: string) => void;
  meldeRechte: (erfolg: boolean, erledigt: string, fehler: string, folge?: string) => void;
  pruefeStand: () => void;
  verwerfeBericht: () => void;
}

// --- Person suchen und Rolle vergeben ---------------------------------------

interface SuchStand {
  status: 'leer' | 'sucht' | 'ok' | 'fehler';
  treffer: PersonenTreffer[];
  ausgeblendet: number;
  http: number;
}

const SUCHE_LEER: SuchStand = { status: 'leer', treffer: [], ausgeblendet: 0, http: 0 };

/** Die Rollen, die man Personen VERGIBT. „User" braucht keinen Eintrag — dafür ist die Auswahl in der Zeile da. */
const VERGEBBAR: UserRole[] = ['Organizer', 'Admin'];

function RolleVergeben(props: AktionsProps): React.ReactElement {
  const { roles, addRole, searchUsers, lastPersonSearchStatus } = useRoles();
  const { t } = useLanguage();
  const { confirmDialog } = useDialog();
  const { laeuft, setLaeuft, melde, meldeRechte, pruefeStand, verwerfeBericht } = props;

  const [text, setText] = React.useState('');
  const [international, setInternational] = React.useState(false);
  const [suche, setSuche] = React.useState<SuchStand>(SUCHE_LEER);
  const [nochmal, setNochmal] = React.useState(0);
  const [gewaehlt, setGewaehlt] = React.useState<PersonenTreffer | null>(null);
  const [rolle, setRolle] = React.useState<UserRole>('Organizer');
  // Nur die jüngste Anfrage zählt: Wer schnell tippt, schickt mehrere ab, und
  // die langsamste darf die neueste nicht überschreiben.
  const zaehler = React.useRef(0);

  React.useEffect(() => {
    if (gewaehlt) return undefined;
    const q = text.trim();
    if (q.length < 2) {
      zaehler.current++;
      setSuche(SUCHE_LEER);
      return undefined;
    }
    const nr = ++zaehler.current;
    setSuche(s => ({ ...s, status: 'sucht' }));
    const timer = window.setTimeout(() => {
      searchUsers(q, international).then(
        res => {
          if (nr !== zaehler.current) return;
          setSuche(res === null
            ? { status: 'fehler', treffer: [], ausgeblendet: 0, http: lastPersonSearchStatus() }
            : { status: 'ok', treffer: res.treffer, ausgeblendet: res.ausgeblendet, http: 200 });
        },
        () => {
          if (nr === zaehler.current) setSuche({ status: 'fehler', treffer: [], ausgeblendet: 0, http: 0 });
        },
      );
    }, 400);
    return () => {
      window.clearTimeout(timer);
      // Eine Antwort, die nach dem Verlassen der Seite eintrifft, darf nichts mehr setzen.
      zaehler.current++;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [text, international, gewaehlt, nochmal]);

  const bestehende = (mail: string): RoleAssignment | undefined => roles.filter(r => gleich(r.userEmail, mail))[0];
  const bestehend = gewaehlt ? bestehende(gewaehlt.email) : undefined;

  async function vergeben(): Promise<void> {
    if (!gewaehlt || laeuft || bestehend) return;
    const name = gewaehlt.displayName || gewaehlt.email;
    if (rolle === 'Admin') {
      // Admin ist die einzige Rolle, die andere Rollen vergibt — und die auf
      // der Rollenliste Vollzugriff bekommt. Das ist keine Kleinigkeit.
      const ja = await confirmDialog(
        t(`${name} zum Admin machen? Admins verwalten alle Rollen und haben Vollzugriff auf die Rollenliste — auch für deine eigene.`,
          `Make ${name} an admin? Admins manage all roles and have full control of the roles list — including yours.`),
        { confirmLabel: t('Zum Admin machen', 'Make admin') },
      );
      if (!ja) return;
    }
    setLaeuft(true);
    try {
      const ok = await addRole(gewaehlt.email, name, rolle);
      pruefeStand();
      meldeRechte(
        ok,
        t(`${name} ist jetzt ${rolleLabel(rolle)}.`, `${name} is now ${rolleLabel(rolle)}.`),
        t(`${name} konnte nicht eingetragen werden — die Rollenliste hat den Eintrag nicht angenommen. Versuch es gleich noch einmal; bleibt es dabei, fehlt dir ein Recht auf der Liste.`,
          `${name} could not be added — the roles list did not accept the entry. Try again shortly; if it persists, you lack a right on the list.`),
      );
      if (ok) {
        verwerfeBericht();
        setGewaehlt(null);
        setText('');
      }
    } catch (e) {
      console.warn('[AIUC] Rolle vergeben:', e);
      melde('fehler', t('Beim Vergeben ist etwas schiefgegangen. Prüf die Liste und versuch es noch einmal.', 'Something went wrong while assigning. Check the list and try again.'));
    } finally {
      setLaeuft(false);
    }
  }

  const beschreibung = (r: UserRole): string => r === 'Admin'
    ? t('Alles, was ein Organizer darf — und verwaltet die Rollen.', 'Everything an organizer may do — and manages the roles.')
    : t('Legt Use Cases an und pflegt sie: Texte, Links, Bilder, Status. Sieht das Protokoll.', 'Creates and maintains use cases: texts, links, images, status. Sees the log.');

  return (
    <div className="dex-ui-section">
      <div className="dex-ui-section-title">{t('Wer soll dazukommen?', 'Who should be added?')}</div>

      {gewaehlt ? (
        <div className="dex-ui-stack">
          <div className="dex-ui-row dex-ui-row--framed dex-ui-row--static">
            <Avatar name={gewaehlt.displayName || gewaehlt.email} />
            <span className="dex-ui-row-main">
              <span className="dex-ui-row-title" style={{ display: 'block' }}>{gewaehlt.displayName || gewaehlt.email}</span>
              <span className="dex-ui-row-sub" style={{ display: 'block' }}>
                {gewaehlt.email}{gewaehlt.jobTitle ? ` · ${gewaehlt.jobTitle}` : ''}
              </span>
            </span>
            <span className="dex-ui-row-actions">
              <button type="button" className="dex-ui-iconbtn" disabled={laeuft} onClick={() => setGewaehlt(null)} aria-label={t('Auswahl entfernen', 'Clear selection')} title={t('Auswahl entfernen', 'Clear selection')}>
                <X size={15} />
              </button>
            </span>
          </div>

          {bestehend ? (
            <div className="dex-ui-callout dex-ui-callout--warn" role="status">
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <span>
                {t(`Diese Person ist schon eingetragen: ${rolleLabel(bestehend.role)}. Ändern geht bewusst nur über die Auswahl in der Liste unten — so stuft sich niemand aus Versehen herab.`,
                  `This person is already listed: ${rolleLabel(bestehend.role)}. Changing it deliberately only works via the dropdown in the list below — that way nobody is downgraded by accident.`)}
              </span>
            </div>
          ) : (
            <>
              <div>
                <div className="dex-ui-label">{t('Welche Rolle bekommt die Person?', 'Which role does the person get?')}</div>
                <div style={{ display: 'grid', gap: 8 }}>
                  {VERGEBBAR.map(r => (
                    <button key={r} type="button" className={cx('dex-ui-choice', rolle === r && 'is-active')} style={{ padding: '10px 12px', alignItems: 'center' }} aria-pressed={rolle === r} disabled={laeuft} onClick={() => setRolle(r)}>
                      <span className="dex-ui-choice-body" style={{ display: 'block' }}>
                        <span className="dex-ui-choice-title" style={{ display: 'block' }}>{rolleLabel(r)}</span>
                        <span className="dex-ui-choice-desc" style={{ display: 'block' }}>{beschreibung(r)}</span>
                      </span>
                      <span className="dex-ui-choice-check">{rolle === r && <Check size={12} />}</span>
                    </button>
                  ))}
                </div>
                <p className="dex-ui-help">{t('User braucht keinen Eintrag: Wer nicht in der Liste steht, ist User.', 'User needs no entry: whoever is not in the list is a user.')}</p>
              </div>
              <div>
                <button type="button" className="btn btn-primary" disabled={laeuft} onClick={() => { void vergeben(); }}>
                  <Plus size={15} /> {laeuft ? t('Wird vergeben …', 'Assigning …') : t(`Als ${rolleLabel(rolle)} eintragen`, `Add as ${rolleLabel(rolle)}`)}
                </button>
              </div>
            </>
          )}
        </div>
      ) : (
        <div>
          <div className="dex-ui-searchbar" style={{ maxWidth: 480 }}>
            <span className="dex-ui-searchbar-icon"><Search size={15} /></span>
            <input
              id="rolle-suche"
              type="text"
              className="dex-ui-input"
              value={text}
              autoComplete="off"
              onChange={e => setText(e.target.value)}
              placeholder={t('Name oder E-Mail-Adresse …', 'Name or email address …')}
              aria-label={t('Person suchen', 'Search for a person')}
            />
          </div>
          <div style={{ marginTop: 8 }}>
            <label className="dex-ui-switch">
              <input type="checkbox" checked={international} onChange={e => setInternational(e.target.checked)} />
              <span className="dex-ui-switch-track" />
              <span className="dex-ui-switch-label" style={{ fontWeight: 500, fontSize: '0.82rem' }}>{t('Auch andere Deloitte-Adressen suchen', 'Also search other Deloitte addresses')}</span>
            </label>
          </div>

          <div style={{ marginTop: 10 }} aria-live="polite">
            {suche.status === 'sucht' && suche.treffer.length === 0 && (
              <div className="dex-ui-progress dex-ui-progress--indeterminate" aria-label={t('Sucht …', 'Searching …')}><div className="dex-ui-progress-bar" /></div>
            )}

            {suche.status === 'fehler' && (
              <div className="dex-ui-callout dex-ui-callout--warn" role="alert">
                <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
                <span className="dex-ui-callout-body">
                  {t(`Die Personensuche ist gerade nicht möglich${suche.http ? ` (HTTP ${suche.http})` : ''}. Das heißt nicht, dass es die Person nicht gibt — SharePoint hat die Suche nicht beantwortet, meist wegen Drosselung. Warte kurz und versuch es noch einmal.`,
                    `The people search is not possible right now${suche.http ? ` (HTTP ${suche.http})` : ''}. That does not mean the person does not exist — SharePoint did not answer the search, usually because of throttling. Wait a moment and try again.`)}{' '}
                  <button type="button" className="dex-ui-textlink" onClick={() => setNochmal(n => n + 1)}>{t('Erneut suchen', 'Search again')}</button>
                </span>
              </div>
            )}

            {suche.status === 'ok' && suche.treffer.length === 0 && (
              <div className="dex-ui-callout dex-ui-callout--neutral" role="status">
                <span className="dex-ui-callout-icon"><Info size={16} /></span>
                <span className="dex-ui-callout-body">
                  {suche.ausgeblendet > 0
                    ? t(`Es gibt ${suche.ausgeblendet} Treffer, aber ohne @deloitte.de-Adresse. Schalte „Auch andere Deloitte-Adressen suchen“ ein, wenn du sie meinst.`,
                      `There ${suche.ausgeblendet === 1 ? 'is 1 hit' : `are ${suche.ausgeblendet} hits`}, but without an @deloitte.de address. Switch on “Also search other Deloitte addresses” if you mean them.`)
                    : t(`Keine Treffer für „${text.trim()}“. Versuch den vollen Namen oder die E-Mail-Adresse.`, `No hits for “${text.trim()}”. Try the full name or the email address.`)}
                </span>
              </div>
            )}

            {suche.treffer.length > 0 && (
              <div className="dex-ui-stack" style={{ gap: 6, opacity: suche.status === 'sucht' ? 0.6 : 1 }}>
                {suche.treffer.map(h => {
                  const schon = bestehende(h.email);
                  return (
                    <button key={h.email} type="button" className="dex-ui-row dex-ui-row--framed dex-ui-rowbtn" onClick={() => { setGewaehlt(h); setText(''); }}>
                      <Avatar name={h.displayName || h.email} />
                      <span className="dex-ui-row-main">
                        <span className="dex-ui-row-title" style={{ display: 'block' }}>{h.displayName || h.email}</span>
                        <span className="dex-ui-row-sub" style={{ display: 'block' }}>{h.email}{h.jobTitle ? ` · ${h.jobTitle}` : ''}</span>
                      </span>
                      {schon && <span className={rollePille(schon.role)}>{t('hat schon', 'already has')}: {rolleLabel(schon.role, true)}</span>}
                    </button>
                  );
                })}
                {suche.ausgeblendet > 0 && (
                  <p className="dex-ui-help" style={{ margin: 0 }}>
                    {t(`${suche.ausgeblendet} weitere Treffer ohne @deloitte.de-Adresse sind ausgeblendet.`, `${suche.ausgeblendet} more hits without an @deloitte.de address are hidden.`)}
                  </p>
                )}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// --- Die Liste der Vergebenen -----------------------------------------------

type Befund = 'luecke' | 'ueberschuss' | 'alias' | 'ok' | 'nicht-geprueft';

interface ListeProps extends AktionsProps {
  bericht: RechteBericht | null;
}

function RollenListe(props: ListeProps): React.ReactElement {
  const { roles, isRolesLoading, rolesReadStatus, updateRole, removeRole, isSelf } = useRoles();
  const { t, isDe } = useLanguage();
  const { confirmDialog } = useDialog();
  const { laeuft, setLaeuft, melde, meldeRechte, pruefeStand, verwerfeBericht, bericht } = props;

  const [suche, setSuche] = React.useState('');
  const [filter, setFilter] = React.useState<'' | UserRole>('');

  const zaehl = (r: UserRole): number => roles.filter(x => x.role === r).length;

  // Wer mehrfach in der Liste steht: Gültig ist die ERSTE Zeile — eine spätere
  // Änderung an der zweiten bliebe wirkungslos (`RoleContext` nimmt per
  // `filter(...)[0]` die erste). Das ist kein Schönheitsfehler, sondern ein
  // Rechte-Fehler.
  const doppelte = React.useMemo(() => {
    const nachAdresse = new Map<string, RoleAssignment[]>();
    roles.forEach(r => {
      const k = (r.userEmail || '').trim().toLowerCase();
      if (!k) return;
      nachAdresse.set(k, (nachAdresse.get(k) || []).concat([r]));
    });
    const out: RoleAssignment[][] = [];
    nachAdresse.forEach(a => { if (a.length > 1) out.push(a); });
    return out;
  }, [roles]);

  const sichtbar = React.useMemo(() => {
    const q = suche.trim().toLowerCase();
    return roles
      .filter(r => !filter || r.role === filter)
      .filter(r => !q || (r.userName || '').toLowerCase().indexOf(q) >= 0 || (r.userEmail || '').toLowerCase().indexOf(q) >= 0)
      .slice()
      .sort((a, b) => (rolleRang(b.role) - rolleRang(a.role)) || (a.userName || a.userEmail).localeCompare(b.userName || b.userEmail, 'de'));
  }, [roles, suche, filter]);

  // Der Befund je Adresse aus dem letzten „Rechte prüfen".
  const befund = React.useMemo(() => {
    const m = new Map<string, Befund>();
    if (!bericht) return m;
    bericht.okAdressen.forEach(a => m.set(a, 'ok'));
    bericht.aliase.forEach(a => m.set(a.email.trim().toLowerCase(), 'alias'));
    bericht.ueberschuss.forEach(a => m.set(a.email.trim().toLowerCase(), 'ueberschuss'));
    bericht.luecken.forEach(a => m.set(a.email.trim().toLowerCase(), 'luecke'));
    return m;
  }, [bericht]);

  const befundPille = (r: RoleAssignment): React.ReactElement => {
    if (!bericht) return <span className="dex-ui-muted">—</span>;
    const b = befund.get((r.userEmail || '').trim().toLowerCase()) || 'nicht-geprueft';
    if (b === 'luecke') return <span className="dex-ui-pill dex-ui-pill--sm dex-ui-pill--red">{t('Rechte fehlen', 'Rights missing')}</span>;
    if (b === 'ueberschuss') return <span className="dex-ui-pill dex-ui-pill--sm dex-ui-pill--orange">{t('Zu viele Rechte', 'Too many rights')}</span>;
    if (b === 'alias') return <span className="dex-ui-pill dex-ui-pill--sm dex-ui-pill--blue" title={t('Rechte da, aber unter einer anderen Schreibweise der Adresse', 'Rights present, but under a different spelling of the address')}>{t('Andere Schreibweise', 'Other spelling')}</span>;
    if (b === 'ok') return r.role === 'User'
      ? <span className="dex-ui-pill dex-ui-pill--sm dex-ui-pill--gray">{t('ohne Rechte', 'no rights')}</span>
      : <span className="dex-ui-pill dex-ui-pill--sm dex-ui-pill--green">{t('Rechte ok', 'Rights ok')}</span>;
    return <span className="dex-ui-muted" title={t('Nach dem letzten Prüflauf hinzugekommen', 'Added after the last check')}>{t('nicht geprüft', 'not checked')}</span>;
  };

  const aenderungsText = (r: RoleAssignment, neu: UserRole): string => {
    const name = r.userName || r.userEmail;
    if (neu === 'User') {
      return t(`${name} auf „User“ setzen? Die Person verliert ihre Rechte auf den Listen der Plattform — Use Cases pflegen und Protokoll gehören dann nicht mehr dazu. Der Eintrag bleibt stehen; wer gar nichts mehr sein soll, wird stattdessen entfernt.`,
        `Set ${name} to “User”? The person loses their rights on the platform lists — maintaining use cases and the log no longer apply. The entry stays; whoever should be nothing at all is removed instead.`);
    }
    if (r.role === 'Admin') {
      return t(`${name} von Admin auf ${rolleLabel(neu)} setzen? Die Person verliert den Vollzugriff auf die Rollenliste und kann keine Rollen mehr verwalten.`,
        `Change ${name} from admin to ${rolleLabel(neu)}? The person loses full control of the roles list and can no longer manage roles.`);
    }
    if (neu === 'Admin') {
      return t(`${name} zum Admin machen? Admins verwalten alle Rollen und haben Vollzugriff auf die Rollenliste — auch für deine eigene.`,
        `Make ${name} an admin? Admins manage all roles and have full control of the roles list — including yours.`);
    }
    return t(`${name} zum ${rolleLabel(neu)} machen? Die Person kann dann Use Cases anlegen, ändern und löschen und sieht das Protokoll.`,
      `Make ${name} a ${rolleLabel(neu)}? The person can then create, change and delete use cases and sees the log.`);
  };

  async function aendern(r: RoleAssignment, neu: UserRole): Promise<void> {
    if (neu === r.role || laeuft) return;
    const name = r.userName || r.userEmail;
    const ja = await confirmDialog(aenderungsText(r, neu), {
      confirmLabel: t('Rolle ändern', 'Change role'),
      danger: rolleRang(neu) < rolleRang(r.role),
    });
    if (!ja) return;
    setLaeuft(true);
    try {
      const ok = await updateRole(r.id, neu);
      pruefeStand();
      meldeRechte(
        ok,
        t(`${name} ist jetzt ${rolleLabel(neu)}.`, `${name} is now ${rolleLabel(neu)}.`),
        t(`Die Rolle von ${name} konnte nicht geändert werden — die Zeile trägt weiterhin ${rolleLabel(r.role)}. Versuch es noch einmal.`,
          `The role of ${name} could not be changed — the row still carries ${rolleLabel(r.role)}. Try again.`),
      );
      if (ok) verwerfeBericht();
    } catch (e) {
      console.warn('[AIUC] Rolle ändern:', e);
      melde('fehler', t('Beim Ändern ist etwas schiefgegangen. Lade die Liste neu und prüf, was jetzt gilt.', 'Something went wrong while changing. Reload the list and check what applies now.'));
    } finally {
      setLaeuft(false);
    }
  }

  async function entfernen(r: RoleAssignment): Promise<void> {
    if (laeuft) return;
    const name = r.userName || r.userEmail;
    const ja = await confirmDialog(
      t(`Die Rolle von ${name} entfernen? Die Person verliert damit auch ihre Rechte auf den Listen der Plattform und ist danach wieder ein normaler User.`,
        `Remove the role of ${name}? The person also loses their rights on the platform lists and is a regular user again.`),
      { confirmLabel: t('Entfernen', 'Remove'), danger: true },
    );
    if (!ja) return;
    setLaeuft(true);
    try {
      const ok = await removeRole(r.id);
      pruefeStand();
      meldeRechte(
        ok,
        t(`Die Rolle von ${name} ist entfernt.`, `The role of ${name} has been removed.`),
        t(`Die Rolle von ${name} konnte nicht entfernt werden — die Zeile steht weiterhin in der Rollenliste. Versuch es noch einmal.`,
          `The role of ${name} could not be removed — the row is still in the roles list. Try again.`),
        // Die Zeile ist weg, also findet auch „Rechte prüfen" die Person nicht
        // mehr. Der Weg zurück: erneut eintragen und noch einmal entfernen — das
        // wiederholt den Entzug.
        t('Die Person hat damit womöglich weiterhin Zugriff auf die Listen. Trag sie erneut ein und entferne sie noch einmal — dann wird der Entzug wiederholt.',
          'The person may therefore still have access to the lists. Add them again and remove them once more — that repeats the revocation.'),
      );
      if (ok) verwerfeBericht();
    } catch (e) {
      console.warn('[AIUC] Rolle entfernen:', e);
      melde('fehler', t('Beim Entfernen ist etwas schiefgegangen. Lade die Liste neu und prüf, was jetzt gilt.', 'Something went wrong while removing. Reload the list and check what applies now.'));
    } finally {
      setLaeuft(false);
    }
  }

  const kpi = (wert: '' | UserRole, label: string, n: number): React.ReactElement => (
    <button
      key={wert || 'alle'}
      type="button"
      className={cx('dex-ui-kpi', 'is-clickable', 'dex-ui-btn-reset', filter === wert && 'is-active')}
      style={{ textAlign: 'left' }}
      aria-pressed={filter === wert}
      onClick={() => setFilter(filter === wert ? '' : wert)}
    >
      <div className="dex-ui-kpi-value">{n}</div>
      <div className="dex-ui-kpi-label">{label}</div>
    </button>
  );

  return (
    <div className="dex-ui-section">
      <div className="dex-ui-section-title">{t('Vergebene Rollen', 'Assigned roles')}</div>

      {isRolesLoading ? (
        <div className="dex-ui-progress dex-ui-progress--indeterminate"><div className="dex-ui-progress-bar" /></div>
      ) : roles.length === 0 && rolesReadStatus !== 'ok' ? (
        // Nicht „Noch niemand eingetragen": Die Liste konnte nicht gelesen werden,
        // und leer und unlesbar sind zwei verschiedene Aussagen.
        <p className="dex-ui-help" style={{ margin: 0 }}>{t('Die Rollenliste ist nicht lesbar — deshalb lässt sich nicht sagen, wer eingetragen ist.', 'The roles list is not readable — so it cannot be said who is listed.')}</p>
      ) : roles.length === 0 ? (
        <p className="dex-ui-help" style={{ margin: 0 }}>{t('Noch niemand eingetragen.', 'Nobody assigned yet.')}</p>
      ) : (
        <>
          {doppelte.length > 0 && (
            <div className="dex-ui-callout dex-ui-callout--warn" role="alert" style={{ marginBottom: 12 }}>
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <span className="dex-ui-callout-body">
                <strong>
                  {doppelte.length === 1
                    ? t('Eine Person hat mehrere Einträge in der Rollenliste', 'One person has several entries in the roles list')
                    : t(`${doppelte.length} Personen haben mehrere Einträge in der Rollenliste`, `${doppelte.length} people have several entries in the roles list`)}
                </strong>
                <br />
                {t('Gültig ist die erste Zeile — eine Änderung an einer späteren bleibt wirkungslos. Entferne die überzähligen: ',
                  'The first row wins — a change to a later one has no effect. Remove the surplus ones: ')}
                {doppelte.slice(0, 8).map(a => `${a[0].userName || a[0].userEmail} (${a.map(x => rolleLabel(x.role, true)).join(', ')})`).join('; ')}
                {doppelte.length > 8 ? ` … +${doppelte.length - 8}` : ''}
              </span>
            </div>
          )}

          <div className="dex-ui-kpi-row" style={{ marginBottom: 12 }}>
            {kpi('', t('Alle', 'All'), roles.length)}
            {ROLLEN_ALLE.map(r => kpi(r, rolleLabel(r), zaehl(r)))}
          </div>

          <div className="dex-ui-toolbar">
            <div className="dex-ui-searchbar">
              <span className="dex-ui-searchbar-icon"><Search size={15} /></span>
              <input
                type="text"
                id="rolle-filter"
                className="dex-ui-input dex-ui-input--sm"
                value={suche}
                onChange={e => setSuche(e.target.value)}
                placeholder={t('In der Liste suchen …', 'Search the list …')}
                aria-label={t('In der Liste suchen', 'Search the list')}
                style={{ paddingRight: 34 }}
              />
              {suche && (
                <button type="button" className="dex-ui-iconbtn" onClick={() => setSuche('')} aria-label={t('Suche leeren', 'Clear search')} style={{ position: 'absolute', right: 3, top: '50%', transform: 'translateY(-50%)', width: 32, height: 32 }}>
                  <X size={14} />
                </button>
              )}
            </div>
            <span className="dex-ui-toolbar-spacer" />
            <span className="dex-ui-muted">
              {sichtbar.length === roles.length ? `${roles.length}` : `${sichtbar.length} ${t('von', 'of')} ${roles.length}`}
            </span>
          </div>

          {sichtbar.length === 0 ? (
            <div className="dex-ui-callout dex-ui-callout--neutral" role="status">
              <span className="dex-ui-callout-icon"><Info size={16} /></span>
              <span className="dex-ui-callout-body">
                {t('Keine Einträge für diese Auswahl.', 'No entries for this selection.')}{' '}
                <button type="button" className="dex-ui-textlink" onClick={() => { setSuche(''); setFilter(''); }}>{t('Auswahl zurücksetzen', 'Reset selection')}</button>
              </span>
            </div>
          ) : (
            <div className="dex-ui-table-wrap">
              <table className="dex-ui-table">
                <thead>
                  <tr>
                    <th>{t('Person', 'Person')}</th>
                    <th>{t('Rolle', 'Role')}</th>
                    <th>{t('Vergeben', 'Assigned')}</th>
                    <th>{t('SharePoint-Rechte', 'SharePoint rights')}</th>
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {sichtbar.map(r => {
                    const ich = isSelf(r.userEmail);
                    const name = r.userName || r.userEmail;
                    const datum = datumText(r.assignedDate, isDe);
                    return (
                      <tr key={r.id}>
                        <td>
                          <span className="dex-ui-person">
                            <Avatar name={name} />
                            <span style={{ minWidth: 0, maxWidth: 300 }}>
                              <span className="dex-ui-person-name" style={{ display: 'block' }}>
                                {name}{ich && <span className="dex-ui-pill dex-ui-pill--sm dex-ui-pill--blue" style={{ marginLeft: 8 }}>{t('du', 'you')}</span>}
                              </span>
                              <span className="dex-ui-person-sub" style={{ display: 'block' }}>{r.userEmail}</span>
                            </span>
                          </span>
                        </td>
                        <td>
                          {ich ? (
                            <span className={rollePille(r.role)} title={t('Die eigene Rolle ändert ein anderer Admin — wer sich selbst herabstuft, käme nicht mehr an diese Seite.', 'Another admin changes your own role — whoever downgrades themselves could no longer reach this page.')}>{rolleLabel(r.role, true)}</span>
                          ) : (
                            <select
                              className="dex-ui-select dex-ui-select--sm"
                              value={r.role}
                              disabled={laeuft}
                              aria-label={`${t('Rolle ändern', 'Change role')}: ${name}`}
                              onChange={e => { void aendern(r, e.target.value as UserRole); }}
                            >
                              {ROLLEN_ALLE.map(x => <option key={x} value={x}>{rolleLabel(x)}</option>)}
                            </select>
                          )}
                        </td>
                        <td className="dex-ui-muted">
                          {r.assignedBy ? `${t('von', 'by')} ${r.assignedBy}` : ''}{r.assignedBy && datum ? ' · ' : ''}{datum}
                        </td>
                        <td>{befundPille(r)}</td>
                        <td className="is-actions">
                          {!ich && (
                            <button type="button" className="dex-ui-iconbtn dex-ui-iconbtn--danger" disabled={laeuft} onClick={() => { void entfernen(r); }} title={t('Rolle entfernen', 'Remove role')} aria-label={`${t('Rolle entfernen', 'Remove role')}: ${name}`}>
                              <Trash2 size={15} />
                            </button>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}
    </div>
  );
}

// --- Rechte prüfen ----------------------------------------------------------

interface PruefProps {
  laeuft: boolean;
  setLaeuft: (b: boolean) => void;
  bericht: RechteBericht | null;
  setBericht: (b: RechteBericht | null) => void;
}

function RechtePruefung(props: PruefProps): React.ReactElement {
  const { auditRoleRights, lastRightsReadError, repairRoleRights, revokeExcessRights, rolesReadStatus } = useRoles();
  const { t, isDe } = useLanguage();
  const { confirmDialog } = useDialog();
  const { laeuft, setLaeuft, bericht, setBericht } = props;

  const [phase, setPhase] = React.useState<'ruht' | 'prueft' | 'setzt' | 'entzieht'>('ruht');
  const [fortschritt, setFortschritt] = React.useState({ done: 0, total: 0 });
  // Warum das Lesen gescheitert ist. Steht getrennt vom Bericht: Ohne Bericht
  // gibt es KEINE Aussage über die Rechte — weder „in Ordnung" noch „fehlt".
  const [lesefehler, setLesefehler] = React.useState('');
  // Was der letzte Schritt (nachsetzen / entziehen) getan hat.
  const [aktion, setAktion] = React.useState<{ ton: 'ok' | 'warn'; text: string } | null>(null);

  const fortschrittFn = (done: number, total: number): void => setFortschritt({ done, total });

  /** Ein Lauf: lesen, prüfen, Bericht setzen. Gibt zurück, ob er gelesen werden konnte. */
  async function lauf(): Promise<boolean> {
    setPhase('prueft');
    setFortschritt({ done: 0, total: 0 });
    setLesefehler('');
    try {
      const res = await auditRoleRights(fortschrittFn);
      if (res === null) {
        setBericht(null);
        setLesefehler(lastRightsReadError() || t('keine Antwort', 'no answer'));
        return false;
      }
      setBericht(res);
      return true;
    } catch (e) {
      console.warn('[AIUC] Rechte prüfen:', e);
      setBericht(null);
      setLesefehler(e instanceof Error ? e.message : t('unbekannter Fehler', 'unknown error'));
      return false;
    }
  }

  async function pruefen(): Promise<void> {
    if (laeuft) return;
    setLaeuft(true);
    setAktion(null);
    try {
      await lauf();
    } finally {
      setPhase('ruht');
      setLaeuft(false);
    }
  }

  async function nachsetzen(): Promise<void> {
    if (!bericht || laeuft || bericht.luecken.length === 0) return;
    const n = bericht.luecken.length;
    const ja = await confirmDialog(
      t(`Bei ${n} ${n === 1 ? 'Person' : 'Personen'} fehlende Rechte nachsetzen? Es wird nur ergänzt, nichts entzogen. Jedes Recht wird danach nachgelesen — ein Recht gilt erst, wenn es dasteht.`,
        `Re-grant the missing rights for ${n} ${n === 1 ? 'person' : 'people'}? Rights are only added, nothing is revoked. Each right is read back afterwards — a right only counts once it is there.`),
      { confirmLabel: t('Jetzt nachsetzen', 'Re-grant now') },
    );
    if (!ja) return;
    setLaeuft(true);
    setPhase('setzt');
    setFortschritt({ done: 0, total: n });
    try {
      const res = await repairRoleRights(bericht.luecken, fortschrittFn);
      const offenText = res.offen.map(o => `${o.name || o.email} (${o.fehlt.map(k => listeMitRecht(k, o.rolle, isDe)).join(', ')})${o.kontoAufloesbar ? '' : (isDe ? ' — Konto nicht auflösbar' : ' — account not resolvable')}`).join('; ');
      setAktion({
        ton: res.offen.length > 0 ? 'warn' : 'ok',
        text: isDe
          ? `${res.behoben.length} von ${n} nachgesetzt und nachgelesen.${res.offen.length > 0 ? ` NICHT setzbar: ${offenText}. Häufige Ursachen: Das Konto ist ausgeschieden, die Person hat die Site nie besucht und SharePoint löst sie nicht auf, oder die Anfragen wurden gedrosselt — dann in ein paar Minuten erneut prüfen.` : ''}`
          : `${res.behoben.length} of ${n} re-granted and read back.${res.offen.length > 0 ? ` NOT grantable: ${offenText}. Common causes: the account has left, the person never visited the site and SharePoint cannot resolve them, or requests were throttled — check again in a few minutes.` : ''}`,
      });
      // Die Kontrolle danach ist ein NEUER Lesevorgang, nicht das Ergebnis des
      // Schreibens: Der Bericht unten zeigt, was jetzt gilt.
      setPhase('prueft');
      const gelesen = await lauf();
      if (!gelesen) {
        setAktion(a => a && { ton: 'warn', text: `${a.text} ${isDe ? 'Die Kontrolle danach konnte nicht lesen — der Stand unten fehlt.' : 'The check afterwards could not read — the state below is missing.'}` });
      }
    } catch (e) {
      console.warn('[AIUC] Rechte nachsetzen:', e);
      setAktion({ ton: 'warn', text: t('Beim Nachsetzen ist etwas schiefgegangen. Prüf die Rechte noch einmal — es kann sein, dass ein Teil schon gesetzt ist.', 'Something went wrong while re-granting. Check the rights again — part of it may already be set.') });
    } finally {
      setPhase('ruht');
      setLaeuft(false);
    }
  }

  async function entziehen(): Promise<void> {
    if (!bericht || laeuft || bericht.ueberschuss.length === 0) return;
    const items = bericht.ueberschuss;
    const namen = items.map(u => u.name || u.email).join(', ');
    const ja = await confirmDialog(
      t(`Diese Personen tragen mehr Rechte, als ihre Rolle vorsieht: ${namen}. Die überzähligen Rechte auf den Listen werden entzogen; was die Rolle vorsieht, bleibt. Wenn jemand diese Rechte mit Absicht direkt in SharePoint bekommen hat, geht das damit verloren. Jetzt entziehen?`,
        `These people hold more rights than their role provides: ${namen}. The surplus rights on the lists are revoked; what the role provides stays. If someone was given these rights on purpose directly in SharePoint, they are lost. Revoke now?`),
      { confirmLabel: t('Jetzt entziehen', 'Revoke now'), danger: true },
    );
    if (!ja) return;
    setLaeuft(true);
    setPhase('entzieht');
    setFortschritt({ done: 0, total: items.length });
    try {
      const res = await revokeExcessRights(items, fortschrittFn);
      const offenText = res.offen.map(o => `${o.name || o.email} (${o.listen.map(k => listenName(k, isDe)).join(', ')})`).join('; ');
      setAktion({
        ton: res.offen.length > 0 ? 'warn' : 'ok',
        text: isDe
          ? `Bei ${res.erledigt} von ${items.length} entzogen.${res.offen.length > 0 ? ` NICHT entzogen: ${offenText}. Entweder hat sich die Rolle seit der Prüfung geändert, oder SharePoint hat den Entzug nicht bestätigt — prüf noch einmal.` : ''}`
          : `Revoked for ${res.erledigt} of ${items.length}.${res.offen.length > 0 ? ` NOT revoked: ${offenText}. Either the role changed since the check, or SharePoint did not confirm the revocation — check again.` : ''}`,
      });
      setPhase('prueft');
      const gelesen = await lauf();
      if (!gelesen) {
        setAktion(a => a && { ton: 'warn', text: `${a.text} ${isDe ? 'Die Kontrolle danach konnte nicht lesen — der Stand unten fehlt.' : 'The check afterwards could not read — the state below is missing.'}` });
      }
    } catch (e) {
      console.warn('[AIUC] Rechte entziehen:', e);
      setAktion({ ton: 'warn', text: t('Beim Entziehen ist etwas schiefgegangen. Prüf die Rechte noch einmal.', 'Something went wrong while revoking. Check the rights again.') });
    } finally {
      setPhase('ruht');
      setLaeuft(false);
    }
  }

  const laufend = phase !== 'ruht';
  const prozent = fortschritt.total > 0 ? Math.min(100, Math.round((fortschritt.done / fortschritt.total) * 100)) : 0;
  const laufText = ((): string => {
    if (phase === 'prueft') return fortschritt.total > 0 ? t(`Prüft ${fortschritt.done} von ${fortschritt.total} …`, `Checking ${fortschritt.done} of ${fortschritt.total} …`) : t('Liest Rollen und Berechtigungen …', 'Reading roles and permissions …');
    if (phase === 'setzt') return t(`Setzt nach: ${fortschritt.done} von ${fortschritt.total} …`, `Re-granting: ${fortschritt.done} of ${fortschritt.total} …`);
    return t(`Entzieht: ${fortschritt.done} von ${fortschritt.total} …`, `Revoking: ${fortschritt.done} of ${fortschritt.total} …`);
  })();

  const luecken = bericht ? bericht.luecken : [];
  const ueberschuss = bericht ? bericht.ueberschuss : [];
  const aliase = bericht ? bericht.aliase : [];
  const sauber = !!bericht && luecken.length === 0 && ueberschuss.length === 0;

  return (
    <div className="dex-ui-section">
      <div className="dex-ui-section-title">{t('SharePoint-Rechte der Rollen', 'SharePoint rights of the roles')}</div>
      <p className="dex-ui-section-desc">
        {t('Eine Rolle wirkt erst, wenn die Rechte auf den Listen wirklich stehen: Lesen auf der Rollenliste (sonst bleibt die Person trotz Eintrag ein normaler User), Bearbeiten auf den Use Cases und Beitragen im Protokoll (sonst scheitert das Pflegen). Die Prüfung liest das nach und ändert nichts — Nachsetzen ist ein eigener Schritt.',
          'A role only applies once the rights on the lists are really set: read on the roles list (otherwise the person stays a regular user despite the entry), edit on the use cases and contribute on the log (otherwise maintaining fails). The check reads this back and changes nothing — re-granting is a separate step.')}
      </p>

      <div className="dex-ui-inline">
        <button type="button" className="btn btn-secondary" disabled={laeuft || rolesReadStatus !== 'ok'} onClick={() => { void pruefen(); }}>
          <ShieldCheck size={15} /> {phase === 'prueft' && !aktion ? t('Prüft …', 'Checking …') : t('Rechte prüfen', 'Check rights')}
        </button>
        {bericht && !laufend && (
          <span className="dex-ui-muted">{t(`${bericht.geprueft} Einträge geprüft`, `${bericht.geprueft} entries checked`)}</span>
        )}
      </div>

      {laufend && (
        <div style={{ marginTop: 12 }} role="status" aria-live="polite">
          {fortschritt.total > 0
            ? <div className="dex-ui-progress"><div className="dex-ui-progress-bar" style={{ width: `${prozent}%` }} /></div>
            : <div className="dex-ui-progress dex-ui-progress--indeterminate"><div className="dex-ui-progress-bar" /></div>}
          <p className="dex-ui-help">{laufText}</p>
        </div>
      )}

      {aktion && (
        <div className={cx('dex-ui-callout', aktion.ton === 'ok' ? 'dex-ui-callout--success' : 'dex-ui-callout--warn')} role="status" style={{ marginTop: 12 }}>
          <span className="dex-ui-callout-icon">{aktion.ton === 'ok' ? <Check size={16} /> : <AlertCircle size={16} />}</span>
          <span className="dex-ui-callout-body">{aktion.text}</span>
        </div>
      )}

      {/* Nicht lesbar ist KEIN „in Ordnung": ohne Bericht keine Aussage. */}
      {lesefehler && !laufend && (
        <div className="dex-ui-callout dex-ui-callout--danger" role="alert" style={{ marginTop: 12 }}>
          <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
          <span className="dex-ui-callout-body">
            {t(`Die Rechte konnten nicht gelesen werden (${lesefehler}). Darüber lässt sich deshalb nichts sagen — weder, dass alles stimmt, noch, dass etwas fehlt. Versuch es in ein paar Minuten noch einmal; bleibt es dabei, fehlt dir das Recht, die Berechtigungen der Listen zu sehen.`,
              `The rights could not be read (${lesefehler}). Nothing can be said about them — neither that everything is fine nor that something is missing. Try again in a few minutes; if it persists, you lack the right to see the permissions of the lists.`)}
          </span>
        </div>
      )}

      {bericht && !laufend && (
        <div className="dex-ui-stack" style={{ marginTop: 12 }}>
          {sauber && (
            <div className="dex-ui-callout dex-ui-callout--success" role="status">
              <span className="dex-ui-callout-icon"><Check size={16} /></span>
              <span className="dex-ui-callout-body">
                {t(`${bericht.geprueft} Einträge geprüft — bei allen stehen die Rechte, die ihre Rolle braucht.`, `${bericht.geprueft} entries checked — everyone has the rights their role needs.`)}
              </span>
            </div>
          )}

          {luecken.length > 0 && (
            <div className="dex-ui-callout dex-ui-callout--warn" role="alert">
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <span className="dex-ui-callout-body">
                <strong>{luecken.length === 1 ? t('Bei 1 Person fehlt ein Recht', 'One person lacks a right') : t(`Bei ${luecken.length} Personen fehlt mindestens ein Recht`, `${luecken.length} people lack at least one right`)}</strong>
                <ul style={{ margin: '6px 0 8px', paddingLeft: 18 }}>
                  {luecken.map(l => (
                    <li key={l.email}>
                      {l.name || l.email} <span className="dex-ui-muted">({rolleLabel(l.rolle, true)})</span>: {l.fehlt.map(k => listeMitRecht(k, l.rolle, isDe)).join(', ')}
                      {!l.kontoAufloesbar && <> — <em>{t('Konto nicht auflösbar (ausgeschieden oder gedrosselt)', 'account not resolvable (left or throttled)')}</em></>}
                    </li>
                  ))}
                </ul>
                <button type="button" className="btn btn-primary dex-ui-btn-sm" disabled={laeuft} onClick={() => { void nachsetzen(); }}>
                  {t('Fehlende Rechte nachsetzen', 'Re-grant missing rights')}
                </button>
              </span>
            </div>
          )}

          {ueberschuss.length > 0 && (
            <div className="dex-ui-callout dex-ui-callout--warn" role="alert">
              <span className="dex-ui-callout-icon"><AlertCircle size={16} /></span>
              <span className="dex-ui-callout-body">
                <strong>{ueberschuss.length === 1 ? t('Eine Person hat mehr Rechte, als ihre Rolle vorsieht', 'One person has more rights than their role provides') : t(`${ueberschuss.length} Personen haben mehr Rechte, als ihre Rolle vorsieht`, `${ueberschuss.length} people have more rights than their role provides`)}</strong>
                <ul style={{ margin: '6px 0 8px', paddingLeft: 18 }}>
                  {ueberschuss.map(u => (
                    <li key={u.email}>
                      {u.name || u.email} <span className="dex-ui-muted">({rolleLabel(u.rolle, true)})</span>:{' '}
                      {u.rolle === 'Organizer' && u.listen.length === 1 && u.listen[0] === 'roles'
                        ? t('kann die Rollenliste bearbeiten (Rest einer früheren Admin-Rolle?) — vorgesehen ist nur Lesen', 'can edit the roles list (left over from a former admin role?) — only read is provided')
                        : t(`trägt noch direkte Rechte auf: ${u.listen.map(k => listenName(k, true)).join(', ')} — als User braucht sie keine`, `still holds direct rights on: ${u.listen.map(k => listenName(k, false)).join(', ')} — as a user they need none`)}
                    </li>
                  ))}
                </ul>
                <button type="button" className="btn btn-secondary dex-ui-btn-sm" disabled={laeuft} onClick={() => { void entziehen(); }}>
                  {t('Überzählige Rechte entziehen', 'Revoke surplus rights')}
                </button>
              </span>
            </div>
          )}

          {aliase.length > 0 && (
            <div className="dex-ui-callout dex-ui-callout--info" role="status">
              <span className="dex-ui-callout-icon"><Info size={16} /></span>
              <span className="dex-ui-callout-body">
                <strong>{t('Rechte da — aber unter einer anderen Schreibweise der Adresse', 'Rights present — but under a different spelling of the address')}</strong>
                <ul style={{ margin: '6px 0 6px', paddingLeft: 18 }}>
                  {aliase.map(a => (
                    <li key={a.email}>{a.name || a.email}: {t('Rollenliste', 'roles list')} <code>{a.email}</code> · SharePoint <code>{a.spEmail}</code></li>
                  ))}
                </ul>
                {t('Die Rechte stimmen; die Zeile in der Rollenliste führt nur eine andere Adresse als das Konto. Am saubersten: die Zeile entfernen und die Person über die Suche neu eintragen — die Suche liefert die Adresse, die SharePoint führt.',
                  'The rights are fine; the row in the roles list merely carries a different address than the account. Cleanest: remove the row and add the person again via the search — the search returns the address SharePoint holds.')}
              </span>
            </div>
          )}

          {bericht.ohneAdresse > 0 && (
            <div className="dex-ui-callout dex-ui-callout--neutral" role="status">
              <span className="dex-ui-callout-icon"><Info size={16} /></span>
              <span className="dex-ui-callout-body">
                {t(`${bericht.ohneAdresse} ${bericht.ohneAdresse === 1 ? 'Zeile hat' : 'Zeilen haben'} keine E-Mail-Adresse und ${bericht.ohneAdresse === 1 ? 'konnte' : 'konnten'} nicht geprüft werden.`,
                  `${bericht.ohneAdresse} ${bericht.ohneAdresse === 1 ? 'row has' : 'rows have'} no email address and could not be checked.`)}
              </span>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Reiter „Rechte je Rolle"
// ---------------------------------------------------------------------------

function MatrixZelle(props: { wert: MatrixWert; isDe: boolean }): React.ReactElement {
  const { wert, isDe } = props;
  if (wert === true) {
    return <span role="img" aria-label={isDe ? 'ja' : 'yes'} style={{ display: 'inline-flex', color: 'var(--dex-green-dark, #4a7c1f)' }}><Check size={16} /></span>;
  }
  if (wert === false) {
    return <span role="img" aria-label={isDe ? 'nein' : 'no'} className="dex-ui-muted">—</span>;
  }
  return <span className="dex-ui-pill dex-ui-pill--sm dex-ui-pill--orange">{isDe ? wert.de : wert.en}</span>;
}

function RechteMatrix(): React.ReactElement {
  const { currentUserRole } = useRoles();
  const { t, isDe } = useLanguage();

  return (
    <div>
      <div className="dex-ui-callout dex-ui-callout--info" style={{ marginBottom: 16 }}>
        <span className="dex-ui-callout-icon"><Users size={16} /></span>
        <span className="dex-ui-callout-body">
          {t('Was darf wer? Das ist eine reine Ansicht — sie ändert nichts. Welche Rolle jemand hat, steht in der Rollenliste; ohne Eintrag ist jede Person ein User. Die Rolle wirkt erst mit Leserecht auf der Rollenliste, und die Rechte auf den Listen prüft und setzt der andere Reiter.',
            'Who may do what? This is a plain view — it changes nothing. Which role someone has is stored in the roles list; without an entry every person is a user. The role only applies with read access to the roles list, and the other tab checks and sets the rights on the lists.')}
        </span>
      </div>

      <div className="dex-ui-table-wrap dex-ui-table-wrap--sticky">
        <table className="dex-ui-table">
          <thead>
            <tr>
              <th style={{ minWidth: 260 }}>{t('Funktion', 'Function')}</th>
              {MATRIX_SPALTEN.map(sp => {
                const rolle: UserRole = sp === 'admin' ? 'Admin' : sp === 'organizer' ? 'Organizer' : 'User';
                return (
                  <th key={sp} style={{ textAlign: 'center', minWidth: 96 }}>
                    {rolleLabel(rolle)}
                    {currentUserRole === rolle && <div style={{ marginTop: 4 }}><span className="dex-ui-pill dex-ui-pill--sm dex-ui-pill--green" style={{ textTransform: 'none', letterSpacing: 0 }}>{t('deine Rolle', 'your role')}</span></div>}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {RECHTE_MATRIX.map(k => (
              <React.Fragment key={k.key}>
                <tr>
                  <td colSpan={1 + MATRIX_SPALTEN.length} style={{ background: 'var(--dex-gray-50, #fafafa)', fontWeight: 700, fontSize: '0.72rem', textTransform: 'uppercase', letterSpacing: '0.06em', color: 'var(--dex-gray-500, #757575)' }}>
                    {isDe ? k.de : k.en}
                  </td>
                </tr>
                {k.zeilen.map(z => (
                  <tr key={z.key}>
                    <td>
                      <div style={{ fontWeight: 600 }}>{isDe ? z.de : z.en}</div>
                      <div className="dex-ui-help" style={{ margin: '2px 0 0' }}>{isDe ? z.descDe : z.descEn}</div>
                    </td>
                    {MATRIX_SPALTEN.map(sp => (
                      <td key={sp} style={{ textAlign: 'center' }}><MatrixZelle wert={z[sp]} isDe={isDe} /></td>
                    ))}
                  </tr>
                ))}
              </React.Fragment>
            ))}
          </tbody>
        </table>
      </div>

      <p className="dex-ui-help" style={{ marginTop: 10 }}>
        <span style={{ display: 'inline-flex', verticalAlign: 'middle', color: 'var(--dex-green-dark, #4a7c1f)' }}><Check size={14} /></span> {t('ja', 'yes')} · — {t('nein', 'no')} · <span className="dex-ui-pill dex-ui-pill--sm dex-ui-pill--orange">{t('Text', 'text')}</span> {t('ja, mit der genannten Einschränkung', 'yes, with the stated limitation')}
      </p>
      <p className="dex-ui-help" style={{ marginTop: 4 }}>
        {t('Die erste Person, die die Plattform öffnet, solange die Rollenliste leer ist, wird automatisch Admin. Ist die Rollenliste nicht lesbar, wird niemand befördert.',
          'The first person to open the platform while the roles list is empty automatically becomes admin. If the roles list is not readable, nobody is promoted.')}
      </p>
    </div>
  );
}
